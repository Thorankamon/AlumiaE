#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Serveur local pour AlumiaNote.
- Sert la page alumia-note.html (et les autres fichiers du dossier)
- Stocke chaque note dans une vraie arborescence de dossiers sur le disque :
      DonneesNotes/<Bloc-notes>/<Catégorie>/<Titre__id>.json   (source de vérité, relue par l'app)
      DonneesNotes/<Bloc-notes>/<Catégorie>/<Titre__id>.docx   (vrai fichier Word, régénéré à chaque sauvegarde)
- Expose une petite API utilisée par la page :
      GET  /api/load   -> renvoie toutes les notes (reconstruites depuis les fichiers)
      POST /api/save   -> reçoit toutes les notes, réécrit l'arborescence + les .docx

Lancement :
    pip install python-docx        (une seule fois)
    python serveur_notes.py
Puis ouvrir dans le navigateur :  http://localhost:8791/alumia-note.html
"""

import json
import os
import re
import io
import base64
import hashlib
import http.server
import socketserver
from html.parser import HTMLParser
from urllib.parse import urlparse

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

PORT = 8791
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "DonneesNotes")
INDEX_PATH = os.path.join(DATA_DIR, "_index.json")

os.makedirs(DATA_DIR, exist_ok=True)

# ───────────────────────── Utilitaires fichiers ─────────────────────────

def sanitize(name, fallback="sans_nom"):
    name = (name or "").strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = name.strip(". ")
    if not name:
        name = fallback
    return name[:80]


def short_id(id_str):
    return re.sub(r"[^a-zA-Z0-9]", "", str(id_str))[-8:] or "id"


def content_hash(title, content):
    h = hashlib.md5()
    h.update((title or "").encode("utf-8"))
    h.update((content or "").encode("utf-8"))
    return h.hexdigest()


def load_index():
    if os.path.exists(INDEX_PATH):
        try:
            with open(INDEX_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None


def write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


# ───────────────────────── Conversion HTML -> DOCX ─────────────────────────

FONT_SIZE_MAP = {"1": 8, "2": 10, "3": 12, "4": 14, "5": 18, "6": 24, "7": 32}

NAMED_COLORS = {
    "black": (0, 0, 0), "white": (255, 255, 255), "red": (220, 38, 38),
    "green": (22, 163, 74), "blue": (37, 99, 235), "yellow": (234, 179, 8),
}


def parse_color(value):
    if not value:
        return None
    value = value.strip().lower()
    if value.startswith("#"):
        v = value[1:]
        if len(v) == 3:
            v = "".join(c * 2 for c in v)
        if len(v) == 6:
            try:
                return tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))
            except ValueError:
                return None
    return NAMED_COLORS.get(value)


class HtmlToDocx(HTMLParser):
    def __init__(self, doc):
        super().__init__(convert_charrefs=True)
        self.doc = doc
        self.paragraph = None
        self.bold = 0
        self.italic = 0
        self.underline = 0
        self.strike = 0
        self.color = None
        self.size = None
        self.align = None
        self.list_stack = []
        self.skip_depth = 0

    def new_paragraph(self):
        self.paragraph = self.doc.add_paragraph()
        if self.align == "center":
            self.paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif self.align == "right":
            self.paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def ensure_paragraph(self):
        if self.paragraph is None:
            self.new_paragraph()

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if self.skip_depth > 0:
            self.skip_depth += 1
            return
        cls = attrs.get("class", "")
        if tag in ("div", "p"):
            if "audio-block" in cls:
                self.skip_depth = 1
                self.ensure_paragraph()
                run = self.paragraph.add_run("[Note audio - disponible dans AlumiaNote]")
                run.italic = True
                return
            style = attrs.get("style", "")
            m = re.search(r"text-align:\s*(\w+)", style)
            if m:
                self.align = m.group(1)
            self.new_paragraph()
        elif tag == "br":
            self.ensure_paragraph()
            self.paragraph.add_run().add_break()
        elif tag in ("b", "strong"):
            self.bold += 1
        elif tag in ("i", "em"):
            self.italic += 1
        elif tag == "u":
            self.underline += 1
        elif tag in ("strike", "s", "del"):
            self.strike += 1
        elif tag == "font":
            if attrs.get("color"):
                self.color = attrs["color"]
            if attrs.get("size"):
                self.size = attrs["size"]
        elif tag == "span":
            style = attrs.get("style", "")
            mc = re.search(r"(?<!-)color:\s*([#\w]+)", style)
            if mc:
                self.color = mc.group(1)
        elif tag in ("ul", "ol"):
            self.list_stack.append(tag)
        elif tag == "li":
            self.new_paragraph()
            level = max(1, len(self.list_stack))
            bullet = "•  " if (self.list_stack and self.list_stack[-1] == "ul") else "-  "
            self.paragraph.paragraph_format.left_indent = Inches(0.25 * level)
            self.paragraph.add_run(bullet)
        elif tag == "img":
            self.ensure_paragraph()
            src = attrs.get("src", "")
            if src.startswith("data:image"):
                try:
                    _, b64data = src.split(",", 1)
                    imgdata = base64.b64decode(b64data)
                    self.doc.add_picture(io.BytesIO(imgdata), width=Inches(4))
                    self.paragraph = None
                except Exception:
                    pass
        elif re.match(r"h[1-6]$", tag):
            self.new_paragraph()
            self.bold += 1
            self.size = {"h1": "7", "h2": "6", "h3": "5"}.get(tag, "4")

    def handle_endtag(self, tag):
        if self.skip_depth > 0:
            self.skip_depth -= 1
            return
        if tag in ("b", "strong"):
            self.bold = max(0, self.bold - 1)
        elif tag in ("i", "em"):
            self.italic = max(0, self.italic - 1)
        elif tag == "u":
            self.underline = max(0, self.underline - 1)
        elif tag in ("strike", "s", "del"):
            self.strike = max(0, self.strike - 1)
        elif tag == "font":
            self.color = None
            self.size = None
        elif tag == "span":
            self.color = None
        elif tag in ("ul", "ol"):
            if self.list_stack:
                self.list_stack.pop()
        elif re.match(r"h[1-6]$", tag):
            self.bold = max(0, self.bold - 1)
            self.size = None

    def handle_data(self, data):
        if self.skip_depth > 0 or not data:
            return
        self.ensure_paragraph()
        run = self.paragraph.add_run(data)
        run.bold = self.bold > 0
        run.italic = self.italic > 0
        run.underline = self.underline > 0
        if self.strike:
            run.font.strike = True
        if self.size and self.size in FONT_SIZE_MAP:
            run.font.size = Pt(FONT_SIZE_MAP[self.size])
        if self.color:
            rgb = parse_color(self.color)
            if rgb:
                run.font.color.rgb = RGBColor(*rgb)


def html_to_docx_file(title, content_html, target_path):
    if not DOCX_OK:
        return
    doc = Document()
    heading = doc.add_paragraph()
    run = heading.add_run(title or "Sans titre")
    run.bold = True
    run.font.size = Pt(20)
    doc.add_paragraph()
    parser = HtmlToDocx(doc)
    try:
        parser.feed(content_html or "")
    except Exception:
        pass
    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    doc.save(target_path)


# ───────────────────────── Sauvegarde / relecture ─────────────────────────

def save_data(data):
    old_index = load_index() or {"notebooks": []}
    old_pages = {}
    for nb in old_index.get("notebooks", []):
        for sec in nb.get("sections", []):
            for pg in sec.get("pages", []):
                old_pages[pg["id"]] = pg

    seen_dirs = set()
    seen_files = set()
    new_index = {"notebooks": []}

    for nb in data.get("notebooks", []):
        nb_name = sanitize(nb.get("name"), "Bloc-notes")
        new_nb = {"id": nb["id"], "name": nb.get("name", ""), "sections": []}
        for sec in nb.get("sections", []):
            sec_name = sanitize(sec.get("name"), "Categorie")
            folder = os.path.join(DATA_DIR, nb_name, sec_name)
            seen_dirs.add(folder)
            new_sec = {"id": sec["id"], "name": sec.get("name", ""),
                       "color": sec.get("color", "#1eb8cc"), "pages": []}
            for pg in sec.get("pages", []):
                title = pg.get("title") or "Sans titre"
                content = pg.get("content", "")
                base_name = f"{sanitize(title, 'page')}__{short_id(pg['id'])}"
                json_path = os.path.join(folder, base_name + ".json")
                docx_path = os.path.join(folder, base_name + ".docx")

                old = old_pages.get(pg["id"])
                h = content_hash(title, content)
                needs_docx = (old is None) or (old.get("_hash") != h) or (not os.path.exists(docx_path))

                # si le fichier a changé de nom/emplacement, supprimer les anciens
                old_path = old.get("_path") if old else None
                if old_path and old_path != json_path and os.path.exists(old_path):
                    try:
                        os.remove(old_path)
                        old_docx = old_path[:-5] + ".docx"
                        if os.path.exists(old_docx):
                            os.remove(old_docx)
                    except OSError:
                        pass

                write_json(json_path, {
                    "id": pg["id"], "title": title, "content": content,
                    "date": pg.get("date"), "_hash": h,
                })
                if needs_docx:
                    try:
                        html_to_docx_file(title, content, docx_path)
                    except Exception as e:
                        print("Erreur génération docx:", e)

                seen_files.add(json_path)
                seen_files.add(docx_path)
                new_sec["pages"].append({
                    "id": pg["id"], "title": title, "date": pg.get("date"),
                    "_path": json_path, "_hash": h,
                })
            new_nb["sections"].append(new_sec)
        new_index["notebooks"].append(new_nb)

    # nettoyage des anciens fichiers/dossiers qui ne servent plus
    for pg in old_pages.values():
        old_path = pg.get("_path")
        if old_path and old_path not in seen_files and os.path.exists(old_path):
            try:
                os.remove(old_path)
                old_docx = old_path[:-5] + ".docx"
                if os.path.exists(old_docx):
                    os.remove(old_docx)
            except OSError:
                pass
    for root, dirs, files in os.walk(DATA_DIR, topdown=False):
        if root == DATA_DIR:
            continue
        try:
            if not os.listdir(root):
                os.rmdir(root)
        except OSError:
            pass

    write_json(INDEX_PATH, new_index)


def load_data():
    index = load_index()
    if index is None:
        return None
    result = {"notebooks": []}
    for nb in index.get("notebooks", []):
        new_nb = {"id": nb["id"], "name": nb.get("name", ""), "sections": []}
        for sec in nb.get("sections", []):
            new_sec = {"id": sec["id"], "name": sec.get("name", ""),
                       "color": sec.get("color", "#1eb8cc"), "pages": []}
            for pg in sec.get("pages", []):
                path = pg.get("_path")
                content = ""
                if path and os.path.exists(path):
                    try:
                        with open(path, "r", encoding="utf-8") as f:
                            saved = json.load(f)
                            content = saved.get("content", "")
                    except Exception:
                        content = ""
                new_sec["pages"].append({
                    "id": pg["id"], "title": pg.get("title", ""),
                    "content": content, "date": pg.get("date"),
                })
            new_nb["sections"].append(new_sec)
        result["notebooks"].append(new_nb)
    return result


# ───────────────────────── Serveur HTTP ─────────────────────────

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/load":
            try:
                data = load_data()
                self._send_json({"ok": True, "data": data})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, 500)
            return
        super().do_GET()

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/api/save":
            try:
                length = int(self.headers.get("Content-Length", 0))
                raw = self.rfile.read(length)
                data = json.loads(raw.decode("utf-8"))
                save_data(data)
                self._send_json({"ok": True})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, 500)
            return
        self.send_response(404)
        self.end_headers()

    def log_message(self, fmt, *args):
        print("[serveur]", fmt % args)


def main():
    if not DOCX_OK:
        print("⚠️  python-docx n'est pas installé. Les notes seront quand même sauvegardées")
        print("   (en .json), mais les fichiers .docx ne seront pas générés.")
        print("   Pour les avoir : pip install python-docx\n")
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        print(f"✅ Serveur AlumiaNote lancé sur http://localhost:{PORT}/alumia-note.html")
        print(f"   Données stockées dans : {DATA_DIR}")
        print("   (Ctrl+C pour arrêter)")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nArrêt du serveur.")


if __name__ == "__main__":
    main()
