#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────
#  Lanceur AlumiaNote (Linux)
#   1. vérifie / installe Python 3 et python-docx
#   2. lance serveur_notes.py
#   3. ouvre alumia-note.html dans le navigateur
#  À placer dans le même dossier que serveur_notes.py et alumia-note.html
#  Ctrl+C dans le terminal = arrête le serveur
# ─────────────────────────────────────────────────────────────

cd "$(dirname "$(readlink -f "$0")")" || exit 1

PORT=8791
URL="http://localhost:${PORT}/alumia-note.html"
SERVER_PID=""

say()  { printf '\033[1;36m▶ %s\033[0m\n' "$*"; }
ok()   { printf '\033[1;32m✓ %s\033[0m\n' "$*"; }
warn() { printf '\033[1;33m⚠ %s\033[0m\n' "$*"; }
err()  { printf '\033[1;31m✗ %s\033[0m\n' "$*" >&2; }

for f in serveur_notes.py alumia-note.html; do
  if [ ! -f "$f" ]; then
    err "Fichier introuvable : $f (le lanceur doit être dans le même dossier)"
    exit 1
  fi
done

# ── Droits d'installation ──
SUDO=""
if [ "$(id -u)" -ne 0 ] && command -v sudo >/dev/null 2>&1; then
  SUDO="sudo"
fi

# pkg_install <nom apt> <nom dnf> <nom pacman> <nom zypper>
pkg_install() {
  if [ "$(id -u)" -ne 0 ] && [ -z "$SUDO" ]; then
    warn "Pas de droits root ni de sudo : installation automatique impossible."
    return 1
  fi
  if command -v apt-get >/dev/null 2>&1; then
    $SUDO apt-get update -qq && $SUDO apt-get install -y "$1"
  elif command -v dnf >/dev/null 2>&1; then
    $SUDO dnf install -y "$2"
  elif command -v pacman >/dev/null 2>&1; then
    $SUDO pacman -S --noconfirm --needed "$3"
  elif command -v zypper >/dev/null 2>&1; then
    $SUDO zypper --non-interactive install "$4"
  else
    warn "Gestionnaire de paquets non reconnu."
    return 1
  fi
}

has_docx() { "$1" -c "import docx" >/dev/null 2>&1; }
port_open() { (exec 3<>"/dev/tcp/127.0.0.1/${PORT}") 2>/dev/null; }

# ── 1. Python 3 ──
if ! command -v python3 >/dev/null 2>&1; then
  say "Python 3 est absent, installation..."
  pkg_install python3 python3 python python3
  if ! command -v python3 >/dev/null 2>&1; then
    err "Impossible d'installer Python 3. Installe-le à la main puis relance."
    exit 1
  fi
fi
ok "Python trouvé : $(python3 --version 2>&1)"
PY="python3"

# ── 2. python-docx (pour générer les fichiers Word) ──
if has_docx python3; then
  ok "python-docx déjà installé"
elif [ -x .venv/bin/python ] && has_docx .venv/bin/python; then
  PY=".venv/bin/python"
  ok "python-docx trouvé dans .venv"
else
  say "python-docx est absent, installation..."
  # d'abord via le gestionnaire de paquets du système
  pkg_install python3-docx python3-docx python-docx python3-python-docx

  if has_docx python3; then
    ok "python-docx installé"
  else
    # sinon : environnement virtuel local (.venv), sans toucher au système
    say "Installation dans un environnement virtuel (.venv)..."
    rm -rf .venv
    if ! python3 -m venv .venv >/dev/null 2>&1; then
      rm -rf .venv
      pkg_install python3-venv python3 python python3
      python3 -m venv .venv >/dev/null 2>&1
    fi
    if [ -x .venv/bin/pip ]; then
      .venv/bin/pip install --quiet python-docx
    fi
    if has_docx .venv/bin/python; then
      PY=".venv/bin/python"
      ok "python-docx installé dans .venv"
    else
      warn "python-docx n'a pas pu être installé : les notes seront sauvegardées en .json, sans fichiers .docx."
    fi
  fi
fi

# ── 3. Serveur ──
if port_open; then
  ok "Le serveur tourne déjà sur le port ${PORT}"
else
  say "Lancement du serveur..."
  # allow_reuse_address : évite l'erreur "Address already in use" si on relance juste après un arrêt
  "$PY" -c "import socketserver, runpy; socketserver.TCPServer.allow_reuse_address = True; runpy.run_path('serveur_notes.py', run_name='__main__')" &
  SERVER_PID=$!

  for _ in $(seq 1 50); do
    port_open && break
    if ! kill -0 "$SERVER_PID" 2>/dev/null; then
      err "Le serveur s'est arrêté au démarrage (voir le message ci-dessus)."
      exit 1
    fi
    sleep 0.2
  done
  if ! port_open; then
    err "Le serveur ne répond pas sur le port ${PORT}."
    kill "$SERVER_PID" 2>/dev/null
    exit 1
  fi
  ok "Serveur prêt"
fi

# ── 4. Navigateur ──
say "Ouverture de ${URL}"
if command -v firefox >/dev/null 2>&1; then
  nohup firefox "$URL" >/dev/null 2>&1 &
elif command -v xdg-open >/dev/null 2>&1; then
  nohup xdg-open "$URL" >/dev/null 2>&1 &
else
  warn "Aucun navigateur détecté. Ouvre à la main : ${URL}"
fi

# ── 5. On garde le serveur en vie jusqu'à Ctrl+C ──
if [ -n "$SERVER_PID" ]; then
  cleanup() {
    echo
    say "Arrêt du serveur..."
    kill "$SERVER_PID" 2>/dev/null
    wait "$SERVER_PID" 2>/dev/null
    exit 0
  }
  trap cleanup INT TERM
  echo "   (Ctrl+C pour arrêter le serveur)"
  wait "$SERVER_PID"
fi
