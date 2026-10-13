#!/usr/bin/env bash
# Report Kit installer for Ubuntu/Linux (x86_64). No sudo needed.
#
#   git clone https://github.com/<you>/report-kit ~/report-kit && ~/report-kit/install.sh
#
# Everything goes inside the kit folder (tools/, .venv/, node/) plus fonts in
# ~/.local/share/fonts, the `report` command in ~/.local/bin and the Claude Code
# skill in ~/.claude/skills/report. Safe to re-run: present pieces are skipped.
set -euo pipefail

KIT="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
TOOLS="$KIT/tools"; BIN="$KIT/bin"
mkdir -p "$TOOLS" "$BIN" "$HOME/.local/bin"

TYPST=0.15.1; PANDOC=3.12.1; QUARTO=1.10.19; D2=v0.9.0; LT=6.6; NODE=v20.20.2

# grep -q exits on the first match, which SIGPIPEs fc-list and, under pipefail,
# turns a found font into a failed check. Read the whole list instead.
has_font() { fc-list 2>/dev/null | grep -i "$1" >/dev/null; }

step() { printf '\n\033[1m▸ %s\033[0m\n' "$*"; }
ok()   { printf '  \033[32m✓\033[0m %s\n' "$*"; }
warn() { printf '  \033[33m!\033[0m %s\n' "$*"; }
die()  { printf '  \033[31m✗ %s\033[0m\n' "$*"; exit 1; }

fetch_tar() {  # url, target dir name inside tools/
  local url="$1" dir="$2"
  if [ -e "$TOOLS/$dir" ]; then ok "$dir (already there)"; return; fi
  printf '  downloading %s ...\n' "${url##*/}"
  local try
  for try in 1 2 3; do
    rm -f "$TOOLS/dl.part"
    curl -fsSL --retry 3 "$url" -o "$TOOLS/dl.part" && break
    warn "download attempt $try failed; retrying"; sleep 3
  done
  [ -s "$TOOLS/dl.part" ] || die "could not download $url"
  case "$url" in
    *.zip)    mv "$TOOLS/dl.part" "$TOOLS/dl.zip" && unzip -q "$TOOLS/dl.zip" -d "$TOOLS" && rm "$TOOLS/dl.zip" ;;
    *.tar.xz) tar -xJf "$TOOLS/dl.part" -C "$TOOLS" && rm "$TOOLS/dl.part" ;;
    *)        tar -xzf "$TOOLS/dl.part" -C "$TOOLS" && rm "$TOOLS/dl.part" ;;
  esac
  [ -e "$TOOLS/$dir" ] || die "download of $dir failed ($url)"
  ok "$dir"
}

step "Checking basics"
[ "$(uname -m)" = "x86_64" ] || die "this installer supports x86_64 Linux only"
for c in curl tar unzip python3; do command -v "$c" >/dev/null || die "$c is missing (ask the lab admin for: $c)"; done
ok "curl, tar, unzip, python3 $(python3 -c 'import sys;print(".".join(map(str,sys.version_info[:2])))')"
free_gb=$(df -Pk "$KIT" | awk 'NR==2 {print int($4/1048576)}')
[ "${free_gb:-0}" -ge 3 ] || die "only ${free_gb} GB free where the kit lives; it needs about 3 GB (lab PCs often have a home-folder quota: ask for more, or clone the kit to a bigger disk and re-run)"
ok "${free_gb} GB free"

step "Document tools (Typst, Pandoc, Quarto, D2)"
fetch_tar "https://github.com/typst/typst/releases/download/v$TYPST/typst-x86_64-unknown-linux-musl.tar.xz" typst-x86_64-unknown-linux-musl
fetch_tar "https://github.com/jgm/pandoc/releases/download/$PANDOC/pandoc-$PANDOC-linux-amd64.tar.gz" "pandoc-$PANDOC"
fetch_tar "https://github.com/quarto-dev/quarto-cli/releases/download/v$QUARTO/quarto-$QUARTO-linux-amd64.tar.gz" "quarto-$QUARTO"
fetch_tar "https://github.com/terrastruct/d2/releases/download/$D2/d2-$D2-linux-amd64.tar.gz" "d2-$D2"
ln -sf "../tools/typst-x86_64-unknown-linux-musl/typst" "$BIN/typst"
ln -sf "../tools/pandoc-$PANDOC/bin/pandoc" "$BIN/pandoc"
ln -sf "../tools/quarto-$QUARTO/bin/quarto" "$BIN/quarto"
ln -sf "../tools/d2-$D2/bin/d2" "$BIN/d2"

step "Node.js and Mermaid (diagrams)"
node_major=$(node -v 2>/dev/null | sed 's/v\([0-9]*\).*/\1/' || echo 0)
if [ "${node_major:-0}" -lt 18 ]; then
  fetch_tar "https://nodejs.org/dist/$NODE/node-$NODE-linux-x64.tar.xz" "node-$NODE-linux-x64"
  export PATH="$TOOLS/node-$NODE-linux-x64/bin:$PATH"
  ln -sf "../tools/node-$NODE-linux-x64/bin/node" "$BIN/node"
else
  ok "system node $(node -v)"
fi
retry() {  # retry a flaky network step 3 times, showing the last error
  local n=1
  until "$@" >"$TOOLS/last.log" 2>&1; do
    [ $n -ge 3 ] && { tail -15 "$TOOLS/last.log"; return 1; }
    warn "attempt $n failed (network?); retrying"; n=$((n + 1)); sleep 3
  done
}
if [ ! -x "$KIT/node/node_modules/.bin/mmdc" ] || [ ! -d "$HOME/.cache/puppeteer" ]; then
  (cd "$KIT/node" && retry npm ci --no-audit --no-fund) || die "npm install of mermaid-cli failed (error above)"
fi
ln -sf "../node/node_modules/.bin/mmdc" "$BIN/mmdc"
ok "mermaid-cli"

step "Python environment"
if [ ! -x "$KIT/.venv/bin/python" ]; then
  if ! python3 -m venv "$KIT/.venv" 2>/dev/null; then
    rm -rf "$KIT/.venv"
    warn "python3-venv is not installed; using uv instead (no sudo)"
    command -v uv >/dev/null || curl -LsSf https://astral.sh/uv/install.sh | sh >/dev/null
    "$HOME/.local/bin/uv" venv -q "$KIT/.venv" && "$HOME/.local/bin/uv" pip install -q --python "$KIT/.venv/bin/python" pip
  fi
fi
if ! "$KIT/.venv/bin/pip" install -q -r "$KIT/requirements.txt" >"$TOOLS/pip.log" 2>&1; then
  if grep -qiE "no matching distribution|resolutionimpossible|requires-python|could not find a version" "$TOOLS/pip.log"; then
    warn "pinned versions don't fit this Python; installing the newest compatible ones"
    sed 's/==.*//' "$KIT/requirements.txt" | "$KIT/.venv/bin/pip" install -q -r /dev/stdin \
      || die "Python packages failed to install (see above)"
  else
    tail -5 "$TOOLS/pip.log"; die "Python packages failed to install (error above)"
  fi
fi
"$KIT/.venv/bin/python" -c "import matplotlib, pymupdf, docx, yaml, bibtexparser" || die "Python packages failed to install"
ok "Python packages"

step "Grammar checker (LanguageTool) and Java"
fetch_tar "https://languagetool.org/download/LanguageTool-$LT.zip" "LanguageTool-$LT"
java_major=$(java -version 2>&1 | sed -n 's/.*version "\([0-9]*\).*/\1/p' | head -1)
if [ "${java_major:-0}" -lt 17 ]; then
  if [ ! -x "$TOOLS/jre/bin/java" ]; then
    printf '  downloading a private Java 21 runtime ...\n'
    mkdir -p "$TOOLS/jre"
    curl -fsSL "https://api.adoptium.net/v3/binary/latest/21/ga/linux/x64/jre/hotspot/normal/eclipse" \
      | tar -xz -C "$TOOLS/jre" --strip-components=1
  fi
  ln -sf "../tools/jre/bin/java" "$BIN/java"
  ok "private Java 21"
else
  ok "system Java $java_major"
fi

step "Fonts (Times New Roman, Arial; Microsoft's free core fonts)"
FONTDIR="$HOME/.local/share/fonts/msttcore"
if has_font "Times New Roman" && has_font "Arial"; then
  ok "already installed"
else
  mkdir -p "$FONTDIR" "$TOOLS/cab"
  if ! command -v cabextract >/dev/null; then
    (cd "$TOOLS/cab" && { apt-get download cabextract >/dev/null 2>&1; apt-get download libmspack0t64 >/dev/null 2>&1 || apt-get download libmspack0 >/dev/null 2>&1; } ; for d in *.deb; do dpkg -x "$d" root; done) || true
    CAB="$TOOLS/cab/root/usr/bin/cabextract"
    export LD_LIBRARY_PATH="$TOOLS/cab/root/usr/lib/x86_64-linux-gnu:${LD_LIBRARY_PATH:-}"
  else
    CAB=cabextract
  fi
  if [ -x "$CAB" ] || command -v "$CAB" >/dev/null; then
    for f in times32 arial32; do
      for url in "https://downloads.sourceforge.net/corefonts/$f.exe" \
                 "https://downloads.sourceforge.net/project/corefonts/the%20fonts/final/$f.exe"; do
        curl -fsSL --retry 4 --retry-all-errors --retry-delay 2 "$url" -o "$TOOLS/cab/$f.exe" && break
      done
      [ -s "$TOOLS/cab/$f.exe" ] && "$CAB" -q -F '*.ttf' -L -d "$FONTDIR" "$TOOLS/cab/$f.exe" || true
    done
    fc-cache -f >/dev/null 2>&1 || true
  fi
  if has_font "Times New Roman"; then ok "Times New Roman, Arial"
  else warn "could not install Times New Roman; reports use Liberation Serif (same letter widths, same page breaks)"; fi
fi

step "Commands and Claude Code skill"
[ "$KIT" = "$HOME/report-kit" ] || { [ -e "$HOME/report-kit" ] || ln -s "$KIT" "$HOME/report-kit"; }
ln -sf "$KIT/bin/report" "$HOME/.local/bin/report"
mkdir -p "$HOME/.claude/skills"
if [ -e "$HOME/.claude/skills/report" ] && [ ! -L "$HOME/.claude/skills/report" ]; then
  mv "$HOME/.claude/skills/report" "$HOME/.claude/skills/report.bak.$(date +%s)"
fi
ln -sfn "$KIT/claude-skill/report" "$HOME/.claude/skills/report"
ok "report → ~/.local/bin/report;  skill → ~/.claude/skills/report"
case ":$PATH:" in *":$HOME/.local/bin:"*) ;; *) warn "~/.local/bin is not on PATH yet: open a new terminal (or log out and in)";; esac

step "Optional pieces"
if command -v soffice >/dev/null && python3 -c "import uno" 2>/dev/null; then
  ok "LibreOffice: Word files get a filled-in TOC and a layout check"
else
  warn "LibreOffice (with python3-uno) not found: Word fills the TOC when it opens the file; everything else works"
fi
if command -v claude >/dev/null; then ok "Claude Code $(claude --version 2>/dev/null | head -1)"
else warn "Claude Code not installed. To use \"make the report for this project\":  curl -fsSL https://claude.ai/install.sh | bash   then run: claude"; fi

step "Self-test: building the demo report"
DEMO="$(mktemp -d)"; cp -r "$KIT/examples/demo" "$DEMO/"
if (cd "$DEMO/demo/report" && PATH="$BIN:$PATH" "$BIN/report" build >"$DEMO/log" 2>&1) && [ -s "$DEMO/demo/report/out/report.pdf" ]; then
  ok "demo report built ($(grep -o '[0-9]* pages' "$DEMO/log" | head -1)) — the kit works"
  rm -rf "$DEMO"
else
  tail -25 "$DEMO/log"; die "self-test failed (full log: $DEMO/log)"
fi

printf '\n\033[1mDone.\033[0m In any project: open Claude Code and say "make the report for this project",\nor run:  report init  then  report build\n'
