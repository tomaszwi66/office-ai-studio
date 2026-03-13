# Office AI Studio

> Local AI-powered office automation for people who work with files, documents and data every day.  
> No cloud. No subscription. Runs entirely on your machine via [Ollama](https://ollama.com).

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![Ollama](https://img.shields.io/badge/Ollama-required-orange)](https://ollama.com)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)]()
[![License](https://img.shields.io/badge/License-MIT-green)]()

---

## Overview

Office workers spend hours on repetitive tasks: renaming files, extracting data from documents, summarising reports, turning meeting notes into action items. Office AI Studio automates all of this - with or without AI - from a single desktop app built on Python and Ollama.

Everything runs locally. Your files never leave your machine.

---

## Features

**⬇ Smart Drop Zone** - drag any file onto the app and instantly choose what to do: summarize, extract data, translate, rename, copy, hash, and more. Works on multiple files at once.

**⚡ Pipeline Builder** - build multi-step AI workflows visually. Each step passes its output to the next. Drag files directly onto step cards. Save and reload named pipelines.

**» Scripts** - write, save and run Python automation scripts. Five starter scripts included. AI can generate a new script from a description or explain any existing one.

**⏲ Auto Tasks** - create automation rules based on file patterns. Drop files onto a task card to run it instantly. Tracks run count and last-run time.

**⊞ Data Tools** - open, preview, merge, clean and export CSV files without Excel. AI can analyse the data and describe patterns or quality issues.

**◎ Meeting Notes** - paste raw notes and get three outputs: action items with owners, executive summary, and a ready-to-send follow-up email draft.

**≡ File Manager** - browse the filesystem with a text preview pane and right-click context menu.

**◉ Chat** - multi-turn conversation with any local Ollama model.

**📝 Notepad AI** - full-screen editor with an AI assistant panel. Insert or replace text with AI output.

**>_ Terminal AI** - real shell with `?? <question>` to ask AI and an auto-explain mode for command output.

---

## Quickstart

```bash
# 1. Install Ollama
#    https://ollama.com/download

# 2. Pull a model
ollama pull llama3.2:3b

# 3. Start Ollama server
ollama serve

# 4. Install Python dependencies
pip install -r requirements.txt

# 5. Run
python office_ai_studio.py
```

---

## Requirements

| Package | Required | Purpose |
|---------|----------|---------|
| `requests` | ✅ Yes | Ollama API communication |
| `tkinterdnd2` | Recommended | Drag & drop file support |
| `chardet` | Recommended | Auto-detect file encoding |
| `python-docx` | Optional | Read .docx files |
| `openpyxl` | Optional | Read .xlsx files |
| `pandas` | Optional | Advanced data analysis |

`tkinter` is bundled with Python on Windows. On Linux:
```bash
sudo apt install python3-tk
```

---

## Model Compatibility

The app tries `/api/chat` first and automatically falls back to `/api/generate` for base models without a chat template.

| Model | Status |
|-------|--------|
| `llama3.2:3b` | ✅ |
| `llama3.2:1b` | ✅ |
| `SpeakLeash/bielik-*-instruct` | ✅ |
| `mistral:*` | ✅ |
| `qwen2.5:*` | ✅ |
| Base models (no chat template) | ✅ generate fallback |

---

## Data Storage

All data is stored locally in `~/.office_ai_studio/`:

| File | Contents |
|------|----------|
| `pipelines.json` | Saved pipelines |
| `history.json` | AI run history (last 500 entries) |
| `scripts.json` | Saved automation scripts |
| `tasks.json` | Auto task rules |

---

## Keyboard Shortcuts

| Shortcut | Context | Action |
|----------|---------|--------|
| `Ctrl+Enter` | Chat, Notepad AI | Send message |
| `F5` | Scripts | Run script |
| `↑` / `↓` | Terminal AI | Navigate command history |
| `?? ` prefix | Terminal AI | Ask AI a question |
| Double-click | File Manager | Open file / enter folder |

---

## Author

**Tomasz Wietrzykowski**  
[GitHub](https://github.com/tomaszwi66) · [X / Twitter](https://x.com/twf24) · [WebSim](https://websim.com/@TomaszW)

---

## License

MIT - free to use, modify and distribute.
