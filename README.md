# Chat Formatter for Word

Turn text copied from an AI chat into properly formatted Word content.

Paste a DeepSeek / ChatGPT / Gemini answer into Word, select it, right-click —
and it becomes real headings, real tables, bold text, code blocks and lists.
No button, no dialog, no manual cleanup.

## What it converts

| Input | Output |
|---|---|
| `# Heading` … `##### Heading` | Word Heading 1–5 styles |
| `Text` or `===` underline | Heading 1 / Heading 2 (Setext) |
| `**bold**`, `*italic*`, `~~strike~~` | Real character formatting |
| `` `inline code` `` | Consolas 10pt, dark red |
| ```` ```code``` ```` blocks | Shaded monospace paragraphs |
| `\| a \| b \|` + separator row | A real Word table |
| Tab-separated rows (Excel, Sheets) | A real Word table |
| `- item`, `* item`, `+ item` | Bulleted paragraphs (RTL aware) |
| `1. item` | Numbered paragraphs, **numbers preserved** |
| `> quote` | Indented italic blockquote with a rule |
| `[label](url)` | Underlined blue text |
| `---`, `***`, `___` | Removed |

Header rows are bold, shaded, and repeat across page breaks. Tables stretch to
the page width. Persian and English content are both handled — indentation
flips to the right side automatically for RTL text.

## Install

Download the release ZIP, extract it, and double-click **`install.bat`**.

That copies the template into Word's `STARTUP` folder and registers the folder
as a trusted location so macros run without a security prompt.

Restart Word afterwards.

## Use

1. Copy an answer from the AI chat
2. Paste it into Word
3. Select the text, then **right-click**

Alternatively:

- `Alt+F8` → `FormatChatText` formats the whole document
- `Alt+F8` → `FormatSelection` formats just the selection
- `Alt+F8` → `ToggleAutoFormat` turns right-click auto-format on or off
- `Alt+F8` → `ShowSettings` displays the current configuration

Undo with `Ctrl+Z` — the whole run is a single undo record.

## Requirements

- Windows
- Microsoft Word 2016 or newer
- No Python, no .NET, no internet, no admin rights (unless Word lives in a
  protected path)

## Uninstall

Run **`uninstall.bat`**.

## Build from source

The `.dotm` is committed, so most people never need this. To rebuild it after
editing the VBA:

```powershell
powershell -ExecutionPolicy Bypass -File tools\build.ps1
```

Requires Word installed. The script temporarily enables Word's "Trust access to
the VBA project object model" setting, injects `src/*.bas` and `src/*.cls`,
then restores the original setting.

## Layout

```
install.bat              one-click installer
uninstall.bat            one-click uninstaller
ChatFormatter.dotm       the compiled template (this is what Word loads)
src/ChatFormatter.bas    all the formatting logic
src/ThisDocument.cls     event handlers that make auto-format automatic
tools/build.ps1          compiles src/ into ChatFormatter.dotm
```

## How the automatic part works

The template lives in Word's `STARTUP` folder, so Word loads it as a global
add-in on every launch. `ThisDocument` exposes `WindowBeforeRightClick`, which
calls back into `ChatFormatter.CF_RightClickHandler`. That handler checks a
cheap heuristic — does the selection contain `**`, `##`, backticks, pipes or
tabs across multiple lines? — and only then formats. Right-clicking ordinary
prose does nothing.

Turn it off permanently with `Alt+F8` → `ToggleAutoFormat`.

## Security note

This is a macro-enabled template. Install it from a source you trust. The
installer marks Word's `STARTUP` folder as a trusted location, which means
macros placed there run automatically — that is how every Word add-in works,
but be aware of it if you share a machine.

## License

MIT. See `LICENSE`.