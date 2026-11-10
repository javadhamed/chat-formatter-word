# Chat Formatter for Word — macOS
# قالب‌بندی چت برای ورد (مک)

> Paste an answer from an AI chat into Word, open the **Add-ins** tab (or **Tools → Macro → FormatChatText**), and it turns into
> real Word formatting: headings, bold/italic, tables, lists, clean dashes, code blocks.

---

## Install

Both **`ChatFormatter-Mac-Installer.command`** (standalone single file) and
**`install.command`** (inside the release ZIP) automatically remove any
previous installation first, then install the latest template.

1. Unzip the download (or download **`ChatFormatter-Mac-Installer.command`** directly).
2. Double-click **`install.command`** (or **`ChatFormatter-Mac-Installer.command`**).
   - If macOS says it lacks execute permission after downloading directly via a browser, run `chmod +x ~/Downloads/ChatFormatter-Mac-Installer.command` once in Terminal (or run `bash ~/Downloads/ChatFormatter-Mac-Installer.command`).
   - If macOS refuses to open it, right-click it and choose **Open**, then confirm.
3. Open Word. The first time, Word asks whether to enable the macros in the
   template — choose **Yes**. After that it loads silently on every launch.

No administrator rights, no registry edits, no background service.

If the prompt never appears: **Word → Settings → Security → Trusted Locations**
(Prior to Word 16: **Preferences → Security → Trusted Locations**) and add

```
~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word
```

## Uninstall

Double-click **`uninstall.command`**. It removes the template and your settings.

---

## Use it

Copy text from ChatGPT, DeepSeek, Gemini, Claude — any AI chat — paste it into
Word, then run **Format Chat Text** from the **Add-ins** tab (or **Tools → Macro → Macros... → FormatChatText**).

| Macro | What it does |
|---|---|
| `FormatChatText` | Formats the whole document |
| `FormatSelection` | Formats only the selection |
| `ShowSettings` | Shows the current configuration |

**Undo is one step.** The whole conversion is a single Word undo record, so
`⌘Z` puts everything back.

---

## Settings

Settings live in `~/.chatformatter/settings.ini`:

```ini
Tables = 1
```

`Tables` controls whether tab-separated and pipe tables are converted when
running `FormatSelection` (`0` or `1`). When running `FormatChatText`, Word
asks whether to convert tables too.

---

## What it converts

| Input | Output |
|---|---|
| `# Heading` … `##### Heading` | Word Heading 1–5 |
| `Text` over `===` or `---` | Heading (Setext) |
| `**bold**`, `*italic*`, `~~strike~~` | Real character formatting |
| `` `code` `` | Monospace, coloured inline code |
| `[text](url)` | Blue underlined text |
| `> quote` | Indented quote with a border |
| `- item`, `* item` | Real bullet list |
| `1. item`, `۱. مورد` (or inline `۱. ... ۲. ...`) | Real numbered list (split vertically) |
| `—`, `——`, `–`, prose `--` | Standard `-` hyphen (outside code & URLs) |
| Fenced ```` ```tsv ```` blocks, tab-separated rows | Real Word tables |
| Pipe tables `\| a \| b \|` | Real Word tables (RTL-aware) |
| Code fences with other languages | Shaded monospace block |

Persian and Arabic text is handled: RTL lists get their indent on the right,
and headings keep the correct direction.

---

## Troubleshooting

**Add-ins button or macro missing.** Macros may be disabled — see Install.
Ensure `ChatFormatter.dotm` is in Word's Startup folder and Word was restarted.

**Tables stay as text.** Choose **Yes** when `FormatChatText` asks whether to
convert tables, or set `Tables = 1` for `FormatSelection`. For TSV tables, make
sure rows are separated by real **Tab** characters (not spaces). Markdown pipe
tables do not have this requirement.

**Word will not load the add-in.** Quit Word completely, run `install.command`
again, then start Word. Check `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/ChatFormatter.dotm`
exists.

**A large document takes a while.** The first run on a very long document is
the slowest; later runs on the same document are fast, and running it twice
never changes anything further.

---

English details are in the main `README.md`. فارسی همین راهنما در ادامه است.

---

## نصب

هر دو فایل **`ChatFormatter-Mac-Installer.command`** (تک‌فایلی مستقل) و **`install.command`** (داخل ZIP) به‌صورت خودکار ابتدا نسخهٔ قبلی را پاک کرده و سپس نسخهٔ جدید را نصب می‌کنند.

۱. فایل زیپ را باز کنید (یا فایل **`ChatFormatter-Mac-Installer.command`** را مستقیم دانلود کنید).
۲. روی **`install.command`** (یا **`ChatFormatter-Mac-Installer.command`**) دوبار کلیک کنید.
   - اگر مرورگر اجازهٔ اجرا (`+x`) نداده بود، یک بار در Terminal بزنید `chmod +x ~/Downloads/ChatFormatter-Mac-Installer.command` یا `bash ~/Downloads/ChatFormatter-Mac-Installer.command`.
   - اگر macOS باز نکرد، راست‌کلیک کنید و **Open** را بزنید.
۳. Word را باز کنید. بار اول از شما پرسیده می‌شود ماکروها فعال شوند؛ **Yes**
   را بزنید. بعد از آن هر بار خودکار بارگذاری می‌شود.

نصب نیاز به دسترسی مدیر (administrator) ندارد و هیچ سرویس پس‌زمینه‌ای
نمی‌سازد.

## استفاده

متن را از ChatGPT، DeepSeek، Gemini یا Claude کپی کنید، در Word جای‌گذاری
کنید و از **تب افزونه‌ها (Add-ins)** گزینهٔ **Format Chat Text** (یا از
منوی **Tools → Macro → FormatChatText**) را اجرا کنید. با یک `⌘Z` همه‌چیز
به حالت قبل برمی‌گردد.

## تنظیمات

فایل `~/.chatformatter/settings.ini`:

```ini
Tables = 1
```

مقدار `Tables` تبدیل جدول‌ها را هنگام اجرای `FormatSelection` کنترل می‌کند.
هنگام اجرای `FormatChatText` مستقیماً از شما دربارهٔ تبدیل جدول‌ها سؤال می‌شود.