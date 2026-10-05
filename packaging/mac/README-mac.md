# Chat Formatter for Word — macOS
# قالب‌بندی چت برای ورد (مک)

> Paste an answer from an AI chat into Word, right-click it, and it turns into
> real Word formatting: headings, bold/italic, tables, lists, code blocks.

---

## Install

1. Unzip the download.
2. Double-click **`install.command`**. If macOS refuses to open it, right-click
   it and choose **Open**, then confirm.
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
Word, select it, and **right-click**. That is the whole workflow.

Right-click works on Word for Mac 16.52 and later. If your Word build does not
show it, use the macros instead:

| Macro | What it does |
|---|---|
| `FormatChatText` | Formats the whole document |
| `FormatSelection` | Formats only the selection |
| `ToggleAutoFormat` | Turns the automatic right-click mode on/off |
| `ShowSettings` | Opens the settings file |

**Undo is one step.** The whole conversion is a single Word undo record, so
`⌘Z` puts everything back.

---

## Settings

Settings live in `~/.chatformatter/settings.ini`:

```ini
AutoFormat = 1
Tables = 1
```

`AutoFormat` controls the right-click entry, `Tables` controls whether
tab-separated and fenced tables are converted. Both accept `0` or `1`.
`ShowSettings` opens the file in your default editor.

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
| `1. item` | Real numbered list |
| Fenced ```` ```tsv ```` blocks, tab-separated rows | Real Word tables |
| Pipe tables `\| a \| b \|` | Real Word tables |
| Code fences with other languages | Shaded monospace block |

Persian and Arabic text is handled: RTL lists get their indent on the right,
and headings keep the correct direction.

---

## Troubleshooting

**Right-click entry missing.** Macros may be disabled — see Install. Or your
Word build is older than 16.52; use `FormatChatText`.

**Nothing happens when I right-click.** Automatic mode is off. Run
`ToggleAutoFormat`, or turn it on with `AutoFormat = 1` in the settings file.

**Tables stay as text.** Set `Tables = 1`, and make sure the rows are
separated by real **Tab** characters (copy them straight from the chat — do
not use spaces). Markdown pipe tables do not have this requirement.

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

۱. فایل زیپ را باز کنید.
۲. روی **`install.command`** دوبار کلیک کنید. اگر macOS باز نکرد، راست‌کلیک
   کنید و **Open** را بزنید.
۳. Word را باز کنید. بار اول از شما پرسیده می‌شود ماکروها فعال شوند؛ **Yes**
   را بزنید. بعد از آن هر بار خودکار بارگذاری می‌شود.

نصب نیاز به دسترسی مدیر (administrator) ندارد و هیچ سرویس پس‌زمینه‌ای
نمی‌سازد.

## استفاده

متن را از ChatGPT، DeepSeek، Gemini یا Claude کپی کنید، در Word بچسبانید،
انتخاب کنید و **راست‌کلیک** کنید. تمام کار همین است.

اگر راست‌کلیک در نسخهٔ Word شما کار نکرد، از ماکرو `FormatChatText` استفاده
کنید. با یک `⌘Z` همه‌چیز به حالت قبل برمی‌گردد.

## تنظیمات

فایل `~/.chatformatter/settings.ini`:

```ini
AutoFormat = 1
Tables = 1
```

`AutoFormat` حالت خودکار راست‌کلیک را کنترل می‌کند و `Tables` تبدیل جدول‌ها
را. با ماکرو `ShowSettings` این فایش را باز می‌کنید.