# Chat Formatter for Word / قالب‌بندی چت برای ورد

> **English** | [فارسی](#فارسی)

---

## English

### The Problem This Solves

**If you've ever copied a long answer from ChatGPT, DeepSeek, Gemini, or any AI chat into Word, you know the pain:**

- Headings stay as plain text with `#` symbols
- Bold/italic/inline code markers (`**`, `*`, `` ` ``) remain visible
- Markdown tables (`| a | b |`) don't become real Word tables
- Code blocks, blockquotes, bullet/numbered lists — all need manual cleanup
- You spend **hours** fixing formatting instead of using the content

**This repository exists to solve that in seconds.**

Paste an AI answer into Word, select it, **right-click** — and it instantly becomes:
- Real Word headings (styles Heading 1–5)
- Real character formatting (bold, italic, strikethrough, inline code)
- Real Word tables (Markdown pipe tables AND tab-separated from Excel/Sheets)
- Real bullet/numbered lists (with RTL support for Persian/Arabic)
- Proper code blocks with monospace font and shading
- Blockquotes with indentation and border
- Links as underlined blue text

**Undo with `Ctrl+Z`** — the entire conversion is a single undo record.

---

### What It Converts

| Input | Output |
|-------|--------|
| `# Heading` … `##### Heading` | Word Heading 1–5 styles |
| `Text` / `===` underline | Heading 1 / Heading 2 (Setext) |
| `**bold**`, `*italic*`, `~~strike~~` | Real character formatting |
| `` `inline code` `` | Consolas 10pt, dark red |
| ```` ```code``` ```` blocks | Shaded monospace paragraphs |
| `\| a \| b \|` + separator row | Real Word table |
| Tab-separated rows (Excel/Sheets) | Real Word table |
| `- item`, `* item`, `+ item` | Bulleted paragraphs (RTL aware) |
| `1. item` | Numbered paragraphs, **numbers preserved** |
| `> quote` | Indented italic blockquote with right border |
| `[label](url)` | Underlined blue text |
| `---`, `***`, `___` | Removed |

Header rows are bold, shaded, and repeat across page breaks. Tables stretch to page width. Persian and English content are both handled — indentation flips to the right side automatically for RTL text.

---

### Install

**Windows**

1. Download the latest release ZIP from the [Releases page](../../releases)
2. Extract it anywhere
3. **Double-click `install.bat`**

**macOS**

1. Download and unzip the release
2. **Double-click `install.command`** — if macOS refuses to open it, right-click it and choose **Open**
3. Open Word and answer **Yes** the first time it asks to enable the macros

Both copy `ChatFormatter.dotm` into Word's `STARTUP` folder, so Word loads it as a global add-in on every launch. The Windows installer also registers that folder as a trusted location so macros run without security prompts.

**Restart Word** afterwards. Settings live in `~/.chatformatter/settings.ini` on both platforms.

---

### Use

**Automatic (recommended):**
1. Copy an answer from the AI chat
2. Paste it into Word
3. **Select the text, then right-click** — it formats instantly

**Manual (via `Alt+F8`):**
| Macro | Action |
|-------|--------|
| `FormatChatText` | Format entire document (asks: include tables?) |
| `FormatSelection` | Format only the current selection |
| `ToggleAutoFormat` | Enable/disable right-click auto-format |
| `ShowSettings` | Display current configuration |

---

### Requirements

- Windows or macOS
- Microsoft Word 2016 or newer (on macOS the right-click entry needs 16.52+)
- **No Python, no .NET, no internet, no admin rights** (unless Word is in a protected path)

---

### Uninstall

Run **`uninstall.bat`** (Windows) or **`uninstall.command`** (macOS). It removes the template and your `settings.ini`; the trusted-location entry on Windows is left in place because it is harmless.

---

### Build from Source

The `.dotm` is committed so most users never need this. To rebuild after editing the VBA:

```
python3 tools/package_win.py     # dist/win/ and dist/ChatFormatter-Win-<version>.zip
python3 tools/package_mac.py     # dist/mac/ and dist/ChatFormatter-Mac-<version>.zip
```

Needs Python 3 and `olefile` (`pip install olefile`). The build writes the VBA project into a copy of `tools/seed.dotm` using a pure-Python MS-OVBA writer, so **it needs neither Word nor PowerShell** and runs identically on every platform. The template is platform neutral: the same `.dotm` ships to Windows and macOS, only the installer differs.

Check the result before shipping:

```
python3 tests/lint_vba.py src/ChatFormatter.bas src/ThisDocument.cls
python3 tests/test_cases.py
python3 tests/test_sample.py
python3 tests/test_word_output.py <original.docx> <formatted.docx>
```

`test_word_output.py` compares a document the add-in already formatted against the reference model; it needs a real Word run to produce the second file.

---

### Repository Layout

```
install.bat              one-click installer (Windows)
uninstall.bat            one-click uninstaller (Windows)
packaging/mac/           install.command, uninstall.command, README-mac.md
ChatFormatter.dotm       compiled template (what Word loads)
src/ChatFormatter.bas    all formatting logic, one linear pass
src/ThisDocument.cls     event handlers for auto-format
tools/vbabuild.py        MS-OVBA compressor and CFB writer
tools/build_dotm.py      injects src/ into tools/seed.dotm
tools/package_win.py     builds the Windows release ZIP
tools/package_mac.py     builds the macOS release ZIP
tests/                   reference model and test suites
```

---

### How Automatic Mode Works

The template lives in Word's `STARTUP` folder, so Word loads it as a global add-in on every launch. `ThisDocument` exposes `WindowBeforeRightClick`, which calls `ChatFormatter.CF_RightClickHandler`. That handler runs a cheap heuristic — does the selection contain `**`, `##`, backticks, pipes, or tabs across multiple lines? — and only formats when it looks like AI chat markdown. Right-clicking ordinary prose does nothing.

Turn it off permanently with `Alt+F8` → `ToggleAutoFormat`.

---

### Security Note

This is a macro-enabled template (`.dotm`). Install it from a source you trust. The installer marks Word's `STARTUP` folder as a trusted location — that's how every Word add-in works, but be aware if you share a machine.

---

### License

MIT. See `LICENSE`.

---

## فارسی

### مشکلِ که این پروژه حل می‌کند

**اگر تا به حال پاسخ طولانی‌ای را از ChatGPT، DeepSeek، Gemini یا هر چت‌بات هوش مصنوعی دیگر در ورد کپی کرده باشید، دردش را می‌دانید:**

- تیترها به صورت متن ساده با نماد `#` باقی می‌مانند
- نشانه‌های پرته/خمیده/کد‌درون‌خطی (`**`، `*`، `` ` ``) همچنان قابل دیدن هستند
- جداول مارک‌داون (`| a | b |`) تبدیل به جداول واقعی ورد نمی‌شوند
- بلوک‌های کد، نقل‌قول‌ها، لیست‌های پرتی/شماره‌دار — همه نیاز به اصلاح دستی دارند
- شما **ساعات** وقت‌تان را صرف ویرایش قالب‌بندی می‌کنید به جای استفاده از محتوا

**این مخزن برای حل همین مشکل در عرض چند ثانیه ساخته شده است.**

یک پاسخ هوش مصنوعی را در ورد پیست کنید، انتخاب کنید، **راست‌کلیک** کنید — و بلافاصله تبدیل می‌شود به:
- تیترهای واقعی ورد (استایل‌های Heading ۱–۵)
- قالب‌بندی واقعی کاراکتر (پرته، خمیده، خط‌خورده، کد‌درون‌خطی)
- جداول واقعی ورد (جداول مارک‌داون Pipe **و** جداول Tab-separated از اکسل/شیتز)
- لیست‌های پرتی/شماره‌دار واقعی (با پشتیبانی RTL برای فارسی/عربی)
- بلوک‌های کد با قلم مونوسپیس و سایه‌دهی
- نقل‌قول‌ها با تورفتگی و کادر
- لینک‌ها به صورت متن زیرخط‌دار آبی

**واچنین با `Ctrl+Z`** — کل تبدیل یک رکورد Undo واحد است.

---

### مواردی که تبدیل می‌شوند

| ورودی | خروجی |
|--------|-------|
| `# Heading` … `##### Heading` | استایل‌های ورد Heading ۱–۵ |
| متن ساده / زیرخط `===` | Heading ۱ / Heading ۲ (Setext) |
| `**پرته**`، `*خمیده*`، `~~خط‌خورده~~` | قالب‌بندی واقعی کاراکتر |
| `` `کد‌درون‌خطی` `` | Consolas ۱۰نقطه، قرمز تیره |
| ```` ```بلوک‌کد``` ```` | پاراگراف‌های مونوسپیس سایه‌دار |
| `\| الف \| ب \|` + ردیف جداکننده | جدول واقعی ورد |
| ردیف‌های جدا شده با Tab (اکسل/شیتز) | جدول واقعی ورد |
| `- آیتم`، `* آیتم`، `+ آیتم` | پاراگراف‌های پرتی (آگاه به RTL) |
| `۱. آیتم` | پاراگراف‌های شماره‌دار، **شماره‌ها حفظ می‌شوند** |
| `> نقل‌قول` | بلوک نقل‌قول خمیده تورفتگی‌دار با کادر راست |
| `[برچسب](آدرس)` | متن زیرخط‌دار آبی |
| `---`، `***`، `___` | حذف می‌شوند |

ردیف‌های سرتیتر پرته، سایه‌دار و در فصفح تکرار می‌شوند. جداول عرض صفحه را می‌پوشانند. محتوای فارسی و انگلیسی هر دو پشتیبانی می‌شوند — تورفتگی برای متن‌های RTL به‌صورت خودکار به راست می‌رود.

---

### نصب

**ویندوز**

۱. آخرین ZIP منتشر شده را از [صفحه Releases](../../releases) دانلود کنید
۲. آن را در محل دلخواه استخراج کنید
۳. **روی `install.bat` دابل‌کلیک کنید**

**مک**

۱. فایل زیپ را دانلود و باز کنید
۲. **روی `install.command` دابل‌کلیک کنید** — اگر مک باز نکرد، راست‌کلیک کنید و **Open** را بزنید
۳. ورد را باز کنید و بار اول، وقتی پرسید ماکروها فعال شوند، **Yes** را بزنید

در هر دو سیستم `ChatFormatter.dotm` در پوشه `STARTUP` ورد کپی می‌شود، پس ورد آن را در هر اجرا به‌عنوان یک add-in سراسری بارگذاری می‌کند. نصب‌کنندهٔ ویندوز آن پوشه را هم به‌عنوان Trusted Location ثبت می‌کند تا ماکروها بدون پرامپت امنیتی اجرا شوند.

**ورد را ری‌استارت کنید.** تنظیمات در هر دو سیستم در `~/.chatformatter/settings.ini` است.

---

### استفاده

**خودکار (پیشنهادی):**
۱. یک پاسخ از چت هوش مصنوعی کپی کنید
۲. در ورد پیست کنید
۳. **متن را انتخاب کرده و راست‌کلیک کنید** — بلافاصله قالب‌بندی می‌شود

**دستی (از طریق `Alt+F8`):**

| ماکرو | عملکرد |
|--------|--------|
| `FormatChatText` | قالب‌بندی کل سند (پرسش: جداول شامل شوند؟) |
| `FormatSelection` | قالب‌بندی فقط انتخاب فعلی |
| `ToggleAutoFormat` | فعال/غیرفعال کردن قالب‌بندی خودکار راست‌کلیک |
| `ShowSettings` | نمایش تنظیمات فعلی |

---

### پیش‌نیازها

- ویندوز یا مک
- مایکروسافت ورد ۲۰۱۶ یا جدیدتر (روی مک، راست‌کلیک نیاز به نسخهٔ ۱۶.۵۲ به بعد دارد)
- **بدون پایتون، بدون دات‌نت، بدون اینترنت، بدون حق ادمین** (مگر اینکه ورد در مسیر محافظت‌شده باشد)

---

### حذف نصب

**`uninstall.bat`** (ویندوز) یا **`uninstall.command`** (مک) را اجرا کنید. این‌ها قالب و فایل `settings.ini` شما را حذف می‌کنند؛ رکورد Trusted Location روی ویندوز باقی می‌ماند چون بی‌اثر است.

---

### ساخت از کد منبع

فایل `.dotm` در مخزن ذخیره شده است بنابراین اکثر کاربران نیازی به ساخت ندارند. برای بازسازی پس از ویرایش VBA:

```
python3 tools/package_win.py     # dist/win/ و dist/ChatFormatter-Win-<version>.zip
python3 tools/package_mac.py     # dist/mac/ و dist/ChatFormatter-Mac-<version>.zip
```

به پایتون ۳ و `olefile` (`pip install olefile`) نیاز دارد. بیلد، پروژهٔ VBA را با یک نویسندهٔ خالص پایتونی MS-OVBA داخل کپی‌ای از `tools/seed.dotm` می‌نویسد؛ بنابراین **نه به ورد نیاز دارد و نه به PowerShell** و روی همهٔ سیستم‌ها یکسان کار می‌کند. خودِ قالب مستقل از پلتفرم است: همان `.dotm` برای ویندوز و مک ارسال می‌شود و فقط نصب‌کننده فرق می‌کند.

قبل از انتشار بررسی کنید:

```
python3 tests/lint_vba.py src/ChatFormatter.bas src/ThisDocument.cls
python3 tests/test_cases.py
python3 tests/test_sample.py
python3 tests/test_word_output.py <original.docx> <formatted.docx>
```

`test_word_output.py` سندی را که افزونه فرمت کرده با مدل مرجع مقایسه می‌کند؛ برای ساختن فایل دوم یک اجرای واقعی ورد لازم است.

---

### چیدمان مخزن

```
install.bat              نصب‌کننده یک‌کلیک (ویندوز)
uninstall.bat            حذف‌کننده یک‌کلیک (ویندوز)
packaging/mac/           install.command، uninstall.command، README-mac.md
ChatFormatter.dotm       قالب کامپایل‌شده (چه چیزی ورد بارگذاری می‌کند)
src/ChatFormatter.bas    کل منطق قالب‌بندی، در یک پاس خطی
src/ThisDocument.cls     هندلرهای رویداد برای قالب‌بندی خودکار
tools/vbabuild.py        فشرده‌ساز MS-OVBA و نویسندهٔ CFB
tools/build_dotm.py      تزریق src/ به tools/seed.dotm
tools/package_win.py     ساخت ZIP نسخهٔ ویندوز
tools/package_mac.py     ساخت ZIP نسخهٔ مک
tests/                   مدل مرجع و مجموعه تست‌ها
```

---

### نحوه کارکرد حالت خودکار

قالب در پوشه `STARTUP` ورد قرار دارد، بنابراین ورد آن را در هر اجرا به عنوان یک add-in سراسری بارگذاری می‌کند. `ThisDocument` رویداد `WindowBeforeRightClick` را افشا می‌کند که `ChatFormatter.CF_RightClickHandler` را صدا می‌زند. این هندلر یک اکتشافی ارزان انجام می‌دهد — آیا انتخاب حاوی `**`، `##`، بک‌تیک، Pipe یا Tab در چند خط است؟ — و فقط وقتی شبیه مارک‌داون چت هوش مصنوعی باشد قالب‌بندی می‌کند. راست‌کلیک روی متن عادی هیچ کاری انجام نمی‌دهد.

برای غیرفعالسازی دائمی: `Alt+F8` → `ToggleAutoFormat`.

---

### نکته امنیتی

این یک قالب ماکرو‌دار (`.dotm`) است. از منبعی که به آن اعتماد دارید نصب کنید. نصب‌کننده پوشه `STARTUP` ورد را به عنوان Trusted Location علامت‌گذاری می‌کند — این نحوه کار هر add-in وردی است، اما اگر دستگاه را به اشتراک می‌گذارید از آن آگاه باشید.

---

### مجوز

MIT. فایل `LICENSE` را ببینید.

---

### مشارکت

پیشنهادات، گزارش باگ و Pull Requestها خوش‌آمدند. لطفاً قبل از ارسال PR، تست‌های `tests/` را اجرا کنید و قالب را با `python3 tools/package_win.py` و `python3 tools/package_mac.py` بسازید.

---

### تحسین

این پروژه برای همه کسانی ساخته شده که می‌خواهند خروجی هوش مصنوعی را به سرعت در اسناد حرفه‌ای تبدیل کنند — بدون دردسر قالب‌بندی دستی.