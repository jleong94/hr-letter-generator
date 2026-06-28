# hr-letter-generator

A simple, no-experience-needed tool that generates a professional **resignation letter**
as both **Word (.docx)** and **PDF (.pdf)**, in **English, Bahasa Melayu, or Chinese (Simplified)**.

You answer a few questions, and it:

- writes the letter to an experienced-HR / recruiter standard, in the language you choose,
- auto-calculates your **last working day** from the notice date + notice period (in months),
- saves both files into a `generated_letters/` folder.

All dates use the format **`yyyy-MM-dd`** (e.g. `2026-06-28`).

---

## How to use

### Windows
1. Double-click **`run.cmd`**.
2. Answer the prompts in the window that opens.
3. The finished letter opens in the `generated_letters` folder.

### macOS
1. Double-click **`run.command`**.
   - First time only: if macOS blocks it, right-click → **Open** → **Open**.
   - If double-click does nothing, open Terminal in this folder and run `chmod +x run.command run.sh`, then try again.
2. Answer the prompts in the Terminal window.
3. The finished letter opens in the `generated_letters` folder.

> Terminal users (macOS/Linux) can also run `./run.sh`.

The launcher checks your computer the first time it runs and **installs anything missing
automatically**:
- **Python 3** — via `winget` on Windows or **Homebrew** on macOS. If it can't be installed
  automatically, the window shows you exactly what to do.
- The required Python packages (`python-docx`, `fpdf2`) — installed into a local `.venv`
  folder so nothing else on your computer is touched.

The first run takes a minute while things install; later runs start immediately.

---

## What it asks you

- **Language:** English (default), Bahasa Melayu, or 中文 (简体) — the whole letter is written in this language
- **Your details:** full name, job title, department, email, phone
- **Recipient:** name, title, company / organisation, address
- **Notice:** the date you are giving notice (defaults to today) and the notice period in months
- **Optional:** a key achievement to mention in the letter

It then shows the **computed last working day** so you can accept it or type a different date,
and a short review before generating.

> Tip: type your name, company and achievement in the language you selected — they appear in the
> letter exactly as you type them. (The on-screen prompts themselves stay in English.)

### How the last working day is calculated

`last working day = notice date + notice period (months) − 1 day`

For example, giving notice on `2026-06-01` with a `1` month notice period gives a last working
day of `2026-06-30`. You can always override the computed date during the prompts. Entering `0`
months produces a "with immediate effect" letter.

---

## Output

Files are saved in `generated_letters/`. The name uses a prefix for the chosen language plus your
name, e.g.:

```
Resignation_Letter_Jane_Doe.docx             (English)
Surat_Peletakan_Jawatan_Ahmad_bin_Ali.docx   (Bahasa Melayu)
辞职信_张伟.pdf                                 (中文)
```

(If a file with the same name already exists, a number is added so nothing is overwritten.)

---

## Notes

- The **Word** file is fully editable and supports any characters.
- For **English / Bahasa Melayu**, the PDF uses a standard built-in font; very unusual symbols may
  not appear in the PDF — use the Word file for those.
- For **Chinese**, the PDF embeds a Chinese font found on your computer (e.g. Microsoft YaHei or
  SimSun on Windows, PingFang on macOS) — these ship with virtually all systems. In the rare case
  none is found, the Word (.docx) file is still produced with the full Chinese letter.
- Nothing is sent over the internet except the one-time package install. No paid services are used.

---

## Troubleshooting

- **Windows – "Python has been installed… close this window and run again."** The first install
  needs a fresh window to detect Python. Close the window and double-click `run.cmd` again.
- **Windows – winget not available.** Install Python 3 from <https://www.python.org/downloads/>
  and tick *Add python.exe to PATH*, then run `run.cmd` again.
- **macOS – "cannot be opened because it is from an unidentified developer."** Right-click
  `run.command` → **Open** → **Open** (only needed the first time).
- **macOS – double-click does nothing.** Open Terminal in this folder and run
  `chmod +x run.command run.sh`, then try again.
- **No internet on the first run.** The one-time package install needs internet. After the first
  successful run, the tool works offline.
- **Chinese PDF is missing.** If the computer has no Chinese font, the PDF is skipped but the Word
  (`.docx`) file still contains the complete Chinese letter (open it and *Save As PDF* from Word if
  you need a PDF).

---

## Files in this project

| File | What it is |
|------|------------|
| `run.cmd` | Windows launcher (double-click) |
| `run.command` | macOS launcher (double-click in Finder) |
| `run.sh` | macOS / Linux launcher (terminal) |
| `generate_letter.py` | the program that asks the questions and writes the letter |
| `requirements.txt` | the Python packages installed automatically |
| `.gitignore`, `.gitattributes` | repo housekeeping (ignores `.venv/` & `generated_letters/`; keeps script line endings correct) |
| `generated_letters/` | where finished letters are saved (created on first run; not committed) |

---

## Requirements (handled automatically)

- Python 3.9+
- `python-docx`, `fpdf2` (see `requirements.txt`)
- For Chinese PDFs: a CJK font already present on the system (standard on Windows & macOS)
