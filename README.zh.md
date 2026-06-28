# hr-letter-generator

[English](README.md) | [Bahasa Melayu](README.ms.md) | **中文**

一个简单、无需任何经验的小工具，可生成专业的**辞职信**，同时输出 **Word (.docx)** 和 **PDF (.pdf)**
两种格式，支持**英文、马来文或简体中文**。

您只需回答几个问题，它便会：

- 以资深 HR／招聘专员的标准撰写信件，并使用您所选择的语言，
- 根据通知日期 + 通知期（以月计）自动计算您的**最后工作日**，
- 将两个文件一并保存到 `generated_letters/` 文件夹中。

所有日期均采用 **`yyyy-MM-dd`** 格式（例如 `2026-06-28`）。

---

## 使用方法

### Windows
1. 双击 **`run.cmd`**。
2. 在打开的窗口中回答提示问题。
3. 生成的信件会在 `generated_letters` 文件夹中打开。

### macOS
1. 双击 **`run.command`**。
   - 仅首次：如果 macOS 阻止运行，请右键点击 → **Open** → **Open**。
   - 如果双击没有反应，请在此文件夹中打开 Terminal 并运行 `chmod +x run.command run.sh`，然后再试一次。
2. 在 Terminal 窗口中回答问题。
3. 生成的信件会在 `generated_letters` 文件夹中打开。

> Terminal 用户（macOS/Linux）也可以运行 `./run.sh`。

首次运行时，启动器会检查您的电脑并**自动安装任何缺失的组件**：
- **Python 3** —— 在 Windows 上通过 `winget`，在 macOS 上通过 **Homebrew** 安装。如果无法自动
  安装，窗口会明确告诉您该如何操作。
- 所需的 Python 软件包（`python-docx`、`fpdf2`）—— 安装到本地的 `.venv` 文件夹中，不会影响电脑上
  的其他任何东西。

首次运行需要约一分钟来完成安装；之后再运行会立即启动。

---

## 程序会询问的内容

- **语言：** 英文（默认）、马来文或中文（简体）—— 整封信都会以此语言撰写
- **您的信息：** 全名、职位、部门、电子邮箱、电话
- **收件人：** 姓名、职位、公司／机构、地址
- **通知：** 您提出辞职的日期（默认为今天）以及通知期（以月计）
- **可选：** 一项希望在信中提及的主要成就

随后，它会显示**计算出的最后工作日**，您可以接受它或输入其他日期，并在生成前显示一份简要的核对信息。

> 提示：请使用您所选择的语言输入姓名、公司和成就 —— 它们会原样出现在信中。（屏幕上的提示语本身保持英文。）

### 最后工作日的计算方式

`最后工作日 = 通知日期 + 通知期（月）− 1 天`

例如，在 `2026-06-01` 提出辞职、通知期为 `1` 个月，则最后工作日为 `2026-06-30`。在回答问题时，您
随时可以覆盖计算出的日期。输入 `0` 个月会生成一封“即时生效”的信件。

---

## 输出

文件保存在 `generated_letters/` 中。文件名采用与所选语言对应的前缀，加上您的姓名，例如：

```
Resignation_Letter_Jane_Doe.docx             (英文)
Surat_Peletakan_Jawatan_Ahmad_bin_Ali.docx   (马来文)
辞职信_张伟.pdf                                 (中文)
```

（如果同名文件已存在，会自动加上编号，因此不会覆盖任何文件。）

---

## 说明

- **Word** 文件可完全编辑，并支持任何字符。
- 对于**英文／马来文**，PDF 使用标准内置字体；极少见的特殊符号可能无法在 PDF 中显示 —— 这类情况请
  使用 Word 文件。
- 对于**中文**，PDF 会嵌入您电脑上的中文字体（例如 Windows 上的 Microsoft YaHei 或 SimSun，macOS
  上的 PingFang）—— 几乎所有系统都自带这些字体。万一找不到任何中文字体，仍会生成包含完整中文信件的
  Word (.docx) 文件。
- 除首次的一次性软件包安装外，不会通过互联网发送任何内容。不使用任何收费服务。

---

## 疑难解答

- **Windows —— “Python has been installed… close this window and run again.”** 首次安装需要一个
  新窗口才能检测到 Python。请关闭窗口并重新双击 `run.cmd`。
- **Windows —— winget 不可用。** 请从 <https://www.python.org/downloads/> 安装 Python 3，并勾选
  *Add python.exe to PATH*，然后重新运行 `run.cmd`。
- **macOS —— “cannot be opened because it is from an unidentified developer.”** 右键点击
  `run.command` → **Open** → **Open**（仅首次需要）。
- **macOS —— 双击没有反应。** 在此文件夹中打开 Terminal 并运行 `chmod +x run.command run.sh`，
  然后再试一次。
- **首次运行时没有网络。** 首次的软件包安装需要联网。首次成功运行后，本工具即可离线使用。
- **缺少中文 PDF。** 如果电脑上没有中文字体，PDF 会被跳过，但 Word (`.docx`) 文件仍包含完整的中文
  信件（如需 PDF，可打开该文件并在 Word 中“另存为 PDF”）。

---

## 项目文件

| 文件 | 说明 |
|------|------------|
| `run.cmd` | Windows 启动器（双击） |
| `run.command` | macOS 启动器（在 Finder 中双击） |
| `run.sh` | macOS / Linux 启动器（终端） |
| `generate_letter.py` | 负责询问问题并撰写信件的程序 |
| `requirements.txt` | 自动安装的 Python 软件包 |
| `.gitignore`、`.gitattributes` | 仓库维护（忽略 `.venv/` 和 `generated_letters/`；保持脚本换行符正确） |
| `generated_letters/` | 保存生成信件的位置（首次运行时创建；不纳入版本控制） |

---

## 系统要求（自动处理）

- Python 3.9+
- `python-docx`、`fpdf2`（见 `requirements.txt`）
- 中文 PDF 需要：系统上已有的 CJK 字体（Windows 和 macOS 标准自带）
