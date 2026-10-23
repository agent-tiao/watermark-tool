# watermark-tool

简体中文 | [English](README.en.md)

把 **Word / PDF** 转成**栅格化 PDF**，并把 logo 水印**烧进每一页像素**的防去水印工具；同时提供标准 **Agent Skill**，可被 Claude Code、opencode 等 AI 客户端直接调用。

## 简介

普通 PDF 水印是"独立图层/对象"，用 Acrobat、pdfcpu 等工具可以一键删除。`watermark-tool` 换了一个思路：

1. 先把每一页渲染成位图；
2. 把半透明 logo 以斜向平铺、隔行错位的方式压到像素上；
3. 重新打包成 PDF。

这样水印不再是可删除的对象，去水印只能逐页做图像修复 —— 有损、费时、成本高。

**特点**

- 支持 `.docx` / `.doc` / `.pdf`，按文件类型自动选择流程
- 水印可调透明度、密度、角度，支持交叉平铺
- 全程本地运行，文件不上传第三方
- 附带 Agent Skill，AI 可直接完成批量加水印
- Apache-2.0 开源，可自由修改和分发

**代价**：文字不可复制/搜索、文件体积变大。

## 原理

```
Word/PDF  →  (Word 转 PDF)  →  逐页栅格化(DPI)  →  叠加斜向平铺水印  →  重新打包成 PDF
```

- **栅格化**：正文与水印都变成像素，水印不再可分离、无法一键删除。
- **平铺 + 旋转 + 半透明**：铺满全页无处可裁、不遮挡正文、但足够显眼。
- 代价：文字不可复制/搜索、文件体积变大。

## 目录结构

```
watermark-tool/
├─ skills/
│  └─ pdf-watermark/         # Agent Skill
│     ├─ SKILL.md
│     └─ scripts/
│        └─ watermark_pdf.py # 主脚本
├─ assets/
│  └─ sample_logo.png        # 通用示例水印
├─ input/                    # 待处理源文件（已 gitignore）
├─ output/                   # 生成结果（已 gitignore）
├─ requirements.txt
├─ LICENSE
├─ README.md
├─ README.en.md
└─ .gitignore
```

> 发布到公开仓库前请确认：`input/`、`output/` 与品牌 logo 已在 `.gitignore` 中排除，避免泄露客户文档或内部素材。

## 环境依赖

```bash
python -m pip install -r requirements.txt
```

- **Windows + 已安装 MS Word**（用于 `.docx → .pdf` 转换；只处理 `.pdf` 时不需要）
- Python 包：PyMuPDF、Pillow、pywin32、numpy

## 快速开始

在仓库根目录执行：

```bash
# 处理 Word
python skills/pdf-watermark/scripts/watermark_pdf.py "input/xxx.docx" -l "assets/sample_logo.png" --white-to-alpha -d 300 --opacity 0.24 --tile-size 0.26 --gap 0.08 -o "output/xxx_watermarked.pdf"

# 处理 PDF（直接栅格化加水印）
python skills/pdf-watermark/scripts/watermark_pdf.py "input/xxx.pdf" -l "assets/sample_logo.png" -d 300 --opacity 0.24 --tile-size 0.26 --gap 0.08 -o "output/xxx_watermarked.pdf"
```

把 `assets/sample_logo.png` 换成自己的 logo 即可（透明 PNG 效果最佳；白底 JPG 需加 `--white-to-alpha`）。

## 作为 Agent Skill 使用

`skills/pdf-watermark/` 遵循 [Agent Skills](https://agentskills.io) 规范（`SKILL.md` + 脚本），可被 Claude Code、opencode 等客户端加载。

**安装到全局**

```bash
# macOS / Linux
mkdir -p ~/.claude/skills
cp -r skills/pdf-watermark ~/.claude/skills/

# opencode
mkdir -p ~/.config/opencode/skills
cp -r skills/pdf-watermark ~/.config/opencode/skills/
```

```powershell
# Windows PowerShell
New-Item -ItemType Directory -Force "$HOME\.claude\skills" | Out-Null
Copy-Item -Recurse "skills\pdf-watermark" "$HOME\.claude\skills\"
```

也可以放到项目内：`.claude/skills/` 或 `.opencode/skills/`。安装后重启客户端。

**使用**：直接对 AI 说

> 用 pdf-watermark 给 `input/report.pdf` 加水印，透明度 0.24，输出到 `output/`。

AI 会按 `SKILL.md` 中的流程选择参数、执行脚本并检查结果。

## 输入类型兼容

脚本按输入类型自动选择流程：

| 输入 | 处理方式 |
|---|---|
| `.docx` / `.doc` | 先调用 MS Word 转 PDF，再加水印 |
| `.pdf` | 直接栅格化加水印 |
| 无后缀但文件头为 `%PDF-` | 按 PDF 处理（防后缀丢失） |

## 参数说明

| 参数 | 默认值 | 说明 |
|---|---|---|
| `input` | 必填 | 输入文件（.docx / .doc / .pdf） |
| `-o, --output` | `<名称>_watermarked.pdf` | 输出 PDF 路径 |
| `-l, --logo` | 必填 | 水印 logo 图片（PNG 透明最佳） |
| `-d, --dpi` | `300` | 栅格化分辨率。300=打印质量，150~200=屏幕足够，体积更小 |
| `--tile-size` | `0.22` | 水印宽度占页宽比例，越小铺得越密 |
| `--rotate` | `45` | 水印旋转角度 |
| `--opacity` | `0.12` | 水印透明度 0~1，越大越显眼、越难去 |
| `--gap` | `0.15` | 水印之间的间距（占水印尺寸比例） |
| `--no-stagger` | 关 | 关闭隔行错位（默认开启"砖块式"错位） |
| `--white-to-alpha` | 关 | 把接近白色的像素当作透明（**白底 JPG logo 必加**） |
| `--crosshatch` | 关 | 交叉平铺：正反两个方向叠加，更防删，但 logo 会互相重叠 |
| `--angle2` | `-45` | 交叉平铺时第二层角度 |
| `--opacity2` | 同 `--opacity` | 交叉平铺时第二层透明度 |

## 推荐参数

- **单层不重叠（推荐，品牌清晰）**
  ```bash
  --opacity 0.24 --tile-size 0.26 --gap 0.08 --rotate 45
  ```
- **交叉平铺（更防删，但 logo 会重叠、品牌辨识度下降）**
  ```bash
  --crosshatch --opacity 0.22 --opacity2 0.15 --tile-size 0.22 --gap 0.08
  ```

## 使用建议与局限

- **无法 100% 防删**。纯 PDF 去水印工具（Acrobat 删水印 / pdfcpu 等）对本工具产物**无效**（没有可删对象）；但 AI/PS 图像去水印**仍能去除一部分**，只是有损、需逐页处理、成本高。
- 水印**越明显越难去**；太淡可被整体调对比度抹除。
- 栅格化后**不可复制、不可搜索**、不支持无障碍阅读；若需可检索文档，本方案不适用。
- 通过聊天工具发送时，请选 **「发送文件」**（而非转成图片/链接），避免二次压缩。
- Windows 上若 Word 偶发 COM 报错，先结束残留的 `WINWORD` 进程再重试。

## License

[Apache-2.0](LICENSE)。
