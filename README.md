# chapter-review-builder

把一章教材（PDF、课件或笔记）编译成可直接打印的 Word 复习资料。输入是 Markdown 源稿，输出是 `.docx`；公式为 Word 原生 OMML，可编辑、可缩放。

适用：单章复习资料的初做、返工、补题、重排、质检。不适用于通用 Word 排版美化。

## 成品结构

固定顺序，页数舒适区 8–12 页，硬上限 15 页，超过必须拆册（上 / 中 / 下 / 拓展）。

| 位置 | 内容 |
|---|---|
| 页眉 | 科目、教材版本、原书页码 |
| 核心考点 | 本章解决的问题与考试落点 |
| 本章脉络 | 层级缩进的树状结构，可含行内公式，每行可点击跳转 |
| 正文 | 按节展开：讲解、注释、国内考法、例题；每个知识点配一道题 |
| 附录 | 易错点、课后习题、一页速查 |

## 跳转

文档自带内部链接，不需要手写标记：

- 首页「速达」指向 本章脉络 / 易错点 / 附录题库 / 一页速查
- 「本章脉络」的每一行连到对应小节
- 小节标题下的「对应题目」连到附录里那一道题

标题带大纲级别，Word 左侧导航窗格按三级展开，标题左侧可折叠。链接由标题名与附录标签推导：附录条目写成 `【课后习题 6.5 第 1 题】` 的册子才有题目跳转，写成 `【习题 3.1】`（那是题号）的没有。

## 公式

链路：LaTeX → MathML → `MML2OMML.XSL` → OMML。

`latex2mathml` 默认把 `1/2`、`x/2` 输出为斜杠分式（OMML 中 `m:type="lin"`），Word 显示为压扁的斜杠。引擎在出口统一改为上下堆叠（`bar`），`audit` 统计 `lin` 数量作为回归指标。

## 构建与校验

```bash
# 只重建本章，附页数统计（并行会话必须带 --only）
python "<skill>\scripts\复习资料_重建.py" --only 高等数学_第6章 --pdf

# md 级前置检查
python "<skill>\scripts\复习资料_lint_md.py" 复习资料\_生产\_源稿\高等数学_第6章.md

# 构建后审计
python "<skill>\scripts\复习资料_audit.py" <md> <docx>
```

`<skill>` 指本目录，即 `SKILL.md` 所在层。

三层检查，职责不重叠：

1. **lint**（构建前）：私有方言、行内矩阵、结构块内裸公式等写法问题。
2. **audit**（构建后）：公式转换失败数、私有区字符、悬空引用。
3. **视觉验收**：公式挤压、上下标重叠脚本查不出，需渲染 PDF 后人工抽检关键页。

退出码 0 才可交付；WARN 不阻断，但须在自检报告里逐条说明。

## 安装

```bash
cd <宿主 skills 目录>
git clone https://github.com/SpringU26325/Skill-chapter-review-builder.git chapter-review-builder
```

仓库根即 `SKILL.md` 所在层，整个目录就是 skill 包，不需要二次打包。

用 GitHub 的 `Code → Download ZIP` 亦可，解压后须把目录名改为 `chapter-review-builder`（压缩包默认带 `-main` 后缀，目录名即 skill 名）。

## 环境

| 项 | 要求 |
|---|---|
| 系统 | Windows |
| Office | Microsoft Office |
| Python | `python-docx`、`latex2mathml`、`lxml`、`pymupdf` |

WPS 不支持：接管 `Word.Application` 后 COM 层静默失败，且不带 `MML2OMML.XSL`。页数与 PDF 导出、公式转换各依赖其一。

## 已知限制

- 版式只有一套：七种容器加自动分页，不支持自定义风格。
- 一次处理一章，页码范围由人工指定。
- 教材 PDF 文字层不可信，OCR 常把数学符号识别错（`λ` 识别为 `入`）。公式必须回原书按页面图重建。
- Word COM 偶发阻塞，脚本设 180 秒熔断并清理自身拉起的进程，阈值由 `CRB_RENDER_TIMEOUT` 调整。

## 目录约定

```text
Ebooks\                  教材 PDF，只读
复习资料\<科目>\          本科目素材：真题、笔记、课件等原件
复习资料\<科目>\_成品\    成品 docx
复习资料\_生产\_源稿\     各章 Markdown、<科目>_notation.md
复习资料\_生产\_render\   渲染过程文件，交付即清
复习资料\_生产\_locks\    并发锁，删了保护失效
复习资料\_生产\_台账.md   进度、页码、分册
复习资料\_生产\_踩坑清单.md  引擎、公式、COM 的已知问题
```

命令细节见 `SKILL.md`，md 方言见 `references/md-dialect.md`，写作风格见 `references/wording.md`。
