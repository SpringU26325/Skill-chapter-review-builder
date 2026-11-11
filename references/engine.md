# 引擎内部机制（只在改 `scripts\复习资料_build.py` 时读）

> **什么时候读**：要改编排版引擎的时候。**日常写复习资料不需要读这一份。**
> 只写 md 时必须守住的几条，见 `md-dialect.md` 第 7 节。

## 一、引擎已自动处理（写内容时不用管，但别绕过）

| 坑 | 引擎的处理 |
|---|---|
| 矩阵定界符丢失（`A=[` 被合成一个 run，括号只有一行高） | `tex2omml` 自建 `m:d` + `m:m`：按 `\\` 分行、`&` 分列，单元格递归转 OMML |
| `\|` 吞内容（`‖v-w‖²` 变成 `‖²`） | `_prep()` 全局替换 `\|` → `\Vert` |
| `\qquad` / `\quad` 失效（`<mspace>` 被 XSL 丢弃） | token 层切分，插 `m:r/m:t`，内容用 EM SPACE（U+2003）+ `xml:space="preserve"` |
| 粗体里嵌公式（`**零向量 $\mathbf{0}$**` 漏出源码） | `**...**` 解析后**递归**再解析内部 `$...$` |
| 标题里嵌公式（`### 给定 $\mathbf{b}$` 漏出源码） | `h1/h2/h3` 与容器标题一律走 `add_inline()`，不直接 `add_run(text)` |
| 公式没指定数学字体（继承等线 → 字形回退、像教材截图） | `polish_math()`：遍历所有 `m:r` 插 `w:rFonts ascii/hAnsi/cs="Cambria Math"`；`m:nor` 的中文正体改回等线 |
| Unicode 数学字母（`𝐚` U+1D41A、`𝑥` U+1D465） | `polish_math()` 做 NFKC 归一化还原 ASCII，按码位区间补 `m:sty`（b / i / bi） |
| 相邻矩阵贴死（`[A B]=[I 0][…]` 括号叠一起） | token 循环记 `prev`：相邻矩阵插 1em（`\u2003`），矩阵后接字母插 thin space（`\u2009`） |
| 代码块行距过大（`│ ├─` 被切断，树"一节一节"） | 代码块行距固定 **1.0** |
| 文末空页（文档以表格结尾，Word 隐式补段推到下页） | `trim_trailing_empty()`：删多余空段，显式补一个 1pt 段落 |
| 空 run 缺字体（`m:sSub` 的空底数） | `polish_math()` 对空 run 也补字体，否则"缺字体 run = 0"指标被打破 |

## 二、排版留白的三条经验

- **大 box（>半页）按主题拆成两个小 box**，否则整块 `cantSplit` 推到下页，前一页留 40% 空白
- **代码块超过约 14 行会被切成两截**（等宽 9pt、行距 1.0 时一页约 14 行）：注释合并到行尾、去掉空行与中间变量
- **速查卡前的附录页若只剩几行**（`page_break_before` 造成 85% 留白）：用**加题**补，不要删内容

## 三、python-docx 实现细节

- `pPr` / `tblPr` / `tcPr` 的子元素**有 schema 顺序**，直接 `append` 会写坏文件。用 `insert_element_before(el, *后续标签名)`
- 页眉不要用 tab 右对齐（会被 `Header` 样式的默认制表位干扰）。用 **1×2 无边框表格**，右格右对齐，表格 bottom border 当页眉线
- 容器块用单格表格 + `tcBorders` 左边框 + `w:shd` 底纹 + `w:cantSplit`（防框体被分页腰斩）
- 代码块必须给每个 run 补 `xml:space="preserve"`，否则连续空格被折叠，ASCII 树错位

## 四、LaTeX 转换的两个硬坑

1. **`\left| … \begin{bmatrix} … \end{bmatrix} … \right|` 必崩**：latex2mathml 抛 `ExtraLeftOrMissingRightError`，`add_math` 降级成红色 LaTeX 原文（`公式总数` 里不计失败，只能靠构建日志的 `!! 公式失败` 发现）。求行列式的绝对值时不要写成"竖线包住一个矩阵"，改成行列式单独成块、正文补一句"取绝对值"。
2. **行内公式以 `|` 开头且有 `\det`，`\det` 会被吞**：`$|\det A|$` 渲染成 "| dA|"；同样的 `\det` 放在 `$V=|\det A|$` 里却正常。定界符解析在公式开头时走了另一条分支。统一写 `$|A|$`。
3. 顺带确认：`\dfrac` 没有专门处理（用 `\frac`）；`\\[6pt]` 会漏成字面文本；`\begin{vmatrix}` 虽然引擎支持（定界符 `|`），但与已有各章的 `\det\begin{bmatrix}` 写法不一致，且竖线不会随行高拉伸，仍以 bmatrix 为准。

## 五、改引擎后的影响面

**这是引擎级改动，不是内容级。** 改完 `scripts\复习资料_build.py` 或新增容器类型后：

1. 所有已交付章节**必须全量重建**（`复习资料_重建.py --all`）；
2. 视觉验收从「关键页抽样」升级为**全量逐页看图**；
3. 重建前先备份，并确认没有并行会话正在构建。

只改路径、不改渲染逻辑的改动不受此限，但要用 A/B 对照证明产出逐字节相同：同一份 md、同一符号表、同一环境，改前脚本与改后脚本各建一次，比 `word/document.xml` 的 SHA-256。

## 六、文件被 Word 占用时

`PermissionError: [Errno 13]` 的处理：

1. 先输出到源稿目录（引擎自动降级，退出码 3，**非 0**，不允许当作已交付）；
2. 用 COM `GetActiveObject("Word.Application")` 查该文档的 `Saved` 属性；
3. **只有 `Saved=True`（无未保存修改）** 才允许 `Close(0)` 后替换并重新打开；
4. `Saved=False` 一律停下让用户自己关，**绝不强关有改动的文档**；
5. 查真实持有者用 `scripts\_wholock.py <docx>`，不要靠猜。
