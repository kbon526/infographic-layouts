---
name: infographic-layouts
description: 中文信息图正文页版式复刻库——把一个观点设计成图示。提炼自一套80页"版式集合"deck，给出环形闭环/中心辐射/菱形放射/同心圆关键词云/横向流程/多列架构/大数字网格等原型的几何构造配方与 Office.js/OOXML 实现路径。当用户要复刻这类信息图版式、"把内容做成图示"、或要求"按那套版式排"时使用。不含标题/封面设计与品牌 chrome。
---

用 Claude for PowerPoint 插件（Office.js / OOXML）复刻一套成熟中文信息图模板的**正文页版式**。本 skill 给可直接落地的几何构造配方与实现路径推导。**不含标题/封面设计、目录、页脚等品牌 chrome**——那些沿用现有模板或母版。配色/叙事逻辑可与 `content-page-storyline` 配合使用（那个 skill 管"选哪种图示 + 故事线"，本 skill 管"具体怎么用 API 把图示画出来"）。

## When to use
- 用户要**复刻这类信息图正文页**：环形闭环、中心辐射、菱形放射、关键词云、横向流程、多列架构、大数字网格等
- 用户说"把这个观点 / 这段内容做成一页图示""按那套版式排""做成 XX 那种图"
- 不适用：纯标题/封面/目录设计、单纯改字体配色（走母版）

---

## 0. 铁律（每页都必须遵守）

### 字体（CJK 关键）
- **重点 / 标题 / 强调**：微软雅黑 Bold。OOXML 同时写 latin 与 ea，并 `b="1"`：
  `<a:rPr lang="zh-CN" sz="..." b="1" dirty="0"><a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/></a:rPr>`
- **正文**：微软雅黑 Light（细体），`b="0"`：
  `<a:rPr lang="zh-CN" sz="..." b="0" dirty="0"><a:latin typeface="微软雅黑 Light"/><a:ea typeface="微软雅黑 Light"/></a:rPr>`
- **禁用微软雅黑 Regular 字重**——任何文字非粗即细。
- ⚠️ 中文字形走 `<a:ea>`（East Asian），**只写 `<a:latin>` 不生效**，必须 latin + ea 都写。
- 关键词内联强调：变蓝（`<a:srgbClr val="3665D9"/>`）+ `b="1"`。

### 形状内文字（禁止两层叠加）
- 文本**直接写进背景形状自带的 textFrame**，不要"背景框 + 透明文本框"。
- 形状内文字**居中五件套**（Office.js）：`alignment="Center"`、`verticalAlignment="Middle"`、`autoSizeSetting="AutoSizeNone"`、`wordWrap=false`、四边 `margin*=0`。
- 骨架用 `execute_office_js` 的 `addGeometricShape` 摆位 + 上述五件套写入纯文本；**带格式/变色加粗的正文用 `edit_slide_text` / `edit_slide_xml` 二次写入**（Office.js 写不了富文本）。

### 风格禁忌
- **不要"左侧彩色竖条 + 整段文字"的 accent-bar 卡片**（AI 味重）。强调靠：编号 / 加粗 / 变蓝 / 分隔线 / 纯色块。

---

## 1. 设计系统常量（画布 960×540，单位 pt，实测自模板）

### 调色板
| 角色 | HEX | 用途 |
|---|---|---|
| 主蓝 PRIMARY | `#3665D9` | 强调、胶囊底、环、结论 banner、变蓝关键词 |
| 深环 DARK | `#000000` | 双环中的深色环、深色 hub |
| 白 | `#FFFFFF` | 深底上的文字、卡片底 |
| 浅蓝 TINT1 | `#D7E0F7` | 图标圈外环、浅填充 |
| 浅蓝 TINT2 | `#E6ECFA` | 卡片 / 带背景 |
| 极浅蓝 | `#F2F5FF` / `#F7F9FF` | 大面积带背景 |
| 背景 | 白 + 极淡蓝径向 wash | 角落放超大柔和圆/矩形做 wash（画布外溢出，如 656×656 @ 负坐标） |

> ⚠️ 真实 deck 多有自有品牌色（非本库默认蓝）。落地时整体替换为目标 deck 品牌色，见 §7「品牌色：别套默认蓝」。

### 元件库（所有版式由这几种基元拼成）
- **胶囊 pill**：`RoundRectangle`，高 16–29。两型——①蓝底白字（PRIMARY 填充）②白底蓝边蓝字。圆角调到半高（adj≈0.5）。
- **卡 card**：`RoundRectangle` 白底，细描边或浅阴影；可顶部加蓝色标题带。
- **图标圈**：双层圆描边环——外圈 TINT1 / 内圈 PRIMARY（或反）+ 居中数字/小图标。（API 字符串是 `"Ellipse"`，非 `"Oval"`，见 §5#1。）
- **连接线**：正交连接（正上下左右）用 2pt 细 `Rectangle` 当连接器最稳；斜向连接用 `<p:cxnSp>` + `straightConnector1`（见 §7#11）。虚线 `<a:prstDash val="dash"/>`，箭头用 arrowheads；流程折行处**必须有连接箭头**。（`addLine(x1,y1,x2,y2)` 无效，见 §5#3。）
- **结论 banner**：全宽 `RoundRectangle` ≈850×60，蓝底白字，钉死一句结论。
- **环 / 弧**：`BlockArc`（粗环带）、`Arc`（细弧）；角度用 adj 调起止。
- **几乎每个子组件应打成 Group**，方便整体移动/对齐。

> **⚠️ API 真相速记（写代码前先读 §5/§7，避免照早期描述踩坑）**
> 1. 圆形 API 字符串是 `"Ellipse"`，不是 `"Oval"`（模板里形状名显示 Oval 会误导）。
> 2. `setZOrder` 抛 NotImplement——**层级只能靠创建顺序**，背景 wash / 连接线 / 底框必须**最先**建。
> 3. `addLine(x1,y1,x2,y2)` 无效——正交连接用细 `Rectangle`，斜向用 `cxnSp`。
> 4. CJK 字体两步走：Office.js 只写 `<a:latin>`，须 `edit_slide_xml` 把 latin 镜像到 `<a:ea>`。

### 通用搭建流程
1. 新建幻灯片用「仅标题」或「空白」版式。
2. 先用**公式算出所有 x / pitch**（见下），再 `addGeometricShape` 摆骨架；同类元素用循环保证等距等尺寸。**背景/连接线/底框最先建（靠创建顺序置底）。**
3. 形状内写纯文本 + 居中五件套。
4. `edit_slide_text` 二次写入富文本（关键词变蓝加粗、字体 bold/light）；并把 latin 字体镜像到 ea（§5#4）。
5. `verify_slide_visual` 逐页核对：对齐、咬合、对称、是否落到证据。

### 等距排布公式（务必用公式，别手填）
- **N 列等分**：`pitch = (usableW − itemW) / (n−1)`，`x_i = x0 + i*pitch`。usableW≈ 860（左右各留 ~50）。
- **chevron 互锁**：`pitch = 宽 − overlap`（overlap≈8）。
- **环形 N 节点**：`角度θ_i = startDeg + i*360/N`；`x = cx + R*cos(θ) − w/2`，`y = cy + R*sin(θ) − h/2`（θ 用弧度）。

---

## 2. 原型配方库

每个配方含：逻辑类型 / 代表页（可 `verify_slide_visual` 该 id 对照）/ 结构 / 实测坐标与公式 / **微软 API 实现路径**（哪些好画、哪些难、降级方案）。坐标是模板实测锚点，复刻时按内容缩放。

> N×M 能力矩阵、双卡对比、五步流程矩阵 → 见 `content-page-storyline` 的配方 A–D，本 skill 不重复。

### 逻辑 → 配方 速查索引
先想清"这页要表达哪种逻辑关系"，再据下表选原型；同一逻辑给了多个候选时，按信息量/数据量从上到下挑。

| 逻辑关系 | 首选配方 | 备选 |
|---|---|---|
| 能力矩阵 / 多维对比 | A–D（见 `content-page-storyline`） | — |
| 顺序 / 步骤 / 流程 | **L** 横向阶段流程 | **E** 蛇形折行（步数多）、**Q** 增长曲线（带时间轴） |
| 多阶段 / 体系架构 | **F** 多列阶段架构 | **S** 模块 hub |
| 生态 / 中枢放射 | **G** 中心辐射 hub | **S** 模块 hub |
| 闭环 / 循环 / 双向关系 | **H** 环形闭环 / 双环 | — |
| 一个中心 → N 个维度 | **I** 菱形四角放射 | **G** 中心辐射 |
| 发散 / 头脑风暴 / 标签云 | **J** 同心圆关键词云 | **T** 思维导图 |
| 规模 / 成果 / 数字证据 | **K** 大数字指标网格 | **R** 数据图表（有原始数据时） |
| 多源收敛 / 漏斗 / 沙漏 | **M** 漏斗汇聚 | — |
| 层级 / 占比 / 递进 | **N** 金字塔 | **T** 思维导图 |
| 交集 / 融合 / 聚合 | **O** Venn 重叠圆 | — |
| 两维度交叉定位 | **P** 坐标轴象限 | **P+** 机会缺口象限（重要度×满意度，见 §7） |
| 真实数据（趋势/占比/对比） | **R** 数据图表页（必用真图表） | **K** 大数字（只钉单点结论时） |
| 分类拆解 / 树状展开 | **T** 思维导图 | **N** 金字塔 |

### 配方 E — 横向流程 + 蛇形折行（代表页 283#0「在线问诊」）
- **逻辑**：顺序 / 步骤，步数多则折行。
- **结构**：大容器 `RoundRectangle 849×314 @(55,183)`；内部第一行 chevron/卡左→右，行尾向下箭头折到第二行右→左（蛇形）；每步 = 编号圈 + 步名胶囊。
- **API 路径**：全部 `addGeometricShape` 可画。chevron 用 `"Chevron"`；折行箭头用 `"BentArrow"`（`addLine(x1,y1,…)` 无效，见 §5#3）。第二行**步序反向**排列，箭头方向也反向。
- **坐标**：容器内 padding≈30；行高≈100，两行 y≈216 / 340；每行 N 步用列等分公式。

### 配方 F — 多列阶段架构（代表页 437#0「机构建设」）
- **逻辑**：多阶段 / 多层级的体系架构，列间用横向箭头串联。
- **结构**：左右两个**淡色大圆** `Ellipse 353×402 @(17,94)&(589,94)` 作背景氛围；中部多列卡（列 x≈59 / 151 / 256 / 488 / 721，卡宽≈204），每列纵向堆 2–3 张卡；底部全宽 banner `RoundRect 840×41 @(58,461)`。
- **API 路径**：卡与 banner 全部 `addGeometricShape("RoundRectangle")`；列间箭头用 `"RightArrow"`。背景大圆设很浅的 TINT 填充并**最先创建**（`setZOrder` 不可用，靠创建顺序置底，见 §5#2）。
- **复用**：列数 / 每列卡数变了，用列等分公式重算 x，纵向用固定行距堆叠。

### 配方 G — 中心辐射 hub（代表页 355#0「智能大脑」）
- **逻辑**：生态 / 中枢 + 多维能力放射。
- **结构**：中心同心圆 hub（`Ellipse 202 @(377,199)` 内 / `239 @(359,180)` 中 / `400 @(278,86)` 外），**中心≈(478,300)**；左右两组 `BlockArc 320×320 @(96,137)&(571,141)` 作装饰半环；两侧沿弧线对称排**胶囊列**（每侧 4–6 个 pill）；底部 banner `RoundRect 850×66 @(55,454)`；画布外四角超大矩形做淡蓝 wash。
- **API 路径**：同心圆 = 3 个 `Ellipse` 同心（算 left=cx−r, top=cy−r）；半环 = `BlockArc`（adj 调环宽与起止角）；胶囊列**用环形 N 节点公式**沿半径 R≈170 在左右各 ±一定角度内等分布点。中心文字写进最内 Ellipse 的 textFrame。
- **难点**：BlockArc 起止角靠 adj 手调，先画后 `verify_slide_visual` 微调。

### 配方 H — 环形闭环 / 双环（代表页 311#0「社区核心体验」）
- **逻辑**：闭环 / 循环 / 双向关系（如 生产者↔消费者）。
- **结构**：左**深色环** + 右**蓝色弧** `Arc 421×421 @(450,101)`（两环并接成 ∞/闭环感）；中心双圆 `Ellipse 152 @(404,236)` + 内 `124`，**中心≈(480,312)** + 中心标签；环之间对称**花瓣 freeform** `150×289 @(295,167)&(515,167)`；两侧竖向 stage 胶囊列 `38×215 @(84/838,204)`，每列堆 3–4 个阶段 pill。
- **API 路径**：
  - 环：粗环带用 `BlockArc`（推荐，可调环宽），细弧用 `Arc`；左右两环用不同 adj 起止角拼成闭环。
  - 中心双圆 = 两个同心 `Ellipse`。
  - **花瓣 freeform 无法用 addGeometricShape**——降级：① 用 `"Teardrop"` 或 `"Pie"` 近似；② 要精确则用 `edit_slide_xml` 写 `<a:custGeom>` 自定义路径。
  - 侧列 pill 用列内固定行距堆叠。
- **难点**：最复杂原型。建议先摆环 + 中心 + 侧列（占 80% 观感），花瓣作可选装饰。

### 配方 I — 菱形四角放射（代表页 313#0「用户画像」）
- **逻辑**：一个中心概念 → 四个维度 / 要素。
- **结构**：常**左文右图**——左半正文块 `~331×260 @(54,140)`，右半菱形图 `~505×476 @(410,32)`，**菱形中心≈(662,270)**；中心 `Diamond` 蓝底白字（写两行：动作 + 结果）；四角各一节点（图标圈 + label + 短说明），用**虚线**连到中心。
- **API 路径**：中心 `addGeometricShape("Diamond")`；四节点用环形公式 N=4、startDeg=−90、R≈150 算位置（上下左右）；连接用 `cxnSp`（斜向）或细 `Rectangle`（正交）+ `prstDash="dash"`（`addLine` 无效，见 §5#3）。节点图标圈 = 双层 Ellipse。
- **复用**：维度数变 → 改 N（如 N=6 变六角放射，等分角度）。

### 配方 J — 同心圆关键词云（代表页 411#0「色彩情绪版」）
- **逻辑**：围绕一个主题发散 / 头脑风暴 / 标签聚类。
- **结构**：中心**深色实心圆**写主题（`Ellipse`，center≈(480,292)）；外围 2–3 圈**同心虚线圆**（`Ellipse` 无填充 + `prstDash="dash"`，由内到外半径递增）；圈上分布**前景胶囊**（PRIMARY pill）+ 背景**浅灰关键词**（纯文本，低饱和，作氛围层）。整组占 `550×550 @(205,17)`。
- **API 路径**：同心虚线圆 = 多个 `Ellipse` 无填充 + 虚线描边，同心算 left/top；胶囊/文字沿环形 N 节点公式按圈分布（每圈不同 R 与节点数）。
- **注意**：背景浅灰词刻意低对比是**设计意图**（氛围层），但前景胶囊/主题文字必须达标对比；`verify_slides` 的 contrast 告警里**只修前景层**，背景氛围词可保留。

### 配方 K — 大数字指标网格（落证据，常作底带或整页）
- **逻辑**：规模 / 实力 / 成果，用大数字钉死。
- **结构**：2×3 或 1×N 网格，每格 = 上方小灰标签（微软雅黑 Light）+ 下方**超大数字**（微软雅黑 Bold，变蓝，sz≈4000–6000）+ 单位。
- **API 路径**：网格用列等分 + 行距；每格一个透明/浅底 `RoundRectangle`，数字与标签写进其 textFrame（或分两段用 edit_slide_xml）。数字是主角，占格高 60%。
- **复用**：作整页时配一句结论 banner；作底带时压在中带图示下方 y≈404–515。

---

## 3. 复刻校验清单（每页过一遍）
1. **字体**：标题/强调 = 微软雅黑 Bold(latin+ea+b=1)；正文 = 微软雅黑 Light(latin+ea+b=0)；无 Regular。
2. **文本**：都写在背景形状的 textFrame 里（无透明叠层）；形状内居中五件套到位。
3. **几何**：同类元素等距等尺寸（用了公式）；环/弧对称咬合；列对齐。
4. **配色**：用模板调色板（主蓝 #3665D9 等，或目标 deck 品牌色），背景有淡蓝 wash，强调变色。
5. **禁忌**：无"竖条+整段文字"accent-bar 卡片。
6. **证据**：落到可量化数字（大数字 / 结论 banner）。
7. `verify_slide_visual` 对照代表页 id，逐项确认上述。

## 4. 实现路径速查（哪些能 addGeometricShape，哪些要 OOXML）
- **直接可画**（execute_office_js + addGeometricShape）：矩形/圆角矩形(胶囊/卡/banner)、Ellipse(同心圆/图标圈)、Chevron、Diamond、Triangle、RightArrow/BentArrow、BlockArc、Arc、Pie、Teardrop。（圆是 `"Ellipse"` 非 `"Oval"`，见 §5#1。）
- **连接线**：正交用细 `Rectangle`，斜向用 `cxnSp`/`straightConnector1`（`addLine` 无效，见 §5#3、§7#11）。
- **需 OOXML（edit_slide_xml + custGeom）**：花瓣/异形 freeform、精确曲线连接、波浪历程曲线。先用 Pie/Teardrop/BlockArc 近似，必要时再上 custGeom。
- **富文本/变色加粗/字体**：一律 `edit_slide_text` / `edit_slide_xml`（Office.js 的 textRange 写不了，部分版本读出还为空，见 §7#7）。
- **背景 wash**：超大柔和形状溢出画布 + **最先创建置底**（`setZOrder` 不可用）；或母版渐变背景（注意实底页会遮住，见 §7「背景 wash 的限制」）。

---

## 5. API 实测陷阱（本插件 Office 版 = PowerPointApi 1.5，必看）
实测复刻配方 I 时踩过的坑，照做可一次成型：

1. **圆形类型字符串是 `"Ellipse"`，不是 `"Oval"`。** 模板里形状名显示为「Oval」会误导——`addGeometricShape("Oval",…)` 抛 InvalidArgument。同理用 `"RoundRectangle"`/`"Rectangle"`/`"Diamond"`/`"Chevron"`/`"Triangle"`/`"BlockArc"`/`"Arc"`/`"Pie"`/`"Teardrop"`。
2. **`setZOrder` 是 NotImplement（抛错）。层级只能靠创建顺序**——先建的在底层。所以背景 wash、连接线、底框必须**最先**创建，节点/卡/文字后建。不要调用 `shape.setZOrder(...)`。
3. **`addLine(x1,y1,x2,y2)` 无效。** PowerPoint 签名是 `addLine(connectorType, options{left,top,width,height})`，用包围盒+flip 定义方向，斜线很别扭。**正交连接线（正上下左右）直接用 2pt 细 `Rectangle` 当连接器最稳**；斜向连接用旋转细矩形或 `edit_slide_xml` custGeom。
4. **CJK 字体两步走**：Office.js `font.name="微软雅黑 Light"` 只写 `<a:latin>`，中文字形不生效。建完形状后跑一次 `edit_slide_xml`，把每个 `<a:rPr>` 的 latin 字体**镜像到 `<a:ea>`**（ea 须紧跟 latin 之后）：
   ```javascript
   const rPrs = doc.getElementsByTagNameNS(NS_A,"rPr");
   for(let i=0;i<rPrs.length;i++){
     const rPr=rPrs[i], latin=rPr.getElementsByTagNameNS(NS_A,"latin")[0];
     if(!latin) continue; const tf=latin.getAttribute("typeface");
     let ea=rPr.getElementsByTagNameNS(NS_A,"ea")[0];
     if(!ea){ ea=doc.createElementNS(NS_A,"a:ea"); ea.setAttribute("typeface",tf);
       latin.nextSibling?latin.parentNode.insertBefore(ea,latin.nextSibling):latin.parentNode.appendChild(ea);
     } else ea.setAttribute("typeface",tf);
   }
   ```
5. **`addGeometricShape` 报错时是"部分提交"，不是全回滚。** 一旦某形状参数非法/调用 NotImplement，前面已排队的形状可能已落地、后面的丢失，且重跑会产生**重复形状**。对策：**分阶段 `context.sync()`**（如 标题→中心→连接器→节点 各一段），每段失败只影响该段；失败后**先 `list_slide_shapes` 查实际状态**再补，别盲目重跑。
6. **不要在同一批内读新建 shape 的 `.id`/`.left` 等属性**（未 `load`+`sync` 会抛 PropertyNotLoaded，导致整批失败）。要拿新形状的 id，单独再 `load`+`sync` 一次。

### 配方 I 补正（侧节点标签防溢出）
四角放射里**左/右节点的标签放到圆的正下方居中**（与上/下节点一致），不要放在圆的左/右侧——否则右标签溢出右边距、左标签压住左侧正文块。约束：任何元素右缘 ≤917、左缘 ≥36，且右图区不与左正文块（约 54–374）相交。

---

## 6. 原型配方库(扩展 L-T，全册补全)
基于对全册 80 页的结构聚类 + 跨家族视觉抽样得出。E-K 之外，这套模板还反复用到以下原型；坐标按内容缩放，字体/居中/调色板规则同 0-1 节。

### 配方 L — 横向阶段流程(最高频，5 个子变体)
- 逻辑：顺序/阶段/步骤。横排 N 个阶段，步间必须有方向指示(箭头/chevron/曲线折返)。
- 子变体(按内容挑)：
  1. 卡片+箭头串联：N 张等宽 RoundRectangle，间隙放 RightArrow 或细三角箭头。
  2. 大编号 ghost 卡(代表页55)：每卡左上角超大描边数字 01-04(白底无填充+粗描边或浅灰大字，作背景层先建)，上叠正文。
  3. 交错 staggered 卡(代表页23)：奇偶卡上下错位(y 交替 ±30)，字母/编号 A-E 标序。
  4. 阶段大圆+卫星小圆(代表页12)：每阶段一个大 Ellipse，周围挂 3-4 个小 Ellipse chip(环形 N 节点公式分布)，阶段间 chevron。
  5. 异形头卡(代表页71)：卡顶用盾形/气泡(Pentagon/WedgeRoundRectCallout 近似)作 header，内嵌图标+标签+按钮 pill+强调数字。
- API：等宽用 N 列等分公式；箭头优先 addGeometricShape("RightArrow")；ghost 大数字、卫星圆都先建(置底)。

### 配方 M — 漏斗/对向箭头汇聚(沙漏)
- 逻辑：多源收敛到一点，或一点发散到多路；或两端对向汇聚到中心。
- 结构：(1)漏斗(代表页11)：左侧多输入用虚线三角喇叭(Triangle 描边/虚线，宽端朝多、尖端朝一)收敛到中心圆，再发散。(2)沙漏/X 汇聚(代表页47)：上下左右四个 RightArrow/UpArrow 朝中心 Ellipse 汇聚，两侧挂 list。
- API：喇叭用 Triangle + 虚线描边(`prstDash="dash"`)无填充；中心圆写结论；箭头用方向性 GeometricShape。

### 配方 N — 金字塔(分层/分段矩阵)
- 逻辑：层级/递进/结构占比。
- 结构：(1)分层(代表页62)：顶 Triangle + 下方 2-3 个 Trapezoid 层叠(各层等高、上窄下宽)，右侧配 N 张解说卡，卡与层用浅曲线连。(2)分段网格三角(代表页36)：大三角被横竖切成 3×2 网格，左右斜边挂标签，高亮 1-2 格(蓝/黑底白字)。
- API：层用 Trapezoid 逐层加宽(width_i = w0 + i*Δ，居中 x = cx − width_i/2)；顶用 Triangle。分段网格用矩形裁切叠在三角上，或梯形分层+内部分隔线。

### 配方 O — Venn 重叠圆/聚合圆簇
- 逻辑：交集/融合/多要素聚合。
- 结构：(1)双向 Venn(代表页77)：两个半透明 Ellipse 与中心 Ellipse 相交，两侧用弧线虚线挂一排小图标圈。(2)多 Venn 簇(代表页76)：4-5 组重叠圆(攻/防/知/查/抓)，每组外环绕 pill 标签。
- API：圆用 Ellipse，相交靠坐标重叠+浅色 TINT 近似半透明(无透明 API 时用 OOXML 设 alpha)；弧线挂载用 Arc 无填充+沿弧布 chip。

### 配方 P — 坐标轴/象限关系图
- 逻辑：两维度交叉定位(现在/未来 × 态度/行为)。
- 结构(代表页67)：中心十字轴(两条带箭头的细矩形)，四端标维度名；节点(pill/小圆)按象限分布，可用曲线串成路径；顶部常配 2-3 张模型小卡。
- API：轴用方向箭头形状或细 Rectangle；节点用象限散点坐标；象限标签放四角。

### 配方 Q — 增长曲线/阶梯路径
- 逻辑：随时间/阶段递进上升。
- 结构(代表页24)：底部 x 轴(阶段名)+左 y 轴；一条上升曲线(Arc/Freeform/折线)，沿线由低到高摆 N 张阶段卡+圆点 marker；卡高度跟随曲线递增。
- API：曲线优先 edit_slide_xml custGeom(平滑)，降级用多段细 Rectangle 折线；阶段卡 y 递减(越右越高)。

### 配方 R — 数据图表页(柱/折线/组合/饼) 必用真图表
- 逻辑：呈现真实数据(趋势、占比、对比)。
- 代表页 13：柱状+折线组合图 + 顶部小标题 + 一个大数字(如 10.6%)钉关键结论。
- API：一律用 edit_slide_chart 生成 OOXML 真图表，绝不用形状拼。套模板风格：c:style val=2 继承主题、系列色用主蓝 #3665D9、数据标签 showVal=1、标题/轴/图例字体 ≥14pt 且设 latin+ea 微软雅黑。组合图 = 同一 plotArea 内 barChart + lineChart 共用类目轴。

### 配方 S — 模块 hub(中心多圆+双侧 list+汇聚)
- 逻辑：一个平台/系统由多模块组成，前后端/内外部要素汇入。
- 结构(代表页54)：中部横排 3 个 Ellipse 模块(中间实心蓝高亮)；左右各一竖向胶囊 RoundRectangle，内堆 3-4 行要点(前端/后台)；顶部一胶囊(系统名)虚线箭头向下指三圆，底部一胶囊汇总；背景用梯形渐变做聚光灯/隧道透视(两个对向 Trapezoid 浅蓝渐变，最先建置底)。
- API：三圆用列等分；侧胶囊高瘦 RoundRectangle；连接全用虚线箭头；背景梯形先建。

### 配方 T — 思维导图/树状结构
- 逻辑：层级展开/分类拆解(根→分支→叶)。
- 结构(代表页68)：左侧根节点(蓝圆角)→中列 2-4 个二级节点(黑底白字胶囊)→右列叶节点(白底胶囊)，用曲线连接(每级一组平滑曲线+端点小圆点)。
- API：节点用 RoundRectangle；连接曲线用 edit_slide_xml custGeom 贝塞尔(降级用 Arc/折线)；纵向等分各级节点 y。

### 背景母题(可复用装饰)
- 梯形渐变聚光灯/隧道：两个对向 Trapezoid(或大三角)浅蓝→白渐变，从四角朝中心收，制造纵深；最先建置底。
- 角落柔光大圆/大矩形：超大浅 TINT 形状溢出画布四角作 wash。
- 二者都靠创建顺序置底(本版 setZOrder 不可用，见 5 节)。

### 覆盖说明
本配方库(A-D 见 content-page-storyline；E-T 见本 skill)覆盖该 80 页模板经结构聚类得到的全部主要版式家族。未单列页面均为上述原型的参数化变体(改 N、改列数、换子变体、加减证据带)，按对应配方调参即可复刻。

---

## 7. 实战补丁（复刻真实红色品牌 deck 的踩坑与新配方）
> 来源：复刻一套红色品牌（主色 #C3272C，非本库默认蓝）正文页 deck 的实测。下列为本库 0–6 节之外、反复验证有效的补充。

### 品牌色：别套默认蓝
本库示例用主蓝 #3665D9，但真实 deck 多有自有品牌色。落地时一律改用目标 deck 的品牌色（读 `themePalette` 或现有页实测），把所有"主蓝"角色（强调/胶囊/环/banner/变色关键词/KPI 数字）整体替换；浅蓝 TINT 换成对应浅色（如红系用 #FBEAEB / #F7F7F8）。

### API 踩坑（本插件 Office=1.5，补 5 节）
7. **部分 Office 版 `textRange.text` 读出为空、`getTextFrameOrNullObject` 间歇抛 GeneralException**——不要用 Office.js 读/判正文，一律 `read_slide_text` / `edit_slide_xml` 走 OOXML；写富文本也走 OOXML。
8. **空 txBody 非法**：每个形状的 `<p:txBody>` 至少要有一个 `<a:p>`（纯装饰形状给 `<a:p/>`），否则 `insertSlidesFromBase64` 抛 GeneralException。
9. **改已有形状的填充/描边**：用 DOM 往 `spPr` 里插 `solidFill` 常因子元素顺序错、或误定位到别的 `<a:ext>` 而"不生效"（症状：只有部分卡片变色）。最稳：用 DOMParser 造一个完整 `<p:spPr>`（xfrm+prstGeom+fill+ln）整体 `replaceChild` 掉旧的。
10. **图标(pic)缩放/居中陷阱**：`pic.getElementsByTagNameNS(NS_A,"ext")[0]` 可能取到 blipFill 里 svg 扩展的 `<a:ext uri=...>` 而非 xfrm 的 `<a:ext cx cy>`，导致改尺寸无效（图标仍是原始尺寸、按原比例贴框左上角 → 在圆内偏上偏左）。MS 矢量图标常是 34×34 且锁纵横比；最稳用 Office.js 直接 `shape.width/height/left/top`，居中 = 圆心 − 尺寸/2。
11. **斜向连接线**：用 `<p:cxnSp>` + `straightConnector1`，off=两端外接框左上、ext=(|dx|,|dy|)、`flipV="1"` 切换 "/" 与 "\\"，虚线用 `<a:prstDash val="dash"/>`。放射/菱形四角连线都用它，比 addLine 稳。**先建连线、后建中心圆与节点圆**，让圆遮住线端（本版 `setZOrder` 不可用，全靠创建顺序）。
12. **梯形漏斗的文字别写进 custGeom 形状内**：`flipV` 的 trapezoid 会翻转文字；custGeom 形状内文字也易"顶对齐被裁"。做法：custGeom 画"宽顶梯形"（path: 0,0 → W,0 → W−ins,H → ins,H → close）只作色块，文字用**单独的居中覆盖文本框**叠在其上（anchor=ctr）。

### 新配方 P+ — 重要度×满意度"机会缺口"象限（数据洞察首选）
- **逻辑**：一组需求/属性各有"重要度"与"满意度"两分，定位"高重要·低满意"=切入缺口。比并排双条形信息力强得多。
- **结构**：x=满意度(低→高)、y=重要度(低→高，高在上)；**左上角(高重要低满意)画虚线红"机会缺口"框**；每项一个圆点+标签，落在缺口框内的点变红、框外灰；轴端写"满意度 →""重要度 ↑"。
- **API**：点=小 Ellipse；缺口框=dashed 红 roundRect（先建置底）；轴=细矩形。映射 `plotX=x0+(sat−satMin)*scaleX`、`plotY=yBottom−(imp−impMin)*scaleY`。
- **粗刻度(如 1–3 分)会出现重合点**：对同坐标点做 ±15px 抖动，别让点叠死、标签互压。

### 配方 R 补 — 折线+柱组合，折线标签压在柱上
- 渗透率/占比折线值域(如 3–8%)远小于柱(如 9–45)，折线贴底、数据标签压在柱基难读。
- 修法：①折线系列 dLbls 给**白底标签 chip**（spPr 白 solidFill+浅描边）保证任何背景都可读；②`dLblPos="t"`；③折线走**次数值轴**(第二根 `c:valAx`，`crosses=max`)并设 `max`（把如 8% 的系列轴 max 设 ~12）把折线**抬到短柱之上**，短柱处标签即浮在柱顶上方。
- 反向教训：**别把"真实数据图表页"硬改成手绘形状图**（如增长阶梯）——会丢失原生图表精度，且删 graphicFrame 后其 chart part 会在重新打包时被当无引用资源清除、不可逆。数据页保留/重建原生图表(配方R)，叙事感靠标题、KPI 卡、配色实现。

### 背景 wash 的限制
- 角落柔光/梯形聚光灯母题加在**母版或版式层，会被"每页不透明白底"遮住而不显示**。若 deck 每页有实底，wash 只能逐页加最底层形状或改各页 slide 背景填充——成本高。先确认页面是否透明底再决定是否做。

### 小标题/小节标题的垂直安全区
- 右栏/分区图示(Venn、hub 等)的顶边别贴着小节标题放。小节标题约占 `y≈137–170`，图示顶边从 `y≥178` 起，留 ≥10pt 间隙，否则圆/卡顶边会切进标题带（视觉重叠）。
