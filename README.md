# infographic-layouts

一个用于 **Claude for PowerPoint 插件（Office.js / OOXML）** 的 Skill：复刻一套成熟中文信息图模板的**正文页版式**。

提炼自一套 80 页的"版式集合"deck，给出环形闭环 / 中心辐射 / 菱形放射 / 同心圆关键词云 / 横向流程 / 多列架构 / 大数字网格等原型的**几何构造配方**与 **Office.js/OOXML 实现路径**，并附带大量实测 API 陷阱与降级方案。

## 用途
当你要"把一个观点设计成图示"、复刻这类信息图正文页、或要求"按那套版式排"时使用。不含标题/封面设计与品牌 chrome。

配色/叙事逻辑可与 `content-page-storyline` 配合（那个管"选哪种图示 + 故事线"，本 skill 管"具体怎么用 API 把图示画出来"）。

## 结构
- **§0 铁律** — 字体（微软雅黑 Bold/Light + CJK ea）、形状内文字居中五件套、风格禁忌
- **§1 设计系统常量** — 调色板、元件库、等距排布公式、API 真相速记
- **§2 原型配方库 E–K** + 逻辑→配方速查索引
- **§3 复刻校验清单**
- **§4 实现路径速查** — 哪些能 addGeometricShape，哪些要 OOXML
- **§5 API 实测陷阱**（PowerPointApi 1.5）
- **§6 原型配方库 L–T**（全册补全）
- **§7 实战补丁** — 真实品牌 deck 复刻踩坑与新配方

## 安装
把 `infographic-layouts/` 目录放进你的 Claude Code skills 目录即可（`SKILL.md` 为入口）。
