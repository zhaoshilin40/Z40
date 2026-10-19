---
name: tnu-ppt
description: Build a Tianjin Normal University (天津师范大学) presentation by editing the official TNU template in place — fill existing placeholders and shapes, delete unused template slides, never add self-drawn text boxes. Use when the user wants a 天津师范大学 / TNU 汇报答辩 / 课程汇报 PPT, or names the file 天津师范大学毕业论文汇报答辩通用ppt模板.pptx. Covers slide-page selection, placeholder addressing, title-box widening, native chart and table rewriting, residue detection (untouched template text), and PowerPoint COM rendering for verification.
---

# 天津师范大学 PPT 模板编辑（原页编辑法）

模板文件自带的版式已经排好：页眉红条、校徽、页脚、卡片底纹、装饰图形全都到位。
**正确做法是改它已有形状里的文字；新建文本框就等于在模板上又画了一套版式，一定难看。**

这条规则不是洁癖，是可以量化的：成品里每一个形状都必须能在对应模板页找到同
`shape_id` 的原件。`verify` 子命令会逐页核对这个数字，**必须为 0**。

## 一、先定位模板

模板通常在这个位置（用户提供的原始文件，别复制副本，直接读）：

```
C:\Users\12806\Desktop\天津师范大学毕业论文汇报答辩通用ppt模板.pptx
```

62 页 / 28 个版式 / 画布 13.333×7.5 英寸。`references/catalog.md` 是它的完整形状
清单（每页每形状的 id、几何、占位符 idx、字号、图表类型），**选页时查这份，不要
靠猜**。模板改版后重新生成：

```
python -X utf8 <skill>/scripts/tnu.py catalog <模板.pptx> --out <skill>/references/catalog.md
```

## 二、四个子命令

解释器固定用 `C:\Users\12806\Python314\python.exe`（Windows Store 的 `python`
是占位符，跑不起来）。**必须加 `-X utf8`**，否则中文参数与输出会乱码。

| 命令 | 用途 |
|---|---|
| `inventory <pptx> [--pages 1,6,33] [--out inv.json]` | 机器可读的形状清单，写 spec 前查 shape_id 用 |
| `catalog <pptx> [--out catalog.md]` | 人可读的页面目录，选页用 |
| `fill <spec.json> [--out 成品.pptx]` | 按 spec 填字、换图表、改表格、裁页排序、写备注 |
| `verify <模板> <成品> [--json]` | 核对：新增形状必须 0、文字不溢出、没有模板原文残留 |
| `render <pptx> <outdir> [--w 1600] [--h 900]` | PowerPoint COM 导出逐页 PNG，**交付前必须逐页看过** |

标准流程：`catalog` 选页 → `inventory` 拿 shape_id → 写 spec → `fill` → `verify` →
`render` → 目视复核 → 有问题回改 spec 重跑。**verify 通过不等于好看，PNG 一定要看。**

## 三、spec.json 的写法

```json
{
  "template": "C:/Users/12806/Desktop/天津师范大学毕业论文汇报答辩通用ppt模板.pptx",
  "output":   "D:/out/成品.pptx",
  "keep":     [1, 6, 33],
  "header":   { "33": ["中文短名", "ENGLISH KICKER", "1"] },
  "widen_title": { "33": {"sid": 26, "size": 28, "text": "陈述句标题"} },
  "fill":     { "33": { "24": {"paras": ["第一段", "第二段"], "size": 18, "spacing": 1.5} } },
  "chart":    { "54": { "9": {"categories": ["2015","2024"], "series": {"期刊论文": [1.6, 2.3]}} } },
  "table":    { "57": { "20": {"rows": [["维度","A","B"], ["记录","1","2"]]} } },
  "drop":     { "37": [36, 38] },
  "place":    { "32": { "18": {"left": 2.32, "top": 2.14, "width": 8.65, "height": 3.23} } },
  "notes":    { "1": "讲稿……" }
}
```

- 顶层键的**页号是模板页号**（1–62），不是成品页号。`keep` 决定成品顺序与取舍。
- `fill` 的键是 `shape_id`，也可以写 `"ph:11"` 按**占位符 idx** 寻址。
- 数值单位一律**英寸**。
- spacing 是**行距倍数**（1.5 即 150%），不是百分数；写 1.5 而不是 150。

- 不需要的键整个省略即可。

### shape_id 还是 ph:11

模板大量页共用同一套 `shape_id`，但**占位符 idx 是版式语义、跨页稳定**：

| idx | 含义 | 典型页 |
|---|---|---|
| 10 | PICTURE 大图 | 封面、目录、过渡页 |
| 11 | 页眉中文短名 / 封面主标题 / 目录中文条 | 全部 |
| 12 | 页眉序号 / 封面副题 / 过渡页中文标题 | 全部 |
| 13 | 页眉英文 kicker / 目录英文条 / 过渡页英文副题 | 全部 |

所以**封面、目录、过渡页优先用 `ph:`**；内容页的正文卡片不是占位符，只能按
`shape_id` 填。注意目录页版式2（模板 P06）的 8 个占位符 idx 是乱的：
中文条是 11/16/18/14（从上到下），英文条是 13/17/19/15。查 catalog 确认。

### 排版三个硬坑

1. **居中标题框只有 3.33 英寸宽**（如 P33 的 `26 文本框25` / P32 的 `21`）。
   陈述句标题必然折行。用 `widen_title` 以版面中心 **6.665** 为轴扩到 10 英寸：
   `{"sid": 26, "left": 1.665, "width": 10.0}`。**这是全程唯一该改几何的地方。**
2. **`widen_title` 不会自动填字**。模板原文是「在此输入您的标题」，只改框不改字，
   页面上就还是那句占位话。必须同时给 `text`。
3. **图表右侧的三组说明、饼图右侧三条 NO.x、表格下方通栏**都是独立文本框，
   各有自己的 shape_id。换完图表/表格**必须把它们也填掉**，否则整页半新半旧。

## 四、图与表

`chart` 用 `chart.replace_data()` 换数据，系列数必须与模板一致：

- P54 折线图 = **恰好 3 个系列**，右侧三组说明一一配对，不靠图例
- P55 饼图 = 1 个系列
- P52/P53 = 柱状/条形簇状
- P57 表格 = **5 行 × 5 列**，模板表格不能增行增列

**表格必须整表覆写。** `set_table` 会把没用到的行/列清空——因为模板 5×5 里装着
「类别5 / 85 / 21 / 875」这类示例整数，只写前 4 列的话第 5 列会静默保留假数据，
成品看起来跟真的一样。这是本工作流里最贵的一类错误。

## 五、内容要求（用户明确，2026-10）

- 只讲**文献计量学的最新发展**，不要复述教材六条定律、不要讲板书三问
- **禁金句，用学术语言**。不要「六条定律，一条都没有被推翻」这种句子
- 要有**顶刊文献**与**政策进展**：期刊论文 DOI 经 Crossref
  `https://api.crossref.org/works/<DOI>` 直读核验，查不到的一律剔除
- 政策数字标观测日期；「拟入选公示」不等于「正式立项」，别写错
- 章节骨架（已验证好用）：01 数据基础设施 / 02 方法与技术 / 03 理论与指标 /
  04 评价与治理

## 六、交付前自查

```
python -X utf8 <skill>/scripts/tnu.py verify <模板> <成品>
python -X utf8 <skill>/scripts/tnu.py render <成品> <png目录>
```

`verify` 的三类问题含义：

- **新增形状 N 个** — 用了 `add_textbox` / `add_shape`。这是硬错误，必然难看，回退。
  （删形状不算错；删图片占位符给正文腾地方是正常做法。）
- **OVERFLOW / OUT-OF-CANVAS** — 启发式估算，已按「相对模板同形状」校准，所以模板
  自带的越界（内容页右上 `图片 13` 右缘到 13.49 英寸）不会误报。报了就渲染核对。
- **模板原文残留** — 那格忘了填。判定条件是「文字与模板完全相同」，所以特意保留
  模板某句不会误报。

最后逐页看 PNG。看什么：标题是否换行、正文是否空心或溢出色块、图表系列是否对、
表格有没有残留、页脚页码是否连贯。

## 七、参考文件

- `references/catalog.md` — 62 页全部形状清单（自动生成）
- `scripts/tnu.py` — 引擎，约 700 行，含全部函数与注释
- `scripts/render.ps1` — COM 渲染。**必须保持纯 ASCII 无 BOM**（见文件内注释：
  PowerShell 5.1 会把无 BOM 文件按 GBK 读，中文注释被吃掉后连带注释掉下一行代码）
- `references/example_spec.json` — 三页可跑通的完整样例（封面+目录+内容页，
  verify 零问题），拿它当模板改最快
