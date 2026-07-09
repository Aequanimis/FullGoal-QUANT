# ETF Manager Ecosystem Dashboard

这是一个静态 dashboard 页面，用于展示“中美 ETF TOP10 管理人五维生态评分”。项目不依赖前端框架，也不依赖 CDN，直接双击 `index.html` 即可打开。

## 如何打开

1. 进入项目目录：`C:\Users\lvdon\Desktop\Fullgoal\课题研究\FullGoal-QUANT\7.9 dashboard`
2. 双击 `index.html`
3. 默认进入“中国市场”，默认选中“富国基金”

如果使用本地服务打开，页面会优先读取 `data/etf_scores.json`。如果直接双击打开时浏览器阻止读取本地 JSON，页面会自动使用 `app.js` 内置的同版本 fallback 数据，保证 dashboard 可正常展示和交互。

## 展示模式

页面顶部提供两个模式：

- `画像模式`：默认模式，展示单一管理人的五维雷达图、维度排名、品牌 logo、生态评分和单家公司 insight。
- `对比模式`：在当前市场内选择两家管理人，展示双管理人叠加雷达图、A/B 维度排名、A/B logo 对比卡、五维差异卡和对比 insight。

默认状态：

- 首次打开：`画像模式`、`中国市场`、`富国基金`
- 中国市场对比模式：A = `富国基金`，B = `华夏基金`
- 美国市场对比模式：A = `BlackRock/iShares`，B = `Invesco`

对比模式只做同市场内部对比，不做中美跨市场对比。切换市场时，系统会清空旧市场 A/B 选择，并自动使用新市场默认组合。

## A/B 选择逻辑

在 `对比模式` 下，左侧 TOP10 列表仍按规模排名和市场份额展示，但可选择 A/B 两家管理人：

- 第一个选中的管理人为 `A`，使用金色标识。
- 第二个选中的管理人为 `B`，使用霓虹蓝标识。
- 再次点击已选管理人会取消选择。
- 如果已经选满 A/B，再点击第三家公司，会自动替换 `B`，`A` 保持不变。

左侧排序始终按 `scaleRank` 从小到大，生态评分不参与排序。

## 对比计算口径

五维维度排名基于当前市场 TOP10 内部计算：

- 每个维度分别按分数从高到低排名。
- 同分并列，显示相同名次。
- 第 1 名显示 `BEST`，第 2 名显示 `2nd`，第 3 名显示 `3rd`，第 4 名及以后显示 `4th`、`5th` 等。

对比模式下的五维差异卡计算 `A - B`：

- `A +x.x` 表示 A 在该维度领先。
- `B +x.x` 表示 B 在该维度领先。
- 生态评分只作为结构参考展示，不参与五维维度排名，也不参与左侧排序。

## 数据文件

评分数据位于：

```text
data/etf_scores.json
```

数据按市场分为：

- `markets.china.managers`
- `markets.us.managers`

每个管理人包含：

- `rank`：规模排名，页面标签按这个字段排序
- `name`：管理人名称
- `slug`：logo 文件名使用的英文短名
- `scores`：五维评分
- `ecosystemScore`：生态评分

## 如何修改评分数据

修改 `data/etf_scores.json` 中对应管理人的分数即可。建议保持五维字段名不变：

```text
产品广度
费率优势
交易活跃
多元能力
旗舰产品
```

如果需要双击 `index.html` 后也同步使用最新 JSON，请用本地静态服务打开，例如在该目录运行：

```powershell
python -m http.server 8000
```

然后访问 `http://localhost:8000/`。如果本机 `python` 不可用，也可以继续双击打开，但此时可能使用 `app.js` 内置 fallback 数据。

## Logo 资源说明

当前 dashboard 使用本地透明背景 SVG logo，用于右侧品牌 hero 区域优先展示。原始 logo 包来源路径：

```text
C:\Users\lvdon\Desktop\Fullgoal\ICI课题\课题更正\最终更改\logo\US_CN_20_logos_transparent_svg_pack
```

项目内复制路径：

```text
assets/logos/svg
```

当前命名沿用原始 SVG 包：

- 中国市场：`CN_01_华夏基金.svg` 至 `CN_10_富国基金.svg`
- 美国市场：`US_01_Invesco.svg`、`US_02_BlackRock_iShares.svg`、`US_03_StateStreet_SPDR.svg`、`US_04_VanEck.svg`、`US_05_CharlesSchwab.svg`、`US_06_Vanguard.svg`、`US_07_JPMorganAM.svg`、`US_08_Dimensional.svg`、`US_09_Fidelity.svg`、`US_10_FirstTrust.svg`

每个管理人在 `data/etf_scores.json` 中都有 `logo` 字段，例如：

```json
{
  "name": "富国基金",
  "slug": "fullgoal",
  "logo": "assets/logos/svg/CN_10_富国基金.svg"
}
```

如果要替换某一家 logo：

1. 把新的透明背景 SVG 或 PNG 放入 `assets/logos/svg`，例如：

```text
assets/logos/svg/CN_10_富国基金_new.svg
```

2. 修改 `data/etf_scores.json` 中该管理人的 `logo` 路径，例如：

```json
"logo": "assets/logos/svg/CN_10_富国基金_new.svg"
```

3. 优先使用透明背景 SVG，保持原始宽高比，不要转成低清 PNG。

前端会按以下顺序处理：

1. 优先加载 `data/etf_scores.json` 中的 `logo`。
2. 如果 JSON 没有该字段，则使用 `app.js` 中的本地 logo 映射。
3. 如果图片加载失败或路径为空，则自动 fallback 为原来的大号文字品牌展示，不会报错。

## 页面说明

页面底部已展示口径说明：

“五维评分为市场内 TOP10 标准化结果，生态评分仅供结构参考，不代表中美管理人绝对实力横向可比。”

生态评分仅作为结构参考展示，页面视觉中心为五维雷达图和当前管理人的品牌生态画像。
