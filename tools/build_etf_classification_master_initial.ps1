param(
  [string]$InputPath = "C:\Users\lvdon\Desktop\Fullgoal\ICI课题\P17-P23_中国ETF数据处理_二次修正版\00_数据字段检查与描述性统计_二次修正版.xlsx",
  [string]$InputSheet = "ETF分析池_上市交易_二次修正",
  [string]$OutputDir = "C:\Users\lvdon\Desktop\Fullgoal\ICI课题\ETF分类主表_初版",
  [string]$OutputPath = "C:\Users\lvdon\Desktop\Fullgoal\ICI课题\ETF分类主表_初版\ETF分类主表_初版.xlsx"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

function To-Text($value) {
  if ($null -eq $value) { return "" }
  return ([string]$value).Trim()
}

function To-Number($value) {
  if ($null -eq $value) { return 0.0 }
  if ($value -is [double] -or $value -is [int] -or $value -is [decimal]) { return [double]$value }
  $s = (To-Text $value).Replace(",", "")
  if ($s -eq "") { return 0.0 }
  $d = 0.0
  if ([double]::TryParse($s, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
  if ([double]::TryParse($s, [ref]$d)) { return $d }
  return 0.0
}

function Get-CellValue($row, [string]$name) {
  if ($null -eq $row) { return $null }
  if ($row -is [System.Collections.IDictionary]) {
    if ($row.Contains($name)) { return $row[$name] }
    return $null
  }
  $prop = $row.PSObject.Properties[$name]
  if ($null -ne $prop) { return $prop.Value }
  return $null
}

function Set-CellValue($row, [string]$name, $value) {
  if ($row -is [System.Collections.IDictionary] -and $row.Contains($name)) {
    $row[$name] = $value
  } elseif ($row -is [System.Collections.IDictionary]) {
    $row.Add($name, $value)
  } else {
    $row | Add-Member -NotePropertyName $name -NotePropertyValue $value -Force
  }
}

function Convert-ExcelValue($header, $value) {
  if ($null -eq $value) { return "" }
  if (@("上市日期", "成立日期") -contains $header) {
    if ($value -is [double] -or $value -is [int]) {
      try { return [DateTime]::FromOADate([double]$value).ToString("yyyy-MM-dd") } catch { return To-Text $value }
    }
    return To-Text $value
  }
  if (@("基金规模_亿元", "管理费率", "托管费率") -contains $header) {
    if ((To-Text $value) -eq "") { return "" }
    return [double](To-Number $value)
  }
  return To-Text $value
}

function Contains-Any([string]$text, [string[]]$keywords) {
  foreach ($kw in $keywords) {
    if ($kw -and $text.IndexOf($kw, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
  }
  return $false
}

function Match-Keywords([string]$text, [string[]]$keywords) {
  $hits = New-Object System.Collections.ArrayList
  foreach ($kw in $keywords) {
    if ($kw -and $text.IndexOf($kw, [StringComparison]::OrdinalIgnoreCase) -ge 0) { [void]$hits.Add($kw) }
  }
  return @($hits | Select-Object -Unique)
}

function Join-Keywords($items) {
  return ((@($items) | Where-Object { $_ -and (To-Text $_) -ne "" } | Select-Object -Unique) -join "、")
}

function Add-Tag($tags, [string]$tag) {
  if ($tag -and -not $tags.Contains($tag)) { [void]$tags.Add($tag) }
}

function Combined-Text($row) {
  $parts = @(
    (Get-CellValue $row "证券简称"),
    (Get-CellValue $row "基金简称"),
    (Get-CellValue $row "基金全称"),
    (Get-CellValue $row "跟踪指数"),
    (Get-CellValue $row "业绩比较基准"),
    (Get-CellValue $row "投资类型")
  )
  return (($parts | ForEach-Object { To-Text $_ }) -join " ")
}

function Read-InputRows($excel, [string]$path, [string]$sheetName) {
  if (-not (Test-Path -LiteralPath $path)) { throw "未找到二次修正版主分析池：ETF分析池_上市交易_二次修正。请检查输入文件路径。" }
  $wb = $excel.Workbooks.Open($path, $null, $true)
  try {
    $ws = $null
    foreach ($s in $wb.Worksheets) {
      if ($s.Name -eq $sheetName) { $ws = $s; break }
    }
    if ($null -eq $ws) { throw "未找到二次修正版主分析池：ETF分析池_上市交易_二次修正。请检查输入文件路径。" }
    $used = $ws.UsedRange
    $values = $used.Value2
    $rowCount = $used.Rows.Count
    $colCount = $used.Columns.Count
    $headers = New-Object System.Collections.ArrayList
    for ($c = 1; $c -le $colCount; $c++) { [void]$headers.Add((To-Text $values[1, $c])) }
    $rows = New-Object System.Collections.ArrayList
    for ($r = 2; $r -le $rowCount; $r++) {
      $row = [ordered]@{}
      for ($c = 1; $c -le $colCount; $c++) {
        $h = [string]$headers[$c - 1]
        $row[$h] = Convert-ExcelValue $h $values[$r, $c]
      }
      [void]$rows.Add($row)
    }
    return [pscustomobject]@{ Headers = @($headers); Rows = $rows }
  } finally {
    $wb.Close($false)
  }
}

function Sum-Scale($rows) {
  $sum = 0.0
  foreach ($row in @($rows)) { $sum += To-Number (Get-CellValue $row "基金规模_亿元") }
  return [Math]::Round($sum, 4)
}

function Median($numbers) {
  $arr = @($numbers | Sort-Object)
  if ($arr.Count -eq 0) { return 0 }
  $mid = [int][Math]::Floor($arr.Count / 2)
  if ($arr.Count % 2 -eq 1) { return [Math]::Round([double]$arr[$mid], 4) }
  return [Math]::Round((([double]$arr[$mid - 1] + [double]$arr[$mid]) / 2), 4)
}

function Classify-ETF($row) {
  $text = Combined-Text $row
  $investment = To-Text (Get-CellValue $row "投资类型")
  $indexName = To-Text (Get-CellValue $row "跟踪指数")
  $tags = New-Object System.Collections.ArrayList
  $reviewReasons = New-Object System.Collections.ArrayList
  $basis = ""
  $confidence = "高"

  $moneyKeywords = @("货币", "现金", "保证金", "快线", "添益", "场内货币", "收益宝", "财富宝", "银华日利", "华宝添益", "交易货币")
  $commodityKeywords = @("黄金ETF", "黄金基金", "豆粕ETF", "商品ETF", "能源化工", "原油ETF", "白银ETF")
  $bondKeywords = @("国债", "政金债", "信用债", "公司债", "地方债", "城投债", "可转债", "短融", "债券", "债")
  $crossKeywords = @("恒生", "港股", "港股通", "H股", "纳斯达克", "标普", "道琼斯", "日经", "德国", "法国", "海外", "中概", "东南亚", "QDII", "MSCI美国", "美国", "香港", "日韩", "越南", "沙特", "亚太", "全球")

  $asset = "股票ETF"
  $assetHits = @()
  if ($investment.Contains("货币市场型基金") -or $text.Contains("交易型货币市场基金") -or (Contains-Any $text $moneyKeywords)) {
    $asset = "货币ETF"; $assetHits = @(Match-Keywords $text $moneyKeywords); Add-Tag $tags "货币"; Add-Tag $tags "场内货币"; $basis = "命中货币ETF规则"
  } elseif ($investment.Contains("商品型基金") -or (Contains-Any $text $commodityKeywords)) {
    $asset = "商品ETF"; $assetHits = @(Match-Keywords $text $commodityKeywords); Add-Tag $tags "商品"; $basis = "命中商品ETF规则"
  } elseif ($investment.Contains("被动指数型债券基金") -or (Contains-Any $text $bondKeywords)) {
    $asset = "债券ETF"; $assetHits = @(Match-Keywords $text $bondKeywords); Add-Tag $tags "债券"; $basis = "命中债券ETF规则"
  } elseif ($investment.Contains("国际(QDII)股票型基金") -or (Contains-Any $text $crossKeywords)) {
    $asset = "跨境ETF"; $assetHits = @(Match-Keywords $text $crossKeywords); Add-Tag $tags "跨境"; $basis = "命中跨境ETF规则"
  } elseif ($investment -eq "" -and $indexName -eq "") {
    $asset = "其他ETF"; $confidence = "低"; [void]$reviewReasons.Add("投资类型和跟踪指数缺失，无法确认资产类型"); $basis = "无法判断资产类型"
  }

  $auto = $asset
  $isEnhanced = "否"; $isSmart = "否"; $isBroad = "否"; $isTheme = "否"; $isStrategy = "否"
  $isCross = if ($asset -eq "跨境ETF") { "是" } else { "否" }
  $isBond = if ($asset -eq "债券ETF") { "是" } else { "否" }
  $isCommodity = if ($asset -eq "商品ETF") { "是" } else { "否" }
  $isMoney = if ($asset -eq "货币ETF") { "是" } else { "否" }
  $isMulti = if (Contains-Any $text @("多资产", "资产配置")) { "是" } else { "否" }

  $keywords = New-Object System.Collections.ArrayList
  foreach ($h in $assetHits) { [void]$keywords.Add($h) }

  if ($asset -eq "股票ETF") {
    $enhanceKeywords = @("增强策略", "指数增强", "增强")
    $strategyKeywords = @("红利", "低波", "低波动", "红利低波", "质量", "价值", "成长", "质量成长", "现金流", "自由现金流", "股息", "高股息", "基本面", "等权", "等权重", "ESG", "动量", "回购", "分红", "央企红利", "国企红利", "红利质量", "红利价值", "价值100", "成长100", "低估值", "高分红", "Smart Beta", "价值稳健", "成长创新", "质量低波", "红利低波动", "股东回报", "龙头红利", "优选", "精选", "高质量", "盈利", "盈利质量", "央企股东回报", "国企股东回报", "央企分红", "国企分红", "央企价值", "国企价值")
    $themeKeywords = @("新能源", "光伏", "芯片", "半导体", "人工智能", "机器人", "军工", "医药", "消费", "金融", "证券", "银行", "地产", "房地产", "传媒", "计算机", "通信", "汽车", "电池", "储能", "双碳", "绿色", "电力", "煤炭", "有色", "钢铁", "基建", "农业", "酒", "食品", "云计算", "软件", "游戏", "数字经济", "高端制造", "创新药", "医疗器械", "工业母机", "机床", "工程机械", "畜牧养殖", "稀有金属", "稀土", "稀土产业", "工业有色", "化工", "材料", "环保", "教育", "物流", "旅游", "黄金股", "矿业", "交运", "信息技术", "科技", "互联网", "生物", "医疗", "养老", "家电", "机械", "电子", "保险", "石油天然气", "油气产业", "卫星产业", "航天", "航空", "通用航空", "低空经济", "碳中和", "低碳经济", "虚拟现实", "元宇宙", "粮食产业", "中药", "央企创新", "央企创新驱动", "国企改革", "国新央企", "长江保护", "长江经济带", "长三角", "粤港澳", "一带一路", "央企科技", "央企现代能源", "央企结构调整", "央企共赢", "专精特新", "新材料", "新经济", "新消费", "新能车", "智能车", "智能驾驶", "物联网", "大数据", "信创")
    $broadKeywords = @("沪深300", "中证500", "中证800", "中证1000", "中证2000", "中证A500", "中证A100", "中证A50", "中证A股", "上证50", "上证180", "上证380", "上证综指", "上证指数", "深证100", "深证50", "创业板指", "创业板50", "创业板200", "创业板综", "科创50", "科创100", "科创200", "科创板50", "科创板100", "科创板200", "科创板综合", "科创综指", "科创创业50", "双创50", "双创", "北证50", "MSCI中国A股", "MSCI中国A50互联互通", "MSCI A50", "富时中国A50", "国证A指", "万得全A", "中小板", "中小100")
    $genericKeywords = @("核心", "创新", "优势", "产业", "主题", "信息", "龙头", "成长", "精选", "优选")
    $commodityStockConflict = @("黄金股", "有色金属", "有色", "稀有金属", "稀土", "矿业", "煤炭", "钢铁", "石油天然气", "油气", "油气产业")

    $enhanceHits = @(Match-Keywords $text $enhanceKeywords)
    $strategyHits = @(Match-Keywords $text $strategyKeywords)
    $themeHits = @(Match-Keywords $text $themeKeywords)
    $broadHits = @(Match-Keywords $text $broadKeywords)
    $genericHits = @(Match-Keywords $text $genericKeywords)
    foreach ($h in ($enhanceHits + $strategyHits + $themeHits + $broadHits + $genericHits)) { [void]$keywords.Add($h) }

    if ($enhanceHits.Count -gt 0 -and $broadHits.Count -gt 0) {
      $auto = "宽基ETF"; $isEnhanced = "是"; $isStrategy = "是"; $isBroad = "是"; $confidence = "中"
      Add-Tag $tags "宽基"; Add-Tag $tags "指数增强"
      $basis = "宽基增强/指数增强，主分类按宽基处理"
    } elseif ($strategyHits.Count -gt 0) {
      $auto = "策略ETF"; $isStrategy = "是"; $isSmart = "是"
      Add-Tag $tags "Smart Beta"
      foreach ($h in $strategyHits | Select-Object -First 2) { Add-Tag $tags $h }
      $basis = "命中明确策略因子关键词"
      if ($themeHits.Count -gt 0 -or $broadHits.Count -gt 0) { $confidence = "中" }
    } elseif ($themeHits.Count -gt 0) {
      $auto = "行业主题ETF"; $isTheme = "是"
      Add-Tag $tags "行业主题"
      foreach ($h in $themeHits | Select-Object -First 2) { Add-Tag $tags $h }
      $basis = "命中明确行业/主题关键词"
      if ($broadHits.Count -gt 0) { $confidence = "中" }
    } elseif ($broadHits.Count -gt 0) {
      $auto = "宽基ETF"; $isBroad = "是"
      Add-Tag $tags "宽基"
      $basis = "命中纯宽基或板块宽基关键词"
    } else {
      $auto = "其他ETF"; $confidence = "低"; Add-Tag $tags "待复核"; $basis = "未命中明确宽基、策略或行业主题规则"
      [void]$reviewReasons.Add("自动分类为其他ETF")
    }

    if ($enhanceHits.Count -gt 0) { $isEnhanced = "是" }
    $hitKinds = 0
    if ($broadHits.Count -gt 0) { $hitKinds++ }
    if ($strategyHits.Count -gt 0) { $hitKinds++ }
    if ($themeHits.Count -gt 0) { $hitKinds++ }
    if ($hitKinds -ge 3) { $confidence = "低"; [void]$reviewReasons.Add("同时命中宽基、策略、行业主题关键词") }
    if ($genericHits.Count -gt 0 -and $broadHits.Count -eq 0 -and $strategyHits.Count -eq 0 -and $themeHits.Count -eq 0) {
      $confidence = "低"; [void]$reviewReasons.Add("只命中泛词，缺少明确分类依据")
    }
    if ((Contains-Any $text @("央企", "国企")) -and -not (Contains-Any $text @("央企红利", "国企红利", "央企股东回报", "国企股东回报", "央企分红", "国企分红", "央企价值", "国企价值", "央企创新", "央企创新驱动", "央企科技", "央企现代能源", "央企结构调整", "国企改革", "国新央企", "央企共赢", "央企主题"))) {
      $confidence = "低"; [void]$reviewReasons.Add("央企/国企含义不清")
    }
    if ($indexName -eq "") { $confidence = "低"; [void]$reviewReasons.Add("跟踪指数缺失") }
    if (Contains-Any $text $commodityStockConflict) { [void]$reviewReasons.Add("股票ETF含商品/资源相关行业词，建议人工确认") }
  } else {
    if ($asset -eq "跨境ETF") { Add-Tag $tags "跨境" }
    if ($asset -eq "债券ETF") { Add-Tag $tags "债券" }
    if ($asset -eq "商品ETF") { Add-Tag $tags "商品" }
    if ($asset -eq "货币ETF") { Add-Tag $tags "货币" }
    if ($asset -eq "其他ETF") { $auto = "其他ETF" }
  }

  if ($auto -eq "其他ETF") { $confidence = "低" }
  $needReview = if ($reviewReasons.Count -gt 0 -or $confidence -eq "低" -or $auto -eq "其他ETF") { "是" } else { "否" }
  if ($reviewReasons.Count -eq 0 -and $needReview -eq "是") { [void]$reviewReasons.Add("低置信度分类") }

  return [pscustomobject]@{
    Asset = $asset; Auto = $auto; Tags = (($tags | Select-Object -Unique) -join ";")
    IsEnhanced = $isEnhanced; IsSmart = $isSmart; IsBroad = $isBroad; IsTheme = $isTheme; IsStrategy = $isStrategy
    IsCross = $isCross; IsBond = $isBond; IsCommodity = $isCommodity; IsMoney = $isMoney; IsMulti = $isMulti
    Confidence = $confidence; Basis = $basis; Keywords = (Join-Keywords $keywords)
    NeedReview = $needReview; ReviewReason = (($reviewReasons | Select-Object -Unique) -join "；")
  }
}

function New-SummaryRows($rows, [string]$field) {
  $totalCount = @($rows).Count
  $totalScale = Sum-Scale $rows
  $groups = @($rows | Group-Object { To-Text (Get-CellValue $_ $field) })
  $out = New-Object System.Collections.ArrayList
  foreach ($g in $groups) {
    $items = @($g.Group)
    $scale = Sum-Scale $items
    $avg = if ($items.Count -gt 0) { [Math]::Round($scale / $items.Count, 4) } else { 0 }
    $med = Median (@($items | ForEach-Object { To-Number (Get-CellValue $_ "基金规模_亿元") }))
    $row = [ordered]@{}
    $row[$field] = $g.Name
    $row["ETF数量"] = $items.Count
    $row["数量占比"] = if ($totalCount -gt 0) { [Math]::Round($items.Count / $totalCount, 6) } else { 0 }
    $row["总规模_亿元"] = $scale
    $row["规模占比"] = if ($totalScale -gt 0) { [Math]::Round($scale / $totalScale, 6) } else { 0 }
    $row["平均规模_亿元"] = $avg
    $row["规模中位数_亿元"] = $med
    [void]$out.Add($row)
  }
  return @($out | Sort-Object -Property "总规模_亿元" -Descending)
}

function New-LabelStats($rows) {
  $map = @{}
  foreach ($row in $rows) {
    $labels = (To-Text (Get-CellValue $row "ETF标签")).Split(";") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    foreach ($label in $labels) {
      if (-not $map.ContainsKey($label)) { $map[$label] = New-Object System.Collections.ArrayList }
      [void]$map[$label].Add($row)
    }
  }
  $out = New-Object System.Collections.ArrayList
  foreach ($label in ($map.Keys | Sort-Object)) {
    $items = @($map[$label])
    $examples = (($items | Select-Object -First 5 | ForEach-Object { To-Text (Get-CellValue $_ "基金简称") }) -join "；")
    [void]$out.Add([ordered]@{
      "标签" = $label
      "产品数量" = $items.Count
      "总规模_亿元" = (Sum-Scale $items)
      "代表产品示例" = $examples
    })
  }
  return @($out | Sort-Object -Property "总规模_亿元" -Descending)
}

function Xml-Escape($value) {
  $s = To-Text $value
  $s = [regex]::Replace($s, "[\x00-\x08\x0B\x0C\x0E-\x1F]", "")
  return [System.Security.SecurityElement]::Escape($s)
}

function Col-Letter([int]$index) {
  $n = $index; $letters = ""
  while ($n -gt 0) {
    $rem = [int](($n - 1) % 26)
    $letters = ([char]([int](65 + $rem))) + $letters
    $n = [Math]::Floor(($n - 1) / 26)
  }
  return $letters
}

function Write-CellXml($writer, [int]$rowIndex, [int]$colIndex, $value, [int]$styleId) {
  $ref = (Col-Letter $colIndex) + [string]$rowIndex
  if ($null -eq $value) { $value = "" }
  if (($value -is [double] -or $value -is [int] -or $value -is [decimal]) -and -not [double]::IsNaN([double]$value)) {
    $num = ([double]$value).ToString("0.############", [Globalization.CultureInfo]::InvariantCulture)
    if ($styleId -gt 0) { $writer.Write("<c r=""$ref"" s=""$styleId""><v>$num</v></c>") } else { $writer.Write("<c r=""$ref""><v>$num</v></c>") }
  } else {
    $text = Xml-Escape $value
    if ($styleId -gt 0) { $writer.Write("<c r=""$ref"" s=""$styleId"" t=""inlineStr""><is><t>$text</t></is></c>") } else { $writer.Write("<c r=""$ref"" t=""inlineStr""><is><t>$text</t></is></c>") }
  }
}

function Write-WorksheetXml([string]$path, $rows, [string[]]$headers) {
  $dataRows = @($rows)
  if ($null -eq $headers -or $headers.Count -eq 0) {
    if ($dataRows.Count -gt 0 -and $dataRows[0] -is [System.Collections.IDictionary]) { $headers = @($dataRows[0].Keys) } else { $headers = @("说明") }
  }
  $rowCount = $dataRows.Count + 1
  $colCount = $headers.Count
  $lastRef = (Col-Letter $colCount) + [string]$rowCount
  $writer = New-Object System.IO.StreamWriter($path, $false, (New-Object System.Text.UTF8Encoding($false)))
  try {
    $writer.Write("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">")
    $writer.Write("<sheetViews><sheetView workbookViewId=""0""><pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/></sheetView></sheetViews>")
    $writer.Write("<sheetFormatPr defaultColWidth=""15"" defaultRowHeight=""15""/>")
    $writer.Write("<sheetData><row r=""1"">")
    for ($c = 1; $c -le $colCount; $c++) { Write-CellXml $writer 1 $c $headers[$c - 1] 1 }
    $writer.Write("</row>")
    for ($r = 0; $r -lt $dataRows.Count; $r++) {
      $excelRow = $r + 2
      $writer.Write("<row r=""$excelRow"">")
      $row = $dataRows[$r]
      for ($c = 1; $c -le $colCount; $c++) { Write-CellXml $writer $excelRow $c (Get-CellValue $row $headers[$c - 1]) 0 }
      $writer.Write("</row>")
    }
    $writer.Write("</sheetData><autoFilter ref=""A1:$lastRef""/></worksheet>")
  } finally { $writer.Close() }
}

function Export-SimpleXlsx([string]$path, $sheetDefs) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
  $tempRoot = Join-Path $env:TEMP ("etf_master_" + [Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $tempRoot | Out-Null
  try {
    foreach ($d in @("_rels","docProps","xl","xl\_rels","xl\worksheets")) { New-Item -ItemType Directory -Path (Join-Path $tempRoot $d) | Out-Null }
    $sheets = @($sheetDefs)
    for ($i = 0; $i -lt $sheets.Count; $i++) { Write-WorksheetXml (Join-Path $tempRoot ("xl\worksheets\sheet" + ($i + 1) + ".xml")) $sheets[$i].Rows $sheets[$i].Headers }
    $overrides = ""; for ($i = 1; $i -le $sheets.Count; $i++) { $overrides += "<Override PartName=""/xl/worksheets/sheet$i.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>" }
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "[Content_Types].xml"), "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/><Default Extension=""xml"" ContentType=""application/xml""/><Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/><Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/><Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/><Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml""/>$overrides</Types>", (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "_rels\.rels"), "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/><Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml""/></Relationships>", (New-Object System.Text.UTF8Encoding($false)))
    $sheetXml = ""; $relsXml = ""
    for ($i = 0; $i -lt $sheets.Count; $i++) {
      $idx = $i + 1; $name = Xml-Escape $sheets[$i].Name
      $sheetXml += "<sheet name=""$name"" sheetId=""$idx"" r:id=""rId$idx""/>"
      $relsXml += "<Relationship Id=""rId$idx"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$idx.xml""/>"
    }
    $styleRel = $sheets.Count + 1
    $relsXml += "<Relationship Id=""rId$styleRel"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>"
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "xl\workbook.xml"), "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheets>$sheetXml</sheets></workbook>", (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "xl\_rels\workbook.xml.rels"), "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">$relsXml</Relationships>", (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "xl\styles.xml"), "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><fonts count=""2""><font><sz val=""11""/><name val=""Calibri""/></font><font><b/><color rgb=""FFFFFFFF""/><sz val=""11""/><name val=""Calibri""/></font></fonts><fills count=""3""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FF366092""/></patternFill></fill></fills><borders count=""1""><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs><cellXfs count=""2""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""2"" borderId=""0"" xfId=""0"" applyFont=""1"" applyFill=""1""/></cellXfs><cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles></styleSheet>", (New-Object System.Text.UTF8Encoding($false)))
    $now = [DateTime]::UtcNow.ToString("s") + "Z"
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "docProps\core.xml"), "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy><dcterms:created xsi:type=""dcterms:W3CDTF"">$now</dcterms:created><dcterms:modified xsi:type=""dcterms:W3CDTF"">$now</dcterms:modified></cp:coreProperties>", (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "docProps\app.xml"), "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes""><Application>Microsoft Excel</Application></Properties>", (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $path)
  } finally {
    if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
  }
}

function New-SheetDef([string]$name, $rows, [string[]]$headers) {
  return [pscustomobject]@{ Name = $name; Rows = @($rows); Headers = $headers }
}

if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $input = Read-InputRows $excel $InputPath $InputSheet
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

$sourceRows = $input.Rows
$sourceHeaders = [string[]]$input.Headers
$requiredOriginal = @("Wind代码","证券简称","基金简称","基金全称","基金代码","交易代码","基金管理人","管理人简称","基金类型","投资类型","业绩比较基准","基金规模_亿元","成立日期","上市日期","基金上市地点","跟踪指数","跟踪指数代码","基金托管人","管理费率","托管费率")
$newFields = @("ETF资产类型","ETF分类_自动","ETF标签","是否增强策略","是否SmartBeta策略","是否宽基ETF","是否行业主题ETF","是否策略ETF","是否跨境ETF","是否债券ETF","是否商品ETF","是否货币ETF","是否多资产ETF","分类置信度","分类依据","命中关键词","是否需要人工复核","复核原因","人工修正分类","最终ETF分类","最终分类来源")
$missingOriginal = @($requiredOriginal | Where-Object { $sourceHeaders -notcontains $_ })

$master = New-Object System.Collections.ArrayList
foreach ($src in $sourceRows) {
  $row = [ordered]@{}
  foreach ($h in $requiredOriginal) { $row[$h] = Get-CellValue $src $h }
  $cls = Classify-ETF $src
  $row["ETF资产类型"] = $cls.Asset
  $row["ETF分类_自动"] = $cls.Auto
  $row["ETF标签"] = $cls.Tags
  $row["是否增强策略"] = $cls.IsEnhanced
  $row["是否SmartBeta策略"] = $cls.IsSmart
  $row["是否宽基ETF"] = $cls.IsBroad
  $row["是否行业主题ETF"] = $cls.IsTheme
  $row["是否策略ETF"] = $cls.IsStrategy
  $row["是否跨境ETF"] = $cls.IsCross
  $row["是否债券ETF"] = $cls.IsBond
  $row["是否商品ETF"] = $cls.IsCommodity
  $row["是否货币ETF"] = $cls.IsMoney
  $row["是否多资产ETF"] = $cls.IsMulti
  $row["分类置信度"] = $cls.Confidence
  $row["分类依据"] = $cls.Basis
  $row["命中关键词"] = $cls.Keywords
  $row["是否需要人工复核"] = $cls.NeedReview
  $row["复核原因"] = $cls.ReviewReason
  $row["人工修正分类"] = ""
  $row["最终ETF分类"] = $cls.Auto
  $row["最终分类来源"] = "自动分类"
  [void]$master.Add($row)
}

$totalCount = @($master).Count
$totalScale = Sum-Scale $master
$quality = New-Object System.Collections.ArrayList
function Add-Check([string]$item, [bool]$ok, $value, [string]$note) {
  [void]$quality.Add([ordered]@{ "检查项" = $item; "结果" = if ($ok) { "通过" } else { "失败" }; "数值" = $value; "说明" = $note })
}
$ofCount = @($master | Where-Object { (To-Text (Get-CellValue $_ "Wind代码")).EndsWith(".OF") }).Count
$missingStart = @($master | Where-Object { (To-Text (Get-CellValue $_ "成立日期")) -eq "" }).Count
$missingList = @($master | Where-Object { (To-Text (Get-CellValue $_ "上市日期")) -eq "" }).Count
$badScale = @($master | Where-Object { (To-Number (Get-CellValue $_ "基金规模_亿元")) -le 0 }).Count
$emptyAuto = @($master | Where-Object { (To-Text (Get-CellValue $_ "ETF分类_自动")) -eq "" }).Count
$emptyFinal = @($master | Where-Object { (To-Text (Get-CellValue $_ "最终ETF分类")) -eq "" }).Count
$reviewRows = @($master | Where-Object { (To-Text (Get-CellValue $_ "是否需要人工复核")) -eq "是" })
$otherCount = @($master | Where-Object { (To-Text (Get-CellValue $_ "ETF分类_自动")) -eq "其他ETF" }).Count
Add-Check "主分析池数量是否为1574" ($totalCount -eq 1574) $totalCount "必须为1574"
Add-Check "总规模是否约为46949.9638亿元" ([Math]::Abs($totalScale - 46949.9638) -le 0.1) $totalScale "误差不超过0.1亿元"
Add-Check "是否存在.OF产品" ($ofCount -eq 0) $ofCount "必须为0"
Add-Check "成立日期是否有缺失" ($missingStart -eq 0) $missingStart "必须为0"
Add-Check "上市日期是否有缺失" ($missingList -eq 0) $missingList "必须为0"
Add-Check "基金规模是否有缺失或小于等于0" ($badScale -eq 0) $badScale "必须为0"
Add-Check "ETF分类_自动是否有空值" ($emptyAuto -eq 0) $emptyAuto "必须为0"
Add-Check "最终ETF分类是否有空值" ($emptyFinal -eq 0) $emptyFinal "必须为0"
Add-Check "待人工复核数量" $true $reviewRows.Count "提示项"
Add-Check "其他ETF数量" $true $otherCount "提示项"
foreach ($cat in @("宽基ETF","行业主题ETF","策略ETF","跨境ETF","债券ETF","商品ETF","货币ETF")) {
  $cnt = @($master | Where-Object { (To-Text (Get-CellValue $_ "ETF分类_自动")) -eq $cat }).Count
  Add-Check "$cat`数量" $true $cnt "提示项"
}
$fatal = @($quality | Where-Object { (Get-CellValue $_ "结果") -eq "失败" })
if ($fatal.Count -gt 0) {
  $msg = ($fatal | ForEach-Object { (Get-CellValue $_ "检查项") + "=" + (Get-CellValue $_ "数值") }) -join "; "
  throw "质量检查失败，已停止输出：$msg"
}

$autoSummary = New-SummaryRows $master "ETF分类_自动"
$assetSummary = New-SummaryRows $master "ETF资产类型"
$classificationStats = New-Object System.Collections.ArrayList
foreach ($r in $autoSummary) {
  $nr = [ordered]@{ "汇总维度" = "ETF分类_自动" }
  foreach ($k in $r.Keys) { $nr[$k] = $r[$k] }
  [void]$classificationStats.Add($nr)
}
foreach ($r in $assetSummary) {
  $nr = [ordered]@{ "汇总维度" = "ETF资产类型" }
  foreach ($k in $r.Keys) { $nr[$k] = $r[$k] }
  [void]$classificationStats.Add($nr)
}
$labelStats = New-LabelStats $master

$ruleRows = New-Object System.Collections.ArrayList
$missingOriginalText = if ($missingOriginal.Count -gt 0) { $missingOriginal -join "、" } else { "无" }
foreach ($pair in @(
  @("本次输入文件", $InputPath),
  @("本次输入sheet", $InputSheet),
  @("本次输出文件", $OutputPath),
  @("本次是否改变主分析池", "否，保留二次修正版1574只上市ETF主分析池"),
  @("主分析池数量", $totalCount),
  @("主分析池总规模", $totalScale),
  @("缺失原始字段", $missingOriginalText),
  @("ETF资产类型规则", "按货币、商品、债券、跨境、股票、其他的优先级互斥判断"),
  @("ETF分类_自动规则", "非股票大类直接对应；股票ETF再分宽基、行业主题、策略、其他ETF"),
  @("策略ETF判断规则", "识别红利、低波、质量、价值、成长、现金流、等权、ESG、股东回报等明确策略因子"),
  @("行业主题ETF判断规则", "识别行业、产业、区域、科技制造、能源材料、央企/国企主题等明确关键词；泛词不单独作为依据"),
  @("宽基ETF判断规则", "识别沪深300、中证500、中证1000、中证A500、上证50、创业板、科创、双创、北证等纯宽基或板块宽基"),
  @("央企/国企类判断规则", "央企/国企红利、价值、股东回报归策略；央企创新、科技、现代能源、国企改革等归行业主题；仅央企/国企需复核"),
  @("人工复核规则", "其他ETF、低置信度、多类关键词冲突、只命中泛词、央企国企含义不清、跟踪指数缺失、商品/资源相关股票主题均标记复核"),
  @("后续人工修正使用方式", "在人工修正分类中填写校正类别，后续可用人工修正分类覆盖最终ETF分类")
)) { [void]$ruleRows.Add([ordered]@{ "项目" = $pair[0]; "说明" = $pair[1] }) }

$sampleRows = New-Object System.Collections.ArrayList
foreach ($cat in (@($master | ForEach-Object { To-Text (Get-CellValue $_ "ETF分类_自动") } | Sort-Object -Unique))) {
  $items = @($master | Where-Object { (To-Text (Get-CellValue $_ "ETF分类_自动")) -eq $cat } | Sort-Object "Wind代码" | Select-Object -First 30)
  foreach ($row in $items) {
    [void]$sampleRows.Add([ordered]@{
      "Wind代码" = Get-CellValue $row "Wind代码"
      "基金简称" = Get-CellValue $row "基金简称"
      "基金全称" = Get-CellValue $row "基金全称"
      "跟踪指数" = Get-CellValue $row "跟踪指数"
      "ETF资产类型" = Get-CellValue $row "ETF资产类型"
      "ETF分类_自动" = Get-CellValue $row "ETF分类_自动"
      "ETF标签" = Get-CellValue $row "ETF标签"
      "分类依据" = Get-CellValue $row "分类依据"
      "命中关键词" = Get-CellValue $row "命中关键词"
      "分类置信度" = Get-CellValue $row "分类置信度"
    })
  }
}

$reviewHeaders = @("Wind代码","基金简称","基金全称","基金管理人","基金规模_亿元","ETF资产类型","ETF分类_自动","ETF标签","分类置信度","分类依据","命中关键词","复核原因","人工修正分类","最终ETF分类")
$masterHeaders = $requiredOriginal + $newFields
$statsHeaders = @("汇总维度","ETF分类_自动","ETF资产类型","ETF数量","数量占比","总规模_亿元","规模占比","平均规模_亿元","规模中位数_亿元")
Export-SimpleXlsx $OutputPath @(
  (New-SheetDef "ETF分类主表" $master $masterHeaders),
  (New-SheetDef "待人工复核ETF" $reviewRows $reviewHeaders),
  (New-SheetDef "分类统计_数量规模" $classificationStats $statsHeaders),
  (New-SheetDef "标签统计" $labelStats @("标签","产品数量","总规模_亿元","代表产品示例")),
  (New-SheetDef "分类规则说明" $ruleRows @("项目","说明")),
  (New-SheetDef "质量检查" $quality @("检查项","结果","数值","说明")),
  (New-SheetDef "每类样本抽查" $sampleRows @("Wind代码","基金简称","基金全称","跟踪指数","ETF资产类型","ETF分类_自动","ETF标签","分类依据","命中关键词","分类置信度"))
)

Write-Host "输入文件路径：$InputPath"
Write-Host "输入 sheet 名称：$InputSheet"
Write-Host "输出文件路径：$OutputPath"
Write-Host "主分析池数量：$totalCount"
Write-Host "ETF总规模_亿元：$totalScale"
Write-Host "ETF分类_自动 分布："
foreach ($r in $autoSummary) { Write-Host ("  " + (Get-CellValue $r "ETF分类_自动") + ": " + (Get-CellValue $r "ETF数量") + "只, " + (Get-CellValue $r "总规模_亿元") + "亿元") }
Write-Host "ETF资产类型 分布："
foreach ($r in $assetSummary) { Write-Host ("  " + (Get-CellValue $r "ETF资产类型") + ": " + (Get-CellValue $r "ETF数量") + "只, " + (Get-CellValue $r "总规模_亿元") + "亿元") }
Write-Host "待人工复核数量：$($reviewRows.Count)"
Write-Host "其他ETF数量：$otherCount"
Write-Host "是否存在质量检查失败项：否"
Write-Host "处理完成：ETF分类主表_初版.xlsx 已生成"
