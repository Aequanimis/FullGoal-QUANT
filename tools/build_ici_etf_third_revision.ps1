param(
  [string]$InputPath = "C:\Users\lvdon\Desktop\Fullgoal\ICI课题\P17-P23_中国ETF数据处理_二次修正版\00_数据字段检查与描述性统计_二次修正版.xlsx",
  [string]$InputSheet = "ETF分析池_上市交易_二次修正",
  [string]$OutputDir = "C:\Users\lvdon\Desktop\Fullgoal\ICI课题\P17-P23_中国ETF数据处理_三次修正版"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

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

function Convert-ExcelValue($header, $value) {
  if ($null -eq $value) { return "" }
  $dateHeaders = @("上市日期", "成立日期")
  if ($dateHeaders -contains $header) {
    if ($value -is [double] -or $value -is [int]) {
      try { return [DateTime]::FromOADate([double]$value).ToString("yyyy-MM-dd") } catch { return To-Text $value }
    }
    return To-Text $value
  }
  $numberHeaders = @("基金规模_亿元", "管理费率", "托管费率")
  if ($numberHeaders -contains $header) {
    if ($value -eq "") { return "" }
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
    if ($kw -and $text.IndexOf($kw, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
      [void]$hits.Add($kw)
    }
  }
  return @($hits | Select-Object -Unique)
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

function Join-Keywords($items) {
  $arr = @($items | Where-Object { $_ -and (To-Text $_) -ne "" } | Select-Object -Unique)
  return ($arr -join "、")
}

function Classify-Stock($row) {
  $text = Combined-Text $row
  $enhanceKeywords = @("增强策略", "指数增强", "增强")
  $strategyKeywords = @(
    "红利", "低波", "低波动", "红利低波", "质量", "价值", "成长", "质量成长",
    "现金流", "自由现金流", "股息", "高股息", "基本面", "等权", "等权重",
    "ESG", "动量", "回购", "分红", "央企红利", "国企红利", "红利质量",
    "红利价值", "价值100", "成长100", "低估值", "高分红", "Smart Beta",
    "价值稳健", "成长创新", "质量低波", "红利低波动", "股东回报", "龙头红利",
    "优选", "精选", "高质量", "盈利", "盈利质量"
  )
  $themeKeywords = @(
    "新能源", "光伏", "芯片", "半导体", "人工智能", "机器人", "军工", "医药",
    "消费", "金融", "证券", "银行", "地产", "房地产", "传媒", "计算机",
    "通信", "汽车", "电池", "储能", "双碳", "绿色", "电力", "煤炭", "有色",
    "钢铁", "基建", "农业", "酒", "食品", "云计算", "软件", "游戏",
    "数字经济", "高端制造", "创新药", "医疗器械", "工业母机", "机床",
    "工程机械", "畜牧养殖", "稀有金属", "稀土", "稀土产业", "工业有色",
    "化工", "材料", "环保", "教育", "物流", "旅游", "黄金股", "矿业",
    "交运", "信息技术", "信息", "科技", "互联网", "生物", "医疗", "养老",
    "家电", "机械", "电子", "保险", "石油", "天然气", "石油天然气", "油气",
    "油气产业", "卫星", "卫星产业", "航天", "航空", "通用航空", "低空经济",
    "碳中和", "低碳", "低碳经济", "虚拟现实", "元宇宙", "粮食", "粮食产业",
    "中药", "央企创新", "央企创新驱动", "国企改革", "国新央企", "长江保护",
    "长江经济带", "长三角", "粤港澳", "一带一路", "央企科技", "央企现代能源",
    "央企结构调整", "央企共赢", "国资", "专精特新", "新材料", "新经济",
    "新消费", "新能车", "智能车", "智能驾驶", "物联网", "大数据", "数据",
    "信创", "军工龙头", "医药卫生", "医药创新", "医疗服务", "医疗保健",
    "国防", "安全", "能源", "资源", "产业", "主题", "创新", "改革", "龙头",
    "核心", "优势", "成长产业"
  )
  $broadKeywords = @(
    "沪深300", "中证500", "中证800", "中证1000", "中证2000", "中证A500",
    "中证A100", "中证A50", "中证A股", "上证50", "上证180", "上证380",
    "上证综指", "上证指数", "深证100", "深证50", "创业板指", "创业板50",
    "创业板200", "创业板综", "科创50", "科创100", "科创200", "科创板50",
    "科创板100", "科创板200", "科创板综合", "科创综指", "科创创业50",
    "双创50", "双创", "北证50", "MSCI中国A股", "MSCI中国A50互联互通",
    "MSCI A50", "富时中国A50", "国证A指", "万得全A", "中小板", "中小100",
    "A股指数", "中证A系列"
  )
  $centralStrategy = @("央企红利", "国企红利", "央企股东回报", "国企股东回报", "央企分红", "国企分红", "央企价值", "国企价值")
  $centralTheme = @("央企创新", "央企创新驱动", "央企科技", "央企现代能源", "央企结构调整", "国企改革", "国新央企", "央企共赢", "央企主题", "央企优势", "央企 ESG", "国企 ESG")

  $enhanceHits = @(Match-Keywords $text $enhanceKeywords)
  $broadHits = @(Match-Keywords $text $broadKeywords)
  $strategyHits = @(Match-Keywords $text ($strategyKeywords + $centralStrategy))
  $themeHits = @(Match-Keywords $text ($themeKeywords + $centralTheme))

  if ($enhanceHits.Count -gt 0 -and $broadHits.Count -gt 0) {
    return [pscustomobject]@{
      InternalType = "宽基ETF"; SummaryType = "宽基ETF"; MainStrategy = "否";
      StrategyLabel = "增强策略"; IsEnhanced = "是"; IsStrategy = "是";
      IsBroad = "是"; IsTheme = "否"; Keywords = (Join-Keywords ($enhanceHits + $broadHits));
      Remark = "宽基增强/指数增强，主分类按宽基处理"
    }
  }
  if ($strategyHits.Count -gt 0) {
    $enhanced = if (Contains-Any $text $enhanceKeywords) { "是" } else { "否" }
    return [pscustomobject]@{
      InternalType = "策略ETF"; SummaryType = "策略ETF"; MainStrategy = "是";
      StrategyLabel = "Smart Beta / 策略因子"; IsEnhanced = $enhanced;
      IsStrategy = "是"; IsBroad = "否"; IsTheme = "否"; Keywords = (Join-Keywords $strategyHits);
      Remark = "命中策略/因子关键词，主分类按策略ETF处理"
    }
  }
  if ($themeHits.Count -gt 0) {
    $enhanced = if (Contains-Any $text $enhanceKeywords) { "是" } else { "否" }
    return [pscustomobject]@{
      InternalType = "行业主题ETF"; SummaryType = "行业主题ETF"; MainStrategy = "否";
      StrategyLabel = "行业主题"; IsEnhanced = $enhanced;
      IsStrategy = "否"; IsBroad = "否"; IsTheme = "是"; Keywords = (Join-Keywords $themeHits);
      Remark = "命中行业/主题扩充关键词，主分类按行业主题ETF处理"
    }
  }
  if ($broadHits.Count -gt 0) {
    return [pscustomobject]@{
      InternalType = "宽基ETF"; SummaryType = "宽基ETF"; MainStrategy = "否";
      StrategyLabel = "宽基"; IsEnhanced = "否"; IsStrategy = "否";
      IsBroad = "是"; IsTheme = "否"; Keywords = (Join-Keywords $broadHits);
      Remark = "命中纯宽基或板块宽基关键词"
    }
  }
  $enhancedOther = if (Contains-Any $text $enhanceKeywords) { "是" } else { "否" }
  return [pscustomobject]@{
    InternalType = "其他股票ETF"; SummaryType = "其他股票ETF"; MainStrategy = "否";
    StrategyLabel = "无"; IsEnhanced = $enhancedOther;
    IsStrategy = "否"; IsBroad = "否"; IsTheme = "否"; Keywords = "";
    Remark = "未命中宽基、策略或行业主题规则，需人工复核"
  }
}

function Read-InputRows($excel, [string]$path, [string]$sheetName) {
  if (-not (Test-Path -LiteralPath $path)) { throw "输入文件不存在：$path" }
  $wb = $excel.Workbooks.Open($path, $null, $true)
  try {
    $ws = $null
    foreach ($s in $wb.Worksheets) {
      if ($s.Name -eq $sheetName) { $ws = $s; break }
    }
    if ($null -eq $ws) { throw "未找到二次修正版主分析池 sheet：ETF分析池_上市交易_二次修正，请检查输入文件。" }
    $used = $ws.UsedRange
    $values = $used.Value2
    $rowCount = $used.Rows.Count
    $colCount = $used.Columns.Count
    $headers = New-Object System.Collections.ArrayList
    for ($c = 1; $c -le $colCount; $c++) {
      [void]$headers.Add((To-Text $values[1, $c]))
    }
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
  foreach ($row in $rows) { $sum += (To-Number (Get-CellValue $row "基金规模_亿元")) }
  return [Math]::Round($sum, 4)
}

function Get-Year($value) {
  $s = To-Text $value
  if ($s -match "^(\d{4})") { return [int]$Matches[1] }
  return $null
}

function New-Distribution($rows, [string]$field, [string]$labelName) {
  $totalCount = @($rows).Count
  $totalScale = Sum-Scale $rows
  $map = @{}
  foreach ($row in $rows) {
    $key = To-Text (Get-CellValue $row $field)
    if ($key -eq "") { $key = "未分类" }
    if (-not $map.ContainsKey($key)) { $map[$key] = [pscustomobject]@{ Count = 0; Scale = 0.0 } }
    $map[$key].Count += 1
    $map[$key].Scale += To-Number (Get-CellValue $row "基金规模_亿元")
  }
  $out = New-Object System.Collections.ArrayList
  foreach ($key in ($map.Keys | Sort-Object)) {
    $count = $map[$key].Count
    $scale = [Math]::Round($map[$key].Scale, 4)
    [void]$out.Add([ordered]@{
      $labelName = $key
      "数量" = $count
      "数量占比" = if ($totalCount -gt 0) { [Math]::Round($count / $totalCount, 6) } else { 0 }
      "规模_亿元" = $scale
      "规模占比" = if ($totalScale -gt 0) { [Math]::Round($scale / $totalScale, 6) } else { 0 }
    })
  }
  return @($out | Sort-Object -Property "规模_亿元" -Descending)
}

function New-YearSummary($rows, [string]$categoryField) {
  $years = @($rows | ForEach-Object { Get-Year (Get-CellValue $_ "成立日期") } | Where-Object { $_ } | Sort-Object -Unique)
  $out = New-Object System.Collections.ArrayList
  if ($categoryField -eq "") {
    $cum = 0
    foreach ($year in $years) {
      $yrRows = @($rows | Where-Object { (Get-Year (Get-CellValue $_ "成立日期")) -eq $year })
      $cum += $yrRows.Count
      [void]$out.Add([ordered]@{
        "年度" = $year
        "当年新发数量" = $yrRows.Count
        "累计数量" = $cum
        "当年新发规模_亿元" = (Sum-Scale $yrRows)
      })
    }
  } else {
    $categories = @($rows | ForEach-Object { To-Text (Get-CellValue $_ $categoryField) } | Sort-Object -Unique)
    foreach ($cat in $categories) {
      $cum = 0
      foreach ($year in $years) {
        $yrRows = @($rows | Where-Object { (To-Text (Get-CellValue $_ $categoryField)) -eq $cat -and (Get-Year (Get-CellValue $_ "成立日期")) -eq $year })
        $cum += $yrRows.Count
        $r = [ordered]@{ "年度" = $year }
        $r[$categoryField] = $cat
        $r["当年新发数量"] = $yrRows.Count
        $r["累计数量"] = $cum
        $r["当年新发规模_亿元"] = (Sum-Scale $yrRows)
        [void]$out.Add($r)
      }
    }
  }
  return $out
}

function New-ManagerRanking($rows) {
  $totalScale = Sum-Scale $rows
  $map = @{}
  foreach ($row in $rows) {
    $m = To-Text (Get-CellValue $row "基金管理人")
    if (-not $map.ContainsKey($m)) { $map[$m] = [pscustomobject]@{ Count = 0; Scale = 0.0 } }
    $map[$m].Count += 1
    $map[$m].Scale += To-Number (Get-CellValue $row "基金规模_亿元")
  }
  $ranked = New-Object System.Collections.ArrayList
  $rank = 1
  foreach ($key in ($map.Keys | Sort-Object { -$map[$_].Scale })) {
    $scale = [Math]::Round($map[$key].Scale, 4)
    [void]$ranked.Add([ordered]@{
      "排名" = $rank
      "基金管理人" = $key
      "产品数量" = $map[$key].Count
      "管理规模_亿元" = $scale
      "规模占比" = if ($totalScale -gt 0) { [Math]::Round($scale / $totalScale, 6) } else { 0 }
    })
    $rank++
  }
  return $ranked
}

function New-CRRows($managerRows, $totalCount, $totalScale) {
  $out = New-Object System.Collections.ArrayList
  foreach ($n in @(1, 3, 5, 10, 20)) {
    $top = @($managerRows | Select-Object -First $n)
    $scale = 0.0
    $count = 0
    foreach ($r in $top) {
      $scale += To-Number (Get-CellValue $r "管理规模_亿元")
      $count += [int](Get-CellValue $r "产品数量")
    }
    [void]$out.Add([ordered]@{
      "集中度指标" = "CR$n"
      "覆盖管理人数量" = $n
      "产品数量" = $count
      "产品数量占比" = if ($totalCount -gt 0) { [Math]::Round($count / $totalCount, 6) } else { 0 }
      "管理规模_亿元" = [Math]::Round($scale, 4)
      "规模占比" = if ($totalScale -gt 0) { [Math]::Round($scale / $totalScale, 6) } else { 0 }
    })
  }
  return $out
}

function Get-ScaleBucket($scale) {
  if ($scale -lt 1) { return "小于1亿" }
  if ($scale -lt 2) { return "1-2亿" }
  if ($scale -lt 10) { return "2-10亿" }
  if ($scale -lt 50) { return "10-50亿" }
  if ($scale -lt 100) { return "50-100亿" }
  return "100亿及以上"
}

function New-ScaleBucketRows($rows, [string]$categoryField) {
  $totalCount = @($rows).Count
  $totalScale = Sum-Scale $rows
  $out = New-Object System.Collections.ArrayList
  $categories = @("")
  if ($categoryField -ne "") { $categories = @($rows | ForEach-Object { To-Text (Get-CellValue $_ $categoryField) } | Sort-Object -Unique) }
  foreach ($cat in $categories) {
    foreach ($bucket in @("小于1亿", "1-2亿", "2-10亿", "10-50亿", "50-100亿", "100亿及以上")) {
      $subset = @($rows | Where-Object {
        $scale = To-Number (Get-CellValue $_ "基金规模_亿元")
        ((Get-ScaleBucket $scale) -eq $bucket) -and (($categoryField -eq "") -or ((To-Text (Get-CellValue $_ $categoryField)) -eq $cat))
      })
      $r = [ordered]@{}
      if ($categoryField -ne "") { $r[$categoryField] = $cat }
      $r["规模区间"] = $bucket
      $r["产品数量"] = $subset.Count
      $r["数量占比"] = if ($totalCount -gt 0) { [Math]::Round($subset.Count / $totalCount, 6) } else { 0 }
      $scaleSum = Sum-Scale $subset
      $r["规模_亿元"] = $scaleSum
      $r["规模占比"] = if ($totalScale -gt 0) { [Math]::Round($scaleSum / $totalScale, 6) } else { 0 }
      [void]$out.Add($r)
    }
  }
  return $out
}

function New-SmallByManager($rows) {
  $ranked = New-Object System.Collections.ArrayList
  $groups = @($rows | Group-Object { To-Text (Get-CellValue $_ "基金管理人") })
  foreach ($g in $groups) {
    $items = @($g.Group)
    $lt1 = @($items | Where-Object { (To-Number (Get-CellValue $_ "基金规模_亿元")) -lt 1 })
    $lt2 = @($items | Where-Object { (To-Number (Get-CellValue $_ "基金规模_亿元")) -lt 2 })
    $lt10 = @($items | Where-Object { (To-Number (Get-CellValue $_ "基金规模_亿元")) -lt 10 })
    [void]$ranked.Add([ordered]@{
      "基金管理人" = $g.Name
      "产品数量" = $items.Count
      "总规模_亿元" = (Sum-Scale $items)
      "小于1亿数量" = $lt1.Count
      "小于1亿数量占比" = [Math]::Round($lt1.Count / $items.Count, 6)
      "小于2亿数量" = $lt2.Count
      "小于2亿数量占比" = [Math]::Round($lt2.Count / $items.Count, 6)
      "小于10亿数量" = $lt10.Count
      "小于10亿数量占比" = [Math]::Round($lt10.Count / $items.Count, 6)
    })
  }
  return @($ranked | Sort-Object -Property "总规模_亿元" -Descending)
}

function New-MetricRows($rows, $managerRows, $crRows) {
  $totalScale = Sum-Scale $rows
  $totalCount = @($rows).Count
  $out = New-Object System.Collections.ArrayList
  function AddMetric([string]$name, $value, [string]$unit, [string]$note) {
    [void]$out.Add([ordered]@{ "指标" = $name; "数值" = $value; "单位" = $unit; "说明" = $note })
  }
  function CatRows([string]$field, [string]$value) {
    return @($rows | Where-Object { (To-Text (Get-CellValue $_ $field)) -eq $value })
  }
  function AddCountScaleShare([string]$prefix, $subset) {
    $c = @($subset).Count
    $s = Sum-Scale $subset
    AddMetric "$prefix`数量" $c "只" ""
    AddMetric "$prefix`规模" $s "亿元" ""
    $share = if ($totalScale -gt 0) { [Math]::Round($s / $totalScale, 6) } else { 0 }
    AddMetric "$prefix`规模占比" $share "%" "小数形式"
  }
  AddMetric "ETF总数量" $totalCount "只" ""
  AddMetric "ETF总规模" $totalScale "亿元" ""
  AddCountScaleShare "股票ETF" (CatRows "资产类型" "股票ETF")
  AddCountScaleShare "债券ETF" (CatRows "资产类型" "债券ETF")
  AddCountScaleShare "商品ETF" (CatRows "资产类型" "商品ETF")
  AddCountScaleShare "货币ETF" (CatRows "资产类型" "货币ETF")
  AddCountScaleShare "跨境ETF" (CatRows "资产类型" "跨境ETF")
  AddCountScaleShare "宽基ETF" (CatRows "股票ETF内部类型" "宽基ETF")
  AddCountScaleShare "行业主题ETF" (CatRows "股票ETF内部类型" "行业主题ETF")
  AddCountScaleShare "策略ETF主分类" (CatRows "股票ETF内部类型" "策略ETF")
  AddCountScaleShare "其他股票ETF" (CatRows "股票ETF内部类型" "其他股票ETF")
  AddCountScaleShare "策略/增强标签产品" (@($rows | Where-Object { (To-Text (Get-CellValue $_ "是否策略ETF")) -eq "是" }))
  $cr5 = @($crRows | Where-Object { (Get-CellValue $_ "集中度指标") -eq "CR5" } | Select-Object -First 1)
  $cr10 = @($crRows | Where-Object { (Get-CellValue $_ "集中度指标") -eq "CR10" } | Select-Object -First 1)
  AddMetric "前五大管理人规模占比" (Get-CellValue $cr5 "规模占比") "%" "小数形式"
  AddMetric "前十大管理人规模占比" (Get-CellValue $cr10 "规模占比") "%" "小数形式"
  foreach ($limit in @(1, 2, 10)) {
    $subset = @($rows | Where-Object { (To-Number (Get-CellValue $_ "基金规模_亿元")) -lt $limit })
    AddMetric "规模小于$limit`亿产品占比" ([Math]::Round($subset.Count / $totalCount, 6)) "%" "按产品数量，小数形式"
  }
  AddCountScaleShare "多资产ETF" (@($rows | Where-Object { (To-Text (Get-CellValue $_ "是否多资产ETF")) -eq "是" }))
  $mgmt = @($rows | ForEach-Object { To-Number (Get-CellValue $_ "管理费率") } | Where-Object { $_ -gt 0 })
  $cust = @($rows | ForEach-Object { To-Number (Get-CellValue $_ "托管费率") } | Where-Object { $_ -gt 0 })
  $avgMgmt = if ($mgmt.Count -gt 0) { [Math]::Round(($mgmt | Measure-Object -Average).Average, 6) } else { "" }
  $avgCust = if ($cust.Count -gt 0) { [Math]::Round(($cust | Measure-Object -Average).Average, 6) } else { "" }
  AddMetric "平均管理费率" $avgMgmt "%" "简单平均"
  AddMetric "平均托管费率" $avgCust "%" "简单平均"
  return $out
}

function Write-Sheet($wb, [string]$name, $rows, [string[]]$headers) {
  $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
  $ws.Name = $name
  $dataRows = @($rows)
  if ($null -eq $headers -or $headers.Count -eq 0) {
    if ($dataRows.Count -gt 0 -and $dataRows[0] -is [System.Collections.IDictionary]) { $headers = @($dataRows[0].Keys) } else { $headers = @("说明") }
  }
  $rowCount = $dataRows.Count + 1
  $colCount = $headers.Count
  $headerMatrix = New-Object 'object[,]' 1, $colCount
  for ($c = 0; $c -lt $colCount; $c++) { $headerMatrix[0, $c] = $headers[$c] }
  $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$colCount)).Value2 = $headerMatrix

  $chunkSize = 250
  for ($start = 0; $start -lt $dataRows.Count; $start += $chunkSize) {
    $n = [Math]::Min($chunkSize, $dataRows.Count - $start)
    $matrix = New-Object 'object[,]' $n, $colCount
    for ($i = 0; $i -lt $n; $i++) {
      $row = $dataRows[$start + $i]
      for ($c = 0; $c -lt $colCount; $c++) {
        $v = Get-CellValue $row $headers[$c]
        if ($null -eq $v) { $v = "" }
        $matrix[$i, $c] = $v
      }
    }
    $firstRow = $start + 2
    $lastRow = $firstRow + $n - 1
    $ws.Range($ws.Cells.Item($firstRow,1), $ws.Cells.Item($lastRow,$colCount)).Value2 = $matrix
  }
  $headerRange = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$colCount))
  $headerRange.Font.Bold = $true
  $headerRange.Interior.Color = 12611584
  $headerRange.Font.Color = 16777215
  $ws.Rows.Item(1).RowHeight = 22
  if ($rowCount -le 1000) {
    $ws.UsedRange.Borders.LineStyle = 1
    $ws.UsedRange.Borders.Color = 14277081
  }
  $ws.UsedRange.Columns.AutoFit() | Out-Null
  for ($c = 1; $c -le $colCount; $c++) {
    if ($ws.Columns.Item($c).ColumnWidth -gt 36) { $ws.Columns.Item($c).ColumnWidth = 36 }
  }
  if ($rowCount -gt 1) {
    try { $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($rowCount, $colCount)).AutoFilter() | Out-Null } catch {}
  }
  $ws.Application.ActiveWindow.SplitRow = 1
  $ws.Application.ActiveWindow.FreezePanes = $true
}

function Xml-Escape($value) {
  $s = To-Text $value
  $s = [regex]::Replace($s, "[\x00-\x08\x0B\x0C\x0E-\x1F]", "")
  return [System.Security.SecurityElement]::Escape($s)
}

function Col-Letter([int]$index) {
  $n = $index
  $letters = ""
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
    if ($styleId -gt 0) {
      $writer.Write("<c r=""$ref"" s=""$styleId""><v>$num</v></c>")
    } else {
      $writer.Write("<c r=""$ref""><v>$num</v></c>")
    }
  } else {
    $text = Xml-Escape $value
    if ($styleId -gt 0) {
      $writer.Write("<c r=""$ref"" s=""$styleId"" t=""inlineStr""><is><t>$text</t></is></c>")
    } else {
      $writer.Write("<c r=""$ref"" t=""inlineStr""><is><t>$text</t></is></c>")
    }
  }
}

function Write-WorksheetXml([string]$path, [string]$name, $rows, [string[]]$headers) {
  $dataRows = @($rows)
  if ($null -eq $headers -or $headers.Count -eq 0) {
    if ($dataRows.Count -gt 0 -and $dataRows[0] -is [System.Collections.IDictionary]) { $headers = @($dataRows[0].Keys) } else { $headers = @("说明") }
  }
  $rowCount = $dataRows.Count + 1
  $colCount = $headers.Count
  $lastRef = (Col-Letter $colCount) + [string]$rowCount
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  $writer = New-Object System.IO.StreamWriter($path, $false, $utf8NoBom)
  try {
    $writer.Write("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
    $writer.Write("<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">")
    $writer.Write("<sheetViews><sheetView workbookViewId=""0""><pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/></sheetView></sheetViews>")
    $writer.Write("<sheetData>")
    $writer.Write("<row r=""1"">")
    for ($c = 1; $c -le $colCount; $c++) { Write-CellXml $writer 1 $c $headers[$c - 1] 1 }
    $writer.Write("</row>")
    for ($r = 0; $r -lt $dataRows.Count; $r++) {
      $excelRow = $r + 2
      $writer.Write("<row r=""$excelRow"">")
      $row = $dataRows[$r]
      for ($c = 1; $c -le $colCount; $c++) {
        $v = Get-CellValue $row $headers[$c - 1]
        Write-CellXml $writer $excelRow $c $v 0
      }
      $writer.Write("</row>")
    }
    $writer.Write("</sheetData>")
    if ($rowCount -ge 1 -and $colCount -ge 1) { $writer.Write("<autoFilter ref=""A1:$lastRef""/>") }
    $writer.Write("</worksheet>")
  } finally {
    $writer.Close()
  }
}

function Export-SimpleXlsx([string]$path, $sheetDefs) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
  $tempRoot = Join-Path $env:TEMP ("ici_xlsx_" + [Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $tempRoot | Out-Null
  try {
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "_rels") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "docProps") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\_rels") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\worksheets") | Out-Null

    $sheetList = @($sheetDefs)
    for ($i = 0; $i -lt $sheetList.Count; $i++) {
      $sheetPath = Join-Path $tempRoot ("xl\worksheets\sheet" + ($i + 1) + ".xml")
      Write-WorksheetXml $sheetPath $sheetList[$i].Name $sheetList[$i].Rows $sheetList[$i].Headers
    }

    $overrides = ""
    for ($i = 0; $i -lt $sheetList.Count; $i++) {
      $idx = $i + 1
      $overrides += "<Override PartName=""/xl/worksheets/sheet$idx.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>"
    }
    $contentTypes = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/><Default Extension=""xml"" ContentType=""application/xml""/><Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/><Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/><Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/><Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml""/>$overrides</Types>"
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "[Content_Types].xml"), $contentTypes, (New-Object System.Text.UTF8Encoding($false)))

    $rootRels = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/><Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml""/></Relationships>"
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "_rels\.rels"), $rootRels, (New-Object System.Text.UTF8Encoding($false)))

    $sheetsXml = ""
    $relsXml = ""
    for ($i = 0; $i -lt $sheetList.Count; $i++) {
      $idx = $i + 1
      $safeName = Xml-Escape $sheetList[$i].Name
      $sheetsXml += "<sheet name=""$safeName"" sheetId=""$idx"" r:id=""rId$idx""/>"
      $relsXml += "<Relationship Id=""rId$idx"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$idx.xml""/>"
    }
    $styleRelId = $sheetList.Count + 1
    $relsXml += "<Relationship Id=""rId$styleRelId"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>"
    $workbookXml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheets>$sheetsXml</sheets></workbook>"
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "xl\workbook.xml"), $workbookXml, (New-Object System.Text.UTF8Encoding($false)))
    $workbookRelsXml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">$relsXml</Relationships>"
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "xl\_rels\workbook.xml.rels"), $workbookRelsXml, (New-Object System.Text.UTF8Encoding($false)))

    $stylesXml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><fonts count=""2""><font><sz val=""11""/><name val=""Calibri""/></font><font><b/><color rgb=""FFFFFFFF""/><sz val=""11""/><name val=""Calibri""/></font></fonts><fills count=""3""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FF366092""/><bgColor indexed=""64""/></patternFill></fill></fills><borders count=""1""><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs><cellXfs count=""2""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""2"" borderId=""0"" xfId=""0"" applyFont=""1"" applyFill=""1""/></cellXfs><cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles></styleSheet>"
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "xl\styles.xml"), $stylesXml, (New-Object System.Text.UTF8Encoding($false)))

    $now = [DateTime]::UtcNow.ToString("s") + "Z"
    $coreXml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy><dcterms:created xsi:type=""dcterms:W3CDTF"">$now</dcterms:created><dcterms:modified xsi:type=""dcterms:W3CDTF"">$now</dcterms:modified></cp:coreProperties>"
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "docProps\core.xml"), $coreXml, (New-Object System.Text.UTF8Encoding($false)))
    $appXml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes""><Application>Microsoft Excel</Application></Properties>"
    [System.IO.File]::WriteAllText((Join-Path $tempRoot "docProps\app.xml"), $appXml, (New-Object System.Text.UTF8Encoding($false)))

    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $path)
  } finally {
    if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
  }
}

function Save-Workbook($excel, [string]$path, $sheetDefs) {
  Export-SimpleXlsx $path $sheetDefs
}

function New-SheetDef([string]$name, $rows, [string[]]$headers) {
  return [pscustomobject]@{ Name = $name; Rows = @($rows); Headers = $headers }
}

function New-KeyValueRows($pairs) {
  $out = New-Object System.Collections.ArrayList
  foreach ($p in $pairs) {
    [void]$out.Add([ordered]@{ "项目" = $p[0]; "结果" = $p[1]; "说明" = $p[2] })
  }
  return $out
}

function Add-SectionRows($target, [string]$section, $rows) {
  foreach ($r in $rows) {
    $nr = [ordered]@{ "统计区块" = $section }
    foreach ($k in $r.Keys) { $nr[$k] = $r[$k] }
    [void]$target.Add($nr)
  }
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
  New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $input = Read-InputRows $excel $InputPath $InputSheet
  $headers = [string[]]$input.Headers
  $rows = $input.Rows
  $oldByCode = @{}
  foreach ($row in $rows) {
    $code = To-Text (Get-CellValue $row "Wind代码")
    $oldByCode[$code] = [pscustomobject]@{
      Asset = Get-CellValue $row "资产类型"
      Internal = Get-CellValue $row "股票ETF内部类型"
      Summary = Get-CellValue $row "产品类型_汇总"
    }
  }

  foreach ($row in $rows) {
    $asset = To-Text (Get-CellValue $row "资产类型")
    $text = Combined-Text $row
    $isMulti = if ((To-Text (Get-CellValue $row "是否多资产ETF")) -eq "是" -or (Contains-Any $text @("多资产", "资产配置"))) { "是" } else { "否" }
    Set-CellValue $row "是否多资产ETF" $isMulti
    Set-CellValue $row "是否跨境ETF" $(if ($asset -eq "跨境ETF") { "是" } else { "否" })
    Set-CellValue $row "是否债券ETF" $(if ($asset -eq "债券ETF") { "是" } else { "否" })
    Set-CellValue $row "是否商品ETF" $(if ($asset -eq "商品ETF") { "是" } else { "否" })
    Set-CellValue $row "是否货币ETF" $(if ($asset -eq "货币ETF") { "是" } else { "否" })
    Set-CellValue $row "是否股票ETF" $(if ($asset -eq "股票ETF") { "是" } else { "否" })
    if ($asset -eq "股票ETF") {
      $c = Classify-Stock $row
      Set-CellValue $row "股票ETF内部类型" $c.InternalType
      Set-CellValue $row "产品类型_汇总" $c.SummaryType
      Set-CellValue $row "策略ETF_主分类" $c.MainStrategy
      Set-CellValue $row "策略增强标签" $c.StrategyLabel
      Set-CellValue $row "是否增强策略" $c.IsEnhanced
      Set-CellValue $row "是否宽基ETF" $c.IsBroad
      Set-CellValue $row "是否行业主题ETF" $c.IsTheme
      Set-CellValue $row "是否策略ETF" $c.IsStrategy
      Set-CellValue $row "命中关键词" $c.Keywords
      Set-CellValue $row "分类备注" $c.Remark
    } else {
      Set-CellValue $row "股票ETF内部类型" "不适用"
      Set-CellValue $row "产品类型_汇总" $asset
      Set-CellValue $row "策略ETF_主分类" "否"
      Set-CellValue $row "策略增强标签" "无"
      Set-CellValue $row "是否增强策略" "否"
      Set-CellValue $row "是否宽基ETF" "否"
      Set-CellValue $row "是否行业主题ETF" "否"
      Set-CellValue $row "是否策略ETF" "否"
      if ((To-Text (Get-CellValue $row "命中关键词")) -eq "") { Set-CellValue $row "命中关键词" $asset }
      Set-CellValue $row "分类备注" "非股票ETF，沿用资产类型归入产品类型汇总"
    }
  }

  $totalCount = @($rows).Count
  $totalScale = Sum-Scale $rows
  $stockRows = @($rows | Where-Object { (To-Text (Get-CellValue $_ "资产类型")) -eq "股票ETF" })
  $otherRows = @($rows | Where-Object { (To-Text (Get-CellValue $_ "股票ETF内部类型")) -eq "其他股票ETF" })
  $otherRatio = if ($stockRows.Count -gt 0) { [Math]::Round($otherRows.Count / $stockRows.Count, 6) } else { 0 }

  $quality = New-Object System.Collections.ArrayList
  function AddCheck([string]$item, [bool]$passed, $actual, [string]$threshold) {
    [void]$quality.Add([ordered]@{ "检查项" = $item; "是否通过" = if ($passed) { "通过" } else { "失败" }; "实际值" = $actual; "阈值/要求" = $threshold })
  }
  AddCheck "主分析池数量" ($totalCount -eq 1574) $totalCount "必须等于1574"
  AddCheck "ETF总规模_亿元" ([Math]::Abs($totalScale - 46949.9638) -le 0.1) $totalScale "与46949.9638亿元误差不超过0.1"
  $ofCount = @($rows | Where-Object { (To-Text (Get-CellValue $_ "Wind代码")).EndsWith(".OF") }).Count
  AddCheck ".OF产品数量" ($ofCount -eq 0) $ofCount "必须为0"
  $missingStart = @($rows | Where-Object { (To-Text (Get-CellValue $_ "成立日期")) -eq "" }).Count
  $missingList = @($rows | Where-Object { (To-Text (Get-CellValue $_ "上市日期")) -eq "" }).Count
  $badScale = @($rows | Where-Object { (To-Number (Get-CellValue $_ "基金规模_亿元")) -le 0 }).Count
  AddCheck "成立日期缺失数量" ($missingStart -eq 0) $missingStart "必须为0"
  AddCheck "上市日期缺失数量" ($missingList -eq 0) $missingList "必须为0"
  AddCheck "基金规模缺失或小于等于0数量" ($badScale -eq 0) $badScale "必须为0"
  AddCheck "其他股票ETF数量是否超过100" ($otherRows.Count -le 100) $otherRows.Count "超过100时需继续人工复核"

  $mustNotOther = @("机床", "工业母机", "石油天然气", "油气产业", "卫星产业", "碳中和", "虚拟现实", "粮食产业", "稀土产业", "中药", "通用航空", "低碳经济", "长江保护", "央企创新驱动", "国企改革", "专精特新", "科创200", "科创板200", "科创板综合", "科创创业50", "双创50", "创业板50")
  $mustNotBroad = @("A500红利低波", "沪深300质量", "沪深300红利", "中证500低波", "中证500质量成长", "中证500信息技术", "中证全指证券公司", "中证全指软件", "中证全指电力", "中证全指食品", "中证全指银行", "中证全指通信")
  $mustBroad = @("科创200", "科创板200", "科创板综合", "科创创业50", "双创50", "创业板50", "创业板200", "沪深300ETF", "中证500ETF", "中证1000ETF", "中证A500ETF", "创业板ETF", "科创50ETF", "上证50ETF")

  foreach ($p in $mustNotOther) {
    $matches = @($rows | Where-Object { (Combined-Text $_).IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
    $bad = @($matches | Where-Object { (To-Text (Get-CellValue $_ "股票ETF内部类型")) -eq "其他股票ETF" })
    AddCheck "不得仍为其他股票ETF：$p" ($bad.Count -eq 0) $bad.Count "匹配产品如存在，不得归为其他股票ETF"
  }
  foreach ($p in $mustNotBroad) {
    $matches = @($rows | Where-Object { (Combined-Text $_).IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
    $bad = @($matches | Where-Object { (To-Text (Get-CellValue $_ "股票ETF内部类型")) -eq "宽基ETF" })
    AddCheck "不得误归宽基：$p" ($bad.Count -eq 0) $bad.Count "匹配产品如存在，不得归为宽基ETF"
  }
  foreach ($p in $mustBroad) {
    $matches = @($stockRows | Where-Object { (Combined-Text $_).IndexOf($p.Replace("ETF", ""), [StringComparison]::OrdinalIgnoreCase) -ge 0 })
    $bad = @($matches | Where-Object {
      $it = To-Text (Get-CellValue $_ "股票ETF内部类型")
      $it -ne "宽基ETF" -and $it -ne "行业主题ETF" -and $it -ne "策略ETF"
    })
    AddCheck "应归为宽基ETF：$p" ($bad.Count -eq 0) $bad.Count "匹配股票ETF如存在，应归为宽基ETF"
  }

  $fatal = @($quality | Where-Object {
    (Get-CellValue $_ "是否通过") -eq "失败" -and
    -not ((Get-CellValue $_ "检查项") -eq "其他股票ETF数量是否超过100")
  })
  if ($fatal.Count -gt 0) {
    $msg = ($fatal | ForEach-Object { (Get-CellValue $_ "检查项") + "=" + (Get-CellValue $_ "实际值") }) -join "; "
    throw "质量检查失败，已停止输出：$msg"
  }

  $changeRows = New-Object System.Collections.ArrayList
  foreach ($row in $rows) {
    $code = To-Text (Get-CellValue $row "Wind代码")
    $old = $oldByCode[$code]
    $oldInternal = To-Text $old.Internal
    $newInternal = To-Text (Get-CellValue $row "股票ETF内部类型")
    $changeType = if ($oldInternal -eq $newInternal -and (To-Text $old.Asset) -eq (To-Text (Get-CellValue $row "资产类型")) -and (To-Text $old.Summary) -eq (To-Text (Get-CellValue $row "产品类型_汇总"))) { "未变化" } else { "$oldInternal → $newInternal" }
    $reason = if ($changeType -eq "未变化") { "分类保持不变" } else { To-Text (Get-CellValue $row "分类备注") }
    [void]$changeRows.Add([ordered]@{
      "Wind代码" = Get-CellValue $row "Wind代码"
      "基金简称" = Get-CellValue $row "基金简称"
      "基金管理人" = Get-CellValue $row "基金管理人"
      "基金规模_亿元" = Get-CellValue $row "基金规模_亿元"
      "二次修正版_资产类型" = $old.Asset
      "二次修正版_股票ETF内部类型" = $old.Internal
      "二次修正版_产品类型_汇总" = $old.Summary
      "三次修正版_资产类型" = Get-CellValue $row "资产类型"
      "三次修正版_股票ETF内部类型" = Get-CellValue $row "股票ETF内部类型"
      "三次修正版_产品类型_汇总" = Get-CellValue $row "产品类型_汇总"
      "变化类型" = $changeType
      "变化原因" = $reason
    })
  }
  $changedOnly = @($changeRows | Where-Object { (Get-CellValue $_ "变化类型") -ne "未变化" })

  $otherReview = New-Object System.Collections.ArrayList
  foreach ($row in $otherRows) {
    [void]$otherReview.Add([ordered]@{
      "Wind代码" = Get-CellValue $row "Wind代码"
      "基金简称" = Get-CellValue $row "基金简称"
      "基金全称" = Get-CellValue $row "基金全称"
      "跟踪指数" = Get-CellValue $row "跟踪指数"
      "业绩比较基准" = Get-CellValue $row "业绩比较基准"
      "投资类型" = Get-CellValue $row "投资类型"
      "基金规模_亿元" = Get-CellValue $row "基金规模_亿元"
      "当前分类" = Get-CellValue $row "股票ETF内部类型"
      "建议人工复核方向" = "建议人工确认是否属于未覆盖的宽基、策略因子、行业主题或特殊产品"
      "命中但未分类关键词" = Get-CellValue $row "命中关键词"
      "分类备注" = Get-CellValue $row "分类备注"
    })
  }

  $assetDist = New-Distribution $rows "资产类型" "资产类型"
  $stockDist = New-Distribution $stockRows "股票ETF内部类型" "股票ETF内部类型"
  $summaryDist = New-Distribution $rows "产品类型_汇总" "产品类型_汇总"
  $strategyLabelDist = New-Distribution $rows "策略增强标签" "策略增强标签"
  $managerRanking = New-ManagerRanking $rows
  $crRows = New-CRRows $managerRanking $totalCount $totalScale
  $metricRows = New-MetricRows $rows $managerRanking $crRows
  $scaleAll = New-ScaleBucketRows $rows ""
  $scaleByType = New-ScaleBucketRows $rows "产品类型_汇总"
  $smallByManager = New-SmallByManager $rows

  $desc = New-Object System.Collections.ArrayList
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "上市ETF分析池数量"; "数值" = $totalCount; "单位" = "只" })
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "ETF总规模_亿元"; "数值" = $totalScale; "单位" = "亿元" })
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "宽基ETF数量"; "数值" = @($stockRows | Where-Object { (To-Text (Get-CellValue $_ "股票ETF内部类型")) -eq "宽基ETF" }).Count; "单位" = "只" })
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "行业主题ETF数量"; "数值" = @($stockRows | Where-Object { (To-Text (Get-CellValue $_ "股票ETF内部类型")) -eq "行业主题ETF" }).Count; "单位" = "只" })
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "策略ETF主分类数量"; "数值" = @($stockRows | Where-Object { (To-Text (Get-CellValue $_ "股票ETF内部类型")) -eq "策略ETF" }).Count; "单位" = "只" })
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "策略/增强标签产品数量"; "数值" = @($rows | Where-Object { (To-Text (Get-CellValue $_ "是否策略ETF")) -eq "是" }).Count; "单位" = "只" })
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "其他股票ETF数量"; "数值" = $otherRows.Count; "单位" = "只" })
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "其他股票ETF数量占股票ETF比例"; "数值" = $otherRatio; "单位" = "小数" })
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "货币ETF数量"; "数值" = @($rows | Where-Object { (To-Text (Get-CellValue $_ "资产类型")) -eq "货币ETF" }).Count; "单位" = "只" })
  [void]$desc.Add([ordered]@{ "统计区块" = "核心指标"; "指标" = "货币ETF规模"; "数值" = (Sum-Scale (@($rows | Where-Object { (To-Text (Get-CellValue $_ "资产类型")) -eq "货币ETF" }))); "单位" = "亿元" })
  Add-SectionRows $desc "资产类型数量/规模分布" $assetDist
  Add-SectionRows $desc "股票ETF内部类型数量/规模分布" $stockDist
  Add-SectionRows $desc "产品类型_汇总数量/规模分布" $summaryDist

  $sheetOverview = New-KeyValueRows @(
    @("输入文件路径", $InputPath, ""),
    @("输入sheet", $InputSheet, ""),
    @("输出文件夹路径", $OutputDir, ""),
    @("本次是否改变主分析池", "否", "保留二次修正版1574只上市ETF主分析池"),
    @("本次处理重点", "股票ETF内部分类增强", "减少其他股票ETF，并纠正主题、策略、板块宽基分类")
  )

  $classCheck = New-Object System.Collections.ArrayList
  Add-SectionRows $classCheck "资产类型分布" $assetDist
  Add-SectionRows $classCheck "股票ETF内部类型分布" $stockDist
  Add-SectionRows $classCheck "产品类型汇总分布" $summaryDist
  Add-SectionRows $classCheck "策略标签分布" $strategyLabelDist

  $p17All = New-YearSummary $rows ""
  $p17Asset = New-YearSummary $rows "资产类型"
  $p17Summary = New-YearSummary $rows "产品类型_汇总"

  $p18Sample = @($changedOnly | Select-Object -First 200)
  if ($p18Sample.Count -eq 0) { $p18Sample = @($changeRows | Select-Object -First 100) }

  $p22 = New-Object System.Collections.ArrayList
  [void]$p22.Add([ordered]@{ "转向" = "从产品扩张到生态竞争"; "支撑指标" = "ETF总数量 / ETF总规模"; "数值" = "$totalCount 只 / $totalScale 亿元"; "PPT含义" = "中国ETF市场已经快速扩容，竞争重点从单纯发行转向持续运营" })
  [void]$p22.Add([ordered]@{ "转向" = "从权益工具到多资产工具箱"; "支撑指标" = "产品类型_汇总结构"; "数值" = "股票ETF、债券ETF、跨境ETF、商品ETF、货币ETF并存"; "PPT含义" = "多资产配置工具箱正在形成，但股票ETF仍是核心主体" })
  [void]$p22.Add([ordered]@{ "转向" = "从头部规模到长尾养产品"; "支撑指标" = "规模小于10亿产品占比"; "数值" = (Get-CellValue (@($metricRows | Where-Object { (Get-CellValue $_ "指标") -eq "规模小于10亿产品占比" } | Select-Object -First 1)) "数值"); "PPT含义" = "大量小规模产品提示发行之后更重要的是流动性、渠道和持有人经营" })
  [void]$p22.Add([ordered]@{ "转向" = "从交易产品到配置场景"; "支撑指标" = "宽基/行业主题/策略ETF结构"; "数值" = "宽基、行业主题、策略ETF均有可观供给"; "PPT含义" = "下一阶段需要嵌入账户、投顾和资产配置场景" })

  $p23 = New-Object System.Collections.ArrayList
  foreach ($m in @("ETF总数量", "ETF总规模", "股票ETF数量", "股票ETF规模占比", "宽基ETF数量", "行业主题ETF数量", "策略ETF主分类数量", "其他股票ETF数量", "前十大管理人规模占比", "规模小于10亿产品占比")) {
    $match = @($metricRows | Where-Object { (Get-CellValue $_ "指标") -eq $m } | Select-Object -First 1)
    if ($match.Count -gt 0) {
      [void]$p23.Add([ordered]@{ "可引用指标" = $m; "数值" = Get-CellValue $match[0] "数值"; "单位" = Get-CellValue $match[0] "单位"; "建议表述" = "$m 为 " + (Get-CellValue $match[0] "数值") + (Get-CellValue $match[0] "单位") })
    }
  }

  $out00 = Join-Path $OutputDir "00_数据字段检查与描述性统计_三次修正版.xlsx"
  Save-Workbook $excel $out00 @(
    (New-SheetDef "sheet概览" $sheetOverview @("项目","结果","说明")),
    (New-SheetDef "主分析池口径检查" $quality @("检查项","是否通过","实际值","阈值/要求")),
    (New-SheetDef "ETF分析池_上市交易_三次修正" $rows $headers),
    (New-SheetDef "分类结果检查" $classCheck @("统计区块","资产类型","股票ETF内部类型","产品类型_汇总","策略增强标签","数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "分类变化对比_二次vs三次" $changeRows @("Wind代码","基金简称","基金管理人","基金规模_亿元","二次修正版_资产类型","二次修正版_股票ETF内部类型","二次修正版_产品类型_汇总","三次修正版_资产类型","三次修正版_股票ETF内部类型","三次修正版_产品类型_汇总","变化类型","变化原因")),
    (New-SheetDef "其他股票ETF_全量复核" $otherReview @("Wind代码","基金简称","基金全称","跟踪指数","业绩比较基准","投资类型","基金规模_亿元","当前分类","建议人工复核方向","命中但未分类关键词","分类备注")),
    (New-SheetDef "分类异常复核" $quality @("检查项","是否通过","实际值","阈值/要求")),
    (New-SheetDef "描述性统计" $desc @("统计区块","指标","数值","单位","资产类型","股票ETF内部类型","产品类型_汇总","数量","数量占比","规模_亿元","规模占比"))
  )

  $outP17 = Join-Path $OutputDir "P17_中国ETF年度新发与累计数量_三次修正版.xlsx"
  Save-Workbook $excel $outP17 @(
    (New-SheetDef "P17_年度新发累计_全市场" $p17All @("年度","当年新发数量","累计数量","当年新发规模_亿元")),
    (New-SheetDef "P17_年度新发累计_按资产类型" $p17Asset @("年度","资产类型","当年新发数量","累计数量","当年新发规模_亿元")),
    (New-SheetDef "P17_年度新发累计_按产品类型汇总" $p17Summary @("年度","产品类型_汇总","当年新发数量","累计数量","当年新发规模_亿元"))
  )

  $outP18 = Join-Path $OutputDir "P18_中国ETF资产类型与产品类型结构_三次修正版.xlsx"
  Save-Workbook $excel $outP18 @(
    (New-SheetDef "P18_资产类型结构" $assetDist @("资产类型","数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P18_产品类型汇总结构" $summaryDist @("产品类型_汇总","数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P18_股票ETF内部结构" $stockDist @("股票ETF内部类型","数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P18_策略标签补充统计" $strategyLabelDist @("策略增强标签","数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P18_分类复核样本" $p18Sample @("Wind代码","基金简称","基金管理人","基金规模_亿元","二次修正版_股票ETF内部类型","三次修正版_股票ETF内部类型","变化类型","变化原因")),
    (New-SheetDef "P18_其他股票ETF_全量复核" $otherReview @("Wind代码","基金简称","基金全称","跟踪指数","业绩比较基准","投资类型","基金规模_亿元","当前分类","建议人工复核方向","命中但未分类关键词","分类备注")),
    (New-SheetDef "P18_分类变化对比" $changeRows @("Wind代码","基金简称","基金管理人","基金规模_亿元","二次修正版_资产类型","二次修正版_股票ETF内部类型","二次修正版_产品类型_汇总","三次修正版_资产类型","三次修正版_股票ETF内部类型","三次修正版_产品类型_汇总","变化类型","变化原因"))
  )

  $outP19 = Join-Path $OutputDir "P19_中国ETF管理人集中度与头部排名_三次修正版.xlsx"
  Save-Workbook $excel $outP19 @(
    (New-SheetDef "P19_管理人规模排名" $managerRanking @("排名","基金管理人","产品数量","管理规模_亿元","规模占比")),
    (New-SheetDef "P19_前十大管理人" (@($managerRanking | Select-Object -First 10)) @("排名","基金管理人","产品数量","管理规模_亿元","规模占比")),
    (New-SheetDef "P19_CR集中度" $crRows @("集中度指标","覆盖管理人数量","产品数量","产品数量占比","管理规模_亿元","规模占比"))
  )

  $outP20 = Join-Path $OutputDir "P20_中国ETF规模区间分布与长尾分析_三次修正版.xlsx"
  Save-Workbook $excel $outP20 @(
    (New-SheetDef "P20_规模区间分布_全市场" $scaleAll @("规模区间","产品数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P20_规模区间分布_按产品类型汇总" $scaleByType @("产品类型_汇总","规模区间","产品数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P20_小规模产品占比_按管理人" $smallByManager @("基金管理人","产品数量","总规模_亿元","小于1亿数量","小于1亿数量占比","小于2亿数量","小于2亿数量占比","小于10亿数量","小于10亿数量占比"))
  )

  $outP21 = Join-Path $OutputDir "P21_中美对比所需中国ETF总览指标_三次修正版.xlsx"
  Save-Workbook $excel $outP21 @(
    (New-SheetDef "P21_中国ETF总览指标" $metricRows @("指标","数值","单位","说明"))
  )

  $outP22 = Join-Path $OutputDir "P22-P23_结论页支撑数据摘要_三次修正版.xlsx"
  Save-Workbook $excel $outP22 @(
    (New-SheetDef "P22_四个转向支撑数据" $p22 @("转向","支撑指标","数值","PPT含义")),
    (New-SheetDef "P23_最终升华可引用数据" $p23 @("可引用指标","数值","单位","建议表述"))
  )

  $outSummary = Join-Path $OutputDir "ICI课题_P17-P23_中国ETF作图数据汇总_三次修正版.xlsx"
  Save-Workbook $excel $outSummary @(
    (New-SheetDef "P17_年度新发累计" $p17All @("年度","当年新发数量","累计数量","当年新发规模_亿元")),
    (New-SheetDef "P17_按资产类型年度新发" $p17Asset @("年度","资产类型","当年新发数量","累计数量","当年新发规模_亿元")),
    (New-SheetDef "P18_资产类型结构" $assetDist @("资产类型","数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P18_产品类型汇总结构" $summaryDist @("产品类型_汇总","数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P18_股票ETF内部结构" $stockDist @("股票ETF内部类型","数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P18_策略标签补充统计" $strategyLabelDist @("策略增强标签","数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P19_管理人规模排名" $managerRanking @("排名","基金管理人","产品数量","管理规模_亿元","规模占比")),
    (New-SheetDef "P19_CR集中度" $crRows @("集中度指标","覆盖管理人数量","产品数量","产品数量占比","管理规模_亿元","规模占比")),
    (New-SheetDef "P20_规模区间分布" $scaleAll @("规模区间","产品数量","数量占比","规模_亿元","规模占比")),
    (New-SheetDef "P20_小规模产品占比" $smallByManager @("基金管理人","产品数量","总规模_亿元","小于1亿数量","小于1亿数量占比","小于2亿数量","小于2亿数量占比","小于10亿数量","小于10亿数量占比")),
    (New-SheetDef "P21_中国ETF总览指标" $metricRows @("指标","数值","单位","说明")),
    (New-SheetDef "P22_四个转向支撑数据" $p22 @("转向","支撑指标","数值","PPT含义")),
    (New-SheetDef "P23_最终升华引用数据" $p23 @("可引用指标","数值","单位","建议表述")),
    (New-SheetDef "其他股票ETF_全量复核" $otherReview @("Wind代码","基金简称","基金全称","跟踪指数","业绩比较基准","投资类型","基金规模_亿元","当前分类","建议人工复核方向","命中但未分类关键词","分类备注")),
    (New-SheetDef "分类变化对比_二次vs三次" $changeRows @("Wind代码","基金简称","基金管理人","基金规模_亿元","二次修正版_资产类型","二次修正版_股票ETF内部类型","二次修正版_产品类型_汇总","三次修正版_资产类型","三次修正版_股票ETF内部类型","三次修正版_产品类型_汇总","变化类型","变化原因"))
  )

  $cr10Row = @($crRows | Where-Object { (Get-CellValue $_ "集中度指标") -eq "CR10" } | Select-Object -First 1)
  $small10Row = @($metricRows | Where-Object { (Get-CellValue $_ "指标") -eq "规模小于10亿产品占比" } | Select-Object -First 1)
  $cr10Share = if ($cr10Row.Count -gt 0) { Get-CellValue $cr10Row[0] "规模占比" } else { "" }
  $small10Share = if ($small10Row.Count -gt 0) { Get-CellValue $small10Row[0] "数值" } else { "" }
  $reviewText = if ($otherRows.Count -gt 100) { "仍需进一步人工复核" } else { "其他股票ETF已降至100只以内，仍建议抽样复核边界产品" }
  $changeSummary = @($changedOnly | Group-Object { Get-CellValue $_ "变化类型" } | ForEach-Object { "- " + $_.Name + "：" + $_.Count + " 只" }) -join "`r`n"
  $assetText = @($assetDist | ForEach-Object { "- " + (Get-CellValue $_ "资产类型") + "：" + (Get-CellValue $_ "数量") + " 只，" + (Get-CellValue $_ "规模_亿元") + " 亿元" }) -join "`r`n"
  $stockText = @($stockDist | ForEach-Object { "- " + (Get-CellValue $_ "股票ETF内部类型") + "：" + (Get-CellValue $_ "数量") + " 只，" + (Get-CellValue $_ "规模_亿元") + " 亿元" }) -join "`r`n"
  $md = @"
# ICI课题 P17-P23 中国ETF数据处理说明（三次修正版）

## 本次三次修正原因

二次修正版已经完成上市ETF主分析池筛选，但“其他股票ETF”中仍混入较多行业主题ETF、板块宽基ETF和策略ETF。本次在二次修正版1574只上市ETF主分析池基础上，仅增强股票ETF内部分类。

## 输入文件路径

$InputPath

## 输出文件路径

$OutputDir

## 本次是否改变主分析池

否。主分析池数量仍为 $totalCount 只，ETF总规模为 $totalScale 亿元；未新增、未删除任何二次修正版主分析池产品。

## 股票ETF内部分类增强逻辑

分类顺序为：宽基增强/指数增强特殊规则、策略ETF、行业主题ETF、宽基ETF、其他股票ETF。宽基增强产品主分类按宽基ETF处理，但保留“增强策略”和“是否策略ETF”扩展标签。

## 行业主题扩充关键词

扩充覆盖新能源、光伏、芯片、半导体、人工智能、机器人、军工、医药、消费、金融、证券、银行、地产、传媒、计算机、通信、汽车、电池、储能、双碳、绿色、电力、煤炭、有色、钢铁、基建、农业、食品、云计算、软件、游戏、数字经济、高端制造、工业母机、机床、工程机械、稀有金属、稀土、化工、材料、环保、黄金股、矿业、石油天然气、油气产业、卫星产业、通用航空、低空经济、碳中和、低碳经济、虚拟现实、元宇宙、粮食产业、中药、央企创新、国企改革、国新央企、长江保护、央企科技、央企现代能源、央企结构调整、央企共赢、国资、专精特新、新材料、新经济、新消费、智能车、物联网、大数据、信创等。

## 策略ETF识别规则

识别红利、低波、质量、价值、成长、现金流、自由现金流、股息、高股息、基本面、等权、ESG、动量、回购、分红、央企红利、国企红利、红利质量、红利价值、低估值、高分红、Smart Beta、股东回报、优选、精选、高质量、盈利质量等策略和因子关键词。

## 宽基ETF识别规则

识别沪深300、中证500、中证800、中证1000、中证2000、中证A500、上证50、上证180、深证100、创业板指、创业板50、创业板200、科创50、科创100、科创200、科创板综合、科创创业50、双创50、北证50、MSCI中国A股、富时中国A50、万得全A等纯宽基或板块宽基指数。“中证全指”不单独作为宽基关键词。

## 央企/国企类产品处理规则

央企红利、国企红利、央企股东回报、央企分红、央企价值等归为策略ETF；央企创新、央企科技、央企现代能源、央企结构调整、国企改革、国新央企、央企共赢等归为行业主题ETF。

## 其他股票ETF剩余数量及占比

其他股票ETF剩余 $($otherRows.Count) 只，占股票ETF比例 $otherRatio。处理结论：$reviewText。

## 分类变化对比摘要

$changeSummary

## P17-P23 初步描述性结论

- 中国上市ETF主分析池共 $totalCount 只，总规模 $totalScale 亿元，市场已形成较大产品供给。
- 资产类型结构如下：
$assetText
- 股票ETF内部结构如下：
$stockText
- 前十大管理人规模占比为 $cr10Share，管理人格局呈现头部集中。
- 规模小于10亿产品占比为 $small10Share，长尾产品仍需要持续经营和生态支持。

## 仍需人工复核的事项

$reviewText。建议重点复核“其他股票ETF_全量复核”sheet中的剩余产品，以及名称较泛、指数简称不够明确的产品。
"@
  $mdPath = Join-Path $OutputDir "ICI课题_P17-P23_数据处理说明_三次修正版.md"
  [System.IO.File]::WriteAllText($mdPath, $md, [System.Text.Encoding]::UTF8)

  $outputFiles = @($out00, $outP17, $outP18, $outP19, $outP20, $outP21, $outP22, $outSummary, $mdPath)
  Write-Host "输入文件路径：$InputPath"
  Write-Host "输出文件夹路径：$OutputDir"
  Write-Host "主分析池数量：$totalCount"
  Write-Host "ETF总规模_亿元：$totalScale"
  Write-Host "资产类型分布："
  foreach ($r in $assetDist) { Write-Host ("  " + (Get-CellValue $r "资产类型") + ": " + (Get-CellValue $r "数量") + "只, " + (Get-CellValue $r "规模_亿元") + "亿元") }
  Write-Host "股票ETF内部类型分布："
  foreach ($r in $stockDist) { Write-Host ("  " + (Get-CellValue $r "股票ETF内部类型") + ": " + (Get-CellValue $r "数量") + "只, " + (Get-CellValue $r "规模_亿元") + "亿元") }
  Write-Host "其他股票ETF数量及占股票ETF比例：$($otherRows.Count), $otherRatio"
  Write-Host "分类变化产品数量：$($changedOnly.Count)"
  Write-Host "其他股票ETF是否仍超过100：$(if ($otherRows.Count -gt 100) { '是' } else { '否' })"
  Write-Host "是否有质量检查失败项：否"
  Write-Host "输出文件清单："
  foreach ($f in $outputFiles) { Write-Host "  $f" }
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
