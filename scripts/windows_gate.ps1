[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$RepositoryRoot,
  [Parameter(Mandatory = $true)]
  [string]$EvidenceRoot
)

$ErrorActionPreference = 'Stop'
$ArtifactsRoot = Join-Path $RepositoryRoot 'artifacts'
$ScratchRoot = Join-Path $env:RUNNER_TEMP ('node-edge-log-' + [guid]::NewGuid().ToString('N'))
$ReportNames = @(
  'normalized_events.ndjson',
  'suppressed_lines.csv',
  'session_inventory.csv',
  'hourly_route_metrics.csv',
  'run_summary.json'
)

New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
New-Item -ItemType Directory -Force -Path $ScratchRoot | Out-Null

function Assert-True {
  param([bool]$Condition, [string]$Message)
  if (-not $Condition) {
    throw $Message
  }
}

function Get-Sha256 {
  param([string]$PathValue)
  return (Get-FileHash -Algorithm SHA256 -LiteralPath $PathValue).Hash.ToLowerInvariant()
}

function Get-TreeHash {
  param([string]$Root)
  $lines = Get-ChildItem -LiteralPath $Root -File -Recurse |
    ForEach-Object {
      $relative = [System.IO.Path]::GetRelativePath($Root, $_.FullName).Replace('\', '/')
      $relative + [char]0 + (Get-Sha256 $_.FullName)
    } |
    Sort-Object
  $text = [string]::Join("`n", $lines)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
  $hash = [System.Security.Cryptography.SHA256]::HashData($bytes)
  return [Convert]::ToHexString($hash).ToLowerInvariant()
}

function Get-ZipFileEntries {
  param([string]$ArchivePath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
  try {
    return @($archive.Entries |
      Where-Object { -not [string]::IsNullOrEmpty($_.Name) } |
      ForEach-Object { $_.FullName.Replace('\', '/') } |
      Sort-Object)
  } finally {
    $archive.Dispose()
  }
}

function Assert-SequenceEqual {
  param([object[]]$Actual, [object[]]$Expected, [string]$Label)
  $difference = Compare-Object -ReferenceObject $Expected -DifferenceObject $Actual
  if ($difference) {
    throw "$Label member list mismatch: $($difference | Out-String)"
  }
}

function Get-WorkbookSheetNames {
  param([string]$WorkbookPath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
  try {
    $entry = $archive.GetEntry('xl/workbook.xml')
    Assert-True ($null -ne $entry) 'workbook.xml missing'
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
      [xml]$xml = $reader.ReadToEnd()
    } finally {
      $reader.Dispose()
    }
    return @($xml.SelectNodes("//*[local-name()='sheet']") | ForEach-Object { $_.GetAttribute('name') })
  } finally {
    $archive.Dispose()
  }
}

function Get-WorkbookText {
  param([string]$WorkbookPath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
  $values = [System.Collections.Generic.List[string]]::new()
  $shared = @()
  try {
    $sharedEntry = $archive.GetEntry('xl/sharedStrings.xml')
    if ($null -ne $sharedEntry) {
      $reader = [System.IO.StreamReader]::new($sharedEntry.Open())
      try {
        [xml]$sharedXml = $reader.ReadToEnd()
      } finally {
        $reader.Dispose()
      }
      $shared = @($sharedXml.SelectNodes("//*[local-name()='si']") | ForEach-Object {
        [string]::Join('', @($_.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }))
      })
    }
    foreach ($entry in $archive.Entries) {
      if (-not $entry.FullName.StartsWith('xl/worksheets/') -or -not $entry.FullName.EndsWith('.xml')) {
        continue
      }
      $reader = [System.IO.StreamReader]::new($entry.Open())
      try {
        [xml]$xml = $reader.ReadToEnd()
      } finally {
        $reader.Dispose()
      }
      foreach ($cell in $xml.SelectNodes("//*[local-name()='c']")) {
        $type = $cell.GetAttribute('t')
        if ($type -eq 'inlineStr') {
          $parts = $cell.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }
          if ($parts.Count -gt 0) {
            $values.Add([string]::Join('', $parts))
          }
        } elseif ($type -eq 'str') {
          $node = $cell.SelectSingleNode("./*[local-name()='v']")
          if ($null -ne $node -and -not [string]::IsNullOrEmpty($node.InnerText)) {
            $values.Add($node.InnerText)
          }
        } elseif ($type -eq 's') {
          $node = $cell.SelectSingleNode("./*[local-name()='v']")
          if ($null -ne $node -and $node.InnerText -match '^\d+$') {
            $index = [int]$node.InnerText
            if ($index -lt $shared.Count) {
              $values.Add($shared[$index])
            }
          }
        }
      }
    }
  } finally {
    $archive.Dispose()
  }
  return @($values)
}

function Assert-SpecificationShape {
  param([string]$WorkbookPath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
  try {
    $entry = $archive.GetEntry('xl/worksheets/sheet1.xml')
    Assert-True ($null -ne $entry) 'specification worksheet missing'
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
      [xml]$xml = $reader.ReadToEnd()
    } finally {
      $reader.Dispose()
    }
    foreach ($cell in $xml.SelectNodes("//*[local-name()='c']")) {
      $reference = $cell.GetAttribute('r')
      Assert-True ($reference -match '^[AB]\d+$') "specification contains a value outside two columns: $reference"
    }
  } finally {
    $archive.Dispose()
  }
}

function Assert-NaturalText {
  param([string[]]$Texts, [string]$Label)
  $quoteCharacters = @(
    [char]34, [char]39, [char]96, '“', '”', '‘', '’', '＂', '＇',
    '「', '」', '『', '』', '«', '»', '‹', '›', '〝', '〞', '〟', '《', '》', '〈', '〉'
  )
  $space = '[\t \u00a0\u1680\u2000-\u200a\u202f\u205f\u3000]+'
  $han = '[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]'
  $boundary = "(?:$han$space[A-Za-z0-9]|[A-Za-z0-9]$space$han|[A-Za-z]$space[0-9]|[0-9]$space[A-Za-z])"
  $riskTerms = @('此外', '至关重要', '深入探讨', '彰显', '赋能', '无缝', '不断演变的格局', '不仅', '不只是', '值得注意的是', '专家认为', '行业报告显示', '观察者指出', '未来展望', '挑战与未来', '——')
  $processTerms = @(
    '制题', '返修', '去AI', '修改题目', '规则调整', '规则变化', 'Windows复现', 'Windows验证', 'GitHub Actions',
    'Reference', 'reference.zip', 'reference_members', 'validation', '自证', '固定控制量', '不变量',
    '连续运行', '重复运行', '两次干净运行', '两个空output目录', '双干净目录', '动态改参',
    '失败清理', '失败关闭', '失败收口', '附件哈希', '飞书回读',
    ('record' + '_id'), ('file' + '_token')
  )
  foreach ($text in $Texts) {
    foreach ($character in $quoteCharacters) {
      Assert-True (-not $text.Contains([string]$character)) "$Label contains a forbidden quote"
    }
    Assert-True (-not [regex]::IsMatch($text, $boundary)) "$Label contains a mixed boundary space"
    foreach ($term in ($riskTerms + $processTerms)) {
      Assert-True (-not $text.Contains($term)) "$Label contains forbidden term $term"
    }
  }
}

function Assert-NoLinuxArtifacts {
  param([string]$Root)
  $bannedExtensions = @('.sh', '.bash', '.zsh', '.so', '.elf', '.deb', '.rpm', '.appimage')
  foreach ($file in Get-ChildItem -LiteralPath $Root -File -Recurse) {
    $extension = $file.Extension.ToLowerInvariant()
    Assert-True (-not $bannedExtensions.Contains($extension)) "Linux or shell artifact found: $($file.FullName)"
    $stream = [System.IO.File]::OpenRead($file.FullName)
    try {
      if ($stream.Length -ge 4) {
        $bytes = [byte[]]::new(4)
        [void]$stream.Read($bytes, 0, 4)
        $isElf = $bytes[0] -eq 0x7f -and $bytes[1] -eq 0x45 -and $bytes[2] -eq 0x4c -and $bytes[3] -eq 0x46
        Assert-True (-not $isElf) "ELF binary found: $($file.FullName)"
      }
    } finally {
      $stream.Dispose()
    }
  }
}

function Expand-TaskWorkspace {
  param([string]$Name)
  $workspace = Join-Path $ScratchRoot $Name
  New-Item -ItemType Directory -Force -Path $workspace | Out-Null
  Expand-Archive -LiteralPath (Join-Path $ArtifactsRoot '输入数据包.zip') -DestinationPath $workspace
  Expand-Archive -LiteralPath (Join-Path $ArtifactsRoot 'reference.zip') -DestinationPath $workspace
  return $workspace
}

function Save-ExpectedReports {
  param([string]$Workspace)
  $expectedRoot = Join-Path $Workspace 'expected_reports'
  New-Item -ItemType Directory -Force -Path $expectedRoot | Out-Null
  foreach ($name in $ReportNames) {
    $source = Join-Path $Workspace "output/reports/$name"
    Assert-True (Test-Path -LiteralPath $source -PathType Leaf) "reference report missing: $name"
    Copy-Item -LiteralPath $source -Destination (Join-Path $expectedRoot $name)
  }
  Remove-Item -LiteralPath (Join-Path $Workspace 'output/reports') -Recurse -Force
  return $expectedRoot
}

function Get-CsvSemanticRows {
  param([string]$CsvPath)
  return @(Import-Csv -LiteralPath $CsvPath | ForEach-Object { $_ | ConvertTo-Json -Compress } | Sort-Object)
}

function Get-NdjsonSemanticRows {
  param([string]$NdjsonPath)
  return @(Get-Content -LiteralPath $NdjsonPath | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object {
    $_ | ConvertFrom-Json | ConvertTo-Json -Compress -Depth 20
  } | Sort-Object)
}

function Get-SummarySemanticRows {
  param([string]$JsonPath)
  $row = Get-Content -LiteralPath $JsonPath -Raw | ConvertFrom-Json
  return @(
    "accepted_event_count=$($row.accepted_event_count)",
    "metric_rows=$($row.metric_rows)",
    "parsed_json_lines=$($row.parsed_json_lines)",
    "raw_log_lines=$($row.raw_log_lines)",
    "session_count=$($row.session_count)",
    "status_5xx_total=$($row.status_5xx_total)",
    "suppression_count=$($row.suppression_count)",
    "total_bytes=$($row.total_bytes)"
  ) | Sort-Object
}

function Get-ReportSemantics {
  param([string]$ReportRoot)
  return [ordered]@{
    normalized_events = Get-NdjsonSemanticRows (Join-Path $ReportRoot 'normalized_events.ndjson')
    suppressed_lines = Get-CsvSemanticRows (Join-Path $ReportRoot 'suppressed_lines.csv')
    session_inventory = Get-CsvSemanticRows (Join-Path $ReportRoot 'session_inventory.csv')
    hourly_route_metrics = Get-CsvSemanticRows (Join-Path $ReportRoot 'hourly_route_metrics.csv')
    run_summary = Get-SummarySemanticRows (Join-Path $ReportRoot 'run_summary.json')
  }
}

function Assert-ReportSemantics {
  param([object]$Actual, [object]$Expected, [string]$Label)
  foreach ($key in @('normalized_events', 'suppressed_lines', 'session_inventory', 'hourly_route_metrics', 'run_summary')) {
    Assert-SequenceEqual $Actual[$key] $Expected[$key] "$Label $key"
  }
}

function Invoke-CleanRun {
  param([string]$Name)
  $workspace = Expand-TaskWorkspace $Name
  $expectedRoot = Save-ExpectedReports $workspace
  $inputRoot = Join-Path $workspace 'input_data'
  $outputRoot = Join-Path $workspace 'output'
  $beforeHash = Get-TreeHash $inputRoot
  & (Join-Path $outputRoot 'run.ps1') -InputRoot $inputRoot -OutputRoot $outputRoot | Out-Host
  $afterHash = Get-TreeHash $inputRoot
  Assert-True ($beforeHash -eq $afterHash) 'input files changed during execution'
  $actual = Get-ReportSemantics (Join-Path $outputRoot 'reports')
  $expected = Get-ReportSemantics $expectedRoot
  Assert-ReportSemantics $actual $expected $Name
  $hashes = [ordered]@{}
  foreach ($report in $ReportNames) {
    $hashes[$report] = Get-Sha256 (Join-Path $outputRoot "reports/$report")
  }
  return [ordered]@{
    directory_name = $Name
    input_tree_sha256 = $beforeHash
    report_sha256 = $hashes
    semantics = $actual
    exit_code = 0
  }
}

function Invoke-NegativeRun {
  $workspace = Expand-TaskWorkspace '缺失输入 中文 空格'
  Save-ExpectedReports $workspace | Out-Null
  $missing = Join-Path $workspace 'input_data/logs/edge_access.ndjson.gz'
  Remove-Item -LiteralPath $missing
  $outputRoot = Join-Path $workspace 'output'
  $failed = $false
  try {
    & (Join-Path $outputRoot 'run.ps1') -InputRoot (Join-Path $workspace 'input_data') -OutputRoot $outputRoot | Out-Host
  } catch {
    $failed = $true
  }
  $reportRoot = Join-Path $outputRoot 'reports'
  $residual = [System.Collections.Generic.List[string]]::new()
  if (Test-Path -LiteralPath $reportRoot) {
    foreach ($name in Get-ChildItem -LiteralPath $reportRoot -File | Select-Object -ExpandProperty Name) {
      $residual.Add($name)
    }
  }
  Assert-True $failed 'missing compressed log did not fail'
  Assert-True ($residual.Count -eq 0) 'negative run left report files'
  return [ordered]@{
    case = '缺少edge_access.ndjson.gz'
    nonzero_exit = $true
    residual_reports = $residual
    pass = $true
  }
}

function Write-Evidence {
  param([object]$Evidence)
  $json = $Evidence | ConvertTo-Json -Depth 30
  Set-Content -LiteralPath (Join-Path $EvidenceRoot 'evidence.json') -Value $json -Encoding utf8NoBOM
}

$evidence = [ordered]@{
  task_asset_id = 'node_edge_log_session_report'
  runner = 'windows-2025'
  generated_at_utc = [DateTime]::UtcNow.ToString('o')
  pass = $false
}

try {
  $version = (& node --version).Trim()
  Assert-True ($LASTEXITCODE -eq 0) 'node version command failed'
  Assert-True ($version -match '^v24\.') "unexpected Node.js version: $version"
  Set-Content -LiteralPath (Join-Path $EvidenceRoot 'software-version.txt') -Value $version -Encoding utf8NoBOM
  $evidence.software_version = $version

  $manifest = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'manifest.json') -Raw | ConvertFrom-Json
  Assert-True ($manifest.task_asset_id -eq 'node_edge_log_session_report') 'task asset id mismatch'
  $attachmentHashes = [ordered]@{}
  foreach ($name in @('输入数据包.zip', 'reference.zip', '关键标准答案.xlsx', '任务规格转化.xlsx')) {
    $path = Join-Path $ArtifactsRoot $name
    Assert-True (Test-Path -LiteralPath $path -PathType Leaf) "attachment missing: $name"
    $actualHash = Get-Sha256 $path
    $expectedHash = [string]$manifest.attachments.$name
    Assert-True ($actualHash -eq $expectedHash) "attachment hash mismatch: $name"
    $attachmentHashes[$name] = $actualHash
  }
  $evidence.attachment_sha256 = $attachmentHashes

  $inputMembers = Get-ZipFileEntries (Join-Path $ArtifactsRoot '输入数据包.zip')
  $expectedInputMembers = @(
    'input_data/README.md',
    'input_data/catalog/route_catalog.csv',
    'input_data/logs/edge_access.ndjson.gz',
    'input_data/policy/consent_windows.csv',
    'input_data/policy/privacy_policy.json',
    'input_data/rules/session_reporting_contract.md',
    'input_data/starter/process_edge_logs.mjs'
  ) | Sort-Object
  Assert-SequenceEqual $inputMembers $expectedInputMembers 'input archive'

  $referenceMembers = Get-ZipFileEntries (Join-Path $ArtifactsRoot 'reference.zip')
  $expectedReferenceMembers = @(
    'output/reports/hourly_route_metrics.csv',
    'output/reports/normalized_events.ndjson',
    'output/reports/run_summary.json',
    'output/reports/session_inventory.csv',
    'output/reports/suppressed_lines.csv',
    'output/run.ps1',
    'output/src/process_edge_logs.mjs'
  ) | Sort-Object
  Assert-SequenceEqual $referenceMembers $expectedReferenceMembers 'reference archive'
  $evidence.archive_members = [ordered]@{ input = $inputMembers; reference = $referenceMembers }

  $answerSheets = Get-WorkbookSheetNames (Join-Path $ArtifactsRoot '关键标准答案.xlsx')
  $expectedAnswerSheets = @('交付物答案清单', '固定字段答案', '固定集合答案', '固定数值答案', '允许变体答案')
  Assert-SequenceEqual $answerSheets $expectedAnswerSheets 'answer workbook sheets'
  $specSheets = Get-WorkbookSheetNames (Join-Path $ArtifactsRoot '任务规格转化.xlsx')
  Assert-SequenceEqual $specSheets @('任务规格转化') 'specification workbook sheets'
  Assert-SpecificationShape (Join-Path $ArtifactsRoot '任务规格转化.xlsx')
  $specText = Get-WorkbookText (Join-Path $ArtifactsRoot '任务规格转化.xlsx')
  Assert-True ($specText.Contains('node_edge_log_session_report')) 'name-derived task id missing from specification'
  foreach ($forbidden in @('学科', '难度', '任务名称', '任务概要', '预计工时', 'Windows验证过程', ('线上题目' + 'ID'))) {
    Assert-True (-not $specText.Contains($forbidden)) "forbidden specification field found: $forbidden"
  }
  $evidence.workbook_sheets = [ordered]@{ answer = $answerSheets; specification = $specSheets }

  $naturalTexts = [System.Collections.Generic.List[string]]::new()
  foreach ($file in Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'task') -File) {
    $naturalTexts.Add((Get-Content -LiteralPath $file.FullName -Raw))
  }
  foreach ($name in @('README.md', 'SOURCES.md')) {
    $naturalTexts.Add((Get-Content -LiteralPath (Join-Path $RepositoryRoot $name) -Raw))
  }
  foreach ($text in Get-WorkbookText (Join-Path $ArtifactsRoot '关键标准答案.xlsx')) { $naturalTexts.Add($text) }
  foreach ($text in $specText) { $naturalTexts.Add($text) }

  $scanRoot = Expand-TaskWorkspace '兼容扫描 中文 空格'
  foreach ($relative in @('input_data/README.md', 'input_data/rules/session_reporting_contract.md')) {
    $naturalTexts.Add((Get-Content -LiteralPath (Join-Path $scanRoot $relative) -Raw))
  }
  Assert-NaturalText @($naturalTexts) 'candidate-facing text'
  $answerControlTerms = @('Reference', 'reference.zip', 'reference_members', 'validation', '自证', '固定控制量', '不变量', '连续运行', '重复运行', '两次干净运行', '两个空output目录', '动态改参', '失败清理', '失败关闭', '失败收口')
  $deliveryTexts = [System.Collections.Generic.List[string]]::new()
  foreach ($relative in @('input_data/starter/process_edge_logs.mjs', 'output/src/process_edge_logs.mjs', 'output/run.ps1', 'output/reports/run_summary.json')) {
    $deliveryTexts.Add((Get-Content -LiteralPath (Join-Path $scanRoot $relative) -Raw))
  }
  foreach ($text in @($naturalTexts) + @($deliveryTexts)) {
    foreach ($term in $answerControlTerms) {
      Assert-True (-not $text.Contains($term)) "candidate-visible content contains answer-control term $term"
    }
  }
  $evidence.natural_language = [ordered]@{ source_count = $naturalTexts.Count; pass = $true }
  $evidence.answer_control_language = [ordered]@{ source_count = $naturalTexts.Count + $deliveryTexts.Count; pass = $true }
  Assert-NoLinuxArtifacts $scanRoot
  $sourceText = Get-Content -LiteralPath (Join-Path $scanRoot 'output/src/process_edge_logs.mjs') -Raw
  foreach ($requiredToken in @('createReadStream', 'createGunzip', 'readline.createInterface', 'createHmac')) {
    Assert-True ($sourceText.Contains($requiredToken)) "required Node.js behavior missing: $requiredToken"
  }
  Assert-True (-not $sourceText.Contains('child_process')) 'external process dependency found in Node.js source'
  $evidence.compatibility_scan = [ordered]@{ linux_artifacts = 0; required_node_tokens = 4; pass = $true }

  $businessRun = Invoke-CleanRun '边缘 小时报表'
  $evidence.business_run = $businessRun
  $evidence.negative = Invoke-NegativeRun
  $evidence.pass = $true
  Write-Evidence $evidence
} catch {
  $evidence.error = $_.Exception.Message
  $evidence.pass = $false
  Write-Evidence $evidence
  throw
}
