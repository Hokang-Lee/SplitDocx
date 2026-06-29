# split_docx_final.ps1
# 機能:
# - セクション分割
# - 先頭インデント除去 + (番号) 先頭付与（Mode2/3）
# - 元文書の開始ページ番号を測定（ブックマーク + Word COM）
# - フッター: 「P」+ PAGE フィールド（重複ラベル自動削除）
# - 各分割DOCXの開始ページ番号を sectPr/pgNumType@start で設定
# - PDF変換(任意)

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$OutputDir = "",
    [int]$FirstPageNumber = 0,
    [ValidateSet("OriginalPage","SectionNumber")]
    [string]$FooterSource = "OriginalPage"  # OriginalPage: 元文書の開始ページ番号 / SectionNumber: セクション番号
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Namespaces
$NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Assemblies
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------- IO helpers ----------
function Load-XmlDocument([string]$Path) {
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.Load($Path)
    return $xml
}

function Save-XmlDocumentUtf8([xml]$XmlDoc, [string]$Path) {
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false
    $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
    try { $XmlDoc.Save($writer) } finally { $writer.Close() }
}

function Expand-Docx([string]$DocxPath, [string]$DestDir) {
    if (Test-Path -LiteralPath $DestDir) { Remove-Item -LiteralPath $DestDir -Recurse -Force }
    [System.IO.Compression.ZipFile]::ExtractToDirectory($DocxPath, $DestDir)
}

function Compress-Docx([string]$SrcDir, [string]$DocxPath) {
    $stream = $null; $zip = $null
    try {
        if (Test-Path -LiteralPath $DocxPath) { Remove-Item -LiteralPath $DocxPath -Force }
        $parent = Split-Path -Parent $DocxPath
        if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }

        $stream = [System.IO.File]::Open($DocxPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite)
        $zip = New-Object System.IO.Compression.ZipArchive(
            $stream,
            [System.IO.Compression.ZipArchiveMode]::Create,
            $false,
            [System.Text.Encoding]::UTF8
        )

        $ctName = "[Content_Types].xml"
        $ctFile = Join-Path $SrcDir $ctName
        if (Test-Path -LiteralPath $ctFile) {
            $entry = $zip.CreateEntry($ctName, [System.IO.Compression.CompressionLevel]::Optimal)
            $es = $entry.Open()
            try {
                $bytes = [System.IO.File]::ReadAllBytes($ctFile)
                $es.Write($bytes, 0, $bytes.Length)
            } finally { $es.Close() }
        }

        Get-ChildItem -LiteralPath $SrcDir -Recurse -File |
            Where-Object { $_.Name -ne $ctName } |
            ForEach-Object {
                $rel = $_.FullName.Substring($SrcDir.Length).TrimStart([char]92,[char]47).Replace([char]92,[char]47)
                $entry = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
                $es = $entry.Open()
                try {
                    $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
                    $es.Write($bytes, 0, $bytes.Length)
                } finally { $es.Close() }
            }
    }
    finally {
        if ($zip) { $zip.Dispose() }
        if ($stream) { $stream.Dispose() }
    }
}

# ---------- WordprocessingML helpers ----------
function Get-WBody([xml]$DocXml) {
    foreach ($child in $DocXml.DocumentElement.ChildNodes) {
        if ($child.LocalName -eq "body") { return $child }
    }
    throw "w:body not found"
}

function Get-ParaText([System.Xml.XmlNode]$Para) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $Para.GetElementsByTagName("t", $NS_W)) {
        if ($t.InnerText) { [void]$sb.Append($t.InnerText) }
    }
    return $sb.ToString()
}

function Get-RunInfo([System.Xml.XmlNode]$Run) {
    $text = ""; $color = ""; $themeColor = ""; $italic = $false
    foreach ($rc in $Run.ChildNodes) {
        if ($rc.LocalName -eq "t") {
            $text += $rc.InnerText
        }
        elseif ($rc.LocalName -eq "rPr") {
            foreach ($rp in $rc.ChildNodes) {
                if     ($rp.LocalName -eq "color") { $color = $rp.GetAttribute("val",$NS_W); $themeColor = $rp.GetAttribute("themeColor",$NS_W) }
                elseif ($rp.LocalName -eq "i")     { $italic = $true }
            }
        }
    }
    return @{ Text=$text; Color=$color; ThemeColor=$themeColor; Italic=$italic }
}

function Get-ParagraphMeta([System.Xml.XmlNode]$Para) {
    $styleId=$null; $ilvl=$null; $numId=$null
    $pPr = $null
    foreach ($c in $Para.ChildNodes) { if ($c.LocalName -eq "pPr") { $pPr = $c; break } }
    if ($pPr -ne $null) {
        foreach ($c in $pPr.ChildNodes) {
            if     ($c.LocalName -eq "pStyle") { $styleId = $c.GetAttribute("val",$NS_W) }
            elseif ($c.LocalName -eq "numPr")  {
                foreach ($n in $c.ChildNodes) {
                    if     ($n.LocalName -eq "ilvl")  { $ilvl  = $n.GetAttribute("val",$NS_W) }
                    elseif ($n.LocalName -eq "numId") { $numId = $n.GetAttribute("val",$NS_W) }
                }
            }
        }
    }
    return @{ StyleId=$styleId; Ilvl=$ilvl; NumId=$numId }
}

function Test-BlueItalicNumberPara([System.Xml.XmlNode]$Para) {
    foreach ($r in $Para.ChildNodes) {
        if ($r.LocalName -ne "r") { continue }
        $info = Get-RunInfo $r
        $txt = ("" + $info.Text).Trim()
        if (-not $txt) { continue }
        if (-not $info.Italic) { continue }

        $looksLikeNumber = $false
        if ($txt -match '^$?\d+$?$') { $looksLikeNumber = $true }
        if (-not $looksLikeNumber) { continue }

        $isBlueish = $false
        if ($info.Color -match '^(4F81BD|5B9BD5|00B0F0|9CC2E5|8DB4E2|95B3D7|7F7FFF)$') { $isBlueish = $true }
        if (-not $isBlueish -and $info.ThemeColor) { $isBlueish = $true }

        if ($isBlueish) { return $true }
    }
    return $false
}

function Get-ListParagraphStyleId([string]$UnpackedDir) {
    $stylesPath = Join-Path $UnpackedDir "word\styles.xml"
    if (-not (Test-Path -LiteralPath $stylesPath)) { return $null }
    $stylesXml = Load-XmlDocument $stylesPath

    foreach ($style in $stylesXml.DocumentElement.GetElementsByTagName("style",$NS_W)) {
        $nameEl = $null
        foreach ($c in $style.ChildNodes) { if ($c.LocalName -eq "name") { $nameEl = $c; break } }
        if ($nameEl -ne $null) {
            $nameVal = $nameEl.GetAttribute("val",$NS_W)
            if ($nameVal -eq "List Paragraph") {
                return $style.GetAttribute("styleId",$NS_W)
            }
        }
    }
    return $null
}

function Debug-Paragraphs([xml]$DocXml, [int]$Max = 120) {
    $body = Get-WBody $DocXml
    $children = @($body.ChildNodes)
    Write-Host "  -> Debug paragraph scan start" -ForegroundColor DarkGray

    for ($i = 0; $i -lt [Math]::Min($children.Count, $Max); $i++) {
        $node = $children[$i]
        if ($node.LocalName -ne "p") { continue }
        $text = (Get-ParaText $node).Trim()
        $meta = Get-ParagraphMeta $node
        if ($text -match '\d' -or $meta.NumId -or $meta.StyleId) {
            Write-Host ("     para[{0}] text='{1}' style={2} ilvl={3} numId={4}" -f $i, $text, $meta.StyleId, $meta.Ilvl, $meta.NumId) -ForegroundColor DarkGray
            foreach ($r in $node.ChildNodes) {
                if ($r.LocalName -ne "r") { continue }
                $runXml = $r.InnerXml
                if ($runXml -match 'w:tab|w:br|w:cr' -or (Get-RunInfo $r).Text) {
                    Write-Host ("         run xml='{0}'" -f $runXml) -ForegroundColor DarkGray
                }
            }
        }
    }
    Write-Host "  -> Debug paragraph scan end" -ForegroundColor DarkGray
}

# ---------- Paragraph normalization & numbering ----------
function Normalize-ParagraphIndent([xml]$OwnerDoc, [System.Xml.XmlNode]$Para) {
    # pPr を取得 or 生成
    $pPr = $null
    foreach ($c in $Para.ChildNodes) { if ($c.LocalName -eq "pPr") { $pPr = $c; break } }
    if ($pPr -eq $null) {
        $pPr = $OwnerDoc.CreateElement("w","pPr",$NS_W)
        if ($Para.HasChildNodes) { [void]$Para.InsertBefore($pPr,$Para.FirstChild) } else { [void]$Para.AppendChild($pPr) }
    }

    # 不要プロパティの除去
    $removeList = @()
    foreach ($c in $pPr.ChildNodes) {
        if ($c.LocalName -in @("numPr","ind","tabs","pStyle")) { $removeList += $c }
    }
    foreach ($n in $removeList) { [void]$pPr.RemoveChild($n) }

    # ind を 0 明示
    $ind = $OwnerDoc.CreateElement("w","ind",$NS_W)
    $null = $ind.SetAttribute("left",$NS_W,"0")
    $null = $ind.SetAttribute("right",$NS_W,"0")
    $null = $ind.SetAttribute("firstLine",$NS_W,"0")
    $null = $ind.SetAttribute("hanging",$NS_W,"0")
    [void]$pPr.AppendChild($ind)
}

function Remove-LeadingEmptyRuns([System.Xml.XmlNode]$Para) {
    while ($true) {
        $firstContent = $null
        foreach ($c in $Para.ChildNodes) {
            if ($c.LocalName -eq "pPr") { continue }
            $firstContent = $c; break
        }
        if ($firstContent -eq $null) { break }

        $remove = $false
        if ($firstContent.LocalName -eq "r") {
            $hasVisibleText = $false
            $hasOnlyIgnorable = $true
            foreach ($rc in $firstContent.ChildNodes) {
                if     ($rc.LocalName -eq "t") { if ($rc.InnerText -ne "") { $hasVisibleText = $true; $hasOnlyIgnorable = $false } }
                elseif ($rc.LocalName -in @("tab","br","cr")) { }      # 先頭の tab/br は無視扱い
                elseif ($rc.LocalName -eq "rPr") { }                   # 書式だけ
                else { $hasOnlyIgnorable = $false }
            }
            if (-not $hasVisibleText -and $hasOnlyIgnorable) { $remove = $true }
        }

        if ($remove) { [void]$Para.RemoveChild($firstContent); continue }
        break
    }
}

function Remove-LeadingTabsFromFirstTextRun([System.Xml.XmlNode]$Para) {
    foreach ($c in $Para.ChildNodes) {
        if ($c.LocalName -eq "pPr") { continue }
        if ($c.LocalName -ne "r") { break }

        $removeNodes = @()
        foreach ($rc in $c.ChildNodes) {
            if     ($rc.LocalName -eq "t")  { if ($rc.InnerText -ne "") { break } }
            elseif ($rc.LocalName -in @("tab","br","cr")) { $removeNodes += $rc }
            elseif ($rc.LocalName -eq "rPr") { continue }
            else { break }
        }
        foreach ($n in $removeNodes) { [void]$c.RemoveChild($n) }
        break
    }
}

function Prepend-PlainNumberText([xml]$OwnerDoc, [System.Xml.XmlNode]$Para, [int]$Number) {
    # "(番号)" を Meiryo UI / 薄い青 / イタリック / 12pt で先頭に挿入（スペースなし）
    $r  = $OwnerDoc.CreateElement("w","r",$NS_W)
    $rPr= $OwnerDoc.CreateElement("w","rPr",$NS_W)

    $rFonts = $OwnerDoc.CreateElement("w","rFonts",$NS_W)
    $null = $rFonts.SetAttribute("ascii",$NS_W,"Meiryo UI")
    $null = $rFonts.SetAttribute("hAnsi",$NS_W,"Meiryo UI")
    $null = $rFonts.SetAttribute("eastAsia",$NS_W,"Meiryo UI")
    $null = $rFonts.SetAttribute("cs",$NS_W,"Meiryo UI")

    $color = $OwnerDoc.CreateElement("w","color",$NS_W); $null = $color.SetAttribute("val",$NS_W,"5B9BD5")
    $iTag  = $OwnerDoc.CreateElement("w","i",$NS_W)
    $iCs   = $OwnerDoc.CreateElement("w","iCs",$NS_W)
    $sz    = $OwnerDoc.CreateElement("w","sz",$NS_W);   $null = $sz.SetAttribute("val",$NS_W,"24")
    $szCs  = $OwnerDoc.CreateElement("w","szCs",$NS_W); $null = $szCs.SetAttribute("val",$NS_W,"24")

    [void]$rPr.AppendChild($rFonts)
    [void]$rPr.AppendChild($color)
    [void]$rPr.AppendChild($iTag)
    [void]$rPr.AppendChild($iCs)
    [void]$rPr.AppendChild($sz)
    [void]$rPr.AppendChild($szCs)

    $t = $OwnerDoc.CreateElement("w","t",$NS_W)
    $t.InnerText = "(" + $Number + ")"

    [void]$r.AppendChild($rPr)
    [void]$r.AppendChild($t)

    $insertBefore = $null
    foreach ($c in $Para.ChildNodes) {
        if ($c.LocalName -ne "pPr") { $insertBefore = $c; break }
    }
    if ($insertBefore) { [void]$Para.InsertBefore($r,$insertBefore) } else { [void]$Para.AppendChild($r) }
}

# ---------- Section detection ----------
function Get-SectionParas([xml]$DocXml, [string]$ListParaStyleId, [int]$StartNumber) {
    # 段落配列を取得
    $body = Get-WBody $DocXml
    $children = @($body.ChildNodes)

    # 結果: @{ Index=<paraIndex>; Number=<sectionNumber>; Mode=<"plainText"|"blueItalic"|"numberedParagraph"> }
    $result = [System.Collections.Generic.List[hashtable]]::new()

    # Mode1: 段落テキストの先頭が "(N)" のプレーンテキスト
    for ($i = 0; $i -lt $children.Count; $i++) {
        $node = $children[$i]
        if ($node.LocalName -ne "p") { continue }

        $text = (Get-ParaText $node).Trim()
        if ($text.Length -ge 3 -and $text[0] -eq '(') {
            $closePos = $text.IndexOf(')')
            if ($closePos -gt 1) {
                $numText = $text.Substring(1, $closePos - 1).Trim()
                $num = 0
                if ([int]::TryParse($numText, [ref]$num)) {
                    $result.Add(@{ Index=$i; Number=$num; Mode="plainText" }) | Out-Null
                }
            }
        }
    }
    if ($result.Count -gt 0) {
        Write-Host "  -> Mode1: plain text (N)" -ForegroundColor DarkGray
        return $result
    }

    # Mode2: 青系かつイタリックの単独(番号)ラン
    $seqNum = $StartNumber
    for ($i = 0; $i -lt $children.Count; $i++) {
        $node = $children[$i]
        if ($node.LocalName -ne "p") { continue }
        if (Test-BlueItalicNumberPara $node) {
            $result.Add(@{ Index=$i; Number=$seqNum; Mode="blueItalic" }) | Out-Null
            $seqNum++
        }
    }
    if ($result.Count -gt 0) {
        Write-Host ("  -> Mode2: blue italic number (start=" + $StartNumber + ")") -ForegroundColor DarkGray
        return $result
    }

    # Mode3: 番号付き段落のフォールバック（レベル0/1、スタイルが List Paragraph の場合を優先）
    $seqNum = $StartNumber
    for ($i = 0; $i -lt $children.Count; $i++) {
        $node = $children[$i]
        if ($node.LocalName -ne "p") { continue }

        $meta = Get-ParagraphMeta $node
        if (-not $meta.NumId) { continue }

        $levelOk = $true
        if ($meta.Ilvl -ne $null -and $meta.Ilvl -ne "") {
            $levelOk = @("0","1") -contains $meta.Ilvl
        }
        if (-not $levelOk) { continue }

        $styleOk = $true
        if ($ListParaStyleId -and $meta.StyleId) {
            $styleOk = ($meta.StyleId -eq $ListParaStyleId)
        }
        if (-not $styleOk) { continue }

        $result.Add(@{ Index=$i; Number=$seqNum; Mode="numberedParagraph" }) | Out-Null
        $seqNum++
    }
    if ($result.Count -gt 0) {
        Write-Host ("  -> Mode3: numbered paragraph fallback (start=" + $StartNumber + ")") -ForegroundColor DarkGray
    }

    return $result
}

# ---------- Original start-page measurement & footer shaping ----------
# ブックマークID最大値
function Get-MaxBookmarkId([xml]$DocXml) {
    $max = 0
    foreach ($b in $DocXml.GetElementsByTagName("bookmarkStart",$NS_W)) {
        $idStr = $b.GetAttribute("id",$NS_W)
        if ($idStr -and ($idStr -as [int]) -ne $null) {
            $id = [int]$idStr
            if ($id -gt $max) { $max = $id }
        }
    }
    return $max
}

# セクション先頭段落に一時ブックマークを挿入
function Inject-SectionBookmarks([string]$WorkDir, [System.Collections.Generic.List[hashtable]]$Sections) {
    $docPath = Join-Path $WorkDir "word\document.xml"
    $xml = Load-XmlDocument $docPath
    $body = Get-WBody $xml
    $children = @($body.ChildNodes)

    $nextId = (Get-MaxBookmarkId -DocXml $xml) + 1000
    $bmNames = @{}

    foreach ($sec in $Sections) {
        $idx = $sec.Index
        if ($idx -lt 0 -or $idx -ge $children.Count) { continue }
        $p = $children[$idx]
        if ($p.LocalName -ne "p") { continue }

        $bmName = ("AIspltSec_{0:D3}" -f $sec.Number)

        $bmStart = $xml.CreateElement("w","bookmarkStart",$NS_W)
        $null = $bmStart.SetAttribute("id",$NS_W,($nextId.ToString()))
        $null = $bmStart.SetAttribute("name",$NS_W,$bmName)

        $bmEnd = $xml.CreateElement("w","bookmarkEnd",$NS_W)
        $null = $bmEnd.SetAttribute("id",$NS_W,($nextId.ToString()))

        $insertBefore = $null
        foreach ($c in $p.ChildNodes) { if ($c.LocalName -ne "pPr") { $insertBefore = $c; break } }

        if ($insertBefore) {
            [void]$p.InsertBefore($bmStart,$insertBefore)
            [void]$p.InsertBefore($bmEnd,$insertBefore)
        } else {
            [void]$p.AppendChild($bmStart)
            [void]$p.AppendChild($bmEnd)
        }

        $bmNames[$sec.Number] = $bmName
        $nextId++
    }

    Save-XmlDocumentUtf8 -XmlDoc $xml -Path $docPath
    return $bmNames
}

function Get-BookmarkPageNumbers([string]$DocxPath, [string[]]$BookmarkNames) {
    $result = @{}
    $word = $null
    $doc = $null

    # Word WdInformation constants
    # 3  = wdActiveEndPageNumber
    # 4  = wdNumberOfPagesInDocument
    # 1  = wdActiveEndAdjustedPageNumber
    # 2  = wdActiveEndSectionNumber
    $wdActiveEndPageNumber = 3
    $wdActiveEndAdjustedPageNumber = 1
    $wdActiveEndSectionNumber = 2

    try {
        $DocxPath = [System.IO.Path]::GetFullPath($DocxPath)
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        $doc = $word.Documents.Open($DocxPath, $false, $true, $false)

        foreach ($bmName in $BookmarkNames) {
            try {
                if (-not $doc.Bookmarks.Exists($bmName)) {
                    Write-Host ("Bookmark not found: " + $bmName) -ForegroundColor Yellow
                    continue
                }

                $range = $doc.Bookmarks.Item($bmName).Range

                $rawPage = [int]$range.Information($wdActiveEndPageNumber)
                $adjPage = [int]$range.Information($wdActiveEndAdjustedPageNumber)
                $secNo   = [int]$range.Information($wdActiveEndSectionNumber)

                Write-Host ("Bookmark " + $bmName + " raw=" + $rawPage + " adjusted=" + $adjPage + " section=" + $secNo) -ForegroundColor DarkGray

                # adjusted page を優先
                if ($adjPage -gt 0) {
                    $result[$bmName] = $adjPage
                } elseif ($rawPage -gt 0) {
                    $result[$bmName] = $rawPage
                }
            }
            catch {
                Write-Host ("Bookmark read failed: " + $bmName + " : " + $_.Exception.Message) -ForegroundColor Yellow
            }
        }

        $doc.Close($false)
        $doc = $null
        $word.Quit()
        $word = $null
    }
    finally {
        try { if ($doc) { $doc.Close($false) } } catch {}
        try { if ($word) { $word.Quit() } } catch {}
        try { if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } } catch {}
        try { if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } } catch {}
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    return $result
}

# sectPr に開始ページ番号を設定
function Set-SectionPageStart([string]$SectionDir, [int]$Start) {
    $docPath = Join-Path $SectionDir "word\document.xml"
    if (-not (Test-Path -LiteralPath $docPath)) { return }

    $xml = Load-XmlDocument $docPath
    $body = Get-WBody $xml
    if ($body -eq $null) { return }

    $sectPr = $null

    # 1) body直下の sectPr を優先
    foreach ($child in $body.ChildNodes) {
        if ($child.LocalName -eq "sectPr") {
            $sectPr = $child
        }
    }

    # 2) なければ最後の段落の pPr/sectPr を探す
    if ($sectPr -eq $null) {
        for ($i = $body.ChildNodes.Count - 1; $i -ge 0; $i--) {
            $node = $body.ChildNodes[$i]
            if ($node.LocalName -ne "p") { continue }

            $pPr = $null
            foreach ($c in $node.ChildNodes) {
                if ($c.LocalName -eq "pPr") {
                    $pPr = $c
                    break
                }
            }
            if ($pPr -eq $null) { continue }

            foreach ($c2 in $pPr.ChildNodes) {
                if ($c2.LocalName -eq "sectPr") {
                    $sectPr = $c2
                    break
                }
            }
            if ($sectPr -ne $null) { break }
        }
    }

    # 3) それでもなければ body直下に新規作成
    if ($sectPr -eq $null) {
        $sectPr = $xml.CreateElement("w","sectPr",$NS_W)
        [void]$body.AppendChild($sectPr)
    }

    # 既存の pgNumType を削除
    $toRemove = @()
    foreach ($c in $sectPr.ChildNodes) {
        if ($c.LocalName -eq "pgNumType") {
            $toRemove += $c
        }
    }
    foreach ($n in $toRemove) {
        [void]$sectPr.RemoveChild($n)
    }

    # 新しい pgNumType を追加
    $pg = $xml.CreateElement("w","pgNumType",$NS_W)
    [void]$pg.SetAttribute("start",$NS_W,$Start.ToString())
    [void]$sectPr.AppendChild($pg)

    Save-XmlDocumentUtf8 -XmlDoc $xml -Path $docPath
}

# フッター整形の前処理: runテキスト抽出 & ラベル検出
function Get-RunPlainText([System.Xml.XmlNode]$Run) {
    if ($Run -eq $null) { return "" }
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $Run.ChildNodes) {
        if ($c.LocalName -eq "t") {
            [void]$sb.Append($c.InnerText)
        } elseif ($c.LocalName -eq "tab") {
            [void]$sb.Append("`t")
        } elseif ($c.LocalName -eq "br" -or $c.LocalName -eq "cr") {
            [void]$sb.Append(" ")
        }
    }
    return $sb.ToString()
}

function Test-PageLabelRun([System.Xml.XmlNode]$Run) {
    if ($Run -eq $null -or $Run.LocalName -ne "r") { return $false }
    $t = ((Get-RunPlainText $Run) -replace '\s+', '')
    return ($t -imatch '^(p\.?|page)$')
}

function Remove-PageLabelRunsBeforeNode([System.Xml.XmlNode]$AnchorNode) {
    if ($AnchorNode -eq $null) { return $false }
    $para = $AnchorNode.ParentNode
    if ($para -eq $null) { return $false }

    $changed = $false

    while ($true) {
        $prev = $AnchorNode.PreviousSibling
        if (-not $prev) { break }

        # run 以外は終了
        if ($prev.LocalName -ne "r") { break }

        $txt = (Get-RunPlainText $prev)
        $compact = ($txt -replace '\s+', '').ToLower()

        # 削除対象:
        # "", "p", ".", "p.", "page"
        if ($compact -eq "" -or $compact -eq "p" -or $compact -eq "." -or $compact -eq "p." -or $compact -eq "page") {
            [void]$para.RemoveChild($prev)
            $changed = $true
            continue
        }

        break
    }

    return $changed
}


# Simple field: 「P」+ PAGE 体裁へ（直前のラベルrunは除去）
function Ensure-SimpleFieldPPlusPAGE([xml]$FooterXml) {
    $changed = $false
    foreach ($fld in @($FooterXml.GetElementsByTagName("fldSimple",$NS_W))) {
        $instr = $fld.GetAttribute("instr",$NS_W)
        if ($instr -and $instr -match '(^| )PAGE( |$)') {
            $p = $fld.ParentNode; if (-not $p) { continue }
            $kids = @($p.ChildNodes)
            for ($i=0; $i -lt $kids.Count; $i++) {
                if ($kids[$i] -eq $fld) {
                    if ($i-1 -ge 0 -and (Test-PageLabelRun $kids[$i-1])) { [void]$p.RemoveChild($kids[$i-1]) }
                    $r = $FooterXml.CreateElement("w","r",$NS_W)
                    $t = $FooterXml.CreateElement("w","t",$NS_W); $t.InnerText = "P"
                    [void]$r.AppendChild($t)
                    [void]$p.InsertBefore($r, $fld)
                    $changed = $true
                    break
                }
            }
        }
    }
    return $changed
}

function Normalize-FooterPageLabel([xml]$FooterXml) {
    $changed = $false

    foreach ($para in $FooterXml.GetElementsByTagName("p",$NS_W)) {
        $nodes = @($para.ChildNodes)

        for ($i = 0; $i -lt $nodes.Count; $i++) {
            $node = $nodes[$i]
            $isPageField = $false

            # simple field
            if ($node.LocalName -eq "fldSimple") {
                $instr = $node.GetAttribute("instr",$NS_W)
                if ($instr -and $instr -match '(^| )PAGE( |$)') {
                    $isPageField = $true
                }
            }

            # complex field
            if (-not $isPageField -and $node.LocalName -eq "r") {
                $isBegin = $false
                foreach ($rc in $node.ChildNodes) {
                    if ($rc.LocalName -eq "fldChar" -and $rc.GetAttribute("fldCharType",$NS_W) -eq "begin") {
                        $isBegin = $true
                        break
                    }
                }

                if ($isBegin) {
                    for ($j = $i + 1; $j -lt $nodes.Count; $j++) {
                        $r2 = $nodes[$j]
                        if ($r2.LocalName -ne "r") { continue }

                        foreach ($rc2 in $r2.ChildNodes) {
                            if ($rc2.LocalName -eq "instrText" -and $rc2.InnerText -match '(^| )PAGE( |$)') {
                                $isPageField = $true
                                break
                            }
                            if ($rc2.LocalName -eq "fldChar" -and $rc2.GetAttribute("fldCharType",$NS_W) -eq "end") {
                                break
                            }
                        }
                        if ($isPageField) { break }
                    }
                }
            }

            if (-not $isPageField) { continue }

            # 既存ラベルを広めに除去
            if (Remove-PageLabelRunsBeforeNode -AnchorNode $node) {
                $changed = $true
            }

            # 直前に P がなければ追加
            $prev = $node.PreviousSibling
            $hasP = $false
            if ($prev -and $prev.LocalName -eq "r") {
                $prevText = ((Get-RunPlainText $prev) -replace '\s+', '')
                if ($prevText -eq "P") {
                    $hasP = $true
                }
            }

            if (-not $hasP) {
                $r = $FooterXml.CreateElement("w","r",$NS_W)
                $t = $FooterXml.CreateElement("w","t",$NS_W)
                $t.InnerText = "P"
                [void]$r.AppendChild($t)
                [void]$para.InsertBefore($r, $node)
                $changed = $true
            }

            $nodes = @($para.ChildNodes)
        }
    }

    return $changed
}

# Complex field: begin..instrText(PAGE)..end ブロックの直前に「P」を追加（重複ラベルは削除）
function Ensure-ComplexFieldPPlusPAGE([xml]$FooterXml) {
    $changed = $false

    foreach ($para in $FooterXml.GetElementsByTagName("p",$NS_W)) {
        $nodes = @($para.ChildNodes)

        for ($i = 0; $i -lt $nodes.Count; $i++) {
            $node = $nodes[$i]
            $isBegin = $false
            $isPAGE = $false
            $endAt = -1

            if ($node.LocalName -eq "r") {
                foreach ($rc in $node.ChildNodes) {
                    if ($rc.LocalName -eq "fldChar" -and $rc.GetAttribute("fldCharType",$NS_W) -eq "begin") {
                        $isBegin = $true
                        break
                    }
                }
            }
            if (-not $isBegin) { continue }

            for ($j = $i + 1; $j -lt $nodes.Count; $j++) {
                $r2 = $nodes[$j]
                if ($r2.LocalName -ne "r") { continue }

                foreach ($rc2 in $r2.ChildNodes) {
                    if ($rc2.LocalName -eq "instrText" -and $rc2.InnerText -match '(^| )PAGE( |$)') {
                        $isPAGE = $true
                    }
                    if ($rc2.LocalName -eq "fldChar" -and $rc2.GetAttribute("fldCharType",$NS_W) -eq "end") {
                        $endAt = $j
                        break
                    }
                }
                if ($endAt -ge 0) { break }
            }

            if (-not $isPAGE) { continue }

            # すでに直前に P / p / page があるなら追加しない
            if ($i - 1 -ge 0) {
                $prevNode = $nodes[$i - 1]
                if (Test-PageLabelRun $prevNode) {
                    continue
                }
            }

            $beginRun = $nodes[$i]

            $r = $FooterXml.CreateElement("w","r",$NS_W)
            $t = $FooterXml.CreateElement("w","t",$NS_W)
            $t.InnerText = "P"
            [void]$r.AppendChild($t)
            [void]$para.InsertBefore($r, $beginRun)

            $changed = $true

            # 配列を更新して以降の走査位置を進める
            $nodes = @($para.ChildNodes)
            $i = $endAt + 1
        }
    }

    return $changed
}

# すべての footer*.xml に対して「P」+ PAGE 体裁を適用
function Fix-FooterToDynamicPPage([string]$SectionDir) {
    $dir = Join-Path $SectionDir "word"
    if (-not (Test-Path -LiteralPath $dir)) { return }

    $files = Get-ChildItem -LiteralPath $dir -File -Filter "footer*.xml" -ErrorAction SilentlyContinue
    foreach ($ff in $files) {
        $xml = Load-XmlDocument $ff.FullName
        $changed = Normalize-FooterPageLabel -FooterXml $xml
        if ($changed) {
            Save-XmlDocumentUtf8 -XmlDoc $xml -Path $ff.FullName
        }
    }
}

# ---------- Section doc builder & PDF ----------
function Save-SectionDocXml(
    [xml]$OrigDocXml,
    [int[]]$ParaIndices,
    [string]$OutPath,
    [int]$SectionNumber,
    [string]$SectionMode
) {
    $newDoc = New-Object System.Xml.XmlDocument
    $newDoc.PreserveWhitespace = $true
    $newDoc.LoadXml($OrigDocXml.OuterXml)

    $body = Get-WBody $newDoc
    $orig = @($body.ChildNodes)

    # 末尾のセクションプロパティを退避
    $sectPr = $null
    for ($i = $orig.Count - 1; $i -ge 0; $i--) {
        if ($orig[$i].LocalName -eq "sectPr") { $sectPr = $orig[$i].CloneNode($true); break }
    }

    # 中身を一旦空にして、対象段落のみコピー
    $body.RemoveAll()
    $firstParaNode = $null
    foreach ($idx in $ParaIndices) {
        if ($idx -lt 0 -or $idx -ge $orig.Count) { continue }
        if ($orig[$idx].LocalName -eq "sectPr") { continue }
        $imported = $newDoc.ImportNode($orig[$idx], $true)
        [void]$body.AppendChild($imported)
        if ($firstParaNode -eq $null -and $imported.LocalName -eq "p") { $firstParaNode = $imported }
    }

    # 先頭段落の整形と "(番号)" 付与（plainText検出時は既存の見出しを尊重して付与しない）
    if ($firstParaNode -ne $null -and $SectionMode -ne "plainText") {
        Normalize-ParagraphIndent -OwnerDoc $newDoc -Para $firstParaNode
        Remove-LeadingEmptyRuns -Para $firstParaNode
        Remove-LeadingTabsFromFirstTextRun -Para $firstParaNode
        Prepend-PlainNumberText -OwnerDoc $newDoc -Para $firstParaNode -Number $SectionNumber
    }

    if ($sectPr) { [void]$body.AppendChild($newDoc.ImportNode($sectPr, $true)) }
    Save-XmlDocumentUtf8 -XmlDoc $newDoc -Path $OutPath
}

function Convert-ToPdf([string]$DocxPath, [string]$PdfPath) {
    $word = $null; $doc = $null
    try {
        $DocxPath = [System.IO.Path]::GetFullPath($DocxPath)
        $PdfPath  = [System.IO.Path]::GetFullPath($PdfPath)
        $pdfDir = Split-Path -Parent $PdfPath
        if ($pdfDir -and -not (Test-Path -LiteralPath $pdfDir)) { New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null }
        if (Test-Path -LiteralPath $PdfPath) { Remove-Item -LiteralPath $PdfPath -Force }

        Write-Host ("    PDF src : " + $DocxPath) -ForegroundColor DarkGray
        Write-Host ("    PDF dest: " + $PdfPath) -ForegroundColor DarkGray

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false; $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($DocxPath, $false, $true, $false)
        $doc.ExportAsFixedFormat($PdfPath, 17)  # wdExportFormatPDF=17
        $doc.Close($false); $word.Quit()

        if (-not (Test-Path -LiteralPath $PdfPath)) { throw "PDF file was not created: $PdfPath" }
        return $true
    }
    catch {
        try { if ($doc) { $doc.Close($false) } } catch {}
        try { if ($word) { $word.Quit() } } catch {}
        Write-Host ("    PDF error: " + $_.Exception.Message) -ForegroundColor Yellow
        return $false
    }
    finally {
        if ($doc)  { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
        if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ---------- main (前段) ----------
if (-not [System.IO.Path]::IsPathRooted($InputFile)) { $InputFile = Join-Path (Get-Location).Path $InputFile }
if (-not (Test-Path -LiteralPath $InputFile)) { Write-Error ("File not found: " + $InputFile); exit 1 }
$InputFile = (Get-Item -LiteralPath $InputFile).FullName
if ([System.IO.Path]::GetExtension($InputFile).ToLower() -ne ".docx") { Write-Error "Not a .docx file"; exit 1 }

if ($OutputDir -eq "") { $OutputDir = Join-Path (Split-Path $InputFile -Parent) "split_output" }
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
$startNum = 1
$mStart = [System.Text.RegularExpressions.Regex]::Match($baseName, '(\d{3})-\d{3}')
if ($mStart.Success) { $startNum = [int]$mStart.Groups[1].Value }

Write-Host ("Input : " + $InputFile)
Write-Host ("Output: " + $OutputDir)
Write-Host ("Start#: " + $startNum)
Write-Host ""

$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("docx_split_" + [System.IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null

try {
    Write-Host "[1/5] Unpacking..." -ForegroundColor Cyan
    $unpackedDir = Join-Path $tmpRoot "unpacked"
    Expand-Docx -DocxPath $InputFile -DestDir $unpackedDir

    Write-Host "[2/5] Detecting section boundaries..." -ForegroundColor Cyan
    $dp = Join-Path $unpackedDir "word\document.xml"
    $docXml = Load-XmlDocument $dp
    $listParaStyleId = Get-ListParagraphStyleId -UnpackedDir $unpackedDir
    Write-Host ("  -> ListParagraph styleId: " + $listParaStyleId)
    $sections = @(Get-SectionParas -DocXml $docXml -ListParaStyleId $listParaStyleId -StartNumber $startNum)
    if ($sections.Count -eq 0) { throw "No section headers found." }
    Write-Host ("  -> " + $sections.Count + " sections found")
    foreach ($s in $sections) { Write-Host ("     (" + $s.Number + ") at para[" + $s.Index + "] mode=" + $s.Mode) }

    # 3/5: 元文書での開始ページ番号を測定（必要な場合）
    $origPageMap = @{}
    if ($FooterSource -eq "OriginalPage") {
        Write-Host "[3/5] Marking and measuring original start pages..." -ForegroundColor Cyan
        $measureDir = Join-Path $tmpRoot "measure"
        Copy-Item -LiteralPath $unpackedDir -Destination $measureDir -Recurse -Force

        $numberToBmName = Inject-SectionBookmarks -WorkDir $measureDir -Sections $sections

        $measureDocx = Join-Path $tmpRoot "measure.docx"
        Compress-Docx -SrcDir $measureDir -DocxPath $measureDocx

        $bmPageMap = Get-BookmarkPageNumbers -DocxPath $measureDocx -BookmarkNames @($numberToBmName.Values)

        $origPageMap = @{}
        foreach ($num in $numberToBmName.Keys) {
            $bm = $numberToBmName[$num]
            if ($bmPageMap.ContainsKey($bm)) {
                $origPageMap[$num] = $bmPageMap[$bm]
            }
        }

        if ($FirstPageNumber -gt 0) {
            $adjust = $FirstPageNumber - 1
            foreach ($k in @($origPageMap.Keys)) {
                $origPageMap[$k] = [int]$origPageMap[$k] + $adjust
            }
        }

        $csvPath = Join-Path $OutputDir "section_page_map.csv"
        "SectionNumber,OriginalPage" | Out-File -FilePath $csvPath -Encoding utf8
        foreach ($k in ($origPageMap.Keys | Sort-Object)) {
            "$k,$($origPageMap[$k])" | Out-File -FilePath $csvPath -Append -Encoding utf8
        }
        Write-Host ("  -> Page map saved: " + $csvPath)
    }

    # 4/5: セクションごとにDOCXを作成
    Write-Host "[4/5] Building section DOCX files..." -ForegroundColor Cyan
    $bodyChildren = @((Get-WBody $docXml).ChildNodes)

    for ($si = 0; $si -lt $sections.Count; $si++) {
        $sec = $sections[$si]
        Write-Host ("  -> Section " + [int]$sec.Number + " start") -ForegroundColor Cyan

        $startIdx = [int]$sec.Index
        if ($si -lt ($sections.Count - 1)) {
            $endIdx = [int]$sections[$si+1].Index - 1
        } else {
            $endIdx = $bodyChildren.Count - 1
        }

        Write-Host ("     indices: " + $startIdx + " .. " + $endIdx) -ForegroundColor DarkGray

        if ($startIdx -gt $endIdx) { continue }

        $indices = New-Object System.Collections.Generic.List[int]
        for ($k = $startIdx; $k -le $endIdx; $k++) { $indices.Add($k) | Out-Null }

        $sectionDir = Join-Path $tmpRoot ("sec_" + ("{0:D3}" -f [int]$sec.Number))
        Write-Host ("     copy template -> " + $sectionDir) -ForegroundColor DarkGray
        Copy-Item -LiteralPath $unpackedDir -Destination $sectionDir -Recurse -Force

        $outDocXml = Join-Path $sectionDir "word\document.xml"
        Write-Host "     save section xml" -ForegroundColor DarkGray
        Save-SectionDocXml -OrigDocXml $docXml -ParaIndices $indices.ToArray() -OutPath $outDocXml -SectionNumber ([int]$sec.Number) -SectionMode $sec.Mode

        Write-Host "     fix footer" -ForegroundColor DarkGray
        Fix-FooterToDynamicPPage -SectionDir $sectionDir

        if ($FooterSource -eq "OriginalPage" -and $origPageMap.ContainsKey([int]$sec.Number)) {
            $startPage = [int]$origPageMap[[int]$sec.Number]
        } else {
            $startPage = [int]$sec.Number
        }

        Write-Host ("     set start page = " + $startPage) -ForegroundColor DarkGray
        Set-SectionPageStart -SectionDir $sectionDir -Start $startPage

        $outDocx = Join-Path $OutputDir ("{0:D3}.docx" -f [int]$sec.Number)
        Write-Host ("     compress docx -> " + $outDocx) -ForegroundColor DarkGray
        Compress-Docx -SrcDir $sectionDir -DocxPath $outDocx
        Write-Host ("  -> DOCX done: " + $outDocx) -ForegroundColor Green

        $outPdf = [System.IO.Path]::ChangeExtension($outDocx, ".pdf")
        Write-Host ("     convert pdf -> " + $outPdf) -ForegroundColor DarkGray
        if (-not (Convert-ToPdf -DocxPath $outDocx -PdfPath $outPdf)) {
            Write-Host ("     PDF skipped or failed for: " + $outDocx) -ForegroundColor Yellow
        } else {
            Write-Host ("     PDF done: " + $outPdf) -ForegroundColor Green
        }
    }

    Write-Host ""
    Write-Host "All sections exported." -ForegroundColor Green
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
finally {
    if (Test-Path -LiteralPath $tmpRoot) {
        try { Remove-Item -LiteralPath $tmpRoot -Recurse -Force } catch {}
    }
    Write-Host ""
    Write-Host "Temp cleaned." -ForegroundColor DarkGray
    Write-Host "Done." -ForegroundColor Green
}
# --- END OF FILE ---