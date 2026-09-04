# Create the same embedded-task-pane workbook shape used by office-addin-dev-settings.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ManifestPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
    throw "Manifest not found: $ManifestPath"
}

$manifest = New-Object System.Xml.XmlDocument
$manifest.Load($ManifestPath)
$namespace = New-Object System.Xml.XmlNamespaceManager($manifest.NameTable)
$namespace.AddNamespace("o", "http://schemas.microsoft.com/office/appforoffice/1.1")
$manifestId = $manifest.SelectSingleNode("/o:OfficeApp/o:Id", $namespace).InnerText.Trim()
$manifestVersion = $manifest.SelectSingleNode("/o:OfficeApp/o:Version", $namespace).InnerText.Trim()
if ($manifestId -notmatch "^[0-9A-Fa-f]{8}(-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}$") {
    throw "Manifest ID is not a GUID: $manifestId"
}
if ($manifestVersion -notmatch "^\d+\.\d+\.\d+\.\d+$") {
    throw "Manifest version is not an Office four-part version: $manifestVersion"
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Add-ZipText(
    [IO.Compression.ZipArchive]$Archive,
    [string]$Path,
    [string]$Text
) {
    $entry = $Archive.CreateEntry($Path, [IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $writer = New-Object IO.StreamWriter($stream, [Text.UTF8Encoding]::new($false))
        try {
            $writer.Write($Text)
        } finally {
            $writer.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

$parent = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($OutputPath))
if (-not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue

$contentTypes = '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" /><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Override PartName="/xl/worksheets/sheet.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" /><Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml" /><Override PartName="/xl/webextensions/webextension.xml" ContentType="application/vnd.ms-office.webextension+xml" /></Types>'
$rootRelationships = '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml" Id="rWorkbook" /><Relationship Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="/xl/webextensions/taskpanes.xml" Id="rTaskpanes" /></Relationships>'
$workbook = '<?xml version="1.0" encoding="utf-8"?><x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheets><x:sheet name="땡땡엑셀 시작" sheetId="1" r:id="rSheet" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" /></x:sheets></x:workbook>'
$workbookRelationships = '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet.xml" Id="rSheet" /></Relationships>'
$worksheet = '<?xml version="1.0" encoding="utf-8"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData /></x:worksheet>'
$taskpanes = '<?xml version="1.0" encoding="utf-8"?><wetp:taskpanes xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"><wetp:taskpane dockstate="" visibility="1" width="350" row="1"><wetp:webextensionref r:id="rWebextension" /></wetp:taskpane></wetp:taskpanes>'
$taskpaneRelationships = '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="/xl/webextensions/webextension.xml" Id="rWebextension" /></Relationships>'
$webextension = "<?xml version=`"1.0`" encoding=`"utf-8`"?><we:webextension xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`" xmlns:we=`"http://schemas.microsoft.com/office/webextensions/webextension/2010/11`" id=`"{$manifestId}`"><we:reference id=`"$manifestId`" version=`"$manifestVersion`" store=`"developer`" storeType=`"Registry`" /><we:alternateReferences /><we:properties></we:properties><we:bindings /><we:snapshot /></we:webextension>"

$file = [IO.File]::Open($OutputPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
try {
    $archive = New-Object IO.Compression.ZipArchive(
        $file,
        [IO.Compression.ZipArchiveMode]::Create,
        $false)
    try {
        Add-ZipText $archive "[Content_Types].xml" $contentTypes
        Add-ZipText $archive "_rels/.rels" $rootRelationships
        Add-ZipText $archive "xl/workbook.xml" $workbook
        Add-ZipText $archive "xl/_rels/workbook.xml.rels" $workbookRelationships
        Add-ZipText $archive "xl/worksheets/sheet.xml" $worksheet
        Add-ZipText $archive "xl/webextensions/taskpanes.xml" $taskpanes
        Add-ZipText $archive "xl/webextensions/_rels/taskpanes.xml.rels" $taskpaneRelationships
        Add-ZipText $archive "xl/webextensions/webextension.xml" $webextension
    } finally {
        $archive.Dispose()
    }
} finally {
    $file.Dispose()
}

Write-Host "Embedded WEF bootstrap workbook: $OutputPath"
