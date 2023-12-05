# Function for parse html file out of html elements
function Parse-HTML {
    param([string]$htmlContent)

    # Remove scripts and styles
    $htmlContent = $htmlContent -replace '<script[^>]*>[\s\S]*?</script>|<style[^>]*>[\s\S]*?</style>', ''

    # Convert unordered lists
    $htmlContent = $htmlContent -replace '<ul>', "`n" -replace '</ul>', "`n"
    $htmlContent = $htmlContent -replace '<li>', "• " -replace '</li>', "`n"

    # Convert ordered lists
    $counter = 0
    $htmlContent = $htmlContent -replace '<ol>', {}
    $htmlContent = $htmlContent -replace '<li>', {
        $counter++
        "$counter. "
    } -replace '</li>', "`n"

    # Handle emphasis
    $htmlContent = $htmlContent -replace '<strong>', '<b>' -replace '</strong>', '</b>'
    $htmlContent = $htmlContent -replace '<em>', '<i>' -replace '</em>', '</i>'
    $htmlContent = $htmlContent -replace '<b>', "`n<b>" -replace '</b>', "</b>`n"
    $htmlContent = $htmlContent -replace '<i>', "`n<i>" -replace '</i>', "</i>`n"

    # Convert paragraph tags, headers, and <div>, <article> tags
    $htmlContent = $htmlContent -replace '<h[1-6]>', "`n" -replace '</h[1-6]>', "`n`n"
    $htmlContent = $htmlContent -replace '<p>', "`n" -replace '</p>', "`n`n"
    $htmlContent = $htmlContent -replace '<div>', "`n" -replace '</div>', "`n`n"
    $htmlContent = $htmlContent -replace '<article>', "`n" -replace '</article>', "`n`n"

    # Remove all other HTML tags
    $htmlContent = $htmlContent -replace '<[^>]+>', ''

    # Decode HTML entities
    $textContent = [System.Web.HttpUtility]::HtmlDecode($htmlContent)

    # Remove unnecessary white spaces
    $textContent = $textContent -replace '(\s*\r?\n\s*)+', "`n"

    return $textContent.Trim()
}


# Function for generate document in .xml with content
function Generate-DOCXXml {
    param([string]$textContent)

     # Split content into paragraphs
     $paragraphs = $textContent -split "`n" | Where-Object { $_.Trim() -ne '' }

    $documentXmlContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
            xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14 w15 wp14">
    <w:body>
"@
    # Split the text into fragments considering <b> and <i> tags
    $textContent -split '(\n|<b>|</b>|<i>|</i>)' | ForEach-Object {
        if ($_ -match '^<b>') { $inBold = $true }
        elseif ($_ -match '^</b>') { $inBold = $false }
        elseif ($_ -match '^<i>') { $inItalic = $true }
        elseif ($_ -match '^</i>') { $inItalic = $false }
        elseif ($_ -match '\S') {
            $textFragment = $_.Replace('<b>', '').Replace('</b>', '').Replace('<i>', '').Replace('</i>', '')
            $documentXmlContent += "<w:p><w:r>"
            if ($inBold) { $documentXmlContent += "<w:rPr><w:b/></w:rPr>" }
            if ($inItalic) { $documentXmlContent += "<w:rPr><w:i/></w:rPr>" }
            $documentXmlContent += "<w:t xml:space='preserve'>$textFragment</w:t></w:r></w:p>"
        }
    }

    $documentXmlContent += @"
    </w:body>
</w:document>
"@

    return $documentXmlContent
}


# Function for creating directory structure inside .zip file
function Create-DOCXStructure {
    param([string]$outputPath)

    Write-Host "Creating directory structure..."
    $docxDir = Join-Path $outputPath "docx"
    $wordDir = Join-Path $docxDir "word"
    $relsDir = Join-Path $docxDir "_rels"
    $wordRelsDir = Join-Path $wordDir "_rels"
    
    New-Item -ItemType Directory -Path $wordDir, $relsDir, $wordRelsDir -Force | Out-Null

    return $docxDir, $wordDir, $relsDir, $wordRelsDir
}


# Function for creating [Content_Types].xml
function Create-ContentTypesXml {
    param([string]$docxDir)

    Write-Host "Creating file [Content_Types].xml..."
    $contentTypesXmlPath = Join-Path $docxDir "[Content_Types].xml"
    $contentTypesXmlContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
    <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
    <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"@
    [System.IO.File]::WriteAllText($contentTypesXmlPath, $contentTypesXmlContent)
}

# Function for creating all other neccesery files into .zip file, change .zip file into .docx
function Create-DOCXPackage {
    param(
        [string]$documentXml,
        [string]$outputPath
    )

    $docxDir, $wordDir, $relsDir, $wordRelsDir = Create-DOCXStructure -outputPath $outputPath

    # Creating a document.xml file in the word directory
    $documentXmlPath = Join-Path $wordDir "document.xml"
    Set-Content -Path $documentXmlPath -Value $documentXml -Force

    # Create the [Content_Types].xml file
    Create-ContentTypesXml -docxDir $docxDir

    # Create a master .rels relationship file in the _rels directory
    $relsPath = Join-Path $relsDir ".rels"
    $relsContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@
    Set-Content -Path $relsPath -Value $relsContent -Force

    # Create a document.xml.rels file in the word/_rels directory
    $documentXmlRelsPath = Join-Path $wordRelsDir "document.xml.rels"
    $documentXmlRelsContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
</Relationships>
"@
    Set-Content -Path $documentXmlRelsPath -Value $documentXmlRelsContent -Force

    # Packing everything into a ZIP archive
    $zipPath = Join-Path $outputPath "output.zip"
    Compress-Archive -Path "$docxDir\*" -DestinationPath $zipPath
    Write-Host "The ZIP archive has been created: $zipPath"
    
    # Change the extension from .zip to .docx
    $docxPath = [System.IO.Path]::ChangeExtension($zipPath, ".docx")
    Rename-Item -Path $zipPath -NewName $docxPath
    Write-Host "The DOCX file has been created: $docxPath"

    # Clean up temporary files
    Remove-Item -Path $docxDir -Recurse -Force

    # Return the path to the DOCX file
    return $docxPath
}

# Get paths from the user
$htmlPath = Read-Host "Provide the path to the HTML file"
$outputPath = Read-Host "Provide the path to the output directory"

# Preparation and use of functions
$htmlContent = Get-Content -Path $htmlPath -Raw
$textContent = Parse-HTML -htmlContent $htmlContent
$documentXml = Generate-DOCXXml -textContent $textContent
$docxPath = Create-DOCXPackage -documentXml $documentXml -outputPath $outputPath

