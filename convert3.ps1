$OUTPUT_DIR = "C:\Users\$env:username\Desktop\OneNoteExport"
$ASSETS_DIR = "assets"

$current_notebook = "" # set globally below

write-host $OUTPUT_PATH
$OneNote = New-Object -ComObject OneNote.Application
[xml]$Hierarchy = ""
$OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)

function handlePage($page, $path, $assets_path, $i) {
    # Make the dir if not exists
    $abspath = [IO.Path]::Combine($OUTPUT_DIR, $path)
    New-Item -ItemType Directory -Force -Path $abspath | Out-Null
    Write-Host PAGE: $page.name at $path
    $pagename = $page.name.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
    $pagename = -join(([string]$i).PadLeft(2,'0'), "_", $pagename)
    $filename_docx = -join($pagename, ".docx")
    $filename_md = -join($pagename, ".md")
    $filename_htm = -join($pagename, ".htm")
    $filename_assets = -join($pagename, "_files")
    $path_docx = [IO.Path]::Combine($OUTPUT_DIR, $path, $filename_docx)
    $path_htm = [IO.Path]::Combine($OUTPUT_DIR, $path, $filename_htm)
    $path_md = [IO.Path]::Combine($OUTPUT_DIR, $path, $filename_md)
    $path_assets = [IO.Path]::Combine($OUTPUT_DIR, $path, $filename_assets)
    Write-Host Creating DOCX: $path_docx
    $OneNote.Publish($page.ID, $path_docx, 5, "")
    $OneNote.Publish($page.ID, $path_htm, 7, "")
    pandoc.exe -i $path_docx -o $path_md -t markdown-simple_tables-multiline_tables-grid_tables --wrap=none
    rm $path_docx
    rm $path_htm
    if (Test-Path $path_assets) {
        $path_assets_wild = Join-Path $path_assets "*"
        $section_assets = [IO.Path]::Combine($OUTPUT_DIR, $path, "assets")
        New-Item -ItemType Directory -Force -Path $section_assets | Out-Null
        $items = Get-ChildItem $path_assets_wild
        $items | Rename-Item -NewName { $pagename + "_" + $_.Name };
        Move-Item -Force $path_assets_wild $section_assets
        rmdir $path_assets

        # Replace image links with the right path
        Write-Host Fixing image links
        $LINK_REGEX = '!\[[^\]]*\]\((?<filename>.*?)(?=\"|\))(?<optionalpart>\".*\")?\)(\{.*\})?'
        $path_md2 = -join($path_md, "2")
        $prefix = -join($pagename, "_")
        (Get-Content $path_md) `
            -replace $LINK_REGEX, "![[$prefix${filename}]]" `
            -replace 'second regex', 'second replacement' |
          Out-File $path_md2
    }
    sleep 5
}

function handleSection($section, $path) {
    $path = Join-Path $path $section.name
    Write-Host Section: $path
    $i=0
    foreach($page in $section.ChildNodes) {
        handlePage $page $path $assets_path $i
        $i++
    }
}

function handleSectionGroup($sg, $path) {
    $path = Join-Path $path $sg.name
    if ($sg.isRecycleBin -ne 'true') {
        Write-Host Section Group: $path
        foreach ($section in $sg.ChildNodes) {
            handleSection $section $path
        }
        foreach ($sg2 in $sg.SectionGroup) {
            handleSectionGroup $sg2 $path
        }
    }
}

foreach ($notebook in $Hierarchy.Notebooks.Notebook ) {
    $current_notebook = $notebook.Name
    Write-Host Notebook: $current_notebook
    $path = $current_notebook
    foreach ($sectiongroup in $notebook.SectionGroup) {
        handleSectionGroup $sectiongroup $path
    }
    foreach ($section in $notebook.Section) {
        handleSection $section $path
    }
}