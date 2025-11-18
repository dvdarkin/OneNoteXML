# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT

param(
    [Parameter(Mandatory=$true)]
    [string]$NotebookName,
    [string]$OutputPath = "output\XML"
)

Write-Host "OneNote XML Export for Notebook: $NotebookName"
Write-Host "Output: $OutputPath"

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$onenote = $null

try {
    $onenote = New-Object -ComObject OneNote.Application
    Write-Host "Connected to OneNote"
    
    $hierarchyXml = ""
    $onenote.GetHierarchy("", 4, [ref]$hierarchyXml)
    Write-Host "Got hierarchy XML"
    
    $xml = [xml]$hierarchyXml
    
    # Find the specified notebook (excluding web-based ones)
    $targetNotebook = $xml.Notebooks.Notebook | Where-Object { 
        $_.name -eq $NotebookName -and $_.path -notlike "*web*" 
    }
    
    if ($targetNotebook) {
        Write-Host "Found '$NotebookName' notebook"
        
        $notebookDir = Join-Path $OutputPath ($NotebookName + "_XML")
        if (-not (Test-Path $notebookDir)) {
            New-Item -ItemType Directory -Path $notebookDir -Force | Out-Null
        }
        
        $sectionCount = 0
        $totalPages = 0
        
        foreach ($section in $targetNotebook.Section) {
            $sectionCount++
            Write-Host "Section $sectionCount : $($section.name)"
            
            # Sanitize section name for folder
            $sectionName = $section.name -replace '[^\w\s.-]', '_'
            $sectionDir = Join-Path $notebookDir $sectionName
            if (-not (Test-Path $sectionDir)) {
                New-Item -ItemType Directory -Path $sectionDir -Force | Out-Null
            }
            
            if ($section.Page) {
                $pageIndex = 0
                foreach ($page in $section.Page) {
                    $pageIndex++
                    $totalPages++
                    $pageName = if ($page.name) { $page.name } else { "Page_$pageIndex" }
                    
                    # Sanitize page name for filename
                    $safePageName = $pageName -replace '[^\w\s.-]', '_'
                    
                    try {
                        $pageXml = ""
                        $onenote.GetPageContent($page.ID, [ref]$pageXml)
                        
                        $pageFile = Join-Path $sectionDir "$safePageName.xml"
                        $pageXml | Out-File $pageFile -Encoding UTF8
                        Write-Host "  Page $pageIndex : $pageName -> XML"
                    } catch {
                        Write-Host "  Page $pageIndex FAILED: $($_.Exception.Message)"
                    }
                }
            } else {
                Write-Host "  No pages in this section"
            }
        }
        
        Write-Host ""
        Write-Host "Export complete:"
        Write-Host "  Notebook: $NotebookName"
        Write-Host "  Sections: $sectionCount"
        Write-Host "  Total Pages: $totalPages"
        Write-Host "  Output Directory: $notebookDir"
        
    } else {
        Write-Host "ERROR: Notebook '$NotebookName' not found"
        Write-Host ""
        Write-Host "Available notebooks:"
        foreach ($notebook in $xml.Notebooks.Notebook) {
            if ($notebook.path -notlike "*web*") {
                Write-Host "  - $($notebook.name)"
            }
        }
        exit 1
    }
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($onenote) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($onenote) | Out-Null
    }
}