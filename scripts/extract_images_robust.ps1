# Copyright (c) 2025 Denis Darkin
# SPDX-License-Identifier: MIT

param(
    [Parameter(Mandatory=$true)]
    [string]$NotebookName,
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,
    [Parameter(Mandatory=$true)]
    [string]$MapFile
)

function Get-ImageFormat {
    param([byte[]]$ImageBytes)
    
    if ($ImageBytes.Length -lt 4) {
        return $null
    }
    
    # Check PNG signature (89 50 4E 47)
    if ($ImageBytes[0] -eq 0x89 -and $ImageBytes[1] -eq 0x50 -and $ImageBytes[2] -eq 0x4E -and $ImageBytes[3] -eq 0x47) {
        return "png"
    }
    
    # Check JPEG signature (FF D8 FF)
    if ($ImageBytes[0] -eq 0xFF -and $ImageBytes[1] -eq 0xD8 -and $ImageBytes[2] -eq 0xFF) {
        return "jpeg"
    }
    
    # Check GIF signature (47 49 46)
    if ($ImageBytes[0] -eq 0x47 -and $ImageBytes[1] -eq 0x49 -and $ImageBytes[2] -eq 0x46) {
        return "gif"
    }
    
    # Check BMP signature (42 4D)
    if ($ImageBytes[0] -eq 0x42 -and $ImageBytes[1] -eq 0x4D) {
        return "bmp"
    }
    
    # Check WebP signature (52 49 46 46 ... 57 45 42 50)
    if ($ImageBytes.Length -ge 12 -and 
        $ImageBytes[0] -eq 0x52 -and $ImageBytes[1] -eq 0x49 -and $ImageBytes[2] -eq 0x46 -and $ImageBytes[3] -eq 0x46 -and
        $ImageBytes[8] -eq 0x57 -and $ImageBytes[9] -eq 0x45 -and $ImageBytes[10] -eq 0x42 -and $ImageBytes[11] -eq 0x50) {
        return "webp"
    }
    
    # Default to null if unknown
    return $null
}

Write-Host "OneNote Robust Image Extraction for Notebook: $NotebookName"
Write-Host "Output: $OutputPath"
Write-Host "Map File: $MapFile"

if (-not (Test-Path $MapFile)) {
    Write-Host "ERROR: Image extraction map file not found: $MapFile"
    exit 1
}

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$onenote = $null

try {
    # Read the image extraction map
    $mapContent = Get-Content $MapFile -Raw | ConvertFrom-Json
    $imageCount = ($mapContent | Get-Member -MemberType NoteProperty).Count
    
    if ($imageCount -eq 0) {
        Write-Host "No images found in extraction map"
        exit 0
    }
    
    Write-Host "Found $imageCount images to extract"
    
    $onenote = New-Object -ComObject OneNote.Application
    Write-Host "Connected to OneNote"
    
    # Get hierarchy to find all pages
    $hierarchyXml = ""
    $onenote.GetHierarchy("", 4, [ref]$hierarchyXml)
    $hierarchyDoc = [xml]$hierarchyXml
    
    # Find the target notebook
    $targetNotebook = $hierarchyDoc.Notebooks.Notebook | Where-Object { 
        $_.name -eq $NotebookName -and $_.path -notlike "*web*" 
    }
    
    if (-not $targetNotebook) {
        Write-Host "ERROR: Notebook '$NotebookName' not found"
        exit 1
    }
    
    Write-Host "Processing notebook: $($targetNotebook.name)"
    
    # Debug: List all available sections
    Write-Host "Available sections in notebook:"
    foreach ($section in $targetNotebook.Section) {
        Write-Host "  - '$($section.name)'"
    }
    Write-Host ""
    
    $successCount = 0
    $failureCount = 0
    $extractedImages = @{}  # Track which images we've extracted to avoid duplicates
    
    # For each image in the map
    foreach ($callbackId in $mapContent.PSObject.Properties.Name) {
        $imageInfo = $mapContent.$callbackId
        $sectionName = $imageInfo.section
        $pageName = $imageInfo.page
        $altText = $imageInfo.alt_text
        $targetPath = Join-Path $OutputPath (Split-Path $imageInfo.target_path -Leaf)
        
        # Skip if we already extracted this exact file
        $targetFileName = Split-Path $imageInfo.target_path -Leaf
        if ($extractedImages.ContainsKey($targetFileName)) {
            Write-Host "Skipping duplicate: $targetFileName (already extracted)"
            continue
        }
        
        Write-Host "Processing: CallbackID $callbackId"
        Write-Host "  Alt: '$altText'"
        Write-Host "  Section: $sectionName > $pageName"
        Write-Host "  Target: $targetFileName"
        
        try {
            # Find the section - try multiple variations
            $section = $targetNotebook.Section | Where-Object { $_.name -eq $sectionName }
            
            # If exact match fails, try with spaces instead of underscores
            if (-not $section) {
                $sectionNameWithSpaces = $sectionName -replace '_', ' '
                $section = $targetNotebook.Section | Where-Object { $_.name -eq $sectionNameWithSpaces }
                if ($section) {
                    Write-Host "  Found section with spaces: '$sectionNameWithSpaces'"
                }
            }
            
            # If still not found, try with underscores instead of dots
            if (-not $section) {
                $sectionNameWithUnderscores = $sectionName -replace '\.', '_'
                $section = $targetNotebook.Section | Where-Object { $_.name -eq $sectionNameWithUnderscores }
                if ($section) {
                    Write-Host "  Found section with underscores: '$sectionNameWithUnderscores'"
                }
            }
            
            # If still not found, try case-insensitive match
            if (-not $section) {
                $section = $targetNotebook.Section | Where-Object { $_.name -ieq $sectionName }
                if ($section) {
                    Write-Host "  Found section with case-insensitive match: '$($section.name)'"
                }
            }
            
            if (-not $section) {
                Write-Host "  ERROR: Section '$sectionName' not found (tried variations too)"
                $failureCount++
                continue
            }
            
            # Find the page
            $page = $section.Page | Where-Object { $_.name -eq $pageName }
            if (-not $page) {
                Write-Host "  ERROR: Page '$pageName' not found in section '$sectionName'"
                $failureCount++
                continue
            }
            
            # Get page content with binary data
            $pageXml = ""
            $onenote.GetPageContent($page.ID, [ref]$pageXml, 3)  # 3 = include binary data
            
            # Parse XML to find our image
            $pageDoc = [xml]$pageXml
            $ns = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
            
            # Create namespace manager for XPath
            $nsMgr = New-Object System.Xml.XmlNamespaceManager($pageDoc.NameTable)
            $nsMgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote")
            
            # Find all images in the page
            $allImages = $pageDoc.SelectNodes("//one:Image", $nsMgr)
            Write-Host "  Found $($allImages.Count) images in page"
            
            $foundImage = $false
            $imageIndex = 0
            
            foreach ($imageNode in $allImages) {
                $imageIndex++
                
                # Get the alt attribute and CallbackID
                $imgAlt = $imageNode.GetAttribute("alt")
                $callbackNode = $imageNode.SelectSingleNode("one:CallbackID", $nsMgr)
                $imgCallbackId = ""
                if ($callbackNode) {
                    $imgCallbackId = $callbackNode.GetAttribute("callbackID")
                }
                
                Write-Host "    Image $imageIndex : Alt='$imgAlt', CallbackID='$imgCallbackId'"
                
                # Check for Data element with base64 content
                $dataNode = $imageNode.SelectSingleNode("one:Data", $nsMgr)
                if ($dataNode -and $dataNode.InnerText) {
                    $dataLength = $dataNode.InnerText.Length
                    Write-Host "      Has data: $dataLength characters"
                    
                    # Match by CallbackID first (most reliable), then fall back to alt text
                    $isMatch = $false
                    
                    if ($imgCallbackId -and $imgCallbackId -eq $callbackId) {
                        Write-Host "      MATCH by CallbackID!"
                        $isMatch = $true
                    } elseif ($altText -and $altText.Trim() -ne "" -and $imgAlt -eq $altText) {
                        Write-Host "      MATCH by alt text!"
                        $isMatch = $true
                    } elseif ((-not $altText -or $altText.Trim() -eq "" -or $altText -eq "Image") -and $imageIndex -eq 1) {
                        # For images with no alt text, take the first one with data
                        Write-Host "      MATCH by position (first image with data)"
                        $isMatch = $true
                    }
                    
                    if ($isMatch) {
                        try {
                            $imageBytes = [System.Convert]::FromBase64String($dataNode.InnerText)

                            # Detect image format from binary header (for informational purposes)
                            $imageFormat = Get-ImageFormat $imageBytes
                            if ($imageFormat) {
                                Write-Host "      Detected format: $imageFormat"
                            }

                            # Save with the extension specified by Python (from target_path)
                            # This ensures markdown links match actual filenames
                            # Note: File extension doesn't need to match actual format -
                            # browsers/viewers handle JPEG data in .png files just fine
                            [System.IO.File]::WriteAllBytes($targetPath, $imageBytes)
                            $fileSizeKB = [Math]::Round($imageBytes.Length / 1024, 1)
                            Write-Host "      SUCCESS: Saved $fileSizeKB KB to $targetFileName"
                            $successCount++
                            $extractedImages[$targetFileName] = $true
                            $foundImage = $true
                            break
                        } catch {
                            Write-Host "      ERROR converting base64: $($_.Exception.Message)"
                        }
                    }
                } else {
                    Write-Host "      No data element found"
                }
            }
            
            if (-not $foundImage) {
                Write-Host "  FAILED: No matching image found"
                $failureCount++
            }
            
        } catch {
            Write-Host "  FAILED: $($_.Exception.Message)"
            $failureCount++
        }
        
        Write-Host ""
    }
    
    Write-Host "=========================================="
    Write-Host "Image extraction complete:"
    Write-Host "  Notebook: $NotebookName"
    Write-Host "  Total Images in Map: $imageCount"
    Write-Host "  Successfully Extracted: $successCount"
    Write-Host "  Failed: $failureCount"
    Write-Host "  Unique Files Created: $($extractedImages.Count)"
    
    if ($successCount -gt 0) {
        Write-Host ""
        Write-Host "Extracted images:"
        Get-ChildItem $OutputPath -File | ForEach-Object {
            $sizeKB = [Math]::Round($_.Length/1024, 1)
            Write-Host "  - $($_.Name) ($sizeKB KB)"
        }
    }
    
    if ($failureCount -gt 0) {
        Write-Host ""
        Write-Host "Note: Some images failed because:"
        Write-Host "  - Images are stored externally (linked, not embedded)"
        Write-Host "  - Images were deleted or moved in OneNote"
        Write-Host "  - Alt text matching failed (different between export and live data)"
        Write-Host "  - Image format not supported for extraction"
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