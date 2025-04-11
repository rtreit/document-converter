<#
.SYNOPSIS
  Batch converts Word (.docx) files to Markdown and fixes Pandoc image syntax issues.

.DESCRIPTION
  This script processes all Word (.docx) files in a specified folder, converting them to
  Markdown using Pandoc and then fixing the image syntax by removing attribute blocks.
  It creates an 'images' folder for each document's media and ensures proper image references.

.PARAMETER FolderPath
  The path to the folder containing .docx files to process. Defaults to current directory.

.PARAMETER SkipImageFix
  If specified, the script will convert the documents but skip the image syntax fixing step.

.PARAMETER NoBackup
  If specified, the script will not create backup files during the image syntax fixing process.

.EXAMPLE
  .\Convert-DocxToMarkdown.ps1
  Processes all .docx files in the current directory.

.EXAMPLE
  .\Convert-DocxToMarkdown.ps1 -FolderPath "C:\MyDocuments"
  Processes all .docx files in the C:\MyDocuments folder.

.EXAMPLE
  .\Convert-DocxToMarkdown.ps1 -FolderPath "C:\MyDocuments" -SkipImageFix
  Only converts the documents without fixing image syntax.

.NOTES
  Requires Pandoc to be installed and accessible in the PATH.
  The script will create a subfolder called 'images' for each document's extracted media.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Position=0)]
    [string]$FolderPath = (Get-Location).Path,
    
    [Parameter()]
    [switch]$SkipImageFix,
    
    [Parameter()]
    [switch]$NoBackup
)

#region Helper Functions

function Test-PandocInstalled {
    try {
        $pandocVersion = pandoc --version
        if ($LASTEXITCODE -eq 0) {
            $versionLine = $pandocVersion[0]
            Write-Host "Using $versionLine" -ForegroundColor Green
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

function Fix-PandocImageSyntax {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter()]
        [switch]$NoBackup
    )
    
    Write-Host "  Fixing image syntax in: $FilePath" -ForegroundColor Cyan
    
    # Create backup if requested
    if (-not $NoBackup) {
        $backupFilePath = "$FilePath.bak"
        try {
            Copy-Item -Path $FilePath -Destination $backupFilePath -Force -ErrorAction Stop
            Write-Host "  Backup created: $backupFilePath" -ForegroundColor Gray
        } catch {
            Write-Error "Failed to create backup file '$backupFilePath'. Error: $_"
            return $false
        }
    }
    
    try {
        # Read the entire file content
        $fileContent = Get-Content -Path $FilePath -Raw -Encoding UTF8
        
        # Store original content for change detection
        $originalContent = $fileContent
        
        # Fix common Pandoc image pattern
        # This matches ![alt](path){width="X" height="Y"} and removes the attributes
        $regexPattern = '(!\[[^\]]*\]\([^\)]*\))\{[^}]*\}'
        $fileContent = [regex]::Replace($fileContent, $regexPattern, '$1')
        
        # Count changes (approximate)
        $changesMade = ($originalContent -ne $fileContent)
        
        # Write changes if any were made
        if ($changesMade) {
            Set-Content -Path $FilePath -Value $fileContent -Encoding UTF8 -NoNewline
            Write-Host "  Image syntax fixed successfully." -ForegroundColor Green
            return $true
        } else {
            Write-Host "  No image syntax issues found to fix." -ForegroundColor Yellow
            return $true
        }
    } catch {
        Write-Error "Error processing file '$FilePath': $_"
        return $false
    }
}

#endregion

# Validate folder path
if (-not (Test-Path -Path $FolderPath -PathType Container)) {
    Write-Error "Folder not found: $FolderPath"
    exit 1
}

# Check if Pandoc is installed
if (-not (Test-PandocInstalled)) {
    Write-Error "Pandoc is not installed or not in PATH. Please install Pandoc before running this script."
    Write-Host "You can download Pandoc from: https://pandoc.org/installing.html" -ForegroundColor Yellow
    exit 1
}

# Get all .docx files in the specified folder
$docxFiles = Get-ChildItem -Path $FolderPath -Filter "*.docx" -File

if ($docxFiles.Count -eq 0) {
    Write-Warning "No .docx files found in $FolderPath"
    exit 0
}

Write-Host "Found $($docxFiles.Count) Word document(s) to process." -ForegroundColor Cyan

# Process each file
foreach ($docxFile in $docxFiles) {
    # Extract filename without extension
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($docxFile.Name)
    $mdFileName = "$baseName.md"
    $mdFilePath = Join-Path -Path $FolderPath -ChildPath $mdFileName
    
    # Create image folder path for this document
    $imagesFolder = Join-Path -Path $FolderPath -ChildPath "images"
    if (-not (Test-Path -Path $imagesFolder -PathType Container)) {
        New-Item -Path $imagesFolder -ItemType Directory | Out-Null
        Write-Host "Created images folder: $imagesFolder" -ForegroundColor Gray
    }
    
    # Convert docx to markdown using Pandoc
    Write-Host "Converting: $($docxFile.FullName) -> $mdFileName" -ForegroundColor Cyan
    
    # Build the pandoc command
    $pandocCmd = "pandoc '$($docxFile.FullName)' -t markdown --extract-media=./images -o '$mdFilePath'"
    
    try {
        # Execute Pandoc command
        Invoke-Expression $pandocCmd
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  Conversion successful." -ForegroundColor Green
            
            # Fix image syntax if not skipped
            if (-not $SkipImageFix) {
                $fixResult = Fix-PandocImageSyntax -FilePath $mdFilePath -NoBackup:$NoBackup
                if (-not $fixResult) {
                    Write-Warning "  Image syntax fixing failed or was skipped for $mdFileName"
                }
            }
        } else {
            Write-Error "Pandoc conversion failed for $($docxFile.Name) with exit code $LASTEXITCODE"
        }
    } catch {
        Write-Error "Error executing Pandoc: $_"
    }
    
    Write-Host "--------------------------------------------------"
}

Write-Host "Processing complete!" -ForegroundColor Green

# Optional: Display summary of processed files
$mdFiles = Get-ChildItem -Path $FolderPath -Filter "*.md" -File
Write-Host "Total Markdown files: $($mdFiles.Count)" -ForegroundColor Cyan

# Verify images folder has content
$imageFiles = Get-ChildItem -Path (Join-Path -Path $FolderPath -ChildPath "images") -File -Recurse -ErrorAction SilentlyContinue
if ($imageFiles) {
    Write-Host "Total extracted image files: $($imageFiles.Count)" -ForegroundColor Cyan
} else {
    Write-Warning "No images were extracted. This might be normal if your documents don't contain images."
}