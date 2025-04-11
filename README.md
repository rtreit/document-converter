# Document Converter

A collection of utilities for converting between different document formats.

## Features

### Convert-DocxToMarkdown.ps1

A PowerShell script that batch converts Word (.docx) files to Markdown format.

#### Key Features:
- Converts Word documents to Markdown using Pandoc
- Automatically extracts and organizes images into an 'images' folder
- Fixes Pandoc image syntax issues by removing attribute blocks
- Creates backups of original files (optional)
- Supports batch processing of multiple documents

#### Requirements:
- PowerShell
- [Pandoc](https://pandoc.org/installing.html) installed and accessible in PATH

#### Usage:

Basic usage (processes all .docx files in current directory):
```powershell
.\scripts\Convert-DocxToMarkdown.ps1
```

Specify a folder containing .docx files:
```powershell
.\scripts\Convert-DocxToMarkdown.ps1 -FolderPath "C:\MyDocuments"
```

Convert without fixing image syntax:
```powershell
.\scripts\Convert-DocxToMarkdown.ps1 -FolderPath "C:\MyDocuments" -SkipImageFix
```

Convert without creating backup files:
```powershell
.\scripts\Convert-DocxToMarkdown.ps1 -NoBackup
```

#### Parameters:

| Parameter | Description |
|-----------|-------------|
| `-FolderPath` | Path to the folder containing .docx files to process. Defaults to current directory. |
| `-SkipImageFix` | If specified, the script will convert documents but skip the image syntax fixing step. |
| `-NoBackup` | If specified, the script will not create backup files during the image syntax fixing process. |

## License

See the [LICENSE](LICENSE) file for details.