# File Page Count in SharePoint Document Library

## Overview

This PowerShell script automates the process of retrieving the page count of PDF, Word, and PowerPoint files stored in a SharePoint Online document library. It uses PnP PowerShell to connect to SharePoint, extracts files, calculates page/slides count, and updates the respective metadata column in SharePoint.

## Features

✅ Connects to SharePoint Online using PnP PowerShell.
✅ Downloads PDF, Word, and PowerPoint files from SharePoint.
✅ Extracts page count for:

- PDF using PdfSharp.
- Word documents using Word Interop.
- PowerPoint files using PowerPoint Interop.
  ✅ Updates a custom SharePoint column "FilePagesCount" with the calculated page/slide count.
  ✅ Supports bulk processing of files.

## Prerequisites

- Install the required PowerShell modules:
  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser -Force
  ```
- Install PdfSharp for PDF processing.
- Ensure Microsoft Word and PowerPoint are installed on the machine.
- SharePoint Online access with necessary permissions.

## Usage

1. Update the `$SiteURL` and `$LibraryName` variables in the script.
2. Run the script using PowerShell:
   ```powershell
   .\FilePageCountScript.ps1
   ```

## How It Works

- The script connects to SharePoint and retrieves files from the specified document library.
- It downloads each file to a local directory.
- Depending on the file type:
  - PDFs: Extracts page count using PdfSharp.
  - Word documents: Uses Microsoft Word Interop to compute page count.
  - PowerPoint files: Uses Microsoft PowerPoint Interop to count slides.
- The extracted page count is updated in the "FilePagesCount" column in SharePoint.

## Expected Output

- Console log of files processed and updated:
  ```
  ✅ Updated Report1.pdf with 12 pages
  ✅ Updated Document1.docx with 5 pages
  ✅ Updated Presentation1.pptx with 15 slides
  ```

## Notes

- The script assumes that "FilePagesCount" is an existing column in the SharePoint document library.
- Ensure that the script has the necessary permissions to read and write files.
- PowerPoint and Word applications must not be open while running the script.

## Disconnecting

- After execution, the script will automatically disconnect from SharePoint using:
  ```powershell
  Disconnect-PnPOnline
  ```
