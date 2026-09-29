# MacroShelf

A growing collection of practical Excel VBA macros to automate everyday tasks.

**Browse the bilingual macro page:** https://tipek41.github.io/MacroShelf/

## Available macro

### Merge PDFs in a folder
Select a folder and merge its top-level PDF files into one PDF, sorted by filename. The source PDFs remain unchanged. The macro skips its previous output and reports files it could not read.

- [Download the VBA module](pdf-merge.bas)
- Requirements: Excel for Windows and the Adobe Acrobat desktop application. Acrobat Reader alone does not support the document-editing OLE automation used by this macro.
- The macro processes files locally and does not upload PDFs to a web service.

## Install

1. In Excel, press **Alt+F11** to open the VBA editor.
2. Choose **File → Import File…** and select `pdf-merge.bas`.
3. Run `MergePDFsInFolder` and choose the folder containing the PDFs.

## GitHub Pages

This repository's site is the `index.html` file in the root. Enable Pages from **Settings → Pages → Deploy from a branch → main → /(root)**.
