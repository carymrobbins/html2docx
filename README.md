html2docx
=========

html2docx utilizes Microsoft Word's ability to convert HTML to DOCX.

*Disclaimer* - The files generated by this utility only work in Microsoft Word and are not compatible
with other word processing variants (OpenOffice, etc.)

Setup
-----

```bash
cd path/to/html2docx

# Dependencies
./install_apt_requirements

# Build
xbuild /p:Configuration=Release

# Usage
cd html2docx/bin/Release
mono html2docx.exe inputfile.html outputfile.docx
```

