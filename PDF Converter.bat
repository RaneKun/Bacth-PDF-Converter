@echo off
setlocal enabledelayedexpansion

:: ============================================================================
:: PDF Converter - Converts various file formats to PDF (Keeps Originals)
:: Supported formats: .docx, .pptx, .txt, .doc, .ppt, .png, .jpg, .jpeg
:: Author: Rane
:: ============================================================================

:: Get current folder where the script is located
set "folder=%~dp0"
echo =========================================
echo         PDF CONVERTER STARTED           
echo =========================================
echo Target Folder: %folder%
echo.

:: ----------------------------------------------------------------------------
:: Convert .docx and .doc files using Microsoft Word
:: ----------------------------------------------------------------------------
for %%F in ("%folder%*.docx" "%folder%*.doc") do (
    echo [Word] Converting: %%~nxF
    powershell -NoProfile -Command ^
        "$file = '%%F';" ^
        "$pdf = [System.IO.Path]::ChangeExtension($file, 'pdf');" ^
        "$word = New-Object -ComObject Word.Application;" ^
        "$word.Visible = $false;" ^
        "$doc = $word.Documents.Open($file);" ^
        "$doc.SaveAs([ref]$pdf, [ref]17);" ^
        "$doc.Close($false);" ^
        "$word.Quit();"
    if errorlevel 1 (
        echo [ERROR] Failed to convert: %%~nxF
    ) else (
        echo [SUCCESS] Converted: %%~nxF to PDF
    )
)

:: ----------------------------------------------------------------------------
:: Convert .pptx and .ppt files using Microsoft PowerPoint
:: ----------------------------------------------------------------------------
for %%F in ("%folder%*.pptx" "%folder%*.ppt") do (
    echo [PowerPoint] Converting: %%~nxF
    powershell -NoProfile -Command ^
        "$file = '%%F';" ^
        "$pdf = [System.IO.Path]::ChangeExtension($file, 'pdf');" ^
        "$wshell = New-Object -ComObject wscript.shell;" ^
        "$ppt = New-Object -ComObject PowerPoint.Application;" ^
        "$ppt.DisplayAlerts = 'ppAlertsNone';" ^
        "$pres = $ppt.Presentations.Open($file, 0, 0, 0);" ^
        "$null = $wshell.AppActivate('Microsoft PowerPoint');" ^
        "$null = $wshell.SendKeys('{ESC}');" ^
        "$pres.SaveAs($pdf, 32);" ^
        "$pres.Close();" ^
        "$ppt.Quit();"
    if errorlevel 1 (
        echo [ERROR] Failed to convert: %%~nxF
    ) else (
        echo [SUCCESS] Converted: %%~nxF to PDF
    )
)

:: ----------------------------------------------------------------------------
:: Convert .txt files using Microsoft Word
:: ----------------------------------------------------------------------------
for %%F in ("%folder%*.txt") do (
    echo [Text] Converting: %%~nxF
    powershell -NoProfile -Command ^
        "$file = '%%F';" ^
        "$pdf = [System.IO.Path]::ChangeExtension($file, 'pdf');" ^
        "$word = New-Object -ComObject Word.Application;" ^
        "$word.Visible = $false;" ^
        "$doc = $word.Documents.Open($file);" ^
        "$doc.SaveAs([ref]$pdf, [ref]17);" ^
        "$doc.Close($false);" ^
        "$word.Quit();"
    if errorlevel 1 (
        echo [ERROR] Failed to convert: %%~nxF
    ) else (
        echo [SUCCESS] Converted: %%~nxF to PDF
    )
)

:: ----------------------------------------------------------------------------
:: Convert image files (.png, .jpg, .jpeg) to PDF using Word
:: ----------------------------------------------------------------------------
for %%F in ("%folder%*.png" "%folder%*.jpg" "%folder%*.jpeg") do (
    echo [Image] Converting: %%~nxF
    powershell -NoProfile -Command ^
        "$file = '%%F';" ^
        "$pdf = [System.IO.Path]::ChangeExtension($file, 'pdf');" ^
        "$word = New-Object -ComObject Word.Application;" ^
        "$word.Visible = $false;" ^
        "$doc = $word.Documents.Add();" ^
        "$doc.PageSetup.TopMargin = 0;" ^
        "$doc.PageSetup.BottomMargin = 0;" ^
        "$doc.PageSetup.LeftMargin = 0;" ^
        "$doc.PageSetup.RightMargin = 0;" ^
        "$selection = $word.Selection;" ^
        "$shape = $selection.InlineShapes.AddPicture($file).ConvertToShape();" ^
        "$imgWidth = $shape.Width;" ^
        "$imgHeight = $shape.Height;" ^
        "$pageWidth = $doc.PageSetup.PageWidth;" ^
        "$pageHeight = $doc.PageSetup.PageHeight;" ^
        "$ratio = [Math]::Min($pageWidth/$imgWidth, $pageHeight/$imgHeight);" ^
        "$shape.Width = $imgWidth * $ratio;" ^
        "$shape.Height = $imgHeight * $ratio;" ^
        "$shape.RelativeHorizontalPosition = 0;" ^
        "$shape.RelativeVerticalPosition = 0;" ^
        "$shape.Left = -999995;" ^
        "$shape.Top = -999995;" ^
        "$shape.RelativeHorizontalPosition = 1;" ^
        "$shape.RelativeVerticalPosition = 1;" ^
        "$shape.Left = ($pageWidth - $shape.Width) / 2;" ^
        "$shape.Top = ($pageHeight - $shape.Height) / 2;" ^
        "$doc.SaveAs([ref]$pdf, [ref]17);" ^
        "$doc.Close($false);" ^
        "$word.Quit();"
    if errorlevel 1 (
        echo [ERROR] Failed to convert: %%~nxF
    ) else (
        echo [SUCCESS] Converted: %%~nxF to PDF
    )
)

echo.
echo =========================================
echo        CONVERSION COMPLETE              
echo =========================================
pause