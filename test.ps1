$params = $args
function ConvertTo-PDF {
    param(
        $TextDocumentPath
    )
	
    Add-Type -AssemblyName System.Drawing
    $doc = New-Object System.Drawing.Printing.PrintDocument
    $doc.DocumentName = $TextDocumentPath
    $doc.PrinterSettings = new-Object System.Drawing.Printing.PrinterSettings
    $doc.PrinterSettings.PrinterName = 'Microsoft Print to PDF'
    $doc.PrinterSettings.PrintToFile = $true
    $file = [io.fileinfo]$TextDocumentPath
    $pdf = [io.path]::Combine($file.DirectoryName, $file.BaseName) + '.pdf'
    $doc.PrinterSettings.PrintFileName = $pdf
    $doc.Print()
    $doc.Dispose()
}
Set-Location "H:\My Drive\(1) ScaleTickets2022\week_38_2022\jgl"
ConvertTo-PDF $params[0]
Set-Location "C:\Users\Zach\Documents\projects\invoicing"