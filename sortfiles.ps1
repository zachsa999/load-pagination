$params = $args
$rootPath = 'I:\My Drive\ScaleTickets2023\' 
$projectPath = 'C:\Users\Zach\projects\load-pagination\'
$weekFolderName = 'week_' + (get-date -UFormat %V) + '_' + (Get-Date -Format "yyyy") + '\'
$altweekFolderName = 'week_' + (get-date -UFormat %V) + '_' + (Get-Date -Format "yyyy") + '_1' + '\'
$customers = Import-Csv -Path .\customers.txt -Header "filename", "1", "2", "3", "4", "5" -Delimiter "|"
$weekPathName = $rootPath + $weekFolderName

function OutputPS {
    param (
        $Text
    )
    if ($params.Contains("-v")) {
        Write-Host ("{0}" -f $Text)
    }
} 
    
function MoveFiles {
    Set-Location $rootPath

    OutputPS "Entering Movefiles()"
    OutputPS ('Moving files into corresponding folder in ' + $weekFolderName)
    # create week file after testing if exists
    if (!(Test-Path -Path ($weekFolderName))) {
        New-Item -Path ($weekFolderName) -ItemType Directory
        OutputPS ("folder does not exist Creating folder " + $weekFolderName)
    }
    else {
        New-Item -Path ($altweekFolderName) -ItemType Directory
        OutputPS ("folder does not exist Creating folder " + $altweekFolderName)
    }

    # Rename all files to lowercase for simple
    Get-ChildItem -File | ForEach-Object {
        Rename-Item $_ -NewName $_.ToString().ToLower()
    }

    Get-ChildItem -File | ForEach-Object {
        $fileMoved = 0
        foreach ($customer in $customers) {
            # Test not null/
            $file = $_.ToString()
            # If a customer matches the email name
            if ($file.Contains($customer.filename)`
                    -or $file.Contains($customer.1)`
                    -or $file.Contains($customer.2)`
                    -or $file.Contains($customer.3)`
                    -or $file.Contains($customer.4)`
                    -or $file.Contains($customer.5)) {
                if (!(Test-Path -Path ($weekFolderName + $customer.filename))) {
                    New-Item -Path ($weekFolderName + $customer.filename) -ItemType Directory
                }
                Move-Item $_ -Destination ($weekFolderName + $customer.filename) 
                $fileMoved = 1
            }
        }
        if ($fileMoved -eq 0) {
            Move-Item $_ -Destination $weekFolderName
        }
    }
    OutputPS "Exiting Movefiles()"
}

function MoveRest {
    Set-Location $weekPathName
    Get-ChildItem -File | ForEach-Object {
        $fileName = $_.ToString()

        Write-Host ("folder does not exist! Please Enter a customer for: '{0}'" -f $filename)

        if ($fileName -match "\.pdf$") {
            Start-Process $fileName
        }

        $newDirName = Read-Host -Prompt "Folder Name:: "

        if ([string]::IsNullOrEmpty($newDirName)) {
            Write-Host "Directory name cannot be null or empty."
            return
        }

        if (!(Test-Path $newDirName)) {
            New-Item -Path $newDirName -ItemType Directory
            Write-Host ("folder does not exist. Creating folder {0}" -f $newDirName)
        }
        
        Move-Item $_ -Destination $newDirName
    }
}
function MergeFiles {
    Set-Location $weekPathName
    OutputPS "Inside MergeFiles()"
    Get-ChildItem -Directory | ForEach-Object {
        OutputPS ("operating on " + $_)
        # Write-Output $files    
        Set-Location $_
        magick mogrify -format pdf *.jp*
        magick mogrify -format pdf *.png
        Remove-Item *.jp*
        Remove-Item *.png
        Set-Location $weekPathName
        Merge-Pdf -OutputPath $_ -Path $_
    }
    OutputPS "exiting MergeFiles()"
    Set-Location -Path $projectPath
}

function OCR {
    Set-Location $weekPathName
    OutputPS "Inside OCR()"
    Get-ChildItem -Directory | ForEach-Object {
        $targetFile = $_.ToString() + "/Merged.pdf"
        $outputFile = $_.ToString() + "/ocrMerged1.pdf"
        OutputPS ("operating on " + $targetFile)   
        ocrmypdf --redo-ocr --jbig2-lossy --optimize 3 --rotate-pages --rotate-pages-threshold 2.5 --output-type pdf $targetFile $outputFile
        $removeMerged = $_.ToString() + "\Merged.pdf"
        Remove-Item $removeMerged
    }
    OutputPS "exiting OCR()"
    Set-Location $projectPath
}


Set-Location $rootPath
MoveFiles
MoveRest
MergeFiles
OCR

Set-Location $projectPath

 