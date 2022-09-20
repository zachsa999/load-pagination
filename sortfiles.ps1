$params = $args
$rootPath = 'H:\My Drive\(1) ScaleTickets2022\' 
$projectPath = 'C:\Users\Zach\Documents\projects\invoicing\'
$weekFolderName = 'week_' + (get-date -UFormat %V) + '_' + (Get-Date -Format "yyyy") + '\'
$customers = Import-Csv -Path .\customers.txt -Header "filename", "1", "2", "3", "4", "5" -Delimiter "|"

function OutputPS {
    param (
        $Text
    )
    if ($params.Contains("-v")) {
        Write-Host ("{0}" -f $Text)
    }
} 
    
function MoveFiles {
    OutputPS "Entering Movefiles()"
    OutputPS ('Moving files into corresponding folder in ' + $weekFolderName)

    # create week file after testing if exists
    if (!(Test-Path -Path ($weekFolderName))) {
        New-Item -Path ($weekFolderName) -ItemType Directory
        OutputPS ("folder does not exist Creating folder " + $weekFolderName)
    }

    # Rename all files to lowercase for simple
    Get-ChildItem -File | ForEach-Object {
        Rename-Item $_ -NewName $_.ToString().ToLower()
    }

    Get-ChildItem -File | ForEach-Object {

        OutputPS ("Processing " + $_ + " File")
        foreach ($customer in $customers) {
            OutputPS (".. ..")
            OutputPS ("`X Processing " + $customer.filename + " Folder")
            # Test not null
            $file = $_.ToString()
            # If a customer matches the email name
            if ($file.Contains($customer.filename)`
                    -or $file.Contains($customer.1)`
                    -or $file.Contains($customer.2)`
                    -or $file.Contains($customer.3)`
                    -or $file.Contains($customer.4)`
                    -or $file.Contains($customer.5)) {
                # Creates File by 
                # tests for existing folder
                if (!(Test-Path -Path ($weekFolderName + $customer.filename))) {
                    New-Item -Path ($weekFolderName + $customer.filename) -ItemType Directory
                    OutputPS ("folder does not exist Creating folder " + $customer.filename)
                }

                # move $_ into $customer file
                OutputPS ("Filepath = " + $_)

                Move-Item $_ -Destination ($weekFolderName + $customer.filename) 
                OutputPS ("^" + $_ + " Moved into " + $customer.filename)
            }
        }

        OutputPS ( $weekFolderName).GetType()
        Move-Item $_ -Destination $weekFolderName
    }
    OutputPS "Exiting Movefiles()"
}

function LooseEnds {
    Get-ChildItem -File | ForEach-Object {
        # extract metadata
        $fileName = $_.ToString()
        $pos1 = $fileName.IndexOf(')_subject(')
        $rightPart = $fileName.Substring($pos1 + 9)
        $pos2 = $rightPart.IndexOf('_date(')
        $leftPart = $rightPart.Substring(0, $pos2)

        $newDirName = Read-Host -Prompt "$leftPart new dir name?"

        if (!(Test-Path -Path ($newDirNam))) {
            New-Item -Path $newDirName -ItemType Directory
            New-Item -Path ($weekFolderName) -ItemType Directory
            OutputPS ("folder does not exist Creating folder " + $newDirNam)
        }
        Move-Item $_ -Destination $newDirName
    }
}

function MergeFiles {
    OutputPS "Inside MergeFiles()"
    Get-ChildItem -Directory | ForEach-Object {
        OutputPS ("operating on " + $_)
        # Write-Output $files    
        Merge-Pdf -OutputPath $_ -Path $_
    }
    OutputPS "exiting MergeFiles()"
}
function OCR {
    OutputPS "Inside OCR()"
    Get-ChildItem -Directory | ForEach-Object {
        $targetFile = $_.ToString() + "/Merged.pdf"
        $outputFile = $_.ToString() + "/ocrMerged1.pdf"
        OutputPS ("operating on " + $targetFile)   
        # ocrmypdf --redo-ocr --jbig2-lossy --optimize 3 --rotate-pages --rotate-pages-threshold 2.5 $targetFile $outputFile
        ocrmypdf --redo-ocr --jbig2-lossy --optimize 3 --rotate-pages --rotate-pages-threshold 2.5 --output-type pdf $targetFile $outputFile
    }
    OutputPS "exiting OCR()"
}


Set-Location "H:\My Drive\(1) ScaleTickets2022\week_38_2022"
# MergeFiles
OCR




# Set-Location -Path $rootPath
# OutputPS -Text "Begin File"
# # MoveFiles
# Set-Location -Path $weekFolderName
# LooseEnds

Set-Location -Path $projectPath

