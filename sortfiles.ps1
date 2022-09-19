$params = $args
$rootPath = 'H:\My Drive\(1) ScaleTickets2022\' 
$weekFolderName = 'week_' + (get-date -UFormat %V) + '_' + (Get-Date -Format "yyyy") + '\'


function OutputPS {
    param (
        $Text
    )
    if ($params.Contains("-v")) {
        Write-Host ("{0}" -f $Text)
    }
} 
    
function MoveFiles {
    Set-Location -Path $rootPath
    $customers = Import-Csv -Path .\customers.txt -Header "filename", "1", "2", "3", "4", "5" -Delimiter "|"
    OutputPS "Entering Movefiles()"
    OutputPS ('Moving files into corresponding folder in ' + $rootPath + $weekFolderName)

    # create week file after testing if exists
    if (!(Test-Path -Path ($rootPath + $weekFolderName))) {
        New-Item -Path ($rootPath + $weekFolderName) -ItemType Directory
        OutputPS ("folder does not exist Creating folder " + $weekFolderName)

    }

    # Rename all files to lowercase for simple
    Get-ChildItem -Path ($rootPath) -File | ForEach-Object {
        Rename-Item $_ -NewName $_.ToString().ToLower()

    }

    Get-ChildItem -Path ($rootPath) -File | ForEach-Object {

        OutputPS ("Processing " + $_ + " File")
        foreach ($customer in $customers) {
            OutputPS (".. ..")

            OutputPS ("`X Processing " + $customer.filename + " Folder")
            
            # Test not null
            $file = $_.ToString()

            # If a customer matches the email name
            if ($.Contains($customer.filename)`
                    -or $file.Contains($customer.1)`
                    -or $file.Contains($customer.2)`
                    -or $file.Contains($customer.3)`
                    -or $file.Contains($customer.4)`
                    -or $file.Contains($customer.5)) {
                
                # Creates File by 
                # tests for existing folder
                if (!(Test-Path -Path ($rootPath + $weekFolderName + $customer.filename))) {
                    New-Item -Path ($rootPath + $weekFolderName + $customer.filename) -ItemType Directory
                    OutputPS ("folder does not exist Creating folder " + $customer.filename)
                }

                # move $_ into $customer file
                OutputPS ("Filepath = " + $rootPath + $_)

                Move-Item $_ -Destination ($rootPath + $weekFolderName + $customer.filename) 
                OutputPS ("^" + $_ + " Moved into " + $customer.filename)
            }
        }
        OutputPS ($rootPath + $_)
        OutputPS ($rootPath + $weekFolderName).GetType()
        Move-Item $_ -Destination ($rootPath + $weekFolderName)
    }
    OutputPS "Exiting Movefiles()"
}

OutputPS -Text "Begin File"

# CreateLoadFolders
MoveFiles

