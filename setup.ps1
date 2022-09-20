Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

Install-Module -Name MergePdf
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
Get-PSRepository  

choco install python3
choco install --pre tesseract
choco install ghostscript
choco install pngquant

pip install ocrmypdf

