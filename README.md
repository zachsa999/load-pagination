# Simple broken Powershell script to automate load sorting
## Usage
create a customers.txt file at the root of this directory with format
```
customername|misspelling1|cutstotmer|thereneeds|tobesix|columnsTotal
```
the only important bit is that there is six columns. fill the unneeded columns with junk like `alksdaopisdfo`. I haven't tested it with special characters, stick to alphanumeric only.

## Setup
### Ensure you have powerhsell > 7 installed 
You can try running `setup.ps1`, I have not tested it, and a lot of the commands require reboot. All the setup steps require elevated permissions.

### Install [jBig2](https://github.com/ocrmypdf/OCRmyPDF/issues/748)
A little convoluted see the github thread.
Move it to a folder in ProgramFiles and add the fold;er to environment variables

### Install [Chocolatey](https://chocolatey.org)

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
```

### Install [MergePDF](https://anthony-f-tannous.medium.com/merge-pdf-files-b02685a4f410)
```powershell
Install-Module -Name MergePdf
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
Get-PSRepository
```
### Install [ocrmypdf](https://ocrmypdf.readthedocs.io/en/latest/index.html)
```powershell
choco install python3
choco install --pre tesseract
choco install ghostscript
choco install pngquant
pip install ocrmypdf
```
Congratulations! After a reboot `sortfiles.ps1` should be ready to run.