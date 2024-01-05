Import-Module PSExcel

function checkConnection{   # check connection between remote pc and local machine
    param (
        [string]$ipAddress
    )
    Test-Connection -ComputerName $ipAddress -Count 1 -Quiet      # checking connection between remote and local pc
}

function TempDir{           # create temp dir for temporary txt files
    if (Test-Path ".\temp"){                           # create or remove temp directory to store all temporary .txt files of connected USB devides by each user 
        Remove-Item ".\temp" -Recurse -Force
        mkdir ".\temp"
    }
    else {
        mkdir ".\temp"
    }
}

function checkMappedDisks{  # check mapped disks
    param (
        [string]$hostname
    )

    $desiredUsername = Invoke-Command -ComputerName $hostname -ScriptBlock {            # invoke command on remote pc
        $(Get-CimInstance Win32_ComputerSystem | Select-Object username).username -split '\\' | Select-Object -Last 1;      # get username
    }

    $mappedDisks = Invoke-Command -ErrorAction SilentlyContinue -ComputerName $hostname -ScriptBlock {      # invoke command on remote pc
        $path = Get-ChildItem -Path "Registry::HKEY_USERS" | 
        Where-Object { ($_.Name -like "HKEY_USERS\S-1-5-21-*") -and ($_.Name -notlike "HKEY_USERS\*_Classes")  -and ($_.Name -notlike "HKEY_USERS\S-1-5-21-*-*-*-500") } | 
        Select-Object -ExpandProperty Name;                                     # get path to mapped disk in regedit

        Get-ChildItem -Path "Registry::\$path\Network" | ForEach-Object {       # get mapped disks
            [PSCustomObject]@{
                Disk_Letter = $_.PSChildName
                Disk_Path = $_.GetValue("RemotePath")
            }
        }
    }
    
    $data = $mappedDisks | ForEach-Object {
        "{0}`t{1}`t{2}`t{3}`t{4}" -f $user, $desiredUsername, $_.Disk_Letter, $_.Disk_Path, $hostname
    }
    
    if($null -ne $mappedDisks){
        $data | Out-File -FilePath ".\temp\$hostname.txt" -Encoding UTF8 -Append        # create temp txt file with user and plugged devices
    }
}

function mergeTxt{          # concenate temp txt files into one
    $txtFiles = Get-ChildItem -Path ".\temp"                                                                 # get all .txt files in temp directory

    $headers = "User STATLOOK`tUser POWERSHELL`tMapped Drive Letter`tPath`tHostname"
    $headers | Out-File -FilePath ".\mapped_disks.txt" -Encoding UTF8 -Append                                # create headers to txt file (used later when exporting to csv)
    $newLine = ""

    foreach($txtFile in $txtFiles){
        Get-Content ".\temp\$txtFile" | Out-File -FilePath ".\mapped_disks.txt" -Encoding UTF8 -Append       # append each temp txt files to the main txt file
        $newLine | Out-File -FilePath ".\mapped_disks.txt" -Encoding UTF8 -Append                            # append emtpy row to the main txt file
        Remove-Item ".\temp\$txtFile" -Force                                                                 # delete all txt files 
    }
}

function exportToCSV{       # export txt file to CSV for better and easier view
    if (Test-Path ".\mapped_disks.xlsx"){                   
        Remove-Item ".\mapped_disks.xlsx"                   # remove .XLSX file if exist
    }

    if (Test-Path ".\mapped_disks.csv"){                    # check if old version of .csv file exist (yes - remove it | no - do nothing)
        Remove-Item ".\mapped_disks.csv"

        Get-Content .\mapped_disks.txt >> .\mapped_disks.csv                        # make CSV from TXT
        $csv = Import-Csv ".\mapped_disks.csv"                                      # read data from CSV
        $xlsxPath = ".\mapped_disks.xlsx"
        $csv | Export-Excel -Path $xlsxPath -AutoSize                               # export CSV to XLSX
        
        Remove-Item ".\mapped_disks.txt"
        Remove-Item ".\temp"
        Remove-Item ".\mapped_disks.csv"
    }

    else{
        Get-Content .\mapped_disks.txt >> .\mapped_disks.csv                        # make CSV from TXT
        $csv = Import-Csv ".\mapped_disks.csv"                                      # read data from CSV
        $xlsxPath = ".\mapped_disks.xlsx"
        $csv | Export-Excel -Path $xlsxPath -AutoSize                               # export CSV to XLSX
        
        Remove-Item ".\mapped_disks.txt"
        Remove-Item ".\temp"
        Remove-Item ".\mapped_disks.csv" 
    }
}

$data = Get-Content ".\users.txt" -Encoding utf8   # load txt file with all domain users
TempDir                                                                                 # create temp dir to store txt files

foreach ($line in $data) {
    $dane_host = -split $line                               # create table [[username, ip_address, name, surname], ...]
    $hostname = $dane_host[0]                               # declare hostname - index 0 from the table above
    $ipAddress = $dane_host[1]                              # declare ip_address - index 1 from the table above
    $name = $dane_host[2]; $surname = $dane_host[3]         # declare name and surname - index 2 and 3 from the table above
    $user = $name + " " + $surname                          # declare username - concatenated index 2 and 3 from the table above

    if(checkConnection($ipAddress) -eq 1){
        checkMappedDisks($hostname)                         # check for mapped disks
    }
    else{
        Continue
    }
}

mergeTxt; exportToCSV;                                     