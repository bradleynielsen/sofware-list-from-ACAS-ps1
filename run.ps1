<# # Description:
Prototype script to compile list of software identified by Nessus plugin 20811.

# Instructions:
1. create a folder called 'csv' in root repo folder
    _C:\Program Files\Development_Software\Developer Solutions\PowerShell\sofware-list-from-ACAS-ps1\csv_

2. add ACAS csv report files into ~\sofware-list-from-ACAS-ps1\csv

3. run 'run.ps1' by right clicking and selecting "Run with PowerShell"
    _C:\Program Files\Development_Software\Developer Solutions\PowerShell\sofware-list-from-ACAS-ps1\run.ps1_

4. Your results will show in ~\sofware-list-from-ACAS-ps1\results as two files:
    _A list of each instance of software installed on each host, for each repository_    
    _A list of of unique software names installed across for all hosts_ #>

Add-Type -AssemblyName PresentationFramework
$scriptRoot     = $PSScriptRoot
$csvpath        = "$scriptRoot\csv"
$resultspath    = "$scriptRoot\results\"


#init
$date = Get-Date -Format yyyy-MM-dd_mmss
$allResultsFilePath     = $resultspath + "host_level - " + $date+".csv"
$uniqueResultsFilePath  = $resultspath + "system_level - " + $date+".csv"
$softwareUseTable       = @()
$softwareNamesTable     = @()
$uniqueSoftwareTable    = @()

if (!(test-path -path $csvpath)) {
    mkdir $csvpath
    "Making csv folder"
    [System.Windows.MessageBox]::Show('`csv` folder created. Add your ACAS reports to it before continuing.', 'Missing CSV files')
}

if (!(test-path -path $resultspath)) {
    mkdir $resultspath
    "Making results folder"
}


"Starting Import"
Get-ChildItem $csvpath  -Filter *.csv | 
Foreach-Object {
    "Importing $_"
    $csvImport = Import-Csv $_.FullName -WarningAction Continue
    foreach ($line in $csvImport) {
        if ($line.plugin -eq "20811") {
            
            $hostname   = $line."NetBIOS Name"
            $repository = $line."Repository"
            $pluginText = $line."Plugin Text"
                        
            foreach ($ptLine in ($pluginText -split "\r?\n|\r")) {
                $softwareLine = $ptLine.Trim()
                
                #test if line is sw or header
                if ($softwareLine.Contains(']')) {
                    #split the line by bracket
                    $softwareName = $softwareLine.split("[")[0]
                    
                    #prevent null sw version from throwing error
                    try {
                        $softwareVersion = ($softwareLine.split("[")[1]).replace("version ", "").replace("]", "").trim() 
                    }
                    catch {
                        #the line has no version
                        $softwareVersion = ""
                    }
                    $softwareUseTable += [PSCustomObject]@{
                        softwareName    = $softwareName
                        softwareVersion = $softwareVersion
                        hostname        = $hostname
                        repository      = $repository
                    }
                    $softwareUseTable

                }
            }
        }  
    }
}


#Build unique software list
"Building unique software list"
foreach ($row in $softwareUseTable) {
    $softwareNamesTable += [PSCustomObject]@{
        softwareName    = $row.softwareName
        softwareVersion = $row.softwareVersion
    }
}
#sort, and remove duplicates
"Sorting and removing duplicates"
$uniqueSoftwareTable = $softwareNamesTable | sort-object -Property softwareName -Unique


"Software count for all sites: " + $softwareUseTable.Length
"Unique software count: "        + $uniqueSoftwareTable.Length

"Exporting reports"
$softwareUseTable    | Export-Csv -Path $allResultsFilePath -NoTypeInformation
$uniqueSoftwareTable | Export-Csv -Path $uniqueResultsFilePath -NoTypeInformation


explorer $scriptRoot\results

