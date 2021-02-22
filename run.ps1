Add-Type -AssemblyName PresentationFramework
$scriptRoot     = $PSScriptRoot
$csvpath        = "$scriptRoot\csv"
$resultspath    = "$scriptRoot\results\"


#init
$date = Get-Date -Format yyyy-MM-dd_mmss
$allResultsFilePath     = $resultspath + "all_results - " + $date+".csv"
$uniqueResultsFilePath  = $resultspath + "unique_results - " + $date+".csv"
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

