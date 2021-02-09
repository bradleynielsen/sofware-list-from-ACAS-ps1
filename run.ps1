Add-Type -AssemblyName PresentationFramework
$scriptRoot     = $PSScriptRoot
$csvpath        = "$scriptRoot\csv"
$resultspath    = "$scriptRoot\results"


#init
$date = Get-Date -Format yyyy-MM-dd
$resultsFilePath     = $resultspath + "results - " + $date+".csv"
$softwareUseTable    = @()
$uniqueSoftwareTable = @()

if (!(test-path -path $csvpath)) {
    mkdir $csvpath
    "Making csv folder"
    [System.Windows.MessageBox]::Show('`csv` folder created. Add your ACAS reports to it before continuing.', 'Missing CSV files')
}

if (!(test-path -path $resultspath)) {
    mkdir $resultspath
    "Making results folder"
}


#import scans
Get-ChildItem $csvpath  -Filter *.csv | 
Foreach-Object {
    $csvImport = Import-Csv $_.FullName
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

$softwareUseTable

foreach ($row in $softwareUseTable) {
    #
    
}


#explorer $scriptRoot\results
#pause
exit