Get-Service bits
Get-Service wuauserv   
Get-Service cryptsvc

Stop-Service bits
Stop-Service wuauserv   
Stop-Service cryptsvc

Start-Service bits
Start-Service wuauserv   
Start-Service cryptsvc

Get-Service bits
Get-Service wuauserv   
Get-Service cryptsvc

Write-Host 'Scanning Please wait...'

Function Get-Hash($Path){
    
    $Stream = New-Object System.IO.FileStream($Path,[System.IO.FileMode]::Open) 
    
    $StringBuilder = New-Object System.Text.StringBuilder 
    $HashCreate = [System.Security.Cryptography.HashAlgorithm]::Create("SHA256").ComputeHash($Stream)
    $HashCreate | Foreach {
        $StringBuilder.Append($_.ToString("x2")) | Out-Null
    }
    $Stream.Close() 
    $StringBuilder.ToString() 
}

$DataFolder = "$env:ProgramData\WSUS Offline Catalog"
$CabPath = "C:\UpdateScaner\wsusscn2.cab"

# Create download dir
mkdir $DataFolder -Force | Out-Null

# Check if cab exists
$CabExists = Test-Path $CabPath


Write-Verbose "Creating Windows Update session"
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateServiceManager  = New-Object -ComObject Microsoft.Update.ServiceManager 

$UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $CabPath, 1) 

Write-Verbose "Creating Windows Update Searcher"
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()  
$UpdateSearcher.ServerSelection = 3
$UpdateSearcher.ServiceID = $UpdateService.ServiceID.ToString()
 
Write-Verbose "Searching for missing updates"
$SearchResult = $UpdateSearcher.Search("IsInstalled=0")

$Updates = $SearchResult.Updates

$UpdateSummary = [PSCustomObject]@{

    ComputerName = $env:COMPUTERNAME    
    MissingUpdatesCount = $Updates.Count
    Vulnerabilities = $Updates | Foreach {
        $_.CveIDs
    }
    MissingUpdates = $Updates | Select Title, Description, MsrcSeverity, @{Name="KBArticleIDs";Expression={$_.KBArticleIDs}} | ConvertTo-Html | Out-File -FilePath "C:\UpdateScaner\result.html"
}

Return $UpdateSummary