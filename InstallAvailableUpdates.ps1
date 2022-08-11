$Sysinfo = New-Object -ComObject Microsoft.Update.SystemInfo
$pending = $Sysinfo.RebootRequired
if ($pending){ 
	shutdown.exe /r /f /t 120
	exit 55
}

$testnet = Test-NetConnection -ComputerName www.catalog.update.microsoft.com -CommonTCPPort HTTP
if($testnet.TcpTestSucceeded -eq "True"){}Else{exit 44}

$Session = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session"))#,$Computer))
$UpdateSearcher = $Session.CreateUpdateSearcher()

$Criteria = "IsHidden=0 and IsInstalled=0 and IsAssigned=1"
$SearchResult = $UpdateSearcher.Search($Criteria).Updates

if($SearchResult.count -ne 0){
    foreach ($entry in $SearchResult){
	    if ($entry.isDownloaded -eq $false){		
		$updateSession = New-Object -ComObject 'Microsoft.Update.Session'
		$updatesToDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl'
		$updatesToDownload.Add($entry) | Out-Null
        	$downloader = $updateSession.CreateUpdateDownloader()
        	$downloader.Updates = $updatesToDownload
        	$downloadResult = $downloader.Download()
		if ($downloadResult.ResultCode -eq 2) {
			$updatesToInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
			$updatesToInstall.Add($entry) | Out-Null
			
			$installer = New-Object -ComObject 'Microsoft.Update.Installer'
			$installer.Updates = $updatesToInstall        
			$installResult = $installer.Install()
				
			Clear-Variable updatesToInstall -Force -ErrorAction SilentlyContinue
		    }
		Clear-Variable updatesToDownload -Force -ErrorAction SilentlyContinue
	    }
	    else{
		$updatesToInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
		$updatesToInstall.Add($entry) | Out-Null 

        	$installer = New-Object -ComObject 'Microsoft.Update.Installer'
        	$installer.Updates = $updatesToInstall        
        	$installResult = $installer.Install()
		Clear-Variable updatesToInstall -Force -ErrorAction SilentlyContinue
	    }
    }
    $Sysinfo = New-Object -ComObject Microsoft.Update.SystemInfo
    $pending = $Sysinfo.RebootRequired
    if ($pending) { shutdown.exe /r /f /t 120 }
    exit
}
else{
    exit
}
