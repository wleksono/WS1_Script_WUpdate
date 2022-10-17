$Sysinfo = New-Object -ComObject Microsoft.Update.SystemInfo
$pending = $Sysinfo.RebootRequired

if ($pending){
	try{
	    shutdown.exe /r /f /t 120
	}
	catch{
	    exit 1
	}
	exit 2
}

Clear-Variable Sysinfo -Force -ErrorAction SilentlyContinue
Clear-Variable pending -Force -ErrorAction SilentlyContinue

$Session = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session"))#,$Computer))
$UpdateSearcher = $Session.CreateUpdateSearcher()

$Criteria = "IsHidden=0 and IsInstalled=0 and IsAssigned=1"

try{

    $SearchResult = $UpdateSearcher.Search($Criteria).Updates

}catch{
	"Update Search Failed"
	exit 1
}

if ((Get-Service DoSvc).Status -eq "Stopped") {Start-Service DoSvc}

if ((Get-Service DoSvc).Status -eq "Running") {Restart-Service DoSvc -Force}

$retrycount = 3

if($SearchResult.count -ne 0){
	foreach ($entry in $SearchResult){	
		$updateSession = New-Object -ComObject 'Microsoft.Update.Session'
		$updatesToDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl'
		$updatesToDownload.Add($entry) | Out-Null

		$a=0
		do{
			$a++
			$downloader = $updateSession.CreateUpdateDownloader()
			$downloader.ClientApplicationID = "InstallUpdate"
			$downloader.IsForced = $True
			$downloader.Updates = $updatesToDownload

			try{
				$downloadResult = $downloader.Download()
			}catch{"Attempt to download update failed"}

		}Until ($a -ge $retrycount -or $downloadResult.ResultCode -eq 2)

		Clear-Variable updatesToDownload -Force -ErrorAction SilentlyContinue
	}

	Clear-Variable entry -Force -ErrorAction SilentlyContinue

	foreach ($entry in $SearchResult){
		$updatesToInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
		$updatesToInstall.Add($entry) | Out-Null

		$a=0
		do{
			$a++

			$installer = New-Object -ComObject 'Microsoft.Update.Installer'
			$installer.ClientApplicationID = "InstallUpdate"
			$installer.IsForced = $True         
			$installer.Updates = $updatesToInstall

			try{
				$installResult = $installer.Install()
			}catch{"Attempt to install update failed"}
		}Until($a -ge $retrycount -or $installResult.ResultCode -eq 2 -or $installResult.ResultCode -eq 7)

		Clear-Variable updatesToInstall -Force -ErrorAction SilentlyContinue
	}

	$Sysinfo = New-Object -ComObject Microsoft.Update.SystemInfo
	$pending = $Sysinfo.RebootRequired
	if ($pending) {
		try{
			shutdown.exe /r /f /t 120
		}
		catch{
			exit 1
		}
	}
}
exit
