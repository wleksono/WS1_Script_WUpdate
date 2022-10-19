$LogFilePath = "C:\Temp\ws1"
if (!(Test-Path -Path $LogFilePath))
{
    New-Item -Path $LogFilePath -ItemType Directory | Out-Null
}
    
$Logfile = $LogFilePath+"\installUpdates.log"
    
Function Log([string]$level, [string]$logstring)
{
    $rightSide = [string]::join("   ", ($level, $logstring))

    $date = Get-Date -Format g
    $logEntry = [string]::join("    ", ($date, $rightSide)) 
    Add-content $Logfile -value $logEntry
}

Log "Info" "Update Script Start"

$Sysinfo = New-Object -ComObject Microsoft.Update.SystemInfo
$pending = $Sysinfo.RebootRequired

if ($pending){
	try{
	    shutdown.exe /r /f /t 120
	}
	catch{
		Log "Error" "$($_.Exception)"
	    exit 1
	}
	Log "Warning" "Rebooting Device..."
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
	Log "Error" "$($_.Exception)"
	exit 1
}

Log "Success" "Found $($SearchResult.count) Update(s)"

if ((Get-Service DoSvc).Status -eq "Stopped") {Start-Service DoSvc | Out-Null}

if ((Get-Service DoSvc).Status -eq "Running") {Restart-Service DoSvc -Force | Out-Null}

$retrycount = 3

if($SearchResult.count -ne 0){
	foreach ($entry in $SearchResult){	
		$updateSession = New-Object -ComObject 'Microsoft.Update.Session'
		$updatesToDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl'
		$updatesToDownload.Add($entry) | Out-Null

		Log "Start" "KB$($entry.KBArticleIDs) Download Processing..."

		$a=0
		do{
			$a++
			$downloader = $updateSession.CreateUpdateDownloader()
			$downloader.ClientApplicationID = "InstallUpdate"
			$downloader.IsForced = $True
			$downloader.Updates = $updatesToDownload

			Log "Start" "KB$($entry.KBArticleIDs) Download Starting..."

			try{
				$downloadResult = $downloader.Download()
			}catch{Log "Warning" "$($_.Exception)"}

			if ($downloadResult.ResultCode -eq 2){
				Log "Success" "KB$($entry.KBArticleIDs) Downloaded"
			}
			elseif ($a -ge $retrycount){
				Log "Error" "KB$($entry.KBArticleIDs) Download retry limit reached"
			}
			else{
				Log "Warning" "KB$($entry.KBArticleIDs) Download retry $($a)"
			}
			
		}Until ($a -ge $retrycount -or $downloadResult.ResultCode -eq 2)

		Clear-Variable updatesToDownload -Force -ErrorAction SilentlyContinue
	}

	Clear-Variable entry -Force -ErrorAction SilentlyContinue

	foreach ($entry in $SearchResult){
		$updatesToInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
		$updatesToInstall.Add($entry) | Out-Null

		Log "Start" "KB$($entry.KBArticleIDs) Install Processing..."

		$a=0
		do{
			$a++

			$installer = New-Object -ComObject 'Microsoft.Update.Installer'
			$installer.ClientApplicationID = "InstallUpdate"
			$installer.IsForced = $True         
			$installer.Updates = $updatesToInstall

			if(!$entry.IsDownloaded){
				Log "Error" "KB$($entry.KBArticleIDs) not Downloaded"
				break
			}

			Log "Start" "KB$($entry.KBArticleIDs) Install Starting..."

			try{
				$installResult = $installer.Install()
			}catch {Log "Warning" "$($_.Exception)"}

			if ($installResult.ResultCode -eq 2 -or $installResult.ResultCode -eq 7){
				Log "Success" "KB$($entry.KBArticleIDs) Installed"
			}
			elseif ($a -ge $retrycount){
				Log "Error" "KB$($entry.KBArticleIDs) Install retry limit reached"
			}
			else{
				Log "Warning" "KB$($entry.KBArticleIDs) Install retry $($a)"
			}

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
			Log "Error" "$($_.Exception)"
			exit 1
		}
		Log "Info" "Rebooting Device after Update..."
	}
}
Log "Info" "Update Script End"
exit