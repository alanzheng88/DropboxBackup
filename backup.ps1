# Prerequisites for backing up with zip files: (will be used sometime in the future)
# 	Ensure chocolatey is installed by running the following:
# 	iex ((new-object net.webclient).DownloadString('https://chocolatey.org/install.ps1'))
# 	Install Write-Zip by running the following in powershell:
# 	cinst pscx

$ErrorActionPreference = 'Stop'

# environment variables set for user on localhost
$foldersToBackupPaths = (Get-Content $($env:backupListTxt)) -Split '`r`n'
$backupCredsContent = Get-Content $($env:backupCredsTxt)
$creds = $backupCredsContent.Split(';')
$username = $creds[0]
$password = $creds[1]
$emailFrom = $creds[0]
$emailTo = $creds[2].Split(',')
$stmpServer = $creds[3]
$smtpPort = $creds[4]

function Get-DropBox {
	$hostFile = Join-Path (Split-Path (Get-ItemProperty HKCU:\Software\Dropbox).InstallPath) "host.db"
	$encodedPath = [System.Convert]::FromBase64String((Get-Content $hostFile)[1])
	return [System.Text.Encoding]::UTF8.GetString($encodedPath)
}

function Check-AllFoldersExist ($foldersToBackupPaths) {
	$foldersNotExist = @()
	$foldersToBackupPaths | %{
		$folderPath = $_
		if (-not (Test-Path $folderPath)) {
			$foldersNotExist += $folderPath
		}
	}
	
	if ($foldersNotExist.Count -gt 0) {
		Write-Host "Please check if all folders to backup exist. The following folders cannot be found."
		Write-Host $foldersNotExist
		exit -1
	}
}

function Get-DropboxFolder($dropboxLocation, $fullFilePath) {
	$fullFilePath = $fullFilePath -replace '\w{1}:', "$dropboxLocation"
	$parentPath = Split-Path -Path $fullFilePath -Parent
	return $parentPath
}

function Create-MissingDropboxFolders($foldersToBackupPaths, $dropboxLocation) {
	Get-ChildItem $foldersToBackupPaths -Recurse -File | %{
		$dropboxFolder = Get-DropboxFolder $dropboxLocation $_.FullName
		if (-not (Test-Path $dropboxFolder)) {
			Write-Host "Creating $dropboxFolder because it does not exist`r`n"
			New-Item $dropboxFolder -Type Directory | Out-Null
		}
	}
}

function Backup($foldersToBackupPaths, $dropboxLocation) {
	$startTime = Get-Date -Format g
	Write-Host "`r`nBacking up the following folders -- Time: $startTime`r`n"
	Write-Host $foldersToBackupPaths
	$originalFiles = @(Get-ChildItem $foldersToBackupPaths -Recurse -File)
	$dropboxFiles = @(Get-ChildItem $dropboxLocation -Recurse -File)
	$diff = Compare-Object -ReferenceObject $originalFiles -DifferenceObject $dropboxFiles -property Name, LastWriteTime -PassThru |
		Where-Object { $_.SideIndicator -eq "<=" }
	if ([string]::IsNullOrEmpty($diff)) {
		$endTime = Get-Date -Format g
		Write-Host "`r`nNothing to backup -- Time: $endTime`r`n"
		Exit 0
	}
	$diff | %{ Copy-Item $_.FullName -Destination (Get-DropboxFolder $dropboxLocation $_.FullName) -Force }
	$diff | Format-List -Property FullName, LastWriteTime
	$endTime = Get-Date -Format g
	Write-Host "`r`nFinished backing up -- Time: $endTime`r`n"
}

function Send-Log($from, $to, $attachment, $body, $smtpServer, $smtpPort, $username, $password) {
	$subject = "Dropbox Backup Automation Log"
	$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
	$credentials = New-Object System.Management.Automation.PSCredential($username, $securePassword)
	Write-Host "Sending logs to the following: $to"
	Send-MailMessage -From $from -To $to -Subject $subject -Body $body -SmtpServer $smtpServer -port $smtpPort -UseSsl -Credential $credentials
	Write-Host "Sending logs complete"
}

# Main
$stdoutLog = "C:\Users\Alan\Desktop\backuplog.txt"
Start-Transcript -Path $stdoutLog | Out-Null
$dropboxLocation = Get-DropBox
Check-AllFoldersExist $foldersToBackupPaths
Create-MissingDropboxFolders $foldersToBackupPaths $dropboxLocation
Backup $foldersToBackupPaths $dropboxLocation
Stop-Transcript | Out-Null
$log = Get-Content $stdoutLog | Out-String
Send-Log $emailFrom $emailTo $stdoutLog $log $stmpServer $smtpPort $username $password
# Read-Host -Prompt "Press enter to exit"