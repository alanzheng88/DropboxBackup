# Prerequisites for backing up with zip files:
# 	Ensure chocolatey is installed by running the following:
# 	iex ((new-object net.webclient).DownloadString('https://chocolatey.org/install.ps1'))
# 	Install Write-Zip by running the following in powershell:
# 	cinst pscx

$ErrorActionPreference = 'Stop'

# Backup specific folders
$foldersToBackupPaths = @(
	"D:\Coop"
	"D:\books"
	"D:\cmpt 213"
)

# Exclude the following files from being backed up
$exclusions = @(
	"D:\cmpt 275\CMPT275 SVN"
)

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
	$diff | %{ Copy-Item $_.FullName -Destination (Get-DropboxFolder $dropboxLocation $_.FullName) -Force }
	$diff | Format-List -Property FullName, LastWriteTime
	$endTime = Get-Date -Format g
	Write-Host "`r`nFinished backing up -- Time: $endTime`r`n"
}

# Main
$stdoutLog = "C:\Users\Alan\Desktop\backuplog.txt"
Start-Transcript -Path $stdoutLog
$dropboxLocation = Get-DropBox
Check-AllFoldersExist $foldersToBackupPaths
Create-MissingDropboxFolders $foldersToBackupPaths $dropboxLocation
Backup $foldersToBackupPaths $dropboxLocation
Stop-Transcript
Read-Host -Prompt "Press enter to exit"