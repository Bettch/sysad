function Get-InstalledApps {

	$NAME = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
	$NAME = $NAME -replace '\\','_'
	Write-Output "Generating Report for $NAME"

	$REG_PATH = @()

	if (Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"){
		$REG_PATH += "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
	}
	if (Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"){
		$REG_PATH += "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
	}
	if (Test-Path "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"){
		$REG_PATH += "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
	}
	if (Test-Path "HKCU:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"){
		$REG_PATH += "HKCU:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
	}

	$OUTPUT = Get-ItemProperty $REG_PATH | 
				Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, SystemComponent, ParentDisplayName | 
					Where-Object { $_.SystemComponent -ne 1 -and $_.DisplayName -and !$_.ParentDisplayName -and ($_ -match 
									# Filter for desired apps for uninstallation
									'Google Chrome|Libre|Opera|Adobe|FireFox|Mozilla|WPS Office') } |
						Sort-Object Publisher, DisplayName

	$count = 1
	$OUTPUT = $OUTPUT | ForEach-Object {
		$_ | Select-Object @{Name = 'Line'; Expression = {$count}}, *
		$count++
	}

	$OUTPUT | Format-Table -AutoSize

	$OUTPUT_stats = $OUTPUT | Measure-Object

	Write-Output "Found $(($OUTPUT_stats).Count) installed applications"
	Write-Output "Done"
	
	$DisplayNames = $OUTPUT | Select-Object -ExpandProperty DisplayName
	$DisplayNames
}

function Get-DefaultApps {
	$DefaultApps = Get-AppxPackage -AllUsers |
		Select-Object Name, Version, Publisher |
			Format-Table -AutoSize

	$DefaultApps
}

Get-InstalledApps
Get-DefaultApps