##Office365 Office 2013 Installer##
##Adam Witt 2015##
##Version 1.1.2##

##Select Version
write-Host "Please Select From The Following" -foregroundcolor "black" -backgroundcolor "white"
Write-Host "1) Professional Plus" -foregroundcolor "magenta"
write-Host "2) Business Premium" -foregroundcolor "magenta"
$version = Read-Host "What version would you like to install? (number)"


##Version Set
if ($version -eq 1) {$xml = 'pro.bat'
	Write-Host "Office 2013 Professional Plus Selected..." -foregroundcolor "yellow"
}
else {if ($version -eq 2) {$xml = 'business.bat'
		Write-Host "Office 2013 Home & Business Selected..." -foregroundcolor "yellow"
	}
	else {clear
		write-host "Invalid option selected!!!" -foregroundcolor "red"
		./"Office365 Installer.ps1"
		exit
	}
}

##Checking for Installer files
write-Host "Checking for Install Files..." -foregroundcolor "green"
$path = convert-path .
$pathcat = $path  + "\Office"
$test = test-Path $pathcat

#Downloading/Installing depending on previous
if($test -eq $False) {Write-Host "Downloading Install Files..." -foregroundcolor "green"
}
else {Write-Host "Setup Files Exist..." -foregroundcolor "green"
}
##Installing Office
##powershell.exe -noprofile -executionpolicy bypass -file '.\Office365 Installer.ps1'
if($test -eq $False) {./setup.exe /download /$xml
}
else {write-Host "Installing Office..." -foregroundcolor "green"
	start-Process $xml
	write-Host "Installing Office Finished!!!" -foregroundcolor "green"
	exit
}

##Check after file download
$path = convert-path .
$pathcat = $path  + "\Office"
$test = test-Path $pathcat

##If download fails
if($test -eq $True) {Write-Host "Download Failed!!!" -foregroundcolor "Red"
	start-sleep -s 3
	clear
	Write-Host "Attempting to restart Installer..." -foregroundcolor "darkyellow"
	start-sleep -s 4
	clear
	Write-Host "Installer Restarted" -backgroundcolor "green"
	./"Office365 Installer.ps1"
}

##Install Office
else {Write-Host "Setup Files Downloaded..." -foregroundcolor "green"
	write-Host "Installing Office..." -foregroundcolor "green"
	start-Process $xml
	write-Host "Installing Office Finished!!!" -foregroundcolor "green"
	exit
}

