##Office365 Office 2013 Installer##
##Adam Witt 2015##
##Version 1.1.2##

##Select version using a number for the install##
write-Host "Please Select From The Following" -foregroundcolor "black" -backgroundcolor "white"
Write-Host "1) Professional Plus" -foregroundcolor "magenta"
write-Host "2) Business Premium" -foregroundcolor "magenta"
$version = Read-Host "What version would you like to install? (number)"


##Set the version selected in the last step##
if ($version -eq 1) {$xml = 'pro'
	Write-Host "Office 2013 Professional Plus Selected..." -foregroundcolor "yellow"
}
else {if ($version -eq 2) {$xml = 'business'
		Write-Host "Office 2013 Home & Business Selected..." -foregroundcolor "yellow"
	}
	else {clear
		write-host "Invalid option selected!!!" -foregroundcolor "red"
		./"Office365 Installer.ps1"
		exit
	}
}

##Checking for the installer files##
write-Host "Checking for Install Files..." -foregroundcolor "green"
$path = convert-path .
$pathcat = $path  + "\Office"
$test = test-Path $pathcat

#Checking if files exist and downloading if needed##
if($test -eq $False) {Write-Host "Downloading Install Files..." -foregroundcolor "green"
	if($xml -eq 'pro') {powershell.exe -noprofile -executionpolicy bypass -file '.\pro.ps1'}
	else {if ($xml -eq 'business') {powershell.exe -noprofile -executionpolicy bypass -file '.\business.ps1'}
		else {write-Host "ERROR" -ForegroundColor "darkred"}
		exit
	}
	else {Write-Host "Setup Files Exist..." -foregroundcolor "green"
	if($xml -eq 'pro') {powershell.exe -noprofile -executionpolicy bypass -file '.\proinst.ps1'}
	else {if ($xml -eq 'business') {powershell.exe -noprofile -executionpolicy bypass -file '.\businessinst.ps1'}
		else {write-Host "ERROR" -ForegroundColor "darkred"}
		exit
	}

	##Installing Office 2013##
	#if($test -eq $False) {if($xml = 'pro') {powershell.exe -noprofile -executionpolicy bypass -file '.\proinst.ps1'
	#	}
	#}
	#else {write-Host "Installing Office..." -foregroundcolor "green"
	#	./setup.exe /configure /$xml
	#	write-Host "Installing Office Finished!!!" -foregroundcolor "green"
	#	exit
	#}

#	##Check after file download
#	$path = convert-path .
#	$pathcat = $path  + "\Office"
#	$test = test-Path $pathcat

##Check to make sure the download was successful##
if($test -eq $True) {Write-Host "Download Failed!!!" -foregroundcolor "Red"
	start-sleep -s 4
	clear
	Write-Host "Attempting to restart Installer..." -foregroundcolor "darkyellow"
	start-sleep -s 4
	clear
	./"Office365 Installer.ps1"
}

	##Install Office
	#else {Write-Host "Setup Files Downloaded..." -foregroundcolor "green"
	#	write-Host "Installing Office..." -foregroundcolor "green"
	#	./setup.exe /configure /$xml
	#	write-Host "Installing Office Finished!!!" -foregroundcolor "green"
	#	exit
}
}