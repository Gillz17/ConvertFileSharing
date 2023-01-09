#=====================================================
#Created by: Zach McGill
#Created: 9/8/2021
#ConvertFileScanningData.ps1
#Converts scanner output into ProLaw readable data to be uploaded
#Also moves all of the old files to the archive folder in the user's filescanning drive
#=====================================================

#The path for the error file
$errorFile = "H:\filescanning\scan_errorlog.txt"
#New-Item -ItemType File -Force -Path $errorFile

if((Test-Path "H:\filescanning\upload.txt") -and (Get-Content "H:\filescanning\upload.txt")){
	$isEmpty = $false
	$override = Read-Host "The upload file is not empty, do you want to overwrite the data(y/n)?"
	if($override = "y"){
		Clear-Content "H:\filescanning\upload.txt"
	}
}else{
	$isEmpty = $true
}

#Read in each line of raw_data and process it
foreach($line in Get-Content "H:\filescanning\raw_data.txt"){
	#If the line starts with $N, its a new user scanning
	if($line.Contains('$N')){
        Write-Host "NAME: " $line -ForegroundColor Red
		$name = $line[2..300] -join ''
	#If the line starts with $L, we are at a new cabinet location
	}elseif($line.Contains('$L')){
        Write-Host "LOCATION: " $location -ForegroundColor Red		
        $location = $line[2..300] -join ''
		#Cabinets do not have a specific location so we need to clear this variable
		$specificLocation = ""
		$hasSpecLoc = $false
		#Check to see if the location is called RECALL, if it is set the boolean to true
		if($line.contains("RECALL")){
		    $hasSpecLoc = $true 
		}
	}else{
		#The location is in RECALL so the next line is the specific location
		if($hasSpecLoc -eq $true){
			$specificLocation = $line
			$hasSpecLoc = $false
		#Checks to see if we have a user, location, and matter number otherwise skip the line
		}elseif($name -eq "" -OR $location -eq ""){
			$emptyUserLocation = "Name or Location not set for line " + ($line.Count)
			$emptyUserLocation | Add-Content -Path $errorFile
		}else{
			$entry = $location + "|" + $specificLocation + "|" + $(Get-Date -Format "MM/dd/yyyy") + "|" + $name + "|" + $line
			$entry | Add-Content -Path "H:\filescanning\upload.txt"
		}
	}
}

Write-Host "Completed Conversion!`n" -ForegroundColor Green

#Pauses for the user to upload the file to ProLaw before we move it
Write-Host "Run ProLaw Import Now" -ForegroundColor Yellow
pause

#Check to see if the archive folder exists
if(Test-Path "H:\filescanning\archive"){
	Write-Host "The archive folder already exists" -ForegroundColor Blue
}else{
	#Create the archive folder if it does not exist
	New-Item -ItemType Directory -Force -Path "H:\filescanning\archive"
}

#Moves the items to the archive folder
#Also checks to increment the file name so that it increments with the files
Write-Host "Archiving Files" -ForegroundColor Yellow

#All of the new file names/paths
$archiveUploadFile = "H:\filescanning\archive\$(Get-Date -Format yyyy-MM-dd)-upload.txt"
$archiveRejectionFile = "H:\filescanning\archive\$(Get-Date -Format yyyy-MM-dd)-uploadRej.txt"
$archiveLogFile = "H:\filescanning\archive\$(Get-Date -Format yyyy-MM-dd)-uploadLog.txt"
$archiveScanErrorFile = "H:\filescanning\archive\$(Get-Date -Format yyyy-MM-dd)-scan_errorlog.txt"

#The original file names/paths
$orgUploadFile = "H:\filescanning\upload.txt"
$orgRejectionFile = "H:\filescanning\upload.rej"
$orgLogFile = "H:\filescanning\upload.log"
$orgScanErrorFile = "H:\filescanning\scan_errorlog.txt"

$i = 0
#Determine if the files exist
While ((Test-Path $archiveUploadFile) -or (Test-Path $archiveRejectionFile) -or 
            (Test-Path $archiveLogFile) -or (Test-Path $archiveScanErrorFile)){
    $i += 1
    $archiveUploadFile = "H:\filescanning\archive\$(Get-Date -Format yyyy-MM-dd)-upload ($i).txt"
    $archiveRejectionFile = "H:\filescanning\archive\$(Get-Date -Format yyyy-MM-dd)-uploadRej ($i).txt"
    $archiveLogFile = "H:\filescanning\archive\$(Get-Date -Format yyyy-MM-dd)-uploadLog ($i).txt"
    $archiveScanErrorFile = "H:\filescanning\archive\$(Get-Date -Format yyyy-MM-dd)-scan_errorlog ($i).txt"
}
#Move the files to their new homes in the archive folder
Move-Item -Path $orgUploadFile -Destination $archiveUploadFile
Move-Item -Path $orgRejectionFile -Destination $archiveRejectionFile
Move-Item -Path $orgLogFile -Destination $archiveLogFile
Move-Item -Path $orgScanErrorFile -Destination $archiveScanErrorFile

Write-Host "Clear raw_data"
Clear-Content "H:\filescanning\raw_data.txt"

Write-Host "All items completed" -ForegroundColor Green

Write-Host -NoNewLine "Closing in: "
$x = 5
do{
    Write-Host -NoNewLine $x".."
    Start-Sleep 1
    $x -= 1
}while($x -ne 0)
exit