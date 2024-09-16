# PID VID tracer, gets the USB info on a windows machine and adds vendor and product text when available. See instructions at end how to run


$ProcessFolder = "E:\Temp\UsbCheck\" #!! Replace with your own
$ExcelFilePath = $ProcessFolder + "usb_devices.xlsx"
$CsvFilePath = $ProcessFolder + "usb_devices.csv"
Set-Location -Path $ProcessFolder

$idsFilePath = $ProcessFolder + "usb.ids"  
# download actual ids info uncomment if needed
# $url = "https://usb-ids.gowdy.us/usb.ids"
# Invoke-WebRequest -Uri $url -OutFile $idsFilePath

$fileContent = Get-Content $idsFilePath 

# Generate local computer USB info
$usbDevices = Get-PnpDevice | Where-Object {$_.InstanceId -match '^USB'}; $usbDevices | ft InstanceID, Class, FriendlyName


$allDevices = $usbDevices | ForEach-Object {
    $deviceID = $_.InstanceId
    $vid = if ($deviceID -match "VID_([0-9A-F]+)") { "$($matches[1])" } else { $null }
    $PID_ = if ($deviceID -match "PID_([0-9A-F]+)") { "$($matches[1])" } else { $null }
    $PidText=""
    
    $VIDline = 0
    
    $VIDpat = "^" + $VID

    # Search and output line VID vendor
    #$VidText = Select-String -Path $idsFilePath -Pattern $VIDpat
    $Vid1 = Select-String -Path $idsFilePath -Pattern $VIDpat
    $VidText = $Vid1 | ForEach-Object { $_.Line }

    $Vid1 | ForEach-Object {   # line in file for tests
        $VIDline = $($_.LineNumber)
    }

    if ($VIDText) { # VID found

        # Find next VID below 
        $pattern = "^[0-9A-Fa-f]{4}"  # any 4 digit hex pattern

        #Find end of VID entries
        $VIDline++ # Skip actual VIDline
        $linesToSearch = $fileContent[$VIDline..($fileContent.Count - 1)]

        # Find the first PID occurrence of 
        $PidT = $linesToSearch | Select-String -Pattern $PID_ | Select-Object -First 1
        #$PidText = $PidT | ForEach-Object { $_.Line.TrimStart() }
        $PidText = [string]$PidT.Line
        $PidText = $PidText.Trim()
        # If a match is found, display the line number and content
        if (-not $Pidtext) {Write-Output "No PID for this VID" } 
              
    }
        
    [PSCustomObject]@{
        DeviceName = $_.FriendlyName
        VID = $vid
        PID = $PID_
        Status = $_.Status
        VidName = $VidText
        PidName = $PidText
        DeviceID = $deviceID
    }
} | Sort-Object VID, PID_

$okDevices = $allDevices | Where-Object {
    $_.Status -notmatch "Unknown|Fehler"
} | Sort-Object VID, PID_


$allDevices | Export-Excel -Path $ExcelFilepath -WorksheetName "AllDevices2" -AutoSize  # Excel report, comment out if not used
$allDevices | Export-Csv -Path $CsvFilepath -NoTypeInformation -Encoding UTF8


# Mercurial / Tortoisehg is used to see differences and history, use your own VCS or diff if conveniant
# Mercurial init
# hg ini
# hg add

# 1st run connect device -> comment in first hg line ->  comment out 2nd hg line -> run script
# 2nd run disconnecet desired device -> comment out first hg line (#) -> comment in 2nd hg line -> start script again gain

hg ci -m "USB List with connected device"
#hg ci -m "USB device unconnected"
