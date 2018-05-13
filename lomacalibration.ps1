<#$notepad = "HKCU:\Software\Microsoft\Notepad"
Set-ItemProperty -Path $notepad -Name lfFaceName -Value "Courier New"
Set-ItemProperty -Path $notepad -Name iPointSize -Value 120#>

$port = new-Object System.IO.Ports.SerialPort COM4, 115200, None, 8, one
$port.ReadTimeout = 15000

Function OpenPort(){
	Try {
		$port.open()
	}
	Catch {
		Write-Host Access to port is denied -ForegroundColor Red
		$exit = Read-Host -Prompt 'Press any key to exit'
		Break
	}
}

Function ReadLine(){
	Try {
		return $port.ReadLine()
		#$temp = $port.ReadLine()
		#Write-Host $temp
		#return $temp
}
	Catch {
		Write-Host The ReadLine operation has timed out
		$exit = Read-Host -Prompt 'Press any key to exit'
		Break}
}

Function Clear0([string]$command){
	$port.WriteLine($command)
	$temp = ReadLine
	$temp = ReadLine
	$temp = ReadLine
	$temp = ReadLine
	$temp = ReadLine
	$temp = ReadLine
}

Function Save0(){
	$port.WriteLine('SAVE')
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	$temp = $port.ReadLine()
	Write-Host 'Save to flash complete'
}

Function Calibrate([string]$command){
	$port.WriteLine($command)	
	$temp = ReadLine
	$temp = ReadLine
	$temp = ReadLine
	$port.DiscardInBuffer()
	$port.DiscardOutBuffer()
}

Function GetRegisterValue([string]$command){
	$port.WriteLine($command)
	$temp = ReadLine
	$temp = ReadLine
	$temp = ReadLine
	$value = ReadLine
	$temp = ReadLine
	$port.DiscardInBuffer()
	$port.DiscardOutBuffer()
	return $value
}

Function Extract([string]$data){
	$ind_L = $data.IndexOf('Low') + 9
	$ind_M = $data.IndexOf('Mid') + 8
	$ind_H = $data.IndexOf('High') + 9
	$Low = $data.Substring($ind_L, 2).Replace(',','')
	$Mid = $data.Substring($ind_M, 2).Replace(',','')
	$High = $data.Substring($ind_H, $data.Length - $ind_H - 1).Replace(',','')
	if ($Low.Length -eq 1) {$Low = (-join('0', $Low))}
	if ($Mid.Length -eq 1) {$Mid = (-join('0', $Mid))}
	if ($High.Length -eq 1) {$High = (-join('0', $High))}
	$hex = (-join($High, ' ', $Mid, ' ', $Low))
	return $hex
}

Function Hex_Dec_Digit([string]$hex){
	if ($hex -eq 'a') { return 10 }
	if ($hex -eq 'b') { return 11 }
	if ($hex -eq 'c') { return 12 }
	if ($hex -eq 'd') { return 13 }
	if ($hex -eq 'e') { return 14 }
	if ($hex -eq 'f') { return 15 }
	return [int]$hex
}

Function Hex_Dec_String([string]$hexspace){
	$hex = $hexspace.Replace(' ', '')
	$length = $hex.Length
	if ($length -eq 1){
		return Hex_Dec_Digit $hex
	}
	else {
		$start = $hex.Substring(0, $length - 1)
		$end = $hex.Substring($length - 1)
		return 16 * (Hex_Dec_String $start) + (Hex_Dec_Digit $end)
	}
}

Function Display([string]$head, [string]$hex){
	$dec = Hex_Dec_String $hex
	Write-Host $head ':' $hex ' | ' $dec -foregroundcolor Red
	Add-Content $file (-join($head, ': ', $hex, ' | ', $dec))
	Add-Content $masterfile (-join($head, ': ', $hex, ' | ', $dec))
}

OpenPort
$folder = split-path -parent $MyInvocation.MyCommand.Definition
$masterfile_name = Read-Host -Prompt 'Name master log file name'
$masterfile = New-Item $folder\$masterfile_name.txt -type file -Force
$file_name = Read-Host -Prompt 'Scan to name the new .txt file'
while ($file_name.Length -ne 0){
	$date = Get-Date
	$file = New-Item $folder\$file_name.txt -type file -force
	Add-Content $file (-join('Serial Number: ', $file_name))
	Add-Content $masterfile (-join('Serial Number: ', $file_name))
	Add-Content $file $date
	Add-Content $masterfile $date

	Start-Sleep -s 3
	$port.DiscardInBuffer()
	$port.DiscardOutBuffer()
	Write-Host Calibrating I AC OFFSET Start

	#0: Save SN
	Calibrate (-join('SN=', $file_name))

	#1: SCON
	Calibrate 'SCON'

	#2: 0x000000
	Clear0 'IACOFFA=0x000000'
	Clear0 'IACOFFB=0x000000'
	Clear0 'IACOFFC=0x000000'

	#3: IAOFFCAL
	Calibrate 'IAOFFCAL'
	Start-Sleep -s 10

	#4: IACOFF[A|B|C]
	$IACOFFA = Extract (GetRegisterValue 'IACOFFA')
	$IACOFFB = Extract (GetRegisterValue 'IACOFFB')
	$IACOFFC = Extract (GetRegisterValue 'IACOFFC')
	Display 'IACOFFA' $IACOFFA
	Display 'IACOFFB' $IACOFFB
	Display 'IACOFFC' $IACOFFC

	#5: SINGLE
	Calibrate 'SINGLE'
	Start-Sleep -s 8

	#6: IRMS & VRMS
	$IRMSA = Extract (GetRegisterValue 'IRMSA')
	$IRMSB = Extract (GetRegisterValue 'IRMSB')
	$IRMSC = Extract (GetRegisterValue 'IRMSC')
	$VRMSA = Extract (GetRegisterValue 'VRMSA')
	$VRMSB = Extract (GetRegisterValue 'VRMSB')
	$VRMSC = Extract (GetRegisterValue 'VRMSC')
	Display 'IRMSA' $IRMSA
	Display 'IRMSB' $IRMSB
	Display 'IRMSC' $IRMSC
	Display 'VRMSA' $VRMSA
	Display 'VRMSB' $VRMSB
	Display 'VRMSC' $VRMSC

	Write-Host Calibrating I AC OFFSET Finish
	Write-Host ''

#--------------------------------------------------------
#--------------------------------------------------------
#--------------------------------------------------------

	$continue = Read-Host -Prompt 'Ready to calibrate I GAIN and V GAIN, scan barcode to continue'
	if ($continue -eq $file_name) {
		Write-Host Calibrating I GAIN and V GAIN Start
		#1: SCON  
		Calibrate 'SCON'
        
		#2: 0x000040
		Clear0 'IGAINA=0x000040'
		Clear0 'IGAINB=0x000040'
		Clear0 'IGAINC=0x000040'
		Clear0 'VGAINA=0x000040'
		Clear0 'VGAINB=0x000040'
		Clear0 'VGAINC=0x000040'
            
		#3: ICAL & VCAL
		Calibrate 'VCALA'
		Calibrate 'VCALB'
		Calibrate 'VCALC'
		Start-Sleep -s 10
		Calibrate 'ICALA' 
		Calibrate 'ICALB' 
		Calibrate 'ICALC'
		Start-Sleep -s 10

		#4: IGAIN VGAIN
		$IGAINA = Extract (GetRegisterValue 'IGAINA')
		$IGAINA = Extract (GetRegisterValue 'IGAINA')
		$IGAINB = Extract (GetRegisterValue 'IGAINB')
		$IGAINC = Extract (GetRegisterValue 'IGAINC')
		$VGAINA = Extract (GetRegisterValue 'VGAINA')
		$VGAINB = Extract (GetRegisterValue 'VGAINB')
		$VGAINC = Extract (GetRegisterValue 'VGAINC')
		Display 'IGAINA' $IGAINA
		Display 'IGAINB' $IGAINB
		Display 'IGAINC' $IGAINC
		Display 'VGAINA' $VGAINA
		Display 'VGAINB' $VGAINB
		Display 'VGAINC' $VGAINC

		#5: SINGLE
		Calibrate 'SINGLE'
		Start-Sleep -s 8

		#6: IRMS & VRMs
		$IRMSA = Extract (GetRegisterValue 'IRMSA')
		$IRMSB = Extract (GetRegisterValue 'IRMSB')
		$IRMSC = Extract (GetRegisterValue 'IRMSC')
		$VRMSA = Extract (GetRegisterValue 'VRMSA')
		$VRMSB = Extract (GetRegisterValue 'VRMSB')
		$VRMSC = Extract (GetRegisterValue 'VRMSC')
		Display 'IRMSA' $IRMSA
		Display 'IRMSB' $IRMSB
		Display 'IRMSC' $IRMSC
		Display 'VRMSA' $VRMSA
		Display 'VRMSB' $VRMSB
		Display 'VRMSC' $VRMSC
		
		Write-Host Calibrating I GAIN and V GAIN Finish
		Save0
		Write-Host ''
		Write-Host '-----------------------------------------'
		Write-Host ''

		Add-Content $masterfile '--------------------'
		Add-Content $masterfile '--------------------'
		Add-Content $masterfile '--------------------'

		$file_name = Read-Host -Prompt 'Scan to name the new .txt file'
	}
	else {
		Write-Host Barcode does not match
		$exit = Read-Host -Prompt 'Press any key to exit'
		Break
	}
}

$exit = Read-Host -Prompt 'Press any key to exit'
