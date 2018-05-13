$notepad = "HKCU:\Software\Microsoft\Notepad"
Set-ItemProperty -Path $notepad -Name lfFaceName -Value "Courier New"
Set-ItemProperty -Path $notepad -Name iPointSize -Value 120

$folder = split-path -parent $MyInvocation.MyCommand.Definition
$file_name = Read-Host -Prompt 'Name the new .txt file'
$file = New-Item $folder\$file_name.txt -type file -force

$barcode = Read-Host 'Scan Barcode:'
$count = 0
while ($barcode -ne ''){
	$count = $count + 1
	Write-Host $count ': ' $barcode -BackgroundColor Red
	Add-Content $file (-join($count, ': ', $barcode))
	$barcode = Read-Host 'Scan Barcode:'
}
Write-Host 'total count of' $count 'meters' -BackgroundColor Red
Add-Content $file (-join('total count of ', $count, ' meters'))

$exit = Read-Host -Prompt 'Press any key to exit'