# Before running the script:
# * Run: Import-Module Azure
# * Run Install-Module -Name AzureRM -AllowClobber
# * Authenticate to Azure in PowerShell using Login-AzureRmAccount


Write-Output("==== Users ====")
$users = Get-AzureRmADUser -First 10

$excel = New-Object -ComObject excel.application
$excel.visible = $True
$workbook = $excel.WorkBooks.Add()
$workbook.WorkSheets.Item(1).Name = "Users"
$worksheet = $workbook.worksheets.Item(1)
$row = 1
$column = 1

foreach($user in $users){
	$user.PSObject.Properties | ForEach-Object {
		$excel.cells.item($row,$column) = $_.Name
		$column = $column + 1
	}

	break
}

$column = 1
$row = $row + 1
$users
foreach($user in $users){
	
	$user.PSObject.Properties | ForEach-Object {
		if ($_.Value -ne $null -and $_.Value.count -ne 0 -and $_.Name -ne "Id"){
		$excel.cells.item($row,$column) = $_.Value
		}else{
		$excel.cells.item($row,$column) = ""
		}
		$column = $column + 1
	}
	$column = 1
	$row = $row + 1
}

