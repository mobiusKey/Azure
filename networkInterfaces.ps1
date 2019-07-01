# Before running the script:
# * Run: Import-Module Azure
# * Run Install-Module -Name AzureRM -AllowClobber
# * Authenticate to Azure in PowerShell using Login-AzureRmAccount


Write-Output("==== Network Security Groups ====")
$securityGroups = Get-AzureRmNetworkSecurityGroup
$securityGroups

$excel = New-Object -ComObject excel.application
$excel.visible = $True
$workbook = $excel.WorkBooks.Add()
$workbook.WorkSheets.Item(1).Name = "Network Security Groups"
$worksheet = $workbook.worksheets.Item(1)
$row = 1
$column = 1

foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
		$excel.cells.item($row,$column) = $_.Name
		$column = $column + 1
	}

	break
}

$column = 1
$row = $row + 1



foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
		if ($_.Value -ne $null -and $_.Value.count -ne 0 -and (-Not($_.Name -eq "SecurityRules")) -and (-Not($_.Name -eq "DefaultSecurityRules")) -and (-Not($_.Name -eq "NetworkInterfaces")) -and (-Not($_.Name -eq "Subnets"))){
		$excel.cells.item($row,$column) = $_.Value
		}
		else{
			if((($_.Name -eq "SecurityRules")) -or (($_.Name -eq "DefaultSecurityRules")) -or (($_.Name -eq "NetworkInterfaces")) -or (($_.Name -eq "Subnets"))){
				Write-Output("<------------->")
				Write-Output $_.Name
				Write-Output $_.Value
			}else{
		
		$excel.cells.item($row,$column) = ""
		}
		}
		$column = $column + 1
	}
	$column = 1
	$row = $row + 1
}

$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Security Rules"

$column = 1
$row = 1
Write-Output("-------Security Rules-------")
foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "SecurityRules"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					Write-Output $_.Name
					$excel.cells.item($row,$column) = $_.Name
					$column = $column + 1
				}
				break
				}
				
			}
		}
		$column = $column + 1
	}
	$column = 1
	$row = $row + 1

foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "SecurityRules"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					switch($_.Name){
						"SourcePortRange" {$excel.cells.item($row,$column) =$value.SourcePortRange[0]; break}
						default {}
					}
					Write-Output $_.Name
					$excel.cells.item($row,$column) = $_.Value
					$column = $column + 1
				}
				$row = $row + 1
				$column = 1
				}
				
			}
		}
		$column = $column + 1
	}