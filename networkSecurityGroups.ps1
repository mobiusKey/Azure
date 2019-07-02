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

# Security Rules
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Security Rules"

$column = 1
$row = 1

foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "SecurityRules"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					$excel.cells.item($row,$column) = $_.Name
					$column = $column + 1
				}
				break
				}
				
			}
		}
		$column = 1
	}
	$column = 1
	$row = $row + 1

foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "SecurityRules"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					
					switch($_.Name){
						"Description" {$excel.cells.item($row,$column) = $value.Description; break}
						"Protocol" {$excel.cells.item($row,$column) = $value.Protocol; break}
						"SourcePortRange" {$excel.cells.item($row,$column) = $value.SourcePortRange[0]; break}
						"DestinationPortRange" {$excel.cells.item($row,$column) = $value.DestinationPortRange[0]; break}
						"SourceAddressPrefix" {$excel.cells.item($row,$column) = $value.SourceAddressPrefix[0]; break}
						"DestinationAddressPrefix" {$excel.cells.item($row,$column) = $value.DestinationAddressPrefix[0]; break}
						"Access" {$excel.cells.item($row,$column) = $value.Access; break}
						"Priority" {$excel.cells.item($row,$column) = $value.Priority; break}
						"Direction" {$excel.cells.item($row,$column) = $value.Direction; break}
						"ProvisioningState" {$excel.cells.item($row,$column) = $value.ProvisioningState; break}
						"SourceApplicationSecurityGroups" {$excel.cells.item($row,$column) = $value.SourceApplicationSecurityGroups[0]; break}
						"SourceApplicationSecurityGroupsText" {$excel.cells.item($row,$column) = $value.SourceApplicationSecurityGroupsText; break}
						"DestinationApplicationSecurityGroups" {$excel.cells.item($row,$column) = $value.DestinationApplicationSecurityGroups[0]; break}
						"DestinationApplicationSecurityGroupsText" {$excel.cells.item($row,$column) = $value.DestinationApplicationSecurityGroupsText; break}
						"Name" {$excel.cells.item($row,$column) = $value.Name; break}
						"Etag" {$excel.cells.item($row,$column) = $value.Etag; break}
						"Id" {$excel.cells.item($row,$column) = $value.Id; break}
						default {$excel.cells.item($row,$column) = "Error"}
					}
					
					$column = $column + 1
				}
				$row = $row + 1
				$column = 1
				}
				
			}
		}
		$column = 1
	}
	
	
# Default Security Rules
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Default Security Rules"

$column = 1
$row = 1
foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "DefaultSecurityRules"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					$excel.cells.item($row,$column) = $_.Name
					$column = $column + 1
				}
				break
				}
				break
				
			}
		}
		$column = $column + 1
	}
	$column = 1
	$row = $row + 1

foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "DefaultSecurityRules"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					
					switch($_.Name){
						"Description" {$excel.cells.item($row,$column) = $value.Description; break}
						"Protocol" {$excel.cells.item($row,$column) = $value.Protocol; break}
						"SourcePortRange" {$excel.cells.item($row,$column) = $value.SourcePortRange[0]; break}
						"DestinationPortRange" {$excel.cells.item($row,$column) = $value.DestinationPortRange[0]; break}
						"SourceAddressPrefix" {$excel.cells.item($row,$column) = $value.SourceAddressPrefix[0]; break}
						"DestinationAddressPrefix" {$excel.cells.item($row,$column) = $value.DestinationAddressPrefix[0]; break}
						"Access" {$excel.cells.item($row,$column) = $value.Access; break}
						"Priority" {$excel.cells.item($row,$column) = $value.Priority; break}
						"Direction" {$excel.cells.item($row,$column) = $value.Direction; break}
						"ProvisioningState" {$excel.cells.item($row,$column) = $value.ProvisioningState; break}
						"SourceApplicationSecurityGroups" {$excel.cells.item($row,$column) = $value.SourceApplicationSecurityGroups[0]; break}
						"SourceApplicationSecurityGroupsText" {$excel.cells.item($row,$column) = $value.SourceApplicationSecurityGroupsText; break}
						"DestinationApplicationSecurityGroups" {$excel.cells.item($row,$column) = $value.DestinationApplicationSecurityGroups[0]; break}
						"DestinationApplicationSecurityGroupsText" {$excel.cells.item($row,$column) = $value.DestinationApplicationSecurityGroupsText; break}
						"Name" {$excel.cells.item($row,$column) = $value.Name; break}
						"Etag" {$excel.cells.item($row,$column) = $value.Etag; break}
						"Id" {$excel.cells.item($row,$column) = $value.Id; break}
						default {$excel.cells.item($row,$column) = "Error"}
					}
					
					$column = $column + 1
				}
				$row = $row + 1
				$column = 1
				}
				
			}
		}
		$column = 1
	}
	
# Network Interfaces
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Network Interfaces"

$column = 1
$row = 1
#Subnets not showing up for some reason
foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "NetworkInterfaces"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					$_.Name
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
			if($_.Name -eq "NetworkInterfaces"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					
					switch($_.Name){
						"Description" {$excel.cells.item($row,$column) = $value.Description; break}
						"Protocol" {$excel.cells.item($row,$column) = $value.Protocol; break}
						"SourcePortRange" {$excel.cells.item($row,$column) = $value.SourcePortRange[0]; break}
						"DestinationPortRange" {$excel.cells.item($row,$column) = $value.DestinationPortRange[0]; break}
						"SourceAddressPrefix" {$excel.cells.item($row,$column) = $value.SourceAddressPrefix[0]; break}
						"DestinationAddressPrefix" {$excel.cells.item($row,$column) = $value.DestinationAddressPrefix[0]; break}
						"Access" {$excel.cells.item($row,$column) = $value.Access; break}
						"Priority" {$excel.cells.item($row,$column) = $value.Priority; break}
						"Direction" {$excel.cells.item($row,$column) = $value.Direction; break}
						"ProvisioningState" {$excel.cells.item($row,$column) = $value.ProvisioningState; break}
						"SourceApplicationSecurityGroups" {$excel.cells.item($row,$column) = $value.SourceApplicationSecurityGroups[0]; break}
						"SourceApplicationSecurityGroupsText" {$excel.cells.item($row,$column) = $value.SourceApplicationSecurityGroupsText; break}
						"DestinationApplicationSecurityGroups" {$excel.cells.item($row,$column) = $value.DestinationApplicationSecurityGroups[0]; break}
						"DestinationApplicationSecurityGroupsText" {$excel.cells.item($row,$column) = $value.DestinationApplicationSecurityGroupsText; break}
						"Name" {$excel.cells.item($row,$column) = $value.Name; break}
						"Etag" {$excel.cells.item($row,$column) = $value.Etag; break}
						"Id" {$excel.cells.item($row,$column) = $value.Id; break}
						default {$excel.cells.item($row,$column) = "Error"}
					}
					
					$column = $column + 1
				}
				$row = $row + 1
				$column = 1
				}
				
			}
		}
		$column = 1
	}
	