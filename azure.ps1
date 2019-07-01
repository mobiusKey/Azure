# Before running the script:
# * Run: Import-Module Azure
# * Authenticate to Azure in PowerShell using Login-AzureRmAccount


$CurrentSubscription = Get-AzureRmContext
$banner = "========= Subscription: " + $CurrentSubscription.Subscription.Id + " =========" | Out-String
Write-Output $banner
$CurrentSubscription

Write-Output("----- Account Info -----")
$CurrentSubscription.Account

Write-Output("----- Tenant Info -----")
$CurrentSubscription.Tenant

Write-Output("---- Role Assingnments ----")
Get-AzureRmRoleAssignment

Write-Output("----- Subscription Info -----")
$CurrentSubscription.Subscription
Write-Output("---- Resource Groups ----")
# omits tags, are they important?
$resourceGroup = Get-AzureRmResourceGroup
Get-AzureRmResourceGroup | Format-Table ResourceGroupName, Location, ProvisioningState

Write-Output("==== Web Apps ====")
Get-AzureRmWebApp

Write-Output("==== Virtual Machines ====")
$vms = Get-AzureRmVM
foreach($vm in $vms){
	Write-Output("-- Resource Group Name --")
	Get-AzureRmVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name
	Write-Output("-- Hardware Profile --")
	$vm.HardwareProfile
	Write-Output("-- OS Profile --")
	$vm.OSProfile
	Write-Output("-- Image Reference --")
	$vm.StorageProfile.ImageReference
}

Write-Output("==== Storage ====")
$SAs = Get-AzureRmStorageAccount
$SAs
# not sure if this should be used in a config review
<# foreach($sa in $SAs){
	Get-AzureRmStorageAccountKey -ResourceGroupName $sa.ResourceGroupName -StorageAccountName $sa.StorageAccountName
} #>

Write-Output("==== Networking ====")
$networkInterfaces = Get-AzureRmNetworkInterface
$networkInterfaces
Write-Output("-- Public IP Space --")
Get-AzureRmPublicIPAddress

Write-Output("==== Network Security Groups ====")
# TODO: Gives no output needs fixing
# perhaps because I have no networkInterfaces
foreach($vm in $vms){
	$vm.Name
	$ni = Get-AzureRmNetworkInterface | where {$_.Id -eq $vm.NetworkInterfaceIDs}
	Write-Output("Network Security Group for " + $vm.Name + ":")
	Get-AzureRmNetworkSecurityGroup | where {$_.Id -eq $ni.NetworkSecurityGroup.Id}
	Get-AzureRmNetworkSecurityGroup
}

Write-Output("==== SQL ====")
foreach( $rg in $resourceGroup){
	foreach($ss in Get-AzureRmSqlServer -ResourceGroupName $rg.ResourceGroupName){
		Write-Output("Server: " + $ss.ServerName + " Resource Group Name: " + $rg.ResourceGroupName)
		Get-AzureRmSqlServer -ServerName $ss.ServerName -ResourceGroupName $rg.ResourceGroupName
		
		Write-Output("-- Databases --")
		Get-AzureRmSqlDatabase -ServerName $ss.ServerName -ResourceGroupName $rg.ResourceGroupName
		
		Write-Output("-- SQL Firewall Rules --")
		Get-AzureRmSqlServerFirewallRule -ServerName $ss.ServerName -ResourceGroupName $rg.ResourceGroupName
		
		Write-Output("-- Threat Detection Policy --")
		Get-AzureRmSqlServerThreatDetectionPolicy -ServerName $ss.ServerName -ResourceGroupName $rg.ResourceGroupName
		
		
		}
	}

Write-Output("==== Users ====")
# Get-AzureRmADUser

$excel = New-Object -ComObject excel.application
$excel.visible = $True
$workbook = $excel.WorkBooks.Add()
$workbook.WorkSheets.Item(1).Name = "Resource Groups"
$row = 1
$column = 1
# Bad WaySet Column Headers
foreach($rg in $resourceGroup){
	$rg.PSObject.Properties | ForEach-Object {
		$excel.cells.item($row,$column) = $_.Name
		$column = $column + 1
	}
	break
}
$column = 1
$row = $row + 1
$list = {}
# TODO: Fix breaks when tag is an empty list {} but output still works
foreach($rg in $resourceGroup){
	$rg.PSObject.Properties | ForEach-Object {
		if ($_.Value -ne $null -or $_.Value -ne $list){
		$excel.cells.item($row,$column) = $_.Value
		}else{
		$excel.cells.item($row,$column) = ""
		}
		$column = $column + 1
	}
	$column = 1
	$row = $row + 1
	$rg
}