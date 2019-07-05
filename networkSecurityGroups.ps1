# Before running the script:
# * Run: Import-Module Azure
# * Run Install-Module -Name AzureRM -AllowClobber
# * Authenticate to Azure in PowerShell using Login-AzureRmAccount

# TODO: Fix Tags output

function arrayToString {
	if($args[0] -ne $null){ 
	$string = ""
	$args[0]
	foreach($element in $args[0]){
		$string += $element + " "
	}
	return $string
	}
	return "NULL"
}


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
		$column += 1
	}

	break
}

$column = 1
$worksheet.Rows($row).RowHeight = 15
$row += 1



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
		$column += 1
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1
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
					$column += 1
				}
				break
				}
				
			}
		}
		$column = 1
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1

foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "SecurityRules"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					
					switch($_.Name){
						"Description" {$excel.cells.item($row,$column) = $value.Description; break}
						"Protocol" {$excel.cells.item($row,$column) = $value.Protocol; break}
						"SourcePortRange" {$excel.cells.item($row,$column) = arrayToString($value.SourcePortRange); break}
						"DestinationPortRange" {$excel.cells.item($row,$column) = arrayToString($value.DestinationPortRange); break}
						"SourceAddressPrefix" {$excel.cells.item($row,$column) = arrayToString($value.SourceAddressPrefix); break}
						"DestinationAddressPrefix" {$excel.cells.item($row,$column) = arrayToString($value.DestinationAddressPrefix); break}
						"Access" {$excel.cells.item($row,$column) = $value.Access; break}
						"Priority" {$excel.cells.item($row,$column) = $value.Priority; break}
						"Direction" {$excel.cells.item($row,$column) = $value.Direction; break}
						"ProvisioningState" {$excel.cells.item($row,$column) = $value.ProvisioningState; break}
						"SourceApplicationSecurityGroups" {$excel.cells.item($row,$column) = arrayToString($value.SourceApplicationSecurityGroups); break}
						"SourceApplicationSecurityGroupsText" {$excel.cells.item($row,$column) = $value.SourceApplicationSecurityGroupsText; break}
						"DestinationApplicationSecurityGroups" {$excel.cells.item($row,$column) = arrayToString($value.DestinationApplicationSecurityGroups); break}
						"DestinationApplicationSecurityGroupsText" {$excel.cells.item($row,$column) = $value.DestinationApplicationSecurityGroupsText; break}
						"Name" {$excel.cells.item($row,$column) = $value.Name; break}
						"Etag" {$excel.cells.item($row,$column) = $value.Etag; break}
						"Id" {$excel.cells.item($row,$column) = $value.Id; break}
						default {$excel.cells.item($row,$column) = "Error"}
					}
					
					$column += 1
				}
				$worksheet.Rows($row).RowHeight = 15
				$row += 1
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
					$column += 1
				}
				break
				}
				break
				
			}
		}
		$column += 1
	}
	$column = 1
	$row += 1

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
					
					$column += 1
				}
				$row += 1
				$column = 1
				}
				
			}
		}
		$column = 1
	}
	
# Network Interfaces
# TODO: Understand why in Get-AzureRmNetworkSecurityGroup the network interfaces is blank
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Network Interfaces"

$column = 1
$row = 1


Write-Output("====== Interfaces ======")
$interfaces = Get-AzureRmNetworkInterface
$interfaces


foreach($interface in $interfaces){
	$interface.PSObject.Properties | ForEach-Object {
		$excel.cells.item($row,$column) = $_.Name
		$column += 1
			
	}
	break
}

	$column = 1
	$row += 1

foreach($interface in $interfaces){
	$interface.PSObject.Properties | ForEach-Object {
		$value = $_.Value
		switch($_.Name){
			"VirtualMachine" {$excel.cells.item($row,$column) = $value.Id; break}
			# contains if IPv4, Name, Primary, private addresss,  PrivateIpAllocationMethod 
			"IpConfigurations" {$excel.cells.item($row,$column) = $value.Name; break}
			"TapConfigurations" {$excel.cells.item($row,$column) = arrayToString($value); break}
			# contains DnsServers, AppliedDnsServers, Internal, DnsNameLabel, InternalFqdn, InternalDomainNameSuffix, DnsServersText, AppliedDnsServersText
			"DnsSettings" {$excel.cells.item($row,$column) = $value.DnsNameLabel; break}
			"MacAddress" {$excel.cells.item($row,$column) = $value; break}
			"Primary" {$excel.cells.item($row,$column) = $value; break}
			"EnableAcceleratedNetworking" {$excel.cells.item($row,$column) = $value; break}
			"EnableIPForwarding" {$excel.cells.item($row,$column) = $value; break}
			"HostedWorkloads" {$excel.cells.item($row,$column) = arrayToString($value); break}
			# contains ResourceGroupName Name Location ProvisioningState
			#"NetworkSecurityGroup" {$sg = Get-AzureRmEffectiveNetworkSecurityGroup	-ResourceGroupName $interface.ResourceGroupName  -NetworkInterfaceName $interface.Name; $excel.cells.item($row,$column) = $sg.NetworkSecurityGroup.Id; break}
			"NetworkSecurityGroup" {$excel.cells.item($row,$column) = $value.Id; break}
			"ProvisioningState" {$excel.cells.item($row,$column) = $value; break}
			"VirtualMachineText" {$excel.cells.item($row,$column) = $value; break}
			"IpConfigurationsText" {$excel.cells.item($row,$column) = $value; break}
			"TapConfigurationsText" {$excel.cells.item($row,$column) = $value; break}
			"DnsSettingsText" {$excel.cells.item($row,$column) = $value; break}
			"NetworkSecurityGroupText" {$excel.cells.item($row,$column) = $value; break}
			"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}
			"Location" {$excel.cells.item($row,$column) = $value; break}
			"ResourceGuid" {$excel.cells.item($row,$column) = $value; break}
			"Type" {$excel.cells.item($row,$column) = $value; break}
			"Tag" {$excel.cells.item($row,$column) = $value.Name; break}
			"TagsTable" {$excel.cells.item($row,$column) = $value; break}
			"Name" {$excel.cells.item($row,$column) = $value; break}
			"Etag" {$excel.cells.item($row,$column) = $value; break}
			"Id" {$excel.cells.item($row,$column) = $value; break}
			default {$excel.cells.item($row,$column) = "Error"}
		}
		
		$column += 1
				}
		$worksheet.Rows($row).RowHeight = 15
		$row += 1
		$column = 1
	}
	
# Subnets
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Subnets in Security Groups"

$column = 1
$row = 1

foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "Subnets"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					$excel.cells.item($row,$column) = $_.Name
					$column += 1
				}
				break
				}
				
			}
		}
		$column = 1
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1

foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "Subnets"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					
					switch($_.Name){
						"AddressPrefix" {$excel.cells.item($row,$column) = $value.AddressPrefix; break}
						"IpConfigurations" {$excel.cells.item($row,$column) = $value.IpConfigurations.Name; break}
						"ServiceAssociationLinks" {$excel.cells.item($row,$column) = arrayToString($value.ServiceAssociationLinks); break}
						"ResourceNavigationLinks" {$excel.cells.item($row,$column) = arrayToString($value.ResourceNavigationLinks); break}
						# TODO: Can a subnet have multiple network security groups?
						"NetworkSecurityGroup" {$excel.cells.item($row,$column) = $group.Id; break}
						"RouteTable" {$excel.cells.item($row,$column) = $value.RouteTable; break}
						"ServiceEndpoints" {$excel.cells.item($row,$column) = arrayToString($value.ServiceEndpoints); break}
						"ServiceEndpointPolicies" {$excel.cells.item($row,$column) = arrayToString($value.ServiceEndpointPolicies); break}
						"Delegations" {$excel.cells.item($row,$column) = arrayToString($value.Delegations); break}
						"InterfaceEndpoints" {$excel.cells.item($row,$column) = arrayToString($value.InterfaceEndpoints); break}
						"ProvisioningState" {$excel.cells.item($row,$column) = $value.ProvisioningState; break}
						"IpConfigurationsText" {$excel.cells.item($row,$column) = $value.IpConfigurationsText; break}
						"ServiceAssociationLinksText" {$excel.cells.item($row,$column) = $value.ServiceAssociationLinksText; break}
						"ResourceNavigationLinksText" {$excel.cells.item($row,$column) = $value.ResourceNavigationLinksText; break}
						"NetworkSecurityGroupText" {$excel.cells.item($row,$column) = $value.NetworkSecurityGroupText; break}
						"RouteTableText" {$excel.cells.item($row,$column) = $value.RouteTableText; break}
						"ServiceEndpointText" {$excel.cells.item($row,$column) = $value.ServiceEndpointText; break}
						"InterfaceEndpointsText" {$excel.cells.item($row,$column) = $value.InterfaceEndpointsText; break}
						"DelegationsText" {$excel.cells.item($row,$column) = $value.DelegationsText; break}
						"Name" {$excel.cells.item($row,$column) = $value.Name; break}
						"Etag" {$excel.cells.item($row,$column) = $value.Etag; break}
						"Id" {$excel.cells.item($row,$column) = $value.Id; break}
						default {$excel.cells.item($row,$column) = "Error"}
					}
					
					$column += 1
				}
				$worksheet.Rows($row).RowHeight = 15
				$row += 1
				$column = 1
				}
				
			}
		}
		$column = 1
	}
	
	
# Virtual Networks
Write-Output("========= Virtual Networks ==========")

$networks = Get-AzureRmVirtualNetwork
$network
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Virtual Networks"
$subnets = New-Object System.Collections.ArrayList
$column = 1
$row = 1


foreach($network in $networks){
	$network.PSObject.Properties | ForEach-Object {
					$excel.cells.item($row,$column) = $_.Name
					Write-Output($_.Name + ": " + $_.Value)
					$column += 1
		}
		
		$column = 1
		break
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1


foreach($network in $networks){
	$network.PSObject.Properties | ForEach-Object {
		$value = $_.Value
		Write-Output($_.Name + ": " +$value)
		switch($_.Name){
		# TODO: Fix Address Space output
		"AddressSpace" {$excel.cells.item($row,$column) = $network.AddressSpace.AddressPrefixesText; break}
		"DhcpOptions" {$excel.cells.item($row,$column) = $value; break}
		"Subnets" {
			$substring = ""
			foreach($sub in $value){
			$substring += $sub.Id + " "
			$excel.cells.item($row,$column) = $substring
		}
		break}
		"VirtualNetworkPeerings" {$excel.cells.item($row,$column) = $value.Name; break}
		"ProvisioningState" {$excel.cells.item($row,$column) = $value; break}
		"EnableDdosProtection" {$excel.cells.item($row,$column) = $value; break}
		"EnableVmProtection" {$excel.cells.item($row,$column) = $value; break}
		"DdosProtectionPlan" {$excel.cells.item($row,$column) = $value; break}
		"AddressSpaceText" {$excel.cells.item($row,$column) = $value; break}
		"DhcpOptionsText" {$excel.cells.item($row,$column) = $value; break}
		"SubnetsText" {$excel.cells.item($row,$column) = $value; break}
		"VirtualNetworkPeeringsText" {$excel.cells.item($row,$column) = $value; break}	
		"EnableDdosProtectionText" {$excel.cells.item($row,$column) = $value; break}	
		"DdosProtectionPlanText" {$excel.cells.item($row,$column) = $value; break}	
		"EnableVmProtectionText" {$excel.cells.item($row,$column) = $value; break}	
		"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}	
		"Location" {$excel.cells.item($row,$column) = $value; break}	
		"ResourceGuid" {$excel.cells.item($row,$column) = $value; break}	
		"Type" {$excel.cells.item($row,$column) = $value; break}	
		"Tag" {$excel.cells.item($row,$column) = $value.Name; break}
		"TagsTable" {$excel.cells.item($row,$column) = $value; break}	
		"Name" {$excel.cells.item($row,$column) = $value; break}	
		"Etag" {$excel.cells.item($row,$column) = $value; break}	
		"Id" {$excel.cells.item($row,$column) = $value; break}
		
		default {$excel.cells.item($row,$column) = "Error"}
	}
	$column += 1
		}
					$column = 1
					$worksheet.Rows($row).RowHeight = 15
					$row += 1
					foreach($net in $network.Subnets){
						$subnets.add($net)
					}
	}
	
Write-Output("==== All Subnets =====")
$subnets
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Subnets"
$column = 1
$row = 1
	
	# TODO: programatically add column names and values instead of hardcoded column names etc
$excel.cells.item($row,$column) = "AddressPrefix"
$excel.cells.item($row,$column + 1) = "IpConfigurations"	
$excel.cells.item($row,$column + 2) = "ServiceAssociationLinks"	
$excel.cells.item($row,$column + 3) = "ResourceNavigationLinks"	
$excel.cells.item($row,$column + 4) = "NetworkSecurityGroup"	
$excel.cells.item($row,$column + 5) = "RouteTable"	
$excel.cells.item($row,$column + 6) = "ServiceEndpoints"	
$excel.cells.item($row,$column + 7) = "ServiceEndpointPolicies"	
$excel.cells.item($row,$column + 8) = "Delegations"	
$excel.cells.item($row,$column + 9) = "InterfaceEndpoints"	
$excel.cells.item($row,$column + 10) = "ProvisioningState"	
$excel.cells.item($row,$column + 11) = "IpConfigurationsText"	
$excel.cells.item($row,$column + 12) = "ServiceAssociationLinksText"	
$excel.cells.item($row,$column + 13) = "ResourceNavigationLinksText"	
$excel.cells.item($row,$column + 14) = "NetworkSecurityGroupText"	
$excel.cells.item($row,$column + 15) = "RouteTableText"	
$excel.cells.item($row,$column + 16) = "ServiceEndpointText"	
$excel.cells.item($row,$column + 17) = "ServiceEndpointPoliciesText"	
$excel.cells.item($row,$column + 18) = "InterfaceEndpointsText"	
$excel.cells.item($row,$column + 19) = "DelegationsText"	
$excel.cells.item($row,$column + 20) = "Name"	
$excel.cells.item($row,$column + 21) = "Etag"	
$excel.cells.item($row,$column + 22) = "Id"

$column = 1
$row += 1
foreach($subnet in $subnets){
$excel.cells.item($row,$column) = arrayToString($subnet.AddressPrefix)
$excel.cells.item($row,$column + 1) = $subnet.IpConfigurations.Name
$excel.cells.item($row,$column + 2) = arrayToString($subnet.ServiceAssociationLinks)
$excel.cells.item($row,$column + 3) = arrayToString($subnet.ResourceNavigationLinks)
$excel.cells.item($row,$column + 4) = $subnet.NetworkSecurityGroup.Id
$excel.cells.item($row,$column + 5) = $subnet.RouteTable.Id
$excel.cells.item($row,$column + 6) = arrayToString($subnet.ServiceEndpoints)
$excel.cells.item($row,$column + 7) = arrayToString($subnet.ServiceEndpointPolicies)	
$excel.cells.item($row,$column + 8) = arrayToString($subnet.Delegations)
$excel.cells.item($row,$column + 9) = arrayToString($subnet.InterfaceEndpoints)
$excel.cells.item($row,$column + 10) = $subnet.ProvisioningState	
$excel.cells.item($row,$column + 11) = $subnet.IpConfigurationsText	
$excel.cells.item($row,$column + 12) = $subnet.ServiceAssociationLinksText	
$excel.cells.item($row,$column + 13) = $subnet.ResourceNavigationLinksText	
$excel.cells.item($row,$column + 14) = $subnet.NetworkSecurityGroupText
$excel.cells.item($row,$column + 15) = $subnet.RouteTableText
$excel.cells.item($row,$column + 16) = $subnet.ServiceEndpointText
$excel.cells.item($row,$column + 17) = $subnet.ServiceEndpointPoliciesText	
$excel.cells.item($row,$column + 18) = $subnet.InterfaceEndpointsText
$excel.cells.item($row,$column + 19) = $subnet.DelegationsText
$excel.cells.item($row,$column + 20) = $subnet.Name
$excel.cells.item($row,$column + 21) = $subnet.Etag	
$excel.cells.item($row,$column + 22) = $subnet.Id

$worksheet.Rows($row).RowHeight = 15
$row += 1
			
}
	
# Virtual Machines
# TODO: Check Extensions for Endpoint protections
Write-Output("========= Virtual Machines ==========")

$vms = Get-AzureRmVM
$vms
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Virtual Machines"
$column = 1
$row = 1


foreach($vm in $vms){
	$vm.PSObject.Properties | ForEach-Object {
		if($_.Name -eq "OSProfile" -or $_.Name -eq "StorageProfile"){
			if($_.Name -eq "OSProfile"){
				$excel.cells.item($row,$column) = $_.Name
				#Using this value for the OSProfile field
				# $excel.cells.item($row,$column + 1) = "ComputerName"
				$excel.cells.item($row,$column + 1) = "AdminUsername"
				$excel.cells.item($row,$column + 2) = "AdminPassword"
				$excel.cells.item($row,$column + 3) = "CustomData"
				$excel.cells.item($row,$column + 4) = "WindowsConfiguration"
				$excel.cells.item($row,$column + 5) = "Secrets"
					$column = $column + 6
			}
			if($_.Name -eq "StorageProfile"){
				$excel.cells.item($row,$column) = $_.Name
				$excel.cells.item($row,$column + 1) = "Offer"
				$excel.cells.item($row,$column + 2) = "Sku"
				$excel.cells.item($row,$column + 3) = "Version"
					$column = $column + 4
			}
		}
		else{
		$excel.cells.item($row,$column) = $_.Name
		$column += 1
		}	
		Write-Output($_.Name + ": " + $_.Value)
		}
		
		$column = 1
		break
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1

foreach($vm in $vms){
	$vm.PSObject.Properties | ForEach-Object {
	$value = $_.Value
	Write-Output($_.Name + ": " +$value)
	switch($_.Name){
	
	"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}
	"Id" {$excel.cells.item($row,$column) = $value; break}
	"VmId" {$excel.cells.item($row,$column) = $value; break}	
	"Name" {$excel.cells.item($row,$column) = $value; break}	
	"Type" {$excel.cells.item($row,$column) = $value; break}	
	"Location" {$excel.cells.item($row,$column) = $value; break}	
	"LicenseType" {$excel.cells.item($row,$column) = $value; break}	
	"Tags" {$excel.cells.item($row,$column) = $value[0]; break}	
	"AvailabilitySetReference" {$excel.cells.item($row,$column) = $value; break}	
	"DiagnosticsProfile" {$excel.cells.item($row,$column) = $value; break}	
	"Extensions" {$excel.cells.item($row,$column) = $value[0]; break}	
	"HardwareProfile" {$excel.cells.item($row,$column) = $value.VmSize; break}	
	"InstanceView"	{$excel.cells.item($row,$column) = $value; break}
	"NetworkProfile" {
		$string = ""
		foreach($interface in $value.NetworkInterfaces){
			$string += $interface.Id
		
		}
		$excel.cells.item($row,$column) = $string
		break
	}	
	# TODO: Check if Password Authentication is enabled?
	# TODO: Check Windows and Linux Configurations? here or in another excel workbook?
	"OSProfile" {
		$excel.cells.item($row,$column) = $value.ComputerName
		$excel.cells.item($row,$column + 1) = $value.AdminUsername
		$excel.cells.item($row,$column + 2) = $value.AdminPassword
		$excel.cells.item($row,$column + 3) = $value.CustomData
		$excel.cells.item($row,$column + 4) = $value.WindowsConfiguration
		$excel.cells.item($row,$column + 5) = arrayToString($value.Secrets)
		$column = $column + 5
		break
	}	
	"Plan" {$excel.cells.item($row,$column) = $value; break}	
	"ProvisioningState" {$excel.cells.item($row,$column) = $value; break}	
	"StorageProfile" {
		$excel.cells.item($row,$column) = $value.ImageReference.Publisher
		$excel.cells.item($row,$column + 1) = $value.ImageReference.Offer
		$excel.cells.item($row,$column + 2) = $value.ImageReference.Sku
		$excel.cells.item($row,$column + 3) = $value.ImageReference.Version
		#skipping Id to avoid confusion with VM Id
		$column = $column + 3
		break
	}	
	"DisplayHint" {$excel.cells.item($row,$column) = $value; break}	
	"Identity" {$excel.cells.item($row,$column) = $value; break}	
	"Zones" {$excel.cells.item($row,$column) = $value[0]; break}	
	"FullyQualifiedDomainName" {$excel.cells.item($row,$column) = $value; break}
	"AdditionalCapabilities" {$excel.cells.item($row,$column) = $value; break}	
	"RequestId" {$excel.cells.item($row,$column) = $value; break}	
	"StatusCode" {$excel.cells.item($row,$column) = $value; break}
						
	default {$excel.cells.item($row,$column) = "Error"}
					}
					$column += 1
		}
					$column = 1
					$worksheet.Rows($row).RowHeight = 15
					$row += 1
	}
	
	
	# Application Gateways and Firewalls (Web App and Next Gen Firewalls)
	Write-Output("========= Application Gateways ==========")

$gateways = Get-AzureRmApplicationGateway
$gateways
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Application Gateways"

$column = 1
$row = 1


foreach($gate in $gateways){
	$gate.PSObject.Properties | ForEach-Object {
					$excel.cells.item($row,$column) = $_.Name
					Write-Output($_.Name + ": " + $_.Value)
					$column += 1
		}
		
		$column = 1
		break
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1

# TODO: Are WAFs that the user creates 
	foreach($gate in $gateways){
	$gate.PSObject.Properties | ForEach-Object {
		$value = $_.Value
		Write-Output($_.Name + ": " +$value)
		switch($_.Name){
			
			
			"Sku" {$excel.cells.item($row,$column) = $value.Name; break} # Object
			"SslPolicy"	{$excel.cells.item($row,$column) = $value.Id; break} # Object
			"GatewayIPConfigurations" {$excel.cells.item($row,$column) = $value[0].Name; break}	# List of Objects
			"AuthenticationCertificates" {$excel.cells.item($row,$column) = arrayToString($value); break}	
			"SslCertificates" {$excel.cells.item($row,$column) = arrayToString($value); break}	
			"TrustedRootCertificates" {$excel.cells.item($row,$column) = arrayToString($value); break}	# List of Objects
			"FrontendIPConfigurations" {$excel.cells.item($row,$column) = $value[0].Name; break}	#List of Objects
			"FrontendPorts" {$excel.cells.item($row,$column) = $value[0].Port; break}	# List of Objects
			"Probes" {$excel.cells.item($row,$column) = arrayToString($value); break}	# List of Objects
			"BackendAddressPools" {$excel.cells.item($row,$column) = $value[0].Name; break}	# List of Objects
			"BackendHttpSettingsCollection"	{$excel.cells.item($row,$column) = $value[0].Name; break} # List of Objects
			"HttpListeners"	{$excel.cells.item($row,$column) = $value[0].Name; break} # List of Objects
			"UrlPathMaps" {$excel.cells.item($row,$column) = arrayToString($value); break}	# List of Objects
			"RequestRoutingRules" {$excel.cells.item($row,$column) = $value.Id; break}	# Object 
			"RedirectConfigurations" {$excel.cells.item($row,$column) = $value.Id; break}	
			"WebApplicationFirewallConfiguration" {$excel.cells.item($row,$column) = $value.Id; break} # Object
			"AutoscaleConfiguration" {$excel.cells.item($row,$column) = $value.Id; break}	# Object
			"CustomErrorConfigurations"	{$excel.cells.item($row,$column) = arrayToString($value); break}
			"EnableHttp2" {$excel.cells.item($row,$column) = $value; break}	
			"EnableFips" {$excel.cells.item($row,$column) = $value; break}	
			"Zones" {$excel.cells.item($row,$column) = arrayToString($value); break}	
			"OperationalState" {$excel.cells.item($row,$column) = $value; break}	
			"ProvisioningState" {$excel.cells.item($row,$column) = $value; break}	
			"GatewayIpConfigurationsText" {$excel.cells.item($row,$column) = $value; break}	
			"AuthenticationCertificatesText" {$excel.cells.item($row,$column) = $value; break}	
			"SslCertificatesText" {$excel.cells.item($row,$column) = $value; break}	
			"FrontendIpConfigurationsText" {$excel.cells.item($row,$column) = $value; break}	
			"FrontendPortsText" {$excel.cells.item($row,$column) = $value; break}	
			"BackendAddressPoolsText" {$excel.cells.item($row,$column) = $value; break}	
			"BackendHttpSettingsCollectionText" {$excel.cells.item($row,$column) = $value; break}	
			"HttpListenersText" {$excel.cells.item($row,$column) = $value; break}	
			"RequestRoutingRulesText" {$excel.cells.item($row,$column) = $value; break}	
			"ProbesText" {$excel.cells.item($row,$column) = $value; break}	
			"UrlPathMapsText" {$excel.cells.item($row,$column) = $value; break}	
			"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}	
			"Location" {$excel.cells.item($row,$column) = $value; break}	
			"ResourceGuid" {$excel.cells.item($row,$column) = $value; break}	
			"Type" {$excel.cells.item($row,$column) = $value.Id; break}	
			"Tag" {$excel.cells.item($row,$column) = $value.Name; break}	
			"TagsTable" {$excel.cells.item($row,$column) = $value; break}	
			"Name" {$excel.cells.item($row,$column) = $value; break}	
			"Etag" {$excel.cells.item($row,$column) = $value; break}	
			"Id" {$excel.cells.item($row,$column) = $value; break}

			
			default {$excel.cells.item($row,$column) = "Error"}
					}
					$column += 1
		}
		$column = 1
		$worksheet.Rows($row).RowHeight = 15
		$row += 1
	}
	
# Firewalls
Write-Output("========= Firewalls ==========")
	
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Firewalls"

$column = 1
$row = 1
	
	$firewalls = Get-AzureRmFirewall
	foreach($wall in $firewalls){
	$wall.PSObject.Properties | ForEach-Object {
					$excel.cells.item($row,$column) = $_.Name
					Write-Output($_.Name + ": " + $_.Value)
					$column += 1
		}
		
		$column = 1
		break
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1
	
foreach($wall in $firewalls){
	$wall.PSObject.Properties | ForEach-Object {
		$value = $_.Value
		Write-Output($_.Name + ": " +$value)
		switch($_.Name){
			"IpConfigurations" {$excel.cells.item($row,$column) = $value.Id; break} # List of Objects
			"ApplicationRuleCollections" {$excel.cells.item($row,$column) = $value.Id; break} # List of Objects
			"NatRuleCollections" {$excel.cells.item($row,$column) = $value.Id; break} # List of Objects
			"NetworkRuleCollections" {$excel.cells.item($row,$column) = $value.Id; break}  # List of Objects
			"ProvisioningState" {$excel.cells.item($row,$column) = $value; break}
			"IpConfigurationsText" {$excel.cells.item($row,$column) = $value; break}
			"ApplicationRuleCollectionsText" {$excel.cells.item($row,$column) = $value; break}
			"NatRuleCollectionsText" {$excel.cells.item($row,$column) = $value; break}
			"NetworkRuleCollectionsText" {$excel.cells.item($row,$column) = $value; break}
			"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}
			"Location" {$excel.cells.item($row,$column) = $value; break}
			"ResourceGuid" {$excel.cells.item($row,$column) = $value; break}
			"Type" {$excel.cells.item($row,$column) = $value; break}
			"Tag" {$excel.cells.item($row,$column) = $value.Name; break}
			"TagsTable" {$excel.cells.item($row,$column) = $value; break}
			"Name" {$excel.cells.item($row,$column) = $value; break}
			"Etag" {$excel.cells.item($row,$column) = $value; break}
			"Id" {$excel.cells.item($row,$column) = $value; break}

			default {$excel.cells.item($row,$column) = "Error"}
		}
		$column += 1
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1
}


Write-Output("========= Web Applications ==========")
	
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Web Apps"

$column = 1
$row = 1
	
$applications = Get-AzureRmWebApp
foreach($app in $applications){
	$app.PSObject.Properties | ForEach-Object {
		$excel.cells.item($row,$column) = $_.Name
		Write-Output($_.Name + ": " + $_.Value)
		$column += 1
	}
	$column = 1
	break
}

$column = 1
$worksheet.Rows($row).RowHeight = 15
$row += 1
	
foreach($app in $applications){
	$app.PSObject.Properties | ForEach-Object {
		$value = $_.Value
		Write-Output($_.Name + ": " +$value)
		switch($_.Name){
		
			"GitRemoteName" {$excel.cells.item($row,$column) = $value; break}
			"GitRemoteUri" {$excel.cells.item($row,$column) = $value; break}
			"GitRemoteUsername" {$excel.cells.item($row,$column) = $value; break}
			"GitRemotePassword" {$excel.cells.item($row,$column) = $value; break} # secure string
			"SnapshotInfo" {$excel.cells.item($row,$column) = $value; break}
			"State" {$excel.cells.item($row,$column) = $value; break}
			"HostNames" {$excel.cells.item($row,$column) = arrayToString($value); break}
			"RepositorySiteName" {$excel.cells.item($row,$column) = $value; break}
			"UsageState" {$excel.cells.item($row,$column) = $value; break} # System.Nullable[Microsoft.Azure.Management.WebSites.Models.UsageState
			"Enabled" {$excel.cells.item($row,$column) = $value; break}
			"EnabledHostNames" {$excel.cells.item($row,$column) = arrayToString($value); break}
			"AvailabilityState" {$excel.cells.item($row,$column) = $value; break}
			"HostNameSslStates" {$excel.cells.item($row,$column) = $value.Name; break} # List of Objects
			"ServerFarmId" {$excel.cells.item($row,$column) = $value; break}
			"Reserved" {$excel.cells.item($row,$column) = $value; break}
			"IsXenon" {$excel.cells.item($row,$column) = $value; break}
			"LastModifiedTimeUtc" {$excel.cells.item($row,$column) = $value; break} # datetime
			"SiteConfig" {$excel.cells.item($row,$column) = $value; break} # object
			"TrafficManagerHostNames" {$excel.cells.item($row,$column) = arrayToString($value); break}
			"ScmSiteAlsoStopped" {$excel.cells.item($row,$column) = $value; break}
			"TargetSwapSlot" {$excel.cells.item($row,$column) = $value; break}
			"HostingEnvironmentProfile" {$excel.cells.item($row,$column) = $value; break} # object
			"ClientAffinityEnabled" {$excel.cells.item($row,$column) = $value; break}
			"ClientCertEnabled" {$excel.cells.item($row,$column) = $value; break}
			"HostNamesDisabled" {$excel.cells.item($row,$column) = $value; break}
			"OutboundIpAddresses" {$excel.cells.item($row,$column) = $value; break}
			"PossibleOutboundIpAddresses" {$excel.cells.item($row,$column) = $value; break}
			"ContainerSize" {$excel.cells.item($row,$column) = $value; break}
			"DailyMemoryTimeQuota" {$excel.cells.item($row,$column) = $value; break}
			"SuspendedTill" {$excel.cells.item($row,$column) = $value; break} # datetime
			"MaxNumberOfWorkers" {$excel.cells.item($row,$column) = $value; break}
			"CloningInfo" {$excel.cells.item($row,$column) = $value; break} # object
			"ResourceGroup" {$excel.cells.item($row,$column) = $value; break}
			"IsDefaultContainer" {$excel.cells.item($row,$column) = $value; break}
			"DefaultHostName" {$excel.cells.item($row,$column) = $value; break}
			"SlotSwapStatus" {$excel.cells.item($row,$column) = $value; break}
			"HttpsOnly" {$excel.cells.item($row,$column) = $value; break}
			"Identity" {$excel.cells.item($row,$column) = $value; break}
			"Id" {$excel.cells.item($row,$column) = $value; break}
			"Name" {$excel.cells.item($row,$column) = $value; break}
			"Kind" {$excel.cells.item($row,$column) = $value; break}
			"Location" {$excel.cells.item($row,$column) = $value; break}
			"Type" {$excel.cells.item($row,$column) = $value; break}
			"Tags" {$excel.cells.item($row,$column) = $value.Name; break}

			default {$excel.cells.item($row,$column) = "Error"}
		}
		$column += 1
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1
}


# Public IP addresses

Write-Output("========= Public IP Addresses ==========")
	
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Public IPs"
$column = 1
$row = 1	
$ips = Get-AzureRmPublicIpAddress

foreach($ip in $ips){
	$ip.PSObject.Properties | ForEach-Object {
		$excel.cells.item($row,$column) = $_.Name
		Write-Output($_.Name + ": " + $_.Value)
		$column += 1
	}
	$column = 1
	break
}

$column = 1
$worksheet.Rows($row).RowHeight = 15
$row += 1
	
foreach($ip in $ips){
	$ip.PSObject.Properties | ForEach-Object {
		$value = $_.Value
		Write-Output($_.Name + ": " + $value)
		switch($_.Name){
			"PublicIpAllocationMethod" {$excel.cells.item($row,$column) = $value; break}
			"Sku" {$excel.cells.item($row,$column) = $value.Name; break} # object
			"IpConfiguration" {$excel.cells.item($row,$column) = $value.Id; break}
			"DnsSettings" {$excel.cells.item($row,$column) = $value.Fqdn; break} #object 
			"IpTags" {$excel.cells.item($row,$column) = $value.Id; break} # list of objects
			"IpAddress" {$excel.cells.item($row,$column) = $value; break}
			"PublicIpAddressVersion" {$excel.cells.item($row,$column) = $value; break}
			"IdleTimeoutInMinutes" {$excel.cells.item($row,$column) = $value; break}
			"Zones" {$excel.cells.item($row,$column) = $value.Id; break}
			"ProvisioningState" {$excel.cells.item($row,$column) = $value; break}
			"PublicIpPrefix" {$excel.cells.item($row,$column) = $value; break}
			"IpConfigurationText" {$excel.cells.item($row,$column) = $value; break}
			"DnsSettingsText" {$excel.cells.item($row,$column) = $value; break}
			"IpTagsText" {$excel.cells.item($row,$column) = $value; break}
			"SkuText" {$excel.cells.item($row,$column) = $value; break}
			"PublicIpPrefixText" {$excel.cells.item($row,$column) = $value; break}
			"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}
			"Location" {$excel.cells.item($row,$column) = $value; break}
			"ResourceGuid" {$excel.cells.item($row,$column) = $value; break}
			"Type" {$excel.cells.item($row,$column) = $value; break}
			"Tag" {$excel.cells.item($row,$column) = $value; break}
			"TagsTable" {$excel.cells.item($row,$column) = $value; break}
			"Name" {$excel.cells.item($row,$column) = $value; break}
			"Etag" {$excel.cells.item($row,$column) = $value; break}
			"Id" {$excel.cells.item($row,$column) = $value; break}

			default {$excel.cells.item($row,$column) = "Error"}
		}
		$column += 1
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1
}
	
	
Write-Output("Done") 