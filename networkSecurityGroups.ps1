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
$worksheet.Rows($row).RowHeight = 15
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
	$worksheet.Rows($row).RowHeight = 15
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
	$worksheet.Rows($row).RowHeight = 15
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
				$worksheet.Rows($row).RowHeight = 15
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
					$column = $column + 1
				
		}
		break
	}

	$column = 1
	$row = $row + 1

foreach($interface in $interfaces){
	$interface.PSObject.Properties | ForEach-Object {
					$value = $_.Value
					switch($_.Name){
						"VirtualMachine" {$excel.cells.item($row,$column) = $value.Id; break}
						# contains if IPv4, Name, Primary, private addresss,  PrivateIpAllocationMethod 
						"IpConfigurations" {$excel.cells.item($row,$column) = $value.Name; break}
						"TapConfigurations" {$excel.cells.item($row,$column) = $value[0]; break}
						# contains DnsServers, AppliedDnsServers, Internal, DnsNameLabel, InternalFqdn, InternalDomainNameSuffix, DnsServersText, AppliedDnsServersText
						"DnsSettings" {$excel.cells.item($row,$column) = $value.DnsNameLabel; break}
						"MacAddress" {$excel.cells.item($row,$column) = $value; break}
						"Primary" {$excel.cells.item($row,$column) = $value; break}
						"EnableAcceleratedNetworking" {$excel.cells.item($row,$column) = $value; break}
						"EnableIPForwarding" {$excel.cells.item($row,$column) = $value; break}
						"HostedWorkloads" {$excel.cells.item($row,$column) = $value[0]; break}
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
						"Tag" {$excel.cells.item($row,$column) = $value; break}
						"TagsTable" {$excel.cells.item($row,$column) = $value; break}
						"Name" {$excel.cells.item($row,$column) = $value; break}
						"Etag" {$excel.cells.item($row,$column) = $value; break}
						"Id" {$excel.cells.item($row,$column) = $value; break}
						default {$excel.cells.item($row,$column) = "Error"}
					}
					
					$column = $column + 1
				}
		$worksheet.Rows($row).RowHeight = 15
		$row = $row + 1
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
					$column = $column + 1
				}
				break
				}
				
			}
		}
		$column = 1
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row = $row + 1

foreach($group in $securityGroups){
	$group.PSObject.Properties | ForEach-Object {
			if($_.Name -eq "Subnets"){
				foreach($value in $_.Value){
					$value.PSObject.Properties | ForEach-Object{
					
					switch($_.Name){
						"AddressPrefix" {$excel.cells.item($row,$column) = $value.AddressPrefix; break}
						"IpConfigurations" {$excel.cells.item($row,$column) = $value.IpConfigurations[0]; break}
						"ServiceAssociationLinks" {$excel.cells.item($row,$column) = $value.ServiceAssociationLinks[0]; break}
						"ResourceNavigationLinks" {$excel.cells.item($row,$column) = $value.ResourceNavigationLinks[0]; break}
						# check if this is right
						"NetworkSecurityGroup" {$excel.cells.item($row,$column) = $group.Id; break}
						"RouteTable" {$excel.cells.item($row,$column) = $value.RouteTable; break}
						"ServiceEndpoints" {$excel.cells.item($row,$column) = $value.ServiceEndpoints[0]; break}
						"ServiceEndpointPolicies" {$excel.cells.item($row,$column) = $value.ServiceEndpointPolicies[0]; break}
						"Delegations" {$excel.cells.item($row,$column) = $value.Delegations[0]; break}
						"InterfaceEndpoints" {$excel.cells.item($row,$column) = $value.InterfaceEndpoints[0]; break}
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
					
					$column = $column + 1
				}
				$worksheet.Rows($row).RowHeight = 15
				$row = $row + 1
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
					$column = $column + 1
		}
		
		$column = 1
		break
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row = $row + 1

foreach($network in $networks){
	$network.PSObject.Properties | ForEach-Object {
						$value = $_.Value
						Write-Output($_.Name + ": " +$value)
						switch($_.Name){
						"AddressSpace" {$excel.cells.item($row,$column) = $value.AddressPrefixes[0]; break}
						"DhcpOptions" {$excel.cells.item($row,$column) = $value; break}
						"Subnets" {$excel.cells.item($row,$column) = $value.Name;break}
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
						"Tag" {$excel.cells.item($row,$column) = $value; break}
						"TagsTable" {$excel.cells.item($row,$column) = $value; break}	
						"Name" {$excel.cells.item($row,$column) = $value; break}	
						"Etag" {$excel.cells.item($row,$column) = $value; break}	
						"Id" {$excel.cells.item($row,$column) = $value; break}
						
						default {$excel.cells.item($row,$column) = "Error"}
					}
					$column = $column + 1
		}
					$column = 1
					$worksheet.Rows($row).RowHeight = 15
					$row = $row + 1
					$subnets.add($network.Subnets)
	}
	
	Write-Output("==== All Subnets =====")
	$subnets
	$worksheet = $workbook.worksheets.add($worksheet)
	$worksheet.Name = "Subnets"
	$column = 1
	$row = 1
	
	# TODO: programatically add column names and values instead of hardcoded column names etc
			$excel.cells.item($row,$column) ="AddressPrefix"
			$excel.cells.item($row,$column + 1) ="IpConfigurations"	
			$excel.cells.item($row,$column + 2) ="ServiceAssociationLinks"	
			$excel.cells.item($row,$column + 3) ="ResourceNavigationLinks"	
			$excel.cells.item($row,$column + 4) ="NetworkSecurityGroup"	
			$excel.cells.item($row,$column + 5) ="RouteTable"	
			$excel.cells.item($row,$column + 6) ="ServiceEndpoints"	
			$excel.cells.item($row,$column + 7) ="ServiceEndpointPolicies"	
			$excel.cells.item($row,$column + 8) ="Delegations"	
			$excel.cells.item($row,$column + 9) ="InterfaceEndpoints"	
			$excel.cells.item($row,$column + 10) ="ProvisioningState"	
			$excel.cells.item($row,$column + 11) ="IpConfigurationsText"	
			$excel.cells.item($row,$column + 12) ="ServiceAssociationLinksText"	
			$excel.cells.item($row,$column + 13) ="ResourceNavigationLinksText"	
			$excel.cells.item($row,$column + 14) ="NetworkSecurityGroupText"	
			$excel.cells.item($row,$column + 15) ="RouteTableText"	
			$excel.cells.item($row,$column + 16) ="ServiceEndpointText"	
			$excel.cells.item($row,$column + 17) ="ServiceEndpointPoliciesText"	
			$excel.cells.item($row,$column + 18) ="InterfaceEndpointsText"	
			$excel.cells.item($row,$column + 19) ="DelegationsText"	
			$excel.cells.item($row,$column + 20) ="Name"	
			$excel.cells.item($row,$column + 21) ="Etag"	
			$excel.cells.item($row,$column + 22) ="Id"

			$column = 1
			$row = $row + 1
			foreach($subnet in $subnets){
			$excel.cells.item($row,$column) = $subnet.AddressPrefix
			$excel.cells.item($row,$column + 1) = $subnet.IpConfigurations.Id
			$excel.cells.item($row,$column + 2) = $subnet.ServiceAssociationLinks
			$excel.cells.item($row,$column + 3) = $subnet.ResourceNavigationLinks
			$excel.cells.item($row,$column + 4) = $subnet.NetworkSecurityGroup.Id
			$excel.cells.item($row,$column + 5) = $subnet.RouteTable
			$excel.cells.item($row,$column + 6) = $subnet.ServiceEndpoints
			$excel.cells.item($row,$column + 7) = $subnet.ServiceEndpointPolicies	
			$excel.cells.item($row,$column + 8) = $subnet.Delegations
			$excel.cells.item($row,$column + 9) = $subnet.InterfaceEndpoints
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
			$row = $row + 1
			}
	
Write-Output("Done") 