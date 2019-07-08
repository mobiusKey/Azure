# Install-Module PSSQLite
# Import-Module

function encode {
	if($args[0] -ne $null){
	$encoded = [System.Web.HttpUtility]::HtmlEncode($args[0])
	return $encoded
	}
	return $null
}

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

$DataSource = Get-Location
$DataSource = $DataSource.Path + "\Test.SQLite"

# Network Security Groups
# TABLES NSG, SecurityRules, DefaultSecurityRules, Subnets, NetworkInterfaces
$Query = "CREATE TABLE NetworkSecurityGroup (
	SecurityRules TEXT,
	DefaultSecurityRules TEXT,
	NetworkInterfaces TEXT,
	Subnets TEXT,
	ProvisioningState TEXT,
	SecurityRulesText TEXT,
	DefaultSecurityRulesText TEXT,
	NetworkInterfacesText TEXT,
	SubnetsText TEXT,
	ResourceGroupName TEXT,
	Location TEXT,
	ResourceGuid TEXT,
	Type TEXT,
	Tag TEXT,
	TagsTable TEXT,
	Name TEXT,
	Etag TEXT,
	Id TEXT PRIMARY KEY NOT NULL
	)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query

# Security Rules

$Query = "CREATE TABLE SecurityRules (
	Description TEXT, 
	Protocol TEXT, 
	SourcePortRange TEXT, 
	DestinationPortRange TEXT, 
	SourceAddressPrefix TEXT, 
	DestinationAddressPrefix TEXT, 
	Access TEXT, 
	Priority TEXT, 
	Direction TEXT, 
	ProvisioningState TEXT, 
	SourceApplicationSecurityGroups TEXT, 
	DestinationApplicationSecurityGroups TEXT, 
	SourceApplicationSecurityGroupsText TEXT, 
	DestinationApplicationSecurityGroupsText TEXT, 
	Name TEXT, 
	Etag TEXT, 
	Id TEXT PRIMARY KEY NOT NULL,
	NetworkSecurityGroup TEXT NOT NULL,
	FOREIGN KEY(NetworkSecurityGroup) REFERENCES NetworkSecurityGroup(Id))"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query


# Default Security Rules

$Query = "CREATE TABLE DefaultSecurityRules (
	Description TEXT, 
	Protocol TEXT, 
	SourcePortRange TEXT, 
	DestinationPortRange TEXT, 
	SourceAddressPrefix TEXT, 
	DestinationAddressPrefix TEXT, 
	Access TEXT, 
	Priority TEXT, 
	Direction TEXT, 
	ProvisioningState TEXT, 
	SourceApplicationSecurityGroups TEXT, 
	DestinationApplicationSecurityGroups TEXT, 
	SourceApplicationSecurityGroupsText TEXT, 
	DestinationApplicationSecurityGroupsText TEXT, 
	Name TEXT, 
	Etag TEXT, 
	Id TEXT PRIMARY KEY,
	NetworkSecurityGroup TEXT NOT NULL,
	FOREIGN KEY(NetworkSecurityGroup) REFERENCES NetworkSecurityGroup(Id)
	)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query


#VirtualMachines
$Query = "CREATE TABLE VirtualMachine (
	ResourceGroupName TEXT,
	Id TEXT PRIMARY KEY,
	VmId TEXT,
	Name TEXT,
	Type TEXT,
	Location TEXT,
	LicenseType TEXT,
	Tags TEXT,
	AvailabilitySetReference TEXT,
	DiagnosticsProfile TEXT,
	Extensions TEXT,
	HardwareProfile TEXT,
	InstanceView TEXT,
	NetworkProfile TEXT,
	OSProfile TEXT,
	AdminUsername TEXT,
	AdminPassword TEXT,
	CustomData TEXT,
	WindowsConfiguration TEXT,
	Secrets TEXT,
	Plan TEXT,
	ProvisioningState TEXT,
	StorageProfile TEXT,
	Offer TEXT,
	Sku TEXT,
	Version TEXT,
	DisplayHint TEXT,
	Identity TEXT,
	Zones TEXT,
	FullyQualifiedDomainName TEXT,
	AdditionalCapabilities TEXT,
	RequestId TEXT,
	StatusCode TEXT
	)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query


#Subnets

$Query = "CREATE TABLE Subnet (
	AddressPrefix TEXT,
	IpConfigurations TEXT,
	ServiceAssociationLinks TEXT,
	ResourceNavigationLinks TEXT,
	NetworkSecurityGroup TEXT,
	RouteTable TEXT,
	ServiceEndpoints TEXT,
	ServiceEndpointPolicies TEXT,
	Delegations TEXT,
	InterfaceEndpoints TEXT,
	ProvisioningState TEXT,
	IpConfigurationsText TEXT,
	ServiceAssociationLinksText TEXT,
	ResourceNavigationLinksText TEXT,
	NetworkSecurityGroupText TEXT,
	RouteTableText TEXT,
	ServiceEndpointText TEXT,
	ServiceEndpointPoliciesText TEXT,
	InterfaceEndpointsText TEXT,
	DelegationsText TEXT,
	Name TEXT,
	Etag TEXT,
	Id TEXT PRIMARY KEY 
)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query

#Network Interfaces
$Query = "CREATE TABLE NetworkInterface (
	VirtualMachine TEXT, 
	IpConfigurations TEXT, 
	TapConfigurations TEXT, 
	DnsSettings TEXT, 
	MacAddress TEXT, 
	PrimaryBool TEXT, 
	EnableAcceleratedNetworking TEXT, 
	EnableIPForwarding TEXT, 
	HostedWorkloads TEXT, 
	NetworkSecurityGroup TEXT, 
	ProvisioningState TEXT, 
	VirtualMachineText TEXT, 
	IpConfigurationsText TEXT,
	TapConfigurationsText TEXT, 
	DnsSettingsText TEXT, 
	NetworkSecurityGroupText TEXT, 
	ResourceGroupName TEXT, 
	Location TEXT, 
	ResourceGuid TEXT, 
	Type TEXT, 
	Tag TEXT, 
	TagsTable TEXT, 
	Name TEXT, 
	Etag TEXT, 
	Id TEXT,
	FOREIGN KEY(VirtualMachine) REFERENCES VirtualMachine(Id),
	FOREIGN KEY(NetworkSecurityGroup) REFERENCES NetworkSecurityGroup(Id) 
	)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query


# Inserting Data 
$groups = Get-AzureRmNetworkSecurityGroup

foreach($group in $groups){
	$Query = 'INSERT INTO NetworkSecurityGroup (SecurityRules, DefaultSecurityRules, NetworkInterfaces, Subnets, ProvisioningState,SecurityRulesText, DefaultSecurityRulesText, NetworkInterfacesText, SubnetsText, ResourceGroupName, Location, ResourceGuid, Type, Tag, TagsTable, Name, Etag, Id)  
	VALUES ("' + (encode($group.SecurityRules.Id)) 
	$Query += '", "' + (encode($group.DefaultSecurityRules.Id)) 
	$Query += '", "' + (encode($group.NetworkInterfaces.Id)) 
	$Query += '", "' + (encode($group.Subnets.Id)) 
	$Query += '", "' + (encode($group.ProvisioningState))
	$Query += '", "' + (encode($group.SecurityRulesText)) 
	$Query += '", "' + (encode($group.DefaultSecurityRulesText)) 
	$Query += '", "' + (encode($group.NetworkInterfacesText)) 
	$Query += '", "' + (encode($group.SubnetsText))
	$Query += '", "' + (encode($group.ResourceGroupName)) 
	$Query += '", "' + (encode($group.Location))
	$Query += '", "' + (encode($group.ResourceGuid)) 
	$Query += '", "' + (encode($group.Type))
	$Query += '", "' + (encode($group.Tag))
	# TODO create relationship table for Tags Table
	$Query += '", "' + (encode($group.TagsTable))
	$Query += '", "' + (encode($group.Name))
	$Query += '", "' + (encode($group.Etag))
	$Query += '", "' + (encode($group.Id)) + '")'
		Invoke-SqliteQuery -DataSource $DataSource -Query $Query

	foreach($rule in $group.SecurityRules){

		$Query = 'INSERT INTO SecurityRules (Description, Protocol, SourcePortRange, DestinationPortRange, SourceAddressPrefix, DestinationAddressPrefix, Access, Priority, Direction, ProvisioningState, SourceApplicationSecurityGroups, DestinationApplicationSecurityGroups, SourceApplicationSecurityGroupsText, DestinationApplicationSecurityGroupsText, Name, Etag, Id, NetworkSecurityGroup)  
		VALUES ("' + (encode($rule.Description)) 
		$Query += '", "' + (encode($rule.Protocol)) 
		$Query += '", "' + (arrayToString($rule.SourcePortRange))
		$Query += '", "' + (arrayToString($rule.DestinationPortRange)) 
		$Query += '", "' + ($rule.SourceAddressPrefix)
		$Query += '", "' + ($rule.DestinationAddressPrefix) 
		$Query += '", "' + (encode($rule.Access)) 
		$Query += '", "' + (encode($rule.Priority)) 
		$Query += '", "' + (encode($rule.Direction))
		$Query += '", "' + (encode($rule.ProvisioningState)) 
		$Query += '", "' + (encode($rule.SourceApplicationSecurityGroups))
		$Query += '", "' + (encode($rule.DestinationApplicationSecurityGroups))
		$Query += '", "' + (encode($rule.SourceApplicationSecurityGroupsText))
		$Query += '", "' + (encode($rule.DestinationApplicationSecurityGroupsText))
		$Query += '", "' + (encode($rule.Name))
		$Query += '", "' + (encode($rule.Etag))
		$Query += '", "' + (encode($rule.Id))
		$Query += '", "' + (encode($group.Id)) + '")'
		Invoke-SqliteQuery -DataSource $DataSource -Query $Query

	}
	
	foreach($rule in $group.DefaultSecurityRules){

		$Query = 'INSERT INTO DefaultSecurityRules (Description, Protocol, SourcePortRange, DestinationPortRange, SourceAddressPrefix, DestinationAddressPrefix, Access, Priority, Direction, ProvisioningState, SourceApplicationSecurityGroups, DestinationApplicationSecurityGroups, SourceApplicationSecurityGroupsText, DestinationApplicationSecurityGroupsText, Name, Etag, Id, NetworkSecurityGroup)  
		VALUES ("' + (encode($rule.Description)) 
		$Query += '", "' + (encode($rule.Protocol)) 
		$Query += '", "' + (arrayToString($rule.SourcePortRange))
		$Query += '", "' + (arrayToString($rule.DestinationPortRange)) 
		$Query += '", "' + ($rule.SourceAddressPrefix)
		$Query += '", "' + ($rule.DestinationAddressPrefix) 
		$Query += '", "' + (encode($rule.Access)) 
		$Query += '", "' + (encode($rule.Priority)) 
		$Query += '", "' + (encode($rule.Direction))
		$Query += '", "' + (encode($rule.ProvisioningState)) 
		$Query += '", "' + (encode($rule.SourceApplicationSecurityGroups))
		$Query += '", "' + (encode($rule.DestinationApplicationSecurityGroups))
		$Query += '", "' + (encode($rule.SourceApplicationSecurityGroupsText))
		$Query += '", "' + (encode($rule.DestinationApplicationSecurityGroupsText))
		$Query += '", "' + (encode($rule.Name))
		$Query += '", "' + (encode($rule.Etag))
		$Query += '", "' + (encode($rule.Id))
		$Query += '", "' + (encode($group.Id)) + '")'
		
		Invoke-SqliteQuery -DataSource $DataSource -Query $Query

	}
}

# VirtualMachines
$vms = Get-AzureRmVM

foreach($vm in $vms){
		$Query = 'INSERT INTO VirtualMachine (ResourceGroupName, Id, VmId, Name, Type, Location, LicenseType, Tags, AvailabilitySetReference, DiagnosticsProfile, Extensions, HardwareProfile, InstanceView, NetworkProfile, OSProfile, AdminUsername, AdminPassword, CustomData, WindowsConfiguration, Secrets, Plan, ProvisioningState, StorageProfile, Offer, Sku, Version, DisplayHint, Identity, Zones, FullyQualifiedDomainName, AdditionalCapabilities, RequestId, StatusCode) 
		VALUES ("' + $vm.ResourceGroupName + '", "' + $vm.Id+ '", "' + $vm.VmId+ '", "' + $vm.Name+ '", "' + $vm.Type+ '", "' + $vm.Location+ '", "' + $vm.LicenseType+ '", "' + $vm.Tags+ '", "' + $vm.AvailabilitySetReference+ '", "' + $vm.DiagnosticsProfile+ '", "' + $vm.Extensions+ '", "' + $vm.HardwareProfile+ '", "' + $vm.InstanceView+ '", "' + $vm.NetworkProfile+ '", "' + $vm.OSProfile+ '", "' + $vm.AdminUsername+ '", "' + $vm.AdminPassword+ '", "' + $vm.CustomData+ '", "' + $vm.WindowsConfiguration+ '", "' + $vm.Secrets+ '", "' + $vm.Plan+ '", "' + $vm.ProvisioningState+ '", "' + $vm.StorageProfile+ '", "' + $vm.Offer+ '", "' + $vm.Sku+ '", "' + $vm.Version+ '", "' + $vm.DisplayHint+ '", "' + $vm.Identity+ '", "' + $vm.Zones+ '", "' + $vm.FullyQualifiedDomainName+ '", "' + $vm.AdditionalCapabilities+ '", "' + $vm.RequestId+ '", "' + $vm.StatusCode + '")'
	Invoke-SqliteQuery -DataSource $DataSource -Query $Query
	}

# Network Interfaces

$interfaces = Get-AzureRmNetworkInterface
foreach($interface in $interfaces){

	$Query = 'INSERT INTO NetworkInterface(	VirtualMachine TEXT, 
	IpConfigurations, 
	TapConfigurations, 
	DnsSettings, 
	MacAddress, 
	PrimaryBool, 
	EnableAcceleratedNetworking, 
	EnableIPForwarding, 
	HostedWorkloads, 
	NetworkSecurityGroup, 
	ProvisioningState, 
	VirtualMachine, 
	IpConfigurations,
	TapConfigurations, 
	DnsSettings, 
	NetworkSecurityGroup, 
	ResourceGroupName, 
	Location, 
	ResourceGuid, 
	Type, 
	Tag, 
	TagsTable, 
	Name, 
	Etag, 
	Id)  
	VALUES ("' + (encode($interface.IpConfigurations.Id)) 
	$Query += '", "' + (encode($interface.TapConfigurations)) 
	$Query += '", "' + (encode($interface.DnsSettings)) 
	$Query += '", "' + (arrayToString($rule.SourcePortRange))
	$Query += '", "' + (arrayToString($rule.DestinationPortRange)) 
	$Query += '", "' + ($rule.SourceAddressPrefix)
	$Query += '", "' + ($rule.DestinationAddressPrefix) 
	$Query += '", "' + (encode($rule.Access)) 
	$Query += '", "' + (encode($rule.Priority)) 
	$Query += '", "' + (encode($rule.Direction))
	$Query += '", "' + (encode($rule.ProvisioningState)) 
	$Query += '", "' + (encode($rule.SourceApplicationSecurityGroups))
	$Query += '", "' + (encode($rule.DestinationApplicationSecurityGroups))
	$Query += '", "' + (encode($rule.SourceApplicationSecurityGroupsText))
	$Query += '", "' + (encode($rule.DestinationApplicationSecurityGroupsText))
	$Query += '", "' + (encode($rule.Name))
	$Query += '", "' + (encode($rule.Etag))
	$Query += '", "' + (encode($rule.Id))
	$Query += '", "' + (encode($group.Id)) + '")'
	
	Invoke-SqliteQuery -DataSource $DataSource -Query $Query

}