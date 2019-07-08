# Install-Module PSSQLite
# Import-Module

# Note that most values are never NULL but are of length 0
# TODO add Function Apps and API apps
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
	Plan TEXT,
	ProvisioningState TEXT,
	StorageProfile TEXT,
	Identity TEXT,
	Zones TEXT,
	FullyQualifiedDomainName TEXT,
	AdditionalCapabilities TEXT,
	RequestId TEXT,
	StatusCode TEXT
	)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query

# OSProfile
$Query = "CREATE TABLE OSProfile (
	ComputerName TEXT,
	AdminUsername TEXT,
	AdminPassword TEXT,
	CustomData TEXT,
	WindowsConfiguration TEXT,
	LinuxConfiguration TEXT,
	AllowExtensionOperations TEXT,
	Secrets TEXT,
	Id TEXT PRIMARY KEY,
	FOREIGN KEY(Id) REFERENCES VirtualMachine(Id)
)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query

# TODO: StorageProfile

# Subnets
# TODO: Create RouteTable TABLE
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
	Id TEXT PRIMARY KEY,
	VirtualNetwork TEXT,
	FOREIGN KEY(VirtualNetwork) REFERENCES VirtualNetwork(Id)
	FOREIGN KEY (NetworkSecurityGroup) REFERENCES NetworkSecurityGroup(Id)
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

# Virtual Networks
$Query = "CREATE TABLE VirtualNetwork (
	AddressSpace TEXT,
	DhcpOptions TEXT,
	Subnets TEXT,
	VirtualNetworkPeerings TEXT,
	ProvisioningState TEXT,
	EnableDdosProtection TEXT,
	EnableVmProtection TEXT,
	DdosProtectionPlan TEXT,
	AddressSpaceText TEXT,
	DhcpOptionsText TEXT,
	SubnetsText TEXT,
	VirtualNetworkPeeringsText TEXT,
	EnableDdosProtectionText TEXT,
	DdosProtectionPlanText TEXT,
	EnableVmProtectionText TEXT,
	ResourceGroupName TEXT,
	Location TEXT,
	ResourceGuid TEXT,
	Type TEXT,
	Tag TEXT,
	TagsTable TEXT,
	Name TEXT,
	Etag TEXT,
	Id TEXT PRIMARY KEY 
)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query

$Query = "CREATE TABLE WebApp (
	GitRemoteName TEXT,
	GitRemoteUri TEXT,
	GitRemoteUsername TEXT,
	GitRemotePassword TEXT,
	SnapshotInfo TEXT,
	State TEXT,
	HostNames TEXT,
	RepositorySiteName TEXT,
	UsageState TEXT,
	Enabled TEXT,
	EnabledHostNames TEXT,
	AvailabilityState TEXT,
	HostNameSslStates TEXT,
	ServerFarmId TEXT,
	Reserved TEXT,
	IsXenon TEXT,
	LastModifiedTimeUtc TEXT,
	SiteConfig TEXT,
	TrafficManagerHostNames TEXT,
	ScmSiteAlsoStopped TEXT,
	TargetSwapSlot TEXT,
	HostingEnvironmentProfile TEXT,
	ClientAffinityEnabled TEXT,
	ClientCertEnabled TEXT,
	HostNamesDisabled TEXT,
	OutboundIpAddresses TEXT,
	PossibleOutboundIpAddresses TEXT,
	ContainerSize TEXT,
	DailyMemoryTimeQuota TEXT,
	SuspendedTill TEXT,
	MaxNumberOfWorkers TEXT,
	CloningInfo TEXT,
	ResourceGroup TEXT,
	IsDefaultContainer TEXT,
	DefaultHostName TEXT,
	SlotSwapStatus TEXT,
	HttpsOnly TEXT,
	Identity TEXT,
	Id TEXT PRIMARY KEY,
	Name TEXT,
	Kind TEXT,
	Location TEXT,
	Type TEXT,
	Tags TEXT
)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query

# Application Gateways

$Query = "CREATE TABLE ApplicationGateway (
	Sku TEXT,
	SslPolicy TEXT,
	GatewayIPConfigurations TEXT,
	AuthenticationCertificates TEXT,
	SslCertificates TEXT,
	TrustedRootCertificates TEXT,
	FrontendIPConfigurations TEXT,
	FrontendPorts TEXT,
	Probes TEXT,
	BackendAddressPools TEXT,
	BackendHttpSettingsCollection TEXT,
	HttpListeners TEXT,
	UrlPathMaps TEXT,
	RequestRoutingRules TEXT,
	RedirectConfigurations TEXT,
	WebApplicationFirewallConfiguration TEXT,
	AutoscaleConfiguration TEXT,
	CustomErrorConfigurations TEXT,
	EnableHttp2 TEXT,
	EnableFips TEXT,
	Zones TEXT,
	OperationalState TEXT,
	ProvisioningState TEXT,
	GatewayIpConfigurationsText TEXT,
	AuthenticationCertificatesText TEXT,
	SslCertificatesText TEXT,
	FrontendIpConfigurationsText TEXT,
	FrontendPortsText TEXT,
	BackendAddressPoolsText TEXT,
	BackendHttpSettingsCollectionText TEXT,
	HttpListenersText TEXT,
	RequestRoutingRulesText TEXT,
	ProbesText TEXT,
	UrlPathMapsText TEXT,
	ResourceGroupName TEXT,
	Location TEXT,
	ResourceGuid TEXT,
	Type TEXT,
	Tag TEXT,
	TagsTable TEXT,
	Name TEXT,
	Etag TEXT,
	Id TEXT PRIMARY KEY
)"

Invoke-SqliteQuery -DataSource $DataSource -Query $Query

# Firewalls

$Query = "CREATE TABLE Firewall (
	IpConfigurations TEXT,
	ApplicationRuleCollections TEXT,
	NatRuleCollections TEXT,
	NetworkRuleCollections TEXT,
	ProvisioningState TEXT,
	IpConfigurationsText TEXT,
	ApplicationRuleCollectionsText TEXT,
	NatRuleCollectionsText TEXT,
	NetworkRuleCollectionsText TEXT,
	ResourceGroupName TEXT,
	Location TEXT,
	ResourceGuid TEXT,
	Type TEXT,
	Tag TEXT,
	TagsTable TEXT,
	Name TEXT,
	Etag TEXT,
	Id TEXT PRIMARY KEY
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
		$Query = 'INSERT INTO VirtualMachine (ResourceGroupName, Id, VmId, Name, Type, Location, LicenseType, Tags, AvailabilitySetReference, DiagnosticsProfile, Extensions, HardwareProfile, InstanceView, NetworkProfile, OSProfile, Plan, ProvisioningState, StorageProfile, Identity, Zones, FullyQualifiedDomainName, AdditionalCapabilities, RequestId, StatusCode) 
		VALUES ("' + $vm.ResourceGroupName
		$Query += '", "' + $vm.Id
		$Query += '", "' + $vm.VmId
		$Query += '", "' + $vm.Name
		$Query += '", "' + $vm.Type
		$Query += '", "' + $vm.Location
		$Query += '", "' + $vm.LicenseType
		$Query += '", "' + $vm.Tags
		$Query += '", "' + $vm.AvailabilitySetReference
		$Query += '", "' + $vm.DiagnosticsProfile
		$Query += '", "' + $vm.Extensions
		# HardwareProfile only contains VmSize
		$Query += '", "' + $vm.HardwareProfile.VmSize
		$Query += '", "' + $vm.InstanceView
		$Query += '", "' + $vm.NetworkProfile
		$Query += '", "' + $vm.OSProfile
		$Query += '", "' + $vm.Plan
		$Query += '", "' + $vm.ProvisioningState
		$Query += '", "' + $vm.StorageProfile
		$Query += '", "' + $vm.Identity
		$Query += '", "' + $vm.Zones
		$Query += '", "' + $vm.FullyQualifiedDomainName
		$Query += '", "' + $vm.AdditionalCapabilities
		$Query += '", "' + $vm.RequestId
		$Query += '", "' + $vm.StatusCode + '")'
	Invoke-SqliteQuery -DataSource $DataSource -Query $Query

	# OSProfile
	$Query = 'INSERT INTO OSProfile(	
	ComputerName,
	AdminUsername,
	AdminPassword,
	CustomData,
	WindowsConfiguration,
	LinuxConfiguration,
	Secrets,
	AllowExtensionOperations,
	Id)  
	VALUES ("' 
	$Query += (encode($vm.OSProfile.ComputerName)) 
	$Query += '", "' + (encode($vm.OSProfile.AdminUsername)) 
	$Query += '", "' + (encode($vm.OSProfile.AdminPassword))
	$Query += '", "' + (encode($vm.OSProfile.CustomData)) 
	$Query += '", "' + (encode($vm.WindowsConfiguration)) 
	$Query += '", "' + (encode($vm.LinuxConfiguration)) 
	$Query += '", "' + (arrayToString($vm.secrets)) 
	$Query += '", "' + (encode($vm.AllowExtensionOperations)) 
	$Query += '", "' + (encode($vm.Id)) 
	$Query += '")'
	# TODO go down the rest of the rabbit hole that is Linux and Windows Configs
	
	Invoke-SqliteQuery -DataSource $DataSource -Query $Query
	# TODO: StorageProfile
}

# Network Interfaces

$interfaces = Get-AzureRmNetworkInterface
foreach($interface in $interfaces){

	$Query = 'INSERT INTO NetworkInterface(	
	VirtualMachine, 
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
	VALUES ("' + (encode($interface.VirtualMachine)) 
	$Query += '", "' + (encode($interface.IpConfigurations)) 
	$Query += '", "' + (encode($interface.TapConfigurations)) 
	# Create Table for DnsSettings
	$Query += '", "' + (encode($interface.DnsSettings.DnsServersText)) 
	$Query += '", "' + (encode($interface.MacAddress)) 
	$Query += '", "' + (encode($interface.Primary)) 
	$Query += '", "' + (encode($interface.EnableAcceleratedNetworking)) 
	$Query += '", "' + (encode($interface.EnableIPForwarding)) 
	$Query += '", "' + (encode($interface.HostedWorkloads)) 
	$Query += '", "' + (encode($interface.NetworkSecurityGroup.Id)) 
	$Query += '", "' + (encode($interface.ProvisioningState)) 
	$Query += '", "' + (encode($interface.VirtualMachineText)) 
	$Query += '", "' + (encode($interface.IpConfigurationsText)) 
	$Query += '", "' + (encode($interface.TapConfigurationsText)) 
	$Query += '", "' + (encode($interface.DnsSettingsText)) 
	$Query += '", "' + (encode($interface.NetworkSecurityGroupText)) 
	$Query += '", "' + (encode($interface.ResourceGroupName)) 
	$Query += '", "' + (encode($interface.Location)) 
	$Query += '", "' + (encode($interface.ResourceGuid)) 
	$Query += '", "' + (encode($interface.Type)) 
	$Query += '", "' + (encode($interface.Tag)) 
	$Query += '", "' + (encode($interface.TagsTable)) 
	$Query += '", "' + (encode($interface.Name))
	$Query += '", "' + (encode($interface.Etag))
	$Query += '", "' + (encode($interface.Id)) + '")'
	
	Invoke-SqliteQuery -DataSource $DataSource -Query $Query

}

# Virtual Networks

$networks = Get-AzureRmVirtualNetwork
foreach($network in $networks){
	$Query = 'INSERT INTO VirtualNetwork (
		AddressSpace,
		DhcpOptions,
		ProvisioningState,
		EnableDdosProtection,
		EnableVmProtection,
		DdosProtectionPlan,
		AddressSpaceText,
		DhcpOptionsText,
		SubnetsText,
		VirtualNetworkPeeringsText,
		EnableDdosProtectionText,
		DdosProtectionPlanText,
		EnableVmProtectionText,
		ResourceGroupName,
		Location,
		ResourceGuid,
		Type,
		Tag,
		TagsTable,
		Name,
		Etag,
		Id
	)
	VALUES ("' + (arrayToString($network.AddressSpace.AddressPrefixes))
		# TODO verify Dhcp options
		$Query += '", "' + (encode($network.DhcpOptions.Id))
		# Foreign key will be added to Subnets for virtual networks
		# TODO Create Virtual Network Peerings Table
		$Query += '", "' + (encode($network.ProvisioningState))
		$Query += '", "' + (encode($network.EnableDdosProtection))
		$Query += '", "' + (encode($network.EnableVmProtection))
		$Query += '", "' + (encode($network.DdosProtectionPlan))
		$Query += '", "' + (encode($network.AddressSpaceText))
		$Query += '", "' + (encode($network.DhcpOptionsText))
		$Query += '", "' + (encode($network.SubnetsText))
		$Query += '", "' + (encode($network.VirtualNetworkPeeringsText))
		$Query += '", "' + (encode($network.EnableDdosProtectionText))
		$Query += '", "' + (encode($network.DdosProtectionText))
		$Query += '", "' + (encode($network.EnableVmProtectionText))
		$Query += '", "' + (encode($network.ResourceGroupName))
		$Query += '", "' + (encode($network.Location))
		$Query += '", "' + (encode($network.ResourceGuid))
		$Query += '", "' + (encode($network.Type))
		$Query += '", "' + (encode($network.Tag))
		$Query += '", "' + (encode($network.TagsTable))
		$Query += '", "' + (encode($network.Name))
		$Query += '", "' + (encode($network.Etag))
		$Query += '", "' + (encode($network.Id)) + '")'

	Invoke-SqliteQuery -DataSource $DataSource -Query $Query
	
	foreach($subnet in $network.Subnets){
		$Query = 'INSERT INTO Subnet (
	AddressPrefix,
	IpConfigurations,
	ServiceAssociationLinks,
	ResourceNavigationLinks,
	NetworkSecurityGroup,
	RouteTable,
	ServiceEndpoints,
	ServiceEndpointPolicies,
	Delegations,
	InterfaceEndpoints,
	ProvisioningState,
	IpConfigurationsText,
	ServiceAssociationLinksText,
	ResourceNavigationLinksText,
	NetworkSecurityGroupText,
	RouteTableText,
	ServiceEndpointText,
	ServiceEndpointPoliciesText,
	InterfaceEndpointsText,
	DelegationsText,
	Name,
	Etag,
	Id,
	VirtualNetwork
	)
	VALUES ("' + (arrayToString($subnet.AddressPrefix))
	# TODO verify IpConfigurations, ServiceAssociationLinks, ResourceNavigationLinks, RouteTable, ServiceEndpoints, etc
		$Query += '", "' + (encode($subnet.IpConfigurations))
		$Query += '", "' + (encode($subnet.ServiceAssociationLinks))
		$Query += '", "' + (encode($subnet.ResourceNavigationLinks))
		$Query += '", "' + (encode($subnet.NetworkSecurityGroup.Id))
		$Query += '", "' + (encode($subnet.RouteTable.Id))
		$Query += '", "' + (encode($subnet.ServiceEndpoints))
		$Query += '", "' + (encode($subnet.ServiceEndpointPolicies))
		$Query += '", "' + (encode($subnet.Delegations))
		$Query += '", "' + (encode($subnet.InterfaceEndpoints))
		$Query += '", "' + (encode($subnet.ProvisioningState))
		$Query += '", "' + (encode($subnet.IpConfigurationsText))
		$Query += '", "' + (encode($subnet.ServiceAssociationLinksText))
		$Query += '", "' + (encode($subnet.ResourceNavigationLinksText))
		$Query += '", "' + (encode($subnet.NetworkSecurityGroupText))
		$Query += '", "' + (encode($subnet.RouteTableText))
		$Query += '", "' + (encode($subnet.ServiceEndpointText))
		$Query += '", "' + (encode($subnet.ServiceEndpointPoliciesText))
		$Query += '", "' + (encode($subnet.InterfaceEndpointsText))
		$Query += '", "' + (encode($subnet.DelegationsText))
		$Query += '", "' + (encode($subnet.Name))
		$Query += '", "' + (encode($subnet.Etag))
		$Query += '", "' + (encode($subnet.Id))
		$Query += '", "' + (encode($network.Id)) + '")'

	Invoke-SqliteQuery -DataSource $DataSource -Query $Query
	}
}

# Web Apps
$apps = Get-AzureRmWebApp
foreach($app in $apps){
	$Query = 'INSERT INTO WebApp(	
		GitRemoteName,
		GitRemoteUri,
		GitRemoteUsername,
		GitRemotePassword,
		SnapshotInfo,
		State,
		HostNames,
		RepositorySiteName,
		UsageState,
		Enabled,
		EnabledHostNames,
		AvailabilityState,
		HostNameSslStates,
		ServerFarmId,
		Reserved,
		IsXenon,
		LastModifiedTimeUtc,
		SiteConfig,
		TrafficManagerHostNames,
		ScmSiteAlsoStopped,
		TargetSwapSlot,
		HostingEnvironmentProfile,
		ClientAffinityEnabled,
		ClientCertEnabled,
		HostNamesDisabled,
		OutboundIpAddresses,
		PossibleOutboundIpAddresses,
		ContainerSize,
		DailyMemoryTimeQuota,
		SuspendedTill,
		MaxNumberOfWorkers,
		CloningInfo,
		ResourceGroup,
		IsDefaultContainer,
		DefaultHostName,
		SlotSwapStatus,
		HttpsOnly,
		Identity,
		Id,
		Name,
		Kind,
		Location,
		Type,
		Tags)  
		VALUES ("' 
	$Query += (encode($app.GitRemoteName)) 
	$Query += '", "' + (encode($app.GitRemoteUri)) 
	$Query += '", "' + (encode($app.GitRemoteUserName)) 
	$Query += '", "' + (encode($app.GitRemotePassword)) 
	$Query += '", "' + (encode($app.SnapshotInfo)) 
	$Query += '", "' + (encode($app.State)) 
	$Query += '", "' + (arrayToString($app.HostNames)) 
	$Query += '", "' + (encode($app.RepositorySiteName)) 
	$Query += '", "' + (encode($app.UsageState)) 
	$Query += '", "' + (encode($app.Enabled)) 
	$Query += '", "' + (arrayToString($app.EnabledHostNames)) 
	$Query += '", "' + (encode($app.AvailabilityState)) 
	# TODO: make this table
	$Query += '", "' + (arrayToString($app.HostNameSslStates.SslState)) 
	$Query += '", "' + (encode($app.ServerFarmId)) 
	$Query += '", "' + (encode($app.Reserved)) 
	$Query += '", "' + (encode($app.IsXenon)) 
	$Query += '", "' + (encode($app.LastModifiedTimeUtc)) 
	$Query += '", "' + (encode($app.SiteConfig)) 
	$Query += '", "' + (encode($app.TrafficManagerHostNames)) 
	$Query += '", "' + (encode($app.ScmSiteAlsoStopped)) 
	$Query += '", "' + (encode($app.TargetSwapSlot)) 
	$Query += '", "' + (encode($app.HostingEnvironmentProfile)) 
	$Query += '", "' + (encode($app.ClientAffinityEnabled)) 
	$Query += '", "' + (encode($app.ClientCertEnabled)) 
	$Query += '", "' + (encode($app.HostNamesDisabled)) 
	$Query += '", "' + (arrayToString($app.OutboundIpAddresses)) 
	$Query += '", "' + (arrayToString($app.PossibleOutboundIpAddresses)) 
	$Query += '", "' + (encode($app.ContainerSize)) 
	$Query += '", "' + (encode($app.DailyMemoryTimeQuota)) 
	$Query += '", "' + (encode($app.SuspendedTill)) 
	$Query += '", "' + (encode($app.MaxNumberOfWorkers)) 
	$Query += '", "' + (encode($app.CloningInfo)) 
	$Query += '", "' + (encode($app.ResourceGroup)) 
	$Query += '", "' + (encode($app.IsDefaultContainer)) 
	$Query += '", "' + (encode($app.DefaultHostName)) 
	$Query += '", "' + (encode($app.SlotSwapStatus)) 
	$Query += '", "' + (encode($app.HttpsOnly)) 
	$Query += '", "' + (encode($app.Identity)) 
	$Query += '", "' + (encode($app.Id)) 
	$Query += '", "' + (encode($app.Name)) 
	$Query += '", "' + (encode($app.Kind)) 
	$Query += '", "' + (encode($app.Location)) 
	$Query += '", "' + (encode($app.Type)) 
	$Query += '", "' + (encode($app.Tags)) 
	$Query += '")'

	Invoke-SqliteQuery -DataSource $DataSource -Query $Query
}

# Application Gateway
$gateways = Get-AzureRmApplicationGateway
foreach($gate in $gateways){
	$Query = 'INSERT INTO ApplicationGateway(	
		Sku,
		SslPolicy,
		GatewayIPConfigurations,
		AuthenticationCertificates,
		SslCertificates,
		TrustedRootCertificates,
		FrontendIPConfigurations,
		FrontendPorts,
		Probes,
		BackendAddressPools,
		BackendHttpSettingsCollection,
		HttpListeners,
		UrlPathMaps,
		RequestRoutingRules,
		RedirectConfigurations,
		WebApplicationFirewallConfiguration,
		AutoscaleConfiguration,
		CustomErrorConfigurations,
		EnableHttp2,
		EnableFips,
		Zones,
		OperationalState,
		ProvisioningState,
		GatewayIpConfigurationsText,
		AuthenticationCertificatesText,
		SslCertificatesText,
		FrontendIpConfigurationsText,
		FrontendPortsText,
		BackendAddressPoolsText,
		BackendHttpSettingsCollectionText,
		HttpListenersText,
		RequestRoutingRulesText,
		ProbesText,
		UrlPathMapsText,
		ResourceGroupName,
		Location,
		ResourceGuid,
		Type,
		Tag,
		TagsTable,
		Name,
		Etag,
		Id)
		VALUES ("' 
		$Query += (encode($gate.Sku)) 
		$Query += '", "' + (encode($gate.SslPolicy)) 
		$Query += '", "' + (encode($gate.GatewayIPConfigurations)) 
		$Query += '", "' + (encode($gate.AuthenticationCertificates)) 
		$Query += '", "' + (encode($gate.SslCertificates)) 
		$Query += '", "' + (encode($gate.TrustedRootCertificates)) 
		$Query += '", "' + (encode($gate.FrontendIPConfigurations)) 
		# TODO: fix this
		$Query += '", "' + (encode($gate.FrontendPorts[0].Port)) 
		$Query += '", "' + (encode($gate.Probes)) 
		$Query += '", "' + (encode($gate.BackendAddressPools)) 
		$Query += '", "' + (encode($gate.BackendHttpSettingsCollection)) 
		$Query += '", "' + (encode($gate.HttpListeners)) 
		$Query += '", "' + (encode($gate.UrlPathMaps)) 
		$Query += '", "' + (encode($gate.RequestRoutingRules)) 
		$Query += '", "' + (encode($gate.RedirectConfigurations)) 
		$Query += '", "' + (encode($gate.WebApplicationFirewallConfiguration)) 
		$Query += '", "' + (encode($gate.AutoscaleConfiguration)) 
		$Query += '", "' + (encode($gate.CustomErrorConfigurations)) 
		$Query += '", "' + (encode($gate.EnableHttp2)) 
		$Query += '", "' + (encode($gate.EnableFips)) 
		$Query += '", "' + (encode($gate.Zones)) 
		$Query += '", "' + (encode($gate.OperationalState)) 
		$Query += '", "' + (encode($gate.ProvisioningState)) 
		$Query += '", "' + (encode($gate.GatewayIpConfigurationsText)) 
		$Query += '", "' + (encode($gate.AuthenticationCertificatesText)) 
		$Query += '", "' + (encode($gate.SslCertificatesText)) 
		$Query += '", "' + (encode($gate.FrontendIpConfigurationsText)) 
		$Query += '", "' + (encode($gate.FrontendPortsText)) 
		$Query += '", "' + (encode($gate.BackendAddressPoolsText)) 
		$Query += '", "' + (encode($gate.BackendHttpSettingsCollectionText)) 
		$Query += '", "' + (encode($gate.HttpListenersText)) 
		$Query += '", "' + (encode($gate.RequestRoutingRulesText)) 
		$Query += '", "' + (encode($gate.ProbesText)) 
		$Query += '", "' + (encode($gate.UrlPathMapsText)) 
		$Query += '", "' + (encode($gate.ResourceGroupName)) 
		$Query += '", "' + (encode($gate.Location)) 
		$Query += '", "' + (encode($gate.ResourceGuid)) 
		$Query += '", "' + (encode($gate.Type)) 
		$Query += '", "' + (encode($gate.Tag)) 
		$Query += '", "' + (encode($gate.TagsTable)) 
		$Query += '", "' + (encode($gate.Name)) 
		$Query += '", "' + (encode($gate.Etag)) 
		$Query += '", "' + (encode($gate.Id)) 		 
		$Query += '")'
		
		Invoke-SqliteQuery -DataSource $DataSource -Query $Query
}

# Firewalls
$firewalls = Get-AzureRmFirewall
foreach($wall in $firewalls){

$Query = 'INSERT INTO Firewall(	
	IpConfigurations,
	ApplicationRuleCollections,
	NatRuleCollections,
	NetworkRuleCollections,
	ProvisioningState,
	IpConfigurationsText,
	ApplicationRuleCollectionsText,
	NatRuleCollectionsText,
	NetworkRuleCollectionsText,
	ResourceGroupName,
	Location,
	ResourceGuid,
	Type,
	Tag,
	TagsTable,
	Name,
	Etag,
	Id)
	VALUES ("' 
	$Query += (encode($wall.IpConfigurations)) 
	$Query += '", "' + (encode($wall.ApplicationRuleCollections)) 
	$Query += '", "' + (encode($wall.NatRuleCollections))
	$Query += '", "' + (encode($wall.NetworkRuleCollections)) 
	$Query += '", "' + (encode($wall.ProvisioningState)) 
	$Query += '", "' + (encode($wall.IpConfigurationsText))   
	$Query += '", "' + (encode($wall.ApplicationRuleCollectionsText))   
	$Query += '", "' + (encode($wall.NatRuleCollectionsText))   
	$Query += '", "' + (encode($wall.NetworkRuleCollectionsText))   
	$Query += '", "' + (encode($wall.ResourceGroupName)) 
	$Query += '", "' + (encode($wall.Location)) 
	$Query += '", "' + (encode($wall.ResourceGuid)) 
	$Query += '", "' + (encode($wall.Type)) 
	$Query += '", "' + (encode($wall.Tag)) 
	$Query += '", "' + (encode($wall.TagsTable)) 
	$Query += '", "' + (encode($wall.Name)) 
	$Query += '", "' + (encode($wall.Etag)) 
	$Query += '", "' + (encode($wall.Id)) 
	$Query += '")'
	
	Invoke-SqliteQuery -DataSource $DataSource -Query $Query
}