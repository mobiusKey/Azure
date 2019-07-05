# Install-Module PSSQLite
# Import-Module
# Create VM Table


$DataSource = Get-Location
$DataSource = $DataSource.Path + "\Test.SQLite"

$Query = "CREATE TABLE VMS (
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

$vms = Get-AzureRmVM

foreach($vm in $vms){
		$Query = 'INSERT INTO VMS (ResourceGroupName, Id, VmId, Name, Type, Location, LicenseType, Tags, AvailabilitySetReference, DiagnosticsProfile, Extensions, HardwareProfile, InstanceView, NetworkProfile, OSProfile, AdminUsername, AdminPassword, CustomData, WindowsConfiguration, Secrets, Plan, ProvisioningState, StorageProfile, Offer, Sku, Version, DisplayHint, Identity, Zones, FullyQualifiedDomainName, AdditionalCapabilities, RequestId, StatusCode) 
		VALUES ("' + $vm.ResourceGroupName + '", "' + $vm.Id+ '", "' + $vm.VmId+ '", "' + $vm.Name+ '", "' + $vm.Type+ '", "' + $vm.Location+ '", "' + $vm.LicenseType+ '", "' + $vm.Tags+ '", "' + $vm.AvailabilitySetReference+ '", "' + $vm.DiagnosticsProfile+ '", "' + $vm.Extensions+ '", "' + $vm.HardwareProfile+ '", "' + $vm.InstanceView+ '", "' + $vm.NetworkProfile+ '", "' + $vm.OSProfile+ '", "' + $vm.AdminUsername+ '", "' + $vm.AdminPassword+ '", "' + $vm.CustomData+ '", "' + $vm.WindowsConfiguration+ '", "' + $vm.Secrets+ '", "' + $vm.Plan+ '", "' + $vm.ProvisioningState+ '", "' + $vm.StorageProfile+ '", "' + $vm.Offer+ '", "' + $vm.Sku+ '", "' + $vm.Version+ '", "' + $vm.DisplayHint+ '", "' + $vm.Identity+ '", "' + $vm.Zones+ '", "' + $vm.FullyQualifiedDomainName+ '", "' + $vm.AdditionalCapabilities+ '", "' + $vm.RequestId+ '", "' + $vm.StatusCode + '")'
	Invoke-SqliteQuery -DataSource $DataSource -Query $Query
	}