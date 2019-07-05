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


Write-Output("==== SQL Servers ====")
$servers = Get-AzureRmSqlServer
$server

$excel = New-Object -ComObject excel.application
$excel.visible = $True
$workbook = $excel.WorkBooks.Add()
$workbook.WorkSheets.Item(1).Name = "SQL Servers"
$worksheet = $workbook.worksheets.Item(1)
$row = 1
$column = 1

foreach($server in $servers){
	$server.PSObject.Properties | ForEach-Object {
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
	
foreach($server in $servers){
	$server.PSObject.Properties | ForEach-Object {
		$value = $_.Value
		Write-Output($_.Name + ": " + $value)
		switch($_.Name){
			"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}
			"ServerName" {$excel.cells.item($row,$column) = $value; break}
			"Location" {$excel.cells.item($row,$column) = $value; break}
			"SqlAdministratorLogin" {$excel.cells.item($row,$column) = $value; break}
			"SqlAdministratorPassword" {$excel.cells.item($row,$column) = $value; break}
			"ServerVersion" {$excel.cells.item($row,$column) = $value; break}
			"Tags" {$excel.cells.item($row,$column) = $value.Name; break}
			"Identity" {$excel.cells.item($row,$column) = $value; break}
			"FullyQualifiedDomainName" {$excel.cells.item($row,$column) = $value; break}
			"ResourceId"  {$excel.cells.item($row,$column) = $value; break}

			default {$excel.cells.item($row,$column) = "Error"}
		}
		$column += 1
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1
}
	
# SQL Databases
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Databases"
$column = 1
$row = 1	

foreach($server in $servers){
	$databases = Get-AzureRmSqlDatabase -ResourceGroupName $server.ResourceGroupName -ServerName $server.ServerName
	
	foreach($db in $databases){
		$db.PSObject.Properties | ForEach-Object {
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
	
	foreach($db in $databases){
		$db.PSObject.Properties | ForEach-Object {
			$value = $_.Value
			Write-Output($_.Name + ": " + $value)
			switch($_.Name){
				"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}
				"ServerName" {$excel.cells.item($row,$column) = $value; break}
				"DatabaseName" {$excel.cells.item($row,$column) = $value; break}
				"Location" {$excel.cells.item($row,$column) = $value; break}
				"DatabaseId" {$excel.cells.item($row,$column) = $value.Guid; break}
				"Edition" {$excel.cells.item($row,$column) = $value; break}
				"CollationName" {$excel.cells.item($row,$column) = $value; break}
				"CatalogCollation" {$excel.cells.item($row,$column) = $value; break}
				"MaxSizeBytes" {$excel.cells.item($row,$column) = $value; break}
				"Status" {$excel.cells.item($row,$column) = $value; break}
				"CreationDate" {$excel.cells.item($row,$column) = $value; break}
				"CurrentServiceObjectiveId" {$excel.cells.item($row,$column) = $value.Guid; break}
				"CurrentServiceObjectiveName" {$excel.cells.item($row,$column) = $value; break}
				"RequestedServiceObjectiveName" {$excel.cells.item($row,$column) = $value; break}
				"RequestedServiceObjectiveId" {$excel.cells.item($row,$column) = $value; break}
				"ElasticPoolName" {$excel.cells.item($row,$column) = $value; break}
				"EarliestRestoreDate" {$excel.cells.item($row,$column) = $value; break}
				"Tags" {$excel.cells.item($row,$column) = $value.Name; break}
				"ResourceId" {$excel.cells.item($row,$column) = $value; break}
				"CreateMode" {$excel.cells.item($row,$column) = $value; break}
				"ReadScale" {$excel.cells.item($row,$column) = $value; break}
				"ZoneRedundant" {$excel.cells.item($row,$column) = $value; break}
				"Capacity" {$excel.cells.item($row,$column) = $value; break}
				"Family" {$excel.cells.item($row,$column) = $value; break}
				"SkuName" {$excel.cells.item($row,$column) = $value; break}
				"LicenseType" {$excel.cells.item($row,$column) = $value; break}

				default {$excel.cells.item($row,$column) = "Error"}
			}
		$column += 1
		}
		$column = 1
		$worksheet.Rows($row).RowHeight = 15
		$row += 1
	}
}	
	
# Storage Accounts
$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Storage Accounts"
$column = 1
$row = 1	
$SAs = Get-AzureRmStorageAccount

foreach($sa in $SAs){
	$sa.PSObject.Properties | ForEach-Object {
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

foreach($sa in $SAs){
	$sa.PSObject.Properties | ForEach-Object {
		$value = $_.Value
		Write-Output($_.Name + ": " + $value)
		switch($_.Name){
			"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}
			"StorageAccountName" {$excel.cells.item($row,$column) = $value; break}
			"Id" {$excel.cells.item($row,$column) = $value; break}
			"Location" {$excel.cells.item($row,$column) = $value; break}
			"Sku" {$excel.cells.item($row,$column) = $value; break}
			"Kind" {$excel.cells.item($row,$column) = $value; break}
			"Encryption" {$excel.cells.item($row,$column) = $value; break}
			"AccessTier" {$excel.cells.item($row,$column) = $value; break}
			"CreationTime" {$excel.cells.item($row,$column) = $value; break}
			"CustomDomain" {$excel.cells.item($row,$column) = $value; break}
			"Identity" {$excel.cells.item($row,$column) = $value; break}
			"LastGeoFailoverTime" {$excel.cells.item($row,$column) = $value; break}
			"PrimaryEndpoints" {$excel.cells.item($row,$column) = $value; break}
			"PrimaryLocation" {$excel.cells.item($row,$column) = $value; break}
			"ProvisioningState" {$excel.cells.item($row,$column) = $value; break}
			"SecondaryEndpoints" {$excel.cells.item($row,$column) = $value; break}
			"SecondaryLocation" {$excel.cells.item($row,$column) = $value; break}
			"StatusOfPrimary" {$excel.cells.item($row,$column) = $value; break}
			"StatusOfSecondary" {$excel.cells.item($row,$column) = $value; break}
			"Tags" {$excel.cells.item($row,$column) = $value.Name; break}
			"EnableHttpsTrafficOnly" {$excel.cells.item($row,$column) = $value; break}
			"NetworkRuleSet" {$excel.cells.item($row,$column) = $value.DefaultAction; break}
			"Context" {$excel.cells.item($row,$column) = $value; break}
			"ExtendedProperties" {$excel.cells.item($row,$column) = $value; break}


			default {$excel.cells.item($row,$column) = "Error"}
		}
	$column += 1
	}
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1
}
	
	# Auditing policies
	Write-Output("======== Auditing ==========")
	
	Write-Output("=== Servers ===")
	$servers
	Write-Output("=== Databases ===")
	$databases
	$worksheet = $workbook.worksheets.add($worksheet)
$worksheet.Name = "Auditing Policies"
$column = 1
$row = 1	

foreach($server in $servers){
	$databases = Get-AzureRmSqlDatabase -ResourceGroupName $server.ResourceGroupName -ServerName $server.ServerName
	foreach($db in $databases){
		$policy = Get-AzureRmSqlDatabaseAuditing -ResourceGroupName $server.ResourceGroupName -ServerName $server.ServerName -DatabaseName $db.DatabaseName
		foreach($pol in $policy){
		$policy.PSObject.Properties | ForEach-Object {
			$excel.cells.item($row,$column) = $_.Name
			Write-Output($_.Name + ": " + $_.Value)
			$column += 1
		}
		
		$column = 1
		break
		}
	}
	}
	
	$column = 1
	$worksheet.Rows($row).RowHeight = 15
	$row += 1
	
<# 	foreach($db in $databases){
		$db.PSObject.Properties | ForEach-Object {
			$value = $_.Value
			Write-Output($_.Name + ": " + $value)
			switch($_.Name){
				"ResourceGroupName" {$excel.cells.item($row,$column) = $value; break}
				"ServerName" {$excel.cells.item($row,$column) = $value; break}
				"DatabaseName" {$excel.cells.item($row,$column) = $value; break}
				"Location" {$excel.cells.item($row,$column) = $value; break}
				"DatabaseId" {$excel.cells.item($row,$column) = $value.Guid; break}
				"Edition" {$excel.cells.item($row,$column) = $value; break}
				"CollationName" {$excel.cells.item($row,$column) = $value; break}
				"CatalogCollation" {$excel.cells.item($row,$column) = $value; break}
				"MaxSizeBytes" {$excel.cells.item($row,$column) = $value; break}
				"Status" {$excel.cells.item($row,$column) = $value; break}
				"CreationDate" {$excel.cells.item($row,$column) = $value; break}
				"CurrentServiceObjectiveId" {$excel.cells.item($row,$column) = $value.Guid; break}
				"CurrentServiceObjectiveName" {$excel.cells.item($row,$column) = $value; break}
				"RequestedServiceObjectiveName" {$excel.cells.item($row,$column) = $value; break}
				"RequestedServiceObjectiveId" {$excel.cells.item($row,$column) = $value; break}
				"ElasticPoolName" {$excel.cells.item($row,$column) = $value; break}
				"EarliestRestoreDate" {$excel.cells.item($row,$column) = $value; break}
				"Tags" {$excel.cells.item($row,$column) = $value.Name; break}
				"ResourceId" {$excel.cells.item($row,$column) = $value; break}
				"CreateMode" {$excel.cells.item($row,$column) = $value; break}
				"ReadScale" {$excel.cells.item($row,$column) = $value; break}
				"ZoneRedundant" {$excel.cells.item($row,$column) = $value; break}
				"Capacity" {$excel.cells.item($row,$column) = $value; break}
				"Family" {$excel.cells.item($row,$column) = $value; break}
				"SkuName" {$excel.cells.item($row,$column) = $value; break}
				"LicenseType" {$excel.cells.item($row,$column) = $value; break}

				default {$excel.cells.item($row,$column) = "Error"}
			}
		$column += 1
		}
		$column = 1
		$worksheet.Rows($row).RowHeight = 15
		$row += 1
	}
}	 #>
	
	# Threat Detection 
	
	
Write-Output("Done") 
