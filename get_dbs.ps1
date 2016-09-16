# Check if the server has cluster services installed (1 = yes, 0 = no)
Function Check-HostHasClusterService {
	$SqlServer = $env:COMPUTERNAME
	$cluster=$null
	$hascluster=$null

	$cluster = Get-WmiObject win32_service -computerName $SqlServer | Where-Object {$_.Name -match "ClusSvc*" -and $_.PathName -match "clussvc.exe"} | Foreach{$_.Caption}

	if ($cluster) {
		$hascluster = 1
	}
	else	{
		$hascluster = 0
	}
	return $hascluster
}

# Check if this host is part of a cluster as a node
Function Check-HostIsNodeOfCluster {
	$SqlServer = $env:COMPUTERNAME
	$HostIsNodeOfCluster=$null

	if (Check-HostHasClusterService -eq 1) {
		if((Get-WmiObject -class MSCluster_Node -namespace "root\mscluster" -computername $SqlServer | Where-Object {$_.Name -eq $SqlServer} | Foreach {$_.name}) -eq $SqlServer) {
			$HostIsNodeOfCluster=1
		}
		else {
			$HostIsNodeOfCluster=0
		}
	}
	else {
		$HostIsNodeOfCluster=0
	}

	return $HostIsNodeOfCluster
}

# Get all SQL instance connections on the current host (even is SQL is stopped)
Function Get-SQLConnections {
	$SqlServer = $env:COMPUTERNAME
	$isclustered=Check-HostIsNodeOfCluster
	$sqlvirtualname=@()
	$clusteredinstance=$null
	$definstancecluster=$null
	$clusteredinstances=@()
	$dummyarray=@()
	$instances=New-Object System.Collections.ArrayList
	$instance=$null
	$localinstance=@()
	$named=$null
	$namedinstances=@()
	$sqlconnections=@()
	$sqlconnection=$null
	$allinstances=@()

	# If the host is a node of a cluster 
	if ($isclustered -eq 1) {
	    # Get all clustered instances
		$sqlvirtualname=Get-WmiObject -class "MSCluster_Resource" -namespace "root\mscluster" -computername $SqlServer `
						| where {$_.type -eq "SQL Server"} `
						| select @{n='ServerInstance';e={("{0}\{1}" -f $_.PrivateProperties.VirtualServerName,$_.PrivateProperties.InstanceName).TrimEnd('\')}}
		
		# If there are actual clustered instances
		if ($sqlvirtualname) {
			# Let's process them one by one
			foreach ($clusteredinstance in $sqlvirtualname) {
				# If it is a default instance installed on the cluster, remove "MSSQLSERVER" from the connection name
				if ($clusteredinstance.ServerInstance -like "*\MSSQLSERVER*") {
					$definstancecluster=$clusteredinstance.ServerInstance.ToString().Replace("\MSSQLSERVER","")
					$clusteredinstances+=$definstancecluster
				}
				# If it is not a default instance, the value can be used as a connection name
				else {
					$clusteredinstances+=$clusteredinstance.ServerInstance
				}
			}
		}
	}

	# It is possible to have non-clustered instances on a cluster node, so let's find them

	# F*cking powershell, need to create a dummy array so i can add it to an arraylist
	$dummyarray = @(Get-WmiObject win32_service -computerName $SqlServer `
					| Where-Object {$_.Name -match "mssql*" -and $_.PathName -match "sqlservr.exe"} `
					| Foreach {$_.Name})
	
	# Add all items from dummy array to the arraylist, so we can do some manipulation
	if ($dummyarray) {
		$instances.AddRange($dummyarray) | Out-Null
	}
	
	# Remove the clustered instances from the total list of instances on this host, else we will have duplicate values
	if ($clusteredinstances) {
		for ($i = 0; $i -lt $instances.Count; $i++) {
			foreach ($clusteredinstance in $sqlvirtualname) {
				$instance=$instances[$i]
				
				# Match takes place based on instance name without the host name
				if ($clusteredinstance.ServerInstance -match ($instance | %{$_.split("$")[-1]} | %{$_.trimStart("(")} | %{$_.trimEnd(")")})) {
					# If we have a match, remove this particular instance from arraylist
					$instances.Remove($instance)
				}
			}
		}
	}
	
	# If there are non-clustered instances, let's also process them
	if ($instances) {
		foreach ($instance in $instances) {
			# If it is a default instance use the hostname as the connection name
			if ($instance -eq "SQL Server (MSSQLSERVER)" -or $instance -eq "MSSQLSERVER") {
				$localinstance += $SqlServer
			}
			# Build the connection name for named instances
			else {
			  $named = $instance | %{$_.split("$")[-1]} | %{$_.trimStart("(")} | %{$_.trimEnd(")")}
			  $namedinstances += $Sqlserver + "\" + $named
			}
		}
	}
	
	# If there is no value for local or named instance return NULL
	if (!$localinstance -and !$namedinstances) {
		$sqlconnections=@()
	}
	# If there is only a value for $namedinstance then only return the namedinstances
	elseif (!$localinstance -and $namedinstances) {
		$sqlconnections=$namedinstances
	}
	# If there is only a value for $localinstance then only return the localinstance
	elseif($localinstance -and !$namedinstances) {
		$sqlconnections=$localinstance
	}
	# For anything else, return the localinstance and namedinstances
	else {
		$sqlconnections=$localinstance + $namedinstances
	}

	# Return all instances found
	$allinstances=$sqlconnections + $clusteredinstances

	return $allinstances
}

# Get all SQL instance connections on the current host (only when SQL is running)
Function Get-SQLActiveConnections {
	$SqlServer = $env:COMPUTERNAME
	$sqlvirtualname=@()
	$localactiveinstances=@()
	$localinstance=@()
	$sqlconnections=@()
	$sqlconnection=$null
	$dummyarray=@()
	$instances=New-Object System.Collections.ArrayList
	$namedinstances=@()
	$clusteredinstances=@()
	$definstancecluster=$null
	$named=$null
	$instance=$null
	$allinstances=@()
	$isclustered=Check-HostIsNodeOfCluster

	# If the host is a node of a cluster
	if ($isclustered -eq 1) {
		$sqlvirtualname= gwmi -class "MSCluster_Resource" -namespace "root\mscluster" -computername $SqlServer  | where {$_.type -eq "SQL Server" -and $_.State -eq 2} | `
			select @{n='ServerInstance';e={("{0}\{1}" -f $_.PrivateProperties.VirtualServerName,$_.PrivateProperties.InstanceName).TrimEnd('\')}}, `
			@{n='Node';e={$(gwmi -namespace "root\mscluster" -computerName $SqlServer -query "ASSOCIATORS OF {MSCluster_Resource.Name='$($_.Name)'} WHERE AssocClass = MSCluster_NodeToActiveResource" | Select -ExpandProperty Name)}}

		$localactivesqlvirtualname=$sqlvirtualname | Where-Object {$_.Node -eq $SqlServer} | Select ServerInstance

		if ($localactivesqlvirtualname) {
		 	foreach ($clusteredinstance in $localactivesqlvirtualname) {
				# If it is a default instance installed on the cluster, remove "MSSQLSERVER" from the connection name
				if ($clusteredinstance.ServerInstance -like "*\MSSQLSERVER*") {
					$definstancecluster=$clusteredinstance.ServerInstance.ToString().Replace("\MSSQLSERVER","")
					$clusteredinstances+=$definstancecluster
				}
				# If it is not a default instance, the value can be used as a connection name
				else {
					$clusteredinstances+=$clusteredinstance.ServerInstance
				}
			}
		}
	}
	
	# If the host is not a node of a cluster

    $dummyarray = @(Get-WmiObject win32_service -computerName $SqlServer `
							| Where-Object {$_.Name -match "mssql*" -and $_.PathName -match "sqlservr.exe" -and $_.State -eq "Running"} `
							| Foreach {$_.Name})

	# Add all items from dummy array to the arraylist, so we can do some manipulation
	if ($dummyarray) {
		$instances.AddRange($dummyarray) | Out-Null
	}
	
	# Remove the clustered instances from the total list of instances on this host, else we will have duplicate values
	if ($clusteredinstances) {
		for ($i = 0; $i -lt $instances.Count; $i++) {
			foreach ($clusteredinstance in $localactivesqlvirtualname) {
				$instance=$instances[$i]
				
				# Match takes place based on instance name without the host name
				if ($clusteredinstance.ServerInstance -match ($instance | %{$_.split("$")[-1]} | %{$_.trimStart("(")} | %{$_.trimEnd(")")})) {
					# If we have a match, remove this particular instance from arraylist
					$instances.Remove($instance)
				}
			}
		}
	}
	
	# If there are non-clustered instances, let's also process them
	if ($instances) {
		foreach ($instance in $instances) {
			# If it is a default instance use the hostname as the connection name
			if ($instance -eq "SQL Server (MSSQLSERVER)" -or $instance -eq "MSSQLSERVER") {
				$localinstance += $SqlServer
			}
			# Build the connection name for named instances
			else {
			  $named = $instance | %{$_.split("$")[-1]} | %{$_.trimStart("(")} | %{$_.trimEnd(")")}
			  $namedinstances += $Sqlserver + "\" + $named
			}
		}
	}
	
	# If there is no value for local or named instance return NULL
	if (!$localinstance -and !$namedinstances) {
		$sqlconnections=@()
	}
	# If there is only a value for $namedinstance then only return the namedinstances
	elseif (!$localinstance -and $namedinstances) {
		$sqlconnections=$namedinstances
	}
	# If there is onlu a value for $localinstance then only return the localinstance
	elseif($localinstance -and !$namedinstances) {
		$sqlconnections=$localinstance
	}
	# For anything else, return the localinstance and namedinstances
	else {
		$sqlconnections=$localinstance + $namedinstances
	}
    
	# Return all instances found
	$allinstances=$sqlconnections + $clusteredinstances
	
	return $allinstances
}

# Standard function for running queries to SQL Server, it will return a datatable object
Function RunSQLStatement {
	Param
	(
		[parameter(Position=1,mandatory=$true)] [string]$SqlServer,
		[parameter(Position=2,mandatory=$true)] [string]$SqlCatalog,
		[parameter(Position=3,mandatory=$true)] [string]$SqlQuery
	)

	TRY {
		$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
		$SqlConnection.ConnectionString = "Server = $SqlServer; Database = $SqlCatalog; Integrated Security = True; Application Name = .NET SQLClient Data Provider - Nagios Checks"
		$SqlConnection.Open()

		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
		$SqlCmd.CommandText = $SqlQuery
		$SqlCmd.CommandTimeout = 60
		$SqlCmd.Connection = $SqlConnection

		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd

		$DataSet = New-Object System.Data.DataSet
		$SqlAdapter.Fill($DataSet) | Out-Null

		$SqlConnection.Close()

		$ResultSet=$DataSet.Tables[0]
	}
	CATCH {
		$ErrorMessage=$Error[0].Exception ;
		$NagiosMessage="Could not query the instance: " + $SqlServer + " Errormessage: " + $ErrorMessage.Message
		Write-Host $NagiosMessage	
		exit ($global:EC_WARNING)
	}

	return $ResultSet
}