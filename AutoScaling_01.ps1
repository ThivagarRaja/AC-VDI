######## Variables ##########

# View Verbose data
# Set to "Continue" to see Verbose data
# set to "SilentlyContinue" to hide Verbose data
$VerbosePreference = "Continue"

# Server start threshold
# Number of available sessions to trigger a server start or shutdown
$serverStartThreshold = 30

# Peak time and Threshold settings
# Set usePeak to $true to enable peak time, $false to disable
# Set useBreadthFirstDuringPeak to $true to change load balancing to Breadth-First and start all Session Hosts
#   This setting will not change the max session limit
#   useBreadthFirstDuringPeak requires $usePeak set to $true
# Set the Peak Threshold, the spare capacity during peak hours (not required if useBreadthFirstDuringPeak is used)
# Set the Start and Stop Peak Time, use a 24 hour format of Hour:Minute:Seconds (08:30:00)
# Set the time zone to use, use "Get-TimeZone -ListAvailable" to list ID's
$usePeak = $true
$useBreadthFirstDuringPeak = $true
$peakServerStartThreshold = 4
$startPeakTime = '08:00:00'
$endPeakTime = '18:00:00'
$timeZone = "Eastern Daylight Time"
$peakDay = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'

# Host Pool Name
$hostPoolName = 'HP'

# Session Host Resource Group
# Session Hosts and Host Pools can exist in different Resource Groups, but are commonly the same
# Host Pool Resource Group and the resource group of the Session host VM's.
$hostPoolRg = 'Test-South'
$sessionHostVmRg= 'Test-South'

############## Functions ####################

Function Start-SessionHost {
    param (
        $sessionHosts,
        $hostsToStart
    )

    # Number of off session hosts accepting connections
    $offSessionHosts = $sessionHosts | Where-Object { $_.Status -eq "Unavailable" -or $_.Status -eq "Shutdown" }
    $offSessionHostsCount = $offSessionHosts.count
    Write-Verbose "Off Session Hosts $offSessionHostsCount"
    Write-Verbose ($offSessionHosts | Out-String)

    if ($offSessionHosts.Count -eq 0 ) {
        Write-Error "Start threshold met, but there are no hosts available to start"
    }
    else {
        if ($hostsToStart -gt $offSessionHostsCount) {
            $hostsToStart = $offSessionHostsCount
        }
        Write-Verbose "Conditions met to start a host"
        $counter = 0
        while ($counter -lt $hostsToStart) {
            $startServerName = ($offSessionHosts | Select-Object -Index $counter).name
            Write-Verbose "Server to start $startServerName"
            try {
                # Start the VM
                $vmName = ($startServerName -split { $_ -eq '.' -or $_ -eq '/' })[1]
                Start-AzVM -ErrorAction Stop -ResourceGroupName $sessionHostVmRg -Name $vmName -NoWait
            }
            catch {
                $ErrorMessage = $_.Exception.message
                Write-Error ("Error starting the session host: " + $ErrorMessage)
                Break
            }
            $counter++
        }
    }
}
function Stop-SessionHost {
    param (
        $SessionHosts,
        $hostsToStop
    )
    # Get computers running with no users
    $emptyHosts = $sessionHosts | Where-Object { $_.Session -eq 0 -and $_.Status -eq 'Available' }
    $emptyHostsCount = $emptyHosts.count
    Write-Verbose "Evaluating servers to shut down"

    if ($emptyHostsCount -eq 0) {
        Write-error "No hosts available to shut down"
    }
    else { 
        if ($hostsToStop -ge $emptyHostsCount) {
            $hostsToStop = $emptyHostsCount
        }
        Write-Verbose "Conditions met to stop a host"
        $counter = 0
        while ($counter -lt $hostsToStop) {
            $shutServerName = ($emptyHosts | Select-Object -Index $counter).Name 
            Write-Verbose "Shutting down server $shutServerName"
            try {
                # Stop the VM
                $vmName = ($shutServerName -split { $_ -eq '.' -or $_ -eq '/' })[1]
                Stop-AzVM -ErrorAction Stop -ResourceGroupName $sessionHostVmRg -Name $vmName -Force -NoWait
            }
            catch {
                $ErrorMessage = $_.Exception.message
                Write-Error ("Error stopping the VM: " + $ErrorMessage)
                Break
            }
            $counter++
        }
    }
}  

########## Script Execution ##########

# Get Host Pool 
try {
    $hostPool = Get-AzWvdHostPool -ResourceGroupName $hostPoolRg -Name $hostPoolName 
    Write-Verbose "HostPool:"
    Write-Verbose $hostPool.Name
}
catch {
    $ErrorMessage = $_.Exception.message
    Write-Error ("Error getting host pool details: " + $ErrorMessage)
    Break
}

# Check if peak time and adjust threshold
# Warning! will not adjust for DST
$isPeakTime = $false
if ($usePeak -eq $true) {
    # Get the current date adjusted by the time zone
    $utcDate = ((get-date).ToUniversalTime())
    $tZ = Get-TimeZone $timeZone
    $date = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcDate, $tZ)
    write-verbose "Date and Time"
    write-verbose $date
    # Get the current day of the week adjusted for the time zone
    $utcOffset = $tz.BaseUtcOffset.TotalHours
    $dateDay = (((get-date).ToUniversalTime()).AddHours($utcOffset)).dayofweek
    Write-Verbose $dateDay
    # Slice and dice to get the peak start and end time adjusted for the time zone
    $startPeakTimeSplit = $startPeakTime.Split(":")
    $startPeakTime = (get-date $date -Hour $startPeakTimeSplit[0] -minute $startPeakTimeSplit[1] -second $startPeakTimeSplit[2])
    $endPeakTimeSplit = $endPeakTime.Split(":")
    $endPeakTime = (get-date $date -Hour $endPeakTimeSplit[0] -minute $endPeakTimeSplit[1] -second $endPeakTimeSplit[2])
    # Adjust threshold if in the peak time window
    if ($date -gt $startPeakTime -and $date -lt $endPeakTime -and $dateDay -in $peakDay) {
        Write-Verbose "Adjusting threshold for peak hours"
        $serverStartThreshold = $peakServerStartThreshold
        Write-Verbose "Setting Peak Time to True"
        $isPeakTime = $true
    } 
}

# Find the total number of session hosts
# Exclude servers in drain mode and do not allow new connections
try {
    $sessionHosts = Get-AzWvdSessionHost -ResourceGroupName $hostPoolRg -HostPoolName $hostPoolName | Where-Object { $_.AllowNewSession -eq $true }
    # Get current active user sessions
    $currentSessions = 0
    foreach ($sessionHost in $sessionHosts) {
        $count = $sessionHost.session
        $currentSessions += $count
    }
    Write-Verbose "CurrentSessions"
    Write-Verbose $currentSessions
}
catch {
    $ErrorMessage = $_.Exception.message
    Write-Error ("Error getting session hosts details: " + $ErrorMessage)
    Break
}

# Number of running and available session hosts
# Host shut down are excluded
$runningSessionHosts = $sessionHosts | Where-Object { $_.Status -eq "Available" }
$runningSessionHostsCount = $runningSessionHosts.count
Write-Verbose "Running Session Host $runningSessionHostsCount"
Write-Verbose ($runningSessionHosts | Out-string)

#region Breadth-First During Peak
# This section only runs if using peak time and if $useBreadthFirstDuringPeak is set to $true
if (($isPeakTime -eq $true) -and ($useBreadthFirstDuringPeak -eq $true)) {
    # Set the host pool to BreadthFirst
    if ($hostPool.LoadBalancerType -eq "DepthFirst") {
        try {
            Write-Verbose "Setting host pool load balancing algorithm to BreadthFirst"
            Update-AzWvdHostPool -Name $hostPoolName -ResourceGroupName $hostPoolRg -LoadBalancerType 'BreadthFirst' 
        }
        catch {
            $ErrorMessage = $_.Exception.message
            Write-Error ("Error setting the host pool to BreadthFirst: " + $ErrorMessage)
            Break
        }
    }
    #Start all available session hosts
    if ($sessionHosts.Count -gt $runningSessionHostsCount) {
        Write-Verbose "Starting all available session hosts"
        $hostsToStart = $sessionHosts.Count - $runningSessionHostsCount
        Write-Verbose "Hosts to start: $hostsToStart"
        Start-SessionHost -sessionHosts $sessionHosts -hostsToStart $hostsToStart
    }
}
#endregion

#region Depth-First During Peak
if (($isPeakTime -eq $false) -or ($useBreadthFirstDuringPeak -eq $false)) {
    # Verify load balancing is set to Depth-first, update if not
    if ($hostPool.LoadBalancerType -ne "DepthFirst") {
        try {
            Write-Verbose "Setting host pool load balancing algorithm to DepthFirst"
            Update-AzWvdHostPool -Name $hostPoolName -ResourceGroupName $hostPoolRg -LoadBalancerType 'DepthFirst' 
        }
        catch {
            $ErrorMessage = $_.Exception.message
            Write-Error ("Error setting the host pool to BreadthFirst: " + $ErrorMessage)
            Break
        }
    }

    # Get the Max Session Limit on the host pool
    # This is the total number of sessions per session host
    $maxSession = $hostPool.MaxSessionLimit
    Write-Verbose "MaxSession:"
    Write-Verbose $maxSession

    # Target number of servers required running based on active sessions, Threshold and maximum sessions per host
    $sessionHostTarget = [math]::Ceiling((($currentSessions + $serverStartThreshold) / $maxSession))

    if ($runningSessionHostsCount -lt $sessionHostTarget) {
        Write-Verbose "Running session host count $runningSessionHostsCount is less than session host target count $sessionHostTarget, run start function"
        $hostsToStart = ($sessionHostTarget - $runningSessionHostsCount)
        Start-SessionHost -sessionHosts $sessionHosts -hostsToStart $hostsToStart
    }
    elseif ($runningSessionHostsCount -gt $sessionHostTarget) {
        Write-Verbose "Running session hosts count $runningSessionHostsCount is greater than session host target count $sessionHostTarget, run stop function"
        $hostsToStop = ($runningSessionHostsCount - $sessionHostTarget)
        Stop-SessionHost -SessionHosts $sessionHosts -hostsToStop $hostsToStop
    }
    else {
        Write-Verbose "Running session host count $runningSessionHostsCount matches session host target count $sessionHostTarget, doing nothing"
    }
}
#endregion