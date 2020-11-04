# !TODO Add Documentation Synopsis
# !TODO Add option to specify host list

[CmdletBinding()]
param (
    [TypeName]$ParameterName
)

$Computers = Get-ADComputer -filter { (enabled -eq "true") } | Select-Object -ExpandProperty Name | Sort-Object

Write-Verbose $Computers.Count

$pool = [RunspaceFactory]::CreateRunspacePool(1, 500)
$pool.ApartmentState = "MTA"
$pool.open()
$runspaces = @()

# This script block defines what each space is going to be doing
$scriptblock = {
    param ([string]$Computer);
    # Check to see of the computer is running
    if (test-connection -computername $Computer -BufferSize 16 -quiet -count 1) {
        Try {
            # If the computer is running find the process we want to close and close it
            # !TODO Change process to Parameter
            invoke-command $Computer { get-process gssmenu | stop-process -force  } -ErrorAction Stop
        }
        catch {
            Write-Verbose "Can not connect to Host $Computer"
        }
    }
}

# Create a runspace for each computer and start running it
Foreach ($Computer in $Computers) {
    $runspace = [powershell]::create()
    $null = $runspace.addscript($scriptblock)
    $null = $runspace.addargument($Computer)
    $runspace.runspacepool = $pool
    $runspaces += [pscustomobject]@{pipe = $runspace; Status = $runspace.begininvoke() }
}

# Initialize variables for the progress bar
$currentcount = 0;
$totalcount = ($runspaces | measure-object).count;
Write-Host $totalcount

# Check to see if runspace has finished and update progress bar as they do
while ($runspaces.status -ne $null) {
    $completed = $runspaces | Where-Object { $_.status.iscompleted -eq $true };
    # Update progress bar
    $currentcount = $currentcount + ($completed | measure-object).count;
    write-progress -activity "Checking Computers..." -percentcomplete (([int]$currentcount / [int]$totalcount) * 100);

    # Clear completed runspaces
    foreach ($runspace in $completed) {
        $runspace.pipe.endinvoke($runspace.status)
        $runspace.status = $null
    }
}

# Clean-up Runspace Pool
$pool.close();
$pool.dispose();