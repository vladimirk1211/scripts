##############################
# Set variables
##############################

# Original pool name
[string]$AppPoolName = 'DefaultAppPool'

# New pool name
[string]$AppPoolNewName = 'DefaultAppPool2'

# Path to appcmd.exe
[string]$AppCmd = 'C:\windows\system32\inetsrv\appcmd.exe'

##############################
# Set functions
##############################


function Get-IISApps {
    <#
    .SYNOPSIS
        The function will return a list of applications associated with the AppPool.
    .DESCRIPTION
        The function get a full list of applications by using appcmd.exe.
        Then function will leave in the list only those applications that mention name of AppPool AppPoolName.
        The result is returned as an array of application names.
    .EXAMPLE
        [array]$List = Get-IISApps -AppPoolName 'DefaultAppPool'
    #>
    [CmdletBinding()]
    [OutputType([array])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppPoolName,

        [string]
        $AppCmd = 'C:\windows\system32\inetsrv\appcmd.exe'
    )
    process {
        # Exporting the list of applications to xml
        [string]$AppCmdCommand = $AppCmd + ' list app /config:* /xml'
        [xml]$ListAppsXml = (cmd.exe /c $AppCmdCommand)

        # Get application names
        [array]$ListAppNames = @()
        ($ListAppsXml.appcmd.ChildNodes ) | ForEach-Object {
            if ($_.'APPPOOL.NAME' -eq $AppPoolName){
                $ListAppNames += $_.'APP.NAME'
            }
        }

        # Return an array of application names
        $ListAppNames
    }
}


function Stop-IISPool {
    <#
    .SYNOPSIS
        The function stoping the AppPool
    .DESCRIPTION
        The function will execute stop command for the AppPool and waiting while AppPool status will be "Stopped".
        While AppPool status is not "Stopped" function wil be get status of AppPool every 100 ms.
        If AppPool status is not "Stopped" a long time, than function will be display a message about this every 5 seconds.
    .EXAMPLE
        Stop-IISPool -AppPoolName 'DefaultAppPool' -Format 'yyyy.MM.ddTHH:mm:ss.fffZ'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppPoolName,

        [string]
        $Format = 'yyyy.MM.dd HH:mm:ss.fff'
    )
    process {
        [string]$Message = $null
        [int]$ThisTimer = 0
        [int]$TimePrint = 0
        $null = (Import-Module WebAdministration)
        Do {
            [string]$AppPoolState = (Get-WebAppPoolState -Name $AppPoolName).Value
	        if ($AppPoolState -eq 'Started') {
                $null = (Stop-WebAppPool -Name $AppPoolName)
            }
	        Start-Sleep -Milliseconds $ThisTimer
	        $AppPoolState = (Get-WebAppPoolState -Name $AppPoolName).Value
	        if ($AppPoolState -ne 'Stopped') {
	            $ThisTimer = 100
	            $TimePrint += 100
	        }
	        if ($TimePrint -eq 5000) {
                [string]$ThisMessage = '[' + (get-date -Format $Format) + ']' + ' INFO ' + 'Waiting for the AppPool "' + $AppPoolName + '" to stop ...' + [System.Environment]::NewLine
                Write-Warning -Message $ThisMessage
                $Message += $ThisMessage
                $TimePrint = 0
            }
        } while ($AppPoolState -ne 'Stopped')
        $Message += '[' + (get-date -Format $Format) + ']' + ' INFO ' + 'AppPool "' + $AppPoolName + '" is stopped.'
        write-host $Message
    }
}


function Start-IISPool {
    <#
    .SYNOPSIS
        The function starting the AppPool
    .DESCRIPTION
        The function will execute start command for the AppPool and waiting while AppPool status will be "Started".
        While AppPool status is not "Started" function wil be get status of AppPool every 100 ms.
        If AppPool status is not "Started" a long time, than function will be display a message about this every 5 seconds.
    .EXAMPLE
        Start-IISPool -AppPoolName 'DefaultAppPool' -Format 'yyyy.MM.ddTHH:mm:ss.fffZ'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppPoolName,

        [string]
        $Format = 'yyyy.MM.dd HH:mm:ss.fff'
    )
    process {
        [string]$Message = $null
        $ThisTimer = 0
        $TimePrint = 0
        $null = (Import-Module WebAdministration)
        Do {
            $AppPoolState = (Get-WebAppPoolState -Name $AppPoolName).Value
	        if ($AppPoolState -eq 'Stopped') {
                $null = (Start-WebAppPool -Name $AppPoolName)
            }
	        Start-Sleep -Milliseconds $ThisTimer
	        $AppPoolState = (Get-WebAppPoolState -Name $AppPoolName).Value
	        if ($AppPoolState -ne 'Started') {
	            $ThisTimer = 100
	            $TimePrint += 100
	        }
	        if ($TimePrint -eq 5000) {
                [string]$ThisMessage = '[' + (get-date -Format $Format) + ']' + ' INFO ' + 'Waiting for the AppPool "' + $AppPoolName + '" to start ...' + [System.Environment]::NewLine
                Write-Warning -Message $ThisMessage
                $Message += $ThisMessage
                $TimePrint = 0
            }
        } while ($AppPoolState -ne 'Started')
        $Message += '[' + (get-date -Format $Format) + ']' + ' INFO ' + 'AppPool "' + $AppPoolName + '" is started.'
        write-host $Message
    }
}


function Rename-IISPool {
    <#
    .SYNOPSIS
        The function will renaming the AppPool
    .DESCRIPTION
        The function will creating a new AppPool with the same characteristics as the original AppPool.
        All applications what associated with the original AppPool will be reconfigured to the new AppPool.
        The original AppPool will be deleted.
        The procedure is as follows:
        1. Creating a new AppPool.
        2. Stoping the original AppPool if it is running.
        3. Reconfiguring applications what associated with the original AppPool to the new AppPool.
        4. Starting a new AppPool if the original AppPool was started.
        5. Deleting the original AppPool.
    .EXAMPLE
        Rename-IISPool -AppPoolName 'DefaultAppPool' -AppPoolNewName 'NewDefaultAppPool' -AppCmd 'C:\windows\system32\inetsrv\appcmd.exe'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppPoolName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppPoolNewName,

        [string]
        $AppCmd = 'C:\windows\system32\inetsrv\appcmd.exe',

        [string]
        $Format = 'yyyy.MM.dd HH:mm:ss.fff'
    )
    process {    
    
        # Setting the default encoding
        $PSDefaultParameterValues['*:Encoding'] = 'utf8'

        # Connecting a module to work with IIS
        $null = (Import-Module WebAdministration)

        # Check the existing of pools
        [bool]$ExistAppPoolOriginal =  Test-Path ('IIS:\AppPools\' + $AppPoolName)
        [bool]$ExistAppPoolNew =  Test-Path ('IIS:\AppPools\' + $AppPoolNewName)

        if (-not $ExistAppPoolOriginal) {
            [string]$Message = '[' + (get-date -Format $Format) + ']' + ' DEBUG ' + 'AppPool "' + $AppPoolName + '" does not exist. AppPool "' + $AppPoolNewName + '" will not be created.'
            write-host $Message
        }

        if ($ExistAppPoolNew) {
            [string]$Message = '[' + (get-date -Format $Format) + ']' + ' DEBUG ' + 'AppPool "' + $AppPoolNewName + '" already exists. Nothing to do.'
            write-host $Message
        }

        if ($ExistAppPoolOriginal -and (-not $ExistAppPoolNew)) {
            # Getting a list of applications what associated with a AppPool
            [array]$ListAppNames = Get-IISApps -AppPoolName $AppPoolName -AppCmd $AppCmd

            # Remembering the state of the AppPool
            [string]$AppPoolState = (Get-WebAppPoolState -Name $AppPoolName).Value
    
            # Exporting AppPool to a temporary file
            [string]$TmpPath = New-TemporaryFile
            [string]$AppCmdCommand = $AppCmd + ' list apppool ' + $AppPoolName + ' /config:* /xml > ' + $TmpPath
            try {
                [string]$ExportLog = (cmd.exe /c $AppCmdCommand)
            } finally {
                $ErrorCodeExport = $Lastexitcode
            }
            if ($ErrorCodeExport -eq 0) {
                # Renaming AppPool in a temporary file
                [xml]$NewAppPool = Get-Content $TmpPath
                $NewAppPool.appcmd.APPPOOL.'APPPOOL.NAME' = $AppPoolNewName
                $NewAppPool.appcmd.APPPOOL.add | Where-Object { $_.name -eq $AppPoolName } | ForEach-Object {
                    $_.name = $AppPoolNewName
                }
                $NewAppPool.Save($TmpPath)
    
                # Importing AppPool - Creating AppPool with a new name.
                [string]$AppCmdCommand = $AppCmd + ' add apppool /in < ' + $TmpPath
                try {
                    [string]$ImportLog = (cmd.exe /c $AppCmdCommand)
                } finally {
                    $ErrorCodeImport = $Lastexitcode
                }
                if ($ErrorCodeImport -eq 0) {
                    # Deleting a temporary file
                    Remove-Item -Force $TmpPath
        
                    # AppPool creation check - for slow servers
                    $ThisTimer = 0
                    $TimePrint = 0
                    Do {
                        [bool]$AppPoolExist = (Test-Path ('IIS:\AppPools\' + $AppPoolNewName))
                        Start-Sleep -Milliseconds $ThisTimer
                        if (-not ($AppPoolExist)) {
                            $ThisTimer = 100
	                        $TimePrint += 100
                        }
                        if ($TimePrint -eq 5000) {
                            [string]$Message = '[' + (get-date -Format $Format) + ']' + ' INFO ' + 'The AppPool import command was successful, but the AppPool "' + $AppPoolNewName + '" has not yet been created. We are waiting ...'
                            write-host $Message
                            $TimePrint = 0
                        }
                    } while (-not ($AppPoolExist))
					[string]$Message = '[' + (get-date -Format $Format) + ']' + ' INFO ' + 'AppPool "' + $AppPoolNewName + '" created.'
                    write-host $Message
    
                    # Stopping the original AppPool
                    if ($AppPoolState -ne 'Stopped') {
                        Stop-IISPool -AppPoolName $AppPoolName
                    }
    
                    # If the AppPool is associated with applications, then the applications will be reconfigure to the new AppPool.
                    if ($ListAppNames.Count -gt 0) {
                        Foreach ($AppName in $ListAppNames) {
                            # Reconfiguring the application
                            [string]$AppCmdCommand = $AppCmd + ' set app "' +  $AppName + '" /applicationPool:"' + $AppPoolNewName + '"'
                            $ReconfigAppLog = (cmd.exe /c $AppCmdCommand)
    
                            # Checking that the reconfiguration was completed correctly
                            [string]$AppCmdCommand = $AppCmd + ' list app /config:* /xml'
                            [xml]$AppXml = (cmd.exe /c $AppCmdCommand)
                            [bool]$AppReconfigTrue = $true
                            $AppXml.appcmd.APP | Where-Object { $_.'APP.NAME' -eq $AppName } | ForEach-Object {
                                if ($_.'APPPOOL.NAME' -ne $AppPoolNewName){
                                    $AppReconfigTrue = $false
                                }
                            }
    
                            # Displaying message
                            if ($AppReconfigTrue) {
                                [string]$Message = '[' + (get-date -Format $Format) + ']' + ' INFO ' + 'App "' + $AppName + '" was reconfigured successfully.'
                                write-host $Message
                            } else {
                                [string]$Message = '[' + (get-date -Format $Format) + ']' + ' ERROR ' + 'App "' + $AppName + '" not reconfigured!'
                                write-host $Message
                            }
                        }
                    }
    
                    # Starting a new AppPool
                    if ($AppPoolState -ne 'Stopped') {
                        Start-IISPool -AppPoolName $AppPoolNewName
                    }
    
                    # Deleting original AppPool
                    [string]$AppCmdCommand = $AppCmd + ' delete apppool "' + $AppPoolName + '"'
                    [string]$RemoveLog = (cmd.exe /c $AppCmdCommand)
    
                    # Checking if the AppPool has been deleted
                    if (-not (Test-Path ('IIS:\AppPools\' + $AppPoolName))) {
                        [string]$Message = '[' + (get-date -Format $Format) + ']' + ' INFO ' + 'AppPool "' + $AppPoolName + '" deleted.'
                        write-host $Message
                    } else {
                        [string]$Message = '[' + (get-date -Format $Format) + ']' + ' ERROR ' + 'AppPool "' + $AppPoolName + '" not deleted! Result of executing the delete AppPool command: ' + $RemoveLog
                        write-host $Message
                    }
                } else {
                    [string]$Message = '[' + (get-date -Format $Format) + ']' + ' ERROR ' + 'An error occurred while importing AppPool "' + $AppPoolNewName + '"! AppPool has not been renamed. Error text: ' + $ImportLog
                    write-host $Message
                }
            } else {
                [string]$Message = '[' + (get-date -Format $Format) + ']' + ' ERROR ' + 'An error occurred while exporting AppPool "' + $AppPoolNewName + '"! AppPool has not been renamed. Error text: ' + $ExportLog
                write-host $Message
            }
        }
    }
}

##############################
# Executable code
##############################

Rename-IISPool -AppPoolName $AppPoolName -AppPoolNewName $AppPoolNewName -AppCmd $AppCmd
