
$tui = $false

<#
From: https://stackoverflow.com/a/51692402/3969362
#>
function LoadModule ($m) {

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object {$_.Name -eq $m}) {
        #write-host "Module $m is already imported."
        return $true
    }
    else
    {
        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m})
        {
            Import-Module $m -Verbose
            return $true
        }
        else
        {
            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object {$_.Name -eq $m})
            {
                Write-Host "To enable advanced features run the following command from an admin PowerShell console:"
                Write-Host "Install-Module $m" -ForegroundColor Yellow
                Install-Module -Name $m -Force -Verbose -Scope CurrentUser
                Import-Module $m -Verbose
            }
            else
            {
                # If the module is not imported, not available and not in the online gallery then abort
                write-host "Module $m not imported, not available and not in an online gallery, exiting."
            }
            return $false
        }
    }
}

<#
Display com ports on our console host
#>
function DisplayComPorts()
{
    Get-CimInstance -Class Win32_SerialPort | Format-Table Name, Description, DeviceID, PNPDeviceID | Out-String | Write-Host
}

<#
Look for a COM port belonging to given device and monitor it.
See: https://learn.microsoft.com/en-us/dotnet/api/system.io.ports.serialport
TODO: Use SerialPort.DataReceived Event and SerialPort.ErrorReceived Event instead of pooling
#>
function MonitorPort()
{

    $wasMonitoring = $false
    # TODO: Have those as parameters from the script
    # TODO: Use a map with description
    # See: https://github.com/raspberrypi/usb-pid/blob/main/Readme.md
    $pnpDeviceIds = @(
        "USB\\VID_2E8A&PID_0009" # RPi Pico 2 W USB CDC stdio - Raspberry Pi Pico SDK CDC UART
        "USB\\VID_2E8A&PID_000A" # RPi Pico W USB CDC stdio - Raspberry Pi Pico SDK CDC UART (RP2040)
        "USB\\VID_2E8A&PID_000C" # RPi RP2040 Debug Probe - Raspberry Pi RP2040 CMSIS-DAP debug adapter
    )

    $pnpDeviceIdsRe = $pnpDeviceIds -join '|'

    Write-Host "Looking for known COM ports"

    # See: https://stackoverflow.com/a/63999061/3969362
    # Find COM port matching given PID and VID - in our case we used the ones matching our Pico 2 W
    $device = Get-CimInstance -Class Win32_SerialPort | Select-Object Name, Description, DeviceID, PNPDeviceID | Where-Object { $_.PNPDeviceID -match "($pnpDeviceIdsRe)" } | Select-Object -First 1

    if (!$device) {
            Write-Host "No known COM port found"
            Write-Host "Check your device VID and PID and edit this script accordingly"
            DisplayComPorts
            return
    } else {
            Write-Host "Found device: $($device.Description)"
            Write-Host "Opening port: $($device.DeviceID)"
    }

    # Open the serial port matching the device we found
    $port = new-Object System.IO.Ports.SerialPort $device.DeviceID,115200,None,8,One
    # We need RTS: Request To Send
    $port.Handshake = 2
    # We need DTR: Data Terminal Ready
    $port.DtrEnable = 1
    # Don't wait for ever on read and write operations
    $port.ReadTimeout = 1000
    $port.WriteTimeout = 1000

    try
    {
            $port.Open()
            Write-Host "Monitoring $($device.DeviceID)..."
            $wasMonitoring = $true

            #Register-ObjectEvent -InputObject $port -EventName "DataReceived" -SourceIdentifier "SerialPort.DataReceived"


            #$counter = 0
            $charInput = ""
            do {
                    # $dataEvent = Wait-Event -SourceIdentifier "SerialPort.DataReceived" -Timeout 1
                    # if ($dataEvent)
                    # {
                    #     Remove-Event -EventIdentifier $dataEvent.EventIdentifier
                    # }

                    #Write-Host "Write: ${counter}"
                    #$port.Write($counter)
                    # Wait long enough for our board to respond
                    #Start-Sleep -Milliseconds 100;
                    #Write-Host "ReadExisting"
                    $read = $port.ReadExisting()
                    #Write-Host "ReadExisting -done"
                    if ($read) {
                        #TODO: have a different color to distinguish for our logs
                        # See: https://stackoverflow.com/questions/20541456/list-of-all-colors-available-for-powershell
                        # Gray is the default
                        # DarkYellow DarkGray DarkCyan DarkGreen DarkMagenta Yellow
                        # Can't use us custo color for some reason 0xFFFAEBD7
                        Write-Host $read -NoNewline -Foreground DarkCyan
                    }

                    # Check if user pressed a key
                    # Boken in PS 7: $host.UI.RawUI.KeyAvailable
                    # See: https://github.com/PowerShell/PSReadLine/issues/959
                    if ([Console]::KeyAvailable) {
                        # Get next key
                        # See: https://learn.microsoft.com/en-us/dotnet/api/system.management.automation.host.pshostrawuserinterface.readkey
                        $key = $host.UI.RawUI.ReadKey();
                        #$Host.UI.RawUI.ReadKey("IncludeKeyUp,NoEcho").Character))
                        #Write-Host $key
                        # Check if return was pressed
                        # See: https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
                        #if ($key.VirtualKeyCode -eq 0x0D)
                        if ($key.Character -eq "`r")
                        {
                            if ($charInput -eq "")
                            {
                                # Send new line if return is pressed without input
                                $port.Write("`n")
                            }
                            else
                            {
                                # Send input when return is pressed
                                $port.Write($charInput)
                                $charInput = ""
                            }
                        }
                        else
                        {
                            $charInput = "$charInput$($key.Character)"
                        }
                    }

                    # Discard events from our queue otherwise WMI events would accumulate
                    Get-Event | Remove-Event
            }
            while ($port.IsOpen)
    }
    catch
    {
        Write-Host $_
    }
    finally
    {
            Write-Host "Monitoring stopped"
            # Make sure close the port even when doing Ctrl + C
            # That also runs when the port is gone for instance when the board is unplugged
            Write-Host "Closing port: $($device.DeviceID)"
            try {$port.Close()} catch {$_} # Defensive
    }

    return $wasMonitoring
}

function TryMonitorPort()
{
    $wasMonitoring = $false
    try
    {
       $wasMonitoring = MonitorPort
    }
    catch
    {
        # Print the exception
        Write-Host $_
    }
    finally
    {

    }

    return $wasMonitoring
}

try
{
    # Check if Terminal UI is installed
    <#
    if (LoadModule("Microsoft.PowerShell.ConsoleGuiTools"))
    {
        Write-Host "Advanced features enabled"
        $tui = $true
    }
    #>

    # Register for device connection events
    # See: https://stackoverflow.com/a/16374535/3969362
    # Though we use the PowerShell event queue instead of the action callback as it would block further WMI event from being processed
    $query = "SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2"
    $null = Register-CimIndicationEvent -Query $query -SourceIdentifier "Device.Connected"

    DisplayComPorts

    # Main application loop
    do
    {
        # Check if the COM port we want is there and monitor it
        # Keep doing data
        while (TryMonitorPort) {}
        # Wait for device connection event before trying to find our COM port again
        # Discard output see: https://stackoverflow.com/a/18413183/3969362
        # TODO: do timeout and process keys to specify a port to connect to for instance see above how we did it
        $wmiEvent = Wait-Event -SourceIdentifier "Device.Connected"
        Remove-Event -EventIdentifier $wmiEvent.EventIdentifier
        #Start-Sleep 3
    }
    While($True)

}
finally
{
    # Cancel all event subscriptions
    # See: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/unregister-event
    Get-EventSubscriber -Force | Unregister-Event -Force
    #
    Write-Host "Exit"
}
