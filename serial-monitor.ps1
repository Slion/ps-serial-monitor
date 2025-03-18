
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
    $pnpDeviceIds = @(
        "USB\\VID_2E8A&PID_0009" # RPi Pico 2 W USB CDC stdio
        "USB\\VID_2E8A&PID_000C" # RPi RP2040 Debug Probe
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
            #$counter = 0
            $charInput = ""
            do {
                    #Write-Host "Write: ${counter}"
                    #$port.Write($counter)
                    # Wait long enough for our board to respond
                    #Start-Sleep -Milliseconds 100;
                    $read = $port.ReadExisting()
                    if ($read) {
                        #TODO: have a different color to distinguish for our logs
                        # See: https://stackoverflow.com/questions/20541456/list-of-all-colors-available-for-powershell
                        # Gray is the default
                        # DarkYellow DarkGray DarkCyan DarkGreen DarkMagenta
                        # Can't use us custo color for some reason 0xFFFAEBD7
                        Write-Host -NoNewline $read -Foreground DarkCyan
                    }

                    # Check if user pressed a key
                    if ($Host.UI.RawUI.KeyAvailable) {
                        # Get next key
                        # See: https://learn.microsoft.com/en-us/dotnet/api/system.management.automation.host.pshostrawuserinterface.readkey
                        $key = $Host.UI.RawUI.ReadKey();
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
        $_
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

Try
{
    # Register for device connection events
    # See: https://stackoverflow.com/a/16374535/3969362
    # Though we use the PowerShell event queue instead of the action callback as it would block further WMI event from being processed
    $query = "SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2"
    $null = Register-WMIEvent -Query $query -SourceIdentifier "Device.Connected"

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
Finally
{
    # Cancel all event subscriptions
    # See: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/unregister-event
    Get-EventSubscriber -Force | Unregister-Event -Force
    #
    Write-Host "Exit"
}
