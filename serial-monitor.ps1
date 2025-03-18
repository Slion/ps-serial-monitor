$query = "SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2"
$wmi_event = Register-WMIEvent -Query $query -Action {
    Write-Host "A device has been inserted"

        # See: https://stackoverflow.com/a/63999061/3969362
    # Find COM port matching given PID and VID - in our case we used the ones matching our Pico 2 W
    $usb_vid = '2E8A'
    $usb_pid = '0009'
    $device = Get-CimInstance -Class Win32_SerialPort | Select-Object Name, Description, DeviceID, PNPDeviceID | Where-Object { $_.PNPDeviceID -match "USB\\VID_${usb_vid}&PID_${usb_pid}" } | Select-Object -First 1

    if (!$device) {
            Write-Host "No COM port matching VID:${usb_vid} PID:${usb_pid}"
            Write-Host "Check your device VID and PID and edit this script accordingly"
            Get-CimInstance -Class Win32_SerialPort | Format-Table Name, Description, DeviceID, PNPDeviceID | Out-String | Write-Host
            Exit
    } else {
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

    Try {
            $port.Open()
            #$counter = 0
            do {
                    #Write-Host "Write: ${counter}"
                    #$port.Write($counter)
                    # Wait long enough for our board to respond
                    #Start-Sleep -Milliseconds 100;
                    $read = $port.ReadExisting()
                    if ($read) {
                        Write-Host -NoNewline $read
                    }


                    # Don't do that too fast so that one can read the ouput
                    #Start-Sleep -Milliseconds 1000;
            }
            while ($port.IsOpen)

    } Finally {
            # Make sure close the port even when doing Ctrl + C
            # That also runs when the port is gone for instance when the board is unplugged
            Write-Host "Finally"
            $port.Close()
    }

}

Try {
    While($True) {
        #echo "looping"
        #Start-Sleep 3
    }
} Finally {
    Write-Host "Exit"
    Quit
}
