# Send-TVCommand: control a serial-connected Panasonic TV with PowerShell.
# Version 1.0
# cbfox01@syr.edu
function Send-TVCommand {
    param ([switch]$On,[switch]$Off,[String]$Raw)
    if ($PSBoundParameters.count -gt 1) {
        Write-Error -Category InvalidArgument -Message "This function accepts only one parameter per execution to process!" # Terminates.
    }
    if ($PSBoundParameters.count -eq 0) {
        Write-Host "Send-TVCommand is a sample PowerShell function demonstrating RS-232 control of serial-connected Panasonic commercial-class televisions. This is presently configured and verified to work with TH-*CQE1 (e.g. the 65`" model is TH-65CQE1) series displays."
        Write-Host "The function accepts three possible parameters:"
        Write-Host "`t-On`tTurn the attached TV on."
        Write-Host "`t-Off`tTurn the attached TV off."
        Write-Host "`t-Raw`tAccepts a raw command string to send to the TV. Review the available commands here:"
        Write-Host "`t`thttps://panasonic.net/cns/prodisplays/support/download/pdf/SQE1_CQE1_SerialCommandList.pdf"
        Write-Host "A sample incantation to query the currently set input:"
        Write-Host "`tSend-TVCommand -Raw QMI"
        Write-Host "A sample incantation to set the picture mode to Sports:"
        Write-Host "`tSend-TVCommand -Raw VPC:MENSPT"
        Write-Host "A sample incantation to set the backlight to 25%. Note the need to include leading and trailing zeroes."
        Write-Host "`tSend-TVCommand -Raw VPC:BLT025"
        Write-Host "NOTE: Only these commands are available when the TV is in Standby (i.e. off but AC connected):"
        Write-Host "`tPON: Turns the TV on if it is off, leaves it on if it is on. Returns result code PON in either case."
        Write-Host "`tPOF: Turns the TV off if it is on, leaves it off if it is off. Returns result code POF in either case."
        Write-Host "`tQPW: Returns current power status. Returns result code QPW:0 if the TV is off (standby) and QPW:1 if the TV is on."
        Write-Host "`tQRV: Returns information about the TV's software version."
        Write-Host "`tQID: Returns information about the TV size, model, and market."
        Write-Host "NOTE: At this time Send-TVCommand accepts only one parameter (i.e. command) per execution to process."
        Return # Terminates.
    }

    # Define convenience variables.
    $stx = [char]02 # "Start of text" control character.
    $etx = [char]03 # "End of text" control character.
    $enc = [System.Text.Encoding]::ASCII # ASCII encoding for RS-232 send/receive

    # Configure and open the serial port.
    # Our serial port is COM1, and Panasonic TH-*CQE1 series expects 9600 baud, no parity, eight data bits and one stop bit.
    $port = New-Object System.IO.Ports.SerialPort COM1,9600,None,8,One
    $port.Open()

    # Process parameter.
    switch ($PSBoundParameters.Keys) {
        "On" {$cmd = "PON"}
        "Off" {$cmd = "POF"}
        "Raw" {$cmd = $Raw}
    }

    # Create the command byte stream, with start and end text characters.
    $out = $enc.GetBytes("$($stx)$($cmd)$($etx)")

    # Write the byte stream to the serial port.
    $port.Write($out,0,$out.length)

    # Wait up to 10 seconds for a result to be received.
    $timeout = 10 # Seconds
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    while (($timer.Elapsed.TotalSeconds -lt $timeout) -and ($port.BytesToRead -eq 0)) {
        Start-Sleep -Seconds 1
    }
    $timer.Stop()

    # If no result received, error, otherwise collect result code received from TV with control characters stripped.
    if ($port.BytesToRead -eq 0) {
        $port.Close() # Close the serial port.
        Write-Error -Category OperationTimeout -Message "No reply from TV within 10 seconds." # Terminates.
    } else {
        # Collect the received result from the serial port buffer, strip stx and etx control characters.
        $in = $port.ReadExisting().Trim(@($stx,$etx))
        $port.Close() # Close the serial port.
    }

    #Return the received result.
    Return $in # Terminates.
}