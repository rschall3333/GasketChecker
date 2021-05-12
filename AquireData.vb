Option Strict Off
Option Explicit On 
Module AquireData
    Public Sub DataAquisition()
        Dim sTxHead01 As String
        Dim sRxHead01 As String
        Dim sTxHead02 As String
        Dim sRxHead02 As String
        Dim sTxLKG3001 As String
        Dim sRxLKG3001 As String
        Dim StringCompareBool As Boolean
        Dim port02Availabile As Boolean     ' 0 - No, 1 - Yes
        Dim port03Availabile As Boolean     ' 0 - No, 1 - Yes
        '********************************************************
        '***************** Open com port02 **********************
        '********************************************************
        keyenceLK5001 = New Rs232
        Try
            '// Setup Parameters
            With keyenceLK5001
                .Port = 2
                .BaudRate = 9600
                .DataBit = 8
                .StopBit = Rs232.DataStopBit.StopBit_1
                .Parity = Rs232.DataParity.Parity_None
                .Timeout = 1500
            End With
            '// Initializes port
            keyenceLK5001.Open()
            '// Set state of RTS / DTS
            keyenceLK5001.Dtr = False
            keyenceLK5001.Rts = False
        Catch Ex As Exception When (AutoModeBool And boolCycleRunning)
            MessageBox.Show(Ex.Message, "Connection Error", MessageBoxButtons.OK)
        End Try

        '************** Head01-Output02 ********************************
        port02Availabile = keyenceLK5001.IsPortAvailable(2)
        If port02Availabile Then
            port02Head01Out02Err = False
            '// Clear Tx/Rx Buffers for Head01 reading
            keyenceLK5001.PurgeBuffer(Rs232.PurgeBuffers.TxClear Or Rs232.PurgeBuffers.RXClear)

            sTxHead01 = "M2"      ' M1 sends a single value from the Keyence Head01
            sTxHead01 += ControlChars.Cr      ' Add a CR
            'Request a single value from Keyence Head01
            keyenceLK5001.Write(sTxHead01)

            ' Get Data from Keyence Head01
            keyenceLK5001.Read(10)
            sRxHead01 = keyenceLK5001.InputStreamString()
            StringCompareBool = sRxHead01.StartsWith("-")

            If (StringCompareBool = False) Then
                Try
                    ' Convert Value to Double
                    keyenceHead01Output02 = CDbl(keyenceLK5001.InputStreamString)
                Catch ex As Exception
                    'MessageBox.Show("Open value not numeric")
                End Try
            End If

            If (AutoModeBool And boolCycleRunning) Or CheckZeroPageUp Then
                'Call left cone Compute Value
                leftConeCompute(keyenceHead01Output02)
            ElseIf (AutoModeBool = False) Then
                'Call left cone Compute Value
                leftConeCompute(keyenceHead01Output02)
            End If
        Else
            port02Head01Out02Err = True
        End If

        '************** Head01-Output03 ********************************
        If port02Availabile Then
            port02Head01Out03Err = False
            '// Clear Tx/Rx Buffers for Head01 reading
            keyenceLK5001.PurgeBuffer(Rs232.PurgeBuffers.TxClear Or Rs232.PurgeBuffers.RXClear)

            sTxHead01 = "M3"      ' M1 sends a single value from the Keyence Head01
            sTxHead01 += ControlChars.Cr      ' Add a CR
            'Request a single value from Keyence Head01
            keyenceLK5001.Write(sTxHead01)

            ' Get Data from Keyence Head01
            keyenceLK5001.Read(10)
            sRxHead01 = keyenceLK5001.InputStreamString()
            StringCompareBool = sRxHead01.StartsWith("-")

            If (StringCompareBool = False) Then
                Try
                    ' Convert Value to Double
                    keyenceHead01Output03 = CDbl(keyenceLK5001.InputStreamString)
                Catch ex As Exception
                    'MessageBox.Show("Open value not numeric")
                End Try
            End If

            If AutoModeBool And boolCycleRunning Or CheckZeroPageUp Then
                'Call right cone Compute Value
                rightConeCompute(keyenceHead01Output03)
            ElseIf (AutoModeBool = False) Then
                'Call right cone Compute Value
                rightConeCompute(keyenceHead01Output03)
            End If
        Else
            port02Head01Out03Err = True
        End If

        '************** Head02-Output04 ********************************
        port02Availabile = keyenceLK5001.IsPortAvailable(2)
        If port02Availabile Then
            port02Head02Out04Err = False
            '// Clear Tx/Rx Buffers for Head02 reading
            keyenceLK5001.PurgeBuffer(Rs232.PurgeBuffers.TxClear Or Rs232.PurgeBuffers.RXClear)

            sTxHead02 = "M4"      ' M4 sends a single value from the Keyence Head01
            sTxHead02 += ControlChars.Cr      ' Add a CR
            'Request a single value from Keyence Head02
            keyenceLK5001.Write(sTxHead02)

            ' Get Data from Keyence Head02
            keyenceLK5001.Read(10)
            sRxHead02 = keyenceLK5001.InputStreamString
            StringCompareBool = sRxHead02.StartsWith("-")
            If (StringCompareBool = False) Then
                Try
                    ' Convert Value to Double
                    keyenceHead02Output04 = CDbl(keyenceLK5001.InputStreamString)
                Catch ex As Exception
                    'MessageBox.Show("Open value not numeric")
                End Try
            End If

            If AutoModeBool And boolCycleRunning Or CheckZeroPageUp Then
                'Call right cone Compute Value
                PadCompute(keyenceHead02Output04)
            ElseIf (AutoModeBool = False) Then
                'Call right cone Compute Value
                PadCompute(keyenceHead02Output04)
            End If
        Else
            port02Head02Out04Err = True
        End If
        keyenceLK5001.Close()
        '********************************************************
        '***************** COM port2 closed *********************
        '********************************************************

        '********************************************************
        '***************** COM port5 open ***********************
        '********************************************************

        '************** LK-G87 - Output02 ********************************
        ' Open com port06 for Keyence Head01
        keyenceLKG3001 = New Rs232
        If (laserDisplacementERROR = False) Then
            Try
                '// Setup Parameters
                With keyenceLKG3001
                    .Port = 5
                    .BaudRate = 9600
                    .DataBit = 8
                    .StopBit = Rs232.DataStopBit.StopBit_1
                    .Parity = Rs232.DataParity.Parity_None
                    .Timeout = 1500
                End With
                '// Initializes port
                keyenceLKG3001.Open()
                '// Set state of RTS / DTS
                keyenceLKG3001.Dtr = False
                keyenceLKG3001.Rts = False

            Catch Ex2 As Exception
                MessageBox.Show(Ex2.Message, "Connection Error", MessageBoxButtons.OK)
            End Try
        End If
        port03Availabile = keyenceLKG3001.IsPortAvailable(5)
        If port03Availabile Then
            port05Head01Out01Err = False
            If (laserDisplacementERROR = False) Then
                '// Clear Tx/Rx Buffers for Head01 reading
                keyenceLKG3001.PurgeBuffer(Rs232.PurgeBuffers.TxClear Or Rs232.PurgeBuffers.RXClear)

                sTxLKG3001 = "M1"      ' M1 sends a single value from the Keyence Head01
                sTxLKG3001 += ControlChars.Cr      ' Add a CR
                'Request a single value from Keyence Head01
                keyenceLKG3001.Write(sTxLKG3001)

                ' Get Data from Keyence Head01
                keyenceLKG3001.Read(12)
                sRxLKG3001 = keyenceLKG3001.InputStreamString
                sRxLKG3001 = sRxLKG3001.Remove(0, 3)

                StringCompareBool = sRxLKG3001.StartsWith("-F")
                ' Convert Value to Double
                If (StringCompareBool = False) Then
                    Try
                        keyenceLKG3001Value = CDbl(sRxLKG3001)
                    Catch ex As Exception
                        'MessageBox.Show("Open value not numeric")
                    End Try
                End If
                If AutoModeBool And boolCycleRunning Or CheckZeroPageUp Then
                    'Call left cone Compute Value
                    FlangeCompute(keyenceLKG3001Value)
                ElseIf (AutoModeBool = False) Then
                    'Call left cone Compute Value
                    FlangeCompute(keyenceLKG3001Value)
                End If
            End If
        Else
            port05Head01Out01Err = True
        End If
        ' Close com port for reading Head01 and Head02
        keyenceLKG3001.Close()
        '********************************************************
        '***************** COM port5 closed *********************
        '********************************************************

        ' Update the Form with new data
        If AutoModeBool And boolCycleRunning Then
            If ProductionCollection Then
                If AutoPartCount <= 5 Then
                    BufferData()
                End If
                If ProductionCollection And writeStreamOneShot = False Then
                    writeProductionPacket()
                End If
            End If
            If ValidationCollection Then
                writeValidationPacket()
            End If
        End If
    End Sub
    Public Sub BufferData()
        padDataArray(AutoPartCount) = padDbl
        flangeDataArray(AutoPartCount) = flangeDbl
        leftConeDataArray(AutoPartCount) = leftConeDbl
        rightConeDataArray(AutoPartCount) = rightConeDbl
        'Increment Part Count
        AutoPartCount = AutoPartCount + 1
    End Sub
End Module