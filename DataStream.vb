Option Strict Off
Option Explicit On 
Imports System
Imports System.IO




Module DataStream
    Private Const FILE_NAME01 As String = "C:\Streaming Production Data\D13137_Production_Stream.csv"
    Private Const FILE_NAME02 As String = "C:\Streaming Validation Data\D13137_Validation_Stream.csv"

    Public Sub writeProductionPacket()
        Dim productionPacket As String
        Dim packetTimeStamp As String
        Dim timeStamp As DateTime

        ' time stamp
        timeStamp = DateTime.Now
        packetTimeStamp = timeStamp.Now.ToString
        '
        If (StreamDataNoteBool = False) Then

            productionPacket = packetTimeStamp _
                                & "," & padDbl.ToString _
                                & "," & flangeDbl.ToString _
                                & "," & leftConeDbl.ToString _
                                & "," & rightConeDbl.ToString _
                                & "," & PartNumberStr _
                                & "," & OperatorIdStr _
                                & "," & MoldNumberStr _
                                & "," & ShopOrderStr _
                                & "," & PartTimeStr _
                                & "," & PartDateStr

        ElseIf StreamDataNoteBool Then

            productionPacket = "NOTE: " & "," & StreamingDataNote

            ' Reset Streaming note bit
            StreamDataNoteBool = False

        End If

        ' write to file
        'If File.Exists(FILE_NAME) Then
        '    Dim sr As StreamWriter = New StreamWriter(FILE_NAME)
        '    sr.WriteLine(productionPacket)
        '    sr.Close()
        'Else
        'If File.Exists(FILE_NAME) Then
        Dim sr As StreamWriter = File.AppendText(FILE_NAME01)
        sr.WriteLine(productionPacket)

        sr.Close()
        'End If
    End Sub
    Public Sub writeValidationPacket()
        Dim validationPacket As String
        Dim packetTimeStamp As String
        Dim timeStamp As DateTime

        ' time stamp
        timeStamp = DateTime.Now
        packetTimeStamp = timeStamp.Now.ToString
        '
        If (StreamDataNoteBool = False) Then

            validationPacket = packetTimeStamp _
                            & "," & padDbl.ToString _
                            & "," & flangeDbl.ToString _
                            & "," & leftConeDbl.ToString _
                            & "," & rightConeDbl.ToString _
                            & "," & PartNumberStr _
                            & "," & OperatorIdStr _
                            & "," & MoldNumberStr _
                            & "," & CavityNumberStr _
                            & "," & HeatNumberStr _
                            & "," & RunNumberStr

        ElseIf StreamDataNoteBool Then

            validationPacket = "NOTE: " & "," & StreamingDataNote

            

        End If

        ' write to file
        'If File.Exists(FILE_NAME) Then
        '    Dim sr As StreamWriter = New StreamWriter(FILE_NAME)
        '    sr.WriteLine(validationPacket)
        '    sr.Close()
        'Else
        Dim sr As StreamWriter = File.AppendText(FILE_NAME02)
        sr.WriteLine(validationPacket)
        sr.Close()
        writeStreamOneShot = True
        'End If
        'increment count
        If (StreamDataNoteBool = False) Then
            AutoPartCount = AutoPartCount + 1
            AutoPartCountRunning = AutoPartCountRunning + 1
        End If
        ' Reset Streaming note bit
        StreamDataNoteBool = False

    End Sub





End Module