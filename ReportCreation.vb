Option Strict Off
Option Explicit On 
Imports Microsoft.Office.Core
Module ReportCreation
    Public objApp As Excel.Application
    Public objBook As Excel.Workbook
    Public Sub CreateNewReport()
        Dim objBooks As Excel.Workbooks
        Dim objSheets As Excel.Sheets
        Dim objSheet As Excel._Worksheet
        Dim range As Excel.Range
        Dim i As Integer = 0
        Dim fileName As String
        Dim timeStamp As DateTime
        Dim TimeStr As String
        Dim DateStr As String
        Dim strTemp As String

        'Conditional Formatting constants
        'Pad:
        Dim PadUpperLim02 As Double = 0.018
        Dim PadUpperLim01 As Double = 0.0175
        Dim PadLowerLim01 As Double = 0.0155
        Dim PadLowerLim02 As Double = 0.015

        'Flange:
        Dim FlangeUpperLim02 As Double = 0.012
        Dim FlangeUpperLim01 As Double = 0.0115
        Dim FlangeLowerLim01 As Double = 0.0095
        Dim FlangeLowerLim02 As Double = 0.009

        'Left Cone:
        Dim LeftConeUpperLim02 As Double = 0.008
        Dim LeftConeUpperLim01 As Double = 0.0075
        Dim LeftConeLowerLim01 As Double = 0.0055
        Dim LeftConeLowerLim02 As Double = 0.005

        'Right Cone:
        Dim RightConeUpperLim02 As Double = 0.008
        Dim RightConeUpperLim01 As Double = 0.0075
        Dim RightConeLowerLim01 As Double = 0.0055
        Dim RightConeLowerLim02 As Double = 0.005

        'start a new workbook.
        objBooks = objApp.Workbooks
        objBook = objBooks.Add
        objSheets = objBook.Worksheets
        objSheet = objSheets(1)

        ' Make Excel Invisible
        objApp.Visible = False

        'Protection:
        'objSheets.Protect(Password:="9999", Contents:=True, _
        '                DrawingObjects:=True, Scenarios:=True, _
        '                AllowFormattingCells:=False, AllowSorting:=True)

        ' Format Spreadsheet:
        'Title
        objSheet.Range("C1").Value = "D13137 Production Report"
        objSheet.Range("A1").ColumnWidth = 7.0
        objSheet.Range("B1").ColumnWidth = 11.86
        objSheet.Range("C1").ColumnWidth = 12.89
        objSheet.Range("D1").ColumnWidth = 9.86
        objSheet.Range("E1").ColumnWidth = 20.0
        objSheet.Range("F1").ColumnWidth = 12.57
        objSheet.Range("G1").ColumnWidth = 6.71
        objSheet.Range("H1").ColumnWidth = 11.57
        objSheet.Range("C1", "E1").Merge()
        objSheet.Range("C2", "E2").Merge()
        objSheet.Range("C1", "C2").Merge()
        objSheet.Range("C1").VerticalAlignment = 2
        objSheet.Range("C1").HorizontalAlignment = 3
        objSheet.Range("C1").Font.Size = 18
        objSheet.Range("C1").Font.Bold = True
        'objSheet.Range("A2", "H2").Borders.Value = 1
        With objSheet.Range("A2", "H2").Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThick
            .ColorIndex = 5
        End With

        'Margins
        objSheet.PageSetup.TopMargin = 0
        objSheet.PageSetup.HeaderMargin = 0
        objSheet.PageSetup.RightMargin = 0
        objSheet.PageSetup.FooterMargin = 0
        objSheet.PageSetup.BottomMargin = 0

        'Date and Time:
        'Date:
        objSheet.Range("G1", "H1").Merge()
        objSheet.Range("G1").Font.Bold = True
        objSheet.Range("G1").Formula = "=TODAY()"
        objSheet.Range("G1").NumberFormat = "mmmm d, yyyy"
        'Time:
        objSheet.Range("G2", "H2").Merge()
        objSheet.Range("G2").Font.Bold = True
        objSheet.Range("G2").Formula = "=NOW()"
        objSheet.Range("G2").NumberFormat = "h:mm:ss AM/PM"

        ' Part Information 
        'Part Number:
        objSheet.Range("B4").VerticalAlignment = 2
        objSheet.Range("B4").HorizontalAlignment = 4
        objSheet.Range("B4").Value = "Part Number:"
        objSheet.Range("C4", "D4").Merge()
        objSheet.Range("C4").VerticalAlignment = 2
        objSheet.Range("C4").HorizontalAlignment = 2
        objSheet.Range("C4").Value = PartNumberStr

        'Operator Id:
        objSheet.Range("B5").VerticalAlignment = 2
        objSheet.Range("B5").HorizontalAlignment = 4
        objSheet.Range("B5").Value = "Operator Id:"
        objSheet.Range("C5", "D5").Merge()
        objSheet.Range("C5").VerticalAlignment = 2
        objSheet.Range("C5").HorizontalAlignment = 2
        objSheet.Range("C5").Value = OperatorIdStr

        'Mold Number:
        objSheet.Range("B6").VerticalAlignment = 2
        objSheet.Range("B6").HorizontalAlignment = 4
        objSheet.Range("B6").Value = "Mold Number:"
        objSheet.Range("C6", "D6").Merge()
        objSheet.Range("C6").VerticalAlignment = 2
        objSheet.Range("C6").HorizontalAlignment = 2
        objSheet.Range("C6").Value = MoldNumberStr

        'Shop Order:
        objSheet.Range("F4").VerticalAlignment = 2
        objSheet.Range("F4").HorizontalAlignment = 4
        objSheet.Range("F4").Value = "Shop Order:"
        objSheet.Range("G4", "H4").Merge()
        objSheet.Range("G4").VerticalAlignment = 2
        objSheet.Range("G4").HorizontalAlignment = 2
        objSheet.Range("G4").Value = ShopOrderStr

        'Sample Time:
        objSheet.Range("F5").VerticalAlignment = 2
        objSheet.Range("F5").HorizontalAlignment = 4
        objSheet.Range("F5").Value = "Sample Time:"
        objSheet.Range("G5", "H5").Merge()
        objSheet.Range("G5").VerticalAlignment = 2
        objSheet.Range("G5").HorizontalAlignment = 2
        objSheet.Range("G5").Value = PartTimeStr

        'Sample Date:
        objSheet.Range("F6").VerticalAlignment = 2
        objSheet.Range("F6").HorizontalAlignment = 4
        objSheet.Range("F6").Value = "Sample Date:"
        objSheet.Range("G6", "H6").Merge()
        objSheet.Range("G6").VerticalAlignment = 2
        objSheet.Range("G6").HorizontalAlignment = 2
        objSheet.Range("G6").Value = PartDateStr

        'Data Headings X-Axis
        objSheet.Range("B9").VerticalAlignment = 2
        objSheet.Range("B9").HorizontalAlignment = 3
        objSheet.Range("B9").Font.Size = 11
        objSheet.Range("B9").Font.Bold = True
        objSheet.Range("B9").Value = "Pad"
        objSheet.Range("B9").Borders.Value = 1
        objSheet.Range("B9").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        objSheet.Range("B9").Borders.Weight = Excel.XlBorderWeight.xlThick

        objSheet.Range("D9").VerticalAlignment = 2
        objSheet.Range("D9").HorizontalAlignment = 3
        objSheet.Range("D9").Font.Size = 11
        objSheet.Range("D9").Font.Bold = True
        objSheet.Range("D9").Value = "Flange"
        objSheet.Range("D9").Borders.Value = 1
        objSheet.Range("D9").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        objSheet.Range("D9").Borders.Weight = Excel.XlBorderWeight.xlThick

        objSheet.Range("F9").VerticalAlignment = 2
        objSheet.Range("F9").HorizontalAlignment = 3
        objSheet.Range("F9").Font.Size = 11
        objSheet.Range("F9").Font.Bold = True
        objSheet.Range("F9").Value = "Left Cone"
        objSheet.Range("F9").Borders.Value = 1
        objSheet.Range("F9").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        objSheet.Range("F9").Borders.Weight = Excel.XlBorderWeight.xlThick

        objSheet.Range("H9").VerticalAlignment = 2
        objSheet.Range("H9").HorizontalAlignment = 3
        objSheet.Range("H9").Font.Size = 11
        objSheet.Range("H9").Font.Bold = True
        objSheet.Range("H9").Value = "Right Cone"
        objSheet.Range("H9").Borders.Value = 1
        objSheet.Range("H9").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        objSheet.Range("H9").Borders.Weight = Excel.XlBorderWeight.xlThick

        'Data Headings Y-Axis
        objSheet.Range("A10").VerticalAlignment = 2
        objSheet.Range("A10").HorizontalAlignment = 2
        objSheet.Range("A10").Value = "Nominal"

        objSheet.Range("A11").VerticalAlignment = 2
        objSheet.Range("A11").HorizontalAlignment = 2
        objSheet.Range("A11").Value = "USL"

        objSheet.Range("A12").VerticalAlignment = 2
        objSheet.Range("A12").HorizontalAlignment = 2
        objSheet.Range("A12").Value = "LSL"

        objSheet.Range("A13").VerticalAlignment = 2
        objSheet.Range("A13").HorizontalAlignment = 4
        objSheet.Range("A13").Value = "1"

        objSheet.Range("A14").VerticalAlignment = 2
        objSheet.Range("A14").HorizontalAlignment = 4
        objSheet.Range("A14").Value = "2"

        objSheet.Range("A15").VerticalAlignment = 2
        objSheet.Range("A15").HorizontalAlignment = 4
        objSheet.Range("A15").Value = "3"

        objSheet.Range("A16").VerticalAlignment = 2
        objSheet.Range("A16").HorizontalAlignment = 4
        objSheet.Range("A16").Value = "4"

        objSheet.Range("A17").VerticalAlignment = 2
        objSheet.Range("A17").HorizontalAlignment = 4
        objSheet.Range("A17").Value = "5"

        objSheet.Range("A22").VerticalAlignment = 2
        objSheet.Range("A22").HorizontalAlignment = 2
        objSheet.Range("A22").Value = "Mean"

        objSheet.Range("A23").VerticalAlignment = 2
        objSheet.Range("A23").HorizontalAlignment = 2
        objSheet.Range("A23").Value = "Sigma"

        objSheet.Range("A24").VerticalAlignment = 2
        objSheet.Range("A24").HorizontalAlignment = 2
        objSheet.Range("A24").Value = "Max"

        objSheet.Range("A25").VerticalAlignment = 2
        objSheet.Range("A25").HorizontalAlignment = 2
        objSheet.Range("A25").Value = "Min"

        objSheet.Range("A26").VerticalAlignment = 2
        objSheet.Range("A26").HorizontalAlignment = 2
        objSheet.Range("A26").Value = "Range"

        'Statistics:
        'Pad
        objSheet.Range("B22", "B26").VerticalAlignment = 2
        objSheet.Range("B22", "B26").HorizontalAlignment = 3
        objSheet.Range("B22").Formula = "=AVERAGE(B13:B17)"
        objSheet.Range("B23").Formula = "=STDEV(B13:B17)"
        objSheet.Range("B24").Formula = "=MAX(B13:B17)"
        objSheet.Range("B25").Formula = "=MIN(B13:B17)"
        objSheet.Range("B26").Formula = "=B24-B25"

        'Flange
        objSheet.Range("D22", "D26").VerticalAlignment = 2
        objSheet.Range("D22", "D26").HorizontalAlignment = 3
        objSheet.Range("D22").Formula = "=AVERAGE(D13:D17)"
        objSheet.Range("D23").Formula = "=STDEV(D13:D17)"
        objSheet.Range("D24").Formula = "=MAX(D13:D17)"
        objSheet.Range("D25").Formula = "=MIN(D13:D17)"
        objSheet.Range("D26").Formula = "=D24-D25"

        'Left Cone
        objSheet.Range("F22", "F26").VerticalAlignment = 2
        objSheet.Range("F22", "F26").HorizontalAlignment = 3
        objSheet.Range("F22").Formula = "=AVERAGE(F13:F17)"
        objSheet.Range("F23").Formula = "=STDEV(F13:F17)"
        objSheet.Range("F24").Formula = "=MAX(F13:F17)"
        objSheet.Range("F25").Formula = "=MIN(F13:F17)"
        objSheet.Range("F26").Formula = "=F24-F25"

        'Right Cone
        objSheet.Range("H22", "H26").VerticalAlignment = 2
        objSheet.Range("H22", "H26").HorizontalAlignment = 3
        objSheet.Range("H22").Formula = "=AVERAGE(H13:H17)"
        objSheet.Range("H23").Formula = "=STDEV(H13:H17)"
        objSheet.Range("H24").Formula = "=MAX(H13:H17)"
        objSheet.Range("H25").Formula = "=MIN(H13:H17)"
        objSheet.Range("H26").Formula = "=H24-H25"

        'Number representation
        objSheet.Range("B10", "B26").NumberFormat = "0.00000"
        objSheet.Range("D10", "D26").NumberFormat = "0.00000"
        objSheet.Range("F10", "F26").NumberFormat = "0.00000"
        objSheet.Range("H10", "H26").NumberFormat = "0.00000"

        'Dimensional Data
        'Pad:
        objSheet.Range("B10", "B17").VerticalAlignment = 2
        objSheet.Range("B10", "B17").HorizontalAlignment = 3
        objSheet.Range("B10").Value = "0.0165"
        objSheet.Range("B11").Formula = "=B10+0.0015"
        objSheet.Range("B12").Formula = "=B10-0.0015"

        'Flange:
        objSheet.Range("D10", "D17").VerticalAlignment = 2
        objSheet.Range("D10", "D17").HorizontalAlignment = 3
        objSheet.Range("D10").Value = "0.0105"
        objSheet.Range("D11").Formula = "=D10+0.0015"
        objSheet.Range("D12").Formula = "=D10-0.0015"

        'Left Cone:
        objSheet.Range("F10", "F17").VerticalAlignment = 2
        objSheet.Range("F10", "F17").HorizontalAlignment = 3
        objSheet.Range("F10").Value = "0.0065"
        objSheet.Range("F11").Formula = "=F10+0.0015"
        objSheet.Range("F12").Formula = "=F10-0.0015"

        'Right Cone:
        objSheet.Range("H10", "H17").VerticalAlignment = 2
        objSheet.Range("H10", "H17").HorizontalAlignment = 3
        objSheet.Range("H10").Value = "0.0065"
        objSheet.Range("H11").Formula = "=H10+0.0015"
        objSheet.Range("H12").Formula = "=H10-0.0015"

        'Conditional Formating
        'For i = 13 To 17
        '    strTemp = "B" & i.ToString
        '    objSheet.Range(strTemp).Interior.ColorIndex = 5 'Blue
        '    With objSheet.Range(strTemp).FormatConditions
        '        .Delete()
        '        .Add( _
        '            Type:=Excel.XlFormatConditionType.xlExpression, _
        '             Formula1:="=IF(strTemp<0.015, strTemp>0.018)")
        '        .Item(1).Interior.ColorIndex = 3  'Red
        '        .Add( _
        '            Type:=Excel.XlFormatConditionType.xlExpression, _
        '            Formula1:="=IF(strTemp>=0.015, strTemp<=0.0155)")
        '        .Item(2).Interior.ColorIndex = 6   'Yellow
        '        .Add( _
        '            Type:=Excel.XlFormatConditionType.xlExpression, _
        '            Formula1:="=IF(strTemp>=0.0175, strTemp<=0.018)")
        '        .Item(3).Interior.ColorIndex = 6   'Yellow
        '    End With
        'Next i

        'Populate Arrays for Testing
        'For i = 0 To 4
        '    padDataArray(i) = i + 1
        '    flangeDataArray(i) = i + 6
        '    leftConeDataArray(i) = i + 11
        '    rightConeDataArray(i) = i + 16
        'Next i
        'Populate Data Fields
        For i = 0 To 4
            Dim PadPosition As String = "B"
            Dim FlangePosition As String = "D"
            Dim LeftConePosition As String = "F"
            Dim RightConePosition As String = "H"

            PadPosition = PadPosition & (i + 13).ToString ' Data starts at cell 13
            FlangePosition = FlangePosition & (i + 13).ToString
            LeftConePosition = LeftConePosition & (i + 13).ToString
            RightConePosition = RightConePosition & (i + 13).ToString

            ' Conditional Formatting
            'Pad:
            If (padDataArray(i) >= PadLowerLim01) And (padDataArray(i) < PadLowerLim02) Then
                'objSheet.Range(PadPosition).Interior.ColorIndex = 6 'Yellow
                objSheet.Range(PadPosition).Font.ColorIndex = 44 ' yellow
            ElseIf (padDataArray(i) > PadUpperLim01) And (padDataArray(i) <= PadUpperLim02) Then
                'objSheet.Range(PadPosition).Interior.ColorIndex = 6 'Yellow
                objSheet.Range(PadPosition).Font.ColorIndex = 44 ' yellow
            ElseIf (padDataArray(i) >= PadLowerLim02) And (padDataArray(i) <= PadUpperLim01) Then
                'objSheet.Range(PadPosition).Interior.ColorIndex = 5 'Blue
                objSheet.Range(PadPosition).Font.ColorIndex = 5 'Blue
            Else
                'objSheet.Range(PadPosition).Interior.ColorIndex = 3 'Red
                objSheet.Range(PadPosition).Font.ColorIndex = 3 'Red
            End If

            'Flange:
            If (flangeDataArray(i) >= FlangeLowerLim01) And (flangeDataArray(i) < FlangeLowerLim02) Then
                'objSheet.Range(FlangePosition).Interior.ColorIndex = 6 'Yellow
                objSheet.Range(FlangePosition).Font.ColorIndex = 44 'Yellow
            ElseIf (flangeDataArray(i) > FlangeUpperLim01) And (flangeDataArray(i) <= FlangeUpperLim02) Then
                'objSheet.Range(FlangePosition).Interior.ColorIndex = 6 'Yellow
                objSheet.Range(FlangePosition).Font.ColorIndex = 44 'Yellow
            ElseIf (flangeDataArray(i) >= FlangeLowerLim02) And (flangeDataArray(i) <= FlangeUpperLim01) Then
                'objSheet.Range(FlangePosition).Interior.ColorIndex = 5 'Blue
                objSheet.Range(FlangePosition).Font.ColorIndex = 5 'Blue
            Else
                'objSheet.Range(FlangePosition).Interior.ColorIndex = 3 'Red
                objSheet.Range(FlangePosition).Font.ColorIndex = 3 'Red
            End If

            'Left Cone:
            If (leftConeDataArray(i) >= LeftConeLowerLim01) And (leftConeDataArray(i) < LeftConeLowerLim02) Then
                'objSheet.Range(LeftConePosition).Interior.ColorIndex = 6 'Yellow
                objSheet.Range(LeftConePosition).Font.ColorIndex = 44 'Yellow
            ElseIf (leftConeDataArray(i) > LeftConeUpperLim01) And (leftConeDataArray(i) <= LeftConeUpperLim02) Then
                'objSheet.Range(LeftConePosition).Interior.ColorIndex = 6 'Yellow
                objSheet.Range(LeftConePosition).Font.ColorIndex = 44 'Yellow
            ElseIf (leftConeDataArray(i) >= LeftConeLowerLim02) And (leftConeDataArray(i) <= LeftConeUpperLim01) Then
                'objSheet.Range(LeftConePosition).Interior.ColorIndex = 5 'Blue
                objSheet.Range(LeftConePosition).Font.ColorIndex = 5 'Blue
            Else
                'objSheet.Range(LeftConePosition).Interior.ColorIndex = 3 'Red
                objSheet.Range(LeftConePosition).Font.ColorIndex = 3 'Red
            End If

            'Right Cone:
            If (rightConeDataArray(i) >= RightConeLowerLim01) And (rightConeDataArray(i) < RightConeLowerLim02) Then
                'objSheet.Range(RightConePosition).Interior.ColorIndex = 6 'Yellow
                objSheet.Range(RightConePosition).Font.ColorIndex = 44 'Yellow
            ElseIf (rightConeDataArray(i) > RightConeUpperLim01) And (rightConeDataArray(i) <= RightConeUpperLim02) Then
                'objSheet.Range(RightConePosition).Interior.ColorIndex = 6 'Yellow
                objSheet.Range(RightConePosition).Font.ColorIndex = 44 'Yellow
            ElseIf (rightConeDataArray(i) >= RightConeLowerLim02) And (rightConeDataArray(i) <= RightConeUpperLim01) Then
                'objSheet.Range(RightConePosition).Interior.ColorIndex = 5 'Blue
                objSheet.Range(RightConePosition).Font.ColorIndex = 5 'Blue
            Else
                'objSheet.Range(RightConePosition).Interior.ColorIndex = 3 'Red
                objSheet.Range(RightConePosition).Font.ColorIndex = 3 'Red
            End If

            objSheet.Range(PadPosition).Value = padDataArray(i).ToString
            objSheet.Range(FlangePosition).Value = flangeDataArray(i).ToString
            objSheet.Range(LeftConePosition).Value = leftConeDataArray(i).ToString
            objSheet.Range(RightConePosition).Value = rightConeDataArray(i).ToString
        Next i

        'Notes Section
        If ProductionReportBool Then
            ' Create Notes Section
            objSheet.Range("A28", "H28").Merge()
            objSheet.Range("A28", "A30").Merge()
            objSheet.Range("A28").VerticalAlignment = 1
            objSheet.Range("A28").WrapText = True
            ' Add note
            objSheet.Range("A28").Value = ProductionReportNote

            'Resert Production Note Bit
            ProductionReportBool = False
        End If

        ' Make Excel Invisible
        objApp.Visible = True
        timeStamp = timeStamp.Now
        TimeStr = timeStamp.Hour & timeStamp.Minute & timeStamp.Second
        DateStr = timeStamp.Month & timeStamp.Day & timeStamp.Year
        fileName = "C:\Production Reports\" & "D13137_" & DateStr & "_" & TimeStr & ".xls"

        ' Make report read-only
        objSheet.Protect()
        'Save a copy
        objBook.SaveCopyAs(fileName)

        'Clean up a little.
        range = Nothing
        objSheet = Nothing
        objSheets = Nothing
        objBooks = Nothing

        ' Reset Part Count
        AutoPartCount = 0

        Dim newform03 As CheckZero = New CheckZero
        newform03.Show()              'show instance of PartData
    End Sub
End Module