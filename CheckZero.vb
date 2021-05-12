Public Class CheckZero
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents CheckZeroButton As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents leftConeCheckLabel As System.Windows.Forms.Label
    Friend WithEvents rightConeCheckLabel As System.Windows.Forms.Label
    Friend WithEvents flangeCheckLabel As System.Windows.Forms.Label
    Friend WithEvents PadCheckLabel As System.Windows.Forms.Label
    Friend WithEvents EditPartInfoButton As System.Windows.Forms.Button
    Friend WithEvents DataAcqTmr As System.Windows.Forms.Timer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents PadTextBox As System.Windows.Forms.TextBox
    Friend WithEvents RightConeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents LeftConeTextBoxLabel As System.Windows.Forms.Label
    Friend WithEvents LeftConeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents FlangeTextBox As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CheckZero))
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.CheckZeroButton = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.PadCheckLabel = New System.Windows.Forms.Label
        Me.flangeCheckLabel = New System.Windows.Forms.Label
        Me.rightConeCheckLabel = New System.Windows.Forms.Label
        Me.leftConeCheckLabel = New System.Windows.Forms.Label
        Me.EditPartInfoButton = New System.Windows.Forms.Button
        Me.DataAcqTmr = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.FlangeTextBox = New System.Windows.Forms.TextBox
        Me.PadTextBox = New System.Windows.Forms.TextBox
        Me.RightConeTextBox = New System.Windows.Forms.TextBox
        Me.LeftConeTextBoxLabel = New System.Windows.Forms.Label
        Me.LeftConeTextBox = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.AutoSize = False
        Me.TextBox1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.DarkRed
        Me.TextBox1.Location = New System.Drawing.Point(8, 8)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(560, 56)
        Me.TextBox1.TabIndex = 45
        Me.TextBox1.TabStop = False
        Me.TextBox1.Text = "Remove part from nest and click ""ZERO"" to zero all the gauges. Update part inform" & _
        "ation for next sample at this time."
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'CheckZeroButton
        '
        Me.CheckZeroButton.BackColor = System.Drawing.Color.Transparent
        Me.CheckZeroButton.Enabled = False
        Me.CheckZeroButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckZeroButton.Location = New System.Drawing.Point(176, 0)
        Me.CheckZeroButton.Name = "CheckZeroButton"
        Me.CheckZeroButton.Size = New System.Drawing.Size(278, 48)
        Me.CheckZeroButton.TabIndex = 1
        Me.CheckZeroButton.Text = "CHECK ZERO"
        Me.CheckZeroButton.Visible = False
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(8, 160)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(159, 80)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "OK"
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(168, 160)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(200, 80)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "ZERO"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.PadCheckLabel)
        Me.GroupBox1.Controls.Add(Me.flangeCheckLabel)
        Me.GroupBox1.Controls.Add(Me.rightConeCheckLabel)
        Me.GroupBox1.Controls.Add(Me.leftConeCheckLabel)
        Me.GroupBox1.Location = New System.Drawing.Point(128, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(278, 32)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Visible = False
        '
        'PadCheckLabel
        '
        Me.PadCheckLabel.Enabled = False
        Me.PadCheckLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PadCheckLabel.ForeColor = System.Drawing.Color.DarkGray
        Me.PadCheckLabel.Location = New System.Drawing.Point(8, 66)
        Me.PadCheckLabel.Name = "PadCheckLabel"
        Me.PadCheckLabel.Size = New System.Drawing.Size(264, 24)
        Me.PadCheckLabel.TabIndex = 3
        Me.PadCheckLabel.Text = "Pad Zero Check: "
        Me.PadCheckLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.PadCheckLabel.Visible = False
        '
        'flangeCheckLabel
        '
        Me.flangeCheckLabel.Enabled = False
        Me.flangeCheckLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.flangeCheckLabel.ForeColor = System.Drawing.Color.DarkGray
        Me.flangeCheckLabel.Location = New System.Drawing.Point(8, 93)
        Me.flangeCheckLabel.Name = "flangeCheckLabel"
        Me.flangeCheckLabel.Size = New System.Drawing.Size(264, 24)
        Me.flangeCheckLabel.TabIndex = 2
        Me.flangeCheckLabel.Text = "Flange Zero Check:"
        Me.flangeCheckLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.flangeCheckLabel.Visible = False
        '
        'rightConeCheckLabel
        '
        Me.rightConeCheckLabel.Enabled = False
        Me.rightConeCheckLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rightConeCheckLabel.ForeColor = System.Drawing.Color.DarkGray
        Me.rightConeCheckLabel.Location = New System.Drawing.Point(8, 39)
        Me.rightConeCheckLabel.Name = "rightConeCheckLabel"
        Me.rightConeCheckLabel.Size = New System.Drawing.Size(264, 24)
        Me.rightConeCheckLabel.TabIndex = 1
        Me.rightConeCheckLabel.Text = "Right Cone Zero Check: "
        Me.rightConeCheckLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rightConeCheckLabel.Visible = False
        '
        'leftConeCheckLabel
        '
        Me.leftConeCheckLabel.Enabled = False
        Me.leftConeCheckLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.leftConeCheckLabel.ForeColor = System.Drawing.Color.DarkGray
        Me.leftConeCheckLabel.Location = New System.Drawing.Point(8, 12)
        Me.leftConeCheckLabel.Name = "leftConeCheckLabel"
        Me.leftConeCheckLabel.Size = New System.Drawing.Size(264, 15)
        Me.leftConeCheckLabel.TabIndex = 0
        Me.leftConeCheckLabel.Text = "Left Cone Zero Check: "
        Me.leftConeCheckLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.leftConeCheckLabel.Visible = False
        '
        'EditPartInfoButton
        '
        Me.EditPartInfoButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EditPartInfoButton.Location = New System.Drawing.Point(368, 160)
        Me.EditPartInfoButton.Name = "EditPartInfoButton"
        Me.EditPartInfoButton.Size = New System.Drawing.Size(200, 80)
        Me.EditPartInfoButton.TabIndex = 46
        Me.EditPartInfoButton.Text = "EDIT PART INFO"
        '
        'DataAcqTmr
        '
        Me.DataAcqTmr.Enabled = True
        Me.DataAcqTmr.Interval = 500
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.FlangeTextBox)
        Me.GroupBox2.Controls.Add(Me.PadTextBox)
        Me.GroupBox2.Controls.Add(Me.RightConeTextBox)
        Me.GroupBox2.Controls.Add(Me.LeftConeTextBoxLabel)
        Me.GroupBox2.Controls.Add(Me.LeftConeTextBox)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 64)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(560, 88)
        Me.GroupBox2.TabIndex = 52
        Me.GroupBox2.TabStop = False
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(444, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 59
        Me.Label3.Text = "Flange:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(304, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 58
        Me.Label2.Text = "Pad:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(166, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 57
        Me.Label1.Text = "Right Cone:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'FlangeTextBox
        '
        Me.FlangeTextBox.AutoSize = False
        Me.FlangeTextBox.BackColor = System.Drawing.Color.Black
        Me.FlangeTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FlangeTextBox.ForeColor = System.Drawing.SystemColors.Window
        Me.FlangeTextBox.Location = New System.Drawing.Point(428, 40)
        Me.FlangeTextBox.Name = "FlangeTextBox"
        Me.FlangeTextBox.Size = New System.Drawing.Size(120, 32)
        Me.FlangeTextBox.TabIndex = 56
        Me.FlangeTextBox.Text = "0.00000"
        Me.FlangeTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PadTextBox
        '
        Me.PadTextBox.AutoSize = False
        Me.PadTextBox.BackColor = System.Drawing.Color.Black
        Me.PadTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PadTextBox.ForeColor = System.Drawing.SystemColors.Window
        Me.PadTextBox.Location = New System.Drawing.Point(289, 40)
        Me.PadTextBox.Name = "PadTextBox"
        Me.PadTextBox.Size = New System.Drawing.Size(120, 32)
        Me.PadTextBox.TabIndex = 55
        Me.PadTextBox.Text = "0.00000"
        Me.PadTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'RightConeTextBox
        '
        Me.RightConeTextBox.AutoSize = False
        Me.RightConeTextBox.BackColor = System.Drawing.Color.Black
        Me.RightConeTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RightConeTextBox.ForeColor = System.Drawing.SystemColors.Window
        Me.RightConeTextBox.Location = New System.Drawing.Point(151, 40)
        Me.RightConeTextBox.Name = "RightConeTextBox"
        Me.RightConeTextBox.Size = New System.Drawing.Size(120, 32)
        Me.RightConeTextBox.TabIndex = 54
        Me.RightConeTextBox.Text = "0.00000"
        Me.RightConeTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LeftConeTextBoxLabel
        '
        Me.LeftConeTextBoxLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LeftConeTextBoxLabel.Location = New System.Drawing.Point(26, 16)
        Me.LeftConeTextBoxLabel.Name = "LeftConeTextBoxLabel"
        Me.LeftConeTextBoxLabel.Size = New System.Drawing.Size(88, 16)
        Me.LeftConeTextBoxLabel.TabIndex = 53
        Me.LeftConeTextBoxLabel.Text = " Left Cone:"
        Me.LeftConeTextBoxLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LeftConeTextBox
        '
        Me.LeftConeTextBox.AutoSize = False
        Me.LeftConeTextBox.BackColor = System.Drawing.Color.Black
        Me.LeftConeTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LeftConeTextBox.ForeColor = System.Drawing.SystemColors.Window
        Me.LeftConeTextBox.Location = New System.Drawing.Point(12, 40)
        Me.LeftConeTextBox.Name = "LeftConeTextBox"
        Me.LeftConeTextBox.Size = New System.Drawing.Size(120, 32)
        Me.LeftConeTextBox.TabIndex = 52
        Me.LeftConeTextBox.Text = "0.00000"
        Me.LeftConeTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'CheckZero
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(578, 246)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.EditPartInfoButton)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.CheckZeroButton)
        Me.Controls.Add(Me.TextBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CheckZero"
        Me.Text = "ZERO Prompt"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub CheckZero_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckZeroPageUp = True
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Get data from gauges
        DataAquisition()

        ' Zero the left cone
        zeroLeftConeDbl = rawLeftConeDbl
        ' Zero the right cone
        zeroRightConeDbl = rawRightConeDbl
        ' Zero the pad
        zeroPadDbl = rawPadDbl
        ' Zero the flange
        zeroFlangeDbl = rawFlangeDbl

        'Get data from gauges
        DataAquisition()

        'Update screen to 
    End Sub

    'Private Sub CheckZeroButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckZeroButton.Click
    '    Dim leftConeZeroDiff As Double
    '    Dim rightConeZeroDiff As Double
    '    Dim PadZeroDiff As Double
    '    Dim FlangeZeroDiff As Double
    '    Dim ZeroRangeHigh As Double = 0.00005
    '    Dim ZeroRangeLow As Double = -0.00005
    '     = True
    '    'Get data from gauges
    '    DataAquisition()

    '    'Left Cone Check
    '    leftConeZeroDiff = (zeroLeftConeDbl - rawLeftConeDbl) * 0.5
    '    If (leftConeZeroDiff < ZeroRangeLow) Or (leftConeZeroDiff > ZeroRangeHigh) Then
    '        leftConeCheckLabel.ForeColor = Color.DarkRed
    '        leftConeCheckLabel.Text = "Left Cone Zero Check: FAILED"
    '    Else
    '        leftConeCheckLabel.ForeColor = Color.ForestGreen
    '        leftConeCheckLabel.Text = "Left Cone Zero Check: PASSED"
    '    End If

    '    'Right Cone Check
    '    rightConeZeroDiff = (zeroRightConeDbl - rawRightConeDbl) * 0.5
    '    If (rightConeZeroDiff < ZeroRangeLow) Or (rightConeZeroDiff > ZeroRangeHigh) Then
    '        rightConeCheckLabel.ForeColor = Color.DarkRed
    '        rightConeCheckLabel.Text = "Right Cone Zero Check: FAILED"
    '    Else
    '        rightConeCheckLabel.ForeColor = Color.ForestGreen
    '        rightConeCheckLabel.Text = "Right Cone Zero Check: PASSED"
    '    End If

    '    ' Pad Check
    '    PadZeroDiff = zeroPadDbl - rawPadDbl
    '    If (PadZeroDiff < ZeroRangeLow) Or (PadZeroDiff > ZeroRangeHigh) Then
    '        PadCheckLabel.ForeColor = Color.DarkRed
    '        PadCheckLabel.Text = "Pad Zero Check: FAILED"
    '    Else
    '        PadCheckLabel.ForeColor = Color.ForestGreen
    '        PadCheckLabel.Text = "Pad Zero Check: PASSED"
    '    End If

    '    'Flange Check
    '    FlangeZeroDiff = (zeroFlangeDbl / 1000) - (rawFlangeDbl / 1000)
    '    If (FlangeZeroDiff < ZeroRangeLow) Or (FlangeZeroDiff > ZeroRangeHigh) Then
    '        flangeCheckLabel.ForeColor = Color.DarkRed
    '        flangeCheckLabel.Text = "Flange Zero Check: FAILED"
    '    Else
    '        flangeCheckLabel.ForeColor = Color.ForestGreen
    '        flangeCheckLabel.Text = "Pad Zero Check: PASSED"
    '    End If
    '    CheckZeroPageUp = False
    'End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.DataAcqTmr.Enabled = False
        CheckZeroPageUp = False
        boolCycleComplete = True
        disableManualButton = False
        AutoPartCount = 0
        SuspendDOUT = False
        Me.Close()                  ' Close part data form
    End Sub

    Private Sub EditPartInfoButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditPartInfoButton.Click
        ' Create an instance of PartData
        Dim newform08 As PartData = New PartData
        newform08.Show()              'show instance of PartData
    End Sub

    Private Sub DataAcqTmr_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataAcqTmr.Tick

        'If (laserDisplacementERROR = False) Then
        'Get Data
        DataAquisition()
        'Update check zero datafields
        updateDataFields()
        'End If
    End Sub
    Private Sub updateDataFields()
        ' Write data to the Form
        If port02Head01Out02Err Then
            LeftConeTextBox.Text = "xxxxxxx"
        Else
            LeftConeTextBox.Text = Format(leftConeDbl, "0.00000")
        End If
        If port02Head01Out03Err Then
            RightConeTextBox.Text = "xxxxxxx"
        Else
            RightConeTextBox.Text = Format(rightConeDbl, "0.00000")
        End If
        If port02Head02Out04Err Then
            PadTextBox.Text = "xxxxxxx"
        Else
            PadTextBox.Text = Format(padDbl, "0.00000")
        End If

        If laserDisplacementERROR Or port05Head01Out01Err Then
            FlangeTextBox.Text = "xxxxxxx"
        Else
            FlangeTextBox.Text = Format(flangeDbl, "0.00000")
        End If
    End Sub
End Class
