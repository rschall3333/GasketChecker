Imports System.Convert
Imports System

Public Class Main
    Inherits System.Windows.Forms.Form

    Dim DataTranslationDIN As Object
    Dim DataTranslationDOUT As Integer

    Dim CycleComplete As Integer = 127          'Constant for cycle complete
    Dim CycleRunning As Integer = 126           'Constant for cycle running
    Dim VacuumFault As Integer = 123            'Constant for vacuum fault
    Dim TakeMeasurement As Integer = 122        'Constant for take measurement
    Dim laserDisplacementErrorLim As Integer = 200
    
    Dim leftConeOutOfZero As Boolean
    Dim rightConeOutOfZero As Boolean
    Dim PadOutOfZero As Boolean
    Dim FlangeOutOfZero As Boolean

    Dim leftConeZeroCompare As Double
    Dim rightConeZeroCompare As Double
    Dim PadZeroCompare As Double
    Dim FlangeZeroCompare As Double
    Dim zeroValuesInTolerance As Boolean        ' 0 - error, 1 - ok
    Dim startTime As Integer

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
    Friend WithEvents MainPanel As System.Windows.Forms.Panel
    Friend WithEvents MainBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ContactGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents Picture01GroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ContactLabel As System.Windows.Forms.Label
    Friend WithEvents InfoBox As System.Windows.Forms.GroupBox
    Friend WithEvents InfoPanel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TimeGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblHour As System.Windows.Forms.Label
    Friend WithEvents lblColon As System.Windows.Forms.Label
    Friend WithEvents lblMinute As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblMonth As System.Windows.Forms.Label
    Friend WithEvents lblDay As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Timers.Timer
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents lblMinuteZeroFill As System.Windows.Forms.Label
    Friend WithEvents lblMinuteSingleDigit As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents DataBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DataBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents PadTextBox As System.Windows.Forms.TextBox
    Friend WithEvents FlangeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents RightConeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents LeftConeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents DIN As AxDTAcq32Lib.AxDTAcq32
    Friend WithEvents DOUT As AxDTAcq32Lib.AxDTAcq32
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents StatusTextBox As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents JobOrderLabel As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents VersionLabel As System.Windows.Forms.Label
    Friend WithEvents DateLabel As System.Windows.Forms.Label
    Friend WithEvents AutoManualBox As System.Windows.Forms.GroupBox
    Friend WithEvents ManualButton As System.Windows.Forms.Button
    Friend WithEvents AutoButton As System.Windows.Forms.Button
    Friend WithEvents PartControlGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents PartNumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents OperatorIdentificationTextBox As System.Windows.Forms.TextBox
    Friend WithEvents MoldNumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents OperatorIDLabel As System.Windows.Forms.Label
    Friend WithEvents MoldNumberLabel As System.Windows.Forms.Label
    Friend WithEvents CavityNumberLabel As System.Windows.Forms.Label
    Friend WithEvents CavityNumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents HeatNumberLabel As System.Windows.Forms.Label
    Friend WithEvents HeatNumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PartDateLabel As System.Windows.Forms.Label
    Friend WithEvents PartDateTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents PartNumberLabel As System.Windows.Forms.Label
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents zeroButton As System.Windows.Forms.Button
    Friend WithEvents RunNumberLabel As System.Windows.Forms.Label
    Friend WithEvents RunNumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PartTimeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PartTimeLabel As System.Windows.Forms.Label
    Friend WithEvents ShopOrderTextBox As System.Windows.Forms.TextBox
    Friend WithEvents ShopOrderLabel As System.Windows.Forms.Label
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents addNoteToDataStrButton As System.Windows.Forms.Button
    Friend WithEvents AddNoteProdReportButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents PartCountLabel As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main))
        Me.MainPanel = New System.Windows.Forms.Panel
        Me.MainBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.PartControlGroupBox = New System.Windows.Forms.GroupBox
        Me.ShopOrderLabel = New System.Windows.Forms.Label
        Me.ShopOrderTextBox = New System.Windows.Forms.TextBox
        Me.PartTimeLabel = New System.Windows.Forms.Label
        Me.PartTimeTextBox = New System.Windows.Forms.TextBox
        Me.RunNumberTextBox = New System.Windows.Forms.TextBox
        Me.RunNumberLabel = New System.Windows.Forms.Label
        Me.PartNumberLabel = New System.Windows.Forms.Label
        Me.zeroButton = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.PartDateLabel = New System.Windows.Forms.Label
        Me.PartDateTextBox = New System.Windows.Forms.TextBox
        Me.HeatNumberLabel = New System.Windows.Forms.Label
        Me.HeatNumberTextBox = New System.Windows.Forms.TextBox
        Me.CavityNumberLabel = New System.Windows.Forms.Label
        Me.CavityNumberTextBox = New System.Windows.Forms.TextBox
        Me.MoldNumberLabel = New System.Windows.Forms.Label
        Me.OperatorIDLabel = New System.Windows.Forms.Label
        Me.MoldNumberTextBox = New System.Windows.Forms.TextBox
        Me.OperatorIdentificationTextBox = New System.Windows.Forms.TextBox
        Me.PartNumberTextBox = New System.Windows.Forms.TextBox
        Me.AutoManualBox = New System.Windows.Forms.GroupBox
        Me.ManualButton = New System.Windows.Forms.Button
        Me.AutoButton = New System.Windows.Forms.Button
        Me.DataBox2 = New System.Windows.Forms.GroupBox
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.PartCountLabel = New System.Windows.Forms.Label
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.AddNoteProdReportButton = New System.Windows.Forms.Button
        Me.addNoteToDataStrButton = New System.Windows.Forms.Button
        Me.StatusTextBox = New System.Windows.Forms.TextBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.DOUT = New AxDTAcq32Lib.AxDTAcq32
        Me.DIN = New AxDTAcq32Lib.AxDTAcq32
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.FlangeTextBox = New System.Windows.Forms.TextBox
        Me.PadTextBox = New System.Windows.Forms.TextBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.RightConeTextBox = New System.Windows.Forms.TextBox
        Me.LeftConeTextBox = New System.Windows.Forms.TextBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ContactGroupBox = New System.Windows.Forms.GroupBox
        Me.TimeGroupBox = New System.Windows.Forms.GroupBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.lblMonth = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblDay = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblYear = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.lblMinuteSingleDigit = New System.Windows.Forms.Label
        Me.lblMinuteZeroFill = New System.Windows.Forms.Label
        Me.lblMinute = New System.Windows.Forms.Label
        Me.lblHour = New System.Windows.Forms.Label
        Me.lblColon = New System.Windows.Forms.Label
        Me.InfoBox = New System.Windows.Forms.GroupBox
        Me.InfoPanel1 = New System.Windows.Forms.Panel
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.VersionLabel = New System.Windows.Forms.Label
        Me.DateLabel = New System.Windows.Forms.Label
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.JobOrderLabel = New System.Windows.Forms.Label
        Me.ContactLabel = New System.Windows.Forms.Label
        Me.Picture01GroupBox = New System.Windows.Forms.GroupBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Timer1 = New System.Timers.Timer
        Me.DataBox1 = New System.Windows.Forms.GroupBox
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.MainPanel.SuspendLayout()
        Me.MainBox1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.PartControlGroupBox.SuspendLayout()
        Me.AutoManualBox.SuspendLayout()
        Me.DataBox2.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        CType(Me.DOUT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DIN, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContactGroupBox.SuspendLayout()
        Me.TimeGroupBox.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.InfoBox.SuspendLayout()
        Me.InfoPanel1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.Picture01GroupBox.SuspendLayout()
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainPanel
        '
        Me.MainPanel.BackColor = System.Drawing.Color.Silver
        Me.MainPanel.Controls.Add(Me.MainBox1)
        Me.MainPanel.Location = New System.Drawing.Point(10, 10)
        Me.MainPanel.Name = "MainPanel"
        Me.MainPanel.Size = New System.Drawing.Size(1004, 730)
        Me.MainPanel.TabIndex = 0
        '
        'MainBox1
        '
        Me.MainBox1.BackColor = System.Drawing.Color.Silver
        Me.MainBox1.Controls.Add(Me.GroupBox1)
        Me.MainBox1.Controls.Add(Me.DataBox2)
        Me.MainBox1.Controls.Add(Me.ContactGroupBox)
        Me.MainBox1.Location = New System.Drawing.Point(10, 10)
        Me.MainBox1.Name = "MainBox1"
        Me.MainBox1.Size = New System.Drawing.Size(984, 706)
        Me.MainBox1.TabIndex = 0
        Me.MainBox1.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.PartControlGroupBox)
        Me.GroupBox1.Controls.Add(Me.AutoManualBox)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Location = New System.Drawing.Point(10, 10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(964, 122)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'PartControlGroupBox
        '
        Me.PartControlGroupBox.BackColor = System.Drawing.Color.Transparent
        Me.PartControlGroupBox.Controls.Add(Me.ShopOrderLabel)
        Me.PartControlGroupBox.Controls.Add(Me.ShopOrderTextBox)
        Me.PartControlGroupBox.Controls.Add(Me.PartTimeLabel)
        Me.PartControlGroupBox.Controls.Add(Me.PartTimeTextBox)
        Me.PartControlGroupBox.Controls.Add(Me.RunNumberTextBox)
        Me.PartControlGroupBox.Controls.Add(Me.RunNumberLabel)
        Me.PartControlGroupBox.Controls.Add(Me.PartNumberLabel)
        Me.PartControlGroupBox.Controls.Add(Me.zeroButton)
        Me.PartControlGroupBox.Controls.Add(Me.Button2)
        Me.PartControlGroupBox.Controls.Add(Me.Button1)
        Me.PartControlGroupBox.Controls.Add(Me.PartDateLabel)
        Me.PartControlGroupBox.Controls.Add(Me.PartDateTextBox)
        Me.PartControlGroupBox.Controls.Add(Me.HeatNumberLabel)
        Me.PartControlGroupBox.Controls.Add(Me.HeatNumberTextBox)
        Me.PartControlGroupBox.Controls.Add(Me.CavityNumberLabel)
        Me.PartControlGroupBox.Controls.Add(Me.CavityNumberTextBox)
        Me.PartControlGroupBox.Controls.Add(Me.MoldNumberLabel)
        Me.PartControlGroupBox.Controls.Add(Me.OperatorIDLabel)
        Me.PartControlGroupBox.Controls.Add(Me.MoldNumberTextBox)
        Me.PartControlGroupBox.Controls.Add(Me.OperatorIdentificationTextBox)
        Me.PartControlGroupBox.Controls.Add(Me.PartNumberTextBox)
        Me.PartControlGroupBox.Location = New System.Drawing.Point(190, 5)
        Me.PartControlGroupBox.Name = "PartControlGroupBox"
        Me.PartControlGroupBox.Size = New System.Drawing.Size(768, 112)
        Me.PartControlGroupBox.TabIndex = 14
        Me.PartControlGroupBox.TabStop = False
        '
        'ShopOrderLabel
        '
        Me.ShopOrderLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ShopOrderLabel.Location = New System.Drawing.Point(540, 15)
        Me.ShopOrderLabel.Name = "ShopOrderLabel"
        Me.ShopOrderLabel.Size = New System.Drawing.Size(85, 22)
        Me.ShopOrderLabel.TabIndex = 46
        Me.ShopOrderLabel.Text = "Shop Order:"
        Me.ShopOrderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ShopOrderTextBox
        '
        Me.ShopOrderTextBox.BackColor = System.Drawing.Color.LightGray
        Me.ShopOrderTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ShopOrderTextBox.Location = New System.Drawing.Point(625, 15)
        Me.ShopOrderTextBox.Name = "ShopOrderTextBox"
        Me.ShopOrderTextBox.ReadOnly = True
        Me.ShopOrderTextBox.Size = New System.Drawing.Size(135, 20)
        Me.ShopOrderTextBox.TabIndex = 45
        Me.ShopOrderTextBox.TabStop = False
        Me.ShopOrderTextBox.Text = "Shop Order Number"
        '
        'PartTimeLabel
        '
        Me.PartTimeLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartTimeLabel.Location = New System.Drawing.Point(540, 46)
        Me.PartTimeLabel.Name = "PartTimeLabel"
        Me.PartTimeLabel.Size = New System.Drawing.Size(85, 22)
        Me.PartTimeLabel.TabIndex = 44
        Me.PartTimeLabel.Text = "Sample Time:"
        Me.PartTimeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PartTimeTextBox
        '
        Me.PartTimeTextBox.BackColor = System.Drawing.Color.LightGray
        Me.PartTimeTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartTimeTextBox.Location = New System.Drawing.Point(625, 46)
        Me.PartTimeTextBox.Name = "PartTimeTextBox"
        Me.PartTimeTextBox.ReadOnly = True
        Me.PartTimeTextBox.Size = New System.Drawing.Size(135, 20)
        Me.PartTimeTextBox.TabIndex = 43
        Me.PartTimeTextBox.TabStop = False
        Me.PartTimeTextBox.Text = "Time"
        '
        'RunNumberTextBox
        '
        Me.RunNumberTextBox.BackColor = System.Drawing.Color.LightGray
        Me.RunNumberTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunNumberTextBox.Location = New System.Drawing.Point(625, 76)
        Me.RunNumberTextBox.Name = "RunNumberTextBox"
        Me.RunNumberTextBox.ReadOnly = True
        Me.RunNumberTextBox.Size = New System.Drawing.Size(135, 20)
        Me.RunNumberTextBox.TabIndex = 42
        Me.RunNumberTextBox.TabStop = False
        Me.RunNumberTextBox.Text = "Run"
        '
        'RunNumberLabel
        '
        Me.RunNumberLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunNumberLabel.Location = New System.Drawing.Point(540, 76)
        Me.RunNumberLabel.Name = "RunNumberLabel"
        Me.RunNumberLabel.Size = New System.Drawing.Size(85, 22)
        Me.RunNumberLabel.TabIndex = 41
        Me.RunNumberLabel.Text = "Run:"
        Me.RunNumberLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PartNumberLabel
        '
        Me.PartNumberLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartNumberLabel.Location = New System.Drawing.Point(310, 15)
        Me.PartNumberLabel.Name = "PartNumberLabel"
        Me.PartNumberLabel.Size = New System.Drawing.Size(80, 22)
        Me.PartNumberLabel.TabIndex = 40
        Me.PartNumberLabel.Text = "Part Number:"
        Me.PartNumberLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'zeroButton
        '
        Me.zeroButton.BackColor = System.Drawing.Color.Silver
        Me.zeroButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.zeroButton.ForeColor = System.Drawing.Color.Black
        Me.zeroButton.Location = New System.Drawing.Point(155, 13)
        Me.zeroButton.Name = "zeroButton"
        Me.zeroButton.Size = New System.Drawing.Size(150, 91)
        Me.zeroButton.TabIndex = 36
        Me.zeroButton.TabStop = False
        Me.zeroButton.Text = "ZERO"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Silver
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(5, 59)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(150, 45)
        Me.Button2.TabIndex = 35
        Me.Button2.TabStop = False
        Me.Button2.Text = "DATA COLLECTION"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(5, 13)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(150, 45)
        Me.Button1.TabIndex = 34
        Me.Button1.TabStop = False
        Me.Button1.Text = "EDIT PART INFO"
        '
        'PartDateLabel
        '
        Me.PartDateLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartDateLabel.Location = New System.Drawing.Point(540, 76)
        Me.PartDateLabel.Name = "PartDateLabel"
        Me.PartDateLabel.Size = New System.Drawing.Size(85, 22)
        Me.PartDateLabel.TabIndex = 33
        Me.PartDateLabel.Text = "Sample Date:"
        Me.PartDateLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PartDateTextBox
        '
        Me.PartDateTextBox.BackColor = System.Drawing.Color.LightGray
        Me.PartDateTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartDateTextBox.Location = New System.Drawing.Point(625, 76)
        Me.PartDateTextBox.Name = "PartDateTextBox"
        Me.PartDateTextBox.ReadOnly = True
        Me.PartDateTextBox.Size = New System.Drawing.Size(135, 20)
        Me.PartDateTextBox.TabIndex = 32
        Me.PartDateTextBox.TabStop = False
        Me.PartDateTextBox.Text = "Date"
        '
        'HeatNumberLabel
        '
        Me.HeatNumberLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HeatNumberLabel.Location = New System.Drawing.Point(540, 46)
        Me.HeatNumberLabel.Name = "HeatNumberLabel"
        Me.HeatNumberLabel.Size = New System.Drawing.Size(85, 22)
        Me.HeatNumberLabel.TabIndex = 31
        Me.HeatNumberLabel.Text = "Heat Number:"
        Me.HeatNumberLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'HeatNumberTextBox
        '
        Me.HeatNumberTextBox.BackColor = System.Drawing.Color.LightGray
        Me.HeatNumberTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HeatNumberTextBox.Location = New System.Drawing.Point(625, 46)
        Me.HeatNumberTextBox.Name = "HeatNumberTextBox"
        Me.HeatNumberTextBox.ReadOnly = True
        Me.HeatNumberTextBox.Size = New System.Drawing.Size(135, 20)
        Me.HeatNumberTextBox.TabIndex = 30
        Me.HeatNumberTextBox.TabStop = False
        Me.HeatNumberTextBox.Text = "Heat Number"
        '
        'CavityNumberLabel
        '
        Me.CavityNumberLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CavityNumberLabel.Location = New System.Drawing.Point(540, 15)
        Me.CavityNumberLabel.Name = "CavityNumberLabel"
        Me.CavityNumberLabel.Size = New System.Drawing.Size(85, 22)
        Me.CavityNumberLabel.TabIndex = 29
        Me.CavityNumberLabel.Text = "Cavity Number:"
        Me.CavityNumberLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CavityNumberTextBox
        '
        Me.CavityNumberTextBox.BackColor = System.Drawing.Color.LightGray
        Me.CavityNumberTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CavityNumberTextBox.Location = New System.Drawing.Point(625, 15)
        Me.CavityNumberTextBox.Name = "CavityNumberTextBox"
        Me.CavityNumberTextBox.ReadOnly = True
        Me.CavityNumberTextBox.Size = New System.Drawing.Size(135, 20)
        Me.CavityNumberTextBox.TabIndex = 28
        Me.CavityNumberTextBox.TabStop = False
        Me.CavityNumberTextBox.Text = "Cavity Number"
        '
        'MoldNumberLabel
        '
        Me.MoldNumberLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MoldNumberLabel.Location = New System.Drawing.Point(310, 76)
        Me.MoldNumberLabel.Name = "MoldNumberLabel"
        Me.MoldNumberLabel.Size = New System.Drawing.Size(80, 22)
        Me.MoldNumberLabel.TabIndex = 27
        Me.MoldNumberLabel.Text = "Mold Number:"
        Me.MoldNumberLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'OperatorIDLabel
        '
        Me.OperatorIDLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OperatorIDLabel.Location = New System.Drawing.Point(310, 46)
        Me.OperatorIDLabel.Name = "OperatorIDLabel"
        Me.OperatorIDLabel.Size = New System.Drawing.Size(80, 22)
        Me.OperatorIDLabel.TabIndex = 26
        Me.OperatorIDLabel.Text = "Operator ID:"
        Me.OperatorIDLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'MoldNumberTextBox
        '
        Me.MoldNumberTextBox.BackColor = System.Drawing.Color.LightGray
        Me.MoldNumberTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MoldNumberTextBox.Location = New System.Drawing.Point(390, 76)
        Me.MoldNumberTextBox.Name = "MoldNumberTextBox"
        Me.MoldNumberTextBox.ReadOnly = True
        Me.MoldNumberTextBox.Size = New System.Drawing.Size(140, 20)
        Me.MoldNumberTextBox.TabIndex = 24
        Me.MoldNumberTextBox.TabStop = False
        Me.MoldNumberTextBox.Text = "Mold Number"
        '
        'OperatorIdentificationTextBox
        '
        Me.OperatorIdentificationTextBox.BackColor = System.Drawing.Color.LightGray
        Me.OperatorIdentificationTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OperatorIdentificationTextBox.Location = New System.Drawing.Point(390, 46)
        Me.OperatorIdentificationTextBox.Name = "OperatorIdentificationTextBox"
        Me.OperatorIdentificationTextBox.ReadOnly = True
        Me.OperatorIdentificationTextBox.Size = New System.Drawing.Size(140, 20)
        Me.OperatorIdentificationTextBox.TabIndex = 23
        Me.OperatorIdentificationTextBox.TabStop = False
        Me.OperatorIdentificationTextBox.Text = "Operator Identification"
        '
        'PartNumberTextBox
        '
        Me.PartNumberTextBox.BackColor = System.Drawing.Color.LightGray
        Me.PartNumberTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartNumberTextBox.Location = New System.Drawing.Point(390, 16)
        Me.PartNumberTextBox.Name = "PartNumberTextBox"
        Me.PartNumberTextBox.ReadOnly = True
        Me.PartNumberTextBox.Size = New System.Drawing.Size(140, 20)
        Me.PartNumberTextBox.TabIndex = 22
        Me.PartNumberTextBox.TabStop = False
        Me.PartNumberTextBox.Text = "Part Number"
        '
        'AutoManualBox
        '
        Me.AutoManualBox.BackColor = System.Drawing.Color.Transparent
        Me.AutoManualBox.Controls.Add(Me.ManualButton)
        Me.AutoManualBox.Controls.Add(Me.AutoButton)
        Me.AutoManualBox.Location = New System.Drawing.Point(5, 5)
        Me.AutoManualBox.Name = "AutoManualBox"
        Me.AutoManualBox.Size = New System.Drawing.Size(180, 112)
        Me.AutoManualBox.TabIndex = 13
        Me.AutoManualBox.TabStop = False
        '
        'ManualButton
        '
        Me.ManualButton.BackColor = System.Drawing.Color.Transparent
        Me.ManualButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ManualButton.ForeColor = System.Drawing.Color.Black
        Me.ManualButton.Location = New System.Drawing.Point(7, 57)
        Me.ManualButton.Name = "ManualButton"
        Me.ManualButton.Size = New System.Drawing.Size(165, 45)
        Me.ManualButton.TabIndex = 0
        Me.ManualButton.Text = "MANUAL"
        '
        'AutoButton
        '
        Me.AutoButton.BackColor = System.Drawing.Color.Gray
        Me.AutoButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.AutoButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AutoButton.ForeColor = System.Drawing.Color.Lime
        Me.AutoButton.Location = New System.Drawing.Point(7, 11)
        Me.AutoButton.Name = "AutoButton"
        Me.AutoButton.Size = New System.Drawing.Size(165, 45)
        Me.AutoButton.TabIndex = 0
        Me.AutoButton.Text = "AUTO"
        '
        'DataBox2
        '
        Me.DataBox2.BackColor = System.Drawing.Color.Silver
        Me.DataBox2.Controls.Add(Me.Panel4)
        Me.DataBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataBox2.Location = New System.Drawing.Point(10, 110)
        Me.DataBox2.Name = "DataBox2"
        Me.DataBox2.Size = New System.Drawing.Size(964, 425)
        Me.DataBox2.TabIndex = 1
        Me.DataBox2.TabStop = False
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.DarkGray
        Me.Panel4.Controls.Add(Me.GroupBox7)
        Me.Panel4.Controls.Add(Me.GroupBox6)
        Me.Panel4.Controls.Add(Me.StatusTextBox)
        Me.Panel4.Controls.Add(Me.ProgressBar1)
        Me.Panel4.Controls.Add(Me.DOUT)
        Me.Panel4.Controls.Add(Me.DIN)
        Me.Panel4.Controls.Add(Me.Label6)
        Me.Panel4.Controls.Add(Me.Label5)
        Me.Panel4.Controls.Add(Me.FlangeTextBox)
        Me.Panel4.Controls.Add(Me.PadTextBox)
        Me.Panel4.Controls.Add(Me.PictureBox3)
        Me.Panel4.Controls.Add(Me.Label4)
        Me.Panel4.Controls.Add(Me.RightConeTextBox)
        Me.Panel4.Controls.Add(Me.LeftConeTextBox)
        Me.Panel4.Controls.Add(Me.PictureBox2)
        Me.Panel4.Controls.Add(Me.Label1)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel4.Location = New System.Drawing.Point(3, 30)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(958, 392)
        Me.Panel4.TabIndex = 3
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.PartCountLabel)
        Me.GroupBox7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox7.Location = New System.Drawing.Point(8, 327)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(129, 38)
        Me.GroupBox7.TabIndex = 50
        Me.GroupBox7.TabStop = False
        '
        'PartCountLabel
        '
        Me.PartCountLabel.BackColor = System.Drawing.Color.Transparent
        Me.PartCountLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartCountLabel.Location = New System.Drawing.Point(15, 12)
        Me.PartCountLabel.Name = "PartCountLabel"
        Me.PartCountLabel.Size = New System.Drawing.Size(95, 20)
        Me.PartCountLabel.TabIndex = 35
        Me.PartCountLabel.Text = "Part:"
        Me.PartCountLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.AddNoteProdReportButton)
        Me.GroupBox6.Controls.Add(Me.addNoteToDataStrButton)
        Me.GroupBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox6.Location = New System.Drawing.Point(777, 5)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(170, 116)
        Me.GroupBox6.TabIndex = 49
        Me.GroupBox6.TabStop = False
        '
        'AddNoteProdReportButton
        '
        Me.AddNoteProdReportButton.BackColor = System.Drawing.Color.Silver
        Me.AddNoteProdReportButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddNoteProdReportButton.ForeColor = System.Drawing.Color.Black
        Me.AddNoteProdReportButton.Location = New System.Drawing.Point(10, 61)
        Me.AddNoteProdReportButton.Name = "AddNoteProdReportButton"
        Me.AddNoteProdReportButton.Size = New System.Drawing.Size(150, 45)
        Me.AddNoteProdReportButton.TabIndex = 49
        Me.AddNoteProdReportButton.TabStop = False
        Me.AddNoteProdReportButton.Text = "ADD NOTE TO PRODUCTION REPORT"
        '
        'addNoteToDataStrButton
        '
        Me.addNoteToDataStrButton.BackColor = System.Drawing.Color.Silver
        Me.addNoteToDataStrButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.addNoteToDataStrButton.ForeColor = System.Drawing.Color.Black
        Me.addNoteToDataStrButton.Location = New System.Drawing.Point(10, 15)
        Me.addNoteToDataStrButton.Name = "addNoteToDataStrButton"
        Me.addNoteToDataStrButton.Size = New System.Drawing.Size(150, 45)
        Me.addNoteToDataStrButton.TabIndex = 48
        Me.addNoteToDataStrButton.TabStop = False
        Me.addNoteToDataStrButton.Text = "ADD NOTE TO STREAMING DATA FILE"
        '
        'StatusTextBox
        '
        Me.StatusTextBox.BackColor = System.Drawing.Color.LightGray
        Me.StatusTextBox.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.StatusTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusTextBox.ForeColor = System.Drawing.Color.DarkBlue
        Me.StatusTextBox.Location = New System.Drawing.Point(0, 347)
        Me.StatusTextBox.Name = "StatusTextBox"
        Me.StatusTextBox.Size = New System.Drawing.Size(958, 22)
        Me.StatusTextBox.TabIndex = 18
        Me.StatusTextBox.TabStop = False
        Me.StatusTextBox.Text = "TextBox1"
        Me.StatusTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 369)
        Me.ProgressBar1.Maximum = 4000
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(958, 23)
        Me.ProgressBar1.Step = 95
        Me.ProgressBar1.TabIndex = 17
        '
        'DOUT
        '
        Me.DOUT.ContainingControl = Me
        Me.DOUT.Enabled = True
        Me.DOUT.Location = New System.Drawing.Point(604, 82)
        Me.DOUT.Name = "DOUT"
        Me.DOUT.OcxState = CType(resources.GetObject("DOUT.OcxState"), System.Windows.Forms.AxHost.State)
        Me.DOUT.Size = New System.Drawing.Size(100, 50)
        Me.DOUT.TabIndex = 16
        Me.DOUT.TabStop = False
        '
        'DIN
        '
        Me.DIN.ContainingControl = Me
        Me.DIN.Enabled = True
        Me.DIN.Location = New System.Drawing.Point(93, 176)
        Me.DIN.Name = "DIN"
        Me.DIN.OcxState = CType(resources.GetObject("DIN.OcxState"), System.Windows.Forms.AxHost.State)
        Me.DIN.Size = New System.Drawing.Size(100, 50)
        Me.DIN.TabIndex = 15
        Me.DIN.TabStop = False
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label6.Location = New System.Drawing.Point(732, 270)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(170, 20)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "(Flange 0.0105"" +-0.0015"") "
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label5.Location = New System.Drawing.Point(362, 300)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(155, 20)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "(Pad 0.0165"" +-0.0015"") "
        '
        'FlangeTextBox
        '
        Me.FlangeTextBox.AutoSize = False
        Me.FlangeTextBox.BackColor = System.Drawing.Color.LightGray
        Me.FlangeTextBox.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.FlangeTextBox.ForeColor = System.Drawing.Color.ForestGreen
        Me.FlangeTextBox.Location = New System.Drawing.Point(717, 210)
        Me.FlangeTextBox.Name = "FlangeTextBox"
        Me.FlangeTextBox.ReadOnly = True
        Me.FlangeTextBox.Size = New System.Drawing.Size(200, 55)
        Me.FlangeTextBox.TabIndex = 33
        Me.FlangeTextBox.TabStop = False
        Me.FlangeTextBox.Text = "0.01653"
        Me.FlangeTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PadTextBox
        '
        Me.PadTextBox.AutoSize = False
        Me.PadTextBox.BackColor = System.Drawing.Color.LightGray
        Me.PadTextBox.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.PadTextBox.ForeColor = System.Drawing.Color.ForestGreen
        Me.PadTextBox.Location = New System.Drawing.Point(342, 240)
        Me.PadTextBox.Name = "PadTextBox"
        Me.PadTextBox.ReadOnly = True
        Me.PadTextBox.Size = New System.Drawing.Size(200, 55)
        Me.PadTextBox.TabIndex = 33
        Me.PadTextBox.TabStop = False
        Me.PadTextBox.Text = "0.01657"
        Me.PadTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(572, 225)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(165, 115)
        Me.PictureBox3.TabIndex = 5
        Me.PictureBox3.TabStop = False
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label4.Location = New System.Drawing.Point(425, 98)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(200, 20)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "(Right Cone 0.0065"" +-0.0015"") "
        '
        'RightConeTextBox
        '
        Me.RightConeTextBox.AutoSize = False
        Me.RightConeTextBox.BackColor = System.Drawing.Color.Yellow
        Me.RightConeTextBox.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.RightConeTextBox.ForeColor = System.Drawing.Color.Red
        Me.RightConeTextBox.Location = New System.Drawing.Point(410, 38)
        Me.RightConeTextBox.Name = "RightConeTextBox"
        Me.RightConeTextBox.ReadOnly = True
        Me.RightConeTextBox.Size = New System.Drawing.Size(200, 55)
        Me.RightConeTextBox.TabIndex = 33
        Me.RightConeTextBox.TabStop = False
        Me.RightConeTextBox.Text = "0.00812"
        Me.RightConeTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LeftConeTextBox
        '
        Me.LeftConeTextBox.AutoSize = False
        Me.LeftConeTextBox.BackColor = System.Drawing.Color.LightGray
        Me.LeftConeTextBox.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.LeftConeTextBox.ForeColor = System.Drawing.Color.ForestGreen
        Me.LeftConeTextBox.Location = New System.Drawing.Point(35, 38)
        Me.LeftConeTextBox.Name = "LeftConeTextBox"
        Me.LeftConeTextBox.ReadOnly = True
        Me.LeftConeTextBox.Size = New System.Drawing.Size(200, 55)
        Me.LeftConeTextBox.TabIndex = 33
        Me.LeftConeTextBox.TabStop = False
        Me.LeftConeTextBox.Text = "0.00681"
        Me.LeftConeTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(240, 63)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(210, 100)
        Me.PictureBox2.TabIndex = 0
        Me.PictureBox2.TabStop = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label1.Location = New System.Drawing.Point(50, 98)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(188, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "(Left Cone 0.0065"" +-0.0015"") "
        '
        'ContactGroupBox
        '
        Me.ContactGroupBox.Controls.Add(Me.TimeGroupBox)
        Me.ContactGroupBox.Controls.Add(Me.InfoBox)
        Me.ContactGroupBox.Controls.Add(Me.ContactLabel)
        Me.ContactGroupBox.Controls.Add(Me.Picture01GroupBox)
        Me.ContactGroupBox.Location = New System.Drawing.Point(10, 535)
        Me.ContactGroupBox.Name = "ContactGroupBox"
        Me.ContactGroupBox.Size = New System.Drawing.Size(964, 155)
        Me.ContactGroupBox.TabIndex = 0
        Me.ContactGroupBox.TabStop = False
        '
        'TimeGroupBox
        '
        Me.TimeGroupBox.Controls.Add(Me.Panel1)
        Me.TimeGroupBox.Location = New System.Drawing.Point(9, 10)
        Me.TimeGroupBox.Name = "TimeGroupBox"
        Me.TimeGroupBox.Size = New System.Drawing.Size(180, 135)
        Me.TimeGroupBox.TabIndex = 3
        Me.TimeGroupBox.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkGray
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Location = New System.Drawing.Point(2, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(174, 123)
        Me.Panel1.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Panel3)
        Me.GroupBox3.Location = New System.Drawing.Point(5, 65)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(165, 35)
        Me.GroupBox3.TabIndex = 10
        Me.GroupBox3.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.Controls.Add(Me.lblMonth)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.lblDay)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.lblYear)
        Me.Panel3.Location = New System.Drawing.Point(5, 9)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(154, 22)
        Me.Panel3.TabIndex = 0
        '
        'lblMonth
        '
        Me.lblMonth.BackColor = System.Drawing.Color.DarkGray
        Me.lblMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMonth.ForeColor = System.Drawing.Color.Black
        Me.lblMonth.Location = New System.Drawing.Point(0, 4)
        Me.lblMonth.Name = "lblMonth"
        Me.lblMonth.Size = New System.Drawing.Size(36, 15)
        Me.lblMonth.TabIndex = 3
        Me.lblMonth.Text = "XX"
        Me.lblMonth.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.DarkGray
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(39, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(10, 15)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "/"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDay
        '
        Me.lblDay.BackColor = System.Drawing.Color.DarkGray
        Me.lblDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDay.ForeColor = System.Drawing.Color.Black
        Me.lblDay.Location = New System.Drawing.Point(45, 4)
        Me.lblDay.Name = "lblDay"
        Me.lblDay.Size = New System.Drawing.Size(40, 15)
        Me.lblDay.TabIndex = 5
        Me.lblDay.Text = "XX"
        Me.lblDay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.DarkGray
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(82, 4)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(10, 15)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "/"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblYear
        '
        Me.lblYear.BackColor = System.Drawing.Color.DarkGray
        Me.lblYear.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear.ForeColor = System.Drawing.Color.Black
        Me.lblYear.Location = New System.Drawing.Point(93, 4)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(65, 15)
        Me.lblYear.TabIndex = 6
        Me.lblYear.Text = "XXXX"
        Me.lblYear.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Panel2)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 15)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(165, 35)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkGray
        Me.Panel2.Controls.Add(Me.lblMinuteSingleDigit)
        Me.Panel2.Controls.Add(Me.lblMinuteZeroFill)
        Me.Panel2.Controls.Add(Me.lblMinute)
        Me.Panel2.Controls.Add(Me.lblHour)
        Me.Panel2.Controls.Add(Me.lblColon)
        Me.Panel2.Location = New System.Drawing.Point(5, 9)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(154, 22)
        Me.Panel2.TabIndex = 8
        '
        'lblMinuteSingleDigit
        '
        Me.lblMinuteSingleDigit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMinuteSingleDigit.ForeColor = System.Drawing.Color.Black
        Me.lblMinuteSingleDigit.Location = New System.Drawing.Point(94, 4)
        Me.lblMinuteSingleDigit.Name = "lblMinuteSingleDigit"
        Me.lblMinuteSingleDigit.Size = New System.Drawing.Size(13, 15)
        Me.lblMinuteSingleDigit.TabIndex = 4
        Me.lblMinuteSingleDigit.Text = "9"
        Me.lblMinuteSingleDigit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblMinuteZeroFill
        '
        Me.lblMinuteZeroFill.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMinuteZeroFill.ForeColor = System.Drawing.Color.Black
        Me.lblMinuteZeroFill.Location = New System.Drawing.Point(82, 4)
        Me.lblMinuteZeroFill.Name = "lblMinuteZeroFill"
        Me.lblMinuteZeroFill.Size = New System.Drawing.Size(13, 15)
        Me.lblMinuteZeroFill.TabIndex = 3
        Me.lblMinuteZeroFill.Text = "0"
        Me.lblMinuteZeroFill.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblMinute
        '
        Me.lblMinute.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMinute.ForeColor = System.Drawing.Color.Black
        Me.lblMinute.Location = New System.Drawing.Point(82, 4)
        Me.lblMinute.Name = "lblMinute"
        Me.lblMinute.Size = New System.Drawing.Size(36, 15)
        Me.lblMinute.TabIndex = 2
        Me.lblMinute.Text = "XX"
        Me.lblMinute.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblHour
        '
        Me.lblHour.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHour.ForeColor = System.Drawing.Color.Black
        Me.lblHour.Location = New System.Drawing.Point(32, 4)
        Me.lblHour.Name = "lblHour"
        Me.lblHour.Size = New System.Drawing.Size(40, 15)
        Me.lblHour.TabIndex = 0
        Me.lblHour.Text = "XX"
        Me.lblHour.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblColon
        '
        Me.lblColon.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblColon.ForeColor = System.Drawing.Color.Black
        Me.lblColon.Location = New System.Drawing.Point(74, 3)
        Me.lblColon.Name = "lblColon"
        Me.lblColon.Size = New System.Drawing.Size(8, 15)
        Me.lblColon.TabIndex = 1
        Me.lblColon.Text = ":"
        Me.lblColon.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'InfoBox
        '
        Me.InfoBox.Controls.Add(Me.InfoPanel1)
        Me.InfoBox.Location = New System.Drawing.Point(775, 10)
        Me.InfoBox.Name = "InfoBox"
        Me.InfoBox.Size = New System.Drawing.Size(180, 135)
        Me.InfoBox.TabIndex = 2
        Me.InfoBox.TabStop = False
        '
        'InfoPanel1
        '
        Me.InfoPanel1.BackColor = System.Drawing.Color.DarkGray
        Me.InfoPanel1.Controls.Add(Me.GroupBox5)
        Me.InfoPanel1.Controls.Add(Me.GroupBox4)
        Me.InfoPanel1.Location = New System.Drawing.Point(2, 8)
        Me.InfoPanel1.Name = "InfoPanel1"
        Me.InfoPanel1.Size = New System.Drawing.Size(174, 123)
        Me.InfoPanel1.TabIndex = 0
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.VersionLabel)
        Me.GroupBox5.Controls.Add(Me.DateLabel)
        Me.GroupBox5.Location = New System.Drawing.Point(5, 40)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(165, 75)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        '
        'VersionLabel
        '
        Me.VersionLabel.Location = New System.Drawing.Point(5, 14)
        Me.VersionLabel.Name = "VersionLabel"
        Me.VersionLabel.Size = New System.Drawing.Size(150, 20)
        Me.VersionLabel.TabIndex = 3
        Me.VersionLabel.Text = "Version: AK56597_Av2"
        Me.VersionLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DateLabel
        '
        Me.DateLabel.Location = New System.Drawing.Point(5, 44)
        Me.DateLabel.Name = "DateLabel"
        Me.DateLabel.Size = New System.Drawing.Size(150, 20)
        Me.DateLabel.TabIndex = 4
        Me.DateLabel.Text = "1-26-2006"
        Me.DateLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.JobOrderLabel)
        Me.GroupBox4.Location = New System.Drawing.Point(5, 5)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(165, 32)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        '
        'JobOrderLabel
        '
        Me.JobOrderLabel.Location = New System.Drawing.Point(35, 10)
        Me.JobOrderLabel.Name = "JobOrderLabel"
        Me.JobOrderLabel.Size = New System.Drawing.Size(100, 20)
        Me.JobOrderLabel.TabIndex = 1
        Me.JobOrderLabel.Text = "JO# 56597"
        Me.JobOrderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ContactLabel
        '
        Me.ContactLabel.Location = New System.Drawing.Point(10, 135)
        Me.ContactLabel.Name = "ContactLabel"
        Me.ContactLabel.Size = New System.Drawing.Size(945, 15)
        Me.ContactLabel.TabIndex = 1
        Me.ContactLabel.Text = "www.nobletool.com     /     1535 Stanley Ave. Dayton, OH 45404     /     (937) 46" & _
        "1-4040"
        Me.ContactLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Picture01GroupBox
        '
        Me.Picture01GroupBox.Controls.Add(Me.PictureBox1)
        Me.Picture01GroupBox.Location = New System.Drawing.Point(440, 6)
        Me.Picture01GroupBox.Name = "Picture01GroupBox"
        Me.Picture01GroupBox.Size = New System.Drawing.Size(75, 125)
        Me.Picture01GroupBox.TabIndex = 0
        Me.Picture01GroupBox.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(5, 10)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(64, 110)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.SynchronizingObject = Me
        '
        'DataBox1
        '
        Me.DataBox1.Location = New System.Drawing.Point(0, 0)
        Me.DataBox1.Name = "DataBox1"
        Me.DataBox1.TabIndex = 0
        Me.DataBox1.TabStop = False
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 600
        '
        'Main
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Highlight
        Me.ClientSize = New System.Drawing.Size(1016, 741)
        Me.Controls.Add(Me.MainPanel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Name = "Main"
        Me.Text = "Noble Tool Corp."
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MainPanel.ResumeLayout(False)
        Me.MainBox1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.PartControlGroupBox.ResumeLayout(False)
        Me.AutoManualBox.ResumeLayout(False)
        Me.DataBox2.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        CType(Me.DOUT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DIN, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContactGroupBox.ResumeLayout(False)
        Me.TimeGroupBox.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.InfoBox.ResumeLayout(False)
        Me.InfoPanel1.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.Picture01GroupBox.ResumeLayout(False)
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Initialize Data Translation Inputs
        DIN.Board = DIN.get_BoardList(0)
        DIN.SubSysType = OLSS_DIN
        DIN.SubSysElement = 0
        DIN.DataFlow = OL_DF_SINGLEVALUE
        DIN.Config()
        'Initialize Data Translation Ouput
        DOUT.Board = DIN.get_BoardList(0)
        DOUT.SubSysType = OLSS_DIN
        DOUT.SubSysType = OLSS_DOUT
        DOUT.SubSysElement = 0
        DOUT.DataFlow = OL_DF_SINGLEVALUE
        DOUT.Config()
        'ProgressBar1.Step()
        'Fill text boxes
        PartNumberTextBox.Text = PartNumberStr

        ManualButton.PerformClick()

        ValidationCollection = False
        ProductionCollection = True
        fillTextBoxes()
        AutoPartCount = 0

        ' Create a new instance of Excel.
        objApp = New Excel.Application

    End Sub
    Private Sub UpdateIO()
        'Digital Input
        DataTranslationDIN = DIN.GetSingleValue(0, 1)
        If DataTranslationDIN = CycleComplete Then
            If ValidationCollection And ((AutoPartCountRunning - AutoPartCountPrevious) < 5) Then
                boolCycleComplete = True
                'Enable Manual Button
                disableManualButton = False
                If AutoModeBool Then
                    EnableButtons()
                End If
            End If

            If AutoPartCount < 5 And ProductionCollection Then
                boolCycleComplete = True
                'Enable Manual Button
                disableManualButton = False
                If AutoModeBool Then
                    EnableButtons()
                End If
            End If
            boolCycleRunning = False
            ' Turn Progress Bar Off
            ProgressBar1.Visible = False
            ' Reset Progress Bar
            ProgressBar1.Value = 0
            ' Turn status text box on
            StatusTextBox.Visible = True
            StatusTextBox.ForeColor = Color.Navy
            StatusTextBox.Text = "Cycle Complete"
            'Reset data aquisition one shot
            DataAcquisitionOneShot = False
            If ProductionCollection Then
                If AutoPartCount = 5 Then
                    CreateNewReport()
                    'Suspend DOUT
                    ' When creating reports, go to "PLC Manual" until Zero Prompt closed
                    DOUT.PutSingleValue(0, 1, 1)
                    SuspendDOUT = True
                End If
            End If
            If ValidationCollection Then
                If (Math.Abs(AutoPartCountRunning - AutoPartCountPrevious) = 5) Then
                    DOUT.PutSingleValue(0, 1, 1)
                    SuspendDOUT = True
                    ' Check Zero    
                    Dim newform03 As CheckZero = New CheckZero
                    newform03.Show()              'show instance of PartData
                    AutoPartCountPrevious = AutoPartCountRunning
                End If
            End If
        ElseIf DataTranslationDIN = CycleRunning And boolCycleComplete Then
            If ProductionCollection And AutoPartCount > 5 Then
                AutoPartCount = 0
            End If
            boolCycleComplete = False
            boolCycleRunning = True
            ' Disble Manual Button
            DisableButtons()
            disableManualButton = True
            ' Turn Progress Bar On
            ProgressBar1.Visible = True
            ' Turn StatusTextBox Off
            StatusTextBox.Visible = False
            'PartCountLabel.Visible = False
            writeStreamOneShot = False
        ElseIf DataTranslationDIN = TakeMeasurement And boolCycleRunning Then
            If (DataAcquisitionOneShot = False) Then
                DataAquisition()    ' Acquire data from Keyence Gauges
                DataAcquisitionOneShot = True
            End If

            'Enable One Shot
            DataAcquisitionOneShot = True
            ' Update the Form with new data
            If AutoModeBool And boolCycleRunning Then
                updateDataFields()
            ElseIf (AutoModeBool = False) Then
                updateDataFields()
            End If
        End If
        If DataTranslationDIN = VacuumFault Then
            boolVacuumFault = True
            boolCycleRunning = False
            ' Turn Progress Bar Off
            ProgressBar1.Visible = False
            ' Turn status text box on
            StatusTextBox.Visible = True
            ' Turn status text box on
            StatusTextBox.Visible = True
            StatusTextBox.ForeColor = Color.DarkRed
            StatusTextBox.Text = "Vacuum Fault"
        ElseIf DataTranslationDIN <> VacuumFault Then
            boolVacuumFault = False
        End If
        ' Alarm input: 250, 251, 254, 255
        If DataTranslationDIN > laserDisplacementErrorLim Then
            laserDisplacementERROR = True
        ElseIf laserDisplacementERROR And (DataTranslationDIN < laserDisplacementErrorLim) Then
            'startTime = Second(Delay00)
            laserDisplacementERROR = False
        End If
        'TestTextBox00.Text = DataTranslationDIN
        If (SuspendDOUT = False) Then
            DOUT.PutSingleValue(0, 1, DataTranslationDOUT)
        End If
    End Sub
    Private Sub UpdateDateValues()
        Dim dtNow As DateTime
        Dim currentTime As String
        'Set DateTime Object to Now
        dtNow = DateTime.Now
        'Display the Time
        If (dtNow.Hour > 12) Then
            lblHour.Text = dtNow.Hour - 12
        Else
            lblHour.Text = dtNow.Hour
        End If
        lblMinute.Text = dtNow.Minute
        lblMinuteSingleDigit.Text = dtNow.Minute
        'Zero Fill on Single Digit Minutes (i.e 0-9)
        If (dtNow.Minute <= 9) Then
            lblMinuteZeroFill.Visible = True
            lblMinuteSingleDigit.Visible = True
            lblMinute.Visible = False
        Else
            lblMinuteZeroFill.Visible = False
            lblMinuteSingleDigit.Visible = False
            lblMinute.Visible = True
        End If
        'Display the Date
        lblMonth.Text = dtNow.Month
        lblDay.Text = dtNow.Day
        lblYear.Text = dtNow.Year
    End Sub

    Private Sub Timer1_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed
        ' Query Data Translation I/O
        UpdateIO()
        ' Update Time and Date Information
        UpdateDateValues()
        If boolCycleRunning Then
            ProgressBar1.PerformStep() ' Increment Progress Bar
        End If
        'Update text boxes on part data change
        If partDataChanged Then
            fillTextBoxes()
            partDataChanged = False
        End If
        If laserDisplacementERROR And (AutoModeBool = False) Then
            FlangeTextBox.BackColor = Color.LightGray
            FlangeTextBox.ForeColor = Color.Black
            FlangeTextBox.Text = "xxxxxxx"
        End If

        If CheckZeroPageUp Then
            DisableButtons()
        End If

        If ProductionCollection Then
            AddNoteProdReportButton.Enabled = True
        ElseIf ValidationCollection Then
            AddNoteProdReportButton.Enabled = False
        End If
    End Sub

    Private Sub fillTextBoxes()
        If ProductionCollection Then
            PartNumberTextBox.Text = PartNumberStr
            OperatorIdentificationTextBox.Text = OperatorIdStr
            MoldNumberTextBox.Text = MoldNumberStr

            '****Disable Cavity****
            CavityNumberLabel.Visible = False
            CavityNumberTextBox.Visible = False
            CavityNumberTextBox.Text = " "
            '**********************

            '*****Disable Heat*****
            HeatNumberLabel.Visible = False
            HeatNumberTextBox.Visible = False
            HeatNumberTextBox.Text = " "
            '**********************

            '*****Disable Run******
            RunNumberLabel.Visible = False
            RunNumberTextBox.Visible = False
            RunNumberTextBox.Text = " "
            '**********************

            '**Enable Shop Order***
            ShopOrderLabel.Visible = True
            ShopOrderTextBox.Visible = True
            ShopOrderTextBox.Text = ShopOrderStr
            '**********************

            '*****Enable Time******
            PartTimeLabel.Visible = True
            PartTimeTextBox.Visible = True
            PartTimeTextBox.Text = PartTimeStr
            '**********************

            '*****Disable Date*****
            PartDateLabel.Visible = True
            PartDateTextBox.Visible = True
            PartDateTextBox.Text = PartDateStr
            '**********************

            PartDateTextBox.Text = PartDateStr
        ElseIf ValidationCollection Then
            PartNumberTextBox.Text = PartNumberStr
            OperatorIdentificationTextBox.Text = OperatorIdStr
            MoldNumberTextBox.Text = MoldNumberStr

            '**Disable Shop Order**
            ShopOrderLabel.Visible = False
            ShopOrderTextBox.Visible = False
            ShopOrderTextBox.Text = " "
            '**********************

            '*****Disable Time*****
            PartTimeLabel.Visible = False
            PartTimeTextBox.Visible = False
            PartTimeTextBox.Text = " "
            '**********************

            '*****Disable Date*****
            PartDateLabel.Visible = False
            PartDateTextBox.Visible = False
            PartDateTextBox.Text = " "
            '**********************

            '****Enable Cavity*****
            CavityNumberLabel.Visible = True
            CavityNumberTextBox.Visible = True
            CavityNumberTextBox.Text = CavityNumberStr
            '**********************

            '*****Enable Heat******
            HeatNumberLabel.Visible = True
            HeatNumberTextBox.Visible = True
            HeatNumberTextBox.Text = HeatNumberStr
            '**********************

            '******Enable Run******
            RunNumberLabel.Visible = True
            RunNumberTextBox.Visible = True
            RunNumberTextBox.Text = RunNumberStr
            '**********************
        End If



    End Sub
    Private Sub updateDataFields()
        Dim StringTemp As String
        ' Range comparison for left cone
        If (leftConeDbl < 0.005) Or (leftConeDbl > 0.008) Then
            LeftConeTextBox.BackColor = Color.Yellow
            LeftConeTextBox.ForeColor = Color.Red
        Else
            LeftConeTextBox.BackColor = Color.LightGray
            LeftConeTextBox.ForeColor = Color.ForestGreen
        End If

        ' Range comparison for right cone
        If (rightConeDbl < 0.005) Or (rightConeDbl > 0.008) Then
            RightConeTextBox.BackColor = Color.Yellow
            RightConeTextBox.ForeColor = Color.Red
        Else
            RightConeTextBox.BackColor = Color.LightGray
            RightConeTextBox.ForeColor = Color.ForestGreen
        End If

        ' Range comparison for pad
        If (padDbl < 0.015) Or (padDbl > 0.018) Then
            PadTextBox.BackColor = Color.Yellow
            PadTextBox.ForeColor = Color.Red
        Else
            PadTextBox.BackColor = Color.LightGray
            PadTextBox.ForeColor = Color.ForestGreen
        End If

        ' Range comparison for flange
        If (flangeDbl < 0.009) Or (flangeDbl > 0.012) Or laserDisplacementERROR Then
            FlangeTextBox.BackColor = Color.Yellow
            FlangeTextBox.ForeColor = Color.Red
        Else
            FlangeTextBox.BackColor = Color.LightGray
            FlangeTextBox.ForeColor = Color.ForestGreen
        End If
        ' Write data to the Form
        If port02Head01Out02Err Then
            LeftConeTextBox.BackColor = Color.LightGray
            LeftConeTextBox.ForeColor = Color.Black
            LeftConeTextBox.Text = "xxxxxxx"
        Else
            LeftConeTextBox.Text = Format(leftConeDbl, "0.00000")
        End If
        If port02Head01Out03Err Then
            RightConeTextBox.BackColor = Color.LightGray
            RightConeTextBox.ForeColor = Color.Black
            RightConeTextBox.Text = "xxxxxxx"
        Else
            RightConeTextBox.Text = Format(rightConeDbl, "0.00000")
        End If
        If port02Head02Out04Err Then
            PadTextBox.BackColor = Color.LightGray
            PadTextBox.ForeColor = Color.Black
            PadTextBox.Text = "xxxxxxx"
        Else
            PadTextBox.Text = Format(padDbl, "0.00000")
        End If

        If laserDisplacementERROR Or port05Head01Out01Err Then
            FlangeTextBox.BackColor = Color.LightGray
            FlangeTextBox.ForeColor = Color.Black
            FlangeTextBox.Text = "xxxxxxx"
        Else
            FlangeTextBox.Text = Format(flangeDbl, "0.00000")
        End If
        If ProductionCollection Then
            Select Case AutoPartCount
                Case 1
                    PartCountLabel.Visible = True
                    PartCountLabel.Text = "Part 1 of 5"
                Case 2
                    PartCountLabel.Visible = True
                    PartCountLabel.Text = "Part 2 of 5"
                Case 3
                    PartCountLabel.Visible = True
                    PartCountLabel.Text = "Part 3 of 5"
                Case 4
                    PartCountLabel.Visible = True
                    PartCountLabel.Text = "Part 4 of 5"
                Case 5
                    PartCountLabel.Visible = True
                    PartCountLabel.Text = "Part 5 of 5"
            End Select
        ElseIf ValidationCollection Then
            PartCountLabel.Visible = True
            PartCountLabel.Text = "Part: " & AutoPartCountRunning.ToString
        End If


    End Sub

    Private Sub AutoButton_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoButton.Click
        DataTranslationDOUT = 0
        AutoButton.ForeColor = Color.Lime
        ManualButton.ForeColor = Color.Black
        AutoButton.BackColor = Color.Gray
        ManualButton.BackColor = Color.Transparent
        ' Make Part Count Visible
        If AutoPartCount > 0 Then
            PartCountLabel.Visible = True
            'Else
            '    PartCountLabel.Visible = False
        End If
        AutoModeBool = True
        'Disable Timer2
        Timer2.Enabled = False
        ErrorStateBool = False
        ' Disable Production Report button
        Button2.Enabled = True
        ' Disable Part Edit button
        Button1.Enabled = True
        ' Enable Zero Button
        zeroButton.Enabled = False
    End Sub

    Private Sub ManualButton_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ManualButton.Click
        DataTranslationDOUT = 1
        AutoButton.ForeColor = Color.Black
        ManualButton.ForeColor = Color.Lime
        AutoButton.BackColor = Color.Transparent
        ManualButton.BackColor = Color.Gray
        ' Hide Part Count
        PartCountLabel.Visible = False
        AutoModeBool = False
        ' Enable Timer2
        Timer2.Enabled = True
        ErrorStateBool = False
        ' Disable Production Report button
        Button2.Enabled = False
        ' Disable Part Edit button
        Button1.Enabled = False
        ' Enable Zero Button
        zeroButton.Enabled = True

        AutoPartCount = 0
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If (laserDisplacementERROR = False) Then
            If (AutoModeBool = False) And (laserDisplacementERROR = False) Then
                DataAquisition()
                updateDataFields()
            End If
        End If
    End Sub

    Public Sub zeroButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles zeroButton.Click
        ' Zero the left cone
        zeroLeftConeDbl = rawLeftConeDbl
        ' Zero the right cone
        zeroRightConeDbl = rawRightConeDbl
        ' Zero the pad
        zeroPadDbl = rawPadDbl
        ' Zero the flange
        zeroFlangeDbl = rawFlangeDbl
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ' Create an instance of PartData
        Dim newform02 As ProductionReport = New ProductionReport
        newform02.Show()              'show instance of PartData
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' Create an instance of PartData
        Dim newform07 As PartData = New PartData
        newform07.Show()              'show instance of PartData
    End Sub

    Private Sub DisableButtons()
        ManualButton.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
    End Sub

    Private Sub EnableButtons()
        ManualButton.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
    End Sub

    Private Sub addNoteToDataStrButton_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles addNoteToDataStrButton.Click
        StreamDataNoteBool = True
        Dim newform04 As Note = New Note
        newform04.Show()              'show instance of PartData
    End Sub

    Private Sub AddNoteProdReportButton_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddNoteProdReportButton.Click
        ProductionReportBool = True
        Dim newform05 As Note = New Note
        newform05.Show()              'show instance of PartData
    End Sub
End Class
