Public Class ProductionReport
    Inherits System.Windows.Forms.Form
    Dim ChangeState As Boolean
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
    Friend WithEvents CreateReportButton As System.Windows.Forms.Button
    Friend WithEvents ResetReportCountButton As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents OkButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ValidationRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents ProductionRadioButton As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ProductionReport))
        Me.CreateReportButton = New System.Windows.Forms.Button
        Me.ResetReportCountButton = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.OkButton = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ProductionRadioButton = New System.Windows.Forms.RadioButton
        Me.ValidationRadioButton = New System.Windows.Forms.RadioButton
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CreateReportButton
        '
        Me.CreateReportButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CreateReportButton.Location = New System.Drawing.Point(197, 184)
        Me.CreateReportButton.Name = "CreateReportButton"
        Me.CreateReportButton.Size = New System.Drawing.Size(188, 50)
        Me.CreateReportButton.TabIndex = 1
        Me.CreateReportButton.Text = "CREATE PRODUCTION REPORT"
        '
        'ResetReportCountButton
        '
        Me.ResetReportCountButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ResetReportCountButton.Location = New System.Drawing.Point(386, 184)
        Me.ResetReportCountButton.Name = "ResetReportCountButton"
        Me.ResetReportCountButton.Size = New System.Drawing.Size(188, 50)
        Me.ResetReportCountButton.TabIndex = 2
        Me.ResetReportCountButton.Text = "RESET PRODUCTION REPORT"
        '
        'TextBox1
        '
        Me.TextBox1.AutoSize = False
        Me.TextBox1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(8, 8)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(568, 72)
        Me.TextBox1.TabIndex = 45
        Me.TextBox1.TabStop = False
        Me.TextBox1.Text = "Choose between Validation or Production data collection. Select ""CREATE REPORT"" t" & _
        "o create report with curent part data. Select ""RESET REPORT"" to reset the part c" & _
        "ount and clear data buffers. "
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'OkButton
        '
        Me.OkButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OkButton.Location = New System.Drawing.Point(8, 184)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.Size = New System.Drawing.Size(188, 50)
        Me.OkButton.TabIndex = 4
        Me.OkButton.Text = "OK"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ProductionRadioButton)
        Me.GroupBox1.Controls.Add(Me.ValidationRadioButton)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 80)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(568, 96)
        Me.GroupBox1.TabIndex = 47
        Me.GroupBox1.TabStop = False
        '
        'ProductionRadioButton
        '
        Me.ProductionRadioButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.ProductionRadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ProductionRadioButton.Location = New System.Drawing.Point(166, 56)
        Me.ProductionRadioButton.Name = "ProductionRadioButton"
        Me.ProductionRadioButton.Size = New System.Drawing.Size(236, 24)
        Me.ProductionRadioButton.TabIndex = 48
        Me.ProductionRadioButton.Text = "Production Data Collection"
        '
        'ValidationRadioButton
        '
        Me.ValidationRadioButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.ValidationRadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ValidationRadioButton.Location = New System.Drawing.Point(166, 16)
        Me.ValidationRadioButton.Name = "ValidationRadioButton"
        Me.ValidationRadioButton.Size = New System.Drawing.Size(236, 24)
        Me.ValidationRadioButton.TabIndex = 47
        Me.ValidationRadioButton.Text = "Validation Data Collection"
        '
        'ProductionReport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(582, 239)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.OkButton)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.ResetReportCountButton)
        Me.Controls.Add(Me.CreateReportButton)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ProductionReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Data Collection"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub ProductionReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ProductionCollection Then
            ProductionRadioButton.Checked = True
            ResetReportCountButton.Enabled = True
        ElseIf ValidationCollection Then
            ValidationRadioButton.Checked = True
            ResetReportCountButton.Enabled = False
        End If
        ChangeState = False
    End Sub

    Private Sub CreateReportButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateReportButton.Click
        CreateNewReport()
    End Sub

    Private Sub ResetReportCountButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetReportCountButton.Click
        Dim i As Integer = 0
        AutoPartCount = 0
        For i = 0 To 4
            padDataArray(i) = 0
            flangeDataArray(i) = 0
            leftConeDataArray(i) = 0
            rightConeDataArray(i) = 0
        Next i
    End Sub

    Private Sub OkButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OkButton.Click

        If ChangeState Then
            'MessageBox.Show("Changed data collection setting, update part information", MessageBoxIcon.Exclamation, MessageBoxButtons.OK) ', MessageBoxIcon.Exclamation)
            MessageBox.Show("Data collection option have Changed. Update part information", "Part Information", _
           MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End If
        Me.Close()
    End Sub

    Private Sub ValidationRadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValidationRadioButton.CheckedChanged
        If ValidationRadioButton.Checked Then
            ValidationCollection = True
            ProductionCollection = False
            ProductionRadioButton.Checked = False
            ChangeState = True
            partDataChanged = True
            ResetReportCountButton.Enabled = False
            AutoPartCount = 0
            AutoPartCountRunning = 0
            AutoPartCountPrevious = 0
        ElseIf ProductionRadioButton.Checked Then
            ValidationCollection = False
            ProductionCollection = True
            ValidationRadioButton.Checked = False
            partDataChanged = True
            ChangeState = True
            AutoPartCount = 0
            ResetReportCountButton.Enabled = True
        End If
    End Sub
End Class
