Public Class PartData
    Inherits System.Windows.Forms.Form

    Private Sub ProductionReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PartNumberTextBox2.Text = PartNumberStr
        OperatorIDTextBox2.Text = OperatorIdStr
        MoldNumberTextBox2.Text = MoldNumberStr
        If ProductionCollection Then
            'Disable Cavity
            CavityaNumberTextBox2.Visible = False
            CavityNumberLabel2.Visible = False
            'Disable Heat
            HeatNumberTextBox2.Visible = False
            HeatNumberLabel2.Visible = False
            'Disable Run Number
            RunNumberTextBox.Visible = False
            RunNumberLabel.Visible = False

            'Enable Shop Order
            ShopOrderTextBox.Visible = True
            ShopOrderLabel.Visible = True
            ShopOrderTextBox.Text = ShopOrderStr
            'Enable Sample Time
            PartTimeTextBox.Visible = True
            PartTimeLabel.Visible = True
            PartTimeTextBox.Text = PartTimeStr
            'Enable Sample Date
            PartDateTextBox.Visible = True
            SampleDateLabel.Visible = True
            PartDateTextBox.Text = PartDateStr
        ElseIf ValidationCollection Then
            'Disable Shop Order
            ShopOrderTextBox.Visible = False
            ShopOrderLabel.Visible = False
            'Disable Sample Time
            PartTimeTextBox.Visible = False
            PartTimeLabel.Visible = False
            'Disable Sample Date
            PartDateTextBox.Visible = False
            SampleDateLabel.Visible = False

            'Enable Cavity
            CavityaNumberTextBox2.Visible = True
            CavityNumberLabel2.Visible = True
            CavityaNumberTextBox2.Text = CavityNumberStr
            'Enable Heat
            HeatNumberTextBox2.Visible = True
            HeatNumberLabel2.Visible = True
            HeatNumberTextBox2.Text = HeatNumberStr
            'Enable Run Number
            RunNumberTextBox.Visible = True
            RunNumberLabel.Visible = True
            RunNumberTextBox.Text = RunNumberStr
        End If
    End Sub


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
    Friend WithEvents PartNumberTextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents PartNumberLabel2 As System.Windows.Forms.Label
    Friend WithEvents OkButton1 As System.Windows.Forms.Button
    Friend WithEvents OperatorIDLabel2 As System.Windows.Forms.Label
    Friend WithEvents OperatorIDTextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents MoldNumberLabel2 As System.Windows.Forms.Label
    Friend WithEvents MoldNumberTextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents CavityNumberLabel2 As System.Windows.Forms.Label
    Friend WithEvents CavityaNumberTextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents HeatNumberLabel2 As System.Windows.Forms.Label
    Friend WithEvents HeatNumberTextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ShopOrderTextBox As System.Windows.Forms.TextBox
    Friend WithEvents ShopOrderLabel As System.Windows.Forms.Label
    Friend WithEvents PartTimeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PartTimeLabel As System.Windows.Forms.Label
    Friend WithEvents SampleDateLabel As System.Windows.Forms.Label
    Friend WithEvents RunNumberLabel As System.Windows.Forms.Label
    Friend WithEvents RunNumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PartDateTextBox As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(PartData))
        Me.PartNumberTextBox2 = New System.Windows.Forms.TextBox
        Me.PartNumberLabel2 = New System.Windows.Forms.Label
        Me.OkButton1 = New System.Windows.Forms.Button
        Me.OperatorIDLabel2 = New System.Windows.Forms.Label
        Me.OperatorIDTextBox2 = New System.Windows.Forms.TextBox
        Me.MoldNumberLabel2 = New System.Windows.Forms.Label
        Me.MoldNumberTextBox2 = New System.Windows.Forms.TextBox
        Me.CavityNumberLabel2 = New System.Windows.Forms.Label
        Me.CavityaNumberTextBox2 = New System.Windows.Forms.TextBox
        Me.HeatNumberLabel2 = New System.Windows.Forms.Label
        Me.HeatNumberTextBox2 = New System.Windows.Forms.TextBox
        Me.RunNumberLabel = New System.Windows.Forms.Label
        Me.RunNumberTextBox = New System.Windows.Forms.TextBox
        Me.ShopOrderTextBox = New System.Windows.Forms.TextBox
        Me.ShopOrderLabel = New System.Windows.Forms.Label
        Me.PartTimeTextBox = New System.Windows.Forms.TextBox
        Me.PartTimeLabel = New System.Windows.Forms.Label
        Me.SampleDateLabel = New System.Windows.Forms.Label
        Me.PartDateTextBox = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'PartNumberTextBox2
        '
        Me.PartNumberTextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartNumberTextBox2.Location = New System.Drawing.Point(136, 32)
        Me.PartNumberTextBox2.Name = "PartNumberTextBox2"
        Me.PartNumberTextBox2.Size = New System.Drawing.Size(190, 22)
        Me.PartNumberTextBox2.TabIndex = 0
        Me.PartNumberTextBox2.Text = ""
        '
        'PartNumberLabel2
        '
        Me.PartNumberLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartNumberLabel2.Location = New System.Drawing.Point(22, 32)
        Me.PartNumberLabel2.Name = "PartNumberLabel2"
        Me.PartNumberLabel2.Size = New System.Drawing.Size(96, 20)
        Me.PartNumberLabel2.TabIndex = 33
        Me.PartNumberLabel2.Text = "Part Number :"
        Me.PartNumberLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'OkButton1
        '
        Me.OkButton1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OkButton1.Location = New System.Drawing.Point(104, 272)
        Me.OkButton1.Name = "OkButton1"
        Me.OkButton1.Size = New System.Drawing.Size(152, 40)
        Me.OkButton1.TabIndex = 6
        Me.OkButton1.Text = "OK"
        '
        'OperatorIDLabel2
        '
        Me.OperatorIDLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OperatorIDLabel2.Location = New System.Drawing.Point(22, 72)
        Me.OperatorIDLabel2.Name = "OperatorIDLabel2"
        Me.OperatorIDLabel2.Size = New System.Drawing.Size(96, 20)
        Me.OperatorIDLabel2.TabIndex = 33
        Me.OperatorIDLabel2.Text = "Operator ID :"
        Me.OperatorIDLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'OperatorIDTextBox2
        '
        Me.OperatorIDTextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OperatorIDTextBox2.Location = New System.Drawing.Point(136, 72)
        Me.OperatorIDTextBox2.Name = "OperatorIDTextBox2"
        Me.OperatorIDTextBox2.Size = New System.Drawing.Size(190, 22)
        Me.OperatorIDTextBox2.TabIndex = 1
        Me.OperatorIDTextBox2.Text = ""
        '
        'MoldNumberLabel2
        '
        Me.MoldNumberLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MoldNumberLabel2.Location = New System.Drawing.Point(22, 112)
        Me.MoldNumberLabel2.Name = "MoldNumberLabel2"
        Me.MoldNumberLabel2.Size = New System.Drawing.Size(96, 20)
        Me.MoldNumberLabel2.TabIndex = 33
        Me.MoldNumberLabel2.Text = "Mold Number :"
        Me.MoldNumberLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'MoldNumberTextBox2
        '
        Me.MoldNumberTextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MoldNumberTextBox2.Location = New System.Drawing.Point(136, 112)
        Me.MoldNumberTextBox2.Name = "MoldNumberTextBox2"
        Me.MoldNumberTextBox2.Size = New System.Drawing.Size(190, 22)
        Me.MoldNumberTextBox2.TabIndex = 2
        Me.MoldNumberTextBox2.Text = ""
        '
        'CavityNumberLabel2
        '
        Me.CavityNumberLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CavityNumberLabel2.Location = New System.Drawing.Point(6, 152)
        Me.CavityNumberLabel2.Name = "CavityNumberLabel2"
        Me.CavityNumberLabel2.Size = New System.Drawing.Size(112, 20)
        Me.CavityNumberLabel2.TabIndex = 35
        Me.CavityNumberLabel2.Text = "Cavity Number :"
        Me.CavityNumberLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CavityaNumberTextBox2
        '
        Me.CavityaNumberTextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CavityaNumberTextBox2.Location = New System.Drawing.Point(136, 152)
        Me.CavityaNumberTextBox2.Name = "CavityaNumberTextBox2"
        Me.CavityaNumberTextBox2.Size = New System.Drawing.Size(190, 22)
        Me.CavityaNumberTextBox2.TabIndex = 3
        Me.CavityaNumberTextBox2.Text = ""
        '
        'HeatNumberLabel2
        '
        Me.HeatNumberLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HeatNumberLabel2.Location = New System.Drawing.Point(6, 192)
        Me.HeatNumberLabel2.Name = "HeatNumberLabel2"
        Me.HeatNumberLabel2.Size = New System.Drawing.Size(112, 20)
        Me.HeatNumberLabel2.TabIndex = 37
        Me.HeatNumberLabel2.Text = "Heat Number :"
        Me.HeatNumberLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'HeatNumberTextBox2
        '
        Me.HeatNumberTextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HeatNumberTextBox2.Location = New System.Drawing.Point(136, 192)
        Me.HeatNumberTextBox2.Name = "HeatNumberTextBox2"
        Me.HeatNumberTextBox2.Size = New System.Drawing.Size(190, 22)
        Me.HeatNumberTextBox2.TabIndex = 4
        Me.HeatNumberTextBox2.Text = ""
        '
        'RunNumberLabel
        '
        Me.RunNumberLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunNumberLabel.Location = New System.Drawing.Point(6, 232)
        Me.RunNumberLabel.Name = "RunNumberLabel"
        Me.RunNumberLabel.Size = New System.Drawing.Size(112, 20)
        Me.RunNumberLabel.TabIndex = 39
        Me.RunNumberLabel.Text = "Run :"
        Me.RunNumberLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'RunNumberTextBox
        '
        Me.RunNumberTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunNumberTextBox.Location = New System.Drawing.Point(136, 232)
        Me.RunNumberTextBox.Name = "RunNumberTextBox"
        Me.RunNumberTextBox.Size = New System.Drawing.Size(190, 22)
        Me.RunNumberTextBox.TabIndex = 5
        Me.RunNumberTextBox.Text = ""
        '
        'ShopOrderTextBox
        '
        Me.ShopOrderTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ShopOrderTextBox.Location = New System.Drawing.Point(136, 152)
        Me.ShopOrderTextBox.Name = "ShopOrderTextBox"
        Me.ShopOrderTextBox.Size = New System.Drawing.Size(190, 22)
        Me.ShopOrderTextBox.TabIndex = 3
        Me.ShopOrderTextBox.Text = ""
        '
        'ShopOrderLabel
        '
        Me.ShopOrderLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ShopOrderLabel.Location = New System.Drawing.Point(6, 152)
        Me.ShopOrderLabel.Name = "ShopOrderLabel"
        Me.ShopOrderLabel.Size = New System.Drawing.Size(112, 20)
        Me.ShopOrderLabel.TabIndex = 41
        Me.ShopOrderLabel.Text = "Shop Order :"
        Me.ShopOrderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PartTimeTextBox
        '
        Me.PartTimeTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartTimeTextBox.Location = New System.Drawing.Point(136, 192)
        Me.PartTimeTextBox.Name = "PartTimeTextBox"
        Me.PartTimeTextBox.Size = New System.Drawing.Size(190, 22)
        Me.PartTimeTextBox.TabIndex = 4
        Me.PartTimeTextBox.Text = ""
        '
        'PartTimeLabel
        '
        Me.PartTimeLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartTimeLabel.Location = New System.Drawing.Point(6, 192)
        Me.PartTimeLabel.Name = "PartTimeLabel"
        Me.PartTimeLabel.Size = New System.Drawing.Size(112, 20)
        Me.PartTimeLabel.TabIndex = 43
        Me.PartTimeLabel.Text = "Sample Time :"
        Me.PartTimeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'SampleDateLabel
        '
        Me.SampleDateLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SampleDateLabel.Location = New System.Drawing.Point(6, 232)
        Me.SampleDateLabel.Name = "SampleDateLabel"
        Me.SampleDateLabel.Size = New System.Drawing.Size(112, 20)
        Me.SampleDateLabel.TabIndex = 44
        Me.SampleDateLabel.Text = "Sample Date :"
        Me.SampleDateLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PartDateTextBox
        '
        Me.PartDateTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartDateTextBox.Location = New System.Drawing.Point(136, 232)
        Me.PartDateTextBox.Name = "PartDateTextBox"
        Me.PartDateTextBox.Size = New System.Drawing.Size(190, 22)
        Me.PartDateTextBox.TabIndex = 5
        Me.PartDateTextBox.Text = ""
        '
        'PartData
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(354, 319)
        Me.ControlBox = False
        Me.Controls.Add(Me.PartDateTextBox)
        Me.Controls.Add(Me.SampleDateLabel)
        Me.Controls.Add(Me.PartTimeLabel)
        Me.Controls.Add(Me.PartTimeTextBox)
        Me.Controls.Add(Me.ShopOrderLabel)
        Me.Controls.Add(Me.ShopOrderTextBox)
        Me.Controls.Add(Me.RunNumberLabel)
        Me.Controls.Add(Me.RunNumberTextBox)
        Me.Controls.Add(Me.HeatNumberLabel2)
        Me.Controls.Add(Me.HeatNumberTextBox2)
        Me.Controls.Add(Me.CavityNumberLabel2)
        Me.Controls.Add(Me.CavityaNumberTextBox2)
        Me.Controls.Add(Me.MoldNumberLabel2)
        Me.Controls.Add(Me.MoldNumberTextBox2)
        Me.Controls.Add(Me.OperatorIDLabel2)
        Me.Controls.Add(Me.OperatorIDTextBox2)
        Me.Controls.Add(Me.OkButton1)
        Me.Controls.Add(Me.PartNumberLabel2)
        Me.Controls.Add(Me.PartNumberTextBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PartData"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Part Data"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub OkButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OkButton1.Click
        partDataChanged = True      ' Set part data changed flag
        Me.Close()                  ' Close part data form
    End Sub

    Private Sub PartNumberTextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PartNumberTextBox2.TextChanged
        PartNumberStr = PartNumberTextBox2.Text         ' write to global string
    End Sub

    Private Sub OperatorIDTextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OperatorIDTextBox2.TextChanged
        OperatorIdStr = OperatorIDTextBox2.Text         ' write to global string     
    End Sub

    Private Sub MoldNumberTextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MoldNumberTextBox2.TextChanged
        MoldNumberStr = MoldNumberTextBox2.Text         ' write to global string
    End Sub

    Private Sub CavityaNumberTextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CavityaNumberTextBox2.TextChanged
        CavityNumberStr = CavityaNumberTextBox2.Text    ' write to global string
    End Sub

    Private Sub HeatNumberTextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HeatNumberTextBox2.TextChanged
        HeatNumberStr = HeatNumberTextBox2.Text         ' write to global string
    End Sub
    Private Sub RunNumberTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunNumberTextBox.TextChanged
        RunNumberStr = RunNumberTextBox.Text         ' write to global string
    End Sub

    Private Sub ShopOrderTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShopOrderTextBox.TextChanged
        ShopOrderStr = ShopOrderTextBox.Text             ' write to global string
    End Sub
    Private Sub PartTimeTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PartTimeTextBox.TextChanged
        PartTimeStr = PartTimeTextBox.Text             ' write to global string
    End Sub

    Private Sub PartDateTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PartDateTextBox.TextChanged
        PartDateStr = PartDateTextBox.Text             ' write to global string
    End Sub
End Class
