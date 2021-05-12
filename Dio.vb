Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents _picLEDIN_0 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_1 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_2 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_3 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_4 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_5 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_6 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_7 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_8 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_9 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_10 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_11 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_12 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_13 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_14 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDIN_15 As System.Windows.Forms.PictureBox
	Public WithEvents lblInData As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents piclighton As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_0 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_15 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_14 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_13 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_12 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_11 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_10 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_9 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_8 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_7 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_6 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_5 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_4 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_3 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_2 As System.Windows.Forms.PictureBox
	Public WithEvents _picLEDOUT_1 As System.Windows.Forms.PictureBox
	Public WithEvents lblOutData As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents piclightoff As System.Windows.Forms.PictureBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents DIN As AxDTAcq32Lib.AxDTAcq32
	Public WithEvents DOUT As AxDTAcq32Lib.AxDTAcq32
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents picLEDIN As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	Public WithEvents picLEDOUT As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Frame3 = New System.Windows.Forms.GroupBox
		Me._picLEDIN_0 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_1 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_2 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_3 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_4 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_5 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_6 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_7 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_8 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_9 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_10 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_11 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_12 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_13 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_14 = New System.Windows.Forms.PictureBox
		Me._picLEDIN_15 = New System.Windows.Forms.PictureBox
		Me.lblInData = New System.Windows.Forms.Label
		Me.piclighton = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_0 = New System.Windows.Forms.PictureBox
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me._picLEDOUT_15 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_14 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_13 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_12 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_11 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_10 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_9 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_8 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_7 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_6 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_5 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_4 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_3 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_2 = New System.Windows.Forms.PictureBox
		Me._picLEDOUT_1 = New System.Windows.Forms.PictureBox
		Me.lblOutData = New System.Windows.Forms.Label
		Me.piclightoff = New System.Windows.Forms.PictureBox
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.DIN = New AxDTAcq32Lib.AxDTAcq32
		Me.DOUT = New AxDTAcq32Lib.AxDTAcq32
		Me.Label1 = New System.Windows.Forms.Label
		Me.picLEDIN = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.picLEDOUT = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		CType(Me.DIN, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DOUT, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.picLEDIN, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.picLEDOUT, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.SystemColors.Window
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Digital I/O"
		Me.ClientSize = New System.Drawing.Size(614, 278)
		Me.Location = New System.Drawing.Point(15, 98)
		Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Icon = CType(resources.GetObject("Form1.Icon"), System.Drawing.Icon)
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form1"
		Me.Frame3.BackColor = System.Drawing.SystemColors.Window
		Me.Frame3.Text = "Digital Input"
		Me.Frame3.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Frame3.Size = New System.Drawing.Size(609, 105)
		Me.Frame3.Location = New System.Drawing.Point(0, 16)
		Me.Frame3.TabIndex = 21
		Me.Frame3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame3.Enabled = True
		Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame3.Visible = True
		Me.Frame3.Name = "Frame3"
		Me._picLEDIN_0.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_0.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_0.Location = New System.Drawing.Point(568, 24)
		Me._picLEDIN_0.TabIndex = 37
		Me._picLEDIN_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_0.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_0.CausesValidation = True
		Me._picLEDIN_0.Enabled = True
		Me._picLEDIN_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_0.TabStop = True
		Me._picLEDIN_0.Visible = True
		Me._picLEDIN_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_0.Name = "_picLEDIN_0"
		Me._picLEDIN_1.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_1.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_1.Location = New System.Drawing.Point(536, 24)
		Me._picLEDIN_1.TabIndex = 36
		Me._picLEDIN_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_1.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_1.CausesValidation = True
		Me._picLEDIN_1.Enabled = True
		Me._picLEDIN_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_1.TabStop = True
		Me._picLEDIN_1.Visible = True
		Me._picLEDIN_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_1.Name = "_picLEDIN_1"
		Me._picLEDIN_2.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_2.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_2.Location = New System.Drawing.Point(504, 24)
		Me._picLEDIN_2.TabIndex = 35
		Me._picLEDIN_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_2.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_2.CausesValidation = True
		Me._picLEDIN_2.Enabled = True
		Me._picLEDIN_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_2.TabStop = True
		Me._picLEDIN_2.Visible = True
		Me._picLEDIN_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_2.Name = "_picLEDIN_2"
		Me._picLEDIN_3.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_3.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_3.Location = New System.Drawing.Point(472, 24)
		Me._picLEDIN_3.TabIndex = 34
		Me._picLEDIN_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_3.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_3.CausesValidation = True
		Me._picLEDIN_3.Enabled = True
		Me._picLEDIN_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_3.TabStop = True
		Me._picLEDIN_3.Visible = True
		Me._picLEDIN_3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_3.Name = "_picLEDIN_3"
		Me._picLEDIN_4.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_4.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_4.Location = New System.Drawing.Point(416, 24)
		Me._picLEDIN_4.TabIndex = 33
		Me._picLEDIN_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_4.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_4.CausesValidation = True
		Me._picLEDIN_4.Enabled = True
		Me._picLEDIN_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_4.TabStop = True
		Me._picLEDIN_4.Visible = True
		Me._picLEDIN_4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_4.Name = "_picLEDIN_4"
		Me._picLEDIN_5.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_5.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_5.Location = New System.Drawing.Point(384, 24)
		Me._picLEDIN_5.TabIndex = 32
		Me._picLEDIN_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_5.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_5.CausesValidation = True
		Me._picLEDIN_5.Enabled = True
		Me._picLEDIN_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_5.TabStop = True
		Me._picLEDIN_5.Visible = True
		Me._picLEDIN_5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_5.Name = "_picLEDIN_5"
		Me._picLEDIN_6.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_6.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_6.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_6.Location = New System.Drawing.Point(352, 24)
		Me._picLEDIN_6.TabIndex = 31
		Me._picLEDIN_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_6.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_6.CausesValidation = True
		Me._picLEDIN_6.Enabled = True
		Me._picLEDIN_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_6.TabStop = True
		Me._picLEDIN_6.Visible = True
		Me._picLEDIN_6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_6.Name = "_picLEDIN_6"
		Me._picLEDIN_7.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_7.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_7.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_7.Location = New System.Drawing.Point(320, 24)
		Me._picLEDIN_7.TabIndex = 30
		Me._picLEDIN_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_7.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_7.CausesValidation = True
		Me._picLEDIN_7.Enabled = True
		Me._picLEDIN_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_7.TabStop = True
		Me._picLEDIN_7.Visible = True
		Me._picLEDIN_7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_7.Name = "_picLEDIN_7"
		Me._picLEDIN_8.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_8.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_8.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_8.Location = New System.Drawing.Point(264, 24)
		Me._picLEDIN_8.TabIndex = 29
		Me._picLEDIN_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_8.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_8.CausesValidation = True
		Me._picLEDIN_8.Enabled = True
		Me._picLEDIN_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_8.TabStop = True
		Me._picLEDIN_8.Visible = True
		Me._picLEDIN_8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_8.Name = "_picLEDIN_8"
		Me._picLEDIN_9.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_9.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_9.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_9.Location = New System.Drawing.Point(232, 24)
		Me._picLEDIN_9.TabIndex = 28
		Me._picLEDIN_9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_9.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_9.CausesValidation = True
		Me._picLEDIN_9.Enabled = True
		Me._picLEDIN_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_9.TabStop = True
		Me._picLEDIN_9.Visible = True
		Me._picLEDIN_9.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_9.Name = "_picLEDIN_9"
		Me._picLEDIN_10.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_10.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_10.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_10.Location = New System.Drawing.Point(200, 24)
		Me._picLEDIN_10.TabIndex = 27
		Me._picLEDIN_10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_10.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_10.CausesValidation = True
		Me._picLEDIN_10.Enabled = True
		Me._picLEDIN_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_10.TabStop = True
		Me._picLEDIN_10.Visible = True
		Me._picLEDIN_10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_10.Name = "_picLEDIN_10"
		Me._picLEDIN_11.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_11.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_11.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_11.Location = New System.Drawing.Point(168, 24)
		Me._picLEDIN_11.TabIndex = 26
		Me._picLEDIN_11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_11.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_11.CausesValidation = True
		Me._picLEDIN_11.Enabled = True
		Me._picLEDIN_11.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_11.TabStop = True
		Me._picLEDIN_11.Visible = True
		Me._picLEDIN_11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_11.Name = "_picLEDIN_11"
		Me._picLEDIN_12.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_12.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_12.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_12.Location = New System.Drawing.Point(112, 24)
		Me._picLEDIN_12.TabIndex = 25
		Me._picLEDIN_12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_12.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_12.CausesValidation = True
		Me._picLEDIN_12.Enabled = True
		Me._picLEDIN_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_12.TabStop = True
		Me._picLEDIN_12.Visible = True
		Me._picLEDIN_12.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_12.Name = "_picLEDIN_12"
		Me._picLEDIN_13.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_13.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_13.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_13.Location = New System.Drawing.Point(80, 24)
		Me._picLEDIN_13.TabIndex = 24
		Me._picLEDIN_13.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_13.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_13.CausesValidation = True
		Me._picLEDIN_13.Enabled = True
		Me._picLEDIN_13.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_13.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_13.TabStop = True
		Me._picLEDIN_13.Visible = True
		Me._picLEDIN_13.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_13.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_13.Name = "_picLEDIN_13"
		Me._picLEDIN_14.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_14.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_14.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_14.Location = New System.Drawing.Point(48, 24)
		Me._picLEDIN_14.TabIndex = 23
		Me._picLEDIN_14.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_14.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_14.CausesValidation = True
		Me._picLEDIN_14.Enabled = True
		Me._picLEDIN_14.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_14.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_14.TabStop = True
		Me._picLEDIN_14.Visible = True
		Me._picLEDIN_14.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_14.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_14.Name = "_picLEDIN_14"
		Me._picLEDIN_15.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDIN_15.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDIN_15.Size = New System.Drawing.Size(33, 33)
		Me._picLEDIN_15.Location = New System.Drawing.Point(16, 24)
		Me._picLEDIN_15.TabIndex = 22
		Me._picLEDIN_15.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDIN_15.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDIN_15.CausesValidation = True
		Me._picLEDIN_15.Enabled = True
		Me._picLEDIN_15.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDIN_15.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDIN_15.TabStop = True
		Me._picLEDIN_15.Visible = True
		Me._picLEDIN_15.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDIN_15.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDIN_15.Name = "_picLEDIN_15"
		Me.lblInData.BackColor = System.Drawing.SystemColors.Window
		Me.lblInData.Text = "Digital Input Value"
		Me.lblInData.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblInData.Size = New System.Drawing.Size(121, 17)
		Me.lblInData.Location = New System.Drawing.Point(248, 72)
		Me.lblInData.TabIndex = 38
		Me.lblInData.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblInData.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblInData.Enabled = True
		Me.lblInData.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInData.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInData.UseMnemonic = True
		Me.lblInData.Visible = True
		Me.lblInData.AutoSize = False
		Me.lblInData.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInData.Name = "lblInData"
		Me.piclighton.BackColor = System.Drawing.SystemColors.Window
		Me.piclighton.ForeColor = System.Drawing.SystemColors.WindowText
		Me.piclighton.Size = New System.Drawing.Size(33, 33)
		Me.piclighton.Location = New System.Drawing.Point(48, 240)
		Me.piclighton.Image = CType(resources.GetObject("piclighton.Image"), System.Drawing.Image)
		Me.piclighton.TabIndex = 20
		Me.piclighton.Visible = False
		Me.piclighton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.piclighton.Dock = System.Windows.Forms.DockStyle.None
		Me.piclighton.CausesValidation = True
		Me.piclighton.Enabled = True
		Me.piclighton.Cursor = System.Windows.Forms.Cursors.Default
		Me.piclighton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.piclighton.TabStop = True
		Me.piclighton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.piclighton.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.piclighton.Name = "piclighton"
		Me._picLEDOUT_0.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_0.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_0.Location = New System.Drawing.Point(568, 168)
		Me._picLEDOUT_0.TabIndex = 4
		Me._picLEDOUT_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_0.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_0.CausesValidation = True
		Me._picLEDOUT_0.Enabled = True
		Me._picLEDOUT_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_0.TabStop = True
		Me._picLEDOUT_0.Visible = True
		Me._picLEDOUT_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_0.Name = "_picLEDOUT_0"
		Me.Frame2.BackColor = System.Drawing.SystemColors.Window
		Me.Frame2.Text = "Digital Output"
		Me.Frame2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Frame2.Size = New System.Drawing.Size(609, 105)
		Me.Frame2.Location = New System.Drawing.Point(0, 136)
		Me.Frame2.TabIndex = 1
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.Enabled = True
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Name = "Frame2"
		Me._picLEDOUT_15.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_15.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_15.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_15.Location = New System.Drawing.Point(16, 32)
		Me._picLEDOUT_15.TabIndex = 19
		Me._picLEDOUT_15.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_15.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_15.CausesValidation = True
		Me._picLEDOUT_15.Enabled = True
		Me._picLEDOUT_15.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_15.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_15.TabStop = True
		Me._picLEDOUT_15.Visible = True
		Me._picLEDOUT_15.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_15.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_15.Name = "_picLEDOUT_15"
		Me._picLEDOUT_14.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_14.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_14.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_14.Location = New System.Drawing.Point(48, 32)
		Me._picLEDOUT_14.TabIndex = 18
		Me._picLEDOUT_14.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_14.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_14.CausesValidation = True
		Me._picLEDOUT_14.Enabled = True
		Me._picLEDOUT_14.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_14.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_14.TabStop = True
		Me._picLEDOUT_14.Visible = True
		Me._picLEDOUT_14.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_14.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_14.Name = "_picLEDOUT_14"
		Me._picLEDOUT_13.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_13.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_13.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_13.Location = New System.Drawing.Point(80, 32)
		Me._picLEDOUT_13.TabIndex = 17
		Me._picLEDOUT_13.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_13.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_13.CausesValidation = True
		Me._picLEDOUT_13.Enabled = True
		Me._picLEDOUT_13.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_13.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_13.TabStop = True
		Me._picLEDOUT_13.Visible = True
		Me._picLEDOUT_13.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_13.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_13.Name = "_picLEDOUT_13"
		Me._picLEDOUT_12.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_12.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_12.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_12.Location = New System.Drawing.Point(112, 32)
		Me._picLEDOUT_12.TabIndex = 16
		Me._picLEDOUT_12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_12.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_12.CausesValidation = True
		Me._picLEDOUT_12.Enabled = True
		Me._picLEDOUT_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_12.TabStop = True
		Me._picLEDOUT_12.Visible = True
		Me._picLEDOUT_12.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_12.Name = "_picLEDOUT_12"
		Me._picLEDOUT_11.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_11.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_11.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_11.Location = New System.Drawing.Point(168, 32)
		Me._picLEDOUT_11.TabIndex = 15
		Me._picLEDOUT_11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_11.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_11.CausesValidation = True
		Me._picLEDOUT_11.Enabled = True
		Me._picLEDOUT_11.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_11.TabStop = True
		Me._picLEDOUT_11.Visible = True
		Me._picLEDOUT_11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_11.Name = "_picLEDOUT_11"
		Me._picLEDOUT_10.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_10.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_10.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_10.Location = New System.Drawing.Point(200, 32)
		Me._picLEDOUT_10.TabIndex = 14
		Me._picLEDOUT_10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_10.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_10.CausesValidation = True
		Me._picLEDOUT_10.Enabled = True
		Me._picLEDOUT_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_10.TabStop = True
		Me._picLEDOUT_10.Visible = True
		Me._picLEDOUT_10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_10.Name = "_picLEDOUT_10"
		Me._picLEDOUT_9.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_9.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_9.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_9.Location = New System.Drawing.Point(232, 32)
		Me._picLEDOUT_9.TabIndex = 13
		Me._picLEDOUT_9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_9.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_9.CausesValidation = True
		Me._picLEDOUT_9.Enabled = True
		Me._picLEDOUT_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_9.TabStop = True
		Me._picLEDOUT_9.Visible = True
		Me._picLEDOUT_9.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_9.Name = "_picLEDOUT_9"
		Me._picLEDOUT_8.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_8.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_8.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_8.Location = New System.Drawing.Point(264, 32)
		Me._picLEDOUT_8.TabIndex = 12
		Me._picLEDOUT_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_8.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_8.CausesValidation = True
		Me._picLEDOUT_8.Enabled = True
		Me._picLEDOUT_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_8.TabStop = True
		Me._picLEDOUT_8.Visible = True
		Me._picLEDOUT_8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_8.Name = "_picLEDOUT_8"
		Me._picLEDOUT_7.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_7.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_7.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_7.Location = New System.Drawing.Point(320, 32)
		Me._picLEDOUT_7.TabIndex = 11
		Me._picLEDOUT_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_7.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_7.CausesValidation = True
		Me._picLEDOUT_7.Enabled = True
		Me._picLEDOUT_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_7.TabStop = True
		Me._picLEDOUT_7.Visible = True
		Me._picLEDOUT_7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_7.Name = "_picLEDOUT_7"
		Me._picLEDOUT_6.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_6.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_6.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_6.Location = New System.Drawing.Point(352, 32)
		Me._picLEDOUT_6.TabIndex = 10
		Me._picLEDOUT_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_6.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_6.CausesValidation = True
		Me._picLEDOUT_6.Enabled = True
		Me._picLEDOUT_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_6.TabStop = True
		Me._picLEDOUT_6.Visible = True
		Me._picLEDOUT_6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_6.Name = "_picLEDOUT_6"
		Me._picLEDOUT_5.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_5.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_5.Location = New System.Drawing.Point(384, 32)
		Me._picLEDOUT_5.TabIndex = 9
		Me._picLEDOUT_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_5.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_5.CausesValidation = True
		Me._picLEDOUT_5.Enabled = True
		Me._picLEDOUT_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_5.TabStop = True
		Me._picLEDOUT_5.Visible = True
		Me._picLEDOUT_5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_5.Name = "_picLEDOUT_5"
		Me._picLEDOUT_4.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_4.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_4.Location = New System.Drawing.Point(416, 32)
		Me._picLEDOUT_4.TabIndex = 8
		Me._picLEDOUT_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_4.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_4.CausesValidation = True
		Me._picLEDOUT_4.Enabled = True
		Me._picLEDOUT_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_4.TabStop = True
		Me._picLEDOUT_4.Visible = True
		Me._picLEDOUT_4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_4.Name = "_picLEDOUT_4"
		Me._picLEDOUT_3.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_3.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_3.Location = New System.Drawing.Point(472, 32)
		Me._picLEDOUT_3.TabIndex = 7
		Me._picLEDOUT_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_3.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_3.CausesValidation = True
		Me._picLEDOUT_3.Enabled = True
		Me._picLEDOUT_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_3.TabStop = True
		Me._picLEDOUT_3.Visible = True
		Me._picLEDOUT_3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_3.Name = "_picLEDOUT_3"
		Me._picLEDOUT_2.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_2.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_2.Location = New System.Drawing.Point(504, 32)
		Me._picLEDOUT_2.TabIndex = 6
		Me._picLEDOUT_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_2.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_2.CausesValidation = True
		Me._picLEDOUT_2.Enabled = True
		Me._picLEDOUT_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_2.TabStop = True
		Me._picLEDOUT_2.Visible = True
		Me._picLEDOUT_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_2.Name = "_picLEDOUT_2"
		Me._picLEDOUT_1.BackColor = System.Drawing.SystemColors.Window
		Me._picLEDOUT_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picLEDOUT_1.Size = New System.Drawing.Size(33, 33)
		Me._picLEDOUT_1.Location = New System.Drawing.Point(536, 32)
		Me._picLEDOUT_1.TabIndex = 5
		Me._picLEDOUT_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picLEDOUT_1.Dock = System.Windows.Forms.DockStyle.None
		Me._picLEDOUT_1.CausesValidation = True
		Me._picLEDOUT_1.Enabled = True
		Me._picLEDOUT_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._picLEDOUT_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picLEDOUT_1.TabStop = True
		Me._picLEDOUT_1.Visible = True
		Me._picLEDOUT_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._picLEDOUT_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picLEDOUT_1.Name = "_picLEDOUT_1"
		Me.lblOutData.BackColor = System.Drawing.SystemColors.Window
		Me.lblOutData.Text = "Digital Output Value"
		Me.lblOutData.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblOutData.Size = New System.Drawing.Size(121, 25)
		Me.lblOutData.Location = New System.Drawing.Point(248, 72)
		Me.lblOutData.TabIndex = 2
		Me.lblOutData.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOutData.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblOutData.Enabled = True
		Me.lblOutData.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblOutData.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblOutData.UseMnemonic = True
		Me.lblOutData.Visible = True
		Me.lblOutData.AutoSize = False
		Me.lblOutData.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblOutData.Name = "lblOutData"
		Me.piclightoff.BackColor = System.Drawing.SystemColors.Window
		Me.piclightoff.ForeColor = System.Drawing.SystemColors.WindowText
		Me.piclightoff.Size = New System.Drawing.Size(33, 33)
		Me.piclightoff.Location = New System.Drawing.Point(8, 240)
		Me.piclightoff.Image = CType(resources.GetObject("piclightoff.Image"), System.Drawing.Image)
		Me.piclightoff.TabIndex = 0
		Me.piclightoff.Visible = False
		Me.piclightoff.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.piclightoff.Dock = System.Windows.Forms.DockStyle.None
		Me.piclightoff.CausesValidation = True
		Me.piclightoff.Enabled = True
		Me.piclightoff.Cursor = System.Windows.Forms.Cursors.Default
		Me.piclightoff.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.piclightoff.TabStop = True
		Me.piclightoff.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.piclightoff.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.piclightoff.Name = "piclightoff"
		Me.Timer1.Interval = 100
		Me.Timer1.Enabled = True
		DIN.OcxState = CType(resources.GetObject("DIN.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DIN.Location = New System.Drawing.Point(104, 248)
		Me.DIN.Name = "DIN"
		DOUT.OcxState = CType(resources.GetObject("DOUT.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DOUT.Location = New System.Drawing.Point(216, 248)
		Me.DOUT.Name = "DOUT"
		Me.Label1.BackColor = System.Drawing.SystemColors.Window
		Me.Label1.Text = "Click lights to change Digital Output value"
		Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Label1.Size = New System.Drawing.Size(305, 17)
		Me.Label1.Location = New System.Drawing.Point(296, 240)
		Me.Label1.TabIndex = 3
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(Frame3)
		Me.Controls.Add(piclighton)
		Me.Controls.Add(_picLEDOUT_0)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(piclightoff)
		Me.Controls.Add(DIN)
		Me.Controls.Add(DOUT)
		Me.Controls.Add(Label1)
		Me.Frame3.Controls.Add(_picLEDIN_0)
		Me.Frame3.Controls.Add(_picLEDIN_1)
		Me.Frame3.Controls.Add(_picLEDIN_2)
		Me.Frame3.Controls.Add(_picLEDIN_3)
		Me.Frame3.Controls.Add(_picLEDIN_4)
		Me.Frame3.Controls.Add(_picLEDIN_5)
		Me.Frame3.Controls.Add(_picLEDIN_6)
		Me.Frame3.Controls.Add(_picLEDIN_7)
		Me.Frame3.Controls.Add(_picLEDIN_8)
		Me.Frame3.Controls.Add(_picLEDIN_9)
		Me.Frame3.Controls.Add(_picLEDIN_10)
		Me.Frame3.Controls.Add(_picLEDIN_11)
		Me.Frame3.Controls.Add(_picLEDIN_12)
		Me.Frame3.Controls.Add(_picLEDIN_13)
		Me.Frame3.Controls.Add(_picLEDIN_14)
		Me.Frame3.Controls.Add(_picLEDIN_15)
		Me.Frame3.Controls.Add(lblInData)
		Me.Frame2.Controls.Add(_picLEDOUT_15)
		Me.Frame2.Controls.Add(_picLEDOUT_14)
		Me.Frame2.Controls.Add(_picLEDOUT_13)
		Me.Frame2.Controls.Add(_picLEDOUT_12)
		Me.Frame2.Controls.Add(_picLEDOUT_11)
		Me.Frame2.Controls.Add(_picLEDOUT_10)
		Me.Frame2.Controls.Add(_picLEDOUT_9)
		Me.Frame2.Controls.Add(_picLEDOUT_8)
		Me.Frame2.Controls.Add(_picLEDOUT_7)
		Me.Frame2.Controls.Add(_picLEDOUT_6)
		Me.Frame2.Controls.Add(_picLEDOUT_5)
		Me.Frame2.Controls.Add(_picLEDOUT_4)
		Me.Frame2.Controls.Add(_picLEDOUT_3)
		Me.Frame2.Controls.Add(_picLEDOUT_2)
		Me.Frame2.Controls.Add(_picLEDOUT_1)
		Me.Frame2.Controls.Add(lblOutData)
		Me.picLEDIN.SetIndex(_picLEDIN_0, CType(0, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_1, CType(1, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_2, CType(2, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_3, CType(3, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_4, CType(4, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_5, CType(5, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_6, CType(6, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_7, CType(7, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_8, CType(8, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_9, CType(9, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_10, CType(10, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_11, CType(11, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_12, CType(12, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_13, CType(13, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_14, CType(14, Short))
		Me.picLEDIN.SetIndex(_picLEDIN_15, CType(15, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_0, CType(0, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_15, CType(15, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_14, CType(14, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_13, CType(13, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_12, CType(12, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_11, CType(11, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_10, CType(10, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_9, CType(9, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_8, CType(8, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_7, CType(7, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_6, CType(6, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_5, CType(5, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_4, CType(4, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_3, CType(3, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_2, CType(2, Short))
		Me.picLEDOUT.SetIndex(_picLEDOUT_1, CType(1, Short))
		CType(Me.picLEDOUT, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.picLEDIN, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DOUT, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DIN, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As Form1
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As Form1
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New Form1()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Dim bit(16) As Integer
	Dim InValue As Object
	Dim OutValue As Integer
	
	
	'This example demonstrates the use of single value data flow
	'mode with both a Digital Input and a Digital Output subsystem.
	'As the user clicks on the light icons of the Digital Output box
	'the Digital output value is changed and written to the DOUT
	'subsystem.
	'A timer control is used to read the digital input data at
	'regular time intervals.  The light icons in the digital input
	'box will appear to turn on and off, depending of the
	'value read from the DIN subsystem.
	'
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim Prompt, Title As String
		Dim NL As String
		Dim BoardName As String
		Dim i As Object
		Dim numss As Short
		
		'fill array of binary values used for controlling
		'on/off of light icons
		bit(0) = 1
		For i = 1 To 15
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			bit(i) = bit(i - 1) * 2
		Next i
		
		'if the dtol control was not configured at design time,
		'configure it now
		If DIN.hDass = 0 Or DOUT.hDass = 0 Then
			On Error GoTo Trap
			
			'ask user for board name
			NL = Chr(10)
			Title = "Digital Input/Output"
			Prompt = "Please Enter Board Name" & NL & NL
			Prompt = Prompt & "(use the name assigned at install time)"
			BoardName = InputBox(Prompt, Title)
			If BoardName = "" Then End
			
			DIN.Board = BoardName
			DIN.SubSysType = OLSS_DIN 'set type = DIN
			DIN.SubSysElement = 0
			
			DOUT.Board = DIN.Board
			DOUT.SubSysType = OLSS_DOUT 'set type = DOUT
			'determine # of Digital Output subsystems on device
			numss = DOUT.GetDevCaps(OLDC_DOUTELEMENTS)
			
			'select last element
			DOUT.SubSysElement = numss - 1
			
			On Error GoTo 0
		End If
		On Error GoTo 0
		
		'set up DIN
		DIN.DataFlow = OL_DF_SINGLEVALUE 'set flow for single value
		DIN.Config()
		'UPGRADE_WARNING: Couldn't resolve default property of object InValue. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		InValue = 0 'set initial input value
		'UPGRADE_WARNING: Couldn't resolve default property of object InValue. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lblInData.Text = Hex(InValue) & " Hex"
		
		'display only icons upto the # of bits of Resolution
		For i = 0 To 15
			PicLEDIN(i).Image = piclightoff.Image
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If i > DIN.Resolution - 1 Then
				PicLEDIN(i).Visible = False
			End If
		Next i
		
		
		'set up DOUT
		DOUT.DataFlow = OL_DF_SINGLEVALUE 'set flow for single value
		OutValue = 0 ' set initial output value
		lblOutData.Text = Hex(OutValue) & " Hex"
		
		'write initial value to Digital Output subsystem
		DOUT.Config()
		DOUT.PutSingleValue(0, 1, OutValue)
		
		'display only icons upto the # of bits of Resolution
		For i = 0 To 15
			picLEDOUT(i).Image = piclightoff.Image
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If i > DOUT.Resolution - 1 Then
				picLEDOUT(i).Visible = False
			End If
		Next i
		
		Exit Sub
		
Trap: 
		MsgBox("Board/Subsystem not available", 48, Title)
		End
	End Sub
	
	Private Sub picLEDOUT_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picLEDOUT.Click
		Dim Index As Short = picLEDOUT.GetIndex(eventSender)
		
		'switch on/off light & toggle corresponding bit
		'of output value
		
		'turn off
		If picLEDOUT(Index).Image.equals(piclighton.Image) Then
			picLEDOUT(Index).Image = piclightoff.Image
			OutValue = OutValue And (Not bit(Index))
			'turn on
		Else
			picLEDOUT(Index).Image = piclighton.Image
			OutValue = OutValue Or bit(Index)
		End If
		
		'display hexidecimal notation of digital output value
		lblOutData.Text = Hex(OutValue) & " Hex"
		
		'write value to digital output subsystem
		DOUT.PutSingleValue(0, 1, OutValue)
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		Dim oldvalue As Integer
		Dim i As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object InValue. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		oldvalue = InValue 'save current digital input value
		'UPGRADE_WARNING: Couldn't resolve default property of object InValue. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		InValue = DIN.GetSingleValue(0, 1)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object InValue. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lblInData.Text = Hex(InValue) & " Hex" 'display value in hexidecimal notation
		
		'if value hasn't changed,
		'then don't change light settings
		'UPGRADE_WARNING: Couldn't resolve default property of object InValue. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If InValue = oldvalue Then
			Exit Sub
		End If
		
		'turn on/off lights
		For i = 0 To 15
			'bit set ?
			If InValue And bit(i) Then
				PicLEDIN(i).Image = piclighton.Image
			Else
				PicLEDIN(i).Image = piclightoff.Image
			End If
			
		Next i
		
		
	End Sub
End Class