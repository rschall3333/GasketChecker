Option Strict Off
Option Explicit On
Friend Class BoardChooser
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
	Public WithEvents Board_ok As System.Windows.Forms.Button
	Public WithEvents Combo1 As System.Windows.Forms.ComboBox
	Public WithEvents selectboard As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(BoardChooser))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Board_ok = New System.Windows.Forms.Button
		Me.Combo1 = New System.Windows.Forms.ComboBox
		Me.selectboard = New System.Windows.Forms.Label
		Me.Text = "Data Translation"
		Me.ClientSize = New System.Drawing.Size(392, 159)
		Me.Location = New System.Drawing.Point(4, 31)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
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
		Me.Name = "BoardChooser"
		Me.Board_ok.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Board_ok.Text = "OK"
		Me.Board_ok.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Board_ok.Size = New System.Drawing.Size(41, 25)
		Me.Board_ok.Location = New System.Drawing.Point(272, 80)
		Me.Board_ok.TabIndex = 2
		Me.Board_ok.BackColor = System.Drawing.SystemColors.Control
		Me.Board_ok.CausesValidation = True
		Me.Board_ok.Enabled = True
		Me.Board_ok.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Board_ok.Cursor = System.Windows.Forms.Cursors.Default
		Me.Board_ok.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Board_ok.TabStop = True
		Me.Board_ok.Name = "Board_ok"
		Me.Combo1.Size = New System.Drawing.Size(225, 21)
		Me.Combo1.Location = New System.Drawing.Point(32, 80)
		Me.Combo1.TabIndex = 1
		Me.Combo1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo1.BackColor = System.Drawing.SystemColors.Window
		Me.Combo1.CausesValidation = True
		Me.Combo1.Enabled = True
		Me.Combo1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Combo1.IntegralHeight = True
		Me.Combo1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Combo1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Combo1.Sorted = False
		Me.Combo1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.Combo1.TabStop = True
		Me.Combo1.Visible = True
		Me.Combo1.Name = "Combo1"
		Me.selectboard.Text = "Select a Data Translation Board:"
		Me.selectboard.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.selectboard.Size = New System.Drawing.Size(289, 33)
		Me.selectboard.Location = New System.Drawing.Point(16, 16)
		Me.selectboard.TabIndex = 0
		Me.selectboard.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.selectboard.BackColor = System.Drawing.SystemColors.Control
		Me.selectboard.Enabled = True
		Me.selectboard.ForeColor = System.Drawing.SystemColors.ControlText
		Me.selectboard.Cursor = System.Windows.Forms.Cursors.Default
		Me.selectboard.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.selectboard.UseMnemonic = True
		Me.selectboard.Visible = True
		Me.selectboard.AutoSize = False
		Me.selectboard.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.selectboard.Name = "selectboard"
		Me.Controls.Add(Board_ok)
		Me.Controls.Add(Combo1)
		Me.Controls.Add(selectboard)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As BoardChooser
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As BoardChooser
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New BoardChooser()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Private Sub Board_ok_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Board_ok.Click
		Form1.DefInstance.ADC.Board = Combo1.Text
		Form1.DefInstance.CT.Board = Combo1.Text
		Form1.DefInstance.DIO.Board = Combo1.Text
		Form1.DefInstance.Enabled = True
		Hide()
	End Sub
	
	Private Sub BoardChooser_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim DIO As Object
		Dim ADC As Object
		Dim CT As Object
		Dim BoardIndex As Object
		Dim i As Short
		
		'initialize list of boards
		For i = 0 To Form1.DefInstance.ADC.numBoards - 1
			Combo1.Items.Add(Form1.DefInstance.ADC.get_BoardList(i))
		Next i
		
		'UPGRADE_WARNING: Couldn't resolve default property of object BoardIndex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		BoardIndex = -1
		Combo1.SelectedIndex = 0
		Exit Sub
		
starterr: 
		'UPGRADE_WARNING: Couldn't resolve default property of object CT.LastError. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If CT.LastError <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object CT.LastErrDescription. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			MsgBox(CT.LastErrDescription, 48, "Error")
			'UPGRADE_WARNING: Couldn't resolve default property of object CT.ClearError. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			CT.ClearError()
			'UPGRADE_WARNING: Couldn't resolve default property of object ADC.LastError. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf ADC.LastError <> 0 Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object ADC.LastErrDescription. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			MsgBox(ADC.LastErrDescription, 48, "Error")
			'UPGRADE_WARNING: Couldn't resolve default property of object ADC.ClearError. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			ADC.ClearError()
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object DIO.LastErrDescription. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			MsgBox(DIO.LastErrDescription, 48, "Error")
			'UPGRADE_WARNING: Couldn't resolve default property of object DIO.ClearError. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			DIO.ClearError()
		End If
		End
		
	End Sub
End Class