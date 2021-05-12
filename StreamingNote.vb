Public Class Note
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
    Friend WithEvents StreamingNoteOKButton As System.Windows.Forms.Button
    Friend WithEvents StreamingNoteTextBox As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.StreamingNoteTextBox = New System.Windows.Forms.TextBox
        Me.StreamingNoteOKButton = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'StreamingNoteTextBox
        '
        Me.StreamingNoteTextBox.AutoSize = False
        Me.StreamingNoteTextBox.Location = New System.Drawing.Point(8, 8)
        Me.StreamingNoteTextBox.Name = "StreamingNoteTextBox"
        Me.StreamingNoteTextBox.Size = New System.Drawing.Size(560, 168)
        Me.StreamingNoteTextBox.TabIndex = 0
        Me.StreamingNoteTextBox.Text = "Add Note Here"
        '
        'StreamingNoteOKButton
        '
        Me.StreamingNoteOKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StreamingNoteOKButton.Location = New System.Drawing.Point(208, 184)
        Me.StreamingNoteOKButton.Name = "StreamingNoteOKButton"
        Me.StreamingNoteOKButton.Size = New System.Drawing.Size(159, 56)
        Me.StreamingNoteOKButton.TabIndex = 1
        Me.StreamingNoteOKButton.Text = "OK"
        '
        'Note
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(576, 244)
        Me.ControlBox = False
        Me.Controls.Add(Me.StreamingNoteOKButton)
        Me.Controls.Add(Me.StreamingNoteTextBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "Note"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Note"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub StreamingNote_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        StreamingNoteTextBox.Text = "Add Note Here"
    End Sub

    Private Sub StreamingNoteTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StreamingNoteTextBox.TextChanged
        If StreamDataNoteBool Then
            StreamingDataNote = StreamingNoteTextBox.Text
        ElseIf ProductionReportBool Then
            ProductionReportNote = "NOTE: " & StreamingNoteTextBox.Text
        End If

    End Sub

    Private Sub StreamingNoteOKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StreamingNoteOKButton.Click
        If StreamDataNoteBool And ProductionCollection Then
            writeProductionPacket()
        ElseIf StreamDataNoteBool And ValidationCollection Then
            writeValidationPacket()
        End If
        Me.Close()
    End Sub
End Class
