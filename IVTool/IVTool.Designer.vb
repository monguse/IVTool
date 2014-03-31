<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_IVTool
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tb_DocNumber = New System.Windows.Forms.TextBox()
        Me.btn_Start = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rtb_Messages = New System.Windows.Forms.RichTextBox()
        Me.cb_ReplaceTB = New System.Windows.Forms.CheckBox()
        Me.btn_DocSelect = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tb_FolderPath = New System.Windows.Forms.TextBox()
        Me.btn_FolderSelect = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Select Document"
        '
        'tb_DocNumber
        '
        Me.tb_DocNumber.Location = New System.Drawing.Point(8, 25)
        Me.tb_DocNumber.Name = "tb_DocNumber"
        Me.tb_DocNumber.Size = New System.Drawing.Size(297, 20)
        Me.tb_DocNumber.TabIndex = 1
        '
        'btn_Start
        '
        Me.btn_Start.Location = New System.Drawing.Point(311, 98)
        Me.btn_Start.Name = "btn_Start"
        Me.btn_Start.Size = New System.Drawing.Size(75, 36)
        Me.btn_Start.TabIndex = 3
        Me.btn_Start.Text = "Start"
        Me.btn_Start.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rtb_Messages)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 144)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(378, 232)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Messages"
        '
        'rtb_Messages
        '
        Me.rtb_Messages.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.rtb_Messages.HideSelection = False
        Me.rtb_Messages.Location = New System.Drawing.Point(6, 19)
        Me.rtb_Messages.Name = "rtb_Messages"
        Me.rtb_Messages.ReadOnly = True
        Me.rtb_Messages.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.rtb_Messages.Size = New System.Drawing.Size(366, 207)
        Me.rtb_Messages.TabIndex = 0
        Me.rtb_Messages.Text = ""
        '
        'cb_ReplaceTB
        '
        Me.cb_ReplaceTB.AutoSize = True
        Me.cb_ReplaceTB.Location = New System.Drawing.Point(6, 19)
        Me.cb_ReplaceTB.Name = "cb_ReplaceTB"
        Me.cb_ReplaceTB.Size = New System.Drawing.Size(115, 17)
        Me.cb_ReplaceTB.TabIndex = 4
        Me.cb_ReplaceTB.Text = "Replace Titleblock"
        Me.cb_ReplaceTB.UseVisualStyleBackColor = True
        '
        'btn_DocSelect
        '
        Me.btn_DocSelect.Location = New System.Drawing.Point(311, 23)
        Me.btn_DocSelect.Name = "btn_DocSelect"
        Me.btn_DocSelect.Size = New System.Drawing.Size(75, 23)
        Me.btn_DocSelect.TabIndex = 5
        Me.btn_DocSelect.Text = "Browse..."
        Me.btn_DocSelect.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Select Output Folder"
        '
        'tb_FolderPath
        '
        Me.tb_FolderPath.Location = New System.Drawing.Point(8, 64)
        Me.tb_FolderPath.Name = "tb_FolderPath"
        Me.tb_FolderPath.Size = New System.Drawing.Size(297, 20)
        Me.tb_FolderPath.TabIndex = 2
        '
        'btn_FolderSelect
        '
        Me.btn_FolderSelect.Location = New System.Drawing.Point(311, 62)
        Me.btn_FolderSelect.Name = "btn_FolderSelect"
        Me.btn_FolderSelect.Size = New System.Drawing.Size(75, 23)
        Me.btn_FolderSelect.TabIndex = 8
        Me.btn_FolderSelect.Text = "Browse..."
        Me.btn_FolderSelect.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cb_ReplaceTB)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 90)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(297, 48)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Options"
        '
        'frm_IVTool
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(393, 383)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.btn_FolderSelect)
        Me.Controls.Add(Me.tb_FolderPath)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btn_DocSelect)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btn_Start)
        Me.Controls.Add(Me.tb_DocNumber)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_IVTool"
        Me.Text = "Batch Duwoofer"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tb_DocNumber As System.Windows.Forms.TextBox
    Friend WithEvents btn_Start As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rtb_Messages As System.Windows.Forms.RichTextBox
    Friend WithEvents cb_ReplaceTB As System.Windows.Forms.CheckBox
    Friend WithEvents btn_DocSelect As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tb_FolderPath As System.Windows.Forms.TextBox
    Friend WithEvents btn_FolderSelect As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox

End Class
