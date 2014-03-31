Public Class frm_IVTool

    Public Sub Cout(iStr As String)
        Me.rtb_Messages.AppendText(iStr)
    End Sub

    Private Sub btn_Start_Click(sender As Object, e As EventArgs) Handles btn_Start.Click
        If tb_DocNumber.Text <> "" And tb_FolderPath.Text <> "" Then
            Me.btn_Start.Enabled = False
            Me.Cursor = Cursors.WaitCursor
            IVTouch.ProcessAssembly()
            Me.btn_Start.Enabled = True
            Me.Cursor = Cursors.Default
        Else
            If tb_DocNumber.Text = "" Then
                Cout("Error: No document selected" & vbNewLine)
            End If
            If tb_FolderPath.Text = "" Then
                Cout("Error: No folder selected" & vbNewLine)
            End If
        End If

    End Sub

    Private Sub tb_DocNumber_Leave(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tb_DocNumber.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(DirectCast(sender, System.Windows.Forms.TextBox), True, True, False, True)
        End If
    End Sub

    Private Sub frm_IVTool_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        IVTouch.LoadIV()
    End Sub

    Private Sub frm_IVTool_Closing(sender As Object, e As EventArgs) Handles MyBase.FormClosing
        IVTouch.CloseIV()
    End Sub

    Private Sub btn_DocSelect_Click(sender As Object, e As EventArgs) Handles btn_DocSelect.Click
        Dim fd As OpenFileDialog = New OpenFileDialog

        fd.Title = "Select Document"
        fd.InitialDirectory = "C:\VaultWorkspace\Designs"
        fd.Filter = "Inventor Files(*.iam;*.ipt)|*.iam;*.ipt|All Files(*.*)|*.*"
        fd.FilterIndex = 1
        fd.RestoreDirectory = True

        If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            tb_DocNumber.Text = fd.FileName
        End If
    End Sub

    Private Sub btn_FolderSelect_Click(sender As Object, e As EventArgs) Handles btn_FolderSelect.Click
        Dim fd As FolderBrowserDialog = New FolderBrowserDialog

        fd.RootFolder = Environment.SpecialFolder.MyComputer
        fd.Description = "Select Folder"
        fd.ShowNewFolderButton = True

        If fd.ShowDialog = Windows.Forms.DialogResult.OK Then
            tb_FolderPath.Text = fd.SelectedPath
        End If
    End Sub
End Class
