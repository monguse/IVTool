Public Class frm_IVTool

    Public Sub Cout(iStr As String)
        Me.rtb_Messages.AppendText(iStr + vbNewLine)
    End Sub

    Private Sub btn_Start_Click(sender As Object, e As EventArgs) Handles btn_Start.Click
        Cout("Searching...")
        Me.btn_Start.Enabled = False
        IVTouch.ProcessAssembly(tb_DocNumber.Text, Me.cb_Recursive.CheckState, Me.cb_ReplaceTB.CheckState)
        Me.btn_Start.Enabled = True
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
End Class
