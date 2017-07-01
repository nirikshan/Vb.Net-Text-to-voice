Public Class Form1
    Dim SAPI As Object

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SAPI = CreateObject("SAPI.spvoice")
        For i As Integer = 1 To 100
            System.Threading.Thread.Sleep(2)
            ProgressBar1.Value = i
        Next

        SAPI.Speak(RichTextBox1.Text)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
