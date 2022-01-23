Imports System.IO
Public Class frmHelp
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Load instructions
        WebBrowser1.DocumentText = File.ReadAllText("Help.htm")
    End Sub
End Class