Public Class Form3
    Public ctrls As List(Of KeyValuePair(Of String, CheckBox)) = New List(Of KeyValuePair(Of String, CheckBox))
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Programatically create a state selection dialog
        Dim x As Integer, y As Integer, inc As Integer, cb As CheckBox

        If Not ctrls.Any Then
            x = 30 : y = 30 : inc = 20
            With Me
                For Each st In Form1.states
                    cb = New CheckBox
                    With cb
                        .Text = st : .Location = New Point(x, y)
                    End With
                    .Controls.Add(cb)
                    ctrls.Add(New KeyValuePair(Of String, CheckBox)(st, cb))
                    y += inc
                Next
                cb = New CheckBox
                With cb
                    .Text = "All" : .Location = New Point(x, y)
                End With
                .Controls.Add(cb)
                ctrls.Add(New KeyValuePair(Of String, CheckBox)("All", cb))
            End With
        End If
    End Sub
    Public Function Selected() As List(Of String)
        ' Return a list of selected checkboxes
        Dim result As New List(Of String)
        For i = 0 To Me.ctrls.Count - 2
            Dim box As KeyValuePair(Of String, CheckBox) = Me.ctrls.Item(i)    ' extract the checkbox
            If box.Value.Checked Or Me.ctrls.Item(Me.ctrls.Count - 1).Value.Checked Then
                result.Add(box.Key)
            End If
        Next
        Return result
    End Function
End Class