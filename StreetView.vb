Imports System.Text.RegularExpressions

Public Class StreetView
    Private Sub btnEncode_Click(sender As Object, e As EventArgs) Handles btnEncode.Click
        ' Decode, and recode, a Google Street View URL
        decode()    ' encode contents of control
    End Sub

    Private Sub btnPaste_Click(sender As Object, e As EventArgs) Handles btnPaste.Click
        Me.tbEncodedURL.Text = Clipboard.GetText    ' copy clipboard to control
        decode()    ' encode contents of control
    End Sub
    Private Sub decode()
        ' Typical link https://www.google.com/maps/@-38.1363332,145.1300837,3a,40.8y,202.9h,92.66t/data=!3m6!1e1!3m4!1sfo0PetpJ-gCu30Kkmy_mtg!2e0!7i16384!8i8192
        ' Doco at https://www.trekview.org/blog/2020/decoding-google-street-view-urls/

        Dim pattern As String = "@([-\.0-9]+),([-\.0-9]+),([\.\d]+)a,([\.\d]+)y,([\.\d]+)h,([\.\d]+)t"
        Dim matches As MatchCollection
        Dim regex As New Regex(pattern)

        matches = regex.Matches(Me.tbEncodedURL.Text)     ' analyse encoded URL
        If matches.Count = 1 Then
            Dim groups As GroupCollection = matches(0).Groups
            Debug.Assert(groups.Count = 7, "Something wrong with regexp")
            Dim lat As Single = CSng(groups(1).Value)
            Dim lon As Single = CSng(groups(2).Value)
            Dim fov As Integer = CInt(groups(4).Value)
            Dim heading As Integer = CInt(groups(5).Value)
            Dim pitch As Integer = CInt(groups(6).Value) - 90   ' Why -90 ?
            Dim url As String = $"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={lat:f6},{lon:f6}&fov={fov }&heading={heading }&pitch={pitch }"
            Me.tbDecodedURL.Text = url
            Clipboard.SetText(url)     ' copy to clipboard
        Else
            Me.tbDecodedURL.Text = "Not a valid Street View URL"
        End If
    End Sub
End Class