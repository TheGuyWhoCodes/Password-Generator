Public Class Form2

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim theWebSite As String
        theWebSite = "http://www.Zohaibart.blogspot.com"
        Call Shell("explorer.exe " & theWebSite, vbNormalFocus)
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)

    End Sub
End Class