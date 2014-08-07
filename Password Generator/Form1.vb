Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = Generatecode()

    End Sub

    Public Function Generatecode() As Object
        Dim intRnd As Object
        Dim strName As Object
        Dim intNameLength As Object
        Dim intLenght As Object
        Dim strInputString As Object
        Dim inStep As Object
        

        strInputString = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

        intLenght = Len(strInputString)

        intNameLength = 8

        Randomize()

        strName = ""

        For inStep = 1 To intNameLength

            intRnd = Int((intLenght * Rnd()) + 1)

            strName = strName & Mid(strInputString, intRnd, 1)

        Next

        Generatecode = strName

    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = Generatecode2()

    End Sub

    Public Function Generatecode2() As Object
        Dim intRnd As Object
        Dim strName As Object
        Dim intNameLength As Object
        Dim intLenght As Object
        Dim strInputString As Object
        Dim inStep As Object


        strInputString = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!@#$$%^&*(){}|\][;:',<.>/?=+-_"

        intLenght = Len(strInputString)

        intNameLength = 8

        Randomize()

        strName = ""

        For inStep = 1 To intNameLength

            intRnd = Int((intLenght * Rnd()) + 1)

            strName = strName & Mid(strInputString, intRnd, 1)

        Next

        Generatecode2 = strName
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        MsgBox("Text has been copied to your clipboard")
        Clipboard.SetDataObject(TextBox1.Text, True)
       
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        Form2.Show()

    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        Form3.Show()


    End Sub
End Class








