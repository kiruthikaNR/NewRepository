Public Class Form2


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myFileDlog2 As New OpenFileDialog()
        myFileDlog2.InitialDirectory = "c:\"

        'specifies what type of data files to look for
        myFileDlog2.Filter = "All Files (*.*)|*.*" & _
            "|Zip Files (*.zip)|*.zip"

        'specifies which data type is focused on start up
        myFileDlog2.FilterIndex = 2

        'Gets or sets a value indicating whether the dialog box restores the current directory before closing.
        myFileDlog2.RestoreDirectory = True

        'seperates message outputs for files found or not found
        If myFileDlog2.ShowDialog() = _
            DialogResult.OK Then
            If Dir(myFileDlog2.FileName) = "" Then
                MsgBox("File Not Found", _
                       MsgBoxStyle.Critical)
            End If
        End If

        'Adds the file directory to the text box
        eagleFile.Text = myFileDlog2.FileName
    End Sub

    Private Sub combotextchanged(ByVal sender As Object, ByVal _e As EventArgs) Handles ComboBox2.TextChanged
        Dim cb As ComboBox = CType(sender, ComboBox)
        Dim selitem = CType(ComboBox2.SelectedItem, String)
        If selitem = "Yes" Then
            ErepLabel.Visible = True
            eagleFile.Visible = True
            Button1.Visible = True
        Else
            ErepLabel.Visible = False
            eagleFile.Visible = False
            Button1.Visible = False
        End If


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim myFileDlog2 As New OpenFileDialog()
        myFileDlog2.InitialDirectory = "c:\"

        'specifies what type of data files to look for
        myFileDlog2.Filter = "All Files (*.*)|*.*" & _
            "|Zip Files (*.zip)|*.zip"

        'specifies which data type is focused on start up
        myFileDlog2.FilterIndex = 2

        'Gets or sets a value indicating whether the dialog box restores the current directory before closing.
        myFileDlog2.RestoreDirectory = True

        'seperates message outputs for files found or not found
        If myFileDlog2.ShowDialog() = _
            DialogResult.OK Then
            If Dir(myFileDlog2.FileName) = "" Then
                MsgBox("File Not Found", _
                       MsgBoxStyle.Critical)
            End If
        End If

        'Adds the file directory to the text box
        TextBox2.Text = myFileDlog2.FileName
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim myFileDlog2 As New OpenFileDialog()
        myFileDlog2.InitialDirectory = "c:\"

        'specifies what type of data files to look for
        myFileDlog2.Filter = "All Files (*.*)|*.*" & _
            "|Zip Files (*.zip)|*.zip"

        'specifies which data type is focused on start up
        myFileDlog2.FilterIndex = 2

        'Gets or sets a value indicating whether the dialog box restores the current directory before closing.
        myFileDlog2.RestoreDirectory = True

        'seperates message outputs for files found or not found
        If myFileDlog2.ShowDialog() = _
            DialogResult.OK Then
            If Dir(myFileDlog2.FileName) = "" Then
                MsgBox("File Not Found", _
                       MsgBoxStyle.Critical)
            End If
        End If

        'Adds the file directory to the text box
        TextBox3.Text = myFileDlog2.FileName
    End Sub

    Private Sub SourceDatafileBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SourceDatafileBox.SelectedIndexChanged

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Hide()
        Form1.Show()
    End Sub


    Private Sub rfNote_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class