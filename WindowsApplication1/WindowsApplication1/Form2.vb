Public Class Form2


    
   

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
        'TextBox2.Text = myFileDlog2.FileName
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
        'TextBox3.Text = myFileDlog2.FileName
    End Sub

    Private Sub SourceDatafileBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Hide()
        Form1.Show()
    End Sub


    Private Sub rfNote_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub
End Class