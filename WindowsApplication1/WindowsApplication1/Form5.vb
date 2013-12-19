Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.IO
Imports Word = Microsoft.Office.Interop.Word
Imports System.Configuration
Imports System.Configuration.ConfigurationSettings
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Public Class Form5

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim FolderBrowserDialog1 As New FolderBrowserDialog
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim openFileDialog2 As New OpenFileDialog
        openFileDialog2.Multiselect = False
        If openFileDialog2.ShowDialog = DialogResult.OK Then
            For x = 0 To openFileDialog2.FileNames.Count - 1
                TextBox2.Text = openFileDialog2.FileNames(x)
            Next
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim openFileDialog3 As New OpenFileDialog
        openFileDialog3.Multiselect = False
        If openFileDialog3.ShowDialog = DialogResult.OK Then
            For x = 0 To openFileDialog3.FileNames.Count - 1
                TextBox3.Text = openFileDialog3.FileNames(x)
            Next
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim openFileDialog4 As New OpenFileDialog
        openFileDialog4.Multiselect = False
        If openFileDialog4.ShowDialog = DialogResult.OK Then
            For x = 0 To openFileDialog4.FileNames.Count - 1
                TextBox4.Text = openFileDialog4.FileNames(x)
            Next
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If (TextBox1.Text <> "") Then
            If File.Exists(TextBox1.Text & "\A_JE_PREP.LOG") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN THE A LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
            If File.Exists(TextBox1.Text & "\A_JE_PREP.zip") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN Zipped A LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
            If File.Exists(TextBox1.Text & "\B_TB_PREP.LOG") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN THE B LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
            If File.Exists(TextBox1.Text & "\B_TB_PREP.zip") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN Zipped B LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
            If File.Exists(TextBox1.Text & "\C_WORKLOG.LOG") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN THE C LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
            If File.Exists(TextBox1.Text & "\C_WORKLOG.zip") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN Zipped C LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
            If File.Exists(TextBox1.Text & "\D_JE_ROLL.LOG") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN THE D LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
            If File.Exists(TextBox1.Text & "\D_JE_ROLL.zip") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN Zipped D LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
            If File.Exists(TextBox1.Text & "\E_EXPORTS.LOG") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN THE E LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
            If File.Exists(TextBox1.Text & "\E_EXPORTS.zip") = False Then
                MessageBox.Show(StrConv("FOLDER DOES NOT CONTAIN Zipped E LOG", vbProperCase), "Incorrect Entry!!")
                TextBox1.Focus()
                Exit Sub
            End If
        ElseIf (TextBox1.Text = "") Then
            MessageBox.Show(StrConv("PLEASE select the work folder", vbProperCase), "Incomplete Entry!!")
            Exit Sub
        End If

        'If (TextBox2.Text <> "") Then
        '    If (TextBox2.Text.Substring(TextBox2.Text.Length - 13, 9) <> "B_TB_PREP") Then
        '        MessageBox.Show(StrConv("PLEASE ENTER THE CORRECT B LOG", vbProperCase), "Incorrect Entry!!")
        '        TextBox2.Focus()
        '        Exit Sub
        '    End If
        'ElseIf (TextBox2.Text = "") Then
        '    MessageBox.Show(StrConv("PLEASE ENTER B LOG", vbProperCase), "Incorrect Entry!!")
        '    Exit Sub
        'End If

        'If (TextBox3.Text <> "") Then
        '    If (TextBox3.Text.Substring(TextBox3.Text.Length - 13, 9) <> "C_WORKLOG") Then
        '        MessageBox.Show(StrConv("PLEASE ENTER THE CORRECT C LOG", vbProperCase), "Incorrect Entry!!")
        '        TextBox3.Focus()
        '        Exit Sub
        '    End If
        'ElseIf (TextBox3.Text = "") Then
        '    MessageBox.Show(StrConv("PLEASE ENTER C LOG", vbProperCase), "Incorrect Entry!!")
        '    Exit Sub
        'End If

        'If (TextBox4.Text <> "") Then
        '    If (TextBox4.Text.Substring(TextBox4.Text.Length - 13, 9) <> "D_JE_ROLL") Then
        '        MessageBox.Show(StrConv("PLEASE ENTER THE CORRECT D LOG", vbProperCase), "Incorrect Entry!!")
        '        TextBox4.Focus()
        '        Exit Sub
        '    End If
        'ElseIf (TextBox4.Text = "") Then
        '    MessageBox.Show(StrConv("PLEASE ENTER D LOG", vbProperCase), "Incorrect Entry!!")
        '    Exit Sub
        'End If

        'If (TextBox7.Text <> "") Then
        '    If (TextBox7.Text.Substring(TextBox7.Text.Length - 13, 9) <> "E_EXPORTS") Then
        '        MessageBox.Show(StrConv("PLEASE ENTER THE CORRECT E LOG", vbProperCase), "Incorrect Entry!!")
        '        TextBox7.Focus()
        '        Exit Sub
        '    End If
        'ElseIf (TextBox7.Text = "") Then
        '    MessageBox.Show(StrConv("PLEASE ENTER E LOG", vbProperCase), "Incorrect Entry!!")
        '    Exit Sub
        'End If

        If (TextBox5.Text <> "") Then
            If (UCase(TextBox5.Text.Substring(TextBox5.Text.Length - 4)) <> ".ZIP") Then
                MessageBox.Show(StrConv("PLEASE ENTER THE CORRECT ROLLFORWARD attachment", vbProperCase), "Incorrect Entry!!")
                TextBox5.Focus()
                Exit Sub
            End If
        ElseIf (TextBox5.Text = "") Then
            MessageBox.Show(StrConv("Please Enter the Rollforward Sheet", vbProperCase), "Incorrect Entry!!")
            TextBox5.Focus()
            Exit Sub
        End If

        If (TextBox6.Text <> "") Then
            If (UCase(TextBox6.Text.Substring(TextBox6.Text.Length - 4)) <> ".ACL") Then
                MessageBox.Show(StrConv("PLEASE ENTER THE CORRECT ACL attachment", vbProperCase), "Incorrect Entry!!")
                TextBox6.Focus()
                Exit Sub
            End If
        ElseIf (TextBox6.Text = "") Then
            MessageBox.Show(StrConv("Please Enter the ACL", vbProperCase), "Incorrect Entry!!")
            TextBox6.Focus()
            Exit Sub
        End If

        Me.Hide()
        Form1.Focus()
        'Form1.Button3.Select()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim openFileDialog5 As New OpenFileDialog
        openFileDialog5.Multiselect = False
        If openFileDialog5.ShowDialog = DialogResult.OK Then
            For x = 0 To openFileDialog5.FileNames.Count - 1
                TextBox5.Text = openFileDialog5.FileNames(x)
            Next
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim openFileDialog6 As New OpenFileDialog
        openFileDialog6.Multiselect = False
        If openFileDialog6.ShowDialog = DialogResult.OK Then
            For x = 0 To openFileDialog6.FileNames.Count - 1
                TextBox6.Text = openFileDialog6.FileNames(x)
            Next
        End If
    End Sub


    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim openFileDialog7 As New OpenFileDialog
        openFileDialog7.Multiselect = False
        If openFileDialog7.ShowDialog = DialogResult.OK Then
            For x = 0 To openFileDialog7.FileNames.Count - 1
                TextBox7.Text = openFileDialog7.FileNames(x)
            Next
        End If
    End Sub
End Class