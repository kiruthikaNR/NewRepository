Imports System.Text.RegularExpressions
Public Class Form6

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If (TextBox1.Text = "") Then
            MessageBox.Show("Please enter the engagement code", "Incomplete!!")
            TextBox1.Focus()
            Exit Sub
        ElseIf Regex.IsMatch(TextBox1.Text, "^[0-9 ]+$") = False Then
            MessageBox.Show("Engagement Code can only contain numbers", "Invalid Entry!!")
            TextBox1.Focus()
            Exit Sub
        End If
        If (spoc.Text = "") Then
            MessageBox.Show("Please enter the name of the DA Spoc", "Incomplete!!")
            spoc.Focus()
            Exit Sub
            'ElseIf Regex.IsMatch(TextBox2.Text, "[a-z][A-Z]") = False Then
        Else
            For a = 0 To 9
                If spoc.Text.ToString.Contains(Trim(Str(a))) = True Then
                    MessageBox.Show("DA Spoc name can only contain alphabets", "Invalid Entry!!")
                    spoc.Focus()
                    Exit Sub
                End If
            Next a
        End If
        If (ucr.Text = "") Then
            MessageBox.Show("Please enter the name of the US CAAT Reviewer", "Incomplete!!")
            ucr.Focus()
            Exit Sub
            'ElseIf Regex.IsMatch(TextBox2.Text, "[a-z][A-Z]") = False Then
        Else
            For a = 0 To 9
                If ucr.Text.ToString.Contains(Trim(Str(a))) = True Then
                    MessageBox.Show("US CAAT Reviewer name can only contain alphabets", "Invalid Entry!!")
                    ucr.Focus()
                    Exit Sub
                End If
            Next a
        End If
        If (TextBox6.Text = "") Then
            MessageBox.Show("Please enter the name of the US CAAT Manager", "Incomplete!!")
            TextBox6.Focus()
            Exit Sub
            'ElseIf Regex.IsMatch(TextBox2.Text, "[a-z][A-Z]") = False Then
        Else
            For a = 0 To 9
                If TextBox6.Text.ToString.Contains(Trim(Str(a))) = True Then
                    MessageBox.Show("US CAAT Manager name can only contain alphabets", "Invalid Entry!!")
                    TextBox6.Focus()
                    Exit Sub
                End If
            Next a
        End If
        If (ListBox1.SelectedIndex = -1) Then
            MessageBox.Show("Please enter the CAAT preparer's name", "Incomplete!!")
            ListBox1.Focus()
            Exit Sub
        End If
        If ListBox1.SelectedItem = "Surender U" Then TextBox3.Text = "Surender.U@in.ey.com"
        If ListBox1.SelectedItem = "Sharada Lc" Then TextBox3.Text = "Sharada.Lc@in.ey.com"
        If ListBox1.SelectedItem = "Sriram C Kaushik" Then TextBox3.Text = "Sriram.Kaushik@in.ey.com"
        If ListBox1.SelectedItem = "Priya R S" Then TextBox3.Text = "priya.rs@in.ey.com"
        'If ListBox1.SelectedItem = "Apurv Sharma" Then TextBox3.Text = "apurv.sharma@in.ey.com"
        'If ListBox1.SelectedItem = "Sajag Goel" Then TextBox3.Text = "sajag.goel@in.ey.com"
        'If ListBox1.SelectedItem = "Ishita Srivastava" Then TextBox3.Text = "ishita.srivastava@in.ey.com"
        'If ListBox1.SelectedItem = "Jasleen Kaur" Then TextBox3.Text = "jasleen.kaur@in.ey.com"

        If (ListBox2.SelectedIndex = -1) Then
            MessageBox.Show("Please enter the reviewer's name", "Incomplete!!")
            ListBox2.Focus()
            Exit Sub
        End If
        If ListBox2.SelectedItem = "Surender U" Then TextBox4.Text = "Surender.U@in.ey.com"
        If ListBox2.SelectedItem = "Sharada Lc" Then TextBox4.Text = "Sharada.Lc@in.ey.com"
        If ListBox2.SelectedItem = "Sriram C Kaushik" Then TextBox4.Text = "Sriram.Kaushik@in.ey.com"
        If ListBox2.SelectedItem = "Priya R S" Then TextBox4.Text = "priya.rs@in.ey.com"
        'If ListBox2.SelectedItem = "Suvidha Kaul" Then TextBox4.Text = "suvidha.kaul@in.ey.com"
        'If ListBox2.SelectedItem = "Smriti Jain" Then TextBox4.Text = "smriti3.jain@in.ey.com"
        'If ListBox2.SelectedItem = "Apurv Sharma" Then TextBox4.Text = "apurv.sharma@in.ey.com"
        'If ListBox2.SelectedItem = "Sajag Goel" Then TextBox4.Text = "sajag.goel@in.ey.com"
        'If ListBox2.SelectedItem = "Ishita Srivastava" Then TextBox4.Text = "ishita.srivastava@in.ey.com"
        'If ListBox2.SelectedItem = "Jasleen Kaur" Then TextBox4.Text = "jasleen.kaur@in.ey.com"

        Me.Hide()
        Form1.Focus()
    End Sub

    Private Sub other_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        Dim cb As CheckBox = CType(sender, CheckBox)

        If cb.Checked = True Then
            other.Visible = True
        Else
            other.Visible = False
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Dim str2 As String
        If CheckBox1.Checked = True Then
            str2 = InputBox("Enter the name of CAAT Preparer", "CAAT Preparer", "")
            ListBox1.Items.Add(str2)
            ListBox1.SelectedItem = str2
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        Dim str2 As String
        If CheckBox5.Checked = True Then
            str2 = InputBox("Enter the name of CAAT Reviewer", "CAAT Reviewer", "")
            ListBox2.Items.Add(str2)
            ListBox2.SelectedItem = str2
        End If
    End Sub
End Class