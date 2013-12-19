<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.COACheck = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.SourceDatafileBox = New System.Windows.Forms.CheckedListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ACLScriptBox = New System.Windows.Forms.CheckedListBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.JournalEntryBox = New System.Windows.Forms.CheckedListBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.eagleFile = New System.Windows.Forms.TextBox
        Me.ErepLabel = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.BRule = New System.Windows.Forms.ComboBox
        Me.NewLogo = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        CType(Me.NewLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'COACheck
        '
        Me.COACheck.AutoSize = True
        Me.COACheck.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.COACheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.COACheck.Location = New System.Drawing.Point(12, 161)
        Me.COACheck.Name = "COACheck"
        Me.COACheck.Size = New System.Drawing.Size(63, 17)
        Me.COACheck.TabIndex = 2
        Me.COACheck.Text = "COA file:"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.CheckBox1.Location = New System.Drawing.Point(168, 161)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(62, 17)
        Me.CheckBox1.TabIndex = 3
        Me.CheckBox1.Text = "Present"
        Me.CheckBox1.UseVisualStyleBackColor = False
        '
        'SourceDatafileBox
        '
        Me.SourceDatafileBox.FormattingEnabled = True
        Me.SourceDatafileBox.HorizontalScrollbar = True
        Me.SourceDatafileBox.Items.AddRange(New Object() {"Control totals were not provided by the client. Control totals and record counts " & _
                        "have been taken from the data imported into ACL.", "Control totals were provided by the client. Control totals and record counts have" & _
                        " been matched with the data imported into ACL.", "Manual formatting was performed on the trial balance and journal entry files prio" & _
                        "r to importing to ACL. Refer to Appendix B for complete details", "Input JE files were combined in DOS and a single file was imported in ACL.", "The record count mentioned for the JE file is of the combined file.", "Monarch was used to capture the Trial Balance files prior to importing in ACL. Re" & _
                        "fer to Appendix B for complete details.", "The beginning and the ending balances nets to $XX and $XX respectively . . Non-ze" & _
                        "ro balances were due to rounding of transactions to two decimal places.", "There was one more client provided file ""XXXX"" which was not utilized by EY as it" & _
                        " was not required for analysis.", "The TB used for the CAAT is the same as used for the audit at w/p ""X""", "The TB used for the CAAT was reconciled at an account type level to the TB used i" & _
                        "n the audit at w/p ""X"" "})
        Me.SourceDatafileBox.Location = New System.Drawing.Point(168, 190)
        Me.SourceDatafileBox.Name = "SourceDatafileBox"
        Me.SourceDatafileBox.Size = New System.Drawing.Size(397, 64)
        Me.SourceDatafileBox.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.Label1.Location = New System.Drawing.Point(7, 190)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(139, 15)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Source Data File Notes :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.Label2.Location = New System.Drawing.Point(245, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(160, 20)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Other Configurations:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.White
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label3.Location = New System.Drawing.Point(7, 262)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(123, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "ACL Script Notes :"
        '
        'ACLScriptBox
        '
        Me.ACLScriptBox.FormattingEnabled = True
        Me.ACLScriptBox.HorizontalScrollbar = True
        Me.ACLScriptBox.Items.AddRange(New Object() {resources.GetString("ACLScriptBox.Items"), resources.GetString("ACLScriptBox.Items1"), "EY noted that Business Unit was assigned according to the name of the journal ent" & _
                        "ry and trial balance files. Refer to A_JE_Prep and B_TB_Prep log for details.", resources.GetString("ACLScriptBox.Items2"), "Within A_JE_PREP and B_TB_PREP, EY excluded the line items with blank records whi" & _
                        "ch were not required for analysis.", resources.GetString("ACLScriptBox.Items3"), "GL account numbers beginning with ‘XXXX’, ‘XXXX’ were excluded from the journal e" & _
                        "ntry and trial balance file, as these were statistical accounts. Refer to A_JE_P" & _
                        "rep and B_TB_Prep log for details.", resources.GetString("ACLScriptBox.Items4"), resources.GetString("ACLScriptBox.Items5")})
        Me.ACLScriptBox.Location = New System.Drawing.Point(168, 262)
        Me.ACLScriptBox.Name = "ACLScriptBox"
        Me.ACLScriptBox.Size = New System.Drawing.Size(397, 64)
        Me.ACLScriptBox.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label4.Location = New System.Drawing.Point(612, 87)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(174, 17)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Journal Entry Files Notes :"
        '
        'JournalEntryBox
        '
        Me.JournalEntryBox.FormattingEnabled = True
        Me.JournalEntryBox.HorizontalScrollbar = True
        Me.JournalEntryBox.Items.AddRange(New Object() {"Field headers were not provided in the JE data; hence ACL automatically assigns f" & _
                        "ield headers as Field_1, Field_2, Field_3… in order of occurrences of the field." & _
                        "", "The field ""XXXX"" had the same value ""XXXX"" for all the line items. Hence not mapp" & _
                        "ed in GAT.", "The client provided fields ""XXXX"" and ""XXXX"" had no values for any of the line it" & _
                        "ems. Hence these fields were not mapped in the analysis.", "Fields ""XXXX"" and ""XXXX"" were not mapped in GAT, as all three user defined fields" & _
                        " were utilized and there was no way to capture this information in the GAT.", resources.GetString("JournalEntryBox.Items"), "EY had not mapped the fields Intercompany and Prof Fees in GAT as all the user de" & _
                        "fined fields were utilized.", resources.GetString("JournalEntryBox.Items1"), "EY noted that the effective date was always on or after the entry date, and appea" & _
                        "red to be the date the entry was posted and not necessarily the effective date o" & _
                        "f the entry.", resources.GetString("JournalEntryBox.Items2"), "There were a few line items for which the XXXX was exceeding the GAT length limit" & _
                        " of XXXX. These exceeded XXXX for a few line items was mapped in ACL as ""XXXX"". " & _
                        "This field was mapped in GAT as XXXX ", "EY noted that Source field had two values i.e. XXXX and XXXX which was not a stan" & _
                        "dard value as Source.", resources.GetString("JournalEntryBox.Items3"), "The client provided ""XXXX"" field was not utilized in the analysis as this is the " & _
                        "GL account name and available in the trial balance file. ", resources.GetString("JournalEntryBox.Items4")})
        Me.JournalEntryBox.Location = New System.Drawing.Point(612, 114)
        Me.JournalEntryBox.Name = "JournalEntryBox"
        Me.JournalEntryBox.Size = New System.Drawing.Size(397, 64)
        Me.JournalEntryBox.TabIndex = 12
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.White
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label5.Location = New System.Drawing.Point(612, 190)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(102, 17)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Buisness Rule:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.White
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label6.Location = New System.Drawing.Point(608, 331)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(114, 17)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Used EAGLe ? : "
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"Yes", "No"})
        Me.ComboBox2.Location = New System.Drawing.Point(611, 351)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox2.TabIndex = 16
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(925, 351)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(55, 21)
        Me.Button1.TabIndex = 17
        Me.Button1.Text = "Browse"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'eagleFile
        '
        Me.eagleFile.Location = New System.Drawing.Point(741, 352)
        Me.eagleFile.Name = "eagleFile"
        Me.eagleFile.Size = New System.Drawing.Size(178, 20)
        Me.eagleFile.TabIndex = 18
        Me.eagleFile.Visible = False
        '
        'ErepLabel
        '
        Me.ErepLabel.AutoSize = True
        Me.ErepLabel.BackColor = System.Drawing.Color.White
        Me.ErepLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.ErepLabel.Location = New System.Drawing.Point(738, 333)
        Me.ErepLabel.Name = "ErepLabel"
        Me.ErepLabel.Size = New System.Drawing.Size(181, 15)
        Me.ErepLabel.TabIndex = 19
        Me.ErepLabel.Text = "EAGLe Analysis Report (As .zip):"
        Me.ErepLabel.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label7.Location = New System.Drawing.Point(609, 259)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(126, 17)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "GAT ScreenShots:"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(741, 258)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(237, 20)
        Me.TextBox2.TabIndex = 22
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(978, 256)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(31, 23)
        Me.Button2.TabIndex = 21
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.White
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label8.Location = New System.Drawing.Point(609, 294)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(96, 17)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "File Structure:"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(741, 291)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(237, 20)
        Me.TextBox3.TabIndex = 25
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(978, 289)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(31, 23)
        Me.Button3.TabIndex = 24
        Me.Button3.Text = "..."
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(531, 394)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(134, 36)
        Me.Button4.TabIndex = 26
        Me.Button4.Text = "Done"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label9.Location = New System.Drawing.Point(7, 104)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(182, 42)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "Roll-Forward Note : " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Check if yes)"
        '
        'CheckBox2
        '
        Me.CheckBox2.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.CheckBox2.Location = New System.Drawing.Point(168, 99)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(438, 52)
        Me.CheckBox2.TabIndex = 28
        Me.CheckBox2.Text = "The initial client provided data had roll-forward differences. EY obtained update" & _
            "d files and was able to roll forward all the GL accounts to obtain reasonable as" & _
            "surance it is complete"
        Me.CheckBox2.UseVisualStyleBackColor = False
        '
        'BRule
        '
        Me.BRule.FormattingEnabled = True
        Me.BRule.Items.AddRange(New Object() {"Manual", "System"})
        Me.BRule.Location = New System.Drawing.Point(720, 189)
        Me.BRule.Name = "BRule"
        Me.BRule.Size = New System.Drawing.Size(121, 21)
        Me.BRule.TabIndex = 14
        '
        'NewLogo
        '
        Me.NewLogo.Image = Global.WindowsApplication1.My.Resources.Resources.logo
        Me.NewLogo.Location = New System.Drawing.Point(-4, -3)
        Me.NewLogo.Name = "NewLogo"
        Me.NewLogo.Size = New System.Drawing.Size(10, 10)
        Me.NewLogo.TabIndex = 29
        Me.NewLogo.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(-4, -3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1076, 462)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = Global.WindowsApplication1.My.Resources.Resources.EYL
        Me.PictureBox3.Location = New System.Drawing.Point(-4, -3)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(111, 94)
        Me.PictureBox3.TabIndex = 30
        Me.PictureBox3.TabStop = False
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1065, 441)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.NewLogo)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.ErepLabel)
        Me.Controls.Add(Me.eagleFile)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.BRule)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.JournalEntryBox)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ACLScriptBox)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SourceDatafileBox)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.COACheck)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "Form2"
        Me.Text = "Form2"
        CType(Me.NewLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents COACheck As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents SourceDatafileBox As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ACLScriptBox As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents JournalEntryBox As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents eagleFile As System.Windows.Forms.TextBox
    Friend WithEvents ErepLabel As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents BRule As System.Windows.Forms.ComboBox
    Friend WithEvents NewLogo As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
End Class
