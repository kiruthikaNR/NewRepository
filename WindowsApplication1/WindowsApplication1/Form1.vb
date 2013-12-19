Option Explicit Off

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


Public Class Form1

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'VALIDATIONS
        Dim a, START_INDEX1, END_INDEX1, LEN1 As Integer
        Dim start3 As String
        If (Form6.TextBox1.Text = "") Then
            MessageBox.Show("Please fill up the User Information form", "Incomplete entry")
            Form6.TextBox1.Focus()
            Exit Sub
        ElseIf (Form5.TextBox1.Text = "") Then
            MessageBox.Show("Please fill up the Entry of logs", "Incomplete entry")
            Form5.TextBox1.Focus()
            Exit Sub
        End If

        Dim STrA As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\A_JE_PREP.LOG")
        Dim STrB As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\B_TB_PREP.LOG")
        Dim STrC As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\C_WORKLOG.LOG")
        Dim STrD As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\D_JE_ROLL.LOG")
        'Dim STrA As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text.Substring(0, Len(Form5.TextBox1.Text) - 4) & ".LOG")
        'Dim STrB As IO.TextReader = System.IO.File.OpenText(Form5.TextBox2.Text.Substring(0, Len(Form5.TextBox2.Text) - 4) & ".LOG")
        'Dim STrC As IO.TextReader = System.IO.File.OpenText(Form5.TextBox3.Text.Substring(0, Len(Form5.TextBox3.Text) - 4) & ".LOG")
        'Dim STrD As IO.TextReader = System.IO.File.OpenText(Form5.TextBox4.Text.Substring(0, Len(Form5.TextBox4.Text) - 4) & ".LOG")
        Dim TRA As String = STrA.ReadToEnd
        Dim TRB As String = STrB.ReadToEnd
        Dim TRC As String = STrC.ReadToEnd
        Dim TRD As String = STrD.ReadToEnd
        'Dim MyFileLine1 As String = Split(TRD, vbCrLf)(12)


        temp1 = TRD.Substring(TRD.IndexOf("@ ASSIGN CLIENT_NAME ="))
        START_INDEX1 = TRD.IndexOf("@ ASSIGN CLIENT_NAME =") + 24
        END_INDEX1 = 0
        b = START_INDEX1 + 1
        Do While start3 <> Chr(34)
            start3 = TRD.Substring(b, 1)
            If start3 = Chr(34) Then
                END_INDEX1 = b
                Exit Do
            End If
            b = b + 1
        Loop
        LEN1 = END_INDEX1 - START_INDEX1
        temp = TRD.Substring(START_INDEX1, LEN1)
        Dim myclientname As String = temp

        temp1 = TRD.Substring(TRD.IndexOf("@ ASSIGN PERIOD      ="))
        START_INDEX1 = TRD.IndexOf("@ ASSIGN PERIOD      =") + 24
        END_INDEX1 = 0
        b = START_INDEX1 + 1
        start3 = ""
        Do While start3 <> Chr(34)
            start3 = TRD.Substring(b, 1)
            If start3 = Chr(34) Then
                END_INDEX1 = b
                Exit Do
            End If
            b = b + 1
        Loop
        LEN1 = END_INDEX1 - START_INDEX1
        temp = TRD.Substring(START_INDEX1, LEN1)
        'Dim MyFileLine2 As String = Split(TRD, vbCrLf)(14)
        Dim myPOA As String = temp
        'STrD.Close()
        'MsgBox(MyFileLine)
        START_POA = Trim(myPOA).Substring(0, myPOA.IndexOf(" "))
        temp = Trim(myPOA).Substring(Len(START_POA) + 1)
        end_poa = temp.Substring(temp.IndexOf(" ") + 1, Len(temp) - temp.IndexOf(" ") - 1)
        'NUMBER OF JE AND TB FILES

        Dim count_JE As Integer = Regex.Matches(TRA, "ACTIVATE").Count
        Dim count_TB As Integer = Regex.Matches(TRB, "ACTIVATE").Count
        Dim CLOSEOUT As Integer = Regex.Matches(TRB, "%RevExpTotal%").Count
        Dim countA_FIL As Integer
        Dim countB_FIL As Integer

        'Counting number of files except RAW - JE

        Dim splitfile() As String = TRA.Split(Chr(10))
        Dim astrlist As ArrayList = New ArrayList
        Dim check As Hashtable = New Hashtable
        Dim filename As String = ""
        For Each s In splitfile
            filename = Regex.Match(s, "\b\S*.fil\b |\b\S*.FIL\b").Value
            If ((Not filename.Contains("Raw")) And filename <> "") Then
                If Not (check(filename) = True) Then
                    check.Add(filename, True)
                    countA_FIL = countA_FIL + 1
                    astrlist.Add(filename)

                End If
            End If
        Next

        'Counting number of files except RAW - TB 

        Dim splitfileTB() As String = TRB.Split(Chr(10))
        Dim astrlist2 As ArrayList = New ArrayList
        Dim check2 As Hashtable = New Hashtable
        Dim filename2 As String = ""
        For Each s In splitfileTB
            filename2 = Regex.Match(s, "\b\S*.fil\b |\b\S*.FIL\b").Value
            If ((Not filename2.Contains("Raw")) And filename2 <> "") Then
                If Not (check2(filename2) = True) Then
                    check2.Add(filename2, True)
                    countB_FIL = countB_FIL + 1
                    astrlist2.Add(filename2)
                    'MsgBox(filename2)
                End If
            End If
        Next


        Dim LOG_PATH() As String = Form5.TextBox5.Text.Split("\")
        Dim U_BOUND As Integer = UBound(LOG_PATH)
        Dim str2 As String = ""
        For a3 = 0 To U_BOUND - 1
            str2 = str2 & LOG_PATH(a3) & "\"
        Next a3

        roll_name = Form5.TextBox5.Text.Substring(Len(str2), Len(Form5.TextBox5.Text) - Len(str2))

        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        'Dim oTable As Word.Table
        Dim oPara2 As Word.Paragraph
        Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph
        Dim oPara5 As Word.Paragraph

        'Start Word and open the document template.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add
        'Form6.eFrom.Format = DateTimePickerFormat.Custom
        'Form6.eFrom.CustomFormat = ""
        Dim effecPeriod As String = " from " & Form6.eFrom.Value.Date.ToString("MM/dd/yyyy") & " Through " & Form6.eTo.Value.Date.ToString("MM/dd/yyyy")


        'Insert a HEADER at the beginning of the document.
        Dim section As Microsoft.Office.Interop.Word.Section
        For Each section In oDoc.Sections
            section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Text = vbNewLine & vbNewLine & "EY Global Talent Hub JE CAAT" & Chr(10) & myclientname & effecPeriod
            section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Font.Bold = True
            section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Font.Name = "ARIAL"
            section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Font.Size = 10
            section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.InsertParagraphAfter()
            Clipboard.SetImage(My.Resources.EYL())
            section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Paragraphs(1).Range.Paste()
            section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Paragraphs(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        Next


        If (Form2.CheckBox2.Checked) Then
            Dim oParaTemp As Word.Paragraph = oDoc.Content.Paragraphs.Add
            oParaTemp.Range.Text = "Note : The initial client provided data had roll-forward differences. EY obtained updated files and was able to roll forward all the GL accounts to obtain reasonable assurance it is complete."
            oParaTemp.Range.Font.Name = "Times New Roman"
            oParaTemp.Range.Font.Bold = False
            oParaTemp.Format.SpaceAfter = 6
            oParaTemp.Range.Font.Size = 10
            oParaTemp.Range.Font.Italic = True
            oParaTemp.Range.Words(1).Font.Bold = True

            'oParaTemp.Range.Words(5).Font.ColorIndex = Word.WdColorIndex.wdDarkYellow
            oParaTemp.Range.InsertParagraphAfter()
        End If

        'Insert a paragraph at the beginning of the document.
        oPara2 = oDoc.Content.Paragraphs.Add
        oPara2.Range.Text = "Objective"
        oPara2.Range.Font.Name = "Times New Roman"
        oPara2.Range.Font.Bold = True
        oPara2.Range.Font.Underline = True
        oPara2.Format.SpaceAfter = 0
        oPara2.Range.Font.Size = 10
        oPara2.Range.InsertParagraphAfter()

        'date format
        Dim poa1 As String = Form6.eFrom.Value
        Dim poa1_1 As String() = poa1.Split("/")
        Dim poa1_2 As String() = poa1.Split(" ")

        If poa1_1(0) = "1" Then
            poa1_1_mon = "01"
            start_poa_mon = "January"
        ElseIf poa1_1(0) = "2" Then
            poa1_1_mon = "02"
            start_poa_mon = "February"
        ElseIf poa1_1(0) = "3" Then
            poa1_1_mon = "03"
            start_poa_mon = "March"
        ElseIf poa1_1(0) = "4" Then
            poa1_1_mon = "04"
            start_poa_mon = "April"
        ElseIf poa1_1(0) = "5" Then
            poa1_1_mon = "05"
            start_poa_mon = "May"
        ElseIf poa1_1(0) = "6" Then
            poa1_1_mon = "06"
            start_poa_mon = "June"
        ElseIf poa1_1(0) = "7" Then
            poa1_1_mon = "07"
            start_poa_mon = "July"
        ElseIf poa1_1(0) = "8" Then
            poa1_1_mon = "08"
            start_poa_mon = "August"
        ElseIf poa1_1(0) = "9" Then
            poa1_1_mon = "09"
            start_poa_mon = "September"
        ElseIf poa1_1(0) = "10" Then
            poa1_1_mon = "10"
            start_poa_mon = "October"
        ElseIf poa1_1(0) = "11" Then
            poa1_1_mon = "11"
            start_poa_mon = "November"
        ElseIf poa1_1(0) = "12" Then
            poa1_1_mon = "12"
            start_poa_mon = "December"
        End If

        If Len(poa1_1(1)) = 1 Then
            poa1_1_date = "0" & poa1_1(1)
        Else
            poa1_1_date = poa1_1(1)
        End If

        START_POA = poa1_1_mon & "/" & poa1_1_date & "/" & poa1_2(0).Substring(Len(poa1_1(0)) + Len(poa1_1(1)) + 2, 4)
        START_POA_word = start_poa_mon & " " & poa1_1_date & "," & poa1_2(0).Substring(Len(poa1_1(0)) + Len(poa1_1(1)) + 2, 4)

        Dim eoa1 As String = Form6.eTo.Value
        Dim eoa1_1 As String() = eoa1.Split("/")
        Dim eoa1_2 As String() = eoa1.Split(" ")

        If eoa1_1(0) = "1" Then
            eoa1_1_mon = "01"
            start_poa_mon = "January"
        ElseIf eoa1_1(0) = "2" Then
            eoa1_1_mon = "02"
            start_poa_mon = "February"
        ElseIf eoa1_1(0) = "3" Then
            eoa1_1_mon = "03"
            start_poa_mon = "March"
        ElseIf eoa1_1(0) = "4" Then
            eoa1_1_mon = "04"
            start_poa_mon = "April"
        ElseIf eoa1_1(0) = "5" Then
            eoa1_1_mon = "05"
            start_poa_mon = "May"
        ElseIf eoa1_1(0) = "6" Then
            eoa1_1_mon = "06"
            start_poa_mon = "June"
        ElseIf eoa1_1(0) = "7" Then
            eoa1_1_mon = "07"
            start_poa_mon = "July"
        ElseIf eoa1_1(0) = "8" Then
            eoa1_1_mon = "08"
            start_poa_mon = "August"
        ElseIf eoa1_1(0) = "9" Then
            eoa1_1_mon = "09"
            start_poa_mon = "September"
        ElseIf eoa1_1(0) = "10" Then
            eoa1_1_mon = "10"
            start_poa_mon = "October"
        ElseIf eoa1_1(0) = "11" Then
            eoa1_1_mon = "11"
            start_poa_mon = "November"
        ElseIf eoa1_1(0) = "12" Then
            eoa1_1_mon = "12"
            start_poa_mon = "December"
        End If

        If Len(eoa1_1(1)) = 1 Then
            eoa1_1_date = "0" & eoa1_1(1)
        Else
            eoa1_1_date = eoa1_1(1)
        End If

        'end_poa = temp.Substring(temp.IndexOf(" ") + 1, Len(temp) - temp.IndexOf(" ") - 1)
        end_poa = eoa1_1_mon & "/" & eoa1_1_date & "/" & eoa1_2(0).Substring(Len(eoa1_1(0)) + Len(eoa1_1(1)) + 2, 4)
        end_POA_word = start_poa_mon & " " & eoa1_1_date & "," & eoa1_2(0).Substring(Len(eoa1_1(0)) + Len(eoa1_1(1)) + 2, 4)

        'RECEIPT DATE

        Dim RCPT_DATE As String = Form6.DateTimePicker1.Value
        Dim RD1_1 As String() = RCPT_DATE.Split("/")
        Dim RD1_2 As String() = RCPT_DATE.Split(" ")

        If RD1_1(0) = "1" Then
            RD1_1_mon = "01"
            RD_mon = "January"
        ElseIf RD1_1(0) = "2" Then
            RD1_1_mon = "02"
            RD_mon = "February"
        ElseIf RD1_1(0) = "3" Then
            RD1_1_mon = "03"
            RD_mon = "March"
        ElseIf RD1_1(0) = "4" Then
            RD1_1_mon = "04"
            RD_mon = "April"
        ElseIf RD1_1(0) = "5" Then
            RD1_1_mon = "05"
            RD_mon = "May"
        ElseIf RD1_1(0) = "6" Then
            RD1_1_mon = "06"
            RD_mon = "June"
        ElseIf RD1_1(0) = "7" Then
            RD1_1_mon = "07"
            RD_mon = "July"
        ElseIf RD1_1(0) = "8" Then
            RD1_1_mon = "08"
            RD_mon = "August"
        ElseIf RD1_1(0) = "9" Then
            RD1_1_mon = "09"
            RD_mon = "September"
        ElseIf RD1_1(0) = "10" Then
            RD1_1_mon = "10"
            RD_mon = "October"
        ElseIf RD1_1(0) = "11" Then
            RD1_1_mon = "11"
            RD_mon = "November"
        ElseIf RD1_1(0) = "12" Then
            RD1_1_mon = "12"
            RD_mon = "December"
        End If

        If Len(RD1_1(1)) = 1 Then
            RD1_1_date = "0" & RD1_1(1)
        Else
            RD1_1_date = RD1_1(1)
        End If

        RD_word = RD_mon & " " & RD1_1_date & "," & RD1_2(0).Substring(Len(RD1_1(0)) + Len(RD1_1(1)) + 2, 4)



        'date format done


        '** \endofdoc is a predefined bookmark.

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Text = "To perform journal entry analysis for " & myclientname & " for the current period effective " & START_POA & "through" & end_poa
        oPara2.Format.SpaceAfter = 6
        oPara2.Range.Font.Name = "Times New Roman"
        oPara2.Range.Font.Size = 10
        oPara2.Range.Font.Bold = False
        oPara2.Range.Font.Underline = False
        oPara2.Format.SpaceAfter = 4
        oPara2.Range.InsertParagraphAfter()

        'FIRST TABLE OF THE MEMO

        
        'Changed - 1
        Dim TestDate As Date = Today()


        Dim otable1 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 10, 2)
        otable1.Borders.Enable = True
        otable1.Columns.Item(1).Width = oWord.CentimetersToPoints(5.27)
        otable1.Columns.Item(2).Width = oWord.CentimetersToPoints(12.49)
        otable1.Rows.Height = oWord.CentimetersToPoints(0.51)
        otable1.Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly

        otable1.Cell(1, 1).Range.Text = "Client / Engagement Name:"
        otable1.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(1, 1).Range.Font.Size = 10
        otable1.Cell(1, 1).Range.Bold = True
        otable1.Cell(1, 1).Range.Underline = False
        otable1.Cell(1, 1).Range.Italic = False
        otable1.Cell(1, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(1, 2).Range.Text = myclientname
        otable1.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(1, 2).Range.Font.Size = 10
        otable1.Cell(1, 2).Range.Bold = False
        otable1.Cell(1, 2).Range.Italic = False
        otable1.Cell(1, 2).Range.Underline = False

        otable1.Cell(2, 1).Range.Text = "Client / Engagement Code:"
        otable1.Cell(2, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(2, 1).Range.Font.Size = 10
        otable1.Cell(2, 1).Range.Bold = True
        otable1.Cell(2, 1).Range.Underline = False
        otable1.Cell(2, 1).Range.Italic = False
        otable1.Cell(2, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(2, 2).Range.Text = Form6.TextBox1.Text
        otable1.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(2, 2).Range.Font.Size = 10
        otable1.Cell(2, 2).Range.Bold = False
        otable1.Cell(2, 2).Range.Underline = False
        otable1.Cell(2, 2).Range.Italic = False


        otable1.Cell(3, 1).Range.Text = "DA Single Point Of Contact"
        otable1.Cell(3, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(3, 1).Range.Font.Size = 10
        otable1.Cell(3, 1).Range.Bold = True
        otable1.Cell(3, 1).Range.Underline = False
        otable1.Cell(3, 1).Range.Italic = False
        otable1.Cell(3, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(3, 2).Range.Text = Form6.spoc.Text
        otable1.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(3, 2).Range.Font.Size = 10
        otable1.Cell(3, 2).Range.Bold = False
        otable1.Cell(3, 2).Range.Underline = False
        otable1.Cell(3, 2).Range.Italic = False

        otable1.Cell(4, 1).Range.Text = "GTH CAAT Preparer:"
        otable1.Cell(4, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(4, 1).Range.Font.Size = 10
        otable1.Cell(4, 1).Range.Bold = True
        otable1.Cell(4, 1).Range.Underline = False
        otable1.Cell(4, 1).Range.Italic = False
        otable1.Cell(4, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(4, 2).Range.Text = Form6.ListBox1.Text
        otable1.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(4, 2).Range.Font.Size = 10
        otable1.Cell(4, 2).Range.Bold = False
        otable1.Cell(4, 2).Range.Underline = False
        otable1.Cell(4, 2).Range.Italic = False

        otable1.Cell(5, 1).Range.Text = "GTH CAAT Reviewer:"
        otable1.Cell(5, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(5, 1).Range.Font.Size = 10
        otable1.Cell(5, 1).Range.Bold = True
        otable1.Cell(5, 1).Range.Underline = False
        otable1.Cell(5, 1).Range.Italic = False
        otable1.Cell(5, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(5, 2).Range.Text = Form6.ListBox2.Text
        otable1.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(5, 2).Range.Font.Size = 10
        otable1.Cell(5, 2).Range.Bold = False
        otable1.Cell(5, 2).Range.Underline = False
        otable1.Cell(5, 2).Range.Italic = False

        otable1.Cell(6, 1).Range.Text = "US CAAT Reviewer:"
        otable1.Cell(6, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(6, 1).Range.Font.Size = 10
        otable1.Cell(6, 1).Range.Bold = True
        otable1.Cell(6, 1).Range.Underline = False
        otable1.Cell(6, 1).Range.Italic = False
        otable1.Cell(6, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(6, 2).Range.Text = Form6.ucr.Text
        otable1.Cell(6, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(6, 2).Range.Font.Size = 10
        otable1.Cell(6, 2).Range.Bold = False
        otable1.Cell(6, 2).Range.Underline = False
        otable1.Cell(6, 2).Range.Italic = False

        otable1.Cell(7, 1).Range.Text = "Data Receipt Date:"
        otable1.Cell(7, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(7, 1).Range.Font.Size = 10
        otable1.Cell(7, 1).Range.Bold = True
        otable1.Cell(7, 1).Range.Underline = False
        otable1.Cell(7, 1).Range.Italic = False
        otable1.Cell(7, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(7, 2).Range.Text = RD_word
        otable1.Cell(7, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(7, 2).Range.Font.Size = 10
        otable1.Cell(7, 2).Range.Bold = False
        otable1.Cell(7, 2).Range.Underline = False
        otable1.Cell(7, 2).Range.Italic = False

        otable1.Cell(8, 1).Range.Text = "Period of Analysis:"
        otable1.Cell(8, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(8, 1).Range.Font.Size = 10
        otable1.Cell(8, 1).Range.Bold = True
        otable1.Cell(8, 1).Range.Underline = False
        otable1.Cell(8, 1).Range.Italic = False
        otable1.Cell(8, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(8, 2).Range.Text = START_POA_word & " - " & end_POA_word
        otable1.Cell(8, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(8, 2).Range.Font.Size = 10
        otable1.Cell(8, 2).Range.Bold = False
        otable1.Cell(8, 2).Range.Underline = False
        otable1.Cell(8, 2).Range.Italic = False


        otable1.Cell(9, 1).Range.Text = "JE Module Delivery Date"
        otable1.Cell(9, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(9, 1).Range.Font.Size = 10
        otable1.Cell(9, 1).Range.Bold = True
        otable1.Cell(9, 1).Range.Underline = False
        otable1.Cell(9, 1).Range.Italic = False
        otable1.Cell(9, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        today_date = System.DateTime.Today.ToString("MM/dd/yyyy")
        today_date_month = Mid(today_date, 1, 2)
        today_date_day = Mid(today_date, 4, 2)
        today_date_year = Mid(today_date, 7, 4)

        If today_date_month = "1" Then
            today_date_month = "January"
        ElseIf today_date_month = "2" Then
            today_date_month = "February"
        ElseIf today_date_month = "3" Then
            today_date_month = "March"
        ElseIf today_date_month = "4" Then
            today_date_month = "April"
        ElseIf today_date_month = "5" Then
            today_date_month = "May"
        ElseIf today_date_month = "6" Then
            today_date_month = "June"
        ElseIf today_date_month = "7" Then
            today_date_month = "July"
        ElseIf today_date_month = "8" Then
            today_date_month = "August"
        ElseIf today_date_month = "9" Then
            today_date_month = "September"
        ElseIf today_date_month = "10" Then
            today_date_month = "October"
        ElseIf today_date_month = "11" Then
            today_date_month = "November"
        ElseIf today_date_month = "12" Then
            today_date_month = "December"
        End If

        If Len(today_date_day) = 1 Then today_date_day = "0" & today_date_day

        otable1.Cell(9, 2).Range.Text = today_date_month & " " & today_date_day & "," & today_date_year
        otable1.Cell(9, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(9, 2).Range.Font.Size = 10
        otable1.Cell(9, 2).Range.Bold = False
        otable1.Cell(9, 2).Range.Underline = False
        otable1.Cell(9, 2).Range.Italic = False
        otable1.Cell(9, 2).Range.Font.ColorIndex = Word.WdColorIndex.wdDarkRed


        otable1.Cell(10, 1).Range.Text = "Reviewer Sign-off Date"
        otable1.Cell(10, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(10, 1).Range.Font.Size = 10
        otable1.Cell(10, 1).Range.Bold = True
        otable1.Cell(10, 1).Range.Underline = False
        otable1.Cell(10, 1).Range.Italic = False
        otable1.Cell(10, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(10, 2).Range.Text = ""
        otable1.Cell(10, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(10, 2).Range.Font.Size = 10
        otable1.Cell(10, 2).Range.Bold = False
        otable1.Cell(10, 2).Range.Underline = False
        otable1.Cell(10, 2).Range.Italic = False

        oPara2.Format.SpaceAfter = 6
        oPara2.Range.InsertParagraphAfter()

        'Insert another paragraph.
        oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara3.Range.Text = "Method of Analysis: "
        oPara3.Range.Font.Bold = False
        oPara3.Format.SpaceAfter = 0
        oPara3.Range.Font.Name = "Times New Roman"
        oPara3.Range.Font.Bold = True
        oPara3.Range.Font.Underline = True
        oPara3.Range.Font.Italic = False
        oPara3.Range.Font.Size = 10
        oPara3.Range.InsertParagraphAfter()

        Dim otable3 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 8)
        otable3.Borders.Enable = True
        otable3.Columns.Item(1).Width = oWord.CentimetersToPoints(0.87)
        otable3.Columns.Item(3).Width = oWord.CentimetersToPoints(0.87)
        otable3.Columns.Item(5).Width = oWord.CentimetersToPoints(0.87)
        otable3.Columns.Item(7).Width = oWord.CentimetersToPoints(0.87)
        otable3.Columns.Item(2).Width = oWord.CentimetersToPoints(4.61)
        otable3.Columns.Item(4).Width = oWord.CentimetersToPoints(3.05)
        otable3.Columns.Item(6).Width = oWord.CentimetersToPoints(3.05)
        otable3.Columns.Item(8).Width = oWord.CentimetersToPoints(3.58)
        otable3.Rows.Height = oWord.CentimetersToPoints(0.4)

        otable3.Cell(1, 1).Range.Text = If(Form6.gl.Checked, ChrW(9746), ChrW(9744))
        otable3.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable3.Cell(1, 1).Range.Font.Size = 12
        otable3.Cell(1, 1).Range.Bold = False
        otable3.Cell(1, 1).Range.Underline = False
        otable3.Cell(1, 1).Range.Italic = False
        otable3.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable3.Cell(1, 3).Range.Text = If(Form6.CheckBox2.Checked, ChrW(9746), ChrW(9744))
        otable3.Cell(1, 3).Range.Font.Name = "Times New Roman"
        otable3.Cell(1, 3).Range.Font.Size = 12
        otable3.Cell(1, 3).Range.Bold = False
        otable3.Cell(1, 3).Range.Underline = False
        otable3.Cell(1, 3).Range.Italic = False
        otable3.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable3.Cell(1, 5).Range.Text = If(Form6.CheckBox3.Checked, ChrW(9746), ChrW(9744))
        otable3.Cell(1, 5).Range.Font.Name = "Times New Roman"
        otable3.Cell(1, 5).Range.Font.Size = 12
        otable3.Cell(1, 5).Range.Bold = False
        otable3.Cell(1, 5).Range.Underline = False
        otable3.Cell(1, 5).Range.Italic = False
        otable3.Cell(1, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable3.Cell(1, 7).Range.Text = If(Form6.CheckBox4.Checked, ChrW(9746), ChrW(9744))
        otable3.Cell(1, 7).Range.Font.Name = "Times New Roman"
        otable3.Cell(1, 7).Range.Font.Size = 12
        otable3.Cell(1, 7).Range.Bold = False
        otable3.Cell(1, 7).Range.Underline = False
        otable3.Cell(1, 7).Range.Italic = False
        otable3.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable3.Cell(1, 2).Range.Text = "EY/Global Analytics Module"
        otable3.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable3.Cell(1, 2).Range.Font.Size = 10
        otable3.Cell(1, 2).Range.Bold = False
        otable3.Cell(1, 2).Range.Underline = False
        otable3.Cell(1, 2).Range.Italic = False
        otable3.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable3.Cell(1, 4).Range.Text = "ACL"
        otable3.Cell(1, 4).Range.Font.Name = "Times New Roman"
        otable3.Cell(1, 4).Range.Font.Size = 10
        otable3.Cell(1, 4).Range.Bold = False
        otable3.Cell(1, 4).Range.Underline = False
        otable3.Cell(1, 4).Range.Italic = False
        otable3.Cell(1, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable3.Cell(1, 6).Range.Text = "MS Access"
        otable3.Cell(1, 6).Range.Font.Name = "Times New Roman"
        otable3.Cell(1, 6).Range.Font.Size = 10
        otable3.Cell(1, 6).Range.Bold = False
        otable3.Cell(1, 6).Range.Underline = False
        otable3.Cell(1, 6).Range.Italic = False
        otable3.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable3.Cell(1, 8).Range.Text = "Other:" & If(Form6.CheckBox4.Checked, Form6.other.Text, " ")
        otable3.Cell(1, 8).Range.Font.Name = "Times New Roman"
        otable3.Cell(1, 8).Range.Font.Size = 10
        otable3.Cell(1, 8).Range.Bold = False
        otable3.Cell(1, 8).Range.Underline = False
        otable3.Cell(1, 8).Range.Italic = False
        otable3.Cell(1, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable3.Rows.Height = oWord.CentimetersToPoints(0.4)

        oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara3.Range.Text = "NOTE:  Agreed upon per discussion with Financial Audit team."
        oPara3.Range.Font.Bold = False
        oPara3.Format.SpaceAfter = 6
        oPara3.Range.Font.Name = "Times New Roman"
        oPara3.Range.Font.Bold = False
        oPara3.Range.Font.Underline = False
        oPara3.Range.Font.Size = 10
        oPara3.Range.Font.Italic = True
        oPara3.Range.InsertParagraphAfter()
        oPara3.Range.Words(1).Font.Bold = True

        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara4.Range.Text = "Source Data Files:"
        oPara4.Range.Font.Bold = False
        oPara4.Format.SpaceAfter = 0
        oPara4.Range.Font.Name = "Times New Roman"
        oPara4.Range.Font.Bold = True
        oPara4.Range.Font.Underline = True
        oPara4.Range.Font.Italic = False
        oPara4.Range.Font.Size = 10
        oPara4.Range.InsertParagraphAfter()


        Dim otable2 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, count_JE + count_TB + +countA_FIL + countB_FIL + 1, 5)
        otable2.Borders.Enable = True
        otable2.Columns.Item(1).Width = oWord.CentimetersToPoints(6.44)
        otable2.Columns.Item(2).Width = oWord.CentimetersToPoints(2)
        otable2.Columns.Item(3).Width = oWord.CentimetersToPoints(2.75)
        otable2.Columns.Item(4).Width = oWord.CentimetersToPoints(3.25)
        otable2.Columns.Item(5).Width = oWord.CentimetersToPoints(3.32)
        otable2.Rows.Item(1).Height = oWord.CentimetersToPoints(0.49)

        otable2.Cell(1, 1).Range.Text = "Data File Name"
        otable2.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable2.Cell(1, 1).Range.Font.Size = 10
        otable2.Cell(1, 1).Range.Bold = True
        otable2.Cell(1, 1).Range.Underline = False
        otable2.Cell(1, 1).Range.Italic = True
        otable2.Cell(1, 1).Shading.BackgroundPatternColor = RGB(224, 224, 224)

        otable2.Cell(1, 2).Range.Text = "Record Count"
        otable2.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable2.Cell(1, 2).Range.Font.Size = 10
        otable2.Cell(1, 2).Range.Bold = True
        otable2.Cell(1, 2).Range.Underline = False
        otable2.Cell(1, 2).Range.Italic = True
        otable2.Cell(1, 2).Shading.BackgroundPatternColor = RGB(224, 224, 224)

        otable2.Cell(1, 3).Range.Text = "Control Total"
        otable2.Cell(1, 3).Range.Font.Name = "Times New Roman"
        otable2.Cell(1, 3).Range.Font.Size = 10
        otable2.Cell(1, 3).Range.Bold = True
        otable2.Cell(1, 3).Range.Underline = False
        otable2.Cell(1, 3).Range.Italic = True
        otable2.Cell(1, 3).Shading.BackgroundPatternColor = RGB(224, 224, 224)

        otable2.Cell(1, 4).Range.Text = "Description"
        otable2.Cell(1, 4).Range.Font.Name = "Times New Roman"
        otable2.Cell(1, 4).Range.Font.Size = 10
        otable2.Cell(1, 4).Range.Bold = True
        otable2.Cell(1, 4).Range.Underline = False
        otable2.Cell(1, 4).Range.Italic = True
        otable2.Cell(1, 4).Shading.BackgroundPatternColor = RGB(224, 224, 224)

        otable2.Cell(1, 5).Range.Text = "ACL Table Name" & " (*.fil)"
        otable2.Cell(1, 5).Range.Font.Name = "Times New Roman"
        otable2.Cell(1, 5).Range.Font.Size = 10
        otable2.Cell(1, 5).Range.Bold = True
        otable2.Cell(1, 5).Range.Underline = False
        otable2.Cell(1, 5).Range.Italic = True
        otable2.Cell(1, 5).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        Dim C As Integer

        'EXTRACTING JE DATA FROM A LOG

        Dim i As Integer = TRA.IndexOf("Opening file name ")
        Dim i_JE1 As Integer = TRA.IndexOf("ACTIVATE")
        Do While (i_JE1 <> -1)

            'GETTING THE RAW ACL TABLE NAMES

            TRA_TEMP1 = TRA.Substring(i + 18)
            je_file_name = TRA_TEMP1.SUBSTRING(0, UCase(TRA_TEMP1).INDEXOF(".FIL"))

            'GETTING THE RECORD COUNT OF ACL TABLE
            start1 = UCase(TRA_TEMP1).INDEXOF(" TO ") + 4
            end1 = UCase(TRA_TEMP1).INDEXOF(UCase(" records produced")) + 17

            TRA_TEMP2 = TRA_TEMP1.SUBSTRING(start1, end1 - start1)
            If Regex.Matches(TRA_TEMP2, "met the test: ").Count = 0 Then
                ENDINDEX = UCase(TRA_TEMP1).INDEXOF(UCase(" records produced"))
            Else
                ENDINDEX = UCase(TRA_TEMP1).INDEXOF(UCase(" met the test: "))
            End If



            If ENDINDEX <> -1 Then
                STARTINDEX = 0
                b = ENDINDEX - 1
                Do While start3 <> " "
                    start3 = UCase(TRA_TEMP1).substring(b, 1)
                    If start3 = " " Then
                        STARTINDEX = b
                        start3 = ""
                        Exit Do
                    End If
                    b = b - 1
                Loop
                C = C + 1

                'TRA_TEMP = TRA_TEMP1.SUBSTRING(1, TRA_TEMP1.INDEXOF(" as supplied in the table layout"))
                'Dim endIndex As Integer = UCase(TRA_TEMP).IndexOf(".FIL")
                'If endIndex <> -1 Then
                '    startindex = 0
                '    b = endIndex - 1
                '    Do While start3 <> " "
                '        start3 = UCase(TRA_TEMP).substring(b, 4)
                '        If start3 = " " Then
                '            startindex = b
                '            start3 = ""
                '            Exit Do
                '        End If
                '        b = b - 1
                '    Loop

                '    C = C + 1
                Dim extraction_je_count As String = TRA_TEMP1.Substring(STARTINDEX, endIndex - STARTINDEX).Trim
                Dim str1 As String = ""
                str1 = String.Format("{0:0,0}", FormatNumber(CDbl(extraction_je_count), 0))

                otable2.Cell(2 + C - 1, 2).Range.Text = str1
                otable2.Cell(2 + C - 1 - 1, 2).Range.Font.Name = "Times New Roman"
                otable2.Cell(2 + C - 1, 2).Range.Font.Size = 10
                otable2.Cell(2 + C - 1, 2).Range.Bold = False
                otable2.Cell(2 + C - 1, 2).Range.Underline = False
                otable2.Cell(2 + C - 1, 2).Range.Italic = False

                otable2.Cell(2 + C - 1, 5).Range.Text = je_file_name
                otable2.Cell(2 + C - 1, 5).Range.Font.Name = "Times New Roman"
                otable2.Cell(2 + C - 1, 5).Range.Font.Size = 10
                otable2.Cell(2 + C - 1, 5).Range.Bold = False
                otable2.Cell(2 + C - 1, 5).Range.Underline = False
                otable2.Cell(2 + C - 1, 5).Range.Italic = False

                temp1 = TRD.Substring(TRD.IndexOf("The total of EY_Amount is:"))
                START_INDEX1 = TRD.IndexOf("The total of EY_Amount is:")
                END_INDEX1 = 0
                b = START_INDEX1 + 1
                Do While start3 <> "@"
                    start3 = TRD.Substring(b, 1)
                    If start3 = "@" Then
                        END_INDEX1 = b
                        start3 = ""
                        Exit Do
                    End If
                    b = b + 1
                Loop
                LEN1 = END_INDEX1 - START_INDEX1
                temp = TRD.Substring(START_INDEX1 + 27, LEN1 - 27)

                je_amount = "Amount: $" & FormatNumber(CDbl(temp), 2)

                otable2.Cell(2 + C - 1, 3).Range.Text = "Amount: $" & FormatNumber(CDbl(temp), 2)
                otable2.Cell(2 + C - 1, 3).Range.Font.Name = "Times New Roman"
                otable2.Cell(2 + C - 1, 3).Range.Font.Size = 10
                otable2.Cell(2 + C - 1, 3).Range.Bold = False
                otable2.Cell(2 + C - 1, 3).Range.Underline = False
                otable2.Cell(2 + C - 1, 3).Range.Italic = False

                otable2.Cell(2 + C - 1, 4).Range.Text = "JE Activity for " & myPOA
                otable2.Cell(2 + C - 1, 4).Range.Font.Name = "Times New Roman"
                otable2.Cell(2 + C - 1, 4).Range.Font.Size = 10
                otable2.Cell(2 + C - 1, 4).Range.Bold = False
                otable2.Cell(2 + C - 1, 4).Range.Underline = False
                otable2.Cell(2 + C - 1, 4).Range.Italic = False

                i = TRA.IndexOf("Opening file name ", i + 1)
                i_JE1 = TRA.IndexOf("ACTIVATE", i_JE1 + 1)
            Else
                Exit Do
            End If
        Loop

        If count_JE > 1 Then
            With otable2
                .Cell(2, 2).Merge(.Cell(2 + count_JE - 1, 2))
                .Cell(2, 3).Merge(.Cell(2 + count_JE - 1, 3))
                .Cell(2, 3).Range.Text = je_amount
                .Cell(2, 4).Merge(.Cell(2 + count_JE - 1, 4))
                .Cell(2, 5).Merge(.Cell(2 + count_JE - 1, 5))
                .Cell(2, 4).Range.Text = "JE Activity for " & myPOA
            End With
        End If

        'EXTRACTING TB DATA FROM B LOG
        Dim i_TB As Integer = TRB.IndexOf("Opening file name")
        Dim i_TB1 As Integer = TRB.IndexOf("ACTIVATE")
        Dim D As Integer = 0

        Do While (i_TB1 <> -1)


            'GETTING THE RAW ACL TABLE NAMES

            TRB_TEMP1 = TRB.Substring(i_TB + 18)
            TB_file_name = TRB_TEMP1.SUBSTRING(0, UCase(TRB_TEMP1).INDEXOF(".FIL"))

            'GETTING THE RECORD COUNT OF ACL TABLE
            start1 = UCase(TRB_TEMP1).INDEXOF(" TO ") + 4
            end1 = UCase(TRB_TEMP1).INDEXOF(UCase(" records produced")) + 17

            TRB_TEMP2 = TRB_TEMP1.SUBSTRING(start1, end1 - start1)
            If Regex.Matches(TRB_TEMP2, "met the test: ").Count = 0 Then
                ENDINDEX = UCase(TRB_TEMP1).INDEXOF(UCase(" records produced"))
            Else
                ENDINDEX = UCase(TRB_TEMP1).INDEXOF(UCase(" met the test: "))
            End If

            If ENDINDEX <> -1 Then
                STARTINDEX = 0
                b = ENDINDEX - 1
                Do While start3 <> " "
                    start3 = UCase(TRB_TEMP1).substring(b, 1)
                    If start3 = " " Then
                        STARTINDEX = b
                        start3 = ""
                        Exit Do
                    End If
                    b = b - 1
                Loop
                Try
                    C = C + 1

                    'startIndex = TRB_TEMP.IndexOf("of") + 3
                    'endIndex = TRB_TEMP.IndexOf("met the test:")
                    'If endIndex <> -1 Then
                    '    startindex = 0
                    '    b = endIndex - 1
                    '    Do While start3 <> "of"
                    '        start3 = TRB_TEMP.substring(b, 2)
                    '        If start3 = "of" Then
                    '            startindex = b
                    '            start3 = ""
                    '            Exit Do
                    '        End If
                    '        b = b - 1
                    '    Loop
                    '    Try
                    '        D = D + 1
                    Dim extraction_tb_count As String = TRB_TEMP1.Substring(STARTINDEX, endIndex - STARTINDEX).Trim

                    'Dim TB_file_name As String = TRB_TEMP.Substring(startIndex1, endIndex1 - startIndex1).Trim
                    Dim str1 = String.Format("{0:0,0}", FormatNumber(CDbl(extraction_tb_count), 0))
                    otable2.Cell(2 + C + D - 1, 2).Range.Text = str1
                    otable2.Cell(2 + C + D - 1, 2).Range.Font.Name = "Times New Roman"
                    otable2.Cell(2 + C + D - 1, 2).Range.Font.Size = 10
                    otable2.Cell(2 + C + D - 1, 2).Range.Bold = False
                    otable2.Cell(2 + C + D - 1, 2).Range.Underline = False
                    otable2.Cell(2 + C + D - 1, 2).Range.Italic = False

                    otable2.Cell(2 + C + D - 1, 5).Range.Text = TB_file_name
                    otable2.Cell(2 + C + D - 1, 5).Range.Font.Name = "Times New Roman"
                    otable2.Cell(2 + C + D - 1, 5).Range.Font.Size = 10
                    otable2.Cell(2 + C + D - 1, 5).Range.Bold = False
                    otable2.Cell(2 + C + D - 1, 5).Range.Underline = False
                    otable2.Cell(2 + C + D - 1, 5).Range.Italic = False

                    TRD_TEMP = TRD.Substring(TRD.IndexOf("@ TOTAL FIELDS EY_BegBal EY_EndBal"))
                    If UCase(TB_file_name).Contains("BEG") Then
                        temp1 = "Beginning Balance: $"
                        temp2 = TRD_TEMP.Substring(TRD_TEMP.IndexOf("The total of EY_BegBal is:  ") + 28, TRD_TEMP.IndexOf("The total of EY_EndBal is:  ") - 28 - TRD_TEMP.IndexOf("The total of EY_BegBal is:  "))
                        temp = temp1 & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2))
                        name1 = "Beginning trial balance as on " & Chr(10) & START_POA
                    ElseIf UCase(TB_file_name).Contains("END") Then
                        temp1 = "Ending Balance: $"
                        temp2 = TRD_TEMP.Substring(TRD_TEMP.IndexOf("The total of EY_EndBal is:  ") + 28, 5)
                        temp = temp1 & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2))
                        name1 = "Ending trial balance as on " & Chr(10) & end_poa
                    Else
                        temp1 = "Beginning Balance: $"
                        temp2 = TRD_TEMP.Substring(TRD_TEMP.IndexOf("The total of EY_BegBal is:  ") + 28, TRD_TEMP.IndexOf("The total of EY_EndBal is:  ") - 28 - TRD_TEMP.IndexOf("The total of EY_BegBal is:  "))
                        temp_1 = temp1 & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2))
                        temp3 = "Ending Balance: $"
                        temp4 = TRD_TEMP.Substring(TRD_TEMP.IndexOf("The total of EY_EndBal is:  ") + 28, 5)
                        temp_2 = temp3 & String.Format("{0:0,0}", FormatNumber(CDbl(temp4), 2))
                        temp = temp1 & vbCrLf & temp2
                        name1 = "Beginning trial balance as on " & Chr(10) & START_POA & " and " & Chr(10) & "Ending trial balance as on " & Chr(10) & end_poa
                    End If
                Catch ex As Exception
                    If (TypeOf Err.GetException() Is ArgumentOutOfRangeException) Then
                        temp1 = "Beginning Balance: $"
                        temp2 = TRD.Substring(TRD.IndexOf("The total of EY_BegBal is:") + 28, TRD.IndexOf("The total of EY_EndBal is:") - TRD.IndexOf("The total of EY_BegBal is:") - 28)
                        'temp2 = TRB_TEMP.Substring(TRB_TEMP.IndexOf("The total of EY_BegBal is:  ") + 28, TRB_TEMP.IndexOf("The total of EY_EndBal is:  ") - 28 - TRB_TEMP.IndexOf("The total of EY_BegBal is:  "))
                        temp_1 = temp1 & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2))
                        temp3 = "Ending Balance: $"
                        TEMP4 = TRD.Substring(TRD.IndexOf("The total of EY_EndBal is:") + 28, TRD.IndexOf("@ TOTAL FIELDS COUNT") - TRD.IndexOf("The total of EY_EndBal is:") - 28)
                        'temp4 = TRB_TEMP.Substring(TRB_TEMP.IndexOf("The total of EY_EndBal is:  ") + 28, 5)
                        temp_2 = temp3 & String.Format("{0:0,0}", FormatNumber(CDbl(TEMP4), 2))
                        temp = Replace(temp_1, Chr(10), " ") & vbCrLf & Replace(temp_2, Chr(10), " ")
                        name1 = "Beginning trial balance as on " & Chr(10) & START_POA & " and " & Chr(10) & "Ending trial balance as on " & Chr(10) & end_poa
                    End If
                End Try

                tb_amount = temp


                otable2.Cell(2 + C + D - 1, 3).Range.Text = temp
                otable2.Cell(2 + C + D - 1, 3).Range.Font.Name = "Times New Roman"
                otable2.Cell(2 + C + D - 1, 3).Range.Font.Size = 10
                otable2.Cell(2 + C + D - 1, 3).Range.Bold = False
                otable2.Cell(2 + C + D - 1, 3).Range.Underline = False
                otable2.Cell(2 + C + D - 1, 3).Range.Italic = False

                otable2.Cell(2 + C + D - 1, 4).Range.Text = name1
                otable2.Cell(2 + C + D - 1, 4).Range.Font.Name = "Times New Roman"
                otable2.Cell(2 + C + D - 1, 4).Range.Font.Size = 10
                otable2.Cell(2 + C + D - 1, 4).Range.Bold = False
                otable2.Cell(2 + C + D - 1, 4).Range.Underline = False
                otable2.Cell(2 + C + D - 1, 4).Range.Italic = False

                i_TB = TRB.IndexOf("Opening file name", i_TB + 1)
                i_TB1 = TRB.IndexOf("ACTIVATE", i_TB + 1)
            Else
                Exit Do
            End If
        Loop

        If count_TB > 1 Then
            With otable2
                .Cell(2 + count_JE, 2).Merge(.Cell(2 + count_TB + count_JE - 1, 2))
                .Cell(2 + count_JE, 3).Merge(.Cell(2 + count_TB + count_JE - 1, 3))
                .Cell(2 + count_JE, 1).Merge(.Cell(2 + count_TB + count_JE - 1, 1))
                temp1 = "Beginning Balance: $"
                temp2 = TRD.Substring(TRD.IndexOf("The total of EY_BegBal is:") + 28, TRD.IndexOf("The total of EY_EndBal is:") - TRD.IndexOf("The total of EY_BegBal is:") - 28)
                'temp2 = TRB_TEMP.Substring(TRB_TEMP.IndexOf("The total of EY_BegBal is:  ") + 28, TRB_TEMP.IndexOf("The total of EY_EndBal is:  ") - 28 - TRB_TEMP.IndexOf("The total of EY_BegBal is:  "))
                temp_1 = temp1 & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2))
                temp3 = "Ending Balance: $"
                TEMP4 = TRD.Substring(TRD.IndexOf("The total of EY_EndBal is:") + 28, TRD.IndexOf("@ TOTAL FIELDS COUNT") - TRD.IndexOf("The total of EY_EndBal is:") - 28)
                'temp4 = TRB_TEMP.Substring(TRB_TEMP.IndexOf("The total of EY_EndBal is:  ") + 28, 5)
                temp_2 = temp3 & String.Format("{0:0,0}", FormatNumber(CDbl(TEMP4), 2))
                temp = Replace(temp_1, Chr(10), " ") & vbCrLf & Replace(temp_2, Chr(10), " ")
                'name1 = "Beginning trial balance as on " & Chr(10) & START_POA & " and " & Chr(10) & "Ending trial balance as on " & Chr(10) & end_poa
                name1 = "Beginning trial balance as on " & START_POA & " and " & "Ending trial balance as on " & end_poa
                .Cell(2 + count_JE, 3).Range.Text = temp
                .Cell(2 + count_JE, 4).Merge(.Cell(2 + count_TB + count_JE - 1, 4))
                .Cell(2 + count_JE, 5).Merge(.Cell(2 + count_TB + count_JE - 1, 5))
                .Cell(2 + count_JE, 4).Range.Text = name1
            End With
        End If


        Dim countnow As Integer = C + D + 1
        For index As Integer = 1 To countA_FIL

            otable2.Cell(countnow, 1).Range.Text = " "
            otable2.Cell(countnow, 2).Range.Text = " "
            otable2.Cell(countnow, 3).Range.Text = " "
            otable2.Cell(countnow, 4).Range.Text = "<@#$$@> All JE Files"
            otable2.Cell(countnow, 4).Range.Words(1).Shading.BackgroundPatternColorIndex = Word.WdColorIndex.wdYellow
            otable2.Cell(countnow, 4).Range.Bold = False
            otable2.Cell(countnow, 4).Range.Underline = False
            otable2.Cell(countnow, 4).Range.Italic = False
            ' MsgBox(astrlist(index - 1))
            otable2.Cell(countnow, 5).Range.Text = astrlist(index - 1)
            otable2.Cell(countnow, 5).Range.Bold = False
            otable2.Cell(countnow, 5).Range.Underline = False
            otable2.Cell(countnow, 5).Range.Italic = False

            countnow = countnow + 1
        Next

        otable2.Cell(C + D + 1, 4).Merge(otable2.Cell(countnow - 1, 4))
        otable2.Cell(C + D + 1, 5).Merge(otable2.Cell(countnow - 1, 5))
        Dim indexafterje As Integer = countnow

        For index As Integer = 1 To countB_FIL


            otable2.Cell(countnow, 1).Range.Text = " "
            otable2.Cell(countnow, 2).Range.Text = " "
            otable2.Cell(countnow, 3).Range.Text = " "
            otable2.Cell(countnow, 4).Range.Text = "<@#$$@> All TB Files"
            otable2.Cell(countnow, 4).Range.Words(1).Shading.BackgroundPatternColorIndex = Word.WdColorIndex.wdYellow
            otable2.Cell(countnow, 4).Range.Bold = False
            otable2.Cell(countnow, 4).Range.Underline = False
            otable2.Cell(countnow, 4).Range.Italic = False
            otable2.Cell(countnow, 5).Range.Text = astrlist2(index - 1)
            otable2.Cell(countnow, 5).Range.Bold = False
            otable2.Cell(countnow, 5).Range.Underline = False
            otable2.Cell(countnow, 5).Range.Italic = False

            countnow = countnow + 1
        Next

        otable2.Cell(indexafterje, 4).Merge(otable2.Cell(countnow - 1, 4))
        otable2.Cell(indexafterje, 5).Merge(otable2.Cell(countnow - 1, 5))

        oPara7 = oDoc.Content.Paragraphs.Add
        oPara7.Range.Text = "Note(s):"
        oPara7.Format.SpaceAfter = 0
        oPara7.Range.Font.Name = "Times New Roman"
        oPara7.Range.Font.Bold = False
        oPara7.Range.Font.Underline = False
        oPara7.Range.Font.Size = 10
        oPara7.Range.Font.Italic = True
        oPara7.Range.InsertParagraphAfter()


        Dim ind As Integer = 1
        If (Form2.SourceDatafileBox.SelectedItems.Count <> 0) Then
            For Each Entry In Form2.SourceDatafileBox.CheckedItems
                oPara7 = oDoc.Content.Paragraphs.Add
                oPara7.Range.Text = "   " & ind.ToString() & ". " & Entry.ToString()
                oPara7.Format.SpaceAfter = 1.5
                oPara7.Range.Font.Name = "Times New Roman"
                oPara7.Range.Font.Bold = False
                oPara7.Range.Font.Underline = False
                oPara7.Range.Font.Size = 10
                oPara7.Range.Font.Italic = True
                oPara7.Range.InsertParagraphAfter()
                ind = ind + 1
            Next

        End If



        'oPara7 = oDoc.Content.Paragraphs.Add
        'oPara7.Range.Text = "1." & Chr(9) & "<@#$$@> No Control totals were provided by the client. Control totals and Record count has been taken from the data imported into ACL " & Chr(10) & "2. <@#$$@>" & Chr(10) & "3. <@#$$@>"
        'oPara7.Format.SpaceAfter = 10
        'oPara7.Range.Font.Name = "Times New Roman"
        'oPara7.Range.Font.Bold = False
        'oPara7.Range.Font.Underline = False
        'oPara7.Range.Font.Size = 10
        'oPara7.Range.Font.Italic = True
        'oPara7.Range.InsertParagraphAfter()

        oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara5.Range.Text = "ACL Scripts:"
        oPara5.Range.Font.Bold = False
        oPara5.Format.SpaceAfter = 0
        oPara5.Range.Font.Name = "Times New Roman"
        oPara5.Range.Font.Bold = True
        oPara5.Range.Font.Underline = True
        oPara5.Range.Font.Italic = False
        oPara5.Range.Font.Size = 10
        oPara5.Range.InsertParagraphAfter()



        Dim otable4 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 5)
        otable4.Borders.Enable = True

        otable4.Columns.Item(1).Width = oWord.CentimetersToPoints(0.6)
        otable4.Columns.Item(2).Width = oWord.CentimetersToPoints(2.95)
        otable4.Columns.Item(3).Width = oWord.CentimetersToPoints(6.63)
        otable4.Columns.Item(4).Width = oWord.CentimetersToPoints(4.39)
        otable4.Columns.Item(5).Width = oWord.CentimetersToPoints(3.18)
        otable4.Rows.Item(7).Height = oWord.CentimetersToPoints(0.4)

        With otable4
            .Cell(7, 1).Merge(.Cell(7, 5))
            .Cell(8, 1).Merge(.Cell(8, 5))
        End With
        otable4.Cell(1, 1).Range.Text = "#"
        otable4.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Font.Size = 10
        otable4.Cell(1, 1).Range.Bold = True
        otable4.Cell(1, 1).Range.Underline = False
        otable4.Cell(1, 1).Range.Italic = True
        otable4.Cell(1, 1).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        For a = 2 To 6
            otable4.Cell(a, 1).Range.Text = a - 1
            otable4.Cell(a, 1).Range.Font.Name = "Times New Roman"
            otable4.Cell(a, 1).Range.Font.Size = 10
            otable4.Cell(a, 1).Range.Bold = False
            otable4.Cell(a, 1).Range.Underline = False
            otable4.Cell(a, 1).Range.Italic = False
        Next
        otable4.Cell(1, 2).Range.Text = "Script"
        otable4.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 2).Range.Font.Size = 10
        otable4.Cell(1, 2).Range.Bold = True
        otable4.Cell(1, 2).Range.Underline = False
        otable4.Cell(1, 2).Range.Italic = True
        otable4.Cell(1, 2).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        otable4.Cell(2, 2).Range.Text = "A_JE_ PREP"
        otable4.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otable4.Cell(2, 2).Range.Font.Size = 10
        otable4.Cell(2, 2).Range.Bold = False
        otable4.Cell(2, 2).Range.Underline = False
        otable4.Cell(2, 2).Range.Italic = False

        otable4.Cell(3, 2).Range.Text = "B_TB_ PREP"
        otable4.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otable4.Cell(3, 2).Range.Font.Size = 10
        otable4.Cell(3, 2).Range.Bold = False
        otable4.Cell(3, 2).Range.Underline = False
        otable4.Cell(3, 2).Range.Italic = False

        otable4.Cell(4, 2).Range.Text = "C_WORKLOG "
        otable4.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otable4.Cell(4, 2).Range.Font.Size = 10
        otable4.Cell(4, 2).Range.Bold = False
        otable4.Cell(4, 2).Range.Underline = False
        otable4.Cell(4, 2).Range.Italic = False

        otable4.Cell(5, 2).Range.Text = "D_JE_ROLL"
        otable4.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otable4.Cell(5, 2).Range.Font.Size = 10
        otable4.Cell(5, 2).Range.Bold = False
        otable4.Cell(5, 2).Range.Underline = False
        otable4.Cell(5, 2).Range.Italic = False

        otable4.Cell(6, 2).Range.Text = "E_EXPORTS"
        otable4.Cell(6, 2).Range.Font.Name = "Times New Roman"
        otable4.Cell(6, 2).Range.Font.Size = 10
        otable4.Cell(6, 2).Range.Bold = False
        otable4.Cell(6, 2).Range.Underline = False
        otable4.Cell(6, 2).Range.Italic = False

        otable4.Cell(1, 3).Range.Text = "Purpose"
        otable4.Cell(1, 3).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 3).Range.Font.Size = 10
        otable4.Cell(1, 3).Range.Bold = True
        otable4.Cell(1, 3).Range.Underline = False
        otable4.Cell(1, 3).Range.Italic = True
        otable4.Cell(1, 3).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        otable4.Cell(2, 3).Range.Text = "Formats source JE data file and consolidates into one table and joins to COA to obtain Account Type and Account Class information or compresses the data." & Chr(10) & "NOTE: A manual import of the JE data files must be performed prior to running this script. "
        otable4.Cell(2, 3).Range.Font.Name = "Times New Roman"
        otable4.Cell(2, 3).Range.Font.Size = 10
        otable4.Cell(2, 3).Range.Bold = False
        otable4.Cell(2, 3).Range.Underline = False
        otable4.Cell(2, 3).Range.Italic = False

        otable4.Cell(3, 3).Range.Text = "Formats source TB data file and consolidates into one table and joins to COA to obtain Account Type and Account Class information and performs close out." & Chr(10) & "NOTE: A manual import of the TB data files must be performed prior to running this script."
        otable4.Cell(3, 3).Range.Font.Name = "Times New Roman"
        otable4.Cell(3, 3).Range.Font.Size = 10
        otable4.Cell(3, 3).Range.Bold = False
        otable4.Cell(3, 3).Range.Underline = False
        otable4.Cell(3, 3).Range.Italic = False

        otable4.Cell(4, 3).Range.Text = "Performs data quality checks on the journal entry data file.  "
        otable4.Cell(4, 3).Range.Font.Name = "Times New Roman"
        otable4.Cell(4, 3).Range.Font.Size = 10
        otable4.Cell(4, 3).Range.Bold = False
        otable4.Cell(4, 3).Range.Underline = False
        otable4.Cell(4, 3).Range.Italic = False

        otable4.Cell(5, 3).Range.Text = "Performs Trial Balance roll-forward to validate completeness and extracts to Excel format.  "
        otable4.Cell(5, 3).Range.Font.Name = "Times New Roman"
        otable4.Cell(5, 3).Range.Font.Size = 10
        otable4.Cell(5, 3).Range.Bold = False
        otable4.Cell(5, 3).Range.Underline = False
        otable4.Cell(5, 3).Range.Italic = False

        otable4.Cell(6, 3).Range.Text = "Exports EY_JE and EY_TB for importing into global analytics tool"
        otable4.Cell(6, 3).Range.Font.Name = "Times New Roman"
        otable4.Cell(6, 3).Range.Font.Size = 10
        otable4.Cell(6, 3).Range.Bold = False
        otable4.Cell(6, 3).Range.Underline = False
        otable4.Cell(6, 3).Range.Italic = False

        'UPDATING 4TH COLUMN

        otable4.Cell(1, 4).Range.Text = "Input Files Created" & Chr(32) & " (*.FIL)"
        otable4.Cell(1, 4).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 4).Range.Font.Size = 10
        otable4.Cell(1, 4).Range.Bold = True
        otable4.Cell(1, 4).Range.Underline = False
        otable4.Cell(1, 4).Range.Italic = True
        otable4.Cell(1, 4).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        otable4.Cell(2, 4).Range.Text = "EY_JE"
        otable4.Cell(2, 4).Range.Font.Name = "Times New Roman"
        otable4.Cell(2, 4).Range.Font.Size = 10
        otable4.Cell(2, 4).Range.Bold = False
        otable4.Cell(2, 4).Range.Underline = False
        otable4.Cell(2, 4).Range.Italic = False

        otable4.Cell(3, 4).Range.Text = "EY_TB"
        otable4.Cell(3, 4).Range.Font.Name = "Times New Roman"
        otable4.Cell(3, 4).Range.Font.Size = 10
        otable4.Cell(3, 4).Range.Bold = False
        otable4.Cell(3, 4).Range.Underline = False
        otable4.Cell(3, 4).Range.Italic = False

        otable4.Cell(4, 4).Range.Text = "JE_SUMMARY"
        otable4.Cell(4, 4).Range.Font.Name = "Times New Roman"
        otable4.Cell(4, 4).Range.Font.Size = 10
        otable4.Cell(4, 4).Range.Bold = False
        otable4.Cell(4, 4).Range.Underline = False
        otable4.Cell(4, 4).Range.Italic = False

        otable4.Cell(5, 4).Range.Text = myclientname & " " & effecPeriod & " TB Rollforward.xls" & Chr(10) & "Unmatched_Roll_Trans.xls"
        otable4.Cell(5, 4).Range.Font.Name = "Times New Roman"
        otable4.Cell(5, 4).Range.Font.Size = 10
        otable4.Cell(5, 4).Range.Bold = False
        otable4.Cell(5, 4).Range.Underline = False
        otable4.Cell(5, 4).Range.Italic = False

        otable4.Cell(6, 4).Range.Text = "EY_JE.txt" & Chr(10) & "EY_TB.txt"
        otable4.Cell(6, 4).Range.Font.Name = "Times New Roman"
        otable4.Cell(6, 4).Range.Font.Size = 10
        otable4.Cell(6, 4).Range.Bold = False
        otable4.Cell(6, 4).Range.Underline = False
        otable4.Cell(6, 4).Range.Italic = False

        'UPDATING 5TH COLUMN

        otable4.Cell(1, 5).Range.Text = "ACL Logs" & Chr(32) & " (*.LOG)"
        otable4.Cell(1, 5).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 5).Range.Font.Size = 10
        otable4.Cell(1, 5).Range.Bold = True
        otable4.Cell(1, 5).Range.Underline = False
        otable4.Cell(1, 5).Range.Italic = True
        otable4.Cell(1, 5).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        'Dim filename As Object = Form5.TextBox1.Text


        'otable4.Cell(2, 5).Range.InsertFile(filename, Attachment:=True)
        a_zip = Form5.TextBox1.Text & "\A_JE_PREP.zip"
        b_zip = Form5.TextBox1.Text & "\B_TB_PREP.zip"
        c_zip = Form5.TextBox1.Text & "\C_WORKLOG.zip"
        d_zip = Form5.TextBox1.Text & "\D_JE_ROLL.zip"
        e_zip = Form5.TextBox1.Text & "\E_EXPORTS.zip"

        otable4.Cell(2, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=a_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="A_JE_PREP")
        otable4.Cell(3, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=b_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="B_JE_PREP")
        otable4.Cell(4, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=c_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="C_WORKLOG")
        otable4.Cell(5, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=d_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="D_JE_ROLL")
        otable4.Cell(6, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=e_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="E_EXPORTS")

        otable4.Cell(2, 5).Range.Font.Underline = False
        otable4.Cell(3, 5).Range.Font.Underline = False
        otable4.Cell(4, 5).Range.Font.Underline = False
        otable4.Cell(5, 5).Range.Font.Underline = False
        otable4.Cell(6, 5).Range.Font.Underline = False

        otable4.Cell(7, 1).Range.Text = "ACL File:"
        otable4.Cell(7, 1).Range.Font.Name = "Times New Roman"
        otable4.Cell(7, 1).Range.Font.Size = 10
        otable4.Cell(7, 1).Range.Bold = True
        otable4.Cell(7, 1).Range.Underline = False
        otable4.Cell(7, 1).Range.Italic = True
        otable4.Cell(7, 1).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        ACL_NAME = Form5.TextBox6.Text.Substring(0, Len(Form5.TextBox6.Text) - 4) & ".zip"

        otable4.Cell(8, 1).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=ACL_NAME, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:=myclientname & " " & myPOA & " ACL ")
        otable4.Cell(8, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable4.Cell(8, 1).Range.Underline = False

        oPara7 = oDoc.Content.Paragraphs.Add
        oPara7.Range.Text = "Note(s)"
        oPara7.Format.SpaceAfter = 0
        oPara7.Range.Font.Name = "Times New Roman"
        oPara7.Range.Font.Bold = False
        oPara7.Range.Font.Underline = False
        oPara7.Range.Font.Size = 10
        oPara7.Range.Font.Italic = True
        oPara7.Range.InsertParagraphAfter()


        ind = 1
        If (Form2.ACLScriptBox.SelectedItems.Count <> 0) Then
            For Each Entry In Form2.ACLScriptBox.CheckedItems
                oPara7 = oDoc.Content.Paragraphs.Add
                oPara7.Range.Text = "   " & ind.ToString() & ". " & Entry.ToString()
                oPara7.Format.SpaceAfter = 1.5
                oPara7.Range.Font.Name = "Times New Roman"
                oPara7.Range.Font.Bold = False
                oPara7.Range.Font.Underline = False
                oPara7.Range.Font.Size = 10
                oPara7.Range.Font.Italic = True
                oPara7.Range.InsertParagraphAfter()
                ind = ind + 1
            Next

        End If

        oPara7 = oDoc.Content.Paragraphs.Add
        oPara7.Range.Text = ""
        oPara7.Format.SpaceAfter = 1.5
        oPara7.Range.Font.Name = "Times New Roman"
        oPara7.Range.Font.Bold = True
        oPara7.Range.Font.Underline = True
        oPara7.Range.Font.Size = 1
        oPara7.Range.Font.Italic = False
        oPara7.Range.InsertParagraphAfter()

        oPara8 = oDoc.Content.Paragraphs.Add
        oPara8.Range.Text = ""
        oPara8.Format.SpaceAfter = 0
        oPara8.Range.Font.Name = "Times New Roman"
        oPara8.Range.Font.Bold = True
        oPara8.Range.Font.Underline = True
        oPara8.Range.Font.Size = 1
        oPara8.Range.Font.Italic = False
        oPara8.Range.InsertParagraphAfter()

        'MAPPING OF JE SOURCE FILES

        Dim aTemp As String = STrA.ReadToEnd
        Dim udCount As Integer = Regex.Matches(aTemp, "UDF").Count()
        Dim otable6 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 15 + udCount, 3)
        otable6.Borders.Enable = True

        otable6.Columns.Item(1).Width = oWord.CentimetersToPoints(7.01)
        otable6.Columns.Item(2).Width = oWord.CentimetersToPoints(4.29)
        otable6.Columns.Item(3).Width = oWord.CentimetersToPoints(6.46)
        otable6.Rows.Height = oWord.CentimetersToPoints(0.4)

        otable6.Cell(1, 1).Merge(otable6.Cell(1, 3))

        otable6.Cell(1, 1).Range.Text = "JE Transaction Files"
        otable6.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(1, 1).Range.Font.Size = 10
        otable6.Cell(1, 1).Range.Bold = True
        otable6.Cell(1, 1).Range.Underline = False
        otable6.Cell(1, 1).Range.Italic = True
        otable6.Cell(1, 1).Shading.BackgroundPatternColor = RGB(0, 0, 0)
        otable6.Cell(1, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        otable6.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0

        otable6.Cell(2, 1).Range.Text = "EY/Global Analytics Field Name"
        otable6.Cell(2, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(2, 1).Range.Font.Size = 10
        otable6.Cell(2, 1).Range.Bold = True
        otable6.Cell(2, 1).Range.Underline = False
        otable6.Cell(2, 1).Range.Italic = True
        otable6.Cell(2, 1).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        otable6.Cell(2, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

        otable6.Cell(2, 2).Range.Text = "ACL Field Name"
        otable6.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(2, 2).Range.Font.Size = 10
        otable6.Cell(2, 2).Range.Bold = True
        otable6.Cell(2, 2).Range.Underline = False
        otable6.Cell(2, 2).Range.Italic = True
        otable6.Cell(2, 2).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        otable6.Cell(2, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

        otable6.Cell(2, 3).Range.Text = "Client Data Field Name"
        otable6.Cell(2, 3).Range.Font.Name = "Times New Roman"
        otable6.Cell(2, 3).Range.Font.Size = 10
        otable6.Cell(2, 3).Range.Bold = True
        otable6.Cell(2, 3).Range.Underline = False
        otable6.Cell(2, 3).Range.Italic = True
        otable6.Cell(2, 3).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        otable6.Cell(2, 3).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

        otable6.Cell(3, 1).Range.Text = "Journal Entry Number"
        otable6.Cell(3, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(3, 1).Range.Font.Size = 10
        otable6.Cell(3, 1).Range.Bold = True
        otable6.Cell(3, 1).Range.Underline = False
        otable6.Cell(3, 1).Range.Italic = False

        otable6.Cell(3, 2).Range.Text = "EY_JENum"
        otable6.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(3, 2).Range.Font.Size = 10
        otable6.Cell(3, 2).Range.Bold = False
        otable6.Cell(3, 2).Range.Underline = False
        otable6.Cell(3, 2).Range.Italic = False

        Try
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_JENum       COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)

            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+") - 1) & "<@#$$@>"
            End If
        Catch ex As Exception
            temp = "<@#$$@>"
        End Try

        Dim test As String = temp

        If (test.Contains("<@#$$@>")) Then
            otable6.Cell(3, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(3, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(3, 3).Range.Font.Size = 10
            otable6.Cell(3, 3).Range.Bold = False
            otable6.Cell(3, 3).Range.Underline = False
            otable6.Cell(3, 3).Range.Italic = False
            otable6.Cell(3, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        Else
            otable6.Cell(3, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(3, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(3, 3).Range.Font.Size = 10
            otable6.Cell(3, 3).Range.Bold = False
            otable6.Cell(3, 3).Range.Underline = False
            otable6.Cell(3, 3).Range.Italic = False
        End If



        Unique_je = Replace(temp, Chr(10), "")

        otable6.Cell(4, 1).Range.Text = "General Ledger Account Number"
        otable6.Cell(4, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(4, 1).Range.Font.Size = 10
        otable6.Cell(4, 1).Range.Bold = True
        otable6.Cell(4, 1).Range.Underline = False
        otable6.Cell(4, 1).Range.Italic = False

        otable6.Cell(4, 2).Range.Text = "EY_Acct"
        otable6.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(4, 2).Range.Font.Size = 10
        otable6.Cell(4, 2).Range.Bold = False
        otable6.Cell(4, 2).Range.Underline = False
        otable6.Cell(4, 2).Range.Italic = False
        Try
            a1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Acct        COMPUTED")))
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Acct        COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+")) & "<@#$$@>"
            End If
        Catch ex As Exception
            temp = "<@#$$@>"
        End Try
        test = temp
        If (test.Contains("<@#$$@>")) Then
            otable6.Cell(4, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(4, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(4, 3).Range.Font.Size = 10
            otable6.Cell(4, 3).Range.Bold = False
            otable6.Cell(4, 3).Range.Underline = False
            otable6.Cell(4, 3).Range.Italic = False
            otable6.Cell(4, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        Else
            otable6.Cell(4, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(4, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(4, 3).Range.Font.Size = 10
            otable6.Cell(4, 3).Range.Bold = False
            otable6.Cell(4, 3).Range.Underline = False
            otable6.Cell(4, 3).Range.Italic = False

        End If


        otable6.Cell(5, 1).Range.Text = "Amount"
        otable6.Cell(5, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(5, 1).Range.Font.Size = 10
        otable6.Cell(5, 1).Range.Bold = True
        otable6.Cell(5, 1).Range.Underline = False
        otable6.Cell(5, 1).Range.Italic = False

        otable6.Cell(5, 2).Range.Text = "EY_Amount"
        otable6.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(5, 2).Range.Font.Size = 10
        otable6.Cell(5, 2).Range.Bold = False
        otable6.Cell(5, 2).Range.Underline = False
        otable6.Cell(5, 2).Range.Italic = False

        Try
            a1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Amount      COMPUTED")))
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Amount      COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+")) & "<@#$$@>"
            End If
        Catch ex As Exception
            temp = "<@#$$@>"
        End Try
        test = temp

        If (test.Contains("<@#$$@>")) Then
            otable6.Cell(5, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(5, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(5, 3).Range.Font.Size = 10
            otable6.Cell(5, 3).Range.Bold = False
            otable6.Cell(5, 3).Range.Underline = False
            otable6.Cell(5, 3).Range.Italic = False
            otable6.Cell(5, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        Else
            otable6.Cell(5, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(5, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(5, 3).Range.Font.Size = 10
            otable6.Cell(5, 3).Range.Bold = False
            otable6.Cell(5, 3).Range.Underline = False
            otable6.Cell(5, 3).Range.Italic = False

        End If


        otable6.Cell(6, 1).Range.Text = "Business Unit"
        otable6.Cell(6, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(6, 1).Range.Font.Size = 10
        otable6.Cell(6, 1).Range.Bold = True
        otable6.Cell(6, 1).Range.Underline = False
        otable6.Cell(6, 1).Range.Italic = False

        otable6.Cell(6, 2).Range.Text = "EY_BusUnit"
        otable6.Cell(6, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(6, 2).Range.Font.Size = 10
        otable6.Cell(6, 2).Range.Bold = False
        otable6.Cell(6, 2).Range.Underline = False
        otable6.Cell(6, 2).Range.Italic = False

        Try
            a1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_BusUnit     COMPUTED")))
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_BusUnit     COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+")) & "<@#$$@>"
            End If
        Catch ex As Exception
            temp = "<@#$$@>"
        End Try
        test = temp


        If (test.Contains("<@#$$@>")) Then

            otable6.Cell(6, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(6, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(6, 3).Range.Font.Size = 10
            otable6.Cell(6, 3).Range.Bold = False
            otable6.Cell(6, 3).Range.Underline = False
            otable6.Cell(6, 3).Range.Italic = False
            otable6.Cell(6, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        Else

            otable6.Cell(6, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(6, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(6, 3).Range.Font.Size = 10
            otable6.Cell(6, 3).Range.Bold = False
            otable6.Cell(6, 3).Range.Underline = False
            otable6.Cell(6, 3).Range.Italic = False
        End If



        otable6.Cell(7, 1).Range.Text = "Effective Date"
        otable6.Cell(7, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(7, 1).Range.Font.Size = 10
        otable6.Cell(7, 1).Range.Bold = True
        otable6.Cell(7, 1).Range.Underline = False
        otable6.Cell(7, 1).Range.Italic = False

        otable6.Cell(7, 2).Range.Text = "EY_EffectiveDt"
        otable6.Cell(7, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(7, 2).Range.Font.Size = 10
        otable6.Cell(7, 2).Range.Bold = False
        otable6.Cell(7, 2).Range.Underline = False
        otable6.Cell(7, 2).Range.Italic = False

        Try
            a1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_EffectiveDt COMPUTED")))
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_EffectiveDt COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+")) & "<@#$$@>"
            End If

        Catch ex As Exception
            temp = "<@#$$@>"
        End Try

        test = temp

        If (test.Contains("<@#$$@>")) Then
            otable6.Cell(7, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(7, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(7, 3).Range.Font.Size = 10
            otable6.Cell(7, 3).Range.Bold = False
            otable6.Cell(7, 3).Range.Underline = False
            otable6.Cell(7, 3).Range.Italic = False
            otable6.Cell(7, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        Else
            otable6.Cell(7, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(7, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(7, 3).Range.Font.Size = 10
            otable6.Cell(7, 3).Range.Bold = False
            otable6.Cell(7, 3).Range.Underline = False
            otable6.Cell(7, 3).Range.Italic = False

        End If



        otable6.Cell(8, 1).Range.Text = "Entry Date"
        otable6.Cell(8, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(8, 1).Range.Font.Size = 10
        otable6.Cell(8, 1).Range.Bold = True
        otable6.Cell(8, 1).Range.Underline = False
        otable6.Cell(8, 1).Range.Italic = False

        otable6.Cell(8, 2).Range.Text = "EY_EntryDt"
        otable6.Cell(8, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(8, 2).Range.Font.Size = 10
        otable6.Cell(8, 2).Range.Bold = False
        otable6.Cell(8, 2).Range.Underline = False
        otable6.Cell(8, 2).Range.Italic = False

        Try
            a1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_EntryDt     COMPUTED")))
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_EntryDt     COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+")) & "<@#$$@>"
            End If

        Catch ex As Exception
            temp = "<@#$$@>"
        End Try

        test = temp

        If (test.Contains("<@#$$@>")) Then
            otable6.Cell(8, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(8, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(8, 3).Range.Font.Size = 10
            otable6.Cell(8, 3).Range.Bold = False
            otable6.Cell(8, 3).Range.Underline = False
            otable6.Cell(8, 3).Range.Italic = False
            otable6.Cell(8, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed

        Else
            otable6.Cell(8, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(8, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(8, 3).Range.Font.Size = 10
            otable6.Cell(8, 3).Range.Bold = False
            otable6.Cell(8, 3).Range.Underline = False
            otable6.Cell(8, 3).Range.Italic = False

        End If


        otable6.Cell(9, 1).Range.Text = "Period"
        otable6.Cell(9, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(9, 1).Range.Font.Size = 10
        otable6.Cell(9, 1).Range.Bold = True
        otable6.Cell(9, 1).Range.Underline = False
        otable6.Cell(9, 1).Range.Italic = False

        otable6.Cell(9, 2).Range.Text = "EY_Period"
        otable6.Cell(9, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(9, 2).Range.Font.Size = 10
        otable6.Cell(9, 2).Range.Bold = False
        otable6.Cell(9, 2).Range.Underline = False
        otable6.Cell(9, 2).Range.Italic = False

        Try
            a1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Period      COMPUTED")))
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Period      COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+") - 1) & "<@#$$@>"
            End If

        Catch ex As Exception
            temp = "<@#$$@>"
        End Try

        period_src = temp

        test = temp

        If (test.Contains("<@#$$@>")) Then
            otable6.Cell(9, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(9, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(9, 3).Range.Font.Size = 10
            otable6.Cell(9, 3).Range.Bold = False
            otable6.Cell(9, 3).Range.Underline = False
            otable6.Cell(9, 3).Range.Italic = False
            otable6.Cell(9, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        Else
            otable6.Cell(9, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(9, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(9, 3).Range.Font.Size = 10
            otable6.Cell(9, 3).Range.Bold = False
            otable6.Cell(9, 3).Range.Underline = False
            otable6.Cell(9, 3).Range.Italic = False
        End If




        otable6.Cell(10, 1).Range.Text = "Preparer ID"
        otable6.Cell(10, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(10, 1).Range.Font.Size = 10
        otable6.Cell(10, 1).Range.Bold = True
        otable6.Cell(10, 1).Range.Underline = False
        otable6.Cell(10, 1).Range.Italic = False

        otable6.Cell(10, 2).Range.Text = "EY_PreparerID"
        otable6.Cell(10, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(10, 2).Range.Font.Size = 10
        otable6.Cell(10, 2).Range.Bold = False
        otable6.Cell(10, 2).Range.Underline = False
        otable6.Cell(10, 2).Range.Italic = False

        Try
            a1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_PreparerID  COMPUTED")))
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_PreparerID  COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+")) & "<@#$$@>"
            End If

        Catch ex As Exception
            temp = "<@#$$@>"
        End Try
        test = temp

        If (test.Contains("<@#$$@>")) Then
            otable6.Cell(10, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(10, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(10, 3).Range.Font.Size = 10
            otable6.Cell(10, 3).Range.Bold = False
            otable6.Cell(10, 3).Range.Underline = False
            otable6.Cell(10, 3).Range.Italic = False
            otable6.Cell(10, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        Else
            otable6.Cell(10, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(10, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(10, 3).Range.Font.Size = 10
            otable6.Cell(10, 3).Range.Bold = False
            otable6.Cell(10, 3).Range.Underline = False
            otable6.Cell(10, 3).Range.Italic = False
        End If




        otable6.Cell(11, 1).Range.Text = "Source"
        otable6.Cell(11, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(11, 1).Range.Font.Size = 10
        otable6.Cell(11, 1).Range.Bold = True
        otable6.Cell(11, 1).Range.Underline = False
        otable6.Cell(11, 1).Range.Italic = False

        otable6.Cell(11, 2).Range.Text = "EY_Source"
        otable6.Cell(11, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(11, 2).Range.Font.Size = 10
        otable6.Cell(11, 2).Range.Bold = False
        otable6.Cell(11, 2).Range.Underline = False
        otable6.Cell(11, 2).Range.Italic = False

        Try
            a1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Source      COMPUTED")))
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Source      COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+")) & "<@#$$@>"
            End If

        Catch ex As Exception
            temp = "<@#$$@>"
        End Try
        test = temp

        If (test.Contains("<@#$$@>")) Then
            otable6.Cell(11, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(11, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(11, 3).Range.Font.Size = 10
            otable6.Cell(11, 3).Range.Bold = False
            otable6.Cell(11, 3).Range.Underline = False
            otable6.Cell(11, 3).Range.Italic = False
            otable6.Cell(11, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        Else
            otable6.Cell(11, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(11, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(11, 3).Range.Font.Size = 10
            otable6.Cell(11, 3).Range.Bold = False
            otable6.Cell(11, 3).Range.Underline = False
            otable6.Cell(11, 3).Range.Italic = False
        End If



        otable6.Cell(12, 1).Range.Text = "Journal Entry /Transaction Description"
        otable6.Cell(12, 1).Range.Font.Name = "Times New Roman"
        otable6.Cell(12, 1).Range.Font.Size = 10
        otable6.Cell(12, 1).Range.Bold = True
        otable6.Cell(12, 1).Range.Underline = False
        otable6.Cell(12, 1).Range.Italic = False

        otable6.Cell(12, 2).Range.Text = "EY_JE_Desc"
        otable6.Cell(12, 2).Range.Font.Name = "Times New Roman"
        otable6.Cell(12, 2).Range.Font.Size = 10
        otable6.Cell(12, 2).Range.Bold = False
        otable6.Cell(12, 2).Range.Underline = False
        otable6.Cell(12, 2).Range.Italic = False

        Try
            a1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Desc        COMPUTED")))
            TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_Desc        COMPUTED")) + 23)
            TRA_TEMP = TRA_TEMP1.Substring(TRA_TEMP1.indexof(Chr(10)) + 1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
            END_INDEX1 = TRA_TEMP.INDEXOF(",")
            START_INDEX1 = 0
            For a = END_INDEX1 To 1 Step -1
                start3 = TRA_TEMP.substring(a, 1)
                If start3 = "(" Then
                    START_INDEX1 = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX1 - START_INDEX1
            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
            If TRA_TEMP.indexof("+") <> -1 Then
                temp = temp & " " & TRA_TEMP.substring(TRA_TEMP.INDEXOF("+"), TRA_TEMP.INDEXOF(Chr(10)) - TRA_TEMP.INDEXOF("+")) & "<@#$$@>"
            End If
        Catch ex As Exception
            temp = "<@#$$@>"
        End Try
        test = temp

        If (test.Contains("<@#$$@>")) Then
            otable6.Cell(12, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(12, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(12, 3).Range.Font.Size = 10
            otable6.Cell(12, 3).Range.Bold = False
            otable6.Cell(12, 3).Range.Underline = False
            otable6.Cell(12, 3).Range.Italic = False
            otable6.Cell(12, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed

        Else
            otable6.Cell(12, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable6.Cell(12, 3).Range.Font.Name = "Times New Roman"
            otable6.Cell(12, 3).Range.Font.Size = 10
            otable6.Cell(12, 3).Range.Bold = False
            otable6.Cell(12, 3).Range.Underline = False
            otable6.Cell(12, 3).Range.Italic = False

        End If





        'Try
        '    TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_UserDefinedField1 COMPUTED")) + 23)
        '    TRA_TEMP = TRA_TEMP1.Substring(1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
        '    END_INDEX1 = TRA_TEMP.INDEXOF(",")
        '    START_INDEX1 = 0
        '    For a = END_INDEX1 To 1 Step -1
        '        start3 = TRA_TEMP.substring(a, 1)
        '        If start3 = "(" Then
        '            START_INDEX1 = a
        '            Exit For
        '        End If
        '    Next a
        '    LEN1 = END_INDEX1 - START_INDEX1
        '    If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
        'Catch ex As Exception
        '    temp = "<@#$$@>"
        'End Try



        'Try
        '    TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_UserDefinedField2 COMPUTED")) + 23)
        '    TRA_TEMP = TRA_TEMP1.Substring(1, TRA_TEMP1.IndexOf("COMPUTED") - 1)
        '    END_INDEX1 = TRA_TEMP.INDEXOF(",")
        '    START_INDEX1 = 0
        '    For a = END_INDEX1 To 1 Step -1
        '        start3 = TRA_TEMP.substring(a, 1)
        '        If start3 = "(" Then
        '            START_INDEX1 = a
        '            Exit For
        '        End If
        '    Next a
        '    LEN1 = END_INDEX1 - START_INDEX1
        '    If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
        'Catch ex As Exception
        '    temp = "<@#$$@>"
        'End Try



        'Try
        '    TRA_TEMP1 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_UserDefinedField3 COMPUTED")))
        '    TRA_TEMP = TRA_TEMP1.SUBSTRING(1, TRA_TEMP1.INDEXOF("@ EXTRACT"))
        '    END_INDEX1 = TRA_TEMP.INDEXOF(",")
        '    START_INDEX1 = 0
        '    For a = END_INDEX1 To 1 Step -1
        '        start3 = TRA_TEMP.substring(a, 1)
        '        If start3 = "(" Then
        '            START_INDEX1 = a
        '            Exit For
        '        End If
        '    Next a
        '    LEN1 = END_INDEX1 - START_INDEX1
        '    If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = TRA_TEMP.SUBSTRING(START_INDEX1, LEN1)
        'Catch ex As Exception
        '    temp = "<@#$$@>"
        'End Try


        If (udCount > 0) Then

            For index As Integer = 1 To 10
                otable6.Cell(13 + index, 1).Range.Text = "User Defined Field" & index.ToString()
                otable6.Cell(13 + index, 1).Range.Font.Name = "Times New Roman"
                otable6.Cell(13 + index, 1).Range.Font.Size = 10
                otable6.Cell(13 + index, 1).Range.Bold = True
                otable6.Cell(13 + index, 1).Range.Underline = False
                otable6.Cell(13 + index, 1).Range.Italic = False

                otable6.Cell(13 + index, 2).Range.Text = "<@#$$@>"
                otable6.Cell(13 + index, 2).Range.Font.Name = "Times New Roman"
                otable6.Cell(13 + index, 2).Range.Font.Size = 10
                otable6.Cell(13 + index, 2).Range.Bold = False
                otable6.Cell(13 + index, 2).Range.Underline = False
                otable6.Cell(13 + index, 2).Range.Italic = False
                otable6.Cell(13 + index, 2).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
            Next


        End If



        oPara9 = oDoc.Content.Paragraphs.Add
        oPara9.Range.Text = ""
        oPara9.Format.SpaceAfter = 10
        oPara9.Range.Font.Name = "Times New Roman"
        oPara9.Range.Font.Bold = True
        oPara9.Range.Font.Underline = True
        oPara9.Range.Font.Size = 10
        oPara9.Range.Font.Italic = False
        oPara9.Range.InsertParagraphAfter()


        oPara9 = oDoc.Content.Paragraphs.Add
        oPara9.Range.Text = "Note(s):"
        oPara9.Format.SpaceAfter = 0
        oPara9.Range.Font.Name = "Times New Roman"
        oPara9.Range.Font.Bold = False
        oPara9.Range.Font.Underline = False
        oPara9.Range.Font.Size = 10
        oPara9.Range.Font.Italic = True
        oPara9.Range.InsertParagraphAfter()

        ind = 1
        If (Form2.JournalEntryBox.SelectedItems.Count <> 0) Then
            For Each Entry In Form2.JournalEntryBox.CheckedItems
                oPara7 = oDoc.Content.Paragraphs.Add
                oPara7.Range.Text = "   " & ind.ToString() & ". " & Entry.ToString()
                oPara7.Format.SpaceAfter = 1.5
                oPara7.Range.Font.Name = "Times New Roman"
                oPara7.Range.Font.Bold = False
                oPara7.Range.Font.Underline = False
                oPara7.Range.Font.Size = 10
                oPara7.Range.Font.Italic = True
                oPara7.Range.InsertParagraphAfter()
                ind = ind + 1
            Next

        End If

        'TB FIELDS MAPPING

        temp_text = "Trial Balance Files"
        temp = TRB.Substring(TRB.IndexOf("EY_BegBal      COMPUTED") + 23)
        temp1 = temp.substring(temp.indexof(Chr(10)) + 1, temp.indexof("COMPUTED"))
        If Regex.Matches(temp1, "0.00").Count = 0 Then
            temp = TRB.Substring(TRB.IndexOf("EY_EndBal      COMPUTED") + 23)
            temp1 = temp.substring(temp.indexof(Chr(10)) + 1, temp.indexof("COMPUTED"))
            If Regex.Matches(temp1, "0.00").Count <> 0 Then FlaG = "y" Else flag = "n"
        Else : flag = "n"
        End If

        Dim tbCount As Integer = (Regex.Matches(STrB.ReadToEnd(), "@ ACTIVATE")).Count
        Dim colCount As Integer
        If (tbCount > 1) Then
            colCount = 4
        Else
            colCount = 3
        End If

        If FlaG = "n" Then

            Dim otable8 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            otable8.Borders.Enable = True
            otable8.AllowAutoFit = True
            otable8.Columns.Item(1).Width = oWord.CentimetersToPoints(6.6)
            otable8.Columns.Item(2).Width = oWord.CentimetersToPoints(3.81)
            otable8.Columns.Item(3).Width = oWord.CentimetersToPoints(7.41)
            otable8.Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly
            otable8.Rows.Height = oWord.CentimetersToPoints(0.6)

            otable8.Cell(1, 1).Merge(otable8.Cell(1, 3))

            otable8.Cell(1, 1).Range.Text = "Trial Balance Files"
            otable8.Cell(1, 1).Range.Font.Name = "Times New Roman"
            otable8.Cell(1, 1).Range.Font.Size = 10
            otable8.Cell(1, 1).Range.Bold = True
            otable8.Cell(1, 1).Range.Underline = False
            otable8.Cell(1, 1).Range.Italic = True
            otable8.Cell(1, 1).Shading.BackgroundPatternColor = RGB(0, 0, 0)
            otable8.Cell(1, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            otable8.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0

            otable8.Cell(2, 1).Range.Text = "EY/Global Analytics Field Name"
            otable8.Cell(2, 1).Range.Font.Name = "Times New Roman"
            otable8.Cell(2, 1).Range.Font.Size = 10
            otable8.Cell(2, 1).Range.Bold = True
            otable8.Cell(2, 1).Range.Underline = False
            otable8.Cell(2, 1).Range.Italic = True
            otable8.Cell(2, 1).Shading.BackgroundPatternColor = RGB(224, 224, 224)
            otable8.Cell(2, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

            otable8.Cell(2, 2).Range.Text = "ACL Field Name"
            otable8.Cell(2, 2).Range.Font.Name = "Times New Roman"
            otable8.Cell(2, 2).Range.Font.Size = 10
            otable8.Cell(2, 2).Range.Bold = True
            otable8.Cell(2, 2).Range.Underline = False
            otable8.Cell(2, 2).Range.Italic = True
            otable8.Cell(2, 2).Shading.BackgroundPatternColor = RGB(224, 224, 224)
            otable8.Cell(2, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

            otable8.Cell(2, 3).Range.Text = "Client Data Field Name"
            otable8.Cell(2, 3).Range.Font.Name = "Times New Roman"
            otable8.Cell(2, 3).Range.Font.Size = 10
            otable8.Cell(2, 3).Range.Bold = True
            otable8.Cell(2, 3).Range.Underline = False
            otable8.Cell(2, 3).Range.Italic = True
            otable8.Cell(2, 3).Shading.BackgroundPatternColor = RGB(224, 224, 224)
            otable8.Cell(2, 3).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

            otable8.Cell(3, 1).Range.Text = "General Ledger Account Number"
            otable8.Cell(3, 1).Range.Font.Name = "Times New Roman"
            otable8.Cell(3, 1).Range.Font.Size = 10
            otable8.Cell(3, 1).Range.Bold = True
            otable8.Cell(3, 1).Range.Underline = False
            otable8.Cell(3, 1).Range.Italic = False

            otable8.Cell(3, 2).Range.Text = "EY_Acct"
            otable8.Cell(3, 2).Range.Font.Name = "Times New Roman"
            otable8.Cell(3, 2).Range.Font.Size = 10
            otable8.Cell(3, 2).Range.Bold = False
            otable8.Cell(3, 2).Range.Underline = False
            otable8.Cell(3, 2).Range.Italic = False
            Try
                a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_Acct        COMPUTED")))
                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_Acct        COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.IndexOf(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try
            test = temp

            If (test.Contains("<@#$$@>")) Then
                otable8.Cell(3, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
                otable8.Cell(3, 3).Range.Font.Name = "Times New Roman"
                otable8.Cell(3, 3).Range.Font.Size = 10
                otable8.Cell(3, 3).Range.Bold = False
                otable8.Cell(3, 3).Range.Underline = False
                otable8.Cell(3, 3).Range.Italic = False
                otable8.Cell(3, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
            Else
                otable8.Cell(3, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
                otable8.Cell(3, 3).Range.Font.Name = "Times New Roman"
                otable8.Cell(3, 3).Range.Font.Size = 10
                otable8.Cell(3, 3).Range.Bold = False
                otable8.Cell(3, 3).Range.Underline = False
                otable8.Cell(3, 3).Range.Italic = False

            End If




            otable8.Cell(4, 1).Range.Text = "General Ledger Account Name"
            otable8.Cell(4, 1).Range.Font.Name = "Times New Roman"
            otable8.Cell(4, 1).Range.Font.Size = 10
            otable8.Cell(4, 1).Range.Bold = True
            otable8.Cell(4, 1).Range.Underline = False
            otable8.Cell(4, 1).Range.Italic = False

            otable8.Cell(4, 2).Range.Text = "EY_AcctName"
            otable8.Cell(4, 2).Range.Font.Name = "Times New Roman"
            otable8.Cell(4, 2).Range.Font.Size = 10
            otable8.Cell(4, 2).Range.Bold = False
            otable8.Cell(4, 2).Range.Underline = False
            otable8.Cell(4, 2).Range.Italic = False

            Try
                If UCase(TRB).IndexOf(UCase("EY_AcctDesc    COMPUTED")) <> -1 Then
                    a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctDesc    COMPUTED")))
                    trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctDesc    COMPUTED")) + 23)
                    trb_temp = trb_temp1.Substring(trb_temp1.indexof(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                    END_INDEX1 = trb_temp.INDEXOF(",")
                    START_INDEX1 = 0
                    For a = END_INDEX1 To 1 Step -1
                        start3 = trb_temp.substring(a, 1)
                        If start3 = "(" Then
                            START_INDEX1 = a
                            Exit For
                        End If
                    Next a
                    LEN1 = END_INDEX1 - START_INDEX1
                    If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                    If trb_temp.indexof("+") <> -1 Then
                        temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                    End If
                ElseIf UCase(TRB).IndexOf(UCase("EY_AcctName    COMPUTED")) <> -1 Then
                    a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctName    COMPUTED")))
                    trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctName    COMPUTED")) + 23)
                    trb_temp = trb_temp1.Substring(trb_temp1.indexof(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                    END_INDEX1 = trb_temp.INDEXOF(",")
                    START_INDEX1 = 0
                    For a = END_INDEX1 To 1 Step -1
                        start3 = trb_temp.substring(a, 1)
                        If start3 = "(" Then
                            START_INDEX1 = a
                            Exit For
                        End If
                    Next a
                    LEN1 = END_INDEX1 - START_INDEX1
                    If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                    If trb_temp.indexof("+") <> -1 Then
                        temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                    End If
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            test = temp

            If (test.Contains("<@#$$@>")) Then
                otable8.Cell(4, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
                otable8.Cell(4, 3).Range.Font.Name = "Times New Roman"
                otable8.Cell(4, 3).Range.Font.Size = 10
                otable8.Cell(4, 3).Range.Bold = False
                otable8.Cell(4, 3).Range.Underline = False
                otable8.Cell(4, 3).Range.Italic = False
                otable8.Cell(4, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
            Else
                otable8.Cell(4, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
                otable8.Cell(4, 3).Range.Font.Name = "Times New Roman"
                otable8.Cell(4, 3).Range.Font.Size = 10
                otable8.Cell(4, 3).Range.Bold = False
                otable8.Cell(4, 3).Range.Underline = False
                otable8.Cell(4, 3).Range.Italic = False

            End If


            otable8.Cell(5, 1).Range.Text = "Account Type"
            otable8.Cell(5, 1).Range.Font.Name = "Times New Roman"
            otable8.Cell(5, 1).Range.Font.Size = 10
            otable8.Cell(5, 1).Range.Bold = True
            otable8.Cell(5, 1).Range.Underline = False
            otable8.Cell(5, 1).Range.Italic = False

            otable8.Cell(5, 2).Range.Text = "EY_Accttype"
            otable8.Cell(5, 2).Range.Font.Name = "Times New Roman"
            otable8.Cell(5, 2).Range.Font.Size = 10
            otable8.Cell(5, 2).Range.Bold = False
            otable8.Cell(5, 2).Range.Underline = False
            otable8.Cell(5, 2).Range.Italic = False

            Try
                a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctType    COMPUTED")))
                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctType    COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.indexof(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            test = temp

            If (test.Contains("<@#$$@>")) Then
                otable8.Cell(5, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
                otable8.Cell(5, 3).Range.Font.Name = "Times New Roman"
                otable8.Cell(5, 3).Range.Font.Size = 10
                otable8.Cell(5, 3).Range.Bold = False
                otable8.Cell(5, 3).Range.Underline = False
                otable8.Cell(5, 3).Range.Italic = False
                otable8.Cell(5, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
            Else
                otable8.Cell(5, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
                otable8.Cell(5, 3).Range.Font.Name = "Times New Roman"
                otable8.Cell(5, 3).Range.Font.Size = 10
                otable8.Cell(5, 3).Range.Bold = False
                otable8.Cell(5, 3).Range.Underline = False
                otable8.Cell(5, 3).Range.Italic = False

            End If


            otable8.Cell(6, 1).Range.Text = "Account Class"
            otable8.Cell(6, 1).Range.Font.Name = "Times New Roman"
            otable8.Cell(6, 1).Range.Font.Size = 10
            otable8.Cell(6, 1).Range.Bold = True
            otable8.Cell(6, 1).Range.Underline = False
            otable8.Cell(6, 1).Range.Italic = False

            otable8.Cell(6, 2).Range.Text = "EY_AcctClass"
            otable8.Cell(6, 2).Range.Font.Name = "Times New Roman"
            otable8.Cell(6, 2).Range.Font.Size = 10
            otable8.Cell(6, 2).Range.Bold = False
            otable8.Cell(6, 2).Range.Underline = False
            otable8.Cell(6, 2).Range.Italic = False

            Try
                a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctClass   COMPUTED")))
                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctClass   COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.indexof(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable8.Cell(6, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable8.Cell(6, 3).Range.Font.Name = "Times New Roman"
            otable8.Cell(6, 3).Range.Font.Size = 10
            otable8.Cell(6, 3).Range.Bold = False
            otable8.Cell(6, 3).Range.Underline = False
            otable8.Cell(6, 3).Range.Italic = False

            otable8.Cell(7, 1).Range.Text = "Beginning Balance"
            otable8.Cell(7, 1).Range.Font.Name = "Times New Roman"
            otable8.Cell(7, 1).Range.Font.Size = 10
            otable8.Cell(7, 1).Range.Bold = True
            otable8.Cell(7, 1).Range.Underline = False
            otable8.Cell(7, 1).Range.Italic = False

            otable8.Cell(7, 2).Range.Text = "EY_BegBal"
            otable8.Cell(7, 2).Range.Font.Name = "Times New Roman"
            otable8.Cell(7, 2).Range.Font.Size = 10
            otable8.Cell(7, 2).Range.Bold = False
            otable8.Cell(7, 2).Range.Underline = False
            otable8.Cell(7, 2).Range.Italic = False

            Try
                a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_BegBal      COMPUTED")))
                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_BegBal      COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.indexof(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "0.00" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable8.Cell(7, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable8.Cell(7, 3).Range.Font.Name = "Times New Roman"
            otable8.Cell(7, 3).Range.Font.Size = 10
            otable8.Cell(7, 3).Range.Bold = False
            otable8.Cell(7, 3).Range.Underline = False
            otable8.Cell(7, 3).Range.Italic = False

            otable8.Cell(8, 1).Range.Text = "Ending Balance"
            otable8.Cell(8, 1).Range.Font.Name = "Times New Roman"
            otable8.Cell(8, 1).Range.Font.Size = 10
            otable8.Cell(8, 1).Range.Bold = True
            otable8.Cell(8, 1).Range.Underline = False
            otable8.Cell(8, 1).Range.Italic = False

            otable8.Cell(8, 2).Range.Text = "EY_EndBal"
            otable8.Cell(8, 2).Range.Font.Name = "Times New Roman"
            otable8.Cell(8, 2).Range.Font.Size = 10
            otable8.Cell(8, 2).Range.Bold = False
            otable8.Cell(8, 2).Range.Underline = False
            otable8.Cell(8, 2).Range.Italic = False

            Try
                a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_EndBal      COMPUTED")))
                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_EndBal      COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.indexof(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "0.00 <@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable8.Cell(8, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable8.Cell(8, 3).Range.Font.Name = "Times New Roman"
            otable8.Cell(8, 3).Range.Font.Size = 10
            otable8.Cell(8, 3).Range.Bold = False
            otable8.Cell(8, 3).Range.Underline = False
            otable8.Cell(8, 3).Range.Italic = False

            oPara10 = oDoc.Content.Paragraphs.Add
            oPara10.Range.Text = ""
            oPara10.Format.SpaceAfter = 0
            oPara10.Range.Font.Name = "Times New Roman"
            oPara10.Range.Font.Bold = True
            oPara10.Range.Font.Underline = True
            oPara10.Range.Font.Size = 10
            oPara10.Range.Font.Italic = False
            oPara10.Range.InsertParagraphAfter()

        End If
        oPara12 = oDoc.Content.Paragraphs.Add
        'oPara12.Range.Text = "Notes: 1.	Field headers were not provided in the JE data; hence ACL automatically assigns field headers as Field_1, Field_2, Field_3… in order of occurrences of the field.2.	The field “XXXX” had the same value “XXXX” for all the line items. Hence not mapped in GAT.3.	The client provided fields “XXXX” and “XXXX” had no values for any of the line items. Hence these fields were not mapped in the analysis.4.	Fields “XXXX” and “XXXX” were not mapped in GAT, as all three user defined fields were utilized and there was no way to capture this information in the GAT.5.	Apart from the journal entry description field ‘XXXX’ the client provided data also contained three other fields ‘XXXX’, ‘XXXX’ and ‘XXXX’ which were similar to the description field. These fields have been mapped as User Defined fields. The Assurance team should consider this field as well during their description word tests.6.	EY had not mapped the fields Intercompany and Prof Fees in GAT as all the user defined fields were utilized.7.	The user defined fields in GAT accepts only 50 characters. Since, the fields XXXX, XXXX and XXXX had length greater 50 characters, the rest of the characters were truncated in GAT. Only the first 50 characters were captured.8.	EY noted that the effective date was always on or after the entry date, and appeared to be the date the entry was posted and not necessarily the effective date of the entry.9.	There were XXXX line items for which the XXXX was exceeding the GAT length limit of 200.These line items are attached below for reference. These line items were mapped in ACL as “XXXX”. This field was not mapped in GAT as all the user defined fields in GAT were utilized.10.	There were a few line items for which the XXXX was exceeding the GAT length limit of XXXX. These exceeded XXXX for a few line items was mapped in ACL as “XXXX”. This field was mapped in GAT as XXXX 11.	EY noted that Source field had two values i.e. XXXX and XXXX which was not a standard value as Source.12.	The Inter-Company accounts and Prof-Fees accounts were identified as separate fields in the ACL for analysis. However due to restrict number of user defined fields available in GAT, these fields were not mapped as separate fields in the tool. The Inter-Company accounts and Prof-Fees accounts were identified in the report parameters in GAT directly from the General Ledger Account Numbers. Refer to Appendix A for the screenshot of parameters added in GAT13.	The client provided “XXXX” field was not utilized in the analysis as this is the GL account name and available in the trial balance file. 14.	EY noted that there were few other client provided fields “GL_BRANCH_NO”, “PAYMENT_NO”, “ARCADJ_NO” and “PAYABLE_NO” present in the data. However, all the fields could not be mapped as all the user defined fields were utilized."
        oPara12.Format.SpaceAfter = 0
        oPara12.Range.Font.Name = "Times New Roman"
        oPara12.Range.Font.Bold = True
        oPara12.Range.Font.Underline = True
        oPara12.Range.Font.Size = 10
        oPara12.Range.Font.Italic = False
        oPara12.Range.InsertParagraphAfter()



        Dim otable11 As Word.Table

        If FlaG = "y" Then

            otable10 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 9, 4)
            otable10.Borders.Enable = True
            otable10.AllowAutoFit = True
            otable10.Columns.Item(1).Width = oWord.CentimetersToPoints(5.5)
            otable10.Columns.Item(2).Width = oWord.CentimetersToPoints(3.25)
            otable10.Columns.Item(3).Width = oWord.CentimetersToPoints(4.5)
            otable10.Columns.Item(4).Width = oWord.CentimetersToPoints(4.5)
            otable10.Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly
            otable10.Rows.Height = oWord.CentimetersToPoints(0.55)

            With otable10
                .Cell(1, 1).Merge(otable10.Cell(1, 4))
                .Cell(2, 1).Merge(otable10.Cell(3, 1))
                .Cell(2, 2).Merge(otable10.Cell(3, 2))
                .Cell(2, 3).Merge(otable10.Cell(2, 4))
            End With


            otable10.Cell(1, 1).Range.Text = "Trial Balance Files "
            otable10.Cell(1, 1).Range.Font.Name = "Times New Roman"
            otable10.Cell(1, 1).Range.Font.Size = 10
            otable10.Cell(1, 1).Range.Bold = True
            otable10.Cell(1, 1).Range.Underline = False
            otable10.Cell(1, 1).Range.Italic = True
            otable10.Cell(1, 1).Shading.BackgroundPatternColor = RGB(0, 0, 0)
            otable10.Cell(1, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            otable10.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0

            otable10.Cell(2, 1).Range.Text = "EY/Global Analytics Field Name"
            otable10.Cell(2, 1).Range.Font.Name = "Times New Roman"
            otable10.Cell(2, 1).Range.Font.Size = 10
            otable10.Cell(2, 1).Range.Bold = True
            otable10.Cell(2, 1).Range.Underline = False
            otable10.Cell(2, 1).Range.Italic = True
            otable10.Cell(2, 1).Shading.BackgroundPatternColor = RGB(224, 224, 224)
            otable10.Cell(2, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

            otable10.Cell(2, 2).Range.Text = "ACL Field Name"
            otable10.Cell(2, 2).Range.Font.Name = "Times New Roman"
            otable10.Cell(2, 2).Range.Font.Size = 10
            otable10.Cell(2, 2).Range.Bold = True
            otable10.Cell(2, 2).Range.Underline = False
            otable10.Cell(2, 2).Range.Italic = True
            otable10.Cell(2, 2).Shading.BackgroundPatternColor = RGB(224, 224, 224)
            otable10.Cell(2, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

            otable10.Cell(2, 3).Range.Text = "Client Data Field Name"
            otable10.Cell(2, 3).Range.Font.Name = "Times New Roman"
            otable10.Cell(2, 3).Range.Font.Size = 10
            otable10.Cell(2, 3).Range.Bold = True
            otable10.Cell(2, 3).Range.Underline = False
            otable10.Cell(2, 3).Range.Italic = True
            otable10.Cell(2, 3).Shading.BackgroundPatternColor = RGB(224, 224, 224)
            otable10.Cell(2, 3).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

            otable10.Cell(3, 3).Range.Text = "Opening Trial Balance"
            otable10.Cell(3, 3).Range.Font.Name = "Times New Roman"
            otable10.Cell(3, 3).Range.Font.Size = 10
            otable10.Cell(3, 3).Range.Bold = True
            otable10.Cell(3, 3).Range.Underline = False
            otable10.Cell(3, 3).Range.Italic = True
            otable10.Cell(3, 3).Shading.BackgroundPatternColor = RGB(224, 224, 224)
            otable10.Cell(3, 3).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

            otable10.Cell(3, 4).Range.Text = "Closing Trial Balance"
            otable10.Cell(3, 4).Range.Font.Name = "Times New Roman"
            otable10.Cell(3, 4).Range.Font.Size = 10
            otable10.Cell(3, 4).Range.Bold = True
            otable10.Cell(3, 4).Range.Underline = False
            otable10.Cell(3, 4).Range.Italic = True
            otable10.Cell(3, 4).Shading.BackgroundPatternColor = RGB(224, 224, 224)
            otable10.Cell(3, 4).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

            otable10.Cell(4, 1).Range.Text = "General Ledger Account Number"
            otable10.Cell(4, 1).Range.Font.Name = "Times New Roman"
            otable10.Cell(4, 1).Range.Font.Size = 10
            otable10.Cell(4, 1).Range.Bold = True
            otable10.Cell(4, 1).Range.Underline = False
            otable10.Cell(4, 1).Range.Italic = False

            otable10.Cell(4, 2).Range.Text = "EY_Acct"
            otable10.Cell(4, 2).Range.Font.Name = "Times New Roman"
            otable10.Cell(4, 2).Range.Font.Size = 10
            otable10.Cell(4, 2).Range.Bold = False
            otable10.Cell(4, 2).Range.Underline = False
            otable10.Cell(4, 2).Range.Italic = False

            Try
                a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_Acct        COMPUTED")))
                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_Acct        COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.IndexOf(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(4, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(4, 3).Range.Font.Name = "Times New Roman"
            otable10.Cell(4, 3).Range.Font.Size = 10
            otable10.Cell(4, 3).Range.Bold = False
            otable10.Cell(4, 3).Range.Underline = False
            otable10.Cell(4, 3).Range.Italic = False

            trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("@ EXTRACT ")))

            Try
                a1 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_Acct        COMPUTED")))
                trb_temp2 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_Acct        COMPUTED")) + 23)
                trb_temp = trb_temp2.Substring(trb_temp2.indexof(Chr(10)) + 1, trb_temp2.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(4, 4).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(4, 4).Range.Font.Name = "Times New Roman"
            otable10.Cell(4, 4).Range.Font.Size = 10
            otable10.Cell(4, 4).Range.Bold = False
            otable10.Cell(4, 4).Range.Underline = False
            otable10.Cell(4, 4).Range.Italic = False

            otable10.Cell(5, 1).Range.Text = "General Ledger Account Name"
            otable10.Cell(5, 1).Range.Font.Name = "Times New Roman"
            otable10.Cell(5, 1).Range.Font.Size = 10
            otable10.Cell(5, 1).Range.Bold = True
            otable10.Cell(5, 1).Range.Underline = False
            otable10.Cell(5, 1).Range.Italic = False

            otable10.Cell(5, 2).Range.Text = "EY_AcctName"
            otable10.Cell(5, 2).Range.Font.Name = "Times New Roman"
            otable10.Cell(5, 2).Range.Font.Size = 10
            otable10.Cell(5, 2).Range.Bold = False
            otable10.Cell(5, 2).Range.Underline = False
            otable10.Cell(5, 2).Range.Italic = False

            Try
                If UCase(TRB).IndexOf(UCase("EY_AcctDesc    COMPUTED")) <> -1 Then
                    trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctDesc    COMPUTED")) + 23)
                    trb_temp = trb_temp1.Substring(trb_temp1.IndexOf(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                    END_INDEX1 = trb_temp.INDEXOF(",")
                    START_INDEX1 = 0
                    For a = END_INDEX1 To 1 Step -1
                        start3 = trb_temp.substring(a, 1)
                        If start3 = "(" Then
                            START_INDEX1 = a
                            Exit For
                        End If
                    Next a
                    LEN1 = END_INDEX1 - START_INDEX1
                    If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                    If trb_temp.indexof("+") <> -1 Then
                        temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                    End If
                ElseIf UCase(TRB).IndexOf(UCase("EY_AcctName    COMPUTED")) <> -1 Then
                    trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctName    COMPUTED")) + 23)
                    trb_temp = trb_temp1.Substring(trb_temp1.IndexOf(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                    END_INDEX1 = trb_temp.INDEXOF(",")
                    START_INDEX1 = 0
                    For a = END_INDEX1 To 1 Step -1
                        start3 = trb_temp.substring(a, 1)
                        If start3 = "(" Then
                            START_INDEX1 = a
                            Exit For
                        End If
                    Next a
                    LEN1 = END_INDEX1 - START_INDEX1
                    If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                    If trb_temp.indexof("+") <> -1 Then
                        temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                    End If
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(5, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(5, 3).Range.Font.Name = "Times New Roman"
            otable10.Cell(5, 3).Range.Font.Size = 10
            otable10.Cell(5, 3).Range.Bold = False
            otable10.Cell(5, 3).Range.Underline = False
            otable10.Cell(5, 3).Range.Italic = False

            trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("@ EXTRACT ")))

            Try
                If UCase(trb_temp1).IndexOf(UCase("EY_AcctDesc    COMPUTED")) <> -1 Then
                    trb_temp2 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_AcctDesc    COMPUTED")) + 23)
                    trb_temp = trb_temp2.Substring(trb_temp2.indexof(Chr(10)) + 1, trb_temp2.IndexOf("COMPUTED") - 1)
                    END_INDEX1 = trb_temp.INDEXOF(",")
                    START_INDEX1 = 0
                    For a = END_INDEX1 To 1 Step -1
                        start3 = trb_temp.substring(a, 1)
                        If start3 = "(" Then
                            START_INDEX1 = a
                            Exit For
                        End If
                    Next a
                    LEN1 = END_INDEX1 - START_INDEX1
                    If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                    If trb_temp.indexof("+") <> -1 Then
                        temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                    End If
                ElseIf UCase(trb_temp1).IndexOf(UCase("EY_AcctName    COMPUTED")) <> -1 Then
                    trb_temp2 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_AcctName    COMPUTED")) + 23)
                    trb_temp = trb_temp2.Substring(trb_temp2.indexof(Chr(10)) + 1, trb_temp2.IndexOf("COMPUTED") - 1)
                    END_INDEX1 = trb_temp.INDEXOF(",")
                    START_INDEX1 = 0
                    For a = END_INDEX1 To 1 Step -1
                        start3 = trb_temp.substring(a, 1)
                        If start3 = "(" Then
                            START_INDEX1 = a
                            Exit For
                        End If
                    Next a
                    LEN1 = END_INDEX1 - START_INDEX1
                    If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                    If trb_temp.indexof("+") <> -1 Then
                        temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                    End If
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(5, 4).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(5, 4).Range.Font.Name = "Times New Roman"
            otable10.Cell(5, 4).Range.Font.Size = 10
            otable10.Cell(5, 4).Range.Bold = False
            otable10.Cell(5, 4).Range.Underline = False
            otable10.Cell(5, 4).Range.Italic = False

            otable10.Cell(6, 1).Range.Text = "Account Type"
            otable10.Cell(6, 1).Range.Font.Name = "Times New Roman"
            otable10.Cell(6, 1).Range.Font.Size = 10
            otable10.Cell(6, 1).Range.Bold = True
            otable10.Cell(6, 1).Range.Underline = False
            otable10.Cell(6, 1).Range.Italic = False

            otable10.Cell(6, 2).Range.Text = "EY_Accttype"
            otable10.Cell(6, 2).Range.Font.Name = "Times New Roman"
            otable10.Cell(6, 2).Range.Font.Size = 10
            otable10.Cell(6, 2).Range.Bold = False
            otable10.Cell(6, 2).Range.Underline = False
            otable10.Cell(6, 2).Range.Italic = False

            Try
                If UCase(TRB).IndexOf(UCase("EY_AcctType    COMPUTED")) = -1 Then
                    temp = "<@#$$@>"
                    Exit Try
                End If

                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctType    COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.IndexOf(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(6, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(6, 3).Range.Font.Name = "Times New Roman"
            otable10.Cell(6, 3).Range.Font.Size = 10
            otable10.Cell(6, 3).Range.Bold = False
            otable10.Cell(6, 3).Range.Underline = False
            otable10.Cell(6, 3).Range.Italic = False

            trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("@ EXTRACT ")))

            Try
                If UCase(trb_temp1).IndexOf(UCase("EY_AcctType    COMPUTED")) = -1 Then
                    temp = "<@#$$@>"
                    Exit Try
                End If

                trb_temp2 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_AcctType    COMPUTED")) + 23)
                trb_temp = trb_temp2.Substring(trb_temp2.indexof(Chr(10)) + 1, trb_temp2.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(6, 4).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(6, 4).Range.Font.Name = "Times New Roman"
            otable10.Cell(6, 4).Range.Font.Size = 10
            otable10.Cell(6, 4).Range.Bold = False
            otable10.Cell(6, 4).Range.Underline = False
            otable10.Cell(6, 4).Range.Italic = False

            otable10.Cell(7, 1).Range.Text = "Account Class"
            otable10.Cell(7, 1).Range.Font.Name = "Times New Roman"
            otable10.Cell(7, 1).Range.Font.Size = 10
            otable10.Cell(7, 1).Range.Bold = True
            otable10.Cell(7, 1).Range.Underline = False
            otable10.Cell(7, 1).Range.Italic = False

            otable10.Cell(7, 2).Range.Text = "EY_AcctClass"
            otable10.Cell(7, 2).Range.Font.Name = "Times New Roman"
            otable10.Cell(7, 2).Range.Font.Size = 10
            otable10.Cell(7, 2).Range.Bold = False
            otable10.Cell(7, 2).Range.Underline = False
            otable10.Cell(7, 2).Range.Italic = False

            Try
                If UCase(TRB).IndexOf(UCase("EY_AcctClass   COMPUTED")) = -1 Then
                    temp = "<@#$$@>"
                    Exit Try
                End If

                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctClass   COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.IndexOf(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(7, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(7, 3).Range.Font.Name = "Times New Roman"
            otable10.Cell(7, 3).Range.Font.Size = 10
            otable10.Cell(7, 3).Range.Bold = False
            otable10.Cell(7, 3).Range.Underline = False
            otable10.Cell(7, 3).Range.Italic = False

            trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("@ EXTRACT ")))

            Try
                If UCase(trb_temp1).IndexOf(UCase("EY_AcctClass   COMPUTED")) = -1 Then
                    temp = "<@#$$@>"
                    Exit Try
                End If

                trb_temp2 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_AcctClass   COMPUTED")) + 23)
                trb_temp = trb_temp2.Substring(trb_temp2.indexof(Chr(10)) + 1, trb_temp2.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(7, 4).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(7, 4).Range.Font.Name = "Times New Roman"
            otable10.Cell(7, 4).Range.Font.Size = 10
            otable10.Cell(7, 4).Range.Bold = False
            otable10.Cell(7, 4).Range.Underline = False
            otable10.Cell(7, 4).Range.Italic = False

            otable10.Cell(8, 1).Range.Text = "Beginning Balance"
            otable10.Cell(8, 1).Range.Font.Name = "Times New Roman"
            otable10.Cell(8, 1).Range.Font.Size = 10
            otable10.Cell(8, 1).Range.Bold = True
            otable10.Cell(8, 1).Range.Underline = False
            otable10.Cell(8, 1).Range.Italic = False

            otable10.Cell(8, 2).Range.Text = "EY_BegBal"
            otable10.Cell(8, 2).Range.Font.Name = "Times New Roman"
            otable10.Cell(8, 2).Range.Font.Size = 10
            otable10.Cell(8, 2).Range.Bold = False
            otable10.Cell(8, 2).Range.Underline = False
            otable10.Cell(8, 2).Range.Italic = False

            Try
                a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_BegBal      COMPUTED")))
                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_BegBal      COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.IndexOf(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "0.00 <@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(8, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(8, 3).Range.Font.Name = "Times New Roman"
            otable10.Cell(8, 3).Range.Font.Size = 10
            otable10.Cell(8, 3).Range.Bold = False
            otable10.Cell(8, 3).Range.Underline = False
            otable10.Cell(8, 3).Range.Italic = False

            trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("@ EXTRACT ")))

            Try
                a1 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_BegBal      COMPUTED")))
                trb_temp2 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_BegBal      COMPUTED")) + 23)
                trb_temp = trb_temp2.Substring(trb_temp2.indexof(Chr(10)) + 1, trb_temp2.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "0.00 <@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(8, 4).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(8, 4).Range.Font.Name = "Times New Roman"
            otable10.Cell(8, 4).Range.Font.Size = 10
            otable10.Cell(8, 4).Range.Bold = False
            otable10.Cell(8, 4).Range.Underline = False
            otable10.Cell(8, 4).Range.Italic = False

            otable10.Cell(9, 1).Range.Text = "Ending Balance"
            otable10.Cell(9, 1).Range.Font.Name = "Times New Roman"
            otable10.Cell(9, 1).Range.Font.Size = 10
            otable10.Cell(9, 1).Range.Bold = True
            otable10.Cell(9, 1).Range.Underline = False
            otable10.Cell(9, 1).Range.Italic = False

            otable10.Cell(9, 2).Range.Text = "EY_EndBal"
            otable10.Cell(9, 2).Range.Font.Name = "Times New Roman"
            otable10.Cell(9, 2).Range.Font.Size = 10
            otable10.Cell(9, 2).Range.Bold = False
            otable10.Cell(9, 2).Range.Underline = False
            otable10.Cell(9, 2).Range.Italic = False

            Try
                a1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_EndBal      COMPUTED")))
                trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_EndBal      COMPUTED")) + 23)
                trb_temp = trb_temp1.Substring(trb_temp1.IndexOf(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "0.00 <@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(9, 3).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(9, 3).Range.Font.Name = "Times New Roman"
            otable10.Cell(9, 3).Range.Font.Size = 10
            otable10.Cell(9, 3).Range.Bold = False
            otable10.Cell(9, 3).Range.Underline = False
            otable10.Cell(9, 3).Range.Italic = False

            trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("@ EXTRACT ")))

            Try
                a1 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_EndBal      COMPUTED")))
                trb_temp2 = trb_temp1.Substring(UCase(trb_temp1).IndexOf(UCase("EY_EndBal      COMPUTED")) + 23)
                trb_temp = trb_temp2.Substring(trb_temp2.indexof(Chr(10)) + 1, trb_temp2.IndexOf("@ EXTRACT") - 1)
                END_INDEX1 = trb_temp.INDEXOF(",")
                START_INDEX1 = 0
                For a = END_INDEX1 To 1 Step -1
                    start3 = trb_temp.substring(a, 1)
                    If start3 = "(" Then
                        START_INDEX1 = a
                        Exit For
                    End If
                Next a
                LEN1 = END_INDEX1 - START_INDEX1
                If LEN1 <= 0 Then temp = "0.00 <@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
                If trb_temp.indexof("+") <> -1 Then
                    temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+")) & "<@#$$@>"
                End If
            Catch EX As Exception
                temp = "<@#$$@>"
            End Try

            otable10.Cell(9, 4).Range.Text = Replace(Replace(Replace(temp, "(", ""), ")", ""), ",", "")
            otable10.Cell(9, 4).Range.Font.Name = "Times New Roman"
            otable10.Cell(9, 4).Range.Font.Size = 10
            otable10.Cell(9, 4).Range.Bold = False
            otable10.Cell(9, 4).Range.Underline = False
            otable10.Cell(9, 4).Range.Italic = False

            oPara11 = oDoc.Content.Paragraphs.Add
            oPara11.Range.Text = ""
            oPara11.Format.SpaceAfter = 10
            oPara11.Range.Font.Name = "Times New Roman"
            oPara11.Range.Font.Bold = True
            oPara11.Range.Font.Underline = True
            oPara11.Range.Font.Size = 10
            oPara11.Range.Font.Italic = False
            oPara11.Range.InsertParagraphAfter()
        End If

        'BUSINESS RULES UPDATE

        oPara11 = oDoc.Content.Paragraphs.Add
        oPara11.Range.Text = "Business Rules:"
        oPara11.Format.SpaceAfter = 0
        oPara11.Range.Font.Name = "Times New Roman"
        oPara11.Range.Font.Bold = True
        oPara11.Range.Font.Underline = True
        oPara11.Range.Font.Size = 10
        oPara11.Range.Font.Italic = False
        oPara11.Range.InsertParagraphAfter()

        rng = oDoc.Bookmarks.Item("\endofdoc").Range
        otable10 = oDoc.Tables.Add(Range:=rng, NumRows:=9, NumColumns:=1)
        otable10.Borders.Enable = True
        otable10.AllowAutoFit = True
        otable10.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone
        otable10.Columns.Width = oWord.CentimetersToPoints(17.8)


        otable10.Cell(1, 1).Range.InsertParagraph()
        otable10.Cell(1, 1).Range.Paragraphs(1).Range.Text = ""
        otable10.Cell(1, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable10.Cell(1, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(1).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(1).Range.Bold = True
        otable10.Cell(1, 1).Range.Paragraphs(1).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(1).Range.Italic = False

        temp = TRA.Substring(UCase(TRA).IndexOf("TO EY_JE ") + 9)
        If UCase(temp.SUBSTRING(0, 2)) = "IF" Then
            temp1 = "Exclusions <@#$$@> " & temp.SUBSTRING(1, temp.INDEXOF(Chr(10)))
        Else
            temp1 = " No Exclusions <@#$$@>"
        End If

        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(2).Range.Text = "1.  " & "Identify and order journal entry fields to arrive at a unique journal entry"
        otable10.Cell(1, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
        otable10.Cell(1, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(2).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(2).Range.Bold = True
        otable10.Cell(1, 1).Range.Paragraphs(2).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(2).Range.Italic = False

        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(3).Range.Text = "     " & "• " & "The field " & Unique_je & " was identified as the unique journal entry identifier. <@#$$@>" & " . " & temp1
        otable10.Cell(1, 1).Range.Paragraphs(3).Format.SpaceAfter = 10
        otable10.Cell(1, 1).Range.Paragraphs(3).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(3).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(3).Range.Bold = False
        otable10.Cell(1, 1).Range.Paragraphs(3).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(3).Range.Italic = False

        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(4).Range.Text = "2.  " & " System/Manual Identification"
        otable10.Cell(1, 1).Range.Paragraphs(4).Format.SpaceAfter = 0
        otable10.Cell(1, 1).Range.Paragraphs(4).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(4).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(4).Range.Bold = True
        otable10.Cell(1, 1).Range.Paragraphs(4).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(4).Range.Italic = False

        '        If UCase(TRA).IndexOf(UCase("EY_SysMan      COMPUTED")) = -1 Then
        '            temp = "All entries are considered as manual entries."
        '            GoTo t1
        '        End If

        '        TRA_TEMP2 = TRA.Substring(UCase(TRA).IndexOf(UCase("EY_SysMan      COMPUTED")) + 23)
        '        TRA_TEMP1 = TRA_TEMP2.SUBSTRING(1, TRA_TEMP2.INDEXOF("COMPUTED") - 1)
        '        END_INDEX1 = TRA_TEMP1.INDEXOF("EY_") - 1
        '        START_INDEX1 = 0
        '        For a = END_INDEX1 To 1 Step -1
        '            start3 = TRA_TEMP1.substring(a, 1)
        '            If start3 = "(" Then
        '                START_INDEX1 = 1
        '                Exit For
        '            End If
        '        Next a
        '        LEN1 = END_INDEX1 - START_INDEX1
        '        If LEN1 <= 0 Then
        '            temp = "All entries are considered as manual entries."
        '        ElseIf Trim(UCase(temp)) = Chr(34) & "MANUAL" & Chr(34) Then
        '            temp = "All entries are considered as manual entries."
        '        Else
        '            temp = TRA_TEMP1.substring(START_INDEX1, LEN1)
        '        End If

        't1:     otable10.Cell(1, 1).Range.InsertParagraphAfter()
        '        otable10.Cell(1, 1).Range.Paragraphs(5).Range.Text = Chr(9) & "•   " & Replace(temp, Chr(10), " ") & "<@#$$@>"
        '        otable10.Cell(1, 1).Range.Paragraphs(5).Format.SpaceAfter = 10
        '        otable10.Cell(1, 1).Range.Paragraphs(5).Range.Font.Name = "Times New Roman"
        '        otable10.Cell(1, 1).Range.Paragraphs(5).Range.Font.Size = 10
        '        otable10.Cell(1, 1).Range.Paragraphs(5).Range.Bold = False
        '        otable10.Cell(1, 1).Range.Paragraphs(5).Range.Underline = False
        '        otable10.Cell(1, 1).Range.Paragraphs(5).Range.Italic = False

        '        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        '        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Text = "3." & Chr(9) & "Account Type "
        '        otable10.Cell(1, 1).Range.Paragraphs(6).Format.SpaceAfter = 0
        '        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Font.Name = "Times New Roman"
        '        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Font.Size = 10
        '        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Bold = True
        '        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Underline = False
        '        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Italic = False

        '        Try
        '            trb_temp1 = TRB.Substring(UCase(TRB).IndexOf(UCase("EY_AcctType    COMPUTED")) + 23)
        '            trb_temp = trb_temp1.Substring(trb_temp1.indexof(Chr(10)) + 1, trb_temp1.IndexOf("COMPUTED") - 1)
        '            END_INDEX1 = trb_temp.INDEXOF(",")
        '            START_INDEX1 = 0
        '            For a = END_INDEX1 To 1 Step -1
        '                start3 = trb_temp.substring(a, 1)
        '                If start3 = "(" Then
        '                    START_INDEX1 = a
        '                    Exit For
        '                End If
        '            Next a
        '            LEN1 = END_INDEX1 - START_INDEX1
        '            If LEN1 <= 0 Then temp = "<@#$$@>" Else temp = trb_temp.SUBSTRING(START_INDEX1, LEN1)
        '            If trb_temp.indexof("+") <> -1 Then
        '                temp = temp & " " & trb_temp.substring(trb_temp.INDEXOF("+"), trb_temp.INDEXOF(Chr(10)) - trb_temp.INDEXOF("+") - 1) & "<@#$$@>"
        '            End If
        '        Catch EX As Exception
        '            temp = "<@#$$@>"
        '        End Try

        If (Not (Form2.BRule.SelectedText Is Nothing)) Then

            If (Form2.BRule.SelectedItem.ToString() = "Manual") Then
                otable10.Cell(1, 1).Range.InsertParagraphAfter()
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Text = "     " & "• " & "All values were classified as manual  entries, as information regarding how to distinguish between system and manual entries was not provided. Assurance may update in the Global Analytics tool if additional information is obtained."
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Font.Name = "Times New Roman"
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Font.Size = 10
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Bold = False
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Underline = False
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Italic = False
            Else
                otable10.Cell(1, 1).Range.InsertParagraphAfter()
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Text = "     " & "• All entries with a XXXX field starting with XXXX or XXXX or XXXX were classified as System entries. All other values were classified as Manual entries. " & Chr(10) & "OR" & Chr(10) & " All the entries with a source value of Manual or Spreadsheet were classified as manual entries." & Chr(10) & "• All other source values were classified as system entries."
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Font.Name = "Times New Roman"
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Font.Size = 10
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Bold = False
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Underline = False
                otable10.Cell(1, 1).Range.Paragraphs(5).Range.Italic = False
            End If

        End If

        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Text = "3.  " & "GL Account Number"
        otable10.Cell(1, 1).Range.Paragraphs(6).Format.SpaceAfter = 0
        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Bold = True
        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(6).Range.Italic = False

        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(7).Range.Text = ""
        otable10.Cell(1, 1).Range.Paragraphs(7).Format.SpaceAfter = 0
        otable10.Cell(1, 1).Range.Paragraphs(7).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(7).Range.Font.Size = 1
        otable10.Cell(1, 1).Range.Paragraphs(7).Range.Bold = False
        otable10.Cell(1, 1).Range.Paragraphs(7).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(7).Range.Italic = False

        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(8).Range.Text = "    " & "• " & "Account Number was taken as the combination of the fields Business Unit and Account separated by a hyphen (-). For example, if the business unit is UNC01 and the account is 4225; then the computed G L account number will be UNC01-4225. "
        otable10.Cell(1, 1).Range.Paragraphs(8).Format.SpaceAfter = 0
        otable10.Cell(1, 1).Range.Paragraphs(8).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(8).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(8).Range.Bold = False
        otable10.Cell(1, 1).Range.Paragraphs(8).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(8).Range.Italic = False


        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(9).Range.Text = "3.  " & "Amount"
        otable10.Cell(1, 1).Range.Paragraphs(9).Format.SpaceAfter = 0
        otable10.Cell(1, 1).Range.Paragraphs(9).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(9).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(9).Range.Bold = True
        otable10.Cell(1, 1).Range.Paragraphs(9).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(9).Range.Italic = False

        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(10).Range.Text = "    " & "• " & "The amount field was computed using debit/credit indicator field, ""XX"", and the amount field, ""XXXX"". Debit amounts were identified by ""S"" and credit amounts were identified by ""H"" .The credit amounts were multiplied by -1 and debit amounts were taken as it is to create the amount field."
        otable10.Cell(1, 1).Range.Paragraphs(10).Format.SpaceAfter = 0
        otable10.Cell(1, 1).Range.Paragraphs(10).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(10).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(10).Range.Bold = False
        otable10.Cell(1, 1).Range.Paragraphs(10).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(10).Range.Italic = False


        'otable10.Cell(1, 1).Range.InsertParagraphAfter()
        'otable10.Cell(1, 1).Range.Paragraphs(7).Range.Text = Chr(9) & "•   " & "The account type was identified by applying the following logic on the field " & Chr(34) & "ACCOUNT_TYPE" & Chr(34) & ":" & Replace(temp, Chr(10), " ") & "<@#$$@>"
        'otable10.Cell(1, 1).Range.Paragraphs(7).Format.SpaceAfter = 10
        'otable10.Cell(1, 1).Range.Paragraphs(7).Range.Font.Name = "Times New Roman"
        'otable10.Cell(1, 1).Range.Paragraphs(7).Range.Font.Size = 10
        'otable10.Cell(1, 1).Range.Paragraphs(7).Range.Bold = False
        'otable10.Cell(1, 1).Range.Paragraphs(7).Range.Underline = False
        'otable10.Cell(1, 1).Range.Paragraphs(7).Range.Italic = False

        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(11).Range.Text = "5.  " & "Period"
        otable10.Cell(1, 1).Range.Paragraphs(11).Format.SpaceAfter = 0
        otable10.Cell(1, 1).Range.Paragraphs(11).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(11).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(11).Range.Bold = True
        otable10.Cell(1, 1).Range.Paragraphs(11).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(11).Range.Italic = False

        otable10.Cell(1, 1).Range.InsertParagraphAfter()
        otable10.Cell(1, 1).Range.Paragraphs(12).Range.Text = "     " & "• " & "The period field was computed as the month of the Effective date (ex. effective date of ""02/01/2013"" had a period value of ""02""."
        otable10.Cell(1, 1).Range.Paragraphs(12).Format.SpaceAfter = 1.5
        otable10.Cell(1, 1).Range.Paragraphs(12).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(12).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(12).Range.Bold = False
        otable10.Cell(1, 1).Range.Paragraphs(12).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(12).Range.Italic = False

        otable10.Cell(1, 1).Range.InsertParagraphAfter()

        otable10.Cell(1, 1).Range.Paragraphs(13).Range.Text = "     OR" & vbNewLine & "    •  A leading zero was added to the client provided ""Acct_Period"" field to make the length consistent to 2 characters (e.g. period 1 was computed as ""01"")." & vbNewLine & "     OR" & vbNewLine & "     •  The field Effective Date was used to define the period based on the following logic:"
        otable10.Cell(1, 1).Range.Paragraphs(13).Format.SpaceAfter = 1.5
        otable10.Cell(1, 1).Range.Paragraphs(13).Range.Font.Name = "Times New Roman"
        otable10.Cell(1, 1).Range.Paragraphs(13).Range.Font.Size = 10
        otable10.Cell(1, 1).Range.Paragraphs(13).Range.Bold = False
        otable10.Cell(1, 1).Range.Paragraphs(13).Range.Underline = False
        otable10.Cell(1, 1).Range.Paragraphs(13).Range.Italic = False

        otable10.Cell(1, 1).Range.InsertParagraphAfter()



        Dim newdoc As New Word.Document
        newdoc = oWord.Documents.Add
        'oPara8 = newdoc.Content.Paragraphs.Add


        otable11 = newdoc.Tables.Add(newdoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        otable11.Borders.Enable = True
        otable11.Columns.Item(1).Width = oWord.CentimetersToPoints(1.48)
        otable11.Columns.Item(2).Width = oWord.CentimetersToPoints(14.23)
        otable11.Rows.Height = oWord.CentimetersToPoints(0.11)
        otable11.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
        otable11.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter


        otable11.Cell(1, 1).Range.Text = "Period"
        otable11.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Bold = True
        otable11.Cell(1, 1).Range.Underline = False
        otable11.Cell(1, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable11.Cell(1, 2).Range.Text = "Period Criteria"
        otable11.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 2).Range.Font.Size = 10
        otable11.Cell(1, 2).Range.Bold = True
        otable11.Cell(1, 2).Range.Underline = False
        otable11.Cell(1, 2).Shading.BackgroundPatternColor = RGB(192, 192, 192)


        otable11.Cell(2, 1).Range.Text = """07"""
        otable11.Cell(2, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(2, 1).Range.Font.Size = 10
        otable11.Cell(2, 1).Range.Bold = False
        otable11.Cell(2, 1).Range.Underline = False

        otable11.Cell(2, 2).Range.Text = "If EY_EffectiveDt lies between ""08/01/2013"" to ""08/31/2013"""
        otable11.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(2, 2).Range.Font.Size = 10
        otable11.Cell(2, 2).Range.Bold = False
        otable11.Cell(2, 2).Range.Underline = False

        otable11.Cell(3, 1).Range.Text = """08"""
        otable11.Cell(3, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(3, 1).Range.Font.Size = 10
        otable11.Cell(3, 1).Range.Bold = False
        otable11.Cell(3, 1).Range.Underline = False

        otable11.Cell(3, 2).Range.Text = "If EY_EffectiveDt lies between ""09/01/2013"" to ""09/30/2013"""
        otable11.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(3, 2).Range.Font.Size = 10
        otable11.Cell(3, 2).Range.Bold = False
        otable11.Cell(3, 2).Range.Underline = False

        otable11.Cell(4, 1).Range.Text = """09"""
        otable11.Cell(4, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(4, 1).Range.Font.Size = 10
        otable11.Cell(4, 1).Range.Bold = False
        otable11.Cell(4, 1).Range.Underline = False

        otable11.Cell(4, 2).Range.Text = "If EY_EffectiveDt lies between ""10/01/2013"" to ""10/31/2013"""
        otable11.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(4, 2).Range.Font.Size = 10
        otable11.Cell(4, 2).Range.Bold = False
        otable11.Cell(4, 2).Range.Underline = False

        otable11.Cell(5, 1).Range.Text = """00"""
        otable11.Cell(5, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(5, 1).Range.Font.Size = 10
        otable11.Cell(5, 1).Range.Bold = False
        otable11.Cell(5, 1).Range.Underline = False

        otable11.Cell(5, 2).Range.Text = "for any dates lying outside the audit period (EY noted that no entries met this criteria)"
        otable11.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(5, 2).Range.Font.Size = 10
        otable11.Cell(5, 2).Range.Bold = False
        otable11.Cell(5, 2).Range.Underline = False


        newdoc.ActiveWindow.Selection.WholeStory()
        newdoc.ActiveWindow.Selection.Copy()
        otable10.Cell(2, 1).Range.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)


        newdoc.SaveAs2("c:\temp\test.doc")
        newdoc.Close()


        otable10.Cell(3, 1).Range.InsertParagraphAfter()
        otable10.Cell(3, 1).Range.Paragraphs(1).Range.Text = "6.  " & "Account Type"
        otable10.Cell(3, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable10.Cell(3, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable10.Cell(3, 1).Range.Paragraphs(1).Range.Font.Size = 10
        otable10.Cell(3, 1).Range.Paragraphs(1).Range.Bold = True
        otable10.Cell(3, 1).Range.Paragraphs(1).Range.Underline = False
        otable10.Cell(3, 1).Range.Paragraphs(1).Range.Italic = False

        otable10.Cell(3, 1).Range.InsertParagraphAfter()
        otable10.Cell(3, 1).Range.Paragraphs(2).Range.Text = "     " & "•  EY formatted the GL account number to include the currency, GL account number and business unit (ex. USD-10007-9790). The following account type assignment was performed using the GL account number only (10007)."
        otable10.Cell(3, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
        otable10.Cell(3, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
        otable10.Cell(3, 1).Range.Paragraphs(2).Range.Font.Size = 10
        otable10.Cell(3, 1).Range.Paragraphs(2).Range.Bold = False
        otable10.Cell(3, 1).Range.Paragraphs(2).Range.Underline = False
        otable10.Cell(3, 1).Range.Paragraphs(2).Range.Italic = False

        otable10.Cell(3, 1).Range.InsertParagraphAfter()
        otable10.Cell(3, 1).Range.Paragraphs(3).Range.Text = "     •  The account type was identified by applying the following logic to the GL account number: "
        otable10.Cell(3, 1).Range.Paragraphs(3).Format.SpaceAfter = 1.5
        otable10.Cell(3, 1).Range.Paragraphs(3).Range.Font.Name = "Times New Roman"
        otable10.Cell(3, 1).Range.Paragraphs(3).Range.Font.Size = 10
        otable10.Cell(3, 1).Range.Paragraphs(3).Range.Bold = False
        otable10.Cell(3, 1).Range.Paragraphs(3).Range.Underline = False
        otable10.Cell(3, 1).Range.Paragraphs(3).Range.Italic = False



        Dim newdoc3 As New Word.Document
        newdoc3 = oWord.Documents.Add
        'oPara8 = newdoc.Content.Paragraphs.Add

        Dim otable12 As Word.Table

        otable12 = newdoc3.Tables.Add(newdoc3.Bookmarks.Item("\endofdoc").Range, 7, 2)
        otable12.Borders.Enable = True
        otable12.Columns.Item(1).Width = oWord.CentimetersToPoints(1.48)
        otable12.Columns.Item(2).Width = oWord.CentimetersToPoints(14.23)
        otable12.Rows.Height = oWord.CentimetersToPoints(0.11)
        otable12.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
        otable12.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter

        otable12.Cell(1, 1).Range.Text = "Account Type"
        otable12.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable12.Cell(1, 1).Range.Font.Size = 10
        otable12.Cell(1, 1).Range.Bold = True
        otable12.Cell(1, 1).Range.Underline = False
        otable12.Cell(1, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable12.Cell(1, 2).Range.Text = "Account Type Criteria"
        otable12.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable12.Cell(1, 2).Range.Font.Size = 10
        otable12.Cell(1, 2).Range.Bold = True
        otable12.Cell(1, 2).Range.Underline = False
        otable12.Cell(1, 2).Shading.BackgroundPatternColor = RGB(192, 192, 192)


        otable12.Cell(2, 1).Range.Text = "Assets"
        otable12.Cell(2, 1).Range.Font.Name = "Times New Roman"
        otable12.Cell(2, 1).Range.Font.Size = 10
        otable12.Cell(2, 1).Range.Bold = False
        otable12.Cell(2, 1).Range.Underline = False

        otable12.Cell(2, 2).Range.Text = "GL accounts starting with "" XXXX """
        otable12.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otable12.Cell(2, 2).Range.Font.Size = 10
        otable12.Cell(2, 2).Range.Bold = False
        otable12.Cell(2, 2).Range.Underline = False

        otable12.Cell(3, 1).Range.Text = "Liabilities"
        otable12.Cell(3, 1).Range.Font.Name = "Times New Roman"
        otable12.Cell(3, 1).Range.Font.Size = 10
        otable12.Cell(3, 1).Range.Bold = False
        otable12.Cell(3, 1).Range.Underline = False

        otable12.Cell(3, 2).Range.Text = "GL accounts starting with ""XXXX """
        otable12.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otable12.Cell(3, 2).Range.Font.Size = 10
        otable12.Cell(3, 2).Range.Bold = False
        otable12.Cell(3, 2).Range.Underline = False

        otable12.Cell(4, 1).Range.Text = "Equity"
        otable12.Cell(4, 1).Range.Font.Name = "Times New Roman"
        otable12.Cell(4, 1).Range.Font.Size = 10
        otable12.Cell(4, 1).Range.Bold = False
        otable12.Cell(4, 1).Range.Underline = False

        otable12.Cell(4, 2).Range.Text = "GL accounts starting with "" XXXX """
        otable12.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otable12.Cell(4, 2).Range.Font.Size = 10
        otable12.Cell(4, 2).Range.Bold = False
        otable12.Cell(4, 2).Range.Underline = False

        otable12.Cell(5, 1).Range.Text = "Revenue"
        otable12.Cell(5, 1).Range.Font.Name = "Times New Roman"
        otable12.Cell(5, 1).Range.Font.Size = 10
        otable12.Cell(5, 1).Range.Bold = False
        otable12.Cell(5, 1).Range.Underline = False

        otable12.Cell(5, 2).Range.Text = "GL accounts starting with "" XXXX """
        otable12.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otable12.Cell(5, 2).Range.Font.Size = 10
        otable12.Cell(5, 2).Range.Bold = False
        otable12.Cell(5, 2).Range.Underline = False

        otable12.Cell(6, 1).Range.Text = "Expenses"
        otable12.Cell(6, 1).Range.Font.Name = "Times New Roman"
        otable12.Cell(6, 1).Range.Font.Size = 10
        otable12.Cell(6, 1).Range.Bold = False
        otable12.Cell(6, 1).Range.Underline = False

        otable12.Cell(6, 2).Range.Text = "GL accounts starting with "" XXXX """
        otable12.Cell(6, 2).Range.Font.Name = "Times New Roman"
        otable12.Cell(6, 2).Range.Font.Size = 10
        otable12.Cell(6, 2).Range.Bold = False
        otable12.Cell(6, 2).Range.Underline = False

        otable12.Cell(7, 1).Range.Text = "Undefined"
        otable12.Cell(7, 1).Range.Font.Name = "Times New Roman"
        otable12.Cell(7, 1).Range.Font.Size = 10
        otable12.Cell(7, 1).Range.Bold = False
        otable12.Cell(7, 1).Range.Underline = False

        otable12.Cell(7, 2).Range.Text = "if none of the above (EY noted that no accounts met this criteria)"
        otable12.Cell(7, 2).Range.Font.Name = "Times New Roman"
        otable12.Cell(7, 2).Range.Font.Size = 10
        otable12.Cell(7, 2).Range.Bold = False
        otable12.Cell(7, 2).Range.Underline = False

        newdoc3.ActiveWindow.Selection.WholeStory()
        newdoc3.ActiveWindow.Selection.Copy()
        otable10.Cell(4, 1).Range.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)


        newdoc3.SaveAs2("c:\temp\test.doc")
        newdoc3.Close()



        otable10.Cell(5, 1).Range.InsertParagraphAfter()
        otable10.Cell(5, 1).Range.Paragraphs(1).Range.Text = "7.  " & "Intercompany"
        otable10.Cell(5, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable10.Cell(5, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable10.Cell(5, 1).Range.Paragraphs(1).Range.Font.Size = 10
        otable10.Cell(5, 1).Range.Paragraphs(1).Range.Bold = True
        otable10.Cell(5, 1).Range.Paragraphs(1).Range.Underline = False
        otable10.Cell(5, 1).Range.Paragraphs(1).Range.Italic = False


        otable10.Cell(5, 1).Range.InsertParagraphAfter()
        otable10.Cell(5, 1).Range.Paragraphs(2).Range.Text = "     •  GL accounts ranging from XXXX to XXXX or XXXX to XXXX were mapped as ""Intercompany""."
        otable10.Cell(5, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
        otable10.Cell(5, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
        otable10.Cell(5, 1).Range.Paragraphs(2).Range.Font.Size = 10
        otable10.Cell(5, 1).Range.Paragraphs(2).Range.Bold = False
        otable10.Cell(5, 1).Range.Paragraphs(2).Range.Underline = False
        otable10.Cell(5, 1).Range.Paragraphs(2).Range.Italic = False

        otable10.Cell(5, 1).Range.InsertParagraphAfter()
        otable10.Cell(5, 1).Range.Paragraphs(3).Range.Text = "     OR" & vbNewLine & "     •  The following GL account numbers were identified as InterCompany accounts:"
        otable10.Cell(5, 1).Range.Paragraphs(3).Format.SpaceAfter = 0
        otable10.Cell(5, 1).Range.Paragraphs(3).Range.Font.Name = "Times New Roman"
        otable10.Cell(5, 1).Range.Paragraphs(3).Range.Font.Size = 10
        otable10.Cell(5, 1).Range.Paragraphs(3).Range.Bold = False
        otable10.Cell(5, 1).Range.Paragraphs(3).Range.Underline = False
        otable10.Cell(5, 1).Range.Paragraphs(3).Range.Italic = False


        Dim newdoc1 As New Word.Document
        newdoc1 = oWord.Documents.Add
        otable11 = newdoc1.Tables.Add(newdoc1.Bookmarks.Item("\endofdoc").Range, 11, 2)

        otable11.Borders.Enable = True
        otable11.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
        otable11.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter

        otable11.Cell(1, 1).Range.Text = "Account Type"
        otable11.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Bold = True
        otable11.Cell(1, 1).Range.Underline = False
        otable11.Cell(1, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable11.Cell(1, 2).Range.Text = "Account Type Criteria"
        otable11.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 2).Range.Font.Size = 10
        otable11.Cell(1, 2).Range.Bold = True
        otable11.Cell(1, 2).Range.Underline = False
        otable11.Cell(1, 2).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable11.Cell(2, 1).Range.Text = "XXXX"
        otable11.Cell(2, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(2, 1).Range.Font.Size = 10
        otable11.Cell(2, 1).Range.Bold = False
        otable11.Cell(2, 1).Range.Underline = False

        otable11.Cell(2, 2).Range.Text = "XXXX"
        otable11.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(2, 2).Range.Font.Size = 10
        otable11.Cell(2, 2).Range.Bold = False
        otable11.Cell(2, 2).Range.Underline = False
        otable11.Cell(1, 2).Shading.BackgroundPatternColor = RGB(192, 192, 192)


        otable11.Cell(3, 1).Range.Text = ""
        otable11.Cell(3, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(3, 1).Range.Font.Size = 10
        otable11.Cell(3, 1).Range.Bold = True
        otable11.Cell(3, 1).Range.Underline = False

        otable11.Cell(4, 1).Range.Text = ""
        otable11.Cell(4, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(4, 1).Range.Font.Size = 10
        otable11.Cell(4, 1).Range.Bold = False
        otable11.Cell(4, 1).Range.Underline = False

        otable11.Cell(4, 2).Range.Text = ""
        otable11.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(4, 2).Range.Font.Size = 10
        otable11.Cell(4, 2).Range.Bold = False
        otable11.Cell(4, 2).Range.Underline = False

        otable11.Cell(5, 1).Range.Text = ""
        otable11.Cell(5, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(5, 1).Range.Font.Size = 10
        otable11.Cell(5, 1).Range.Bold = False
        otable11.Cell(5, 1).Range.Underline = False

        otable11.Cell(5, 2).Range.Text = ""
        otable11.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(5, 2).Range.Font.Size = 10
        otable11.Cell(5, 2).Range.Bold = False
        otable11.Cell(5, 2).Range.Underline = False

        otable11.Cell(6, 1).Range.Text = ""
        otable11.Cell(6, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(6, 1).Range.Font.Size = 10
        otable11.Cell(6, 1).Range.Bold = False
        otable11.Cell(6, 1).Range.Underline = False

        otable11.Cell(6, 2).Range.Text = ""
        otable11.Cell(6, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(6, 2).Range.Font.Size = 10
        otable11.Cell(6, 2).Range.Bold = False
        otable11.Cell(6, 2).Range.Underline = False

        otable11.Cell(7, 1).Range.Text = ""
        otable11.Cell(7, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(7, 1).Range.Font.Size = 10
        otable11.Cell(7, 1).Range.Bold = False
        otable11.Cell(7, 1).Range.Underline = False

        otable11.Cell(7, 2).Range.Text = ""
        otable11.Cell(7, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(7, 2).Range.Font.Size = 10
        otable11.Cell(7, 2).Range.Bold = False
        otable11.Cell(7, 2).Range.Underline = False

        otable11.Cell(8, 1).Range.Text = ""
        otable11.Cell(8, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(8, 1).Range.Font.Size = 10
        otable11.Cell(8, 1).Range.Bold = False
        otable11.Cell(8, 1).Range.Underline = False

        otable11.Cell(8, 2).Range.Text = ""
        otable11.Cell(8, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(8, 2).Range.Font.Size = 10
        otable11.Cell(8, 2).Range.Bold = False
        otable11.Cell(8, 2).Range.Underline = False

        otable11.Cell(9, 1).Range.Text = ""
        otable11.Cell(9, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(9, 1).Range.Font.Size = 10
        otable11.Cell(9, 1).Range.Bold = False
        otable11.Cell(9, 1).Range.Underline = False

        otable11.Cell(9, 2).Range.Text = ""
        otable11.Cell(9, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(9, 2).Range.Font.Size = 10
        otable11.Cell(9, 2).Range.Bold = False
        otable11.Cell(9, 2).Range.Underline = False

        otable11.Cell(10, 1).Range.Text = ""
        otable11.Cell(10, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(10, 1).Range.Font.Size = 10
        otable11.Cell(10, 1).Range.Bold = False
        otable11.Cell(10, 1).Range.Underline = False

        otable11.Cell(10, 2).Range.Text = ""
        otable11.Cell(10, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(10, 2).Range.Font.Size = 10
        otable11.Cell(10, 2).Range.Bold = False
        otable11.Cell(10, 2).Range.Underline = False

        otable11.Cell(11, 1).Range.Text = ""
        otable11.Cell(11, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(11, 1).Range.Font.Size = 10
        otable11.Cell(11, 1).Range.Bold = False
        otable11.Cell(11, 1).Range.Underline = False

        otable11.Cell(11, 2).Range.Text = ""
        otable11.Cell(11, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(11, 2).Range.Font.Size = 10
        otable11.Cell(11, 2).Range.Bold = False
        otable11.Cell(11, 2).Range.Underline = False



        newdoc1.ActiveWindow.Selection.WholeStory()
        newdoc1.ActiveWindow.Selection.Copy()
        otable10.Cell(6, 1).Range.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)
        otable10.Cell(6, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        newdoc1.SaveAs2("c:\temp\test1.doc")

        otable10.Cell(7, 1).Range.InsertParagraphAfter()
        otable10.Cell(7, 1).Range.Paragraphs(1).Range.Text = "Professional Fees"
        otable10.Cell(7, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable10.Cell(7, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable10.Cell(7, 1).Range.Paragraphs(1).Range.Font.Size = 10
        otable10.Cell(7, 1).Range.Paragraphs(1).Range.Bold = True
        otable10.Cell(7, 1).Range.Paragraphs(1).Range.Underline = False
        otable10.Cell(7, 1).Range.Paragraphs(1).Range.Italic = False


        otable10.Cell(7, 1).Range.InsertParagraphAfter()
        otable10.Cell(7, 1).Range.Paragraphs(2).Range.Text = "     •  GL accounts ranging from XXXX to XXXX or XXXX to XXXX were mapped as ""Professional Fees""."
        otable10.Cell(7, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
        otable10.Cell(7, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
        otable10.Cell(7, 1).Range.Paragraphs(2).Range.Font.Size = 10
        otable10.Cell(7, 1).Range.Paragraphs(2).Range.Bold = False
        otable10.Cell(7, 1).Range.Paragraphs(2).Range.Underline = False
        otable10.Cell(7, 1).Range.Paragraphs(2).Range.Italic = False

        otable10.Cell(7, 1).Range.InsertParagraphAfter()
        otable10.Cell(7, 1).Range.Paragraphs(3).Range.Text = "     OR" & vbNewLine & "     •  The following GL account numbers were identified as Professional Fees accounts:"
        otable10.Cell(7, 1).Range.Paragraphs(3).Format.SpaceAfter = 0
        otable10.Cell(7, 1).Range.Paragraphs(3).Range.Font.Name = "Times New Roman"
        otable10.Cell(7, 1).Range.Paragraphs(3).Range.Font.Size = 10
        otable10.Cell(7, 1).Range.Paragraphs(3).Range.Bold = False
        otable10.Cell(7, 1).Range.Paragraphs(3).Range.Underline = False
        otable10.Cell(7, 1).Range.Paragraphs(3).Range.Italic = False


        newdoc1.ActiveWindow.Selection.Copy()
        otable10.Cell(8, 1).Range.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)
        otable10.Cell(8, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        newdoc1.Close()



        otable10.Cell(9, 1).Range.InsertParagraphAfter()
        otable10.Cell(9, 1).Range.Paragraphs(1).Range.Text = "9. " & "Begining Balance"
        otable10.Cell(9, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable10.Cell(9, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable10.Cell(9, 1).Range.Paragraphs(1).Range.Font.Size = 10
        otable10.Cell(9, 1).Range.Paragraphs(1).Range.Bold = True
        otable10.Cell(9, 1).Range.Paragraphs(1).Range.Underline = False
        otable10.Cell(9, 1).Range.Paragraphs(1).Range.Italic = False

        otable10.Cell(9, 1).Range.InsertParagraphAfter()
        otable10.Cell(9, 1).Range.Paragraphs(2).Range.Text = "     •  For the purpose of the TB close out, the beginning balances of all revenue and expense accounts were added to the beginning balance of the GL account XXXX –Account Name . This adjustment was made to close prior period net income to retained earnings. Additionally, the beginning balances of all revenue and expense accounts were set to $0.00. """
        otable10.Cell(9, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
        otable10.Cell(9, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
        otable10.Cell(9, 1).Range.Paragraphs(2).Range.Font.Size = 10
        otable10.Cell(9, 1).Range.Paragraphs(2).Range.Bold = False
        otable10.Cell(9, 1).Range.Paragraphs(2).Range.Underline = False
        otable10.Cell(9, 1).Range.Paragraphs(2).Range.Italic = False


        otable10.Cell(9, 1).Range.InsertParagraphAfter()
        otable10.Cell(9, 1).Range.Paragraphs(3).Range.Text = "10. " & "Account Class"
        otable10.Cell(9, 1).Range.Paragraphs(3).Format.SpaceAfter = 0
        otable10.Cell(9, 1).Range.Paragraphs(3).Range.Font.Name = "Times New Roman"
        otable10.Cell(9, 1).Range.Paragraphs(3).Range.Font.Size = 10
        otable10.Cell(9, 1).Range.Paragraphs(3).Range.Bold = True
        otable10.Cell(9, 1).Range.Paragraphs(3).Range.Underline = False
        otable10.Cell(9, 1).Range.Paragraphs(3).Range.Italic = False


        If (Form2.CheckBox1.Checked = True) Then
            otable10.Cell(9, 1).Range.InsertParagraphAfter()
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Text = "     •  For the purpose of the TB close out, the beginning balances of all revenue and expense accounts were added to the beginning balance of the GL account XXXX –Account Name . This adjustment was made to close prior period net income to retained earnings. Additionally, the beginning balances of all revenue and expense accounts were set to $0.00."
            otable10.Cell(9, 1).Range.Paragraphs(4).Format.SpaceAfter = 5
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Font.Name = "Times New Roman"
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Font.Size = 10
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Bold = False
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Underline = False
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Italic = False

            otable10.Cell(9, 1).Range.InsertParagraphAfter()
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Text = "[Attachment of COA File]"
            otable10.Cell(9, 1).Range.Paragraphs(5).Format.SpaceAfter = 10
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Font.Name = "Times New Roman"
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Font.Size = 10
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Bold = False
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Underline = False
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Italic = False
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Font.ColorIndex = Word.WdColorIndex.wdDarkRed
        Else
            otable10.Cell(9, 1).Range.InsertParagraphAfter()
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Text = "     •  The account class information was taken from the chart of accounts file. This was required for the EAGLe procedures documented below on page 5. The file utilized is attached within Appendix B to retain for future analytics. "
            otable10.Cell(9, 1).Range.Paragraphs(4).Format.SpaceAfter = 0
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Font.Name = "Times New Roman"
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Font.Size = 10
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Bold = False
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Underline = False
            otable10.Cell(9, 1).Range.Paragraphs(4).Range.Italic = False

            otable10.Cell(9, 1).Range.InsertParagraphAfter()
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Text = "     •  EY prefixed the abbreviated account type (A, L, Q, R or E) to each account class for ease of mapping within the Global Analytics Tool. "
            otable10.Cell(9, 1).Range.Paragraphs(5).Format.SpaceAfter = 10
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Font.Name = "Times New Roman"
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Font.Size = 10
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Bold = False
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Underline = False
            otable10.Cell(9, 1).Range.Paragraphs(5).Range.Italic = False
        End If






        oPara12 = oDoc.Content.Paragraphs.Add
        oPara12.Range.Text = vbNewLine & "Statements of Fact:"
        oPara12.Format.SpaceAfter = 0
        oPara12.Range.Font.Name = "Times New Roman"
        oPara12.Range.Font.Bold = True
        oPara12.Range.Font.Underline = True
        oPara12.Range.Font.Size = 10
        oPara12.Range.Font.Italic = False
        oPara12.Range.InsertParagraphAfter()

        rng = oDoc.Bookmarks.Item("\endofdoc").Range
        otable11 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 1)
        otable11.Borders.Enable = True
        otable11.Borders.InsideColor = RGB(255, 255, 255)
        otable11.Columns.Width = oWord.CentimetersToPoints(17.8)
        'otable11.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(1).Range.Text = "Control Totals:"
        otable11.Cell(1, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(1).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(1).Range.Bold = True
        otable11.Cell(1, 1).Range.Paragraphs(1).Range.Underline = True
        otable11.Cell(1, 1).Range.Paragraphs(1).Range.Italic = False

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(2).Range.Text = "     •  " & "Control totals were not provided. However, the control totals for JE dump and Trial Balance is as mentioned above (source data files) <@#$$@>"
        otable11.Cell(1, 1).Range.Paragraphs(2).Format.SpaceAfter = 10
        otable11.Cell(1, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(2).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(2).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(2).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(2).Range.Italic = False

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(3).Range.Text = "Validation Results:"
        otable11.Cell(1, 1).Range.Paragraphs(3).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(3).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(3).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(3).Range.Bold = True
        otable11.Cell(1, 1).Range.Paragraphs(3).Range.Underline = True
        otable11.Cell(1, 1).Range.Paragraphs(3).Range.Italic = False

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(4).Range.Text = "EY noted the following."
        otable11.Cell(1, 1).Range.Paragraphs(4).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(4).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(4).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(4).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(4).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(4).Range.Italic = True

        temp1 = TRD.Substring(TRD.IndexOf("The total of EY_Amount is:"))
        START_INDEX1 = TRD.IndexOf("The total of EY_Amount is:")
        END_INDEX1 = 0
        b = START_INDEX1 + 1
        Do While start3 <> Chr(10)
            start3 = TRD.Substring(b, 1)
            If start3 = Chr(10) Then
                END_INDEX1 = b
                start3 = ""
                Exit Do
            End If
            b = b + 1
        Loop
        LEN1 = END_INDEX1 - START_INDEX1
        temp = TRD.Substring(START_INDEX1 + 27, LEN1 - 27)


        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(5).Range.Text = "     •  " & "The JE amount summed to $" & String.Format("{0:0,0}", FormatNumber(CDbl(temp), 2)) & " Non-zero balances were due to rounding of transactions to two decimal places."
        otable11.Cell(1, 1).Range.Paragraphs(5).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(5).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(5).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(5).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(5).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(5).Range.Italic = False

        temp1 = TRD.Substring(TRD.IndexOf("The total of EY_BegBal is:") + 28, TRD.IndexOf("The total of EY_EndBal is:") - TRD.IndexOf("The total of EY_BegBal is:") - 28)
        temp2 = TRD.Substring(TRD.IndexOf("The total of EY_EndBal is:") + 28, TRD.IndexOf("@ TOTAL FIELDS COUNT") - TRD.IndexOf("The total of EY_EndBal is:") - 28)

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(6).Range.Text = "     •  " & "The beginning and ending trial balances summed to $" & String.Format("{0:0,0}", FormatNumber(CDbl(temp1), 2)) & " and $" & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2)) & " respectively. " & "Non-zero balances were due to rounding of transactions to two decimal places."
        otable11.Cell(1, 1).Range.Paragraphs(6).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(6).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(6).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(6).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(6).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(6).Range.Italic = False

        temp1 = TRD.Substring(TRD.IndexOf("met the test: DIFFERENCE <> 0"))
        END_INDEX1 = TRD.IndexOf("met the test: DIFFERENCE <> 0")
        START_INDEX1 = 0
        For a = END_INDEX1 To 1 Step -1
            start3 = TRD.Substring(a, 2)
            If start3 = "of" Then
                START_INDEX1 = a
                Exit For
            End If
        Next a
        LEN1 = END_INDEX1 - START_INDEX1
        temp1 = TRD.Substring(START_INDEX1 + 3, LEN1 - 3)
        END_INDEX1 = START_INDEX1 - 1
        START_INDEX1 = 0
        For a = END_INDEX1 - 1 To 1 Step -1
            start3 = TRD.Substring(a, 1)
            If start3 = " " Then
                START_INDEX1 = a
                Exit For
            End If
        Next a
        LEN1 = END_INDEX1 - START_INDEX1

        temp2 = TRD.Substring(START_INDEX1, LEN1)

        temp3 = Val(temp1) - Val(temp2)

        'CALCULATING SUM OF DIFFERENCES

        temp = TRD.Substring(TRD.IndexOf("@ CLASSIFY ON EY_AcctType ACCUMULATE EY_BegBal EY_Amount EY_EndBal ROLLFORWARD_BALANCE DIFFERENCE") + 98, TRD.IndexOf("@ EXTRACT FIELDS ALL TO " & Chr(34) & "Trial Balance Rollforward") - TRD.IndexOf("@ CLASSIFY ON EY_AcctType ACCUMULATE EY_BegBal EY_Amount EY_EndBal ROLLFORWARD_BALANCE DIFFERENCE") - 98)
        temp = temp.SUBSTRING(0, Len(temp) - 4)
        END_INDEX = Len(temp)
        START_INDEX = 0
        start3 = ""
        For a = END_INDEX - 1 To 1 Step -1
            start3 = temp.Substring(a, 1)
            If start3 = " " Then
                START_INDEX = a
                Exit For
            End If
        Next a
        LEN1 = END_INDEX - START_INDEX
        TEMP4 = temp.SUBSTRING(START_INDEX, LEN1)

        If temp2 <> 0 Then temp = "     •  " & String.Format("{0:0,0}", FormatNumber(CDbl(temp3), 0)) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(temp1), 0)) & " account balances rolled to the trial balance and " & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 0)) & " do not roll forward. However, these accounts had offsetting differences. Refer to ""Roll Forward Variance Section"" for details."

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(7).Range.Text = temp
        otable11.Cell(1, 1).Range.Paragraphs(7).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(7).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(7).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(7).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(7).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(7).Range.Italic = False

        'CALCULATE UNBALANCED JE NUMBERS

        temp1 = TRC.Substring(TRC.IndexOf("met the test: EY_Amount<>0"))
        END_INDEX1 = TRC.IndexOf("met the test: EY_Amount<>0")
        START_INDEX1 = 0
        For a = END_INDEX1 To 1 Step -1
            start3 = TRC.Substring(a, 2)
            If start3 = "of" Then
                START_INDEX1 = a
                Exit For
            End If
        Next a
        LEN1 = END_INDEX1 - START_INDEX1
        temp1 = TRC.Substring(START_INDEX1 + 3, LEN1 - 3)
        unique_jenum = temp1
        END_INDEX1 = START_INDEX1 - 1
        START_INDEX1 = 0
        For a = END_INDEX1 - 1 To 1 Step -1
            start3 = TRC.Substring(a, 1)
            If start3 = " " Then
                START_INDEX1 = a
                Exit For
            End If
        Next a
        LEN1 = END_INDEX1 - START_INDEX1

        bal_JE = TRC.Substring(START_INDEX1, LEN1)

        non_bal = Val(temp1) - Val(temp2)

        If bal_JE <> 0 Then unique_je_stmnt = "     •  " & String.Format("{0:0,0}", bal_JE) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_jenum), 0)) & " unique JE's net to $0.00. However " & String.Format("{0:0,0}", FormatNumber(CDbl(non_bal), 0)) & " JE numbers that did not sum to zero have insignificant amount." Else Unique_je_stmnt = Chr(9) & "•   " & "All of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_jenum), 0)) & " unique journal entries summed to $0.00."

        'CHECK NUMBER OF BLANK JE NUMBERS

        Try
            temp = TRC.Substring(TRC.IndexOf("@ COUNT IF ALLTRIM(EY_JENum)  = ") + 36, TRC.IndexOf("met the test: ALLTRIM(EY_JENum)  = ") - TRC.IndexOf("@ COUNT IF ALLTRIM(EY_JENum)  = ") - 36)
            START_INDEX = 0
            END_INDEX = Len(temp)
            For a = END_INDEX - 2 To 1 Step -1
                str3 = temp.substring(a, 2)
                If str3 = "of" Then
                    START_INDEX = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX - START_INDEX
            unique_je2 = temp.substring(START_INDEX + 2, LEN1 - 2)

            temp2 = temp.substring(temp.indexof(Chr(10)) + 1, START_INDEX - 1)
        Catch
            temp2 = "0"
        End Try



        'CALCULATE UNMATCHED

        temp1 = TRD.Substring(TRD.IndexOf("UNMATCHED_ROLL_TRANS"))
        temp2 = temp1.substring(temp1.indexof(Chr(10)) + 1, temp1.indexof(" records produced") - temp1.indexof(Chr(10)) - 1)
        START_INDEX = 0
        For a = Len(temp2) - 1 To 1 Step -1
            str3 = temp2.substring(a, 1)
            If str3 = " " Then
                START_INDEX = a
                Exit For
            End If
        Next a
        LEN1 = Len(temp2) - START_INDEX
        temp = temp2.substring(START_INDEX, LEN1)

        Dim xlsapp As Excel.Application
        Dim xlswkbk As Excel.Workbook
        xlsapp = New Excel.Application
        xlsapp.Visible = False
        excel_name = Form5.TextBox5.Text.Substring(0, Len(Form5.TextBox5.Text) - 4) & ".xlsx"
        xlswkbk = xlsapp.Workbooks.Open(excel_name)

        'lastrow = xlswkbk.Worksheets("Unmatched Transactions").UsedRange.Rows.Count
        lastrow = xlswkbk.Worksheets("Unmatched Transactions").range("B1048576").END(Excel.XlDirection.xlUp).ROW
        Unmatched_count = lastrow - 11

        'CHECK IF ROLLFORWARD SHEET CONTAINS CORRECT UNMATCHED UPDATED SHEET

        If Unmatched_count <> temp Then
            MessageBox.Show("The Rollforward sheet attached does not contain correct UNMATCHED sheet", "Warning!!")
        End If

        unmatch_amt = xlswkbk.Worksheets("Unmatched Transactions").range("c" & lastrow).value

        lastrow = xlswkbk.Worksheets("TB Rollforward").range("B1048576").END(Excel.XlDirection.xlUp).ROW

        'Calculating unused GL accounts from Rollforward sheet

        Counter = 0
        For a = lastrow To 9 Step -1
            If xlswkbk.Worksheets("TB Rollforward").range("C" & a).VALUE = "Only in TB" Then
                If xlswkbk.Worksheets("TB Rollforward").range("F" & a).VALUE = 0 Then
                    If xlswkbk.Worksheets("TB Rollforward").range("H" & a).VALUE = 0 Then
                        Counter = Counter + 1
                    End If
                End If
            End If
        Next a

        unused_act = Counter

        xlsapp.Quit()

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(8).Range.Text = unique_je_stmnt
        otable11.Cell(1, 1).Range.Paragraphs(8).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(8).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(8).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(8).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(8).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(8).Range.Italic = False

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(9).Range.Text = "     •  There were  " & String.Format("{0:0,0}", FormatNumber(CDbl(temp), 0)) & " unidentified GL accounts((accounts in the JE data but not in the TB).  Additionally, EY noted that " & String.Format("{0:0,0}", FormatNumber(CDbl(temp), 0)) & " GL account numbers summed to $" & String.Format("{0:0,0}", FormatNumber(CDbl(unmatch_amt), 2))
        otable11.Cell(1, 1).Range.Paragraphs(9).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(9).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(9).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(9).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(9).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(9).Range.Italic = False



        'CHECK NUMBER OF BLANK ACCOUNT NUMBERS


        Try
            temp = TRC.Substring(TRC.IndexOf("@ COUNT IF ALLTRIM(EY_Acct)   = ") + 36, TRC.IndexOf("met the test: ALLTRIM(EY_Acct)   = ") - TRC.IndexOf("@ COUNT IF ALLTRIM(EY_Acct)   = ") - 36)
            START_INDEX = 0
            END_INDEX = Len(temp)
            For a = END_INDEX - 2 To 1 Step -1
                str3 = temp.substring(a, 2)
                If str3 = "of" Then
                    START_INDEX = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX - START_INDEX
            unique_je2 = temp.substring(START_INDEX + 2, LEN1 - 2)

            temp = TRC.Substring(TRC.IndexOf("@ COUNT IF ALLTRIM(EY_Acct)   = ") + 36, TRC.IndexOf(unique_je2 - 4) - TRC.IndexOf("@ COUNT IF ALLTRIM(EY_Acct)   = ") - 36)
            START_INDEX = 0
            END_INDEX = Len(temp)
            For a = END_INDEX - 2 To 1 Step -1
                str3 = temp.substring(a, 2)
                If str3 = " " Then
                    START_INDEX = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX - START_INDEX
            temp2 = temp.substring(START_INDEX + 2, LEN1 - 2)

            'temp2 = temp.substring(temp.indexof(Chr(10)) + 1, START_INDEX - 1)
        Catch
            temp2 = "0"
        End Try






        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(10).Range.Text = "     •  " & "There were <@#$$@> of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_je2), 0)) & " JEs with blanks in JE Description. Additionally, EY noted that this was not a required field, however added to provide additional information."
        otable11.Cell(1, 1).Range.Paragraphs(10).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(10).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(10).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(10).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(10).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(10).Range.Italic = False

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(11).Range.Text = "     •  " & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 0)) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_je2), 0)) & " JEs with blank JE Numbers.Additionally, EY noted that this was not a required field, however added to provide additional information."
        otable11.Cell(1, 1).Range.Paragraphs(11).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(11).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(11).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(11).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(11).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(11).Range.Italic = False


        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(12).Range.Text = "     •  " & "There were " & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 0)) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_je2), 0)) & " JEs with blank account numbers "
        otable11.Cell(1, 1).Range.Paragraphs(12).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(12).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(12).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(12).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(12).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(12).Range.Italic = False

        Try
            temp = TRC.Substring(TRC.IndexOf("@ COUNT IF EY_EffectiveDt = ") + 38, TRC.IndexOf("met the test: EY_EffectiveDt = `19000101`") - TRC.IndexOf("@ COUNT IF EY_EffectiveDt = `19000101`") - 38)
            START_INDEX = 0
            END_INDEX = Len(temp)
            For a = END_INDEX - 2 To 1 Step -1
                str3 = temp.substring(a, 2)
                If str3 = "of" Then
                    START_INDEX = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX - START_INDEX
            unique_je2 = temp.substring(START_INDEX + 2, LEN1 - 2)

            temp2 = temp.substring(0, START_INDEX - 1)
        Catch
            temp2 = "0"
        End Try

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(13).Range.Text = "     •  " & "There were " & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 0)) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_je2), 0)) & " JEs with blank effective dates. Additionally, EY noted that this was not a required field, however added to provide additional information. "
        otable11.Cell(1, 1).Range.Paragraphs(13).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(13).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(13).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(13).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(13).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(13).Range.Italic = False

        Try
            temp = TRC.Substring(TRC.IndexOf("@ COUNT IF EY_EntryDt     = ") + 38, TRC.IndexOf("met the test: EY_EntryDt     = `19000101`") - TRC.IndexOf("@ COUNT IF EY_EntryDt     = `19000101`") - 38)
            START_INDEX = 0
            END_INDEX = Len(temp)
            For a = END_INDEX - 2 To 1 Step -1
                str3 = temp.substring(a, 2)
                If str3 = "of" Then
                    START_INDEX = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX - START_INDEX
            unique_je2 = temp.substring(START_INDEX + 2, LEN1 - 2)

            temp2 = temp.substring(0, START_INDEX - 1)
        Catch
            temp2 = "0"
        End Try

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(14).Range.Text = "     •  " & "There were " & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 0)) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_je2), 0)) & " JEs with blank entry dates. Additionally, EY noted that this was not a required field, however added to provide additional information.   "
        otable11.Cell(1, 1).Range.Paragraphs(14).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(14).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(14).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(14).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(14).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(14).Range.Italic = False

        Try
            temp = TRC.Substring(TRC.IndexOf("@ COUNT IF ALLTRIM(EY_BusUnit) = ") + 36, TRC.IndexOf("met the test: ALLTRIM(EY_BusUnit) = ") - TRC.IndexOf("@ COUNT IF ALLTRIM(EY_BusUnit) = ") - 36)
            START_INDEX = 0
            END_INDEX = Len(temp)
            For a = END_INDEX - 2 To 1 Step -1
                str3 = temp.substring(a, 2)
                If str3 = "of" Then
                    START_INDEX = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX - START_INDEX
            unique_je2 = temp.substring(START_INDEX + 2, LEN1 - 2)

            temp2 = temp.substring(0, START_INDEX - 1)
        Catch
            temp2 = "0"
        End Try
        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(15).Range.Text = "     •  " & "There were " & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 0)) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_je2), 0)) & " JEs with blank business units. Additionally, EY noted that this was not a required field, however added to provide additional information. "
        otable11.Cell(1, 1).Range.Paragraphs(15).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(15).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(15).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(15).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(15).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(15).Range.Italic = False

        Try
            temp = TRC.Substring(TRC.IndexOf("@ COUNT IF ALLTRIM(EY_Source)  = ") + 36, TRC.IndexOf("met the test: ALLTRIM(EY_Source)  = ") - TRC.IndexOf("@ COUNT IF ALLTRIM(EY_Source)  = ") - 36)
            START_INDEX = 0
            END_INDEX = Len(temp)
            For a = END_INDEX - 2 To 1 Step -1
                str3 = temp.substring(a, 2)
                If str3 = "of" Then
                    START_INDEX = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX - START_INDEX
            unique_je2 = temp.substring(START_INDEX + 2, LEN1 - 2)

            temp2 = temp.substring(0, START_INDEX - 1)
        Catch
            temp2 = "0"
        End Try

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(16).Range.Text = "     •  " & "There were " & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 0)) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_je2), 0)) & " JEs with blank source. Additionally, EY noted that this was not a required field, however added to provide additional information. "
        otable11.Cell(1, 1).Range.Paragraphs(16).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(16).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(16).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(16).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(16).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(16).Range.Italic = False

        Try
            temp = TRC.Substring(TRC.IndexOf("@ COUNT IF ALLTRIM(EY_PreparerID)  = ") + 41, TRC.IndexOf("met the test: ALLTRIM(EY_PreparerID)  = ") - TRC.IndexOf("@ COUNT IF ALLTRIM(EY_PreparerID)  = ") - 41)
            START_INDEX = 0
            END_INDEX = Len(temp)
            For a = END_INDEX - 2 To 1 Step -1
                str3 = temp.substring(a, 2)
                If str3 = "of" Then
                    START_INDEX = a
                    Exit For
                End If
            Next a
            LEN1 = END_INDEX - START_INDEX
            unique_je2 = temp.substring(START_INDEX + 2, LEN1 - 2)

            temp2 = temp.substring(temp.indexof(Chr(10)), START_INDEX - 1)
        Catch
            temp2 = "0"
        End Try

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(17).Range.Text = "     •  " & "There were " & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 0)) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_je2), 0)) & " JEs with blank preparer IDs. Additionally, EY noted that this was not a required field, however added to provide additional information."
        otable11.Cell(1, 1).Range.Paragraphs(17).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(17).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(17).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(17).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(17).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(17).Range.Italic = False



        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(18).Range.Text = "     •  " & "There were " & String.Format("{0:0,0}", FormatNumber(CDbl(unused_act), 0)) & " unused GL accounts(accounts in the TB without balances or JE activity). The beginning and ending balances for all unused GL accounts were equal to $ XX. "
        otable11.Cell(1, 1).Range.Paragraphs(18).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(18).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(18).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(18).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(18).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(18).Range.Italic = False



        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(19).Range.Text = "The above results were communicated to the core assurance team. The core assurance team should evaluate the level of risk of these observations and perform any additional audit procedures where appropriate."
        otable11.Cell(1, 1).Range.Paragraphs(19).Format.SpaceAfter = 5
        otable11.Cell(1, 1).Range.Paragraphs(19).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(19).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(19).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(19).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(19).Range.Italic = False

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(20).Range.Text = "Refer to " & roll_name & " for details of the trial balance roll forward results. "
        otable11.Cell(1, 1).Range.Paragraphs(20).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(20).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(20).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(20).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(20).Range.Italic = False
        'otable11.Cell(1, 1).Range.Paragraphs(11).Range.SetRange(9, Len(otable11.Cell(1, 1).Range.Paragraphs(11).Range.Text) - 55)
        otable11.Cell(1, 1).Range.Paragraphs(20).Range.Font.Bold = False


        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(21).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=Form5.TextBox5.Text, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:=roll_name)
        otable11.Cell(1, 1).Range.Paragraphs(21).Format.SpaceAfter = 10
        otable11.Cell(1, 1).Range.Paragraphs(21).Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(22).Range.Text = "Roll Forward Variances:"
        otable11.Cell(1, 1).Range.Paragraphs(22).Format.SpaceAfter = 1.5
        otable11.Cell(1, 1).Range.Paragraphs(22).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(22).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(22).Range.Bold = True
        otable11.Cell(1, 1).Range.Paragraphs(22).Range.Underline = True
        otable11.Cell(1, 1).Range.Paragraphs(22).Range.Italic = False

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(22 + 1).Range.Text = Chr(9) & "[Screenshot of variances if less than 20]"
        otable11.Cell(1, 1).Range.Paragraphs(22 + 1).Format.SpaceAfter = 5
        otable11.Cell(1, 1).Range.Paragraphs(22 + 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(22 + 1).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(22 + 1).Range.Bold = False
        otable11.Cell(1, 1).Range.Paragraphs(22 + 1).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(22 + 1).Range.Italic = False



        If (Form2.ComboBox2.SelectedText = "Yes") Then
            otable11.Cell(1, 1).Range.InsertParagraphAfter()
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Text = "EAGLe Tool:"
            otable11.Cell(1, 1).Range.Paragraphs(24).Format.SpaceAfter = 10
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Font.Name = "Times New Roman"
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Font.Size = 10
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Bold = True
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Underline = True
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Italic = False

            otable11.Cell(1, 1).Range.InsertParagraphAfter()
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Text = "The Global Analytics output was imported into EY EAGLe, a tool that enables the Assurance team to use the general ledger data obtained during the CAAT to support analytical procedures (planning or substantive interim work), lead sheet generation and the ability to select and investigate journal entries using ""drill down"" capability. EY IT generated a process map listing all the account classes for each specific source (e.g. manual, A/R subledger) and classified balance sheets and income statements. To enable further analysis, the Assurance team should import the Global Analytics file, including processing of the report cube, and use the EAGLe Excel macro to access the data, contacting the Assurance or ITRA coach with questions" & Chr(10) & Chr(9) & "Refer to "
            otable11.Cell(1, 1).Range.Paragraphs(25).Format.SpaceAfter = 5
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Font.Name = "Times New Roman"
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Font.Size = 10
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Bold = False
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Underline = False
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Italic = False

            otable11.Cell(1, 1).Range.Paragraphs(26).Range.InlineShapes.AddOLEObject("zip", Form2.eagleFile.Text)

            otable11.Cell(1, 1).Range.InsertParagraphAfter()
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Text = "This engagement is part of the data analysis limited deployment initiative. As a result, the Assurance team will apply guidance from the ""any any AnEY EAGLe and AAM Application guide"", located in GAAIT (Document ID 100769207), to this engagement which describes how data analysis may be used to identify risks of material misstatement and to obtain audit evidence "
            otable11.Cell(1, 1).Range.Paragraphs(26).Format.SpaceAfter = 5
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Font.Name = "Times New Roman"
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Font.Size = 10
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Bold = False
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Underline = False
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Italic = False
        Else
            otable11.Cell(1, 1).Range.InsertParagraphAfter()
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Text = ""
            otable11.Cell(1, 1).Range.Paragraphs(24).Format.SpaceAfter = 5
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Font.Name = "Times New Roman"
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Font.Size = 10
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Bold = True
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Underline = True
            otable11.Cell(1, 1).Range.Paragraphs(24).Range.Italic = False

            otable11.Cell(1, 1).Range.InsertParagraphAfter()
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Text = ""
            otable11.Cell(1, 1).Range.Paragraphs(25).Format.SpaceAfter = 5
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Font.Name = "Times New Roman"
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Font.Size = 10
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Bold = False
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Underline = False
            otable11.Cell(1, 1).Range.Paragraphs(25).Range.Italic = False

            otable11.Cell(1, 1).Range.InsertParagraphAfter()
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Text = ""
            otable11.Cell(1, 1).Range.Paragraphs(26).Format.SpaceAfter = 5
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Font.Name = "Times New Roman"
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Font.Size = 10
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Bold = False
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Underline = False
            otable11.Cell(1, 1).Range.Paragraphs(26).Range.Italic = False


        End If



        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(27).Range.Text = "Appendices"
        otable11.Cell(1, 1).Range.Paragraphs(27).Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        otable11.Cell(1, 1).Range.Paragraphs(27).Format.SpaceAfter = 5
        otable11.Cell(1, 1).Range.Paragraphs(27).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(27).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(27).Range.Underline = True
        otable11.Cell(1, 1).Range.Paragraphs(27).Range.Italic = False
        otable11.Cell(1, 1).Range.Paragraphs(27).Range.Font.Bold = True

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(28).Range.Text = "Refer to """ & myclientname & """ GAT Screenshots - Appendix A.zip"" (attached below) for details of items underlined in the Global Analytics tool."
        otable11.Cell(1, 1).Range.Paragraphs(28).Format.SpaceAfter = 5
        otable11.Cell(1, 1).Range.Paragraphs(28).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(28).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(28).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(28).Range.Italic = False
        otable11.Cell(1, 1).Range.Paragraphs(28).Range.Font.Bold = False

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(29).Range.InlineShapes.AddOLEObject("zip", Form2.TextBox2.Text)

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(30).Range.Text = "Refer to """ & myclientname & " File Structure - Appendix B. zip"" (attached below) for details on the format of the source files analyzed."
        otable11.Cell(1, 1).Range.Paragraphs(30).Format.SpaceAfter = 5
        otable11.Cell(1, 1).Range.Paragraphs(30).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(30).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(30).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(30).Range.Italic = False
        otable11.Cell(1, 1).Range.Paragraphs(30).Range.Font.Bold = False

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(31).Range.InlineShapes.AddOLEObject("zip", Form2.TextBox3.Text)

        otable11.Cell(1, 1).Range.InsertParagraphAfter()
        otable11.Cell(1, 1).Range.Paragraphs(32).Range.Text = "The actual reports were not included within this memo. These reports were provided directly to the Assurance team for analysis, evaluation of the results and documentation of the conclusions reached."
        otable11.Cell(1, 1).Range.Paragraphs(32).Format.SpaceAfter = 0
        otable11.Cell(1, 1).Range.Paragraphs(32).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Paragraphs(32).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Paragraphs(32).Range.Underline = False
        otable11.Cell(1, 1).Range.Paragraphs(32).Range.Italic = False
        otable11.Cell(1, 1).Range.Paragraphs(32).Range.Font.Bold = False



        If File.Exists(str2 & myclientname & " EY GTH JE Analysis Procedure Memo " & myPOA & ".docx") Then
            YN = MsgBox("The file already exists. Do you want to replace it", vbQuestion + vbYesNo, "Click your response")
            If YN = vbNo Then
                oWord.ActiveDocument.SaveAs(str2 & myclientname & " EY GTH JE Analysis Procedure Memo " & myPOA & "-Copy.docx")
                MessageBox.Show("Memo is prepared in " & Chr(10) & str2 & Chr(10) & " Please have a look and update the manual sections", "Complete!!")
            ElseIf YN = vbYes Then
                oWord.ActiveDocument.SaveAs(str2 & myclientname & " EY GTH JE Analysis Procedure Memo " & myPOA & ".docx")
                MessageBox.Show("Memo is prepared in " & Chr(10) & str2 & Chr(10) & " Please have a look and update the manual sections", "Complete!!")
            End If
        Else
            oWord.ActiveDocument.SaveAs(str2 & myclientname & " EY GTH JE Analysis Procedure Memo " & myPOA & ".docx")
            MessageBox.Show("Memo is prepared in " & Chr(10) & str2 & Chr(10) & " Please have a look and update the manual sections", "Complete!!")
        End If
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Form5.Show()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form6.Show()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        Form2.Show()
    End Sub
End Class
