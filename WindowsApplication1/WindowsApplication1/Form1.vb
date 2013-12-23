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
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        'Dim oTable As Word.Table


        'VALIDATIONS
        Dim a, START_INDEX1, END_INDEX1, LEN1 As Integer
        Dim start3 As String = 0
        If (Form6.TextBox1.Text = "") Then
            MessageBox.Show("Please fill up the User Information form", "Incomplete entry")
            Form6.TextBox1.Focus()
            Exit Sub
        ElseIf (Form5.TextBox1.Text = "") Then
            MessageBox.Show("Please fill up the Entry of logs", "Incomplete entry")
            Form5.TextBox1.Focus()
            Exit Sub
        End If

        Dim STrA10 As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\A10_JE_PREP.LOG")
        Dim STrA20 As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\A20_TB_PREP.LOG")
        Dim STrA30 As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\A30_MAIN.LOG")
        Dim STrC As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\C10_ROLL.LOG")
        'Dim STrA As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text.Substring(0, Len(Form5.TextBox1.Text) - 4) & ".LOG")
        'Dim STrB As IO.TextReader = System.IO.File.OpenText(Form5.TextBox2.Text.Substring(0, Len(Form5.TextBox2.Text) - 4) & ".LOG")
        'Dim STrC As IO.TextReader = System.IO.File.OpenText(Form5.TextBox3.Text.Substring(0, Len(Form5.TextBox3.Text) - 4) & ".LOG")
        'Dim STrD As IO.TextReader = System.IO.File.OpenText(Form5.TextBox4.Text.Substring(0, Len(Form5.TextBox4.Text) - 4) & ".LOG")
        Dim TRA10 As String = STrA10.ReadToEnd
        Dim TRA20 As String = STrA20.ReadToEnd
        Dim TRA30 As String = STrA30.ReadToEnd
        Dim TRC As String = STrC.ReadToEnd
        'Dim MyFileLine1 As String = Split(TRD, vbCrLf)(12)

        temp1 = TRC.Substring(TRC.IndexOf("@ ASSIGN CLIENTNAME_var"))
        START_INDEX1 = TRC.IndexOf("@ ASSIGN CLIENTNAME_var") + 26
        END_INDEX1 = 0
        b = START_INDEX1 + 1
        Do While start3 <> Chr(34)
            start3 = TRC.Substring(b, 1)
            If start3 = Chr(34) Then
                END_INDEX1 = b
                Exit Do
            End If
            b = b + 1
        Loop
        LEN1 = END_INDEX1 - START_INDEX1
        temp = TRC.Substring(START_INDEX1 - 1, LEN1)
        Dim myclientname As String = temp

        temp1 = TRC.Substring(TRC.IndexOf("@ ASSIGN PERIOD_var="))
        START_INDEX1 = TRC.IndexOf("@ ASSIGN PERIOD_var=") + 22
        END_INDEX1 = 0
        b = START_INDEX1 + 1
        start3 = ""
        Do While start3 <> Chr(34)
            start3 = TRC.Substring(b, 1)
            If start3 = Chr(34) Then
                END_INDEX1 = b
                Exit Do
            End If
            b = b + 1
        Loop
        LEN1 = END_INDEX1 - START_INDEX1
        temp = TRC.Substring(START_INDEX1, LEN1)
        Dim myPOA As String = temp
        START_POA = Trim(myPOA).Substring(0, myPOA.IndexOf(" "))
        temp = Trim(myPOA).Substring(Len(START_POA) + 1)
        end_poa = temp.Substring(temp.IndexOf(" ") + 1, Len(temp) - temp.IndexOf(" ") - 1)

        'NUMBER OF JE AND TB FILES
        Dim count_JE As Integer = Regex.Matches(TRA10, "OPEN SRC").Count
        Dim count_TB As Integer = Regex.Matches(TRA20, "OPEN SRC").Count
        Dim CLOSEOUT As Integer = 0 'Regex.Matches(TRB, "%RevExpTotal%").Count
        Dim countA_FIL As Integer
        Dim countB_FIL As Integer


        Dim LOG_PATH() As String = Form5.TextBox5.Text.Split("\")
        Dim U_BOUND As Integer = UBound(LOG_PATH)
        Dim str2 As String = ""
        For a3 = 0 To U_BOUND - 1
            str2 = str2 & LOG_PATH(a3) & "\"
        Next a3

        roll_name = Form5.TextBox5.Text.Substring(Len(str2), Len(Form5.TextBox5.Text) - Len(str2))

        'Start Word and open the document template.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add
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

        'Athithyaa's Edit

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

        Dim RCPT_DATE As String = Form6.drd.Value
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
        oPara2.Range.Text = "To perform journal entry analysis for " & myclientname & " for the current period effective " & START_POA & " through" & end_poa
        oPara2.Format.SpaceAfter = 6
        oPara2.Range.Font.Name = "Times New Roman"
        oPara2.Range.Font.Size = 10
        oPara2.Range.Font.Bold = False
        oPara2.Range.Font.Underline = False
        oPara2.Format.SpaceAfter = 4
        oPara2.Range.InsertParagraphAfter()

        'FIRST TABLE OF THE MEMO

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


        otable1.Cell(3, 1).Range.Text = "Client Contact"
        otable1.Cell(3, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(3, 1).Range.Font.Size = 10
        otable1.Cell(3, 1).Range.Bold = True
        otable1.Cell(3, 1).Range.Underline = False
        otable1.Cell(3, 1).Range.Italic = False
        otable1.Cell(3, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(3, 2).Range.Text = "Coordinated through Financial Audit Team"
        otable1.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(3, 2).Range.Font.Size = 10
        otable1.Cell(3, 2).Range.Bold = False
        otable1.Cell(3, 2).Range.Underline = False
        otable1.Cell(3, 2).Range.Italic = False

        otable1.Cell(4, 1).Range.Text = "Financial Audit Contact:"
        otable1.Cell(4, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(4, 1).Range.Font.Size = 10
        otable1.Cell(4, 1).Range.Bold = True
        otable1.Cell(4, 1).Range.Underline = False
        otable1.Cell(4, 1).Range.Italic = False
        otable1.Cell(4, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(4, 2).Range.Text = ""
        otable1.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(4, 2).Range.Font.Size = 10
        otable1.Cell(4, 2).Range.Bold = False
        otable1.Cell(4, 2).Range.Underline = False
        otable1.Cell(4, 2).Range.Italic = False

        otable1.Cell(5, 1).Range.Text = "CAAT Preparer:"
        otable1.Cell(5, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(5, 1).Range.Font.Size = 10
        otable1.Cell(5, 1).Range.Bold = True
        otable1.Cell(5, 1).Range.Underline = False
        otable1.Cell(5, 1).Range.Italic = False
        otable1.Cell(5, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(5, 2).Range.Text = Form6.ListBox1.Text
        otable1.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(5, 2).Range.Font.Size = 10
        otable1.Cell(5, 2).Range.Bold = True
        otable1.Cell(5, 2).Range.Underline = False
        otable1.Cell(5, 2).Range.Italic = False

        otable1.Cell(6, 1).Range.Text = "CAAT Reviewer:"
        otable1.Cell(6, 1).Range.Font.Name = "Times New Roman"
        otable1.Cell(6, 1).Range.Font.Size = 10
        otable1.Cell(6, 1).Range.Bold = True
        otable1.Cell(6, 1).Range.Underline = False
        otable1.Cell(6, 1).Range.Italic = False
        otable1.Cell(6, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable1.Cell(6, 2).Range.Text = Form6.ListBox2.Text
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

        If (Form6.dor.Checked) Then
            otable1.Cell(7, 2).Range.Text = Form6.drd.Value.ToString() & " " & Form6.dorv.Value.ToString() & " (Date of Re-validations),  " & Form6.dop.Value.ToString() & "(Date of Proceed)"
            otable1.Cell(7, 2).Range.Font.Name = "Times New Roman"
            otable1.Cell(7, 2).Range.Font.Size = 10
            otable1.Cell(7, 2).Range.Bold = False
            otable1.Cell(7, 2).Range.Underline = False
            otable1.Cell(7, 2).Range.Italic = False
        Else
            otable1.Cell(7, 2).Range.Text = Form6.drd.Value.ToString() & " " & Form6.dop.Value.ToString()
            otable1.Cell(7, 2).Range.Font.Name = "Times New Roman"
            otable1.Cell(7, 2).Range.Font.Size = 10
            otable1.Cell(7, 2).Range.Bold = False
            otable1.Cell(7, 2).Range.Underline = False
            otable1.Cell(7, 2).Range.Italic = False

        End If

        
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

        otable1.Cell(9, 2).Range.Text = Form6.jmdd.Value.ToString()
        otable1.Cell(9, 2).Range.Font.Name = "Times New Roman"
        otable1.Cell(9, 2).Range.Font.Size = 10
        otable1.Cell(9, 2).Range.Bold = False
        otable1.Cell(9, 2).Range.Underline = False
        otable1.Cell(9, 2).Range.Italic = False

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
        oPara3.Range.Text = "Purpose of Memorandum "
        oPara3.Range.Font.Bold = False
        oPara3.Format.SpaceAfter = 0
        oPara3.Range.Font.Name = "Times New Roman"
        oPara3.Range.Font.Bold = True
        oPara3.Range.Font.Underline = False
        oPara3.Range.Font.Italic = False
        oPara3.Range.Font.Size = 11
        oPara3.Range.InsertParagraphAfter()

        Dim otable2 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 1)
        otable2.Borders.Enable = True
        
        rng = oDoc.Bookmarks.Item("\endofdoc").Range
        otable2 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 1)
        otable2.Borders.Enable = True
        otable2.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt
        otable2.Borders.InsideColor = RGB(255, 255, 255)
        otable2.Columns.Width = oWord.CentimetersToPoints(17.8)
        'otable11.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)

        otable2.Cell(1, 1).Range.InsertParagraphAfter()
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Text = " "
        otable2.Cell(1, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Font.Size = 11
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Bold = False
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Underline = False
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Italic = False
        otable2.Cell(1, 1).Range.Paragraphs(1).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        otable2.Cell(1, 1).Range.InsertParagraphAfter()
        otable2.Cell(1, 1).Range.Paragraphs(2).Range.Text = "This memorandum and supporting JE CAAT file were prepared by the EY GTH Team for use by the audit team. The memorandum documents the objectives of the work, planned procedures, procedures executed, and our assessment of the client data. This memorandum is intended to guide and assist the audit team in performing the journal entry analysis procedures and should not be considered a standalone work paper. We have provided this memorandum in softcopy so that the audit teams may copy those portions that are deemed relevant to their audit for inclusion in the final work papers. "
        otable2.Cell(1, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
        otable2.Cell(1, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
        otable2.Cell(1, 1).Range.Paragraphs(2).Range.Font.Size = 11
        otable2.Cell(1, 1).Range.Paragraphs(2).Range.Bold = False
        otable2.Cell(1, 1).Range.Paragraphs(2).Range.Underline = False
        otable2.Cell(1, 1).Range.Paragraphs(2).Range.Italic = False
        otable2.Cell(1, 1).Range.Paragraphs(2).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        'Objective
        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara4.Range.Text = "Objective"
        oPara4.Range.Font.Bold = False
        oPara4.Format.SpaceAfter = 0
        oPara4.Range.Font.Name = "Times New Roman"
        oPara4.Range.Font.Bold = True
        oPara4.Range.Font.Underline = False
        oPara4.Range.Font.Italic = False
        oPara4.Range.Font.Size = 11
        oPara4.Range.InsertParagraphAfter()

        Dim otable3 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 1)
        otable3.Borders.Enable = True

        rng = oDoc.Bookmarks.Item("\endofdoc").Range
        otable3 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 1)
        otable3.Borders.Enable = True
        otable3.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt
        otable3.Borders.InsideColor = RGB(255, 255, 255)
        otable3.Columns.Width = oWord.CentimetersToPoints(17.8)
        
        otable3.Cell(1, 1).Range.InsertParagraphAfter()
        otable3.Cell(1, 1).Range.Paragraphs(1).Range.Bold = False
        otable3.Cell(1, 1).Range.Paragraphs(1).Range.Text = vbNewLine & "To evaluate the completeness of the Journal Entry data for " & myclientname & " for the period " & START_POA & "and " & end_poa & " . This memo will accompany the ‘eyje’ file that must be imported into the Global Analytics Tool and reviewed by the Financial Audit team"
        otable3.Cell(1, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable3.Cell(1, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable3.Cell(1, 1).Range.Paragraphs(1).Range.Font.Size = 11
        otable3.Cell(1, 1).Range.Paragraphs(1).Range.Bold = False
        otable3.Cell(1, 1).Range.Paragraphs(1).Range.Underline = False
        otable3.Cell(1, 1).Range.Paragraphs(1).Range.Italic = False
        otable3.Cell(1, 1).Range.Paragraphs(1).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        'Data Completeness, Validation and Observations
        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara4.Range.Text = "Objective"
        oPara4.Range.Font.Bold = False
        oPara4.Format.SpaceAfter = 0
        oPara4.Range.Font.Name = "Times New Roman"
        oPara4.Range.Font.Bold = True
        oPara4.Range.Font.Underline = True
        oPara4.Range.Font.Italic = False
        oPara4.Range.Font.Size = 11
        oPara4.Range.InsertParagraphAfter()

        'Dim otable4 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 1)
        'otable4.Borders.Enable = True

        rng = oDoc.Bookmarks.Item("\endofdoc").Range
        otable4 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 3, 1)
        otable4.Borders.Enable = True
        otable4.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt
        otable4.Borders.InsideColor = RGB(255, 255, 255)
        otable4.Columns.Width = oWord.CentimetersToPoints(17.8)

        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        otable4.Cell(1, 1).Range.Paragraphs(1).Range.Text = "    Any exclusion especially done for the validation."
        otable4.Cell(1, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(1).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(1).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(1).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(1).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(1).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        'Roll forward difference calculation
        temp1 = TRC.Substring(TRC.IndexOf("met the test: DIFFERENCE <> 0"))
        END_INDEX1 = TRC.IndexOf("met the test: DIFFERENCE <> 0")
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

        temp2 = TRC.Substring(START_INDEX1, LEN1)

        temp3 = Val(temp1) - Val(temp2)

        'CALCULATING SUM OF DIFFERENCES

        'temp = TRC.Substring(TRC.IndexOf("@ CLASSIFY ON EY_AcctType ACCUMULATE EY_BegBal EY_Amount EY_EndBal ROLLFORWARD_BALANCE DIFFERENCE") + 98, TRC.IndexOf("@ EXTRACT FIELDS ALL TO " & Chr(34) & "Trial Balance Rollforward") - TRC.IndexOf("@ CLASSIFY ON EY_AcctType ACCUMULATE EY_BegBal EY_Amount EY_EndBal ROLLFORWARD_BALANCE DIFFERENCE") - 98)
        'temp = temp.SUBSTRING(0, Len(temp) - 4)
        'END_INDEX = Len(temp)
        'START_INDEX = 0
        'start3 = ""
        'For a = END_INDEX - 1 To 1 Step -1
        '    start3 = temp.Substring(a, 1)
        '    If start3 = " " Then
        '        START_INDEX = a
        '        Exit For
        '    End If
        'Next a
        'LEN1 = END_INDEX - START_INDEX
        'TEMP4 = temp.SUBSTRING(START_INDEX, LEN1)

        'If temp2 <> 0 Then temp = "     •  " & String.Format("{0:0,0}", FormatNumber(CDbl(temp3), 0)) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(temp1), 0)) & " account balances rolled to the trial balance and " & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 0)) & " do not roll forward. However, these accounts had offsetting differences. Refer to ""Roll Forward Variance Section"" for details."

        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        otable4.Cell(1, 1).Range.Paragraphs(2).Range.Text = "    •   XXXX of XXXX account balances rolled to the Trial Balance and XXXX account balances did not roll. XXX of XXXX account balances that did not roll to the Trial Balance have significant rollforward differences and set off each other. There are XXXX accounts that have activity per the Journal Entry transaction and are not present in the Trial Balance file. Net Activity in XXX of these XXXX unmatched accounts summed to $0.00. Refer to the spreadsheet""" & myclientname & " " & START_POA & " thru " & end_poa & " TB Rollforward.xlsx"" for details of the Trial Balance Rollforward results."
        otable4.Cell(1, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(2).Range.Font.Size = 10
        otable4.Cell(1, 1).Range.Paragraphs(2).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(2).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(2).Range.Italic = False

        otable4.Cell(1, 1).Range.InsertParagraphAfter()

        otable4.Cell(1, 1).Range.Paragraphs(3).Range.Text = " "
        'otable4.Cell(1, 1).Range.Paragraphs(3).Range.

        otable4.Cell(1, 1).Range.Paragraphs(3).Range.Text = " The Financial Audit Team contact, ABC, investigated the rollforward results and instructed us to proceed further despite the above rollforward differences." & vbNewLine
        otable4.Cell(1, 1).Range.Paragraphs(3).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(3).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(3).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(3).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(3).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(3).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(3).Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustifyMed

        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        'otable4.Cell(1, 1).Range.Paragraphs(3).Range.
        otable4.Cell(1, 1).Range.Paragraphs(4).Range.Text = " INSERT ROLLFORWARD SPREADSHEET" & vbNewLine
        otable4.Cell(1, 1).Range.Paragraphs(4).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(4).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(4).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(4).Range.Bold = True
        otable4.Cell(1, 1).Range.Paragraphs(4).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(4).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(4).Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter


        temp1 = TRA10.Substring(TRA10.IndexOf("The total of EY_AMOUNT is:"))
        START_INDEX1 = TRA10.IndexOf("The total of EY_AMOUNT is:")
        END_INDEX1 = 0
        b = START_INDEX1 + 1
        Do While start3 <> Chr(10)
            start3 = TRC.Substring(b, 1)
            If start3 = Chr(10) Then
                END_INDEX1 = b
                start3 = ""
                Exit Do
            End If
            b = b + 1
        Loop
        LEN1 = END_INDEX1 - START_INDEX1
        temp = TRA10.Substring(START_INDEX1 + 27, LEN1 - 27)


        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        'otable4.Cell(1, 1).Range.Paragraphs(3).Range.
        otable4.Cell(1, 1).Range.Paragraphs(5).Range.Text = "     •	  The Journal Entry detail file amounts summed to " & temp
        otable4.Cell(1, 1).Range.Paragraphs(5).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(5).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(5).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(5).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(5).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(5).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(5).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

        'temp1 = Regex.Match(TRA20, "The total of EY_BegBal is:")
        Dim STemp(100) As Char
        Dim ind1 As Integer = TRA20.IndexOf("The total of EY_BEGBAL is:")
        Dim ind2 As Integer = 0
        Do While (TRA20.Chars(ind1) <> Chr(10))
            STemp(ind2) = TRA20.Chars(ind1)
            'MsgBox(TRA20.Chars(ind1))
            ind1 = ind1 + 1
            ind2 = ind2 + 1
        Loop

        Dim Atemp() As String = (STemp.ToString()).Split(" ")
        temp1 = Atemp(Atemp.GetUpperBound(0))
        MsgBox(STemp)

        'STemp = ""
        'Ending Balance
        ind1 = TRA20.IndexOf("The total of EY_ENDBAL is:")
        ind2 = 0
        Do While (TRA20.Chars(ind1) <> Chr(10))
            STemp(ind2) = TRA20.Chars(ind1)
            'MsgBox(TRA20.Chars(ind1))
            ind1 = ind1 + 1
            ind2 = ind2 + 1
        Loop
        Dim Atemp1() As String = (STemp.ToString().Split(" "))
        temp2 = Atemp1(Atemp1.GetUpperBound(0))


        'temp1 = TRA20.Substring(TRA20.IndexOf("The total of EY_BegBal is:") + 28, TRA20.IndexOf("The total of EY_EndBal is:") - TRA20.IndexOf("The total of EY_BegBal is:") - 28)
        'temp2 = TRA20.Substring(TRA20.IndexOf("The total of EY_EndBal is:") + 28, TRA20.IndexOf("@ TOTAL FIELDS COUNT") - TRA20.IndexOf("The total of EY_EndBal is:") - 28)

        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        'otable4.Cell(1, 1).Range.Paragraphs(6).Range.Text = "     •  " & "The beginning and ending trial balances summed to $" & String.Format("{0:0,0}", FormatNumber(CDbl(temp1), 2)) & " and $" & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2)) & " respectively. " & "Non-zero balances were due to rounding of transactions to two decimal places."
        otable4.Cell(1, 1).Range.Paragraphs(6).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Italic = False


        'CALCULATE UNBALANCED JE NUMBERS

        'temp1 = TRC.Substring(TRC.IndexOf("met the test: EY_Amount<>0"))
        'END_INDEX1 = TRC.IndexOf("met the test: EY_Amount<>0")
        'START_INDEX1 = 0
        'For a = END_INDEX1 To 1 Step -1
        '    start3 = TRC.Substring(a, 2)
        '    If start3 = "of" Then
        '        START_INDEX1 = a
        '        Exit For
        '    End If
        'Next a
        'LEN1 = END_INDEX1 - START_INDEX1
        'temp1 = TRC.Substring(START_INDEX1 + 3, LEN1 - 3)
        'unique_jenum = temp1
        'END_INDEX1 = START_INDEX1 - 1
        'START_INDEX1 = 0
        'For a = END_INDEX1 - 1 To 1 Step -1
        '    start3 = TRC.Substring(a, 1)
        '    If start3 = " " Then
        '        START_INDEX1 = a
        '        Exit For
        '    End If
        'Next a
        'LEN1 = END_INDEX1 - START_INDEX1

        'bal_JE = TRC.Substring(START_INDEX1, LEN1)

        'non_bal = Val(temp1) - Val(temp2)

        'If bal_JE <> 0 Then unique_je_stmnt = "     •  " & String.Format("{0:0,0}", bal_JE) & " of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_jenum), 0)) & " unique JE's net to $0.00. However " & String.Format("{0:0,0}", FormatNumber(CDbl(non_bal), 0)) & " JE numbers that did not sum to zero have insignificant amount." Else Unique_je_stmnt = Chr(9) & "•   " & "All of " & String.Format("{0:0,0}", FormatNumber(CDbl(unique_jenum), 0)) & " unique journal entries summed to $0.00."

        'temp1 = TRA20.Substring(TRA20.IndexOf("The total of EY_BegBal is:") + 28, TRA20.IndexOf("The total of EY_EndBal is:") - TRA20.IndexOf("The total of EY_BegBal is:") - 28)
        'temp2 = TRA20.Substring(TRA20.IndexOf("The total of EY_EndBal is:") + 28, TRA20.IndexOf("@ TOTAL FIELDS COUNT") - TRA20.IndexOf("The total of EY_EndBal is:") - 28)


        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        'otable4.Cell(1, 1).Range.Paragraphs(6).Range.Text = "     •  " & "The beginning and ending trial balances summed to $" & String.Format("{0:0,0}", FormatNumber(CDbl(temp1), 2)) & " and $" & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2)) & " respectively. " & "Non-zero balances were due to rounding of transactions to two decimal places."
        otable4.Cell(1, 1).Range.Paragraphs(6).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(6).Range.Italic = False

        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        otable4.Cell(1, 1).Range.Paragraphs(7).Range.Text = "     •  " & "XXXX of XXXX unique journal entries summed to $0.00 and XX unique journal entries did not sum to $0.00. X of these XX unbalanced accounts have immaterial amounts.Refer to the attached spreadsheet """ & myclientname & " " & START_POA & " thru " & end_poa & "Unbalanced Journal Entries.xlsx"" for details of the unbalanced journal entries." & vbNewLine
        otable4.Cell(1, 1).Range.Paragraphs(7).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(7).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(7).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(7).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(7).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(7).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(8).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        otable4.Cell(1, 1).Range.Paragraphs(8).Range.Text = "INSERT UNBALANCED SPREADSHEET"
        otable4.Cell(1, 1).Range.Paragraphs(8).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(8).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(8).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(8).Range.Bold = True
        otable4.Cell(1, 1).Range.Paragraphs(8).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(8).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(8).Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        otable4.Cell(1, 1).Range.Paragraphs(9).Range.Text = "     •	 There are XX of XXXXX line items with zero amounts" & vbNewLine & "     •	 There are XX line items with a blank Preparer ID"
        otable4.Cell(1, 1).Range.Paragraphs(9).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(9).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(9).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(9).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(9).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(9).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(9).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

        'otable4.Cell(1, 1).Range.InsertParagraphAfter()
        'otable4.Cell(1, 1).Range.Paragraphs(9).Range.Text = "     •	 There are XX line items with a blank Preparer ID"
        'otable4.Cell(1, 1).Range.Paragraphs(9).Format.SpaceAfter = 0
        'otable4.Cell(1, 1).Range.Paragraphs(9).Range.Font.Name = "Times New Roman"
        'otable4.Cell(1, 1).Range.Paragraphs(9).Range.Font.Size = 11
        'otable4.Cell(1, 1).Range.Paragraphs(9).Range.Bold = False
        'otable4.Cell(1, 1).Range.Paragraphs(9).Range.Underline = False
        'otable4.Cell(1, 1).Range.Paragraphs(9).Range.Italic = False


        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        otable4.Cell(1, 1).Range.Paragraphs(10).Range.Text = "     •  There are XX line items with a blank JE Description"
        otable4.Cell(1, 1).Range.Paragraphs(10).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(10).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(10).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(10).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(10).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(10).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(10).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        otable4.Cell(1, 1).Range.Paragraphs(11).Range.Text = "     •  Entry Date is as early as MM/DD/YYYY and as late as MM/DD/YYYY."
        otable4.Cell(1, 1).Range.Paragraphs(11).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(11).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(11).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(11).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(11).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(11).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(11).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        otable4.Cell(1, 1).Range.Paragraphs(12).Range.Text = "     •  Effective Date is as early as " & Form6.eFrom.Value.ToString("M/d/yyyy") & "and as late as" & Form6.eFrom.Value.ToString("M/d/yyyy")
        otable4.Cell(1, 1).Range.Paragraphs(12).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(12).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(12).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(12).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(12).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(12).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(12).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        otable4.Cell(1, 1).Range.InsertParagraphAfter()
        otable4.Cell(1, 1).Range.Paragraphs(13).Range.Text = "     •  Field_1 and Field_2 were not provided in the data"
        otable4.Cell(1, 1).Range.Paragraphs(13).Format.SpaceAfter = 0
        otable4.Cell(1, 1).Range.Paragraphs(13).Range.Font.Name = "Times New Roman"
        otable4.Cell(1, 1).Range.Paragraphs(13).Range.Font.Size = 11
        otable4.Cell(1, 1).Range.Paragraphs(13).Range.Bold = False
        otable4.Cell(1, 1).Range.Paragraphs(13).Range.Underline = False
        otable4.Cell(1, 1).Range.Paragraphs(13).Range.Italic = False
        otable4.Cell(1, 1).Range.Paragraphs(13).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        otable4.Cell(1, 1).Range.Paragraphs(13).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


        
        Dim ui_input1 = 1

        If (ui_input1 = 1) Then
            otable4.Cell(1, 1).Range.InsertParagraphAfter()
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Text = "     •	 We reset opening balances of Income Statement accounts to $0.00 and transferred Net Income to the Retained Earnings account ""XXXXXX"". (In case of usage of Reset-Retained) "
            otable4.Cell(1, 1).Range.Paragraphs(14).Format.SpaceAfter = 0
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Font.Name = "Times New Roman"
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Font.Size = 11
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Bold = False
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Underline = False
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Italic = False
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Font.ColorIndex = Word.WdColorIndex.wdBlack
            otable4.Cell(1, 1).Range.Paragraphs(14).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

        Else
            otable4.Cell(1, 1).Range.InsertParagraphAfter()
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Text = " "
            otable4.Cell(1, 1).Range.Paragraphs(14).Format.SpaceAfter = 0
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Font.Name = "Times New Roman"
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Font.Size = 11
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Bold = False
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Underline = False
            otable4.Cell(1, 1).Range.Paragraphs(14).Range.Italic = False
            otable4.Cell(1, 1).Range.Paragraphs(14).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

        End If



        Dim ui_input2 As Boolean = True
        Dim otable11 As Word.Table

        If (ui_input2) Then
            otable4.Cell(1, 1).Range.InsertParagraphAfter()
            otable4.Cell(1, 1).Range.Paragraphs(15).Range.Text = "     •  We reconciled record counts and control totals provided by client for the Journal Entry data as follows:"
            otable4.Cell(1, 1).Range.Paragraphs(15).Format.SpaceAfter = 0
            otable4.Cell(1, 1).Range.Paragraphs(15).Range.Font.Name = "Times New Roman"
            otable4.Cell(1, 1).Range.Paragraphs(15).Range.Font.Size = 11
            otable4.Cell(1, 1).Range.Paragraphs(15).Range.Bold = False
            otable4.Cell(1, 1).Range.Paragraphs(15).Range.Underline = False
            otable4.Cell(1, 1).Range.Paragraphs(15).Range.Italic = False
            otable4.Cell(1, 1).Range.Paragraphs(15).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft


            otable4.Cell(1, 1).Range.InsertParagraphAfter()


            Dim newdoc1 As New Word.Document
            newdoc1 = oWord.Documents.Add
            otable11 = newdoc1.Tables.Add(newdoc1.Bookmarks.Item("\endofdoc").Range, 8, 3)
            otable11.Cell(1, 1).Merge(otable11.Cell(1, 3))

            otable11.Borders.Enable = True
            otable11.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
            otable11.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter


            otable11.Cell(1, 1).Range.Text = "Journal Entry Data Record Counts and Control Totals "
            otable11.Cell(1, 1).Range.Font.Name = "Times New Roman"
            otable11.Cell(1, 1).Range.Font.Size = 10
            otable11.Cell(1, 1).Range.Bold = True
            otable11.Cell(1, 1).Range.Underline = False
            otable11.Cell(1, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)



            otable11.Cell(2, 1).Range.Text = "File Name"
            otable11.Cell(2, 1).Range.Font.Name = "Times New Roman"
            otable11.Cell(2, 1).Range.Font.Size = 10
            otable11.Cell(2, 1).Range.Bold = True
            otable11.Cell(2, 1).Range.Underline = False

            otable11.Cell(2, 2).Range.Text = "Record Count"
            otable11.Cell(2, 2).Range.Font.Name = "Times New Roman"
            otable11.Cell(2, 2).Range.Font.Size = 10
            otable11.Cell(2, 2).Range.Bold = True
            otable11.Cell(2, 2).Range.Underline = False
            'otable11.Cell(2, 2).Shading.BackgroundPatternColor = RGB(192, 192, 192)

            otable11.Cell(2, 3).Range.Text = "Total Amount"
            otable11.Cell(2, 3).Range.Font.Name = "Times New Roman"
            otable11.Cell(2, 3).Range.Font.Size = 10
            otable11.Cell(2, 3).Range.Bold = True
            otable11.Cell(2, 3).Range.Underline = False
            'otable11.Cell(2, 3).Shading.BackgroundPatternColor = RGB(192, 192, 192)


            otable11.Cell(3, 1).Range.Text = "Source_File_1"
            otable11.Cell(3, 1).Range.Font.Name = "Times New Roman"
            otable11.Cell(3, 1).Range.Font.Size = 10
            otable11.Cell(3, 1).Range.Bold = False
            otable11.Cell(3, 1).Range.Underline = False

            otable11.Cell(3, 2).Range.Text = "XXXX"
            otable11.Cell(3, 2).Range.Font.Name = "Times New Roman"
            otable11.Cell(3, 2).Range.Font.Size = 10
            otable11.Cell(3, 2).Range.Bold = False
            otable11.Cell(3, 2).Range.Underline = False

            otable11.Cell(3, 3).Range.Text = "$0.00"
            otable11.Cell(3, 3).Range.Font.Name = "Times New Roman"
            otable11.Cell(3, 3).Range.Font.Size = 10
            otable11.Cell(3, 3).Range.Bold = False
            otable11.Cell(3, 3).Range.Underline = False


            otable11.Cell(4, 1).Range.Text = "Source_File_2"
            otable11.Cell(4, 1).Range.Font.Name = "Times New Roman"
            otable11.Cell(4, 1).Range.Font.Size = 10
            otable11.Cell(4, 1).Range.Bold = False
            otable11.Cell(4, 1).Range.Underline = False

            otable11.Cell(4, 2).Range.Text = "XXXX"
            otable11.Cell(4, 2).Range.Font.Name = "Times New Roman"
            otable11.Cell(4, 2).Range.Font.Size = 10
            otable11.Cell(4, 2).Range.Bold = False
            otable11.Cell(4, 2).Range.Underline = False

            otable11.Cell(4, 3).Range.Text = "$0.00"
            otable11.Cell(4, 3).Range.Font.Name = "Times New Roman"
            otable11.Cell(4, 3).Range.Font.Size = 10
            otable11.Cell(4, 3).Range.Bold = False
            otable11.Cell(4, 3).Range.Underline = False

            otable11.Cell(5, 1).Range.Text = "Source_File_1"
            otable11.Cell(5, 1).Range.Font.Name = "Times New Roman"
            otable11.Cell(5, 1).Range.Font.Size = 10
            otable11.Cell(5, 1).Range.Bold = False
            otable11.Cell(5, 1).Range.Underline = False

            otable11.Cell(5, 2).Range.Text = "XXXX"
            otable11.Cell(5, 2).Range.Font.Name = "Times New Roman"
            otable11.Cell(5, 2).Range.Font.Size = 10
            otable11.Cell(5, 2).Range.Bold = False
            otable11.Cell(5, 2).Range.Underline = False

            otable11.Cell(5, 3).Range.Text = "$0.00"
            otable11.Cell(5, 3).Range.Font.Name = "Times New Roman"
            otable11.Cell(5, 3).Range.Font.Size = 10
            otable11.Cell(5, 3).Range.Bold = False
            otable11.Cell(5, 3).Range.Underline = False

            otable11.Cell(6, 1).Range.Text = "Source_File_1"
            otable11.Cell(6, 1).Range.Font.Name = "Times New Roman"
            otable11.Cell(6, 1).Range.Font.Size = 10
            otable11.Cell(6, 1).Range.Bold = False
            otable11.Cell(6, 1).Range.Underline = False

            otable11.Cell(6, 2).Range.Text = "XXXX"
            otable11.Cell(6, 2).Range.Font.Name = "Times New Roman"
            otable11.Cell(6, 2).Range.Font.Size = 10
            otable11.Cell(6, 2).Range.Bold = False
            otable11.Cell(6, 2).Range.Underline = False

            otable11.Cell(6, 3).Range.Text = "$0.00"
            otable11.Cell(6, 3).Range.Font.Name = "Times New Roman"
            otable11.Cell(6, 3).Range.Font.Size = 10
            otable11.Cell(6, 3).Range.Bold = False
            otable11.Cell(6, 3).Range.Underline = False

            otable11.Cell(7, 1).Range.Text = "Source_File_1"
            otable11.Cell(7, 1).Range.Font.Name = "Times New Roman"
            otable11.Cell(7, 1).Range.Font.Size = 10
            otable11.Cell(7, 1).Range.Bold = False
            otable11.Cell(7, 1).Range.Underline = False

            otable11.Cell(7, 2).Range.Text = "XXXX"
            otable11.Cell(7, 2).Range.Font.Name = "Times New Roman"
            otable11.Cell(7, 2).Range.Font.Size = 10
            otable11.Cell(7, 2).Range.Bold = False
            otable11.Cell(7, 2).Range.Underline = False

            otable11.Cell(7, 3).Range.Text = "$0.00"
            otable11.Cell(7, 3).Range.Font.Name = "Times New Roman"
            otable11.Cell(7, 3).Range.Font.Size = 10
            otable11.Cell(7, 3).Range.Bold = False
            otable11.Cell(7, 3).Range.Underline = False

            otable11.Cell(8, 1).Range.Text = "Totals"
            otable11.Cell(8, 1).Range.Font.Name = "Times New Roman"
            otable11.Cell(8, 1).Range.Font.Size = 10
            otable11.Cell(8, 1).Range.Bold = False
            otable11.Cell(8, 1).Range.Underline = False

            otable11.Cell(8, 2).Range.Text = ""
            otable11.Cell(8, 2).Range.Font.Name = "Times New Roman"
            otable11.Cell(8, 2).Range.Font.Size = 10
            otable11.Cell(8, 2).Range.Bold = False
            otable11.Cell(8, 2).Range.Underline = False


            newdoc1.ActiveWindow.Selection.WholeStory()
            newdoc1.ActiveWindow.Selection.Copy()
            otable4.Cell(2, 1).Range.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)
            otable4.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            newdoc1.SaveAs2("c:\temp\test1.doc")

        Else
            otable4.Cell(2, 1).Range.InsertParagraphAfter()
            otable4.Cell(2, 1).Range.Paragraphs(1).Range.Text = "     •	 We used" & Form6.eFrom.Value.ToString("M/d/yyyy") & "Through" & Form6.eTo.Value.ToString("M/d/yyyy") & "to designate current period journal entries in the EY Global Analytics Tool"
            otable4.Cell(2, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
            otable4.Cell(2, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
            otable4.Cell(2, 1).Range.Paragraphs(1).Range.Font.Size = 11
            otable4.Cell(2, 1).Range.Paragraphs(1).Range.Bold = False
            otable4.Cell(2, 1).Range.Paragraphs(1).Range.Underline = False
            otable4.Cell(2, 1).Range.Paragraphs(1).Range.Italic = False

        End If

        Dim ui_input3 As Boolean = True


        If True Then
            otable4.Cell(3, 1).Range.InsertParagraphAfter()
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Text = "     •	 We identified unmatched accounts as ""Unmatched"" and mapped to their respective Account Type e.g. ""Unmatched Assets"" for the unmatched Assets accounts. (If Unmatched Accounts are present"
            otable4.Cell(3, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Font.Size = 11
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Bold = False
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Underline = False
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Italic = False
        Else
            otable4.Cell(3, 1).Range.InsertParagraphAfter()
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Text = " "
            otable4.Cell(3, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Font.Size = 11
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Bold = False
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Underline = False
            otable4.Cell(3, 1).Range.Paragraphs(1).Range.Italic = False
        End If

        'CALCULATE UNMATCHED

        temp1 = TRC.Substring(TRC.IndexOf("UNMATCHED_ROLL_TRANS"))
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


        otable4.Cell(3, 1).Range.InsertParagraphAfter()
        otable4.Cell(3, 1).Range.Paragraphs(2).Range.Text = "     •	 There are " & String.Format("{0:0,0}", FormatNumber(CDbl(temp), 0)) & " GL Accounts in the TB data without balances or JE activity. These accounts were excluded from further processing by the EY Global Analytics Tool."
        otable4.Cell(3, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
        otable4.Cell(3, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
        otable4.Cell(3, 1).Range.Paragraphs(2).Range.Font.Size = 11
        otable4.Cell(3, 1).Range.Paragraphs(2).Range.Bold = False
        otable4.Cell(3, 1).Range.Paragraphs(2).Range.Underline = False
        otable4.Cell(3, 1).Range.Paragraphs(2).Range.Italic = False
        

        'Conclusion
        oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara5.Range.Text = "Conclusion"
        oPara5.Range.Font.Bold = False
        oPara5.Format.SpaceAfter = 0
        oPara5.Range.Font.Name = "Times New Roman"
        oPara5.Range.Font.Bold = True
        oPara5.Range.Font.Underline = True
        oPara5.Range.Font.Italic = False
        oPara5.Range.Font.Size = 11
        oPara5.Range.InsertParagraphAfter()

        Dim otable5 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 1)
        otable5.Borders.Enable = True
        otable5.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt

        rng = oDoc.Bookmarks.Item("\endofdoc").Range
        otable5 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 3, 1)
        otable5.Borders.Enable = True
        otable5.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt
        otable5.Borders.InsideColor = RGB(255, 255, 255)
        otable5.Columns.Width = oWord.CentimetersToPoints(17.8)

        Dim ui_rollback = 0
        If (ui_rollback <> 0) Then
            otable5.Cell(1, 1).Range.InsertParagraphAfter()
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Text = " "
            otable5.Cell(1, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Font.Size = 11
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Bold = False
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Underline = False
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Italic = False

            otable5.Cell(1, 1).Range.InsertParagraphAfter()
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Text = "Per the procedures performed, we conclude that the journal entry data is valid and can be relied upon by the Financial Audit team, but we were unable to test 100% completeness of the data as XXXX account balances did not roll forward successfully, out of which XX accounts have significant roll forward differences. We recommend that the Financial Audit team independently review the reasonableness of any noted items in the Data Completeness, Validation and Observation sections and conclude on their reliance strategy for the journal entry data."
            otable5.Cell(1, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Font.Size = 11
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Bold = False
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Underline = False
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Italic = False
        Else

            otable5.Cell(1, 1).Range.InsertParagraphAfter()
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Text = " "
            otable5.Cell(1, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Font.Size = 11
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Bold = False
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Underline = False
            otable5.Cell(1, 1).Range.Paragraphs(1).Range.Italic = False

            otable5.Cell(1, 1).Range.InsertParagraphAfter()
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Text = "As per the procedures performed, we conclude that the journal entry data is valid and can be relied upon by the financial audit team. We were able to test 100% completeness of the data as all account balances rolled forward successfully. We recommend that the financial audit team independently review the reasonableness of any noted items in the Data Completeness, Validation and Observation sections and conclude on their reliance strategy for the journal entry data"
            otable5.Cell(1, 1).Range.Paragraphs(2).Format.SpaceAfter = 0
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Font.Name = "Times New Roman"
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Font.Size = 11
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Bold = False
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Underline = False
            otable5.Cell(1, 1).Range.Paragraphs(2).Range.Italic = False
        End If

        'Data Audit Trail
        oPara6 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara6.Range.Text = "Data Audit Trail"
        oPara6.Range.Font.Bold = False
        oPara6.Format.SpaceAfter = 0
        oPara6.Range.Font.Name = "Times New Roman"
        oPara6.Range.Font.Bold = True
        oPara6.Range.Font.Underline = False
        oPara6.Range.Font.Italic = False
        oPara6.Range.Font.Size = 11
        oPara6.Range.InsertParagraphAfter()

        Dim otable6 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 1)
        otable6.Borders.Enable = True
        otable6.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt

        rng = oDoc.Bookmarks.Item("\endofdoc").Range
        otable6 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 1)
        otable6.Borders.Enable = True
        otable6.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt
        otable6.Borders.InsideColor = RGB(255, 255, 255)
        otable6.Columns.Width = oWord.CentimetersToPoints(17.8)

        otable6.Cell(1, 1).Range.InsertParagraphAfter()
        otable6.Cell(1, 1).Range.Paragraphs(1).Range.Text = "The following sections relate to all client data manipulation performed throughout the JE CAAT by the Enterprise Intelligence and Data Analytics team. The items below serve as the audit trail of our procedures and are to be referenced in future runs to leverage the efficiencies gained through recurring execution."
        otable6.Cell(1, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable6.Cell(1, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable6.Cell(1, 1).Range.Paragraphs(1).Range.Font.Size = 11
        otable6.Cell(1, 1).Range.Paragraphs(1).Range.Bold = False
        otable6.Cell(1, 1).Range.Paragraphs(1).Range.Underline = False
        otable6.Cell(1, 1).Range.Paragraphs(1).Range.Italic = False

        'Insert another paragraph.
        oPara7 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara7.Range.Text = "Method of Analysis: "
        oPara7.Range.Font.Bold = False
        oPara7.Format.SpaceAfter = 0
        oPara7.Range.Font.Name = "Times New Roman"
        oPara7.Range.Font.Bold = True
        oPara7.Range.Font.Underline = True
        oPara7.Range.Font.Italic = False
        oPara7.Range.Font.Size = 10
        oPara7.Range.InsertParagraphAfter()

        Dim otable7 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 1, 8)
        otable7.Borders.Enable = True
        otable7.Columns.Item(1).Width = oWord.CentimetersToPoints(0.87)
        otable7.Columns.Item(3).Width = oWord.CentimetersToPoints(0.87)
        otable7.Columns.Item(5).Width = oWord.CentimetersToPoints(0.87)
        otable7.Columns.Item(7).Width = oWord.CentimetersToPoints(0.87)
        otable7.Columns.Item(2).Width = oWord.CentimetersToPoints(4.61)
        otable7.Columns.Item(4).Width = oWord.CentimetersToPoints(3.05)
        otable7.Columns.Item(6).Width = oWord.CentimetersToPoints(3.05)
        otable7.Columns.Item(8).Width = oWord.CentimetersToPoints(3.58)
        otable7.Rows.Height = oWord.CentimetersToPoints(0.4)

        otable7.Cell(1, 1).Range.Text = If(Form6.gl.Checked, ChrW(9746), ChrW(9744))
        otable7.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable7.Cell(1, 1).Range.Font.Size = 12
        otable7.Cell(1, 1).Range.Bold = False
        otable7.Cell(1, 1).Range.Underline = False
        otable7.Cell(1, 1).Range.Italic = False
        otable7.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable7.Cell(1, 3).Range.Text = If(Form6.CheckBox2.Checked, ChrW(9746), ChrW(9744))
        otable7.Cell(1, 3).Range.Font.Name = "Times New Roman"
        otable7.Cell(1, 3).Range.Font.Size = 12
        otable7.Cell(1, 3).Range.Bold = False
        otable7.Cell(1, 3).Range.Underline = False
        otable7.Cell(1, 3).Range.Italic = False
        otable7.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable7.Cell(1, 5).Range.Text = If(Form6.CheckBox3.Checked, ChrW(9746), ChrW(9744))
        otable7.Cell(1, 5).Range.Font.Name = "Times New Roman"
        otable7.Cell(1, 5).Range.Font.Size = 12
        otable7.Cell(1, 5).Range.Bold = False
        otable7.Cell(1, 5).Range.Underline = False
        otable7.Cell(1, 5).Range.Italic = False
        otable7.Cell(1, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable7.Cell(1, 7).Range.Text = If(Form6.CheckBox4.Checked, ChrW(9746), ChrW(9744))
        otable7.Cell(1, 7).Range.Font.Name = "Times New Roman"
        otable7.Cell(1, 7).Range.Font.Size = 12
        otable7.Cell(1, 7).Range.Bold = False
        otable7.Cell(1, 7).Range.Underline = False
        otable7.Cell(1, 7).Range.Italic = False
        otable7.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable7.Cell(1, 2).Range.Text = "EY/Global Analytics Module"
        otable7.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable7.Cell(1, 2).Range.Font.Size = 10
        otable7.Cell(1, 2).Range.Bold = False
        otable7.Cell(1, 2).Range.Underline = False
        otable7.Cell(1, 2).Range.Italic = False
        otable7.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable7.Cell(1, 4).Range.Text = "ACL"
        otable7.Cell(1, 4).Range.Font.Name = "Times New Roman"
        otable7.Cell(1, 4).Range.Font.Size = 10
        otable7.Cell(1, 4).Range.Bold = False
        otable7.Cell(1, 4).Range.Underline = False
        otable7.Cell(1, 4).Range.Italic = False
        otable7.Cell(1, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable7.Cell(1, 6).Range.Text = "MS Access"
        otable7.Cell(1, 6).Range.Font.Name = "Times New Roman"
        otable7.Cell(1, 6).Range.Font.Size = 10
        otable7.Cell(1, 6).Range.Bold = False
        otable7.Cell(1, 6).Range.Underline = False
        otable7.Cell(1, 6).Range.Italic = False
        otable7.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable7.Cell(1, 8).Range.Text = "Other:" & If(Form6.CheckBox4.Checked, Form6.other.Text, " ")
        otable7.Cell(1, 8).Range.Font.Name = "Times New Roman"
        otable7.Cell(1, 8).Range.Font.Size = 10
        otable7.Cell(1, 8).Range.Bold = False
        otable7.Cell(1, 8).Range.Underline = False
        otable7.Cell(1, 8).Range.Italic = False
        otable7.Cell(1, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

        otable7.Rows.Height = oWord.CentimetersToPoints(0.4)

        oPara7 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara7.Range.Text = "NOTE:  Agreed upon per discussion with Financial Audit team."
        oPara7.Range.Font.Bold = False
        oPara7.Format.SpaceAfter = 6
        oPara7.Range.Font.Name = "Times New Roman"
        oPara7.Range.Font.Bold = False
        oPara7.Range.Font.Underline = False
        oPara7.Range.Font.Size = 10
        oPara7.Range.Font.Italic = True
        oPara7.Range.InsertParagraphAfter()
        oPara7.Range.Words(1).Font.Bold = True



        'BreakPoint
        Dim otable8 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, count_JE + count_TB, 5)
        otable8.Borders.Enable = True
        otable8.Columns.Item(1).Width = oWord.CentimetersToPoints(6.44)
        otable8.Columns.Item(2).Width = oWord.CentimetersToPoints(2)
        otable8.Columns.Item(3).Width = oWord.CentimetersToPoints(2.75)
        otable8.Columns.Item(4).Width = oWord.CentimetersToPoints(3.25)
        otable8.Columns.Item(5).Width = oWord.CentimetersToPoints(3.32)
        otable8.Rows.Item(1).Height = oWord.CentimetersToPoints(0.49)

        otable8.Cell(1, 1).Range.Text = "Data File Name"
        otable8.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable8.Cell(1, 1).Range.Font.Size = 10
        otable8.Cell(1, 1).Range.Bold = True
        otable8.Cell(1, 1).Range.Underline = False
        otable8.Cell(1, 1).Range.Italic = True
        otable8.Cell(1, 1).Shading.BackgroundPatternColor = RGB(224, 224, 224)

        otable8.Cell(1, 2).Range.Text = "Record Count"
        otable8.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable8.Cell(1, 2).Range.Font.Size = 10
        otable8.Cell(1, 2).Range.Bold = True
        otable8.Cell(1, 2).Range.Underline = False
        otable8.Cell(1, 2).Range.Italic = True
        otable8.Cell(1, 2).Shading.BackgroundPatternColor = RGB(224, 224, 224)

        otable8.Cell(1, 3).Range.Text = "Control Total"
        otable8.Cell(1, 3).Range.Font.Name = "Times New Roman"
        otable8.Cell(1, 3).Range.Font.Size = 10
        otable8.Cell(1, 3).Range.Bold = True
        otable8.Cell(1, 3).Range.Underline = False
        otable8.Cell(1, 3).Range.Italic = True
        otable8.Cell(1, 3).Shading.BackgroundPatternColor = RGB(224, 224, 224)

        otable8.Cell(1, 4).Range.Text = "Description"
        otable8.Cell(1, 4).Range.Font.Name = "Times New Roman"
        otable8.Cell(1, 4).Range.Font.Size = 10
        otable8.Cell(1, 4).Range.Bold = True
        otable8.Cell(1, 4).Range.Underline = False
        otable8.Cell(1, 4).Range.Italic = True
        otable8.Cell(1, 4).Shading.BackgroundPatternColor = RGB(224, 224, 224)

        otable8.Cell(1, 5).Range.Text = "ACL Table Name" & " (*.fil)"
        otable8.Cell(1, 5).Range.Font.Name = "Times New Roman"
        otable8.Cell(1, 5).Range.Font.Size = 10
        otable8.Cell(1, 5).Range.Bold = True
        otable8.Cell(1, 5).Range.Underline = False
        otable8.Cell(1, 5).Range.Italic = True
        otable8.Cell(1, 5).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        Dim C As Integer

        'EXTRACTING JE DATA FROM A LOG

        Dim i As Integer = TRA10.IndexOf("Opening file name ")
        Dim i_JE1 As Integer = TRA10.IndexOf("ACTIVATE")
        Do While (i_JE1 <> -1)

            'GETTING THE RAW ACL TABLE NAMES

            TRA_TEMP1 = TRA10.Substring(i + 18)
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

                Dim extraction_je_count As String = TRA_TEMP1.Substring(STARTINDEX, endIndex - STARTINDEX).Trim
                Dim str1 As String = ""
                str1 = String.Format("{0:0,0}", FormatNumber(CDbl(extraction_je_count), 0))

                otable8.Cell(2 + C - 1, 2).Range.Text = str1
                otable8.Cell(2 + C - 1 - 1, 2).Range.Font.Name = "Times New Roman"
                otable8.Cell(2 + C - 1, 2).Range.Font.Size = 10
                otable8.Cell(2 + C - 1, 2).Range.Bold = False
                otable8.Cell(2 + C - 1, 2).Range.Underline = False
                otable8.Cell(2 + C - 1, 2).Range.Italic = False

                otable8.Cell(2 + C - 1, 5).Range.Text = je_file_name
                otable8.Cell(2 + C - 1, 5).Range.Font.Name = "Times New Roman"
                otable8.Cell(2 + C - 1, 5).Range.Font.Size = 10
                otable8.Cell(2 + C - 1, 5).Range.Bold = False
                otable8.Cell(2 + C - 1, 5).Range.Underline = False
                otable8.Cell(2 + C - 1, 5).Range.Italic = False

                temp1 = TRA10.Substring(TRA10.IndexOf("The total of EY_AMOUNT is:"))
                START_INDEX1 = TRA10.IndexOf("The total of EY_AMOUNT is:")
                '28:
                END_INDEX1 = 0
                Do While temp1(END_INDEX1) <> Chr(10)
                    END_INDEX1 = END_INDEX1 + 1
                Loop
                END_INDEX1 = END_INDEX1 - 1
                'b = START_INDEX1 + 1
                'Do While start3 <> "@"
                '    start3 = TRA10.Substring(b, 1)
                '    If start3 = "@" Then
                '        END_INDEX1 = b
                '        start3 = ""
                '        Exit Do
                '    End If
                '    b = b + 1
                'Loop
                'LEN1 = END_INDEX1 - START_INDEX1
                temp = TRA10.Substring(START_INDEX1 + 27, END_INDEX1 - 26)

                'MsgBox(temp)
                je_amount = "Amount: $" & FormatNumber(CDbl(temp), 2)


                otable8.Cell(2 + C - 1, 3).Range.Text = "Amount: $" & FormatNumber(CDbl(temp), 2)
                otable8.Cell(2 + C - 1, 3).Range.Font.Name = "Times New Roman"
                otable8.Cell(2 + C - 1, 3).Range.Font.Size = 10
                otable8.Cell(2 + C - 1, 3).Range.Bold = False
                otable8.Cell(2 + C - 1, 3).Range.Underline = False
                otable8.Cell(2 + C - 1, 3).Range.Italic = False

                otable8.Cell(2 + C - 1, 4).Range.Text = "JE Activity "
                otable8.Cell(2 + C - 1, 4).Range.Font.Name = "Times New Roman"
                otable8.Cell(2 + C - 1, 4).Range.Font.Size = 10
                otable8.Cell(2 + C - 1, 4).Range.Bold = False
                otable8.Cell(2 + C - 1, 4).Range.Underline = False
                otable8.Cell(2 + C - 1, 4).Range.Italic = False

                i = TRA10.IndexOf("Opening file name ", i + 1)
                i_JE1 = TRA10.IndexOf("ACTIVATE", i_JE1 + 1)
            Else
                Exit Do
            End If
        Loop

        'If count_JE > 1 Then
        '    With otable8
        '        .Cell(2, 2).Merge(.Cell(2 + count_JE - 1, 2))
        '        .Cell(2, 3).Merge(.Cell(2 + count_JE - 1, 3))
        '        .Cell(2, 3).Range.Text = je_amount
        '        .Cell(2, 4).Merge(.Cell(2 + count_JE - 1, 4))
        '        .Cell(2, 5).Merge(.Cell(2 + count_JE - 1, 5))
        '        .Cell(2, 4).Range.Text = "JE Activity for " & myPOA
        '    End With
        'End If

        'EXTRACTING TB DATA FROM B LOG
        Dim i_TB As Integer = TRA20.IndexOf("Opening file name")
        Dim i_TB1 As Integer = TRA20.IndexOf("ACTIVATE")
        Dim D As Integer = 0

        Do While (i_TB1 <> -1)


            'GETTING THE RAW ACL TABLE NAMES

            TRB_TEMP1 = TRA20.Substring(i_TB + 18)
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

                    Dim extraction_tb_count As String = TRB_TEMP1.Substring(STARTINDEX, endIndex - STARTINDEX).Trim

                    'Dim TB_file_name As String = TRB_TEMP.Substring(startIndex1, endIndex1 - startIndex1).Trim
                    Dim str1 = String.Format("{0:0,0}", FormatNumber(CDbl(extraction_tb_count), 0))
                    otable8.Cell(2 + C + D - 1, 2).Range.Text = str1
                    otable8.Cell(2 + C + D - 1, 2).Range.Font.Name = "Times New Roman"
                    otable8.Cell(2 + C + D - 1, 2).Range.Font.Size = 10
                    otable8.Cell(2 + C + D - 1, 2).Range.Bold = False
                    otable8.Cell(2 + C + D - 1, 2).Range.Underline = False
                    otable8.Cell(2 + C + D - 1, 2).Range.Italic = False

                    otable8.Cell(2 + C + D - 1, 5).Range.Text = TB_file_name
                    otable8.Cell(2 + C + D - 1, 5).Range.Font.Name = "Times New Roman"
                    otable8.Cell(2 + C + D - 1, 5).Range.Font.Size = 10
                    otable8.Cell(2 + C + D - 1, 5).Range.Bold = False
                    otable8.Cell(2 + C + D - 1, 5).Range.Underline = False
                    otable8.Cell(2 + C + D - 1, 5).Range.Italic = False

                    TRD_TEMP = TRA20.Substring(TRA20.IndexOf("@  TOTAL FIELDS EY_BEGBAL EY_ENDBAL"))
                    If UCase(TB_file_name).Contains("BEG") Then
                        temp1 = "Beginning Balance: $"
                        temp2 = TRD_TEMP.Substring(TRD_TEMP.IndexOf("The total of EY_BEGBAL is:  ") + 28, TRD_TEMP.IndexOf("The total of EY_ENDBAL is:  ") - 28 - TRD_TEMP.IndexOf("The total of EY_BEGBAL is:  "))
                        temp = temp1 & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2))
                        name1 = "Beginning trial balance as on " & Chr(10) & START_POA
                    ElseIf UCase(TB_file_name).Contains("END") Then
                        temp1 = "Ending Balance: $"
                        temp2 = TRD_TEMP.Substring(TRD_TEMP.IndexOf("The total of EY_ENDBAL is:  ") + 28, 5)
                        MsgBox("end bal" & temp2)
                        temp = temp & temp1 & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2))
                        name1 = "Ending trial balance as on " & Chr(10) & end_poa
                    Else
                        temp1 = "Beginning Balance: $"
                        temp2 = TRD_TEMP.Substring(TRD_TEMP.IndexOf("The total of EY_BEGBAL is: ") + 28, TRD_TEMP.IndexOf("The total of EY_ENDBAL is:  ") - 28 - TRD_TEMP.IndexOf("The total of EY_BEGBAL is:  "))
                        temp_1 = temp1 & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2))
                        temp3 = "Ending Balance: $"
                        temp4 = TRD_TEMP.Substring(TRD_TEMP.IndexOf("The total of EY_ENDBAL is:  ") + 28, 5)

                        temp_2 = temp3 & String.Format("{0:0,0}", FormatNumber(CDbl(temp4), 2))
                        MsgBox("else " & temp_2)
                        temp = temp1 & vbCrLf & temp2
                        MsgBox(temp)
                        name1 = "Beginning trial balance as on " & Chr(10) & START_POA & " and " & Chr(10) & "Ending trial balance as on " & Chr(10) & end_poa
                    End If
                Catch ex As Exception
                    If (TypeOf Err.GetException() Is ArgumentOutOfRangeException) Then
                        temp1 = "Beginning Balance: $"
                        temp2 = TRA20.Substring(TRA20.IndexOf("The total of EY_BEGBAL is:") + 28, TRA20.IndexOf("The total of EY_ENDBAL is:") - TRA20.IndexOf("The total of EY_BEGBAL is:") - 28)
                        'temp2 = TRB_TEMP.Substring(TRB_TEMP.IndexOf("The total of EY_BegBal is:  ") + 28, TRB_TEMP.IndexOf("The total of EY_EndBal is:  ") - 28 - TRB_TEMP.IndexOf("The total of EY_BegBal is:  "))
                        temp_1 = temp1 & String.Format("{0:0,0}", FormatNumber(CDbl(temp2), 2))
                        temp3 = "Ending Balance: $"
                        TEMP4 = TRA20.Substring(TRA20.IndexOf("The total of EY_ENDBAL is:") + 28, TRA20.IndexOf(">>> COMMAND <3> ") - TRA20.IndexOf("The total of EY_ENDBAL is:") - 28)
                        MsgBox(TEMP4)
                        'temp4 = TRB_TEMP.Substring(TRB_TEMP.IndexOf("The total of EY_EndBal is:  ") + 28, 5)
                        temp_2 = temp3 & String.Format("{0:0,0}", FormatNumber(CDbl(TEMP4), 2))
                        temp = Replace(temp_1, Chr(10), " ") & vbCrLf & Replace(temp_2, Chr(10), " ")
                        name1 = "Beginning trial balance as on " & Chr(10) & START_POA & " and " & Chr(10) & "Ending trial balance as on " & Chr(10) & end_poa
                    End If
                End Try

                tb_amount = temp


                otable8.Cell(2 + C + D - 1, 3).Range.Text = temp
                otable8.Cell(2 + C + D - 1, 3).Range.Font.Name = "Times New Roman"
                otable8.Cell(2 + C + D - 1, 3).Range.Font.Size = 10
                otable8.Cell(2 + C + D - 1, 3).Range.Bold = False
                otable8.Cell(2 + C + D - 1, 3).Range.Underline = False
                otable8.Cell(2 + C + D - 1, 3).Range.Italic = False

                otable8.Cell(2 + C + D - 1, 4).Range.Text = name1
                otable8.Cell(2 + C + D - 1, 4).Range.Font.Name = "Times New Roman"
                otable8.Cell(2 + C + D - 1, 4).Range.Font.Size = 10
                otable8.Cell(2 + C + D - 1, 4).Range.Bold = False
                otable8.Cell(2 + C + D - 1, 4).Range.Underline = False
                otable8.Cell(2 + C + D - 1, 4).Range.Italic = False

                i_TB = TRA20.IndexOf("Opening file name", i_TB + 1)
                i_TB1 = TRA20.IndexOf("ACTIVATE", i_TB + 1)
            Else
                Exit Do
            End If
        Loop










































































































































































































































































































        oParaAthi4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oParaAthi4.Range.Text = "ACL Scripts:"
        oParaAthi4.Range.Font.Bold = False
        oParaAthi4.Format.SpaceAfter = 0
        oParaAthi4.Range.Font.Name = "Times New Roman"
        oParaAthi4.Range.Font.Bold = True
        oParaAthi4.Range.Font.Underline = True
        oParaAthi4.Range.Font.Italic = False
        oParaAthi4.Range.Font.Size = 10
        oParaAthi4.Range.InsertParagraphAfter()


        Dim unbalancedflag As Integer = 1

        Dim otableAthi4 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 10 - unbalancedflag, 5)
        otableAthi4.Borders.Enable = True

        otableAthi4.Columns.Item(1).Width = oWord.CentimetersToPoints(0.6)
        otableAthi4.Columns.Item(2).Width = oWord.CentimetersToPoints(2.95)
        otableAthi4.Columns.Item(3).Width = oWord.CentimetersToPoints(6.63)
        otableAthi4.Columns.Item(4).Width = oWord.CentimetersToPoints(4.39)
        otableAthi4.Columns.Item(5).Width = oWord.CentimetersToPoints(3.18)
        otableAthi4.Rows.Item(7).Height = oWord.CentimetersToPoints(0.4)

        With otableAthi4
            .Cell(9 - unbalancedflag, 1).Merge(.Cell(9 - unbalancedflag, 5))
            .Cell(10 - unbalancedflag, 1).Merge(.Cell(10 - unbalancedflag, 5))
        End With

        otableAthi4.Cell(1, 1).Range.Text = "#"
        otableAthi4.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(1, 1).Range.Font.Size = 10
        otableAthi4.Cell(1, 1).Range.Bold = True
        otableAthi4.Cell(1, 1).Range.Underline = False
        otableAthi4.Cell(1, 1).Range.Italic = True
        otableAthi4.Cell(1, 1).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        For a = 2 To 8 - unbalancedflag
            otableAthi4.Cell(a, 1).Range.Text = a - 1
            otableAthi4.Cell(a, 1).Range.Font.Name = "Times New Roman"
            otableAthi4.Cell(a, 1).Range.Font.Size = 10
            otableAthi4.Cell(a, 1).Range.Bold = False
            otableAthi4.Cell(a, 1).Range.Underline = False
            otableAthi4.Cell(a, 1).Range.Italic = False
        Next

        otableAthi4.Cell(1, 2).Range.Text = "Script"
        otableAthi4.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(1, 2).Range.Font.Size = 10
        otableAthi4.Cell(1, 2).Range.Bold = True
        otableAthi4.Cell(1, 2).Range.Underline = False
        otableAthi4.Cell(1, 2).Range.Italic = True
        otableAthi4.Cell(1, 2).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        otableAthi4.Cell(2, 2).Range.Text = "A10_JE_ PREP"
        otableAthi4.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(2, 2).Range.Font.Size = 10
        otableAthi4.Cell(2, 2).Range.Bold = False
        otableAthi4.Cell(2, 2).Range.Underline = False
        otableAthi4.Cell(2, 2).Range.Italic = False

        otableAthi4.Cell(3, 2).Range.Text = "A20_TB_ PREP"
        otableAthi4.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(3, 2).Range.Font.Size = 10
        otableAthi4.Cell(3, 2).Range.Bold = False
        otableAthi4.Cell(3, 2).Range.Underline = False
        otableAthi4.Cell(3, 2).Range.Italic = False

        otableAthi4.Cell(4, 2).Range.Text = "A30_MAIN"
        otableAthi4.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(4, 2).Range.Font.Size = 10
        otableAthi4.Cell(4, 2).Range.Bold = False
        otableAthi4.Cell(4, 2).Range.Underline = False
        otableAthi4.Cell(4, 2).Range.Italic = False

        otableAthi4.Cell(5, 2).Range.Text = "B10_VALIDATION"
        otableAthi4.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(5, 2).Range.Font.Size = 10
        otableAthi4.Cell(5, 2).Range.Bold = False
        otableAthi4.Cell(5, 2).Range.Underline = False
        otableAthi4.Cell(5, 2).Range.Italic = False

        otableAthi4.Cell(6, 2).Range.Text = "C10_ROLL"
        otableAthi4.Cell(6, 2).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(6, 2).Range.Font.Size = 10
        otableAthi4.Cell(6, 2).Range.Bold = False
        otableAthi4.Cell(6, 2).Range.Underline = False
        otableAthi4.Cell(6, 2).Range.Italic = False

        otableAthi4.Cell(7, 2).Range.Text = "D10_GLOBAL_JE_MAPPING"
        otableAthi4.Cell(7, 2).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(7, 2).Range.Font.Size = 10
        otableAthi4.Cell(7, 2).Range.Bold = False
        otableAthi4.Cell(7, 2).Range.Underline = False
        otableAthi4.Cell(7, 2).Range.Italic = False

        If (unbalancedflag = 0) Then
            otableAthi4.Cell(8, 2).Range.Text = "E10_UNBALANCED_ENTRIES"
            otableAthi4.Cell(8, 2).Range.Font.Name = "Times New Roman"
            otableAthi4.Cell(8, 2).Range.Font.Size = 10
            otableAthi4.Cell(8, 2).Range.Bold = False
            otableAthi4.Cell(8, 2).Range.Underline = False
            otableAthi4.Cell(8, 2).Range.Italic = False
            otableAthi4.Cell(8, 2).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        End If

        otableAthi4.Cell(1, 3).Range.Text = "Purpose"
        otableAthi4.Cell(1, 3).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(1, 3).Range.Font.Size = 10
        otableAthi4.Cell(1, 3).Range.Bold = True
        otableAthi4.Cell(1, 3).Range.Underline = False
        otableAthi4.Cell(1, 3).Range.Italic = True
        otableAthi4.Cell(1, 3).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        Dim rngtbl2 As Word.Range

        otableAthi4.Cell(2, 3).Range.Text = "Formats source JE data files and consolidates into individual table (if necessary)." & vbNewLine & "NOTE: A manual import of the JE data files must be performed prior to running this script."
        otableAthi4.Cell(2, 3).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(2, 3).Range.Font.Size = 10
        otableAthi4.Cell(2, 3).Range.Bold = False
        otableAthi4.Cell(2, 3).Range.Underline = False
        otableAthi4.Cell(2, 3).Range.Italic = False
        otableAthi4.Cell(2, 3).Range.Words(16).Font.Italic = True
        otableAthi4.Cell(2, 3).Range.Words(16).Font.Bold = True
        otableAthi4.Cell(2, 3).Range.Words(16).Font.Underline = True
        rngtbl2 = oWord.ActiveDocument.Range(otableAthi4.Cell(2, 3).Range.Words(17).Start, otableAthi4.Cell(2, 3).Range.Words(33).End)
        rngtbl2.Font.Italic = True


        otableAthi4.Cell(3, 3).Range.Text = "Formats source TB data files and consolidates into individual table (if necessary)." & vbNewLine & "NOTE: A manual import of the TB data files must be performed prior to running this script."
        otableAthi4.Cell(3, 3).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(3, 3).Range.Font.Size = 10
        otableAthi4.Cell(3, 3).Range.Bold = False
        otableAthi4.Cell(3, 3).Range.Underline = False
        otableAthi4.Cell(3, 3).Range.Italic = False
        otableAthi4.Cell(3, 3).Range.Words(16).Font.Italic = True
        otableAthi4.Cell(3, 3).Range.Words(16).Font.Bold = True
        otableAthi4.Cell(3, 3).Range.Words(16).Font.Underline = True
        rngtbl2 = oWord.ActiveDocument.Range(otableAthi4.Cell(3, 3).Range.Words(17).Start, otableAthi4.Cell(3, 3).Range.Words(33).End)
        rngtbl2.Font.Italic = True


        otableAthi4.Cell(4, 3).Range.Text = "Sets up the fields that will be utilized for the processing of the transaction data file."
        otableAthi4.Cell(4, 3).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(4, 3).Range.Font.Size = 10
        otableAthi4.Cell(4, 3).Range.Bold = False
        otableAthi4.Cell(4, 3).Range.Underline = False
        otableAthi4.Cell(4, 3).Range.Italic = False

        otableAthi4.Cell(5, 3).Range.Text = "Contains all the validation checks that must be performed on all EY GTH journal entry CAATs." & vbNewLine & "NOTE: Any exceptions that arise from the validation results should be highlighted to the Sub-Area ITRA SPOC and resolved and documented in the memo for future reference."
        otableAthi4.Cell(5, 3).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(5, 3).Range.Font.Size = 10
        otableAthi4.Cell(5, 3).Range.Bold = False
        otableAthi4.Cell(5, 3).Range.Underline = False
        otableAthi4.Cell(5, 3).Range.Italic = False
        otableAthi4.Cell(5, 3).Range.Words(19).Font.Italic = True
        otableAthi4.Cell(5, 3).Range.Words(19).Font.Bold = True
        otableAthi4.Cell(5, 3).Range.Words(19).Font.Underline = True
        rngtbl2 = oWord.ActiveDocument.Range(otableAthi4.Cell(5, 3).Range.Words(21).Start, otableAthi4.Cell(5, 3).Range.Words(48).End)
        rngtbl2.Font.Italic = True


        otableAthi4.Cell(6, 3).Range.Text = "This script performs the rollforward test using the transaction and the trial balance data files." & vbNewLine & "NOTE: All rollforward differences must be documented and resolved prior to proceeding with the CAAT."
        otableAthi4.Cell(6, 3).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(6, 3).Range.Font.Size = 10
        otableAthi4.Cell(6, 3).Range.Bold = False
        otableAthi4.Cell(6, 3).Range.Underline = False
        otableAthi4.Cell(6, 3).Range.Italic = False
        otableAthi4.Cell(6, 3).Range.Words(18).Font.Italic = True
        otableAthi4.Cell(6, 3).Range.Words(18).Font.Bold = True
        otableAthi4.Cell(6, 3).Range.Words(18).Font.Underline = True
        rngtbl2 = oWord.ActiveDocument.Range(otableAthi4.Cell(6, 3).Range.Words(20).Start, otableAthi4.Cell(6, 3).Range.Words(34).End)
        rngtbl2.Font.Italic = True


        otableAthi4.Cell(7, 3).Range.Text = "This script prepares and exports the transaction and trial balance files for upload into the global analytics tool."
        otableAthi4.Cell(7, 3).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(7, 3).Range.Font.Size = 10
        otableAthi4.Cell(7, 3).Range.Bold = False
        otableAthi4.Cell(7, 3).Range.Underline = False
        otableAthi4.Cell(7, 3).Range.Italic = False

        If (unbalancedflag = 0) Then
            otableAthi4.Cell(8, 3).Range.Text = "This script prepares and exports the unbalanced journal entries."
            otableAthi4.Cell(8, 3).Range.Font.Name = "Times New Roman"
            otableAthi4.Cell(8, 3).Range.Font.Size = 10
            otableAthi4.Cell(8, 3).Range.Bold = False
            otableAthi4.Cell(8, 3).Range.Underline = False
            otableAthi4.Cell(8, 3).Range.Italic = False
            otableAthi4.Cell(8, 3).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        End If

        'UPDATING 4TH COLUMN

        otableAthi4.Cell(1, 4).Range.Text = "Input Files Created" & Chr(32) & " (*.FIL)"
        otableAthi4.Cell(1, 4).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(1, 4).Range.Font.Size = 10
        otableAthi4.Cell(1, 4).Range.Bold = True
        otableAthi4.Cell(1, 4).Range.Underline = False
        otableAthi4.Cell(1, 4).Range.Italic = True
        otableAthi4.Cell(1, 4).Shading.BackgroundPatternColor = RGB(12, 12, 12)


        Dim rngtbl As Word.Range

        otableAthi4.Cell(2, 4).Range.Text = "-EY_JE" & vbNewLine & "-EY_JE_EXCLUDE" & vbNewLine & "-(Write name if any other file was created prior to preparing EY_JE)"
        otableAthi4.Cell(2, 4).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(2, 4).Range.Font.Size = 10
        otableAthi4.Cell(2, 4).Range.Bold = False
        otableAthi4.Cell(2, 4).Range.Underline = False
        otableAthi4.Cell(2, 4).Range.Italic = False
        rngtbl = oWord.ActiveDocument.Range(otableAthi4.Cell(2, 4).Range.Words(12).Start, otableAthi4.Cell(2, 4).Range.Words(27).End)
        rngtbl.Font.ColorIndex = Word.WdColorIndex.wdRed

        otableAthi4.Cell(3, 4).Range.Text = "-EY_TB" & vbNewLine & "-EY_TEMP_TB" & vbNewLine & "-EY_TB_EXCLUDE"
        otableAthi4.Cell(3, 4).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(3, 4).Range.Font.Size = 10
        otableAthi4.Cell(3, 4).Range.Bold = False
        otableAthi4.Cell(3, 4).Range.Underline = False
        otableAthi4.Cell(3, 4).Range.Italic = False

        otableAthi4.Cell(4, 4).Range.Text = "See attached log file."
        otableAthi4.Cell(4, 4).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(4, 4).Range.Font.Size = 10
        otableAthi4.Cell(4, 4).Range.Bold = False
        otableAthi4.Cell(4, 4).Range.Underline = False
        otableAthi4.Cell(4, 4).Range.Italic = False

        otableAthi4.Cell(5, 4).Range.Text = "-EY_JE _GROUPING"
        otableAthi4.Cell(5, 4).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(5, 4).Range.Font.Size = 10
        otableAthi4.Cell(5, 4).Range.Bold = False
        otableAthi4.Cell(5, 4).Range.Underline = False
        otableAthi4.Cell(5, 4).Range.Italic = False

        otableAthi4.Cell(6, 4).Range.Text = "-Trial Balance Rollforward" & vbNewLine & "-UNMATCHED_ROLL_TRANS" & vbNewLine & "-" & myclientname & " " & START_POA & " thru " & end_poa & " TB Rollforward.xlsx"
        otableAthi4.Cell(6, 4).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(6, 4).Range.Font.Size = 10
        otableAthi4.Cell(6, 4).Range.Bold = False
        otableAthi4.Cell(6, 4).Range.Underline = False
        otableAthi4.Cell(6, 4).Range.Italic = False

        otableAthi4.Cell(7, 4).Range.Text = "-EY_JE.txt" & vbNewLine & "-EY_TB.txt"
        otableAthi4.Cell(7, 4).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(7, 4).Range.Font.Size = 10
        otableAthi4.Cell(7, 4).Range.Bold = False
        otableAthi4.Cell(7, 4).Range.Underline = False
        otableAthi4.Cell(7, 4).Range.Italic = False

        If (unbalancedflag = 0) Then
            otableAthi4.Cell(8, 4).Range.Text = "-" & myclientname & " " & START_POA & " " & end_poa & " Unbalanced Journal Entries"
            otableAthi4.Cell(8, 4).Range.Font.Name = "Times New Roman"
            otableAthi4.Cell(8, 4).Range.Font.Size = 10
            otableAthi4.Cell(8, 4).Range.Bold = False
            otableAthi4.Cell(8, 4).Range.Underline = False
            otableAthi4.Cell(8, 4).Range.Italic = False
        End If


        'UPDATING 5TH COLUMN

        otableAthi4.Cell(1, 5).Range.Text = "ACL Logs" & Chr(32) & " (*.LOG)"
        otableAthi4.Cell(1, 5).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(1, 5).Range.Font.Size = 10
        otableAthi4.Cell(1, 5).Range.Bold = True
        otableAthi4.Cell(1, 5).Range.Underline = False
        otableAthi4.Cell(1, 5).Range.Italic = True
        otableAthi4.Cell(1, 5).Shading.BackgroundPatternColor = RGB(12, 12, 12)

        'Dim filename As Object = Form5.TextBox1.Text


        'otableAthi4.Cell(2, 5).Range.InsertFile(filename, Attachment:=True)
        a_zip = Form5.TextBox1.Text & "\A10_JE_PREP.zip"
        b_zip = Form5.TextBox1.Text & "\A20_TB_PREP.zip"
        c_zip = Form5.TextBox1.Text & "\A30_MAIN.zip"
        d_zip = Form5.TextBox1.Text & "\B10_VALIDATION.zip"
        e_zip = Form5.TextBox1.Text & "\C10_ROLL.zip"
        f_zip = Form5.TextBox1.Text & "\D10_GLOBAL_JE_MAPPING.zip"
        g_zip = Form5.TextBox1.Text & "\E10_UNBALANCED_ENTRIES.zip"

        otableAthi4.Cell(2, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=a_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="A10_JE_PREP")
        otableAthi4.Cell(3, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=b_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="A20_JE_PREP")
        otableAthi4.Cell(4, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=c_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="A30_MAIN")
        otableAthi4.Cell(5, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=d_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="B10_VALIDATION")
        otableAthi4.Cell(6, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=e_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="C10_ROLL")
        otableAthi4.Cell(7, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=f_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="D10_GLOBAL_JE_MAPPING")

        If (unbalancedflag = 0) Then
            otableAthi4.Cell(8, 5).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=g_zip, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:="E10_UNBALANCED_ENTRIES")
        End If

        otableAthi4.Cell(2, 5).Range.Font.Underline = False
        otableAthi4.Cell(3, 5).Range.Font.Underline = False
        otableAthi4.Cell(4, 5).Range.Font.Underline = False
        otableAthi4.Cell(5, 5).Range.Font.Underline = False
        otableAthi4.Cell(6, 5).Range.Font.Underline = False
        otableAthi4.Cell(7, 5).Range.Font.Underline = False

        If (unbalancedflag = 0) Then
            otableAthi4.Cell(8, 5).Range.Font.Underline = False
        End If


        otableAthi4.Cell(9 - unbalancedflag, 1).Range.Text = "ACL File:"
        otableAthi4.Cell(9 - unbalancedflag, 1).Range.Font.Name = "Times New Roman"
        otableAthi4.Cell(9 - unbalancedflag, 1).Range.Font.Size = 10
        otableAthi4.Cell(9 - unbalancedflag, 1).Range.Bold = True
        otableAthi4.Cell(9 - unbalancedflag, 1).Range.Underline = False
        otableAthi4.Cell(9 - unbalancedflag, 1).Range.Italic = True
        otableAthi4.Cell(9 - unbalancedflag, 1).Shading.BackgroundPatternColor = RGB(12, 12, 12)


        ACL_NAME = Form5.TextBox6.Text.Substring(0, Len(Form5.TextBox6.Text) - 4) & ".zip"
        ACL_LOG_NAME = "C:\Users\Public\Documents\SametimeFileTransfer\Logs\Pan_Handle2_Oil_GAS_10012012_th" & ".LOG"
        otableAthi4.Cell(10 - unbalancedflag, 1).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=ACL_NAME, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:=myclientname & " " & myPOA & " ACL ")
        otableAthi4.Cell(10 - unbalancedflag, 1).Range.InlineShapes.AddOLEObject(ClassType:="Package", FileName:=ACL_LOG_NAME, DisplayAsIcon:=True, IconFileName:="C:\WINDOWS\system32\packager.dll", IconIndex:=0, IconLabel:=myclientname & " " & myPOA & " LOG")
        otableAthi4.Cell(10 - unbalancedflag, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otableAthi4.Cell(10 - unbalancedflag, 1).Range.Underline = False



        'Journal Entry Table

        Dim udcount As Integer = 3   'to extract the number of user defined feilds

        Dim oParaAthi1 As Word.Paragraph


        oParaAthi1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oParaAthi1.Range.InsertParagraphAfter()
        oParaAthi1.Range.Text = "Global Tool Field Mapping: "
        oParaAthi1.Range.Font.Name = "Times New Roman"
        oParaAthi1.Range.Font.Size = 10
        oParaAthi1.Format.SpaceAfter = 0
        oParaAthi1.Range.Font.Bold = True
        oParaAthi1.Range.Font.Underline = True
        oParaAthi1.Range.Font.Italic = False


        Dim otableAthi1 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 12 + udcount, 3)

        otableAthi1.AllowAutoFit = True

        otableAthi1.Borders.Enable = True

        otableAthi1.Cell(1, 1).Merge(otableAthi1.Cell(1, 3))
        otableAthi1.Cell(1, 1).Range.Text = "JE Transaction Files"
        otableAthi1.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(1, 1).Range.Font.Size = 10
        otableAthi1.Cell(1, 1).Range.Bold = True
        otableAthi1.Cell(1, 1).Range.Underline = False
        otableAthi1.Cell(1, 1).Range.Italic = True
        otableAthi1.Cell(1, 1).Shading.BackgroundPatternColor = RGB(0, 0, 0)
        otableAthi1.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        otableAthi1.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0



        otableAthi1.Cell(2, 1).Range.Text = "EY/Global Analytics Field Name"
        otableAthi1.Cell(2, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(2, 1).Range.Font.Size = 10
        otableAthi1.Cell(2, 1).Range.Bold = True
        otableAthi1.Cell(2, 1).Range.Underline = False
        otableAthi1.Cell(2, 1).Range.Italic = True
        otableAthi1.Cell(2, 1).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        otableAthi1.Cell(2, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

        otableAthi1.Cell(2, 2).Range.Text = "ACL Field Name"
        otableAthi1.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(2, 2).Range.Font.Size = 10
        otableAthi1.Cell(2, 2).Range.Bold = True
        otableAthi1.Cell(2, 2).Range.Underline = False
        otableAthi1.Cell(2, 2).Range.Italic = True
        otableAthi1.Cell(2, 2).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        otableAthi1.Cell(2, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

        otableAthi1.Cell(2, 3).Range.Text = "Client Data Field Name"
        otableAthi1.Cell(2, 3).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(2, 3).Range.Font.Size = 10
        otableAthi1.Cell(2, 3).Range.Bold = True
        otableAthi1.Cell(2, 3).Range.Underline = False
        otableAthi1.Cell(2, 3).Range.Italic = True
        otableAthi1.Cell(2, 3).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        otableAthi1.Cell(2, 3).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

        otableAthi1.Cell(3, 1).Range.Text = "Journal Entry Number"
        otableAthi1.Cell(3, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(3, 1).Range.Font.Size = 10
        otableAthi1.Cell(3, 1).Range.Bold = True
        otableAthi1.Cell(3, 1).Range.Underline = False
        otableAthi1.Cell(3, 1).Range.Italic = False

        otableAthi1.Cell(3, 2).Range.Text = "EY_JENum"
        otableAthi1.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(3, 2).Range.Font.Size = 10
        otableAthi1.Cell(3, 2).Range.Bold = False
        otableAthi1.Cell(3, 2).Range.Underline = False
        otableAthi1.Cell(3, 2).Range.Italic = False

        otableAthi1.Cell(4, 1).Range.Text = "General Ledger Account Number"
        otableAthi1.Cell(4, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(4, 1).Range.Font.Size = 10
        otableAthi1.Cell(4, 1).Range.Bold = True
        otableAthi1.Cell(4, 1).Range.Underline = False
        otableAthi1.Cell(4, 1).Range.Italic = False

        otableAthi1.Cell(4, 2).Range.Text = "EY_Acct"
        otableAthi1.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(4, 2).Range.Font.Size = 10
        otableAthi1.Cell(4, 2).Range.Bold = False
        otableAthi1.Cell(4, 2).Range.Underline = False
        otableAthi1.Cell(4, 2).Range.Italic = False

        otableAthi1.Cell(5, 1).Range.Text = "Amount"
        otableAthi1.Cell(5, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(5, 1).Range.Font.Size = 10
        otableAthi1.Cell(5, 1).Range.Bold = True
        otableAthi1.Cell(5, 1).Range.Underline = False
        otableAthi1.Cell(5, 1).Range.Italic = False

        otableAthi1.Cell(5, 2).Range.Text = "EY_Amount"
        otableAthi1.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(5, 2).Range.Font.Size = 10
        otableAthi1.Cell(5, 2).Range.Bold = False
        otableAthi1.Cell(5, 2).Range.Underline = False
        otableAthi1.Cell(5, 2).Range.Italic = False

        otableAthi1.Cell(6, 1).Range.Text = "Business Unit"
        otableAthi1.Cell(6, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(6, 1).Range.Font.Size = 10
        otableAthi1.Cell(6, 1).Range.Bold = True
        otableAthi1.Cell(6, 1).Range.Underline = False
        otableAthi1.Cell(6, 1).Range.Italic = False

        otableAthi1.Cell(6, 2).Range.Text = "EY_BusUnit"
        otableAthi1.Cell(6, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(6, 2).Range.Font.Size = 10
        otableAthi1.Cell(6, 2).Range.Bold = False
        otableAthi1.Cell(6, 2).Range.Underline = False
        otableAthi1.Cell(6, 2).Range.Italic = False

        otableAthi1.Cell(7, 1).Range.Text = "Effective Date"
        otableAthi1.Cell(7, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(7, 1).Range.Font.Size = 10
        otableAthi1.Cell(7, 1).Range.Bold = True
        otableAthi1.Cell(7, 1).Range.Underline = False
        otableAthi1.Cell(7, 1).Range.Italic = False

        otableAthi1.Cell(7, 2).Range.Text = "EY_EffectiveDt"
        otableAthi1.Cell(7, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(7, 2).Range.Font.Size = 10
        otableAthi1.Cell(7, 2).Range.Bold = False
        otableAthi1.Cell(7, 2).Range.Underline = False
        otableAthi1.Cell(7, 2).Range.Italic = False

        otableAthi1.Cell(8, 1).Range.Text = "Entry Date"
        otableAthi1.Cell(8, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(8, 1).Range.Font.Size = 10
        otableAthi1.Cell(8, 1).Range.Bold = True
        otableAthi1.Cell(8, 1).Range.Underline = False
        otableAthi1.Cell(8, 1).Range.Italic = False

        otableAthi1.Cell(8, 2).Range.Text = "EY_EntryDt"
        otableAthi1.Cell(8, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(8, 2).Range.Font.Size = 10
        otableAthi1.Cell(8, 2).Range.Bold = False
        otableAthi1.Cell(8, 2).Range.Underline = False
        otableAthi1.Cell(8, 2).Range.Italic = False

        otableAthi1.Cell(9, 1).Range.Text = "Period"
        otableAthi1.Cell(9, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(9, 1).Range.Font.Size = 10
        otableAthi1.Cell(9, 1).Range.Bold = True
        otableAthi1.Cell(9, 1).Range.Underline = False
        otableAthi1.Cell(9, 1).Range.Italic = False

        otableAthi1.Cell(9, 2).Range.Text = "EY_Period"
        otableAthi1.Cell(9, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(9, 2).Range.Font.Size = 10
        otableAthi1.Cell(9, 2).Range.Bold = False
        otableAthi1.Cell(9, 2).Range.Underline = False
        otableAthi1.Cell(9, 2).Range.Italic = False

        otableAthi1.Cell(10, 1).Range.Text = "Preparer ID"
        otableAthi1.Cell(10, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(10, 1).Range.Font.Size = 10
        otableAthi1.Cell(10, 1).Range.Bold = True
        otableAthi1.Cell(10, 1).Range.Underline = False
        otableAthi1.Cell(10, 1).Range.Italic = False

        otableAthi1.Cell(10, 2).Range.Text = "EY_PreparerID"
        otableAthi1.Cell(10, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(10, 2).Range.Font.Size = 10
        otableAthi1.Cell(10, 2).Range.Bold = False
        otableAthi1.Cell(10, 2).Range.Underline = False
        otableAthi1.Cell(10, 2).Range.Italic = False

        otableAthi1.Cell(11, 1).Range.Text = "Source"
        otableAthi1.Cell(11, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(11, 1).Range.Font.Size = 10
        otableAthi1.Cell(11, 1).Range.Bold = True
        otableAthi1.Cell(11, 1).Range.Underline = False
        otableAthi1.Cell(11, 1).Range.Italic = False

        otableAthi1.Cell(11, 2).Range.Text = "EY_Source"
        otableAthi1.Cell(11, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(11, 2).Range.Font.Size = 10
        otableAthi1.Cell(11, 2).Range.Bold = False
        otableAthi1.Cell(11, 2).Range.Underline = False
        otableAthi1.Cell(11, 2).Range.Italic = False

        otableAthi1.Cell(12, 1).Range.Text = "Journal Entry /Transaction Description"
        otableAthi1.Cell(12, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(12, 1).Range.Font.Size = 10
        otableAthi1.Cell(12, 1).Range.Bold = True
        otableAthi1.Cell(12, 1).Range.Underline = False
        otableAthi1.Cell(12, 1).Range.Italic = False

        otableAthi1.Cell(12, 2).Range.Text = "EY_JE_Desc"
        otableAthi1.Cell(12, 2).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(12, 2).Range.Font.Size = 10
        otableAthi1.Cell(12, 2).Range.Bold = False
        otableAthi1.Cell(12, 2).Range.Underline = False
        otableAthi1.Cell(12, 2).Range.Italic = False




        oParaAthi1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oParaAthi1.Range.InsertParagraphAfter()

        'Trial Balance Table

        Dim otableAthi2 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)

        otableAthi2.Borders.Enable = True

        otableAthi2.Borders.Enable = True
        otableAthi2.AllowAutoFit = True
        otableAthi2.Columns.Item(1).Width = oWord.CentimetersToPoints(6.6)
        otableAthi2.Columns.Item(2).Width = oWord.CentimetersToPoints(3.81)
        otableAthi2.Columns.Item(3).Width = oWord.CentimetersToPoints(7.41)
        otableAthi2.Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly
        otableAthi2.Rows.Height = oWord.CentimetersToPoints(0.6)


        otableAthi2.Cell(1, 1).Merge(otableAthi2.Cell(1, 3))
        otableAthi2.Cell(1, 1).Range.Text = "Trial Balance Files"
        otableAthi2.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(1, 1).Range.Font.Size = 10
        otableAthi2.Cell(1, 1).Range.Bold = True
        otableAthi2.Cell(1, 1).Range.Underline = False
        otableAthi2.Cell(1, 1).Range.Italic = True
        otableAthi2.Cell(1, 1).Shading.BackgroundPatternColor = RGB(0, 0, 0)
        otableAthi2.Cell(1, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        otableAthi2.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0




        otableAthi2.Cell(2, 1).Range.Text = "EY/Global Analytics Field Name"
        otableAthi2.Cell(2, 1).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(2, 1).Range.Font.Size = 10
        otableAthi2.Cell(2, 1).Range.Bold = True
        otableAthi2.Cell(2, 1).Range.Underline = False
        otableAthi2.Cell(2, 1).Range.Italic = True
        otableAthi2.Cell(2, 1).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        otableAthi2.Cell(2, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

        otableAthi2.Cell(2, 2).Range.Text = "ACL Field Name"
        otableAthi2.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(2, 2).Range.Font.Size = 10
        otableAthi2.Cell(2, 2).Range.Bold = True
        otableAthi2.Cell(2, 2).Range.Underline = False
        otableAthi2.Cell(2, 2).Range.Italic = True
        otableAthi2.Cell(2, 2).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        otableAthi2.Cell(2, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

        otableAthi2.Cell(2, 3).Range.Text = "Client Data Field Name"
        otableAthi2.Cell(2, 3).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(2, 3).Range.Font.Size = 10
        otableAthi2.Cell(2, 3).Range.Bold = True
        otableAthi2.Cell(2, 3).Range.Underline = False
        otableAthi2.Cell(2, 3).Range.Italic = True
        otableAthi2.Cell(2, 3).Shading.BackgroundPatternColor = RGB(224, 224, 224)
        otableAthi2.Cell(2, 3).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

        otableAthi2.Cell(3, 1).Range.Text = "General Ledger Account Number"
        otableAthi2.Cell(3, 1).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(3, 1).Range.Font.Size = 10
        otableAthi2.Cell(3, 1).Range.Bold = True
        otableAthi2.Cell(3, 1).Range.Underline = False
        otableAthi2.Cell(3, 1).Range.Italic = False

        otableAthi2.Cell(3, 2).Range.Text = "EY_Acct"
        otableAthi2.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(3, 2).Range.Font.Size = 10
        otableAthi2.Cell(3, 2).Range.Bold = False
        otableAthi2.Cell(3, 2).Range.Underline = False
        otableAthi2.Cell(3, 2).Range.Italic = False

        otableAthi2.Cell(4, 1).Range.Text = "General Ledger Account Name"
        otableAthi2.Cell(4, 1).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(4, 1).Range.Font.Size = 10
        otableAthi2.Cell(4, 1).Range.Bold = True
        otableAthi2.Cell(4, 1).Range.Underline = False
        otableAthi2.Cell(4, 1).Range.Italic = False

        otableAthi2.Cell(4, 2).Range.Text = "EY_AcctName"
        otableAthi2.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(4, 2).Range.Font.Size = 10
        otableAthi2.Cell(4, 2).Range.Bold = False
        otableAthi2.Cell(4, 2).Range.Underline = False
        otableAthi2.Cell(4, 2).Range.Italic = False

        otableAthi2.Cell(5, 1).Range.Text = "Account Type"
        otableAthi2.Cell(5, 1).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(5, 1).Range.Font.Size = 10
        otableAthi2.Cell(5, 1).Range.Bold = True
        otableAthi2.Cell(5, 1).Range.Underline = False
        otableAthi2.Cell(5, 1).Range.Italic = False

        otableAthi2.Cell(5, 2).Range.Text = "EY_Accttype"
        otableAthi2.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(5, 2).Range.Font.Size = 10
        otableAthi2.Cell(5, 2).Range.Bold = False
        otableAthi2.Cell(5, 2).Range.Underline = False
        otableAthi2.Cell(5, 2).Range.Italic = False

        otableAthi2.Cell(6, 1).Range.Text = "Account Class"
        otableAthi2.Cell(6, 1).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(6, 1).Range.Font.Size = 10
        otableAthi2.Cell(6, 1).Range.Bold = True
        otableAthi2.Cell(6, 1).Range.Underline = False
        otableAthi2.Cell(6, 1).Range.Italic = False

        otableAthi2.Cell(6, 2).Range.Text = "EY_AcctClass"
        otableAthi2.Cell(6, 2).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(6, 2).Range.Font.Size = 10
        otableAthi2.Cell(6, 2).Range.Bold = False
        otableAthi2.Cell(6, 2).Range.Underline = False
        otableAthi2.Cell(6, 2).Range.Italic = False

        otableAthi2.Cell(7, 1).Range.Text = "Beginning Balance"
        otableAthi2.Cell(7, 1).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(7, 1).Range.Font.Size = 10
        otableAthi2.Cell(7, 1).Range.Bold = True
        otableAthi2.Cell(7, 1).Range.Underline = False
        otableAthi2.Cell(7, 1).Range.Italic = False

        otableAthi2.Cell(7, 2).Range.Text = "EY_BegBal"
        otableAthi2.Cell(7, 2).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(7, 2).Range.Font.Size = 10
        otableAthi2.Cell(7, 2).Range.Bold = False
        otableAthi2.Cell(7, 2).Range.Underline = False
        otableAthi2.Cell(7, 2).Range.Italic = False

        otableAthi2.Cell(8, 1).Range.Text = "Ending Balance"
        otableAthi2.Cell(8, 1).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(8, 1).Range.Font.Size = 10
        otableAthi2.Cell(8, 1).Range.Bold = True
        otableAthi2.Cell(8, 1).Range.Underline = False
        otableAthi2.Cell(8, 1).Range.Italic = False

        otableAthi2.Cell(8, 2).Range.Text = "EY_EndBal"
        otableAthi2.Cell(8, 2).Range.Font.Name = "Times New Roman"
        otableAthi2.Cell(8, 2).Range.Font.Size = 10
        otableAthi2.Cell(8, 2).Range.Bold = False
        otableAthi2.Cell(8, 2).Range.Underline = False
        otableAthi2.Cell(8, 2).Range.Italic = False


        'Buisness Rules
        Dim oParaAthi2 As Word.Paragraph

        oParaAthi2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oParaAthi2.Range.InsertParagraphAfter()
        oParaAthi2.Range.Text = "Buisness Rules: "
        oParaAthi2.Range.Font.Name = "Times New Roman"
        oParaAthi2.Range.Font.Size = 10
        oParaAthi2.Format.SpaceAfter = 0
        oParaAthi2.Range.Font.Bold = True
        oParaAthi2.Range.Font.Underline = True
        oParaAthi2.Range.Font.Italic = False

        'Box Table
        Dim otableAthi3 As Word.Table
        rng = oDoc.Bookmarks.Item("\endofdoc").Range
        otableAthi3 = oDoc.Tables.Add(Range:=rng, NumRows:=6, NumColumns:=1)
        otableAthi3.Borders.Enable = True
        otableAthi3.AllowAutoFit = True
        otableAthi3.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone
        otableAthi3.Columns.Width = oWord.CentimetersToPoints(17.8)

        otableAthi3.Cell(1, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(1, 1).Range.Paragraphs(1).Range.Text = "1. Identify and order journal entry fields to arrive at a unique journal entry"
        otableAthi3.Cell(1, 1).Range.Paragraphs(1).Range.Font.Bold = False
        otableAthi3.Cell(1, 1).Range.Paragraphs(1).Range.Font.Underline = False
        otableAthi3.Cell(1, 1).Range.Paragraphs(1).Range.Font.Italic = False


        otableAthi3.Cell(1, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(1, 1).Range.Paragraphs(2).Range.Text = "   • Journal Entry Number - EY_JENum - Field_1 (Mostly)" & vbNewLine & "   • Field_Name - EY_Field_Name - Field_X (Rarely)"
        otableAthi3.Cell(1, 1).Range.Paragraphs(2).Range.Font.Bold = False
        otableAthi3.Cell(1, 1).Range.Paragraphs(2).Range.Font.ColorIndex = Word.WdColorIndex.wdRed
        otableAthi3.Cell(1, 1).Range.Paragraphs(2).Range.Font.Underline = False
        otableAthi3.Cell(1, 1).Range.Paragraphs(2).Range.Font.Italic = False


        otableAthi3.Cell(1, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(1, 1).Range.Paragraphs(3).Range.Text = "2.	Account Type Definition"
        otableAthi3.Cell(1, 1).Range.Paragraphs(3).Range.Font.Bold = True
        otableAthi3.Cell(1, 1).Range.Paragraphs(3).Range.Font.Underline = False
        otableAthi3.Cell(1, 1).Range.Paragraphs(3).Range.Font.Italic = False

        Dim AccountNumberChoice As Integer = 1
        Dim newdoc As New Word.Document
        newdoc = oWord.Documents.Add
        'oPara8 = newdoc.Content.Paragraphs.Add

        otable11 = newdoc.Tables.Add(newdoc.Bookmarks.Item("\endofdoc").Range, 6, 2)
        otable11.Borders.Enable = True
        otable11.Columns.Item(1).Width = oWord.CentimetersToPoints(1.48)
        otable11.Columns.Item(2).Width = oWord.CentimetersToPoints(14.23)
        otable11.Rows.Height = oWord.CentimetersToPoints(0.11)
        otable11.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
        otable11.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter


        otable11.Cell(1, 1).Range.Text = "GL Account Number beginning with"
        otable11.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 1).Range.Font.Size = 10
        otable11.Cell(1, 1).Range.Bold = True
        otable11.Cell(1, 1).Range.Underline = False
        otable11.Cell(1, 1).Shading.BackgroundPatternColor = RGB(192, 192, 192)

        otable11.Cell(1, 2).Range.Text = "Account Type"
        otable11.Cell(1, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(1, 2).Range.Font.Size = 10
        otable11.Cell(1, 2).Range.Bold = True
        otable11.Cell(1, 2).Range.Underline = False
        otable11.Cell(1, 2).Shading.BackgroundPatternColor = RGB(192, 192, 192)


        otable11.Cell(2, 1).Range.Text = If(AccountNumberChoice = 1, "1", "<20000")
        otable11.Cell(2, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(2, 1).Range.Font.Size = 10
        otable11.Cell(2, 1).Range.Bold = False
        otable11.Cell(2, 1).Range.Underline = False

        otable11.Cell(2, 2).Range.Text = "Assets"
        otable11.Cell(2, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(2, 2).Range.Font.Size = 10
        otable11.Cell(2, 2).Range.Bold = False
        otable11.Cell(2, 2).Range.Underline = False

        otable11.Cell(3, 1).Range.Text = If(AccountNumberChoice = 1, "2", ">20000 and <30000")
        otable11.Cell(3, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(3, 1).Range.Font.Size = 10
        otable11.Cell(3, 1).Range.Bold = False
        otable11.Cell(3, 1).Range.Underline = False

        otable11.Cell(3, 2).Range.Text = "Liabilities"
        otable11.Cell(3, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(3, 2).Range.Font.Size = 10
        otable11.Cell(3, 2).Range.Bold = False
        otable11.Cell(3, 2).Range.Underline = False

        otable11.Cell(4, 1).Range.Text = If(AccountNumberChoice = 1, "3", ">30000 and <40000")
        otable11.Cell(4, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(4, 1).Range.Font.Size = 10
        otable11.Cell(4, 1).Range.Bold = False
        otable11.Cell(4, 1).Range.Underline = False

        otable11.Cell(4, 2).Range.Text = "Equity"
        otable11.Cell(4, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(4, 2).Range.Font.Size = 10
        otable11.Cell(4, 2).Range.Bold = False
        otable11.Cell(4, 2).Range.Underline = False

        otable11.Cell(5, 1).Range.Text = If(AccountNumberChoice = 1, "4", ">40000 and <50000")
        otable11.Cell(5, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(5, 1).Range.Font.Size = 10
        otable11.Cell(5, 1).Range.Bold = False
        otable11.Cell(5, 1).Range.Underline = False

        otable11.Cell(5, 2).Range.Text = "Revenue"
        otable11.Cell(5, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(5, 2).Range.Font.Size = 10
        otable11.Cell(5, 2).Range.Bold = False
        otable11.Cell(5, 2).Range.Underline = False

        otable11.Cell(6, 1).Range.Text = If(AccountNumberChoice = 1, "5,6,7,8 or 9", ">50000")
        otable11.Cell(6, 1).Range.Font.Name = "Times New Roman"
        otable11.Cell(6, 1).Range.Font.Size = 10
        otable11.Cell(6, 1).Range.Bold = False
        otable11.Cell(6, 1).Range.Underline = False

        otable11.Cell(6, 2).Range.Text = "Expenses"
        otable11.Cell(6, 2).Range.Font.Name = "Times New Roman"
        otable11.Cell(6, 2).Range.Font.Size = 10
        otable11.Cell(6, 2).Range.Bold = False
        otable11.Cell(6, 2).Range.Underline = False


        newdoc.ActiveWindow.Selection.WholeStory()
        newdoc.ActiveWindow.Selection.Copy()
        otableAthi3.Cell(2, 1).Range.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)
        newdoc.SaveAs2("c:\temp\test.doc")
        newdoc.Close()

        Dim exclusionchoice_ui As Integer = 1

        Dim dynind As Integer = 3 + exclusionchoice_ui

        otableAthi3.Cell(3, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(3, 1).Range.Paragraphs(1).Range.Text = If(exclusionchoice_ui = 1, "3. Exclusions", "")
        otableAthi3.Cell(3, 1).Range.Paragraphs(1).Range.Font.Bold = True
        otableAthi3.Cell(3, 1).Range.Paragraphs(1).Range.Font.Underline = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(1).Range.Font.Italic = False

        otableAthi3.Cell(3, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(3, 1).Range.Paragraphs(2).Range.Text = dynind.ToString & ". System/Manual Identification"
        otableAthi3.Cell(3, 1).Range.Paragraphs(2).Range.Font.Bold = True
        otableAthi3.Cell(3, 1).Range.Paragraphs(2).Range.Font.Underline = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(2).Range.Font.Italic = False

        Dim sysman_ui As Integer = 1
        Dim optstr As String = " "

        If (sysman_ui = 1) Then
            optstr = "   • Manual Entries: All entries where Source/Preparer ID/Field EQUALS ""Phrase""/""Condition""." & vbNewLine & "   • System Entries: All other entries."
        Else
            optstr = "   • All entries were marked as ""Manual""."
        End If

        otableAthi3.Cell(3, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(3, 1).Range.Paragraphs(3).Range.Text = optstr
        otableAthi3.Cell(3, 1).Range.Paragraphs(3).Range.Font.Bold = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(3).Range.Font.Underline = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(3).Range.Font.Italic = False

        dynind = dynind + 1

        Dim point5_choice_ui As Integer = 1

        otableAthi3.Cell(3, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(3, 1).Range.Paragraphs(4).Range.Text = dynind.ToString & ". Inter-Company Accounts"
        otableAthi3.Cell(3, 1).Range.Paragraphs(4).Range.Font.Bold = True
        otableAthi3.Cell(3, 1).Range.Paragraphs(4).Range.Font.Underline = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(4).Range.Font.Italic = False

        Dim optstr1 As String = " "

        If (point5_choice_ui = 1) Then
            optstr1 = "   • N/A"
        ElseIf (point5_choice_ui = 2) Then
            optstr1 = "   • Journal Entry Description CONTAINS: " & vbNewLine & "      • Phrase 1 " & vbNewLine & "      • Phrase 2 "
        ElseIf (point5_choice_ui = 3) Then
            optstr1 = "   • GL Account Number EQUALS: "
        End If

        otableAthi3.Cell(3, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(3, 1).Range.Paragraphs(5).Range.Text = optstr1
        otableAthi3.Cell(3, 1).Range.Paragraphs(5).Range.Font.Bold = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(5).Range.Font.Underline = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(5).Range.Font.Italic = False

        dynind = dynind + 1

        otableAthi3.Cell(3, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(3, 1).Range.Paragraphs(6).Range.Text = dynind.ToString & ". Related Party Accounts"
        otableAthi3.Cell(3, 1).Range.Paragraphs(6).Range.Font.Bold = True
        otableAthi3.Cell(3, 1).Range.Paragraphs(6).Range.Font.Underline = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(6).Range.Font.Italic = False

        dynind = dynind + 1

        otableAthi3.Cell(3, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(3, 1).Range.Paragraphs(7).Range.Text = dynind.ToString & ". Professional Fee Accounts"
        otableAthi3.Cell(3, 1).Range.Paragraphs(7).Range.Font.Bold = True
        otableAthi3.Cell(3, 1).Range.Paragraphs(7).Range.Font.Underline = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(7).Range.Font.Italic = False

        dynind = dynind + 1


        otableAthi3.Cell(3, 1).Range.InsertParagraphAfter()
        otableAthi3.Cell(3, 1).Range.Paragraphs(8).Range.Text = dynind.ToString & ". Report Thresholds and Other Parameters"
        otableAthi3.Cell(3, 1).Range.Paragraphs(8).Range.Font.Bold = True
        otableAthi3.Cell(3, 1).Range.Paragraphs(8).Range.Font.Underline = False
        otableAthi3.Cell(3, 1).Range.Paragraphs(8).Range.Font.Italic = False





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
