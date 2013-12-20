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
        temp = TRC.Substring(START_INDEX1, LEN1)
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
        oPara2.Range.Text = "To perform journal entry analysis for " & myclientname & " for the current period effective " & START_POA & "through" & end_poa
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
            otable1.Cell(7, 2).Range.Text = Form6.drd.Value.ToString() & " " & Form6.dorv.Value.ToString() & " (Date of Re-validations)  ," & Form6.dop.Value.ToString()
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
        oPara3.Range.Font.Underline = True
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
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Text = "This memorandum and supporting JE CAAT file were prepared by the EY GTH Team for use by the audit team. The memorandum documents the objectives of the work, planned procedures, procedures executed, and our assessment of the client data. This memorandum is intended to guide and assist the audit team in performing the journal entry analysis procedures and should not be considered a standalone work paper. We have provided this memorandum in softcopy so that the audit teams may copy those portions that are deemed relevant to their audit for inclusion in the final work papers. "
        otable2.Cell(1, 1).Range.Paragraphs(1).Format.SpaceAfter = 0
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Font.Name = "Times New Roman"
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Font.Size = 11
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Bold = False
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Underline = False
        otable2.Cell(1, 1).Range.Paragraphs(1).Range.Italic = False
        otable2.Cell(1, 1).Range.Paragraphs(1).Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft






















































































































































































































































































































































































        'Journal Entry Table

        Dim udcount As Integer =        'to extract the number of user defined feilds
        Dim oParaAthi1 As Word.Paragraph


        oParaAthi1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oParaAthi1.Range.InsertParagraphAfter()
        oParaAthi1.Range.Text = "Global Tool Field Mapping:"
        oParaAthi1.Range.Font.Name = "Times New Roman"
        oParaAthi1.Range.Font.Size = 10
        oParaAthi1.Format.SpaceAfter = 0
        oParaAthi1.Range.Font.Bold = True
        oParaAthi1.Range.Font.Underline = True
        oParaAthi1.Range.Font.Italic = False

        Dim otableAthi1 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 15 + udCount, 3)
        otableAthi1.Borders.Enable = True

        otableAthi1.Cell(1, 1).Range.Text = "JE Transaction Files"
        otableAthi1.Cell(1, 1).Range.Font.Name = "Times New Roman"
        otableAthi1.Cell(1, 1).Range.Font.Size = 10
        otableAthi1.Cell(1, 1).Range.Bold = True
        otableAthi1.Cell(1, 1).Range.Underline = False
        otableAthi1.Cell(1, 1).Range.Italic = True
        otableAthi1.Cell(1, 1).Shading.BackgroundPatternColor = RGB(0, 0, 0)
        otableAthi1.Cell(1, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
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


        'Trial Balance Table


        Dim oParaAthi2 As Word.Paragraph


        oParaAthi2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oParaAthi2.Range.InsertParagraphAfter()
        oParaAthi2.Range.Text = "Global Tool Field Mapping:"
        oParaAthi2.Range.Font.Name = "Times New Roman"
        oParaAthi2.Range.Font.Size = 10
        oParaAthi2.Format.SpaceAfter = 0
        oParaAthi2.Range.Font.Bold = True
        oParaAthi2.Range.Font.Underline = True
        oParaAthi2.Range.Font.Italic = False

        Dim otableAthi2 As Word.Table = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 15 + udCount, 3)
        otableAthi1.Borders.Enable = True



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
