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

        Dim STrA10 As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\A10_JE_PREP.LOG")
        Dim STrA20 As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\A20_TB_PREP.LOG")
        Dim STrA30 As IO.TextReader = System.IO.File.OpenText(Form5.TextBox1.Text & "\C_WORKLOG.LOG")
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
