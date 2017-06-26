Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop
''' <summary>
''' Test zum Erstellen eines WORD-Dokuments
''' <param name="sender">Typ Object</param> 
''' <param name="e">as EventArgs</param> 
''' </summary>
''' <remarks></remarks>
Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        fWord()
    End Sub
    ''' <summary>
    ''' Die Funktion wird vom Form1 mit Button1 gestartet
    ''' </summary>
    ''' <returns>true oder false</returns>
    ''' <remarks></remarks>
    Public Function fWord() As Boolean
        Dim WordApp As Word.Application
        Dim WordDatei As Word.Document
        Try
            WordApp = New Word.Application()
            WordDatei = WordApp.Documents.Open("C:\tmp\testdok2.doc")
            WordDatei.MailMerge.MainDocumentType = Word.WdMailMergeMainDocType.wdFormLetters
            WordDatei.MailMerge.OpenDataSource(Name:= _
                "C:\tmp\testcsv.csv", _
                ConfirmConversions:=False, ReadOnly:=False, LinkToSource:=True, _
                AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:="", _
                WritePasswordDocument:="", WritePasswordTemplate:="", Revert:=False, _
                Format:=Word.WdOpenFormat.wdOpenFormatAuto, Connection:="", SQLStatement:="", SQLStatement1 _
                :="", SubType:=Word.WdMergeSubType.wdMergeSubTypeOther)
            With WordDatei.MailMerge
                '.Destination = Word.WdMailMergeDestination.wdSendToNewDocument
                .Destination = Word.WdMailMergeDestination.wdSendToEmail
                .MailFormat = Word.WdMailMergeMailFormat.wdMailFormatHTML
                .MailSubject = "Testmail "
                .MailAddressFieldName = "email"
                .SuppressBlankLines = True
                With .DataSource
                    .FirstRecord = Word.WdMailMergeDefaultRecord.wdDefaultFirstRecord
                    .LastRecord = Word.WdMailMergeDefaultRecord.wdDefaultLastRecord
                End With
                .Execute(Pause:=False)
            End With

            WordDatei.Close()
            WordApp.Visible = True
            WordApp.Quit()

            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////
    ' Dokument drucken, muss noch angepasst werden
    Public Function pWord() As Boolean
        Dim WordApp As Word.Application
        Dim WordDatei As Word.Document
        Try
            WordApp = New Word.Application()
            WordDatei = WordApp.Documents.Open("C:\tmp\testdok.doc")
            WordDatei.MailMerge.MainDocumentType = Word.WdMailMergeMainDocType.wdFormLetters
            WordDatei.MailMerge.OpenDataSource(Name:= _
                "C:\tmp\testcsv.csv", _
                ConfirmConversions:=False, ReadOnly:=False, LinkToSource:=True, _
                AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:="", _
                WritePasswordDocument:="", WritePasswordTemplate:="", Revert:=False, _
                Format:=Word.WdOpenFormat.wdOpenFormatAuto, Connection:="", SQLStatement:="", SQLStatement1 _
                :="", SubType:=Word.WdMergeSubType.wdMergeSubTypeOther)
            With WordDatei.MailMerge
                '.Destination = Word.WdMailMergeDestination.wdSendToNewDocument
                .Destination = Word.WdMailMergeDestination.wdSendToEmail
                .MailFormat = Word.WdMailMergeMailFormat.wdMailFormatHTML
                .MailSubject = "Testmail "
                .MailAddressFieldName = "empfänger"
                .SuppressBlankLines = True
                With .DataSource
                    .FirstRecord = Word.WdMailMergeDefaultRecord.wdDefaultFirstRecord
                    .LastRecord = Word.WdMailMergeDefaultRecord.wdDefaultLastRecord
                End With
                .Execute(Pause:=False)
            End With

            WordDatei.Close()
            WordApp.Visible = True
            WordApp.Quit()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary>
    ''' Beendet das Programm
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class
