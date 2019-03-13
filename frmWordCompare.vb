Imports System
Imports System.IO
Imports System.Text


Public Class frmWordCompare

    Public boolCancel As Boolean = True

    Private Sub frmWordCompare_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        'now get template
        Dim boolT As Boolean = ApplyReportTemplate(Me, "Word Compare")

        Call Dev()

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        Call Execute()
        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Sub Dev()

        'Dim doc1 As String
        'Dim doc2 As String
        'Dim strPath As String

        'doc1 = "D:\LabIntegrity\LABIntegrityApps\StudyDoc\Administration\WordCompare\Doc01.docx"
        'doc2 = "D:\LabIntegrity\LABIntegrityApps\StudyDoc\Administration\WordCompare\Doc02.docx"

        'Me.txtDoc1.Text = doc1
        'Me.txtDoc2.Text = doc2


    End Sub

    Sub Execute()

        Cursor.Current = Cursors.WaitCursor

        Me.Visible = False

        Dim wd As New Microsoft.Office.Interop.Word.Application

        Dim wdVer As String
        Dim intVer As Short
        Dim strM As String
        Dim str1 As String
        wdVer = wd.Version
        'convert wdver to integer
        intVer = CInt(wdVer)
        If intVer < 15 Then
            strM = "The WordCompare function works only with Word 2013 or newer."
            MsgBox(strM, vbInformation, "Invalid action")
            wd.Quit()
            GoTo end1
        End If



        Dim doc As Microsoft.Office.Interop.Word.Document
        Dim doc1 As Microsoft.Office.Interop.Word.Document
        Dim doc2 As Microsoft.Office.Interop.Word.Document
        Dim doc3 As Microsoft.Office.Interop.Word.Document
        wd.DisplayAlerts = False

        Dim strPath1 As String = Me.txtDoc1.Text
        Dim strPath2 As String = Me.txtDoc2.Text

        Try

            'strPathT1 and strPathT2 are already open
            'need to open as another document

            doc1 = wd.Documents.Open(strPath1, , True)
            doc2 = wd.Documents.Open(strPath2, , True)


            wd.CompareDocuments(OriginalDocument:=wd.Documents(strPath1), _
    RevisedDocument:=wd.Documents(strPath2), Destination:= _
    Microsoft.Office.Interop.Word.WdCompareDestination.wdCompareDestinationNew, Granularity:=Microsoft.Office.Interop.Word.WdGranularity.wdGranularityWordLevel, _
    CompareFormatting:=Me.chkFormatting.Checked, CompareCaseChanges:=Me.chkCase.Checked, CompareWhitespace:= _
    Me.chkWhiteSpace.Checked, CompareTables:=Me.chkTables.Checked, CompareHeaders:=Me.chkHeaders.Checked, CompareFootnotes:=Me.chkFootnotes.Checked, _
    CompareTextboxes:=Me.chkTextboxes.Checked, CompareFields:=Me.chkFields.Checked, CompareComments:=Me.chkComments.Checked, _
    CompareMoves:=False, RevisedAuthor:="xxx", _
    IgnoreAllComparisonWarnings:=False)

            wd.ActiveWindow.ShowSourceDocuments = Microsoft.Office.Interop.Word.WdShowSourceDocuments.wdShowSourceDocumentsBoth

            doc3 = wd.ActiveDocument

            'for some reason, the first time this is run, Word opens with no Menu or Ribbon.


            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert
            With doc3.ActiveWindow.View.RevisionsFilter
                .Markup = Microsoft.Office.Interop.Word.WdRevisionsMarkup.wdRevisionsMarkupAll
                .View = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
            End With

            doc1.Close(False)
            doc2.Close(False)

            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert
            With doc3.ActiveWindow.View.RevisionsFilter
                .Markup = Microsoft.Office.Interop.Word.WdRevisionsMarkup.wdRevisionsMarkupAll
                .View = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
            End With

            'wd.Visible = True
            Try
                doc3.Activate()
            Catch ex As Exception

            End Try

            'wdPaneRevisionsHoriz

            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert
            With doc3.ActiveWindow.View.RevisionsFilter
                .Markup = Microsoft.Office.Interop.Word.WdRevisionsMarkup.wdRevisionsMarkupAll
                .View = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
            End With

            'LEE: Very strange behavior
            'must do this several times to get what's expected
            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsHoriz
            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert
            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsHoriz
            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert

            With doc3.ActiveWindow.View.RevisionsFilter
                .Markup = Microsoft.Office.Interop.Word.WdRevisionsMarkup.wdRevisionsMarkupAll
                .View = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
            End With

            doc3.ActiveWindow.WindowState = Office.Interop.Word.WdWindowState.wdWindowStateMaximize

            'wd.Visible = True

            doc3.Activate()

            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsHoriz
            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert
            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsHoriz
            doc3.ActiveWindow.View.SplitSpecial = Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneRevisionsVert

            Try
                Dim CB As Microsoft.Office.Core.CommandBar
                Dim docCB As Microsoft.Office.Core.CommandBars = doc3.CommandBars

                For Each CB In docCB
                    str1 = CB.Name
                    If StrComp(str1, "Menu Bar", CompareMethod.Text) = 0 Then
                        CB.Visible = True
                        'MsgBox("True")
                        Exit For
                    End If
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Dim strPwd As String = RandomPswd()
            doc3.Protect(Microsoft.Office.Interop.Word.WdProtectionType.wdAllowOnlyReading, , strPwd)

            strM = "Commpared document opened in Microsoft" & ChrW(8482) & " Word called 'Compare Result....."
            strM = strM & ChrW(10) & ChrW(10) & "You may have to activate the document from the taskbar."
            strM = strM & ChrW(10) & ChrW(10) & "Please note that a Word Automation peculiarity may cause the Word document to display with no menu bar or ribbon."
            strM = strM & ChrW(10) & ChrW(10) & "If this happens, close the 'Compare Result...' Word document and create it again in StudyDoc."
            MsgBox(strM, vbInformation, "All done...")

            wd.Visible = True
            doc3.Activate()

        Catch ex As Exception

        End Try

        wd.DisplayAlerts = True

end1:


    End Sub

    Private Sub cmdSaveSettings_Click(sender As Object, e As EventArgs) Handles cmdSaveSettings.Click

        'MsgBox("Under construction", MsgBoxStyle.Information, "Under construction...")

        Dim strModule As String

        strModule = "Word Compare"

        Try
            Call RecordReportPrelim(Me, strModule)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class