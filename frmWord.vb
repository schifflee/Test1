Public Class frmWord
    Public boolRO As Boolean = True 'Open document as readonly
    Public strPath As String
    Public boolCancel As Boolean = True

    Private Sub frmWord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.WindowState = FormWindowState.Maximized
        Dim t, h, l, w

        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        Me.Top = 0
        Me.Left = 0
        Me.Width = w
        Me.Height = h


        t = Me.Top
        h = Me.Height
        l = Me.Left
        w = Me.Width

        'Me.panWdWB.Top = 0
        'Me.panWdWB.Height = h
        'Me.panWdWB.Width = w - Me.panWord.Left - 30

        'Me.panWdWB.ContextMenuStrip = Me.cmsfrmWord
        Me.panWdWB.Visible = False

        Me.panWord.Top = 0
        Me.panWord.Height = h
        Me.panWord.Width = w - Me.panWord.Left - 30

        Me.panWord.ContextMenuStrip = Me.cmsfrmWord
        Me.panWord.Visible = True

    End Sub

    Sub FormLoad()

        'Me.wbFrmWd.Navigate(strPath)
        Me.afrWord.Open(strPath, boolRO, "Word.Document", "a", "a")
        Me.afrWord.Refresh()


    End Sub

    Private Sub cmdOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpen.Click
        Dim afr As AxDSOFramer.AxFramerControl
        Dim int1 As Short

        Dim t, h, l, w, var1

        afr = Me.afrWord
        'Me.panWord.Visible = False

        't = Me.Top
        'h = Me.Height
        'l = Me.Left
        'w = Me.Width
        'MsgBox(t & ", " & h & ", " & l & ", " & w)

        't = afr.Top
        'h = afr.Height
        'l = afr.Left
        'w = afr.Width
        'MsgBox(t & ", " & h & ", " & l & ", " & w)

        'strPath = "C:\GubbsInc\GuWu\ReportStatements\ReportStatementsGuWu01.doc"
        'strPath = "C:\GubbsInc\GuWu\ReportStatements\ReportStatementsGuWu01.xml"

        boolRO = False 'Read Only

        'Dim oWordApp As Object
        'Dim oWordDoc As Object
        'Dim wd As Object
        'oWordApp = CreateObject("Word.Application")
        ''oWordDoc = oWordApp.Documents.Add
        'wd = oWordApp.documents.open(strPath, , boolRO)

        'Dim wd As New Word.Application
        afr.Open(strPath, boolRO, "Word.Document", "a", "a")
        afr.Refresh()
        'int1 = afr.ActiveDocument.tables.count
        'MsgBox(int1)
        'Me.panWord.Visible = True

    End Sub

    Sub InsertFC()
        Dim boolM As Boolean
        Dim strM As String

        strM = ""
        boolM = False

        Dim pos As Int64
        Dim strT As String
        Dim str1 As String
        Dim strL As String
        Dim strR As String
        Dim wdDoc As Word.Document
        Dim wdApp As Word.Application

        'record position of cursor in text box
        'wd = wb.Document
        wdDoc = Me.afrWord.ActiveDocument
        'wdDoc = Me.wbFrmWd.Document
        wdApp = wdDoc.Application

        'wd_doc_RBS = wd 'frmH.wbRBS.Document
        'wd_app_RBS = wd_doc_RBS.Application

        'MsgBox(wd_app_RBS.Selection.Start)

        Dim frm As New frmFieldCodes

        'Me.Cursor = New Cursor(Cursor.Current.Handle)

        'frm.Location = New Point(Cursor.Position.X, Cursor.Position.Y + 10)

        frm.ShowDialog()

        Me.afrWord.Refresh()

        If frm.boolCancel Then


            'wdapp..Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)

        Else


            wdApp.Selection.TypeText(Text:=Trim(frm.strFC))
            'wd.Content.Text = frm.strFC



        End If

        frm.Dispose()

end1:
    End Sub

    Private Sub cmdFieldCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFieldCode.Click
        Call InsertFC()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Dim wd As Word.Document

        Try
            wd = Me.afrWord.ActiveDocument

            wd.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
            wd.quit(Word.WdSaveOptions.wdDoNotSaveChanges)

        Catch ex As Exception

        End Try

        boolCancel = True
        Me.Close()

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim wd As Word.Document

        wd = Me.afrWord.ActiveDocument

        Me.afrWord.Save()
        Me.afrWord.Close()

        'Try
        '    wd.Close(Word.WdSaveOptions.wdSaveChanges)
        '    wd.quit(Word.WdSaveOptions.wdSaveChanges)

        'Catch ex As Exception

        'End Try

        boolCancel = False
        Me.Close()

    End Sub

    Private Sub cmiFieldCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmiFieldCode.Click
        Call InsertFC()
    End Sub

    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        Try

            Me.panWord.Visible = False
            Me.panWord.Visible = True

            'Me.afrWord.Close()
            'Me.afrWord.Open(strPath, boolRO, "Word.Document", "a", "a")
            'Me.afrWord.Refresh()

        Catch ex As Exception

        End Try

    End Sub
End Class