Imports System
Imports System.IO
Imports System.Text

Public Class frmWordUpdate
    Public strGPath As String
    Public fs As FileStream

    Private Sub frmWordUpdate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim t, l, w, h

        t = Me.Top
        l = Me.Left
        w = Me.Width
        h = Me.Height

        Me.pan1.Width = w - Me.pan1.Left - 50
        Me.pan1.Top = 50
        Me.pan1.Height = (h / 2) - 10

        Me.lblPath.Top = 3
        Me.lblPath.Left = Me.pan1.Left

        Me.pan2.Left = Me.pan1.Left
        Me.pan2.Width = w - Me.pan1.Left - 50
        Me.pan2.Top = Me.pan1.Top + Me.pan1.Height + 10
        Me.pan2.Height = (h / 2) - 10

        Me.pan2.Visible = True
        Me.pan1.Visible = True


    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        Dim strPath As String
        Dim strFilter As String
        Dim strFileName As String
        Dim str1 As String
        Dim str2 As String
        Dim boolGo As Boolean
        Dim boolRO As Boolean

        'strPath = Me.txtArchivePath.Text
        'If Len(strPath) = 0 Then
        'Else
        '    Me.txtArchivePath.Text = "C:\"
        'End If
        'strPath = Me.txtArchivePath.Text

        'strFilter = ".MDB files (*.MDB*)|*.MDB"
        'strFileName = "*.MDB"

        strPath = "C:\"
        strFilter = ".* files (*.*)|*.*"
        strFileName = "*.*"

        str1 = ReturnDirectoryBrowse(True, strPath, strFilter, strFileName) 'true = looking for file

        strGPath = str1
        Me.lblPath.Text = strGPath


        'If Len(str1) = 0 Then
        'Else
        '    boolRO = False
        '    Me.afr1.Open(strGPath, boolRO, "Word.Document", "a", "a")

        'End If
end1:
    End Sub

    Private Sub cmdWordStatements_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWordStatements.Click
        'this routine:
        ' - opens GuWuStatements
        ' - creates a separate word file for every Section Statement

        Dim strPath As String
        Dim strPathT As String
        Dim strPathX As String
        Dim intTables As Short
        'Dim afr As DSOFramer.FramerControl
        Dim Count1 As Short
        Dim boolRO As Boolean
        Dim str1 As String


        strPathT = "C:\GubbsInc\GuWu\Temp\"
        strPathX = "C:\GubbsInc\GuWu\XML\"

        Dim dirInfoT As New System.IO.DirectoryInfo(strPathT)
        If dirInfoT.Exists Then
        Else 'create it
            Directory.CreateDirectory(strPathT)
        End If

        Dim dirInfoX As New System.IO.DirectoryInfo(strPathT)
        If dirInfoX.Exists Then
        Else 'create it
            Directory.CreateDirectory(strPathT)
        End If

        str1 = Me.lblPath.Text
        If StrComp(str1, "Path", CompareMethod.Text) = 0 Then
            MsgBox("Browse for a file", MsgBoxStyle.Information, "Browse for a file...")
            Exit Sub
        Else
            strGPath = Me.lblPath.Text
        End If
        'intTables = Me.afr1.ActiveDocument.tables.count

        'strGPath = "C:\GubbsInc\GuWu\ReportStatements\Prac.doc"
        boolRO = True
        'Me.afr1.Open(strGPath, boolRO, "Word.Document", "a", "a")
        'break these tables down into individual xml files
        'Call dothis(intTables)
        intTables = 0
        Call DoTableCopy(intTables)




    End Sub

    Sub DoTableCopy(ByVal intT As Short)
        Dim Count1 As Short
        Dim Count2 As Short
        Dim strG As String
        Dim strT As String
        Dim strX As String
        Dim fso As New Scripting.FileSystemObject
        Dim fiP As File
        Dim fi As File
        Dim wd As New Word.Application
        Dim wd1 As New Word.Application
        Dim strName As String
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strL1 As String
        Dim strL2 As String
        Dim dtbl1 As DataTable
        Dim rows1() As DataRow
        Dim dtbl2 As DataTable
        Dim rows2() As DataRow
        Dim strF1 As String
        Dim strF2 As String
        Dim intCt As Short
        Dim dtbl3 As DataTable
        Dim rows3() As DataRow
        Dim strF3 As String
        Dim var1, var2, var3
        Dim strNameT As String
        Dim boolRO As Boolean
        Dim idCBS As Int64

        dtbl1 = frmH.tblWordStatements
        dtbl2 = frmH.tblReportStatements
        dtbl3 = frmH.tblConfigBodySections

        strT = "C:\GubbsInc\GuWu\Temp\"
        strX = "C:\GubbsInc\GuWu\XML\"

        strG = "C:\GubbsInc\GuWu\ReportStatements\Prac.doc"



        'wd.Documents.Open(strG)
        'intT = wd.ActiveDocument.Tables.Count
        'wd.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        wd.Documents.Open(strG)
        'wd.Visible = False
        Me.Refresh()

        intT = wd.ActiveDocument.Tables.Count

        intCt = 0

        For Count1 = 2 To intT

            strL1 = Count1 & " of " & intT & "..."
            Me.lblStatus.Text = strL1
            Me.lblStatus.Refresh()

            'make sure there is a tblConfigBodySections id for this table
            'get ind_tblconfigbodysection
            Erase rows3
            strF3 = "INTWORDTABLENUMBER = " & Count1
            rows3 = dtbl3.Select(strF3)
            If rows3.Length = 0 Then
                GoTo next1
            Else
                idCBS = rows3(0).Item("ID_TBLCONFIGBODYSECTIONS")

            End If



            'wd.tables.item(Count2).select()
            wd.Selection.GoTo(What:=Word.WdGoToItem.wdGoToTable, Which:=Word.WdGoToDirection.wdGoToFirst, Count:=Count1, Name:="")
            'Me.afr1.Selection.GoTo(What:=Word.WdGoToItem.wdGoToTable, Which:=Word.WdGoToDirection.wdGoToFirst, Count:=Count1, Name:="")

            wd.Selection.Tables(1).Select()
            int1 = wd.Selection.Tables(1).Rows.Count
            For Count2 = 3 To int1

                strL2 = strL1 & Count2 & " of " & int1
                Me.lblStatus.Text = strL2
                Me.lblStatus.Refresh()

                'goto cell 2,1
                wd.Selection.Tables(1).Cell(Count2, 1)
                str1 = wd.Selection.Tables(1).Cell(Count2, 1).Range.Text

                If Len(str1) = 2 Then
                    Exit For
                End If
                intCt = intCt + 1

                str2 = Mid(str1, 1, Len(str1) - 2)
                strNameT = str2
                str3 = Replace(str2, " ", "_", 1, -1, CompareMethod.Text)
                wd1.Documents.Add() '(Template:="Normal", NewTemplate:=False, DocumentType:=0)
                'wd1.Visible = False
                Me.Refresh()
                'save as
                strName = strX & "WS_" & Format(Count1, "000") & "_" & str3
                wd1.ActiveDocument.SaveAs(FileName:=strName, FileFormat:=Word.WdSaveFormat.wdFormatDocument) ', _
                'wd1.ActiveDocument.SaveAs(FileName:=strName, FileFormat:=Word.WdSaveFormat.wdFormatXML) ', _
                'LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
                ':="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                'SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
                'False)

                wd.Selection.Tables(1).Cell(Count2, 2).Select()
                wd.Selection.Copy()
                'Documents.Add(DocumentType:=wdNewBlankDocument)
                'Selection.PasteAndFormat(wdPasteDefault)

                wd1.Selection.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault)
                wd1.Documents.Close(Word.WdSaveOptions.wdSaveChanges)
                'wd1.Visible = False
                Me.Refresh()

                'record information in database
                Dim newRow As DataRow = dtbl1.NewRow
                newRow.BeginEdit()
                newRow("ID_TBLWORDSTATEMENTS") = intCt
                newRow("ID_TBLCONFIGBODYSECTIONS") = idCBS
                newRow("INTWORDTABLENUMBER") = Count1
                newRow("CHARWORDSTATEMENT") = strName & ".xml"
                newRow("CHARTITLE") = strNameT

                newRow.EndEdit()
                dtbl1.Rows.Add(newRow)

                'now update tblReportStatements dtbl2
                'NO do this later

            Next
next1:

        Next

        Try
            wd1.Quit(Word.WdSaveOptions.wdDoNotSaveChanges)
        Catch ex As Exception

        End Try
        Try
            wd.Quit(Word.WdSaveOptions.wdDoNotSaveChanges)
        Catch ex As Exception

        End Try

        If boolGuWuOracle Then
            Try
                frmH.ta_tblWordStatements.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLWORDSTATEMENTS.Merge(frmH.ds2005.TBLWORDSTATEMENTS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblWordStatementsAcc.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLWORDSTATEMENTS.Merge(frmH.ds2005Acc.TBLWORDSTATEMENTS, True)
            End Try
        End If

        'increase maxid
        Dim maxid As Int64
        Dim tblM As DataTable
        Dim rowM() As DataRow
        Dim strM As String

        tblM = frmH.tblMaxID
        strM = "CHARTABLE = 'tblWordStatements'"
        rowM = tblM.Select(strM)
        rowM(0).BeginEdit()
        rowM(0).Item("NUMMAXID") = intCt
        rowM(0).EndEdit()
        If boolGuWuOracle Then
            Try
                frmH.ta_tblMaxID.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLMAXID.Merge(frmH.ds2005.TBLMAXID, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblMaxIDAcc.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLMAXID.Merge(frmH.ds2005Acc.TBLMAXID, True)
            End Try
        End If

        Me.lblStatus.Text = "Status..."
        Me.lblStatus.Refresh()

    End Sub


    Private Sub cmdReportStatements_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReportStatements.Click
        'this routine populates the new field 'tblWordStatements' that has been added to tblReportStatements

        Dim Count1 As Short
        Dim Count2 As Short
        Dim strG As String
        Dim strT As String
        Dim strX As String
        Dim strName As String
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strL1 As String
        Dim strL2 As String
        Dim dtbl1 As DataTable
        Dim rows1() As DataRow
        Dim dtbl2 As DataTable
        Dim rows2() As DataRow
        Dim strF1 As String
        Dim strF2 As String
        Dim intCt As Short
        Dim dtbl3 As DataTable
        Dim rows3() As DataRow
        Dim strF3 As String
        Dim var1, var2, var3
        Dim strNameT As String
        Dim boolRO As Boolean
        Dim idCBS As Int64
        Dim arr1(0, 0)

        dtbl1 = frmH.tblWordStatements
        dtbl2 = frmH.tblReportStatements
        dtbl3 = frmH.tblConfigBodySections

        'update tblReportStatements
        int1 = dtbl2.Rows.Count
        ReDim arr1(3, int1)
        intCt = 0
        For Count1 = 0 To int1 - 1

            Me.lblStatus.Text = Count1 & " of " & int1 - 1 & "..."
            Me.lblStatus.Refresh()

            var1 = dtbl2.Rows(Count1).Item("ID_TBLCONFIGBODYSECTIONS")
            var2 = dtbl2.Rows(Count1).Item("CHARSTATEMENT")
            strF1 = "ID_TBLCONFIGBODYSECTIONS = " & var1 & " AND CHARTITLE = '" & var2 & "'"
            rows1 = dtbl1.Select(strF1)
            int2 = rows1.Length
            If int2 = 0 Then 'delete record in dtbl2
                intCt = intCt + 1
                var1 = dtbl2.Rows(Count1).Item("id_tblStudies")
                var2 = dtbl2.Rows(Count1).Item("id_tblConfigReportType")
                var3 = dtbl2.Rows(Count1).Item("ID_TBLCONFIGBODYSECTIONS")
                arr1(1, intCt) = var1
                arr1(2, intCt) = var2
                arr1(3, intCt) = var3

            Else
                var3 = rows1(0).Item("ID_TBLWORDSTATEMENTS")
                dtbl2.Rows(Count1).BeginEdit()
                dtbl2.Rows(Count1).Item("ID_TBLWORDSTATEMENTS") = var3
                dtbl2.Rows(Count1).EndEdit()
            End If


        Next

        'now delete table entries
        For Count1 = 1 To intCt
            Me.lblStatus.Text = "Deleting " & Count1 & " of " & intCt & "..."
            Me.lblStatus.Refresh()

            var1 = arr1(1, Count1)
            var2 = arr1(2, Count1)
            var3 = arr1(3, Count1)

            strF1 = "id_tblStudies = " & var1 & " AND id_tblConfigReportType  = " & var2 & " AND ID_TBLCONFIGBODYSECTIONS  = " & var3
            Erase rows2
            rows2 = dtbl2.Select(strF1)
            For Count2 = 0 To rows2.Length - 1
                rows2(0).Delete()
            Next
        Next

        Me.lblStatus.Text = "Status..."

        If boolGuWuOracle Then
            Try
                frmH.ta_tblReportStatements.Update(frmH.tblReportStatements)
            Catch ex As DBConcurrencyException
                '''msgbox("aaData Tab: " & ex.Message)
                frmH.ds2005.TBLREPORTSTATEMENTS.Merge(frmH.ds2005.TBLREPORTSTATEMENTS, True)
            End Try

            Try
                frmH.ta_tblReportStatements.Update(frmH.tblReportStatements)
            Catch ex As DBConcurrencyException
                '''msgbox("aaData Tab: " & ex.Message)
                frmH.ds2005.TBLREPORTSTATEMENTS.Merge(frmH.ds2005.TBLREPORTSTATEMENTS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblReportStatementsAcc.Update(frmH.tblReportStatements)
            Catch ex As DBConcurrencyException
                '''msgbox("aaData Tab: " & ex.Message)
                frmH.ds2005Acc.TBLREPORTSTATEMENTS.Merge(frmH.ds2005Acc.TBLREPORTSTATEMENTS, True)
            End Try

            Try
                frmH.ta_tblReportStatementsAcc.Update(frmH.tblReportStatements)
            Catch ex As DBConcurrencyException
                '''msgbox("aaData Tab: " & ex.Message)
                frmH.ds2005Acc.TBLREPORTSTATEMENTS.Merge(frmH.ds2005Acc.TBLREPORTSTATEMENTS, True)
            End Try
        End If


    End Sub


    Private Sub cmdReportHeaders_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReportHeaders.Click
        'this routine populates the table tblReportHeaders

        Dim Count1 As Short
        Dim Count2 As Short
        Dim strG As String
        Dim strT As String
        Dim strX As String
        Dim strName As String
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strL1 As String
        Dim strL2 As String
        Dim dtbl1 As DataTable
        Dim rows1() As DataRow
        Dim dtbl2 As DataTable
        Dim rows2() As DataRow
        Dim strF1 As String
        Dim strF2 As String
        Dim intCt As Short
        Dim dtbl3 As DataTable
        Dim rows3() As DataRow
        Dim strF3 As String
        Dim var1, var2, var3, var4
        Dim strNameT As String
        Dim boolRO As Boolean
        Dim idCBS As Int64
        Dim arr1(0, 0)

        dtbl1 = frmH.tblReportHeaders
        dtbl2 = frmH.tblReports

        'tblReportHeaders
        'ID_TBLREPORTHEADERS
        'ID_TBLREPORTS
        'ID_TBLSTUDIES
        'BOOLDIFFFIRSTPAGE
        'BOOLINCLUDELOGO
        'CHARHLT
        'CHARHRT
        'CHARHLB
        'CHARHRB
        'CHARFLT
        'CHARFRT
        'CHARFLB
        'CHARFRB
        'UPSIZE_TS


        'tblReports
        'ID_TBLREPORTS
        'CHARREPORTNUMBER
        'DTREPORTDRAFTISSUEDATE
        'DTREPORTFINALISSUEDATE
        'ID_TBLSTUDIES
        'CHARREPORTTEMPLATE
        'ID_TBLCONFIGREPORTTYPE
        'INTCALSTD
        'INTQC
        'INTSHOWBQL
        'INTSHOWCALSTD
        'INTUSERCOMMENTS
        'CHARREPORTTYPE
        'UPSIZE_TS
        'CHARREPORTTITLE
        'BOOLEXCLUDEPSAE



        'update tblReportStatements
        'int1 = dtbl1.Rows.Count
        'If int1 > 0 Then
        '    Exit Sub
        'End If
        int1 = dtbl2.Rows.Count
        intCt = 0
        For Count1 = 0 To int1 - 1 'add a record for each record in dtbl2
            intCt = intCt + 1

            Me.lblStatus.Text = Count1 & " of " & int1 - 1 & "..."
            Me.lblStatus.Refresh()

            var1 = dtbl2.Rows(Count1).Item("ID_TBLREPORTS")
            var2 = dtbl2.Rows(Count1).Item("ID_TBLSTUDIES")
            'strF1 = "ID_TBLREPORTS = " & var1 & " AND ID_TBLSTUDIES = " & var2
            'rows1 = dtbl2.Select(strF1)

            Dim newRow As DataRow = dtbl1.NewRow
            newRow.BeginEdit()

            newRow.Item("ID_TBLREPORTHEADERS") = intCt
            newRow.Item("ID_TBLREPORTS") = var1
            newRow.Item("ID_TBLSTUDIES") = var2
            newRow.Item("BOOLDIFFFIRSTPAGE") = -1
            newRow.Item("BOOLINCLUDELOGO") = 0
            newRow.Item("CHARHLT") = "Project Number: [CORPORATESTUDY/PROJECTNUMBER]"
            newRow.Item("CHARHRT") = "Page [PAGENUMBER]"
            newRow.Item("CHARHLB") = "Final Report"
            newRow.Item("CHARHRB") = "For [SUBMITTEDTO]"
            newRow.Item("CHARFLT") = System.DBNull.Value
            newRow.Item("CHARFRT") = System.DBNull.Value
            newRow.Item("CHARFLB") = System.DBNull.Value
            newRow.Item("CHARFRB") = System.DBNull.Value
            newRow.Item("UPSIZE_TS") = "01-SEP-07"

            newRow.EndEdit()

            dtbl1.Rows.Add(newRow)

        Next

        Me.lblStatus.Text = "Status..."

        If boolGuWuOracle Then
            Try
                frmH.ta_tblReportHeaders.Update(frmH.tblReportHeaders)
            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLREPORTHEADERS.Merge(frmH.ds2005.TBLREPORTHEADERS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblReportHeadersAcc.Update(frmH.tblReportHeaders)
            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLREPORTHEADERS.Merge(frmH.ds2005Acc.TBLREPORTHEADERS, True)
            End Try
        End If


        'increase maxid
        Dim maxid As Int64
        Dim tblM As DataTable
        Dim rowM() As DataRow
        Dim strM As String

        tblM = frmH.tblMaxID
        strM = "CHARTABLE = 'tblReportHeaders'"
        rowM = tblM.Select(strM)
        rowM(0).BeginEdit()
        rowM(0).Item("NUMMAXID") = intCt
        rowM(0).EndEdit()
        If boolGuWuOracle Then
            Try
                frmH.ta_tblMaxID.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLMAXID.Merge(frmH.ds2005.TBLMAXID, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblMaxIDAcc.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLMAXID.Merge(frmH.ds2005Acc.TBLMAXID, True)
            End Try
        End If

        Me.lblStatus.Text = "Status..."
        Me.lblStatus.Refresh()

    End Sub

    Private Sub cmdFieldCodes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmdDirectories_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDirectories.Click
        'this routine does not need to be run because it was performed in Word Statements 
        Dim Count1 As Short
        Dim Count2 As Short
        Dim strG As String
        Dim strT As String
        Dim strX As String
        Dim fso As New Scripting.FileSystemObject
        Dim fiP As Scripting.File
        Dim fi As Scripting.File
        Dim dirX As Scripting.Folder

        'Dim wd As New Word.Application
        'Dim wd1 As New Word.Application
        Dim strName As String
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strL1 As String
        Dim strL2 As String
        Dim dtbl1 As DataTable
        Dim rows1() As DataRow
        Dim dtbl2 As DataTable
        Dim rows2() As DataRow
        Dim strF1 As String
        Dim strF2 As String
        Dim intCt As Short
        Dim dtbl3 As DataTable
        Dim rows3() As DataRow
        Dim strF3 As String
        Dim var1, var2, var3
        Dim strNameT As String
        Dim boolRO As Boolean
        Dim idCBS As Int64
        Dim intT As Short
        Dim strP As String
        Dim strTi As String
        Dim id
        Dim intTN As Short

        dtbl1 = frmH.tblWordStatements
        dtbl2 = frmH.tblReportStatements
        dtbl3 = frmH.tblConfigBodySections

        For Count1 = 0 To dtbl1.Rows.Count - 1
            dtbl1.Rows(Count1).BeginEdit()
            var1 = dtbl1.Rows(Count1).Item("CHARWORDSTATEMENT")
            var2 = var1 & ".XML"
            dtbl1.Rows(Count1).Item("CHARWORDSTATEMENT") = var2
            dtbl1.Rows(Count1).EndEdit()

        Next

        If boolGuWuOracle Then
            Try
                frmH.ta_tblWordStatements.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                '''msgbox("aaData Tab: " & ex.Message)
                frmH.ds2005.TBLWORDSTATEMENTS.Merge(frmH.ds2005.TBLWORDSTATEMENTS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblWordStatementsAcc.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                '''msgbox("aaData Tab: " & ex.Message)
                frmH.ds2005Acc.TBLWORDSTATEMENTS.Merge(frmH.ds2005Acc.TBLWORDSTATEMENTS, True)
            End Try
        End If

        Exit Sub

        strT = "C:\GubbsInc\GuWu\Temp\"
        strX = "C:\GubbsInc\GuWu\XML\"

        strG = "C:\GubbsInc\GuWu\ReportStatements\Prac.doc"

        dirX = fso.GetFolder(strX)


        'wd.Documents.Open(strG)
        'intT = wd.ActiveDocument.Tables.Count
        'wd.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        'wd.Documents.Open(strG)
        'wd.Visible = False
        Me.Refresh()

        'intT = wd.ActiveDocument.Tables.Count
        intT = dtbl1.Rows.Count


        intCt = 0

        For Each fi In dirX.Files
            intCt = intCt + 1

            strL1 = intCt & " of " & dirX.Files.Count & "..."
            Me.lblStatus.Text = strL1
            Me.lblStatus.Refresh()

            strName = fi.Name
            'find id and chartitle
            int1 = InStr(1, strName, "_", CompareMethod.Text)
            int2 = InStr(int1 + 1, strName, "_", CompareMethod.Text)
            id = Mid(strName, int1 + 1, int2 - int1 - 1)
            strTi = Mid(strName, int2 + 1, Len(strName) - int2)
            id = CLng(id)

            'make sure there is a tblConfigBodySections id for this table
            'get id_tblconfigbodysection
            Erase rows2
            strF2 = "INTWORDTABLENUMBER = " & id
            rows2 = dtbl2.Select(strF3)
            If rows2.Length = 0 Then
                GoTo next1
            Else
                idCBS = rows2(0).Item("ID_TBLCONFIGBODYSECTIONS")
            End If

            'get id_tblconfigbodysection
            Erase rows3
            strF3 = "ID_TBLCONFIGBODYSECTIONS = " & idCBS
            rows3 = dtbl3.Select(strF3)
            If rows3.Length = 0 Then
                GoTo next1
            Else
                intTN = rows3(0).Item("INTWORDTABLENUMBER")
            End If

            'record information in database
            Dim newRow As DataRow = dtbl1.NewRow
            newRow.BeginEdit()
            newRow("ID_TBLWORDSTATEMENTS") = id
            newRow("ID_TBLCONFIGBODYSECTIONS") = idCBS
            newRow("INTWORDTABLENUMBER") = intTN
            newRow("CHARTITLE") = strTi
            newRow("CHARWORDSTATEMENT") = strX & fi.Name

            newRow.EndEdit()
            dtbl1.Rows.Add(newRow)

            'now update tblReportStatements dtbl2
            'NO do this later

            'Next
next1:

        Next


        If boolGuWuOracle Then
            Try
                frmH.ta_tblWordStatements.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLWORDSTATEMENTS.Merge(frmH.ds2005.TBLWORDSTATEMENTS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblWordStatementsAcc.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLWORDSTATEMENTS.Merge(frmH.ds2005Acc.TBLWORDSTATEMENTS, True)
            End Try
        End If

        'increase maxid
        Dim maxid As Int64
        Dim tblM As DataTable
        Dim rowM() As DataRow
        Dim strM As String

        tblM = frmH.tblMaxID
        strM = "CHARTABLE = 'tblWordStatements'"
        rowM = tblM.Select(strM)
        rowM(0).BeginEdit()
        rowM(0).Item("NUMMAXID") = intCt
        rowM(0).EndEdit()
        If boolGuWuOracle Then
            Try
                frmH.ta_tblMaxID.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLMAXID.Merge(frmH.ds2005.TBLMAXID, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblMaxIDAcc.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLMAXID.Merge(frmH.ds2005Acc.TBLMAXID, True)
            End Try
        End If

        Me.lblStatus.Text = "Status..."
        Me.lblStatus.Refresh()
    End Sub

    Private Sub cmdStoreXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStoreXML.Click
        'this routine doesn't neet to be run

        Dim Count1 As Short
        Dim Count2 As Short
        Dim strG As String
        Dim strT As String
        Dim strX As String
        Dim fso As New Scripting.FileSystemObject
        Dim fiP As Scripting.File
        Dim fi As Scripting.File
        Dim dirX As Scripting.Folder

        'Dim wd As New Word.Application
        'Dim wd1 As New Word.Application
        Dim strName As String
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strL1 As String
        Dim strL2 As String
        Dim dtbl1 As DataTable
        Dim rows1() As DataRow
        Dim dtbl2 As DataTable
        Dim rows2() As DataRow
        Dim strF1 As String
        Dim strF2 As String
        Dim intCt As Short
        Dim dtbl3 As DataTable
        Dim rows3() As DataRow
        Dim strF3 As String
        Dim var1, var2, var3
        Dim strNameT As String
        Dim boolRO As Boolean
        Dim idCBS As Int64
        Dim intT As Short
        Dim strP As String
        Dim strTi As String
        Dim id
        Dim intTN As Short

        dtbl1 = frmH.tblWordStatements
        dtbl2 = frmH.tblReportStatements
        dtbl3 = frmH.tblConfigBodySections

        For Count1 = 0 To dtbl1.Rows.Count - 1
            dtbl1.Rows(Count1).BeginEdit()
            var1 = dtbl1.Rows(Count1).Item("CHARWORDSTATEMENT")
            var2 = var1 & ".XML"
            dtbl1.Rows(Count1).Item("CHARWORDSTATEMENT") = var2
            dtbl1.Rows(Count1).EndEdit()

        Next

        If boolGuWuOracle Then
            Try
                frmH.ta_tblWordStatements.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                '''msgbox("aaData Tab: " & ex.Message)
                frmH.ds2005.TBLWORDSTATEMENTS.Merge(frmH.ds2005.TBLWORDSTATEMENTS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblWordStatementsAcc.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                '''msgbox("aaData Tab: " & ex.Message)
                frmH.ds2005Acc.TBLWORDSTATEMENTS.Merge(frmH.ds2005Acc.TBLWORDSTATEMENTS, True)
            End Try
        End If

        Exit Sub

        strT = "C:\GubbsInc\GuWu\Temp\"
        strX = "C:\GubbsInc\GuWu\XML\"

        strG = "C:\GubbsInc\GuWu\ReportStatements\Prac.doc"

        dirX = fso.GetFolder(strX)


        'wd.Documents.Open(strG)
        'intT = wd.ActiveDocument.Tables.Count
        'wd.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        'wd.Documents.Open(strG)
        'wd.Visible = False
        Me.Refresh()

        'intT = wd.ActiveDocument.Tables.Count
        intT = dtbl1.Rows.Count


        intCt = 0

        For Each fi In dirX.Files
            intCt = intCt + 1

            strL1 = intCt & " of " & dirX.Files.Count & "..."
            Me.lblStatus.Text = strL1
            Me.lblStatus.Refresh()

            strName = fi.Name
            'find id and chartitle
            int1 = InStr(1, strName, "_", CompareMethod.Text)
            int2 = InStr(int1 + 1, strName, "_", CompareMethod.Text)
            id = Mid(strName, int1 + 1, int2 - int1 - 1)
            strTi = Mid(strName, int2 + 1, Len(strName) - int2)
            id = CLng(id)

            'make sure there is a tblConfigBodySections id for this table
            'get id_tblconfigbodysection
            Erase rows2
            strF2 = "INTWORDTABLENUMBER = " & id
            rows2 = dtbl2.Select(strF3)
            If rows2.Length = 0 Then
                GoTo next1
            Else
                idCBS = rows2(0).Item("ID_TBLCONFIGBODYSECTIONS")
            End If

            'get id_tblconfigbodysection
            Erase rows3
            strF3 = "ID_TBLCONFIGBODYSECTIONS = " & idCBS
            rows3 = dtbl3.Select(strF3)
            If rows3.Length = 0 Then
                GoTo next1
            Else
                intTN = rows3(0).Item("INTWORDTABLENUMBER")
            End If

            'record information in database
            Dim newRow As DataRow = dtbl1.NewRow
            newRow.BeginEdit()
            newRow("ID_TBLWORDSTATEMENTS") = id
            newRow("ID_TBLCONFIGBODYSECTIONS") = idCBS
            newRow("INTWORDTABLENUMBER") = intTN
            newRow("CHARTITLE") = strTi
            newRow("CHARWORDSTATEMENT") = strX & fi.Name

            newRow.EndEdit()
            dtbl1.Rows.Add(newRow)

            'now update tblReportStatements dtbl2
            'NO do this later

            'Next
next1:

        Next

        If boolGuWuOracle Then
            Try
                frmH.ta_tblWordStatements.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLWORDSTATEMENTS.Merge(frmH.ds2005.TBLWORDSTATEMENTS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblWordStatementsAcc.Update(frmH.tblWordStatements)
            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLWORDSTATEMENTS.Merge(frmH.ds2005Acc.TBLWORDSTATEMENTS, True)
            End Try
        End If


        'increase maxid
        Dim maxid As Int64
        Dim tblM As DataTable
        Dim rowM() As DataRow
        Dim strM As String

        tblM = frmH.tblMaxID
        strM = "CHARTABLE = 'tblWordStatements'"
        rowM = tblM.Select(strM)
        rowM(0).BeginEdit()
        rowM(0).Item("NUMMAXID") = intCt
        rowM(0).EndEdit()
        If boolGuWuOracle Then
            Try
                frmH.ta_tblMaxID.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLMAXID.Merge(frmH.ds2005.TBLMAXID, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblMaxIDAcc.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLMAXID.Merge(frmH.ds2005Acc.TBLMAXID, True)
            End Try
        End If

        Me.lblStatus.Text = "Status..."
        Me.lblStatus.Refresh()
    End Sub

    Private Sub cmdFileStream_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFileStream.Click
        'this routine will populate the table tblWordDocs

        Call DoDBUpdate(True)

        Exit Sub

        Dim dtbl1 As DataTable
        Dim dtbl2 As DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strPath As String
        Dim id As Int64
        Dim intL As Int64
        Dim strpathT As String
        Dim Count1 As Int16
        Dim Count10 As Short
        Dim intMax As Int64

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String

        Dim con As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim strT As String

        con.Open(constrIni)
        cmd.ActiveConnection = con

        dtbl1 = frmH.tblWordStatements
        dtbl2 = frmH.tblWorddocs
        strT = "tblWorddocs"

        Dim boolE As Boolean

        intMax = 0
        For Count10 = 0 To dtbl1.Rows.Count - 1

            Count1 = 0

            Me.lblStatus.Text = Count10 & " of " & dtbl1.Rows.Count - 1 & "..."
            Me.lblStatus.Refresh()

            strPath = dtbl1.Rows(Count10).Item("CHARWORDSTATEMENT")
            id = dtbl1.Rows(Count10).Item("ID_TBLWORDSTATEMENTS")

            boolE = True
            Count1 = 0
            Do Until boolE = False
                Count1 = Count1 + 1
                strpathT = "C:\GubbsInc\GuWu\Temp\Temp" & Format(Count1, "00000") & ".xml"
                If File.Exists(strpathT) Then
                Else
                    Exit Do
                End If
            Loop

            Count1 = 0

            fs = File.OpenRead(strPath)
            intL = fs.Length

            Dim b(8) As Byte
            'Dim temp As UTF8Encoding = New UTF8Encoding(True)
            'Dim temp As New UTF8Encoding(False, True)
            Dim temp As System.Text.Encoding = System.Text.Encoding.UTF8

            'Dim utf8 As New UTF8Encoding()
            'Dim utf8ThrowException As New UTF8Encoding(False, True)


            Dim var1, var2

            Count1 = 0
            boolE = False

            Dim strW As String
            Dim int2 As Short
            strW = ""
            boolE = False
            int2 = 0


            Do Until boolE
                If fs.Read(b, 0, b.Length) <= 0 Then
                    boolE = True
                    int2 = 4000
                Else
                    int2 = int2 + 1
                    Try
                        strW = strW & temp.GetString(b)

                    Catch ex As Exception
                        boolE = True
                        Exit Do
                    End Try
                End If

                If int2 > 2000 Then
                    intMax = intMax + 1
                    Count1 = Count1 + 1
                    If intMax >= 46 Then
                        var1 = Len(strW)
                    End If

                    var1 = Len(strW)
                    Dim newRow As DataRow = dtbl2.NewRow
                    newRow.BeginEdit()
                    newRow.Item("ID_TBLWORDDOCS") = intMax
                    newRow.Item("ID_TBLWORDSTATEMENTS") = id
                    newRow.Item("CHARXML") = strW
                    'If boolE Then
                    '    newRow.Item("CHARXML") = strW
                    'Else
                    '    newRow.Item("CHARXML") = RTrim(LTrim((strW)))
                    'End If
                    'newRow.Item("CHARXML") = RTrim(LTrim((strW)))
                    newRow.Item("UPSIZE_TS") = "01-SEP-07"
                    newRow.EndEdit()
                    dtbl2.Rows.Add(newRow)

                    'ID_TBLWORDDOCS
                    'ID_TBLWORDSTATEMENTS
                    'CHARXML
                    'UPSIZE_TS
                    str1 = "("
                    str1 = str1 & intMax & ", "
                    str1 = str1 & id & ", '"
                    str1 = str1 & RTrim(LTrim((strW))) & "', '"
                    'str1 = Replace(str1, """", """""", 1, -1, CompareMethod.Text)
                    'str1 = str1 & "Hi', "
                    str1 = str1 & "01-SEP-07')"

                    str2 = "INSERT INTO " & strT & " (ID_TBLWORDDOCS, ID_TBLWORDSTATEMENTS, CHARXML, UPSIZE_TS)"
                    str3 = " VALUES " & str1 '(" & intMaxID & ", '" & str1 & "',1)"

                    strSQL = str2 & str3

                    '''''''console.writeline(strSQL)

                    'cmd.CommandText = strSQL
                    'cmd.CommandType = CommandTypeEnum.adCmdText

                    'cmd.Execute()


                    strW = ""
                    int2 = 0


                End If

            Loop



            fs.Close()


        Next

        If boolGuWuOracle Then
            Try
                Try
                    frmH.ta_tblWorddocs.Update(frmH.tblWorddocs)
                Catch ex As Exception
                    frmH.ds2005.TBLWORDDOCS.Merge(frmH.ds2005.TBLWORDDOCS, True)
                End Try

            Catch ex As DBConcurrencyException
                ''msgbox("aaContributing Personnel: " & ex.Message)
                frmH.ds2005.TBLWORDDOCS.Merge(frmH.ds2005.TBLWORDDOCS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                Try
                    frmH.ta_tblWorddocsAcc.Update(frmH.tblWorddocs)
                Catch ex As Exception
                    frmH.ds2005Acc.TBLWORDDOCS.Merge(frmH.ds2005Acc.TBLWORDDOCS, True)
                End Try

            Catch ex As DBConcurrencyException
                ''msgbox("aaContributing Personnel: " & ex.Message)
                frmH.ds2005Acc.TBLWORDDOCS.Merge(frmH.ds2005Acc.TBLWORDDOCS, True)
            End Try
        End If

        Me.lblStatus.Text = "Status..."

        Dim tblM As DataTable
        Dim rowM() As DataRow
        Dim strM As String

        tblM = frmH.tblMaxID
        strM = "CHARTABLE = 'tblworddocs'"
        rowM = tblM.Select(strM)
        rowM(0).BeginEdit()
        rowM(0).Item("NUMMAXID") = intMax
        rowM(0).EndEdit()
        If boolGuWuOracle Then
            Try
                frmH.ta_tblMaxID.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLMAXID.Merge(frmH.ds2005.TBLMAXID, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                frmH.ta_tblMaxIDAcc.Update(frmH.tblMaxID)
            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLMAXID.Merge(frmH.ds2005Acc.TBLMAXID, True)
            End Try
        End If

        'sw.Close()
        'sw.Dispose()

        fs.Close()

        'fs1.Dispose()

        'write to file
        'Me.wb1.Navigate(strpathT)


    End Sub

    Private Sub wb1_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles wb1.DocumentCompleted
        fs.close()
    End Sub

    Private Sub cmdRetrieveDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRetrieveDB.Click
        Dim dtbl1 As DataTable
        Dim dtbl2 As DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strPath As String
        Dim id As Int64
        Dim intL As Int64
        Dim strpathT As String
        Dim Count1 As Int16
        Dim strW As String

        dtbl1 = frmH.tblWordStatements
        dtbl2 = frmH.tblWorddocs

        strPath = dtbl1.Rows(0).Item("CHARWORDSTATEMENT")
        id = dtbl1.Rows(0).Item("ID_TBLWORDSTATEMENTS")
        strF = "ID_TBLWORDSTATEMENTS = " & id
        strS = "ID_TBLWORDDOCS ASC"
        rows2 = dtbl2.Select(strF, strS)
        intL = rows2.Length

        Dim boolE As Boolean
        boolE = True
        Count1 = 0
        strpathT = ""
        Do Until boolE = False
            Count1 = Count1 + 1
            strpathT = "C:\GubbsInc\GuWu\Temp\Temp" & Format(Count1, "00000") & ".xml"
            If File.Exists(strpathT) Then
            Else
                Exit Do
            End If
        Loop

        For Count1 = 0 To intL - 1
            strW = strW & rows2(Count1).Item("CHARXML")
        Next

        Dim info As Byte() = New UTF8Encoding(True).GetBytes(strW)

        strW = ""
        fs = File.Create(strpathT)
        fs.Close()
        fs = File.OpenWrite(strpathT)


        ' Add some information to the file.
        fs.Write(info, 0, info.Length)
        fs.Close()

        Me.wb1.Navigate(strpathT)


    End Sub

    Sub DoDBUpdate(ByVal boolAll As Boolean)
        'this routine will populate the table tblWordDocs

        Dim dtbl1 As DataTable
        Dim dtbl2 As DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strPath As String
        Dim id As Int64
        Dim intL As Int64
        Dim strpathT As String
        Dim Count1 As Int16
        Dim Count10 As Short
        Dim intMax As Int64

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim intS As Short
        Dim intE As Short

        Dim var1, var2, var3

        'Dim con As New ADODB.Connection
        'Dim cmd As New ADODB.Command
        Dim strT As String

        'con.Open(constrIni)
        'cmd.ActiveConnection = con

        dtbl1 = frmH.tblWordStatements
        dtbl2 = frmH.tblWorddocs
        strT = "tblWorddocs"

        If boolAll Then
            intS = 0
            intE = dtbl1.Rows.Count - 1
        Else
            intS = 2
            intE = 2
        End If

        Dim boolE As Boolean

        For Count10 = intS To intE
            If boolAll Then
                strPath = dtbl1.Rows(Count10).Item("CHARWORDSTATEMENT")
                id = dtbl1.Rows(Count10).Item("ID_TBLWORDSTATEMENTS")

            Else
                Select Case Count10
                    Case 1
                        'IMPORTANT!
                        'Enter these values manually
                        id = 71
                        intMax = 2000
                        'done entering values

                        strF = "ID_TBLWORDSTATEMENTS = " & id
                        rows1 = dtbl1.Select(strF)
                    Case 2
                        'IMPORTANT!
                        'Enter these values manually
                        id = 73
                        intMax = 3000
                        'done entering values

                        strF = "ID_TBLWORDSTATEMENTS = " & id
                        rows1 = dtbl1.Select(strF)

                End Select
                strPath = rows1(0).Item("CHARWORDSTATEMENT")

            End If

            Me.lblStatus.Text = Count10 & " of " & dtbl1.Rows.Count - 1 & "..."
            Me.lblStatus.Refresh()

            If Me.chkJustView.Checked Then
            Else
                'id = rows1(Count10).Item("ID_TBLWORDSTATEMENTS")

                'boolE = True
                'Count1 = 0
                'Do Until boolE = False
                '    Count1 = Count1 + 1
                '    strpathT = "C:\GubbsInc\GuWu\Temp\Temp" & Format(Count1, "00000") & ".xml"
                '    If File.Exists(strpathT) Then
                '    Else
                '        Exit Do
                '    End If
                'Loop

                Count1 = 0

                fs = File.OpenRead(strPath)
                'Stream.charset = "iso-8859-1"

                intL = fs.Length

                'Dim b(0) As Byte
                Dim strW As String
                Dim int2 As Short
                Dim b(intL) As Byte
                Dim chars() As Char
                Dim c As Char
                Dim Count2 As Int64

                fs.Read(b, 0, b.Length)

                'Dim temp As UTF8Encoding = New UTF8Encoding(True)
                Dim temp As System.Text.Encoding = System.Text.Encoding.UTF8
                Dim utf8Decoder As Decoder = Encoding.UTF8.GetDecoder()
                Dim charCount As Integer = utf8Decoder.GetCharCount(b, 0, b.Length)
                chars = New Char(charCount - 1) {}
                Dim charsDecodedCount As Integer = utf8Decoder.GetChars(b, 0, b.Length, chars, 0)

                strW = ""
                Count1 = 0
                Count2 = 0
                For Each c In chars
                    Count1 = Count1 + 1
                    Count2 = Count2 + 1
                    strW = strW & c
                    If Count1 = 2000 Or Count2 = charCount - 1 Then
                        intMax = intMax + 1
                        var1 = Len(strW)
                        Dim newRow As DataRow = dtbl2.NewRow
                        newRow.BeginEdit()
                        newRow.Item("ID_TBLWORDDOCS") = intMax
                        newRow.Item("ID_TBLWORDSTATEMENTS") = id
                        newRow.Item("CHARXML") = strW
                        'newRow.Item("CHARXML") = RTrim(LTrim((strW)))
                        newRow.Item("UPSIZE_TS") = "01-SEP-07"
                        newRow.EndEdit()
                        dtbl2.Rows.Add(newRow)

                        Count1 = 0
                        strW = ""
                    End If

                Next

                fs.Close()
            End If

        Next

        If boolGuWuOracle Then
            Try
                Try
                    frmH.ta_tblWorddocs.Update(frmH.tblWorddocs)
                Catch ex As Exception
                    frmH.ds2005.TBLWORDDOCS.Merge(frmH.ds2005.TBLWORDDOCS, True)
                End Try

            Catch ex As DBConcurrencyException
                frmH.ds2005.TBLWORDDOCS.Merge(frmH.ds2005.TBLWORDDOCS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                Try
                    frmH.ta_tblWorddocsAcc.Update(frmH.tblWorddocs)
                Catch ex As Exception
                    frmH.ds2005Acc.TBLWORDDOCS.Merge(frmH.ds2005Acc.TBLWORDDOCS, True)
                End Try

            Catch ex As DBConcurrencyException
                frmH.ds2005Acc.TBLWORDDOCS.Merge(frmH.ds2005Acc.TBLWORDDOCS, True)
            End Try
        End If

        Me.lblStatus.Text = "Status..."

        If Me.chkJustView.Checked Then
        Else
            Dim tblM As DataTable
            Dim rowM() As DataRow
            Dim strM As String

            tblM = frmH.tblMaxID
            strM = "CHARTABLE = 'tblworddocs'"
            rowM = tblM.Select(strM)
            rowM(0).BeginEdit()
            rowM(0).Item("NUMMAXID") = intMax
            rowM(0).EndEdit()
            If boolGuWuOracle Then
                Try
                    frmH.ta_tblMaxID.Update(frmH.tblMaxID)
                Catch ex As DBConcurrencyException
                    frmH.ds2005.TBLMAXID.Merge(frmH.ds2005.TBLMAXID, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    frmH.ta_tblMaxIDAcc.Update(frmH.tblMaxID)
                Catch ex As DBConcurrencyException
                    frmH.ds2005Acc.TBLMAXID.Merge(frmH.ds2005Acc.TBLMAXID, True)
                End Try
            End If

            'sw.Close()
            'sw.Dispose()

            fs.Close()

            'refresh table
            If boolGuWuOracle Then
                frmH.tblWorddocs.BeginLoadData()
                frmH.ta_tblWorddocs.Fill(frmH.tblWorddocs)
                frmH.tblWorddocs.EndLoadData()
            ElseIf boolGuWuAccess Then
                frmH.tblWorddocs.BeginLoadData()
                frmH.ta_tblWorddocsAcc.Fill(frmH.tblWorddocs)
                frmH.tblWorddocs.EndLoadData()
            End If

        End If


        Call NavigateWB1(id)


    End Sub

    Private Sub cmdIndDBUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdIndDBUpdate.Click
        Call DoDBUpdate(False)
    End Sub

    Sub NavigateWB1(ByVal id As Int64)



        Dim dgv As DataGridView
        Dim intRow As Short
        Dim strPath As String


        'strPath = Replace(strPath, ".XML", ".DOC", 1, -1, CompareMethod.Text)
        'strPath = Replace(strPath, "C:", "\\Gubbs09\Gubbs09\GubbsInc\GuWu\XML", 1, -1, CompareMethod.Text)
        'strPath = Replace(strPath, "C:", "\\Gubbs09\Gubbs09", 1, -1, CompareMethod.Text)
        '\\Gubbs09\Gubbs09\GubbsInc\GuWu\XML
        'strPath = Replace(strPath, ".xml", ".doc", 1, -1, CompareMethod.Text)
        'If Len(strPath) = 0 Then
        '    Try
        '        'frmH.afrRBS.Close()
        '    Catch ex As Exception

        '    End Try
        '    frmH.wbRBS.Navigate("about:blank")
        '    Exit Sub
        'End If

        Dim var1, var2
        Dim dtbl1 As DataTable
        Dim dtbl2 As DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intL As Int64
        Dim strpathT As String
        Dim Count1 As Int16
        Dim strW As String
        Dim fs As FileStream
        Dim strM As String

        dtbl1 = frmH.tblWordStatements
        dtbl2 = frmH.tblWorddocs

        'strPath = NZ(dgv("CHARWORDSTATEMENT", intRow).Value, "") 'don't need
        'id = dgv("ID_TBLWORDSTATEMENTS", intRow).Value
        strF = "ID_TBLWORDSTATEMENTS = " & id
        strS = "ID_TBLWORDDOCS ASC"
        rows2 = dtbl2.Select(strF, strS)
        intL = rows2.Length

        'If intL = 0 Then
        '    frmH.wbRBS.Navigate("about:blank")
        '    GoTo end1
        'End If

        Dim boolE As Boolean
        boolE = True
        Count1 = 0
        strpathT = ""
        Do Until boolE = False
            Count1 = Count1 + 1
            strpathT = "C:\GubbsInc\GuWu\Temp\Temp" & Format(Count1, "00000") & ".xml"
            If File.Exists(strpathT) Then
            Else
                Exit Do
            End If
        Loop

        strW = ""
        For Count1 = 0 To intL - 1
            strW = strW & rows2(Count1).Item("CHARXML")
        Next

        Dim intA As Int64
        Dim intB As Int64
        intA = Len(strW)

        strW = RTrim(strW)
        intB = Len(strW)

        'Dim info As Byte() = New UTF8Encoding(True).GetBytes(strW)


        ' Add some information to the file.
        Dim info As Byte()
        If intL = 0 Then
            strM = "There is a problem with this data:" & ChrW(10)
            strM = strM & "tblWorddocs: " & strF & ChrW(10)
            strM = strM & "Please contact your GuWu system administrator."
            info = New UTF8Encoding(True).GetBytes(strM)
            strpathT = Replace(strpathT, ".XML", ".TXT", 1, -1, CompareMethod.Text)
        Else
            ' Add some information to the file.
            info = New UTF8Encoding(True).GetBytes(strW)
        End If

        fs = File.Create(strpathT)
        fs.Close()
        fs = File.OpenWrite(strpathT)

        fs.Write(info, 0, info.Length)

        fs.Close()

        'Dim strXSL As String
        'Dim strXSLPath As String
        'strXSLPath = "C:\GubbsInc\GuWu\XML\GuWuWord.xsl"

        ''need to convert strXSLPath to string
        'Dim fs1 As FileStream
        'fs1 = File.OpenRead(strXSLPath)
        ''Stream.charset = "iso-8859-1"

        'intL = fs1.Length

        ''Dim b(0) As Byte
        'Dim int2 As Short
        'Dim b(intL) As Byte
        'Dim chars() As Char
        'Dim c As Char
        'Dim Count2 As Int64

        'fs1.Read(b, 0, b.Length)

        ''Dim temp As UTF8Encoding = New UTF8Encoding(True)
        'Dim temp As System.Text.Encoding = System.Text.Encoding.UTF8
        'Dim utf8Decoder As Decoder = Encoding.UTF8.GetDecoder()
        'Dim charCount As Integer = utf8Decoder.GetCharCount(b, 0, b.Length)
        'chars = New Char(charCount - 1) {}
        'Dim charsDecodedCount As Integer = utf8Decoder.GetChars(b, 0, b.Length, chars, 0)

        'strXSL = ""
        'Count1 = 0
        'Count2 = 0
        'For Each c In chars
        '    Count1 = Count1 + 1
        '    Count2 = Count2 + 1
        '    strXSL = strXSL & c
        'Next

        'fs1.Close()

        'screw it. can't get xml/xsl stuff to work
        ' Me.wb1.DocumentStream = TransformXML(strW, strXSL)



        'Me.wb1.Navigate(strpathT)

        'Call LoadDocStream(strW)


        'clean up temp file
        'File.Delete(strpathT)

end1:

    End Sub

    Friend Shared Function TransformXML(ByVal xmlString As String, ByVal xslString As String) As MemoryStream
        Dim memStream As MemoryStream = Nothing
        Try
            ' Create a xml document from the sent-in string
            Dim xmlDoc As New Xml.XmlDocument
            xmlDoc.LoadXml(xmlString)

            ' Load the xsl as a xmldocument, from the sent-in string
            Dim xslDoc As New Xml.XmlDocument
            xslDoc.LoadXml(xslString)

            ' Create and load an transformation
            Dim trans As New Xml.Xsl.XslCompiledTransform
            trans.Load(xslDoc)

            ' Create a memorystrem to hold the response
            memStream = New MemoryStream()
            Dim srWriter As New StreamWriter(memStream)

            ' Transform according to the xsl and save the result in the memStream
            ' variable
            trans.Transform(xmlDoc, Nothing, memStream)

            ' Set the intial position of the memorystream
            memStream.Position = 0
        Catch ex As Exception
            Console.Write(ex.ToString())
            MsgBox(ex.ToString)
        End Try
        Return memStream
    End Function



    Sub LoadDocStream(ByVal strW As String)

        '     Dim xmlDoc As New Xml.XmlDocument


        '     'XmlDocument(xmlDoc = New XmlDocument())
        '     'Next, you need to load the string into an XmlDocument object to parse it.


        '     xmlDoc.LoadXml(strW)

        '     'Now, onto the XSLT that performs the transformation. Listing 2 contains a sample of the 
        '     'XML generated thus far, and Listing 3 shows the conv.xslt file that converts the XML to HTML. 
        '     'To achieve this you need to declare an XslCompiledTransform object and load the XSLT into it:

        '     Dim xslt As New Xml.Xsl.XslCompiledTransform
        '     'XslCompiledTransform xslt = new XslCompiledTransform();
        '     xslt.Load("c:\\conv.xslt")

        '     'Now, create a memory stream and open an XmlTextWriter on the stream. 
        '     'The xslt.Transform method uses the XmlTextWriter to write the transformed XML to this stream. 
        '     'When you're done, set the stream's position to its beginning.

        '     Dim mem = New MemoryStream()
        '     'Dim myWriter As Xml.XmlTextWriter

        '     Dim myWriter As New Xml.XmlTextWriter(mem, Encoding.ASCII)

        '     'XmlTextWriter myWriter = new XmlTextWriter(mem, Encoding.ASCII);
        'xslt.Transform(xmlDoc, myWriter);
        'mem.Position = 0;

        '     'You can now point the WebBrowser object's DocumentStream property at this stream, 
        '     'which causes the browser to load and render the HTML.

    End Sub

    Private Sub cmdPopulateBLOB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPopulateBLOB.Click
        'this routine will populate the table tblWordDocsBlob

        Dim dtbl1 As DataTable
        Dim dtbl2 As DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strPath As String
        Dim id As Int64
        Dim intL As Int64
        Dim strpathT As String
        Dim Count1 As Int16
        Dim Count10 As Short
        Dim intMax As Int64

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String

        Dim con As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim strT As String



        'Dim fs As New FileStream(strPath, FileMode.Open, FileAccess.Read)

        Dim fs As FileStream

        con.CursorLocation = CursorLocationEnum.adUseServer
        con.Open(constrIni)
        cmd.ActiveConnection = con

        strSQL = "select TBLWORDDOCSBLOB.* from TBLWORDDOCSBLOB"
        rs.Open(strSQL, con, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockOptimistic, CommandTypeEnum.adCmdText)

        dtbl1 = frmH.tblWordStatements
        'dtbl2 = frmH.tblWorddocs
        strT = "tblWorddocsBLOB"

        Dim boolE As Boolean

        intMax = 0
        For Count10 = 0 To dtbl1.Rows.Count - 1

            Count1 = 0

            Me.lblStatus.Text = Count10 & " of " & dtbl1.Rows.Count - 1 & "..."
            Me.lblStatus.Refresh()

            strPath = dtbl1.Rows(Count10).Item("CHARWORDSTATEMENT")
            strPath = Replace(strPath, ".XML", ".DOC", 1, -1, CompareMethod.Text)
            strPath = Replace(strPath, "\XML\", "\DOC\", 1, -1, CompareMethod.Text)

            fs = New FileStream(strPath, FileMode.Open, FileAccess.Read)
            Dim b(fs.Length) As Byte

            fs.Read(b, 0, System.Convert.ToInt64(fs.Length))
            fs.Close()

            id = dtbl1.Rows(Count10).Item("ID_TBLWORDSTATEMENTS")

            boolE = True
            Count1 = 0

            intMax = intMax + 1
            Count1 = Count1 + 1

            rs.AddNew()
            rs.Fields("ID_TBLWORDDOCS").Value = intMax
            rs.Fields("ID_TBLWORDSTATEMENTS").Value = id
            rs.Fields("BLOBDOC").Value = b
            rs.Fields("UPSIZE_TS").Value = "01-SEP-07"
            rs.Update()


            'Dim newRow As DataRow = dtbl2.NewRow
            'newRow.BeginEdit()
            'newRow.Item("ID_TBLWORDDOCS") = intMax
            'newRow.Item("ID_TBLWORDSTATEMENTS") = id
            'newRow.Item("CHARXML") = RTrim(LTrim((strW)))
            'newRow.Item("UPSIZE_TS") = "01-SEP-07"
            'newRow.EndEdit()
            'dtbl2.Rows.Add(newRow)

            'ID_TBLWORDDOCS
            'ID_TBLWORDSTATEMENTS
            'CHARXML
            'UPSIZE_TS
            'str1 = "("
            'str1 = str1 & intMax & ", "
            'str1 = str1 & id & ", '"
            'str1 = str1 & strPath & "', '"
            ''str1 = Replace(str1, """", """""", 1, -1, CompareMethod.Text)
            ''str1 = str1 & "Hi', "
            'str1 = str1 & "01-SEP-07')"

            'str2 = "INSERT INTO " & strT & " (ID_TBLWORDDOCSBLOB, ID_TBLWORDSTATEMENTS, BLOBDOC, UPSIZE_TS)"
            'str3 = " VALUES " & str1 '(" & intMaxID & ", '" & str1 & "',1)"

            'strSQL = str2 & str3

            '''''''console.writeline(strSQL)

            'cmd.CommandText = strSQL
            'cmd.CommandType = CommandTypeEnum.adCmdText

            'cmd.Execute()





            'Try
            '    Try
            '        frmH.ta_tblWorddocs.Update(frmH.tblWorddocs)
            '    Catch ex As Exception
            '        frmH.ds2005.TBLWORDDOCS.Merge(frmH.ds2005.TBLWORDDOCS, True)
            '    End Try

            'Catch ex As DBConcurrencyException
            '    ''msgbox("aaContributing Personnel: " & ex.Message)
            '    frmH.ds2005.TBLWORDDOCS.Merge(frmH.ds2005.TBLWORDDOCS, True)
            'End Try

            Try
                fs.Close()
            Catch ex As Exception

            End Try


        Next

        Me.lblStatus.Text = "Status..."


        'Dim tblM As DataTable
        'Dim rowM() As DataRow
        'Dim strM As String

        'tblM = frmH.tblMaxID
        'strM = "CHARTABLE = 'tblworddocs'"
        'rowM = tblM.Select(strM)
        'rowM(0).BeginEdit()
        'rowM(0).Item("NUMMAXID") = intMax
        'rowM(0).EndEdit()
        'Try
        '    frmH.ta_tblMaxID.Update(frmH.tblMaxID)
        'Catch ex As DBConcurrencyException
        '    '''msgbox("aaData Tab: " & ex.Message)
        '    frmH.ds2005.TBLMAXID.Merge(frmH.ds2005.TBLMAXID, True)
        'End Try

        'sw.Close()
        'sw.Dispose()

        fs.Close()

        'fs1.Dispose()

        'write to file
        'Me.wb1.Navigate(strpathT)

    End Sub
End Class