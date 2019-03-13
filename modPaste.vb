Option Compare Text

Imports System
Imports System.IO
Imports System.Text
Module modPaste

    Function OpenTemplate(ByVal idRT As Int64, ByVal strPathWd As String) As String

        OpenTemplate = strPathWd

        Dim var1, var2
        Dim Count2 As Short
        Dim intStCount As Short
        Dim int1 As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim pos1, pos2
        Dim wtbl As Microsoft.Office.Interop.Word.Table
        Dim boolIn As Boolean
        Dim intCt As Short
        Dim intMv As Short
        Dim len1 As Int64

        boolIn = True

        'generate xml document
        Dim strW As String
        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intL As Short
        Dim Count1 As Short
        Dim fs As FileStream
        Dim strPathT As String
        Dim intV As Int64

        Try

            'must open this table
            'dtbl1 = tblWordStatementsVERSIONS
            'strF = "ID_TBLWORDSTATEMENTS = " & idRT
            'strS = "INTWORDVERSION DESC"
            'rows1 = dtbl1.Select(strF, strS)
            'intV = rows1(0).Item("INTWORDVERSION")

            intV = GetWordVersion(idRT, True)

            Call OpenWordDocs(idRT, intV)

            dtbl2 = tblWorddocs

            'strPath = NZ(dgv("CHARWORDSTATEMENT", intRow).Value, "") 'don't need
            'id = dgv("ID_TBLWORDSTATEMENTS", intRow).Value
            strF = "ID_TBLWORDSTATEMENTS = " & idRT
            strS = "ID_TBLWORDDOCS ASC"
            Erase rows2
            rows2 = dtbl2.Select(strF, strS)
            intL = rows2.Length
            strW = ""
            Dim strBuild As New StringBuilder("")
            For Count1 = 0 To intL - 1
                strBuild = strBuild.Append(rows2(Count1).Item("CHARXML"))
            Next
            strW = strBuild.ToString()

            Dim intA As Int64
            Dim intB As Int64
            intA = Len(strW)

            strW = RTrim(strW)
            intB = Len(strW)

            ' Add some information to the file.
            Dim info As Byte()
            info = New UTF8Encoding(True).GetBytes(strW)

            'strPathT = strPathWd & "_1.xml"
            'strPathWd = GetNewTempFile()'Hmmm strPathWd is global
            strPathT = strPathWd ' Replace(strPathWd, ".xml", "_1.xml", 1, -1, CompareMethod.Text)
            'strPathT = Replace(strPathWd, ".xml", "_1.doc", 1, -1, CompareMethod.Text)

            fs = File.Create(strPathT) 'this will overwrite existing file
            fs.Close()
            fs = File.OpenWrite(strPathT)
            fs.Write(info, 0, info.Length)
            fs.Close()

            OpenTemplate = strPathT

        Catch ex As Exception

        End Try



    End Function

    Sub PasteStatement(ByVal intCBS As Short, ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal boolGuWu As Boolean, ByVal boolStatement As Boolean, ByVal charS As String, ByVal charSectionName As String, ByVal intTNum As Short, ByVal idRT As Int64)

        Dim var1, var2
        Dim Count2 As Short
        Dim intStCount As Short
        Dim int1 As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.selection
        Dim pos1, pos2
        Dim wtbl As Microsoft.Office.Interop.Word.Table
        Dim boolIn As Boolean
        Dim intCt As Short
        Dim intMv As Short
        Dim len1 As Int64

        boolIn = True
        With wd

            'generate xml document
            Dim strW As String
            Dim dtbl2 As System.Data.Datatable
            Dim rows2() As DataRow
            Dim strF As String
            Dim strS As String
            Dim intL As Short
            Dim Count1 As Short
            Dim fs As FileStream
            Dim strPathT As String

            dtbl2 = tblWorddocs

            'strPath = NZ(dgv("CHARWORDSTATEMENT", intRow).Value, "") 'don't need
            'id = dgv("ID_TBLWORDSTATEMENTS", intRow).Value
            strF = "ID_TBLWORDSTATEMENTS = " & idRT
            strS = "ID_TBLWORDDOCS ASC"
            Erase rows2
            rows2 = dtbl2.Select(strF, strS)
            intL = rows2.Length

            Dim strBuild = New StringBuilder("")
            For Count1 = 0 To intL - 1
                strBuild.Append(rows2(Count1).Item("CHARXML"))
            Next
            strW = strBuild.ToString()

            Dim intA As Int64
            Dim intB As Int64
            intA = Len(strW)

            strW = RTrim(strW)
            intB = Len(strW)

            ' Add some information to the file.
            Dim info As Byte()
            info = New UTF8Encoding(True).GetBytes(strW)

            'strPathT = strPathWd & "_1.xml"
            'strPathWd = GetNewTempFile()'Hmmm strPathWd is global
            strPathT = Replace(strPathWd, ".xml", "_1.xml", 1, -1, CompareMethod.Text)
            'strPathT = Replace(strPathWd, ".xml", "_1.doc", 1, -1, CompareMethod.Text)

            fs = File.Create(strPathT) 'this will overwrite existing file
            fs.Close()
            fs = File.OpenWrite(strPathT)
            fs.Write(info, 0, info.Length)
            fs.Close()

            'add Temp1 bookmark
            wrdSelection = wd.Selection()
            With .ActiveDocument.Bookmarks
                .Add(Range:=wrdSelection.Range, Name:="Temp1")
                .ShowHidden = False
            End With
            pos1 = .Selection.Start

            '.Selection.InsertFile(FileName:=strPathT, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False)

            'wd.Visible = True

            If boolEntireReport Then

                .Selection.InsertFile(FileName:=strPathT, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False)
                '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                Try
                    frmH.Activate()
                    'wd.visible = False
                Catch ex As Exception

                End Try

            Else

                'wdSt.Documents.Open(strPathT)
                'wdSt.Selection.WholeStory()
                'wrdSelection = wdSt.Selection
                ''var1 = Len(wrdSelection.Characters.Count)
                'var1 = wrdSelection.Characters.Count

                ''wtbl.Cell(Count2, 2).Select()
                ''var1 = wtbl.Cell(Count2, 2).Range.Text

                'wdSt.Selection.Copy()

                ''wdSt.visible = True

                ' ''add Temp1 bookmark
                ''wrdSelection = wd.Selection()
                ''With .ActiveDocument.Bookmarks
                ''    .Add(Range:=wrdSelection.Range, Name:="Temp1")
                ''    .ShowHidden = False
                ''End With
                ''pos1 = .Selection.Start

                ''.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                ''If Len(var1) = 2 Then 'empty cell
                'If var1 < 2 Then 'empty cell
                'Else
                '    '.Selection.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault)
                '    'the commented line below was from word 2003
                '    '.Selection.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)
                '    .Selection.Paste()
                '    intCt = 0

                'End If

                'pos2 = .Selection.Start

                'If pos2 - pos1 < 2 Then 'empty document
                '    .Selection.TypeBackspace()
                'End If

                ''Try
                ''    wdSt.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)

                ''Catch ex As Exception

                ''End Try

                'Try
                '    frmH.Activate()
                '    'wd.visible = False
                'Catch ex As Exception

                'End Try

                ''determine if an extra paragraph return needs to be inserted
                '.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=True)
                'var1 = .Selection.Text
                'var2 = Asc(var1)
                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                'If var2 = 13 Then
                'Else
                '    .Selection.TypeParagraph()
                'End If

                ''add Temp2 bookmark
                'wrdSelection = wd.Selection()
                'With .ActiveDocument.Bookmarks
                '    .Add(Range:=wrdSelection.Range, Name:="Temp2")
                '    .ShowHidden = False
                'End With


                'If pos2 - pos1 < 5 Then
                'Else
                '    '.ActiveDocument.Bookmarks.item("Temp1").Delete()
                '    '.ActiveDocument.Bookmarks.item("Temp2").Delete()
                '    Try
                '        .ActiveDocument.Bookmarks.Item("Temp1").Delete()
                '    Catch ex As Exception

                '    End Try
                '    Try
                '        .ActiveDocument.Bookmarks.Item("Temp2").Delete()
                '    Catch ex As Exception

                '    End Try
                'End If

            End If


            'End If
        End With


    End Sub

    Sub InsertPageNumber(ByVal wd As Microsoft.Office.Interop.Word.Application)
        Dim wdSel As Microsoft.Office.Interop.Word.Selection
        With wd
            'insert page number over found selection
            wdSel = .Selection
            .Selection.Fields.Add(Range:=wdSel.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
        End With
    End Sub

    Sub InsertTotalPages(ByVal wd As Microsoft.Office.Interop.Word.Application)
        Dim wdSel As Microsoft.Office.Interop.Word.Selection
        With wd
            'insert total pages over found selection
            wdSel = .Selection
            .Selection.Fields.Add(Range:=wdSel.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
        End With
    End Sub

End Module
