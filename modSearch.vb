Option Compare Text

Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO

Module modSearch

    Sub SearchRepeat(ByRef wd As Microsoft.Office.Interop.Word.Application)


        '20180827 LEE:
        'Word is crashing if Repeat Sections are embedded in complicated Word table structures
        'try puttin in normal (draft) view

        Try
            wd.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView
        Catch ex As Exception

        End Try

        Try
            Dim wrdSel As Word.Selection
            Dim strFind1 As String
            Dim strFind2 As String
            Dim str1 As String
            Dim str2 As String
            Dim str3 As String
            Dim strFCID As String

            Dim var1

            Dim Count1 As Integer
            Dim Count2 As Integer

            Dim intA As Integer
            Dim strA1 As String
            Dim strA2 As String
            Dim strAnal As String

            '20181128 LEE:
            Dim strIS1 As String
            Dim strIS2 As String

            Dim strF As String
            Dim strS As String

            'Dim arrAnalytes(16, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
            '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
            '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX

            strF = "IsIntStd = 'No'"
            strS = "INTORDER ASC"
            Dim rowsA() As System.Data.DataRow = tblAnalytesHome.Select(strF, strS)
            intA = rowsA.Length

            strA1 = "Analyte_1"
            strIS1 = "IS_Analyte_1"

            Dim boolQuit As Boolean

            Dim rng1 As Word.Range
            Dim rng2 As Word.Range
            Dim rng3 As Word.Range
            Dim rng4 As Word.Range

            Dim pos1 As Int64
            Dim pos2 As Int64
            Dim pos3 As Int64
            Dim pos4 As Int64
            Dim pos5 As Int64
            Dim pos6 As Int64

            Dim intChar As Int64

            Dim strIsIS As String

            Dim intDo As Int16

            Dim intStartA As Short = 2

            strFind1 = "[RPTS]"

            strFind2 = "[RPTE]"

            boolQuit = False

            Dim boolF As Boolean


            Dim intCtA As Short = 0

            'first must check to see if Analyte_1 is first used cmpd
            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            rng1 = wd.Selection.Range
            rng1.WholeStory()
            For Count1 = 1 To intA

                strAnal = rowsA(Count1 - 1).Item("AnalyteDescription")

                strA2 = "Analyte_" & Count1
                'ensure analyte is used
                If UseAnalyte(strAnal) Then
                    If Count1 = 1 Then
                    Else
                        'replace _1] with _Count1]
                        With rng1.Find
                            .ClearFormatting()
                            .MatchWholeWord = False
                            .Forward = True
                            .Wrap = Word.WdFindWrap.wdFindStop
                            '.Text = "_1]"
                            .Text = "Analyte_1]"
                            .Replacement.ClearFormatting()
                            '.Replacement.Text = "_" & Count1 & "]"
                            .Replacement.Text = "Analyte_" & Count1 & "]"
                            boolF = .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                        End With

                        'replace _1_ with _Count1_
                        With rng1.Find
                            .ClearFormatting()
                            .MatchWholeWord = False
                            .Forward = True
                            .Wrap = Word.WdFindWrap.wdFindStop
                            .Text = "Analyte_1_"
                            .Replacement.ClearFormatting()
                            .Replacement.Text = "Analyte_" & Count1 & "_"
                            boolF = .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                        End With
                        intStartA = Count1 + 1
                    End If
                    Exit For
                End If

            Next Count1

            'check for number of actual cmpds used
            For Count1 = 1 To intA

                strAnal = rowsA(Count1 - 1).Item("AnalyteDescription")
                'ensure analyte is used
                If UseAnalyte(strAnal) Then
                    intCtA = intCtA + 1
                End If

            Next


            If intCtA = 1 Then
                'remove all instances of RPTS and RPTE
                With rng1.Find
                    .ClearFormatting()
                    .MatchWholeWord = False
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindStop
                    .Text = "[RPTS]"
                    .Replacement.ClearFormatting()
                    .Replacement.Text = ""
                    boolF = .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                End With

                'replace _1_ with _Count1_
                With rng1.Find
                    .ClearFormatting()
                    .MatchWholeWord = False
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindStop
                    .Text = "[RPTE]"
                    .Replacement.ClearFormatting()
                    .Replacement.Text = ""
                    boolF = .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                End With

                GoTo end1

            End If

            Try

                With wd

                    .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                    'wd.Visible = True

                    Do Until boolQuit

                        '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                        wrdSel = .Selection

                        With wrdSel.Find
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .Forward = True
                            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                            .Execute(FindText:=strFind1)
                            If .Found Then

                                'for some reason, if cursor is within a table and is the first entry of the table, a space is added when item is deleted
                                'record next character
                                pos5 = wd.Selection.End
                                rng1 = wd.ActiveDocument.Range(Start:=pos5, End:=pos5)
                                'get character of this range
                                var1 = rng1.Characters(1).Text

                                'delete it
                                wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                                'for some reason, if the following character is '[', a space is added when item is deleted
                                If StrComp(var1, "[", vbTextCompare) = 0 Then
                                    wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                                End If

                                'record the position
                                pos1 = wd.Selection.Start

                                'now find end
                                pos2 = 0
                                With wd.Selection.Find

                                    .ClearFormatting()
                                    .MatchWholeWord = True
                                    .Forward = True
                                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                                    .Execute(FindText:=strFind2)
                                    If .Found Then

                                        'delete it
                                        wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                                        'record the position
                                        pos2 = wd.Selection.Start

                                    End If

                                End With

                                If pos2 = 0 Then
                                    Exit Do
                                End If

                                'If intA = 1 Then
                                If intStartA > intA Then

                                    boolQuit = True
                                Else

                                    'set a range
                                    rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                    intChar = rng1.Characters.Count




                                    intDo = 0
                                    Dim INTGROUP As Short

                                    'For Count1 = intStartA To intA
                                    For Count1 = 1 To intA 'must start 1 and check each individual item

                                        strAnal = rowsA(Count1 - 1).Item("AnalyteDescription") 'this has _C1, matrix
                                        INTGROUP = rowsA(Count1 - 1).Item("INTGROUP")

                                        ''20180815 LEE:
                                        ''need to evaluate per table fcid
                                        'rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                        'strFCID = ReturnFCID(rng1)
                                        'rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                        ''see if analyte is included


                                        strA2 = "Analyte_" & Count1
                                        strIS2 = "IS_Analyte_" & Count1

                                        '20180815 LEE:
                                        'need to determine if analyte is used in THIS table in rng1
                                        '20180815 LEE:
                                        'must reset rng1
                                        rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                        If IsAnalInTable(rng1, INTGROUP, strAnal) Then
                                            intDo = intDo + 1

                                            '20180815 LEE:
                                            'must reset rng1
                                            rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)

                                            If intDo = 1 Then
                                                'replace Analyte_1 with actual
                                                Dim strF1 As String = "Analyte_1"
                                                rng1.Find.Execute(FindText:=strF1, ReplaceWith:=strA2, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                                                'reset strA
                                                strA1 = strA2
                                            Else
                                                '20180815 LEE:
                                                'must reset rng1
                                                rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)

                                                'enter a paragraph return
                                                wd.Selection.TypeParagraph()
                                                pos3 = wd.Selection.Start
                                                rng3 = wd.ActiveDocument.Range(Start:=pos3, End:=pos3)

                                                'copy the range
                                                rng1.Copy()

                                                rng3.Select()

                                                'paste the range
                                                rng3.Paste()

                                                pos4 = pos3 + intChar

                                                rng4 = wd.ActiveDocument.Range(Start:=pos3, End:=pos4 - 1)

                                                'now replace
                                                rng4.Find.Execute(FindText:=strA1, ReplaceWith:=strA2, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                                                'move to end of range
                                                rng4 = wd.ActiveDocument.Range(Start:=pos4, End:=pos4)
                                                rng4.Select()

                                                var1 = var1
                                            End If

                                            '20180815 LEE:
                                            'must reset rng1
                                            rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)

                                            '20181128 LEE:
                                            'now do IS_Analyte_1
                                            '"IS_Analyte_"
                                            If intDo = 1 Then
                                                'replace Analyte_1 with actual
                                                Dim strF1 As String = "IS_Analyte_1"
                                                rng1.Find.Execute(FindText:=strF1, ReplaceWith:=strIS2, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                                                'reset strA
                                                strIS1 = strIS2
                                            Else
                                                '20180815 LEE:
                                                'must reset rng1
                                                rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)

                                                rng4 = wd.ActiveDocument.Range(Start:=pos3, End:=pos4 - 1)

                                                'now replace
                                                rng4.Find.Execute(FindText:=strIS1, ReplaceWith:=strIS2, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                                                'move to end of range
                                                rng4 = wd.ActiveDocument.Range(Start:=pos4, End:=pos4)
                                                rng4.Select()

                                                var1 = var1
                                            End If


                                        End If

                                        var1 = var1

                                    Next Count1

                                    var1 = var1

                                End If

                            Else
                                boolQuit = True
                            End If

                        End With

                    Loop

                End With

            Catch ex As Exception
                var1 = ex.Message
            End Try

end1:
        Catch ex As Exception

        End Try

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        Try
            wd.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception

        End Try

    End Sub

    Sub SearchRepeatBU(ByRef wd As Microsoft.Office.Interop.Word.Application)


        '20180827 LEE:
        'Word is crashing if Repeat Sections are embedded in complicated Word table structures
        'try puttin in normal (draft) view

        Try
            wd.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView
        Catch ex As Exception

        End Try

        Try
            Dim wrdSel As Word.Selection
            Dim strFind1 As String
            Dim strFind2 As String
            Dim str1 As String
            Dim str2 As String
            Dim str3 As String
            Dim strFCID As String

            Dim var1

            Dim Count1 As Integer
            Dim Count2 As Integer

            Dim intA As Integer
            Dim strA1 As String
            Dim strA2 As String
            Dim strAnal As String
            Dim rngNew As Word.Range

            Dim strF As String
            Dim strS As String

            'Dim arrAnalytes(16, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
            '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
            '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX

            strF = "IsIntStd = 'No'"
            strS = "INTORDER ASC"
            Dim rowsA() As System.Data.DataRow = tblAnalytesHome.Select(strF, strS)
            intA = rowsA.Length

            strA1 = "Analyte_1"

            Dim boolQuit As Boolean

            Dim rng1 As Word.Range
            Dim rng2 As Word.Range
            Dim rng3 As Word.Range
            Dim rng4 As Word.Range
            Dim c As Word.Range

            Dim pos1 As Int64
            Dim pos2 As Int64
            Dim pos3 As Int64
            Dim pos4 As Int64
            Dim pos5 As Int64
            Dim pos6 As Int64

            Dim INTGROUP As Short
            Dim strUserIS As String
            Dim strUserAnal As String

            Dim intChar As Int64

            Dim strIsIS As String

            Dim intDo As Int16

            Dim intStartA As Short = 2

            strFind1 = "[RPTS]"

            strFind2 = "[RPTE]"

            boolQuit = False

            Dim boolF As Boolean
            Dim strText As String

            Dim intIA As Short
            Dim intIIS As Short


            Dim intCtA As Short = 0

            'first must check to see if Analyte_1 is first used cmpd
            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            rng1 = wd.Selection.Range
            rng1.WholeStory()
            For Count1 = 1 To intA

                strAnal = rowsA(Count1 - 1).Item("AnalyteDescription")

                strA2 = "Analyte_" & Count1
                'ensure analyte is used
                If UseAnalyte(strAnal) Then
                    If Count1 = 1 Then
                    Else
                        'replace _1] with _Count1]
                        With rng1.Find
                            .ClearFormatting()
                            .MatchWholeWord = False
                            .Forward = True
                            .Wrap = Word.WdFindWrap.wdFindStop
                            '.Text = "_1]"
                            .Text = "Analyte_1]"
                            .Replacement.ClearFormatting()
                            '.Replacement.Text = "_" & Count1 & "]"
                            .Replacement.Text = "Analyte_" & Count1 & "]"
                            boolF = .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                        End With

                        'replace _1_ with _Count1_
                        With rng1.Find
                            .ClearFormatting()
                            .MatchWholeWord = False
                            .Forward = True
                            .Wrap = Word.WdFindWrap.wdFindStop
                            .Text = "Analyte_1_"
                            .Replacement.ClearFormatting()
                            .Replacement.Text = "Analyte_" & Count1 & "_"
                            boolF = .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                        End With
                        intStartA = Count1 + 1
                    End If
                    Exit For
                End If

            Next Count1

            'check for number of actual cmpds used
            For Count1 = 1 To intA

                strAnal = rowsA(Count1 - 1).Item("AnalyteDescription")
                'ensure analyte is used
                If UseAnalyte(strAnal) Then
                    intCtA = intCtA + 1
                End If

            Next


            If intCtA = 1 Then
                'remove all instances of RPTS and RPTE
                With rng1.Find
                    .ClearFormatting()
                    .MatchWholeWord = False
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindStop
                    .Text = "[RPTS]"
                    .Replacement.ClearFormatting()
                    .Replacement.Text = ""
                    boolF = .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                End With

                'replace _1_ with _Count1_
                With rng1.Find
                    .ClearFormatting()
                    .MatchWholeWord = False
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindStop
                    .Text = "[RPTE]"
                    .Replacement.ClearFormatting()
                    .Replacement.Text = ""
                    boolF = .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                End With

                GoTo end1

            End If

            Try

                With wd

                    .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                    'wd.Visible = True

                    Do Until boolQuit

                        '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                        wrdSel = .Selection

                        With wrdSel.Find
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .Forward = True
                            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                            .Execute(FindText:=strFind1)
                            If .Found Then

                                'for some reason, if cursor is within a table and is the first entry of the table, a space is added when item is deleted
                                'record next character
                                pos5 = wd.Selection.End
                                rng1 = wd.ActiveDocument.Range(Start:=pos5, End:=pos5)
                                'get character of this range
                                var1 = rng1.Characters(1).Text

                                'delete it
                                wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                                'for some reason, if the following character is '[', a space is added when item is deleted
                                If StrComp(var1, "[", vbTextCompare) = 0 Then
                                    wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                                End If

                                'record the position
                                pos1 = wd.Selection.Start

                                'now find end
                                pos2 = 0
                                With wd.Selection.Find

                                    .ClearFormatting()
                                    .MatchWholeWord = True
                                    .Forward = True
                                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                                    .Execute(FindText:=strFind2)
                                    If .Found Then

                                        'delete it
                                        wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                                        'record the position
                                        pos2 = wd.Selection.Start

                                    End If

                                End With

                                If pos2 = 0 Then
                                    Exit Do
                                End If

                                'If intA = 1 Then
                                If intStartA > intA Then

                                    boolQuit = True
                                Else

                                    'set a range
                                    rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                    intChar = rng1.Characters.Count
                                    strText = rng1.Text

                                    intDo = 0


                                    'For Count1 = intStartA To intA
                                    For Count1 = 1 To intA 'must start 1 and check each individual item

                                        strAnal = rowsA(Count1 - 1).Item("AnalyteDescription") 'this has _C1, matrix
                                        INTGROUP = rowsA(Count1 - 1).Item("INTGROUP")
                                        strUserIS = NZ(rowsA(Count1 - 1).Item("CHARUSERIS"), NZ(rowsA(Count1 - 1).Item("IntStd"), "IntStd"))
                                        strUserAnal = NZ(rowsA(Count1 - 1).Item("CHARUSERANALYTE"), strAnal)

                                        ''20180815 LEE:
                                        ''need to evaluate per table fcid
                                        'rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                        'strFCID = ReturnFCID(rng1)
                                        'rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                        ''see if analyte is included


                                        strA2 = "Analyte_" & Count1

                                        '20180815 LEE:
                                        'need to determine if analyte is used in THIS table in rng1
                                        '20180815 LEE:
                                        'must reset rng1
                                        rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                        If IsAnalInTable(rng1, INTGROUP, strAnal) Then
                                            intDo = intDo + 1

                                            '20180815 LEE:
                                            'must reset rng1
                                            rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)

                                            If intDo = 1 Then

                                                Dim strF1 As String
                                                strF1 = "[Analyte_1]"

                                                intIA = 0
                                                '20181112 LEE:
                                                'Count number times strF1 appears in rng1
                                                rng2 = wd.ActiveDocument.Range(Start:=pos1, End:=pos1)
                                                rng2.Select()

                                                With wd.Selection.Find
                                                    .ClearFormatting()
                                                    .Text = strF1
                                                    ' Loop until Word can no longer
                                                    ' find the search string and
                                                    ' count each instance
                                                    Do While .Execute
                                                        pos3 = wd.Selection.End
                                                        If pos3 > pos2 Or pos3 <= pos1 Then
                                                            Exit Do
                                                        Else
                                                            intIA = intIA + 1
                                                        End If
                                                        wd.Selection.MoveRight()
                                                    Loop
                                                End With
                                                'go back
                                                rng2 = wd.ActiveDocument.Range(Start:=pos1, End:=pos1)
                                                rng2.Select()


                                                '******

                                                'replace Analyte_1 with actual

                                                'strF1 = "Analyte_1"
                                                'rng1.Find.Execute(FindText:=strF1, ReplaceWith:=strA2, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                                                '20181108
                                                'just replace with struseranal

                                                rng1.Find.Execute(FindText:=strF1, ReplaceWith:=strUserAnal, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                                                'reset strA
                                                'rng1.Select() 'debug
                                                strA1 = strA2

                                                '20181112 LEE:
                                                'Hmmm. There may be more than one replacement
                                                pos2 = pos2 - (Len(strF1) * intIA) + (Len(strUserAnal) * intIA)

                                                'debug
                                                rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                                rng1.Select() 'debug
                                                var1 = var1


                                            Else
                                                '20180815 LEE:
                                                'must reset rng1
                                                'rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                                rng1 = wd.ActiveDocument.Range(Start:=pos2, End:=pos2)
                                                rng1.Select()

                                                'enter a paragraph return
                                                wd.Selection.TypeParagraph()

                                                pos1 = wd.Selection.Start

                                                'pos3 = wd.Selection.Start
                                                'rng3 = wd.ActiveDocument.Range(Start:=pos3, End:=pos3)
                                                '20181108 LEE:
                                                'just paste strText
                                                wd.Selection.Text = strText

                                                ''copy the range
                                                'rng1.Copy()

                                                'rng3.Select()

                                                ''paste the range
                                                'rng3.Paste()

                                                pos2 = pos1 + intChar

                                                rng4 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)

                                                'now replace
                                                'rng4.Find.Execute(FindText:=strA1, ReplaceWith:=strA2, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                                                '20181108
                                                'just replace with struseranal
                                                str1 = "[" & strA1 & "]"
                                                rng4.Find.Execute(FindText:=str1, ReplaceWith:=strUserAnal, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                                                'reset pos2
                                                If intIA = 0 Then
                                                    intIA = 1
                                                End If
                                                pos3 = pos2 - (Len(str1) * intIA) + (Len(strUserAnal) * intIA)
                                                pos2 = pos3

                                            End If

                                            '20180815 LEE:
                                            'must reset rng1
                                            ''20181108 LEE:
                                            ''and reset pos2
                                            'pos3 = pos2 - (Len("[ANALYTE]") + 2) + Len(strUserAnal)
                                            'pos2 = pos3
                                            rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                            var1 = var1 'debug

                                            '20181108 LEE:
                                            'now look to see if [INTERNALSTANDARD] exists
                                            str1 = "[INTERNALSTANDARD_1]"

                                            intIIS = 0
                                            '20181112 LEE:
                                            'Count number times strF1 appears in rng1
                                            rng2 = wd.ActiveDocument.Range(Start:=pos1, End:=pos1)
                                            rng2.Select()
                                            'With wd.Selection.Find
                                            '    .ClearFormatting()
                                            '    .Text = str1
                                            '    ' Loop until Word can no longer
                                            '    ' find the search string and
                                            '    ' count each instance
                                            With wd.Selection.Find
                                                .ClearFormatting()
                                                .Text = str1
                                                ' Loop until Word can no longer
                                                ' find the search string and
                                                ' count each instance
                                                Do While .Execute
                                                    pos3 = wd.Selection.End
                                                    If pos3 > pos2 Or pos3 <= pos1 Then
                                                        Exit Do
                                                    Else
                                                        intIIS = intIIS + 1
                                                    End If
                                                    wd.Selection.MoveRight()
                                                Loop
                                            End With
                                            'go back
                                            rng2 = wd.ActiveDocument.Range(Start:=pos1, End:=pos1)
                                            rng2.Select()

                                            'End With
                                            If rng1.Find.Execute(FindText:=str1) Then
                                                'find Internal standard
                                                'must reset rng1
                                                rng4 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                                rng4.Find.Execute(FindText:=str1, ReplaceWith:=strUserIS, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                                                'and reset pos2
                                                If intIIS = 0 Then
                                                    intIIS = 1
                                                End If
                                                pos3 = pos2 - (Len(str1) * intIIS) + (Len(strUserIS) * intIIS)
                                                pos2 = pos3

                                            End If


                                            rng1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
                                            var1 = var1 'debug

                                        End If

                                        var1 = var1

                                    Next Count1

                                    var1 = var1

                                End If

                            Else
                                boolQuit = True
                            End If

                        End With

                    Loop

                End With

            Catch ex As Exception
                var1 = ex.Message
            End Try

end1:
        Catch ex As Exception

        End Try

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        Try
            wd.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception

        End Try

    End Sub

    Sub SearchReplaceCustomFieldCode(ByRef wd As Microsoft.Office.Interop.Word.Application)

        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short

        Dim strFind As String
        Dim varReplace
        Dim dv1 As System.Data.DataView
        Dim int1 As Short
        Dim var8
        Dim strNA1 As String
        Dim strNA2 As String
        Dim strNA3 As String
        Dim strNA4 As String
        Dim charSectionName As String
        Dim mysel As Microsoft.Office.Interop.Word.Selection
        Dim boolNum As Boolean
        Dim num1
        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim strF1 As String
        Dim strF2 As String
        Dim strS1 As String
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim id1 As Int64

        dtbl1 = tblCustomFieldCodes
        dtbl2 = tblFieldCodes

        strF1 = "ID_TBLSTUDIES = " & id_tblStudies
        rows1 = dtbl1.Select(strF1)

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        mysel = wd.Selection

        For Count1 = 0 To rows1.Length - 1

            id1 = NZ(rows1(Count1).Item("ID_TBLFIELDCODES"), -1)
            strF2 = "ID_TBLFIELDCODES = " & id1
            rows2 = dtbl2.Select(strF2)

            If rows2.Length = 0 Then
            Else

                strFind = NZ(rows2(0).Item("CHARFIELDCODE"), "AAAAAA")
                varReplace = NZ(rows1(Count1).Item("CHARVALUE"), "[NA]")

                strNA1 = "Custom Field Codes"
                strNA2 = strFind
                strNA3 = "Custom Field Codes"
                strNA4 = "Data Tab - Custom Field Codes"

                Try
                    If Len(strFind) = 0 Then
                    Else
                        With wd
                            With mysel.Find
                                .ClearFormatting()
                                .Text = strFind
                                With .Replacement
                                    .ClearFormatting()
                                    If boolNum Then
                                        .Text = num1
                                    Else
                                        .Text = varReplace
                                    End If
                                End With
                                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                                If .Found Then
                                    If Len(varReplace) = 0 Or StrComp(NZ(varReplace, "[None]"), "[None]", CompareMethod.Text) = 0 Or StrComp(varReplace, "[NA]", CompareMethod.Text) = 0 Then
                                        'add entries to arrReportNA
                                        ctArrReportNA = ctArrReportNA + 1
                                        arrReportNA(1, ctArrReportNA) = strNA1
                                        arrReportNA(2, ctArrReportNA) = strNA2
                                        arrReportNA(3, ctArrReportNA) = strNA3
                                        arrReportNA(4, ctArrReportNA) = strNA4
                                    End If
                                End If
                            End With
                        End With
                    End If
                Catch ex As Exception
                    Dim aaa
                    aaa = 1
                End Try

            End If

        Next

        mysel = wd.Selection

        'now look for any [NA]s
        With mysel.Find
            .ClearFormatting()
            .Text = "[NA]"
            With .Replacement
                .ClearFormatting()
                .Text = "[NA]"
                .Font.Bold = True
                .Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
            End With
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        End With

    End Sub

    Sub SearchReplaceTableI(ByVal wd As Microsoft.Office.Interop.Word.application)

        'Dim Count1 As Short
        'Dim strFind As String
        'Dim varReplace
        'Dim dv1 as system.data.dataview
        'Dim int1 As Short
        'Dim var8
        'Dim strNA1 As String
        'Dim strNA2 As String
        'Dim strNA3 As String
        'Dim strNA4 As String
        'Dim charSectionName As String
        'Dim mysel As Microsoft.Office.Interop.Word.selection
        'Dim boolNum As Boolean
        'Dim num1

        'charSectionName = "Data Tab"
        '''''''wdd.visible = True

        'dv1 = frmH.dgvWatsonAnalRef.DataSource 'intI2 is analyte column in dgvWatsonAnalRef


        'mysel = wd.Selection

        'For Count1 = 1 To 10

        '    boolNum = False
        '    strFind = ""
        '    Select Case Count1

        '        Case 1
        '            strFind = "[ANALYTE1]"
        '            'var8 = arrAnalytes(1, intI)
        '            var8 = strAnal
        '            varReplace = var8

        '        Case 2
        '            strFind = "[LLOQ]"
        '            int1 = FindRowDVByCol("LLOQ", dv1, "Item")
        '            'var8 = dg.Item(int1, 1)
        '            num1 = dv1.Item(int1).Item(strAnal)
        '            num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
        '            boolNum = True
        '            var8 = num1
        '            If IsDBNull(var8) Then
        '                varReplace = "[NA]"
        '            ElseIf Len(var8) = 0 Then
        '                varReplace = "[NA]"
        '            Else
        '                varReplace = var8
        '            End If
        '            'ctArrReportNA = ctArrReportNA + 1
        '            strNA1 = "Analytical Reference Standard"
        '            strNA2 = "LLOQ"
        '            strNA3 = "Analytical Reference Standard"
        '            strNA4 = "Watson Analytical Reference Standard Table"

        '        Case 3
        '            strFind = "[LLOQUNITS]"
        '            int1 = FindRowDVByCol("LLOQ Units", dv1, "Item")
        '            var8 = dv1.Item(int1).Item(strAnal)
        '            If IsDBNull(var8) Then
        '                varReplace = "[NA]"
        '            ElseIf Len(var8) = 0 Then
        '                varReplace = "[NA]"
        '            Else
        '                varReplace = var8
        '            End If
        '            'ctArrReportNA = ctArrReportNA + 1
        '            strNA1 = "Analytical Reference Standard"
        '            strNA2 = "LLOQ Units"
        '            strNA3 = "Analytical Reference Standard"
        '            strNA4 = "Watson Analytical Reference Standard Table"

        '        Case 4
        '            strFind = "[ULOQ]"
        '            int1 = FindRowDVByCol("ULOQ", dv1, "Item")
        '            num1 = dv1.Item(int1).Item(strAnal)
        '            num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
        '            boolNum = True
        '            var8 = num1
        '            If IsDBNull(var8) Then
        '                varReplace = "[NA]"
        '            ElseIf Len(var8) = 0 Then
        '                varReplace = "[NA]"
        '            Else
        '                varReplace = var8
        '            End If
        '            'ctArrReportNA = ctArrReportNA + 1
        '            strNA1 = "Analytical Reference Standard"
        '            strNA2 = "ULOQ"
        '            strNA3 = "Analytical Reference Standard"
        '            strNA4 = "Watson Analytical Reference Standard Table"

        '        Case 5
        '            strFind = "[ULOQUNITS]"
        '            int1 = FindRowDVByCol("ULOQ Units", dv1, "Item")
        '            var8 = dv1.Item(int1).Item(strAnal)
        '            If IsDBNull(var8) Then
        '                varReplace = "[NA]"
        '            ElseIf Len(var8) = 0 Then
        '                varReplace = "[NA]"
        '            Else
        '                varReplace = var8
        '            End If
        '            'ctArrReportNA = ctArrReportNA + 1
        '            strNA1 = "Analytical Reference Standard"
        '            strNA2 = "ULOQ Units"
        '            strNA3 = "Analytical Reference Standard"
        '            strNA4 = "Watson Analytical Reference Standard Table"

        '        Case 6
        '            strFind = "[SAMPLESIZE]"
        '            varReplace = ReturnSearch(wd, "Temp", 39, strFind, mysel.Range)

        '        Case 7
        '            strFind = "[SAMPLESIZEUNITS]"
        '            varReplace = ReturnSearch(wd, "Temp", 47, strFind, mysel.Range)

        '        Case 8
        '            strFind = "[ANTICOAGULANT]"
        '            varReplace = ReturnSearch(wd, "Temp", 46, strFind, mysel.Range)

        '        Case 9
        '            strFind = "[MATRIX]"
        '            varReplace = ReturnSearch(wd, "Temp", 38, strFind, mysel.Range)

        '        Case 10
        '            strFind = "[SPECIES]"
        '            varReplace = ReturnSearch(wd, "Temp", 40, strFind, mysel.Range)

        '    End Select

        '    If Len(strFind) = 0 Then
        '    Else
        '        With wd
        '            With mysel.Find
        '                .ClearFormatting()
        '                .Text = strFind
        '                With .Replacement
        '                    .ClearFormatting()
        '                    If boolNum Then
        '                        .Text = num1
        '                    Else
        '                        .Text = varReplace
        '                    End If
        '                End With
        '                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        '                If .Found Then
        '                    If Len(varReplace) = 0 Or StrComp(NZ(varReplace, "[None]"), "[None]", CompareMethod.Text) = 0 Or StrComp(varReplace, "[NA]", CompareMethod.Text) = 0 Then
        '                        'add entries to arrReportNA
        '                        ctArrReportNA = ctArrReportNA + 1
        '                        arrReportNA(1, ctArrReportNA) = strNA1
        '                        arrReportNA(2, ctArrReportNA) = strNA2
        '                        arrReportNA(3, ctArrReportNA) = strNA3
        '                        arrReportNA(4, ctArrReportNA) = strNA4
        '                    End If
        '                End If
        '            End With
        '        End With
        '    End If
        'Next
        ''now look for any [NA]s
        'With mysel.Find
        '    .ClearFormatting()
        '    .Text = "[NA]"
        '    With .Replacement
        '        .ClearFormatting()
        '        .Text = "[NA]"
        '        .Font.Bold = True
        '        .Font.Color = Microsoft.Office.Interop.Word.wdcolor.wdColorRed
        '    End With
        '    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        'End With

    End Sub

    '"[ANALYTE] from [LLOQ] [LLOQUNITS] to [ULOQ] [ULOQUNITS]
    Sub SearchReplaceAnalLLOQ(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strAnal As String, boolDoSS As Boolean)

        Dim boolAcro As Boolean = False
        Dim Count1 As Short
        Dim strFind As String
        Dim varReplace
        Dim dv1 As System.Data.DataView
        Dim int1 As Short
        Dim var8
        Dim strNA1 As String
        Dim strNA2 As String
        Dim strNA3 As String
        Dim strNA4 As String
        Dim charSectionName As String
        Dim mysel As Microsoft.Office.Interop.Word.Selection
        Dim boolNum As Boolean
        Dim num1
        Dim str1 As String

        Dim boolDo As Boolean

        charSectionName = "Data Tab"
        ''''''wdd.visible = True

        dv1 = frmH.dgvWatsonAnalRef.DataSource 'intI2 is analyte column in dgvWatsonAnalRef


        mysel = wd.selection

        For Count1 = 1 To 16

            boolNum = False
            strFind = ""
            boolDo = True
            Select Case Count1

                Case 1
                    strFind = "[ANALYTE1]"
                    'var8 = arrAnalytes(1, intI)
                    var8 = strAnal

                    '20181015 LEE:
                    'put nbh back in
                    var8 = Replace(var8, "-", NBHReal, 1, -1, CompareMethod.Text)

                    varReplace = var8

                Case 2
                    strFind = "[LLOQ]"
                    int1 = FindRowDVByCol("LLOQ", dv1, "Item")
                    'var8 = dg.Item(int1, 1)
                    num1 = dv1.Item(int1).Item(strAnal)
                    num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                    boolNum = True
                    var8 = num1
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                    'ctArrReportNA = ctArrReportNA + 1
                    strNA1 = "Analytical Reference Standard"
                    strNA2 = "LLOQ"
                    strNA3 = "Analytical Reference Standard"
                    strNA4 = "Watson Analytical Reference Standard Table"

                Case 3
                    strFind = "[LLOQUNITS]"
                    int1 = FindRowDVByCol("LLOQ Units", dv1, "Item")
                    var8 = dv1.Item(int1).Item(strAnal)

                    int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        var8 = str1
                    End If

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                    'ctArrReportNA = ctArrReportNA + 1
                    strNA1 = "Analytical Reference Standard"
                    strNA2 = "LLOQ Units"
                    strNA3 = "Analytical Reference Standard"
                    strNA4 = "Watson Analytical Reference Standard Table"

                Case 4
                    strFind = "[ULOQ]"
                    int1 = FindRowDVByCol("ULOQ", dv1, "Item")
                    num1 = dv1.Item(int1).Item(strAnal)
                    num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                    boolNum = True
                    var8 = num1
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                    'ctArrReportNA = ctArrReportNA + 1
                    strNA1 = "Analytical Reference Standard"
                    strNA2 = "ULOQ"
                    strNA3 = "Analytical Reference Standard"
                    strNA4 = "Watson Analytical Reference Standard Table"

                Case 5
                    strFind = "[ULOQUNITS]"
                    int1 = FindRowDVByCol("ULOQ Units", dv1, "Item")
                    var8 = dv1.Item(int1).Item(strAnal)

                    int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        var8 = str1
                    End If

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                    'ctArrReportNA = ctArrReportNA + 1
                    strNA1 = "Analytical Reference Standard"
                    strNA2 = "ULOQ Units"
                    strNA3 = "Analytical Reference Standard"
                    strNA4 = "Watson Analytical Reference Standard Table"

                Case 6
                    boolDo = boolDoSS
                    If boolDoSS Then
                        strFind = "[SAMPLESIZE]"
                        varReplace = ReturnSearch(wd, "Temp", 39, strFind, mysel.Range, -1)
                    End If

                Case 7
                    boolDo = boolDoSS
                    If boolDoSS Then
                        strFind = "[SAMPLESIZEUNITS]"
                        varReplace = ReturnSearch(wd, "Temp", 47, strFind, mysel.Range, -1)
                    End If

                Case 8
                    strFind = "[ANTICOAGULANT]"
                    varReplace = ReturnSearch(wd, "Temp", 46, strFind, mysel.Range, -1)

                Case 9
                    strFind = "[MATRIX]"
                    varReplace = ReturnSearch(wd, "Temp", 38, strFind, mysel.Range, -1)

                Case 10
                    strFind = "[SPECIES]"
                    varReplace = ReturnSearch(wd, "Temp", 40, strFind, mysel.Range, -1)


                Case 11
                    strFind = "[UC_ANTICOAGULANT]"
                    varReplace = ReturnSearch(wd, "Temp", 265, strFind, mysel.Range, -1)

                Case 12
                    strFind = "[LC_ANTICOAGULANT]"
                    varReplace = ReturnSearch(wd, "Temp", 266, strFind, mysel.Range, -1)

                Case 13
                    strFind = "[UC_MATRIX]"
                    varReplace = ReturnSearch(wd, "Temp", 263, strFind, mysel.Range, -1)

                Case 14
                    strFind = "[LC_MATRIX]"
                    varReplace = ReturnSearch(wd, "Temp", 264, strFind, mysel.Range, -1)

                Case 15
                    strFind = "[UC_SPECIES]"
                    varReplace = ReturnSearch(wd, "Temp", 261, strFind, mysel.Range, -1)

                Case 16
                    strFind = "[LC_SPECIES]"
                    varReplace = ReturnSearch(wd, "Temp", 262, strFind, mysel.Range, -1)

            End Select

            If Len(strFind) = 0 Or boolDo = False Then
            Else
                With wd
                    With mysel.Find
                        .ClearFormatting()
                        .Text = strFind
                        With .Replacement
                            .ClearFormatting()
                            If boolNum Then
                                .Text = num1
                            Else
                                .Text = varReplace
                            End If
                        End With
                        .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                        If .Found Then
                            If Len(varReplace) = 0 Or StrComp(NZ(varReplace, "[None]"), "[None]", CompareMethod.Text) = 0 Or StrComp(varReplace, "[NA]", CompareMethod.Text) = 0 Then
                                'add entries to arrReportNA
                                ctArrReportNA = ctArrReportNA + 1
                                arrReportNA(1, ctArrReportNA) = strNA1
                                arrReportNA(2, ctArrReportNA) = strNA2
                                arrReportNA(3, ctArrReportNA) = strNA3
                                arrReportNA(4, ctArrReportNA) = strNA4
                            End If
                        End If
                    End With
                End With
            End If
        Next
        'now look for any [NA]s
        With mysel.Find
            .ClearFormatting()
            .Text = "[NA]"
            With .Replacement
                .ClearFormatting()
                .Text = "[NA]"
                .Font.Bold = True
                .Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
            End With
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        End With

    End Sub

    Sub SearchReplaceMethodSummary(ByVal wd, ByVal strAnal, ByVal strA)

        Dim Count1 As Short
        Dim strFind As String
        Dim varReplace
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim int1 As Short
        Dim var8
        Dim strNA1 As String
        Dim strNA2 As String
        Dim strNA3 As String
        Dim strNA4 As String
        Dim charSectionName As String
        Dim mysel As Microsoft.Office.Interop.Word.Selection
        Dim num1 As Object
        Dim boolNum As Boolean
        Dim dtbl As System.Data.DataTable
        Dim strF As String
        Dim rows1() As DataRow
        Dim var1
        Dim str1 As String

        charSectionName = "Method Summary"
        ''''''wdd.visible = True

        dgv = frmH.dgvMethodValData
        dv = dgv.DataSource 'intI is analyte column in dgvMethodValData
        dv1 = frmH.dgvWatsonAnalRef.DataSource 'intI2 is analyte column in dgvWatsonAnalRef

        mysel = wd.selection

        For Count1 = 1 To 100

            boolNum = False
            strFind = ""
            Select Case Count1
                Case 1
                    strFind = "[LABMETHODNAME]"
                    int1 = FindRowDVByCol("Lab Method Title", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = Trim(var8)
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Lab Method Name"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"

                Case 2
                    strFind = "[LABMETHODNUMBER]"
                    int1 = FindRowDVByCol("Lab Method Number", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = Trim(var8)
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Lab Method Number"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"
                Case 3
                    strFind = "[LMAPPENDIXNUMBER]"
                    If boolUseHyperlinks Then
                        varReplace = "Appendix_" & strA
                    Else
                        varReplace = "Appendix" & ChrW(160) & strA
                    End If


                Case 4
                    strFind = "[METHODASSAYPROCEDUREDESCRIPTION]"
                    int1 = FindRowDVByCol("Extraction Procedure Description", dv, "Item")
                    'var8 = dg.Item(int1, 0)
                    var8 = dv.Item(int1).Item(strAnal)

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        'varReplace = UnCapit(Trim(var8), False)
                        '20181015 LEE:
                        'Don't make this uncapitalized
                        varReplace = Trim(var8)
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Extraction Procedure Description"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"

                Case 5
                    strFind = "[ANALYTE]"
                    'var8 = arrAnalytes(1, intI)
                    var8 = strAnal

                    '20181015 LEE:
                    'put nbh back in
                    var8 = Replace(var8, "-", NBHReal, 1, -1, CompareMethod.Text)

                    varReplace = var8

                Case 7
                    strFind = "[SAMPLESIZE]"
                    int1 = FindRowDVByCol("Sample Size", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    ElseIf var8 = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If

                    strNA1 = charSectionName
                    strNA2 = "Sample Size"
                    strNA3 = "Add/Edit Top Level Data"
                    strNA4 = "Data From Watson Table"

                Case 8
                    strFind = "[SAMPLESIZEUNITS]"
                    int1 = FindRowDVByCol("Sample Size Units", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Sample Size Units"
                    strNA3 = "Add/Edit Top Level Data"
                    strNA4 = "Data From Watson Table"

                Case 9
                    strFind = "[METHODCORPORATESTUDY/PROJECTNUMBER]"
                    int1 = FindRowDVByCol("Validation Corporate Study/Project Number", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = Trim(var8)
                    End If
                    'replace hyphens with nbh
                    varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                    'replace spaces with nbs
                    varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                    strNA1 = charSectionName
                    strNA2 = "Validation Corporate Study/Project Number"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"

                Case 10
                    strFind = "[ANTICOAGULANT]"
                    int1 = FindRowDVByCol("Anticoagulant/Preservative", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    Dim Count2 As Short
                    Dim int2 As Short
                    Dim var2

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        If Len(var8) = 0 Then
                        Else
                            int1 = Len(var8)
                            int2 = 0
                            For Count2 = 1 To int1
                                var1 = CStr(Mid(var8, Count2, 1))
                                var2 = Asc(var1)
                                If var2 > 64 And var2 < 91 Then
                                    int2 = int2 + 1
                                Else
                                    Exit For
                                End If
                            Next
                            If int2 > 2 Then 'probably is an acronym, leave capitalized
                                var8 = var8
                            Else
                                var8 = Trim(var8)
                            End If

                            '20181015 LEE:
                            'put nbh back in
                            var8 = Replace(var8, "-", NBHReal, 1, -1)

                        End If

                        varReplace = var8 'UnCapit(Trim(var8), True)
                    End If

                    strNA1 = charSectionName
                    strNA2 = "Anticoagulant/Preservative"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"
                Case 11
                    strFind = "[SPECIES]"
                    int1 = FindRowDVByCol("Species", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8 'UnCapit(Trim(var8), True)
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Species"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"
                Case 12
                    strFind = "[MATRIX]"
                    int1 = FindRowDVByCol("Matrix", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8 'UnCapit(Trim(var8), True)
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Matrix"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"

                Case 20
                    strFind = "[LLOQ]"
                    int1 = FindRowDVByCol("LLOQ", dv1, "Item")
                    'var8 = dg.Item(int1, 1)
                    num1 = dv1.Item(int1).Item(strAnal)
                    num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                    boolNum = True
                    var8 = num1
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                    'ctArrReportNA = ctArrReportNA + 1
                    strNA1 = "Analytical Reference Standard"
                    strNA2 = "LLOQ"
                    strNA3 = "Analytical Reference Standard"
                    strNA4 = "Watson Analytical Reference Standard Table"

                Case 21
                    strFind = "[LLOQUNITS]"
                    int1 = FindRowDVByCol("LLOQ Units", dv1, "Item")
                    var8 = dv1.Item(int1).Item(strAnal)

                    int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        var8 = str1
                    End If

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                    'ctArrReportNA = ctArrReportNA + 1
                    strNA1 = "Analytical Reference Standard"
                    strNA2 = "LLOQ Units"
                    strNA3 = "Analytical Reference Standard"
                    strNA4 = "Watson Analytical Reference Standard Table"

                Case 22
                    strFind = "[ULOQ]"
                    int1 = FindRowDVByCol("ULOQ", dv1, "Item")
                    num1 = dv1.Item(int1).Item(strAnal)
                    num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                    boolNum = True
                    var8 = num1
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                    'ctArrReportNA = ctArrReportNA + 1
                    strNA1 = "Analytical Reference Standard"
                    strNA2 = "ULOQ"
                    strNA3 = "Analytical Reference Standard"
                    strNA4 = "Watson Analytical Reference Standard Table"

                Case 23
                    strFind = "[ULOQUNITS]"
                    int1 = FindRowDVByCol("ULOQ Units", dv1, "Item")
                    var8 = dv1.Item(int1).Item(strAnal)

                    int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        var8 = str1
                    End If

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                    'ctArrReportNA = ctArrReportNA + 1
                    strNA1 = "Analytical Reference Standard"
                    strNA2 = "ULOQ Units"
                    strNA3 = "Analytical Reference Standard"
                    strNA4 = "Watson Analytical Reference Standard Table"

                Case 24
                    strFind = "[UC_ANTICOAGULANT]"
                    int1 = FindRowDVByCol("Anticoagulant/Preservative", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    Dim Count2 As Short
                    Dim int2 As Short
                    Dim var2

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        If Len(var8) = 0 Then
                        Else
                            int1 = Len(var8)
                            int2 = 0
                            For Count2 = 1 To int1
                                var1 = CStr(Mid(var8, Count2, 1))
                                var2 = Asc(var1)
                                If var2 > 64 And var2 < 91 Then
                                    int2 = int2 + 1
                                Else
                                    Exit For
                                End If
                            Next
                            If int2 > 2 Then 'probably is an acronym, leave capitalized
                                var8 = var8
                            Else
                                var8 = CapitAllWords(Trim(var8))
                            End If
                        End If

                        varReplace = var8 'UnCapit(Trim(var8), True)
                    End If

                Case 25
                    strFind = "[LC_ANTICOAGULANT]"
                    int1 = FindRowDVByCol("Anticoagulant/Preservative", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    Dim Count2 As Short
                    Dim int2 As Short
                    Dim var2

                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        If Len(var8) = 0 Then
                        Else
                            int1 = Len(var8)
                            int2 = 0
                            For Count2 = 1 To int1
                                var1 = CStr(Mid(var8, Count2, 1))
                                var2 = Asc(var1)
                                If var2 > 64 And var2 < 91 Then
                                    int2 = int2 + 1
                                Else
                                    Exit For
                                End If
                            Next
                            If int2 > 2 Then 'probably is an acronym, leave capitalized
                                var8 = var8
                            Else
                                var8 = LowerCase(Trim(var8))
                            End If
                        End If

                        varReplace = var8 'UnCapit(Trim(var8), True)
                    End If

                Case 26
                    strFind = "[UC_SPECIES]"
                    int1 = FindRowDVByCol("Species", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = CapitAllWords(var8) 'UnCapit(Trim(var8), True)
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Species"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"

                Case 27
                    strFind = "[LC_SPECIES]"
                    int1 = FindRowDVByCol("Species", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = LowerCase(var8) 'UnCapit(Trim(var8), True)
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Species"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"

                Case 28
                    strFind = "[UC_MATRIX]"
                    int1 = FindRowDVByCol("Matrix", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = CapitAllWords(var8) 'UnCapit(Trim(var8), True)
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Matrix"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"

                Case 29
                    strFind = "[LC_MATRIX]"
                    int1 = FindRowDVByCol("Matrix", dv, "Item")
                    var8 = dv.Item(int1).Item(strAnal)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = LowerCase(var8) 'UnCapit(Trim(var8), True)
                    End If
                    strNA1 = charSectionName
                    strNA2 = "Matrix"
                    strNA3 = "Method Validation Data"
                    strNA4 = "Method Validation Data Table"

            End Select

            If Len(strFind) = 0 Then
            Else
                With wd
                    With mysel.Find
                        .ClearFormatting()
                        .Text = strFind
                        With .Replacement
                            .ClearFormatting()
                            If boolNum Then
                                .Text = num1
                            Else
                                .Text = varReplace
                            End If
                        End With
                        .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                        If .Found Then
                            If Len(varReplace) = 0 Or StrComp(NZ(varReplace, "[None]"), "[None]", CompareMethod.Text) = 0 Or StrComp(varReplace, "[NA]", CompareMethod.Text) = 0 Then
                                'add entries to arrReportNA
                                ctArrReportNA = ctArrReportNA + 1
                                arrReportNA(1, ctArrReportNA) = strNA1
                                arrReportNA(2, ctArrReportNA) = strNA2
                                arrReportNA(3, ctArrReportNA) = strNA3
                                arrReportNA(4, ctArrReportNA) = strNA4
                            End If
                        End If
                    End With
                End With
            End If
        Next
        'now look for any [NA]s
        With mysel.Find
            .ClearFormatting()
            .Text = "[NA]"
            With .Replacement
                .ClearFormatting()
                .Text = "[NA]"
                .Font.Bold = True
                .Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
            End With
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        End With

        'now look for any ^2
        With mysel.Find
            .ClearFormatting()
            .Text = "^2"
            With .Replacement
                .ClearFormatting()
                .Text = "2"
                .Font.Superscript = True
            End With
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        End With

        'now look for any ^2, special carrot character
        With mysel.Find
            .ClearFormatting()
            '.Text = ChrW(94) & "2"
            .Text = "^^2"
            With .Replacement
                .ClearFormatting()
                .Text = "2"
                .Font.Superscript = True
            End With
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        End With

    End Sub

    Function GetAnalyteAll(ByVal idTbl As Short, ByVal boolRespective As Boolean) As String

        Dim dv As system.data.dataview
        Dim tbl1 As System.Data.DataTable
        Dim strF As String
        Dim rows1() As DataRow
        Dim var1, var2, var3, var4
        Dim Count2 As Short
        Dim tbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim drows() As DataRow
        Dim tblN As System.Data.DataTable
        Dim tblA As System.Data.DataTable
        Dim rowsA() As DataRow
        Dim strS As String
        Dim intA As Short
        Dim boolIS As Boolean

        str1 = "[NA]"

        tbl1 = tblReportTable
        strF = "id_tblConfigReportTables = " & idTbl & " AND ID_TBLSTUDIES = " & id_tblStudies
        rows1 = tbl1.Select(strF)
        var4 = rows1(0).Item("ID_TBLREPORTTABLE")

        tblA = tblAnalytesHome
        strF = "IsReplicate = 'No'"
        strS = "IsIntStd ASC, AnalyteDescription ASC"
        rowsA = tblA.Select(strF, strS)
        intA = rowsA.Length

        tbl2 = tblReportTableAnalytes
        int3 = 0

        'strF = "ID_TBLCONFIGREPORTTABLES = " & idTbl & " AND ID_TBLSTUDIES = " & id_tblStudies
        'rows1 = tbl1.Select(strF)
        For Count2 = 0 To intA - 1
            var2 = rowsA(Count2).Item("ANALYTEID")
            str2 = rowsA(Count2).Item("IsIntStd")
            boolIS = False
            If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                boolIS = True
            Else
                boolIS = False
            End If
            If boolIS Then
                int2 = -1
            Else
                strF = "ID_TBLREPORTTABLE = " & var4 & " AND ANALYTEID = " & var2
                rows2 = tbl2.Select(strF)
                int2 = rows2(0).Item("BOOLINCLUDE")
            End If
            'If int2 = -1 Then
            '    int3 = int3 + 1
            '    If boolIS Then
            '        If int3 = 1 Then
            '            str1 = rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
            '        ElseIf int3 = intA And intA > 2 Then
            '            str1 = str1 & ", and " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
            '        ElseIf int3 <> intA And intA > 2 Then
            '            str1 = str1 & ", " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
            '        Else
            '            str1 = str1 & " and " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
            '        End If
            '    Else
            '        If int3 = 1 Then
            '            str1 = rowsA(Count2).Item("AnalyteDescription")
            '        ElseIf int3 = intA And intA > 2 Then
            '            str1 = str1 & ", and " & rowsA(Count2).Item("AnalyteDescription")
            '        ElseIf int3 <> intA And intA > 2 Then
            '            str1 = str1 & ", " & rowsA(Count2).Item("AnalyteDescription")
            '        Else
            '            str1 = str1 & " and " & rowsA(Count2).Item("AnalyteDescription")
            '        End If
            '    End If
            'End If

            If int2 = -1 Then
                int3 = int3 + 1
                If boolIS Then
                    If int3 = 1 Then
                        str1 = rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION") & " (Internal Standard)"
                    ElseIf int3 = intA And intA > 2 Then
                        str1 = str1 & ", and " & rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION") & " (Internal Standard)"
                    ElseIf int3 <> intA And intA > 2 Then
                        str1 = str1 & ", " & rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION") & " (Internal Standard)"
                    Else
                        str1 = str1 & " and " & rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION") & " (Internal Standard)"
                    End If
                Else
                    If int3 = 1 Then
                        str1 = rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION")
                    ElseIf int3 = intA And intA > 2 Then
                        str1 = str1 & ", and " & rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION")
                    ElseIf int3 <> intA And intA > 2 Then
                        str1 = str1 & ", " & rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION")
                    Else
                        str1 = str1 & " and " & rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION")
                    End If
                End If
            End If

        Next

        'var1 = rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION")

        If boolRespective And intA > 1 Then
            GetAnalyteAll = str1 & ", respectively"
        Else
            GetAnalyteAll = str1
        End If


    End Function

    Function GetAnalyte(ByVal idTbl As Short, ByVal boolRespective As Boolean) As String

        Dim dv As system.data.dataview
        Dim tbl1 As System.Data.DataTable
        Dim strF As String
        Dim rows1() As DataRow
        Dim var1, var2, var3, var4
        Dim Count2 As Short
        Dim tbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim drows() As DataRow
        Dim tblN As System.Data.DataTable
        Dim tblA As System.Data.DataTable
        Dim rowsA() As DataRow
        Dim strS As String
        Dim intA As Short

        str1 = "[NA]"

        tbl1 = tblReportTable
        strF = "id_tblConfigReportTables = " & idTbl & " AND ID_TBLSTUDIES = " & id_tblStudies
        rows1 = tbl1.Select(strF)
        var4 = rows1(0).Item("ID_TBLREPORTTABLE")

        tblA = tblAnalytesHome
        strF = "IsIntStd = 'No' AND IsReplicate = 'No'"
        strS = "IsIntStd ASC, AnalyteDescription ASC"
        rowsA = tblA.Select(strF, strS)
        intA = rowsA.Length

        tbl2 = tblReportTableAnalytes
        int3 = 0

        For Count2 = 0 To intA - 1
            var2 = rowsA(Count2).Item("ANALYTEID")
            strF = "ID_TBLREPORTTABLE = " & var4 & " AND ANALYTEID = " & var2
            rows2 = tbl2.Select(strF)
            int2 = rows2(0).Item("BOOLINCLUDE")
            If int2 = -1 Then
                int3 = int3 + 1
                'str3 = "TableID = " & idTbl & " AND AnalyteName = '" & rowsA(Count2).Item("AnalyteDescription") & "'"
                'drows = tblN.Select(str3)
                'var1 = rowsA(Count2).Item("AnalyteDescription")
                var1 = rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION")
                '
                If int3 = 1 Then
                    str1 = var1
                ElseIf int3 = intA And intA > 2 Then
                    str1 = str1 & ", " & var1
                ElseIf int3 <> intA And intA > 2 Then
                    str1 = str1 & ", " & var1
                Else
                    str1 = str1 & " and " & var1
                End If
            End If
        Next

        If boolRespective And intA > 1 Then
            GetAnalyte = str1 & ", respectively"
        Else
            GetAnalyte = str1
        End If


    End Function

    Function GetAnalyteRef(ByVal idTbl As Short, ByVal boolRespective As Boolean) As String

        GetAnalyteRef = "[NA]"

        Dim dv As system.data.dataview
        Dim tbl1 As System.Data.DataTable
        Dim strF As String
        Dim rows1() As DataRow
        Dim var1, var2, var3, var4
        Dim var2a, var2b, var2c, var2d
        Dim Count2 As Short
        Dim tbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim drows() As DataRow
        Dim tblN As System.Data.DataTable
        Dim tblA As System.Data.DataTable
        Dim rowsA() As DataRow
        Dim strS As String
        Dim intA As Short
        Dim dtblP As System.Data.Datatable
        Dim rowsP() As DataRow
        Dim varI
        Dim boolIS As Boolean = False
        Dim arrNames()
        Dim intNames As Short

        str1 = "[NA]"

        tbl1 = tblReportTable
        strF = "id_tblReportTable = " & idTbl ' & " AND ID_TBLSTUDIES = " & id_tblStudies
        rows1 = tbl1.Select(strF)
        var4 = rows1(0).Item("ID_TBLREPORTTABLE")

        'determine if IS is included
        dtblP = tblTableProperties
        rowsP = dtblP.Select(strF)
        varI = rowsP(0).Item("BOOLINCLUDEISTBL")
        If varI = 0 Then
            boolIS = False
        Else
            boolIS = True
        End If

        tblA = tblAnalytesHome
        If boolIS Then
            strF = "(IsIntStd = 'No' OR IsIntStd = 'Yes') AND IsReplicate = 'No'"
        Else
            strF = "IsIntStd = 'No' AND IsReplicate = 'No'"
        End If
        strS = "AnalyteDescription ASC"
        rowsA = tblA.Select(strF, strS)
        intA = rowsA.Length

        tbl2 = tblReportTableAnalytes
        int3 = 0

        intNames = 0
        ReDim arrNames(intA)
        For Count2 = 0 To intA - 1
            var1 = rowsA(Count2).Item("AnalyteDescription")
            var2 = rowsA(Count2).Item("ANALYTEID")
            var2a = rowsA(Count2).Item("ANALYTEINDEX")
            var2b = rowsA(Count2).Item("MASTERASSAYID")
            var2c = rowsA(Count2).Item("IsIntStd")
            var2d = rowsA(Count2).Item("ORIGINALANALYTEDESCRIPTION") '

            If StrComp(var2c, "No", CompareMethod.Text) = 0 Then
                strF = "ID_TBLREPORTTABLE = " & var4 & " AND ANALYTEID = " & var2 & " AND ANALYTEINDEX = " & var2a & " AND MASTERASSAYID = " & var2b
                rows2 = tbl2.Select(strF)
                int2 = rows2(0).Item("BOOLINCLUDE")

            Else
                int2 = -1
                var2d = var1 & " (Internal Standard)"
            End If

            If int2 = -1 Then
                intNames = intNames + 1
                arrNames(intNames) = var2d
            End If

        Next

        str1 = "[NA]"
        int3 = 0
        For Count2 = 1 To intNames
            int3 = int3 + 1
            var1 = arrNames(Count2)
            If int3 = 1 Then
                str1 = var1
            ElseIf int3 = intNames And intNames > 2 Then
                str1 = str1 & ", and " & var1
            ElseIf int3 <> intNames And intNames > 2 Then
                str1 = str1 & ", " & var1
            Else
                str1 = str1 & " and " & var1
            End If
        Next

        If intNames = 0 Then
            GetAnalyteRef = "[NA]"
        Else
            If boolRespective And intNames > 1 Then
                GetAnalyteRef = str1 & ", respectively"
            Else
                GetAnalyteRef = str1
            End If
        End If



    End Function

    Function CrossRefText(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal str1 As String) As Boolean

        'Dim hd As Microsoft.Office.Interop.Word.section.HeadingStyle
        'Dim sty As Microsoft.Office.Interop.Word.section.Style
        Dim var1 As Object, var2
        Dim Count1 As Short

        'this function will insert a cross reference (text portion) to a heading corresponding to str1
        'will return false if no heading exists
        CrossRefText = False
        var1 = wd.ActiveDocument.GetCrossReferenceItems(Microsoft.Office.Interop.Word.WdReferenceType.wdRefTypeHeading)
        'wdRefTypeBookmark
        'wdRefTypeEndnote
        'wdRefTypeFootnote
        'wdRefTypeHeading
        'wdRefTypeNumberedItem


        For Count1 = 1 To UBound(var1)
            var2 = var1(Count1)
            If InStr(1, var2, str1, vbTextCompare) > 0 Then
                wd.Selection.InsertCrossReference(ReferenceType:="Numbered item", _
                    ReferenceKind:=Microsoft.Office.Interop.Word.WdReferenceKind.wdContentText, ReferenceItem:=CStr(Count1), _
                    InsertAsHyperlink:=True, IncludePosition:=False) ', SeparateNumbers:=False, _
                'SeparatorString:=" ")
                CrossRefText = True
                Exit For
            End If
        Next

    End Function

    Function GetAppFig(ByVal strFC As String)



    End Function

    Function GetTableNumber(ByVal idTbl As Short, ByVal boolRespective As Boolean) As String

        Dim dv As system.data.dataview
        Dim tbl1 As System.Data.DataTable
        Dim strF As String
        Dim rows1() As DataRow
        Dim var1, var2, var3, var4
        Dim Count2 As Short
        Dim tbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim drows() As DataRow
        Dim tblN As System.Data.DataTable
        Dim tblA As System.Data.DataTable
        Dim rowsA() As DataRow
        Dim strS As String
        Dim intA As Short

        tblN = tblTableN
        str1 = "[NA]"

        Dim var5, var6, var7

        tbl1 = tblReportTable
        strF = "id_tblConfigReportTables = " & idTbl & " AND ID_TBLSTUDIES = " & id_tblStudies
        rows1 = tbl1.Select(strF)
        var4 = rows1(0).Item("ID_TBLREPORTTABLE")

        tblA = tblAnalytesHome
        strF = "IsIntStd = 'No' AND IsReplicate = 'No'"
        strS = "IsIntStd ASC, AnalyteDescription ASC"
        rowsA = tblA.Select(strF, strS)
        intA = rowsA.Length

        tbl2 = tblReportTableAnalytes
        int3 = 0

        If ctTableN = 0 Then
            'strF = "ID_TBLCONFIGREPORTTABLES = " & idTbl & " AND ID_TBLSTUDIES = " & id_tblStudies
            'rows1 = tbl1.Select(strF)
            For Count2 = 0 To intA - 1
                var2 = rowsA(Count2).Item("ANALYTEID")
                var5 = rowsA(Count2).Item("ANALYTEINDEX")
                var6 = rowsA(Count2).Item("MASTERASSAYID")
                strF = "ID_TBLREPORTTABLE = " & var4 & " AND ANALYTEID = " & var2 & " AND ANALYTEINDEX = " & var5 & " AND MASTERASSAYID = " & var6
                rows2 = tbl2.Select(strF)
                int4 = rows2.Length
                If int4 = 0 Then
                Else
                    int2 = rows2(0).Item("BOOLINCLUDE")
                    If int2 = -1 Then
                        int3 = int3 + 1
                        If boolRespective Then
                            If int3 = 1 Then
                                str1 = "Table" & ChrW(160) & "[NA]"
                            ElseIf int3 = intA And intA > 2 Then
                                str1 = str1 & ", and Table" & ChrW(160) & "[NA]"
                            ElseIf int3 <> intA And intA > 2 Then
                                str1 = str1 & ", Table" & ChrW(160) & "[NA]"
                            Else
                                str1 = str1 & " and Table" & ChrW(160) & "[NA]"
                            End If
                        Else
                            If int3 = 1 Then
                                str1 = "Table" & ChrW(160) & "[NA] for " & rowsA(Count2).Item("AnalyteDescription")
                            ElseIf int3 = intA And intA > 2 Then
                                str1 = str1 & ", and Table" & ChrW(160) & "[NA] for " & rowsA(Count2).Item("AnalyteDescription")
                            ElseIf int3 <> intA And intA > 2 Then
                                str1 = str1 & ", Table" & ChrW(160) & "[NA] for " & rowsA(Count2).Item("AnalyteDescription")
                            Else
                                str1 = str1 & " and Table" & ChrW(160) & "[NA] for " & rowsA(Count2).Item("AnalyteDescription")
                            End If
                        End If
                    End If
                End If
            Next
        Else
            For Count2 = 0 To intA - 1
                var2 = rowsA(Count2).Item("ANALYTEID")
                var5 = rowsA(Count2).Item("ANALYTEINDEX")
                var6 = rowsA(Count2).Item("MASTERASSAYID")
                strF = "ID_TBLREPORTTABLE = " & var4 & " AND ANALYTEID = " & var2 & " AND ANALYTEINDEX = " & var5 & " AND MASTERASSAYID = " & var6
                rows2 = tbl2.Select(strF)
                int4 = rows2.Length
                If int4 = 0 Then
                Else
                    int2 = rows2(0).Item("BOOLINCLUDE")
                    If int2 = -1 Then
                        int3 = int3 + 1
                        If boolRespective Then
                            str3 = "TableID = " & idTbl & " AND AnalyteName = '" & rowsA(Count2).Item("AnalyteDescription") & "'"
                            drows = tblN.Select(str3)

                            'var1 = NZ(drows(0).Item("TableNumber"), "[NA]")

                            If drows.Length = 0 Then
                                var1 = "[NA]"
                            Else
                                If boolUseHyperlinks Then
                                    var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                                    If int3 = 1 Then
                                        str1 = "Table_" & var1
                                    ElseIf int3 = intA And intA > 2 Then
                                        str1 = str1 & ", and Table_" & var1
                                    ElseIf int3 <> intA And intA > 2 Then
                                        str1 = str1 & ", Table_" & var1
                                    Else
                                        str1 = str1 & " and Table_" & var1
                                    End If
                                Else
                                    var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                                    If int3 = 1 Then
                                        str1 = "Table" & ChrW(160) & var1
                                    ElseIf int3 = intA And intA > 2 Then
                                        str1 = str1 & ", and Table" & ChrW(160) & var1
                                    ElseIf int3 <> intA And intA > 2 Then
                                        str1 = str1 & ", Table" & ChrW(160) & var1
                                    Else
                                        str1 = str1 & " and Table" & ChrW(160) & var1
                                    End If
                                End If

                            End If


                        Else
                            str3 = "TableID = " & idTbl & " AND AnalyteName = '" & rowsA(Count2).Item("AnalyteDescription") & "'"
                            drows = tblN.Select(str3)
                            If drows.Length = 0 Then
                                var1 = "[NA]"
                            Else
                                If boolUseHyperlinks Then
                                    var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                                    If int3 = 1 Then
                                        str1 = "Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                    ElseIf int3 = intA And intA > 2 Then
                                        str1 = str1 & ", and Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                    ElseIf int3 <> intA And intA > 2 Then
                                        str1 = str1 & ", Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                    Else
                                        str1 = str1 & " and Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                    End If
                                Else
                                    var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                                    If int3 = 1 Then
                                        str1 = "Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                    ElseIf int3 = intA And intA > 2 Then
                                        str1 = str1 & ", and Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                    ElseIf int3 <> intA And intA > 2 Then
                                        str1 = str1 & ", Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                    Else
                                        str1 = str1 & " and Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                    End If
                                End If

                            End If

                        End If
                    End If
                End If
            Next
        End If

        GetTableNumber = str1


    End Function

    Function GetTableNumberAll(ByVal idTbl As Short, ByVal boolRespective As Boolean) As String

        Dim dv As System.Data.DataView
        Dim tbl1 As System.Data.DataTable
        Dim strF As String
        Dim rows1() As DataRow
        Dim var1, var2, var3, var4
        Dim Count2 As Short
        Dim tbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim drows() As DataRow
        Dim tblN As System.Data.DataTable
        Dim tblA As System.Data.DataTable
        Dim rowsA() As DataRow
        Dim strS As String
        Dim intA As Short
        Dim boolIS As Boolean

        tblN = tblTableN
        str1 = "[NA]"

        tbl1 = tblReportTable
        strF = "id_tblConfigReportTables = " & idTbl & " AND ID_TBLSTUDIES = " & id_tblStudies
        rows1 = tbl1.Select(strF)
        var4 = rows1(0).Item("ID_TBLREPORTTABLE")

        tblA = tblAnalytesHome
        strF = "IsReplicate = 'No'"
        strS = "IsIntStd ASC, AnalyteDescription ASC"
        rowsA = tblA.Select(strF, strS)
        intA = rowsA.Length

        tbl2 = tblReportTableAnalytes
        int3 = 0

        If ctTableN = 0 Then
            'strF = "ID_TBLCONFIGREPORTTABLES = " & idTbl & " AND ID_TBLSTUDIES = " & id_tblStudies
            'rows1 = tbl1.Select(strF)
            For Count2 = 0 To intA - 1
                var2 = rowsA(Count2).Item("ANALYTEID")
                str2 = rowsA(Count2).Item("IsIntStd")
                var3 = rowsA(Count2).Item("AnalyteDescription")
                boolIS = False
                If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                    boolIS = True
                    int2 = -1
                Else
                    boolIS = False
                    strF = "ID_TBLREPORTTABLE = " & var4 & " AND ANALYTEID = " & var2
                    Erase rows2
                    rows2 = tbl2.Select(strF)
                    If rows2.Length = 0 Then
                        int2 = 0
                    Else
                        int2 = rows2(0).Item("BOOLINCLUDE")
                    End If
                End If

                int4 = rows2.Length
                If int4 = 0 Then
                Else
                    int2 = rows2(0).Item("BOOLINCLUDE")
                    If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                        boolIS = True
                    Else
                        boolIS = False
                    End If
                    If boolIS Then
                        int2 = -1
                    Else
                    End If
                    If int2 = -1 Then
                        int3 = int3 + 1
                        If boolIS Then
                            If boolRespective Then
                                If int3 = 1 Then
                                    str1 = "Table [NA]"
                                ElseIf int3 = intA And intA > 2 Then
                                    str1 = str1 & ", and Table [NA]"
                                ElseIf int3 <> intA And intA > 2 Then
                                    str1 = str1 & ", Table [NA]"
                                Else
                                    str1 = str1 & " and Table [NA]"
                                End If
                            Else
                                If int3 = 1 Then
                                    str1 = "Table [NA] for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                ElseIf int3 = intA And intA > 2 Then
                                    str1 = str1 & ", and Table [NA] for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                ElseIf int3 <> intA And intA > 2 Then
                                    str1 = str1 & ", Table [NA] for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                Else
                                    str1 = str1 & " and Table [NA] for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                End If
                            End If
                        Else
                            If boolRespective Then
                                If int3 = 1 Then
                                    str1 = "Table [NA]"
                                ElseIf int3 = intA And intA > 2 Then
                                    str1 = str1 & ", and Table [NA]"
                                ElseIf int3 <> intA And intA > 2 Then
                                    str1 = str1 & ", Table [NA]"
                                Else
                                    str1 = str1 & " and Table [NA]"
                                End If
                            Else
                                If int3 = 1 Then
                                    str1 = "Table [NA] for " & rowsA(Count2).Item("AnalyteDescription")
                                ElseIf int3 = intA And intA > 2 Then
                                    str1 = str1 & ", and Table [NA] for " & rowsA(Count2).Item("AnalyteDescription")
                                ElseIf int3 <> intA And intA > 2 Then
                                    str1 = str1 & ", Table [NA] for " & rowsA(Count2).Item("AnalyteDescription")
                                Else
                                    str1 = str1 & " and Table [NA] for " & rowsA(Count2).Item("AnalyteDescription")
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Else
            For Count2 = 0 To intA - 1
                var2 = rowsA(Count2).Item("ANALYTEID")
                str2 = rowsA(Count2).Item("IsIntStd")
                var3 = rowsA(Count2).Item("AnalyteDescription")
                boolIS = False

                If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                    boolIS = True
                    int2 = -1
                Else
                    boolIS = False
                    strF = "ID_TBLREPORTTABLE = " & var4 & " AND ANALYTEID = " & var2
                    Erase rows2
                    rows2 = tbl2.Select(strF)
                    If rows2.Length = 0 Then
                        int2 = 0
                    Else
                        int2 = rows2(0).Item("BOOLINCLUDE")
                    End If
                End If

                int4 = rows2.Length
                If int4 = 0 Then
                Else
                    If int2 = -1 Then
                        int3 = int3 + 1
                        If boolIS Then
                            str3 = "TableID = " & idTbl & " AND AnalyteName = '" & rowsA(Count2).Item("AnalyteDescription") & "'"
                            drows = tblN.Select(str3)
                            If drows.Length = 0 Then
                                var1 = "[NA]"
                            Else
                                var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                                If boolUseHyperlinks Then
                                    If boolRespective Then
                                        If int3 = 1 Then
                                            str1 = "Table_" & var1
                                        ElseIf int3 = intA And intA > 2 Then
                                            str1 = str1 & ", and Table_" & var1
                                        ElseIf int3 <> intA And intA > 2 Then
                                            str1 = str1 & ", Table_" & var1
                                        Else
                                            str1 = str1 & " and Table_" & var1
                                        End If
                                    Else
                                        If int3 = 1 Then
                                            str1 = "Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                        ElseIf int3 = intA And intA > 2 Then
                                            str1 = str1 & ", and Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                        ElseIf int3 <> intA And intA > 2 Then
                                            str1 = str1 & ", Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                        Else
                                            str1 = str1 & " and Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                        End If
                                    End If
                                Else
                                    If boolRespective Then
                                        If int3 = 1 Then
                                            str1 = "Table" & ChrW(160) & var1
                                        ElseIf int3 = intA And intA > 2 Then
                                            str1 = str1 & ", and Table" & ChrW(160) & var1
                                        ElseIf int3 <> intA And intA > 2 Then
                                            str1 = str1 & ", Table" & ChrW(160) & var1
                                        Else
                                            str1 = str1 & " and Table" & ChrW(160) & var1
                                        End If
                                    Else
                                        If int3 = 1 Then
                                            str1 = "Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                        ElseIf int3 = intA And intA > 2 Then
                                            str1 = str1 & ", and Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                        ElseIf int3 <> intA And intA > 2 Then
                                            str1 = str1 & ", Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                        Else
                                            str1 = str1 & " and Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription") & " (Internal Standard)"
                                        End If
                                    End If
                                End If

                            End If

                        Else
                            str3 = "TableID = " & idTbl & " AND AnalyteName = '" & rowsA(Count2).Item("AnalyteDescription") & "'"
                            drows = tblN.Select(str3)
                            If drows.Length = 0 Then
                                var1 = "[NA]"
                            Else
                                var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                                If boolUseHyperlinks Then
                                    If boolRespective Then
                                        If int3 = 1 Then
                                            str1 = "Table_" & var1
                                        ElseIf int3 = intA And intA > 2 Then
                                            str1 = str1 & ", and Table_" & var1
                                        ElseIf int3 <> intA And intA > 2 Then
                                            str1 = str1 & ", Table_" & var1
                                        Else
                                            str1 = str1 & " and Table_" & var1
                                        End If
                                    Else
                                        If int3 = 1 Then
                                            str1 = "Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                        ElseIf int3 = intA And intA > 2 Then
                                            str1 = str1 & ", and Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                        ElseIf int3 <> intA And intA > 2 Then
                                            str1 = str1 & ", Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                        Else
                                            str1 = str1 & " and Table_" & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                        End If
                                    End If
                                Else
                                    If boolRespective Then
                                        If int3 = 1 Then
                                            str1 = "Table" & ChrW(160) & var1
                                        ElseIf int3 = intA And intA > 2 Then
                                            str1 = str1 & ", and Table" & ChrW(160) & var1
                                        ElseIf int3 <> intA And intA > 2 Then
                                            str1 = str1 & ", Table" & ChrW(160) & var1
                                        Else
                                            str1 = str1 & " and Table" & ChrW(160) & var1
                                        End If
                                    Else
                                        If int3 = 1 Then
                                            str1 = "Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                        ElseIf int3 = intA And intA > 2 Then
                                            str1 = str1 & ", and Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                        ElseIf int3 <> intA And intA > 2 Then
                                            str1 = str1 & ", Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                        Else
                                            str1 = str1 & " and Table" & ChrW(160) & var1 & " for " & rowsA(Count2).Item("AnalyteDescription")
                                        End If
                                    End If
                                End If

                            End If

                        End If
                    End If
                End If
            Next
        End If

        If boolRespective And intA > 1 Then
            GetTableNumberAll = str1 & ", respectively"
        Else
            GetTableNumberAll = str1
        End If

        GetTableNumberAll = str1 'ignore boolrespective for the time being

    End Function

    Function ReturnSearch(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal charSectionName As String, ByVal intPos As Int16, ByVal strFind As String, ByVal rng1 As Microsoft.Office.Interop.Word.Range, ByVal intAppFig As Int16)

        Dim var1, var2, var3, var4, var5, var6, var7, var8, var9
        Dim Count1 As Int64
        Dim Count2 As Int16
        Dim Count3 As Int64
        Dim pos1 As Int64
        Dim pos2 As Int64
        Dim ctTot As Int64
        Dim varReplace
        Dim myRange As Microsoft.Office.Interop.Word.Range
        Dim intRow As Short
        Dim dg As DataGrid
        Dim ts1 As DataGridTableStyle
        Dim dv As System.Data.DataView
        Dim dtbl As System.Data.DataTable
        Dim drows() As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim int5 As Short
        'Dim mySel As Microsoft.Office.Interop.Word.selection
        Dim mySel As Microsoft.Office.Interop.Word.Range
        Dim strFind1 As String
        Dim intIDtblStudies As Int64
        Dim strNA1 As String
        Dim strNA2 As String
        Dim strNA3 As String
        Dim strNA4 As String
        Dim tblN As System.Data.DataTable
        Dim num1 As Object
        Dim boolNum As Boolean
        Dim dt1 As Date
        Dim dt2 As Date
        Dim tblNick As System.Data.DataTable
        Dim rowsNick() As DataRow
        Dim dgv As DataGridView
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim tbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strM As String
        Dim strM1 As String
        Dim strM2 As String
        Dim intEnd As Short
        Dim intM As Short
        Dim intV As Short
        Dim tbl3 As System.Data.DataTable
        Dim rows3() As DataRow
        Dim intT As Short
        Dim bool1 As Boolean
        Dim bool2 As Boolean
        Dim idTS As Int64
        Dim idTbl As Short
        Dim strUnits As String
        Dim strDo As String
        Dim boolHit As Boolean

        Dim intGroup As Short

        Dim dtblProps As System.Data.DataTable
        Dim rowsProps() As DataRow

        Dim rowsFC() As DataRow = tblMethodValidationData.Select("ID_TBLSTUDIES = " & id_tblStudies)

        dtblProps = tblTableProperties

        idTS = id_tblStudies

        'mySel = wd.selection
        mySel = rng1

        'strFind = "zzzzzzzzzzzz"
        'strFind1 = "zzzzzzzzzzzz"
        tblN = tblTableN

        strM = frmH.lblProgress.Text

        intIDtblStudies = idTS

        Select Case intPos

            Case -10
                'strFind = "[ANALREFTABLE]"
                varReplace = ""
                Call AnalRefStandards(wd, strFind)

            Case 1
                'strFind = "[REPORTTITLE]"
                dgv = frmH.dgvReports
                dv = dgv.DataSource
                'intRow = dg.CurrentRowIndex
                str1 = "id_tblStudies = " & intIDtblStudies
                dv.RowFilter = str1
                var8 = dv.Item(0).Item("charReportTitle")
                'format var8
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'replace any hyphens with nbh
                'varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                varReplace = Replace(varReplace, "-", ChrW(173), 1, -1, CompareMethod.Text)
                'chrw(173)
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Report Title"
                strNA3 = "Choose Study & Report"
                strNA4 = "Configured Reports table"
            Case 2
                'strFind = "[REPORTNUMBER]"
                dgv = frmH.dgvReports
                dv = dgv.DataSource
                'intRow = dg.CurrentRowIndex
                str1 = "id_tblStudies = " & intIDtblStudies
                dv.RowFilter = str1
                var8 = dv.Item(0).Item("charReportNumber")
                'format var8
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Report Number"
                strNA3 = "Choose Study & Report"
                strNA4 = "Configured Reports table"
            Case 3
                'strFind = "[REPORTDRAFTDATE]"
                dgv = frmH.dgvReports
                dv = dgv.DataSource
                'intRow = dg.CurrentRowIndex
                str1 = "id_tblStudies = " & intIDtblStudies
                dv.RowFilter = str1
                var8 = dv.Item(0).Item("dtReportDraftIssueDate")
                'format var8
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Format(CDate(var8), LTextDateFormat)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Report Draft Date"
                strNA3 = "Choose Study & Report"
                strNA4 = "Configured Reports table"

            Case 4
                'strFind = "[REPORTISSUEDATE]"
                dgv = frmH.dgvReports
                dv = dgv.DataSource
                'intRow = dg.CurrentRowIndex
                str1 = "id_tblStudies = " & intIDtblStudies
                dv.RowFilter = str1
                var8 = dv.Item(0).Item("dtReportFinalIssueDate")
                'format var8
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Format(CDate(var8), LTextDateFormat)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Report Issue Date"
                strNA3 = "Choose Study & Report"
                strNA4 = "Configured Reports table"

            Case 5
                'strFind = "[ACCEPTEDANALYTICALRUNS]"
                tbl1 = tblAnalytesHome
                dv = frmH.dgvAnalyticalRunSummary.DataSource
                tbl2 = dv.ToTable
                strF = "IsIntStd = 'No'"
                Erase rows1
                rows1 = tbl1.Select(strF)
                int1 = rows1.Length
                var8 = 0

                For Count2 = 0 To int1 - 1
                    var1 = rows1(Count2).Item("AnalyteDescription")
                    strF = "ANALYTE = '" & var1 & "'"
                    Erase rows2
                    rows2 = tbl2.Select(strF)
                    int2 = rows2.Length
                    int3 = 0
                    For Count3 = 0 To int2 - 1
                        var3 = rows2(Count3).Item("Pass/Fail")
                        If StrComp(var3, "Accepted", CompareMethod.Text) = 0 Then
                            int3 = int3 + 1
                        End If
                    Next
                    If int3 > var8 Then
                        var8 = int3
                    End If
                Next
                varReplace = VerboseNumber(var8, False)
                strNA1 = charSectionName
                strNA2 = "Number of Accepted Analytical Runs"
                strNA3 = "Analytical Run Summary"
                strNA4 = "Analytical Run Summary Table"

            Case 6
                'strFind = "[NUMAPCONC]"
                tbl1 = tblAssignedSamples
                intT = 11
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = VerboseNumber(int1, False)
                End If
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision"

            Case 7
                'strFind = "[NUMAPREPLICATES]"
                tbl1 = tblAssignedSamples
                intT = 11
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC", "RUNID", "CHARANALYTE")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                    var2 = tbl2.Rows.Item(0).Item("RUNID")
                    var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                    strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                    rows1 = tbl1.Select(strF)
                    int2 = rows1.Length
                    If int2 = 0 Then
                        varReplace = "[NA]"
                    Else
                        'varReplace = VerboseNumber(int2, False)
                        '20190222 LEE: not verbose
                        varReplace = int2
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision"

            Case 8
                'strFind = "[NUMLINDILREPLICATES]"
                '20190228 LEE: Deprecated
                ''20190228 LEE:
                'varReplace = GetDilQCInfo(1, 1)
                ''intType: 1 = # of Diln QC Replicates
                ''intType: 2 = Diln QC Value(s)
                ''intType: 3 = Diln QC Dilution Factor
                'strNA1 = charSectionName
                'strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                'strNA3 = "Report Table Configuration"
                'strNA4 = "Summary of Interpolated Dilution QC Concentrations"


                'tbl1 = tblAssignedSamples
                'intT = 12

                ' ''20190222 LEE: Diln can also be 31 ad hoc stability
                ' ''strF = "(ID_TBLCONFIGREPORTTABLES = " & intT & " OR ID_TBLCONFIGREPORTTABLES = 31) AND ID_TBLSTUDIES = " & idTS
                ' ''20190222 LEE: Need to filter for diln factor <> 1
                ''strF = "(ID_TBLCONFIGREPORTTABLES = " & intT & " OR ID_TBLCONFIGREPORTTABLES = 31) AND ID_TBLSTUDIES = " & idTS & " AND ALIQUOTFACTOR <> 1"

                'strS = "ID_TBLCONFIGREPORTTABLES ASC"

                ' ''20190222 LEE: Need to find Dilution table in tblTableProperties
                'strF = GetBOOLSTATSNRFilter(12) '12 = Diln table

                'strS = "ID_TBLCONFIGREPORTTABLES ASC"

                'Try
                '    Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                '    tbl2 = dv1.ToTable("a", True, "NOMCONC", "RUNID", "CHARANALYTE")
                '    int1 = tbl2.Rows.Count
                '    If int1 = 0 Then
                '        varReplace = "[NA]"
                '    Else
                '        'loop to find largest value
                '        Dim intMax As Short = 0
                '        For Count2 = 0 To tbl2.Rows.Count - 1
                '            var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                '            var2 = tbl2.Rows.Item(0).Item("RUNID")
                '            var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                '            strF = "(ID_TBLCONFIGREPORTTABLES = " & intT & " OR ID_TBLCONFIGREPORTTABLES = 31) AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                '            rows1 = tbl1.Select(strF)
                '            int2 = rows1.Length
                '            If int2 > intMax Then
                '                intMax = int2
                '            End If
                '        Next

                '        If intMax = 0 Then
                '            varReplace = "[NA]"
                '        Else
                '            varReplace = intMax ' VerboseNumber(intmax, False)
                '        End If

                '    End If
                '    strNA1 = charSectionName
                '    strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                '    strNA3 = "Report Table Configuration"
                '    strNA4 = "Summary of Interpolated Dilution QC Concentrations"
                'Catch ex As Exception
                '    var1 = var1
                'End Try

            Case 9
                'strFind = "[NUMDILVALUE]"
                '20190228 LEE: Deprecated
                ''20190228 LEE:
                'varReplace = GetDilQCInfo(2, 1)
                ''intType: 1 = # of Diln QC Replicates
                ''intType: 2 = Diln QC Value(s)
                ''intType: 3 = Diln QC Dilution Factor
                'strNA1 = charSectionName
                'strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                'strNA3 = "Report Table Configuration"
                'strNA4 = "Summary of Interpolated Dilution QC Concentrations"

                'tbl1 = tblAssignedSamples
                'intT = 12
                ''20190220 LEE: DilnFactor table is now adhocstability
                ''strF = "ID_TBLCONFIGREPORTTABLES = 31 AND ID_TBLSTUDIES = " & idTS
                ''20190222 LEE: Diln can also be 31 ad hoc stability

                'strS = "ID_TBLCONFIGREPORTTABLES ASC"

                ' ''20190222 LEE: Need to find Dilution table in tblTableProperties
                'strF = GetBOOLSTATSNRFilter(12) '12 = Diln table

                ''strF = "(ID_TBLCONFIGREPORTTABLES = " & intT & " OR ID_TBLCONFIGREPORTTABLES = 31) AND ID_TBLSTUDIES = " & idTS
                'strS = "ID_TBLCONFIGREPORTTABLES ASC"
                'Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                'tbl2 = dv1.ToTable("a", True, "NOMCONC", "CHARANALYTE")
                'int1 = tbl2.Rows.Count
                'If int1 = 0 Then
                '    varReplace = "[NA]"
                'Else
                '    'loop to find all values
                '    For Count2 = 0 To tbl2.Rows.Count - 1
                '        var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                '        If Count2 = 0 Then
                '            varReplace = var1
                '        ElseIf Count2 = tbl2.Rows.Count - 1 Then
                '            varReplace = varReplace & " and " & var1
                '        Else
                '            varReplace = varReplace & ", " & var1
                '        End If
                '    Next
                '    var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                '    varReplace = var1
                'End If
                'strNA1 = charSectionName
                'strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                'strNA3 = "Report Table Configuration"
                'strNA4 = "Summary of Interpolated Dilution QC Concentrations"

            Case 10

                '20190228 LEE: Deprecated
                ''strFind = "[NUMDILFACTOR]"
                ''20190228 LEE:
                'varReplace = GetDilQCInfo(3, 1)
                ''intType: 1 = # of Diln QC Replicates
                ''intType: 2 = Diln QC Value(s)
                ''intType: 3 = Diln QC Dilution Factor
                'strNA1 = charSectionName
                'strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                'strNA3 = "Report Table Configuration"
                'strNA4 = "Summary of Interpolated Dilution QC Concentrations"

                ''20180724 LEE:
                ''Redo this

                'tbl1 = tblAssignedSamples
                'tbl2 = tblAnalysisResultsHome
                'intT = 12
                ''20190220 LEE: DilnFactor table is now adhocstability
                ''strF = "(ID_TBLCONFIGREPORTTABLES = " & intT & " OR ID_TBLCONFIGREPORTTABLES = 31) AND ID_TBLSTUDIES = " & idTS & " AND ALIQUOTFACTOR <> 1"

                ' ''20190222 LEE: Need to find Dilution table in tblTableProperties
                'strF = GetBOOLSTATSNRFilter(12) '12 = Diln table

                'strS = "ALIQUOTFACTOR DESC"
                'Dim rowsD() As DataRow = tbl1.Select(strF, strS)
                'Dim dtblT As DataTable = rowsD.CopyToDataTable
                'Dim dtblD As DataTable = dtblT.DefaultView.ToTable("a", True, "ALIQUOTFACTOR")
                ''Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                'Dim dv1 As System.Data.DataView = New DataView(dtblD, "", strS, DataViewRowState.CurrentRows)
                'If dv1.Count = 0 Then
                '    varReplace = "[NA]"
                'Else
                '    Dim intDC As Short = 0
                '    For Count2 = 0 To dv1.Count - 1
                '        var1 = dv1(Count2).Item("ALIQUOTFACTOR")
                '        'var2 = CInt(1 / var1)
                '        var2 = GetDilnFactor(CDec(var1))
                '        If dv1.Count > 2 Then
                '            If Count2 = 0 Then
                '                varReplace = var2
                '            ElseIf Count2 = dv1.Count - 1 Then
                '                varReplace = varReplace & ", and " & var2
                '            Else
                '                varReplace = varReplace & ", " & var2
                '            End If
                '        Else
                '            If Count2 = 0 Then
                '                varReplace = var2
                '            Else
                '                varReplace = varReplace & " and " & var2
                '            End If
                '        End If

                '    Next

                '    'varReplace = var2
                'End If
                'strNA1 = charSectionName
                'strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                'strNA3 = "Report Table Configuration"
                'strNA4 = "Summary of Interpolated Dilution QC Concentrations"

                ''strFind = "[NUMDILFACTOR]"
                'tbl1 = tblAssignedSamples
                'tbl2 = tblAnalysisResultsHome
                'intT = 12
                'strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                'strS = "ID_TBLCONFIGREPORTTABLES ASC"
                'Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                'If dv1.Count = 0 Then
                '    varReplace = "[NA]"
                'Else
                '    var1 = dv1(0).Item("RUNID")
                '    var2 = dv1(0).Item("RUNSAMPLESEQUENCENUMBER")
                '    strF = "RUNID = " & var1 & " AND RUNSAMPLESEQUENCENUMBER = " & var2 ' & " AND ALIQUOTFACTOR <> 1"
                '    rows2 = tbl2.Select(strF)
                '    var1 = rows2(0).Item("ALIQUOTFACTOR")
                '    var2 = CInt(1 / var1)
                '    varReplace = var2
                'End If
                'strNA1 = charSectionName
                'strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                'strNA3 = "Report Table Configuration"
                'strNA4 = "Summary of Interpolated Dilution QC Concentrations"

            Case 11
                'strFind = "[FULLVALIDATIONNUMBER]"
                int1 = FindRowDV("Validation Corporate Study/Project Number", frmH.dgvMethodValData.DataSource)
                dv = frmH.dgvMethodValData.DataSource
                var8 = dv(int1).Item(1)
                'var8 = frmH.dgvMethodValData.Item(int1, 1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Validation Corporate Study/Project Number"
                strNA3 = "Method Validation Data"
                strNA4 = "Validation Corporate Study/Project Number Entry"

            Case 12
                'strFind = "[NUMSEREPLICATES]"
                tbl1 = tblAssignedSamples
                intT = 15
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC", "RUNID", "CHARANALYTE")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                    var2 = tbl2.Rows.Item(0).Item("RUNID")
                    var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                    strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                    rows1 = tbl1.Select(strF)
                    int2 = rows1.Length
                    If int2 = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = VerboseNumber(int2, False)
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "Summary of Suppression/Enhancement"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Suppression/Enhancement"

            Case 13
                'strFind = "[NUMUNIQUELOTS]"
                tbl1 = tblAssignedSamples
                intT = 17
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "CHARHELPER1")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = VerboseNumber(int1, False)

                End If
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments"

            Case 14
                'strFind = "[NUMUNIQUEREPLICATES]"
                tbl1 = tblAssignedSamples
                intT = 17
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC", "RUNID", "CHARANALYTE", "CHARHELPER1")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                    var2 = tbl2.Rows.Item(0).Item("RUNID")
                    var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                    var4 = tbl2.Rows.Item(0).Item("CHARHELPER1")
                    strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "' AND CHARHELPER1 = '" & var4 & "'"
                    rows1 = tbl1.Select(strF)
                    int2 = rows1.Length
                    If int2 = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = VerboseNumber(int2, False)
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments"

            Case 15
                'strFind = "[NUMCOADREPLICATES]"
                tbl1 = tblAssignedSamples
                intT = 24
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC", "RUNID", "CHARANALYTE", "CHARHELPER1")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                    var2 = tbl2.Rows.Item(0).Item("RUNID")
                    var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                    var4 = tbl2.Rows.Item(0).Item("CHARHELPER1")
                    strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "' AND CHARHELPER1 = '" & var4 & "'"
                    rows1 = tbl1.Select(strF)
                    int2 = rows1.Length
                    If int2 = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = VerboseNumber(int2, False)
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated QC Std Concs Containing Coadministered Compounds"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated QC Std Concs Containing Coadministered Compounds"

            Case 16
                'strFind = "[NUMRTSREPLICATES]"
                tbl1 = tblAssignedSamples
                intT = 18
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC", "RUNID", "CHARANALYTE", "CHARHELPER1")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                    var2 = tbl2.Rows.Item(0).Item("RUNID")
                    var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                    var4 = tbl2.Rows.Item(0).Item("CHARHELPER1")
                    strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "' AND CHARHELPER1 = '" & var4 & "'"
                    rows1 = tbl1.Select(strF)
                    int2 = rows1.Length
                    If int2 = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = VerboseNumber(int2, False)
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "Summary of [Temp Descr] Stability in Matrix"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of [Temp Descr] Stability in Matrix"

            Case 17
                'strFind = "[NUMRTSHOURS]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties
                intT = 18
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = NZ(rows1(0).Item("CHARTIMEPERIOD"), "")
                var2 = NZ(rows1(0).Item("CHARTIMEFRAME"), "")

                'look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = LCase(CStr(var1)) & " " & var2

                    '    'look for a sequence of characters that is numeric
                    '    var3 = ""
                    '    var4 = ""
                    '    bool1 = False 'Start
                    '    bool2 = False 'End
                    '    For Count1 = 1 To int1
                    '        var2 = Mid(var1, Count1, 1)
                    '        If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                    '            var2 = "a"
                    '        End If
                    '        If IsNumeric(var2) Then
                    '            var3 = var3 & var2
                    '            If IsNumeric(var3) Then
                    '                var4 = var3
                    '                bool1 = True
                    '            Else
                    '            End If
                    '        Else
                    '            If bool1 Then
                    '                bool2 = True
                    '            End If
                    '        End If
                    '        If bool1 And bool2 Then
                    '            Exit For
                    '        End If
                    '    Next
                End If
                'If bool1 = False Then
                '    varReplace = "[NA]"
                'Else
                '    varReplace = VerboseNumber(var4, False)
                'End If
                strNA1 = charSectionName
                strNA2 = "Summary of [Temp Descr] Stability in Matrix"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of [Temp Descr] Stability in Matrix PERIODTEMP"

            Case 18
                'strFind = "[NUMFTREPLICATES]"
                tbl1 = tblAssignedSamples
                intT = 19
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC", "RUNID", "CHARANALYTE")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                    var2 = tbl2.Rows.Item(0).Item("RUNID")
                    var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                    strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                    rows1 = tbl1.Select(strF)
                    int2 = rows1.Length
                    If int2 = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = VerboseNumber(int2, False)
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "Summary of Freeze/Thaw [#Cycles] Stability in Matrix"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Freeze/Thaw [#Cycles] Stability in Matrix"

            Case 19
                ''strFind = "[INTERNALSTANDARD]"
                ''determine if there is more than one INT STD
                'tbl1 = tblAnalytesHome
                'strF = "IsIntStd = 'Yes'"
                'rows1 = tbl1.Select(strF)
                'int1 = rows1.Length
                'If int1 = 0 Then
                '    var8 = "[NA]"
                'Else
                '    var8 = rows1(0).Item("AnalyteDescription")

                '    str1 = var8
                '    If int1 > 1 Then
                '        For Count2 = 1 To int1 - 1
                '            var1 = rows1(Count2).Item("AnalyteDescription")
                '            var2 = Replace(var1, "-", NBHReal, 1, -1, CompareMethod.Text)
                '            'replace spaces with nbs
                '            var2 = Replace(var2, " ", ChrW(160), 1, -1, CompareMethod.Text)
                '            If Count2 = int1 - 1 And int1 - 1 > 2 Then
                '                str1 = str1 & ", and " & var2
                '            ElseIf Count2 <> int1 - 1 And int1 - 1 > 2 Then
                '                str1 = str1 & ", " & var2
                '            Else
                '                str1 = str1 & " and " & var2
                '            End If
                '        Next
                '    End If
                '    var8 = str1

                'End If


                '*****

                '20180827 LEE

                'strFind = "[INTERNALSTANDARD]"
                'determine if there is more than one analyte

                str1 = ""
                Dim ct1 As Short
                ct1 = 0

                '20161004 LEE: start using Analyte Groups tblAnalyteGroups
                'first inventory included analytes
                'Dim ctA As Short = tblAnalyteGroups.Rows.Count
                'Dim rowsAG() As DataRow = tblAnalyteGroups.Select("", "INTORDER ASC", DataViewRowState.CurrentRows)

                Dim ctA As Short = tblAnalytesHome.Rows.Count
                Dim rowsAG() As DataRow = tblAnalytesHome.Select("", "INTORDER ASC", DataViewRowState.CurrentRows)

                'tblanalytegroups is sorted by intorder

                Dim arrA(4, ctA)

                For Count2 = 0 To ctA - 1
                    'var1 = rowsAG(Count2).Item("ANALYTEDESCRIPTION")
                    'var2 = rowsAG(Count2).Item("ANALYTEDESCRIPTION_C")
                    var1 = rowsAG(Count2).Item("ORIGINALANALYTEDESCRIPTION")
                    var2 = rowsAG(Count2).Item("ANALYTEDESCRIPTION")
                    var3 = rowsAG(Count2).Item("INTSTD")
                    var4 = rowsAG(Count2).Item("CHARUSERIS")
                    'ensure analyte is included in at least one table
                    If UseAnalyte(CStr(var2)) Then
                        ct1 = ct1 + 1
                        arrA(1, ct1) = var1
                        arrA(2, ct1) = var2
                        arrA(3, ct1) = var3
                        arrA(4, ct1) = var4
                    End If
                Next

                ctA = ct1

                For Count2 = 1 To ct1
                    'var1 = arrA(1, Count2)
                    '20170929 LEE: Should be ANALYTEDESCRIPTION_C for multiple matrix
                    'var1 = arrA(3, Count2)
                    var1 = arrA(4, Count2)
                    'replace hyphens with nbh
                    'var2 = Replace(var1, "-", NBH, 1, -1, CompareMethod.Text)
                    'replace spaces with nbs
                    var2 = Replace(var1, " ", ChrW(160), 1, -1, CompareMethod.Text)
                    var2 = Replace(var2, "-", NBHReal, 1, -1, CompareMethod.Text)

                    If Count2 = 1 Then
                        str1 = var2 'arrAnalytes(1, Count2)
                    Else
                        If Count2 = ct1 And ct1 > 2 Then
                            str1 = str1 & ", and " & var2
                        ElseIf Count2 <> ct1 And ct1 > 2 Then
                            str1 = str1 & ", " & var2
                        Else
                            str1 = str1 & " and " & var2
                        End If
                    End If

                Next
                If ctA = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = str1
                End If


                '*****


                'varReplace = var8
                strNA1 = charSectionName
                strNA2 = "Internal Standard"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 20
                'strFind = "[LLOQ]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("LLOQ", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    num1 = var8 'dv.Item(int1).Item(1)
                    num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                    boolNum = True
                    varReplace = DisplayNum(var8, LSigFig, False)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "LLOQ"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 21
                'strFind = "[LLOQUNITS]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("LLOQ Units", dv)
                var8 = dv.Item(int1).Item(1)

                int2 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                str1 = NZ(frmH.dgvStudyConfig(1, int2).Value, "")

                If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                Else
                    var8 = str1
                End If

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "LLOQ Units"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 22
                'strFind = "[ULOQ]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("ULOQ", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    num1 = var8 'dv.Item(int1).Item(1)
                    num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                    boolNum = True
                    varReplace = DisplayNum(var8, LSigFig, False)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "ULOQ"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 23
                'strFind = "[ULOQUNITS]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("ULOQ Units", dv)
                var8 = dv.Item(int1).Item(1)

                int2 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                str1 = NZ(frmH.dgvStudyConfig(1, int2).Value, "")

                If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                Else
                    var8 = str1
                End If

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "ULOQ Units"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 24
                'strFind = "[CALIBRATIONLEVELS]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("Calibration Levels", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Calibration Levels"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 25
                'strFind = "[REGRESSION]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("REGRESSION", dv)
                var8 = NZ(dv.Item(int1).Item(1), "[NA]")
                'var1 = LCase(NZ(dv.Item(int1).Item(1), "[NA]"))
                'If StrComp(var1, "Linear", CompareMethod.Text) = 0 Then
                '    var8 = "least-squares " & var1
                'Else
                '    var8 = LCase(NZ(dv.Item(int1).Item(1), "[NA]"))
                'End If

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Regression"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 26
                'strFind = "[WEIGHTING]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("WEIGHTING", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Regression Weighting"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 27

                'strFind = "[MINIMUMR2]"
                'dv = frmH.dgvWatsonAnalRef.DataSource
                'int1 = FindRowDV("Minimum r^2", dv)
                'var8 = dv.Item(int1).Item(1)
                Dim rows() = tblRegCon.Select("", "RSQUARED ASC")
                If rows.Length = 0 Then
                    var8 = "[None]"
                Else
                    var8 = rows(0).Item("RSQUARED") ' Format(GetMin(arrRegCon, Count2), str1)
                End If
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    If IsNumeric(var8) Then
                        'sometimes R2 is the wrong sigfigs in dgvWatsonAnalRef
                        var1 = SigFigOrDecString(var8, LR2SigFigs, True)
                        str1 = GetRegrDecStr(LR2SigFigs)
                        var3 = Format(CDbl(var1), str1)
                        varReplace = var3
                    Else
                        varReplace = var8
                    End If

                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Minimum r^2"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 28
                'strFind = "[ANALYTEMEANACCURACYMIN]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("Analyte Mean Accuracy Min", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Analyte Mean Accuracy Min"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 29
                'strFind = "[ANALYTEMEANACCURACYMAX]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("Analyte Mean Accuracy Max", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Analyte Mean Accuracy Max"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 30
                'strFind = "[ANALYTEPRECISIONMIN]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("Analyte Precision Min", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Analyte Precision Min"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 31
                'strFind = "[ANALYTEPRECISIONMAX]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("Analyte Precision Max", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Analyte Precision Max"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 32
                'strFind = "[QCMEANACCURACYMIN]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("QC Mean Accuracy Min", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "QC Mean Accuracy Min"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 33
                'strFind = "[QCMEANACCURACYMAX]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("QC Mean Accuracy Max", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "QC Mean Accuracy Max"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 34
                'strFind = "[QCPRECISIONMIN]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("QC Precision Min", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "QC Precision Min"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 35
                'strFind = "[QCPRECISIONMAX]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("QC Precision Max", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "QC Precision Max"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Analytical Reference Standard Table"

            Case 36
                'strFind = "[WATSONSTUDYID]"
                dgv = frmH.dgvDataWatson
                dv = dgv.DataSource
                int1 = FindRowDV("Watson Study ID", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Watson Study ID"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 37
                'strFind = "[WATSONPROJECTID]"
                dgv = frmH.dgvDataWatson
                dv = dgv.DataSource
                int1 = FindRowDV("Watson Project ID", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Watson Project ID"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 38
                'strFind = "[MATRIX]"
                'dg = frmH.dgvDataWatson
                'dv = dg.DataSource
                'int1 = FindRowDV("Matrix", dv)
                ''var8 = dg.Item(int1, 1)
                ''var8 = UnCapit(NZ(dv.Item(int1).Item(1), ""), True)
                'var8 = Trim(LCase(NZ(dv.Item(int1).Item(1), "[NA]")))

                'now find information from dgvDataWatson
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource 'intI is analyte column in dgvMethodValData
                int1 = FindRowDVByCol("Matrix", dv, "Item")
                'var8 = dv.Item(int1).Item(strAnal)
                var8 = Trim(LCase(NZ(dv.Item(int1).Item(1), "[NA]")))

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Matrix"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 39
                'strFind = "[SAMPLESIZE]"
                'dgv = frmH.dgvDataWatson
                '20180724 LEE:
                'new logic
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Sample Size", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                If var8 = 0 Then 'try Meth Val data
                    dgv = frmH.dgvMethodValData
                    dv = dgv.DataSource
                    int1 = FindRowDV("Sample Size", dv)
                    'var8 = dg.Item(int1, 1)
                    var8 = dv.Item(int1).Item(1)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If

                End If

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Sample Size"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 40
                ''strFind = "[SPECIES]"

                'now find information from dgvDataWatson
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource 'intI is analyte column in dgvMethodValData
                int1 = FindRowDVByCol("Species", dv, "Item")
                var8 = Trim(NZ(dv.Item(int1).Item(1), "[NA]"))

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Species"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 41
                'strFind = "[INTEGRATIONTYPE]"
                dgv = frmH.dgvDataWatson
                dv = dgv.DataSource
                int1 = FindRowDV("Integration Type", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Integration Type"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 42
                'strFind = "[INITIALEXTRACTIONDATE]"
                dgv = frmH.dgvDataWatson
                dv = dgv.DataSource
                int1 = FindRowDV("Initial Extraction Date", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Format(CDate(var8), LTextDateFormat)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Initial Extraction Date"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 43
                'strFind = "[LASTEXTRACTIONDATE]"
                dgv = frmH.dgvDataWatson
                dv = dgv.DataSource
                int1 = FindRowDV("Last Extraction Date", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Format(CDate(var8), LTextDateFormat)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Last Extraction Date"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 44
                'strFind = "[ASSAYTECHNIQUE]"
                var8 = frmH.cbxAssayTechnique.Text
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Assay Technique"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Assay Technique Dropdown Box"

            Case 45
                'strFind = "[ASSAYTECHNIQUEACRONYM]"
                var8 = frmH.cbxAssayTechniqueAcronym.Text
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Assay Technique Acronym"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Assay Technique Acronym Dropdown Box"

            Case 46
                'strFind = "[ANTICOAGULANT]"
                str1 = NZ(frmH.cbxAnticoagulant.Text, "")
                'determine if text should be capitalized
                var8 = "[NA]"
                If Len(str1) = 0 Then
                Else
                    int1 = Len(str1)
                    int2 = 0
                    For Count2 = 1 To int1
                        var1 = CStr(Mid(str1, Count2, 1))
                        var2 = Asc(var1)
                        If var2 > 64 And var2 < 91 Then
                            int2 = int2 + 1
                        Else
                            Exit For
                        End If
                    Next
                    If int2 > 2 Then 'probably is an acronym, leave capitalized
                        var8 = str1
                    Else
                        var8 = str1
                    End If
                End If
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Anticoagulant/Preservative"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Anticoagulant Dropdown Box"

            Case 47
                'strFind = "[SAMPLESIZEUNITS]"
                'var8 = frmH.cbxSampleSizeUnits.Text
                int1 = FindRowDV("Sample Size Units", frmH.dgvMethodValData.DataSource)
                dv = frmH.dgvMethodValData.DataSource
                'var8 = frmH.dgvMethodValData(int1, 1).Value
                var8 = dv(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Sample Size Units"
                strNA3 = "Method Validation Data"
                strNA4 = "Sample Size Units Table Entry"

            Case 48
                'strFind = "[SUBMITTEDTO]"
                dtbl = tblCorporateAddresses
                var8 = frmH.cbxSubmittedTo.Text
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    str1 = var8
                    str2 = "charNickName = '" & str1 & "'"
                    tblNick = tblCorporateNickNames
                    rowsNick = tblNick.Select(str2)
                    var1 = rowsNick(0).Item("id_tblCorporateNickNames")
                    str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1 OR " _
        & "id_tblCorporateNickNames = " & var1 & " AND charAddressLabel = 'Company Name'"
                    drows = dtbl.Select(str2, "id_tbladdresslabels ASC")
                    int1 = drows.Length
                    If int1 < 1 Then
                        varReplace = "[NA]"
                    Else
                        var8 = ""
                        For Count2 = 0 To int1 - 1
                            var8 = var8 & drows(Count2).Item("charValue") & " "
                        Next
                        'remove blank space at end
                        var8 = Trim(var8)
                        varReplace = var8
                    End If
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Submitted To"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Submitted To Textbox"

            Case 49
                'strFind = "[INSUPPORTOF]"
                dtbl = tblCorporateAddresses
                var8 = frmH.cbxInSupportOf.Text
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    str1 = var8
                    str2 = "charNickName = '" & str1 & "'"
                    tblNick = tblCorporateNickNames
                    rowsNick = tblNick.Select(str2)
                    var1 = rowsNick(0).Item("id_tblCorporateNickNames")
                    str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1 OR " _
                            & "id_tblCorporateNickNames = " & var1 & " AND charAddressLabel = 'Company Name'"
                    drows = dtbl.Select(str2, "id_tbladdresslabels ASC")
                    int1 = drows.Length
                    If int1 < 1 Then
                        varReplace = "[NA]"
                    Else
                        var8 = ""
                        For Count2 = 0 To int1 - 1
                            var8 = var8 & drows(Count2).Item("charValue") & " "
                        Next
                        'remove blank space at end
                        var8 = Trim(var8)
                        varReplace = var8
                    End If

                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "In Support Of"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "In Support Of Textbox"

            Case 50
                'strFind = "[SUBMITTEDBY]"
                dtbl = tblCorporateAddresses
                var8 = frmH.cbxSubmittedBy.Text
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    str1 = var8
                    str2 = "charNickName = '" & str1 & "'"
                    tblNick = tblCorporateNickNames
                    rowsNick = tblNick.Select(str2)
                    var1 = rowsNick(0).Item("id_tblCorporateNickNames")
                    str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1 OR " _
                            & "id_tblCorporateNickNames = " & var1 & " AND charAddressLabel = 'Company Name'"   ' & " And boolInclude = " & -1"
                    drows = dtbl.Select(str2, "id_tbladdresslabels ASC")
                    int1 = drows.Length
                    If int1 < 1 Then
                        varReplace = "[NA]"
                    Else
                        var8 = ""
                        For Count2 = 0 To int1 - 1
                            var8 = var8 & drows(Count2).Item("charValue") & " "
                        Next
                        'remove blank space at end
                        var8 = Trim(var8)
                        varReplace = var8
                    End If

                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Submitted By"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Submitted By Textbox"

            Case 51
                'strFind = "[CORPORATESTUDY/PROJECTNUMBER]"
                'dtbl = tblData
                dtbl = tblData
                str2 = "id_tblStudies = " & intIDtblStudies
                drows = dtbl.Select(str2)
                int1 = drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var8 = NZ(drows(0).Item("charCorporateStudyID"), "")
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                End If
                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Corporate Study/Project Number"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data table"

            Case 52
                'strFind = "[PROTOCOLNUMBER]"
                'dtbl = tblData
                dtbl = tblData
                str2 = "id_tblStudies = " & intIDtblStudies
                drows = dtbl.Select(str2)
                int1 = drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var8 = drows(0).Item("charProtocolNumber")
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                End If
                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Protocol Number"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data table"

            Case 53
                'strFind = "[SPONSORSTUDYNUMBER]"
                'dtbl = tblData
                dtbl = tblData
                str2 = "id_tblStudies = " & intIDtblStudies
                drows = dtbl.Select(str2)
                int1 = drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var8 = drows(0).Item("charSponsorStudyNumber")
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                End If
                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Sponsor Study Number"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data table"

            Case 54
                'strFind = "[DATAARCHIVALLOCATION]"
                'dtbl = tblData
                dtbl = tblData
                str2 = "id_tblStudies = " & intIDtblStudies
                drows = dtbl.Select(str2)
                int1 = drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var8 = drows(0).Item("charDataArchivalLocation")
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Data Archival Location"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data table"

            Case 55
                'strFind = "[DATEFIRSTSAMPLESRECEIVED]"
                If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    dv = frmH.dgvSampleReceiptWatson.DataSource
                    strNA4 = "StudyDoc Sample Receipt table"
                Else
                    dv = frmH.dgvSampleReceipt.DataSource
                    strNA4 = "Watson Sample Receipt table"
                End If
                int1 = dv.Count 'drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                        var8 = NZ(dv(0).Item("Date Received"), "")
                    Else
                        var8 = NZ(dv(0).Item("dtShipmentReceived"), "")
                    End If
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    Else
                        If IsDate(var8) Then
                            varReplace = Format(CDate(var8), LTextDateFormat)
                        Else
                            varReplace = "[NA]"
                        End If
                    End If
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Date Samples Received - First Shipment"
                strNA3 = "Sample Receipt"

            Case 56
                'strFind = "[METHODDEMONSTRATEDFREEZE/THAWCYCLES]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Demonstrated Freeze/Thaw Cycles", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Demonstrated Freeze/Thaw Cycles"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 57
                'strFind = "[METHODMAXIMUMNUMBEROFFREEZE/THAWCYCLES]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Maximum # of Freeze/thaw Cycles", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Maximum # of Freeze/thaw Cycles"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 58

                '20190213 LEE: Deprecated

                'strFind = "[METHODSTABILITYUNDERSTORAGECONDITIONS]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                'int1 = FindRowDV("Stability Under Storage Conditions", dv)
                int1 = FindRowDV("Bench-top Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Bench-top Stability" 'Stability Under Storage Conditions"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 59
                'strFind = "[METHODISSTABILITY>=MAXIMUMSTORAGEDURATION]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Is Stability >= Maximum Storage Duration", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Is Stability >= Maximum Storage Duration"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 60
                'strFind = "[METHODCORPORATESTUDY/PROJECTNUMBER]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Validation Corporate Study/Project Number", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Validation Corporate Study/Project Number"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 61
                'strFind = "[METHODPROTOCOLNUMBER]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Validation Protocol Number", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Validation Protocol Number"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 62
                'strFind = "[METHODMETHODVALIDATIONTITLE]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Validation Report Title", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Validation Report Title"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 63
                'strFind = "[METHODSPONSORMETHODVALIDATIONSTUDYNUMBER]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Sponsor Method Validation Study Number", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Sponsor Method Validation Study Number"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 64
                'strFind = "[METHODSPONSORMETHODVALIDATIONTITLE]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Sponsor Method Validation Title", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Sponsor Method Validation Title"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 65
                'strFind = "[METHODASSAYPROCEDUREDESCRIPTION]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Extraction Procedure Description", dv)
                'var8 = dg.Item(int1, 0)
                var8 = dv.Item(int1).Item(1)

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    'varReplace = UnCapit(Trim(var8), False)
                    '20181015 LEE:
                    'Don't make this uncapitalized
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Extraction Procedure Description"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 66
                'strFind = "[FINALREPORTREVIEWDATE]"
                'finalreportdate comes from QA event table Final Report critical phase
                dg = frmH.dgQATable
                dv = dg.DataSource
                intRow = FindRowDVByCol("Final Report", dv, "charUserLabel")
                If intRow = -1 Then
                    var8 = "[None]"
                Else
                    var8 = NZ(dv.Item(intRow).Item("dtColumn1"), "")
                End If
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                    'add entries to arrReportNA
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Format(CDate(var8), LTextDateFormat)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Final Report Review Date"
                strNA3 = "QA Events Table"
                strNA4 = "QA Events Table: Final Report"

            Case 67
                'strFind = "[LMAPPENDIXNUMBER]"
                tbl1 = tblAppendix
                strF = "AppendixName = 'LM'"
                rows1 = tbl1.Select(strF)
                If rows1.Length = 0 Then
                    varReplace = "[NA]"
                Else
                    If boolUseHyperlinks Then
                        var1 = rows1(0).Item("AppendixNumber")
                        var2 = "Appendix_" & var1 'AppendixLetter(var1)
                        str1 = var2
                        If rows1.Length = 1 Then
                        ElseIf rows1.Length = 2 Then
                            var1 = rows1(1).Item("AppendixNumber")
                            var2 = "Appendix_" & var1 'AppendixLetter(var1)
                            str1 = str1 & " and " & var2
                        ElseIf rows1.Length > 2 Then
                            For Count2 = 1 To rows1.Length - 2
                                var1 = rows1(Count2).Item("AppendixNumber")
                                var2 = "Appendix_" & var1 'AppendixLetter(var1)
                                str1 = str1 & ", " & var2
                            Next
                            var1 = rows1(rows1.Length - 1).Item("AppendixNumber")
                            var2 = "Appendix_" & var1 'AppendixLetter(var1)
                            str1 = str1 & ", and " & var2

                        End If
                    Else
                        var1 = rows1(0).Item("AppendixNumber")
                        var2 = "Appendix" & ChrW(160) & var1 'AppendixLetter(var1)
                        str1 = var2
                        If rows1.Length = 1 Then
                        ElseIf rows1.Length = 2 Then
                            var1 = rows1(1).Item("AppendixNumber")
                            var2 = "Appendix" & ChrW(160) & var1 'AppendixLetter(var1)
                            str1 = str1 & " and " & var2
                        ElseIf rows1.Length > 2 Then
                            For Count2 = 1 To rows1.Length - 2
                                var1 = rows1(Count2).Item("AppendixNumber")
                                var2 = "Appendix" & ChrW(160) & var1 'AppendixLetter(var1)
                                str1 = str1 & ", " & var2
                            Next
                            var1 = rows1(rows1.Length - 1).Item("AppendixNumber")
                            var2 = "Appendix" & ChrW(160) & var1 'AppendixLetter(var1)
                            str1 = str1 & ", and " & var2

                        End If
                    End If

                    varReplace = str1 'AppendixLetter(int1)
                End If

                strNA1 = charSectionName
                strNA2 = "Laboratory Method Appendix Letter"
                strNA3 = "Generated by StudyDoc"
                strNA4 = "Laboratory Method Appendix Letter"


            Case 68
                'strFind = "[ANALYTICALRUNTABLENUMBER]"
                idTbl = 1
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Analytical Run Table Number"
                strNA3 = "Generated by StudyDoc"
                strNA4 = "Analytical Run Table Number"

            Case 69
                ''strFind = "[REPRESENTATIVEWATSONRUNNUMBERAPPENDIXLETTER]"
                'strFind = "[REPRCHROMAPPENDIXLETTER]"
                tbl1 = tblAppendix
                strF = "AppendixName = 'Chromatogram'"
                Erase rows1
                rows1 = tbl1.Select(strF)
                If rows1.Length = 0 Then
                    varReplace = "[NA]"
                Else



                    var1 = rows1(0).Item("AppendixNumber")

                    ''''wdd.visible = True

                    If boolUseHyperlinks Then
                        var2 = "Appendix_" & var1 'AppendixLetter(var1)
                        str1 = var2
                        If rows1.Length = 1 Then

                        ElseIf rows1.Length = 2 Then
                            var1 = rows1(1).Item("AppendixNumber")
                            var2 = "Appendix_" & var1 'AppendixLetter(var1)
                            str1 = str1 & " and " & var2
                        ElseIf rows1.Length > 2 Then
                            For Count2 = 1 To rows1.Length - 2
                                var1 = rows1(Count2).Item("AppendixNumber")
                                var2 = "Appendix_" & var1 'AppendixLetter(var1)
                                str1 = str1 & ", " & var2
                            Next
                            var1 = rows1(rows1.Length - 1).Item("AppendixNumber")
                            var2 = "Appendix_" & var1 'AppendixLetter(var1)
                            str1 = str1 & ", and " & var2

                        End If
                    Else
                        var2 = "Appendix" & ChrW(160) & var1 'AppendixLetter(var1)
                        str1 = var2
                        If rows1.Length = 1 Then

                        ElseIf rows1.Length = 2 Then
                            var1 = rows1(1).Item("AppendixNumber")
                            var2 = "Appendix" & ChrW(160) & var1 'AppendixLetter(var1)
                            str1 = str1 & " and " & var2
                        ElseIf rows1.Length > 2 Then
                            For Count2 = 1 To rows1.Length - 2
                                var1 = rows1(Count2).Item("AppendixNumber")
                                var2 = "Appendix" & ChrW(160) & var1 'AppendixLetter(var1)
                                str1 = str1 & ", " & var2
                            Next
                            var1 = rows1(rows1.Length - 1).Item("AppendixNumber")
                            var2 = "Appendix" & ChrW(160) & var1 'AppendixLetter(var1)
                            str1 = str1 & ", and " & var2

                        End If
                    End If

                    varReplace = str1
                End If

                strNA1 = charSectionName
                strNA2 = "Representative Watson Run Appendix Letter"
                strNA3 = "Generated by StudyDoc"
                strNA4 = "Representative Watson Run Appendix Letter"

            Case 70
                'strFind = "[ANALYTE]"
                'determine if there is more than one analyte

                str1 = ""
                Dim ct1 As Short
                ct1 = 0

                '20161004 LEE: start using Analyte Groups tblAnalyteGroups
                'first inventory included analytes
                'Dim ctA As Short = tblAnalyteGroups.Rows.Count
                'Dim rowsAG() As DataRow = tblAnalyteGroups.Select("", "INTORDER ASC", DataViewRowState.CurrentRows)

                Dim ctA As Short = tblAnalytesHome.Rows.Count
                Dim rowsAG() As DataRow = tblAnalytesHome.Select("", "INTORDER ASC", DataViewRowState.CurrentRows)

                Dim arrA(3, ctA)

                For Count2 = 0 To ctA - 1
                    var1 = rowsAG(Count2).Item("ORIGINALANALYTEDESCRIPTION")
                    var2 = rowsAG(Count2).Item("ANALYTEDESCRIPTION")
                    var3 = rowsAG(Count2).Item("CHARUSERANALYTE")
                    'ensure analyte is included in at least one table
                    If UseAnalyte(CStr(var2)) Then
                        ct1 = ct1 + 1
                        arrA(1, ct1) = var1
                        arrA(2, ct1) = var2
                        arrA(3, ct1) = var3
                    End If
                Next

                ctA = ct1

                For Count2 = 1 To ct1
                    var1 = arrA(1, Count2)
                    '20170929 LEE: Should be ANALYTEDESCRIPTION_C for multiple matrix
                    var1 = arrA(2, Count2)
                    '20180829 LEE:
                    'now use CHARUSERANALYTE
                    var1 = arrA(3, Count2)
                    'replace hyphens with nbh
                    'var2 = Replace(var1, "-", NBH, 1, -1, CompareMethod.Text)
                    'replace spaces with nbs
                    var2 = Replace(var1, " ", ChrW(160), 1, -1, CompareMethod.Text)
                    var2 = Replace(var2, "-", NBHReal, 1, -1, CompareMethod.Text)

                    If Count2 = 1 Then
                        str1 = var2 'arrAnalytes(1, Count2)
                    Else
                        If Count2 = ct1 And ct1 > 2 Then
                            str1 = str1 & ", and " & var2
                        ElseIf Count2 <> ct1 And ct1 > 2 Then
                            str1 = str1 & ", " & var2
                        Else
                            str1 = str1 & " and " & var2
                        End If
                    End If

                Next
                If ctA = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = str1
                End If

            Case 71
                'strFind = "[NUMBEROFSAMPLES]"
                If frmH.chkManualSampleNumber.CheckState = CheckState.Checked Then
                    varReplace = frmH.txtSRecTotalReport.Text
                    strNA4 = "StudyDoc Sample Receipt Table"
                ElseIf frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    varReplace = frmH.txtSRecTotalReportWatson.Text
                    strNA4 = "Watson Sample Receipt Table"
                Else
                    varReplace = frmH.txtSRecTotal.Text
                    strNA4 = "StudyDoc Sample Receipt Table"
                End If
                strNA1 = charSectionName
                strNA2 = "Sample Count"
                strNA3 = "Sample Receipt"

            Case 72
                'strFind = "[REPRCHROMWATSONRUNNUMBER]"
                dtbl = tblAppFigs
                strF = "ID_TBLSTUDIES = " & idTS & " AND CHARTYPE = 'RC'"
                Erase rows1
                rows1 = dtbl.Select(strF)
                int1 = rows1.Length
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = rows1(0).Item("NUMWATSONRUNNUMBER")
                    If Len(NZ(var1, "")) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var1
                    End If
                End If

                strNA1 = charSectionName
                strNA2 = "Representative Watson Run Number"
                strNA3 = "Appendices"
                strNA4 = "Representative Chromatogram Entry"

            Case 73
                'strFind = "[STORAGETEMP]" 'tblSampleReceipt.charStorageTemp
                If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    dv = frmH.dgvSampleReceiptWatson.DataSource
                    str1 = "Storage Temperature"
                    strNA4 = "Watson Table Storage Temperature"
                Else
                    dv = frmH.dgvSampleReceipt.DataSource
                    str1 = "charStorageTemp"
                    strNA4 = "StudyDoc Table Storage Temperature"
                End If

                If dv.Count < 1 Then
                    varReplace = "[NA]" ' & ChrW(176) & "C"
                Else
                    'check to see if var1 contains a degreeC ending
                    var8 = dv(0).Item(str1)
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    ElseIf IsNumeric(var8) Then
                        varReplace = var8 & " " & ChrW(176) & "C"
                    Else
                        'int1 = InStr(var8, ChrW(176), CompareMethod.Text)
                        'If int1 > 0 Then 'OK
                        '    'var2 = Mid(var8.ToString, Len(var8) - 1, 1)
                        '    'var3 = AscW(var2)
                        '    'If AscW(var2) = 176 Then 'OK
                        '    varReplace = var8
                        'Else
                        '    varReplace = var8 & " " & ChrW(176) & "C"
                        'End If
                        varReplace = var8
                    End If
                End If

                'check to see if left character is a minus sign
                'if so, replace with non-breaking hyphen
                'DON'T NEED TO DO THIS BECAUSE StudyDoc DOES IT LATER
                'If StrComp(varReplace, "[NA]", CompareMethod.Text) = 0 Then
                'Else
                '    'var1 = dv(0).Item("charStorageTemp")
                '    var8 = varReplace 'dv(0).Item(str1)
                '    If IsDBNull(var8) Then
                '        varReplace = "[NA]"
                '        'check to see if var1 contains a minus sign
                '    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                '        varReplace = "[NA]"
                '    Else
                '        str2 = Mid(var8, 1, 1)
                '        If StrComp(str2, "-", CompareMethod.Text) = 0 Then 'OK
                '            str3 = Mid(var8.ToString, 2, Len(var8) - 1)
                '            'var3 = AscW(var2)
                '            'If AscW(var2) = 176 Then 'OK
                '            varReplace = NBH & str3
                '        End If
                '    End If
                'End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Storage Temperature"
                strNA3 = "Sample Receipt"

            Case 74
                'strFind = "[SHIPMENTCOUNT]"
                If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    dv = frmH.dgvSampleReceiptWatson.DataSource
                    strNA4 = "Watson Table Shipment Count"
                Else
                    dv = frmH.dgvSampleReceipt.DataSource
                    strNA4 = "StudyDoc Table Shipment Count"
                End If
                If dv.Count < 1 Then
                    var8 = "[NA]"
                Else
                    var8 = VerboseNumber(dv.Count, False)
                End If
                varReplace = var8
                strNA1 = charSectionName
                strNA2 = "Shipment Count"
                strNA3 = "Sample Receipt"
            Case 75
                'strFind = "[SHIPMENTCONDITION]"
                If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    dv = frmH.dgvSampleReceiptWatson.DataSource
                    str1 = "Sample Condition"
                    strNA4 = "Watson Table Sample Condition"
                Else
                    dv = frmH.dgvSampleReceipt.DataSource
                    str1 = "charCondition"
                    strNA4 = "StudyDoc Table Sample Condition"
                End If
                If dv.Count < 1 Then
                    var8 = "[NA]"
                Else
                    var8 = NZ(dv(0).Item(str1), "[NA]")
                    var8 = UnCapit(var8, True)
                End If
                varReplace = var8
                strNA1 = charSectionName
                strNA2 = "Sample Condition"
                strNA3 = "Sample Receipt"


            Case 77
                'strFind = "[CALIBRATIONLEVELSSECTION]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("Calibration Levels", dv)
                var8 = NZ(dv(int1).Item(1).ToString, "")
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                Else
                    var2 = arrAnalytes(1, 1)
                    str1 = var8 & "-point calibration curves for " & var2
                    If ctAnalytes > 1 Then
                        For Count2 = 2 To ctAnalytes
                            var1 = NZ(dv(int1).Item(Count2).ToString, "[NA]")
                            var2 = arrAnalytes(1, Count2)
                            If Count2 = ctAnalytes And ctAnalytes > 2 Then
                                str1 = str1 & ", and " & var1 & "-point calibration curves for" & ChrW(160) & var2
                            ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                                str1 = str1 & ", " & var1 & "-point calibration curves for" & ChrW(160) & var2
                            Else
                                str1 = str1 & " and " & var1 & "-point calibration curves for" & ChrW(160) & var2
                            End If
                        Next
                    End If
                    varReplace = str1
                End If
                strNA1 = charSectionName
                strNA2 = "Calibration Levels"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 78
                'strFind = "[REGRESSIONSECTION]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("Weighting", dv)
                var1 = LCase(NZ(dv(int1).Item(1).ToString, ""))
                str2 = var1 & "-weighted "
                int2 = FindRowDV("Regression", dv)
                'var1 = NZ(UnCapit(dv(int1).Item(1).ToString, True), "")
                var1 = LCase(NZ(dv(int2).Item(1).ToString, ""))
                If StrComp(var1, "Linear", CompareMethod.Text) = 0 Then
                    var1 = "least-squares " & var1
                End If
                str1 = str2 & var1 & " regression"
                If ctAnalytes > 1 Then
                    For Count2 = 2 To ctAnalytes
                        var1 = LCase(NZ(dv(int1).Item(Count2).ToString, "[NA]"))
                        str2 = var1 & "-weighted "
                        'var1 = NZ(UnCapit(dv(int1).Item(Count2).ToString, False), "[NA]")
                        var1 = LCase(NZ(dv(int2).Item(1).ToString, "[NA]"))
                        If StrComp(var1, "linear", CompareMethod.Text) = 0 Then
                            var1 = "least-squares " & var1
                        End If
                        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", and " & str2 & var1 & " regression"
                        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", " & str2 & var1 & " regression"
                        Else
                            str1 = str1 & " and " & str2 & var1 & " regression"
                        End If
                    Next
                End If
                varReplace = str1
                strNA1 = charSectionName
                strNA2 = "Regression and Weighting"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 79

                'strFind = "[R2SECTION]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("Minimum r^2", dv)
                var1 = NZ(dv(int1).Item(1).ToString, "[NA]")
                var2 = arrAnalytes(1, 1)
                str1 = var1 & " for " & var2
                If ctAnalytes > 1 Then
                    For Count2 = 2 To ctAnalytes
                        var1 = NZ(dv(int1).Item(Count2).ToString, "")
                        var2 = arrAnalytes(1, Count2)
                        'str1 = str1 & " and " & ChrW(8805) & var1 & " for " & var2
                        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", and " & ChrW(8805) & var1 & " for " & var2
                        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", " & ChrW(8805) & var1 & " for " & var2
                        Else
                            str1 = str1 & " and " & ChrW(8805) & var1 & " for " & var2
                        End If
                    Next
                End If
                varReplace = str1
                strNA1 = charSectionName
                strNA2 = "R-squared"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 80
                'strFind = "[REGRESSIONTABLENUMBERSECTION]"
                '*****
                'check to see if regression info is placed in calibr table
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 3"
                rowsProps = dtblProps.Select(strF)
                If rowsProps.Length = 0 Then
                    varReplace = "NA"
                Else
                    var1 = rowsProps(0).Item("BOOLSTATSREGR")
                    If var1 = -1 Then
                        idTbl = 3 'call CALSTDTABLENUMBERSECTION
                        varReplace = GetTableNumber(idTbl, False)
                    Else
                        idTbl = 2
                        varReplace = GetTableNumber(idTbl, False)
                    End If
                    'If ctTableN = 0 Then
                    '    If ctAnalytes = 1 Then
                    '        str1 = "Table [NA]"
                    '    Else
                    '        str1 = "Tables [NA] to [NA]"
                    '    End If
                    'Else
                    '    'idTbl = 2
                    '    str3 = "TableID = " & idTbl
                    '    drows = tblN.Select(str3)
                    '    int1 = drows.Length
                    '    If drows.Length = 0 Then
                    '        var1 = "[NA]"
                    '        var2 = "[NA]"
                    '    Else
                    '        var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                    '        var2 = NZ(drows(int1 - 1).Item("TableNumber"), "[NA]")
                    '    End If

                    '    If ctAnalytes = 1 Then
                    '        str1 = "Table " & var1
                    '    Else
                    '        str1 = "Tables " & var1 & " to " & var2
                    '    End If
                    'End If
                    'varReplace = str1
                End If


                '****
                strNA1 = charSectionName
                strNA2 = "Regression Table Number(s)"
                strNA3 = "Report Table Configuration"
                strNA4 = "Report Table"

            Case 81
                'strFind = "[CALSTDTABLENUMBERSECTION]"
                '****
                idTbl = 3
                varReplace = GetTableNumber(idTbl, False)

                'If ctTableN = 0 Then
                '    str1 = "Table [NA] for " & arrAnalytes(1, 1)
                '    For Count2 = 2 To ctAnalytes
                '        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", and Table [NA] for " & arrAnalytes(1, Count2)
                '        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", Table [NA] for " & arrAnalytes(1, Count2)
                '        Else
                '            str1 = str1 & " and Table [NA] for " & arrAnalytes(1, Count2)
                '        End If
                '    Next
                'Else
                '    idTbl = 3
                '    str3 = "TableID = " & idTbl
                '    drows = tblN.Select(str3)
                '    var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                '    str1 = "Table " & var1 & " for " & arrAnalytes(1, 1)
                '    For Count2 = 2 To ctAnalytes
                '        str3 = "TableID = " & idTbl & " AND AnalyteName = '" & arrAnalytes(1, Count2) & "'"
                '        drows = tblN.Select(str3)
                '        var1 = NZ(drows(0).Item("TableNumber"), "[NA]")

                '        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", and Table " & var1 & " for " & arrAnalytes(1, Count2)
                '        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", Table " & var1 & " for " & arrAnalytes(1, Count2)
                '        Else
                '            str1 = str1 & " and Table " & var1 & " for " & arrAnalytes(1, Count2)
                '        End If
                '    Next
                'End If
                'varReplace = str1
                '****
                strNA1 = charSectionName
                strNA2 = "Calibration Standard Table Number(s)"
                strNA3 = "Report Table Configuration"
                strNA4 = "Report Table"

            Case 82
                'strFind = "[ACCURACYSECTION]"
                Dim intMeanAccuracyMinRow, intMeanAccuracyMaxRow, intLLOQUnitsRow As Short
                dv = frmH.dgvWatsonAnalRef.DataSource
                intMeanAccuracyMinRow = FindRowDV("Analyte Mean Accuracy Min", dv)
                intMeanAccuracyMaxRow = FindRowDV("Analyte Mean Accuracy Max", dv)
                intLLOQUnitsRow = FindRowDV("LLOQ Units", dv)

                var1 = NZ(dv(intMeanAccuracyMinRow).Item(1).ToString, "[NA]")
                str1 = var1 & ChrW(160) & "to" & ChrW(160)
                var1 = NZ(dv(intMeanAccuracyMaxRow).Item(1).ToString, "[NA]")
                str1 = str1 & var1

                'add units
                var1 = NZ(dv(intLLOQUnitsRow).Item(1).ToString, "[NA]")
                str1 = str1 & ChrW(160) & var1

                var2 = arrAnalytes(1, 1)
                str1 = str1 & " for " & var2

                If ctAnalytes > 1 Then
                    For Count2 = 2 To ctAnalytes
                        var1 = NZ(dv(intMeanAccuracyMinRow).Item(Count2).ToString, "[NA]") 'min
                        str2 = var1 & ChrW(160) & "to" & ChrW(160)
                        var1 = NZ(dv(intMeanAccuracyMaxRow).Item(Count2).ToString, "[NA]") 'max
                        str2 = str2 & var1
                        'add units
                        var1 = NZ(dv(intLLOQUnitsRow).Item(1).ToString, "[NA]")
                        str2 = str2 & ChrW(160) & var1
                        var2 = arrAnalytes(1, Count2)
                        str2 = str2 & " for " & var2
                        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", and " & str2
                        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", " & str2
                        Else
                            str1 = str1 & " and " & str2
                        End If
                    Next
                End If
                varReplace = str1
                strNA1 = charSectionName
                strNA2 = "Analyte Mean Accuracy"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 83
                'strFind = "[PRECISIONSECTION]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("Analyte Precision Min", dv)
                var1 = NZ(dv(int1).Item(1).ToString, "[NA]")
                str1 = var1 & ChrW(160) & "to" & ChrW(160)
                int2 = FindRowDV("Analyte Precision Max", dv)
                var1 = NZ(dv(int2).Item(1).ToString, "")
                str1 = str1 & var1
                var2 = arrAnalytes(1, 1)
                str1 = str1 & " for " & var2
                If ctAnalytes > 1 Then
                    For Count2 = 2 To ctAnalytes
                        var1 = NZ(dv(int1).Item(Count2).ToString, "[NA]") 'min
                        str2 = var1 & ChrW(160) & "to" & ChrW(160)
                        'int1 = FindRowDV("Analyte Precision Max", dv)
                        var1 = NZ(dv(int2).Item(Count2).ToString, "[NA]") 'max
                        str2 = str2 & var1
                        var2 = arrAnalytes(1, Count2)
                        str2 = str2 & " for " & var2
                        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", and " & str2
                        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", " & str2
                        Else
                            str1 = str1 & " and " & str2
                        End If
                    Next
                End If
                varReplace = str1
                strNA1 = charSectionName
                strNA2 = "Analyte Mean Accuracy"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 84
                'strFind = "[QCTABLENUMBERSECTION]"

                idTbl = 4
                varReplace = GetTableNumber(idTbl, False)

                ''****
                'If ctTableN = 0 Then
                '    str1 = "Table [NA] for " & arrAnalytes(1, 1)
                '    For Count2 = 2 To ctAnalytes
                '        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", and Table [NA] for " & arrAnalytes(1, Count2)
                '        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", Table [NA] for " & arrAnalytes(1, Count2)
                '        Else
                '            str1 = str1 & " and Table [NA] for " & arrAnalytes(1, Count2)
                '        End If
                '    Next
                'Else
                '    str3 = "TableName = 'Summary of Interpolated QC Std Conc' AND AnalyteName = '" & arrAnalytes(1, 1) & "'"
                '    idTbl = 4
                '    str3 = "TableID = " & idTbl
                '    drows = tblN.Select(str3)
                '    If drows.Length = 0 Then
                '        var1 = "[NA]"
                '    Else
                '        var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                '    End If
                '    str1 = "Table " & var1 & " for " & arrAnalytes(1, 1)
                '    For Count2 = 2 To ctAnalytes
                '        str3 = "TableID = " & idTbl & " AND AnalyteName = '" & arrAnalytes(1, Count2) & "'"
                '        drows = tblN.Select(str3)
                '        If drows.Length = 0 Then
                '            var1 = "[NA]"
                '        Else
                '            var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                '        End If
                '        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", and Table " & var1 & " for " & arrAnalytes(1, Count2)
                '        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", Table " & var1 & " for " & arrAnalytes(1, Count2)
                '        Else
                '            str1 = str1 & " and Table " & var1 & " for " & arrAnalytes(1, Count2)
                '        End If
                '    Next
                'End If
                'varReplace = str1
                ''****
                strNA1 = charSectionName
                strNA2 = "QC Table Number"
                strNA3 = "Report Table Configuration"
                strNA4 = "Report Table"

            Case 85
                'strFind = "[QCACCURACYSECTION]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("QC Mean Accuracy Min", dv)
                var1 = NZ(dv(int1).Item(1).ToString, "[NA]")
                str1 = var1 & ChrW(160) & "to" & ChrW(160)
                int2 = FindRowDV("QC Mean Accuracy Max", dv)
                var1 = NZ(dv(int2).Item(1).ToString, "[NA]")
                str1 = str1 & var1
                var2 = arrAnalytes(1, 1)
                str1 = str1 & " for " & var2
                If ctAnalytes > 1 Then
                    For Count2 = 2 To ctAnalytes
                        var1 = NZ(dv(int1).Item(Count2).ToString, "[NA]") 'min
                        str2 = var1 & ChrW(160) & "to" & ChrW(160)
                        'int1 = FindRowDV("QC Mean Accuracy Max", dv)
                        var1 = NZ(dv(int2).Item(Count2).ToString, "[NA]") 'max
                        str2 = str2 & var1
                        var2 = arrAnalytes(1, Count2)
                        str2 = str2 & " for " & var2
                        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", and " & str2
                        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", " & str2
                        Else
                            str1 = str1 & " and " & str2
                        End If
                    Next
                End If
                varReplace = str1
                strNA1 = charSectionName
                strNA2 = "QC Mean Accuracy"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 86
                'strFind = "[QCPRECISIONSECTION]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("QC Precision Min", dv)
                var1 = NZ(dv(int1).Item(1).ToString, "[NA]")
                str1 = var1 & ChrW(160) & "to" & ChrW(160)
                int2 = FindRowDV("QC Precision Max", dv)
                var1 = NZ(dv(int2).Item(1).ToString, "[NA]")
                str1 = str1 & var1
                var2 = arrAnalytes(1, 1)
                str1 = str1 & " for " & var2
                If ctAnalytes > 1 Then
                    For Count2 = 2 To ctAnalytes
                        var1 = NZ(dv(int1).Item(Count2).ToString, "[NA]") 'min
                        str2 = var1 & ChrW(160) & "to" & ChrW(160)
                        'int1 = FindRowDV("QC Precision Max", dv)
                        var1 = NZ(dv(int2).Item(Count2).ToString, "[NA]") 'max
                        str2 = str2 & var1
                        var2 = arrAnalytes(1, Count2)
                        str2 = str2 & " for " & var2
                        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", and " & str2
                        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", " & str2
                        Else
                            str1 = str1 & " and " & str2
                        End If
                    Next
                End If
                varReplace = str1
                strNA1 = charSectionName
                strNA2 = "QC Mean Accuracy"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 87
                'strFind = "[SAMPLETABLENUMBERSECTION]"

                idTbl = 5
                varReplace = GetTableNumber(idTbl, False)

                ''****
                'If ctTableN = 0 Then
                '    str1 = "Table [NA] for " & arrAnalytes(1, 1)
                '    For Count2 = 2 To ctAnalytes
                '        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", and Table [NA] for " & arrAnalytes(1, Count2)
                '        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", Table [NA] for " & arrAnalytes(1, Count2)
                '        Else
                '            str1 = str1 & " and Table [NA] for " & arrAnalytes(1, Count2)
                '        End If
                '    Next
                'Else
                '    str3 = "TableName = 'Summary of Samples' AND AnalyteName = '" & arrAnalytes(1, 1) & "'"
                '    idTbl = 5
                '    str3 = "TableID = " & idTbl
                '    drows = tblN.Select(str3)
                '    If drows.Length = 0 Then
                '        var1 = "[NA]"
                '    Else
                '        var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                '    End If

                '    str1 = "Table " & var1 & " for " & arrAnalytes(1, 1)
                '    For Count2 = 2 To ctAnalytes
                '        str3 = "TableID = " & idTbl & " AND AnalyteName = '" & arrAnalytes(1, Count2) & "'"
                '        drows = tblN.Select(str3)
                '        If drows.Length = 0 Then
                '            var1 = "[NA]"
                '        Else
                '            var1 = NZ(drows(0).Item("TableNumber"), "[NA]")
                '        End If

                '        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", and Table " & var1 & " for " & arrAnalytes(1, Count2)
                '        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                '            str1 = str1 & ", Table " & var1 & " for " & arrAnalytes(1, Count2)
                '        Else
                '            str1 = str1 & " and Table " & var1 & " for " & arrAnalytes(1, Count2)
                '        End If
                '    Next
                'End If
                'varReplace = str1
                ''****

                strNA1 = charSectionName
                strNA2 = "Samples Table Numbers"
                strNA3 = "Report Table Configuration"
                strNA4 = "Report Table"

            Case 88
                'strFind = "[STUDYSAMPLEBQLSECTION1]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("LLOQ", dv)
                str1 = NZ(dv(int1).Item(1), "[NA]")
                num1 = NZ(dv(int1).Item(1), 0)
                If IsDBNull(num1) Then
                    var8 = "[NA]"
                    int2 = FindRowDV("LLOQ Units", dv)
                    strUnits = NZ(dv(int2).Item(1), "[NA]")

                    int2 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    str2 = NZ(frmH.dgvStudyConfig(1, int2).Value, "")

                    If Len(str2) = 0 Or StrComp(str2, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        strUnits = str2
                    End If

                    str1 = var8 & ChrW(160) & strUnits & " for " & arrAnalytes(1, 1)
                Else
                    num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                    int2 = FindRowDV("LLOQ Units", dv)
                    strUnits = NZ(dv(int2).Item(1), "[NA]")

                    int2 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    str2 = NZ(frmH.dgvStudyConfig(1, int2).Value, "")

                    If Len(str2) = 0 Or StrComp(str2, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        strUnits = str2
                    End If

                    str1 = num1 & ChrW(160) & strUnits & " for " & arrAnalytes(1, 1)
                End If
                If ctAnalytes > 1 Then
                    For Count2 = 2 To ctAnalytes
                        str2 = NZ(dv(int1).Item(Count2), "[NA]")
                        num1 = dv(int1).Item(Count2)
                        If IsDBNull(num1) Then
                            var8 = "[NA]"
                            str2 = var8 & ChrW(160) & strUnits & " for " & arrAnalytes(1, Count2)
                        Else
                            num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                            str2 = num1 & ChrW(160) & strUnits & " for " & arrAnalytes(1, Count2)
                        End If
                        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", and " & str2
                        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", " & str2
                        Else
                            str1 = str1 & " and " & str2
                        End If
                    Next
                End If
                varReplace = str1
                strNA1 = charSectionName
                strNA2 = "LLOQ"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 89
                'strFind = "[STUDYSAMPLEBQLSECTION2]"
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("LLOQ", dv)
                num1 = dv(int1).Item(1)
                If IsDBNull(num1) Then
                    var8 = "[NA]"
                    str1 = """BQL<(" & var8 & ")"" for " & arrAnalytes(1, 1)
                Else
                    num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                    str1 = """BQL<(" & num1 & ")"" for " & arrAnalytes(1, 1)
                End If
                If ctAnalytes > 1 Then
                    For Count2 = 2 To ctAnalytes
                        num1 = dv(int1).Item(Count2)
                        If IsDBNull(num1) Then
                            var8 = "[NA]"
                            str2 = """BQL<(" & var8 & ")"" for " & arrAnalytes(1, Count2)
                        Else
                            num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                            str2 = """BQL<(" & num1 & ")"" for " & arrAnalytes(1, Count2)
                        End If
                        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", and " & str2
                        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", " & str2
                        Else
                            str1 = str1 & " and " & str2
                        End If
                    Next
                End If
                varReplace = str1
                strNA1 = charSectionName
                strNA2 = "LLOQ"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 90
                'strFind = "[PLURALNUMBER]"
                If ctAnalytes = 1 Then
                    varReplace = "number"
                Else
                    varReplace = "numbers"
                End If
            Case 91
                'strFind = "[PLURALSUMMARY]"
                If ctAnalytes = 1 Then
                    varReplace = "summary"
                Else
                    varReplace = "summaries"
                End If

            Case 92
                'strFind = "[NUMBEROFSAMPLESVERBOSESECTION]"
                If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    var9 = frmH.txtSRecTotalReportWatson.Text
                    strNA4 = "Watson Sample Receipt Table"
                Else
                    var9 = frmH.txtSRecTotalReport.Text
                    strNA4 = "StudyDoc Sample Receipt Table"

                End If
                If var9 = 1 Then
                    var8 = VerboseNumber(var9, True) & " sample"
                ElseIf var9 = 0 Then
                    var8 = "Samples"
                Else
                    var8 = VerboseNumber(var9, True) & " samples"
                End If
                varReplace = var8
                strNA1 = charSectionName
                strNA2 = "Sample Count"
                strNA3 = "Sample Receipt"

            Case 93
                'strFind = "[SHIPMENTCOUNTSECTION]"
                If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    dv = frmH.dgvSampleReceiptWatson.DataSource
                    strNA4 = "Watson Sample Receipt Table"
                Else
                    dv = frmH.dgvSampleReceipt.DataSource
                    strNA4 = "StudyDoc Sample Receipt Table"
                End If

                var9 = dv.Count
                If var9 < 1 Then
                    var8 = "[NA] shipment"
                ElseIf var9 = 1 Then
                    var8 = VerboseNumber(var9, False) & " shipment"
                Else
                    var8 = VerboseNumber(var9, False) & " shipments"
                End If
                varReplace = var8
                strNA1 = charSectionName
                strNA2 = "Number of Shipments"
                strNA3 = "Sample Receipt"

            Case 94
                'strFind = "[SHIPMENTRECEIPTDATESECTION]"
                If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    dv = frmH.dgvSampleReceiptWatson.DataSource
                    str1 = "Date Received"
                    strNA4 = "Watson Sample Receipt Table"
                Else
                    dv = frmH.dgvSampleReceipt.DataSource
                    str1 = "dtShipmentReceived"
                    strNA4 = "StudyDoc Sample Receipt Table"
                End If
                If dv.Count < 1 Then
                    varReplace = "on [NA]"
                Else
                    If IsDBNull(dv(0).Item(str1)) Then
                        str2 = "[NA]"
                    Else
                        dt1 = dv(0).Item(str1)
                        str2 = Format(dt1, LTextDateFormat)
                    End If
                    If IsDBNull(dv(dv.Count - 1).Item(str1)) Then
                        str3 = "[NA]"
                    Else
                        dt2 = dv(dv.Count - 1).Item(str1)
                        str3 = Format(dt2, LTextDateFormat)
                    End If
                    If dv.Count = 1 Then
                        varReplace = "on " & str2
                    ElseIf dv.Count = 2 Then
                        varReplace = str2 & " and " & str3
                    Else
                        'dt2 = dv(dv.Count - 1).Item("Date Received")
                        'varReplace = "between " & str2 & " and " & str3
                        varReplace = str2 & " to " & str3
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "Date Received"
                strNA3 = "Sample Receipt"

            Case 95
                'strFind = "[LABMETHODNAME]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Lab Method Title", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Lab Method Name"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 96
                'strFind = "[LABMETHODNUMBER]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Lab Method Number", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBHReal, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Lab Method Number"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 97
                'strFind = "[SHIPMENTSOURCE]"
                dv = frmH.dgvSampleReceipt.DataSource
                If dv.Count = 0 Then
                    var8 = ""
                Else
                    int1 = 0
                    var8 = NZ(dv.Item(int1).Item("CHARSOURCE"), "")
                End If
                If Len(var8) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Sample Receipt Information"
                strNA3 = "Sample Receipt Tab"
                strNA4 = "Sampel Receipt Data Table"

            Case 98
                'strFind = "[QCCONCENTRATIONCOUNT]" 'numQCConcCount
                var8 = NZ(numQCLevels, 0)
                'If var8 = 0 Then
                '    varReplace = "[NA]"
                'Else
                '    varReplace = VerboseNumber(var8, True)
                'End If
                varReplace = VerboseNumber(var8, True)
                strNA1 = charSectionName
                strNA2 = "# of QC Levels"
                strNA3 = "Analytical Reference Standard"
                strNA4 = "Watson Table"

            Case 99
                'strFind = "[REPLICATEDILUTIONQC]" 'numRepDilnQC

                '20190228 LEE: Deprecated

                '20190228 LEE:
                'do new logic
                tbl1 = tblAssignedSamples
                intT = 12
                strS = "ID_TBLCONFIGREPORTTABLES ASC"

                ''20190222 LEE: Need to find Dilution table in tblTableProperties
                strF = GetBOOLSTATSNRFilter(12) '12 = Diln table

                strS = "ID_TBLCONFIGREPORTTABLES ASC"

                Try
                    Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                    tbl2 = dv1.ToTable("a", True, "NOMCONC", "RUNID", "CHARANALYTE")
                    int1 = tbl2.Rows.Count
                    If int1 = 0 Then
                        varReplace = "[NA]"
                    Else
                        'loop to find largest value
                        Dim intMax As Short = 0
                        For Count2 = 0 To tbl2.Rows.Count - 1
                            var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                            var2 = tbl2.Rows.Item(0).Item("RUNID")
                            var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                            strF = "(ID_TBLCONFIGREPORTTABLES = " & intT & " OR ID_TBLCONFIGREPORTTABLES = 31) AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                            rows1 = tbl1.Select(strF)
                            int2 = rows1.Length
                            If int2 > intMax Then
                                intMax = int2
                            End If
                        Next

                        If intMax = 0 Then
                            varReplace = "[NA]"
                        Else
                            varReplace = intMax ' VerboseNumber(intmax, False)
                        End If

                    End If
                    strNA1 = charSectionName
                    strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                    strNA3 = "Report Table Configuration"
                    strNA4 = "Summary of Interpolated Dilution QC Concentrations"
                Catch ex As Exception
                    var1 = var1
                End Try

                'end new logic

                'var8 = NZ(numRepDilnQC, 0)
                ''If var8 = 0 Then
                ''    varReplace = "[NA]"
                ''Else
                ''    varReplace = VerboseNumber(var8, True)
                ''End If
                'varReplace = VerboseNumber(var8, True)
                'strNA1 = charSectionName
                'strNA2 = "# of Diln QC Replicates"
                'strNA3 = "Analytical Reference Standard"
                'strNA4 = "Watson Table"

            Case 100 'replace any 'deg C' with symbol
                'strFind = "degC"
                var8 = ChrW(186) & "C"
                varReplace = var8

            Case 101 'replace any 'deg C' with symbol
                'strFind = "deg C"
                var8 = ChrW(186) & "C"
                varReplace = var8


            Case 102 'replace any '+/-' with symbol
                'strFind = "+/-"
                var8 = ChrW(177)
                varReplace = var8

            Case 103 'replace any '<=' with symbol
                'strFind = "<="
                var8 = ChrW(8804)
                varReplace = var8

            Case 104 'replace any '>=' with symbol
                'strFind = ">="
                var8 = ChrW(8805)
                varReplace = var8

            Case 105
                'strFind = "[NUMFTCYCLES]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties
                intT = 19
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = rows1(0).Item("INTNUMBEROFCYCLES")
                'look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    'look for a sequence of characters that is numeric
                    var3 = ""
                    var4 = ""
                    bool1 = False 'Start
                    bool2 = False 'End
                    For Count1 = 1 To int1
                        var2 = Mid(var1, Count1, 1)
                        If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                            var2 = "a"
                        End If
                        If IsNumeric(var2) Then
                            var3 = var3 & var2
                            If IsNumeric(var3) Then
                                var4 = var3
                                bool1 = True
                            Else
                            End If
                        Else
                            If bool1 Then
                                bool2 = True
                            End If
                        End If
                        If bool1 And bool2 Then
                            Exit For
                        End If
                    Next
                End If
                If bool1 = False Then
                    varReplace = "[NA]"
                Else
                    varReplace = VerboseNumber(var4, False)
                End If
                strNA1 = charSectionName
                strNA2 = "Summary of Freeze/Thaw [#Cycles] Stability in Matrix"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Freeze/Thaw [#Cycles] Stability in Matrix PERIODTEMP Table"

            Case 106
                'strFind = "[FXSRUNID]"
                tbl1 = tblAssignedSamples
                intT = 21
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "RUNID")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = tbl2.Rows.Item(0).Item("RUNID")
                End If
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations Table"

            Case 107
                'strFind = "[FXSTEMP]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties
                intT = 21
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = rows1(0).Item("CHARPERIODTEMP")
                'look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var1

                    ''look for a sequence of characters that is numeric
                    'var3 = ""
                    'var4 = ""
                    'bool1 = False 'Start
                    'bool2 = False 'End
                    'For Count1 = 1 To int1
                    '    var2 = Mid(var1, Count1, 1)
                    '    If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                    '        var2 = "a"
                    '    End If
                    '    If IsNumeric(var2) Then
                    '        var3 = var3 & var2
                    '        If IsNumeric(var3) Then
                    '            var4 = var3
                    '            bool1 = True
                    '        Else
                    '        End If
                    '    Else
                    '        If bool1 Then
                    '            bool2 = True
                    '        End If
                    '    End If
                    '    If bool1 And bool2 Then
                    '        Exit For
                    '    End If
                    'Next
                    'If bool1 = False Then
                    '    'do nothing
                    'Else
                    '    'remove number from var1
                    '    var2 = Replace(var1, CStr(var4), "", 1)
                    '    var1 = Trim(var2)
                    'End If

                    ''now remove hours/days/weeks/months from var1
                    'For Count1 = 1 To 8
                    '    Select Case Count1
                    '        Case 1
                    '            str1 = "hours"
                    '        Case 2
                    '            str1 = "days"
                    '        Case 3
                    '            str1 = "weeks"
                    '        Case 4
                    '            str1 = "months"
                    '        Case 5
                    '            str1 = "hour"
                    '        Case 6
                    '            str1 = "day"
                    '        Case 7
                    '            str1 = "week"
                    '        Case 8
                    '            str1 = "month"
                    '    End Select
                    '    var2 = Replace(var1, str1, "", 1, -1, CompareMethod.Text)
                    '    var1 = Trim(var2)
                    'Next

                    'If Len(NZ(var1, "")) = 0 Then
                    '    varReplace = "[NA]"
                    'Else
                    '    varReplace = UnCapit(var1, False)
                    'End If

                End If
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations PERIOD TEMP"

            Case 108
                'strFind = "[FXSTIME]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties
                intT = 21
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = NZ(rows1(0).Item("CHARTIMEPERIOD"), "")
                var2 = NZ(rows1(0).Item("CHARTIMEFRAME"), "")

                'look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = LCase(CStr(var1)) & " " & var2

                    ''look for a sequence of characters that is numeric
                    'var3 = ""
                    'var4 = ""
                    'bool1 = False 'Start
                    'bool2 = False 'End
                    'For Count1 = 1 To int1
                    '    var2 = Mid(var1, Count1, 1)
                    '    If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                    '        var2 = "a"
                    '    End If
                    '    If IsNumeric(var2) Then
                    '        var3 = var3 & var2
                    '        If IsNumeric(var3) Then
                    '            var4 = var3
                    '            bool1 = True
                    '        Else
                    '        End If
                    '    Else
                    '        If bool1 Then
                    '            bool2 = True
                    '        End If
                    '    End If
                    '    If bool1 And bool2 Then
                    '        Exit For
                    '    End If
                    'Next
                    'If bool1 = False Then
                    '    var2 = "[NA]" & ChrW(160)
                    'Else
                    '    var2 = VerboseNumber(var4, False) & ChrW(160)
                    'End If

                    ''now look for days or hours
                    'If InStr(1, var1, "hours", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "hours"
                    'ElseIf InStr(1, var1, "days", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "days"
                    'ElseIf InStr(1, var1, "weeks", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "weeks"
                    'ElseIf InStr(1, var1, "months", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "months"

                    'ElseIf InStr(1, var1, "hour", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "hour"
                    'ElseIf InStr(1, var1, "day", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "day"
                    'ElseIf InStr(1, var1, "week", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "week"
                    'ElseIf InStr(1, var1, "month", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "month"

                    'Else
                    '    var2 = var2 & "[NA]"
                    'End If

                    varReplace = var2
                End If

                strNA1 = charSectionName
                strNA2 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations PERIOD TEMP"

            Case 110
                'strFind = "[SSSTEMP]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties
                intT = 22
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = rows1(0).Item("CHARPERIODTEMP")
                'look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var1

                    ''look for a sequence of characters that is numeric
                    'var3 = ""
                    'var4 = ""
                    'bool1 = False 'Start
                    'bool2 = False 'End
                    'For Count1 = 1 To int1
                    '    var2 = Mid(var1, Count1, 1)
                    '    If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                    '        var2 = "a"
                    '    End If
                    '    If IsNumeric(var2) Then
                    '        var3 = var3 & var2
                    '        If IsNumeric(var3) Then
                    '            var4 = var3
                    '            bool1 = True
                    '        Else
                    '        End If
                    '    Else
                    '        If bool1 Then
                    '            bool2 = True
                    '        End If
                    '    End If
                    '    If bool1 And bool2 Then
                    '        Exit For
                    '    End If
                    'Next
                    'If bool1 = False Then
                    '    'do nothing
                    'Else
                    '    'remove number from var1
                    '    var2 = Replace(var1, CStr(var4), "", 1)
                    '    var1 = Trim(var2)
                    'End If

                    ''now remove hours/days/weeks/months from var1
                    'For Count1 = 1 To 8
                    '    Select Case Count1
                    '        Case 1
                    '            str1 = "hours"
                    '        Case 2
                    '            str1 = "days"
                    '        Case 3
                    '            str1 = "weeks"
                    '        Case 4
                    '            str1 = "months"
                    '        Case 5
                    '            str1 = "hour"
                    '        Case 6
                    '            str1 = "day"
                    '        Case 7
                    '            str1 = "week"
                    '        Case 8
                    '            str1 = "month"
                    '    End Select
                    '    var2 = Replace(var1, str1, "", 1, -1, CompareMethod.Text)
                    '    var1 = Trim(var2)
                    'Next

                    'If Len(NZ(var1, "")) = 0 Then
                    '    varReplace = "[NA]"
                    'Else
                    '    varReplace = UnCapit(var1, False)
                    'End If

                End If
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Stock Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Stock Solution Stability Assessment PERIOD TEMP"

            Case 111
                'strFind = "[SSSTIME]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties

                intT = 22
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = NZ(rows1(0).Item("CHARTIMEPERIOD"), "")
                var2 = NZ(rows1(0).Item("CHARTIMEFRAME"), "")
                'look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = LCase(CStr(var1)) & " " & var2

                    ''look for a sequence of characters that is numeric
                    'var3 = ""
                    'var4 = ""
                    'bool1 = False 'Start
                    'bool2 = False 'End
                    'For Count1 = 1 To int1
                    '    var2 = Mid(var1, Count1, 1)
                    '    If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                    '        var2 = "a"
                    '    End If
                    '    If IsNumeric(var2) Then
                    '        var3 = var3 & var2
                    '        If IsNumeric(var3) Then
                    '            var4 = var3
                    '            bool1 = True
                    '        Else
                    '        End If
                    '    Else
                    '        If bool1 Then
                    '            bool2 = True
                    '        End If
                    '    End If
                    '    If bool1 And bool2 Then
                    '        Exit For
                    '    End If
                    'Next
                    'If bool1 = False Then
                    '    var2 = "[NA]" & ChrW(160)
                    'Else
                    '    var2 = VerboseNumber(var4, False) & ChrW(160)
                    'End If

                    ''now look for days or hours
                    'If InStr(1, var1, "hours", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "hours"
                    'ElseIf InStr(1, var1, "days", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "days"
                    'ElseIf InStr(1, var1, "weeks", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "weeks"
                    'ElseIf InStr(1, var1, "months", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "months"

                    'ElseIf InStr(1, var1, "hour", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "hour"
                    'ElseIf InStr(1, var1, "day", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "day"
                    'ElseIf InStr(1, var1, "week", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "week"
                    'ElseIf InStr(1, var1, "month", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "month"
                    'Else
                    '    var2 = var2 & "[NA]"
                    'End If

                    'varReplace = var2
                End If

                strNA1 = charSectionName
                strNA2 = "[Period Temp] Stock Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Stock Solution Stability Assessment PERIOD TEMP"

            Case 112
                'strFind = "[NUMSSSREPLICATES]"
                tbl1 = tblAssignedSamples
                intT = 22
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "CHARHELPER2", "RUNID", "CHARANALYTE")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = NZ(tbl2.Rows.Item(0).Item("CHARHELPER2"), "0")
                    var2 = NZ(tbl2.Rows.Item(0).Item("RUNID"), "0")
                    var3 = NZ(tbl2.Rows.Item(0).Item("CHARANALYTE"), "Anal")

                    strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND CHARHELPER2 = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                    rows1 = tbl1.Select(strF)
                    int2 = rows1.Length
                    If int2 = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = VerboseNumber(int2, False)
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Stock Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Stock Solution Stability Assessment"

            Case 113
                'strFind = "[SpSSTEMP]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties

                intT = 23
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = rows1(0).Item("CHARPERIODTEMP")
                'look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var1

                    ''look for a sequence of characters that is numeric
                    'var3 = ""
                    'var4 = ""
                    'bool1 = False 'Start
                    'bool2 = False 'End
                    'For Count1 = 1 To int1
                    '    var2 = Mid(var1, Count1, 1)
                    '    If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                    '        var2 = "a"
                    '    End If
                    '    If IsNumeric(var2) Then
                    '        'if previous character is "-", then grab it
                    '        var2 = Mid(var1, Count1 - 1, 1)
                    '        If StrComp(var2, "-", CompareMethod.Text) = 0 Then
                    '            var3 = Mid(var1, Count1 - 1, Len(var1))
                    '        Else
                    '            var3 = Mid(var1, Count1, Len(var1))
                    '        End If
                    '        Exit For
                    '    End If

                    '    'If IsNumeric(var2) Then
                    '    '    var3 = var3 & var2
                    '    '    If IsNumeric(var3) Then
                    '    '        var4 = var3
                    '    '        bool1 = True
                    '    '    Else
                    '    '    End If
                    '    'Else
                    '    '    If bool1 Then
                    '    '        bool2 = True
                    '    '    End If
                    '    'End If
                    '    'If bool1 And bool2 Then
                    '    '    Exit For
                    '    'End If
                    'Next
                    ''If bool1 = False Then
                    ''    'do nothing
                    ''Else
                    ''    'remove number from var1
                    ''    var2 = Replace(var1, CStr(var4), "", 1)
                    ''    var1 = Trim(var2)
                    ''End If

                    ' ''now remove hours/days/weeks/months from var1
                    ''For Count1 = 1 To 8
                    ''    Select Case Count1
                    ''        Case 1
                    ''            str1 = "hours"
                    ''        Case 2
                    ''            str1 = "days"
                    ''        Case 3
                    ''            str1 = "weeks"
                    ''        Case 4
                    ''            str1 = "months"
                    ''        Case 5
                    ''            str1 = "hour"
                    ''        Case 6
                    ''            str1 = "day"
                    ''        Case 7
                    ''            str1 = "week"
                    ''        Case 8
                    ''            str1 = "month"
                    ''    End Select
                    ''    var2 = Replace(var1, str1, "", 1, -1, CompareMethod.Text)
                    ''    var1 = Trim(var2)
                    ''Next

                    'If Len(NZ(var3, "")) = 0 Then
                    '    varReplace = "[NA]"
                    'Else
                    '    varReplace = var3 ' UnCapit(var1, False)
                    'End If

                End If
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Spiking Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Spiking Solution Stability Assessment PERIOD TEMP"

            Case 114
                'strFind = "[SpSSTIME]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties

                intT = 23
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = NZ(rows1(0).Item("CHARTIMEPERIOD"), "")
                var2 = NZ(rows1(0).Item("CHARTIMEFRAME"), "")
                'look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = LCase(CStr(var1)) & " " & var2

                    ''look for a sequence of characters that is numeric
                    'var3 = ""
                    'var4 = ""
                    'bool1 = False 'Start
                    'bool2 = False 'End
                    'For Count1 = 1 To int1
                    '    var2 = Mid(var1, Count1, 1)
                    '    If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                    '        var2 = "a"
                    '    End If
                    '    If IsNumeric(var2) Then
                    '        var3 = var3 & var2
                    '        If IsNumeric(var3) Then
                    '            var4 = var3
                    '            bool1 = True
                    '        Else
                    '        End If
                    '    Else
                    '        If bool1 Then
                    '            bool2 = True
                    '        End If
                    '    End If
                    '    If bool1 And bool2 Then
                    '        Exit For
                    '    End If
                    'Next
                    'If bool1 = False Then
                    '    var2 = "[NA]" & ChrW(160)
                    'Else
                    '    var2 = VerboseNumber(var4, False) & ChrW(160)
                    'End If

                    ''now look for days or hours
                    'If InStr(1, var1, "hours", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "hours"
                    'ElseIf InStr(1, var1, "days", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "days"
                    'ElseIf InStr(1, var1, "weeks", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "weeks"
                    'ElseIf InStr(1, var1, "months", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "months"

                    'ElseIf InStr(1, var1, "hour", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "hour"
                    'ElseIf InStr(1, var1, "day", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "day"
                    'ElseIf InStr(1, var1, "week", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "week"
                    'ElseIf InStr(1, var1, "month", CompareMethod.Text) > 0 Then
                    '    var2 = var2 & "month"
                    'Else
                    '    var2 = var2 & "[NA]"
                    'End If

                    'varReplace = var2
                End If

                strNA1 = charSectionName
                strNA2 = "[Period Temp] Spiking Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Spiking Solution Stability Assessment PERIOD TEMP"

            Case 115
                'strFind = "[NUMSpSSREPLICATES]"
                tbl1 = tblAssignedSamples
                intT = 23
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC", "CHARANALYTE", "CHARHELPER1")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = NZ(tbl2.Rows.Item(0).Item("NOMCONC"), 0)
                    var3 = NZ(tbl2.Rows.Item(0).Item("CHARANALYTE"), "Anal")
                    var4 = NZ(tbl2.Rows.Item(0).Item("CHARHELPER1"), "Anal")
                    strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND NOMCONC = " & var1 & "AND CHARANALYTE = '" & CleanText(CStr(var3)) & "' AND CHARHELPER1 = '" & var4 & "'"
                    rows1 = tbl1.Select(strF)
                    int2 = rows1.Length
                    If int2 = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = VerboseNumber(int2, False)
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Spiking Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Spiking Solution Stability Assessment Table"

            Case 116
                'strFind = "[LTSTEMP]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties

                intT = 29
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = rows1(0).Item("CHARPERIODTEMP")
                'look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var1

                    ''look for a sequence of characters that is numeric
                    'var3 = ""
                    'var4 = ""
                    'bool1 = False 'Start
                    'bool2 = False 'End
                    'For Count1 = 1 To int1
                    '    var2 = Mid(var1, Count1, 1)
                    '    If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                    '        var2 = "a"
                    '    End If
                    '    If IsNumeric(var2) Then
                    '        'if previous character is "-", then grab it
                    '        var2 = Mid(var1, Count1 - 1, 1)
                    '        If StrComp(var2, "-", CompareMethod.Text) = 0 Then
                    '            var3 = Mid(var1, Count1 - 1, Len(var1))
                    '        Else
                    '            var3 = Mid(var1, Count1, Len(var1))
                    '        End If
                    '        Exit For
                    '    End If

                    '    '    If IsNumeric(var2) Then
                    '    '        var3 = var3 & var2
                    '    '        If IsNumeric(var3) Then
                    '    '            var4 = var3
                    '    '            bool1 = True
                    '    '        Else
                    '    '        End If
                    '    '    Else
                    '    '        If bool1 Then
                    '    '            bool2 = True
                    '    '        End If
                    '    '    End If
                    '    '    If bool1 And bool2 Then
                    '    '        Exit For
                    '    '    End If
                    'Next
                    ''If bool1 = False Then
                    ''    'do nothing
                    ''Else


                    ''    'remove number from var1
                    ''    var2 = Replace(var1, CStr(var4), "", 1)
                    ''    var1 = Trim(var2)
                    ''End If

                    ' ''now remove hours/days/weeks/months from var1
                    ''For Count1 = 1 To 8
                    ''    Select Case Count1
                    ''        Case 1
                    ''            str1 = "hours"
                    ''        Case 2
                    ''            str1 = "days"
                    ''        Case 3
                    ''            str1 = "weeks"
                    ''        Case 4
                    ''            str1 = "months"
                    ''        Case 5
                    ''            str1 = "hour"
                    ''        Case 6
                    ''            str1 = "day"
                    ''        Case 7
                    ''            str1 = "week"
                    ''        Case 8
                    ''            str1 = "month"
                    ''    End Select
                    ''    var2 = Replace(var1, str1, "", 1, -1, CompareMethod.Text)
                    ''    var1 = Trim(var2)
                    ''Next

                    ''If Len(NZ(var1, "")) = 0 Then
                    ''    varReplace = "[NA]"
                    ''Else
                    ''    varReplace = UnCapit(var1, False)
                    ''End If

                    'If Len(NZ(var3, "")) = 0 Then
                    '    varReplace = "[NA]"
                    'Else
                    '    varReplace = var3 ' UnCapit(var1, False)
                    'End If

                End If
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Long-Term QC Std Storage Stability"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Long-Term QC Std Storage Stability PERIOD TEMP"

            Case 117

                'strFind = "[LTSTIME]"
                'tbl1 = tblReportTable
                tbl1 = tblTableProperties
                intT = 29
                strF = "ID_TBLSTUDIES = " & idTS & " AND ID_TBLCONFIGREPORTTABLES = " & intT
                rows1 = tbl1.Select(strF)
                var1 = NZ(rows1(0).Item("CHARTIMEPERIOD"), "")
                var2 = NZ(rows1(0).Item("CHARTIMEFRAME"), "")
                ''look for a number in this text
                int1 = Len(NZ(var1, ""))
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = LCase(CStr(var1)) & " " & LCase(CStr(var2))

                    '    'look for a sequence of characters that is numeric
                    '    var3 = ""
                    '    var4 = ""
                    '    bool1 = False 'Start
                    '    bool2 = False 'End
                    '    For Count1 = 1 To int1
                    '        var2 = Mid(var1, Count1, 1)
                    '        If StrComp(var2, " ", CompareMethod.Text) = 0 Then
                    '            var2 = "a"
                    '        End If
                    '        If IsNumeric(var2) Then
                    '            var3 = var3 & var2
                    '            If IsNumeric(var3) Then
                    '                var4 = var3
                    '                bool1 = True
                    '            Else
                    '            End If
                    '        Else
                    '            If bool1 Then
                    '                bool2 = True
                    '            End If
                    '        End If
                    '        If bool1 And bool2 Then
                    '            Exit For
                    '        End If
                    '    Next
                    '    If bool1 = False Then
                    '        var2 = "[NA]" & Chr(160)
                    '    Else
                    '        var2 = VerboseNumber(var4, False) & ChrW(160)
                    '    End If

                    '    'now look for days or hours
                    '    If InStr(1, var1, "hours", CompareMethod.Text) > 0 Then
                    '        var2 = var2 & "days"
                    '    ElseIf InStr(1, var1, "days", CompareMethod.Text) > 0 Then
                    '        var2 = var2 & "days"
                    '    ElseIf InStr(1, var1, "weeks", CompareMethod.Text) > 0 Then
                    '        var2 = var2 & "weeks"
                    '    ElseIf InStr(1, var1, "months", CompareMethod.Text) > 0 Then
                    '        var2 = var2 & "months"

                    '    ElseIf InStr(1, var1, "hour", CompareMethod.Text) > 0 Then
                    '        var2 = var2 & "hour"
                    '    ElseIf InStr(1, var1, "day", CompareMethod.Text) > 0 Then
                    '        var2 = var2 & "day"
                    '    ElseIf InStr(1, var1, "week", CompareMethod.Text) > 0 Then
                    '        var2 = var2 & "week"
                    '    ElseIf InStr(1, var1, "month", CompareMethod.Text) > 0 Then
                    '        var2 = var2 & "month"
                    '    Else
                    '        var2 = var2 & "[NA]"
                    '    End If

                    '    varReplace = var2
                End If

                strNA1 = charSectionName
                strNA2 = "[Period Temp] Long-Term QC Std Storage Stability"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Long-Term QC Std Storage Stability PERIOD TEMP"

            Case 118
                'strFind = "[NUMLTSREPLICATES]"
                tbl1 = tblAssignedSamples
                intT = 29
                strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND ID_TBLSTUDIES = " & idTS
                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC", "CHARANALYTE", "CHARHELPER1", "CHARHELPER2")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                    var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                    var4 = tbl2.Rows.Item(0).Item("CHARHELPER1")
                    var2 = tbl2.Rows.Item(0).Item("CHARHELPER2")
                    strF = "ID_TBLCONFIGREPORTTABLES = " & intT & " AND NOMCONC = " & var1 & "AND CHARANALYTE = '" & CleanText(CStr(var3)) & "' AND CHARHELPER1 = '" & var4 & "' AND CHARHELPER2 = '" & var2 & "'"
                    rows1 = tbl1.Select(strF)
                    int2 = rows1.Length
                    If int2 = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = VerboseNumber(int2, False)
                    End If
                End If
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Long-Term QC Std Storage Stability"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Long-Term QC Std Storage Stability"

            Case 119

                'strFind = "[ANALYTERESP]"
                'determine if there is more than one analyte

                str1 = ""
                Dim ct1 As Short
                ct1 = 0

                '20161004 LEE: start using Analyte Groups tblAnalyteGroups
                'first inventory included analytes
                Dim ctA As Short = tblAnalyteGroups.Rows.Count

                Dim arrA(2, ctA)

                For Count2 = 0 To ctA - 1
                    var1 = tblAnalyteGroups.Rows(Count2).Item("ANALYTEDESCRIPTION")
                    var2 = tblAnalyteGroups.Rows(Count2).Item("ANALYTEDESCRIPTION_C")
                    'ensure analyte is included in at least one table
                    If UseAnalyte(CStr(var2)) Then
                        ct1 = ct1 + 1
                        arrA(1, ct1) = var1
                        arrA(2, ct1) = var2
                    End If
                Next

                ctA = ct1

                For Count2 = 1 To ctA
                    var1 = arrA(1, Count2)
                    '20170929 LEE: should be using arrA(2, n) for multiple matrix
                    var1 = arrA(2, Count2)
                    'replace hyphens with nbh
                    var2 = Replace(var1, "-", NBHReal, 1, -1, CompareMethod.Text)
                    'replace spaces with nbs
                    var2 = Replace(var2, " ", ChrW(160), 1, -1, CompareMethod.Text)

                    If Count2 = 1 Then
                        str1 = var2 'arrAnalytes(1, Count2)
                    Else
                        If Count2 = ctA And ctA > 2 Then
                            str1 = str1 & ", and " & var2
                        ElseIf Count2 <> ctA And ctA > 2 Then
                            str1 = str1 & ", " & var2
                        Else
                            str1 = str1 & " and " & var2
                        End If
                    End If

                Next

                If ctA > 1 Then
                    str1 = str1 & ", respectively"
                End If

                If ctA = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = str1
                End If

            Case 120

                'strFind = "[INTERNALSTANDARDRESP]" 'add 'respectively' at the end
                str1 = ""
                Dim ct1 As Short
                ct1 = 0

                'first get unique Int Stds
                'first inventory included analytes
                Dim ctA As Short = tblAnalyteGroups.Rows.Count

                Dim arrA(2, ctA)

                For Count2 = 0 To ctA - 1
                    var1 = tblAnalyteGroups.Rows(Count2).Item("ANALYTEDESCRIPTION")
                    var2 = tblAnalyteGroups.Rows(Count2).Item("ANALYTEDESCRIPTION_C")
                    var3 = tblAnalyteGroups.Rows(Count2).Item("INTSTD")
                    'ensure analyte is included in at least one table
                    If UseAnalyte(CStr(var2)) Then
                        If ct1 = 0 Then
                            ct1 = ct1 + 1
                            arrA(1, ct1) = var3
                        Else
                            'must be unique
                            boolHit = False
                            For Count3 = 1 To ct1
                                var4 = arrA(1, Count3)
                                If StrComp(var4, var3, CompareMethod.Text) = 0 Then
                                    boolHit = True
                                End If
                            Next
                            If boolHit Then
                            Else
                                ct1 = ct1 + 1
                                arrA(1, ct1) = var3
                            End If
                        End If

                    End If
                Next

                ctA = ct1

                For Count2 = 1 To ct1
                    var1 = arrA(1, Count2)
                    'replace hyphens with nbh
                    var2 = Replace(var1, "-", NBHReal, 1, -1, CompareMethod.Text)
                    'replace spaces with nbs
                    var2 = Replace(var2, " ", ChrW(160), 1, -1, CompareMethod.Text)

                    If Count2 = 1 Then
                        str1 = var2 'arrAnalytes(1, Count2)
                    Else
                        If Count2 = ct1 And ct1 > 2 Then
                            str1 = str1 & ", and " & var2
                        ElseIf Count2 <> ct1 And ct1 > 2 Then
                            str1 = str1 & ", " & var2
                        Else
                            str1 = str1 & " and " & var2
                        End If
                    End If

                Next

                If ctA > 1 Then
                    str1 = str1 & ", respectively"
                End If

                If ctA = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = str1
                End If

            Case 121
                'strFind = "[TABLENUMBER_4]"
                idTbl = 4
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated QC Std Conc"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated QC Std Conc Table"
            Case 122
                'strFind = "[TABLENUMBER_12]"
                idTbl = 12
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated Dilution QC Concentrations Table"

            Case 123
                'strFind = "[TABLENUMBERALL_13]"
                idTbl = 13
                varReplace = GetTableNumberAll(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Combined Recovery"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Combined Recovery Table"

            Case 124
                'strFind = "[TABLENUMBERALL_14]"
                idTbl = 14
                varReplace = GetTableNumberAll(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of True Recovery"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of True Recovery Table"

            Case 125
                'strFind = "[TABLENUMBERALL_15]"
                idTbl = 15
                varReplace = GetTableNumberAll(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Suppression/Enhancement"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Suppression/Enhancement Table"

            Case 126
                'strFind = "[TABLENUMBER_2]"
                idTbl = 2
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Regression Constants"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Regression Constants Table"

            Case 127
                'strFind = "[ANALYTE_11]"
                idTbl = 11
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision Table"

            Case 128
                'strFind = "[ANALYTE_3]"
                idTbl = 3
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Back-Calculated Calibration Std Conc"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Back-Calculated Calibration Std Conc Table"

            Case 129
                'strFind = "[ANALYTE_4]"
                idTbl = 4
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated QC Std Conc"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated QC Std Conc Table"

            Case 130
                'strFind = "[ANALYTE_12]"
                idTbl = 12
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated Dilution QC Concentrations"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated Dilution QC Concentrations Table"

            Case 131
                'strFind = "[ANALYTE_13]"
                idTbl = 13
                varReplace = GetAnalyteAll(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Combined Recovery"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Combined Recovery Table"

            Case 132
                'strFind = "[ANALYTE_14]"
                idTbl = 14
                varReplace = GetAnalyteAll(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of True Recovery"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of True Recovery Table"

            Case 133
                'strFind = "[ANALYTE_15]"
                idTbl = 15
                varReplace = GetAnalyteAll(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Suppression/Enhancement"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Suppression/Enhancement Table"

            Case 134
                'strFind = "[ANALYTE_2]" 'deprecated
                idTbl = 2
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Regression Constants"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Regression Constants Table"

            Case 135
                'strfind = [MAXRUNSIZE]
                int1 = FindRowDV("Maximum Run Size", frmH.dgvMethodValData.DataSource)
                dv = frmH.dgvMethodValData.DataSource
                var8 = dv(int1).Item(1)
                'var8 = frmH.dgvMethodValData(1, int1).Value
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Maximum Run Size"
                strNA3 = "Method Validation Tab"
                strNA4 = "Maximum Run Size Entry"

            Case 136
                'strFind = "[TABLENUMBER_17]"
                idTbl = 17
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments Table"

            Case 137
                'strFind = "[ANALYTE_17]"
                idTbl = 17
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments Table"

            Case 138
                'strFind = "[TABLENUMBER_18]"
                idTbl = 18
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of [Period Temp] Stability in Matrix"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of [Period Temp] Stability in Matrix Table"

            Case 139
                'strFind = "[ANALYTE_18]"
                idTbl = 18
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of [Period Temp] Stability in Matrix"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of [Period Temp] Stability in Matrix Table"

            Case 140
                'strFind = "[TABLENUMBER_19]"
                idTbl = 19
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Freeze/Thaw [#Cycles] Stability in Matrix"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Freeze/Thaw [#Cycles] Stability in Matrix"

            Case 141
                'strFind = "[ANALYTE_19]"
                idTbl = 19
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "Summary of Freeze/Thaw [#Cycles] Stability in Matrix"
                strNA3 = "Report Table Configuration"
                strNA4 = "Summary of Freeze/Thaw [#Cycles] Stability in Matrix Table"

            Case 142
                'strFind = "[TABLENUMBER_21]"
                idTbl = 21
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations Table"

            Case 143
                'strFind = "[ANALYTE_21]"
                idTbl = 21
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations Table"

            Case 144
                'strFind = "[TABLENUMBER_22]"
                idTbl = 22
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Stock Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Stock Solution Stability Assessment Table"

            Case 145
                'strFind = "[ANALYTE_22]"
                idTbl = 22
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Stock Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Stock Solution Stability Assessment	Table"

            Case 146
                'strFind = "[TABLENUMBER_23]"
                idTbl = 23
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Spiking Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Spiking Solution Stability Assessment Table"

            Case 147
                'strFind = "[ANALYTE_23]"
                idTbl = 23
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Spiking Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Spiking Solution Stability Assessment Table"

            Case 148
                'strFind = "[TABLENUMBER_29]"
                idTbl = 29
                varReplace = GetTableNumber(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Spiking Solution Stability Assessment"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Spiking Solution Stability Assessment Table"

            Case 149
                'strFind = "[ANALYTE_29]"
                idTbl = 29
                varReplace = GetAnalyte(idTbl, True)
                strNA1 = charSectionName
                strNA2 = "[Period Temp] Long-Term QC Std Storage Stability"
                strNA3 = "Report Table Configuration"
                strNA4 = "[Period Temp] Long-Term QC Std Storage Stability Table"

            Case 150
                'strFind = "[PAGENUMBER]"
                'Call InsertPageNumber(wd)
                varReplace = ""

            Case 151
                'strFind = "[TOTALPAGES]"
                'Call InsertTotalPages(wd)
                varReplace = ""

            Case 152
                'strFind = "[SUBMITTEDTOFULL]"
                dtbl = tblCorporateAddresses
                var8 = frmH.cbxSubmittedTo.Text
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    str1 = var8
                    str2 = "charNickName = '" & str1 & "'"
                    tblNick = tblCorporateNickNames
                    rowsNick = tblNick.Select(str2)
                    var1 = rowsNick(0).Item("id_tblCorporateNickNames")
                    str2 = "id_tblCorporateNickNames = " & var1 ' & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = -1"
                    drows = dtbl.Select(str2, "id_tbladdresslabels ASC")
                    int1 = drows.Length
                    If int1 < 1 Then
                        varReplace = "[NA]"
                    Else
                        str1 = GetAddress(var8)
                        varReplace = str1
                    End If
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Submitted To"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Submitted To Textbox"

            Case 153
                'strFind = "[SUBMITTEDBYFULL]"
                dtbl = tblCorporateAddresses
                var8 = frmH.cbxSubmittedBy.Text
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    str1 = var8
                    str2 = "charNickName = '" & str1 & "'"
                    tblNick = tblCorporateNickNames
                    rowsNick = tblNick.Select(str2)
                    var1 = rowsNick(0).Item("id_tblCorporateNickNames")
                    str2 = "id_tblCorporateNickNames = " & var1 ' & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = " & -1
                    drows = dtbl.Select(str2, "id_tbladdresslabels ASC")
                    int1 = drows.Length
                    If int1 < 1 Then
                        varReplace = "[NA]"
                    Else
                        str1 = GetAddress(var8)
                        varReplace = str1
                    End If

                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Submitted By"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Submitted By Textbox"

            Case 154
                'strFind = "[INSUPPORTOFFULL]"
                'strFind = "[SUBMITTEDBYFULL]"
                dtbl = tblCorporateAddresses
                var8 = frmH.cbxInSupportOf.Text
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    str1 = var8
                    str2 = "charNickName = '" & str1 & "'"
                    tblNick = tblCorporateNickNames
                    rowsNick = tblNick.Select(str2)
                    var1 = rowsNick(0).Item("id_tblCorporateNickNames")
                    str2 = "id_tblCorporateNickNames = " & var1 ' & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = " & -1
                    drows = dtbl.Select(str2, "id_tbladdresslabels ASC")
                    int1 = drows.Length
                    If int1 < 1 Then
                        varReplace = "[NA]"
                    Else
                        str1 = GetAddress(var8)
                        varReplace = str1
                    End If

                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "In Support Of"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "In Support Of Textbox"

            Case 155
                'strFind = "[COVERPAGESIGNATURE]" NOT USED!!
                varReplace = ""


            Case 160 'SECTIONBREAKNEXTPAGE
                varReplace = ""
            Case 161 'INSERTPAGEBREAK
                varReplace = ""
            Case 162
                'strFind = "[FIRSTPAGESPECIAL]"
                varReplace = ""
            Case 163
                'strFind = "[CONTRIBUTINGPERSONNELTABLE]"
                varReplace = ""
            Case 164
                'strFind = "[REASSAYTABLENUMBER]"

                idTbl = 6
                varReplace = GetTableNumber(idTbl, False)

                ''****
                strNA1 = charSectionName
                strNA2 = "Summary of Reassayed Samples Table Number"
                strNA3 = "Report Table Configuration"
                strNA4 = "Report Table"

            Case 165
                'strFind = "[REPEATTABLENUMBER]"

                idTbl = 7
                varReplace = GetTableNumber(idTbl, False)

                strNA1 = charSectionName
                strNA2 = "Summary of Repeat Samples Table Number"
                strNA3 = "Report Table Configuration"
                strNA4 = "Report Table"

            Case 166 'strFind = "[INSERTQATABLE]"
                varReplace = ""

            Case 167 '
                'strFind = "[REPORTSTATUS]"
                dgv = frmH.dgvReports
                dv = dgv.DataSource
                'intRow = dg.CurrentRowIndex
                str1 = "id_tblStudies = " & intIDtblStudies
                dv.RowFilter = str1
                var8 = dv.Item(0).Item("dtReportFinalIssueDate")
                'format var8
                If IsDBNull(var8) Then
                    varReplace = "DRAFT"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "DRAFT"
                Else
                    varReplace = "FINAL"
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Report Issue Date"
                strNA3 = "Choose Study & Report"
                strNA4 = "Configured Reports table"

                'ABSTRACTANALYTEINFO
            Case 168 'strFind = "[ABSTRACTANALYTEINFO]"
                varReplace = "[ABSTRACTANALYTEINFO]"

            Case 169
                'strfind = "[ANALYTE_IUPAC]
                bool1 = True
                tbl1 = tblCompanyAnalRefTable
                int1 = FindRow("IUPAC Name", tbl1, "Item")
                dgv = frmH.dgvCompanyAnalRef
                dv = dgv.DataSource
                var1 = arrAnalytes(1, 1)
                var2 = NZ(tbl1.Rows(int1).Item(var1), "")
                If Len(var2) = 0 Then
                    var2 = "[NA]"
                    bool1 = False
                End If

                var8 = ""
                For Count2 = 1 To ctAnalytes
                    var1 = arrAnalytes(1, Count2)
                    var2 = NZ(tbl1.Rows(int1).Item(var1), "")
                    If Len(var2) = 0 Then
                        var2 = "[NA]"
                        bool1 = False
                    End If
                    var8 = var8 & var1 & ":  " & var2 & ChrW(10)
                Next

                If bool1 Then
                    varReplace = var8
                Else
                    varReplace = "[NA]"
                End If

                strNA1 = charSectionName
                strNA2 = "IUPAC Name"
                strNA3 = "Analytical Reference Standards"
                strNA4 = "Company Analytical Standard Table"

            Case 170
                'strFind = "[CALIBRSTANDARDLIST]"

                'get units
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("ULOQ Units", dv)
                var4 = dv.Item(int1).Item(1)

                int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                Else
                    var4 = str1
                End If

                If IsDBNull(var4) Then
                    var4 = "[NA]"
                ElseIf Len(var4) = 0 Or StrComp(NZ(var4, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    var4 = "[NA]"
                Else
                    var4 = var4
                End If

                intGroup = arrAnalytes(15, 1)

                var1 = arrAnalytes(1, 1)
                var2 = ReturnCalibrStds(CStr(var1), intGroup, False)

                str1 = var2 & " " & var4 & " for " & var1

                If ctAnalytes > 1 Then
                    For Count2 = 2 To ctAnalytes
                        var1 = arrAnalytes(1, Count2)
                        intGroup = arrAnalytes(15, Count2)
                        var2 = ReturnCalibrStds(CStr(var1), intGroup, False)
                        str2 = var2 & " " & var4 & " for " & var1
                        If Count2 = ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", and " & str2
                        ElseIf Count2 <> ctAnalytes And ctAnalytes > 2 Then
                            str1 = str1 & ", " & str2
                        Else
                            str1 = str1 & " and " & str2
                        End If
                    Next
                End If
                varReplace = str1

            Case 171
                'strFind = "[QCSTANDARDLIST]"

                'get units
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("ULOQ Units", dv)
                var4 = dv.Item(int1).Item(1)

                int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                Else
                    var4 = str1
                End If

                If IsDBNull(var4) Then
                    var4 = "[NA]"
                ElseIf Len(var4) = 0 Or StrComp(NZ(var4, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    var4 = "[NA]"
                Else
                    var4 = var4
                End If

                'var1 = arrAnalytes(1, 1)
                'var2 = ReturnQCStds(CStr(var1))

                ''20181129 LEE:
                ''reformat
                'str1 = var1 & ":" & ChrW(9) & var2 & fNBSP() & var4

                '20181129 LEE:
                'reformat
                For Count2 = 1 To ctAnalytes
                    var1 = arrAnalytes(1, Count2)
                    var2 = ReturnQCStds(CStr(var1), False)
                    intGroup = arrAnalytes(15, Count2)
                    Dim rows4() As DataRow = TBLSTUDYDOCANALYTES.Select("INTGROUP = " & intGroup & " AND ID_TBLSTUDIES = " & id_tblStudies)
                    If rows4.Length = 0 Then
                    Else
                        var1 = rows4(0).Item("CHARUSERANALYTE")
                    End If
                    If Count2 = 1 Then
                        str1 = var1 & ":" & ChrW(9) & var2 & fNBSP() & var4
                    Else
                        str1 = str1 & ChrW(11) & var1 & ":" & ChrW(9) & var2 & fNBSP() & var4
                    End If
                Next

                varReplace = str1

            Case 172
                'strFind = [TIMEZONE]
                varReplace = LTimeZone

            Case 173
                '[STUDYSAMPLECONCENTRATIONTABLE]
                varReplace = ""

            Case 174
                'strFind = "[SAMPLERECEIPTTABLE1]"
                varReplace = ""

            Case -9
                'strFind = "[METHODSUMMARYSTATEMENT]"
                varReplace = ""

            Case 175
                'strFind = "[CALSTDTABLE1]"
                varReplace = ""
            Case 176
                'strFind = "[CALSTDTABLE2]"
                varReplace = ""
            Case 177
                'strFind = "[QCSECTION]"
                varReplace = DoQCSECTION(wd)
            Case 178
                'strFind = "[DILUTIONQCSECTION]"
                varReplace = DoDILUTIONQCSECTION(wd)
            Case 179
                'strFind = "[QCTABLE1]"
                varReplace = ""
            Case 180
                'strFind = "[ABSTRACTANALYTEINFO]"
                varReplace = "[ABSTRACTANALYTEINFO1]"
            Case 181
                'strFind = "[ABSTRACTANALYTEINFO]"
                varReplace = "[ABSTRACTANALYTEINFO2]"
            Case 182
                'strFind = "[DATENOW]"
                varReplace = Format(Now, Replace(LTextDateFormat, "Y", "y", 1, -1, CompareMethod.Binary))
                'varReplace = Replace(LTextDateFormat, "Y", "y", 1, -1, CompareMethod.Binary)

            Case 183
                'strFind = "[DATETIMENOW]"
                str1 = Format(Now, Replace(LTextDateFormat, "Y", "y", 1, -1, CompareMethod.Binary))
                str2 = Format(Now, "hh:mm:ss tt")
                varReplace = str1 & " " & str2
            Case 184
                'strFind = "[OUTLIERMETHOD]"
                dtbl = tblData
                str2 = "id_tblStudies = " & intIDtblStudies
                drows = dtbl.Select(str2)
                int1 = drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var8 = NZ(drows(0).Item("CHAROUTLIERMETHOD"), "")
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                End If

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Outlier Method"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data table"

            Case 185
                'strFind = "[TABLESECTION]"
                varReplace = ""
            Case 186
                'strFind = "[APPENDIXSECTION]"
                varReplace = ""
            Case 187
                'strFind = "[FIGURESECTION]"
                varReplace = ""
            Case 188
                varReplace = RegressionEquation()
            Case 189
                'strFind = "[DATELASTSAMPLESRECEIVED]"
                If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    dv = frmH.dgvSampleReceiptWatson.DataSource
                    strNA4 = "StudyDoc Sample Receipt table"
                Else
                    dv = frmH.dgvSampleReceipt.DataSource
                    strNA4 = "Watson Sample Receipt table"
                End If
                int1 = dv.Count 'drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                        var8 = NZ(dv(int1 - 1).Item("Date Received"), "")
                    Else
                        var8 = NZ(dv(int1 - 1).Item("dtShipmentReceived"), "")
                    End If
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    Else
                        If IsDate(var8) Then
                            varReplace = Format(CDate(var8), LTextDateFormat)
                        Else
                            varReplace = "[NA]"
                        End If
                    End If
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Date Samples Received - Last Shipment"
                strNA3 = "Sample Receipt"

            Case 190 'strFind = "[BIASORDIFFQC]"
                varReplace = FindBiasDiff("QC")

            Case 191 'strFind = "[BIASORDIFFCALIBR]"
                varReplace = FindBiasDiff("Calibr")

            Case 192 'strFind = "[ANTICOAGULANTMETHOD]"

                '20190213 LEE: Deprecated. Now use CHARANTICOAGULANT

                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource 'intI is analyte column in dgvMethodValData
                int1 = FindRowDVByCol("Anticoagulant/Preservative", dv, "Item")
                'var8 = dv.Item(int1).Item(strAnal)
                var8 = Trim(NZ(dv.Item(int1).Item(1), "[NA]"))

                If IsDBNull(var8) Then
                    varReplace = "[NA]"

                ElseIf Len(var8) = 0 Then
                    varReplace = "[NA]"
                Else
                    If Len(var8) = 0 Then
                    Else
                        int1 = Len(var8)
                        int2 = 0
                        For Count2 = 1 To Len(var8) 'int1
                            var1 = CStr(Mid(var8, Count2, 1))
                            var2 = Asc(var1)
                            If var2 > 64 And var2 < 91 Then
                                int2 = int2 + 1
                            Else
                                Exit For
                            End If
                        Next
                        If int2 > 2 Then 'probably is an acronym, leave capitalized
                            var8 = var8
                        Else
                            'var8 = UnCapit(Trim(var8), True)
                            var8 = LCase(var8)
                        End If
                    End If

                    varReplace = var8 'UnCapit(Trim(var8), True)
                End If

                strNA1 = charSectionName
                strNA2 = "Anticoagulant/Preservative"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 193 'strFind = "[FREEZETHAWSTORAGECOND]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Freeze/Thaw Storage Conditions", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Freeze/Thaw Storage Conditions"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 194 '[STUDYSTARTDATE]
                dtbl = tblData
                str2 = "id_tblStudies = " & intIDtblStudies
                drows = dtbl.Select(str2)
                int1 = drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var8 = drows(0).Item("DTSTUDYSTARTDATE")
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    Else
                        If IsDate(var8) Then
                            varReplace = Format(CDate(var8), LTextDateFormat)
                        Else
                            varReplace = "[NA]"
                        End If
                    End If
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Study Start Date"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data table"

            Case 195 '[STUDYENDDATE]
                dtbl = tblData
                str2 = "id_tblStudies = " & intIDtblStudies
                drows = dtbl.Select(str2)
                int1 = drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var8 = drows(0).Item("DTSTUDYENDDATE")
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    Else
                        If IsDate(var8) Then
                            varReplace = Format(CDate(var8), LTextDateFormat)
                        Else
                            varReplace = "[NA]"
                        End If
                    End If
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Study End Date"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data table"

            Case 196 '[ATTACHMENTSECTION_01]'NOTE: Attachments depricated
                varReplace = ""

            Case 197
                'strFind = "[WATSONSTUDYTITLE]"
                Try
                    intRow = frmH.dgvwStudy.CurrentRow.Index
                    varReplace = frmH.dgvwStudy("StudyTitle", intRow).Value
                Catch ex As Exception
                    varReplace = "NA"
                End Try

                strNA1 = charSectionName
                strNA2 = "Watson Study Title"
                strNA3 = "Choose Study & Report"
                strNA4 = "Study table"

            Case 198 'strFind = "[WATSONREPCHROMSECTION]"
                varReplace = WatsonRepChrom()

            Case 199
                'strFind = "[REGRESSIONTABLENUMBER]"
                'ctArrReportNA = ctArrReportNA + 1

                varReplace = "[NA]"
                idTbl = 2
                varReplace = GetTableNumber(idTbl, False)

                strNA1 = charSectionName
                strNA2 = "Regression Table Number(s)"
                strNA3 = "Report Table Configuration"
                strNA4 = "Report Table"

            Case 200
                'strFind = "[METHODPROCESSSTABILITY]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Process Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Process Stability"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 201

                '20190213 LEE: Deprecated

                'strFind = "[METHODREFRIGERATEDSTABILITYINMATRIX]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                'int1 = FindRowDV("Refrigerated Stability in Matrix", dv) 'deprecated now Reinjection Stability 20181110
                int1 = FindRowDV("Reinjection Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Reinjection Stability" ' "Refrigerated Stability in Matrix"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 202
                'strFind = "[METHODLONGTERMSTORAGESTABILITYINMATRIX]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Long-term Storage Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Long-term Stability in Matrix"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 203
                'strFind = "[METHODDILUTIONINTEGRITY]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Dilution Integrity", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Dilution Integrity"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 204
                'strFind = "[METHODANALYTESELECTIVITY]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Analyte Selectivity", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Analyte Selectivity"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 205
                'strFind = "[METHODINTERNALSTANDARDSELECTIVITY]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Internal Standard Selectivity", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Internal Standard Selectivity"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"


            Case Is >= intAppFig And intAppFig <> -1

                varReplace = GetAppFig(strFind)



            Case 242 '[TABLEOFATTACHMENTS_01]
                varReplace = ""

            Case 243
                'strFind = "[LASTTABLE#]"
                If ctTableN = 0 Then
                    varReplace = GetLastNumber("Tables")
                Else
                    varReplace = ctTableN
                End If
            Case 244
                'strFind = "[LASTFIGURE#]"
                If ctFigures = 0 Then
                    varReplace = GetLastNumber("Figures")
                Else
                    varReplace = ctFigures
                End If
            Case 245
                'strFind = "[LASTAPPENDIX#]"
                If ctAppendix = 0 Then
                    varReplace = GetLastNumber("Appendices")
                Else
                    varReplace = ctAppendix
                End If

            Case 246 'NOTE: Attachments depricated
                'strFind = "[LASTATTACHMENT#]"
                If ctAppendix = 0 Then 'NOTE: Attachments are appendices
                    varReplace = GetLastNumber("Attachments")
                Else
                    varReplace = ctAppendix
                End If

            Case 247
                'strFind = "[TABLEOFTABLES_01]"
                varReplace = ""

            Case 248
                'strFind = "[TABLEOFFIGURES_01]"
                varReplace = ""

            Case 249
                'strFind = "[TABLEOFAPPENDICES_01]"
                varReplace = ""

            Case 250
                'strFind = "[TABLEOFCONTENTS_01]"
                varReplace = ""

            Case 251
                'strFind = "[TABLEOFTABLES_02]"
                varReplace = ""

            Case 252
                'strFind = "[TABLEOFFIGURES_02]"
                varReplace = ""

            Case 253
                'strFind = "[TABLEOFAPPENDICES_02]"
                varReplace = ""

            Case 254
                'strFind = "[TABLEOFATTACHMENTS_02]"
                varReplace = ""

            Case 255
                'strFind = "[TABLEOFCONTENTS_02]"
                varReplace = ""

            Case 256
                'strfind = "[TOTALANALYTICALRUNS]"
                dv = frmH.dgvAnalyticalRunSummary.DataSource
                Dim ttt As System.Data.DataTable = dv.ToTable("a", True, "Watson Run ID")
                varReplace = 0
                Dim cttt As Short
                For cttt = 0 To ttt.Rows.Count - 1
                    var1 = NZ(ttt.Rows(cttt).Item("Watson Run ID"), "")
                    If Len(var1) = 0 Then
                    Else
                        varReplace = varReplace + 1
                    End If
                Next

            Case 257
                'strFind = "[UNKNOWNSAMPLEMAXCONC]"
                'tblSampleDesign
                Dim maxd As Single = -99999999
                For Count1 = 0 To tblSampleDesign.Rows.Count - 1
                    var1 = NZ(tblSampleDesign.Rows(Count1).Item("CONCENTRATION"), 0)
                    var2 = NZ(tblSampleDesign.Rows(Count1).Item("ALIQUOTFACTOR"), 1)
                    If var2 = 0 Then
                        var2 = 1
                    End If
                    var3 = var1 / var2
                    If var3 > maxd Then
                        maxd = var3
                    End If

                Next
                num1 = SigFigOrDec(maxd, LSigFig, False)
                varReplace = DisplayNum(num1, LSigFig, False)

                strNA1 = charSectionName
                strNA2 = "Max Sample Concentration"
                strNA3 = "Field Code"
                strNA4 = "Field Code"

            Case 258
                'strFind = "[UNKNOWNSAMPLEMINCONC]"
                Dim mind As Single = 999999999
                For Count1 = 0 To tblSampleDesign.Rows.Count - 1
                    var1 = NZ(tblSampleDesign.Rows(Count1).Item("CONCENTRATION"), 0)
                    var2 = NZ(tblSampleDesign.Rows(Count1).Item("ALIQUOTFACTOR"), 1)
                    If var2 = 0 Then
                        var2 = 1
                    End If
                    var3 = var1 / var2
                    If var3 < mind Then
                        mind = var3
                    End If
                Next
                'get lloq
                dv = frmH.dgvWatsonAnalRef.DataSource
                int1 = FindRowDV("LLOQ", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    num1 = 0
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    num1 = 0
                Else
                    num1 = var8 'dv.Item(int1).Item(1)
                End If

                If mind < num1 Then
                    varReplace = "BQL"
                Else
                    num1 = SigFigOrDec(mind, LSigFig, False)
                    varReplace = DisplayNum(num1, LSigFig, False)
                End If

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Min Sample Concentration"
                strNA3 = "Field Code"
                strNA4 = "Field Code"


            Case 259
                'strFind = "[INITIALANALYSISDATE]"
                dgv = frmH.dgvDataWatson
                dv = dgv.DataSource
                int1 = FindRowDV("Initial Analysis Date", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Format(CDate(var8), LTextDateFormat)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Initial Analysis Date"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 260
                'strFind = "[LASTANALYSISDATE]"
                dgv = frmH.dgvDataWatson
                dv = dgv.DataSource
                int1 = FindRowDV("Last Analysis Date", dv)
                'var8 = dg.Item(int1, 1)
                var8 = dv.Item(int1).Item(1)

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Format(CDate(var8), LTextDateFormat)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Last Analysis Date"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"


            Case 261

                'strFind = "[UC_SPECIES]"
                'now find information from dgvDataWatson
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource 'intI is analyte column in dgvMethodValData
                int1 = FindRowDVByCol("Species", dv, "Item")
                var8 = Trim(NZ(dv.Item(int1).Item(1), "[NA]"))

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                varReplace = CapitAllWords(varReplace.ToString)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Species"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 262

                'strFind = "[LC_SPECIES]"
                'now find information from dgvDataWatson
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource 'intI is analyte column in dgvMethodValData
                int1 = FindRowDVByCol("Species", dv, "Item")
                var8 = Trim(NZ(dv.Item(int1).Item(1), "[NA]"))

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                varReplace = LowerCase(varReplace.ToString)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Species"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 263

                'strFind = "[UC_MATRIX]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource 'intI is analyte column in dgvMethodValData
                int1 = FindRowDVByCol("Matrix", dv, "Item")
                'var8 = dv.Item(int1).Item(strAnal)
                var8 = Trim(LCase(NZ(dv.Item(int1).Item(1), "[NA]")))

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                varReplace = CapitAllWords(varReplace.ToString)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Matrix"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 264

                'strFind = "[LC_MATRIX]"
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource 'intI is analyte column in dgvMethodValData
                int1 = FindRowDVByCol("Matrix", dv, "Item")
                'var8 = dv.Item(int1).Item(strAnal)
                var8 = Trim(LCase(NZ(dv.Item(int1).Item(1), "[NA]")))

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                varReplace = LowerCase(varReplace.ToString)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Matrix"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data From Watson Table"

            Case 265

                'strFind = "[UC_ANTICOAGULANT]"
                str1 = NZ(frmH.cbxAnticoagulant.Text, "")
                var8 = str1
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                varReplace = CapitAllWords(varReplace.ToString)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Anticoagulant/Preservative"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Anticoagulant Dropdown Box"

            Case 266

                'strFind = "[LC_ANTICOAGULANT]"
                str1 = NZ(frmH.cbxAnticoagulant.Text, "")
                'determine if text should be capitalized
                var8 = str1

                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = var8
                End If

                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)
                'replace spaces with nbs
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)

                varReplace = UnCapit(varReplace.ToString, True)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Anticoagulant/Preservative"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Anticoagulant Dropdown Box"

            Case 267

                'strFind = "[SAMPLERECEIPTDATES_01]"

                Dim boolWatson = False
                If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                    dv = frmH.dgvSampleReceiptWatson.DataSource
                    boolWatson = True
                Else
                    dv = frmH.dgvSampleReceipt.DataSource
                    boolWatson = False
                End If

                varReplace = "NA"

                For Count1 = 0 To dv.Count - 1

                    If boolWatson Then
                        var1 = dv(Count1).Item("Date Received")
                        var1 = Format(var1, LTextDateFormat)
                    Else
                        var1 = dv(Count1).Item("dtShipmentReceived")
                        var1 = Format(var1, LTextDateFormat)
                    End If

                    var8 = var1

                    If Count1 = 0 Then
                        varReplace = var8
                    Else
                        varReplace = varReplace & ChrW(11) & var8
                    End If

                Next Count1

                strNA1 = charSectionName
                strNA2 = "Sample Receipt Dates Stacked"
                strNA3 = "Sample Receipt"
                strNA4 = "Sample Receitp Tables"

            Case 268 '[SPONSORSTUDYTITLE]

                'dtbl = tblData
                dtbl = tblData
                str2 = "id_tblStudies = " & intIDtblStudies
                drows = dtbl.Select(str2)
                int1 = drows.Length
                'format var8
                If int1 = 0 Then
                    varReplace = "[NA]"
                Else
                    var8 = drows(0).Item("charSponsorStudyTitle")
                    If IsDBNull(var8) Then
                        varReplace = "[NA]"
                    ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                        varReplace = "[NA]"
                    Else
                        varReplace = var8
                    End If
                End If
                'replace hyphens with nbh
                varReplace = Replace(varReplace, "-", NBH, 1, -1, CompareMethod.Text)

                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Sponsor Study Title"
                strNA3 = "Add/Edit Top Level Data"
                strNA4 = "Data table"

            Case 269 '[METHODSUMMARYTABLE]

                'generate Method Summary Table
                varReplace = ""

            Case 270
                'strFind = "[CONTRIBUTINGPERSONNELTITLETABLE]"
                varReplace = ""



            Case 271
                strFind = "[CHARDEMONSTRATEDFREEZETHAW]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 272
                strFind = "[CHARMAXNUMBERFREEZETHAW]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 273
                strFind = "[CHARSTABILITYUNDERSTORAGECOND]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 274
                strFind = "[CHARSTABILITYMAXSTORAGEDUR]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 275
                strFind = "[CHARCORPORATESTUDYID]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 276
                strFind = "[CHARPROTOCOLNUMBER]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 277
                strFind = "[CHARMETHODVALIDATIONTITLE]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 278
                strFind = "[CHARSPONSORMETHODVALIDATIONID]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 279
                strFind = "[CHARSPONSORMETHVALTITLE]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 280
                strFind = "[CHARASSAYDESCRIPTION]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 281
                strFind = "[CHARLMTITLE]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 282
                strFind = "[CHARLMNUMBER]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 283
                strFind = "[NUMSAMPLESIZE]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 284
                strFind = "[CHARSAMPLESIZEUNITS]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 285
                strFind = "[CHARANTICOAGULANT]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 286
                strFind = "[CHARSPECIES]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 287
                strFind = "[CHARMATRIX]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 288
                strFind = "[CHARMAXRUNSIZE]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 289
                strFind = "[CHARANALMETHODTYPE]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 290
                strFind = "[CHARQCCONC]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 291
                strFind = "[CHARCALIBRCONC]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 292
                strFind = "[CHARLLOQ]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 293
                strFind = "[CHARULOQ]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 294
                strFind = "[CHARAVERECANAL]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 295
                strFind = "[CHARAVERECIS]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 296
                strFind = "[CHARINTERQCACCRNG]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 297
                strFind = "[CHARINTERQCPRECRNG]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 298
                strFind = "[CHARINTRAQCACCRNG]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 299
                strFind = "[CHARINTRAQCPRECRNG]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 300
                strFind = "[CHARPROCSTABILITY]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 301
                strFind = "[CHARREFRSTAB]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 302
                strFind = "[CHARLTSTORSTAB]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 303
                strFind = "[CHARDILINTEGR]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 304
                strFind = "[CHARANALSELECT]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 305
                strFind = "[CHARISSELECT]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 306
                strFind = "[CHARVALREPORTNUM]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")
            Case 307
                strFind = "[CHARFTSTORCOND]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                varReplace = NZ(rowsFC(0).Item(str1), "[NA]")


                '20190110 LEE
            Case 308 '20190215 LEE: Deprecated. This is replicate of CHARDEMONSTRATEDFREEZETHAW
                strFind = "[METHODFREEZETHAW]"

                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Freeze/Thaw Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Freeze/Thaw Stability"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"


            Case 309
                strFind = "[METHODBENCHTOP]"

                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Bench-top Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Bench-top Stability in Matrix"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 310
                strFind = "[METHODREINJECTION]"

                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Reinjection Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Reinjection Stability in Matrix"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 311
                strFind = "[METHODBATCHREINJECTION]"

                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Batch Reinjection Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Batch Reinjection Stability in Matrix"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 312
                strFind = "[METHODBLOOD]"

                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Whole Blood Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Whole Blood Stability in Matrix"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 313
                strFind = "[METHODSTOCKSOLUTION]"

                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Stock Solution Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Stock Solution Stability in Matrix"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 314
                strFind = "[METHODSPIKINGSOLUTION]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)

                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Spiking Solution Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Spiking Solution Stability in Matrix"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 315
                strFind = "[METHODAUTOSAMPLER]"

                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Autosampler Stability", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Autosampler Stability in Matrix"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 316 '20190213 LEE:
                strFind = "[CHARREPORTNUMBER]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)

                'can't use rowsFC. There is no column in tblMethodValidation
                dgv = frmH.dgvMethodValData
                dv = dgv.DataSource
                int1 = FindRowDV("Validation Report Number", dv)
                var8 = dv.Item(int1).Item(1)
                If IsDBNull(var8) Then
                    varReplace = "[NA]"
                ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                    varReplace = "[NA]"
                Else
                    varReplace = Trim(var8)
                End If
                'ctArrReportNA = ctArrReportNA + 1
                strNA1 = charSectionName
                strNA2 = "Validation Report Number"
                strNA3 = "Method Validation Data"
                strNA4 = "Method Validation Data Table"

            Case 317
                'strFind = "[METHODASSAYTECHNIQUE]"
                '20190213 LEE:
                'this is actually CHARANALMETHODTYPE in tblMethodValidation
                strFind = "[CHARANALMETHODTYPE]"
                str1 = Mid(strFind, 2, Len(strFind) - 2)
                Try
                    varReplace = NZ(rowsFC(0).Item(str1), "NA")
                Catch ex As Exception
                    var1 = var1
                End Try

        End Select

        Select Case intPos
            Case 269, 270, 271 - 315
            Case Else
                If Len(varReplace) = 0 Or StrComp(NZ(varReplace, "[None]"), "[None]", CompareMethod.Text) = 0 Or StrComp(varReplace, "[NA]", CompareMethod.Text) = 0 Or InStr(varReplace, "[NA]", CompareMethod.Text) > 0 Then
                    'add entries to arrReportNA
                    If Len(strNA1) = 0 Then
                    Else
                        ctArrReportNA = ctArrReportNA + 1
                        arrReportNA(1, ctArrReportNA) = strNA1 'section name
                        arrReportNA(2, ctArrReportNA) = strNA2 'report item
                        arrReportNA(3, ctArrReportNA) = strNA3 'tab name
                        arrReportNA(4, ctArrReportNA) = strNA4 'tab item
                        arrReportNA(5, ctArrReportNA) = strFind 'field code
                    End If

                End If
        End Select


        ReturnSearch = varReplace

    End Function


    Function SearchReplace(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal charSectionName As String, ByVal rng1 As Microsoft.Office.Interop.Word.Range, ByVal boolChar As Boolean, ByVal strChar As String, ByVal intSR As Short, ByVal intS As Short, ByVal intE As Short, ByVal boolFromHeader As Boolean, ByVal boolIgnoreTOC As Boolean, ByVal boolIgnoreTableFigs As Boolean, ByVal intType As Short) As String ', ByVal intS, ByVal intE)

        'also returns Search Item if there is a problem

        '20190221 LEE:
        'if intType = 1 then
        '   is footnote

        Dim var1, var2, var3, var8, var9
        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim Count3 As Int16
        Dim strFind As String
        Dim pos1 As Int16
        Dim pos2 As Int16
        Dim ctTot As Int16
        Dim varReplace
        Dim myRange As Microsoft.Office.Interop.Word.Range
        Dim intRow As Short
        Dim dg As DataGrid
        Dim ts1 As DataGridTableStyle
        Dim dv As System.Data.DataView
        Dim dtbl As System.Data.DataTable
        Dim drows() As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        'Dim myRng As Microsoft.Office.Interop.Word.selection
        Dim myRng As Microsoft.Office.Interop.Word.Range
        Dim strFind1 As String
        Dim intIDtblStudies As Int64
        Dim strNA1 As String
        Dim strNA2 As String
        Dim strNA3 As String
        Dim strNA4 As String
        Dim tblN As System.Data.DataTable
        Dim num1 As Object
        Dim boolNum As Boolean
        Dim dt1 As Date
        Dim dt2 As Date
        Dim tblNick As System.Data.DataTable
        Dim rowsNick() As DataRow
        Dim dgv As DataGridView
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim tbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim strF As String
        Dim strM As String
        Dim strM1 As String
        Dim strM2 As String
        Dim intEnd As Short
        Dim bool1 As Boolean
        Dim intM As Short
        Dim intV As Short
        Dim boolFound As Boolean
        Dim intCount As Short
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim boolRng As Boolean
        Dim boolF As Boolean
        Dim boolF1 As Boolean
        Dim boolH As Boolean
        Dim boolAppFig As Boolean = False
        Dim boolApp As Boolean = False
        Dim boolFig As Boolean = False

        Dim intSa As Short
        Dim incrFC As Short

        SearchReplace = strChar

        ''''wdd.visible = True

        'myRng = wd.selection
        myRng = rng1
        'myRng.Select()
        Try
            myRng.Select()
        Catch ex As Exception
            Exit Function
        End Try
        If intType = 1 Then '20190222 LEE:
            wd.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekFootnotes
            myRng.Select()
        End If
        mySel = wd.Selection

        strFind = "zzzzzzzzzzzz"
        strFind1 = "zzzzzzzzzzzz"
        tblN = tblTableN

        strM = frmH.lblProgress.Text
        If intType = 1 Then
            strM = strM & ChrW(10) & "Searching Footnotes..."
        End If
        frmH.lblProgress.Text = strM

        Try

            Dim tblAF As System.Data.DataTable
            Dim rowsAF() As DataRow
            Dim intRowsAF As Short
            Dim intRowAF As Short
            Dim numAF As Short
            Dim strS As String
            Dim boolSkip As Boolean = False

            tblAF = tblAppFigs
            strF = "ID_TBLSTUDIES = " & id_tblStudies ' & " AND CHARFCID IS NOT NULL"
            strS = "INTORDER ASC"
            rowsAF = tblAF.Select(strF, strS)
            intRowsAF = rowsAF.Length

            Dim intNRows As Int16 = 317 ' 315 ' 307 '270 '269 '268

            intIDtblStudies = id_tblStudies
            If intS = 0 And intE = 0 Then
                intSa = -10
                intEnd = intNRows + intRowsAF
            Else
                intSa = intS
                intEnd = intE
            End If

            numAF = intEnd - intRowsAF + 1

            'record original values
            bool1 = frmH.pb1.Visible
            intM = frmH.pb1.Maximum
            intV = frmH.pb1.Value

            frmH.pb1.Maximum = intEnd
            frmH.pb1.Value = 0
            frmH.pb1.Visible = True

            '''wdd.visible = True

            intCount = 0
            incrFC = 0
            For Count1 = intSa To intEnd

                '*** 20170503 LEE: Deprecate ANALYTE_ and TABLENUMBER_
                If Count1 >= 121 And Count1 <= 134 Then
                    GoTo next1
                End If
                If Count1 >= 136 And Count1 <= 149 Then
                    GoTo next1
                End If


                '[LMAPPENDIXNUMBER]  67
                '[ANALYTICALRUNTABLENUMBER]  68
                '[REPRCHROMAPPENDIXLETTER]  69
                '[REGRESSIONTABLENUMBER]  76
                '[REASSAYTABLENUMBER]  164
                '[REPEATTABLENUMBER]  165

                'Case Is = 22 '[REGRESSIONTABLENUMBERSECTION]'80

                'Case Is = 23 '[CALSTDTABLENUMBERSECTION]'81

                'Case Is = 24 '[SAMPLETABLENUMBERSECTION]'87

                'Case Is = 25 '[QCTABLENUMBERSECTION]'84

                'Case Is = 26 '[REGRESSION]'25

                'Case Is = 27 '[WEIGHTING]'26

                'Case Is = 28 '[QCMEANACCURACYMIN]'32

                'Case Is = 29 '[QCMEANACCURACYMAX]'33

                'Case Is = 30 '[QCPRECISIONMIN]'34

                'Case Is = 31 '[QCPRECISIONMAX]'35

                '[INTEGRATIONTYPE]'41

                '[NUMCOADREPLICATES] 15

                If Count1 >= 67 And Count1 <= 69 Then
                    GoTo next1
                End If

                If Count1 = 76 Then
                    GoTo next1
                End If

                If Count1 >= 164 And Count1 <= 165 Then
                    GoTo next1
                End If

                If Count1 = 87 Then
                    GoTo next1
                End If

                If Count1 = 84 Then
                    GoTo next1
                End If

                If Count1 >= 25 And Count1 <= 26 Then
                    GoTo next1
                End If

                If Count1 >= 32 And Count1 <= 35 Then
                    GoTo next1
                End If

                If Count1 >= 80 And Count1 <= 81 Then
                    GoTo next1
                End If

                If Count1 = 41 Then
                    GoTo next1
                End If

                If Count1 = 15 Then
                    GoTo next1
                End If



                '***

                boolSkip = False
                boolAppFig = False

                Try
                    'var1 = myRng.Text.ToString
                Catch ex As Exception

                End Try

                intSR = Count1

                ''''''''''''''console.writeline("Search: " & Count1)
                intCount = intCount + 1
                boolNum = False
                strFind = "zzzzzzzzzzzz"

                boolH = False


                If intSa = intEnd Then 'ignore
                Else
                    If Count1 >= numAF And boolFromHeader = False Then

                        If boolIgnoreTableFigs Then
                            boolAppFig = False
                        Else
                            intRowAF = Count1 - numAF + 1
                            str1 = NZ(rowsAF(intRowAF - 1).Item("CHARFCID"), "azaza")
                            strFind = "[APPFIGREF_" & str1 & "]"
                            varReplace = ""
                            boolAppFig = True
                        End If

                    Else
                        boolAppFig = False
                    End If

                End If

                Select Case Count1

                    Case -10
                        strFind = "[ANALREFTABLE]"
                    Case -9
                        strFind = "[METHODSUMMARYSTATEMENT]" 'this has to happen early
                    Case 1
                        strFind = "[REPORTTITLE]"
                        boolH = True
                    Case 2
                        strFind = "[REPORTNUMBER]"
                        boolH = True '20180815 LEE: LI00016
                    Case 3
                        strFind = "[REPORTDRAFTDATE]"
                        boolH = True '20180815 LEE: LI00016
                    Case 4
                        strFind = "[REPORTISSUEDATE]"
                        boolH = True '20180815 LEE: LI00016
                    Case 5
                        strFind = "[ACCEPTEDANALYTICALRUNS]"
                    Case 6
                        strFind = "[NUMAPCONC]"
                    Case 7
                        strFind = "[NUMAPREPLICATES]"
                    Case 8
                        strFind = "[NUMLINDILREPLICATES]"
                        boolSkip = True ''20190228 LEE: Deprecated
                    Case 9
                        strFind = "[NUMDILVALUE]"
                        boolSkip = True ''20190228 LEE: Deprecated
                    Case 10
                        strFind = "[NUMDILFACTOR]"
                        boolSkip = True ''20190228 LEE: Deprecated
                    Case 11
                        strFind = "[FULLVALIDATIONNUMBER]"
                    Case 12
                        strFind = "[NUMSEREPLICATES]"
                    Case 13
                        strFind = "[NUMUNIQUELOTS]"
                    Case 14
                        strFind = "[NUMUNIQUEREPLICATES]"
                    Case 15
                        strFind = "[NUMCOADREPLICATES]"
                    Case 16
                        strFind = "[NUMRTSREPLICATES]"
                    Case 17
                        strFind = "[NUMRTSHOURS]"
                    Case 18
                        strFind = "[NUMFTREPLICATES]"
                    Case 19
                        strFind = "[INTERNALSTANDARD]"
                    Case 20
                        strFind = "[LLOQ]"
                    Case 21
                        strFind = "[LLOQUNITS]"
                    Case 22
                        strFind = "[ULOQ]"
                    Case 23
                        strFind = "[ULOQUNITS]"
                    Case 24
                        strFind = "[CALIBRATIONLEVELS]"
                    Case 25
                        strFind = "[REGRESSION]"
                    Case 26
                        strFind = "[WEIGHTING]"
                    Case 27
                        strFind = "[MINIMUMR2]"
                    Case 28
                        strFind = "[ANALYTEMEANACCURACYMIN]"
                    Case 29
                        strFind = "[ANALYTEMEANACCURACYMAX]"
                    Case 30
                        strFind = "[ANALYTEPRECISIONMIN]"
                    Case 31
                        strFind = "[ANALYTEPRECISIONMAX]"
                    Case 32
                        strFind = "[QCMEANACCURACYMIN]"
                    Case 33
                        strFind = "[QCMEANACCURACYMAX]"
                    Case 34
                        strFind = "[QCPRECISIONMIN]"
                    Case 35
                        strFind = "[QCPRECISIONMAX]"
                    Case 36
                        strFind = "[WATSONSTUDYID]"
                    Case 37
                        strFind = "[WATSONPROJECTID]"
                    Case 38
                        strFind = "[MATRIX]"
                        boolH = True
                    Case 39
                        strFind = "[SAMPLESIZE]"
                    Case 40
                        strFind = "[SPECIES]"
                        boolH = True
                    Case 41
                        strFind = "[INTEGRATIONTYPE]"
                    Case 42
                        strFind = "[INITIALEXTRACTIONDATE]"
                    Case 43
                        strFind = "[LASTEXTRACTIONDATE]"
                    Case 44
                        strFind = "[ASSAYTECHNIQUE]"
                    Case 45
                        strFind = "[ASSAYTECHNIQUEACRONYM]"
                    Case 46
                        strFind = "[ANTICOAGULANT]"
                    Case 47
                        strFind = "[SAMPLESIZEUNITS]"
                    Case 48
                        strFind = "[SUBMITTEDTO]"
                    Case 49
                        strFind = "[INSUPPORTOF]"
                    Case 50
                        strFind = "[SUBMITTEDBY]"
                    Case 51
                        strFind = "[CORPORATESTUDY/PROJECTNUMBER]"
                        boolH = True
                    Case 52
                        strFind = "[PROTOCOLNUMBER]"
                        boolH = True
                    Case 53
                        strFind = "[SPONSORSTUDYNUMBER]"
                        boolH = True
                    Case 54
                        strFind = "[DATAARCHIVALLOCATION]"
                    Case 55
                        strFind = "[DATEFIRSTSAMPLESRECEIVED]"
                    Case 56
                        boolSkip = True '20190215 LEE: Deprecated. This is [CHARMAXNUMBERFREEZETHAW]
                        strFind = "[METHODDEMONSTRATEDFREEZE/THAWCYCLES]"
                    Case 57
                        boolSkip = True '20190215 LEE: Deprecated. This is [CHARMAXNUMBERFREEZETHAW]
                        strFind = "[METHODMAXIMUMNUMBEROFFREEZE/THAWCYCLES]"
                    Case 58
                        strFind = "[METHODSTABILITYUNDERSTORAGECONDITIONS]"
                        boolSkip = True '20190213 LEE: deprecated
                    Case 59
                        strFind = "[METHODISSTABILITY>=MAXIMUMSTORAGEDURATION]"
                    Case 60
                        strFind = "[METHODCORPORATESTUDY/PROJECTNUMBER]"
                        boolH = True
                    Case 61
                        strFind = "[METHODPROTOCOLNUMBER]"
                        boolH = True
                        boolSkip = True '20190219 LEE: Deprecated. Replaced with CHARPROTOCOLNUMBER
                    Case 62
                        strFind = "[METHODMETHODVALIDATIONTITLE]"
                        boolH = True
                        boolSkip = True '20190219 LEE: Deprecated. Replaced with CHARMETHODVALIDATIONTITLE
                    Case 63
                        strFind = "[METHODSPONSORMETHODVALIDATIONSTUDYNUMBER]"
                        boolH = True
                        boolSkip = True '20190219 LEE: Deprecated. Replaced with CHARSPONSORMETHODVALIDATIONID
                    Case 64
                        strFind = "[METHODSPONSORMETHODVALIDATIONTITLE]"
                        boolH = True
                    Case 65
                        strFind = "[METHODASSAYPROCEDUREDESCRIPTION]"
                    Case 66
                        strFind = "[FINALREPORTREVIEWDATE]"
                    Case 67
                        strFind = "[LMAPPENDIXNUMBER]"
                    Case 68
                        strFind = "[ANALYTICALRUNTABLENUMBER]"
                    Case 69
                        strFind = "[REPRCHROMAPPENDIXLETTER]"
                    Case 70
                        strFind = "[ANALYTE]"
                        boolH = True
                    Case 71
                        strFind = "[NUMBEROFSAMPLES]"
                    Case 72
                        strFind = "[REPRCHROMWATSONRUNNUMBER]"
                    Case 73
                        strFind = "[STORAGETEMP]" 'tblSampleReceipt.charStorageTemp
                    Case 74
                        strFind = "[SHIPMENTCOUNT]"
                    Case 75
                        strFind = "[SHIPMENTCONDITION]"
                    Case 76
                        strFind = "[REGRESSIONTABLENUMBER]"
                    Case 77
                        strFind = "[CALIBRATIONLEVELSSECTION]"
                    Case 78
                        strFind = "[REGRESSIONSECTION]"
                    Case 79
                        strFind = "[R2SECTION]"
                    Case 80
                        strFind = "[REGRESSIONTABLENUMBERSECTION]"
                    Case 81
                        strFind = "[CALSTDTABLENUMBERSECTION]"
                    Case 82
                        strFind = "[ACCURACYSECTION]"
                    Case 83
                        strFind = "[PRECISIONSECTION]"
                    Case 84
                        strFind = "[QCTABLENUMBERSECTION]"
                    Case 85
                        strFind = "[QCACCURACYSECTION]"
                    Case 86
                        strFind = "[QCPRECISIONSECTION]"
                    Case 87
                        strFind = "[SAMPLETABLENUMBERSECTION]"
                    Case 88
                        strFind = "[STUDYSAMPLEBQLSECTION1]"
                    Case 89
                        strFind = "[STUDYSAMPLEBQLSECTION2]"
                    Case 90
                        strFind = "[PLURALNUMBER]"
                    Case 91
                        strFind = "[PLURALSUMMARY]"
                    Case 92
                        strFind = "[NUMBEROFSAMPLESVERBOSESECTION]"
                    Case 93
                        strFind = "[SHIPMENTCOUNTSECTION]"
                    Case 94
                        strFind = "[SHIPMENTRECEIPTDATESECTION]"
                    Case 95
                        strFind = "[LABMETHODNAME]"
                    Case 96
                        strFind = "[LABMETHODNUMBER]"
                    Case 97
                        strFind = "[SHIPMENTSOURCE]"
                    Case 98
                        strFind = "[QCCONCENTRATIONCOUNT]" 'numQCConcCount
                    Case 99
                        strFind = "[REPLICATEDILUTIONQC]" 'numRepDilnQC
                        '20190228 LEE: deprecated
                        boolSkip = True
                    Case 100 'replace any 'deg C' with symbol
                        strFind = "degC"
                        boolH = True
                    Case 101 'replace any 'deg C' with symbol
                        strFind = "deg C"
                        boolH = True
                    Case 102 'replace any '+/-' with symbol
                        strFind = "+/-"
                        boolH = True
                    Case 103 'replace any '<=' with symbol
                        strFind = "<="
                        boolH = True
                    Case 104 'replace any '>=' with symbol
                        strFind = ">="
                        boolH = True

                    Case 105
                        strFind = "[NUMFTCYCLES]"
                    Case 106
                        strFind = "[FXSRUNID]"
                    Case 107
                        strFind = "[FXSTEMP]"
                    Case 108
                        strFind = "[FXSTIME]"
                    Case 110
                        strFind = "[SSSTEMP]"
                    Case 111
                        strFind = "[SSSTIME]"
                    Case 112
                        strFind = "[NUMSSSREPLICATES]"
                    Case 113
                        strFind = "[SpSSTEMP]"
                    Case 114
                        strFind = "[SpSSTIME]"
                    Case 115
                        strFind = "[NUMSpSSREPLICATES]"
                    Case 116
                        strFind = "[LTSTEMP]"
                    Case 117
                        strFind = "[LTSTIME]"
                    Case 118
                        strFind = "[NUMLTSREPLICATES]"
                    Case 119
                        strFind = "[ANALYTERESP]" 'add 'respectively' at the end
                    Case 120
                        strFind = "[INTERNALSTANDARDRESP]" 'add 'respectively' at the end
                    Case 121
                        strFind = "[TABLENUMBER_4]"
                        boolSkip = True
                    Case 122
                        strFind = "[TABLENUMBER_12]"
                        boolSkip = True
                    Case 123
                        strFind = "[TABLENUMBERALL_13]"
                        boolSkip = True
                    Case 124
                        strFind = "[TABLENUMBERALL_14]"
                        boolSkip = True
                    Case 125
                        strFind = "[TABLENUMBERALL_15]"
                        boolSkip = True
                    Case 126
                        strFind = "[TABLENUMBER_2]"
                        boolSkip = True
                    Case 127
                        strFind = "[ANALYTE_11]"
                        boolSkip = True
                        boolH = False
                    Case 128
                        strFind = "[ANALYTE_3]"
                        boolSkip = True
                        boolH = False
                    Case 129
                        strFind = "[ANALYTE_4]"
                        boolSkip = True
                        boolH = False
                    Case 130
                        strFind = "[ANALYTE_12]"
                        boolSkip = True
                        boolH = False
                    Case 131
                        strFind = "[ANALYTE_13]"
                        boolSkip = True
                        boolH = False
                    Case 132
                        strFind = "[ANALYTE_14]"
                        boolSkip = True
                        boolH = False
                    Case 133
                        strFind = "[ANALYTE_15]"
                        boolSkip = True
                        boolH = False
                    Case 134
                        'strFind = "[ANALYTE_2]" 'deprecate
                        strFind = "[ANALYTE_2xxx]"
                        boolSkip = True
                        boolH = False
                    Case 135
                        strFind = "[MAXRUNSIZE]"
                        boolSkip = True '20190222 LEE: deprecated. Use CHARMAXRUNSIZE
                    Case 136
                        strFind = "[TABLENUMBER_17]"
                        boolSkip = True
                    Case 137
                        strFind = "[ANALYTE_17]"
                        boolSkip = True
                        boolH = False
                    Case 138
                        strFind = "[TABLENUMBER_18]"
                        boolSkip = True
                    Case 139
                        strFind = "[ANALYTE_18]"
                        boolSkip = True
                        boolH = False
                    Case 140
                        strFind = "[TABLENUMBER_19]"
                        boolSkip = True
                    Case 141
                        strFind = "[ANALYTE_19]"
                        boolSkip = True
                        boolH = False
                    Case 142
                        strFind = "[TABLENUMBER_21]"
                        boolSkip = True
                    Case 143
                        strFind = "[ANALYTE_21]"
                        boolSkip = True
                        boolH = False
                    Case 144
                        strFind = "[TABLENUMBER_22]"
                        boolSkip = True
                    Case 145
                        strFind = "[ANALYTE_22]"
                        boolSkip = True
                        boolH = False
                    Case 146
                        strFind = "[TABLENUMBER_23]"
                        boolSkip = True
                    Case 147
                        strFind = "[ANALYTE_23]"
                        boolSkip = True
                        boolH = False
                    Case 148
                        strFind = "[TABLENUMBER_29]"
                        boolSkip = True
                    Case 149
                        strFind = "[ANALYTE_29]"
                        boolSkip = True
                        boolH = False
                    Case 150
                        strFind = "[PAGENUMBER]"
                        boolH = True
                    Case 151
                        strFind = "[TOTALPAGES]"
                        boolH = True
                    Case 152
                        strFind = "[SUBMITTEDTOFULL]"
                    Case 153
                        strFind = "[SUBMITTEDBYFULL]"
                    Case 154
                        strFind = "[INSUPPORTOFFULL]"
                    Case 155
                        strFind = "[COVERPAGESIGNATURE]" ' NOT USED!!!

                    Case 160
                        strFind = "[SECTIONBREAKNEXTPAGE]"
                    Case 161
                        strFind = "[INSERTPAGEBREAK]"
                    Case 162 'REMOVED
                        strFind = "[FIRSTPAGESPECIAL]"
                    Case 163
                        strFind = "[CONTRIBUTINGPERSONNELTABLE]"
                    Case 164
                        strFind = "[REASSAYTABLENUMBER]"
                    Case 165
                        strFind = "[REPEATTABLENUMBER]"
                    Case 166
                        strFind = "[INSERTQATABLE]"
                    Case 167
                        strFind = "[REPORTSTATUS]"
                        boolH = True '20180815 LEE: LI00016
                    Case 168
                        strFind = "[ABSTRACTANALYTEINFO]"
                    Case 169
                        strFind = "[ANALYTE_IUPAC]"
                    Case 170
                        strFind = "[CALIBRSTANDARDLIST]"
                    Case 171
                        strFind = "[QCSTANDARDLIST]"
                    Case 172
                        strFind = "[TIMEZONE]"
                    Case 173
                        strFind = "[STUDYSAMPLECONCENTRATIONTABLE]"
                    Case 174
                        strFind = "[SAMPLERECEIPTTABLE1]"
                    Case 175
                        strFind = "[CALSTDTABLE1]"
                    Case 176
                        strFind = "[CALSTDTABLE2]"
                    Case 177
                        strFind = "[QCSECTION]"
                    Case 178
                        strFind = "[DILUTIONQCSECTION]"
                    Case 179
                        strFind = "[QCTABLE1]"
                    Case 180
                        strFind = "[ABSTRACTANALYTEINFO1]"
                    Case 181
                        strFind = "[ABSTRACTANALYTEINFO2]"
                    Case 182
                        strFind = "[DATENOW]"
                    Case 183
                        strFind = "[DATETIMENOW]"
                    Case 184
                        strFind = "[OUTLIERMETHOD]"
                    Case 185
                        strFind = "[TABLESECTION]"
                    Case 186
                        strFind = "[APPENDIXSECTION]"
                    Case 187
                        strFind = "[FIGURESECTION]"
                    Case 188
                        strFind = "[REGRESSIONEQUATION]"
                    Case 189
                        strFind = "[DATELASTSAMPLESRECEIVED]"
                    Case 190
                        strFind = "[BIASORDIFFQC]"
                    Case 191
                        strFind = "[BIASORDIFFCALIBR]"
                    Case 192
                        strFind = "[ANTICOAGULANTMETHOD]"
                        boolSkip = True '20190213 LEE: Deprecated
                    Case 193
                        boolSkip = True '20190215 LEE: Deprecated. This is [CHARDEMONSTRATEDFREEZETHAW]
                        strFind = "[FREEZETHAWSTORAGECOND]"
                    Case 194
                        strFind = "[STUDYSTARTDATE]"
                    Case 195
                        strFind = "[STUDYENDDATE]"
                    Case 196 'NOTE: Attachments depricated
                        strFind = "[ATTACHMENTSECTION_01]"
                    Case 197
                        strFind = "[WATSONSTUDYTITLE]"
                    Case 198
                        strFind = "[WATSONREPCHROMSECTION]"
                    Case 199
                        strFind = "[REGRESSIONTABLENUMBER]"


                    Case 200
                        strFind = "[METHODPROCESSSTABILITY]"
                    Case 201
                        strFind = "[METHODREFRIGERATEDSTABILITYINMATRIX]"
                        boolSkip = True '20190213 LEE: deprecated. Use METHODREINJECTION
                    Case 202
                        strFind = "[METHODLONGTERMSTORAGESTABILITYINMATRIX]"
                    Case 203
                        strFind = "[METHODDILUTIONINTEGRITY]"
                    Case 204
                        strFind = "[METHODANALYTESELECTIVITY]"
                    Case 205
                        strFind = "[METHODINTERNALSTANDARDSELECTIVITY]"


                        'Case Count1 > 195 And Count1 < 247
                        'boolSkip = True

                        'these must be done last
                    Case 242
                        strFind = "[TABLEOFATTACHMENTS_01]"
                    Case 243
                        strFind = "[LASTTABLE#]"
                    Case 244
                        strFind = "[LASTFIGURE#]"
                    Case 245
                        strFind = "[LASTAPPENDIX#]"
                    Case 246
                        strFind = "[LASTATTACHMENT#]"

                    Case 247
                        strFind = "[TABLEOFTABLES_01]"
                    Case 248
                        strFind = "[TABLEOFFIGURES_01]"
                    Case 249
                        strFind = "[TABLEOFAPPENDICES_01]"
                    Case 250
                        strFind = "[TABLEOFCONTENTS_01]"

                    Case 251
                        strFind = "[TABLEOFTABLES_02]"
                    Case 252
                        strFind = "[TABLEOFFIGURES_02]"
                    Case 253
                        strFind = "[TABLEOFAPPENDICES_02]"
                    Case 254
                        strFind = "[TABLEOFATTACHMENTS_02]"
                    Case 255
                        strFind = "[TABLEOFCONTENTS_02]"
                    Case 256
                        strFind = "[TOTALANALYTICALRUNS]"
                    Case 257
                        strFind = "[UNKNOWNSAMPLEMAXCONC]"
                    Case 258
                        strFind = "[UNKNOWNSAMPLEMINCONC]"

                    Case 259
                        strFind = "[INITIALANALYSISDATE]"
                    Case 260
                        strFind = "[LASTANALYSISDATE]"



                    Case 261
                        strFind = "[UC_SPECIES]"
                    Case 262
                        strFind = "[LC_SPECIES]"
                    Case 263
                        strFind = "[UC_MATRIX]"
                    Case 264
                        strFind = "[LC_MATRIX]"
                    Case 265
                        strFind = "[UC_ANTICOAGULANT]"
                    Case 266
                        strFind = "[LC_ANTICOAGULANT]"
                    Case 267
                        strFind = "[SAMPLERECEIPTDATES_01]"

                    Case 268
                        strFind = "[SPONSORSTUDYTITLE]"
                        boolH = True

                    Case 269 '20180816 LEE:
                        strFind = "[METHODSUMMARYTABLE]"

                    Case 270
                        strFind = "[CONTRIBUTINGPERSONNELTITLETABLE]"

                        '20180829 LEE:
                        'MethodValidationStuff
                    Case 271
                        strFind = "[CHARDEMONSTRATEDFREEZETHAW]"
                    Case 272
                        strFind = "[CHARMAXNUMBERFREEZETHAW]"
                    Case 273
                        strFind = "[CHARSTABILITYUNDERSTORAGECOND]"
                    Case 274
                        strFind = "[CHARSTABILITYMAXSTORAGEDUR]"
                    Case 275
                        strFind = "[CHARCORPORATESTUDYID]"
                    Case 276
                        strFind = "[CHARPROTOCOLNUMBER]"
                    Case 277
                        strFind = "[CHARMETHODVALIDATIONTITLE]"
                    Case 278
                        strFind = "[CHARSPONSORMETHODVALIDATIONID]"
                    Case 279
                        strFind = "[CHARSPONSORMETHVALTITLE]"
                    Case 280
                        strFind = "[CHARASSAYDESCRIPTION]"
                    Case 281
                        strFind = "[CHARLMTITLE]"
                    Case 282
                        strFind = "[CHARLMNUMBER]"
                    Case 283
                        strFind = "[NUMSAMPLESIZE]"
                    Case 284
                        strFind = "[CHARSAMPLESIZEUNITS]"
                    Case 285
                        strFind = "[CHARANTICOAGULANT]"
                    Case 286
                        strFind = "[CHARSPECIES]"
                    Case 287
                        strFind = "[CHARMATRIX]"
                    Case 288
                        strFind = "[CHARMAXRUNSIZE]"
                    Case 289
                        strFind = "[CHARANALMETHODTYPE]"
                    Case 290
                        strFind = "[CHARQCCONC]"
                    Case 291
                        strFind = "[CHARCALIBRCONC]"
                    Case 292
                        strFind = "[CHARLLOQ]"
                    Case 293
                        strFind = "[CHARULOQ]"
                    Case 294
                        strFind = "[CHARAVERECANAL]"
                    Case 295
                        strFind = "[CHARAVERECIS]"
                    Case 296
                        strFind = "[CHARINTERQCACCRNG]"
                    Case 297
                        strFind = "[CHARINTERQCPRECRNG]"
                    Case 298
                        strFind = "[CHARINTRAQCACCRNG]"
                    Case 299
                        strFind = "[CHARINTRAQCPRECRNG]"
                    Case 300
                        strFind = "[CHARPROCSTABILITY]"
                    Case 301
                        strFind = "[CHARREFRSTAB]"
                    Case 302
                        strFind = "[CHARLTSTORSTAB]"
                    Case 303
                        strFind = "[CHARDILINTEGR]"
                    Case 304
                        strFind = "[CHARANALSELECT]"
                    Case 305
                        strFind = "[CHARISSELECT]"
                    Case 306
                        strFind = "[CHARVALREPORTNUM]"
                    Case 307
                        strFind = "[CHARFTSTORCOND]"

                        '20190110 LEE
                    Case 308
                        boolSkip = True '20190215 LEE: this is replicate of CHARDEMONSTRATEDFREEZETHAW
                        strFind = "[METHODFREEZETHAW]"
                    Case 309
                        strFind = "[METHODBENCHTOP]"
                    Case 310
                        strFind = "[METHODREINJECTION]"
                    Case 311
                        strFind = "[METHODBATCHREINJECTION]"
                    Case 312
                        strFind = "[METHODBLOOD]"
                    Case 313
                        strFind = "[METHODSTOCKSOLUTION]"
                    Case 314
                        strFind = "[METHODSPIKINGSOLUTION]"
                    Case 315
                        strFind = "[METHODAUTOSAMPLER]"

                    Case 316 '20190213 LEE:
                        strFind = "[CHARREPORTNUMBER]"
                        boolSkip = True '20190220 LEE: Use CHARVALREPORTNUM
                    Case 317
                        strFind = "[METHODASSAYTECHNIQUE]"

                End Select

                If boolSkip Then
                    GoTo next1
                End If

                If boolFromHeader And boolH = False Then
                    GoTo next1
                End If

                If Count1 > 241 And Count1 < 256 Then
                    If boolIgnoreTOC Then
                        GoTo next1
                    End If
                End If

                If boolIgnoreTableFigs Then
                    Select Case Count1

                        Case -9, 67, 68, 69, 72, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 119, 120, 121, 122, 123, 124, 125, 126, 164, 165, 176, 179, 185, 186, 187, 198, 199
                            GoTo next1
                        Case 127, 128, 129, 130, 131, 132, 133, 134, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148, 149
                            GoTo next1
                    End Select
                End If

                If boolChar Then
                    If InStr(1, strChar, strFind, CompareMethod.Text) > 0 Then
                        boolF = False
                        If StrComp(strFind, "[PAGENUMBER]", CompareMethod.Text) = 0 Then
                        ElseIf StrComp(strFind, "[TOTALPAGES]", CompareMethod.Text) = 0 Then
                        Else
                            varReplace = ReturnSearch(wd, charSectionName, Count1, strFind, myRng, numAF)
                            str1 = Replace(strChar, strFind, NZ(varReplace, "[NA]"), 1, -1, CompareMethod.Text)
                            strChar = str1
                            SearchReplace = strChar
                        End If

                    End If

                Else

                    If Count1 >= incrFC Then
                        If intType = 1 Then
                            frmH.lblProgress.Text = "Footnote Field Code Search:" & ChrW(10) & strFind
                        ElseIf boolFromHeader Then
                            'frmH.lblProgress.Text = "Header Field Code Search:" & ChrW(10) & strFind
                            frmH.lblProgress.Text = strM & ChrW(10) & strFind
                        Else
                            frmH.lblProgress.Text = "Field Code Search:" & ChrW(10) & strFind
                        End If
                        frmH.lblProgress.Refresh()
                        'increment incrFC in pb1.refresh block below
                    ElseIf Count1 = intEnd Then
                        If boolFromHeader Then
                            'frmH.lblProgress.Text = "Footer Field Code Search:" & ChrW(10) & strFind
                            frmH.lblProgress.Text = strM & ChrW(10) & strFind
                        Else
                            frmH.lblProgress.Text = "Field Code Search:" & ChrW(10) & strFind
                        End If
                        frmH.lblProgress.Refresh()
                    End If
                    If intCount > frmH.pb1.Maximum Then
                        intCount = 1
                    End If
                    If Count1 >= incrFC Then
                        frmH.pb1.Value = intCount
                        frmH.pb1.Refresh()
                        incrFC = incrFC + 10
                    End If

                    '''wdd.visible = True


                    If StrComp(strFind, strFind1, CompareMethod.Text) = 0 Then 'do not replace
                        var1 = var2
                    Else

                        If boolFromHeader Then 'must reset range
                            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                            wd.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            myRng = wd.Selection.Range
                        ElseIf intType = 1 Then 'must select all
                            'wd.Selection.WholeStory()
                        End If
                        With myRng.Find
                            .ClearFormatting()
                            '.Text = strFind
                            .Forward = True
                            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                            .Execute(FindText:=strFind)

                            If .Found Then
                                boolFound = True
                            Else
                                boolFound = False
                            End If

                        End With

                        '.selection.SetRange(Start:=0, End:=pos2)

                        If boolFound Then

                            If boolAppFig Then
                                varReplace = ""
                                boolRng = False
                            Else
                                Try
                                    varReplace = ReturnSearch(wd, charSectionName, Count1, strFind, myRng, numAF)
                                Catch ex As Exception
                                    varReplace = "[NA]"
                                End Try
                                boolRng = True
                            End If


                            'boolRng = True

                            Select Case Count1
                                Case -9
                                    boolRng = False
                                Case 68
                                    boolRng = False
                                Case 150
                                    boolRng = False
                                Case 151
                                    boolRng = False
                                Case 155 To 166
                                    boolRng = False
                                Case 168
                                    boolRng = False
                                Case 173 To 176
                                    boolRng = False
                                Case 179 To 181
                                    boolRng = False
                                Case 185 To 187
                                    boolRng = False
                                Case 196, 242
                                    boolRng = False
                                Case 247 To 255
                                    boolRng = False
                                Case 269, 270
                                    boolRng = False
                            End Select

                            '''wdd.visible = True


                            Dim boolBig As Boolean = False
                            Dim vReplace
                            Dim strAA As String = ""
                            vReplace = varReplace
                            If Len(varReplace) > 255 Then
                                boolBig = True
                                varReplace = Mid(vReplace, 1, 255)
                                strAA = Mid(vReplace, 266, Len(vReplace))
                            End If

                            If boolRng And boolBig = False Then 'use find range
                                With myRng.Find
                                    .ClearFormatting()
                                    .Text = strFind
                                    With .Replacement
                                        .ClearFormatting()
                                        If boolNum Then
                                            .Text = num1
                                        Else
                                            .Text = varReplace
                                        End If
                                    End With

                                    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)

                                End With
                            Else 'use find selection
                                'must do this because changes must be done at the selection
                                boolF1 = True

                                Do Until boolF1 = False
                                    With mySel.Find
                                        .ClearFormatting()
                                        .Text = strFind
                                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                                        .Forward = True
                                        With .Replacement
                                            .ClearFormatting()
                                            .Text = varReplace
                                        End With
                                        Try

                                            '''wdd.visible = True

                                            Do While .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True, Forward:=True)


                                                boolF1 = .Found
                                                If boolF1 Then
                                                Else
                                                    Exit Do
                                                End If

                                                If boolBig Then
                                                    ''wdd.visible = True
                                                    wd.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                                                    wd.Selection.TypeText(strAA)
                                                End If

                                                'wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                                                Select Case Count1
                                                    Case 68 'ANALYTICALRUNTABLENUMBER
                                                        '''wdd.visible = True
                                                        Call MVValidationDesignDescr(wd)
                                                    Case 150
                                                        'strFind = "[PAGENUMBER]"
                                                        Call InsertPageNumber(wd)
                                                    Case 151
                                                        'strFind = "[TOTALPAGES]"
                                                        Call InsertTotalPages(wd)
                                                    Case 155
                                                        'strfind="[COVERPAGESIGNATURE]" NOT USED!!
                                                        'Call CoverPageSignature(wd)

                                                    Case 250
                                                        'strFind = "[TABLEOFCONTENTS_01]"
                                                        Try
                                                            boolTOC = True
                                                            Call TableofContents_01(wd, Count1)
                                                        Catch ex As Exception

                                                        End Try
                                                    Case 255
                                                        'strFind = "[TABLEOFCONTENTS_02]"
                                                        Try
                                                            boolTOC = True
                                                            Call TableofContents_01(wd, Count1)
                                                        Catch ex As Exception

                                                        End Try


                                                    Case 242
                                                        'STRFIND="[TABLEOFATTACHMENTS_01]
                                                        Try
                                                            boolTOA = True
                                                            Call TableofAttachments_01(wd, Count1)
                                                            Try
                                                                Call FormatIndex(wd, 84, 19)
                                                                'MUST do this here because uses margins of current page/section
                                                                'I know it's redundant, but gotta do it
                                                            Catch ex As Exception

                                                            End Try
                                                        Catch ex As Exception

                                                        End Try
                                                    Case 254
                                                        'STRFIND="[TABLEOFATTACHMENTS_02]
                                                        Try
                                                            boolTOA = True
                                                            Call TableofAttachments_01(wd, Count1)
                                                            Try
                                                                Call FormatIndex(wd, 84, 19)
                                                                'MUST do this here because uses margins of current page/section
                                                                'I know it's redundant, but gotta do it
                                                            Catch ex As Exception

                                                            End Try
                                                        Catch ex As Exception

                                                        End Try

                                                    Case 247
                                                        'strFind = "[TABLEOFTABLES_01]"
                                                        var1 = var1
                                                        '''wdd.visible = True
                                                        Try
                                                            boolTOT = True
                                                            Call TableofTables_01(wd, Count1)
                                                            Try
                                                                Call FormatIndex(wd, 84, 19)
                                                                'MUST do this here because uses margins of current page/section
                                                                'I know it's redundant, but gotta do it
                                                            Catch ex As Exception

                                                            End Try
                                                        Catch ex As Exception

                                                        End Try
                                                    Case 251
                                                        'strFind = "[TABLEOFTABLES_02]"
                                                        var1 = var1
                                                        '''wdd.visible = True
                                                        Try
                                                            boolTOT = True
                                                            Call TableofTables_01(wd, Count1)
                                                            Try
                                                                Call FormatIndex(wd, 84, 19)
                                                                'MUST do this here because uses margins of current page/section
                                                                'I know it's redundant, but gotta do it
                                                            Catch ex As Exception

                                                            End Try
                                                        Catch ex As Exception

                                                        End Try

                                                    Case 248
                                                        'strFind = "[TABLEOFFIGURES_01]"
                                                        Try
                                                            boolTOF = True
                                                            Call TableofFigures_01(wd, Count1)
                                                            Try
                                                                Call FormatIndex(wd, 84, 19)
                                                                'MUST do this here because uses margins of current page/section
                                                                'I know it's redundant, but gotta do it
                                                            Catch ex As Exception

                                                            End Try
                                                        Catch ex As Exception

                                                        End Try
                                                    Case 252
                                                        'strFind = "[TABLEOFFIGURES_02]"
                                                        Try
                                                            boolTOF = True
                                                            Call TableofFigures_01(wd, Count1)
                                                            Try
                                                                Call FormatIndex(wd, 84, 19)
                                                                'MUST do this here because uses margins of current page/section
                                                                'I know it's redundant, but gotta do it
                                                            Catch ex As Exception

                                                            End Try
                                                        Catch ex As Exception

                                                        End Try

                                                    Case 249
                                                        'strFind = "[TABLEOFAPPENDICES_01]"
                                                        Try
                                                            boolTOA = True
                                                            Call TableofAppendices_01(wd, Count1)
                                                            Try
                                                                Call FormatIndex(wd, 84, 19)
                                                                'MUST do this here because uses margins of current page/section
                                                                'I know it's redundant, but gotta do it
                                                            Catch ex As Exception

                                                            End Try
                                                        Catch ex As Exception

                                                        End Try
                                                    Case 253
                                                        'strFind = "[TABLEOFAPPENDICES_02]"
                                                        Try
                                                            boolTOA = True
                                                            Call TableofAppendices_01(wd, Count1)
                                                            Try
                                                                Call FormatIndex(wd, 84, 19)
                                                                'MUST do this here because uses margins of current page/section
                                                                'I know it's redundant, but gotta do it
                                                            Catch ex As Exception

                                                            End Try
                                                        Catch ex As Exception

                                                        End Try

                                                    Case 160 'SECTIONBREAKNEXTPAGE
                                                        wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                                                        ''''wdd.visible = True
                                                    Case 161 'INSERTPAGEBREAK
                                                        wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
                                                    Case 162 'first page special'REMOVED
                                                        Dim dtbl1 As System.Data.DataTable
                                                        Dim rows() As DataRow
                                                        Dim strF1 As String
                                                        Dim boolG As Boolean
                                                        strF1 = "ID_TBLSTUDIES = " & id_tblStudies
                                                        rows = dtbl1.Select(strF1)
                                                        boolG = True
                                                        Try
                                                            var1 = rows(0).Item("BOOLDIFFFIRSTPAGE")
                                                            If var1 = -1 Then
                                                                boolG = True
                                                            Else
                                                                boolG = False
                                                            End If
                                                        Catch ex As Exception

                                                        End Try
                                                        If boolG Then
                                                            'insert a nextpagebreak
                                                            wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                                                            With wd.Selection.PageSetup
                                                                '.DifferentFirstPageHeaderFooter = True'don't do this anymore
                                                            End With
                                                        End If
                                                    Case 163 '[CONTRIBUTINGPERSONNELTABLE
                                                        Call ContributingPersonnel(wd, False)
                                                    Case 270 '[CONTRIBUTINGPERSONNELTITLETABLE
                                                        Call ContributingPersonnel(wd, True)
                                                    Case 166 'insertqatable
                                                        Call QATable(wd)
                                                    Case 168 'ABSTRACTANALYTEINFO
                                                        Call GuWuAbstract01_26(wd) '
                                                        varReplace = ""
                                                    Case 173 '[STUDYSAMPLECONCENTRATIONTABLE]
                                                        Call GuWuStudySample01(wd)
                                                    Case 174 'strFind = "[SAMPLERECEIPTTABLE1]"
                                                        Call GuWuSampleReceiptStatement01(wd)
                                                    Case -9 'strFind = "[METHODSUMMARYSTATEMENT]"
                                                        Call GuWuMethodSummaryStatement01(wd)
                                                    Case 175 'strFind = "[CALSTDTABLE1]"
                                                        Call DoCALSTDTABLE1(wd)
                                                    Case 176 'strFind = "[CALSTDTABLE2]"
                                                        Call DoCALSTDTABLE2(wd)
                                                    Case 179 '
                                                        Call DoQCTABLE1(wd)
                                                    Case 180 '[ABSTRACTANALYTEINFO1]
                                                        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                                                        Call GuWuObjective01_27(wd) '
                                                        varReplace = ""
                                                    Case 181 '[ABSTRACTANALYTEINFO2]
                                                        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                                                        Call ABSTRACTANALYTEINFO2(wd) '
                                                        varReplace = "[ABSTRACTANALYTEINFO2]"

                                                    Case 185
                                                        'strFind = "[TABLESECTION]"
                                                        If boolDoTables Then

                                                            Try
                                                                Call PrepareWatson(wd)
                                                            Catch ex As Exception
                                                                str1 = "There seems to have been a problem creating the figure/table/appendix portion of this report." & ChrW(10) & ChrW(10)
                                                                str1 = str1 & "Try generating the report again." & ChrW(10) & ChrW(10)
                                                                str1 = str1 & "If the problem persists, please contract your StudyDoc Administrator."
                                                                MsgBox(str1, MsgBoxStyle.Information, "Error in report body...")
                                                                ''''wdd.visible = True
                                                                'GoTo end1
                                                            End Try
                                                        End If

                                                    Case 186
                                                        'strFind = "[APPENDIXSECTION]"
                                                        If boolDoTables Then
                                                            Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                                                            boolAppFigSectionStart = True
                                                            Call InsertGraphics("Appendix", wd, False, 0, False)
                                                            var1 = var1 'debug
                                                        End If

                                                    Case 187
                                                        'strFind = "[FIGURESECTION]"
                                                        If boolDoTables Then
                                                            Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                                                            boolAppFigSectionStart = True
                                                            Call InsertGraphics("Figure", wd, False, 0, False)
                                                        End If

                                                    Case 196 'NOTE: Attachments deprecated
                                                        'strFind = "[ATTACHMENTSECTION_01]"
                                                        If boolDoTables Then
                                                            Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                                                            Call InsertGraphics("Appendix", wd, False, 0, False)
                                                        End If

                                                    Case 269 'Method summary table

                                                        Call SummaryTableAppendix(wd)
                                                        var1 = var1

                                                End Select


                                                If boolAppFig Then
                                                    '''wdd.visible = True

                                                    'find appfig stuff
                                                    Dim strFL As String
                                                    Dim rowsAA() As DataRow
                                                    Dim strFF As String
                                                    Dim intAA As Short
                                                    boolApp = rowsAF(intRowAF - 1).Item("BOOLAPPENDIX")
                                                    boolFig = rowsAF(intRowAF - 1).Item("BOOLFIGURE")

                                                    Dim boolIncl As Boolean = True
                                                    Dim intIncl As Short
                                                    Try
                                                        intIncl = rowsAF(intRowAF - 1).Item("BOOLINCLUDEINREPORT")
                                                        If intIncl = 0 Then
                                                            boolIncl = False
                                                        Else
                                                            boolIncl = True
                                                        End If
                                                    Catch ex As Exception
                                                        boolIncl = False
                                                    End Try

                                                    If boolIncl Then

                                                        strFL = "Figure"
                                                        If boolApp Then
                                                            strFL = "Appendix"
                                                            'get appendix number from tblAppendix
                                                            strFF = "CHARFCID = '" & rowsAF(intRowAF - 1).Item("CHARFCID") & "'"
                                                            rowsAA = tblAppendix.Select(strFF)
                                                            If rowsAA.Length = 0 Then
                                                                intAA = 0
                                                            Else
                                                                intAA = rowsAA(0).Item("APPENDIXNUMBER")
                                                            End If
                                                        End If
                                                        If boolFig Then
                                                            strFL = "Figure"
                                                            'get fig number from tblFigures
                                                            strFF = "CHARFCID = '" & rowsAF(intRowAF - 1).Item("CHARFCID") & "'"
                                                            rowsAA = tblFigures.Select(strFF)
                                                            If rowsAA.Length = 0 Then
                                                                intAA = 0
                                                            Else
                                                                intAA = rowsAA(0).Item("FIGURENUMBER")
                                                            End If
                                                        End If

                                                        If intAA = 0 Then
                                                            wd.Selection.TypeText("[NA]")
                                                        Else
                                                            '20170418 LEE: Hyperlinking must loop 1 to rowsAA.length
                                                            For Count2 = 0 To rowsAA.Length - 1
                                                                If boolApp Then
                                                                    intAA = rowsAA(Count2).Item("APPENDIXNUMBER")
                                                                Else
                                                                    intAA = rowsAA(Count2).Item("FIGURENUMBER")
                                                                End If

                                                                If intAA = 0 Then
                                                                    wd.Selection.TypeText("Problem")
                                                                Else
                                                                    If Count2 > 0 Then
                                                                        wd.Selection.TypeText(Text:=", ")
                                                                    End If
                                                                    If boolUseHyperlinks Then
                                                                        Call HyperlinkFigures(wd, intAA, strFL)
                                                                    Else
                                                                        'just enter value
                                                                        If boolApp Then
                                                                            'need to determine if appendix is letter or number
                                                                            Dim vAppS
                                                                            Try
                                                                                Dim vNS
                                                                                vNS = wd.CaptionLabels.Item("Appendix").NumberStyle
                                                                                If InStr(1, vNS.ToString, "Letter", CompareMethod.Text) > 0 Then
                                                                                    vAppS = ChrW(64 + intAA)
                                                                                Else
                                                                                    vAppS = intAA
                                                                                End If

                                                                            Catch ex As Exception
                                                                                vAppS = ChrW(64 + intAA)
                                                                            End Try
                                                                            var1 = strFL & ChrW(160) & vAppS
                                                                        Else
                                                                            var1 = strFL & ChrW(160) & intAA
                                                                        End If

                                                                        wd.Selection.TypeText(Text:=var1.ToString)
                                                                    End If

                                                                End If
                                                            Next
                                                            'If intAA = 0 Then
                                                            '    wd.Selection.TypeText("Problem")
                                                            'Else
                                                            '    Call HyperlinkFigures(wd, intAA, strFL)
                                                            'End If
                                                        End If


                                                    Else
                                                        wd.Selection.TypeText("[App/Fig Not Included]")
                                                    End If



                                                Else
                                                    Exit Do 'KEEP THIS!!!!
                                                End If
                                                var1 = "a" 'debugging

                                                Exit Do


                                            Loop

                                            boolF1 = .Found
                                            If boolF1 Then
                                            Else
                                                Exit Do
                                            End If

                                        Catch ex As Exception
                                            boolF1 = False
                                            var1 = ex.Message
                                        End Try

                                    End With

                                Loop

                            End If
                        End If
                    End If

                    End If
next1:

            Next

            If boolChar Then 'ignore
                var1 = SearchReplace
            Else


                Try

                    'now look for any [NA]s
                    With myRng.Find
                        .ClearFormatting()
                        .Text = "[NA]"
                        With .Replacement
                            .ClearFormatting()
                            .Text = "[NA]"
                            .Font.Bold = True
                            .Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                        End With
                        .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                    End With

                    '20180327 LEE:
                    'Don't know exactly what this is searching for
                    'The replace messes with footnote/endnote references
                    'The next set of code actually does caret2 (^2)
                    'Must comment this out

                    ''now look for any ^2
                    'With myRng.Find
                    '    .ClearFormatting()
                    '    .Text = "^2"
                    '    With .Replacement
                    '        .ClearFormatting()
                    '        .Text = "2"
                    '        .Font.Superscript = True
                    '    End With
                    '    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                    'End With

                    'now look for any ^2, special carrot character
                    With myRng.Find
                        .ClearFormatting()
                        '.Text = ChrW(94) & "2"
                        .Text = "^^2"
                        With .Replacement
                            .ClearFormatting()
                            .Text = "2"
                            .Font.Superscript = True
                        End With
                        .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                    End With

                Catch ex As Exception

                End Try

            End If

        Catch ex As Exception

            str1 = "There was a problem executing the Search/Replace item: " & strFind & ". Report generation will continue."
            str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
            str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
            MsgBox(str1)

        End Try



        strM1 = "Searching " & intEnd & " of " & intEnd
        strM2 = strM & ChrW(10) & strM1
        'frmH.lblProgress.Text = strM2
        'frmH.lblProgress.Refresh()

        'frmH.lblProgress.Text = strM
        'frmH.lblProgress.Refresh()
        'frmH.pb1.Value = intEnd
        If intEnd > frmH.pb1.Maximum Then
            frmH.pb1.Maximum = intEnd + 10
            intM = intEnd + 10
        End If
        frmH.pb1.Value = intEnd
        frmH.pb1.Visible = bool1
        frmH.pb1.Maximum = intM
        frmH.pb1.Value = intV
        If bool1 Then
            frmH.pb1.Refresh()
        Else
            frmH.Refresh()
        End If

        Try
            'clear formatting again
            myRng.Find.ClearFormatting()

        Catch ex As Exception

        End Try

        ''''wdd.visible = True

    End Function

    Function RegressionEquation()

        'find regression for report
        Dim int1 As Short
        Dim int2 As Short
        Dim strRegressionType As String
        Dim str1 As String
        Dim str2 As String
        Dim Count2 As Short

        int1 = FindRow("Regression", tblWatsonAnalRefTable, "Item")
        strRegressionType = tblWatsonAnalRefTable.Rows(int1).Item(gnumAnal)
        RegressionEquation = "y = ax + b"

        str1 = "'Linear"
        For Count2 = 1 To 17
            Select Case Count2
                Case 1
                    str1 = "Linear"
                Case 2
                    str1 = "Isotope Dilution"
                Case 3
                    str1 = "Logistic"
                Case 4
                    str1 = "Quadratic"
                Case 5
                    str1 = "Hyperbolic"
                Case 6
                    str1 = "Burrows Watson"
                Case 7
                    str1 = "Powerfit"
                Case 8
                    str1 = "Logistic (Auto Estimate)"
                Case 9
                    str1 = "4/5 PL"
                Case 10
                    str1 = "Logit-Log"
                Case 11
                    str1 = "SPLINE"
                Case 12
                    str1 = "4PL"
                Case 13
                    str1 = "5PL"
                Case 14
                    str1 = "REGR"
                Case 15
                    str1 = "Log-Log Linear"
                Case 16
                    str1 = "5PL (Auto Estimate)"
                Case 17
                    str1 = "Spline (Auto Smoothed)"
            End Select
            If StrComp(strRegressionType, str1, CompareMethod.Text) = 0 Then
                Exit For
            End If
        Next

        If StrComp(strRegressionType, "Quadratic", CompareMethod.Text) = 0 Then
            RegressionEquation = "y = ax^2 + bx + c"
        ElseIf StrComp(strRegressionType, "Linear", CompareMethod.Text) = 0 Then
            RegressionEquation = "y = ax + b"
        ElseIf StrComp(strRegressionType, "Powerfit", CompareMethod.Text) = 0 Then
            RegressionEquation = "Y = b^[0]*xb^[1]"
        ElseIf InStr(1, strRegressionType, "Logistic", CompareMethod.Text) > 0 Then
            RegressionEquation = "Y = b[0] + b[1]X[1] + b[2]X[2]"
        ElseIf StrComp(strRegressionType, "4PL", CompareMethod.Text) = 0 Then
            RegressionEquation = "Y = d + ((a-d)/(1 + (x/c)^b))"
        ElseIf StrComp(strRegressionType, "5PL", CompareMethod.Text) = 0 Then
            RegressionEquation = "Y = d + ((a-d)/((1 + (x/c)^b))^g)"
        Else
            RegressionEquation = "NA"
        End If


    End Function

    Sub SearchReplaceSigs(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strFind As String, ByVal strRole As String, ByVal strType As String) ', ByVal intS, ByVal intE)

        Dim Count1 As Short
        'Dim pos1 As Int16
        'Dim pos2 As Int16
        'Dim ctTot As Int16
        Dim varReplace
        'Dim myRange as Microsoft.Office.Interop.Word.Range
        Dim intRow As Short
        Dim dg As DataGrid
        Dim ts1 As DataGridTableStyle
        Dim dv As system.data.dataview
        Dim dtbl As System.Data.DataTable
        Dim drows() As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        'Dim myRng As Microsoft.Office.Interop.Word.selection
        Dim myRng As Microsoft.Office.Interop.Word.Range
        Dim strFind1 As String
        Dim intIDtblStudies As Int64
        Dim strNA1 As String
        Dim strNA2 As String
        Dim strNA3 As String
        Dim strNA4 As String
        Dim num1 As Object
        Dim boolNum As Boolean
        Dim dt1 As Date
        Dim dt2 As Date
        Dim rowsNick() As DataRow
        Dim dgv As DataGridView
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim tbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim strF As String
        Dim strM As String
        Dim strM1 As String
        Dim strM2 As String
        Dim intEnd As Short
        Dim bool1 As Boolean
        Dim intM As Short
        Dim intV As Short
        Dim boolFound As Boolean
        Dim intCount As Short
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim boolRng As Boolean

        'myRng = wd.selection
        'myRng = rng1
        'myRng.Select()
        'mySel = wd.Selection
        ''''''wdd.visible = True

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
        wd.Selection.WholeStory()
        'myRng = wd.Selection
        mySel = wd.Selection

        strM = frmH.lblProgress.Text

        'record original values
        bool1 = frmH.pb1.Visible
        intM = frmH.pb1.Maximum
        intV = frmH.pb1.Value

        frmH.pb1.Maximum = 1
        frmH.pb1.Value = 1
        frmH.pb1.Visible = True
        frmH.pb1.Refresh()

        strM1 = "Searching signature blocks: " & ChrW(10) & ChrW(10) & strFind
        frmH.lblProgress.Text = strM1
        frmH.lblProgress.Refresh()


        'first determine if there is something to replace
        With mySel.Find
            .ClearFormatting()
            '.Text = strFind
            .Forward = True
            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
            .Execute(FindText:=strFind)

            If .Found Then
                boolFound = True
            Else
                boolFound = False
            End If

        End With

        '.selection.SetRange(Start:=0, End:=pos2)

        If boolFound Then

            'varReplace = ReturnSearch(wd, charSectionName, Count1, strFind, myRng)

            boolRng = False

            If boolRng Then 'use find range
                With myRng.Find
                    .ClearFormatting()
                    .Text = strFind
                    With .Replacement
                        .ClearFormatting()
                        If boolNum Then
                            .Text = num1
                        Else
                            .Text = varReplace
                        End If
                    End With
                    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)

                End With
            Else 'use find selection

                '''wdd.visible = True

                'move back one character space or the selected find will be skipped
                wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                With mySel.Find
                    .ClearFormatting()
                    .Text = strFind
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop ' wdFindContinue
                    With .Replacement
                        .ClearFormatting()
                        .Text = varReplace
                    End With
                    Do While .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                        'wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                        Select Case strType
                            Case "SignatureBlockTitleComp"
                                Call SignatureBlockTitleComp(wd, strRole)
                            Case "SignatureBlockTitle"
                                Call SignatureBlockTitle(wd, strRole, False)
                            Case "SignatureBlockTitleInLine"
                                Call SignatureBlockTitle(wd, strRole, True)
                            Case "SignatureBlock"
                                Call SignatureBlock(wd, strRole)
                            Case "Name"
                                Call SignatureName(wd, strRole)
                            Case "NameColumn"
                                Call SignatureNameColumn(wd, strRole)
                            Case "WSigBlock"
                                Call WSignatureBlock(wd, strRole)
                            Case "WSigBlockRole"
                                Call WSignatureBlockRole(wd, strRole)
                            Case "RefStyle_01"
                                Call SignatureRefStyle_01(wd, strRole)

                        End Select
                    Loop
                End With

            End If

        End If


    End Sub



    Sub InsertNBS(ByVal wd As Microsoft.Office.Interop.Word.Document)

        With wd
            .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .selection.Find.ClearFormatting()
            .selection.Find.Replacement.ClearFormatting()
            With .selection.Find
                .Text = " "
                .Replacement.Text = "^s"
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            .selection.Find.Execute()
            With .selection
                If .Find.Forward = True Then
                    .Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
                Else
                    .Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                End If
                .Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne)
                If .Find.Forward = True Then
                    .Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                Else
                    .Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
                End If
                .Find.Execute()
            End With
        End With

        'clear formatting again
        wd.selection.Find.ClearFormatting()


    End Sub

    Sub ReplaceDegC(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal rng1 As Microsoft.Office.Interop.Word.Range)

        Dim var8, varReplace
        Dim mySel As Microsoft.Office.Interop.Word.Range
        Dim strFind As String
        Dim boolFound As Boolean

        mySel = rng1
        var8 = ChrW(176) & "C"
        varReplace = var8
        strFind = "deg C"

        ''''wdd.visible = True

        With mySel.Find
            .ClearFormatting()
            .Text = strFind
            With .Replacement
                .ClearFormatting()
                .Text = varReplace
            End With
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        End With

        strFind = "degC"
        With mySel.Find
            .ClearFormatting()
            .Text = strFind
            With .Replacement
                .ClearFormatting()
                .Text = varReplace
            End With
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        End With

    End Sub


    Sub ReplaceHyphens(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal rng1 As Microsoft.Office.Interop.Word.Range)

        Dim var8, varReplace
        Dim mySel As Microsoft.Office.Interop.Word.Range
        Dim strFind As String
        Dim boolFound As Boolean

        'don't do anymore

        'don't do because is replacing legitimate nbh in report body
        Exit Sub

        'Wiki:  http://en.wikipedia.org/wiki/Hyphen
        'NBH: Soft hyphen. Optional. IF IS A MINUS SIGN, IT WILL NOT PRINT!!!
        '       http://www.fileformat.info/info/unicode/char/ad/index.htm
        'chrw(2011): Hard hyphen. Non-breaking-hyphen
        '       http://www.fileformat.info/info/unicode/char/2011/index.htm
        'chrw(8209): 

        'normal hypen = chrw(45)

        '20140226 Gubbs: This is causing too many problems. Merck Intervet Shawn called today and minus signs are not getting printed.
        ' will do only in Analyte Name and others described in Sub ReturnSearch
        '20150805 Larry: Nope, have to get rid of it entirely

        mySel = rng1
        var8 = NBH() 'ChrW(8209) ' NBH 'ChrW(30) 'ChrW(2011)
        varReplace = var8
        strFind = "-"

        '''wdd.visible = True

        With mySel.Find
            .ClearFormatting()
            .Text = strFind
            With .Replacement
                .ClearFormatting()
                .Text = varReplace
            End With
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
        End With

    End Sub


End Module
