Option Compare Text

Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO

    Module modStyle2

    Sub WSignatureBlock(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strRole As String)

        Dim str1 As String
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tblCP As System.Data.DataTable
        Dim var1, var2, var3, var4, var5
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strNick As String
        Dim int1 As Short
        Dim numRT

        strNick = frmH.cbxSubmittedBy.Text
        tblCP = tblContributingPersonnel

        With wd

            'efei

            numRT = WorkingPageWidth(wd)

            ''enter PI info, if applicable
            str1 = "id_tblStudies = " & id_tblStudies & " AND CHARCPROLE = '" & strRole & "'"
            dr1 = tblCP.Select(str1)
            ct1 = dr1.Length
            Dim intRows As Short
            Dim intCols As Short
            Dim numSec1 As Integer
            Dim numSec2 As Integer
            Dim numSec3 As Integer
            Dim numSec4 As Integer
            Dim numSec5 As Integer
            Dim numSec6 As Integer
            Dim numSec7 As Integer
            Dim boolNew As Boolean
            Dim intTRows As Integer

            intRows = dr1.Length
            'calculate number of table rows
            intTRows = (intRows * 2) ' + intRows - 1

            '****
            If intRows = 0 Then
                intStartTable = 1
                var3 = .Selection.Text
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                var1 = "'" & var3 & "' not configured in Contributing Personnel Table"
                .Selection.TypeText(Text:=var1)
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
            Else
                intStartTable = 2
                'Select Case intRows
                '    Case 1
                '        intCols = 2
                '        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow
                '    Case Is > 1
                '        intCols = 3
                '        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed
                'End Select

                intCols = 7
                var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow

                wrdSelection = wd.Selection()
                .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intTRows, NumColumns:= _
                intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=var1)
                '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                .Selection.Tables.Item(1).Select()
                Call GlobalTableParaFormat(wd)

                .Selection.Rows.AllowBreakAcrossPages = False
                .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop


                var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed
                .Selection.Tables.Item(1).AutoFitBehavior(var1)

                With .Selection 'remove initial borders
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                End With
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                With wd.Selection.Tables.Item(1)
                    .LeftPadding = 1 'InchesToPoints(0.02)
                    .RightPadding = 1 'InchesToPoints(0.02)
                    '.WordWrap = True
                    '.FitText = False
                End With


                numSec1 = numRT * 0.35
                numSec2 = numRT * 0.025
                numSec3 = numRT * 0.15
                numSec4 = numRT * 0.025
                numSec5 = numRT * 0.15
                numSec6 = numRT * 0.025
                numSec7 = numRT * 0.275

                .Selection.Columns.Item(1).Width = numSec1
                .Selection.Columns.Item(2).Width = numSec2
                .Selection.Columns.Item(3).Width = numSec3
                .Selection.Columns.Item(4).Width = numSec4
                .Selection.Columns.Item(5).Width = numSec5
                .Selection.Columns.Item(6).Width = numSec6
                .Selection.Columns.Item(7).Width = numSec7

                'format columns
                boolNew = False
                int1 = CInt(intRows / 2)
                Count2 = 0
                Count3 = 1

                'format column 7
                .Selection.Tables.Item(1).Cell(1, 7).Select()
                .Selection.SelectColumn()
                .Selection.ParagraphFormat.TabStops.Add(Position:=250, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                .Selection.Tables.Item(1).Cell(1, 1).Select()
                Count2 = 0
                For Count1 = 0 To intRows - 1
                    Count2 = Count2 + 1
                    'enter item in column 7
                    If IsEven(Count2) Then
                        .Selection.Tables.Item(1).Rows.Item(Count2).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                        .Selection.Tables.Item(1).Rows.Item(Count2).Height = 72 'InchesToPoints(1)

                        .Selection.Tables.Item(1).Cell(Count2, 7).Select()
                        .Selection.Font.Size = .Selection.Font.Size - 2
                        .Selection.TypeText(Text:="Other Time Zone")
                        .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                        .Selection.TypeText(Text:=vbTab)

                    Else
                        .Selection.Tables.Item(1).Cell(Count2, 7).Select()
                        .Selection.Font.Size = .Selection.Font.Size - 2
                        '.Selection.TypeText(Text:="Eastern Time Zone")
                        .Selection.TypeText(Text:=LTimeZone)
                        .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                        .Selection.TypeText(Text:=vbTab)

                    End If

                    Count2 = Count2 + 1

                    If IsEven(Count2) Then
                        .Selection.Tables.Item(1).Rows.Item(Count2).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                        .Selection.Tables.Item(1).Rows.Item(Count2).Height = 72 'InchesToPoints(1)

                        .Selection.Tables.Item(1).Cell(Count2, 7).Select()
                        .Selection.Font.Size = .Selection.Font.Size - 2
                        .Selection.TypeText(Text:="Other Time Zone")
                        .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                        .Selection.TypeText(Text:=vbTab)

                    Else
                        .Selection.Tables.Item(1).Cell(Count2, 7).Select()
                        .Selection.Font.Size = .Selection.Font.Size - 2
                        '.Selection.TypeText(Text:="Eastern Time Zone")
                        .Selection.TypeText(Text:=LTimeZone)
                        .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                        .Selection.TypeText(Text:=vbTab)

                    End If

                    var1 = dr1(Count1).Item("charCPName")
                    var2 = NZ(dr1(Count1).Item("charCPDegree"), "")
                    var5 = NZ(dr1(Count1).Item("charCPTitle"), "")
                    If Len(var2) = 0 Then
                        var3 = var1 & Chr(10) & var5
                    Else
                        var3 = var1 & ", " & var2 & Chr(10) & var5
                    End If

                    .Selection.Tables.Item(1).Cell(Count2, 1).Select()
                    .Selection.TypeText(Text:=var3)
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    .Selection.Tables.Item(1).Cell(Count2, 3).Select()
                    .Selection.Font.Size = .Selection.Font.Size - 2
                    str1 = "Date" & ChrW(10) & "(dd-mmm-yyyy)"
                    .Selection.TypeText(Text:=str1)
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    .Selection.Tables.Item(1).Cell(Count2, 5).Select()
                    .Selection.Font.Size = .Selection.Font.Size - 2
                    str1 = "Time" & ChrW(10) & "(24 hour clock)"
                    .Selection.TypeText(Text:=str1)
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle


                    If IsEven(Count2) Then
                        'Count2 = Count2 + 1
                    Else
                        'Count2 = Count2 + 1
                    End If

                Next


            End If
            '****

        End With

    End Sub

    '

    Sub SignatureRefStyle_01(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strRole As String)

        Dim str1 As String
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tblCP As System.Data.DataTable
        Dim var1, var2, var3, var4, var5
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strNick As String
        Dim int1 As Short
        Dim numRT

        strNick = frmH.cbxSubmittedBy.Text
        tblCP = tblContributingPersonnel

        'e.g.  Elvebak, L.E.

        With wd

            numRT = WorkingPageWidth(wd)

            ''enter PI info, if applicable
            str1 = "id_tblStudies = " & id_tblStudies & " AND CHARCPROLE = '" & strRole & "'"
            dr1 = tblCP.Select(str1)
            ct1 = dr1.Length
            Dim intRows As Short
            Dim intCols As Short
            Dim numSec1, numSec2
            Dim boolNew As Boolean

            intRows = dr1.Length

            If intRows = 0 Then
                var3 = .Selection.Text
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                var1 = "'" & var3 & "' not configured in Contributing Personnel Table"
                .Selection.TypeText(Text:=var1)
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
            Else
                For Count1 = 0 To intRows - 1
                    var1 = dr1(Count1).Item("charCPName")
                    var2 = NZ(dr1(Count1).Item("charCPDegree"), "")

                    If Len(var2) = 0 Then
                        var3 = var1
                    Else
                        var3 = var1 & ", " & var2
                    End If

                    If Count1 = 0 Then
                        str1 = var3
                    Else
                        str1 = str1 & "; " & var3
                    End If

                Next
                .Selection.TypeText(Text:=str1)
            End If
        End With

    End Sub

    Sub SignatureName(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strRole As String)

        Dim str1 As String
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tblCP As System.Data.DataTable
        Dim var1, var2, var3, var4, var5
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strNick As String
        Dim int1 As Short
        Dim numRT

        strNick = frmH.cbxSubmittedBy.Text
        tblCP = tblContributingPersonnel


        With wd

            numRT = WorkingPageWidth(wd)

            ''enter PI info, if applicable
            str1 = "id_tblStudies = " & id_tblStudies & " AND CHARCPROLE = '" & strRole & "'"
            dr1 = tblCP.Select(str1)
            ct1 = dr1.Length
            Dim intRows As Short
            Dim intCols As Short
            Dim numSec1, numSec2
            Dim boolNew As Boolean

            intRows = dr1.Length

            If intRows = 0 Then
                var3 = .Selection.Text
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                var1 = "'" & var3 & "' not configured in Contributing Personnel Table"
                .Selection.TypeText(Text:=var1)
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
            Else
                For Count1 = 0 To intRows - 1
                    var1 = dr1(Count1).Item("charCPName")
                    var2 = NZ(dr1(Count1).Item("charCPDegree"), "")

                    If Len(var2) = 0 Then
                        var3 = var1
                    Else
                        var3 = var1 & ", " & var2
                    End If

                    If Count1 = 0 Then
                        str1 = var3
                    Else
                        str1 = str1 & "; " & var3
                    End If

                Next
                .Selection.TypeText(Text:=str1)
            End If
        End With

    End Sub

    Sub SignatureNameColumn(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strRole As String)
        Dim str1 As String
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tblCP As System.Data.DataTable
        Dim var1, var2, var3, var4, var5
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strNick As String
        Dim int1 As Short
        Dim numRT

        strNick = frmH.cbxSubmittedBy.Text
        tblCP = tblContributingPersonnel


        With wd

            numRT = WorkingPageWidth(wd)

            ''enter PI info, if applicable
            str1 = "id_tblStudies = " & id_tblStudies & " AND CHARCPROLE = '" & strRole & "'"
            dr1 = tblCP.Select(str1)
            ct1 = dr1.Length
            Dim intRows As Short
            Dim intCols As Short
            Dim numSec1, numSec2
            Dim boolNew As Boolean

            intRows = dr1.Length

            If intRows = 0 Then
                var3 = .Selection.Text
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                var1 = "'" & var3 & "' not configured in Contributing Personnel Table"
                .Selection.TypeText(Text:=var1)
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
            Else
                For Count1 = 0 To intRows - 1
                    var1 = dr1(Count1).Item("charCPName")
                    var2 = NZ(dr1(Count1).Item("charCPDegree"), "")

                    If Len(var2) = 0 Then
                        var3 = var1
                    Else
                        var3 = var1 & ", " & var2
                    End If

                    If Count1 = 0 Then
                        str1 = var3
                    Else
                        str1 = str1 & ChrW(13) & var3
                    End If

                Next
                .Selection.TypeText(Text:=str1)
            End If

        End With

    End Sub

    Sub SignatureBlock(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strRole As String)
        Dim str1 As String
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tblCP As System.Data.DataTable
        Dim var1, var2, var3, var4, var5
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strNick As String
        Dim int1 As Short
        Dim numRT

        strNick = frmH.cbxSubmittedBy.Text
        tblCP = tblContributingPersonnel


        With wd

            numRT = WorkingPageWidth(wd)

            ''enter PI info, if applicable
            str1 = "id_tblStudies = " & id_tblStudies & " AND CHARCPROLE = '" & strRole & "'"
            dr1 = tblCP.Select(str1)
            ct1 = dr1.Length
            Dim intRows As Short
            Dim intCols As Short
            Dim numSec1, numSec2
            Dim boolNew As Boolean

            intRows = dr1.Length
            '****
            If intRows = 0 Then
                intStartTable = 1
                var3 = .Selection.Text
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                var1 = "'" & var3 & "' not configured in Contributing Personnel Table"
                .Selection.TypeText(Text:=var1)
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
            Else
                intStartTable = 2
                Select Case intRows
                    Case 1
                        intCols = 2
                        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow
                    Case Is > 1
                        intCols = 3
                        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed
                End Select
                wrdSelection = wd.Selection()
                .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=2, NumColumns:= _
                intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=var1)
                '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                .Selection.Tables.Item(1).Select()
                Call GlobalTableParaFormat(wd)

                .Selection.Rows.AllowBreakAcrossPages = False
                .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                'align rows
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop

                With .Selection 'remove initial borders
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                End With

                numSec1 = numRT * 0.45
                numSec2 = numRT * 0.1

                Select Case intRows
                    Case 1
                        Count3 = 1
                        .Selection.Tables.Item(1).Cell(Count3, 1).Select()

                        .Selection.Tables.Item(1).Rows.Item(Count3).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                        .Selection.Tables.Item(1).Rows.Item(Count3).Height = 54 '72 'InchesToPoints(1)
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=intCols)
                        .Selection.Tables.Item(1).Rows.Item(Count3 + 1).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto
                        If ct1 = 0 Then
                            var1 = "[NA]"
                            var2 = "[NA]"
                            var5 = "[NA]"
                        Else
                            var1 = dr1(0).Item("charCPName")
                            var2 = NZ(dr1(0).Item("charCPDegree"), "")
                            var5 = NZ(dr1(0).Item("charCPTitle"), "")
                        End If
                        If Len(var2) = 0 Then
                            var3 = var1 & " / Date" ' & Chr(10) & var5
                        Else
                            var3 = var1 & ", " & var2 & " / Date" ' & Chr(10) & var5
                        End If
                        .Selection.TypeText(Text:=var3)
                        '.Selection.TypeParagraph()
                        'str1 = GetAddressTitle(strNick, tblCorporateAddresses)
                        '.Selection.TypeText(Text:=str1)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)

                    Case Is > 1
                        .Selection.Columns.Item(2).Width = numSec2
                        .Selection.Columns.Item(1).Width = numSec1
                        .Selection.Columns.Item(3).Width = numSec1

                        'format columns
                        boolNew = False
                        int1 = CInt(intRows / 2)
                        Count2 = 1
                        Count3 = 1
                        .Selection.Tables.Item(1).Cell(Count3, 1).Select()
                        For Count1 = 0 To intRows - 1
                            If Count1 = 0 Or boolNew Then
                                .Selection.Tables.Item(1).Rows.Item(Count3).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                                .Selection.Tables.Item(1).Rows.Item(Count3).Height = 54 '72 'InchesToPoints(1)
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=intCols)
                                .Selection.Tables.Item(1).Rows.Item(Count3 + 1).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto
                                boolNew = False
                            End If
                            var1 = dr1(Count1).Item("charCPName")
                            var2 = NZ(dr1(Count1).Item("charCPDegree"), "")
                            var5 = NZ(dr1(Count1).Item("charCPTitle"), "")
                            If Len(var2) = 0 Then
                                var3 = var1 & " / Date" ' & Chr(10) & var5
                            Else
                                var3 = var1 & ", " & var2 & " / Date" ' & Chr(10) & var5
                            End If
                            .Selection.TypeText(Text:=var3)
                            '.Selection.TypeParagraph()
                            'str1 = GetAddressTitle(strNick, tblCorporateAddresses)
                            '.Selection.TypeText(Text:=str1)
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle
                            If Count1 = intRows - 1 Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                            ElseIf Count1 = Count2 Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=1)
                                Count2 = Count2 + 2
                                Count3 = Count3 + 2
                                boolNew = True
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                                .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                            Else
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=2)
                            End If
                        Next
                End Select

                'optimize column widths


            End If
            '****

            '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

        End With

    End Sub


    Sub SignatureBlockTitleBU(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strRole As String)

        Dim str1 As String
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tblCP As System.Data.DataTable
        Dim var1, var2, var3, var4, var5
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strNick As String
        Dim int1 As Short
        Dim numRT

        strNick = frmH.cbxSubmittedBy.Text
        tblCP = tblContributingPersonnel

        ''''wdd.visible = True

        With wd

            numRT = WorkingPageWidth(wd)

            ''enter PI info, if applicable
            str1 = "id_tblStudies = " & id_tblStudies & " AND CHARCPROLE = '" & strRole & "'"
            dr1 = tblCP.Select(str1)
            ct1 = dr1.Length
            Dim intRows As Short
            Dim intCols As Short
            Dim numSec1, numSec2
            Dim boolNew As Boolean

            intRows = dr1.Length
            '****
            If intRows = 0 Then
                intStartTable = 1
                var3 = .Selection.Text
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                var1 = "'" & var3 & "' not configured in Contributing Personnel Table"
                .Selection.TypeText(Text:=var1)
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
            Else
                intStartTable = 2
                Select Case intRows
                    Case 1
                        intCols = 2
                        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow
                    Case Is > 1
                        intCols = 3
                        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed
                End Select
                wrdSelection = wd.Selection()
                .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=2, NumColumns:= _
                intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=var1)
                '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                .Selection.Tables.Item(1).Select()
                Call GlobalTableParaFormat(wd)

                .Selection.Rows.AllowBreakAcrossPages = False
                .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop

                With .Selection 'remove initial borders
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                End With

                numSec1 = numRT * 0.45
                numSec2 = numRT * 0.1

                Select Case intRows
                    Case 1
                        Count3 = 1
                        .Selection.Tables.Item(1).Cell(Count3, 1).Select()

                        .Selection.Tables.Item(1).Rows.Item(Count3).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                        .Selection.Tables.Item(1).Rows.Item(Count3).Height = 27 '54 '72 'InchesToPoints(1)
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=intCols)
                        .Selection.Tables.Item(1).Rows.Item(Count3 + 1).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto
                        If ct1 = 0 Then
                            var1 = "[NA]"
                            var2 = "[NA]"
                            var5 = "[NA]"
                        Else
                            var1 = dr1(0).Item("charCPName")
                            var2 = NZ(dr1(0).Item("charCPDegree"), "")
                            var5 = NZ(dr1(0).Item("charCPTitle"), "")
                        End If
                        If Len(var2) = 0 Then
                            var3 = var1 & " / Date" & Chr(10) & var5
                        Else
                            var3 = var1 & ", " & var2 & " / Date" & Chr(10) & var5
                        End If
                        .Selection.TypeText(Text:=var3)
                        '.Selection.TypeParagraph()
                        'str1 = GetAddressTitle(strNick, tblCorporateAddresses)
                        '.Selection.TypeText(Text:=str1)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                        '.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                    Case Is > 1
                        .Selection.Columns.Item(2).Width = numSec2
                        .Selection.Columns.Item(1).Width = numSec1
                        .Selection.Columns.Item(3).Width = numSec1

                        'format columns
                        boolNew = False
                        int1 = CInt(intRows / 2)
                        Count2 = 1
                        Count3 = 1
                        .Selection.Tables.Item(1).Cell(Count3, 1).Select()
                        For Count1 = 0 To intRows - 1
                            If Count1 = 0 Or boolNew Then
                                .Selection.Tables.Item(1).Rows.Item(Count3).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                                .Selection.Tables.Item(1).Rows.Item(Count3).Height = 54 '72 'InchesToPoints(1)
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=intCols)
                                .Selection.Tables.Item(1).Rows.Item(Count3 + 1).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto
                                boolNew = False
                            End If
                            var1 = dr1(Count1).Item("charCPName")
                            var2 = NZ(dr1(Count1).Item("charCPDegree"), "")
                            var5 = NZ(dr1(Count1).Item("charCPTitle"), "")
                            If Len(var2) = 0 Then
                                var3 = var1 & " / Date" & Chr(10) & var5
                            Else
                                var3 = var1 & ", " & var2 & " / Date" & Chr(10) & var5
                            End If
                            .Selection.TypeText(Text:=var3)
                            '.Selection.TypeParagraph()
                            'str1 = GetAddressTitle(strNick, tblCorporateAddresses)
                            '.Selection.TypeText(Text:=str1)
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            If Count1 = intRows - 1 Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                                '.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                            ElseIf Count1 = Count2 Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=1)
                                Count2 = Count2 + 2
                                Count3 = Count3 + 2
                                boolNew = True
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                            Else
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=2)
                            End If
                        Next
                End Select

                'optimize column widths


            End If
            '****

            '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

        End With

    End Sub

    Sub SignatureBlockTitle(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strRole As String, boolInLine As Boolean)

        '20180827 LEE:

        Dim str1 As String
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tblCP As System.Data.DataTable
        Dim var1, var2, var3, var4, var5
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strNick As String
        Dim int1 As Short
        Dim numRT

        strNick = frmH.cbxSubmittedBy.Text
        tblCP = tblContributingPersonnel

        ''''wdd.visible = True

        With wd

            numRT = WorkingPageWidth(wd)

            ''enter PI info, if applicable
            str1 = "id_tblStudies = " & id_tblStudies & " AND CHARCPROLE = '" & strRole & "'"
            dr1 = tblCP.Select(str1)
            ct1 = dr1.Length
            Dim intRows As Short
            Dim intCols As Short
            Dim numSec1, numSec2
            Dim boolNew As Boolean

            intRows = dr1.Length

            'only do one row
            If intRows = 0 Then
                var1 = "[NA]"
                var2 = "[NA]"
                var5 = "[NA]"
            Else
                var1 = dr1(0).Item("charCPName")
                var2 = NZ(dr1(0).Item("charCPDegree"), "")
                var5 = NZ(dr1(0).Item("charCPTitle"), "")
            End If

            If boolInLine Then
                If Len(var2) = 0 Then
                    var3 = var1 & ", " & var5
                Else
                    var3 = var1 & ", " & var2 & ", " & var5
                End If
            Else
                If Len(var2) = 0 Then
                    var3 = var1 & ChrW(10) & var5
                Else
                    var3 = var1 & ", " & var2 & ChrW(10) & var5
                End If
            End If

            .Selection.TypeText(Text:=var3)

        End With

    End Sub

    Sub SignatureBlockTitleComp(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strRole As String)

        Dim str1 As String
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tblCP As System.Data.DataTable
        Dim tblD As System.Data.DataTable
        Dim rowsD() As DataRow
        Dim tblCA As System.Data.DataTable
        Dim rowsCA() As DataRow
        Dim intRowsCA As Short
        Dim var1, var2, var3, var4, var5
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strNick As String
        Dim int1 As Short
        Dim numRT
        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim numTblRows As Short
        Dim strComp As String

        strNick = frmH.cbxSubmittedBy.Text
        tblCP = tblContributingPersonnel
        tblD = tblData
        tblCA = tblCorporateAddresses
        numTblRows = 2

        ''get company information
        'strF1 = "ID_TBLSTUDIES = " & id_tblStudies
        'rowsD = tblD.Select(strF1)
        'var2 = rowsD(0).Item("ID_SUBMITTEDBY")
        'strF2 = "ID_TBLCOROPORATENICKNAMES = " & var2 & " AND ID_TBLADDRESSLABLES > 1 AND ID_TBLADDRESSLABLES < 5 AND CHARVALUE <> NULL"
        'rowsCA = tblCA.Select(strF2, "ID_TBLADDRESSLABLES ASC")
        'intRowsCA = rowsCA.Length
        'numTblRows = 2 'intRowsCA + 2
        'strComp = ""
        'For Count1 = 0 To intRowsCA - 1
        '    var1 = NZ(rowsCA(Count1).Item("CHARVALUE"), "NA")
        '    If Count1 = 0 Then
        '        strComp = var1
        '    Else
        '        strComp = strComp & ChrW(10) & var1
        '    End If
        'Next

        With wd

            numRT = WorkingPageWidth(wd)

            ''enter PI info, if applicable
            str1 = "id_tblStudies = " & id_tblStudies & " AND CHARCPROLE = '" & strRole & "'"
            dr1 = tblCP.Select(str1)
            ct1 = dr1.Length
            Dim intRows As Short
            Dim intCols As Short
            Dim numSec1, numSec2
            Dim boolNew As Boolean

            intRows = dr1.Length
            '****

            ''''wdd.visible = True

            If intRows = 0 Then
                intStartTable = 1
                var3 = .Selection.Text
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                var1 = "'" & var3 & "' not configured in Contributing Personnel Table"
                .Selection.TypeText(Text:=var1)
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic

            Else
                intStartTable = 2
                Select Case intRows
                    Case 1
                        intCols = 2
                        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow
                    Case Is > 1
                        intCols = 3
                        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed
                End Select
                wrdSelection = wd.Selection()
                .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=numTblRows, NumColumns:= _
                intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=var1)
                '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                .Selection.Tables.Item(1).Select()
                Call GlobalTableParaFormat(wd)

                .Selection.Rows.AllowBreakAcrossPages = False
                .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop

                With .Selection 'remove initial borders
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                End With

                numSec1 = numRT * 0.45
                numSec2 = numRT * 0.1

                Select Case intRows
                    Case 1
                        Count3 = 1
                        .Selection.Tables.Item(1).Cell(Count3, 1).Select()

                        .Selection.Tables.Item(1).Rows.Item(Count3).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                        .Selection.Tables.Item(1).Rows.Item(Count3).Height = 54 '72 'InchesToPoints(1)
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=intCols)
                        .Selection.Tables.Item(1).Rows.Item(Count3 + 1).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto
                        If ct1 = 0 Then
                            var1 = "[NA]"
                            var2 = "[NA]"
                            var5 = "[NA]"
                        Else
                            var1 = dr1(0).Item("charCPName")
                            var2 = NZ(dr1(0).Item("charCPDegree"), "")
                            var5 = NZ(dr1(0).Item("charCPTitle"), "")
                        End If
                        If Len(var2) = 0 Then
                            var3 = var1 & " / Date" & Chr(10) & var5
                        Else
                            var3 = var1 & ", " & var2 & " / Date" & Chr(10) & var5
                        End If
                        .Selection.TypeText(Text:=var3)
                        .Selection.TypeParagraph()
                        str1 = GetAddressTitle(strNick, tblCorporateAddresses)
                        .Selection.TypeText(Text:=str1)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)

                    Case Is > 1
                        .Selection.Columns.Item(1).Width = numSec1
                        .Selection.Columns.Item(2).Width = numSec2
                        .Selection.Columns.Item(3).Width = numSec1

                        'format columns
                        boolNew = False
                        int1 = CInt(intRows / 2)
                        Count2 = 1
                        Count3 = 1
                        .Selection.Tables.Item(1).Cell(Count3, 1).Select()
                        For Count1 = 0 To intRows - 1
                            If Count1 = 0 Or boolNew Then
                                .Selection.Tables.Item(1).Rows.Item(Count3).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                                .Selection.Tables.Item(1).Rows.Item(Count3).Height = 54 '72 'InchesToPoints(1)
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=intCols)
                                .Selection.Tables.Item(1).Rows.Item(Count3 + 1).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto
                                boolNew = False
                            End If
                            var1 = dr1(Count1).Item("charCPName")
                            var2 = NZ(dr1(Count1).Item("charCPDegree"), "")
                            var5 = NZ(dr1(Count1).Item("charCPTitle"), "")
                            If Len(var2) = 0 Then
                                var3 = var1 & " / Date" & Chr(10) & var5
                            Else
                                var3 = var1 & ", " & var2 & " / Date" & Chr(10) & var5
                            End If
                            .Selection.TypeText(Text:=var3)
                            .Selection.TypeParagraph()
                            str1 = GetAddressTitle(strNick, tblCorporateAddresses)
                            .Selection.TypeText(Text:=str1)
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            If Count1 = intRows - 1 Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                            ElseIf Count1 = Count2 Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=1)
                                Count2 = Count2 + 2
                                Count3 = Count3 + 2
                                boolNew = True
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                            Else
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=2)
                            End If
                        Next
                End Select

                'optimize column widths


            End If
            '****

            '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

        End With

    End Sub


    Sub TableofContents_01(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal intSource As Short)

        Dim strPath As String
        Dim strPathGuWu As String
        Dim dtbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim charS As String
        Dim boolInclude As Boolean
        Dim boolGuWu As Boolean
        Dim boolStatement As Boolean
        Dim dv As System.Data.DataView
        Dim dg As DataGrid
        Dim ct As Short
        Dim intRType As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim var1, var2, var3
        Dim rm, lm, rt
        Dim intr As Short
        Dim intc As Short
        Dim boolCont As Boolean
        Dim pos1, pos2
        Dim rngEnd As Microsoft.Office.Interop.Word.Range
        Dim rngStart As Microsoft.Office.Interop.Word.Range

        boolCont = True 'for legacy purposes
        'frmH.strErrMsg = ""
        'frmH.intErrCount = 0

        'get table heading and style
        Dim int1 As Short
        Dim int2 As Short
        Dim strTitle As String
        int1 = 135
        If boolEntireReport Then
            'prepare dv1 from tblReportStatements
            Dim tblR As System.Data.DataTable
            Dim strFR As String

            tblR = tblReportstatements
            strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
            dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
            dv.RowFilter = arrRBSColumns(2, 0)

        Else
            dv = frmH.dgvReportStatements.DataSource
        End If

        'dv = frmH.dgvReportStatements.DataSource
        int2 = FindRowDVByCol(int1, dv, "ID_TBLCONFIGBODYSECTIONS")
        strTitle = NZ(dv(int2).Item("CHARHEADINGTEXT"), "TABLE OF CONTENTS")
        Dim numSty As Short
        numSty = dv(int2).Item("NUMHEADINGLEVEL")

        'wdd.visible = True

        With wd
            wrdSelection = .Selection()

            Select Case intSource
                Case 250

                    .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=2, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    .Selection.Tables.Item(1).Select()
                    Call GlobalTableParaFormat(wd)

                    .Selection.Rows.AllowBreakAcrossPages = True

                    'remove borders
                    Call removeAllBorders(wd, False)

                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.Rows.HeadingFormat = True
                    var1 = .Selection.Font.Size
                    wrdSelection = .Selection()
                    If numSty = 0 Then
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Font.Size = 14
                    Else
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Heading " & CStr(numSty))
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    End If
                    .Selection.Font.Bold = True
                    .Selection.TypeText(Text:=strTitle)
                    .Selection.TypeParagraph()
                    .Selection.Font.Bold = False
                    wrdSelection = .Selection()
                    wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                    .Selection.TypeParagraph()
                    '.Selection.Font.Size = var1
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    rt = WorkingPageWidth(wd)
                    .Selection.ParagraphFormat.TabStops.ClearAll()
                    .Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                    .Selection.TypeText(Text:=vbTab)
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                    .Selection.TypeText(Text:="Page No.")
                    .Selection.TypeParagraph()
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone

                    .Selection.Tables.Item(1).Cell(2, 1).Select()
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

                    wrdSelection = .Selection()
                    With .ActiveDocument
                        .TablesOfContents.Add(Range:=wrdSelection.Range, RightAlignPageNumbers:= _
                          True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
                          LowerHeadingLevel:=2, IncludePageNumbers:=True, AddedStyles:="", _
                          UseHyperlinks:=True, HidePageNumbersInWeb:=True) ', UseOutlineLevels:=True)
                        ''

                        .TablesOfContents.Item(1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                        .TablesOfContents.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent
                        Try
                            '.TablesOfContents.Format = Microsoft.Office.Interop.Word.WdTocFormat.wdTOCFormal
                            '20170604 LEE: Change to template to match document styles
                            .TablesOfContents.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                        Catch ex As Exception

                        End Try

                    End With

                    .Selection.Tables.Item(1).Select()
                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                Case 255 'no table entry

                    wrdSelection = .Selection()
                    With .ActiveDocument
                        .TablesOfContents.Add(Range:=wrdSelection.Range, RightAlignPageNumbers:= _
                          True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
                          LowerHeadingLevel:=2, IncludePageNumbers:=True, AddedStyles:="", _
                          UseHyperlinks:=True, HidePageNumbersInWeb:=True) ', UseOutlineLevels:=True)
                        ''

                        .TablesOfContents.Item(1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                        .TablesOfContents.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent
                        Try
                            '.TablesOfContents.Format = Microsoft.Office.Interop.Word.WdTocFormat.wdTOCFormal
                            '20170604 LEE: Change to template to match document styles
                            .TablesOfContents.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                        Catch ex As Exception

                        End Try

                    End With

                    If .Selection.Information(Word.WdInformation.wdWithInTable) Then
                        .Selection.Tables.Item(1).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    End If

            End Select

            'record position
            EOTOC = .Selection.Start
            With wd.ActiveDocument.Bookmarks
                .Add(Range:=wrdSelection.Range, Name:="EOTOC")
                .ShowHidden = False
            End With

        End With

    End Sub

    Sub TableofTables_01(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal intSource As Short)

        Dim strPath As String
        Dim strPathGuWu As String
        Dim dtbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim charS As String
        Dim boolInclude As Boolean
        Dim boolGuWu As Boolean
        Dim boolStatement As Boolean
        Dim dv As system.data.dataview
        Dim dg As DataGrid
        Dim ct As Short
        Dim intRType As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim var1, var2, var3
        Dim rm, lm, rt
        Dim intr As Short
        Dim intc As Short
        Dim boolCont As Boolean
        Dim pos1, pos2
        Dim rngEnd As Microsoft.Office.Interop.Word.Range
        Dim rngStart As Microsoft.Office.Interop.Word.Range
        Dim rng1 As Microsoft.Office.Interop.Word.Range

        'get table heading
        Dim int1 As Short
        Dim int2 As Short
        Dim strTitle As String
        int1 = 136

        If boolEntireReport Then
            'prepare dv1 from tblReportStatements
            Dim tblR As System.Data.DataTable
            Dim strFR As String

            tblR = tblReportstatements
            strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
            dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
            dv.RowFilter = arrRBSColumns(2, 0)

        Else
            dv = frmH.dgvReportStatements.DataSource
        End If

        frmH.lblProgress.Text = "formatting Table of Tables..."
        frmH.lblProgress.Refresh()

        'dv = frmH.dgvReportStatements.DataSource
        int2 = FindRowDVByCol(int1, dv, "ID_TBLCONFIGBODYSECTIONS")
        strTitle = NZ(dv(int2).Item("CHARHEADINGTEXT"), "LIST OF SUMMARY TABLES")
        Dim numSty As Short
        numSty = dv(int2).Item("NUMHEADINGLEVEL")

        With wd
            wrdSelection = .Selection()

            Select Case intSource

                Case 247

                    .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=2, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

                    .Selection.Tables.Item(1).Select()

                    Call GlobalTableParaFormat(wd)

                    .Selection.Rows.AllowBreakAcrossPages = True

                    'remove borders
                    Call removeAllBorders(wd, False)

                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.Rows.HeadingFormat = True
                    var1 = .Selection.Font.Size
                    wrdSelection = .Selection()
                    If numSty = 0 Then
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Font.Size = 14
                    Else
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Heading " & CStr(numSty))
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    End If
                    .Selection.Font.Bold = True
                    .Selection.TypeText(Text:=strTitle)
                    .Selection.TypeParagraph()
                    .Selection.Font.Bold = False
                    wrdSelection = .Selection()
                    wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                    .Selection.TypeParagraph()

                    '.Selection.Font.Size = var1
                    .Selection.Font.Bold = False
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    rt = WorkingPageWidth(wd)
                    .Selection.ParagraphFormat.TabStops.ClearAll()
                    .Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                    .Selection.TypeText(Text:=vbTab)
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                    .Selection.TypeText(Text:="Page No.")
                    .Selection.TypeParagraph()
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone

                    .Selection.Tables.Item(1).Cell(2, 1).Select()
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

                    wrdSelection = .Selection()

                    'wdd.visible = True
                    With .ActiveDocument
                        .TablesOfFigures.Add(Range:=wrdSelection.Range, Caption:="Table", _
                          IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
                          False, UpperHeadingLevel:=1, LowerHeadingLevel:=2, _
                          IncludePageNumbers:=True, AddedStyles:="", UseHyperlinks:=True)

                        int1 = .TablesOfFigures.Count


                        .TablesOfFigures.Item(int1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                        .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent

                        ''wdd.visible = True

                        Try
                            '.TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFFormal
                            '20170604 LEE: Change to template to match document styles
                            .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                        Catch ex As Exception

                        End Try

                        '20181017 LEE:
                        'Force this bold = false
                        rng1 = .TablesOfFigures(int1).Range
                        rng1.Font.Bold = False

                        ''trying new applyguwutof format in formatindex
                        'Try
                        '    Call FormatIndex(wd, 84, 19) 'NO!! This is redundant. Do it at end of report
                        'Catch ex As Exception

                        'End Try


                    End With

                    .Selection.Tables.Item(1).Select()
                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                Case 251 'without heading table

                    With .ActiveDocument
                        '.TablesOfFigures.Add(Range:=wrdSelection.Range, Caption:="Table", _
                        '  IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
                        '  False, UpperHeadingLevel:=1, LowerHeadingLevel:=2, _
                        '  IncludePageNumbers:=True, AddedStyles:="", UseHyperlinks:=True)


                        .TablesOfFigures.Add(Range:=wrdSelection.Range, Caption:="Table", _
                 IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
                 False, UpperHeadingLevel:=1, LowerHeadingLevel:=1, _
                 IncludePageNumbers:=True, AddedStyles:="", UseHyperlinks:=True)



                        ''20180828 LEE:
                        ''No! Use style from doc
                        '.TablesOfFigures.Item(1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                        '.TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent

                        ''wdd.visible = True

                        Try
                            '.TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFFormal
                            '20170604 LEE: Change to template to match document styles
                            '20180828 LEE:
                            'No! Use style from doc
                            .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                        Catch ex As Exception

                        End Try


                        '20181017 LEE:
                        'Force this bold = false
                        int1 = .TablesOfFigures.Count
                        rng1 = .TablesOfFigures(int1).Range
                        rng1.Font.Bold = False

                        ''trying new applyguwutof format in formatindex
                        'Try
                        '    Call FormatIndex(wd, 84, 19)
                        'Catch ex As Exception

                        'End Try

                    End With

                    If .Selection.Information(Word.WdInformation.wdWithInTable) Then
                        .Selection.Tables.Item(1).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    End If

            End Select


            'record position
            EOTOT = .Selection.Start
            With wd.ActiveDocument.Bookmarks
                .Add(Range:=wrdSelection.Range, Name:="EOTOT")
                .ShowHidden = False
            End With

        End With

    End Sub

    Sub TableofFigures_01(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal intSource As Short)

        Dim strPath As String
        Dim strPathGuWu As String
        Dim dtbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim charS As String
        Dim boolInclude As Boolean
        Dim boolGuWu As Boolean
        Dim boolStatement As Boolean
        Dim dv As system.data.dataview
        Dim dg As DataGrid
        Dim ct As Short
        Dim intRType As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim var1, var2, var3
        Dim rm, lm, rt
        Dim intr As Short
        Dim intc As Short
        Dim boolCont As Boolean
        Dim pos1, pos2
        Dim rngEnd As Microsoft.Office.Interop.Word.Range
        Dim rngStart As Microsoft.Office.Interop.Word.Range
        Dim rng1 As Microsoft.Office.Interop.Word.Range

        'get table heading
        Dim int1 As Short
        Dim int2 As Short
        Dim strTitle As String
        int1 = 137
        If boolEntireReport Then
            'prepare dv1 from tblReportStatements
            Dim tblR As System.Data.DataTable
            Dim strFR As String

            tblR = tblReportstatements
            strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
            dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
            dv.RowFilter = arrRBSColumns(2, 0)

        Else
            dv = frmH.dgvReportStatements.DataSource
        End If

        frmH.lblProgress.Text = "formatting Table of Figures..."
        frmH.lblProgress.Refresh()


        'dv = frmH.dgvReportStatements.DataSource
        int2 = FindRowDVByCol(int1, dv, "ID_TBLCONFIGBODYSECTIONS")
        strTitle = NZ(dv(int2).Item("CHARHEADINGTEXT"), "LIST OF FIGURES")
        Dim numSty As Short
        numSty = dv(int2).Item("NUMHEADINGLEVEL")

        With wd
            wrdSelection = .Selection()

            Select Case intSource

                Case 248
                    .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=2, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    .Selection.Tables.Item(1).Select()
                    Call GlobalTableParaFormat(wd)

                    .Selection.Rows.AllowBreakAcrossPages = True

                    'remove borders
                    Call removeAllBorders(wd, False)

                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.Rows.HeadingFormat = True
                    var1 = .Selection.Font.Size
                    wrdSelection = .Selection()
                    If numSty = 0 Then
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Font.Size = 14
                    Else
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Heading " & CStr(numSty))
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    End If
                    .Selection.Font.Bold = True
                    .Selection.TypeText(Text:=strTitle)
                    .Selection.TypeParagraph()
                    .Selection.Font.Bold = False
                    wrdSelection = .Selection()
                    wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                    .Selection.TypeParagraph()

                    '.Selection.Font.Size = var1
                    .Selection.Font.Bold = False
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    rt = WorkingPageWidth(wd)
                    .Selection.ParagraphFormat.TabStops.ClearAll()
                    .Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                    .Selection.TypeText(Text:=vbTab)
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                    .Selection.TypeText(Text:="Page No.")
                    .Selection.TypeParagraph()
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
                    .Selection.Tables.Item(1).Cell(2, 1).Select()
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

                    wrdSelection = .Selection()
                    With .ActiveDocument
                        .TablesOfFigures.Add(Range:=wrdSelection.Range, Caption:="Figure", _
                        IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
                         False, UpperHeadingLevel:=1, LowerHeadingLevel:=2, _
                         IncludePageNumbers:=True, AddedStyles:="", UseHyperlinks:=True, _
                         HidePageNumbersInWeb:=True)
                        ''
                        int1 = .TablesOfFigures.Count

                        .TablesOfFigures.Item(int1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                        .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent
                        Try
                            '.TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFFormal
                            '20170604 LEE: Change to template to match document styles
                            .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                        Catch ex As Exception

                        End Try

                        '20181017 LEE:
                        'Force this bold = false
                        rng1 = .TablesOfFigures(int1).Range
                        rng1.Font.Bold = False

                        ''trying new applyguwutof format in formatindex
                        'Try
                        '    Call FormatIndex(wd, 84, 19)
                        'Catch ex As Exception

                        'End Try

                    End With

                    .Selection.Tables.Item(1).Select()
                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                Case 252 'without heading table

                    With .ActiveDocument
                        .TablesOfFigures.Add(Range:=wrdSelection.Range, Caption:="Figure", _
                        IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
                         False, UpperHeadingLevel:=1, LowerHeadingLevel:=2, _
                         IncludePageNumbers:=True, AddedStyles:="", UseHyperlinks:=True, _
                         HidePageNumbersInWeb:=True)
                        ''

                        int1 = .TablesOfFigures.Count

                        .TablesOfFigures.Item(int1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                        .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent
                        Try
                            '.TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFFormal
                            '20170604 LEE: Change to template to match document styles
                            .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                        Catch ex As Exception

                        End Try

                        '20181017 LEE:
                        'Force this bold = false
                        rng1 = .TablesOfFigures(int1).Range
                        rng1.Font.Bold = False

                        ''trying new applyguwutof format in formatindex
                        'Try
                        '    Call FormatIndex(wd, 84, 19)
                        'Catch ex As Exception

                        'End Try

                    End With

                    If .Selection.Information(Word.WdInformation.wdWithInTable) Then
                        .Selection.Tables.Item(1).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    End If

            End Select



            'record position
            EOTOF = .Selection.Start
            With wd.ActiveDocument.Bookmarks
                .Add(Range:=wrdSelection.Range, Name:="EOTOF")
                .ShowHidden = False
            End With

        End With

    End Sub

    Sub TableofAppendices_01(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal intSource As Short)

        Dim strPath As String
        Dim strPathGuWu As String
        Dim dtbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim charS As String
        Dim boolInclude As Boolean
        Dim boolGuWu As Boolean
        Dim boolStatement As Boolean
        Dim dv As system.data.dataview
        Dim dg As DataGrid
        Dim ct As Short
        Dim intRType As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim var1, var2, var3
        Dim rm, lm, rt
        Dim intr As Short
        Dim intc As Short
        Dim boolCont As Boolean
        Dim pos1, pos2
        Dim rngEnd As Microsoft.Office.Interop.Word.Range
        Dim rngStart As Microsoft.Office.Interop.Word.Range
        Dim boolGo As Boolean

        'get table heading
        Dim int1 As Short
        Dim int2 As Short
        Dim strTitle As String
        int1 = 138

        If boolEntireReport Then
            'prepare dv1 from tblReportStatements
            Dim tblR As System.Data.DataTable
            Dim strFR As String

            tblR = tblReportstatements
            strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
            dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
            dv.RowFilter = arrRBSColumns(2, 0)

        Else
            dv = frmH.dgvReportStatements.DataSource
        End If

        frmH.lblProgress.Text = "formatting Table of Appendices..."
        frmH.lblProgress.Refresh()

        'dv = frmH.dgvReportStatements.DataSource
        int2 = FindRowDVByCol(int1, dv, "ID_TBLCONFIGBODYSECTIONS")
        strTitle = NZ(dv(int2).Item("CHARHEADINGTEXT"), "LIST OF APPENDICES")
        Dim numSty As Short
        numSty = dv(int2).Item("NUMHEADINGLEVEL")

        With wd
            wrdSelection = .Selection()

            Select Case intSource

                Case 249

                    .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=2, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    .Selection.Tables.Item(1).Select()
                    Call GlobalTableParaFormat(wd)

                    .Selection.Rows.AllowBreakAcrossPages = True

                    'remove borders
                    Call removeAllBorders(wd, False)

                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.Rows.HeadingFormat = True
                    var1 = .Selection.Font.Size
                    wrdSelection = .Selection()
                    If numSty = 0 Then
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Font.Size = 14
                    Else
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Heading " & CStr(numSty))
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    End If
                    .Selection.Font.Bold = True
                    .Selection.TypeText(Text:=strTitle)
                    .Selection.TypeParagraph()
                    .Selection.Font.Bold = False
                    wrdSelection = .Selection()
                    wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                    .Selection.TypeParagraph()

                    '.Selection.Font.Size = var1
                    .Selection.Font.Bold = False
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    rt = WorkingPageWidth(wd)
                    .Selection.ParagraphFormat.TabStops.ClearAll()
                    .Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                    .Selection.TypeText(Text:=vbTab)
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                    .Selection.TypeText(Text:="Page No.")
                    .Selection.TypeParagraph()
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
                    .Selection.Tables.Item(1).Cell(2, 1).Select()
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

                    wrdSelection = .Selection()
                    boolGo = False
                    Do While boolGo = False
                        With .ActiveDocument

                            Try
                                .TablesOfFigures.Add(Range:=wrdSelection.Range, Caption:="Appendix", _
                                IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
                                 False, UpperHeadingLevel:=1, LowerHeadingLevel:=2, _
                                 IncludePageNumbers:=True, AddedStyles:="", UseHyperlinks:=True, _
                                 HidePageNumbersInWeb:=True)

                                boolGo = True
                                .TablesOfFigures.Item(1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                                .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent
                                Try
                                    '.TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFFormal
                                    '20170604 LEE: Change to template to match document styles
                                    .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                                Catch ex As Exception

                                End Try

                            Catch ex As Exception
                                Try
                                    .CaptionLabels.Add(Name:="Appendix")

                                Catch ex1 As Exception
                                    .selection.typetext(text:="Could not generate Table of Appendices.")
                                    boolGo = False
                                End Try
                            End Try


                            'hmm. format for TofF/A/T are all the same: "Table of Figures"
                            'trying new applyguwutof format in formatindex
                            If boolGo Then
                                'Try
                                '    Call FormatIndex(wd, 84, 19)
                                'Catch ex As Exception

                                'End Try

                            End If

                        End With

                    Loop

                    .Selection.Tables.Item(1).Select()
                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                Case 253 'without table heading

                    boolGo = False

                    Do While boolGo = False

                        With .ActiveDocument

                            Try
                                .TablesOfFigures.Add(Range:=wrdSelection.Range, Caption:="Appendix", _
                                IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
                                 False, UpperHeadingLevel:=1, LowerHeadingLevel:=2, _
                                 IncludePageNumbers:=True, AddedStyles:="", UseHyperlinks:=True, _
                                 HidePageNumbersInWeb:=True)

                                boolGo = True
                                .TablesOfFigures.Item(1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                                .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent
                                Try
                                    '.TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFFormal
                                    '20170604 LEE: Change to template to match document styles
                                    .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                                Catch ex As Exception

                                End Try

                            Catch ex As Exception
                                Try
                                    .CaptionLabels.Add(Name:="Appendix")

                                Catch ex1 As Exception
                                    .selection.typetext(text:="Could not generate Table of Appendices.")
                                    boolGo = False
                                End Try
                            End Try


                            'hmm. format for TofF/A/T are all the same: "Table of Figures"
                            'trying new applyguwutof format in formatindex
                            If boolGo Then
                                'Try
                                '    Call FormatIndex(wd, 84, 19)
                                'Catch ex As Exception

                                'End Try

                            End If

                        End With

                    Loop

                    If .Selection.Information(Word.WdInformation.wdWithInTable) Then
                        .Selection.Tables.Item(1).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    End If

            End Select


            'record position
            EOTOA = .Selection.Start
            With wd.ActiveDocument.Bookmarks
                .Add(Range:=wrdSelection.Range, Name:="EOTOA")
                .ShowHidden = False
            End With


        End With

    End Sub

    Sub TableofAttachments_01(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal intSource As Short)

        Dim strPath As String
        Dim strPathGuWu As String
        Dim dtbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim charS As String
        Dim boolInclude As Boolean
        Dim boolGuWu As Boolean
        Dim boolStatement As Boolean
        Dim dv As system.data.dataview
        Dim dg As DataGrid
        Dim ct As Short
        Dim intRType As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim var1, var2, var3
        Dim rm, lm, rt
        Dim intr As Short
        Dim intc As Short
        Dim boolCont As Boolean
        Dim pos1, pos2
        Dim rngEnd As Microsoft.Office.Interop.Word.Range
        Dim rngStart As Microsoft.Office.Interop.Word.Range
        Dim boolGo As Boolean

        'get table heading
        Dim int1 As Short
        Dim int2 As Short
        Dim strTitle As String
        int1 = 138

        If boolEntireReport Then
            'prepare dv1 from tblReportStatements
            Dim tblR As System.Data.DataTable
            Dim strFR As String

            tblR = tblReportstatements
            strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
            dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
            dv.RowFilter = arrRBSColumns(2, 0)

        Else
            dv = frmH.dgvReportStatements.DataSource
        End If

        frmH.lblProgress.Text = "formatting Table of Attachments..."
        frmH.lblProgress.Refresh()

        'dv = frmH.dgvReportStatements.DataSource
        int2 = FindRowDVByCol(int1, dv, "ID_TBLCONFIGBODYSECTIONS")
        strTitle = NZ(dv(int2).Item("CHARHEADINGTEXT"), "LIST OF ATTACHMENTS")
        If gboolDisplayAttachments Then
            strTitle = Replace(strTitle, "Appendices", "ATTACHMENTS", 1, -1, CompareMethod.Text)
        End If
        Dim numSty As Short
        numSty = dv(int2).Item("NUMHEADINGLEVEL")

        With wd

            wrdSelection = .Selection()

            Select Case intSource
                Case 242
                    .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=2, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    .Selection.Tables.Item(1).Select()
                    Call GlobalTableParaFormat(wd)

                    .Selection.Rows.AllowBreakAcrossPages = True

                    'remove borders
                    Call removeAllBorders(wd, False)

                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.Rows.HeadingFormat = True
                    var1 = .Selection.Font.Size
                    wrdSelection = .Selection()
                    If numSty = 0 Then
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Font.Size = 14
                    Else
                        wrdSelection.Style = .ActiveDocument.Styles.Item("Heading " & CStr(numSty))
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    End If
                    .Selection.Font.Bold = True
                    .Selection.TypeText(Text:=strTitle)
                    .Selection.TypeParagraph()
                    .Selection.Font.Bold = False
                    wrdSelection = .Selection()
                    wrdSelection.Style = .ActiveDocument.Styles.Item("Normal")
                    .Selection.TypeParagraph()

                    '.Selection.Font.Size = var1
                    .Selection.Font.Bold = False
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    rt = WorkingPageWidth(wd)
                    .Selection.ParagraphFormat.TabStops.ClearAll()
                    .Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                    .Selection.TypeText(Text:=vbTab)
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                    .Selection.TypeText(Text:="Page No.")
                    .Selection.TypeParagraph()
                    .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
                    .Selection.Tables.Item(1).Cell(2, 1).Select()
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

                    wrdSelection = .Selection()
                    boolGo = False
                    Do While boolGo = False
                        With .ActiveDocument

                            Try
                                .TablesOfFigures.Add(Range:=wrdSelection.Range, Caption:="Attachment", _
                                IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
                                 False, UpperHeadingLevel:=1, LowerHeadingLevel:=2, _
                                 IncludePageNumbers:=True, AddedStyles:="", UseHyperlinks:=True, _
                                 HidePageNumbersInWeb:=True)

                                boolGo = True
                                .TablesOfFigures.Item(1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                                .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent
                                Try
                                    '.TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFFormal
                                    '20170604 LEE: Change to template to match document styles
                                    .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                                Catch ex As Exception

                                End Try

                            Catch ex As Exception
                                Try
                                    .CaptionLabels.Add(Name:="Attachment")

                                Catch ex1 As Exception
                                    .selection.typetext(text:="Could not generate Table of Attachments.")
                                    boolGo = False
                                End Try
                            End Try


                            'hmm. format for TofF/A/T are all the same: "Table of Figures"
                            'trying new applyguwutof format in formatindex
                            If boolGo Then
                                'Try
                                '    Call FormatIndex(wd, 84, 19)
                                'Catch ex As Exception

                                'End Try

                            End If

                        End With

                    Loop
                    .Selection.Tables.Item(1).Select()
                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                Case 254 'no heading entry

                    wrdSelection = .Selection()
                    boolGo = False
                    Do While boolGo = False
                        With .ActiveDocument

                            Try
                                .TablesOfFigures.Add(Range:=wrdSelection.Range, Caption:="Attachment", _
                                IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
                                 False, UpperHeadingLevel:=1, LowerHeadingLevel:=2, _
                                 IncludePageNumbers:=True, AddedStyles:="", UseHyperlinks:=True, _
                                 HidePageNumbersInWeb:=True)

                                boolGo = True
                                .TablesOfFigures.Item(1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                                .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent
                                Try
                                    '.TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFFormal
                                    '20170604 LEE: Change to template to match document styles
                                    .TablesOfFigures.Format = Microsoft.Office.Interop.Word.WdTofFormat.wdTOFTemplate
                                Catch ex As Exception

                                End Try

                            Catch ex As Exception
                                Try
                                    .CaptionLabels.Add(Name:="Attachment")

                                Catch ex1 As Exception
                                    .selection.typetext(text:="Could not generate Table of Attachments.")
                                    boolGo = False
                                End Try
                            End Try


                            'hmm. format for TofF/A/T are all the same: "Table of Figures"
                            'trying new applyguwutof format in formatindex
                            If boolGo Then
                                'Try
                                '    Call FormatIndex(wd, 84, 19)
                                'Catch ex As Exception

                                'End Try

                            End If

                        End With

                    Loop

                    If .Selection.Information(Word.WdInformation.wdWithInTable) Then
                        .Selection.Tables.Item(1).Select()
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    End If

            End Select


            'record position
            EOTOA = .Selection.Start
            With wd.ActiveDocument.Bookmarks
                .Add(Range:=wrdSelection.Range, Name:="EOTOA")
                .ShowHidden = False
            End With


        End With

    End Sub

    Sub GuWuSampleReceiptStatement01(ByVal wd As Microsoft.Office.Interop.Word.Application)  'Tables

        Dim var1, var2, var3, var8, var9
        Dim pos1, pos2
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFind As String
        Dim intRows As Int16
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim num1 As Object
        Dim drows() As DataRow
        Dim strF As String
        Dim tblN As System.Data.DataTable
        Dim boolWatson As Boolean


        tblN = tblTableN

        With wd


            dgv = frmH.dgvWatsonAnalRef
            'dv = dgv.DataSource

            'search has already been performed
            ''search for [SAMPLERECEIPTTABLE1]
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'pos2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'pos1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            ''now enter table
            'mySel = wd.selection
            'strFind = "[SAMPLERECEIPTTABLE1]"
            'With mySel.Find
            '    .ClearFormatting()
            '    .Text = strFind
            '    With .Replacement
            '        .ClearFormatting()
            '        .Text = ""
            '    End With
            '    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True)
            'End With

            'reset MySel
            mySel = wd.Selection
            If frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
                dv = frmH.dgvSampleReceiptWatson.DataSource
                boolWatson = True
            Else
                dv = frmH.dgvSampleReceipt.DataSource
                boolWatson = False
            End If
            intRows = dv.Count + 1

            .ActiveDocument.Tables.Add(Range:=mySel.Range, NumRows:=intRows, NumColumns:= _
              4, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:= _
              Word.WdAutoFitBehavior.wdAutoFitWindow)
            With .Selection.Tables.Item(1)
                'If .Style <> "Table Grid" Then
                '    .Style = "Table Grid"
                'End If
                '.ApplyStyleHeadingRows = True
                '.ApplyStyleLastRow = True
                '.ApplyStyleFirstColumn = True
                '.ApplyStyleLastColumn = True
            End With
            .Selection.Tables.Item(1).Select()
            Call GlobalTableParaFormat(wd)

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            'first make first row repeat header
            .Selection.Rows.HeadingFormat = True
            .Selection.TypeText(Text:="Date Received")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Sample Count")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Storage Temperature (" & ChrW(176) & "C)")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Sample Condition")


            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

            For Count1 = 2 To intRows
                .Selection.Tables.Item(1).Cell(Count1, 2).Select()
                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            Next

            'do borders
            .Selection.Tables.Item(1).Select()
            .Selection.Rows.AllowBreakAcrossPages = False

            Call removeAllBorders(wd, False)
            .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(intRows, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            'begin entering data
            For Count2 = 2 To intRows
                For Count1 = 1 To 4
                    .Selection.Tables.Item(1).Cell(Count2, Count1).Select()
                    If boolWatson Then
                        Select Case Count1
                            Case 1
                                var1 = dv(Count2 - 2).Item("Date Received")
                                var1 = Format(var1, LDateFormat)
                            Case 2
                                var1 = dv(Count2 - 2).Item("Sample Count")
                            Case 3
                                var1 = dv(Count2 - 2).Item("Storage Temperature")
                            Case 4
                                var1 = dv(Count2 - 2).Item("Sample Condition")
                        End Select
                    Else
                        Select Case Count1
                            Case 1
                                var1 = dv(Count2 - 2).Item("dtShipmentReceived")
                                var1 = Format(var1, LDateFormat)
                            Case 2
                                var1 = dv(Count2 - 2).Item("numSampleNumber")
                            Case 3
                                var1 = dv(Count2 - 2).Item("charStorageTemp")
                            Case 4
                                var1 = dv(Count2 - 2).Item("charCondition")
                        End Select
                    End If
                    .Selection.TypeText(Text:=CStr(NZ(var1, "[NA]")))
                Next
            Next


            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'pos2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'pos1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

        End With


    End Sub

    Sub GuWuStudySample01(ByVal wd As Microsoft.Office.Interop.Word.Application) 'Tables

        Dim var1, var2, var3, var8, var9
        Dim pos1, pos2
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFind As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim num1 As Object
        Dim drows() As DataRow
        Dim strF As String
        Dim tblN As System.Data.DataTable

        tblN = tblTableN

        ctTableN = tblTableN.Rows.Count


        With wd


            intRows = ctAnalytes + 1
            dgv = frmH.dgvWatsonAnalRef
            dv = dgv.DataSource

            'SEARCH has already been performed
            'search for [STUDYSAMPLECONCENTRATIONTABLE]
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'pos2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'pos1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)


            ''now enter table
            'mySel = wd.selection
            'strFind = "[STUDYSAMPLECONCENTRATIONTABLE]"
            'With mySel.Find
            '    .ClearFormatting()
            '    .Text = strFind
            '    With .Replacement
            '        .ClearFormatting()
            '        .Text = ""
            '    End With
            '    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True)
            'End With

            'reset MySel
            mySel = wd.Selection
            intRows = ctAnalytes + 1
            .ActiveDocument.Tables.Add(Range:=mySel.Range, NumRows:=intRows, NumColumns:= _
              4, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:= _
              Word.WdAutoFitBehavior.wdAutoFitWindow)
            With .Selection.Tables.Item(1)
                'If .Style <> "Table Grid" Then
                '    .Style = "Table Grid"
                'End If
                '.ApplyStyleHeadingRows = True
                '.ApplyStyleLastRow = True
                '.ApplyStyleFirstColumn = True
                '.ApplyStyleLastColumn = True
            End With
            .Selection.Tables.Item(1).Select()
            Call GlobalTableParaFormat(wd)

            .Selection.Tables.Item(1).Cell(1, 1).Select()

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            'first make first row header repeat
            .Selection.Rows.HeadingFormat = True
            .Selection.TypeText(Text:="Analyte")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="LLOQ")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Expression")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Table #")


            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

            For Count1 = 2 To intRows
                .Selection.Tables.Item(1).Cell(Count1, 2).Select()
                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            Next

            'do borders
            .Selection.Tables.Item(1).Select()
            .Selection.Rows.AllowBreakAcrossPages = False

            Call removeAllBorders(wd, False)
            .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(intRows, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(2, 1).Select()
            'begin entering data
            For Count1 = 1 To ctAnalytes

                gstrAnal = arrAnalytes(1, Count1)
                gnumAnal = Count1

                If Count1 = 1 Then
                    var1 = Chr(13) & arrAnalytes(1, Count1)
                Else
                    var1 = arrAnalytes(1, Count1)
                End If
                .Selection.Tables.Item(1).Cell(Count1 + 1, 1).Select()
                .Selection.TypeText(Text:=var1)
                int1 = FindRowDV("LLOQ", dv)
                num1 = NZ(dv(int1).Item(Count1), 0)
                num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
                var2 = BQL() & "<" & num1 '.ToString
                .Selection.Tables.Item(1).Cell(Count1 + 1, 2).Select()
                .Selection.TypeText(Text:=CStr(num1))
                .Selection.Tables.Item(1).Cell(Count1 + 1, 3).Select()
                .Selection.TypeText(Text:=NZ(var2, ""))
                If ctTableN = 0 Then
                    var1 = "[NA]"
                Else
                    'strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TableName = 'Summary of Samples'"
                    'strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TableName = 'Summary of Samples'"
                    'drows = tblN.Select(strF)

                    'strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TableName = 'Summary of Samples'"
                    strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TABLEID = 5"
                    drows = tblN.Select(strF)
                    var9 = drows.Length
                    Try
                        If boolUseHyperlinks Then
                            var1 = "Table_" & drows(0).Item("TableNumber")
                        Else
                            var1 = "Table" & ChrW(160) & drows(0).Item("TableNumber")
                        End If

                    Catch ex As Exception
                        var1 = "[NA]"
                    End Try
                    'HEREHERE
                End If
                .Selection.Tables.Item(1).Cell(Count1 + 1, 4).Select()
                .Selection.TypeText(Text:=CStr(NZ(var1, "[NA]")))

            Next

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'pos2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'pos1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

        End With
    End Sub


    Function VerboseReplicate(ByVal num1, ByVal bool) As String
        'bool: True=Capitalize, False=Uncapitalize
        Dim str1 As String


        str1 = "[NA]"
        Select Case num1
            Case 0
                str1 = "No"
            Case 1
                str1 = "Single"
            Case 2
                str1 = "Duplicate"
            Case 3
                str1 = "Triplicate"
            Case 4
                str1 = "Quadruple"
        End Select

        If StrComp(str1, "[NA]", CompareMethod.Text) = 0 Then
            VerboseReplicate = "[NA]"
        Else
            If bool Then
                VerboseReplicate = str1
            Else
                VerboseReplicate = UnCapit(str1, False)
            End If
        End If

    End Function

    Function DoQCSECTION(ByVal wd As Microsoft.Office.Interop.Word.Application) As String
        Dim var1, var2, var3, var8, var9
        Dim pos1, pos2
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFind As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim dvDist1 As system.data.dataview
        Dim dvDist2 As system.data.dataview
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strR As String

        Dim drows() As DataRow
        Dim strF As String
        Dim tblN As System.Data.DataTable

        Dim numQCLevels As Short
        Dim numRepQC As Short
        Dim numRepDilnQC As Short
        Dim boolSame As Boolean
        Dim intRQCLevels
        Dim intRRepQC
        Dim intRRepDilnQC
        Dim tblD As New System.Data.DataTable

        'gets called from QCSECTION

        tblN = tblTableN
        numQCLevels = 3
        numRepQC = 2
        numRepDilnQC = 2
        boolSame = True

        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource
        intRRepQC = FindRowDV("# of QC Replicates", dv)
        intRQCLevels = FindRowDV("# of QC Levels", dv)
        intRRepDilnQC = FindRowDV("# of Dilution QC Replicates", dv)
        'construct a table for evaluation purposes
        For Count1 = 1 To 4
            Dim col As New DataColumn
            Select Case Count1
                Case 1
                    str1 = "Analyte"
                Case 2
                    str1 = "numRepQC"
                Case 3
                    str1 = "numQCLevels"
                Case 4
                    str1 = "numRepDilnQC"
            End Select
            col.ColumnName = str1
            tblD.Columns.Add(str1)
        Next
        'populate the table
        For Count1 = 1 To ctAnalytes

            gstrAnal = arrAnalytes(1, Count1)
            gnumAnal = Count1

            Dim row As DataRow = tblD.NewRow
            row.BeginEdit()
            For Count2 = 1 To 4
                Select Case Count2
                    Case 1
                        str1 = "Analyte"
                        str2 = arrAnalytes(1, Count1)
                    Case 2
                        str1 = "numRepQC"
                        int1 = NZ(dv(intRRepQC).Item(Count1), 0)
                    Case 3
                        str1 = "numQCLevels"
                        int1 = NZ(dv(intRQCLevels).Item(Count1), 0)
                    Case 4
                        str1 = "numRepDilnQC"
                        int1 = NZ(dv(intRRepDilnQC).Item(Count1), 0)
                End Select
                Select Case Count2
                    Case 1
                        row.Item(str1) = str2
                    Case 2, 3, 4
                        row.Item(str1) = int1
                End Select
            Next
            row.EndEdit()
            tblD.Rows.Add(row)
        Next

        intRows = ctAnalytes + 2
        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource

        If ctAnalytes = 1 Then
            'retrieve values
            numRepQC = tblD.Rows.Item(0).Item("numRepQC")
            numQCLevels = tblD.Rows.Item(0).Item("numQCLevels")

            'construct QC section
            strR = VerboseReplicate(numRepQC, True)
            str1 = strR & " quality control (QC) standards at "
            strR = VerboseNumber(numQCLevels, False)
            str1 = str1 & strR & " concentrations were included in each analytical run."
        Else
            'SELECT DISTINCT tbl1.intQCLevels, tbl1.intQCReps
            'FROM tbl1;
            dvDist1 = tblD.DefaultView
            'Dim newTable as System.Data.DataTable = view.ToTable("UniqueLastNames", True, "FirstName", "LastName")
            'first make table with only two columns
            Dim newTblTemp As System.Data.DataTable = dvDist1.ToTable("tblTemp", False, "numQCLevels", "numRepQC")
            'now do a select distinct from newTblTemp
            dvDist2 = newTblTemp.DefaultView
            Dim newTbl As System.Data.DataTable = dvDist2.ToTable("newTbl", True, "numQCLevels", "numRepQC")
            int1 = newTbl.Rows.Count
            If int1 = 0 Then
                str1 = "[NA]"
            Else
                If int1 = 1 Then 'record as normal
                    'retrieve values
                    numRepQC = tblD.Rows.Item(0).Item("numRepQC")
                    numQCLevels = tblD.Rows.Item(0).Item("numQCLevels")

                    'construct QC section
                    strR = VerboseReplicate(numRepQC, True)
                    str1 = strR & " quality control (QC) standards at "
                    strR = VerboseNumber(numQCLevels, False)
                    str1 = str1 & strR & " concentrations were included in each analytical run."
                Else
                    If int1 = 0 Then
                        str1 = "[NA]"
                    Else
                        str1 = ""
                    End If
                    For Count1 = 0 To int1 - 1
                        var1 = newTbl.Rows.Item(Count1).Item("numQCLevels")
                        var2 = newTbl.Rows.Item(Count1).Item("numRepQC")
                        strF = "numQCLevels = " & var1 & " AND numRepQC = " & var2
                        dvDist1.RowFilter = ""
                        dvDist1.RowFilter = strF
                        int2 = dvDist1.Count
                        If Count1 = 0 Then
                            strR = VerboseReplicate(var2, True)
                        Else
                            strR = VerboseReplicate(var2, False)
                        End If
                        str1 = str1 & strR & " quality control (QC) standards at "
                        strR = VerboseNumber(var1, False)
                        str1 = str1 & strR & " concentrations for " & dvDist1(0).Item("Analyte")
                        For Count2 = 0 To int2 - 1
                            str2 = dvDist1(Count2).Item("Analyte")
                            If Count2 = int2 - 1 And int2 - 1 > 2 Then
                                str1 = str1 & ", and " & str2
                            ElseIf Count2 <> int2 - 1 And int2 - 1 > 2 Then
                                str1 = str1 & ", " & str2
                            Else
                                str1 = str1 & " and " & str2
                            End If
                        Next
                        If Count1 = int1 - 1 And int1 - 1 > 2 Then
                            str1 = str1 & ", and " ' & str2
                        ElseIf Count1 <> int1 - 1 And int1 - 1 > 2 Then
                            str1 = str1 & ", " ' & str2
                        ElseIf Count1 = int1 - 1 Then
                            'nothing
                        Else
                            str1 = str1 & " and " ' & str2
                        End If
                    Next
                    str2 = " were included in each analytical run."
                    str1 = str1 & str2
                End If
            End If
        End If

        DoQCSECTION = str1

    End Function

    Function DoDILUTIONQCSECTION(ByVal wd As Microsoft.Office.Interop.Word.Application) As String

        Dim var1, var2, var3, var8, var9
        Dim pos1, pos2
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFind As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim dvDist1 As System.Data.DataView
        Dim dvDist2 As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strR As String

        Dim drows() As DataRow
        Dim strF As String
        Dim tblN As System.Data.DataTable

        Dim numQCLevels As Short
        Dim numRepQC As Short
        Dim numRepDilnQC As Short
        Dim boolSame As Boolean
        Dim intRQCLevels
        Dim intRRepQC
        Dim intRRepDilnQC
        Dim tblD As New System.Data.DataTable


        'do DILUTIONQCSECTION
        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource
        intRRepQC = FindRowDV("# of QC Replicates", dv)
        intRQCLevels = FindRowDV("# of QC Levels", dv)
        intRRepDilnQC = FindRowDV("# of Dilution QC Replicates", dv)
        'construct a table for evaluation purposes
        For Count1 = 1 To 4
            Dim col As New DataColumn
            Select Case Count1
                Case 1
                    str1 = "Analyte"
                Case 2
                    str1 = "numRepQC"
                Case 3
                    str1 = "numQCLevels"
                Case 4
                    str1 = "numRepDilnQC"
            End Select
            col.ColumnName = str1
            tblD.Columns.Add(str1)
        Next
        'populate the table
        For Count1 = 1 To ctAnalytes

            gstrAnal = arrAnalytes(1, Count1)
            gnumAnal = Count1

            Dim row As DataRow = tblD.NewRow
            row.BeginEdit()
            For Count2 = 1 To 4
                Select Case Count2
                    Case 1
                        str1 = "Analyte"
                        str2 = arrAnalytes(1, Count1)
                    Case 2
                        str1 = "numRepQC"
                        int1 = NZ(dv(intRRepQC).Item(Count1), 0)
                    Case 3
                        str1 = "numQCLevels"
                        int1 = NZ(dv(intRQCLevels).Item(Count1), 0)
                    Case 4
                        str1 = "numRepDilnQC"
                        int1 = NZ(dv(intRRepDilnQC).Item(Count1), 0)
                End Select
                Select Case Count2
                    Case 1
                        row.Item(str1) = str2
                    Case 2, 3, 4
                        row.Item(str1) = int1
                End Select
            Next
            row.EndEdit()
            tblD.Rows.Add(row)
        Next

        intRows = ctAnalytes + 2
        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource

        If ctAnalytes = 1 Then
            'construct Diln QC section
            numRepDilnQC = NZ(dv(intRRepDilnQC).Item(1), 0)
            strR = VerboseReplicate(numRepDilnQC, False)
            If StrComp(strR, "No", CompareMethod.Text) = 0 Then
                str1 = ""
            Else
                str1 = "In addition, "
                str1 = str1 & strR & " dilution QCs were included in each of the analytical runs in which dilutions were performed."
            End If

        Else
            'SELECT DISTINCT tbl1.intQCLevels, tbl1.intQCReps
            'FROM tbl1;
            'now do a select distinct from newTblTemp
            dvDist1 = tblD.DefaultView
            int2 = dvDist1.Count 'for testing
            Dim newTbl As System.Data.DataTable = dvDist1.ToTable("newTbl", True, "numRepDilnQC")
            int1 = newTbl.Rows.Count
            If int1 = 0 Then 'this means this study did not have dilution QCs. Return ""
                str1 = ""
            Else
                If int1 = 1 Then 'record as normal
                    numRepDilnQC = NZ(dv(intRRepDilnQC).Item(1), 0)
                    If numRepDilnQC = 0 Then 'this means this study did not have dilution QCs. Return ""
                        str1 = ""
                    Else
                        strR = VerboseReplicate(numRepDilnQC, False)
                        If StrComp(strR, "No", CompareMethod.Text) = 0 Then
                            str1 = ""
                        Else
                            str1 = "In addition, "
                            str1 = str1 & strR & " dilution QCs were included in each of the analytical runs in which dilutions were performed."
                        End If
                    End If
                Else
                    str1 = "In addition, "
                    For Count1 = 0 To int1 - 1
                        Count2 = 0
                        var1 = newTbl.Rows.Item(Count1).Item("numRepDilnQC")
                        If var1 = 0 Then 'skip analytes with no diln QCs
                        Else
                            strF = "numRepDilnQC = " & var1
                            dvDist1.RowFilter = ""
                            dvDist1.RowFilter = strF
                            dvDist1.Sort = "numRepDilnQC"
                            int2 = dvDist1.Count
                            strR = NZ(var1, 0) '20190228 LEE: Do not make verbose VerboseReplicate(var1, False)
                            str1 = str1 & strR & " dilution QCs for " & dvDist1(0).Item("Analyte")
                            For Count2 = 1 To int2 - 1
                                str2 = dvDist1(Count2).Item("Analyte")
                                If Count2 = int2 - 1 And int2 - 1 > 2 Then
                                    str1 = str1 & ", and " & str2
                                ElseIf Count2 <> int2 - 1 And int2 - 1 > 2 Then
                                    str1 = str1 & ", " & str2
                                Else
                                    str1 = str1 & " and " & str2
                                End If
                            Next
                            If Count1 = int1 - 1 And int1 - 1 > 2 Then
                                str1 = str1 & ", and " ' & str2
                            ElseIf Count1 <> int1 - 1 And int1 - 1 > 2 Then
                                str1 = str1 & ", " ' & str2
                            ElseIf Count1 = int1 - 1 Then
                                'nothing
                            Else
                                str1 = str1 & " and " ' & str2
                            End If
                        End If
                    Next
                    str2 = " were included in each of the analytical runs in which dilutions were performed."
                    str1 = str1 & str2
                End If
            End If
        End If

        DoDILUTIONQCSECTION = str1

    End Function

    Sub DoQCTABLE1(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim var1, var2, var3, var8, var9
        Dim pos1, pos2
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFind As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim dvDist1 As System.Data.DataView
        Dim dvDist2 As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strR As String

        Dim drows() As DataRow
        Dim strF As String
        Dim tblN As System.Data.DataTable

        Dim numQCLevels As Short
        Dim numRepQC As Short
        Dim numRepDilnQC As Short
        Dim boolSame As Boolean
        Dim intRQCLevels
        Dim intRRepQC
        Dim intRRepDilnQC
        Dim tblD As New System.Data.DataTable

        tblN = tblTableN

        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource
        intRRepQC = FindRowDV("# of QC Replicates", dv)
        intRQCLevels = FindRowDV("# of QC Levels", dv)
        intRRepDilnQC = FindRowDV("# of Dilution QC Replicates", dv)
        'construct a table for evaluation purposes
        For Count1 = 1 To 4
            Dim col As New DataColumn
            Select Case Count1
                Case 1
                    str1 = "Analyte"
                Case 2
                    str1 = "numRepQC"
                Case 3
                    str1 = "numQCLevels"
                Case 4
                    str1 = "numRepDilnQC"
            End Select
            col.ColumnName = str1
            tblD.Columns.Add(str1)
        Next

        'need to determine number of rows
        '* Check if table is to be generated for this Analyte
        intRows = 0
        For Count1 = 1 To ctAnalytes
            Dim strDo As String = arrAnalytes(1, Count1) 'record column name (Analyte Description)
            If UseAnalyteByTable(CStr(strDo), True, False) Then
                intRows = intRows + 1
            End If
        Next

        Dim intCt As Short = 0

        'populate the table
        For Count1 = 1 To ctAnalytes

            gstrAnal = arrAnalytes(1, Count1)
            gnumAnal = Count1

            If UseAnalyteByTable(CStr(gstrAnal), True, False) Then
            Else
                GoTo nextCount1A
            End If

            intCt = intCt + 1

            Dim row As DataRow = tblD.NewRow
            row.BeginEdit()
            For Count2 = 1 To 4
                Select Case Count2
                    Case 1
                        str1 = "Analyte"
                        str2 = arrAnalytes(1, Count1)
                    Case 2
                        str1 = "numRepQC"
                        int1 = NZ(dv(intRRepQC).Item(Count1), 0)
                    Case 3
                        str1 = "numQCLevels"
                        int1 = NZ(dv(intRQCLevels).Item(Count1), 0)
                    Case 4
                        str1 = "numRepDilnQC"
                        int1 = NZ(dv(intRRepDilnQC).Item(Count1), 0)
                End Select
                Select Case Count2
                    Case 1
                        row.Item(str1) = str2
                    Case 2, 3, 4
                        row.Item(str1) = int1
                End Select
            Next
            row.EndEdit()
            tblD.Rows.Add(row)

nextCount1A:

        Next

        intRows = intRows + 2
        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource


        With wd
            'reset MySel
            mySel = wd.Selection

            int1 = FindRowDV("# of Dilution QC Replicates", dv)


            numRepDilnQC = dv(int1).Item(1)
            mySel = wd.Selection
            .ActiveDocument.Tables.Add(Range:=mySel.Range, NumRows:=intRows, NumColumns:= _
              8, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:= _
              Word.WdAutoFitBehavior.wdAutoFitWindow)
            With .Selection.Tables.Item(1)
                'If .Style <> "Table Grid" Then
                '    .Style = "Table Grid"
                'End If
                '.ApplyStyleHeadingRows = True
                '.ApplyStyleLastRow = True
                '.ApplyStyleFirstColumn = True
                '.ApplyStyleLastColumn = True
            End With
            .Selection.Tables.Item(1).Select()
            Call GlobalTableParaFormat(wd)

            .Selection.Tables.Item(1).Cell(1, 1).Select()

            .Selection.Tables.Item(1).Cell(2, 1).Select()
            .Selection.TypeText(Text:="Analyte")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Min")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Max")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Table #")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=2)
            .Selection.TypeText(Text:="Min")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Max")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Table #")
            '.selection.HomeKey Unit:=Microsoft.Office.Interop.Word.wdunits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.Tables.Item(1).Cell(2, 1).Select()
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

            For Count1 = 2 To intRows
                .Selection.Tables.Item(1).Cell(Count1, 2).Select()
                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            Next

            .Selection.Tables.Item(1).Select()
            .Selection.Rows.AllowBreakAcrossPages = False

            Call removeAllBorders(wd, False)
            .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(2, 2).Select()
            '.selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend, Count:=2)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(2, 6).Select()
            '.selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend, Count:=2)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(2, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(intRows, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(3, 1).Select()
            'begin entering data

            intCt = 0
            For Count1 = 1 To ctAnalytes

                gstrAnal = arrAnalytes(1, Count1)
                gnumAnal = Count1

                If UseAnalyteByTable(CStr(gstrAnal), True, False) Then
                Else
                    GoTo nextCount1
                End If

                intCt = intCt + 1

                If intCt = 1 Then
                    var1 = Chr(13) & arrAnalytes(1, Count1)
                Else
                    var1 = arrAnalytes(1, Count1)
                End If
                .Selection.Tables.Item(1).Cell(intCt + 2, 1).Select()
                .Selection.TypeText(Text:=var1)
                int1 = FindRowDV("Calibration Levels", dv)

                For Count2 = 2 To 8
                    Select Case Count2
                        Case 2
                            int1 = FindRowDV("QC Mean Accuracy Min", dv)
                        Case 3
                            int1 = FindRowDV("QC Mean Accuracy Max", dv)
                        Case 6
                            int1 = FindRowDV("QC Precision Min", dv)
                        Case 7
                            int1 = FindRowDV("QC Precision Max", dv)
                    End Select
                    Select Case Count2
                        Case 2, 3, 6, 7
                            var1 = dv(int1).Item(arrAnalytes(1, Count1))
                        Case 4, 8
                            If ctTableN = 0 Then
                                var1 = "[NA]"
                            Else
                                strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TABLEID = 4"
                                drows = tblN.Select(strF)
                                Try
                                    If boolUseHyperlinks Then
                                        var1 = "Table_" & drows(0).Item("TableNumber")
                                    Else
                                        var1 = "Table" & ChrW(160) & drows(0).Item("TableNumber")
                                    End If

                                Catch ex As Exception
                                    var1 = "[NA]"
                                End Try
                            End If
                        Case 5
                            var1 = ""
                    End Select
                    .Selection.Tables.Item(1).Cell(intCt + 2, Count2).Select()
                    .Selection.TypeText(Text:=CStr(var1))

                Next Count2

nextCount1:

            Next Count1

            'now merge top row portions
            'start from the right, or cell numbers get screwed up
            .Selection.Tables.Item(1).Cell(1, 6).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            Try
                .Selection.Cells.Merge()
            Catch ex As Exception

            End Try
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.TypeText(Text:="Mean Precision (" & ReturnPrecLabel() & ")")


            .Selection.Tables.Item(1).Cell(1, 2).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            Try
                .Selection.Cells.Merge()
            Catch ex As Exception

            End Try
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            str1 = FindBiasDiff("QC")
            .Selection.TypeText(Text:="Mean Accuracy (%" & str1 & ")")

            'move to line below table
            Call MoveOneCellDown(wd)

        End With
    End Sub

    Sub GuWuQCAccuracyPrecision02(ByVal wd, ByVal charSectionName)

        'not used anymore 20071104

        Dim var1, var2, var3, var8, var9
        Dim pos1, pos2
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFind As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim dvDist1 As system.data.dataview
        Dim dvDist2 As system.data.dataview
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strR As String

        Dim drows() As DataRow
        Dim strF As String
        Dim tblN As System.Data.DataTable

        Dim numQCLevels As Short
        Dim numRepQC As Short
        Dim numRepDilnQC As Short
        Dim boolSame As Boolean
        Dim intRQCLevels
        Dim intRRepQC
        Dim intRRepDilnQC
        Dim tblD As New System.Data.DataTable

        tblN = tblTableN
        numQCLevels = 3
        numRepQC = 2
        numRepDilnQC = 2
        boolSame = True

        'need strFind. How about QCACCPRECVERBOSE

        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource
        intRRepQC = FindRowDV("# of QC Replicates", dv)
        intRQCLevels = FindRowDV("# of QC Levels", dv)
        intRRepDilnQC = FindRowDV("# of Dilution QC Replicates", dv)
        'construct a table for evaluation purposes
        For Count1 = 1 To 4
            Dim col As New DataColumn
            Select Case Count1
                Case 1
                    str1 = "Analyte"
                Case 2
                    str1 = "numRepQC"
                Case 3
                    str1 = "numQCLevels"
                Case 4
                    str1 = "numRepDilnQC"
            End Select
            col.ColumnName = str1
            tblD.Columns.Add(str1)
        Next
        'populate the table
        For Count1 = 1 To ctAnalytes

            gstrAnal = arrAnalytes(1, Count1)
            gnumAnal = Count1

            Dim row As DataRow = tblD.NewRow
            row.BeginEdit()
            For Count2 = 1 To 4
                Select Case Count2
                    Case 1
                        str1 = "Analyte"
                        str2 = arrAnalytes(1, Count1)
                    Case 2
                        str1 = "numRepQC"
                        int1 = NZ(dv(intRRepQC).Item(Count1), 0)
                    Case 3
                        str1 = "numQCLevels"
                        int1 = NZ(dv(intRQCLevels).Item(Count1), 0)
                    Case 4
                        str1 = "numRepDilnQC"
                        int1 = NZ(dv(intRRepDilnQC).Item(Count1), 0)
                End Select
                Select Case Count2
                    Case 1
                        row.Item(str1) = str2
                    Case 2, 3, 4
                        row.Item(str1) = int1
                End Select
            Next
            row.EndEdit()
            tblD.Rows.Add(row)
        Next

        With wd
            intRows = ctAnalytes + 2
            dgv = frmH.dgvWatsonAnalRef
            dv = dgv.DataSource

            If ctAnalytes = 1 Then
                'retrieve values
                numRepQC = tblD.Rows.Item(0).Item("numRepQC")
                numQCLevels = tblD.Rows.Item(0).Item("numQCLevels")

                'construct QC section
                strR = VerboseReplicate(numRepQC, True)
                str1 = strR & " quality control (QC) standards at "
                strR = VerboseNumber(numQCLevels, False)
                str1 = str1 & strR & " concentrations were included in each analytical run."
            Else
                'SELECT DISTINCT tbl1.intQCLevels, tbl1.intQCReps
                'FROM tbl1;
                dvDist1 = tblD.DefaultView
                'Dim newTable as System.Data.DataTable = view.ToTable("UniqueLastNames", True, "FirstName", "LastName")
                'first make table with only two columns
                Dim newTblTemp As System.Data.DataTable = dvDist1.ToTable("tblTemp", False, "numQCLevels", "numRepQC")
                'now do a select distinct from newTblTemp
                dvDist2 = newTblTemp.DefaultView
                Dim newTbl As System.Data.DataTable = dvDist2.ToTable("newTbl", True, "numQCLevels", "numRepQC")
                int1 = newTbl.Rows.Count
                If int1 = 0 Then
                    str1 = ""
                Else
                    If int1 = 1 Then 'record as normal
                        'retrieve values
                        numRepQC = tblD.Rows.Item(0).Item("numRepQC")
                        numQCLevels = tblD.Rows.Item(0).Item("numQCLevels")

                        'construct QC section
                        strR = VerboseReplicate(numRepQC, True)
                        str1 = strR & " quality control (QC) standards at "
                        strR = VerboseNumber(numQCLevels, False)
                        str1 = str1 & strR & " concentrations were included in each analytical run."
                    Else
                        str1 = ""
                        For Count1 = 0 To int1 - 1
                            var1 = newTbl.Rows.Item(Count1).Item("numQCLevels")
                            var2 = newTbl.Rows.Item(Count1).Item("numRepQC")
                            strF = "numQCLevels = " & var1 & " AND numRepQC = " & var2
                            dvDist1.RowFilter = ""
                            dvDist1.RowFilter = strF
                            int2 = dvDist1.Count
                            If Count1 = 0 Then
                                strR = VerboseReplicate(var2, True)
                            Else
                                strR = VerboseReplicate(var2, False)
                            End If
                            str1 = str1 & strR & " quality control (QC) standards at "
                            strR = VerboseNumber(var1, False)
                            str1 = str1 & strR & " concentrations for " & dvDist1(0).Item("Analyte")
                            For Count2 = 0 To int2 - 1
                                str2 = dvDist1(Count2).Item("Analyte")
                                If Count2 = int2 - 1 And int2 - 1 > 2 Then
                                    str1 = str1 & ", and " & str2
                                ElseIf Count2 <> int2 - 1 And int2 - 1 > 2 Then
                                    str1 = str1 & ", " & str2
                                Else
                                    str1 = str1 & " and " & str2
                                End If
                            Next
                            If Count1 = int1 - 1 And int1 - 1 > 2 Then
                                str1 = str1 & ", and " ' & str2
                            ElseIf Count1 <> int1 - 1 And int1 - 1 > 2 Then
                                str1 = str1 & ", " ' & str2
                            ElseIf Count1 = int1 - 1 Then
                                'nothing
                            Else
                                str1 = str1 & " and " ' & str2
                            End If
                        Next
                        str2 = " were included in each analytical run."
                        str1 = str1 & str2
                    End If
                End If
            End If

            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            pos2 = .Selection.Bookmarks.item("Temp2").Start
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            pos1 = .Selection.Bookmarks.item("Temp1").Start
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            mySel = wd.selection
            strFind = "[QCSECTION]"
            With mySel.Find
                .ClearFormatting()
                .Text = strFind
                With .Replacement
                    .ClearFormatting()
                    .Text = str1
                End With
                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True)
            End With


            'do DILUTIONQCSECTION

            If ctAnalytes = 1 Then
                'construct Diln QC section
                numRepDilnQC = NZ(dv(intRRepDilnQC).Item(1), 0)
                If numRepDilnQC = 0 Then
                    str1 = ""
                    str1 = "In addition, "
                    strR = VerboseReplicate(numRepDilnQC, False)
                    str1 = str1 & strR & " dilution QCs were included in each of the analytical runs in which dilutions were performed."
                End If
            Else
                'SELECT DISTINCT tbl1.intQCLevels, tbl1.intQCReps
                'FROM tbl1;
                'now do a select distinct from newTblTemp
                dvDist1 = tblD.DefaultView
                int2 = dvDist1.Count 'for testing
                Dim newTbl As System.Data.DataTable = dvDist1.ToTable("newTbl", True, "numRepDilnQC")
                int1 = newTbl.Rows.Count
                If int1 = 0 Then
                    str1 = ""
                Else
                    If int1 = 1 Then 'record as normal
                        numRepDilnQC = NZ(dv(intRRepDilnQC).Item(1), 0)
                        If numRepDilnQC = 0 Then
                            str1 = ""
                        Else
                            str1 = "In addition, "
                            strR = VerboseReplicate(numRepDilnQC, False)
                            str1 = str1 & strR & " dilution QCs were included in each of the analytical runs in which dilutions were performed."
                        End If
                    Else
                        str1 = "In addition, "
                        For Count1 = 0 To int1 - 1
                            Count2 = 0
                            var1 = newTbl.Rows.Item(Count1).Item("numRepDilnQC")
                            If var1 = 0 Then 'skip analytes with no diln QCs
                            Else
                                strF = "numRepDilnQC = " & var1
                                dvDist1.RowFilter = ""
                                dvDist1.RowFilter = strF
                                dvDist1.Sort = "numRepDilnQC"
                                int2 = dvDist1.Count
                                strR = VerboseReplicate(var1, False)
                                str1 = str1 & strR & " dilution QCs for " & dvDist1(0).Item("Analyte")
                                For Count2 = 1 To int2 - 1
                                    str2 = dvDist1(Count2).Item("Analyte")
                                    If Count2 = int2 - 1 And int2 - 1 > 2 Then
                                        str1 = str1 & ", and " & str2
                                    ElseIf Count2 <> int2 - 1 And int2 - 1 > 2 Then
                                        str1 = str1 & ", " & str2
                                    Else
                                        str1 = str1 & " and " & str2
                                    End If
                                Next
                                If Count1 = int1 - 1 And int1 - 1 > 2 Then
                                    str1 = str1 & ", and " ' & str2
                                ElseIf Count1 <> int1 - 1 And int1 - 1 > 2 Then
                                    str1 = str1 & ", " ' & str2
                                ElseIf Count1 = int1 - 1 Then
                                    'nothing
                                Else
                                    str1 = str1 & " and " ' & str2
                                End If
                            End If
                        Next
                        str2 = " were included in each of the analytical runs in which dilutions were performed."
                        str1 = str1 & str2
                    End If
                End If
            End If

            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            pos2 = .Selection.Bookmarks.item("Temp2").Start
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            pos1 = .Selection.Bookmarks.item("Temp1").Start
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            mySel = wd.selection
            strFind = "[DILUTIONQCSECTION]"
            With mySel.Find
                .ClearFormatting()
                .Text = strFind
                With .Replacement
                    .ClearFormatting()
                    .Text = str1
                End With
                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True)
            End With

            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            pos2 = .Selection.Bookmarks.item("Temp2").Start
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            pos1 = .Selection.Bookmarks.item("Temp1").Start
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            mySel = wd.selection
            strFind = "[DILUTIONQCSECTION]"
            With mySel.Find
                .ClearFormatting()
                .Text = strFind
                With .Replacement
                    .ClearFormatting()
                    .Text = str1
                End With
                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True)
            End With

            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            var2 = .Selection.Bookmarks.item("Temp2").Start
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            var1 = .Selection.Bookmarks.item("Temp1").Start
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)


            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
        End With

    End Sub

    Sub DoCALSTDTABLE2(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim var1, var2, var3, var8, var9
        Dim pos1, pos2
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFind As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim int1 As Short

        Dim drows() As DataRow
        Dim strF As String
        Dim tblN As System.Data.DataTable

        tblN = tblTableN

        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource

        'need to determine number of rows
        '* Check if table is to be generated for this Analyte
        intRows = 0
        For Count1 = 1 To ctAnalytes
            Dim strDo As String = arrAnalytes(1, Count1) 'record column name (Analyte Description)
            If UseAnalyteByTable(CStr(strDo), False, False) Then
                intRows = intRows + 1
            End If
        Next

        With wd

            'reset MySel
            mySel = wd.Selection
            intRows = intRows + 2
            .ActiveDocument.Tables.Add(Range:=mySel.Range, NumRows:=intRows, NumColumns:= _
              8, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:= _
              Word.WdAutoFitBehavior.wdAutoFitWindow)
            With .Selection.Tables.Item(1)
                'If .Style <> "Table Grid" Then
                '    .Style = "Table Grid"
                'End If
                '.ApplyStyleHeadingRows = True
                '.ApplyStyleLastRow = True
                '.ApplyStyleFirstColumn = True
                '.ApplyStyleLastColumn = True
            End With
            .Selection.Tables.Item(1).Select()
            Call GlobalTableParaFormat(wd)

            .Selection.Tables.Item(1).Cell(1, 1).Select()

            .Selection.Tables.Item(1).Cell(2, 1).Select()
            .Selection.TypeText(Text:="Analyte")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Min")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Max")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Table #")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=2)
            .Selection.TypeText(Text:="Min")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Max")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Table #")
            '.selection.HomeKey Unit:=Microsoft.Office.Interop.Word.wdunits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.Tables.Item(1).Cell(2, 1).Select()
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

            For Count1 = 2 To intRows
                .Selection.Tables.Item(1).Cell(Count1, 2).Select()
                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            Next

            .Selection.Tables.Item(1).Select()
            .Selection.Rows.AllowBreakAcrossPages = False

            Call removeAllBorders(wd, False)
            .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(2, 2).Select()
            '.selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend, Count:=2)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(2, 6).Select()
            '.selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend, Count:=2)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(2, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(intRows, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(3, 1).Select()
            'begin entering data
            Dim intCt As Short = 0
            For Count1 = 1 To ctAnalytes

                gstrAnal = arrAnalytes(1, Count1)
                gnumAnal = Count1

                If UseAnalyteByTable(CStr(gstrAnal), False, False) Then
                Else
                    GoTo nextcount1
                End If

                intCt = intCt + 1

                If intCt = 1 Then
                    var1 = Chr(13) & arrAnalytes(1, Count1)
                Else
                    var1 = arrAnalytes(1, Count1)
                End If
                .Selection.Tables.Item(1).Cell(intCt + 2, 1).Select()
                .Selection.TypeText(Text:=var1)
                int1 = FindRowDV("Calibration Levels", dv)
                For Count2 = 2 To 8
                    Select Case Count2
                        Case 2
                            int1 = FindRowDV("Analyte Mean Accuracy Min", dv)
                        Case 3
                            int1 = FindRowDV("Analyte Mean Accuracy Max", dv)
                        Case 6
                            int1 = FindRowDV("Analyte Precision Min", dv)
                        Case 7
                            int1 = FindRowDV("Analyte Precision Max", dv)
                    End Select
                    Select Case Count2
                        Case 2, 3, 6, 7
                            var1 = dv(int1).Item(arrAnalytes(1, Count1))
                        Case 4, 8
                            If ctTableN = 0 Then
                                var1 = "[NA]"
                            Else
                                'strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TableName = 'Summary of Back-Calculated Calibration Std Conc'"
                                strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TABLEID = 3"
                                drows = tblN.Select(strF)
                                If drows.Length = 0 Then
                                    var1 = "[NA]"
                                Else
                                    If boolUseHyperlinks Then
                                        var1 = "Table_" & drows(0).Item("TableNumber")
                                    Else
                                        var1 = "Table" & ChrW(160) & drows(0).Item("TableNumber")
                                    End If

                                End If
                            End If
                        Case 5
                            var1 = ""
                    End Select
                    .Selection.Tables.Item(1).Cell(intCt + 2, Count2).Select()
                    .Selection.TypeText(Text:=CStr(var1))

                Next Count2

nextCount1:

            Next Count1

            'now merge top row portions
            'start from the right, or cell numbers get screwed up
            .Selection.Tables.Item(1).Cell(1, 6).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            Try
                .Selection.Cells.Merge()
            Catch ex As Exception

            End Try
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.TypeText(Text:="Precision (" & ReturnPrecLabel() & ")")


            .Selection.Tables.Item(1).Cell(1, 2).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            Try
                .Selection.Cells.Merge()
            Catch ex As Exception

            End Try
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            Dim str1 As String
            str1 = FindBiasDiff("Calibr")
            .Selection.TypeText(Text:="Mean Accuracy (%" & str1 & ")")


            'move to line below table
            Call MoveOneCellDown(wd)

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'pos2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'pos1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)


            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

            ''clear formatting again
            '.selection.Find.ClearFormatting()

        End With
    End Sub

    Function FindBiasDiff(ByVal strType As String) As String

        FindBiasDiff = "Bias"

        Dim intTbl As Short

        Dim tbl1 As System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim row() As DataRow
        Dim row1() As DataRow
        Dim var1

        Dim tbl2 As System.Data.DataTable
        Dim row2() As DataRow

        tbl1 = tblReports
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        row1 = tbl1.Select(strF)
        var1 = NZ(row1(0).Item("CHARREPORTTYPE"), "Sample Analysis")

        Select Case strType
            Case "Calibr"
                intTbl = 3
            Case "QC"
                intTbl = 4
        End Select

        If intTbl = 4 Then 'check for specific table
            If InStr(1, var1, "Val", CompareMethod.Text) > 0 Then
                intTbl = 11
            End If
        End If

        tbl2 = tblReportTable
        strF = "ID_TBLCONFIGREPORTTABLES = " & intTbl & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE <> 0"
        row2 = tbl2.Select(strF)

        If row2.Length = 0 Then
            Exit Function
        End If

        Dim idTR As Int64
        idTR = row2(0).Item("ID_TBLREPORTTABLE")

        Dim tbl As System.Data.DataTable
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short

        tbl = tblTableProperties
        strF = "ID_TBLCONFIGREPORTTABLES = " & intTbl & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idTR
        strS = "ID_TBLREPORTTABLE ASC"
        row = tbl.Select(strF, strS)

        Dim intL As Short
        intL = row.Length

        If intL = 0 Then
            Exit Function
        End If

        int1 = NZ(row(0).Item("BOOLSTATSDIFF"), 0)
        int2 = NZ(row(0).Item("BOOLSTATSBIAS"), 0)
        int3 = NZ(row(0).Item("BOOLSTATSRE"), 0)
        If int1 = -1 Then
            FindBiasDiff = "Diff"
        ElseIf int2 = -1 Then
            FindBiasDiff = "Bias"
        ElseIf int3 = -1 Then
            FindBiasDiff = "RE"
        End If

    End Function

    Sub DoCALSTDTABLE1(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim var1, var2, var3, var8, var9
        Dim pos1, pos2
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFind As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim int1 As Short

        'need to determine number of rows
        '* Check if table is to be generated for this Analyte
        intRows = 0
        For Count1 = 1 To ctAnalytes
            Dim strDo As String = arrAnalytes(1, Count1) 'record column name (Analyte Description)
            If UseAnalyteByTable(CStr(strDo), False, True) Then
                intRows = intRows + 1
            End If
        Next

        With wd

            'reset mySel
            mySel = wd.Selection

            'replace selection with table

            intRows = intRows + 1
            .ActiveDocument.Tables.Add(Range:=mySel.Range, NumRows:=intRows, NumColumns:= _
              6, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:= _
              Word.WdAutoFitBehavior.wdAutoFitWindow)
            With .Selection.Tables.Item(1)
                'If .Style <> "Table Grid" Then
                '    .Style = "Table Grid"
                'End If
                '.ApplyStyleHeadingRows = True
                '.ApplyStyleLastRow = True
                '.ApplyStyleFirstColumn = True
                '.ApplyStyleLastColumn = True
            End With
            .Selection.Tables.Item(1).Select()
            Call GlobalTableParaFormat(wd)

            .Selection.Tables.Item(1).Cell(1, 1).Select()

            .Selection.TypeText(Text:="Analyte")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Calibration Points")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Weighting")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Regression")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="r-squared")
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .Selection.TypeText(Text:="Table #")
            '.selection.HomeKey Unit:=Microsoft.Office.Interop.Word.wdunits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

            For Count1 = 2 To intRows
                .Selection.Tables.Item(1).Cell(Count1, 2).Select()
                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            Next

            .Selection.Tables.Item(1).Select()
            .Selection.Rows.AllowBreakAcrossPages = False

            Call removeAllBorders(wd, False)
            .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(intRows, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .Selection.Tables.Item(1).Cell(2, 1).Select()
            'begin entering data

            Dim drows() As DataRow
            Dim strF As String
            Dim tblN As System.Data.DataTable

            tblN = tblTableN

            dgv = frmH.dgvWatsonAnalRef
            dv = dgv.DataSource

            Dim intCt As Short = 0

            For Count1 = 1 To ctAnalytes

                gstrAnal = arrAnalytes(1, Count1)
                gnumAnal = Count1

                If UseAnalyteByTable(CStr(gstrAnal), False, True) Then
                Else
                    GoTo nextcount1
                End If

                intCt = intCt + 1

                If intCt = 1 Then
                    var1 = Chr(13) & arrAnalytes(1, Count1)
                Else
                    var1 = arrAnalytes(1, Count1)
                End If
                .Selection.Tables.Item(1).Cell(intCt + 1, 1).Select()
                .Selection.TypeText(Text:=var1)
                int1 = FindRowDV("Calibration Levels", dv)
                For Count2 = 2 To 6
                    Select Case Count2
                        Case 2
                            int1 = FindRowDV("Calibration Levels", dv)
                        Case 3
                            int1 = FindRowDV("Weighting", dv)
                        Case 4
                            int1 = FindRowDV("Regression", dv)
                        Case 5
                            int1 = FindRowDV("Minimum r^2", dv)
                    End Select
                    .Selection.Tables.Item(1).Cell(intCt + 1, Count2).Select()
                    If Count2 = 5 Then
                        '8804: <=
                        '8805: >=
                        '177: +-
                        '176: degree
                        var1 = ChrW(8805) '>=
                        .Selection.TypeText(Text:=CStr(var1))
                    End If
                    var1 = dv(int1).Item(arrAnalytes(1, Count1))
                    .Selection.TypeText(Text:=var1)

                Next Count2

                .Selection.Tables.Item(1).Cell(intCt + 1, 6).Select()
                If ctTableN = 0 Then
                    .Selection.TypeText(Text:="[NA]")
                Else
                    'strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TableName = 'Summary of Regression Constants'"
                    strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TABLEID = 2"
                    drows = tblN.Select(strF)
                    '.selection.typetext(Text:=drows(0).Item("TableNumber"))
                    If drows.Length = 0 Then
                        var1 = "[NA]"
                    Else
                        If boolUseHyperlinks Then
                            var1 = "Table_" & CStr(NZ(drows(0).Item("TableNumber"), "[NA]"))
                        Else
                            var1 = "Table" & ChrW(160) & CStr(NZ(drows(0).Item("TableNumber"), "[NA]"))
                        End If

                    End If
                    .Selection.TypeText(Text:=var1)
                End If

nextCount1:

            Next

            'move to line below table
            Call MoveOneCellDown(wd)


        End With
    End Sub

    Sub GuWuCalibrationStandardAccuracyPrecision01(ByVal wd, ByVal charSectionName)
        Dim var1, var2, var3, var8, var9
        Dim pos1, pos2
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFind As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim int1 As Short

        With wd
            'NO need to search
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'pos2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'pos1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            ''search for [CALSTDTABLE1] and replace with ""
            'mySel = wd.selection
            'strFind = "[CALSTDTABLE1]"
            ''With mySel.Find
            ''    .ClearFormatting()
            ''    .Text = strFind
            ''    With .Replacement
            ''        .ClearFormatting()
            ''        .Text = ""
            ''    End With
            ''    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True)
            ''End With

            '.Selection.Find.ClearFormatting()
            '.Selection.Find.Replacement.ClearFormatting()
            'With mySel.Find
            '    .Text = strFind
            '    .Replacement.Text = ""
            '    .Forward = True
            '    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindAsk
            '    .Format = False
            '    .MatchCase = False
            '    .MatchWholeWord = False
            '    .MatchWildcards = False
            '    .MatchSoundsLike = False
            '    .MatchAllWordForms = False
            'End With
            '.Selection.Find.Execute(Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop)
            'With .Selection
            '    If .Find.Forward = True Then
            '        .Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
            '    Else
            '        .Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
            '    End If
            '    .Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne)
            '    If .Find.Forward = True Then
            '        .Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
            '    Else
            '        .Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
            '    End If
            '    .Find.Execute()
            'End With


            'reset mySel
            mySel = wd.selection

            'replace selection with table

            intRows = ctAnalytes + 1
            .ActiveDocument.Tables.Add(Range:=mySel.Range, NumRows:=intRows, NumColumns:= _
              6, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:= _
              Word.WdAutoFitBehavior.wdAutoFitWindow)
            With .Selection.Tables.item(1)
                'If .Style <> "Table Grid" Then
                '    .Style = "Table Grid"
                'End If
                '.ApplyStyleHeadingRows = True
                '.ApplyStyleLastRow = True
                '.ApplyStyleFirstColumn = True
                '.ApplyStyleLastColumn = True
            End With
            .Selection.Tables.Item(1).Select()
            Call GlobalTableParaFormat(wd)

            .Selection.Tables.Item(1).Cell(1, 1).Select()

            .selection.TypeText(Text:="Analyte")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="Calibration Points")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="Weighting")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="Regression")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="r-squared")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="Table #")
            '.selection.HomeKey Unit:=Microsoft.Office.Interop.Word.wdunits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
            .selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .selection.Tables.item(1).cell(1, 1).select()
            .selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

            For Count1 = 2 To intRows
                .selection.Tables.item(1).Cell(Count1, 2).Select()
                .selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                .selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                .selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            Next

            .selection.Tables.item(1).Select()
            .Selection.Rows.AllowBreakAcrossPages = False

            Call removeAllBorders(wd, False)
            .selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            .selection.Tables.item(1).cell(1, 1).select()
            .selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With
            With .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .selection.Tables.item(1).Cell(intRows, 1).Select()
            .selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .selection.Tables.item(1).Cell(2, 1).Select()
            'begin entering data

            Dim drows() As DataRow
            Dim strF As String
            Dim tblN As System.Data.DataTable

            tblN = tblTableN

            dgv = frmH.dgvWatsonAnalRef
            dv = dgv.DataSource
            For Count1 = 1 To ctAnalytes

                gstrAnal = arrAnalytes(1, Count1)
                gnumAnal = Count1

                If Count1 = 1 Then
                    var1 = Chr(13) & arrAnalytes(1, Count1)
                Else
                    var1 = arrAnalytes(1, Count1)
                End If
                .selection.Tables.item(1).cell(Count1 + 1, 1).select()
                .selection.typetext(Text:=var1)
                int1 = FindRowDV("Calibration Levels", dv)
                For Count2 = 2 To 6
                    Select Case Count2
                        Case 2
                            int1 = FindRowDV("Calibration Levels", dv)
                        Case 3
                            int1 = FindRowDV("Weighting", dv)
                        Case 4
                            int1 = FindRowDV("Regression", dv)
                        Case 5
                            int1 = FindRowDV("Minimum r^2", dv)
                    End Select
                    .selection.Tables.item(1).cell(Count1 + 1, Count2).select()
                    If Count2 = 5 Then
                        '8804: <=
                        '8805: >=
                        '177: +-
                        '176: degree
                        var1 = ChrW(8805) '>=
                        .selection.typetext(Text:=CStr(var1))
                    End If
                    var1 = dv(int1).Item(arrAnalytes(1, Count1))
                    .selection.typetext(Text:=var1)

                Next
                .selection.Tables.item(1).cell(Count1 + 1, 6).select()
                If ctTableN = 0 Then
                    .selection.typetext(Text:="[NA]")
                Else
                    'strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TableName = 'Summary of Regression Constants'"
                    strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TABLEID = 2"
                    drows = tblN.Select(strF)
                    '.selection.typetext(Text:=drows(0).Item("TableNumber"))
                    If drows.Length = 0 Then
                        .selection.typetext(Text:="[NA]")
                    Else
                        .selection.typetext(Text:=CStr(NZ(drows(0).Item("TableNumber"), "[NA]")))
                    End If
                End If
            Next

            'search for [CALSTDTABLE2]
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            pos2 = .Selection.Bookmarks.item("Temp2").Start
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            pos1 = .Selection.Bookmarks.item("Temp1").Start
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            mySel = wd.selection
            strFind = "[CALSTDTABLE2]"
            With mySel.Find
                .ClearFormatting()
                .Text = strFind
                With .Replacement
                    .ClearFormatting()
                    .Text = ""
                End With
                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True)
            End With
            'reset MySel
            mySel = wd.selection
            intRows = ctAnalytes + 2
            .ActiveDocument.Tables.Add(Range:=mySel.Range, NumRows:=intRows, NumColumns:= _
              8, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:= _
              Word.WdAutoFitBehavior.wdAutoFitWindow)
            With .Selection.Tables.item(1)
                'If .Style <> "Table Grid" Then
                '    .Style = "Table Grid"
                'End If
                '.ApplyStyleHeadingRows = True
                '.ApplyStyleLastRow = True
                '.ApplyStyleFirstColumn = True
                '.ApplyStyleLastColumn = True
            End With
            .Selection.Tables.Item(1).Select()
            Call GlobalTableParaFormat(wd)

            .Selection.Tables.Item(1).Cell(1, 1).Select()

            .selection.Tables.item(1).cell(2, 1).select()
            .selection.TypeText(Text:="Analyte")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="Min")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="Max")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="Table #")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=2)
            .selection.TypeText(Text:="Min")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="Max")
            .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
            .selection.TypeText(Text:="Table #")
            '.selection.HomeKey Unit:=Microsoft.Office.Interop.Word.wdunits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
            .selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .selection.Tables.item(1).cell(2, 1).select()
            .selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

            For Count1 = 2 To intRows
                .selection.Tables.item(1).Cell(Count1, 2).Select()
                .selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                .selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                .selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            Next

            .selection.Tables.item(1).Select()
            .Selection.Rows.AllowBreakAcrossPages = False

            Call removeAllBorders(wd, False)
            .selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            .selection.Tables.item(1).cell(1, 1).select()
            .selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .selection.Tables.item(1).cell(2, 2).select()
            '.selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend, Count:=2)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            With .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .selection.Tables.item(1).cell(2, 6).select()
            '.selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend, Count:=2)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            With .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .selection.Tables.item(1).cell(2, 1).select()
            .selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .selection.Tables.item(1).Cell(intRows, 1).Select()
            .selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            With .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = wd.Options.DefaultBorderLineStyle
                .LineWidth = wd.Options.DefaultBorderLineWidth
                .Color = wd.Options.DefaultBorderColor
            End With

            .selection.Tables.item(1).Cell(3, 1).Select()
            'begin entering data
            For Count1 = 1 To ctAnalytes

                gstrAnal = arrAnalytes(1, Count1)
                gnumAnal = Count1

                If Count1 = 1 Then
                    var1 = Chr(13) & arrAnalytes(1, Count1)
                Else
                    var1 = arrAnalytes(1, Count1)
                End If
                .selection.Tables.item(1).cell(Count1 + 2, 1).select()
                .selection.typetext(Text:=var1)
                int1 = FindRowDV("Calibration Levels", dv)
                For Count2 = 2 To 8
                    Select Case Count2
                        Case 2
                            int1 = FindRowDV("Analyte Mean Accuracy Min", dv)
                        Case 3
                            int1 = FindRowDV("Analyte Mean Accuracy Max", dv)
                        Case 6
                            int1 = FindRowDV("Analyte Precision Min", dv)
                        Case 7
                            int1 = FindRowDV("Analyte Precision Max", dv)
                    End Select
                    Select Case Count2
                        Case 2, 3, 6, 7
                            var1 = dv(int1).Item(arrAnalytes(1, Count1))
                        Case 4, 8
                            If ctTableN = 0 Then
                                var1 = "[NA]"
                            Else
                                'strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TableName = 'Summary of Back-Calculated Calibration Std Conc'"
                                strF = "AnalyteName = '" & arrAnalytes(1, Count1) & "' AND TABLEID = 3"
                                drows = tblN.Select(strF)
                                If drows.Length = 0 Then
                                    var1 = "[NA]"
                                Else
                                    var1 = drows(0).Item("TableNumber")
                                End If
                            End If
                        Case 5
                            var1 = ""
                    End Select
                    .selection.Tables.item(1).cell(Count1 + 2, Count2).select()
                    .selection.typetext(Text:=CStr(var1))

                Next
            Next

            'now merge top row portions
            'start from the right, or cell numbers get screwed up
            .selection.Tables.item(1).cell(1, 6).select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            Try
                .Selection.Cells.Merge()
            Catch ex As Exception

            End Try
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.TypeText(Text:="Precision (" & ReturnPrecLabel() & ")")


            .selection.Tables.item(1).cell(1, 2).select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            Try
                .Selection.Cells.Merge()
            Catch ex As Exception

            End Try
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            Dim str1 As String
            str1 = FindBiasDiff("Calibr")
            .Selection.TypeText(Text:="Mean Accuracy (%" & str1 & ")")

            '.Selection.TypeText(Text:="Mean Accuracy (%Bias)")

            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            pos2 = .Selection.Bookmarks.item("Temp2").Start
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            pos1 = .Selection.Bookmarks.item("Temp1").Start
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)


            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

            'clear formatting again
            .selection.Find.ClearFormatting()

        End With
    End Sub

    Sub ABSTRACTANALYTEINFO2(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim var1, var2, var3, var8, var9
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim intRows As Short
        'Dim strF As Short
        'Dim strS As Short
        Dim strRepl As String
        Dim Count1 As Short
        Dim strAnal As String
        Dim strFrag As String

        tbl1 = tblAnalytesHome

        Dim strF As String
        Dim strS As String
        strF = "IsIntStd = 'No'"
        strS = "AnalyteDescription ASC"
        rows1 = tbl1.Select(strF, strS)
        intRows = rows1.Length

        strFrag = "in [SAMPLESIZE]" & ChrW(160) & "[SAMPLESIZEUNITS] [LC_ANTICOAGULANT] buffered [LC_SPECIES] [LC_MATRIX] samples over a concentration range of"

        With wd
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'var2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'var1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            'search for [ABSTRACTANALYTEINFO] and replace with space

            For Count1 = 0 To intRows - 1
                'fill in  [LLOQ] [LLOQUNITS] to [ULOQ] [ULOQUNITS] for [ANALYTE]
                '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
                'var2 = .Selection.Bookmarks.item("Temp2").Start
                '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
                'var1 = .Selection.Bookmarks.item("Temp1").Start
                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                Try
                    '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
                    'var2 = .Selection.Bookmarks.Item("Temp2").Start
                    '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
                    'var1 = .Selection.Bookmarks.Item("Temp1").Start
                    '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

                Catch ex As Exception

                End Try
                With .Selection.Find
                    .ClearFormatting()
                    .Text = "[ABSTRACTANALYTEINFO2]"
                    With .Replacement
                        .ClearFormatting()
                        If Count1 <= intRows - 2 Then
                            If intRows > 2 Then
                                .Text = "[ANALYTE1] " & strFrag & " [LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS], [ABSTRACTANALYTEINFO2]"
                            Else
                                .Text = "[ANALYTE1] " & strFrag & " [LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS] and [ABSTRACTANALYTEINFO2]"
                            End If
                        ElseIf Count1 < intRows - 1 Then
                            If intRows > 2 Then
                                .Text = "[ANALYTE1] " & strFrag & " [LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS], [ABSTRACTANALYTEINFO2]"
                            Else
                                .Text = "[ANALYTE1] " & strFrag & " [LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS], and [ABSTRACTANALYTEINFO2]"
                            End If
                        ElseIf Count1 = intRows - 1 Then
                            .Text = "[ANALYTE1] " & strFrag & " [LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS]"
                        End If

                    End With
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .Format = True
                    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                End With
                strAnal = rows1(Count1).Item("AnalyteDescription")
                Call SearchReplaceAnalLLOQ(wd, strAnal, True)
            Next

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            Try
                .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            Catch ex As Exception

            End Try

        End With
    End Sub

    Sub GuWuObjective01_27(ByVal wd) ', ByVal charSectionName)
        Dim var1, var2, var3, var8, var9
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim intRows As Short
        'Dim strF As Short
        'Dim strS As Short
        Dim strRepl As String
        Dim Count1 As Short
        Dim strAnal As String
        Dim strAnal1 As String

        tbl1 = tblAnalytesHome

        Dim strF As String
        Dim strS As String
        strF = "IsIntStd = 'No'"
        strS = "AnalyteDescription ASC"
        rows1 = tbl1.Select(strF, strS)
        intRows = rows1.Length

        With wd
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'var2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'var1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            'search for [ABSTRACTANALYTEINFO] and replace with space

            For Count1 = 0 To intRows - 1

                'check for use
                strAnal = rows1(Count1).Item("AnalyteDescription")
                '20181015 LEE:
                'put nbh back in
                strAnal = Replace(strAnal, "-", NBHReal, 1, -1, CompareMethod.Text)
                If UseAnalyte(strAnal) Then
                Else
                    GoTo nextCount1
                End If

                'get non-C1 name
                strAnal1 = rows1(Count1).Item("OriginalAnalyteDescription")
                '20181015 LEE:
                'put nbh back in
                strAnal1 = Replace(strAnal1, "-", NBHReal, 1, -1, CompareMethod.Text)

                '20181015 LEE:
                'put nbh back in
                strAnal1 = Replace(strAnal1, "-", NBHReal, 1, -1)

                'fill in  [LLOQ] [LLOQUNITS] to [ULOQ] [ULOQUNITS] for [ANALYTE]
                '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
                'var2 = .Selection.Bookmarks.item("Temp2").Start
                '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
                'var1 = .Selection.Bookmarks.item("Temp1").Start
                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                Try
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
                    var2 = .Selection.Bookmarks.item("Temp2").Start
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
                    var1 = .Selection.Bookmarks.item("Temp1").Start
                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

                Catch ex As Exception

                End Try
                With .selection.Find
                    .ClearFormatting()
                    .Text = "[ABSTRACTANALYTEINFO1]"
                    With .Replacement
                        .ClearFormatting()
                        If Count1 <= intRows - 2 Then
                            If intRows > 2 Then
                                .Text = "[ANALYTE1] from [LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS], [ABSTRACTANALYTEINFO1]"
                            Else
                                .Text = "[ANALYTE1] from [LLOQ" & ChrW(160) & "[LLOQUNITS] to [ULOQ" & ChrW(160) & "[ULOQUNITS] and [ABSTRACTANALYTEINFO1]"
                            End If
                        ElseIf Count1 < intRows - 1 Then
                            If intRows > 2 Then
                                .Text = "[ANALYTE1] from [LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ" & ChrW(160) & "[ULOQUNITS], [ABSTRACTANALYTEINFO1]"
                            Else
                                .Text = "[ANALYTE1] from [LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS], and [ABSTRACTANALYTEINFO1]"
                            End If
                        ElseIf Count1 = intRows - 1 Then
                            .Text = "[ANALYTE1] from [LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS]"
                        End If

                    End With
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .Format = True
                    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                End With

                Call SearchReplaceAnalLLOQ(wd, strAnal1, False)

nextCount1:

            Next Count1

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            Try
                .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            Catch ex As Exception

            End Try

        End With
    End Sub

    Sub GuWuAbstract01_26(ByVal wd) ', ByVal charSectionName)

        Dim var1, var2, var3, var8, var9
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim intRows As Short
        'Dim strF As Short
        'Dim strS As Short
        Dim strRepl As String
        Dim Count1 As Short
        Dim gs As DataGridColumnStyle
        Dim int1 As Short
        Dim str1 As String
        Dim strAnal As String

        tbl1 = tblAnalytesHome

        Dim strF As String
        Dim strS As String
        'strF = "id_tblStudies = " & id_tblStudies & " AND BOOLISREPLICATE = 0 AND BOOLISCOADMINISTERED = 0 AND BOOLINCLUDE = -1 AND BOOLISINTSTD = 0"
        strF = "IsIntStd = 'No'"
        strS = "AnalyteDescription ASC"
        rows1 = tbl1.Select(strF, strS)
        intRows = rows1.Length

        '''''''''''wdd.visible = True

        Dim intGroup As Short

        With wd
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'var2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'var1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            'search for [ABSTRACTANALYTEINFO] and replace with space

            '''''''''wdd.visible = True

            For Count1 = 0 To intRows - 1
                strAnal = rows1(Count1).Item("AnalyteDescription")

                '20181015 LEE:
                'put nbh back in
                strAnal = Replace(strAnal, "-", NBHReal, 1, -1, CompareMethod.Text)

                '20181108 LEE:
                'get useranalyte
                intGroup = rows1(Count1).Item("INTGROUP")
                strAnal = GetUserAnalyteName(strAnal, False, intGroup)

                'fill in  [LLOQ] [LLOQUNITS] to [ULOQ] [ULOQUNITS] for [ANALYTE]
                '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
                'var2 = .Selection.Bookmarks.item("Temp2").Start
                '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
                'var1 = .Selection.Bookmarks.item("Temp1").Start
                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                With .selection.Find
                    .ClearFormatting()
                    .Text = "[ABSTRACTANALYTEINFO]"
                    With .Replacement
                        .ClearFormatting()
                        If Count1 <= intRows - 2 Then
                            If intRows > 2 Then
                                .Text = "[LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS] for [ANALYTE1], [ABSTRACTANALYTEINFO]"
                            Else
                                .Text = "[LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS] for [ANALYTE1] and [ABSTRACTANALYTEINFO]"
                            End If
                        ElseIf Count1 < intRows - 1 Then
                            If intRows > 2 Then
                                .Text = "[LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS] for [ANALYTE1], [ABSTRACTANALYTEINFO]"
                            Else
                                .Text = "[LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS] for [ANALYTE1], and [ABSTRACTANALYTEINFO]"
                            End If
                        ElseIf Count1 = intRows - 1 Then
                            .Text = "[LLOQ]" & ChrW(160) & "[LLOQUNITS] to [ULOQ]" & ChrW(160) & "[ULOQUNITS] for [ANALYTE1]"
                        End If

                    End With
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .Format = True
                    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                End With

                Call SearchReplaceAnalLLOQ(wd, strAnal, False)
                var1 = intRows 'debugging

            Next

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

        End With

    End Sub

    Sub GuWuMethodSummaryStatement01(ByVal wd)

        Dim var1, var2, var3, var8, var9
        Dim bm As Microsoft.Office.Interop.Word.Bookmark
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim pos1, pos2, pos3, pos4
        Dim arrBM(2, 0) As String '1=leading bookmark, 2=trailingbookmark
        Dim drows() As DataRow
        Dim drowsA() As DataRow
        Dim str1 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim str2 As String
        Dim gs As DataGridColumnStyle
        Dim str3 As String
        Dim strAnal As String
        Dim tblAnal As System.Data.DataTable
        Dim rowsAnal() As DataRow
        Dim tblAF As System.Data.DataTable
        Dim rowsAF() As DataRow

        Dim dtbl1 As System.Data.DataTable
        Dim strS1 As String
        Dim strF1 As String
        Dim rows1() As DataRow
        Dim boolMult As Boolean

        'need a strFind. How about MethodSummaryStatement

        'if there are more than one LM, then do this
        dtbl1 = tblMethodValidationData
        strF1 = "ID_TBLSTUDIES = " & id_tblStudies
        strS1 = "ID_TBLREPORTS ASC"
        rows1 = dtbl1.Select(strF1)
        Dim dv1 As system.data.dataview = New DataView(dtbl1, strF1, strS1, DataViewRowState.CurrentRows)
        Dim dtbl2 As System.Data.DataTable = dv1.ToTable("a", True, "CHARLMNUMBER")
        If dtbl2.Rows.Count = 1 Then 'donot continue
            Exit Sub
        End If

        boolMult = True

        str1 = "AppendixName = 'LM'"
        drows = tblAppendix.Select(str1)

        tblAF = tblAppFigs
        str1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARTYPE = 'RC'"
        rowsAF = tblAF.Select(str1)

        With wd
            'temp1 and temp2 come from paste events
            'they aren't relevent here
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'pos2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'pos1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            'intitialize temp1 and temp2 for later code
            wrdSelection = wd.Selection()
            With .ActiveDocument.Bookmarks
                .Add(Range:=wrdSelection.Range, Name:="Temp2")
                .ShowHidden = False
            End With
            pos2 = .Selection.Start

            'goto previous paragraph return
            .Selection.Find.ClearFormatting()
            With .Selection.Find
                .Text = "^p"
                .Replacement.Text = ""
                .Forward = False
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindAsk
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            .Selection.Find.Execute()
            'move right one
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
            'clear formatting
            .Selection.Find.ClearFormatting()

            wrdSelection = wd.Selection()
            With .ActiveDocument.Bookmarks
                .Add(Range:=wrdSelection.Range, Name:="Temp1")
                .ShowHidden = False
            End With
            pos1 = .Selection.Start

            'creation of LM will be based on appendix section
            'determine if there are more than one LMs to describe
            Dim dgv As DataGridView
            Dim dv As system.data.dataview
            'Dim dv1 as system.data.dataview
            Dim intCols As Short
            Dim Count1 As Short
            Dim ctBM As Short

            str1 = "AppendixName = 'LM'"
            drowsA = tblAppendix.Select(str1)

            dgv = frmH.dgvMethodValData
            dv = dgv.DataSource
            'intCols = dg.TableStyles(0).GridColumnStyles.Count
            intCols = dgv.Columns.Count
            'if intcols > 2, then there are multiple methods
            'actually boolMulti finds it

            Dim strF As String
            Dim strS As String
            Dim intRows As Short
            strF = "IsIntStd = 'No'"
            strS = "AnalyteDescription ASC"
            'rowsAnal = tblAnal.Select(strF, strS)
            tblAnal = tblAnalytesHome
            rowsAnal = tblAnal.Select(strF, strS)
            intRows = rowsAnal.Length

            ReDim arrBM(2, intCols)
            ctBM = 0
            'need to paste additional blocks
            'save Temp2 bookmark
            'bm = wd.ActiveDocument.Bookmarks.item("Temp2")

            Dim intGroup As Short

            ''
            If intRows = 1 Then 'if you've got this far, then introws > 1
                arrBM(1, 0) = "Temp2"
                'arrBM(2, 0) = "Temp2"
            Else
                For Count1 = 0 To intRows - 2

                    strAnal = rowsAnal(Count1).Item("AnalyteDescription")

                    '20181108 LEE:
                    'get useranalyte
                    intGroup = rowsAnal(Count1).Item("INTGROUP")
                    strAnal = GetUserAnalyteName(strAnal, False, intgroup)
                    'copy and paste appropriate number of sections
                    '*****
                    'go to end of first paragraph
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")

                    'find first paragraph return
                    .Selection.Find.ClearFormatting()
                    With .Selection.Find
                        .Text = "^p"
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                        .Format = True
                        .Execute()
                        '.Execute()
                    End With
                    pos3 = .selection.end

                    'go back to Temp 1
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos3 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

                    wd.Selection.Copy()

                    .Selection.Moveright(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    '.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    .Selection.TypeParagraph()
                    '.Selection.TypeParagraph()
                    '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                    wrdSelection = wd.Selection()
                    'record first bookmark
                    With .ActiveDocument.Bookmarks
                        ctBM = ctBM + 1
                        str1 = "T" & ctBM
                        arrBM(1, Count1) = str1
                        .Add(Range:=wrdSelection.Range, Name:=str1)
                        .ShowHidden = False
                    End With
                    pos1 = .Selection.Start

                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
                    '.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault)
                    .selection.paste()

                    '.Selection.Moveleft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=True)

                    .Selection.Moveright(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=False)
                    pos4 = .Selection.Start
                    'add Temp2 bookmark
                    wrdSelection = wd.Selection()
                    With .ActiveDocument.Bookmarks
                        ctBM = ctBM + 1
                        str1 = "T" & ctBM
                        arrBM(2, Count1) = str1
                        .Add(Range:=wrdSelection.Range, Name:=str1)
                        .ShowHidden = False
                    End With
                    '****
                Next
            End If

            For Count1 = 0 To intRows - 1

                strAnal = rowsAnal(Count1).Item("AnalyteDescription")

                '20181108 LEE:
                'get useranalyte
                intGroup = rowsAnal(Count1).Item("INTGROUP")
                strAnal = GetUserAnalyteName(strAnal, False, intGroup)

                'now do searchreplace on each section

                If Count1 = 0 Then
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:=arrBM(1, Count1))
                    pos4 = .Selection.Bookmarks.item(arrBM(1, Count1)).Start
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
                    pos3 = .Selection.Bookmarks.item("Temp1").Start
                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos4 - pos3, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

                Else
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:=arrBM(2, Count1 - 1))
                    pos4 = .Selection.Bookmarks.item(arrBM(2, Count1 - 1)).Start
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:=arrBM(1, Count1 - 1))
                    pos3 = .Selection.Bookmarks.item(arrBM(1, Count1 - 1)).Start
                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos4 - pos3, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                End If

                If boolShowExample Then
                    str1 = "[NA]"
                Else
                    If Count1 > drowsA.Length - 1 Then
                        If drowsA.Length = 0 Then
                            str1 = "[NA]"
                        Else
                            str1 = drowsA(drowsA.Length - 1).Item("AppendixNumber")
                        End If
                    Else
                        str1 = drowsA(Count1).Item("AppendixNumber")
                    End If

                End If

                Call SearchReplaceMethodSummary(wd, strAnal, str1)

            Next

            'delete all temp bookmarks
            For Count1 = 0 To intRows - 2
                .ActiveDocument.Bookmarks.item(arrBM(1, Count1)).Delete()
                .ActiveDocument.Bookmarks.item(arrBM(2, Count1)).Delete()
            Next

            'select entire section
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            pos2 = .Selection.Bookmarks.item("Temp2").Start
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            pos1 = .Selection.Bookmarks.item("Temp1").Start
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            'now do Watson Representative Chromatography section
            str1 = "AppendixName = 'Chromatogram'"
            drowsA = tblAppendix.Select(str1)
            int1 = drowsA.Length
            Dim dtbl As System.Data.DataTable
            'Dim rowsC() As DataRow
            'Dim strF1 As String
            'dtbl = frmH.tblChromPath
            strF1 = "ID_TBLSTUDIES = " & id_tblStudies
            'rowsC = dtbl.Select(strF1)
            int2 = rowsAF.Length
            'dv = frmH.dgvRepChrom.DataSource
            'int2 = dv.Count

            str1 = "[NA]"
            str1 = WatsonRepChrom()

            Dim mysel As Microsoft.Office.Interop.Word.Selection
            Dim strFind As String
            Dim varReplace As String

            .Selection.Moveleft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            '''''''''wdd.visible = True

            varReplace = str1
            mysel = wd.selection
            strFind = "[WATSONREPCHROMSECTION]"
            With wd

                With mysel.Find
                    .ClearFormatting()
                    .Text = strFind
                    With .Replacement
                        .ClearFormatting()
                        .Text = varReplace
                    End With
                    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
                    If .Found Then
                        If Len(varReplace) = 0 Or StrComp(NZ(varReplace, "[None]"), "[None]", CompareMethod.Text) = 0 Or StrComp(varReplace, "[NA]", CompareMethod.Text) = 0 Or InStr(1, varReplace, "[NA]", CompareMethod.Text) > 0 Then
                            'add entries to arrReportNA
                            ctArrReportNA = ctArrReportNA + 1
                            arrReportNA(1, ctArrReportNA) = "Method Summary"
                            arrReportNA(2, ctArrReportNA) = "Representative Watson Run ID"
                            arrReportNA(3, ctArrReportNA) = "Appendix Tab"
                            arrReportNA(4, ctArrReportNA) = "Chromatography Table"
                        End If
                    End If
                End With

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

            End With

            'Representative raw chromatographic data from Watson run [REPRCHROMWATSONRUNNUMBER] is provided in Appendix [REPRCHROMAPPENDIXLETTER]

            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            pos2 = .Selection.Bookmarks.item("Temp2").Start
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            pos1 = .Selection.Bookmarks.item("Temp1").Start
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

            'clear formatting again
            .selection.Find.ClearFormatting()

        End With

end1:

    End Sub

    Function WatsonRepChrom() As String

        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim var1, var2
        Dim dtbl As System.Data.DataTable
        Dim drowsC() As DataRow
        Dim drowsA() As DataRow
        Dim Count1 As Short
        Dim rowsAF() As DataRow
        Dim intA As Short
        Dim intC As Short
        Dim strA As String

        If gboolDisplayAttachments Then
            strA = "Attachment"
        Else
            strA = "Appendix"
        End If

        WatsonRepChrom = "[NA]"

        str1 = "AppendixName = 'Chromatogram'"
        drowsC = tblAppendix.Select(str1)
        intA = drowsC.Length

        dtbl = tblAppFigs
        str1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARTYPE = 'RC'"
        rowsAF = dtbl.Select(str1)
        int2 = rowsAF.Length

        str1 = WatsonRepChrom

        '''''''''wdd.visible = True

        If intA <> 0 Then
            If intA = 1 Then
                str1 = "Representative raw chromatographic data from Watson Run ID "
                'str1 = str1 & NZ(dv(0).Item("numWatsonRunNumber"), "[NA]") & " for " & NZ(dv(0).Item("charAnalyte"), "[NA]") & " are provided in Appendix "
                Try
                    var2 = NZ(drowsC(0).Item("AnalyteName"), "[ANALYTE]")
                    '20181108 LEE:
                    'get useranalyte
                    var2 = GetUserAnalyteNameNoGroup(CStr(var2))
                Catch ex As Exception
                    var2 = "[NA]"
                End Try

                '20181015 LEE:
                'put nbh back in
                var2 = Replace(var2, "-", NBHReal, 1, -1, CompareMethod.Text)

                If Len(var2) = 0 Or StrComp(var2, "[NA]", CompareMethod.Text) = 0 Then
                    'str1 = str1 & NZ(drowsC(0).Item("RepWatsonID"), "[NA]") & " for [ANALYTE] are provided in Appendix"
                    'str1 = str1 & NZ(drowsC(0).Item("RepWatsonID"), "[NA]") & " for [ANALYTE] are provided in " & strA
                    str1 = str1 & NZ(drowsC(0).Item("RepWatsonID"), "[NA]") & " are provided in " & strA
                Else
                    str1 = str1 & NZ(drowsC(0).Item("RepWatsonID"), "[NA]") & " for " & var2 & " are provided in " & strA
                End If

                var1 = NZ(drowsC(0).Item("AppendixNumber"), "")
                If Len(var1) = 0 Then
                    var2 = " [NA]"
                Else
                    var2 = "_" & var1 'AppendixLetter(var1)'APPENDIXLETTER will be done later, I think
                End If
                str1 = str1 & var2 ' & ". "

            ElseIf intA > 1 Then

                str1 = ""
                For Count1 = 0 To intA - 1
                    If Count1 = 0 Then
                        str1 = str1 & "Representative raw chromatographic data from Watson Run ID "
                    Else
                        str1 = str1 & ". Representative raw chromatographic data from Watson Run ID "
                    End If
                    Try
                        var2 = NZ(drowsC(Count1).Item("AnalyteName"), "")
                    Catch ex As Exception
                        var2 = "[NA]"
                    End Try
                    If Len(var2) = 0 Then 'try arrAnalyte
                        Try
                            var2 = arrAnalytes(1, Count1 + 1)
                        Catch ex As Exception
                            var2 = "[NA]"
                        End Try
                    End If

                    '20181015 LEE:
                    'put nbh back in
                    var2 = Replace(var2, "-", NBHReal, 1, -1)

                    If Len(var2) = 0 Or StrComp(var2, "[NA]", CompareMethod.Text) = 0 Then
                        str1 = str1 & NZ(drowsC(Count1).Item("RepWatsonID"), "[NA]") & " are provided in " & strA
                        'str1 = str1 & NZ(drowsC(Count1).Item("RepWatsonID"), "[NA]") & " for [ANALYTE] are provided in " & strA
                        'str1 = str1 & NZ(drowsC(Count1).Item("RepWatsonID"), "[NA]") & " for [NA] are provided in Appendix"
                    Else
                        str1 = str1 & NZ(drowsC(Count1).Item("RepWatsonID"), "[NA]") & " for " & var2 & " are provided in " & strA
                    End If
                    'str1 = str1 & NZ(drowsC(Count1).Item("RepWatsonID"), "[NA]") & " for " & NZ(drowsC(Count1).Item("AnalyteName"), "[NA]") & " are provided in Appendix "

                    Try
                        var1 = NZ(drowsC(Count1).Item("AppendixNumber"), "[NA]")
                    Catch ex As Exception
                        var1 = ""
                    End Try
                    If Len(var1) = 0 Then
                        var2 = "[NA]"
                    Else
                        var2 = "_" & var1 'AppendixLetter(var1)
                    End If
                    str1 = str1 & var2 ' & ". "

                    'str1 = str1 & drowsC(Count1).Item("AppendixNumber") & ". "
                Next

            End If
        Else 'one representative chromatogram. Allow [ANALYTE]
            If int2 = 0 Then
                str1 = "[NA]"
            ElseIf int2 = 1 Then
                str1 = "Representative raw chromatographic data from Watson Run ID "
                ''str1 = str1 & NZ(dv(0).Item("numWatsonRunNumber"), "[NA]") & " for " & NZ(dv(0).Item("charAnalyte"), "[NA]") & " are provided in Appendix "
                'var2 = NZ(rowsAF(0).Item("charAnalyte"), "[NA]")
                'If Len(var2) = 0 Or StrComp(var2, "[NA]", CompareMethod.Text) = 0 Then
                '    str1 = str1 & NZ(rowsAF(0).Item("numWatsonRunNumber"), "[NA]") & " for [ANALYTE] are provided in Appendix "
                'Else
                '    str1 = str1 & NZ(rowsAF(0).Item("numWatsonRunNumber"), "[NA]") & " for " & var2 & " are provided in Appendix "
                'End If
                'var2 = "[NA]"
                'str1 = str1 & var2 ' & ". "
                var2 = NZ(drowsC(0).Item("AnalyteName"), "")

                '20181108 LEE:
                'get useranalyte
                var2 = GetUserAnalyteNameNoGroup(CStr(var2))

                '20181015 LEE:
                'put nbh back in
                var2 = Replace(var2, "-", NBHReal, 1, -1, CompareMethod.Text)

                '20181015 LEE:
                'put nbh back in
                var2 = Replace(var2, "-", NBHReal, 1, -1, CompareMethod.Text)

                If Len(var2) = 0 Or StrComp(var2, "[NA]", CompareMethod.Text) = 0 Then
                    str1 = str1 & NZ(rowsAF(0).Item("numWatsonRunNumber"), "[NA]") & " for [ANALYTE] are provided in " & strA
                Else
                    str1 = str1 & NZ(rowsAF(0).Item("numWatsonRunNumber"), "[NA]") & " for " & var2 & " are provided in " & strA
                End If

                Try
                    var1 = NZ(drowsC(Count1).Item("AppendixNumber"), "[NA]")
                Catch ex As Exception
                    var1 = ""
                End Try
                If Len(var1) = 0 Then
                    var2 = "[NA]"
                Else
                    var2 = "_" & var1 'AppendixLetter(var1)
                End If
                str1 = str1 & var2 ' & ". "

            ElseIf int2 > 1 Then
                str1 = ""
                'For Count1 = 0 To int1 - 1
                For Count1 = 0 To int2 - 1
                    'str1 = str1 & "Representative raw chromatographic data from Watson Run ID "
                    If Count1 = 0 Then
                        str1 = str1 & "Representative raw chromatographic data from Watson Run ID "
                    Else
                        str1 = str1 & ". Representative raw chromatographic data from Watson Run ID "
                    End If
                    Try
                        var2 = NZ(rowsAF(Count1).Item("charAnalyte"), "")
                    Catch ex As Exception
                        var2 = "[NA]"
                    End Try
                    If Len(var2) = 0 Then 'try arrAnalyte
                        Try
                            var2 = arrAnalytes(1, Count1 + 1)
                        Catch ex As Exception
                            var2 = "[NA]"
                        End Try
                    End If

                    str1 = str1 & NZ(rowsAF(Count1).Item("numWatsonRunNumber"), "[NA]") & " for " & var2 & " are provided in " & strA

                    Try
                        var1 = NZ(drowsC(Count1).Item("AppendixNumber"), "[NA]")
                    Catch ex As Exception
                        var1 = ""
                    End Try
                    If Len(var1) = 0 Then
                        var2 = "[NA]"
                    Else
                        var2 = "_" & var1 'AppendixLetter(var1)
                    End If
                    str1 = str1 & var2 ' & ". "

                    'If Len(var2) = 0 Or StrComp(var2, "[NA]", CompareMethod.Text) = 0 Then
                    '    str1 = str1 & NZ(rowsAF(Count1).Item("numWatsonRunNumber"), "[NA]") & " for [ANALYTE] are provided in Appendix "
                    'Else
                    '    str1 = str1 & NZ(rowsAF(Count1).Item("numWatsonRunNumber"), "[NA]") & " for " & var2 & " are provided in Appendix "
                    'End If
                    'str1 = str1 & NZ(rowsAF(Count1).Item("numWatsonRunNumber"), "[NA]") & " for " & NZ(rowsAF(Count1).Item("charAnalyte"), "[NA]") & " are provided in Appendix "
                    'str1 = str1 & "[NA]"
                Next
            End If
        End If

        WatsonRepChrom = str1

    End Function

    Sub GuWuComplianceStatement01(ByVal wd, ByVal charSectionName)
        Dim var1, var2, var3, var8, var9
        Dim intPers As Short

        With wd
            'don't need this here

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'var2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'var1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            ''call searchreplace(wd, charSectionName) 
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

            'add signature block(s)
            Dim strF As String
            Dim strS As String
            Dim dtbl As System.Data.DataTable
            Dim rows() As DataRow
            Dim int1 As Short
            Dim intRows As Short
            Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
            Dim Count1 As Short
            Dim intRow As Short
            Dim intCol As Short
            Dim num1, num2
            Dim str1 As String
            Dim tblC As System.Data.DataTable
            Dim rowsC() As DataRow
            Dim intC As Short
            Dim strC As String
            Dim strDeg As String

            dtbl = tblContributingPersonnel
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDESIGONCOMPSTATEMENT	 = -1"
            strS = "INTORDER ASC"
            rows = dtbl.Select(strF, strS)
            intPers = rows.Length

            tblC = tblCorporateNickNames
            str1 = NZ(frmH.cbxSubmittedBy.Text, "GA")
            strF = "CHARNICKNAME = '" & str1 & "'"
            rowsC = tblC.Select(strF)
            var1 = rowsC(0).Item("ID_TBLCORPORATENICKNAMES")

            tblC = tblCorporateAddresses
            strF = "ID_TBLCORPORATENICKNAMES = " & var1 & " AND BOOLINCLUDEINTITLE = -1"
            Erase rowsC
            'strS = "ID_TBLADDRESSLABELS ASC"
            'rowsC = tblC.Select(strF, strS)
            strS = "ID_TBLADDRESSLABELS ASC"
            rowsC = tblC.Select(strF, strS)
            intC = rowsC.Length
            strC = ""
            For Count1 = 0 To intC - 1
                var1 = NZ(rowsC(Count1).Item("CHARVALUE"), "No")
                If Count1 = 0 And Count1 = intC - 1 Then
                    strC = var1
                ElseIf Count1 <> intC - 1 Then
                    strC = strC & var1 & ChrW(10)
                ElseIf Count1 > 0 And Count1 = intC - 1 Then
                    strC = strC & var1
                End If
            Next

            If intPers = 0 Then 'ignore
            Else

                If intPers < 2 Then
                    intRows = 1
                ElseIf intPers < 4 Then
                    intRows = 2
                ElseIf intPers < 6 Then
                    intRows = 3
                ElseIf intPers < 8 Then
                    intRows = 4
                End If

                wd.Selection.TypeParagraph()
                wd.Selection.TypeParagraph()
                wd.Selection.TypeParagraph()
                wd.Selection.TypeParagraph()
                wd.Selection.TypeParagraph()
                wrdSelection = wd.Selection()

                .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intRows, NumColumns:=3, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                .Selection.Tables.Item(1).Select()
                Call GlobalTableParaFormat(wd)


                num1 = .Selection.Tables.item(1).Columns.item(1).width * 3
                .Selection.Tables.item(1).Rows.AllowBreakAcrossPages = False
                .Selection.Tables.item(1).Select()
                .Selection.Rows.AllowBreakAcrossPages = False

                With .Selection 'remove initial borders
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                End With

                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                'size columns
                '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)

                For Count1 = 1 To 3
                    Select Case Count1
                        Case 1
                            num2 = 0.45 * num1
                        Case 2
                            num2 = 0.1 * num1
                        Case 3
                            num2 = 0.45 * num1
                    End Select
                    .Selection.Tables.item(1).Columns.item(Count1).PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                    .Selection.Tables.item(1).Columns.item(Count1).PreferredWidth = num2
                Next

                'enter personnel
                intRow = 1
                intCol = 1
                For Count1 = 0 To intPers - 1
                    If Count1 = 0 Then
                    ElseIf IsEven(Count1) Then
                        intRow = intRow + 1
                        intCol = 1
                    Else
                        intCol = intCol + 2
                    End If

                    .Selection.Tables.item(1).Cell(intRow, intCol).Select()
                    .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    strDeg = NZ(rows(Count1).Item("CHARCPDEGREE"), "")
                    If Len(strDeg) = 0 Then
                        str1 = rows(Count1).Item("CHARCPNAME") & " / Date"
                    Else
                        str1 = rows(Count1).Item("CHARCPNAME") & ", " & strDeg & " / Date"
                    End If
                    str1 = str1 & ChrW(10)
                    str1 = str1 & rows(Count1).Item("CHARCPTITLE")
                    'str1 = str1 & ChrW(10) & strC & ChrW(10) & ChrW(10) & ChrW(10) & ChrW(10)
                    str1 = str1 & ChrW(10) & ChrW(10) & ChrW(10) & ChrW(10) & ChrW(10)
                    .Selection.TypeText(str1)
                Next

                'move to next row below table
                .selection.Tables.item(1).select()
                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)


            End If

        End With
    End Sub

    Sub EnterTableNumber(ByVal wd As Microsoft.Office.Interop.Word.Application, ByRef strTName As String, ByVal numRows As Short, ByVal strAnal As String, ByVal strStability As String, intTableID As Int64, intGroup As Short, idT As Int64)

        'strTName

        Dim strTitle As String = strTName

        Dim strM As String
        Dim str1 As String
        Dim str1a As String
        Dim str1b As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim str5 As String
        Dim str6 As String
        Dim str7 As String
        Dim str8 As String
        Dim str9 As String
        Dim str10 As String
        Dim str11 As String
        Dim wdSel As Microsoft.Office.Interop.Word.Selection
        Dim var1, var2
        Dim boolInTable As Boolean
        Dim strFind As String
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim var8
        Dim varReplace
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dtbl As System.Data.DataTable
        Dim drows() As DataRow
        Dim tblNick As System.Data.DataTable
        Dim rowsNick() As DataRow

        Dim rng1 As Word.Range
        Dim rng2 As Word.Range

        Dim wdTbl As Word.Table

        Dim strCaptionTrailer As String = NZ(lcharCaptionTrailer, "")
        Dim boolTab As Boolean = True

        '20181127 LEE
        Dim boolSpace As Boolean = False
        Dim boolSR As Boolean = False 'soft return

        If StrComp(NZ(gstrCAPTIONFOLLOW, "Tab"), "Tab", CompareMethod.Text) = 0 Then
            boolTab = True
            boolSpace = False
            boolSR = False
        ElseIf StrComp(NZ(gstrCAPTIONFOLLOW, "Tab"), "Space", CompareMethod.Text) = 0 Then
            boolTab = False
            boolSpace = True
            boolSR = False
        Else
            boolTab = False
            boolSpace = False
            boolSR = True
        End If

        '20180523 LEE
        Dim boolV As Boolean = wd.Visible

        '20181108 LEE:
        'get useranalyte
        strAnal = GetUserAnalyteName(strAnal, False, intGroup)

        'legend
        'Public boolExcludeTableNumbers As Boolean = False
        'Public boolExcludeTableTitles As Boolean = False
        'Public boolExcludeEntireTableTitle As Boolean = False

        'If boolExcludeEntireTableTitle Then
        '    GoTo end1
        'End If

        'str1a = Replace(strAnal, "-", NBH, 1, -1, CompareMethod.Text) 'non breaking hyphen'doesn't seem to work here
        str1a = strAnal
        str1b = Replace(str1a, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space
        strAnal = str1b

        'CapitAllWords


        'str1 = Replace(strTitle, "[ANALYTE]", Capit(UnCapit(strAnal, False)), 1, -1, CompareMethod.Text)

        '20181015 LEE:
        'put nbh back in
        strAnal = Replace(strAnal, "-", NBHReal, 1, -1, CompareMethod.Text)

        str2 = CapitAllWords(strAnal)
        str1 = Replace(strTitle, "[ANALYTE]", str2, 1, -1, CompareMethod.Text)
        strTitle = str1
        Dim strIS As String
        'need to find strIS
        Dim strF As String
        If intGroup = -1 Then
            'came from is
            strIS = strAnal
        Else
            strF = "INTGROUP = " & intGroup
            Dim rowsA() As DataRow = tblAnalyteGroups.Select(strF)
            If rowsA.Length = 0 Then
                strIS = "NA"
            ElseIf rowsA.Length = 1 Then
                strIS = rowsA(0).Item("INTSTD")
            End If
        End If



        'str1 = Replace(str1, "[INTERNALSTANDARD]", Capit(UnCapit(strIS, False)), 1, -1, CompareMethod.Text)
        str2 = CapitAllWords(strIS)
        str1 = Replace(strTitle, "[INTERNALSTANDARD]", str2, 1, -1, CompareMethod.Text)
        str3 = str1
        'str1a = Replace(str1, "-", NBH, 1, -1, CompareMethod.Text) 'non breaking hyphen
        'str1b = Replace(str1a, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space
        'str1 = str1a


        'Stability stuff
        Dim strF1 As String
        strF1 = "ID_TBLREPORTTABLE = " & idT
        Dim rowsTR() As DataRow = tblTableProperties.Select(strF1)

        '[PERIODTEMP]
        str4 = CapitAllWords(strStability)
        str2 = Replace(str1, "[PERIOD TEMP]", str4, 1, -1, CompareMethod.Text)
        '20181127 LEE:
        'add additional field code
        str2 = Replace(str2, "[STABILITYCONDITIONS]", str4, 1, -1, CompareMethod.Text)

        '20181127 LEE:
        'start doing nbh again
        str2 = Replace(str2, "-", NBHReal, 1, -1, CompareMethod.Text) 'non breaking hyphen

        '[#Cycles]
        Select Case intTableID
            Case 19
                Dim strNumCycles As String = "NA"
                If rowsTR.Length = 0 Then
                Else
                    strNumCycles = NZ(rowsTR(0).Item("INTNUMBEROFCYCLES"), "NA").ToString
                End If
                str3 = Replace(str2, "[#Cycles]", strNumCycles, 1, -1, CompareMethod.Text)
            Case Else
                str3 = str2
        End Select

        '20190220 LEE: Need to add more stability field codes
        For Count1 = 1 To 3
            Select Case Count1
                Case 1
                    str1 = "[Period]"
                    str2 = "CHARTIMEPERIOD"
                Case 2
                    str1 = "[Period Units]"
                    str2 = "CHARTIMEFRAME"
                Case 3
                    str1 = "[Temp]"
                    str2 = "CHARPERIODTEMP"
            End Select

            str4 = NZ(rowsTR(0).Item(str2), "[NA]")

            str3 = Replace(str3, str1, str4, 1, -1, CompareMethod.Text)

        Next


        dgv = frmH.dgvDataWatson
        dv = dgv.DataSource
        strFind = "[SPECIES]"
        int1 = FindRowDV("Species", dv)
        'var8 = dg.Item(int1, 1)
        'if in table, then capitalize first letter
        var8 = Capit(NZ(dv.Item(int1).Item(1), "")) 'UnCapit(NZ(dv.Item(int1).Item(1), ""), True)
        If IsDBNull(var8) Then
            varReplace = "[NA]"
        ElseIf Len(var8) = 0 Then
            varReplace = "[NA]"
        Else
            'varReplace = LowerCase(var8) 'UnCapit(Trim(var8), True)
            'varReplace = Capit(LCase(Trim(var8)))
            varReplace = Trim(var8) ' CapitAllWords(Trim(var8))
            varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)
        End If
        For Count1 = 1 To 3
            Select Case Count1
                Case 1
                    strFind = "[SPECIES]"
                    If InStr(1, str3, strFind, CompareMethod.Text) > 0 Then
                        varReplace = CapitAllWords(Trim(varReplace))
                    End If

                Case 2
                    strFind = "[LC_SPECIES]"
                    If InStr(1, str3, strFind, CompareMethod.Text) > 0 Then
                        varReplace = LowerCase(Trim(varReplace.ToString))
                    End If

                Case 3
                    strFind = "[UC_SPECIES]"
                    If InStr(1, str3, strFind, CompareMethod.Text) > 0 Then
                        varReplace = CapitAllWords(Trim(varReplace))
                    End If

            End Select

            str4 = Replace(str3, strFind, varReplace, 1, -1, CompareMethod.Text)
            str3 = str4
        Next


        strFind = "[MATRIX]"
        int1 = FindRowDV("Matrix", dv)
        'var8 = dg.Item(int1, 1)
        var8 = Capit(NZ(dv.Item(int1).Item(1), ""))
        If IsDBNull(var8) Then
            varReplace = "[NA]"
        ElseIf Len(var8) = 0 Then
            varReplace = "[NA]"
        Else
            'varReplace = LowerCase(var8) 'UnCapit(Trim(var8), True)
            '20160518 LEE:
            'Don't do capit the following tables, which already have a capit combination of matrices if # matrices > 1
            'Select Case intTableID
            '    Case 1
            '        varReplace = Trim(var8)
            '    Case Else
            '        'varReplace = Capit(LCase(Trim(var8)))
            '        varReplace = CapitAllWords(Trim(var8))
            'End Select
            varReplace = Trim(var8)
            varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)
        End If
        For Count1 = 1 To 3
            Select Case Count1
                Case 1
                    strFind = "[MATRIX]"
                    If InStr(1, varReplace, strFind, CompareMethod.Text) > 0 Then
                        varReplace = CapitAllWords(Trim(varReplace))
                    End If

                Case 2
                    strFind = "[LC_MATRIX]"
                    If InStr(1, varReplace, strFind, CompareMethod.Text) > 0 Then
                        varReplace = LowerCase(Trim(varReplace))
                    End If

                Case 3
                    strFind = "[UC_MATRIX]"
                    If InStr(1, varReplace, strFind, CompareMethod.Text) > 0 Then
                        varReplace = CapitAllWords(Trim(varReplace))
                    End If

            End Select
            str5 = Replace(str4, strFind, varReplace, 1, -1, CompareMethod.Text)
            str4 = str5
        Next

        strFind = "[SPONSOR]"
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
            str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = -1"
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
        str6 = Replace(str5, strFind, varReplace, 1, -1, CompareMethod.Text)

        strFind = "[SPONSOR_STUDY_#]"
        'strFind = "[SPONSORSTUDYNUMBER]"
        'dtbl = tblData
        dtbl = tblData
        str2 = "id_tblStudies = " & id_tblStudies 'intIDtblStudies()
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
        str7 = Replace(str6, strFind, varReplace, 1, -1, CompareMethod.Text)


        strFind = "[ANTICOAGULANT]"
        str1 = NZ(frmH.cbxAnticoagulant.Text, "")
        Dim boolAcro As Boolean = False
        'determine if text should be capitalized
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
                boolAcro = True
            Else
                var8 = str1 ' CapitAllWords(str1)
            End If
        End If
        var8 = str1
        '20181015 LEE:
        'put nbh back in
        var8 = Replace(var8, "-", NBHReal, 1, -1)

        If boolAcro Then
            str8 = var8 ' Replace(str7, strFind, varReplace, 1, -1, CompareMethod.Text)
        Else
            If IsDBNull(var8) Then
                varReplace = "[NA]"
            ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                varReplace = "[NA]"
            Else
                varReplace = Trim(var8) ' Capit(var8)
                varReplace = Replace(varReplace, " ", ChrW(160), 1, -1, CompareMethod.Text)
            End If
            For Count1 = 1 To 3
                Select Case Count1
                    Case 1
                        strFind = "[ANTICOAGULANT]"
                        If InStr(1, varReplace, strFind, CompareMethod.Text) > 0 Then

                        End If
                        varReplace = CapitAllWords(Trim(varReplace))
                    Case 2
                        strFind = "[LC_ANTICOAGULANT]"
                        If InStr(1, varReplace, strFind, CompareMethod.Text) > 0 Then

                        End If
                        varReplace = LowerCase(Trim(varReplace))
                    Case 3
                        strFind = "[UC_ANTICOAGULANT]"
                        If InStr(1, varReplace, strFind, CompareMethod.Text) > 0 Then

                        End If
                        varReplace = CapitAllWords(Trim(varReplace))
                End Select
                str8 = Replace(str7, strFind, varReplace, 1, -1, CompareMethod.Text)
                str7 = str8
            Next
        End If

        '

        strFind = "[INSUPPORTOF]" ' "[IN_SUPPORT_OF]"
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
            str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = -1" 
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
        str9 = Replace(str8, strFind, varReplace, 1, -1, CompareMethod.Text)

        '****

        strFind = "[CORPORATESTUDY/PROJECTNUMBER]"

        'dtbl = tblData
        dtbl = tblData
        str2 = "id_tblStudies = " & id_tblStudies 'intIDtblStudies()
        drows = dtbl.Select(str2)
        int1 = drows.Length
        'format var8
        If int1 = 0 Then
            varReplace = "[NA]"
        Else
            var8 = drows(0).Item("CHARCORPORATESTUDYID")
            If IsDBNull(var8) Then
                varReplace = "[NA]"
            ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                varReplace = "[NA]"
            Else
                varReplace = var8
            End If
        End If
        str10 = Replace(str9, strFind, varReplace, 1, -1, CompareMethod.Text)

        '****
        strFind = "[WATSONSTUDYTITLE]"

        'dtbl = tblData
        dtbl = tblwSTUDY
        int1 = dtbl.Rows.Count
        'format var8
        If int1 = 0 Then
            varReplace = "[NA]"
        Else
            var8 = dtbl.Rows(0).Item("STUDYTITLE")
            If IsDBNull(var8) Then
                varReplace = "[NA]"
            ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                varReplace = "[NA]"
            Else
                varReplace = var8
            End If
        End If
        str11 = Replace(str10, strFind, varReplace, 1, -1, CompareMethod.Text)


        '****

        '*****

        Dim dv1 As DataView = frmH.dgvWatsonAnalRef.DataSource
        Dim num1 As Single
        Dim boolNum As Boolean

        Dim rows() As DataRow
        strF = "INTGROUP = " & intGroup
        Erase rows
        rows = tblCalStdGroupsAll.Select(strF)

        If rows.Length > 0 Then

            strFind = "[LLOQ]"
            var1 = rows(0).Item("LLOQ")
            num1 = SigFigOrDecString(CDec(var1), LSigFig, False)
            boolNum = True
            var8 = num1
            If IsDBNull(var8) Then
                varReplace = "[NA]"
            ElseIf Len(var8) = 0 Then
                varReplace = "[NA]"
            Else
                varReplace = var8
            End If
            str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)

            strFind = "[LLOQUNITS]"
            var8 = rows(0).Item("CONCENTRATIONUNITS")
            If IsDBNull(var8) Then
                varReplace = "[NA]"
            ElseIf Len(var8) = 0 Then
                varReplace = "[NA]"
            Else
                varReplace = var8
            End If
            str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)

            strFind = "[ULOQ]"
            var1 = rows(0).Item("ULOQ")
            num1 = SigFigOrDecString(CDec(var1), LSigFig, False)
            boolNum = True
            var8 = num1
            If IsDBNull(var8) Then
                varReplace = "[NA]"
            ElseIf Len(var8) = 0 Then
                varReplace = "[NA]"
            Else
                varReplace = var8
            End If
            str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)

            strFind = "[ULOQUNITS]"
            var8 = rows(0).Item("CONCENTRATIONUNITS")
            If IsDBNull(var8) Then
                varReplace = "[NA]"
            ElseIf Len(var8) = 0 Then
                varReplace = "[NA]"
            Else
                varReplace = var8
            End If
            str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)

        End If

        '20180329 LEE:
        'need to look for custom field codes

        Dim dtbl1 As DataTable
        Dim dtbl2 As DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim id1 As Int64
        Dim strF2 As String


        dtbl1 = tblCustomFieldCodes
        dtbl2 = tblFieldCodes

        strF1 = "ID_TBLSTUDIES = " & id_tblStudies
        rows1 = dtbl1.Select(strF1)

        For Count1 = 0 To rows1.Length - 1

            id1 = NZ(rows1(Count1).Item("ID_TBLFIELDCODES"), -1)
            strF2 = "ID_TBLFIELDCODES = " & id1
            rows2 = dtbl2.Select(strF2)

            If rows2.Length = 0 Then
            Else
                strFind = NZ(rows2(0).Item("CHARFIELDCODE"), "AAAAAA")
                varReplace = NZ(rows1(Count1).Item("CHARVALUE"), "[NA]")
                str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)
            End If

        Next

        '******


        '20180329 LEE:
        'add additional stuff

        dtbl = tblCorporateAddresses

        strFind = "[SUBMITTEDBY]" ' "[IN_SUPPORT_OF]"
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
            str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = -1" 
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
        str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)

        strFind = "[SUBMITTEDTO]" ' "[IN_SUPPORT_OF]"
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
            str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = -1" 
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
        str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)

        '******


        'strFind = "[LLOQ]"
        'int1 = FindRowDVByCol("LLOQ", dv1, "Item")
        ''var8 = dg.Item(int1, 1)
        'var1 = dv1.Item(int1).Item(strAnal)
        'num1 = SigFigOrDecString(CDec(var8), LSigFig, False)
        'boolNum = True
        'var8 = num1
        'If IsDBNull(var8) Then
        '    varReplace = "[NA]"
        'ElseIf Len(var8) = 0 Then
        '    varReplace = "[NA]"
        'Else
        '    varReplace = var8
        'End If
        'str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)

        'strFind = "[LLOQUNITS]"
        'int1 = FindRowDVByCol("LLOQ Units", dv1, "Item")
        'var8 = dv1.Item(int1).Item(strAnal)
        'If IsDBNull(var8) Then
        '    varReplace = "[NA]"
        'ElseIf Len(var8) = 0 Then
        '    varReplace = "[NA]"
        'Else
        '    varReplace = var8
        'End If
        'str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)


        'strFind = "[ULOQ]"
        'int1 = FindRowDVByCol("ULOQ", dv1, "Item")
        'num1 = dv1.Item(int1).Item(strAnal)
        'num1 = SigFigOrDecString(CDec(num1), LSigFig, False)
        'boolNum = True
        'var8 = num1
        'If IsDBNull(var8) Then
        '    varReplace = "[NA]"
        'ElseIf Len(var8) = 0 Then
        '    varReplace = "[NA]"
        'Else
        '    varReplace = var8
        'End If
        'str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)


        'strFind = "[ULOQUNITS]"
        'int1 = FindRowDVByCol("ULOQ Units", dv1, "Item")
        'var8 = dv1.Item(int1).Item(strAnal)
        'If IsDBNull(var8) Then
        '    varReplace = "[NA]"
        'ElseIf Len(var8) = 0 Then
        '    varReplace = "[NA]"
        'Else
        '    varReplace = var8
        'End If
        'str11 = Replace(str11, strFind, varReplace, 1, -1, CompareMethod.Text)


        '*****

        wdTbl = wd.Selection.Tables.Item(1)


        strTitle = str11

        'replace hyphens with nbh's
        'str1 = Replace(strTitle, "-", NBH, 1, -1, CompareMethod.Text)
        str1 = Replace(strTitle, "-", NBH, 1, -1, CompareMethod.Text)
        strTitle = str1

        '''''''''wdd.visible = True

        'legend
        'Public boolExcludeTableNumbers As Boolean = False
        'Public boolExcludeTableTitles As Boolean = False
        'Public boolExcludeEntireTableTitle As Boolean = False

        '20160207 LEE: set cell padding to 0
        '20170224 LEE: Don't set table cell padding here because some tables have different settings
        'set at table creation level
        'With wd.Selection.Tables.Item(1)
        '    .TopPadding = 0 ' InchesToPoints(0)
        '    .BottomPadding = 0 ' InchesToPoints(0)
        '    .LeftPadding = 0 ' InchesToPoints(0)
        '    .RightPadding = 0 ' InchesToPoints(0)
        '    .Spacing = 0
        'End With
        'Call SetCellPadding(wd.Selection.Tables.Item(1))

        If boolExcludeEntireTableTitle Then
            GoTo end1
        End If

        '20180524 LEE:
        'wd.Visible = True

        'wait
        'Call Pause(0.25)

        '20180524 LEE:
        Try

            With wd

                '.Selection.Tables.Item(1).Cell(1, 1).Select()
                .Selection.Tables.Item(1).Rows(1).Select()

                .Selection.InsertRowsAbove(1)

                '20180525 LEE:
                wdTbl.Select()
                .Selection.Font.Name = .ActiveDocument.Styles("Normal").Font.Name
                .Selection.Font.Bold = False

                'if Table Grid style contains 'Unicode'
                'previous row style gets set to it
                'put it back
                'Call RemoveUnicode(wd)

                wdTbl.Rows(1).Select()

                Try
                    .Selection.Cells.Merge()
                Catch ex As Exception
                    var1 = var1
                    wdTbl.Cell(1, 1).Select()
                    .Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    wdTbl.Rows(wdTbl.Rows.Count).Select()
                    wdTbl.Rows(1).Select()
                    .Selection.InsertRowsAbove(1)

                    '20180525 LEE:
                    wdTbl.Select()
                    .Selection.Font.Name = .ActiveDocument.Styles("Normal").Font.Name
                    .Selection.Font.Bold = False

                    'Call RemoveUnicode(wd)

                    wdTbl.Rows(wdTbl.Rows.Count).Select()
                    wdTbl.Rows(1).Select()
                    .Selection.Cells.Merge()
                End Try


                With .Selection.Cells(1)
                    .WordWrap = True
                    .FitText = False
                End With


                wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView

                wd.Selection.Font.Size = NormalFontsize

                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

                .Selection.Borders.Enable = False
                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)

                wdSel = .Selection
                With wd.ActiveDocument.Bookmarks
                    .Add(Range:=wdSel.Range, Name:="tempTitle")
                    .ShowHidden = False
                End With

                '20180529 LEE:
                'Check for unicode
                wdTbl.Cell(1, 1).Select()
                wd.Selection.Style = wd.ActiveDocument.Styles("Normal")
                Call RemoveUnicode(wd)
                .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                If boolExcludeTableNumbers Then
                Else
                    Try

                        Try
                            If BOOLTABLELABELSECTION Then
                                With wd.CaptionLabels("Table")
                                    Call ApplyChapterNumber(wd, "Table")
                                End With
                            Else
                                With wd.CaptionLabels("Table")
                                    .IncludeChapterNumber = False
                                End With
                            End If
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try

                        .Selection.InsertCaption(Label:="Table", TitleAutoText:="", Title:="", Position:=Microsoft.Office.Interop.Word.WdCaptionPosition.wdCaptionPositionAbove)
                    Catch ex As Exception

                        Try


                            wd.CaptionLabels.Add(Name:="Table")
                            With wd.CaptionLabels("Table")
                                .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleArabic
                                .IncludeChapterNumber = False
                            End With

                            Try
                                If BOOLTABLELABELSECTION Then
                                    Call ApplyChapterNumber(wd, "Table")
                                Else
                                    With wd.CaptionLabels("Table")
                                        .IncludeChapterNumber = False
                                    End With
                                End If
                            Catch ex1 As Exception
                                var1 = ex1.Message
                                var1 = var1
                            End Try

                            Try
                                .Selection.InsertCaption(Label:="Table", TitleAutoText:="", Title:="", Position:=Microsoft.Office.Interop.Word.WdCaptionPosition.wdCaptionPositionAbove)
                            Catch ex2 As Exception
                                strM = "There is a problem inserting the Word Table caption for table:" & ChrW(10) & ChrW(10) & strTName
                                MsgBox(strM, vbInformation, "Problem...")
                            End Try

                        Catch ex1 As Exception
                            strM = "There is a problem establishing the Word Table caption for table:" & ChrW(10) & ChrW(10) & strTName
                            MsgBox(strM, vbInformation, "Problem...")
                        End Try




                    End Try

                    'enter nonbreaking space
                    Call NBSP(wd, True)

                    Dim strTS As String = ""

                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

                    If boolExcludeTableNumbers Then
                        strTS = ""
                    Else
                        strTS = .Selection.Text
                    End If

                    '20170626 LEE: Still getting a copy/paste error here (see Try-Catch below)
                    'Selection is getting cut, but for some reason not getting placed in the clipboard
                    'Try pausing to allow Word to catch up

                    Try
                        '.cut() is funny in word 2000
                        Pause(0.1)
                        .Selection.Cut()
                        Pause(0.1)

                    Catch ex As Exception
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        Pause(0.1)
                        .Selection.Cut()
                        Pause(0.1)
                    End Try

                    .Selection.Style = .ActiveDocument.Styles("Normal")

                    '20180525 LEE:
                    .Selection.Font.Name = .ActiveDocument.Styles("Normal").Font.Name
                    .Selection.Font.Bold = False

                End If

                ' ''now go back and delete that extra line left by the original table caption
                '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                '.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                boolInTable = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable)
                If boolInTable Then
                Else
                    .Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                End If
                '.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteDefault)
                '.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteRTF)



                If boolExcludeTableNumbers Then
                Else

                    Try
                        .Selection.PasteSpecial(Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteRTF)
                        Try
                            .Selection.Style = .ActiveDocument.Styles("Caption")
                        Catch ex As Exception
                            strM = "A problem occurred when attempting to apply the 'Caption' style to the Table heading."
                            strM = strM & ChrW(10) & "Please inspect your StudyDoc Word Template to ensure the Word style 'Caption' exists."
                            strM = strM & ChrW(10) & ChrW(10) & ex.Message
                            MsgBox(strM, vbInformation, "Problem...")
                        End Try

                    Catch ex As Exception
                        strM = "A problem occurred when attempting to PasteSpecial the Table Heading Caption."
                        strM = strM & ChrW(10) & ChrW(10) & ex.Message
                        MsgBox(strM, vbInformation, "Problem...")

                    End Try


                End If

                .Selection.Rows.HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast
                .Selection.Rows.Height = 32 'InchesToPoints(0.4)

                If boolExcludeTableNumbers Then
                Else

                    'don't do this
                    'user styles should be maintained
                    '.Selection.Font.Bold = True
                    '.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)

                    Dim numLI As Single 'Left Indent

                    If boolTab Then

                        'For appendix, attachment and table, set left indent depending on font size and font type
                        'Ricerca has Arial 12, which results in crowded caption and label
                        'The current selection is 'caption' style

                        '20170717 LEE:  Hmmm. I think the following should be ommitted and use the document table caption style instead

                        'numLI = ReturnLeftIndent(wd, False, False, True)

                        '.Selection.ParagraphFormat.TabStops.Add(Position:=numLI, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                        'With .Selection.ParagraphFormat
                        '    .LeftIndent = numLI 'InchesToPoints(0.75)
                        '    .SpaceBefore = 0
                        '    .SpaceBeforeAuto = False
                        '    .SpaceAfter = 0
                        '    .SpaceAfterAuto = False
                        '    .FirstLineIndent = -numLI 'InchesToPoints(-0.75)
                        'End With

                        '20180529 LEE:
                        'if tab, ensure that caption has hanging indent
                        Call ApplyTableCaption(wd)


                        If Len(strCaptionTrailer) = 0 Then
                        Else
                            .Selection.TypeText(Text:=strCaptionTrailer)
                        End If
                        .Selection.TypeText(Text:=vbTab)

                    Else

                        If Len(strCaptionTrailer) = 0 Then
                        Else
                            .Selection.TypeText(Text:=strCaptionTrailer)
                        End If
                        If boolSpace Then
                            .Selection.TypeText(Text:=" ")
                        Else
                            .Selection.TypeText(Text:=ChrW(11))
                        End If

                    End If



                End If

                'replace deg C
                var8 = ChrW(176) & "C"
                strFind = "deg C"
                strTitle = Replace(strTitle, strFind, var8, 1, -1, CompareMethod.Text)
                strFind = "degC"
                strTitle = Replace(strTitle, strFind, var8, 1, -1, CompareMethod.Text)

                'replace hyphens
                'don't do because of copy-paste issue in Word


                If boolExcludeTableTitles Then
                Else
                    If boolTab Then
                        '.Selection.TypeText(Text:=strTitle & ChrW(10))
                        '20170726 LEE: adding line feed to title makes title way too large
                        .Selection.TypeText(Text:=strTitle)
                    Else
                        .Selection.TypeText(Text:=strTitle)
                    End If

                End If

                '
                'record total label as strTitle
                Try

                    Dim posS, posE

                    If boolTab Then
                        posE = .Selection.Start - 1
                    Else
                        posE = .Selection.Start ' - 1
                    End If

                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="tempTitle")
                    posS = .Selection.Start

                    rng1 = .ActiveDocument.Range(Start:=posS, End:=posE)
                    '20170713 LEE:
                    'Note: If caption type is "Include Section Number", captured string dash is represented as 'record separator' (chrw(30))
                    'e.g.: 'Table 10-4	Analytical Batch 9 Quality Control Evaluation Data for BCX7343 in Rat K3EDTA Plasma'
                    'is recorded as 'Table 104	Analytical Batch 9 Quality Control Evaluation Data for BCX7343 in Rat K3EDTA Plasma'
                    '    there is actually a chrw(30) in between 10 and 4
                    'this must be accounted for in Function ReturnLabelPosition located in modConserved
                    strTitle = rng1.Text

                Catch ex As Exception

                End Try

                '20180525 LEE:
                'Unicode stuff
                wdTbl.Rows(1).Select()
                var1 = .ActiveDocument.Styles("Caption").Font.Name
                var2 = .ActiveDocument.Styles("Caption").Font.Bold
                .Selection.Font.Name = var1
                .Selection.Font.Bold = var2

                'return to old spot
                .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="temptitle")
                Try
                    .ActiveDocument.Bookmarks.Item("temptitle").Delete()
                Catch ex As Exception
                End Try

                ''''''wdd.visible = True

                If numRows = 0 Then 'skip
                Else
                    'make header rows
                    wd.Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.SelectRow()
                    .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=numRows - 1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                    .Selection.Rows.HeadingFormat = True

                End If


            End With

            '20180525 LEE:
            'while here, replace any degC
            wdTbl.Select()
            rng1 = wd.Selection.Range
            For Count1 = 1 To 2

                Select Case Count1
                    Case 1
                        str1 = "degC"
                        str2 = ChrW(176) & "C"
                    Case 2
                        str1 = "deg C"
                        str2 = ChrW(176) & fNBSP() & "C"
                End Select

                rng1.Find.Execute(FindText:=str1, ReplaceWith:=str2, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                'With rng1.Find
                '    .ClearFormatting()
                '    .Text = str1
                '    .Replacement.ClearFormatting()
                '    .Replacement.Text = str2
                '    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace, Forward:=True, Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue)
                'End With
            Next

        Catch ex As Exception

            '20180524 LEE:
            'must ensure table is selected
            wdTbl.Select()
            wdTbl.Rows(wdTbl.Rows.Count).Select()
            ' wdTbl.Rows(1).Select()

            wd.Visible = True

            strM = "There was a problem preparing Table Caption for table:" & ChrW(10) & ChrW(10) & strTName
            MsgBox(strM, vbInformation, "Problem...")

        End Try


        strTName = strTitle

end1:

        wd.Visible = boolV

        var2 = 1

    End Sub

    Sub MVValidationDesignDescr(ByVal wd As Microsoft.Office.Interop.Word.Application)

        'update 'lists' to 'list' if anal run summary tables > 1
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim dv As system.data.dataview
        Dim strF As String
        Dim var8, var1, var2, var3, var4
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strFind As String
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim varReplace
        Dim pos1 As Int64
        Dim pos2 As Int64

        ''''''''''''wdd.visible = True

        'clear temps
        Try
            wd.ActiveDocument.Bookmarks("Temp1").Delete()
        Catch ex As Exception

        End Try
        Try
            wd.ActiveDocument.Bookmarks("Temp2").Delete()
        Catch ex As Exception

        End Try

        '''''''''wdd.visible = True

        wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
        pos1 = wd.Selection.Start
        mySel = wd.Selection
        With wd.ActiveDocument.Bookmarks
            .Add(Range:=mySel.Range, Name:="Temp1")
        End With

        wd.Selection.Find.ClearFormatting()
        With wd.Selection.Find
            .Text = "^p"
            .Forward = True
            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        wd.Selection.Find.Execute()
        wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
        pos2 = wd.Selection.Start
        With wd.ActiveDocument.Bookmarks
            .Add(Range:=mySel.Range, Name:="Temp2")
        End With

        'select range
        wd.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
        wd.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=pos2 - pos1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

        mySel = wd.Selection


        strFind = "lists"
        varReplace = "list"

        tbl1 = tblAnalytesHome
        dv = frmH.dgvAnalyticalRunSummary.DataSource
        tbl2 = dv.ToTable
        strF = "IsIntStd = 'No'"
        Erase rows1
        rows1 = tbl1.Select(strF)
        int1 = rows1.Length
        If int1 > 1 Then 'change 'lists' to 'list'
            With mySel.Range.Find
                .ClearFormatting()
                .Text = strFind
                With .Replacement
                    .ClearFormatting()
                    .Text = varReplace
                End With
                .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Format:=True, MatchCase:=False, MatchWholeWord:=True)
            End With
        End If

        ''''''wdd.visible = True

        wd.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

        'clear temps
        Try
            wd.ActiveDocument.Bookmarks("Temp1").Delete()
        Catch ex As Exception

        End Try
        Try
            wd.ActiveDocument.Bookmarks("Temp2").Delete()
        Catch ex As Exception

        End Try

    End Sub

    Sub GuWuApprovalStatement01(ByVal wd, ByVal charSectionName)
        Dim var1, var2, var3, var8, var9

        With wd
            'don't need this here

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'var2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'var1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            ''call searchreplace(wd, charSectionName) 
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

            Exit Sub

            'add signature block(s)
            Dim strF As String
            Dim strS As String
            Dim dtbl As System.Data.DataTable
            Dim rows() As DataRow
            Dim int1 As Short
            Dim intRows As Short
            Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
            Dim Count1 As Short
            Dim intRow As Short
            Dim intCol As Short
            Dim num1, num2
            Dim str1 As String
            Dim tblC As System.Data.DataTable
            Dim rowsC() As DataRow
            Dim intC As Short
            Dim strC As String
            Dim strDeg As String
            Dim intPers As Short

            dtbl = tblContributingPersonnel
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDESIGONTABLEPAGE = -1"
            strS = "INTORDER ASC"
            rows = dtbl.Select(strF, strS)
            intPers = rows.Length

            tblC = tblCorporateNickNames
            str1 = NZ(frmH.cbxSubmittedBy.Text, "GA")
            strF = "CHARNICKNAME = '" & str1 & "'"
            rowsC = tblC.Select(strF)
            var1 = rowsC(0).Item("ID_TBLCORPORATENICKNAMES")

            tblC = tblCorporateAddresses
            strF = "ID_TBLCORPORATENICKNAMES = " & var1 & " AND BOOLINCLUDEINTITLE = -1"
            Erase rowsC
            strS = "ID_TBLADDRESSLABELS ASC"
            rowsC = tblC.Select(strF, strS)
            intC = rowsC.Length
            strC = ""
            For Count1 = 0 To intC - 1
                var1 = NZ(rowsC(Count1).Item("CHARVALUE"), "No")
                If Count1 = 0 And Count1 = intC - 1 Then
                    strC = var1
                ElseIf Count1 <> intC - 1 Then
                    strC = strC & var1 & ChrW(10)
                ElseIf Count1 > 0 And Count1 = intC - 1 Then
                    strC = strC & var1
                End If
            Next

            If intPers = 0 Then 'ignore
            Else

                If intPers < 2 Then
                    intRows = 1
                ElseIf intPers < 4 Then
                    intRows = 2
                ElseIf intPers < 6 Then
                    intRows = 3
                ElseIf intPers < 8 Then
                    intRows = 4
                End If

                wd.Selection.TypeParagraph()
                wd.Selection.TypeParagraph()
                wd.Selection.TypeParagraph()
                wd.Selection.TypeParagraph()
                wd.Selection.TypeParagraph()
                wrdSelection = wd.Selection()

                .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intRows, NumColumns:=3, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                .selection.tables.item(1).select()
                Call GlobalTableParaFormat(wd)


                num1 = .Selection.Tables.item(1).Columns.item(1).width * 3
                .Selection.Tables.item(1).Rows.AllowBreakAcrossPages = False
                .Selection.Tables.item(1).Select()
                .Selection.Rows.AllowBreakAcrossPages = False

                With .Selection 'remove initial borders
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                End With

                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                'size columns
                '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)

                For Count1 = 1 To 3
                    Select Case Count1
                        Case 1
                            num2 = 0.45 * num1
                        Case 2
                            num2 = 0.1 * num1
                        Case 3
                            num2 = 0.45 * num1
                    End Select
                    .Selection.Tables.item(1).Columns.item(Count1).PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                    .Selection.Tables.item(1).Columns.item(Count1).PreferredWidth = num2
                Next

                'enter personnel
                intRow = 1
                intCol = 1
                For Count1 = 0 To intPers - 1
                    If Count1 = 0 Then
                    ElseIf IsEven(Count1) Then
                        intRow = intRow + 1
                        intCol = 1
                    Else
                        intCol = intCol + 2
                    End If

                    .Selection.Tables.item(1).Cell(intRow, intCol).Select()
                    .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    strDeg = NZ(rows(Count1).Item("CHARCPDEGREE"), "")
                    If Len(strDeg) = 0 Then
                        str1 = rows(Count1).Item("CHARCPNAME") & " / Date"
                    Else
                        str1 = rows(Count1).Item("CHARCPNAME") & ", " & strDeg & " / Date"
                    End If
                    str1 = str1 & ChrW(10)
                    str1 = str1 & rows(Count1).Item("CHARCPTITLE")
                    str1 = str1 & ChrW(10) & strC & ChrW(10) & ChrW(10) & ChrW(10) & ChrW(10)
                    .Selection.TypeText(str1)
                Next

                'move to next row below table
                .selection.Tables.item(1).select()
                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)

            End If

        End With

    End Sub


    Sub GuWuPreQAStatement01(ByVal wd, ByVal charSectionName)
        Dim var1, var2, var3, var8, var9
        Dim strFind As String
        Dim mySel As Microsoft.Office.Interop.Word.Selection

        With wd


            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            var2 = .Selection.Bookmarks.item("Temp2").Start
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            var1 = .Selection.Bookmarks.item("Temp1").Start
            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            'search for term [INSERTQATABLE]
            strFind = "[INSERTQATABLE]"
            mySel = wd.selection


            With mySel.Find
                .ClearFormatting()
                .Text = strFind
                .Execute()
            End With
            .Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            'now insert QA Table
            Call QATable(wd)

            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

            ''enter signature block
            'Dim strF As String
            'Dim strS As String
            'Dim dtbl as System.Data.DataTable
            'Dim rows() As DataRow
            'Dim int1 As Short
            'Dim intRows As Short
            'Dim wrdSelection As Microsoft.Office.Interop.Word.selection
            'Dim Count1 As Short
            'Dim intRow As Short
            'Dim intCol As Short
            'Dim num1, num2
            'Dim str1 As String
            'Dim tblC as System.Data.DataTable
            'Dim rowsC() As DataRow
            'Dim intC As Short
            'Dim strC As String
            'Dim strDeg As String

            'dtbl = tblContributingPersonnel
            'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARCPROLE = 'QA Representative'"
            'strS = "INTORDER ASC"
            'rows = dtbl.Select(strF, strS)
            'int1 = rows.Length

            'If int1 = 0 Then 'ignore
            'Else

            '    If int1 < 2 Then
            '        intRows = 1
            '    ElseIf int1 < 4 Then
            '        intRows = 2
            '    ElseIf intRows < 6 Then
            '        intRows = 3
            '    ElseIf intRows < 8 Then
            '        intRows = 4
            '    End If

            '    wd.Selection.TypeParagraph()
            '    wd.Selection.TypeParagraph()
            '    wd.Selection.TypeParagraph()
            '    wrdSelection = wd.Selection()

            '    .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intRows, NumColumns:=3, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)

            '    num1 = .Selection.Tables.item(1).Columns.item(1).width * 3
            '    .Selection.Tables.item(1).Rows.AllowBreakAcrossPages = False
            '    .Selection.Tables.item(1).Select()
            '    .Selection.Rows.AllowBreakAcrossPages = False

            '    With .Selection 'remove initial borders
            '        .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '        .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '        .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '        .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '        .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '        .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '        .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '        .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '    End With

            '    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
            '    'size columns
            '    '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)

            '    For Count1 = 1 To 3
            '        Select Case Count1
            '            Case 1
            '                num2 = 0.45 * num1
            '            Case 2
            '                num2 = 0.1 * num1
            '            Case 3
            '                num2 = 0.45 * num1
            '        End Select
            '        .Selection.Tables.item(1).Columns.item(Count1).PreferredWidthType = Microsoft.Office.Interop.Word.wdpreferredwidthtype.wdPreferredWidthPoints
            '        .Selection.Tables.item(1).Columns.item(Count1).PreferredWidth = num2
            '    Next

            '    'enter personnel
            '    intRow = 1
            '    intCol = 1
            '    For Count1 = 0 To int1 - 1
            '        If Count1 = 0 Then
            '        ElseIf IsEven(Count1) Then
            '            intRow = intRow + 1
            '            intCol = 1
            '        Else
            '            intCol = intCol + 2
            '        End If

            '        .Selection.Tables.item(1).Cell(intRow, intCol).Select()
            '        .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle

            '        strDeg = NZ(rows(Count1).Item("CHARCPDEGREE"), "")
            '        If Len(strDeg) = 0 Then
            '            str1 = rows(Count1).Item("CHARCPNAME") & " / Date"
            '        Else
            '            str1 = rows(Count1).Item("CHARCPNAME") & ", " & strDeg & " / Date"
            '        End If
            '        str1 = str1 & ChrW(10)
            '        str1 = str1 & rows(Count1).Item("CHARCPTITLE")
            '        'str1 = str1 & ChrW(10) & strC & ChrW(10) & ChrW(10) & ChrW(10) & ChrW(10)
            '        str1 = str1 & ChrW(10) & ChrW(10) & ChrW(10)
            '        .Selection.TypeText(str1)
            '    Next

            '    .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)

            'End If

        End With

    End Sub


    Sub SignatureSearch(ByVal wd As Microsoft.Office.Interop.Word.Application)

        'the routine will create role-base signature items
        Dim Count1 As Short
        Dim Count2 As Short
        Dim int1, int2, int3 As Short
        Dim str1, str2, str3 As String
        Dim strF As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim intRows As Short
        Dim strFC As String
        Dim dv As system.data.dataview
        Dim tbl1 As System.Data.DataTable

        strF = "id_tblStudies = " & id_tblStudies
        tbl = tblContributingPersonnel
        dv = New DataView(tbl)
        tbl1 = dv.ToTable("a", True, "CHARCPROLE") 'return distinct recordset of roles
        intRows = tbl1.Rows.Count

        If intRows = 0 Then
            GoTo end1
        End If

        'make signature blocks with title and company
        For Count1 = 0 To intRows - 1
            str1 = NZ(tbl1.Rows.Item(Count1).Item("CHARCPROLE"), "")
            str2 = "SignatureBlockTitleComp"
            If Len(str1) = 0 Then
            Else
                'build field code
                strFC = "[" & str1 & "SIGBLOCKTITLECOMP]"
                'initiate search

                Call SearchReplaceSigs(wd, strFC, str1, str2)

            End If
        Next

        'make signature blocks with title
        For Count1 = 0 To intRows - 1
            str1 = NZ(tbl1.Rows.Item(Count1).Item("CHARCPROLE"), "")
            str2 = "SignatureBlockTitle"
            If Len(str1) = 0 Then
            Else
                'build field code
                strFC = "[" & str1 & "SIGBLOCKTITLE]"
                'initiate search

                Call SearchReplaceSigs(wd, strFC, str1, str2)

            End If
        Next

        'make signature blocks with title inline
        For Count1 = 0 To intRows - 1
            str1 = NZ(tbl1.Rows.Item(Count1).Item("CHARCPROLE"), "")
            str2 = "SignatureBlockTitleInline"
            If Len(str1) = 0 Then
            Else

                'build field code
                strFC = "[" & str1 & "SIGBLOCKTITLEINLINE]"
                'initiate search

                Call SearchReplaceSigs(wd, strFC, str1, str2)

            End If
        Next

        'make signature blocks with name only
        For Count1 = 0 To intRows - 1
            str1 = NZ(tbl1.Rows.Item(Count1).Item("CHARCPROLE"), "")
            str2 = "SignatureBlock"
            If Len(str1) = 0 Then
            Else
                'build field code
                strFC = "[" & str1 & "SIGBLOCK]"
                'initiate search

                Call SearchReplaceSigs(wd, strFC, str1, str2)

            End If
        Next

        'return only signature lines
        For Count1 = 0 To intRows - 1
            str1 = NZ(tbl1.Rows.Item(Count1).Item("CHARCPROLE"), "")
            str2 = "Name"
            If Len(str1) = 0 Then
            Else
                'build field code
                strFC = "[" & str1 & "NAME]"
                'initiate search

                Call SearchReplaceSigs(wd, strFC, str1, str2)

            End If
        Next

        'return only signature lines in a column
        For Count1 = 0 To intRows - 1
            str1 = NZ(tbl1.Rows.Item(Count1).Item("CHARCPROLE"), "")
            str2 = "NameColumn"
            If Len(str1) = 0 Then
            Else
                'build field code
                strFC = "[" & str1 & "NAMECOLUMN]"
                'initiate search

                Call SearchReplaceSigs(wd, strFC, str1, str2)

            End If
        Next

        'return reference style signature
        For Count1 = 0 To intRows - 1
            str1 = NZ(tbl1.Rows.Item(Count1).Item("CHARCPROLE"), "")
            str2 = "RefStyle_01"
            If Len(str1) = 0 Then
            Else
                'build field code
                strFC = "[" & str1 & "REFSTYLE_01]"
                'initiate search

                Call SearchReplaceSigs(wd, strFC, str1, str2)

            End If
        Next

        'return Wyeth signature block w/o role
        For Count1 = 0 To intRows - 1
            str1 = NZ(tbl1.Rows.Item(Count1).Item("CHARCPROLE"), "")
            str2 = "WSigBlock"
            If Len(str1) = 0 Then
            Else
                'build field code
                strFC = "[" & str1 & "WSIGBLOCK]"
                'initiate search

                Call SearchReplaceSigs(wd, strFC, str1, str2)

            End If
        Next

        int1 = 1 'debugginn

        'return Wyeth signature block role
        For Count1 = 0 To intRows - 1
            str1 = NZ(tbl1.Rows.Item(Count1).Item("CHARCPROLE"), "")
            str2 = "WSigBlockRole"
            If Len(str1) = 0 Then
            Else
                'build field code
                strFC = "[" & str1 & "WSIGBLOCKROLE]"
                'initiate search

                Call SearchReplaceSigs(wd, strFC, str1, str2)

            End If
        Next


end1:

    End Sub

    Sub WSignatureBlockRole(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strRole As String)
        Dim str1 As String
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tblCP As System.Data.DataTable
        Dim var1, var2, var3, var4, var5
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim strNick As String
        Dim int1 As Short
        Dim numRT

        strNick = frmH.cbxSubmittedBy.Text
        tblCP = tblContributingPersonnel


        With wd

            numRT = WorkingPageWidth(wd)

            ''enter PI info, if applicable
            str1 = "id_tblStudies = " & id_tblStudies & " AND CHARCPROLE = '" & strRole & "'"
            dr1 = tblCP.Select(str1)
            ct1 = dr1.Length
            Dim intRows As Short
            Dim intCols As Short
            Dim numSec1 As Integer
            Dim numSec2 As Integer
            Dim numSec3 As Integer
            Dim numSec4 As Integer
            Dim numSec5 As Integer
            Dim numSec6 As Integer
            Dim numSec7 As Integer
            Dim boolNew As Boolean
            Dim intTRows As Integer

            intRows = dr1.Length
            'calculate number of table rows
            intTRows = (intRows * 2) ' + intRows - 1

            '****
            If intRows = 0 Then
                intStartTable = 1
                var3 = .Selection.Text
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                var1 = "'" & var3 & "' not configured in Contributing Personnel Table"
                .Selection.TypeText(Text:=var1)
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
            Else
                intStartTable = 2
                'Select Case intRows
                '    Case 1
                '        intCols = 2
                '        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow
                '    Case Is > 1
                '        intCols = 3
                '        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed
                'End Select

                intCols = 7
                var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow

                wrdSelection = wd.Selection()
                .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intTRows, NumColumns:= _
                intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=var1)
                '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                .Selection.Tables.Item(1).Select()
                Call GlobalTableParaFormat(wd)

                .Selection.Rows.AllowBreakAcrossPages = False
                .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop


                var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed
                .Selection.Tables.Item(1).AutoFitBehavior(var1)

                With .Selection 'remove initial borders
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                End With
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                With wd.Selection.Tables.Item(1)
                    .LeftPadding = 1 'InchesToPoints(0.02)
                    .RightPadding = 1 'InchesToPoints(0.02)
                    '.WordWrap = True
                    '.FitText = False
                End With


                numSec1 = numRT * 0.35
                numSec2 = numRT * 0.025
                numSec3 = numRT * 0.15
                numSec4 = numRT * 0.025
                numSec5 = numRT * 0.15
                numSec6 = numRT * 0.025
                numSec7 = numRT * 0.275

                .Selection.Columns.Item(1).Width = numSec1
                .Selection.Columns.Item(2).Width = numSec2
                .Selection.Columns.Item(3).Width = numSec3
                .Selection.Columns.Item(4).Width = numSec4
                .Selection.Columns.Item(5).Width = numSec5
                .Selection.Columns.Item(6).Width = numSec6
                .Selection.Columns.Item(7).Width = numSec7

                'format columns
                boolNew = False
                int1 = CInt(intRows / 2)
                Count2 = 0
                Count3 = 1

                'format column 7
                .Selection.Tables.Item(1).Cell(1, 7).Select()
                .Selection.SelectColumn()
                .Selection.ParagraphFormat.TabStops.Add(Position:=250, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                .Selection.Tables.Item(1).Cell(1, 1).Select()
                Count2 = 0
                For Count1 = 0 To intRows - 1
                    Count2 = Count2 + 1
                    'enter item in column 7
                    If IsEven(Count2) Then
                        .Selection.Tables.Item(1).Rows.Item(Count2).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                        .Selection.Tables.Item(1).Rows.Item(Count2).Height = 72 'InchesToPoints(1)

                        .Selection.Tables.Item(1).Cell(Count2, 7).Select()
                        .Selection.Font.Size = .Selection.Font.Size - 2
                        .Selection.TypeText(Text:="Other Time Zone")
                        .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                        .Selection.TypeText(Text:=vbTab)

                    Else
                        .Selection.Tables.Item(1).Cell(Count2, 7).Select()
                        .Selection.Font.Size = .Selection.Font.Size - 2
                        '.Selection.TypeText(Text:="Eastern Time Zone")
                        .Selection.TypeText(Text:=LTimeZone)
                        .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                        .Selection.TypeText(Text:=vbTab)

                    End If

                    Count2 = Count2 + 1

                    If IsEven(Count2) Then
                        .Selection.Tables.Item(1).Rows.Item(Count2).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                        .Selection.Tables.Item(1).Rows.Item(Count2).Height = 72 'InchesToPoints(1)

                        .Selection.Tables.Item(1).Cell(Count2, 7).Select()
                        .Selection.Font.Size = .Selection.Font.Size - 2
                        .Selection.TypeText(Text:="Other Time Zone")
                        .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                        .Selection.TypeText(Text:=vbTab)

                    Else
                        .Selection.Tables.Item(1).Cell(Count2, 7).Select()
                        .Selection.Font.Size = .Selection.Font.Size - 2
                        '.Selection.TypeText(Text:="Eastern Time Zone")
                        .Selection.TypeText(Text:=LTimeZone)
                        .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                        .Selection.TypeText(Text:=vbTab)

                    End If

                    var1 = dr1(Count1).Item("charCPName")
                    var2 = NZ(dr1(Count1).Item("charCPDegree"), "")
                    var4 = NZ(dr1(Count1).Item("charCPRole"), "")
                    var5 = NZ(dr1(Count1).Item("charCPTitle"), "")
                    If Len(var2) = 0 Then
                        var3 = var1 & Chr(10) & var5
                    Else
                        var3 = var1 & ", " & var2 & Chr(10) & var5
                    End If
                    If Len(var4) = 0 Then
                    Else
                        var3 = var3 & ChrW(10) & var4
                    End If

                    .Selection.Tables.Item(1).Cell(Count2, 1).Select()
                    .Selection.TypeText(Text:=var3)
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    .Selection.Tables.Item(1).Cell(Count2, 3).Select()
                    .Selection.Font.Size = .Selection.Font.Size - 2
                    str1 = "Date" & ChrW(10) & "(dd-mmm-yyyy)"
                    .Selection.TypeText(Text:=str1)
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    .Selection.Tables.Item(1).Cell(Count2, 5).Select()
                    .Selection.Font.Size = .Selection.Font.Size - 2
                    str1 = "Time" & ChrW(10) & "(24 hour clock)"
                    .Selection.TypeText(Text:=str1)
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle


                    If IsEven(Count2) Then
                        'Count2 = Count2 + 1
                    Else
                        'Count2 = Count2 + 1
                    End If

                Next


            End If
            '****

        End With

    End Sub


    Sub NoSplitTable(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal DoTable As Boolean, ByVal intSRow As Short, ByVal intSplitRows As Short, ByVal boolAutoFit As Boolean, ByVal boolSmallFont As Boolean)

        Dim boolCell
        Dim cell1
        Dim cell2
        Dim pg1
        Dim pg2
        Dim pgT
        Dim row1
        Dim row2
        Dim var1, var2
        Dim var3, var4
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim ctSplitRows As Short
        Dim boolTableEnd As Boolean
        Dim char1
        Dim char2
        Dim char3
        Dim myRReturn As Microsoft.Office.Interop.Word.Range
        Dim myr As Microsoft.Office.Interop.Word.Range
        Dim myR1 As Microsoft.Office.Interop.Word.Range
        Dim arrRealLegend(3, UBound(arr, 2))
        Dim int1 As Short
        Dim int2 As Short
        Dim bool As Boolean
        Dim str1 As String
        Dim fonts
        Dim rows1 As Long
        Dim intCell1 As Short
        Dim intCell2 As Short
        Dim intRows As Short
        Dim intRow As Short
        Dim intRow1 As Short
        Dim intRow2 As Short
        Dim intRowStart As Short

        Dim Count20 As Short
        Dim Count30 As Short

        Dim boolGo As Boolean
        Dim intIncr As Short

        Dim boolFound As Boolean
        Dim boolDoLast As Boolean

        intSplitRows = intSplitRows + 1 'account for 'continued on next pate

        'If frmH.rbTable.Checked Then
        'Else
        '    Call SplitTableOld(wd, ctHdRows, ctLegend, arr, strT, DoLegend, DoTable, intSRow, intSplitRows, boolAutoFit, boolSmallFont)
        '    GoTo end2
        'End If

        'wdd.visible = True

        intRows = wd.Selection.Tables.Item(1).Rows.Count

        If boolSmallFont Then
            '''''''''wdd.visible = True
            wd.Selection.Font.Size = NormalFontsize - 1
        End If

        'first make headers one fontsize smaller
        wd.Selection.Tables.Item(1).Cell(2, 1).Select()
        wd.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=ctHdRows - 1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        wd.Selection.SelectRow()
        wd.Selection.Font.Size = NormalFontsize - 1

        wd.Selection.Tables.Item(1).Select()
        If boolSmallFont Then
            wd.Selection.Font.Size = NormalFontsize - 1
        End If

        boolCell = True
        'On Error Resume Next
        cell1 = 1
        row1 = 1
        row2 = 1
        Count3 = 0 'iteration counter
        ctRealLegend = 0

        '''''''''wdd.visible = True


        If wd.Selection.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
            intIncr = 15
        Else
            intIncr = 30
        End If


        boolSplitTable = False
        With wd

            If ctLegend = 0 Then
                GoTo end1
            End If

            boolTableEnd = False
            Count20 = 0
            Count30 = 0
            boolDoLast = False

            intRows = .Selection.Tables.Item(1).Rows.Count
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
            .Selection.Tables.Item(1).Cell(intRows, 1).Select()
            pg2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            If pg1 = pg2 Then 'no need to evaluate
                .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToNext, Count:=1, Name:="")
                int1 = .Selection.Tables.Item(1).Rows.Count
                intRows = .Selection.Tables.Item(1).Rows.Count
                .Selection.Tables.Item(1).Cell(int1, 1).Select()
                boolDoLast = True
            End If

            'begin looking for next page
            intRow = 1
            intRowStart = 1
            Count20 = 0

            Do Until intRow = intRows
                Count20 = Count20 + 1
                str1 = strT & vbCr & "Formatting legend..." & Count20
                frmH.lblProgress.Text = str1
                frmH.lblProgress.Refresh()

                intRows = .Selection.Tables.Item(1).Rows.Count

                ctRealLegend = 0

                If boolDoLast Then
                    intRow = intRows
                    intRow1 = intRow
                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                Else
                    intRow = intRow + intIncr
                    If intRow > intRows Then
                        intRow = intRows
                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                        pgT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        If pgT = pg2 Then
                            boolDoLast = True

                        Else
                        End If

                    End If
                    pgT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                    .Selection.Tables.Item(1).Cell(intRows, 1).Select()
                    pg2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                    pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                    If pg1 = pg2 Then
                        boolDoLast = True
                    Else
                        If pgT <> pg1 Then 'intincr is too big
                            If pgT > pg1 Then
                            Else
                                Do Until pgT = pg1
                                    intRow = intRow - 5
                                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                                    pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                                Loop

                            End If

                        End If

                    End If

                    pgT = pg1
                    intRow1 = intRow
                End If

                If boolDoLast Then
                Else
                    Do Until pg1 <> pgT
                        intRow1 = intRow1 + 1
                        .Selection.Tables.Item(1).Cell(intRow1, 1).Select()
                        pgT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                    Loop
                    'intRowStart = intRows + ctLegend
                    intRow1 = intRow1 - 1
                    'intRow1 = intRow1 - ctLegend - 1
                    intRow2 = intRow1
                    'intRow1 = intRow2 - 1

                    '''''''''wdd.visible = True

                    .Selection.Tables.Item(1).Cell(intRow2, 1).Select()



                    'format paragraph force next page
                    '.Selection.ParagraphFormat.PageBreakBefore = True
                    intRow2 = intRow2 + ctLegend
                End If
                .Selection.Tables.Item(1).Cell(intRow1, 1).Select()
                'remove any borders
                .Selection.SelectRow()
                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone

                'insert rows for legend
                .Selection.InsertRowsBelow(ctLegend)
                'ensure bottom is not bordered
                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone

                .Selection.Cells.Split(NumRows:=ctLegend, NumColumns:=1, MergeBeforeSplit:=True)
                'ensure bottom is not bordered
                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                If boolSmallFont Then
                    .Selection.Font.Size = .Selection.Font.Size - 1
                End If
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                'format cells
                .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                With .Selection.ParagraphFormat
                    .LeftIndent = 36 'InchesToPoints(0.63)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                End With
                With .Selection.ParagraphFormat
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .FirstLineIndent = -36 'InchesToPoints(-0.63)
                End With

                'ensure top is  bordered
                .Selection.Tables.Item(1).Cell(intRow1 + 1, 1).Select()
                '.Selection.SelectRow()
                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                'With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                '    .LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle
                'End With
                If DoLegend Then 'enter legend without checking
                    ctRealLegend = ctLegend
                    For Count1 = 1 To ctLegend
                        arrRealLegend(1, Count1) = arr(1, Count1)
                        arrRealLegend(2, Count1) = arr(2, Count1)
                        arrRealLegend(3, Count1) = arr(3, Count1)
                    Next
                Else
                    'determine if legend needs to be entered
                    'move back into previous table
                    .Selection.Tables.Item(1).Cell(intRowStart, 1).Select()
                    .Selection.SelectRow()
                    int1 = intRow1 - intRowStart
                    .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                    myr = .Selection.Range

                    '''''''''wdd.visible = True


                    '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2)
                    ctRealLegend = 0

                    DoTable = False
                    For Count1 = 1 To ctLegend

                        var1 = arr(1, Count1)
                        myr = .Selection.Range
                        boolFound = False
                        With myr.Find
                            .ClearFormatting()
                            .MatchCase = True
                            .MatchWholeWord = True
                            .Forward = True
                            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                            .Execute(FindText:=var1)

                            If .Found Then
                                boolFound = True
                            Else
                                boolFound = False
                            End If

                        End With

                        '''''''''wdd.visible = True

                        If boolFound Then 'legend not needed
                            ctRealLegend = ctRealLegend + 1
                            arrRealLegend(1, ctRealLegend) = arr(1, Count1)
                            arrRealLegend(2, ctRealLegend) = arr(2, Count1)
                            arrRealLegend(3, ctRealLegend) = arr(3, Count1)
                        End If
                    Next
                End If

                intRow = intRow1
                If ctLegend = 0 Or ctRealLegend = 0 Then
                    .Selection.Tables.Item(1).Cell(intRow2, 1).Select()
                    If boolDoLast Then
                    Else
                        .Selection.TypeText("Continued on next page")
                        'format align right
                        .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight ' wdAlignParagraphRight
                    End If

                Else

                    '''''''''wdd.visible = True

                    'enter legend
                    For Count1 = 1 To ctRealLegend
                        intRow = intRow + 1
                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                        var1 = arrRealLegend(1, Count1)
                        If arrRealLegend(3, Count1) Then
                            fonts = .Selection.Font.Size
                            .Selection.Font.Superscript = True
                            .Selection.Font.Size = 12
                            .Selection.TypeText(Text:=CStr(var1))
                            .Selection.Font.Superscript = False
                            .Selection.Font.Size = fonts
                        Else
                            .Selection.TypeText(Text:=CStr(var1))
                        End If
                        .Selection.TypeText(Text:=vbTab)
                        .Selection.TypeText(Text:="=")
                        .Selection.TypeText(Text:=vbTab)
                        var2 = arrRealLegend(2, Count1)
                        'var3 = var1 & " = " & var2
                        .Selection.TypeText(Text:=CStr(var2))
                        char2 = .Selection.Start
                        If Count1 = ctRealLegend Then
                            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)
                            'MoveOneCellDown(wd)
                            myR1 = wd.ActiveDocument.Range(Start:=char2, End:=char2)
                        Else
                            '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                        End If
                    Next
                    .Selection.Tables.Item(1).Cell(intRow + 1, 1).Select()
                    'now enter 'continued next page
                    .Selection.Tables.Item(1).Cell(intRow + 1, 1).Select()
                    If boolDoLast Then
                    Else
                        .Selection.TypeText("Continued on next page")
                        'format align right
                        .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight ' wdAlignParagraphRight
                    End If


                    .Selection.Tables.Item(1).Cell(intRow2, 1).Select()
                    'intRow1 = intRow2
                    'intRow = intRow1
                    'intRowStart = intRow
                    'intRows = intRows + ctLegend

                End If

                intRow1 = intRow2
                intRow = intRow1
                intRowStart = intRow

                If boolDoLast Then
                    Exit Do
                End If
                var1 = intRow 'for testing

                With myr.Find
                    .ClearFormatting()
                End With

            Loop

            frmH.lblProgress.Text = strT


end1:


        End With

end2:

        'clear formatting again
        wd.Selection.Find.ClearFormatting()
    End Sub


    Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, ByVal arr As Object, _
                   ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, _
                   ByVal boolSmallFont As Boolean, ByVal boolCarefulSplit As Boolean, ByVal boolFirstAnova As Boolean, _
                   ByVal intTableID As Int64)

        'Sub SplitTable(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, ByVal arr As Object, ByVal strT As String,
        'ByVal DoLegend As Boolean, ByVal DoTable As Boolean, ByVal intSRow As Short, ByVal intSplitRows As Short, ByVal boolAutoFit As Boolean, ByVal boolSmallFont As Boolean)

        '20180715 LEE:
        'depricate boolFirstAnova
        boolFirstAnova = False

        Dim boolV As Boolean = wd.Visible

        'arr()
        '1= Actual string to search in table
        '2= Not used in SplitTable
        '3= True/False to superscript table
        '4= True: Do not look for item in table, but add buffer row to row count.  False: Look for item in table; if found, add buffer row to row count

        '20180701 LEE:
        'Hints from https://wordmvp.com/FAQs/TblsFldsFms/FastTables.htm
        '
        '1. Working in Normal view when you can helps, especially if you turn off “Background repagination” (Tools + Options + General). 
        'Whatever you do, though, tables in Word 2000 and higher are a lot slower in most respects than in Word 97 – an unfortunate by-product of the new table engine
        ' created so that Word tables could be fully HTML-compatible.
        '2. If using Word 2000 and above, select  Table | Table Properties | Options, and turn off the checkbox: “Automatically resize to fit contents”. 
        'As well as slowing tables down considerably, this setting gives (usually) undesirable results, but unfortunately is automatically switched on in all new tables.
        '.
        '.
        '6. If inserting a large amount of text into the document, make sure background spelling and grammar checking are switched off.


        Dim intRows As Long
        Dim intRow As Long
        Dim Count1 As Long
        Dim Count2 As Long
        Dim Count3 As Long
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1, var2, var3, var4
        Dim tbl As Word.Table
        Dim p1 As Long
        Dim p2 As Long
        Dim pT As Long
        Dim intIncr As Integer
        Dim rowStart As Long 'intRow1
        Dim ctP As Long  'Page Counter
        Dim intNextPageRow As Long
        Dim intLastSearchRow As Long
        Dim intFirstSearchRow As Long
        Dim int1 As Long
        Dim int2 As Long
        Dim myr As Word.Range
        Dim myr1 As Word.Range
        Dim ctRealLegend As Integer
        Dim arrRealLegend   'Array of the list of legend rows to be actually shown on a page of the table
        Dim Fonts
        Dim boolIsLastPage As Boolean
        Dim boolFound As Boolean

        Dim intLs As Int32

        ReDim arrRealLegend(3, ctLegend)

        Dim arrPageBreaks(2, 1000) As Short
        Dim intPageCount As Int16
        '20180628 LEE:
        'Note that code does not account for a report with > 1000 pages
        'code has been updated to evaluate ubound(arrPageBreak,2) when an arrPageBreak entry is being made and increase ubound if needed
        Dim tRow As Int16
        Dim intLastLegend As Int16 = 0

        '20180628 LEE:
        'StudyDoc loses communication with Word if a table is large (like AbbVie study M13099DAA SampleAnalysis table)
        'Will begin adding save events every 10 pages to see if this resolves the problem
        Dim intPBMax As Short = 5
        Dim intPB As Short = 0
        Dim intPBreakMax As Short = 50

        Dim chkR As Int16
        Dim chkM As Int16

        Dim strDateTimeStamp As String
        Dim arrPN(1000)
        Dim intPN As Short = 0
        Dim boolPageNum As Boolean = False
        Dim boolDTStamp As Boolean = LBOOLTABLEDTTIMESTAMP
        Dim boolFullPageNum As Boolean = False
        Dim ctLCheck As Short
        Dim wrdSelection As Word.Selection

        Dim intForcedLegends As Short = 0
        Dim intNL As Short = 0

        Dim strM1 As String
        Dim strM2 As String
        Dim strM3 As String
        Dim strM As String

        Dim numTab1 As Single
        Dim numIndent As Single

        '20180808 LEE:
        Select Case intTableID
            Case 17
                numTab1 = 42
                numIndent = 51
            Case Else
                numTab1 = 33 '25
                numIndent = 42
        End Select

        '20180627 LEE: Added rngT logic to speed up processing of large tables
        Dim rngT As Word.Range
        Dim doc As Word.Document

        doc = wd.ActiveDocument

        Dim boolAAF As Boolean

        Dim rngPB As Word.Range

        '20160516 LEE:
        'Summary of wdNormalView, repagination, and SplitTable
        'http://wordribbon.tips.net/T005975_Turning_Off_Background_Repagination.html
        'Repagination in Word 2007, 2010, and 2013 can be turned off only if Word is in wdNormalView (Draft view)
        'Repagination is a proble when preparing large documents (~>100 pages), because repagination is time consuming
        'However, placing Word in wdNormalView disrupts SplitTable's ability to correctly find the end of the page to insert a pagebreak
        'This is verified from the web reference: 
        '    "You should note that the background repagination option is not applicable in Print Layout view. 
        '    This is because Word must automatically repaginate in that view to enable the proper display of information on the screen.
        'For example, if the SampleAnalysis study is loaded and a Sample Concentrations table is prepared for both compounds, 
        'the 2nd compound table will sometimes, but not always, have incorrect page breaks.

        '20160513 LEE
        'added .wdNormalView to improve performance
        '20160516 LEE
        'commented line out. See note above and JIRA Note: Repagination dtd 20160516.
        'wd.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView
        'Microsoft.Office.Interop.Word.WdViewType.wdPrintView

        '20180201 LEE: LI00016: Alturas: Study POP01: Sample Concentrations table
        'If all cmpds (or just two even), the first table is truncated early when inserting page breaks
        'If a code break is entered before SplitTable, the table is page breaking correctly
        'perhaps need to pause
        'Pause(2)
        'No! Actually it's because table is not being autofitted before page breaks are inserted


        If StrComp(LcharSTPage, "[None]", CompareMethod.Text) = 0 Then
            boolPageNum = False
        Else
            boolPageNum = True
        End If

        If StrComp(LcharSTPage, "Page x", CompareMethod.Text) = 0 Then
            boolFullPageNum = False
        Else
            boolFullPageNum = True
        End If

        'strDateTimeStamp = Format(LTableDateTimeStamp, LDateFormat & " HH:mm:ss tt")
        strDateTimeStamp = Format(gdtReportDate, LDateFormat & " HH:mm:ss tt")


        intPageCount = 1
        If ctHdRows + 1 > UBound(arrPageBreaks, 2) Then
            ReDim Preserve arrPageBreaks(2, ctHdRows + 1)
        End If
        arrPageBreaks(1, intPageCount) = ctHdRows + 1

        'wdd.visible = True

        Try
            With wd

                '''wdd.visible = True

                tbl = .Selection.Tables.Item(1)

                '20180701 LEE
                'Implement time saving trick
                'set back to true at end

                With tbl
                    boolAAF = .AllowAutoFit
                    .AllowAutoFit = False
                End With

                Call SpellingOff(doc, False)

                If boolPageNum Then
                    intSplitRows = intSplitRows + 1
                End If
                If boolDTStamp Then
                    intSplitRows = intSplitRows + 1
                End If
                intSplitRows = intSplitRows + 1 'account for Continued on next page
                'if orientation is portrait, add a few rows
                If .Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait Then
                    intSplitRows = intSplitRows + 2 'for safety in case legends rap
                End If
                'intSplitRows = intSplitRows + 2 'for safety in case legends rap

                'first get rows
                intRows = tbl.Rows.Count
                'tbl.Cell(1, 1).Select()
                'p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                rngT = tbl.Cell(1, 1).Range
                p1 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                'tbl.Cell(intRows, 1).Select()
                'pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                rngT = tbl.Cell(intRows, 1).Range
                pT = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                intLs = pT - p1 + 1

                tbl.Cell(1, 1).Select()
                If p1 = pT Then 'no need to evaluate
                    intRows = tbl.Rows.Count
                    boolIsLastPage = True
                Else
                    boolIsLastPage = False
                End If

                'begin looking for next page
                intRow = ctHdRows + 1
                ctP = 0

                If .Selection.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                    intIncr = 20
                Else
                    intIncr = 40
                End If

                intFirstSearchRow = ctHdRows + 1
                boolIsLastPage = False

                chkM = 200
                chkR = 0

                strM1 = frmH.lblProgress.Text

                Do Until intRow > intRows

                    chkR = chkR + 1
                    If chkR > chkM Then
                        Exit Do
                    End If

                    ctP = ctP + 1
                    If ctP > intLs Then
                        str1 = strM1 & vbCr & "Formatting legend..." & ctP & " of at least " & ctP
                    Else
                        str1 = strM1 & vbCr & "Formatting legend..." & ctP & " of at least " & intLs
                    End If

                    frmH.lblProgress.Text = str1
                    frmH.lblProgress.Refresh()


                    'Set p1 (# of first page)
                    'tbl.Cell(intRow, 1).Select()
                    'p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                    rngT = tbl.Cell(intRow, 1).Range
                    p1 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                    'Set pT (# of last page)
                    intRows = tbl.Rows.Count
                    'tbl.Cell(intRows, 1).Select()
                    'pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                    rngT = tbl.Cell(intRows, 1).Range
                    pT = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                    If p1 = pT Then 'don't look for table split
                        intRow = intRows
                        intLastSearchRow = intRows
                        boolIsLastPage = True
                    Else 'look for table split

                        intRow = intRow + intIncr
                        If intRow > intRows Then
                            intRow = intRows
                        End If

                        tbl.Cell(intRow, 1).Select()

                        'check to see if intincr is too big
                        'p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        '20180701 LEE:
                        'Don't use .selection for p1, Word can hang if table is very large (>50 pages)
                        rngT = tbl.Cell(intRow, 1).Range
                        p2 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                        If p2 > p1 Then 'intincr is too big
                            Do Until p2 = p1
                                intRow = intRow - intIncr
                                intIncr = intIncr - 5
                                intRow = intRow + intIncr
                                'tbl.Cell(intRow, 1).Select()
                                'p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                                rngT = tbl.Cell(intRow, 1).Range
                                p2 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                            Loop
                            tbl.Cell(intRow, 1).Select()
                        End If

                        'Find first row of table that ends up on the next page (intRows).
                        'p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        '20180701 LEE:
                        'Don't use .selection for p1, Word can hang if table is very large (>50 pages)
                        rngT = tbl.Cell(intRow, 1).Range
                        p1 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        p2 = p1 + 1

                        Do Until p1 = p2
                            If intRow >= intRows Then
                                Exit Do
                            End If
                            intRow = intRow + 1
                            'tbl.Cell(intRow, 1).Select()
                            'p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                            rngT = tbl.Cell(intRow, 1).Range

                            p1 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                            If intRow = intRows Then
                                Exit Do  'We are at the end of the table
                            End If
                        Loop

                        '*****
                        'wdd.visible = True
                        'now do search to determine new intsplitrows
                        If DoLegend Then
                            tRow = ctLegend  'Plan for every line of the legend table to be in the legend
                        Else

                            ''Set selection to the proposed first page
                            'rngT stuff: don't need to select
                            'tbl.Cell(intFirstSearchRow, 1).Select()
                            '.Selection.SelectRow()
                            intLastSearchRow = intRow - 1

                            'Select rows
                            int1 = intLastSearchRow - intFirstSearchRow  'int1 is the number of rows in the proposed page, minus 1
                            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            'myr = .Selection.Range

                            rngT = doc.Range(Start:=tbl.Rows(intFirstSearchRow).Range.Start, End:=tbl.Rows(intFirstSearchRow + int1).Range.End)
                            myr = rngT
                            'Now look for reference strings in the table body on page

                            tRow = 0 'tRow: Total rows needed for legend and other info at bottom of page

                            For Count1 = 1 To ctLegend  'For each type of legend reference
                                var1 = arr(4, Count1)
                                If NZ(var1, False) Then
                                    tRow = tRow + 1 'No reference string, but add line for it
                                Else

                                    'Search for reference string in table, and add to the tRow counter if found

                                    var1 = arr(1, Count1) 'grab reference string 
                                    'myr = .Selection.Range 'call this again because previous find action seems to wipe out the range
                                    myr = doc.Range(Start:=tbl.Rows(intFirstSearchRow).Range.Start, End:=tbl.Rows(intFirstSearchRow + int1).Range.End)
                                    boolFound = False
                                    With myr.Find

                                        'Search for reference string in the page of the table
                                        .ClearFormatting()
                                        .MatchCase = True
                                        .MatchWholeWord = True
                                        .Forward = True
                                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                                        .Execute(FindText:=var1)

                                        If .Found Then
                                            boolFound = True
                                        Else
                                            boolFound = False
                                        End If

                                    End With

                                    If boolFound Then 'legend needed
                                        tRow = tRow + 1
                                        If StrComp(var1, "NA", CompareMethod.Text) = 0 Then 'debug
                                            var1 = var1
                                        End If
                                    End If
                                End If

                            Next

                        End If

                        'add a row for safety
                        tRow = tRow + 1

                        If boolIsLastPage Then
                        Else
                            tRow = tRow + 1 'take into account 'continued on next page'
                        End If
                        If boolPageNum Then 'Page Number
                            tRow = tRow + 1
                        End If
                        If boolDTStamp Then 'Date/Time Stamp
                            tRow = tRow + 1
                        End If

                        If intTableID = 3 Then
                            If boolSTATSREGR Then
                                tRow = tRow + 1
                            End If
                        End If
                        '*****

                        If boolFirstAnova Then

                        Else
                            'go back up tRows (to account for legend plus other items at bottom of page)
                            intRow = intRow - tRow

                            tbl.Cell(intRow, 1).Select()
                            If boolCarefulSplit Then
                                'tbl.Cell(intRow, 2).Select() 'Select the proposed last row (if leaving room for legend+)
                                'var1 = .Selection.Text
                                rngT = tbl.Cell(intRow, 2).Range
                                var1 = rngT.Text
                                Do Until Len(var1) = 2
                                    ''wdd.visible = True
                                    intRow = intRow - 1
                                    'tbl.Cell(intRow, 2).Select()
                                    'var1 = .Selection.Text

                                    rngT = tbl.Cell(intRow, 2).Range
                                    var1 = rngT.Text
                                Loop

                                intRow = intRow + 1
                                tbl.Cell(intRow, 1).Select()

                                If intTableID = 11 Then
                                    'continue with column 1
                                    var1 = .Selection.Text
                                    Do Until Len(var1) = 2
                                        intRow = intRow - 1
                                        'tbl.Cell(intRow, 1).Select()
                                        'var1 = .Selection.Text

                                        rngT = tbl.Cell(intRow, 2).Range
                                        var1 = rngT.Text
                                    Loop
                                    tbl.Cell(intRow, 1).Select()
                                End If
                            End If

                            'format this row as page break
                            'pbpart pb 1
                            rngPB = .Selection.Range
                            .Selection.ParagraphFormat.PageBreakBefore = True

                            ' 'wdd.visible = True

                            If intTableID = 11 Then
                                'PageBreakaBefore for some reason  top-borders the selection
                                'must remove underline
                                '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                ''.Selection.Tables.Item(1).Cell(intFirstAnova, 1).Select()

                                rngT = tbl.Cell(intRow, 1).Range
                                rngT.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

                            End If
                        End If

                        'record this row
                        intNextPageRow = intRow
                        'tbl.Cell(intNextPageRow, 1).Select()

                        'go back one row
                        intRow = intRow - 1
                        tbl.Cell(intRow, 1).Select()
                        intLastSearchRow = intRow

                    End If

                    '*** Now we have a New proposed page, with some room reserved for legends and other stuff. ****
                    '20160511 LEE: This code sometimes runs at inappropriate times
                    'see Ricerca 034483
                    If (intLastSearchRow < intFirstSearchRow) Then
                        'Error
                        wd.Visible = True
                        strM = "A problem occurred while trying to add legends to this table:"
                        strM = strM & ChrW(10) & ChrW(10) & strT & ChrW(10) & ChrW(10)
                        strM = strM & ChrW(10) & "Default paragraph spacing, line spacing, or font settings too wide."
                        strM = strM & ChrW(10) & "Because of this, the Legend takes up the full page."
                        strM = strM & "There is no remaining space for the table itself."
                        strM = strM & "The legend will not be printed."
                        strM = strM & ChrW(10)
                        strM = strM & ChrW(10) & "Please reformat your Word Template and re-run the table."
                        MsgBox(strM, vbInformation, "Problem...")
                        wd.Visible = boolV
                        Exit Sub 'No page breaks
                    End If

                    ctRealLegend = 0

                    If DoLegend Then  'DoLegend seems to add the full legend on every page, regardless as to whether references
                        'appear in the page. 20160511 LEE: Yes, that is the purpose of DoLegend

                        For Count1 = 1 To ctLegend
                            ctRealLegend = ctRealLegend + 1
                            arrRealLegend(1, Count1) = arr(1, Count1)
                            arrRealLegend(2, Count1) = arr(2, Count1)
                            arrRealLegend(3, Count1) = arr(3, Count1)
                        Next
                        intForcedLegends = ctRealLegend

                        If boolIsLastPage Then
                        Else
                            ctRealLegend = ctRealLegend + 1 'take into accout 'continued on next page'
                            'ctRealLegend = ctRealLegend + 1 'sometimes wrap isn't occurring properly. Add another line
                        End If
                        If boolPageNum Then
                            ctRealLegend = ctRealLegend + 1
                        End If
                        If boolDTStamp Then
                            ctRealLegend = ctRealLegend + 1
                        End If
                    Else
                        'Look for reference strings in the table body on this new proposed page

                        ''Set selection to the proposed first page
                        'tbl.Cell(intFirstSearchRow, 1).Select()
                        '.Selection.SelectRow()
                        int1 = intLastSearchRow - intFirstSearchRow

                        'Select rows
                        '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        'myr = .Selection.Range

                        rngT = doc.Range(Start:=tbl.Rows(intFirstSearchRow).Range.Start, End:=tbl.Rows(intFirstSearchRow + int1).Range.End)
                        myr = rngT

                        'First Add column header legend items (which are marked in the 4th column of the array as "True")
                        ctRealLegend = 0
                        For Count1 = 1 To ctLegend
                            var1 = arr(4, Count1)
                            If NZ(var1, False) Then
                                ctRealLegend = ctRealLegend + 1
                                arrRealLegend(1, ctRealLegend) = arr(1, Count1)
                                arrRealLegend(2, ctRealLegend) = arr(2, Count1)
                                arrRealLegend(3, ctRealLegend) = arr(3, Count1)
                            End If
                        Next

                        'Now look for reference strings in the table body on page
                        For Count1 = 1 To ctLegend

                            var1 = arr(4, Count1)
                            If NZ(var1, False) Then
                            Else
                                var1 = arr(1, Count1)
                                'myr = .Selection.Range 'call this again because previous find action seems to wipe out the range
                                myr = doc.Range(Start:=tbl.Rows(intFirstSearchRow).Range.Start, End:=tbl.Rows(intFirstSearchRow + int1).Range.End)
                                '''wdd.visible = True

                                boolFound = False
                                With myr.Find
                                    .ClearFormatting()
                                    .MatchCase = True
                                    .MatchWholeWord = True
                                    .Forward = True
                                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                                    .Execute(FindText:=var1)

                                    If .Found Then
                                        boolFound = True
                                    Else
                                        boolFound = False
                                    End If

                                End With

                                If boolFound Then 'legend needed
                                    If StrComp(var1, "NA", CompareMethod.Text) = 0 Then 'debug
                                        var1 = var1
                                    End If
                                    ctRealLegend = ctRealLegend + 1
                                    arrRealLegend(1, ctRealLegend) = arr(1, Count1)
                                    arrRealLegend(2, ctRealLegend) = arr(2, Count1)
                                    arrRealLegend(3, ctRealLegend) = arr(3, Count1)
                                End If
                            End If

                        Next

                        intForcedLegends = ctRealLegend

                        If boolIsLastPage Then
                        Else
                            ctRealLegend = ctRealLegend + 1 'take into account 'continued on next page'
                            'ctRealLegend = ctRealLegend + 1 'sometimes the final product just doesn't wrap properly. Add another line to ctreallegend
                        End If
                        If boolPageNum Then
                            ctRealLegend = ctRealLegend + 1
                        End If
                        If boolDTStamp Then
                            ctRealLegend = ctRealLegend + 1
                        End If
                    End If



                    'add legend rows
                    tbl.Cell(intLastSearchRow, 1).Select()

                    '20180714 LEE:
                    'No! Must go back ctRealLengend rows, then insert
                    '20180716 NO!!!!
                    'intLastSearchRow = intLastSearchRow - ctRealLegend
                    'tbl.Cell(intLastSearchRow, 1).Select()


                    If ctRealLegend = 0 Then
                    Else
                        '1
                        'pbpart insert 1 may happen before a pagebreak if entire initial table is on one page
                        .Selection.InsertRowsBelow(ctRealLegend)

                        '20180526 LEE
                        Call RemoveUnicode(wd)

                        intLastLegend = intLastSearchRow + ctRealLegend

                        'remove any borders
                        Call removeAllBorders(wd, True)

                        myr1 = .Selection.Range

                        If intTableID = 11 Then
                            '20180716 LEE:
                            'it is highly probably there is a bottom border one line above this range
                            'remove it if it exists
                            Dim rngLB As Word.Range
                            rngLB = tbl.Rows(intLastSearchRow - 1).Range
                            rngLB.Select()
                            Call removeAllBorders(wd, True)
                            myr1.Select()

                        End If


                        'here 1
                        intRows = tbl.Rows.Count

                        'remove borders removes a desired border
                        'put it back

                        'tbl.Rows(intLastSearchRow).Select()
                        '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        rngT = tbl.Rows(intLastSearchRow).Range
                        rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        'tbl.Cell(intLastSearchRow, 1).Select()
                        myr1.Select()

                        If boolIsLastPage Then 'check to see if this went across page

                            ''do some bordering stuff
                            'tbl.Cell(intLastSearchRow, 1).Select()
                            'p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                            'border bottom
                            '.Selection.SelectRow()
                            '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                            rngT = tbl.Rows(intLastSearchRow).Range
                            rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                            'tbl.Cell(intLastSearchRow + ctRealLegend, 1).Select()
                            ''un-border bottom
                            '.Selection.SelectRow()
                            '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

                            rngT = tbl.Rows(intLastSearchRow + ctRealLegend).Range
                            rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

                            tbl.Cell(intLastSearchRow + ctRealLegend, 1).Select()
                            p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                            pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                            '''wdd.visible = True

                            If p1 <> pT Then 'must undo last amount of work and re-do row insertion

                                '****
                                If boolPageNum And intPN <> 0 Then
                                    If boolFullPageNum Then
                                        Try
                                            .ActiveDocument.Bookmarks.Item("PN" & intPN).Delete()
                                        Catch ex As Exception

                                        End Try

                                    End If
                                    intPN = intPN - 1
                                End If
                                '****

                                'jeez
                                Dim intT As Int64

                                'first count the number of rows needed
                                intT = tbl.Rows.Count
                                Dim intT1 As Int64

                                'tbl.Cell(intT, 1).Select()
                                'pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                                rngT = tbl.Cell(intT, 1).Range
                                pT = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                                intT1 = 0
                                Do Until pT = p1
                                    intT1 = intT1 + 1
                                    intT = intT - 1
                                    'tbl.Cell(intT, 1).Select()
                                    'pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                                    rngT = tbl.Cell(intT, 1).Range
                                    pT = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                                Loop

                                'delete new rows
                                'tbl.Cell(intLastSearchRow + 1, 1).Select()
                                'For Count1 = 1 To ctRealLegend
                                '    .Selection.Rows.Delete()
                                '    var1 = var1
                                'Next

                                rngT = doc.Range(Start:=tbl.Rows(intLastSearchRow + 1).Range.Start, End:=tbl.Rows(intLastSearchRow + ctRealLegend).Range.End)
                                rngT.Rows.Delete()

                                'the last delete action moved the cursor out of the table. Must move it back into the tbl
                                intRows = tbl.Rows.Count

                                'tbl.Cell(intLastSearchRow, 1).Select()
                                'p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                                'pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                                rngT = tbl.Cell(intLastSearchRow, 1).Range
                                p2 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                                pT = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                                'go back up intT rows
                                intRow = intRows - intT1 - 1
                                'record this row
                                intNextPageRow = intRow
                                tbl.Cell(intNextPageRow, 1).Select()
                                If boolCarefulSplit Then 'look for next blank row

                                    'tbl.Cell(intNextPageRow, 2).Select()
                                    'var1 = .Selection.Text
                                    rngT = tbl.Cell(intNextPageRow, 2).Range
                                    var1 = rngT.Text
                                    Do Until Len(var1) = 2
                                        intRow = intRow - 1
                                        'tbl.Cell(intRow, 2).Select()
                                        'var1 = .Selection.Text
                                        rngT = tbl.Cell(intRow, 2).Range
                                        var1 = rngT.Text
                                    Loop
                                    intRow = intRow + 1
                                    intNextPageRow = intRow
                                    tbl.Cell(intNextPageRow, 1).Select()

                                    If intTableID = 11 Then
                                        'continue with column 1
                                        var1 = .Selection.Text
                                        Do Until Len(var1) = 2
                                            intRow = intRow - 1
                                            'tbl.Cell(intRow, 1).Select()
                                            'var1 = .Selection.Text
                                            rngT = tbl.Cell(intRow, 1).Range
                                            var1 = rngT.Text
                                        Loop

                                        tbl.Cell(intRow, 1).Select()
                                        var1 = var1

                                    End If

                                End If

                                'format this row as page break
                                'pbpart pb 2
                                rngPB = .Selection.Range
                                .Selection.ParagraphFormat.PageBreakBefore = True


                                'go back one row
                                intRow = intRow - 1
                                tbl.Cell(intRow, 1).Select()
                                intLastSearchRow = intRow

                                'insert rows
                                ctRealLegend = ctRealLegend + 1 'must add 1 row for 'next page'
                                'NO!! Been done already
                                'If boolPageNum Then
                                '    ctRealLegend = ctRealLegend + 1
                                'End If
                                'If boolDTStamp Then
                                '    ctRealLegend = ctRealLegend + 1
                                'End If
                                'ctRealLegend = ctRealLegend + 1 'sometimes wrap isn't happening properly. add another line

                                '2
                                'pbpart insert 2
                                .Selection.InsertRowsBelow(ctRealLegend)

                                '20180526 LEE
                                Call RemoveUnicode(wd)

                                intLastLegend = intLastSearchRow + ctRealLegend

                                'remove any underlines
                                Call removeAllBorders(wd, True)

                                myr1 = .Selection.Range

                                'here 2
                                intRows = tbl.Rows.Count
                                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone

                                'need border here
                                'tbl.Rows(intLastSearchRow).Select()
                                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                                rngT = tbl.Rows(intLastSearchRow).Range
                                If rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle Then
                                Else
                                    rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                                End If
                                'rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                                myr1.Select()

                                boolIsLastPage = False

                            Else
                                myr1.Select()
                                'intLastSearchRow = intLastSearchRow + ctRealLegend
                            End If
                        End If  'End of stuff done only on last page


                        'merge the legend columns
                        .Selection.Cells.Split(NumRows:=ctRealLegend, NumColumns:=1, MergeBeforeSplit:=True)
                        '20181014 LEE:
                        'Aack! if rows > 1, the additional rows may get set to unicode
                        Call RemoveUnicode(wd)

                        '20180711 LEE:
                        'merge will reset cell padding, possibly resulting in table rows going to next page
                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        '20180715 LEE:
                        'depricate boolFirstAnova
                        'If boolFirstAnova And boolIsLastPage = False Then
                        '    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        'End If
                        If intTableID = 11 And boolIsLastPage = False Then
                            'NO!!
                            '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        End If

                        '''wdd.visible = True

                        If boolSmallFont Then
                            .Selection.Font.Size = .Selection.Font.Size - 1
                        End If
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                        'format cells
                        .Selection.ParagraphFormat.TabStops.Add(Position:=numTab1, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                        'With .Selection.ParagraphFormat
                        '    .LeftIndent = numIndent 'InchesToPoints(0.63)
                        '    .SpaceBefore = 0
                        '    .SpaceBeforeAuto = False
                        '    .SpaceAfter = 0
                        '    .SpaceAfterAuto = False
                        'End With

                        '20180701 LEE:
                        With .Selection.ParagraphFormat
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LeftIndent = numIndent 'InchesToPoints(0.63)
                            .FirstLineIndent = -numIndent 'InchesToPoints(-0.63)
                        End With

                    End If

                    'ensure top is  bordered

                    If intLastSearchRow = intRows Then
                        '.Selection.Tables.Item(1).Cell(intLastSearchRow, 1).Select()
                        rngT = tbl.Cell(intLastSearchRow, 1).Range

                        If intTableID = 11 Then
                        Else
                            '.Selection.SelectRow()
                            '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                            rngT = tbl.Rows(intLastSearchRow).Range
                            'rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                            If rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle Then
                            Else
                                rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                            End If
                        End If
                    Else
                        '.Selection.Tables.Item(1).Cell(intLastSearchRow + 1, 1).Select()
                        rngT = tbl.Cell(intLastSearchRow + 1, 1).Range
                        If intTableID = 11 Then
                        Else
                            '.Selection.SelectRow()
                            '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                            rngT = tbl.Rows(intLastSearchRow + 1).Range
                            If rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle Then
                            Else
                                rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                            End If
                        End If
                    End If

                    '''wdd.visible = True
                    ''' 
                    '20180629 LEE:
                    '20180629 LEE:
                    If intPageCount > intPBreakMax Then
                        wd.Visible = True
                        Try
                            doc.UndoClear()
                        Catch ex As Exception
                            var1 = var1
                        End Try
                        wd.Visible = boolV
                    End If

                    intRow = intLastSearchRow

                    'p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                    p1 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)


                    'wdd.visible = True

                    ctLCheck = ctRealLegend

                    'first do forced legends

                    For Count1 = 1 To intForcedLegends

                        intRow = intRow + 1

                        'Herehere
                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()

                        var1 = arrRealLegend(1, Count1)
                        '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
                        var1 = Replace(CStr(var1), "-", NBH, 1, -1, vbTextCompare)
                        If arrRealLegend(3, Count1) Then
                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
                            .Selection.Font.Bold = False

                            ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                            Call SetRowToOrig(wd)

                            If intTableID = 3 Then
                                If boolSTATSNR Then
                                    'Call typeInSuperscript(wd, CStr(var1))
                                    Call typeInSuperscriptFontSize12NoSpace(wd, CStr(var1))
                                Else
                                    .Selection.TypeText(Text:=CStr(var1))
                                End If
                            Else
                                'Call typeInSuperscript(wd, CStr(var1))
                                Call typeInSuperscriptFontSize12NoSpace(wd, CStr(var1))
                            End If

                        Else
                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
                            .Selection.Font.Bold = False
                            .Selection.TypeText(Text:=CStr(var1))

                            ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                            Call SetRowToOrig(wd)

                        End If
                        If Len(var1) = 0 Then
                        Else
                            .Selection.TypeText(Text:=vbTab)
                            .Selection.TypeText(Text:="=")
                            .Selection.TypeText(Text:=vbTab)

                            ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                            Call SetRowToOrig(wd)

                        End If
                        var2 = arrRealLegend(2, Count1)
                        var2 = Replace(CStr(var2), "-", NBH, 1, -1, vbTextCompare)
                        '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
                        .Selection.TypeText(Text:=CStr(var2))

                        ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                        Call SetRowToOrig(wd)

                    Next

                    'do extra legends
                    intNL = 0
                    If boolPageNum Then
                        intRow = intRow + 1
                        intNL = intNL + 1
                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()

                        intPN = intPN + 1
                        If boolFullPageNum Then
                            .Selection.TypeText(Text:="Page " & intPN & " of ")

                            ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                            Call SetRowToOrig(wd)

                            wrdSelection = wd.Selection()
                            With wd.ActiveDocument.Bookmarks
                                .Add(Range:=wrdSelection.Range, Name:="PN" & intPN)
                                .ShowHidden = False
                            End With
                            arrPN(intPN) = "PN" & intPN ' .Selection.Start
                            'format align right
                            .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                        Else
                            .Selection.TypeText(Text:="Page " & intPN)
                            'format align right
                            .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

                            ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                            Call SetRowToOrig(wd)

                        End If

                    End If

                    If boolDTStamp Then

                        intRow = intRow + 1
                        intNL = intNL + 1
                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()

                        .Selection.TypeText(Text:=strDateTimeStamp)
                        'format align right
                        .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

                        ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                        Call SetRowToOrig(wd)

                    End If

                    If boolIsLastPage = False Then

                        intRow = intRow + 1
                        intNL = intNL + 1
                        '.Selection.Tables.Item(1).Cell(intRow, 1).Select()
                        '.Selection.TypeText(Text:="Continued on next page")
                        ''format align right
                        '.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

                        rngT = tbl.Cell(intRow, 1).Range
                        rngT.Text = "Continued on next page"
                        'format align right
                        rngT.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

                        tbl.Cell(intRow, 1).Select()
                        ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                        Call SetRowToOrig(wd)


                    End If


                    '*** check to see if legend has crossed over to next page
                    'this can happen if legend rows wrap when they are long

                    intRows = tbl.Rows.Count
                    ''tbl.Cell(intRows, 1).Select()
                    'tbl.Cell(intLastLegend, 1).Select()
                    'p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                    'pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                    rngT = tbl.Cell(intLastLegend, 1).Range
                    p2 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                    pT = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                    If p1 < pT Then 'must undo last amount of work and re-do row insertion

                        '****
                        If boolPageNum And intPN <> 0 Then
                            If boolFullPageNum Then
                                Try
                                    .ActiveDocument.Bookmarks.Item("PN" & intPN).Delete()
                                Catch ex As Exception

                                End Try

                            End If
                            intPN = intPN - 1
                        End If
                        '****

                        'goto pagebreak and undo
                        'rngpb may not exist
                        'pbpart pb 3
                        tbl.Cell(intLastLegend + 1, 1).Select()
                        If rngPB Is Nothing Then
                        Else
                            rngPB.Select()
                        End If

                        .Selection.ParagraphFormat.PageBreakBefore = False

                        'delete new rows
                        'tbl.Cell(intLastSearchRow + 1, 1).Select()
                        'For Count1 = 1 To ctRealLegend
                        '    .Selection.Rows.Delete()
                        '    var1 = var1
                        'Next

                        '20180713 LEE Check here
                        rngT = doc.Range(Start:=tbl.Rows(intLastSearchRow + 1).Range.Start, End:=tbl.Rows(intLastSearchRow + ctRealLegend).Range.End)
                        rngT.Rows.Delete()

                        'the last delete action moved the cursor out of the table. Must move it back into the tbl
                        intRows = tbl.Rows.Count
                        tbl.Cell(intLastSearchRow, 1).Select()

                        'remove any borders
                        'tbl.Rows(intLastSearchRow + 1).Select()
                        .Selection.SelectRow()
                        Call removeAllBorders(wd, True)

                        If intTableID = 11 Then
                        Else
                            'the last delete action moved the cursor out of the table. Must move it back into the tbl
                            intRows = tbl.Rows.Count
                            tbl.Cell(intRows, 1).Select()
                        End If
                        p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                        '*****
                        'now do search to determine new intsplitrows
                        If DoLegend Then
                            tRow = ctLegend
                        Else

                            tbl.Cell(intFirstSearchRow, 1).Select()
                            .Selection.SelectRow()
                            int1 = intLastSearchRow - intFirstSearchRow
                            ''.Selection.MoveDown(Unit:=wdLine, Count:=int1, Extend:=WdMovementType.wdExtend)
                            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            'myr = .Selection.Range

                            rngT = doc.Range(Start:=tbl.Rows(intFirstSearchRow).Range.Start, End:=tbl.Rows(intFirstSearchRow + int1).Range.End)
                            myr = rngT

                            'low look for body items
                            tRow = 0
                            For Count1 = 1 To ctLegend

                                var1 = arr(4, Count1)
                                If NZ(var1, False) Then
                                    tRow = tRow + 1
                                Else
                                    var1 = arr(1, Count1)
                                    'myr = .Selection.Range 'call this again because previous find action seems to wipe out the range
                                    myr = doc.Range(Start:=tbl.Rows(intFirstSearchRow).Range.Start, End:=tbl.Rows(intFirstSearchRow + int1).Range.End)
                                    '''wdd.visible = True

                                    boolFound = False
                                    With myr.Find
                                        .ClearFormatting()
                                        .MatchCase = True
                                        .MatchWholeWord = True
                                        .Forward = True
                                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                                        .Execute(FindText:=var1)

                                        If .Found Then
                                            boolFound = True
                                        Else
                                            boolFound = False
                                        End If

                                    End With

                                    If boolFound Then 'legend needed
                                        If StrComp(var1, "NA", CompareMethod.Text) = 0 Then 'debug
                                            var1 = var1
                                        End If
                                        tRow = tRow + 1
                                    End If
                                End If

                            Next

                        End If

                        'add a row for safety
                        tRow = tRow + 1

                        If boolIsLastPage Then
                        Else
                            tRow = tRow + 1 'take into account 'continued on next page'
                        End If
                        If boolPageNum Then
                            tRow = tRow + 1
                        End If
                        If boolDTStamp Then
                            tRow = tRow + 1
                        End If
                        If intTableID = 3 Then
                            If boolSTATSREGR Then
                                tRow = tRow + 1
                            End If
                        End If
                        '*****

                        'go back up tRows (to account for legend plus other items at bottom of page)
                        intRow = intLastSearchRow - tRow

                        'rngT here
                        If boolCarefulSplit Then
                            tbl.Cell(intRow, 2).Select()
                            var1 = .Selection.Text
                            Do Until Len(var1) = 2
                                intRow = intRow - 1
                                'tbl.Cell(intRow, 2).Select()
                                'var1 = .Selection.Text

                                rngT = tbl.Cell(intRow, 2).Range
                                var1 = rngT.Text

                            Loop
                            intRow = intRow + 1
                            tbl.Cell(intRow, 1).Select()

                            If intTableID = 11 Then
                                'continue with column 1
                                var1 = .Selection.Text
                                Do Until Len(var1) = 2
                                    intRow = intRow - 1
                                    'tbl.Cell(intRow, 1).Select()
                                    'var1 = .Selection.Text

                                    rngT = tbl.Cell(intRow, 1).Range
                                    var1 = rngT.Text
                                Loop
                                tbl.Cell(intRow, 1).Select()
                                var1 = var1
                            End If
                        End If

                        'record this row
                        intNextPageRow = intRow
                        tbl.Cell(intNextPageRow, 1).Select()


                        'format this row as page break
                        'pbpart pb 4
                        rngPB = .Selection.Range
                        .Selection.ParagraphFormat.PageBreakBefore = True

                        'go back one row
                        intRow = intRow - 1
                        tbl.Cell(intRow, 1).Select()
                        intLastSearchRow = intRow

                        '20180714 LEE:
                        myr1 = .Selection.Range

                        're-perform search

                        '****

                        'search for legend stuff again
                        boolIsLastPage = False

                        ctRealLegend = 0

                        'badd!!!
                        If DoLegend Then

                            For Count1 = 1 To ctLegend
                                ctRealLegend = ctRealLegend + 1
                                arrRealLegend(1, Count1) = arr(1, Count1)
                                arrRealLegend(2, Count1) = arr(2, Count1)
                                arrRealLegend(3, Count1) = arr(3, Count1)
                            Next
                            intForcedLegends = ctRealLegend

                            If boolIsLastPage Then
                            Else
                                ctRealLegend = ctRealLegend + 1 'take into accout 'continued on next page'
                                'ctRealLegend = ctRealLegend + 1 'sometimes wrap doesn't happen properly. Add another line'
                            End If
                            If boolPageNum Then
                                ctRealLegend = ctRealLegend + 1
                            End If
                            If boolDTStamp Then
                                ctRealLegend = ctRealLegend + 1
                            End If
                        Else

                            'look for legends
                            tbl.Cell(intFirstSearchRow, 1).Select()
                            .Selection.SelectRow()
                            int1 = intLastSearchRow - intFirstSearchRow
                            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            'myr = .Selection.Range

                            rngT = doc.Range(Start:=tbl.Rows(intFirstSearchRow).Range.Start, End:=tbl.Rows(intFirstSearchRow + int1).Range.End)
                            myr = rngT

                            'look for legends
                            'first look for column header legend items
                            ctRealLegend = 0
                            For Count1 = 1 To ctLegend
                                var1 = arr(4, Count1)
                                If NZ(var1, False) Then
                                    ctRealLegend = ctRealLegend + 1
                                    arrRealLegend(1, ctRealLegend) = arr(1, Count1)
                                    arrRealLegend(2, ctRealLegend) = arr(2, Count1)
                                    arrRealLegend(3, ctRealLegend) = arr(3, Count1)
                                End If
                            Next

                            'low look for body items
                            For Count1 = 1 To ctLegend

                                var1 = arr(4, Count1)
                                If NZ(var1, False) Then
                                Else
                                    var1 = arr(1, Count1)
                                    'myr = .Selection.Range 'call this again because previous find action seems to wipe out the range
                                    myr = doc.Range(Start:=tbl.Rows(intFirstSearchRow).Range.Start, End:=tbl.Rows(intFirstSearchRow + int1).Range.End)
                                    boolFound = False
                                    With myr.Find
                                        .ClearFormatting()
                                        .MatchCase = True
                                        .MatchWholeWord = True
                                        .Forward = True
                                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                                        .Execute(FindText:=var1)

                                        If .Found Then
                                            boolFound = True
                                        Else
                                            boolFound = False
                                        End If

                                    End With

                                    If boolFound Then 'legend needed
                                        If StrComp(var1, "NA", CompareMethod.Text) = 0 Then 'debug
                                            var1 = var1
                                        End If
                                        ctRealLegend = ctRealLegend + 1
                                        arrRealLegend(1, ctRealLegend) = arr(1, Count1)
                                        arrRealLegend(2, ctRealLegend) = arr(2, Count1)
                                        arrRealLegend(3, ctRealLegend) = arr(3, Count1)
                                    End If
                                End If

                            Next

                            intForcedLegends = ctRealLegend

                            'I'm here
                            If boolIsLastPage Then
                            Else
                                ctRealLegend = ctRealLegend + 1 'take into accout 'continued on next page'
                                'ctRealLegend = ctRealLegend + 1 'sometimes wrap still isn't going properly. Add another line
                            End If
                            If boolPageNum Then
                                ctRealLegend = ctRealLegend + 1
                            End If
                            If boolDTStamp Then
                                ctRealLegend = ctRealLegend + 1
                            End If
                        End If



                        '****

                        'insert rows
                        '20180714 LEE: Moved to before If DoLegend
                        'myr1 = .Selection.Range
                        myr1.Select()
                        '3
                        'pbpart insert 4
                        .Selection.InsertRowsBelow(ctRealLegend)

                        '20180526 LEE
                        Call RemoveUnicode(wd)

                        'remove any underlines
                        Call removeAllBorders(wd, True)

                        myr1 = .Selection.Range

                        'here 3
                        intRows = tbl.Rows.Count
                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone

                        'tbl.Rows(intLastSearchRow).Select()
                        '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        rngT = tbl.Rows(intLastSearchRow).Range
                        rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                        myr1.Select()

                        'merge the rows
                        .Selection.Cells.Split(NumRows:=ctRealLegend, NumColumns:=1, MergeBeforeSplit:=True)
                        '20181014 LEE:
                        'Aack! if rows > 1, the additional rows may get set to unicode
                        Call RemoveUnicode(wd)
                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone
                        If boolSmallFont Then
                            .Selection.Font.Size = .Selection.Font.Size - 1
                        End If
                        '            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                        'format cells

                        .Selection.ParagraphFormat.TabStops.Add(Position:=numTab1, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                        'With .Selection.ParagraphFormat
                        '    .LeftIndent = numIndent 'InchesToPoints(0.63)
                        '    .SpaceBefore = 0
                        '    .SpaceBeforeAuto = False
                        '    .SpaceAfter = 0
                        '    .SpaceAfterAuto = False
                        'End With
                        '20180701 LEE:
                        With .Selection.ParagraphFormat
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LeftIndent = numIndent 'InchesToPoints(0.63)
                            .FirstLineIndent = -numIndent 'InchesToPoints(-0.63)
                        End With

                        'do some bordering stuff
                        'tbl.Cell(intLastSearchRow, 1).Select()
                        'p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        ''        p1 = .selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        'border bottom
                        ''Selection.Borders(wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        '.Selection.SelectRow()
                        '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        rngT = tbl.Rows(intLastSearchRow).Range
                        p1 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
                        rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                        'tbl.Cell(intLastSearchRow + ctRealLegend, 1).Select()
                        '.Selection.SelectRow()
                        ''un-border bottom
                        '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

                        rngT = tbl.Rows(intLastSearchRow + ctRealLegend).Range
                        'un-border bottom
                        rngT.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

                        intRow = intLastSearchRow

                        ctLCheck = ctRealLegend

                        '****
                        For Count1 = 1 To intForcedLegends

                            intRow = intRow + 1

                            'Herehere
                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()

                            var1 = arrRealLegend(1, Count1)
                            '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
                            var1 = Replace(CStr(var1), "-", NBH, 1, -1, vbTextCompare)
                            If arrRealLegend(3, Count1) Then
                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
                                .Selection.Font.Bold = False
                                ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                                Call SetRowToOrig(wd)

                                If intTableID = 3 Then
                                    If boolSTATSNR Then
                                        'Call typeInSuperscript(wd, CStr(var1))
                                        Call typeInSuperscriptFontSize12NoSpace(wd, CStr(var1))
                                    Else
                                        .Selection.TypeText(Text:=CStr(var1))
                                    End If
                                Else
                                    'Call typeInSuperscript(wd, CStr(var1))
                                    Call typeInSuperscriptFontSize12NoSpace(wd, CStr(var1))
                                End If

                            Else
                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
                                .Selection.Font.Bold = False
                                .Selection.TypeText(Text:=CStr(var1))

                                ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                                Call SetRowToOrig(wd)


                            End If
                            If Len(var1) = 0 Then
                            Else
                                .Selection.TypeText(Text:=vbTab)
                                .Selection.TypeText(Text:="=")
                                .Selection.TypeText(Text:=vbTab)

                                ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                                Call SetRowToOrig(wd)


                            End If
                            var2 = arrRealLegend(2, Count1)
                            var2 = Replace(CStr(var2), "-", NBH, 1, -1, vbTextCompare)
                            '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=CStr(var2))

                            ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                            Call SetRowToOrig(wd)


                        Next

                        'do extra legends
                        intNL = 0
                        If boolPageNum Then
                            intRow = intRow + 1
                            intNL = intNL + 1
                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()

                            intPN = intPN + 1
                            If boolFullPageNum Then
                                .Selection.TypeText(Text:="Page " & intPN & " of ")

                                ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                                Call SetRowToOrig(wd)


                                wrdSelection = wd.Selection()
                                With wd.ActiveDocument.Bookmarks
                                    .Add(Range:=wrdSelection.Range, Name:="PN" & intPN)
                                    .ShowHidden = False
                                End With
                                arrPN(intPN) = "PN" & intPN ' .Selection.Start
                                'format align right
                                .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            Else
                                .Selection.TypeText(Text:="Page " & intPN)
                                'format align right
                                .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

                                ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                                Call SetRowToOrig(wd)

                            End If

                        End If

                        If boolDTStamp Then

                            intRow = intRow + 1
                            intNL = intNL + 1
                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()

                            .Selection.TypeText(Text:=strDateTimeStamp)
                            'format align right
                            .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

                            ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                            Call SetRowToOrig(wd)

                        End If

                        If boolIsLastPage = False Then

                            intRow = intRow + 1
                            intNL = intNL + 1
                            '.Selection.Tables.Item(1).Cell(intRow, 1).Select()

                            '.Selection.TypeText(Text:="Continued on next page")
                            ''format align right
                            '.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

                            rngT = tbl.Cell(intRow, 1).Range
                            rngT.Text = "Continued on next page"
                            'format align right
                            rngT.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                            ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                            Call SetRowToOrig(wd)


                        End If
                        '****

                        var1 = "a" 'debug
                        intLastLegend = intRow ' intLastSearchRow + ctRealLegend
                    Else
                        intNextPageRow = intLastSearchRow
                    End If

                    '20180629 LEE:
                    If intPageCount > intPBreakMax Then
                        wd.Visible = True
                        Try
                            doc.UndoClear()
                        Catch ex As Exception
                            var1 = var1
                        End Try
                        wd.Visible = boolV
                    End If


                    'End If

                    '***End check for legend page overflow

                    'intNextPageRow = intNextPageRow + ctRealLegend
                    intNextPageRow = intLastLegend + 1
                    intRows = tbl.Rows.Count
                    If intNextPageRow > intRows Then
                        intNextPageRow = intRows
                    End If

                    If boolIsLastPage Then

                    Else
                        If intPageCount > UBound(arrPageBreaks, 2) Then
                            ReDim Preserve arrPageBreaks(2, intPageCount)
                        End If
                        arrPageBreaks(2, intPageCount) = intNextPageRow - 1
                        intPageCount = intPageCount + 1
                        If intPageCount > UBound(arrPageBreaks, 2) Then
                            ReDim Preserve arrPageBreaks(2, intPageCount)
                        End If
                        arrPageBreaks(1, intPageCount) = intNextPageRow
                    End If

                    'top border last data row + 1

                    intFirstSearchRow = intNextPageRow
                    intRow = intFirstSearchRow
                    intRows = tbl.Rows.Count

                    If intPageCount > UBound(arrPageBreaks, 2) Then
                        ReDim Preserve arrPageBreaks(2, intPageCount)
                    End If
                    arrPageBreaks(2, intPageCount) = intRows

                    '20180628 LEE:
                    intPB = intPB + 1
                    If intPB >= intPBMax Then

                        strM3 = frmH.lblProgress.Text
                        strM2 = strM3 & ChrW(10) & "Saving document..."
                        frmH.lblProgress.Text = strM2
                        frmH.lblProgress.Refresh()

                        'save
                        doc.Save()
                        'wait

                        Pause(0.2)
                        intPB = 0

                        frmH.lblProgress.Text = strM3
                        frmH.lblProgress.Refresh()

                    End If
                    '20180628 LEE:
                    If intPageCount > intPBreakMax Then
                        wd.Visible = True
                        Try
                            doc.UndoClear()
                        Catch ex As Exception
                            var1 = var1
                        End Try
                        wd.Visible = boolV
                    End If

                    'tbl.Cell(intFirstSearchRow, 1).Select()
                    'p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                    tbl.Cell(intFirstSearchRow, 1).Select()
                    rngT = tbl.Cell(intFirstSearchRow, 1).Range
                    p1 = rngT.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

                    If boolIsLastPage Then
                        Exit Do
                    End If

                    'Exit if we are at the end (but I still want this loop executed the first time, for 1-line tables
                    If intRow = intRows Then
                        Exit Do
                    End If
                Loop

                If boolFullPageNum Then
                    Dim pos As Int64
                    Dim rngP As Word.Range
                    For Count1 = 1 To intPN

                        Try

                            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="PN" & Count1)
                            .Selection.TypeText(Text:=CStr(intPN))

                            ''20160218 LEE: set row heights to normal in case row heights have been changed for a specific table
                            Call SetRowToOrig(wd)

                        Catch ex As Exception

                        End Try

                    Next

                    'delete bookmarks
                    For Count1 = 1 To intPN

                        Try
                            .ActiveDocument.Bookmarks.Item("PN" & Count1).Delete()
                        Catch ex As Exception

                        End Try
                    Next
                End If

                If gboolReadOnlyTables Then
                    Call ReadOnlyTables(wd, tbl, ctHdRows, arrPageBreaks, intPageCount)
                End If

            End With
        Catch ex As Exception
            var1 = ex.Message
        End Try

        '20180529 LEE:
        'RemoveUnicode is messing with Caption header
        'if tab (vs soft return), ensure that caption has hanging indent
        Call ApplyTableCaption(wd)

        wd.Visible = boolV

        '20180628 LEE:
        'New logic: save after every table
        strM3 = frmH.lblProgress.Text
        strM2 = strM3 & ChrW(10) & "Saving document..."
        frmH.lblProgress.Text = strM2
        frmH.lblProgress.Refresh()
        'save
        doc.Save()
        'wait
        Pause(0.2)

        frmH.lblProgress.Text = strM1
        frmH.lblProgress.Refresh()

    End Sub

    Sub SetRowToOrig(wd As Microsoft.Office.Interop.Word.Application)

        With wd
            .Selection.Rows.HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto
            .Selection.Rows.Height = 0 ' InchesToPoints(0)
        End With


    End Sub

    'NOTE: Ctrl-K Ctrl-C uncomments sections in Visual Studio

    Sub _19_Nov_2015(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean, ByVal boolCarefulSplit As Boolean, ByVal boolFirstAnova As Boolean, ByVal intTableID As Int64)

        '    'Sub SplitTable(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, ByVal arr As Object, ByVal strT As String,
        '    'ByVal DoLegend As Boolean, ByVal DoTable As Boolean, ByVal intSRow As Short, ByVal intSplitRows As Short, ByVal boolAutoFit As Boolean, ByVal boolSmallFont As Boolean)


        '    Dim intRows As Long
        '    Dim intRow As Long
        '    Dim Count1 As Long
        '    Dim Count2 As Long
        '    Dim Count3 As Long
        '    Dim str1 As String
        '    Dim str2 As String
        '    Dim str3 As String
        '    Dim var1, var2, var3, var4
        '    Dim tbl As Word.Table
        '    Dim p1 As Long
        '    Dim p2 As Long
        '    Dim pT As Long
        '    Dim intIncr As Integer
        '    Dim rowStart As Long 'intRow1
        '    Dim ctP As Long
        '    Dim intNextPageRow As Long
        '    Dim intLastSearchRow As Long
        '    Dim intFirstSearchRow As Long
        '    Dim int1 As Long
        '    Dim int2 As Long
        '    Dim myr As Word.Range
        '    Dim myr1 As Word.Range
        '    Dim ctRealLegend As Integer
        '    Dim arrRealLegend
        '    Dim Fonts
        '    Dim boolIsLastPage As Boolean
        '    Dim boolFound As Boolean

        '    Dim intLs As Int32

        '    ReDim arrRealLegend(3, ctLegend)

        '    Dim arrPageBreaks(2, 500) As Short
        '    Dim intPageCount As Int16
        '    Dim tRow As Int16
        '    Dim intLastLegend As Int16 = 0

        '    Dim chkR As Int16
        '    Dim chkM As Int16

        '    Dim strDateTimeStamp As String
        '    Dim arrPN(1000)
        '    Dim intPN As Short = 0
        '    Dim boolPageNum As Boolean = False
        '    Dim boolDTStamp As Boolean = LBOOLTABLEDTTIMESTAMP
        '    Dim boolFullPageNum As Boolean = False
        '    Dim ctLCheck As Short
        '    Dim wrdSelection As Word.Selection

        '    Dim intForcedLegends As Short = 0
        '    Dim intNL As Short = 0

        '    If StrComp(LcharSTPage, "[None]", CompareMethod.Text) = 0 Then
        '        boolPageNum = False
        '    Else
        '        boolPageNum = True
        '    End If

        '    If StrComp(LcharSTPage, "Page x", CompareMethod.Text) = 0 Then
        '        boolFullPageNum = False
        '    Else
        '        boolFullPageNum = True
        '    End If

        '    'strDateTimeStamp = Format(LTableDateTimeStamp, LDateFormat & " HH:mm:ss tt")
        '    strDateTimeStamp = Format(gdtReportDate, LDateFormat & " HH:mm:ss tt")


        '    intPageCount = 1
        '    arrPageBreaks(1, intPageCount) = ctHdRows + 1

        '    'wdd.visible = True

        '    With wd

        '        '''wdd.visible = True

        '        tbl = .Selection.Tables.Item(1)

        '        If boolPageNum Then
        '            intSplitRows = intSplitRows + 1
        '        End If
        '        If boolDTStamp Then
        '            intSplitRows = intSplitRows + 1
        '        End If
        '        intSplitRows = intSplitRows + 1 'account for Continued on next page
        '        'if orientation is portrait, add a few rows
        '        If .Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait Then
        '            intSplitRows = intSplitRows + 2 'for safety in case legends rap
        '        End If
        '        'intSplitRows = intSplitRows + 2 'for safety in case legends rap

        '        'first get rows
        '        intRows = tbl.Rows.Count
        '        tbl.Cell(1, 1).Select()
        '        p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '        tbl.Cell(intRows, 1).Select()
        '        pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '        intLs = pT - p1 + 1

        '        tbl.Cell(1, 1).Select()
        '        If p1 = pT Then 'no need to evaluate
        '            intRows = tbl.Rows.Count
        '            boolIsLastPage = True
        '        Else
        '            boolIsLastPage = False
        '        End If

        '        'begin looking for next page
        '        intRow = ctHdRows + 1
        '        ctP = 0

        '        If .Selection.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
        '            intIncr = 20
        '        Else
        '            intIncr = 40
        '        End If

        '        intFirstSearchRow = ctHdRows + 1
        '        boolIsLastPage = False

        '        chkM = 200
        '        chkR = 0

        '        Do Until intRow >= intRows

        '            chkR = chkR + 1
        '            If chkR > chkM Then
        '                Exit Do
        '            End If

        '            ctP = ctP + 1
        '            If ctP > intLs Then
        '                str1 = strT & vbCr & "Formatting legend..." & ctP & " of " & ctP
        '            Else
        '                str1 = strT & vbCr & "Formatting legend..." & ctP & " of " & intLs
        '            End If

        '            frmH.lblProgress.Text = str1
        '            frmH.lblProgress.Refresh()


        '            tbl.Cell(intRow, 1).Select()
        '            p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '            intRows = tbl.Rows.Count

        '            tbl.Cell(intRows, 1).Select()
        '            pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '            If p1 = pT Then 'don't look for table split
        '                intRow = intRows
        '                intLastSearchRow = intRows
        '                boolIsLastPage = True
        '            Else 'look for table split

        '                intRow = intRow + intIncr
        '                If intRow > intRows Then
        '                    intRow = intRows
        '                End If

        '                tbl.Cell(intRow, 1).Select()

        '                'check to see if intincr is too big
        '                p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                If p2 > p1 Then 'intincr is too big
        '                    Do Until p2 = p1
        '                        intRow = intRow - intIncr
        '                        intIncr = intIncr - 5
        '                        intRow = intRow + intIncr
        '                        tbl.Cell(intRow, 1).Select()
        '                        p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    Loop
        '                End If

        '                p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                p2 = p1 + 1
        '                Do Until p1 = p2
        '                    If intRow >= intRows Then
        '                        Exit Do
        '                    End If
        '                    intRow = intRow + 1
        '                    tbl.Cell(intRow, 1).Select()
        '                    p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    If intRow = intRows Then
        '                        Exit Do
        '                    End If
        '                Loop

        '                '*****
        '                'wdd.visible = True
        '                'now do search to determine new intsplitrows
        '                If DoLegend Then
        '                    tRow = ctLegend
        '                Else

        '                    tbl.Cell(intFirstSearchRow, 1).Select()
        '                    .Selection.SelectRow()
        '                    intLastSearchRow = intRow - 1
        '                    int1 = intLastSearchRow - intFirstSearchRow
        '                    '.Selection.MoveDown(Unit:=wdLine, Count:=int1, Extend:=WdMovementType.wdExtend)
        '                    .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                    myr = .Selection.Range

        '                    'low look for body items
        '                    tRow = 0
        '                    For Count1 = 1 To ctLegend

        '                        var1 = arr(4, Count1)
        '                        If NZ(var1, False) Then
        '                            tRow = tRow + 1
        '                        Else
        '                            var1 = arr(1, Count1)
        '                            myr = .Selection.Range 'call this again because previous find action seems to wipe out the range

        '                            '''wdd.visible = True

        '                            boolFound = False
        '                            With myr.Find
        '                                .ClearFormatting()
        '                                .MatchCase = True
        '                                .MatchWholeWord = True
        '                                .Forward = True
        '                                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                                .Execute(FindText:=var1)

        '                                If .Found Then
        '                                    boolFound = True
        '                                Else
        '                                    boolFound = False
        '                                End If

        '                            End With

        '                            If boolFound Then 'legend needed
        '                                tRow = tRow + 1
        '                            End If
        '                        End If

        '                    Next

        '                End If

        '                'add a row for safety
        '                tRow = tRow + 1

        '                If boolIsLastPage Then
        '                Else
        '                    tRow = tRow + 1 'take into account 'continued on next page'
        '                End If
        '                If boolPageNum Then
        '                    tRow = tRow + 1
        '                End If
        '                If boolDTStamp Then
        '                    tRow = tRow + 1
        '                End If

        '                If intTableID = 3 Then
        '                    If boolSTATSREGR Then
        '                        tRow = tRow + 1
        '                    End If
        '                End If
        '                intSplitRows = tRow

        '                '*****


        '                'now retreat ctLegend rows
        '                'intRow = intRow - intSplitRows
        '                If boolFirstAnova Then

        '                Else
        '                    intRow = intRow - intSplitRows
        '                    tbl.Cell(intRow, 1).Select()
        '                    If boolCarefulSplit Then
        '                        tbl.Cell(intRow, 2).Select()
        '                        var1 = .Selection.Text
        '                        Do Until Len(var1) = 2
        '                            '''wdd.visible = True

        '                            var2 = Len(var1)
        '                            intRow = intRow - 1
        '                            tbl.Cell(intRow, 2).Select()
        '                            var1 = .Selection.Text
        '                        Loop
        '                        intRow = intRow + 1
        '                        tbl.Cell(intRow, 1).Select()

        '                        If intTableID = 11 Then
        '                            'continue with column 1
        '                            var1 = .Selection.Text
        '                            Do Until Len(var1) = 2
        '                                intRow = intRow - 1
        '                                tbl.Cell(intRow, 1).Select()
        '                                var1 = .Selection.Text
        '                            Loop
        '                        End If
        '                    End If

        '                    'format this row as page break
        '                    .Selection.ParagraphFormat.PageBreakBefore = True

        '                    ' 'wdd.visible = True

        '                    If intTableID = 11 Then
        '                        'PageBreakaBefore for some reason  top-borders the selection
        '                        'must remove underline
        '                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        '.Selection.Tables.Item(1).Cell(intFirstAnova, 1).Select()

        '                    End If
        '                End If

        '                'record this row
        '                intNextPageRow = intRow
        '                tbl.Cell(intNextPageRow, 1).Select()

        '                'go back one row
        '                intRow = intRow - 1
        '                tbl.Cell(intRow, 1).Select()
        '                intLastSearchRow = intRow

        '            End If

        '            'select range and look for legends

        '            ctRealLegend = 0

        '            If DoLegend Then

        '                For Count1 = 1 To ctLegend
        '                    ctRealLegend = ctRealLegend + 1
        '                    arrRealLegend(1, Count1) = arr(1, Count1)
        '                    arrRealLegend(2, Count1) = arr(2, Count1)
        '                    arrRealLegend(3, Count1) = arr(3, Count1)
        '                Next
        '                intForcedLegends = ctRealLegend

        '                If boolIsLastPage Then
        '                Else
        '                    ctRealLegend = ctRealLegend + 1 'take into accout 'continued on next page'
        '                    'ctRealLegend = ctRealLegend + 1 'sometimes wrap isn't occurring properly. Add another line
        '                End If
        '                If boolPageNum Then
        '                    ctRealLegend = ctRealLegend + 1
        '                End If
        '                If boolDTStamp Then
        '                    ctRealLegend = ctRealLegend + 1
        '                End If
        '            Else

        '                tbl.Cell(intFirstSearchRow, 1).Select()
        '                .Selection.SelectRow()
        '                int1 = intLastSearchRow - intFirstSearchRow
        '                '.Selection.MoveDown(Unit:=wdLine, Count:=int1, Extend:=WdMovementType.wdExtend)
        '                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                myr = .Selection.Range

        '                'wdd.visible = True


        '                'look for legends
        '                'first look for column header legend items
        '                ctRealLegend = 0
        '                For Count1 = 1 To ctLegend
        '                    var1 = arr(4, Count1)
        '                    If NZ(var1, False) Then
        '                        ctRealLegend = ctRealLegend + 1
        '                        arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                        arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                        arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                    End If
        '                Next

        '                'low look for body items
        '                For Count1 = 1 To ctLegend

        '                    var1 = arr(4, Count1)
        '                    If NZ(var1, False) Then
        '                    Else
        '                        var1 = arr(1, Count1)
        '                        myr = .Selection.Range 'call this again because previous find action seems to wipe out the range

        '                        '''wdd.visible = True

        '                        boolFound = False
        '                        With myr.Find
        '                            .ClearFormatting()
        '                            .MatchCase = True
        '                            .MatchWholeWord = True
        '                            .Forward = True
        '                            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                            .Execute(FindText:=var1)

        '                            If .Found Then
        '                                boolFound = True
        '                            Else
        '                                boolFound = False
        '                            End If

        '                        End With

        '                        If boolFound Then 'legend needed
        '                            ctRealLegend = ctRealLegend + 1
        '                            arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                            arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                            arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                        End If
        '                    End If

        '                Next

        '                intForcedLegends = ctRealLegend

        '                If boolIsLastPage Then
        '                Else
        '                    ctRealLegend = ctRealLegend + 1 'take into account 'continued on next page'
        '                    'ctRealLegend = ctRealLegend + 1 'sometimes the final product just doesn't wrap properly. Add another line to ctreallegend
        '                End If
        '                If boolPageNum Then
        '                    ctRealLegend = ctRealLegend + 1
        '                End If
        '                If boolDTStamp Then
        '                    ctRealLegend = ctRealLegend + 1
        '                End If
        '            End If



        '            'add legend rows
        '            tbl.Cell(intLastSearchRow, 1).Select()
        '            If ctRealLegend = 0 Then
        '            Else

        '                .Selection.InsertRowsBelow(ctRealLegend)
        '                intLastLegend = intLastSearchRow + ctRealLegend

        '                'remove any borders
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '                myr1 = .Selection.Range
        '                intRows = tbl.Rows.Count

        '                'remove borders removes a desired border
        '                'put it back
        '                'wd.Visible = True
        '                tbl.Rows(intLastSearchRow).Select()
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                tbl.Cell(intLastSearchRow, 1).Select()
        '                myr1.Select()

        '                If boolIsLastPage Then 'check to see if this went across page

        '                    'do some bordering stuff
        '                    tbl.Cell(intLastSearchRow, 1).Select()
        '                    p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    'border bottom
        '                    'Selection.Borders(wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                    .Selection.SelectRow()
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

        '                    tbl.Cell(intLastSearchRow + ctRealLegend, 1).Select()
        '                    .Selection.SelectRow()
        '                    'un-border bottom
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '                    p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '                    '''wdd.visible = True

        '                    If p1 <> pT Then 'must undo

        '                        '****
        '                        If boolPageNum And intPN <> 0 Then
        '                            If boolFullPageNum Then
        '                                Try
        '                                    .ActiveDocument.Bookmarks.Item("PN" & intPN).Delete()
        '                                Catch ex As Exception

        '                                End Try

        '                            End If
        '                            intPN = intPN - 1
        '                        End If
        '                        '****

        '                        'jeez
        '                        Dim intT As Int64

        '                        'first count the number of rows needed
        '                        intT = tbl.Rows.Count
        '                        Dim intT1 As Int64

        '                        tbl.Cell(intT, 1).Select()
        '                        pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                        intT1 = 0
        '                        Do Until pT = p1
        '                            intT1 = intT1 + 1
        '                            intT = intT - 1
        '                            tbl.Cell(intT, 1).Select()
        '                            pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                        Loop

        '                        'delete new rows
        '                        tbl.Cell(intLastSearchRow + 1, 1).Select()
        '                        For Count1 = 1 To ctRealLegend
        '                            .Selection.Rows.Delete()
        '                            var1 = var1
        '                        Next

        '                        'the last delete action moved the cursor out of the table. Must move it back into the tbl
        '                        intRows = tbl.Rows.Count

        '                        tbl.Cell(intLastSearchRow, 1).Select()
        '                        p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                        pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '                        'go back up intT rows
        '                        intRow = intRows - intT1 - 1
        '                        'record this row
        '                        intNextPageRow = intRow
        '                        tbl.Cell(intNextPageRow, 1).Select()
        '                        If boolCarefulSplit Then 'look for next blank row
        '                            tbl.Cell(intNextPageRow, 2).Select()
        '                            var1 = .Selection.Text
        '                            Do Until Len(var1) = 2
        '                                intRow = intRow - 1
        '                                tbl.Cell(intRow, 2).Select()
        '                                var1 = .Selection.Text
        '                            Loop
        '                            intRow = intRow + 1
        '                            intNextPageRow = intRow
        '                            tbl.Cell(intNextPageRow, 1).Select()

        '                            If intTableID = 11 Then
        '                                'continue with column 1
        '                                var1 = .Selection.Text
        '                                Do Until Len(var1) = 2
        '                                    intRow = intRow - 1
        '                                    tbl.Cell(intRow, 1).Select()
        '                                    var1 = .Selection.Text
        '                                Loop
        '                            End If
        '                        End If

        '                        'format this row as page break
        '                        .Selection.ParagraphFormat.PageBreakBefore = True

        '                        'go back one row
        '                        intRow = intRow - 1
        '                        tbl.Cell(intRow, 1).Select()
        '                        intLastSearchRow = intRow

        '                        'insert rows
        '                        ctRealLegend = ctRealLegend + 1 'must add 1 row for 'next page'
        '                        'NO!! Been done already
        '                        'If boolPageNum Then
        '                        '    ctRealLegend = ctRealLegend + 1
        '                        'End If
        '                        'If boolDTStamp Then
        '                        '    ctRealLegend = ctRealLegend + 1
        '                        'End If
        '                        'ctRealLegend = ctRealLegend + 1 'sometimes wrap isn't happening properly. add another line

        '                        .Selection.InsertRowsBelow(ctRealLegend)
        '                        intLastLegend = intLastSearchRow + ctRealLegend

        '                        'remove any underlines
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '                        myr1 = .Selection.Range
        '                        intRows = tbl.Rows.Count
        '                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '                        'need border herer
        '                        tbl.Rows(intLastSearchRow).Select()
        '                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

        '                        myr1.Select()

        '                        boolIsLastPage = False

        '                    Else
        '                        myr1.Select()
        '                        'intLastSearchRow = intLastSearchRow + ctRealLegend
        '                    End If
        '                End If


        '                'merge the rows
        '                .Selection.Cells.Split(NumRows:=ctRealLegend, NumColumns:=1, MergeBeforeSplit:=True)
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                If boolFirstAnova And boolIsLastPage = False Then
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                End If
        '                If intTableID = 11 And boolIsLastPage = False Then
        '                    'NO!!
        '                    '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                End If

        '                '''wdd.visible = True

        '                If boolSmallFont Then
        '                    .Selection.Font.Size = .Selection.Font.Size - 1
        '                End If
        '                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                'format cells
        '                .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                With .Selection.ParagraphFormat
        '                    .LeftIndent = 36 'InchesToPoints(0.63)
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                End With
        '                With .Selection.ParagraphFormat
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                    .FirstLineIndent = -36 'InchesToPoints(-0.63)
        '                End With

        '            End If

        '            'ensure top is  bordered
        '            If intLastSearchRow = intRows Then
        '                .Selection.Tables.Item(1).Cell(intLastSearchRow, 1).Select()

        '                If intTableID = 11 Then
        '                Else
        '                    .Selection.SelectRow()
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                End If
        '            Else
        '                .Selection.Tables.Item(1).Cell(intLastSearchRow + 1, 1).Select()
        '                If intTableID = 11 Then
        '                Else
        '                    .Selection.SelectRow()
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                End If
        '            End If

        '            '''wdd.visible = True

        '            intRow = intLastSearchRow

        '            p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)


        '            'wdd.visible = True

        '            ctLCheck = ctRealLegend

        '            'first do forced legends

        '            For Count1 = 1 To intForcedLegends

        '                intRow = intRow + 1

        '                'Herehere
        '                .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                var1 = arrRealLegend(1, Count1)
        '                '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
        '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, vbTextCompare)
        '                If arrRealLegend(3, Count1) Then
        '                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                    .Selection.Font.Bold = False

        '                    Fonts = .Selection.Font.Size
        '                    .Selection.Font.Superscript = True
        '                    '.Selection.Font.Size = 12
        '                    .Selection.TypeText(Text:=CStr(var1))
        '                    .Selection.Font.Superscript = False
        '                    .Selection.Font.Size = Fonts
        '                Else
        '                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                    .Selection.Font.Bold = False
        '                    .Selection.TypeText(Text:=CStr(var1))
        '                End If
        '                If Len(var1) = 0 Then
        '                Else
        '                    .Selection.TypeText(Text:=vbTab)
        '                    .Selection.TypeText(Text:="=")
        '                    .Selection.TypeText(Text:=vbTab)
        '                End If
        '                var2 = arrRealLegend(2, Count1)
        '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, vbTextCompare)
        '                '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
        '                .Selection.TypeText(Text:=CStr(var2))

        '            Next

        '            'do extra legends
        '            intNL = 0
        '            If boolPageNum Then
        '                intRow = intRow + 1
        '                intNL = intNL + 1
        '                .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                intPN = intPN + 1
        '                If boolFullPageNum Then
        '                    .Selection.TypeText(Text:="Page " & intPN & " of ")
        '                    wrdSelection = wd.Selection()
        '                    With wd.ActiveDocument.Bookmarks
        '                        .Add(Range:=wrdSelection.Range, Name:="PN" & intPN)
        '                        .ShowHidden = False
        '                    End With
        '                    arrPN(intPN) = "PN" & intPN ' .Selection.Start
        '                    'format align right
        '                    .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                Else
        '                    .Selection.TypeText(Text:="Page " & intPN)
        '                    'format align right
        '                    .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                End If

        '            End If

        '            If boolDTStamp Then

        '                intRow = intRow + 1
        '                intNL = intNL + 1
        '                .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                .Selection.TypeText(Text:=strDateTimeStamp)
        '                'format align right
        '                .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

        '            End If

        '            If boolIsLastPage = False Then

        '                intRow = intRow + 1
        '                intNL = intNL + 1
        '                .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                .Selection.TypeText(Text:="Continued on next page")
        '                'format align right
        '                .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

        '            End If


        '            '*** check to see if legend has crossed over to next page
        '            'this can happen if legend rows wrap when they are long

        '            intRows = tbl.Rows.Count
        '            'tbl.Cell(intRows, 1).Select()
        '            tbl.Cell(intLastLegend, 1).Select()
        '            p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '            pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '            If p1 < pT Then
        '                'must undo last amount of work and re-do row insertion

        '                '****
        '                If boolPageNum And intPN <> 0 Then
        '                    If boolFullPageNum Then
        '                        Try
        '                            .ActiveDocument.Bookmarks.Item("PN" & intPN).Delete()
        '                        Catch ex As Exception

        '                        End Try

        '                    End If
        '                    intPN = intPN - 1
        '                End If
        '                '****

        '                'goto pagebreak and undo
        '                tbl.Cell(intLastLegend + 1, 1).Select()
        '                .Selection.ParagraphFormat.PageBreakBefore = False


        '                'delete new rows
        '                tbl.Cell(intLastSearchRow + 1, 1).Select()
        '                For Count1 = 1 To ctRealLegend
        '                    .Selection.Rows.Delete()
        '                    var1 = var1
        '                Next
        '                'the last delete action moved the cursor out of the table. Must move it back into the tbl
        '                intRows = tbl.Rows.Count
        '                tbl.Cell(intLastSearchRow, 1).Select()

        '                'remove any borders
        '                'tbl.Rows(intLastSearchRow + 1).Select()
        '                .Selection.SelectRow()
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '                If intTableID = 11 Then
        '                Else
        '                    'the last delete action moved the cursor out of the table. Must move it back into the tbl
        '                    intRows = tbl.Rows.Count
        '                    tbl.Cell(intRows, 1).Select()
        '                End If
        '                p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '                '*****
        '                'now do search to determine new intsplitrows
        '                If DoLegend Then
        '                    tRow = ctLegend
        '                Else

        '                    tbl.Cell(intFirstSearchRow, 1).Select()
        '                    .Selection.SelectRow()
        '                    int1 = intLastSearchRow - intFirstSearchRow
        '                    '.Selection.MoveDown(Unit:=wdLine, Count:=int1, Extend:=WdMovementType.wdExtend)
        '                    .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                    myr = .Selection.Range

        '                    'low look for body items
        '                    tRow = 0
        '                    For Count1 = 1 To ctLegend

        '                        var1 = arr(4, Count1)
        '                        If NZ(var1, False) Then
        '                            tRow = tRow + 1
        '                        Else
        '                            var1 = arr(1, Count1)
        '                            myr = .Selection.Range 'call this again because previous find action seems to wipe out the range

        '                            '''wdd.visible = True

        '                            boolFound = False
        '                            With myr.Find
        '                                .ClearFormatting()
        '                                .MatchCase = True
        '                                .MatchWholeWord = True
        '                                .Forward = True
        '                                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                                .Execute(FindText:=var1)

        '                                If .Found Then
        '                                    boolFound = True
        '                                Else
        '                                    boolFound = False
        '                                End If

        '                            End With

        '                            If boolFound Then 'legend needed
        '                                tRow = tRow + 1
        '                            End If
        '                        End If

        '                    Next

        '                End If

        '                'add a row for safety
        '                tRow = tRow + 1

        '                If boolIsLastPage Then
        '                Else
        '                    tRow = tRow + 1 'take into account 'continued on next page'
        '                End If
        '                If boolPageNum Then
        '                    tRow = tRow + 1
        '                End If
        '                If boolDTStamp Then
        '                    tRow = tRow + 1
        '                End If
        '                If intTableID = 3 Then
        '                    If boolSTATSREGR Then
        '                        tRow = tRow + 1
        '                    End If
        '                End If
        '                intSplitRows = tRow
        '                '*****

        '                'go back up intsplitrows

        '                'intRow = intRows - intSplitRows
        '                intRow = intLastSearchRow - intSplitRows


        '                If boolCarefulSplit Then
        '                    tbl.Cell(intRow, 2).Select()
        '                    var1 = .Selection.Text
        '                    Do Until Len(var1) = 2
        '                        intRow = intRow - 1
        '                        tbl.Cell(intRow, 2).Select()
        '                        var1 = .Selection.Text
        '                    Loop
        '                    intRow = intRow + 1
        '                    tbl.Cell(intRow, 1).Select()

        '                    If intTableID = 11 Then
        '                        'continue with column 1
        '                        var1 = .Selection.Text
        '                        Do Until Len(var1) = 2
        '                            intRow = intRow - 1
        '                            tbl.Cell(intRow, 1).Select()
        '                            var1 = .Selection.Text
        '                        Loop
        '                    End If
        '                End If

        '                'record this row
        '                intNextPageRow = intRow
        '                tbl.Cell(intNextPageRow, 1).Select()


        '                'format this row as page break
        '                .Selection.ParagraphFormat.PageBreakBefore = True

        '                'go back one row
        '                intRow = intRow - 1
        '                tbl.Cell(intRow, 1).Select()
        '                intLastSearchRow = intRow

        '                're-perform search

        '                '****

        '                'search for legend stuff again
        '                boolIsLastPage = False

        '                ctRealLegend = 0

        '                'badd!!!
        '                If DoLegend Then

        '                    For Count1 = 1 To ctLegend
        '                        ctRealLegend = ctRealLegend + 1
        '                        arrRealLegend(1, Count1) = arr(1, Count1)
        '                        arrRealLegend(2, Count1) = arr(2, Count1)
        '                        arrRealLegend(3, Count1) = arr(3, Count1)
        '                    Next
        '                    intForcedLegends = ctRealLegend

        '                    If boolIsLastPage Then
        '                    Else
        '                        ctRealLegend = ctRealLegend + 1 'take into accout 'continued on next page'
        '                        'ctRealLegend = ctRealLegend + 1 'sometimes wrap doesn't happen properly. Add another line'
        '                    End If
        '                    If boolPageNum Then
        '                        ctRealLegend = ctRealLegend + 1
        '                    End If
        '                    If boolDTStamp Then
        '                        ctRealLegend = ctRealLegend + 1
        '                    End If
        '                Else

        '                    'look for legends
        '                    tbl.Cell(intFirstSearchRow, 1).Select()
        '                    .Selection.SelectRow()
        '                    int1 = intLastSearchRow - intFirstSearchRow
        '                    .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                    myr = .Selection.Range

        '                    'look for legends
        '                    'first look for column header legend items
        '                    ctRealLegend = 0
        '                    For Count1 = 1 To ctLegend
        '                        var1 = arr(4, Count1)
        '                        If NZ(var1, False) Then
        '                            ctRealLegend = ctRealLegend + 1
        '                            arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                            arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                            arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                        End If
        '                    Next

        '                    'low look for body items
        '                    For Count1 = 1 To ctLegend

        '                        var1 = arr(4, Count1)
        '                        If NZ(var1, False) Then
        '                        Else
        '                            var1 = arr(1, Count1)
        '                            myr = .Selection.Range 'call this again because previous find action seems to wipe out the range
        '                            boolFound = False
        '                            With myr.Find
        '                                .ClearFormatting()
        '                                .MatchCase = True
        '                                .MatchWholeWord = True
        '                                .Forward = True
        '                                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                                .Execute(FindText:=var1)

        '                                If .Found Then
        '                                    boolFound = True
        '                                Else
        '                                    boolFound = False
        '                                End If

        '                            End With

        '                            If boolFound Then 'legend needed
        '                                ctRealLegend = ctRealLegend + 1
        '                                arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                                arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                                arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                            End If
        '                        End If

        '                    Next

        '                    intForcedLegends = ctRealLegend

        '                    'I'm here
        '                    If boolIsLastPage Then
        '                    Else
        '                        ctRealLegend = ctRealLegend + 1 'take into accout 'continued on next page'
        '                        'ctRealLegend = ctRealLegend + 1 'sometimes wrap still isn't going properly. Add another line
        '                    End If
        '                    If boolPageNum Then
        '                        ctRealLegend = ctRealLegend + 1
        '                    End If
        '                    If boolDTStamp Then
        '                        ctRealLegend = ctRealLegend + 1
        '                    End If
        '                End If



        '                '****

        '                'insert rows
        '                myr1 = .Selection.Range
        '                .Selection.InsertRowsBelow(ctRealLegend)

        '                'remove any underlines
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '                myr1 = .Selection.Range
        '                intRows = tbl.Rows.Count
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '                tbl.Rows(intLastSearchRow).Select()
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

        '                myr1.Select()

        '                'merge the rows
        '                .Selection.Cells.Split(NumRows:=ctRealLegend, NumColumns:=1, MergeBeforeSplit:=True)
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                If boolSmallFont Then
        '                    .Selection.Font.Size = .Selection.Font.Size - 1
        '                End If
        '                '            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                'format cells
        '                .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                With .Selection.ParagraphFormat
        '                    .LeftIndent = 36 'InchesToPoints(0.63)
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                End With
        '                With .Selection.ParagraphFormat
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                    .FirstLineIndent = -36 'InchesToPoints(-0.63)
        '                End With

        '                'do some bordering stuff
        '                tbl.Cell(intLastSearchRow, 1).Select()
        '                p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                '        p1 = .selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                'border bottom
        '                'Selection.Borders(wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                .Selection.SelectRow()
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

        '                tbl.Cell(intLastSearchRow + ctRealLegend, 1).Select()
        '                .Selection.SelectRow()
        '                'un-border bottom
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '                intRow = intLastSearchRow

        '                ctLCheck = ctRealLegend

        '                '****
        '                For Count1 = 1 To intForcedLegends

        '                    intRow = intRow + 1

        '                    'Herehere
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                    var1 = arrRealLegend(1, Count1)
        '                    '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
        '                    var1 = Replace(CStr(var1), "-", NBH, 1, -1, vbTextCompare)
        '                    If arrRealLegend(3, Count1) Then
        '                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                        .Selection.Font.Bold = False

        '                        Fonts = .Selection.Font.Size
        '                        .Selection.Font.Superscript = True
        '                        '.Selection.Font.Size = 12
        '                        .Selection.TypeText(Text:=CStr(var1))
        '                        .Selection.Font.Superscript = False
        '                        .Selection.Font.Size = Fonts
        '                    Else
        '                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                        .Selection.Font.Bold = False
        '                        .Selection.TypeText(Text:=CStr(var1))
        '                    End If
        '                    If Len(var1) = 0 Then
        '                    Else
        '                        .Selection.TypeText(Text:=vbTab)
        '                        .Selection.TypeText(Text:="=")
        '                        .Selection.TypeText(Text:=vbTab)
        '                    End If
        '                    var2 = arrRealLegend(2, Count1)
        '                    var2 = Replace(CStr(var2), "-", NBH, 1, -1, vbTextCompare)
        '                    '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
        '                    .Selection.TypeText(Text:=CStr(var2))

        '                Next

        '                'do extra legends
        '                intNL = 0
        '                If boolPageNum Then
        '                    intRow = intRow + 1
        '                    intNL = intNL + 1
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                    intPN = intPN + 1
        '                    If boolFullPageNum Then
        '                        .Selection.TypeText(Text:="Page " & intPN & " of ")
        '                        wrdSelection = wd.Selection()
        '                        With wd.ActiveDocument.Bookmarks
        '                            .Add(Range:=wrdSelection.Range, Name:="PN" & intPN)
        '                            .ShowHidden = False
        '                        End With
        '                        arrPN(intPN) = "PN" & intPN ' .Selection.Start
        '                        'format align right
        '                        .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                    Else
        '                        .Selection.TypeText(Text:="Page " & intPN)
        '                        'format align right
        '                        .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                    End If

        '                End If

        '                If boolDTStamp Then

        '                    intRow = intRow + 1
        '                    intNL = intNL + 1
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                    .Selection.TypeText(Text:=strDateTimeStamp)
        '                    'format align right
        '                    .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

        '                End If

        '                If boolIsLastPage = False Then

        '                    intRow = intRow + 1
        '                    intNL = intNL + 1
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                    .Selection.TypeText(Text:="Continued on next page")
        '                    'format align right
        '                    .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

        '                End If
        '                '****

        '                var1 = "a" 'debug
        '                intLastLegend = intRow ' intLastSearchRow + ctRealLegend
        '            Else
        '                intNextPageRow = intLastSearchRow
        '            End If

        '            'End If

        '            '***End check for legend page overflow

        '            'intNextPageRow = intNextPageRow + ctRealLegend
        '            intNextPageRow = intLastLegend + 1
        '            intRows = tbl.Rows.Count
        '            If intNextPageRow > intRows Then
        '                intNextPageRow = intRows
        '            End If

        '            If boolIsLastPage Then

        '            Else
        '                arrPageBreaks(2, intPageCount) = intNextPageRow - 1
        '                intPageCount = intPageCount + 1
        '                arrPageBreaks(1, intPageCount) = intNextPageRow
        '            End If

        '            'top border last data row + 1

        '            intFirstSearchRow = intNextPageRow
        '            intRow = intFirstSearchRow
        '            intRows = tbl.Rows.Count

        '            arrPageBreaks(2, intPageCount) = intRows

        '            tbl.Cell(intFirstSearchRow, 1).Select()
        '            p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '            '    p1 = .selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '            If boolIsLastPage Then
        '                Exit Do
        '            End If

        '        Loop

        '        If boolFullPageNum Then
        '            Dim pos As Int64
        '            Dim rngP As Word.Range
        '            For Count1 = 1 To intPN

        '                Try

        '                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="PN" & Count1)
        '                    .Selection.TypeText(Text:=CStr(intPN))

        '                Catch ex As Exception

        '                End Try

        '            Next

        '            'delete bookmarks
        '            For Count1 = 1 To intPN

        '                Try
        '                    .ActiveDocument.Bookmarks.Item("PN" & Count1).Delete()
        '                Catch ex As Exception

        '                End Try
        '            Next
        '        End If

        '        If gboolReadOnlyTables Then
        '            Call ReadOnlyTables(wd, tbl, ctHdRows, arrPageBreaks, intPageCount)
        '        End If

        '    End With

    End Sub





    Sub SplitTable_20150508(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean, ByVal boolCarefulSplit As Boolean, ByVal boolFirstAnova As Boolean, ByVal intTableID As Int64)

        '    'Sub SplitTable(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, ByVal arr As Object, ByVal strT As String,
        '    'ByVal DoLegend As Boolean, ByVal DoTable As Boolean, ByVal intSRow As Short, ByVal intSplitRows As Short, ByVal boolAutoFit As Boolean, ByVal boolSmallFont As Boolean)


        '    Dim vView
        '    vView = wd.ActiveWindow.ActivePane.View.Type
        '    wd.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView

        '    Dim boolBGS As Boolean
        '    boolBGS = wd.Options.BackgroundSave

        '    Dim boolPag As Boolean
        '    boolPag = wd.Options.Pagination
        '    wd.Options.Pagination = False

        '    wd.Options.BackgroundSave = False

        '    Dim intRows As Long
        '    Dim intRow As Long
        '    Dim Count1 As Long
        '    Dim Count2 As Long
        '    Dim Count3 As Long
        '    Dim str1 As String
        '    Dim str2 As String
        '    Dim str3 As String
        '    Dim var1, var2, var3, var4
        '    Dim tbl As Word.Table
        '    Dim p1 As Long
        '    Dim p2 As Long
        '    Dim pT As Long
        '    Dim intIncr As Integer
        '    Dim rowStart As Long 'intRow1
        '    Dim ctP As Long
        '    Dim intNextPageRow As Long
        '    Dim intLastSearchRow As Long
        '    Dim intFirstSearchRow As Long
        '    Dim int1 As Long
        '    Dim int2 As Long
        '    Dim myr As Word.Range
        '    Dim myr1 As Word.Range
        '    Dim ctRealLegend As Integer
        '    Dim arrRealLegend
        '    Dim Fonts
        '    Dim boolIsLastPage As Boolean
        '    Dim boolFound As Boolean

        '    ReDim arrRealLegend(3, ctLegend)

        '    Dim arrPageBreaks(2, 500) As Short
        '    Dim intPageCount As Int16
        '    Dim tRow As Int16
        '    Dim intLastLegend As Int16 = 0

        '    Dim chkR As Int16
        '    Dim chkM As Int16

        '    Dim strDateTimeStamp As String
        '    Dim arrPN(1000)
        '    Dim intPN As Short = 0
        '    Dim boolPageNum As Boolean = False
        '    Dim boolDTStamp As Boolean = LBOOLTABLEDTTIMESTAMP
        '    Dim boolFullPageNum As Boolean = False
        '    Dim ctLCheck As Short
        '    Dim wrdSelection As Word.Selection
        '    Dim intSplitRowsO As Short

        '    Dim intForcedLegends As Short = 0
        '    Dim intNL As Short = 0

        '    If StrComp(LcharSTPage, "[None]", CompareMethod.Text) = 0 Then
        '        boolPageNum = False
        '    Else
        '        boolPageNum = True
        '    End If

        '    If StrComp(LcharSTPage, "Page x", CompareMethod.Text) = 0 Then
        '        boolFullPageNum = False
        '    Else
        '        boolFullPageNum = True
        '    End If

        '    'strDateTimeStamp = Format(LTableDateTimeStamp, LDateFormat & " HH:mm:ss tt")
        '    strDateTimeStamp = Format(gdtReportDate, LDateFormat & " HH:mm:ss tt")


        '    intPageCount = 1
        '    arrPageBreaks(1, intPageCount) = ctHdRows + 1

        '    wd.Visible = True

        '    With wd

        '        '''wdd.visible = True

        '        tbl = .Selection.Tables.Item(1)
        '        intSplitRowsO = intSplitRows

        '        If boolPSL Then
        '        Else
        '            intSplitRows = 1
        '        End If

        '        If boolPageNum Then
        '            intSplitRows = intSplitRows + 1
        '        End If
        '        If boolDTStamp Then
        '            intSplitRows = intSplitRows + 1
        '        End If
        '        intSplitRows = intSplitRows + 1 'account for Continued on next page

        '        'if orientation is portrait, add a few rows
        '        If boolPSL Then
        '            If .Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait Then
        '                intSplitRows = intSplitRows + 2 'for safety in case legends rap
        '            End If
        '        End If

        '        'intSplitRows = intSplitRows + 2 'for safety in case legends rap



        '        'first get rows
        '        Dim numPages As Int64
        '        intRows = tbl.Rows.Count
        '        tbl.Cell(1, 1).Select()
        '        p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '        tbl.Cell(intRows, 1).Select()
        '        pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '        numPages = pT - p1 + 1

        '        tbl.Cell(1, 1).Select()
        '        If p1 = pT Then 'no need to evaluate
        '            intRows = tbl.Rows.Count
        '            boolIsLastPage = True
        '        Else
        '            boolIsLastPage = False
        '        End If

        '        'begin looking for next page
        '        intRow = ctHdRows + 1
        '        ctP = 0

        '        If .Selection.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
        '            intIncr = 20
        '        Else
        '            intIncr = 40
        '        End If

        '        intFirstSearchRow = ctHdRows + 1
        '        boolIsLastPage = False

        '        chkM = 200
        '        chkR = 0

        '        Do Until intRow >= intRows

        '            chkR = chkR + 1
        '            If chkR > chkM Then
        '                Exit Do
        '            End If

        '            ctP = ctP + 1
        '            str1 = strT & vbCr & "Formatting legend..." & ctP & " of " & numPages
        '            frmH.lblProgress.Text = str1
        '            frmH.lblProgress.Refresh()


        '            tbl.Cell(intRow, 1).Select()
        '            p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '            intRows = tbl.Rows.Count

        '            tbl.Cell(intRows, 1).Select()
        '            pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '            If p1 = pT Then 'don't look for table split
        '                intRow = intRows
        '                intLastSearchRow = intRows
        '                boolIsLastPage = True
        '            Else 'look for table split

        '                intRow = intRow + intIncr
        '                If intRow > intRows Then
        '                    intRow = intRows
        '                End If

        '                tbl.Cell(intRow, 1).Select()

        '                'check to see if intincr is too big
        '                p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                If p2 > p1 Then 'intincr is too big
        '                    Do Until p2 = p1
        '                        intRow = intRow - intIncr
        '                        intIncr = intIncr - 5
        '                        intRow = intRow + intIncr
        '                        tbl.Cell(intRow, 1).Select()
        '                        p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    Loop
        '                End If

        '                p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                p2 = p1 + 1
        '                Do Until p1 = p2
        '                    If intRow >= intRows Then
        '                        Exit Do
        '                    End If
        '                    intRow = intRow + 1
        '                    tbl.Cell(intRow, 1).Select()
        '                    p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    If intRow = intRows Then
        '                        Exit Do
        '                    End If
        '                Loop

        '                '*****
        '                'wdd.visible = True
        '                'now do search to determine new intsplitrows
        '                tRow = 0
        '                If DoLegend Then
        '                    If boolPSL Then
        '                        tRow = ctLegend
        '                    Else
        '                        tRow = intSplitRows
        '                    End If
        '                Else

        '                    tbl.Cell(intFirstSearchRow, 1).Select()
        '                    .Selection.SelectRow()
        '                    intLastSearchRow = intRow - 1
        '                    int1 = intLastSearchRow - intFirstSearchRow
        '                    If boolPSL Then
        '                        '.Selection.MoveDown(Unit:=wdLine, Count:=int1, Extend:=WdMovementType.wdExtend)
        '                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                        myr = .Selection.Range
        '                    End If

        '                    'low look for body items
        '                    tRow = 0
        '                    If boolPSL Then
        '                        For Count1 = 1 To ctLegend

        '                            var1 = arr(4, Count1)
        '                            If NZ(var1, False) Then
        '                                tRow = tRow + 1
        '                            Else
        '                                var1 = arr(1, Count1)
        '                                myr = .Selection.Range 'call this again because previous find action seems to wipe out the range

        '                                '''wdd.visible = True

        '                                boolFound = False
        '                                With myr.Find
        '                                    .ClearFormatting()
        '                                    .MatchCase = True
        '                                    .MatchWholeWord = True
        '                                    .Forward = True
        '                                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                                    .Execute(FindText:=var1)

        '                                    If .Found Then
        '                                        boolFound = True
        '                                    Else
        '                                        boolFound = False
        '                                    End If

        '                                End With

        '                                If boolFound Then 'legend needed
        '                                    tRow = tRow + 1
        '                                End If
        '                            End If

        '                        Next

        '                    End If

        '                End If

        '                'add a row for safety
        '                tRow = tRow + 1

        '                If boolIsLastPage Then
        '                Else
        '                    tRow = tRow + 1 'take into account 'continued on next page'
        '                End If
        '                If boolPageNum Then
        '                    tRow = tRow + 1
        '                End If
        '                If boolDTStamp Then
        '                    tRow = tRow + 1
        '                End If

        '                If intTableID = 3 Then
        '                    If boolSTATSREGR Then
        '                        tRow = tRow + 1
        '                    End If
        '                End If
        '                intSplitRows = tRow

        '                '*****


        '                'now retreat ctLegend rows
        '                'intRow = intRow - intSplitRows
        '                If boolFirstAnova Then

        '                Else
        '                    intRow = intRow - intSplitRows
        '                    tbl.Cell(intRow, 1).Select()
        '                    If boolCarefulSplit Then
        '                        tbl.Cell(intRow, 2).Select()
        '                        var1 = .Selection.Text
        '                        Do Until Len(var1) = 2
        '                            '''wdd.visible = True

        '                            var2 = Len(var1)
        '                            intRow = intRow - 1
        '                            tbl.Cell(intRow, 2).Select()
        '                            var1 = .Selection.Text
        '                        Loop
        '                        intRow = intRow + 1
        '                        tbl.Cell(intRow, 1).Select()

        '                        If intTableID = 11 Then
        '                            'continue with column 1
        '                            var1 = .Selection.Text
        '                            Do Until Len(var1) = 2
        '                                intRow = intRow - 1
        '                                tbl.Cell(intRow, 1).Select()
        '                                var1 = .Selection.Text
        '                            Loop
        '                        End If
        '                    End If

        '                    'format this row as page break
        '                    .Selection.ParagraphFormat.PageBreakBefore = True

        '                    ' 'wdd.visible = True

        '                    If intTableID = 11 Then
        '                        'PageBreakaBefore for some reason  top-borders the selection
        '                        'must remove underline
        '                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        '.Selection.Tables.Item(1).Cell(intFirstAnova, 1).Select()

        '                    End If
        '                End If

        '                'record this row
        '                intNextPageRow = intRow
        '                tbl.Cell(intNextPageRow, 1).Select()

        '                'go back one row
        '                intRow = intRow - 1
        '                tbl.Cell(intRow, 1).Select()
        '                intLastSearchRow = intRow

        '            End If

        '            'select range and look for legends

        '            ctRealLegend = 0

        '            If DoLegend Then

        '                For Count1 = 1 To ctLegend
        '                    ctRealLegend = ctRealLegend + 1
        '                    arrRealLegend(1, Count1) = arr(1, Count1)
        '                    arrRealLegend(2, Count1) = arr(2, Count1)
        '                    arrRealLegend(3, Count1) = arr(3, Count1)
        '                Next
        '                intForcedLegends = ctRealLegend

        '                If boolIsLastPage Then
        '                Else
        '                    ctRealLegend = ctRealLegend + 1 'take into accout 'continued on next page'
        '                    'ctRealLegend = ctRealLegend + 1 'sometimes wrap isn't occurring properly. Add another line
        '                End If
        '                If boolPageNum Then
        '                    ctRealLegend = ctRealLegend + 1
        '                End If
        '                If boolDTStamp Then
        '                    ctRealLegend = ctRealLegend + 1
        '                End If
        '            Else

        '                tbl.Cell(intFirstSearchRow, 1).Select()
        '                .Selection.SelectRow()
        '                int1 = intLastSearchRow - intFirstSearchRow
        '                If boolPSL Then
        '                    '.Selection.MoveDown(Unit:=wdLine, Count:=int1, Extend:=WdMovementType.wdExtend)
        '                    .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                    myr = .Selection.Range
        '                End If

        '                'look for legends
        '                'first look for column header legend items
        '                ctRealLegend = 0
        '                For Count1 = 1 To ctLegend
        '                    var1 = arr(4, Count1)
        '                    If NZ(var1, False) Then
        '                        ctRealLegend = ctRealLegend + 1
        '                        arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                        arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                        arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                    End If
        '                Next

        '                'low look for body items
        '                If boolPSL Then
        '                    For Count1 = 1 To ctLegend

        '                        var1 = arr(4, Count1)
        '                        If NZ(var1, False) Then
        '                        Else
        '                            var1 = arr(1, Count1)
        '                            myr = .Selection.Range 'call this again because previous find action seems to wipe out the range

        '                            '''wdd.visible = True

        '                            boolFound = False

        '                            With myr.Find
        '                                .ClearFormatting()
        '                                .MatchCase = True
        '                                .MatchWholeWord = True
        '                                .Forward = True
        '                                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                                .Execute(FindText:=var1)

        '                                If .Found Then
        '                                    boolFound = True
        '                                Else
        '                                    boolFound = False
        '                                End If

        '                            End With


        '                            If boolFound Then 'legend needed
        '                                ctRealLegend = ctRealLegend + 1
        '                                arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                                arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                                arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                            End If
        '                        End If

        '                    Next
        '                End If


        '                intForcedLegends = ctRealLegend

        '                If boolIsLastPage Then
        '                Else
        '                    ctRealLegend = ctRealLegend + 1 'take into account 'continued on next page'
        '                    'ctRealLegend = ctRealLegend + 1 'sometimes the final product just doesn't wrap properly. Add another line to ctreallegend
        '                End If
        '                If boolPageNum Then
        '                    ctRealLegend = ctRealLegend + 1
        '                End If
        '                If boolDTStamp Then
        '                    ctRealLegend = ctRealLegend + 1
        '                End If
        '            End If



        '            'add legend rows
        '            tbl.Cell(intLastSearchRow, 1).Select()
        '            If boolPSL Then


        '            End If

        '            If ctRealLegend = 0 Then
        '            Else

        '                .Selection.InsertRowsBelow(intSplitRows)
        '                intLastLegend = intLastSearchRow + intSplitRows

        '                'remove any borders
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '                myr1 = .Selection.Range
        '                intRows = tbl.Rows.Count

        '                'remove borders removes a desired border
        '                'put it back
        '                'wd.Visible = True
        '                tbl.Rows(intLastSearchRow).Select()
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                tbl.Cell(intLastSearchRow, 1).Select()
        '                myr1.Select()

        '                If boolIsLastPage Then 'check to see if this went across page

        '                    'do some bordering stuff
        '                    tbl.Cell(intLastSearchRow, 1).Select()
        '                    p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    'border bottom
        '                    'Selection.Borders(wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                    .Selection.SelectRow()
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

        '                    tbl.Cell(intLastSearchRow + ctRealLegend, 1).Select()
        '                    .Selection.SelectRow()
        '                    'un-border bottom
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '                    p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '                    '''wdd.visible = True

        '                    If p1 <> pT Then 'must undo

        '                        '****
        '                        If boolPageNum And intPN <> 0 Then
        '                            If boolFullPageNum Then
        '                                Try
        '                                    .ActiveDocument.Bookmarks.Item("PN" & intPN).Delete()
        '                                Catch ex As Exception

        '                                End Try

        '                            End If
        '                            intPN = intPN - 1
        '                        End If
        '                        '****

        '                        'jeez
        '                        Dim intT As Int64

        '                        'first count the number of rows needed
        '                        intT = tbl.Rows.Count
        '                        Dim intT1 As Int64

        '                        tbl.Cell(intT, 1).Select()
        '                        pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                        intT1 = 0
        '                        Do Until pT = p1
        '                            intT1 = intT1 + 1
        '                            intT = intT - 1
        '                            tbl.Cell(intT, 1).Select()
        '                            pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                        Loop

        '                        'delete new rows
        '                        tbl.Cell(intLastSearchRow + 1, 1).Select()
        '                        For Count1 = 1 To ctRealLegend
        '                            .Selection.Rows.Delete()
        '                            var1 = var1
        '                        Next

        '                        'the last delete action moved the cursor out of the table. Must move it back into the tbl
        '                        intRows = tbl.Rows.Count

        '                        tbl.Cell(intLastSearchRow, 1).Select()
        '                        p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                        pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '                        'go back up intT rows
        '                        intRow = intRows - intT1 - 1
        '                        'record this row
        '                        intNextPageRow = intRow
        '                        tbl.Cell(intNextPageRow, 1).Select()
        '                        If boolCarefulSplit Then 'look for next blank row
        '                            tbl.Cell(intNextPageRow, 2).Select()
        '                            var1 = .Selection.Text
        '                            Do Until Len(var1) = 2
        '                                intRow = intRow - 1
        '                                tbl.Cell(intRow, 2).Select()
        '                                var1 = .Selection.Text
        '                            Loop
        '                            intRow = intRow + 1
        '                            intNextPageRow = intRow
        '                            tbl.Cell(intNextPageRow, 1).Select()

        '                            If intTableID = 11 Then
        '                                'continue with column 1
        '                                var1 = .Selection.Text
        '                                Do Until Len(var1) = 2
        '                                    intRow = intRow - 1
        '                                    tbl.Cell(intRow, 1).Select()
        '                                    var1 = .Selection.Text
        '                                Loop
        '                            End If
        '                        End If

        '                        'format this row as page break
        '                        .Selection.ParagraphFormat.PageBreakBefore = True

        '                        'go back one row
        '                        intRow = intRow - 1
        '                        tbl.Cell(intRow, 1).Select()
        '                        intLastSearchRow = intRow

        '                        'insert rows
        '                        ctRealLegend = ctRealLegend + 1 'must add 1 row for 'next page'
        '                        'NO!! Been done already
        '                        'If boolPageNum Then
        '                        '    ctRealLegend = ctRealLegend + 1
        '                        'End If
        '                        'If boolDTStamp Then
        '                        '    ctRealLegend = ctRealLegend + 1
        '                        'End If
        '                        'ctRealLegend = ctRealLegend + 1 'sometimes wrap isn't happening properly. add another line

        '                        .Selection.InsertRowsBelow(intSplitRows)
        '                        intLastLegend = intLastSearchRow + intSplitRows

        '                        'remove any underlines
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '                        myr1 = .Selection.Range
        '                        intRows = tbl.Rows.Count
        '                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '                        'need border herer
        '                        tbl.Rows(intLastSearchRow).Select()
        '                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

        '                        myr1.Select()

        '                        boolIsLastPage = False

        '                    Else
        '                        myr1.Select()
        '                        'intLastSearchRow = intLastSearchRow + ctRealLegend
        '                    End If

        '                End If


        '                'merge the rows
        '                If boolPSL Then
        '                    .Selection.Cells.Split(NumRows:=ctRealLegend, NumColumns:=1, MergeBeforeSplit:=True)


        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '                    If boolFirstAnova And boolIsLastPage = False Then
        '                        .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                    End If
        '                    If intTableID = 11 And boolIsLastPage = False Then
        '                        'NO!!
        '                        '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                    End If

        '                    '''wdd.visible = True

        '                    If boolSmallFont Then
        '                        .Selection.Font.Size = .Selection.Font.Size - 1
        '                    End If
        '                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                    'format cells
        '                    .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                    With .Selection.ParagraphFormat
        '                        .LeftIndent = 36 'InchesToPoints(0.63)
        '                        .SpaceBefore = 0
        '                        .SpaceBeforeAuto = False
        '                        .SpaceAfter = 0
        '                        .SpaceAfterAuto = False
        '                    End With
        '                    With .Selection.ParagraphFormat
        '                        .SpaceBefore = 0
        '                        .SpaceBeforeAuto = False
        '                        .SpaceAfter = 0
        '                        .SpaceAfterAuto = False
        '                        .FirstLineIndent = -36 'InchesToPoints(-0.63)
        '                    End With
        '                End If

        '            End If 'right here

        '            'ensure top is  bordered
        '            If intLastSearchRow = intRows Then
        '                .Selection.Tables.Item(1).Cell(intLastSearchRow, 1).Select()

        '                If intTableID = 11 Then
        '                Else
        '                    .Selection.SelectRow()
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                End If
        '            Else
        '                .Selection.Tables.Item(1).Cell(intLastSearchRow + 1, 1).Select()
        '                If intTableID = 11 Then
        '                Else
        '                    .Selection.SelectRow()
        '                    .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                End If
        '            End If

        '            '''wdd.visible = True

        '            intRow = intLastSearchRow

        '            p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)


        '            'wdd.visible = True

        '            ctLCheck = ctRealLegend

        '            'first do forced legends
        '            If boolPSL Then

        '                For Count1 = 1 To intForcedLegends

        '                    intRow = intRow + 1

        '                    'Herehere
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                    var1 = arrRealLegend(1, Count1)
        '                    '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
        '                    var1 = Replace(CStr(var1), "-", NBH, 1, -1, vbTextCompare)
        '                    If arrRealLegend(3, Count1) Then
        '                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                        .Selection.Font.Bold = False

        '                        typeInSuperscript(wd, CStr(var1))
        '                    Else
        '                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                        .Selection.Font.Bold = False
        '                        .Selection.TypeText(Text:=CStr(var1))
        '                    End If
        '                    If Len(var1) = 0 Then
        '                    Else
        '                        .Selection.TypeText(Text:=vbTab)
        '                        .Selection.TypeText(Text:="=")
        '                        .Selection.TypeText(Text:=vbTab)
        '                    End If
        '                    var2 = arrRealLegend(2, Count1)
        '                    var2 = Replace(CStr(var2), "-", NBH, 1, -1, vbTextCompare)
        '                    '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
        '                    .Selection.TypeText(Text:=CStr(var2))

        '                Next

        '            End If

        '            'do extra legends
        '            intNL = 0
        '            If boolPageNum Then
        '                intRow = intRow + 1
        '                intNL = intNL + 1
        '                .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                intPN = intPN + 1
        '                If boolFullPageNum Then
        '                    .Selection.TypeText(Text:="Page " & intPN & " of ")
        '                    wrdSelection = wd.Selection()
        '                    With wd.ActiveDocument.Bookmarks
        '                        .Add(Range:=wrdSelection.Range, Name:="PN" & intPN)
        '                        .ShowHidden = False
        '                    End With
        '                    arrPN(intPN) = "PN" & intPN ' .Selection.Start
        '                    'format align right
        '                    .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                Else
        '                    .Selection.TypeText(Text:="Page " & intPN)
        '                    'format align right
        '                    .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                End If

        '            End If

        '            If boolDTStamp Then

        '                intRow = intRow + 1
        '                intNL = intNL + 1
        '                .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                .Selection.TypeText(Text:=strDateTimeStamp)
        '                'format align right
        '                .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

        '            End If

        '            If boolIsLastPage = False Then

        '                intRow = intRow + 1
        '                intNL = intNL + 1
        '                .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                .Selection.TypeText(Text:="Continued on next page")
        '                'format align right
        '                .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

        '            End If


        '            '*** check to see if legend has crossed over to next page
        '            'this can happen if legend rows wrap when they are long

        '            intRows = tbl.Rows.Count
        '            'tbl.Cell(intRows, 1).Select()
        '            tbl.Cell(intLastLegend, 1).Select()
        '            p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '            pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '            If p1 < pT Then
        '                'must undo last amount of work and re-do row insertion

        '                '****
        '                If boolPageNum And intPN <> 0 Then
        '                    If boolFullPageNum Then
        '                        Try
        '                            .ActiveDocument.Bookmarks.Item("PN" & intPN).Delete()
        '                        Catch ex As Exception

        '                        End Try

        '                    End If
        '                    intPN = intPN - 1
        '                End If
        '                '****

        '                'goto pagebreak and undo
        '                tbl.Cell(intLastLegend + 1, 1).Select()
        '                .Selection.ParagraphFormat.PageBreakBefore = False


        '                'delete new rows
        '                tbl.Cell(intLastSearchRow + 1, 1).Select()
        '                For Count1 = 1 To ctRealLegend
        '                    .Selection.Rows.Delete()
        '                    var1 = var1
        '                Next
        '                'the last delete action moved the cursor out of the table. Must move it back into the tbl
        '                intRows = tbl.Rows.Count
        '                tbl.Cell(intLastSearchRow, 1).Select()

        '                'remove any borders
        '                'tbl.Rows(intLastSearchRow + 1).Select()
        '                .Selection.SelectRow()
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '                If intTableID = 11 Then
        '                Else
        '                    'the last delete action moved the cursor out of the table. Must move it back into the tbl
        '                    intRows = tbl.Rows.Count
        '                    tbl.Cell(intRows, 1).Select()
        '                End If
        '                p2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                pT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '                '*****
        '                'now do search to determine new intsplitrows
        '                tRow = 0
        '                If DoLegend Then
        '                    If boolPSL Then
        '                        tRow = ctLegend
        '                    Else
        '                        tRow = intSplitRows
        '                    End If

        '                Else

        '                    tbl.Cell(intFirstSearchRow, 1).Select()
        '                    .Selection.SelectRow()
        '                    int1 = intLastSearchRow - intFirstSearchRow
        '                    If boolPSL Then
        '                        '.Selection.MoveDown(Unit:=wdLine, Count:=int1, Extend:=WdMovementType.wdExtend)
        '                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                        myr = .Selection.Range
        '                    End If


        '                    'low look for body items
        '                    tRow = 0
        '                    If boolPSL Then
        '                        For Count1 = 1 To ctLegend

        '                            var1 = arr(4, Count1)
        '                            If NZ(var1, False) Then
        '                                tRow = tRow + 1
        '                            Else
        '                                var1 = arr(1, Count1)
        '                                myr = .Selection.Range 'call this again because previous find action seems to wipe out the range

        '                                '''wdd.visible = True

        '                                boolFound = False

        '                                With myr.Find
        '                                    .ClearFormatting()
        '                                    .MatchCase = True
        '                                    .MatchWholeWord = True
        '                                    .Forward = True
        '                                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                                    .Execute(FindText:=var1)

        '                                    If .Found Then
        '                                        boolFound = True
        '                                    Else
        '                                        boolFound = False
        '                                    End If

        '                                End With
        '                            End If

        '                            If boolFound Then 'legend needed
        '                                tRow = tRow + 1
        '                            End If

        '                        Next
        '                    End If


        '                End If

        '                'add a row for safety
        '                tRow = tRow + 1

        '                If boolIsLastPage Then
        '                Else
        '                    tRow = tRow + 1 'take into account 'continued on next page'
        '                End If
        '                If boolPageNum Then
        '                    tRow = tRow + 1
        '                End If
        '                If boolDTStamp Then
        '                    tRow = tRow + 1
        '                End If
        '                If intTableID = 3 Then
        '                    If boolSTATSREGR Then
        '                        tRow = tRow + 1
        '                    End If
        '                End If
        '                intSplitRows = tRow
        '                '*****

        '                'go back up intsplitrows

        '                'intRow = intRows - intSplitRows
        '                intRow = intLastSearchRow - intSplitRows


        '                If boolCarefulSplit Then
        '                    tbl.Cell(intRow, 2).Select()
        '                    var1 = .Selection.Text
        '                    Do Until Len(var1) = 2
        '                        intRow = intRow - 1
        '                        tbl.Cell(intRow, 2).Select()
        '                        var1 = .Selection.Text
        '                    Loop
        '                    intRow = intRow + 1
        '                    tbl.Cell(intRow, 1).Select()

        '                    If intTableID = 11 Then
        '                        'continue with column 1
        '                        var1 = .Selection.Text
        '                        Do Until Len(var1) = 2
        '                            intRow = intRow - 1
        '                            tbl.Cell(intRow, 1).Select()
        '                            var1 = .Selection.Text
        '                        Loop
        '                    End If
        '                End If

        '                'record this row
        '                intNextPageRow = intRow
        '                tbl.Cell(intNextPageRow, 1).Select()


        '                'format this row as page break
        '                .Selection.ParagraphFormat.PageBreakBefore = True

        '                'go back one row
        '                intRow = intRow - 1
        '                tbl.Cell(intRow, 1).Select()
        '                intLastSearchRow = intRow

        '                're-perform search

        '                '****

        '                'search for legend stuff again
        '                boolIsLastPage = False

        '                ctRealLegend = 0

        '                'badd!!!
        '                If DoLegend Then

        '                    For Count1 = 1 To ctLegend
        '                        ctRealLegend = ctRealLegend + 1
        '                        arrRealLegend(1, Count1) = arr(1, Count1)
        '                        arrRealLegend(2, Count1) = arr(2, Count1)
        '                        arrRealLegend(3, Count1) = arr(3, Count1)
        '                    Next
        '                    intForcedLegends = ctRealLegend

        '                    If boolIsLastPage Then
        '                    Else
        '                        ctRealLegend = ctRealLegend + 1 'take into accout 'continued on next page'
        '                        'ctRealLegend = ctRealLegend + 1 'sometimes wrap doesn't happen properly. Add another line'
        '                    End If
        '                    If boolPageNum Then
        '                        ctRealLegend = ctRealLegend + 1
        '                    End If
        '                    If boolDTStamp Then
        '                        ctRealLegend = ctRealLegend + 1
        '                    End If
        '                Else

        '                    'look for legends
        '                    tbl.Cell(intFirstSearchRow, 1).Select()
        '                    .Selection.SelectRow()
        '                    int1 = intLastSearchRow - intFirstSearchRow
        '                    If boolPSL Then
        '                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                        myr = .Selection.Range
        '                    End If

        '                    'look for legends
        '                    'first look for column header legend items
        '                    ctRealLegend = 0
        '                    For Count1 = 1 To ctLegend
        '                        var1 = arr(4, Count1)
        '                        If NZ(var1, False) Then
        '                            ctRealLegend = ctRealLegend + 1
        '                            arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                            arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                            arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                        End If
        '                    Next

        '                    'low look for body items
        '                    If boolPSL Then
        '                        For Count1 = 1 To ctLegend

        '                            var1 = arr(4, Count1)
        '                            If NZ(var1, False) Then
        '                            Else
        '                                var1 = arr(1, Count1)
        '                                myr = .Selection.Range 'call this again because previous find action seems to wipe out the range
        '                                boolFound = False
        '                                With myr.Find
        '                                    .ClearFormatting()
        '                                    .MatchCase = True
        '                                    .MatchWholeWord = True
        '                                    .Forward = True
        '                                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                                    .Execute(FindText:=var1)

        '                                    If .Found Then
        '                                        boolFound = True
        '                                    Else
        '                                        boolFound = False
        '                                    End If

        '                                End With

        '                                If boolFound Then 'legend needed
        '                                    ctRealLegend = ctRealLegend + 1
        '                                    arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                                    arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                                    arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                                End If
        '                            End If

        '                        Next
        '                    End If


        '                    intForcedLegends = ctRealLegend

        '                    'I'm here
        '                    If boolIsLastPage Then
        '                    Else
        '                        ctRealLegend = ctRealLegend + 1 'take into accout 'continued on next page'
        '                        'ctRealLegend = ctRealLegend + 1 'sometimes wrap still isn't going properly. Add another line
        '                    End If
        '                    If boolPageNum Then
        '                        ctRealLegend = ctRealLegend + 1
        '                    End If
        '                    If boolDTStamp Then
        '                        ctRealLegend = ctRealLegend + 1
        '                    End If
        '                End If



        '                '****

        '                'insert rows
        '                myr1 = .Selection.Range
        '                .Selection.InsertRowsBelow(intSplitRows)

        '                'remove any underlines
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '                myr1 = .Selection.Range
        '                intRows = tbl.Rows.Count
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '                tbl.Rows(intLastSearchRow).Select()
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

        '                myr1.Select()

        '                'merge the rows
        '                .Selection.Cells.Split(NumRows:=ctRealLegend, NumColumns:=1, MergeBeforeSplit:=True)
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '                If boolSmallFont Then
        '                    .Selection.Font.Size = .Selection.Font.Size - 1
        '                End If
        '                '            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                'format cells
        '                .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                With .Selection.ParagraphFormat
        '                    .LeftIndent = 36 'InchesToPoints(0.63)
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                End With
        '                With .Selection.ParagraphFormat
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                    .FirstLineIndent = -36 'InchesToPoints(-0.63)
        '                End With

        '                'do some bordering stuff
        '                tbl.Cell(intLastSearchRow, 1).Select()
        '                p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                '        p1 = .selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                'border bottom
        '                'Selection.Borders(wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        '                .Selection.SelectRow()
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle

        '                tbl.Cell(intLastSearchRow + ctRealLegend, 1).Select()
        '                .Selection.SelectRow()
        '                'un-border bottom
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '                intRow = intLastSearchRow

        '                ctLCheck = ctRealLegend

        '                '****
        '                If boolPSL Then
        '                    For Count1 = 1 To intForcedLegends

        '                        intRow = intRow + 1

        '                        'Herehere
        '                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                        var1 = arrRealLegend(1, Count1)
        '                        '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
        '                        var1 = Replace(CStr(var1), "-", NBH, 1, -1, vbTextCompare)
        '                        If arrRealLegend(3, Count1) Then
        '                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                            .Selection.Font.Bold = False

        '                            typeInSuperscript(wd, CStr(var1))
        '                        Else
        '                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                            .Selection.Font.Bold = False
        '                            .Selection.TypeText(Text:=CStr(var1))
        '                        End If
        '                        If Len(var1) = 0 Then
        '                        Else
        '                            .Selection.TypeText(Text:=vbTab)
        '                            .Selection.TypeText(Text:="=")
        '                            .Selection.TypeText(Text:=vbTab)
        '                        End If
        '                        var2 = arrRealLegend(2, Count1)
        '                        var2 = Replace(CStr(var2), "-", NBH, 1, -1, vbTextCompare)
        '                        '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
        '                        .Selection.TypeText(Text:=CStr(var2))

        '                    Next

        '                Else
        '                    If boolIsLastPage Then

        '                        'add extra rows
        '                        'wd.Visible = True
        '                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                        '.Selection.InsertRowsBelow(intSplitRowsO) '? should be intSplitRows?
        '                        '.Selection.InsertRowsBelow(intSplitRows) '? should be intSplitRows?
        '                        .Selection.InsertRowsBelow(intForcedLegends) '? should be intSplitRows?

        '                        For Count1 = 1 To intForcedLegends

        '                            intRow = intRow + 1

        '                            'Herehere
        '                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                            var1 = arrRealLegend(1, Count1)
        '                            '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
        '                            var1 = Replace(CStr(var1), "-", NBH, 1, -1, vbTextCompare)
        '                            If arrRealLegend(3, Count1) Then
        '                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                                .Selection.Font.Bold = False

        '                                typeInSuperscript(wd, CStr(var1))
        '                            Else
        '                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                                .Selection.Font.Bold = False
        '                                .Selection.TypeText(Text:=CStr(var1))
        '                            End If
        '                            If Len(var1) = 0 Then
        '                            Else
        '                                .Selection.TypeText(Text:=vbTab)
        '                                .Selection.TypeText(Text:="=")
        '                                .Selection.TypeText(Text:=vbTab)
        '                            End If
        '                            var2 = arrRealLegend(2, Count1)
        '                            var2 = Replace(CStr(var2), "-", NBH, 1, -1, vbTextCompare)
        '                            '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
        '                            .Selection.TypeText(Text:=CStr(var2))

        '                        Next
        '                    End If
        '                End If

        '                'do extra legends
        '                intNL = 0
        '                If boolPageNum Then
        '                    intRow = intRow + 1
        '                    intNL = intNL + 1
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                    intPN = intPN + 1
        '                    If boolFullPageNum Then
        '                        .Selection.TypeText(Text:="Page " & intPN & " of ")
        '                        wrdSelection = wd.Selection()
        '                        With wd.ActiveDocument.Bookmarks
        '                            .Add(Range:=wrdSelection.Range, Name:="PN" & intPN)
        '                            .ShowHidden = False
        '                        End With
        '                        arrPN(intPN) = "PN" & intPN ' .Selection.Start
        '                        'format align right
        '                        .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                    Else
        '                        .Selection.TypeText(Text:="Page " & intPN)
        '                        'format align right
        '                        .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                    End If

        '                End If

        '                If boolDTStamp Then

        '                    intRow = intRow + 1
        '                    intNL = intNL + 1
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                    .Selection.TypeText(Text:=strDateTimeStamp)
        '                    'format align right
        '                    .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

        '                End If

        '                If boolIsLastPage = False Then

        '                    intRow = intRow + 1
        '                    intNL = intNL + 1
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                    .Selection.TypeText(Text:="Continued on next page")
        '                    'format align right
        '                    .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

        '                End If
        '                '****

        '                var1 = "a" 'debug
        '                intLastLegend = intRow ' intLastSearchRow + ctRealLegend
        '            Else
        '                intNextPageRow = intLastSearchRow
        '            End If

        '            'End If

        '            '***End check for legend page overflow

        '            'intNextPageRow = intNextPageRow + ctRealLegend
        '            intNextPageRow = intLastLegend + 1
        '            intRows = tbl.Rows.Count
        '            If intNextPageRow > intRows Then
        '                intNextPageRow = intRows
        '            End If

        '            If boolIsLastPage Then

        '            Else
        '                arrPageBreaks(2, intPageCount) = intNextPageRow - 1
        '                intPageCount = intPageCount + 1
        '                arrPageBreaks(1, intPageCount) = intNextPageRow
        '            End If

        '            'top border last data row + 1

        '            intFirstSearchRow = intNextPageRow
        '            intRow = intFirstSearchRow
        '            intRows = tbl.Rows.Count

        '            arrPageBreaks(2, intPageCount) = intRows

        '            tbl.Cell(intFirstSearchRow, 1).Select()
        '            p1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '            '    p1 = .selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '            If boolIsLastPage Then
        '                Exit Do
        '            End If

        '        Loop

        '        If boolFullPageNum Then
        '            Dim pos As Int64
        '            Dim rngP As Word.Range
        '            For Count1 = 1 To intPN

        '                Try

        '                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="PN" & Count1)
        '                    .Selection.TypeText(Text:=CStr(intPN))

        '                Catch ex As Exception

        '                End Try

        '            Next

        '            'delete bookmarks
        '            For Count1 = 1 To intPN

        '                Try
        '                    .ActiveDocument.Bookmarks.Item("PN" & Count1).Delete()
        '                Catch ex As Exception

        '                End Try
        '            Next
        '        End If

        '        If boolPSL Then
        '        Else

        '            If boolPageNum Or boolDTStamp Then
        '                .Selection.Tables.Item(1).Cell(intLastSearchRow + 1, 1).Select() 'intLastSearchRow is last row in table not including legend
        '                .Selection.InsertRowsAbove(ctLegend)
        '            Else
        '                intLastSearchRow = .Selection.Tables.Item(1).Rows.Count
        '                .Selection.Tables.Item(1).Cell(intLastSearchRow + 1, 1).Select() 'intLastSearchRow is last row in table not including legend
        '                .Selection.InsertRowsBelow(ctLegend)
        '                .Selection.Cells.Split(NumRows:=ctLegend, NumColumns:=1, MergeBeforeSplit:=True)
        '                If boolSmallFont Then
        '                    .Selection.Font.Size = .Selection.Font.Size - 1
        '                End If
        '                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                'format cells
        '                .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                With .Selection.ParagraphFormat
        '                    .LeftIndent = 36 'InchesToPoints(0.63)
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                End With
        '                With .Selection.ParagraphFormat
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                    .FirstLineIndent = -36 'InchesToPoints(-0.63)
        '                End With
        '            End If

        '            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '            ''remove underline from these selected rows
        '            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        '            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '            'border the top
        '            .Selection.Tables.Item(1).Cell(intLastSearchRow + 1, 1).Select()
        '            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle



        '            intRow = intLastSearchRow
        '            For Count1 = 1 To ctLegend

        '                intRow = intRow + 1

        '                .Selection.Tables.Item(1).Cell(intRow, 1).Select()

        '                'var1 = arrRealLegend(1, Count1)
        '                var1 = arr(1, Count1)
        '                'var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
        '                var1 = Replace(CStr(var1), "-", NBH, 1, -1, vbTextCompare)
        '                'If arrRealLegend(3, Count1) Then
        '                If arr(3, Count1) Then
        '                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                    .Selection.Font.Bold = False

        '                    typeInSuperscript(wd, CStr(var1))
        '                Else
        '                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                    .Selection.Font.Bold = False
        '                    .Selection.TypeText(Text:=CStr(var1))
        '                End If
        '                If Len(var1) = 0 Then
        '                Else
        '                    .Selection.TypeText(Text:=vbTab)
        '                    .Selection.TypeText(Text:="=")
        '                    .Selection.TypeText(Text:=vbTab)
        '                End If
        '                'var2 = arrRealLegend(2, Count1)
        '                var2 = arr(2, Count1)
        '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, vbTextCompare)
        '                '                var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
        '                .Selection.TypeText(Text:=CStr(var2))

        '            Next

        '        End If

        '        If gboolReadOnlyTables Then
        '            Call ReadOnlyTables(wd, tbl, ctHdRows, arrPageBreaks, intPageCount)
        '        End If

        '    End With

        '    str1 = strT & vbCr & "Repaginating document..."
        '    frmH.lblProgress.Text = str1
        '    frmH.lblProgress.Refresh()

        '    Try
        '        wd.ActiveWindow.View.Type = vView
        '    Catch ex As Exception
        '        MsgBox("Err vView: " & ex.Message)
        '    End Try

        '    wd.Options.BackgroundSave = boolBGS

        '    wd.Options.Pagination = boolPag


    End Sub


    Sub SplitTableOld2(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

        '        Dim boolCell
        '        Dim cell1
        '        Dim cell2
        '        Dim pg1
        '        Dim pg2
        '        Dim pgT
        '        Dim row1
        '        Dim row2
        '        Dim var1, var2
        '        Dim var3, var4
        '        Dim Count1 As Short
        '        Dim Count2 As Short
        '        Dim Count3 As Short
        '        Dim ctSplitRows As Short
        '        Dim boolTableEnd As Boolean
        '        Dim char1
        '        Dim char2
        '        Dim char3
        '        Dim myRReturn As Microsoft.Office.Interop.Word.Range
        '        Dim myr As Microsoft.Office.Interop.Word.Range
        '        Dim myR1 As Microsoft.Office.Interop.Word.Range
        '        Dim arrRealLegend(3, UBound(arr, 2))
        '        Dim int1 As Short
        '        Dim int2 As Short
        '        Dim int3 As Short
        '        Dim bool As Boolean
        '        Dim str1 As String
        '        Dim fonts
        '        Dim rows1 As Long
        '        Dim intCell1 As Short
        '        Dim intCell2 As Short
        '        Dim intRows As Long
        '        Dim intRow As Long
        '        Dim intRow1 As Long
        '        Dim intRow2 As Long
        '        Dim intRowStart As Long
        '        Dim intRowStart1 As Long
        '        Dim intRowStart2 As Long
        '        Dim intRowStart3 As Long

        '        Dim Count20 As Short
        '        Dim Count30 As Short

        '        Dim boolGo As Boolean
        '        Dim intIncr As Short

        '        Dim boolFound As Boolean
        '        Dim boolDoLast As Boolean

        '        intSplitRows = intSplitRows + 1 'take account for 'continue on next page'

        '        'If frmH.rbTable.Checked Then
        '        'Else
        '        '    Call SplitTableOld(wd, ctHdRows, ctLegend, arr, strT, DoLegend, DoTable, intSRow, intSplitRows, boolAutoFit, boolSmallFont)
        '        '    GoTo end2
        '        'End If

        '        'wdd.visible = True

        '        'select whole table
        '        wd.Selection.Tables.Item(1).Select()
        '        'set Normal font

        '        ''''wdd.visible = True

        '        Call SetNormalTable(wd)

        '        intRows = wd.Selection.Tables.Item(1).Rows.Count


        '        ''''''''wdd.visible = True

        '        If boolSmallFont Then
        '            '''''''''''wdd.visible = True
        '            wd.Selection.Font.Size = NormalFontsize - 1
        '        End If

        '        'first make headers one fontsize smaller
        '        wd.Selection.Tables.Item(1).Cell(2, 1).Select()
        '        wd.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=ctHdRows - 2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '        wd.Selection.SelectRow()
        '        wd.Selection.Font.Size = NormalFontsize - 1

        '        wd.Selection.Tables.Item(1).Select()
        '        If boolSmallFont Then
        '            wd.Selection.Font.Size = NormalFontsize - 1
        '        End If

        '        boolCell = True
        '        'On Error Resume Next
        '        cell1 = 1
        '        row1 = 1
        '        row2 = 1
        '        Count3 = 0 'iteration counter
        '        ctRealLegend = 0

        '        '''''''''''wdd.visible = True


        '        If wd.Selection.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
        '            intIncr = 15
        '        Else
        '            intIncr = 30
        '        End If

        '        boolSplitTable = False
        '        With wd

        '            If ctLegend = 0 Then
        '                GoTo end1
        '            End If

        '            boolTableEnd = False
        '            Count20 = 0
        '            Count30 = 0
        '            boolDoLast = False

        '            intRows = .Selection.Tables.Item(1).Rows.Count
        '            .Selection.Tables.Item(1).Cell(1, 1).Select()
        '            pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '            .Selection.Tables.Item(1).Cell(intRows, 1).Select()
        '            pg2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '            .Selection.Tables.Item(1).Cell(1, 1).Select()
        '            If pg1 = pg2 Then 'no need to evaluate
        '                '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToNext, Count:=1, Name:="")
        '                int1 = .Selection.Tables.Item(1).Rows.Count
        '                intRows = .Selection.Tables.Item(1).Rows.Count
        '                .Selection.Tables.Item(1).Cell(int1, 1).Select()
        '                boolDoLast = True
        '            End If

        '            'begin looking for next page
        '            intRow = 1
        '            intRowStart = 1
        '            Count20 = 0

        '            Do Until intRow >= intRows

        '                Count20 = Count20 + 1
        '                str1 = strT & vbCr & "Formatting legend..." & Count20
        '                frmH.lblProgress.Text = str1
        '                frmH.lblProgress.Refresh()

        '                intRows = .Selection.Tables.Item(1).Rows.Count

        '                ctRealLegend = 0

        '                If boolDoLast Then
        '                    intRow = intRows
        '                    intRow1 = intRow
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()
        '                Else
        '                    intRow = intRow + intIncr
        '                    If intRow > intRows Then
        '                        intRow = intRows
        '                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()
        '                        pgT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                        If pgT = pg2 Then
        '                            boolDoLast = True

        '                        Else
        '                        End If

        '                    End If
        '                    pgT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    .Selection.Tables.Item(1).Cell(intRows, 1).Select()
        '                    pg2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()
        '                    pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '                    ''''wdd.visible = True

        '                    If pg1 = pg2 Then
        '                        boolDoLast = True
        '                        intRowStart1 = .Selection.Tables.Item(1).Rows.Count
        '                    Else
        '                        If pgT <> pg1 Then 'intincr is too big
        '                            If pgT > pg1 Then
        '                            Else
        '                                Do Until pgT = pg1
        '                                    intRow = intRow - 5
        '                                    .Selection.Tables.Item(1).Cell(intRow, 1).Select()
        '                                    pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                                Loop

        '                            End If

        '                        End If

        '                    End If

        '                    pgT = pg1
        '                    intRow1 = intRow
        '                End If


        '                If boolDoLast Then
        '                    intRow1 = .Selection.Tables.Item(1).Rows.Count
        '                    intRowStart1 = intRow1

        '                    ''''''wdd.visible = True
        '                    '.Selection.Tables.Item(1).Cell(intRow1, 1).Select()
        '                Else
        '                    'wdd.visible = True
        '                    Do Until pg1 <> pgT
        '                        intRow1 = intRow1 + 1
        '                        .Selection.Tables.Item(1).Cell(intRow1, 1).Select()
        '                        pgT = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    Loop
        '                    'intRowStart = intRows + ctLegend

        '                    ''''''wdd.visible = True


        '                    intRow1 = intRow1 - 1
        '                    'intRow1 = intRow1 - ctLegend - 1
        '                    'intRow1 = intRow1 - intSplitRows - 1
        '                    intRow1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                    'intRowStart = intRow1

        '                    'If intRow2 = 0 Then
        '                    '    intRow2 = intRow1 - intSplitRows - 1
        '                    'Else
        '                    '    'intRow2 = intRow1 - intSplitRows - 1
        '                    'End If
        '                    intRow2 = intRow1 - intSplitRows - 1
        '                    'intRow1 = intRow2 - 1
        '                    .Selection.Tables.Item(1).Cell(intRow2, 1).Select()
        '                    'format paragraph force next page
        '                    .Selection.ParagraphFormat.PageBreakBefore = True
        '                    'intRow2 = intRow2 + ctLegend

        '                    intRow1 = intRow2 - 1
        '                    intRowStart1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                    intRow2 = intRow2 + intSplitRows 'need to set this for a later piece of code

        '                    '''wdd.visible = True

        '                End If

        '                '''''''wdd.visible = True

        '                'intRow1 = intRow1 - 1 'new for 'continued'

        '                .Selection.Tables.Item(1).Cell(intRow1, 1).Select()
        '                var1 = var1
        '                'insert rows for legend
        '                .Selection.InsertRowsBelow(ctLegend + 1) 'take account for 'continued next page'
        '                intRowStart1 = intRowStart1 + ctLegend
        '                'ensure bottom is not bordered
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
        '                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
        '                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
        '                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
        '                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
        '                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone

        '                .Selection.Cells.Split(NumRows:=ctLegend + 1, NumColumns:=1, MergeBeforeSplit:=True)

        '                'ensure bottom is not bordered
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                If boolSmallFont Then
        '                    .Selection.Font.Size = .Selection.Font.Size - 1
        '                End If
        '                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                'format cells
        '                .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                With .Selection.ParagraphFormat
        '                    .LeftIndent = 36 'InchesToPoints(0.63)
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                End With
        '                With .Selection.ParagraphFormat
        '                    .SpaceBefore = 0
        '                    .SpaceBeforeAuto = False
        '                    .SpaceAfter = 0
        '                    .SpaceAfterAuto = False
        '                    .FirstLineIndent = -36 'InchesToPoints(-0.63)
        '                End With

        '                'ensure top is  bordered
        '                .Selection.Tables.Item(1).Cell(intRow1 + 1, 1).Select()
        '                '.Selection.SelectRow()
        '                .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

        '                'With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
        '                '    .LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle
        '                'End With
        '                If DoLegend Then 'enter legend without checking
        '                    ctRealLegend = ctLegend
        '                    For Count1 = 1 To ctLegend
        '                        arrRealLegend(1, Count1) = arr(1, Count1)
        '                        arrRealLegend(2, Count1) = arr(2, Count1)
        '                        arrRealLegend(3, Count1) = arr(3, Count1)
        '                    Next
        '                Else
        '                    'determine if legend needs to be entered
        '                    'move back into previous table
        '                    '.Selection.Tables.Item(1).Cell(intRowStart1, 1).Select()
        '                    '.Selection.SelectRow()
        '                    ''int1 = intRow1 - intRowStart
        '                    'int1 = intRowStart1 - intRowStart
        '                    '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                    'myr = .Selection.Range

        '                    '.Selection.Tables.Item(1).Cell(intRowStart, 1).Select()
        '                    '.Selection.SelectRow()
        '                    ''int1 = intRow1 - intRowStart
        '                    'int1 = intRowStart1 - intRowStart
        '                    '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                    'myr = .Selection.Range

        '                    .Selection.Tables.Item(1).Cell(intRowStart + 1, 1).Select()
        '                    .Selection.SelectRow()
        '                    'int1 = intRow1 - intRowStart
        '                    int1 = intRowStart1 - intRowStart - intSplitRows ' - 1
        '                    .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=int1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                    myr = .Selection.Range

        '                    ''wdd.visible = True

        '                    '''''''''''wdd.visible = True


        '                    '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2)
        '                    ctRealLegend = 0
        '                    'first look for column header legend items
        '                    For Count1 = 1 To ctLegend
        '                        var1 = arr(4, Count1)
        '                        If NZ(var1, False) Then
        '                            ctRealLegend = ctRealLegend + 1
        '                            arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                            arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                            arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                        End If
        '                    Next

        '                    For Count1 = 1 To ctLegend

        '                        var1 = arr(4, Count1)
        '                        If NZ(var1, False) Then
        '                        Else
        '                            var1 = arr(1, Count1)
        '                            myr = .Selection.Range
        '                            boolFound = False
        '                            With myr.Find
        '                                .ClearFormatting()
        '                                .MatchCase = True
        '                                .MatchWholeWord = True
        '                                .Forward = True
        '                                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                                .Execute(FindText:=var1)

        '                                If .Found Then
        '                                    boolFound = True
        '                                Else
        '                                    boolFound = False
        '                                End If

        '                            End With

        '                            If boolFound Then 'legend needed
        '                                ctRealLegend = ctRealLegend + 1
        '                                arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                                arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                                arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                            End If
        '                        End If

        '                    Next
        '                End If

        '                '''''''''''wdd.visible = True

        '                intRow = intRow1 ' - 1
        '                If ctLegend = 0 Or ctRealLegend = 0 Then

        '                    If ctRealLegend < ctLegend Then 'delete some rows
        '                        int2 = intRow
        '                        'int1 = ctLegend - ctRealLegend
        '                        int1 = ctLegend - ctRealLegend - 1 'leave a row for 'continued next page'
        '                        intRow = intRow + 1
        '                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()
        '                        For Count1 = 1 To int1
        '                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()
        '                            int3 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                            wd.Selection.Tables.Item(1).Rows(int3).Delete()
        '                        Next
        '                        intRowStart1 = intRowStart1 - int1
        '                        intRow = int2

        '                    End If

        '                    If boolDoLast Then
        '                    Else
        '                        '.Selection.Tables.Item(1).Cell(intRow2, 1).Select()
        '                        '.Selection.TypeText("Continued on next page")

        '                        .Selection.Tables.Item(1).Cell(intRowStart1, 1).Select()
        '                        .Selection.TypeText("Continued on next page")

        '                        'format align right
        '                        .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight ' wdAlignParagraphRight
        '                    End If

        '                    '''wdd.visible = True

        '                Else

        '                    '''''''''''wdd.visible = True

        '                    'enter legend
        '                    If ctRealLegend < ctLegend Then 'delete some rows
        '                        int2 = intRow
        '                        'int1 = ctLegend - ctRealLegend
        '                        int1 = ctLegend - ctRealLegend ' - 1 'leave a row for 'continued next page'
        '                        intRow = intRow + 1
        '                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()
        '                        For Count1 = 1 To int1
        '                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()
        '                            int3 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                            wd.Selection.Tables.Item(1).Rows(int3).Delete()
        '                        Next
        '                        intRowStart1 = intRowStart1 - int1
        '                        intRow = int2
        '                    End If
        '                    'For Count1 = 1 To ctRealLegend
        '                    intRow = intRow ' - 1 'new for 'continued next page'
        '                    For Count1 = 1 To ctRealLegend + 1
        '                        intRow = intRow + 1

        '                        '''''wdd.visible = True

        '                        '.Selection.Tables.Item(1).Cell(intRow - 1, 1).Select()
        '                        'var1 = var1
        '                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()
        '                        'var1 = var1

        '                        If Count1 = ctRealLegend + 1 Then
        '                            ''wdd.visible = True
        '                            If boolDoLast Then
        '                            Else
        '                                .Selection.TypeText(Text:="Continued on next page")
        '                                'format align right
        '                                .Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight ' wdAlignParagraphRight
        '                            End If

        '                        Else
        '                            var1 = arrRealLegend(1, Count1)
        '                            var1 = Replace(CStr(var1), "-", NBH, 1, -1, CompareMethod.Text)
        '                            If arrRealLegend(3, Count1) Then
        '                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                                .Selection.Font.Bold = False

        '                                fonts = .Selection.Font.Size
        '                                .Selection.Font.Superscript = True
        '                                .Selection.Font.Size = 12
        '                                .Selection.TypeText(Text:=CStr(var1))
        '                                .Selection.Font.Superscript = False
        '                                .Selection.Font.Size = fonts
        '                            Else
        '                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                                .Selection.Font.Bold = False

        '                                .Selection.TypeText(Text:=CStr(var1))
        '                            End If
        '                            If Len(var1) = 0 Then
        '                            Else
        '                                .Selection.TypeText(Text:=vbTab)
        '                                .Selection.TypeText(Text:="=")
        '                                .Selection.TypeText(Text:=vbTab)
        '                            End If
        '                            var2 = arrRealLegend(2, Count1)
        '                            var2 = Replace(CStr(var2), "-", NBH, 1, -1, CompareMethod.Text)
        '                            'var3 = var1 & " = " & var2
        '                            .Selection.TypeText(Text:=CStr(var2))
        '                            char2 = .Selection.Start
        '                            If Count1 = ctRealLegend Then
        '                                myR1 = wd.ActiveDocument.Range(Start:=char2, End:=char2)
        '                            End If

        '                        End If

        '                    Next
        '                    var1 = .Selection.Tables.Item(1).Rows.Count
        '                    If intRow2 > var1 Then
        '                        intRow2 = var1
        '                    End If
        '                    '.Selection.Tables.Item(1).Cell(intRow2, 1).Select()
        '                    .Selection.Tables.Item(1).Cell(intRowStart1 + 1, 1).Select()
        '                    'intRow1 = intRow2
        '                    'intRow = intRow1
        '                    'intRowStart = intRow
        '                    'intRows = intRows + ctLegend

        '                End If

        '                'intRow1 = intRow2
        '                'intRow = intRow1
        '                'intRowStart = intRow

        '                intRow1 = intRowStart1 ' + ctRealLegend
        '                intRow = intRow1
        '                intRowStart = intRow
        '                intRow2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber) + 1

        '                If boolDoLast Then
        '                    Exit Do
        '                End If
        '                var1 = intRow 'for testing

        '                Try
        '                    With myr.Find
        '                        .ClearFormatting()
        '                    End With
        '                Catch ex As Exception

        '                End Try

        '            Loop

        '            frmH.lblProgress.Text = strT


        'end1:


        '        End With

        'end2:

        '        'clear formatting again
        '        wd.Selection.Find.ClearFormatting()

    End Sub


    Sub SplitTableOld(ByVal wd, ByVal ctHdRows, ByVal ctLegend, ByVal arr, ByVal strT, ByVal DoLegend, ByVal DoTable, ByVal intSRow, ByVal intSplitRows, ByVal boolAutoFit, ByVal boolSmallFont)
        '        Dim boolCell
        '        Dim cell1
        '        Dim cell2
        '        Dim pg1
        '        Dim pg2
        '        Dim row1
        '        Dim row2
        '        Dim var1, var2
        '        Dim var3, var4
        '        Dim Count1 As Short
        '        Dim Count2 As Short
        '        Dim Count3 As Short
        '        Dim ctSplitRows As Short
        '        Dim boolTableEnd As Boolean
        '        Dim char1
        '        Dim char2
        '        Dim char3
        '        Dim myRReturn As Microsoft.Office.Interop.Word.Range
        '        Dim myr As Microsoft.Office.Interop.Word.Range
        '        Dim myR1 As Microsoft.Office.Interop.Word.Range
        '        Dim arrRealLegend(3, UBound(arr, 2))
        '        Dim int1 As Short
        '        Dim int2 As Short
        '        Dim bool As Boolean
        '        Dim str1 As String
        '        Dim fonts
        '        Dim rows1 As Long
        '        Dim intCell1 As Short
        '        Dim intCell2 As Short
        '        Dim intRows As Short

        '        Dim Count20 As Short
        '        Dim Count30 As Short


        '        Call MoveOneCellDown(wd)

        '        wd.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        '        '''''''''wdd.visible = True

        '        'enter footer
        '        If wd.ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
        '            wd.ActiveWindow.Panes(2).Close()
        '        End If
        '        If wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
        '            wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        '        End If
        '        wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter


        '        'make footer not linked to previous
        '        wd.Selection.HeaderFooter.LinkToPrevious = False 'Not wd.Selection.HeaderFooter.LinkToPrevious()

        '        'seek main document
        '        wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
        '        wd.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        '        'enter footer
        '        If wd.ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
        '            wd.ActiveWindow.Panes(2).Close()
        '        End If
        '        If wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or wd.ActiveWindow. _
        '            ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
        '            wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        '        End If
        '        wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter

        '        Try
        '            wd.ActiveWindow.ActivePane.View.NextHeaderFooter()
        '        Catch ex As Exception

        '        End Try


        '        'delete contents of footer
        '        wd.Selection.WholeStory()
        '        wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

        '        If boolSmallFont Then
        '            wd.Selection.Font.Size = 10
        '        End If

        '        'add legend to footer
        '        ctRealLegend = ctLegend
        '        For Count1 = 1 To ctLegend
        '            arrRealLegend(1, Count1) = arr(1, Count1)
        '            arrRealLegend(2, Count1) = arr(2, Count1)
        '            arrRealLegend(3, Count1) = arr(3, Count1)
        '        Next
        '        With wd
        '            If ctLegend = 0 Or ctRealLegend = 0 Then
        '            Else

        '                With .Selection.ParagraphFormat
        '                    .LeftIndent = 39.6 '54 'InchesToPoints(0.5)
        '                End With
        '                With .Selection.ParagraphFormat
        '                    .FirstLineIndent = -39.6 '-54 'InchesToPoints(-0.5)
        '                End With
        '                .Selection.ParagraphFormat.TabStops.Add(Position:=25.2, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

        '                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

        '                'enter legend
        '                For Count1 = 1 To ctRealLegend
        '                    var1 = arrRealLegend(1, Count1)
        '                    If arrRealLegend(3, Count1) Then
        '                        'fonts = .Selection.Font.Size
        '                        .Selection.Font.Superscript = True
        '                        .Selection.Font.Size = 12
        '                        .Selection.TypeText(Text:=CStr(var1))
        '                        .Selection.Font.Superscript = False
        '                        .Selection.Font.Size = fonts
        '                    Else
        '                        .Selection.TypeText(Text:=CStr(var1))
        '                    End If
        '                    .Selection.TypeText(Text:=vbTab)
        '                    .Selection.TypeText(Text:="=")
        '                    .Selection.TypeText(Text:=vbTab)
        '                    var2 = arrRealLegend(2, Count1)
        '                    'var3 = var1 & " = " & var2
        '                    .Selection.TypeText(Text:=CStr(var2))
        '                    char2 = .Selection.start
        '                    If Count1 = ctRealLegend Then
        '                    Else
        '                        .Selection.TypeParagraph()
        '                    End If
        '                Next
        '            End If
        '        End With
        '        'seek main document
        '        wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
        '        wd.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        '        Exit Sub

        '        boolCell = True
        '        'On Error Resume Next
        '        cell1 = 1
        '        row1 = 1
        '        row2 = 1
        '        Count3 = 0 'iteration counter
        '        ctRealLegend = 0

        '        '''''''''wdd.visible = True

        '        boolSplitTable = False
        '        With wd
        '            boolTableEnd = False
        '            Count20 = 0
        '            Count30 = 0
        '            Do Until Err.Number <> 0
        '                Count30 = Count30 + 1
        '                If Count30 > 200 Then
        '                    str1 = "There is a problem splitting this table."
        '                    str1 = str1 & vbCr & vbCr & "Please screen shot the entire window and contact your StudyDoc Administrator."
        '                    MsgBox(str1, MsgBoxStyle.Critical, "Split table problem...")
        '                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
        '                    Exit Sub
        '                End If

        '                'goto first ctHdRows rows and paste
        '                .Selection.Tables.item(1).Cell(1, 1).Select()
        '                .Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, ctHdRows, Word.WdMovementType.wdExtend)
        '                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                .Selection.Copy()

        '                .Selection.Tables.item(1).Cell(1, 1).Select()
        '                intRows = .Selection.Tables.item(1).Rows.Count
        '                pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToNext, Count:=1, Name:="")
        '                pg2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                If pg1 = pg2 Then 'no need to evaluate
        '                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToNext, Count:=1, Name:="")
        '                    int1 = .Selection.Tables.item(1).Rows.Count
        '                    intRows = .Selection.Tables.item(1).Rows.Count
        '                    .Selection.Tables.item(1).Cell(int1, 1).Select()

        '                    GoTo skip1
        '                End If

        '                bool = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable)
        '                'MsgBox(bool & ":  " & str1)
        '                If .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable) Then
        '                    'int1 = CInt(.Selection.Information(Microsoft.Office.Interop.Word.wdinformation.wdStartOfRangeRowNumber))
        '                    intCell1 = CInt(.Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber))
        '                    int1 = intCell1 - 4
        '                    .Selection.Tables.item(1).Cell(int1, 1).Select()
        '                Else
        '                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToPrevious, Count:=1, Name:="")
        '                    intRows = .Selection.Tables.item(1).Rows.Count
        '                    If intRows < 4 Then
        '                    Else
        '                        .Selection.Tables.item(1).Cell(intRows - 4, 1).Select()
        '                    End If

        '                End If
        '                pg2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                rows1 = .Selection.Tables.Item(1).Rows.Count
        '                Do Until pg1 <> pg2

        '                    Count30 = Count30 + 1
        '                    If Count30 > 200 Then
        '                        str1 = "There is a problem splitting this table."
        '                        str1 = str1 & vbCr & vbCr & "Please screen shot the entire window and contact your StudyDoc Administrator."
        '                        MsgBox(str1, MsgBoxStyle.Critical, "Split table problem...")
        '                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
        '                        Exit Sub
        '                    End If

        '                    pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)
        '                    row1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                    If Err.Number > 0 Then
        '                        Exit Do
        '                    End If
        '                    cell1 = .Selection.Tables.item(1).Rows.item(row1).Cells.Count
        '                    If Err.Number > 0 Then
        '                        Exit Do
        '                    End If

        '                    '''''''''wdd.visible = True

        '                    If row1 = rows1 Then
        '                        Exit Do
        '                    Else
        '                        .selection.tables.item(1).cell(row1 + 1, 1).select()
        '                        '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
        '                    End If

        '                    pg2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber)

        '                    row2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                    If Err.Number > 0 Then
        '                        Exit Do
        '                    End If
        '                    cell2 = .Selection.Tables.item(1).Rows.item(row1).Cells.Count
        '                    If Err.Number > 0 Then
        '                        Exit Do
        '                    End If
        '                    'cell2 = .Selection.Rows.item(1).Cells.Count
        '                    If pg1 <> pg2 Then 'must split table

        '                        boolSplitTable = True

        '                        'move back until cell 1 contains data
        '                        Count3 = Count3 + 1

        '                        frmH.lblProgress.Text = strT & "(Splitting Table: " & Count3 & ")"
        '                        frmH.Refresh()
        '                        str1 = "aa"
        '                        ctSplitRows = 0
        '                        If .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable) Then
        '                        Else
        '                            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToPrevious, Count:=1, Name:="")
        '                            intRows = .Selection.Tables.item(1).Rows.Count
        '                            If intRows < 4 Then
        '                            Else
        '                                .Selection.Tables.item(1).Cell(intRows - 4, 1).Select()
        '                            End If
        '                        End If
        '                        'Do Until .Selection.Information(Microsoft.Office.Interop.Word.wdinformation.wdWithInTable)
        '                        '    .Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                        'Loop

        '                        var3 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                        ''Do Until Len(str1) = 0 And ctSplitRows > intSplitRows 'ctLegend + 1
        '                        'Do Until ctSplitRows > intSplitRows 'ctLegend + 1

        '                        '    ctSplitRows = ctSplitRows + 1

        '                        '    '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                        '    var3 = .Selection.Information(Microsoft.Office.Interop.Word.wdinformation.wdStartOfRangeRowNumber)
        '                        '    'str1 = .Selection.Tables.item(1).Rows.item(var3).Cells(1)
        '                        '    'str1 = .selection.Tables.item(1).cell(var3, 1).range.text

        '                        '    'str1 = .selection.Tables.item(1).cell(var3 - ctSplitRows, 1).range.text
        '                        '    ''remove any carriage returns
        '                        '    'var1 = Mid(str1, Len(str1), 1)
        '                        '    'int1 = Asc(var1)
        '                        '    'int1 = Len(str1)
        '                        '    'For Count1 = int1 To 1 Step -1
        '                        '    '    var1 = Mid(str1, Count1, 1)
        '                        '    '    int2 = Asc(var1)
        '                        '    '    If int2 < 29 Then 'eliminate
        '                        '    '        str1 = Microsoft.VisualBasic.Left(str1, Len(str1) - 1)
        '                        '    '    End If
        '                        '    'Next
        '                        'Loop
        '                        'move to last evaluated cell
        '                        '.selection.Tables.item(1).cell(var3 - ctSplitRows, 1).select()
        '                        .selection.Tables.item(1).cell(var3 - intSplitRows, 1).select()

        '                        'autofit table to window so that pasting table header won't get screwed up
        '                        If boolAutoFit Then
        '                            autofitWindow(wd, 2)
        '                        End If
        '                        .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)


        '                        If boolAutoFit Then
        '                            'move to lower table and apply autofitwindow
        '                            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToNext, Count:=1, Name:="")

        '                            '.selection.Tables.item(1).select()
        '                            autofitWindow(wd, 2)
        '                            'move back up to previous table
        '                            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToPrevious, Count:=1, Name:="")
        '                            'intRows = .Selection.Tables.item(1).Rows.Count
        '                            '.Selection.Tables.item(1).Cell(intRows, 1).Select()
        '                            '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=3)
        '                        Else
        '                            'move back up to previous table
        '                            '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2)
        '                        End If
        '                        'move back up to previous table
        '                        .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToPrevious, Count:=1, Name:="")
        '                        intRows = .Selection.Tables.item(1).Rows.Count
        '                        .Selection.Tables.item(1).Cell(intRows, 1).Select()

        '                        '''''''''wdd.visible = True

        'here:

        '                        '.Selection.SplitTable()
        '                        'bottom border on previous line
        '                        '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2)
        '                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                        .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

        '                        '''''''''wdd.visible = True

        '                        'move out of table
        '                        Call MoveOneCellDown(wd)

        '                        '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                        '.Selection.Font.Size = 11
        '                        myr = .Selection.Range
        '                        '.Selection.TypeParagraph()
        '                        wd.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView
        '                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                        wd.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView

        '                        myRReturn = .Selection.Range

        '                        '''''''''wdd.visible = True

        '                        If DoLegend Then 'enter legend without checking
        '                            ctRealLegend = ctLegend
        '                            For Count1 = 1 To ctLegend
        '                                arrRealLegend(1, Count1) = arr(1, Count1)
        '                                arrRealLegend(2, Count1) = arr(2, Count1)
        '                                arrRealLegend(3, Count1) = arr(3, Count1)
        '                            Next
        '                        Else
        '                            'determine if legend needs to be entered
        '                            'move back into previous table
        '                            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToPrevious, Count:=1, Name:="")
        '                            intRows = .Selection.Tables.item(1).Rows.Count
        '                            .Selection.Tables.item(1).Cell(intRows, 1).Select()

        '                            '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2)
        '                            ctRealLegend = 0

        '                            For Count1 = 1 To ctLegend
        '                                If DoTable Then
        '                                    .Selection.Tables.item(1).Select()
        '                                Else
        '                                    var1 = .Selection.Tables.item(1).Rows.Count
        '                                    .Selection.Tables.item(1).Cell(ctHdRows + 1, intSRow).Select()
        '                                    .Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, var1 - ctHdRows - 1, Word.WdMovementType.wdExtend)
        '                                End If
        '                                char1 = .Selection.start
        '                                var1 = arr(1, Count1)
        '                                .Selection.Find.ClearFormatting()
        '                                With .Selection.Find
        '                                    .Text = CStr(var1)
        '                                    .Replacement.Text = ""
        '                                    .Forward = True
        '                                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                                    .Format = False
        '                                    .MatchCase = True
        '                                    .MatchWholeWord = True
        '                                    .MatchWildcards = False
        '                                    .MatchSoundsLike = False
        '                                    .MatchAllWordForms = False
        '                                    .Execute()
        '                                End With
        '                                Err.Clear()
        '                                char2 = .Selection.start

        '                                If char1 = char2 Then 'legend not needed
        '                                Else
        '                                    ctRealLegend = ctRealLegend + 1
        '                                    arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                                    arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                                    arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                                End If
        '                            Next
        '                        End If
        '                        'enter legend
        '                        myr.Select()
        '                        If ctLegend = 0 Or ctRealLegend = 0 Then
        '                        Else

        '                            'int1 = .Selection.Information(Microsoft.Office.Interop.Word.wdinformation.wdStartOfRangeRowNumber)
        '                            '.Selection.Tables.item(1).Cell(int1 - 1, 1).Select()
        '                            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToPrevious, Count:=1, Name:="")
        '                            intRows = .Selection.Tables.item(1).Rows.Count
        '                            .Selection.Tables.item(1).Cell(intRows, 1).Select()

        '                            '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
        '                            .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                            .Selection.InsertRowsBelow(1)
        '                            'ensure font is not red
        '                            .selection.font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                            .Selection.Font.Bold = False
        '                            '.Color = Microsoft.Office.Interop.Word.wdcolor.wdColorBlack

        '                            Try
        '                                .Selection.Cells.Merge()
        '                            Catch ex As Exception

        '                            End Try
        '                            'format cells
        '                            .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                            .Selection.ParagraphFormat.TabStops.Add(Position:=36, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)


        '                            '''''''''wdd.visible = True

        '                            With .Selection.ParagraphFormat
        '                                .LeftIndent = 36 'InchesToPoints(0.5)
        '                            End With
        '                            With .Selection.ParagraphFormat
        '                                .FirstLineIndent = -36 'InchesToPoints(-0.5)
        '                            End With
        '                            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                            .Selection.InsertRowsBelow(ctRealLegend)
        '                            int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                            .Selection.Tables.item(1).Cell(int1 - 1, 1).Select()
        '                            '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                            .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        '                            .Selection.Tables.item(1).Cell(int1, 1).Select()
        '                            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)


        '                            'enter legend
        '                            For Count1 = 1 To ctRealLegend
        '                                var1 = arrRealLegend(1, Count1)
        '                                If arrRealLegend(3, Count1) Then
        '                                    'fonts = .Selection.Font.Size
        '                                    .Selection.Font.Superscript = True
        '                                    .Selection.Font.Size = 12
        '                                    .Selection.TypeText(Text:=CStr(var1))
        '                                    .Selection.Font.Superscript = False
        '                                    .Selection.Font.Size = fonts
        '                                Else
        '                                    .Selection.TypeText(Text:=CStr(var1))
        '                                End If
        '                                .Selection.TypeText(Text:=vbTab)
        '                                .Selection.TypeText(Text:="=")
        '                                .Selection.TypeText(Text:=vbTab)
        '                                var2 = arrRealLegend(2, Count1)
        '                                'var3 = var1 & " = " & var2
        '                                .Selection.TypeText(Text:=CStr(var2))
        '                                char2 = .Selection.start
        '                                If Count1 = ctRealLegend Then
        '                                    '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                                    '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)
        '                                    Call MoveOneCellDown(wd)
        '                                    myR1 = wd.ActiveDocument.Range(start:=char2, End:=char2)
        '                                Else
        '                                    .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
        '                                End If
        '                            Next
        '                        End If

        '                        '''''''''wdd.visible = True

        '                        myRReturn.Select()
        '                        '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
        '                        '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                        '.Selection.Paste()
        '                        .Selection.PasteSpecial(Link:=False, DataType:=Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteRTF) ', Placement:=wdInLine)
        '                        .Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)

        '                        'copy and paste_text the table number
        '                        '.Selection.MoveUp Unit:=Microsoft.Office.Interop.Word.WdUnits.wdline, Count:=5
        '                        .Selection.Tables.item(1).Cell(1, 1).Select()
        '                        .Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1)
        '                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdWord, Count:=1)
        '                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdWord, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                        .Selection.Copy()
        '                        '.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
        '                        '.Selection.TypeText(Text:=" ")
        '                        '.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdFormatPlainText)
        '                        .selection.pastespecial(link:=False, datatype:=Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteText)
        '                        '.Selection.PasteSpecial(Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False)


        '                        '''''''''wdd.visible = True

        '                        If Count3 > 1 Then
        '                        Else
        '                            'modify title with 'continued'
        '                            .Selection.Tables.item(1).Cell(1, 1).Select()
        '                            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
        '                            .Selection.MoveEnd(unit:=Microsoft.Office.Interop.Word.WdUnits.wdParagraph)
        '                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
        '                            .Selection.TypeText(Text:=" (continued)")
        '                        End If

        '                        'Exit Sub

        '                    End If
        '                Loop
        '            Loop


        '            frmH.lblProgress.Text = strT
        '            If Err.Number <> 0 Then
        '                Err.Clear()
        '                'On Error GoTo 0 'end error handling
        '            End If

        '            'border bottom row

        '            .Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '            .Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdLine)

        'skip1:

        '            pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber) 'set page number
        '            char1 = .Selection.start
        '            myr = wd.ActiveDocument.Range(start:=char1, End:=char1)
        '            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '            .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        '            int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '            .Selection.Tables.item(1).Cell(int1 + 1, 1).Select()
        '            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)

        '            If ctLegend = 0 Then
        '                int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                .Selection.Tables.item(1).Cell(int1 - 1, 1).Select()
        '                '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
        '                .Selection.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdRow, Word.WdMovementType.wdExtend)
        '                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
        '                .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
        '                'Set myR1 = wd.ActiveDocument.Range(Start:=char2, End:=char2)
        '            Else
        '                myRReturn = .Selection.Range
        '                If DoLegend Then 'enter legend without checking
        '                    ctRealLegend = ctLegend
        '                    For Count1 = 1 To ctLegend
        '                        arrRealLegend(1, Count1) = arr(1, Count1)
        '                        arrRealLegend(2, Count1) = arr(2, Count1)
        '                        arrRealLegend(3, Count1) = arr(3, Count1)
        '                    Next
        '                Else
        '                    ctRealLegend = 0
        '                    int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                    '.Selection.Tables.item(1).Cell(int1 - 1, 1).Select()
        '                    '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)

        '                    For Count1 = 1 To ctLegend
        '                        If DoTable Then
        '                            .Selection.Tables.item(1).Select()
        '                        Else
        '                            var1 = .Selection.Tables.item(1).Rows.Count
        '                            .Selection.Tables.item(1).Cell(ctHdRows + 1, intSRow).Select()
        '                            .Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, var1 - ctHdRows - 1, Word.WdMovementType.wdExtend)
        '                        End If
        '                        char1 = .Selection.start
        '                        var1 = arr(1, Count1)
        '                        .Selection.Find.ClearFormatting()
        '                        With .Selection.Find
        '                            .Text = CStr(var1)
        '                            .Replacement.Text = ""
        '                            .Forward = True
        '                            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '                            .Format = False
        '                            .MatchCase = True
        '                            .MatchWholeWord = False
        '                            .MatchWildcards = False
        '                            .MatchSoundsLike = False
        '                            .MatchAllWordForms = False
        '                            .Execute()
        '                        End With
        '                        Err.Clear()
        '                        char2 = .Selection.start

        '                        If char1 = char2 Then 'legend not needed
        '                        Else
        '                            ctRealLegend = ctRealLegend + 1
        '                            arrRealLegend(1, ctRealLegend) = arr(1, Count1)
        '                            arrRealLegend(2, ctRealLegend) = arr(2, Count1)
        '                            arrRealLegend(3, ctRealLegend) = arr(3, Count1)
        '                        End If
        '                    Next
        '                End If
        '                'enter final legend
        '                myRReturn.Select()
        '                '.Selection.Font.Size = 11
        '                If ctRealLegend = 0 Then
        '                    int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                    .Selection.Tables.item(1).Cell(int1 - 1, 1).Select()
        '                    '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
        '                    .Selection.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdRow, Word.WdMovementType.wdExtend)
        '                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
        '                    .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
        '                    'Set myR1 = wd.ActiveDocument.Range(Start:=char2, End:=char2)
        '                Else
        '                    '.Selection.TypeParagraph
        '                    'int1 = .Selection.Information(Microsoft.Office.Interop.Word.wdinformation.wdStartOfRangeRowNumber)
        '                    '.Selection.Tables.item(1).Cell(int1 - 1, 1).Select()
        '                    '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
        '                    .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone

        '                    .Selection.InsertRowsBelow(1)
        '                    'ensure font is not red
        '                    .selection.font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                    '.Color = Microsoft.Office.Interop.Word.wdcolor.wdColorBlack
        '                    .Selection.Font.Bold = False

        '                    Try
        '                        .Selection.Cells.Merge()
        '                    Catch ex As Exception

        '                    End Try
        '                    'ensure font is not red
        '                    .selection.font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                    '.Color = Microsoft.Office.Interop.Word.wdcolor.wdColorBlack
        '                    .Selection.Font.Bold = False

        '                    'format cells
        '                    .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                    .Selection.ParagraphFormat.TabStops.Add(Position:=36, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                    With .Selection.ParagraphFormat
        '                        .LeftIndent = 36 'InchesToPoints(0.5)
        '                    End With
        '                    With .Selection.ParagraphFormat
        '                        .FirstLineIndent = -36 'InchesToPoints(-0.5)
        '                    End With
        '                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                    .Selection.InsertRowsBelow(ctRealLegend)
        '                    int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                    .Selection.Tables.item(1).Cell(int1 - 1, 1).Select()
        '                    '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                    .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        '                    .Selection.Tables.item(1).Cell(int1, 1).Select()
        '                    '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)


        '                    '.ActiveDocument.Tables.Add(myRReturn, ctRealLegend + 1, 1)
        '                    '.Selection.Tables.item(1).Rows.AllowBreakAcrossPages = False
        '                    '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                    '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=ctRealLegend, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                    ''format line
        '                    '.Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                    '.Selection.ParagraphFormat.TabStops.Add(Position:=36, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                    'With .Selection.ParagraphFormat
        '                    '    .LeftIndent = 36 'InchesToPoints(0.5)
        '                    'End With

        '                    'With .Selection.ParagraphFormat
        '                    '    .FirstLineIndent = -36 'InchesToPoints(-0.5)
        '                    'End With
        '                    '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
        '                    '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)

        '                    'enter legend
        '                    For Count1 = 1 To ctRealLegend
        '                        var1 = arrRealLegend(1, Count1)
        '                        If arrRealLegend(3, Count1) Then
        '                            'fonts = .Selection.Font.Size
        '                            .Selection.Font.Superscript = True
        '                            .Selection.Font.Size = 12
        '                            .Selection.TypeText(Text:=CStr(var1))
        '                            .Selection.Font.Superscript = False
        '                            .Selection.Font.Size = fonts
        '                        Else
        '                            .Selection.TypeText(Text:=CStr(var1))
        '                        End If
        '                        .Selection.TypeText(Text:=vbTab)
        '                        .Selection.TypeText(Text:="=")
        '                        .Selection.TypeText(Text:=vbTab)
        '                        var2 = arrRealLegend(2, Count1)
        '                        'var3 = var1 & " = " & var2
        '                        .Selection.TypeText(Text:=CStr(var2))
        '                        char2 = .Selection.start
        '                        If Count1 = ctRealLegend Then
        '                            '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)
        '                            Call MoveOneCellDown(wd)
        '                            myR1 = wd.ActiveDocument.Range(start:=char2, End:=char2)
        '                        Else
        '                            .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
        '                            '.Selection.TypeParagraph
        '                        End If
        '                    Next

        '                    'ensure legend is not split up
        '                    myr.Select()
        '                    char1 = .Selection.start
        '                    Count1 = 0
        '                    For Count1 = 1 To ctRealLegend + 1
        '                        pg1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber) 'set page number
        '                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
        '                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
        '                        pg2 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber) 'set page number
        '                        char1 = .Selection.start
        '                        If pg1 <> pg2 Then
        '                            myr.Select()
        '                            int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                            .Selection.Tables.item(1).Cell(int1 - (ctRealLegend + 1), 1).Select()
        '                            '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, ctRealLegend + 1)

        '                            '.Selection.SplitTable
        '                            .Selection.SplitTable()

        '                            'bottom border on previous line
        '                            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToPrevious, Count:=1, Name:="")
        '                            intRows = .Selection.tables.item(1).rows.count
        '                            .Selection.Tables.item(1).Cell(intRows, 1).Select()
        '                            '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                            .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

        '                            Call MoveOneCellDown(wd)
        '                            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                            '.Selection.Font.Size = 11
        '                            myr = .Selection.Range
        '                            .Selection.TypeParagraph()
        '                            myRReturn = .Selection.Range

        '                            '***Come back here
        '                            'enter legend
        '                            If ctLegend = 0 Or ctRealLegend = 0 Then
        '                            Else

        '                                'move back to previous table
        '                                .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToPrevious, Count:=1, Name:="")
        '                                intRows = .Selection.tables.item(1).rows.count
        '                                .Selection.Tables.item(1).Cell(intRows, 1).Select()
        '                                'myr.Select()
        '                                '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
        '                                .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
        '                                .Selection.InsertRowsBelow(1)
        '                                Try
        '                                    .Selection.Cells.Merge()
        '                                Catch ex As Exception

        '                                End Try
        '                                'format cells
        '                                .Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                                .Selection.ParagraphFormat.TabStops.Add(Position:=36, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                                With .Selection.ParagraphFormat
        '                                    .LeftIndent = 36 'InchesToPoints(0.5)
        '                                End With
        '                                With .Selection.ParagraphFormat
        '                                    .FirstLineIndent = -36 'InchesToPoints(-0.5)
        '                                End With
        '                                'ensure font is not red
        '                                .selection.font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                                '.Color = Microsoft.Office.Interop.Word.wdcolor.wdColorBlack
        '                                .Selection.Font.Bold = False

        '                                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                                .Selection.InsertRowsBelow(ctRealLegend)
        '                                int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
        '                                .Selection.Tables.item(1).Cell(int1 - 1, 1).Select()
        '                                '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
        '                                .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        '                                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)


        '                                '.ActiveDocument.Tables.Add(myr, ctRealLegend + 1, 1)
        '                                '.Selection.Tables.item(1).Rows.AllowBreakAcrossPages = False
        '                                '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                                '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=ctRealLegend, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                                ''format line
        '                                '.Selection.ParagraphFormat.TabStops.Add(Position:=25, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                                '.Selection.ParagraphFormat.TabStops.Add(Position:=36, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '                                'With .Selection.ParagraphFormat
        '                                '    .LeftIndent = 36 'InchesToPoints(0.5)
        '                                'End With
        '                                'With .Selection.ParagraphFormat
        '                                '    .FirstLineIndent = -36 'InchesToPoints(-0.5)
        '                                'End With
        '                                '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
        '                                '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)


        '                                For Count2 = 1 To ctRealLegend
        '                                    var1 = arrRealLegend(1, Count2)
        '                                    If arrRealLegend(3, Count2) Then
        '                                        'fonts = .Selection.Font.Size
        '                                        .Selection.Font.Superscript = True
        '                                        .Selection.Font.Size = 12
        '                                        .Selection.TypeText(Text:=CStr(var1))
        '                                        .Selection.Font.Superscript = False
        '                                        .Selection.Font.Size = fonts
        '                                    Else
        '                                        .Selection.TypeText(Text:=CStr(var1))
        '                                    End If
        '                                    .Selection.TypeText(Text:=vbTab)
        '                                    .Selection.TypeText(Text:="=")
        '                                    .Selection.TypeText(Text:=vbTab)
        '                                    var2 = arrRealLegend(2, Count2)
        '                                    'var3 = var1 & " = " & var2
        '                                    .Selection.TypeText(Text:=CStr(var2))
        '                                    If Count2 = ctRealLegend Then
        '                                        Call MoveOneCellDown(wd)
        '                                        '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                                        '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)
        '                                        'Set myR1 = wd.ActiveDocument.Range(Start:=.Selection.Start, End:=.Selection.Start)
        '                                    Else
        '                                        .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
        '                                        '.Selection.TypeParagraph
        '                                    End If
        '                                Next
        '                            End If
        '                            .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
        '                            .Selection.Paste()
        '                            .Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)

        '                            '.Selection.MoveUp Unit:=Microsoft.Office.Interop.Word.WdUnits.wdline, Count:=5
        '                            .Selection.Tables.item(1).Cell(1, 1).Select()
        '                            .Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1)
        '                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdWord, Count:=1)
        '                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdWord, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
        '                            .Selection.Copy()
        '                            '.Selection.PasteAndFormat(wdFormatPlainText)
        '                            .Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteText)
        '                            'modify title with 'continued'

        '                            .Selection.Tables.item(1).Cell(1, 1).Select()
        '                            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
        '                            .Selection.MoveEnd(unit:=Microsoft.Office.Interop.Word.WdUnits.wdParagraph)
        '                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
        '                            .Selection.TypeText(Text:=" (continued)")

        '                            'go to end of table
        '                            myR1.Select()
        '                            '.Selection.MoveUp Word.WdUnits.wdline, 1
        '                            '.Selection.EndKey Unit:=Microsoft.Office.Interop.Word.WdUnits.wdline, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
        '                            '.Selection.MoveRight wdCharacter, 1
        '                            Exit For
        '                            'Exit Do
        '                        End If
        '                    Next
        '                End If
        '                '.Selection.EndKey Unit:=Microsoft.Office.Interop.Word.WdUnits.wdline
        '            End If

        '            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)

        'end1:

        '        End With

        '        'clear formatting again
        '        wd.selection.Find.ClearFormatting()

        '    End Sub

        '    'Sub CreateGuWuTOF(ByVal wd As Word.Application)

        '    Try

        '        Try
        '            'first attempt to delete existing GuWuTOF
        '            wd.ActiveDocument.Styles("GuWuTOF").Delete()
        '        Catch ex As Exception

        '        End Try

        '        With wd

        '            .ActiveDocument.Styles.Add(Name:="GuWuTOF", Type:=Microsoft.Office.Interop.Word.WdStyleType.wdStyleTypeParagraph)
        '            .ActiveDocument.Styles("GuWuTOF").AutomaticallyUpdate = False
        '            With .ActiveDocument.Styles("GuWuTOF").Font
        '                '.Name = "Times New Roman"
        '                '.Size = 10
        '                .Bold = False
        '                .Italic = False
        '                '.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
        '                '.UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                .StrikeThrough = False
        '                .DoubleStrikeThrough = False
        '                .Outline = False
        '                .Emboss = False
        '                .Shadow = False
        '                .Hidden = False
        '                .SmallCaps = True
        '                .AllCaps = False
        '                If boolBLUEHYPERLINK Then
        '                    .Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlue
        '                Else
        '                    .Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        '                End If
        '                .Engrave = False
        '                .Superscript = False
        '                .Subscript = False
        '                .Scaling = 100
        '                .Kerning = 0
        '                '.Animation = wdAnimationNone
        '            End With
        '            With .ActiveDocument.Styles("GuWuTOF").ParagraphFormat
        '                .LeftIndent = 72 ' InchesToPoints(1)
        '                .RightIndent = 18 ' InchesToPoints(0.26)
        '                .SpaceBefore = 0
        '                .SpaceBeforeAuto = False
        '                .SpaceAfter = 0
        '                .SpaceAfterAuto = False
        '                .LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle
        '                .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '                .WidowControl = True
        '                .KeepWithNext = False
        '                .KeepTogether = False
        '                .PageBreakBefore = False
        '                .NoLineNumber = False
        '                .Hyphenation = True
        '                .FirstLineIndent = -72 ' InchesToPoints(-1)
        '                Try
        '                    .OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText

        '                Catch ex As Exception

        '                End Try
        '                .CharacterUnitLeftIndent = 0
        '                .CharacterUnitRightIndent = 0
        '                .CharacterUnitFirstLineIndent = 0
        '                .LineUnitBefore = 0
        '                .LineUnitAfter = 0
        '            End With
        '            .ActiveDocument.Styles("GuWuTOF").NoSpaceBetweenParagraphsOfSameStyle = False
        '            .ActiveDocument.Styles("GuWuTOF").ParagraphFormat.TabStops.ClearAll()
        '            .ActiveDocument.Styles("GuWuTOF").ParagraphFormat.TabStops.Add(Position:=72, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '            .ActiveDocument.Styles("GuWuTOF").ParagraphFormat.TabStops.Add(Position:=504, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots)
        '            With .ActiveDocument.Styles("GuWuTOF").ParagraphFormat
        '                With .Shading
        '                    '.Texture = wdTextureNone
        '                    '.ForegroundPatternColor = wdColorAutomatic
        '                    '.BackgroundPatternColor = wdColorAutomatic
        '                End With
        '                '.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        '                '.Borders(wdBorderRight).LineStyle = wdLineStyleNone
        '                '.Borders(wdBorderTop).LineStyle = wdLineStyleNone
        '                '.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        '                'With .Borders
        '                '    .DistanceFromTop = 1
        '                '    .DistanceFromLeft = 4
        '                '    .DistanceFromBottom = 1
        '                '    .DistanceFromRight = 4
        '                '    .Shadow = False
        '                'End With
        '            End With
        '            '.ActiveDocument.Styles("GuWuTOF").LanguageID = wdEnglishUS
        '            .ActiveDocument.Styles("GuWuTOF").NoProofing = False
        '            .ActiveDocument.Styles("GuWuTOF").Frame.Delete()

        '        End With
        '    Catch ex As Exception

        '    End Try

    End Sub

    Sub ModifyTOF(ByVal wd As Word.Application)

        'now modify Table of Figure built-in style

        With wd.ActiveDocument.Styles("Table of Figures")
            .AutomaticallyUpdate = True
            '.BaseStyle = "Normal"
            '.NextParagraphStyle = "Normal"
        End With

        With wd.ActiveDocument.Styles("Table of Figures").Font
            '.Size = 10
            If boolBLUEHYPERLINK Then
                .ColorIndex = BlueHyperlinkColor ' Microsoft.Office.Interop.Word.WdColor.wdColorBlue
            Else
                '.Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
                .ColorIndex = Word.WdColorIndex.wdAuto
            End If

        End With



        'With wd.ActiveDocument.Styles("Table of Figures").ParagraphFormat

        '    .TabStops.ClearAll()

        '    .TabStops.Add(Position:=72, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
        '    .TabStops.Add(Position:=504, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots)

        '    .LeftIndent = 72 ' InchesToPoints(1)
        '    .RightIndent = 18 ' InchesToPoints(0.26)
        '    .SpaceBefore = 0
        '    .SpaceBeforeAuto = False
        '    .SpaceAfter = 0
        '    .SpaceAfterAuto = False
        '    .LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle
        '    .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
        '    .WidowControl = True
        '    .KeepWithNext = False
        '    .KeepTogether = False
        '    .PageBreakBefore = False
        '    .NoLineNumber = False
        '    .Hyphenation = True
        '    .FirstLineIndent = -72 ' InchesToPoints(-1)
        '    Try
        '        .OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText

        '    Catch ex As Exception

        '    End Try
        '    .CharacterUnitLeftIndent = 0
        '    .CharacterUnitRightIndent = 0
        '    .CharacterUnitFirstLineIndent = 0
        '    .LineUnitBefore = 0
        '    .LineUnitAfter = 0

        'End With

    End Sub

    'Sub ApplyGuWuTOF(ByVal wd As Word.Application)


    '    'wdd.visible = True

    '    'wd.Selection.Style = wd.ActiveDocument.Styles("GuWuTOF")

    '    wd.Selection.Tables.Item(1).Cell(2, 1).Select()
    '    'wd.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

    '    wd.Selection.Style = wd.ActiveDocument.Styles("GuWuTOF")

    'End Sub

    Sub ApplyWordTOF(ByVal wd As Word.Application)


        'wdd.visible = True

        'wd.Selection.Style = wd.ActiveDocument.Styles("GuWuTOF")

        wd.Selection.Tables.Item(1).Cell(2, 1).Select()
        'wd.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

        wd.Selection.Style = wd.ActiveDocument.Styles("Table of Figures")

    End Sub


    Function FindTOF(ByVal wd As Word.Application, ByVal strTOF As String) As Boolean

        FindTOF = False

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
        Dim myrng As Word.Selection
        myrng = wd.Selection

        Dim int1 As Short
        Dim ws As Word.Style

        int1 = 0
        Dim strS As String
        With myrng.Find
            .Wrap = Word.WdFindWrap.wdFindContinue
            '.ClearFormatting()
            Do While .Execute(FindText:=strTOF, Forward:=True)
                ws = wd.Selection.Style
                strS = ws.NameLocal.ToString
                If StrComp(strS, "Heading 1", CompareMethod.Text) = 0 Then
                    FindTOF = True
                    Exit Do
                End If
                int1 = int1 + 1
                If int1 > 2 Then
                    Exit Do
                End If

            Loop
        End With


        'With myrng.Find
        '    .ClearFormatting()
        '    '.Text = strFind
        '    .Forward = True
        '    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '    .Execute(FindText:=strTOF)

        '    If .Found Then
        '        FindTOF = True
        '    Else
        '        FindTOF = False
        '    End If

        'End With

    End Function


    Sub FormatIndex(ByVal wd As Word.Application, ByVal li As Single, ByVal ri As Single)

        '20180828 LEE:
        'Discontinue this

        Exit Sub

        'NOTE: This MUST be run on same page as tables of figures, etc. The margins may be different than that of title page

        Dim pw, lm, rm, rt

        With wd

            '.ActiveDocument.Save() 'try this to get rid of annoying hangs for large TOC/F/A
            'find right tab
            With .Selection.PageSetup
                lm = .LeftMargin ' = InchesToPoints(1.4)
                rm = .RightMargin ' = InchesToPoints(1)
                pw = .PageWidth ' = InchesToPoints(8.5)
                rt = pw - lm - rm
            End With

            ''wdd.visible = True

            '.selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            '.Selection.Tables.Item(1).Cell(2, 1).Select()
            With .ActiveDocument.Styles("Table of Figures").ParagraphFormat
                'xxxx
                .LeftIndent = li '54 'InchesToPoints(0.76)
                .RightIndent = ri '19 'InchesToPoints(0.27)
                'xxxx

                .SpaceBefore = 0
                '.SpaceBeforeAuto = False
                .SpaceAfter = 0
                '.SpaceAfterAuto = False
                .LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle

                .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                '.NoLineNumber = False
                Try
                    .NoLineNumber = False
                Catch ex As Exception

                End Try
                '.Hyphenation = True
                Try
                    .Hyphenation = True
                Catch ex As Exception

                End Try
                '.FirstLineIndent = -li '-54 'InchesToPoints(-0.76)
                '.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText

                'xxxx
                Try
                    .FirstLineIndent = -li '-54 'InchesToPoints(-0.76)
                Catch ex As Exception

                End Try
                'xxxx

                Try
                    .OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText ' Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText
                Catch ex As Exception

                End Try

                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                Try
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                Catch ex As Exception

                End Try
            End With

            wd.ActiveDocument.Styles("Table of Figures").ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots)

            Dim tb As Word.TabStop
            Dim var1

            Dim intTabs As Short
            Dim Count1 As Short
            intTabs = wd.ActiveDocument.Styles("Table of Figures").ParagraphFormat.TabStops.Count
            For Count1 = intTabs To 1 Step -1

                tb = wd.ActiveDocument.Styles("Table of Figures").ParagraphFormat.TabStops(Count1)

                If tb.Position = li Or tb.Position = rt Then
                    var1 = 1
                ElseIf tb.Alignment = Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabCenter Then
                    Try
                        tb.Clear()
                    Catch ex As Exception
                        var1 = ex.Message
                        var1 = var1
                    End Try
                Else

                    If tb.Alignment = Word.WdTabAlignment.wdAlignTabRight Then

                        Try
                            tb.Clear()
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try

                    End If

                End If

            Next

        End With

    End Sub

    Sub InsertFigs(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal RowNum As Short, ByVal rows() As DataRow)

        'Dim fso As New Scripting.FileSystemObject
        'Dim fi As Scripting.File
        'Dim fo
        Dim strPath As String
        Dim Count1 As Integer
        Dim Count2 As Integer
        Dim ctArr As Integer
        Dim var1, var2, var3
        Dim str1 As String
        Dim tm, bm, lm, rm, fm, hm, hd, fd
        Dim numPH, numPW
        Dim scale, ct, cb, cl, cr, g
        Dim ht, wid, sht, swid
        Dim arr1(100) As String
        Dim int1 As Short
        Dim pw, ph, rt
        Dim origV, origM
        Dim orient
        Dim strP As String
        Dim picHt, picWd
        Dim boolIWD As Boolean = False
        Dim boolApp As Boolean = False

        int1 = rows(RowNum).Item("BOOLAPPENDIX")
        If int1 = 0 Then
            boolApp = False
        Else
            boolApp = True
        End If

        'this should be first
        Dim varVersion
        varVersion = wd.Version
        'MsgBox(varVersion)

        'tm = Sheets("ReportConfig").Range("TopMargin").Offset(0, 1).Value
        'bm = Sheets("ReportConfig").Range("BottomMargin").Offset(0, 1).Value
        'rm = Sheets("ReportConfig").Range("RightMargin").Offset(0, 1).Value
        'lm = Sheets("ReportConfig").Range("LeftMargin").Offset(0, 1).Value
        'fm = Sheets("ReportConfig").Range("FooterMargin").Offset(0, 1).Value
        'hm = Sheets("ReportConfig").Range("HeaderMargin").Offset(0, 1).Value

        ''wdd.visible = True

        With wd

            scale = rows(RowNum).Item("numScale")
            ct = NZ(rows(RowNum).Item("numCropTop"), 0) * 72
            cb = NZ(rows(RowNum).Item("numCropBottom"), 0) * 72
            cl = NZ(rows(RowNum).Item("numCropLeft"), 0) * 72
            cr = NZ(rows(RowNum).Item("numCropRight"), 0) * 72
            var1 = NZ(rows(RowNum).Item("BOOLINSERTWORDDOCS"), 0)
            If var1 = 0 Then
                boolIWD = True
            Else
                boolIWD = True
            End If

            If var1 = 0 Then
                boolIWD = False
            Else
                boolIWD = True
            End If
            'find document left, right, top, and bottom margin
            'tm = .ActiveDocument.Sections(1).PageSetup.TopMargin
            'bm = .ActiveDocument.Sections(1).PageSetup.bottomMargin
            'lm = .ActiveDocument.Sections(1).PageSetup.leftMargin
            'rm = .ActiveDocument.Sections(1).PageSetup.rightMargin

            With .ActiveDocument.Sections.Item(1).PageSetup
                lm = .LeftMargin ' * 72 ' = InchesToPoints(1.4)
                rm = .RightMargin ' * 72 ' = InchesToPoints(1)
                bm = .BottomMargin ' * 72 ' = InchesToPoints(1)
                tm = .TopMargin ' * 72 ' = InchesToPoints(1)
                pw = .PageWidth ' * 72 ' = InchesToPoints(8.5)
                ph = .PageHeight ' * 72 ' = InchesToPoints(8.5)
                hd = .HeaderDistance
                fd = .FooterDistance

                rt = pw - lm - rm
                orient = .Orientation

                picHt = ph - bm - tm
                picWd = pw - lm - rm

            End With

            'calculate picture height and width
            'assumes Letter size paper
            numPW = pw - lm - rm
            numPH = ph - tm - bm

            'strPath = Sheets("GlobalConfig").Range("DefaultChromatogramPath").Value
            strPath = NZ(rows(RowNum).Item("charPath"), "")

            If Len(strPath) = 0 Then
                GoTo end1
            End If

            If boolIWD Then
            Else
                'ensure strPath ends in "\"
                If StrComp(Right(strPath, 1), "\", vbTextCompare) = 0 Then
                Else
                    strPath = strPath & "\"
                End If
            End If

            Dim boolEx As Boolean = False
            If boolIWD Then
                'ensure file exists
                boolEx = File.Exists(strPath)
            Else
                'ensure directory exists
                boolEx = Directory.Exists(strPath)
            End If


            If boolEx Then

                If boolIWD Then
                    ctArr = 1
                    arr1(ctArr) = strPath

                    origV = frmH.pb1.Value
                    origM = frmH.pb1.Maximum
                    frmH.pb1.Value = 0
                    frmH.pb1.Maximum = ctArr + 1

                    'if passing boolsingle, then need to pass a query of tblAppFigs
                    If boolApp Then
                        Call InsertGraphics("Appendix", wd, True, RowNum, boolIWD)
                    Else
                        '***** NDL - Need to find which figure it is in the list of figures
                        'When producing figures, it pre-filters the rows in InsertGraphicInd() to only the row with the FCID.
                        'But in InsertGraphics, it re-creates the list of figures, and expects the function call to specify
                        'the correct number on *that* list.  It probably needs a more comprehensive fix, but this should fix
                        'it for Word figure insertions.

                        Dim strF As String
                        Dim Count3, intFigureRowNum As Integer
                        Dim rowFigures() As DataRow

                        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLFIGURE <> 0 AND BOOLAPPENDIX = 0 AND BOOLINCLUDEINREPORT <> 0"
                        rowFigures = tblAppFigs.Select(strF)
                        For Count3 = 0 To rowFigures.Length - 1
                            If (rowFigures(Count3).Item("CHARFCID") = rows(RowNum).Item("CHARFCID")) Then
                                intFigureRowNum = Count3
                            End If
                        Next
                        '***** End NDL

                        InsertGraphics("Figure", wd, True, intFigureRowNum, boolIWD)  'NDL: was RowNum before.
                    End If

                Else
                    'fo = fso.GetFolder(strPath)
                    Dim strExt As String
                    strExt = "*.jpg,*.jpeg,*.bmp,*.tif,*.tiff,*.png"
                    Dim strE As String


                    'For Each foundFile As String In My.Computer.FileSystem.GetFiles(strPath, FileIO.SearchOption.SearchAllSubDirectories, "*.dll")

                    '    Listbox1.Items.Add(foundFile)
                    'Next
                    ReDim arr1(100)
                    int1 = 0
                    Count1 = 0
                    ctArr = 0
                    For Each foundFile As String In Directory.GetFiles(strPath, "*.*", SearchOption.TopDirectoryOnly)

                        strExt = GetExt(foundFile)
                        If IsFig(strExt) Then
                            Count1 = Count1 + 1
                            arr1(Count1) = foundFile ' strPath & fi.Name
                        End If

                    Next
                    ctArr = Count1

                    'Dim fo As String() = Directory.GetFiles(strPath)
                    'Count1 = 0
                    'ctArr = 0
                    'int1 = fo.files.count
                    'ReDim arr1(int1)
                    'For Each fi In fo.Files

                    '    strExt = GetExt(fi.Name)
                    '    If IsFig(strExt) Then
                    '        Count1 = Count1 + 1
                    '        arr1(Count1) = strPath & fi.Name
                    '    End If

                    'Next
                    'ctArr = Count1

                    If ctArr = 0 Then
                        str1 = "Hmmm. There doesn't seem to be any figures in" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & strPath & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "Report generation will continue."
                        .Selection.TypeText(Text:=str1)
                        GoTo end1
                    End If

                    'sort arr1 ascending
                    For Count1 = 1 To ctArr - 1
                        var1 = arr1(Count1)
                        For Count2 = Count1 To ctArr
                            var2 = arr1(Count2)
                            If var2 < var1 Then
                                var3 = var1
                                arr1(Count1) = var2
                                arr1(Count2) = var3
                            End If
                        Next
                    Next

                    'insert figures
                    'wd.Selection.Goto What:=wdGoToBookmark, Name:="Chromatograms"

                    'wd.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                    'wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)

                    origV = frmH.pb1.Value
                    origM = frmH.pb1.Maximum
                    frmH.pb1.Value = 0
                    frmH.pb1.Maximum = ctArr + 1

                    '''wdd.visible = True

                    'ctArr = 0

                    For Count1 = 1 To ctArr

                        strP = "Inserting appendices/figures: " & Count1 & " of " & ctArr
                        frmH.lblProgress.Text = strP
                        frmH.pb1.Value = Count1
                        frmH.pb1.Refresh()
                        frmH.lblProgress.Refresh()

                        'Application.StatusBar = "Inserting chromatograms: " & arr1(Count1) & "..."
                        'wd.Selection.InlineShapes.AddPicture Filename:= _
                        'arr1(Count1), LinkToFile:=False, SaveWithDocument:=True
                        'wd.Selection.InlineShapes.AddPicture(arr1(Count1), False, True)

                        var1 = arr1(Count1) 'debug
                        Try
                            wd.Selection.InlineShapes.AddPicture(arr1(Count1), False, True)
                        Catch ex As Exception
                            GoTo next1
                        End Try

                        intILS = intILS + 1

                        With wd.ActiveDocument.InlineShapes.Item(intILS)

                            ht = .Height
                            wid = .Width

                            sht = .ScaleHeight
                            swid = .ScaleWidth

                            If sht = 0 Then
                                sht = 100
                            End If

                            If swid = 0 Then
                                swid = 100
                            End If

                            '20151027 Larry
                            'Do not size anymore
                            'If orient = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait Then
                            '    .Width = numPW
                            'Else : orient = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape
                            '    .Height = numPH
                            'End If

                            '.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                            '.Line.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue
                            .Line.Visible = True ' Office.Core.MsoTriState.msoCTrue ' Microsoft.Office.Core.MsoTriState.msoTrue 'Microsoft.Office.Core.MsoTriState.msoTrue

                            '.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                            .LockAspectRatio = True ' Microsoft.Office.Core.MsoTriState.msoTrue

                            With .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                '.LineWidth = wdLineWidth050pt
                                '.Color = wdColorAutomatic
                            End With
                            With .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                '.LineWidth = wdLineWidth050pt
                                '.Color = wdColorAutomatic
                            End With
                            With .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                '.LineWidth = wdLineWidth050pt
                                '.Color = wdColorAutomatic
                            End With
                            With .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                '.LineWidth = wdLineWidth050pt
                                '.Color = wdColorAutomatic
                            End With
                            .Borders.Shadow = False

                            If ht > picHt Then
                                .Height = picHt * 0.95
                            End If
                            wid = .Width
                            If wid > picWd Then
                                .Width = picWd
                            End If

                            If scale = 100 Then
                            Else
                                .ScaleHeight = sht * (scale / 100)
                                .ScaleWidth = swid * (scale / 100)
                            End If


                            'sht = .scaleheight
                            'swid = .scalewidth

                            '.PictureFormat.ColorType = Microsoft.Office.Core.MsoPictureColorType.msoPictureAutomatic
                            .PictureFormat.CropLeft = cl '0.0#
                            .PictureFormat.CropRight = cr '0.0#
                            .PictureFormat.CropTop = ct '0.0#
                            .PictureFormat.CropBottom = cb '0.0#
                        End With

                        If Count1 = ctArr Then
                            'wd.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                            wd.Selection.TypeParagraph()
                        Else
                            'wd.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                            wd.Selection.TypeParagraph()
                        End If
next1:

                    Next

                End If



            End If


end1:
        End With
    End Sub


    Sub JustTable(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strAnal As String, ByVal strTitle As String, ByVal strDo As String, ByVal strTName As String, ByVal intTableID As Short, ByVal strStability As String, ByVal strMsg As String, strTNameO As String, intGroup As Short, idTR As Int64)

        '''''''''wdd.visible = True

        Dim str1 As String
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim oFontSize
        Dim oColor
        Dim oBold

        '''''''''wdd.visible = True

        With wd
            'create a table with one row

            '.Selection.InsertBreak

            'wrdSelection = wd.Selection()

            '
            wrdSelection = wd.Selection()

            '20180913 LEE:
            Call IncrNextTableNumber(wd)

            If boolPlaceHolder Then
                .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
            Else
                .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
            End If

            .Selection.Tables.Item(1).Select()

            Call SetCellPaddingZero(.Selection.Tables.Item(1))

            Call GlobalTableParaFormat(wd)


            'remove border, but leave top and bottom
            removeBorderButLeaveTopAndBottom(wd)

            If boolPlaceHolder Then
                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
            Else

                'border top and bottom of range
                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                .Selection.Tables.Item(1).Cell(1, 1).Select()
                .Selection.TypeParagraph()
                oColor = .Selection.Font.Color
                oFontSize = .Selection.Font.Size
                oBold = .Selection.Font.Bold
                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                .Selection.Font.Size = 14
                '.Selection.TypeText("Samples have not been assigned")
                .Selection.TypeText(NZ(strMsg, "Samples have not been assigned"))
                .Selection.TypeParagraph()
                Try
                    .Selection.Font.Color = oColor
                    .Selection.Font.Size = oFontSize
                    .Selection.Font.Bold = oBold
                Catch ex As Exception

                End Try
            End If



            'enter table number
            strTName = strTitle
            Call EnterTableNumber(wd, strTName, 0, strAnal, strStability, intTableID, intGroup, idTR)

            Dim charFCID As String
            Dim strF As String
            Dim var1
            strF = "ID_TBLREPORTTABLE = " & idTR
            Dim rowsTR() As DataRow = tblReportTable.Select(strF)
            var1 = rowsTR(0).Item("CHARFCID")
            charFCID = NZ(var1, "NA")


            'enter a table record in tblTableN
            'ctTableN = ctTableN + 1
            Dim dtblr As DataRow = tblTableN.NewRow
            dtblr.BeginEdit()
            dtblr.Item("TableNumber") = ctTableN
            dtblr.Item("AnalyteName") = strDo 'arrAnalytes(1, Count1)
            dtblr.Item("TableName") = strTNameO
            dtblr.Item("TableID") = intTableID
            dtblr.Item("CHARFCID") = charFCID
            dtblr.Item("TableNameNew") = strTName
            tblTableN.Rows.Add(dtblr)


            Call MoveOneCellDown(wd)

            If boolPlaceHolder Then
                .Selection.TypeParagraph()
                .Selection.TypeParagraph()
            End If

        End With

    End Sub

    Sub DeleteTableRows(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim intC As Short
        Dim intT As Short
        Dim Count1 As Short
        Dim intRows As Short
        Dim intD As Short
        Dim tblNum As Short
        Dim wdTbl As Word.Table

        ''''''wdd.visible = True

        With wd


            wdTbl = .Selection.Tables.Item(1)

            intC = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)

            intT = .Selection.Tables.Item(1).Rows.Count

            'If intC = intT Then
            'Else
            '    'delete excess rows
            '    .Selection.Tables.Item(1).Cell(intC + 1, 1).Select()
            '    intD = intT - intC
            '    For Count1 = 1 To intD
            '        .Selection.Rows.Delete()
            '    Next
            'End If

            If intC = intT Then
            Else
                'delete excess rows
                .Selection.Tables.Item(1).Cell(intC + 1, 1).Select()
                intD = intT - intC
                For Count1 = 1 To intD
                    .Selection.Rows.Delete()
                Next
            End If

            'return to selection
            wdTbl.Select()

            intT = .Selection.Tables.Item(1).Rows.Count
            .Selection.Tables.Item(1).Cell(intT, 1).Select()


        End With

    End Sub

    Sub RemoveRows(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal ctTbl As Short)

        'ctTbl is depricated

        Dim boolV As Boolean = wd.Visible

        Dim intRows As Short
        Dim intCols As Short
        Dim boolDo As Boolean
        Dim var1, var2
        Dim Count1 As Int32
        Dim Count2 As Short
        Dim intI As Short

        Dim intR1 As Short
        Dim intR2 As Short

        Dim tbl As Microsoft.Office.Interop.Word.Table

        tbl = wd.Selection.Tables.Item(1)

        '20180419 LEE: Optimized some code

        intRows = tbl.Rows.Count ' wd.Selection.Tables.Item(1).Rows.Count
        intCols = tbl.Columns.Count 'wd.Selection.Tables.Item(1).Columns.Count

        'wd.Selection.Tables.Item(1).Cell(intRows, 1).Select()
        'wd.ActiveDocument.Tables.Item(ctTbl).Cell(intRows, 1).Select()


        ''''wdd.visible = True

        Dim rng1 As Microsoft.Office.Interop.Word.Range

        boolDo = False
        intI = 0
        intR1 = intRows

        For Count1 = intRows To 1 Step -1

            Try
                rng1 = tbl.Rows(Count1).Range ' Selection.Range

                'look for data in rng1
                intCols = rng1.Cells.Count

            Catch ex As Exception
                '"Cannot access individual rows in this collection because the table has vertically merged cells."
                intCols = 5
            End Try

            Try

                Try
                    For Count2 = 1 To intCols

                        var1 = tbl.Cell(Count1, Count2).Range.Text

                        'var1 = rng1.Cells(Count2).Range.Text
                        var2 = AscW(var1)

                        If var2 = 13 Or Len(var1) = 0 Then
                        Else
                            boolDo = True
                            intR2 = Count1 + 1
                            Exit For
                        End If

                    Next
                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1
                End Try

                If boolDo Then
                    Exit For
                Else
                    intI = intI + 1
                End If
            Catch ex As Exception

                var1 = ex.Message
                var1 = var1

            End Try


        Next Count1


        '20180514 LEE
        'LI-00017 has identified a problem that seems to by related to Office 2013 vs 2010. Development upgraded to Office 2013 Mar 27, 2018. 
        'If the study is multi-analyte and the 1st table has a 1-row legend:
        'The RemoveRows function will remove all rows of the 2nd table if the word document visible remains false
        '   Resulting in an error:  The requested member of the collection does not exist
        'To resolve, the RemoveRows function needs to make word document visible just before rows.delete, then put back to original visible state (may be True)
        'Don't know why, but seems to happen only if being called from Ad Hoc Stability _31, though the RemoveRows function is called in the same manner as all other tables


        wd.Visible = True

        'now delete

        If intI = 0 Then
            wd.Selection.SelectRow()
        Else
            Dim myCells As Microsoft.Office.Interop.Word.Range
            With wd
                myCells = .ActiveDocument.Range(Start:=.Selection.Tables(1).Cell(intR2, 1).Range.Start, End:=.Selection.Tables(1).Cell(intR1, 1).Range.End)

                '20180514 LEE: Added to ensure correct selection
                myCells.Select()

                myCells.Rows.Delete()

                'myCells.Select()
                '.Selection.Rows.Delete()
                'delete will take the cursor out of the table
                'put it back in
                tbl.Cell(intR2 - 1, 1).Select()
            End With

            wd.Selection.SelectRow()

            If intI = 0 Then
            Else
                wd.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
            End If

        End If

        wd.Visible = boolV


        'bottom border selection
        ''wdd.visible = True

        wd.Selection.SelectRow()

        If intI = 0 Then
        Else
            wd.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        End If

    End Sub

    Sub RemoveRowsSpecial(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal ctTbl As Short)

        'ctTbl is depricated

        Dim intRows As Short
        Dim intCols As Short
        Dim boolDo As Boolean
        Dim var1, var2
        Dim Count1 As Int32
        Dim Count2 As Short
        Dim intI As Short

        Dim intR1 As Short
        Dim intR2 As Short

        Dim tbl As Microsoft.Office.Interop.Word.Table

        tbl = wd.Selection.Tables.Item(1)

        '20180419 LEE: Optimized some code

        intRows = tbl.Rows.Count ' wd.Selection.Tables.Item(1).Rows.Count
        intCols = tbl.Columns.Count 'wd.Selection.Tables.Item(1).Columns.Count

        'wd.Selection.Tables.Item(1).Cell(intRows, 1).Select()
        'wd.ActiveDocument.Tables.Item(ctTbl).Cell(intRows, 1).Select()


        ''''wdd.visible = True

        Dim rng1 As Microsoft.Office.Interop.Word.Range

        boolDo = False
        intI = 0
        intR1 = intRows


        For Count1 = intRows To 1 Step -1

            rng1 = tbl.Rows(Count1).Range ' Selection.Range

            'look for data in rng1
            intCols = rng1.Cells.Count

            For Count2 = 1 To intCols

                var1 = rng1.Cells(Count2).Range.Text
                var2 = AscW(var1)

                If var2 = 13 Or Len(var1) = 0 Then
                Else
                    boolDo = True
                    intR2 = Count1 + 1
                    Exit For
                End If

            Next

            If boolDo Then
                Exit For
            Else
                intI = intI + 1
            End If

        Next Count1


        'now delete

        If intI = 0 Then
            wd.Selection.SelectRow()
        Else
            Dim myCells As Microsoft.Office.Interop.Word.Range
            With wd
                myCells = .ActiveDocument.Range(Start:=.Selection.Tables(1).Cell(intR2, 1).Range.Start, End:=.Selection.Tables(1).Cell(intR1, 1).Range.End)

                myCells.Select()

                'wd.Visible = True
                var1 = var1

                myCells.Rows.Delete()

                'myCells.Select()
                '.Selection.Rows.Delete()
                'delete will take the cursor out of the table
                'put it back in
                tbl.Cell(intR2 - 1, 1).Select()
            End With

            wd.Selection.SelectRow()

            If intI = 0 Then
            Else
                wd.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
            End If

        End If


        'bottom border selection
        ''wdd.visible = True

        wd.Selection.SelectRow()

        If intI = 0 Then
        Else
            wd.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        End If

    End Sub


End Module
