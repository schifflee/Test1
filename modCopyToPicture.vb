Option Compare Text

Imports Word = Microsoft.Office.Interop.Word
Imports System.Windows
Imports System.IO

Module modCopyToPicture


    Sub ReadOnlyTables(ByVal wd As Word.Application, ByVal tbl As Word.Table, ByVal ctHdRows As Short, ByVal arrPageBreaks As Object, ByVal intPageCount As Short)

        Dim intTotPages As Short
        Dim p1 As Short
        Dim p2 As Short
        Dim var1, var2, var3, var4
        Dim CountA As Short
        Dim sRows As Short
        Dim myr1 As Word.Selection
        Dim intRows As Int32



        With wd



            Dim strEP As String = "C:\Labintegrity\StudyDoc\Temp\"
            strEP = strEP & "gt.png"
            If File.Exists(strEP) Then
                File.Delete(strEP)
            End If

            'go to the beginning of the table and add intPageBreaks number of rows and merge then

            tbl.Cell(ctHdRows + 2, 1).Select()
            .Selection.InsertRowsAbove(intPageCount)
            'myr1 = .Selection.Range
            intRows = tbl.Rows.Count
            .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone

            'must deactive header row


            'merge the rows
            .Selection.Cells.Split(NumRows:=intPageCount, NumColumns:=1, MergeBeforeSplit:=True)
            .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
            .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone

            With .Selection.Cells(1)
                '.TopPadding = 0 ' InchesToPoints(0)
                '.BottomPadding = 0 ' InchesToPoints(0)
                '.LeftPadding = 0 '1 'InchesToPoints(0)
                '.RightPadding = 0 'InchesToPoints(0)
                .WordWrap = True
                .FitText = False
            End With

            Call SetCellPadding(.Selection.Tables.Item(1))

            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft ' wdAlignParagraphLeft

            var1 = ""

            'Must use Excel because copy/paste within Word returns screwy graphics
            '20150611: Look in to this later



            'wb.SaveAs(strEP)

            'wb.SaveAs("C:\Labintegrity\StudyDoc\Temp\Graphics.xlsx")

            Dim strM1 As String
            Dim strM2 As String
            Dim strM3 As String
            Dim vint1
            Dim vint2
            strM1 = frmH.lblProgress.Text
            For CountA = 1 To intPageCount

                strM2 = "Inserting table as figure..." & CountA & " of " & intPageCount & "..."
                strM3 = strM1 & ChrW(10) & strM2
                frmH.lblProgress.Text = strM3
                frmH.lblProgress.Refresh()

                'var2 = arrPageBreaks(1, CountA) + intPageCount 
                vint1 = arrPageBreaks(1, CountA) 'debug
                var2 = arrPageBreaks(1, CountA) + intPageCount + 1
                vint2 = arrPageBreaks(2, CountA) 'debug
                var3 = arrPageBreaks(2, CountA) + intPageCount
                sRows = var3 - var2

                'select a section
                tbl.Cell(var2, 1).Select()
                .Selection.SelectRow()
                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=sRows, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

                'xlROT.Visible = True
                'wd.Visible = True



                'copy the selection and paste it into the appropriate cell as a graphic
                '.Selection.Copy()

                'wd.Visible = True

                'System.Threading.Thread.Sleep(200)

                '.Selection.CopyAsPicture()


                '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend.wdExtend)

                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                '.Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone


                'With wd.Options
                '    .PasteAdjustWordSpacing = True
                '    .PasteAdjustParagraphSpacing = False 'keep as false
                '    .PasteAdjustTableFormatting = False
                '    .PasteSmartStyleBehavior = False
                '    .PasteMergeFromPPT = False
                '    .PasteMergeFromXL = False 'keep as false
                '    .PasteMergeLists = False
                'End With

                'wd.Visible = True
                'My.Computer.Clipboard.Clear()
                Try
                    Clipboard.Clear()
                Catch ex As Exception

                End Try
                'give time to set
                Pause(0.1)


                Dim rng1 As Microsoft.Office.Interop.Word.Range

                rng1 = .Selection.Range

                'rng1.CopyAsPicture()
                rng1.CopyAsPicture()

                'With wd.ActiveDocument.Bookmarks
                '    .Add(Range:=wd.Selection.Range, Name:="BK1")
                '    '.DefaultSorting = wdSortByName
                '    .ShowHidden = False
                'End With


                '.Selection.CopyAsPicture()

                System.Threading.Thread.Sleep(500)

                If My.Computer.Clipboard.ContainsAudio Then
                    var1 = My.Computer.Clipboard.GetAudioStream
                ElseIf My.Computer.Clipboard.ContainsFileDropList Then
                    var1 = My.Computer.Clipboard.GetFileDropList
                ElseIf My.Computer.Clipboard.ContainsImage Then
                    var1 = My.Computer.Clipboard.GetImage
                ElseIf My.Computer.Clipboard.ContainsText Then
                    var1 = My.Computer.Clipboard.GetText
                End If

                'With .Selection

                '    .Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                '    .PasteSpecial(DataType:=Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteOLEObject)
                'End With



                'Dim oDataObj As System.Windows.Forms.IDataObject = Clipboard.GetDataObject



                'System.Threading.Thread.Sleep(200)

                'Dim ppt As New Microsoft.Office.Interop.PowerPoint.Application

                'With ppt
                '    .Presentations.Add()
                '    .Visible = True

                '    '.ActiveWindow.View.PasteSpecial(DataType:=Microsoft.Office.Interop.PowerPoint.PpPasteDataType.ppPasteRTF, DisplayAsIcon:=Microsoft.Office.Core.MsoTriState.msoFalse, IconLabel:="New Bitmap Image")

                '    Try
                '        .ActivePresentation.Slides.Add(.ActivePresentation.Slides.Count + 1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank)
                '        '.ActiveWindow.Selection.SlideRange.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank
                '        .ActiveWindow.View.PasteSpecial(DataType:=Microsoft.Office.Interop.PowerPoint.PpPasteDataType.ppPasteMetafilePicture, DisplayAsIcon:=Microsoft.Office.Core.MsoTriState.msoFalse, IconLabel:="New Bitmap Image")

                '        .ActiveWindow.Selection.Copy()
                '    Catch ex As Exception
                '        MsgBox(ex.Message)
                '    End Try

                'End With

                ''try new wd doc
                'Dim wdN As New Word.Application
                'wdN.Documents.Add()
                'wdN.Visible = True
                'wdN.Selection.PasteSpecial(Link:=False, DataType:=Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteEnhancedMetafile)
                'wdN.ActiveDocument.Shapes(1).Select()
                'wdN.Selection.Cut()


                'wd.Visible = False

                'xlROT.Visible = True

                ''new code for clipboard to .jpg


                'Dim oDataObj As System.Windows.Forms.IDataObject = System.Windows.Forms.Clipboard.GetDataObject()

                'Dim oDataObj As System.Windows.IDataObject = System.Windows.Clipboard.GetDataObject



                'Try
                '    'If oDataObj.GetDataPresent(System.Windows.Forms.DataFormats.Bitmap) Then
                '    Dim oImgObj As System.Drawing.Image = oDataObj.GetData(DataFormats.Serializable, True)
                '    'To Save as Bitmap
                '    'oImgObj.Save("c:\Test.bmp", System.Drawing.Imaging.ImageFormat.Bmp)
                '    'To Save as Jpeg
                '    'oImgObj.Save(strEP & "\gt.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg)
                '    'To Save as Gif
                '    'oImgObj.Save("c:\Test.gif", System.Drawing.Imaging.ImageFormat.Gif)
                '    'To Save as png
                '    oImgObj.Save(strEP, System.Drawing.Imaging.ImageFormat.Png)
                '    'End If
                'Catch ex As Exception
                '    MsgBox(ex.Message)
                'End Try


                'tbl.Cell(CountA + ctHdRows + 1, 1).Select()

                ''insert figure
                'Try
                '    .Selection.InlineShapes.AddPicture(FileName:=strEP, LinkToFile:=False, SaveWithDocument:=True)
                'Catch ex As Exception

                '    Try
                '        .Selection.PasteSpecial(Link:=False, DataType:=Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteMetafilePicture, Placement:=Microsoft.Office.Interop.Word.WdOLEPlacement.wdInLine, DisplayAsIcon:=False)
                '    Catch ex1 As Exception
                '        MsgBox(ex1.Message)
                '    End Try

                '    MsgBox(ex.Message)
                'End Try

                ''end new code


                'old code

                'xlROT.DisplayAlerts = False
                'Dim wb As Microsoft.Office.Interop.Excel.Workbook
                'wb = xlROT.ActiveWorkbook

                'wb.ActiveSheet.PasteSpecial(Format:="Picture (Enhanced Metafile)", Link:=False, DisplayAsIcon:=False)


                'xlROT.Selection.Cut()

                'xlROT.DisplayAlerts = True

                'Try
                '    wb.ActiveSheet.PasteSpecial(Format:="Microsoft Word 97-2003 Document Object", Link:=False, DisplayAsIcon:=False)
                'Catch ex As Exception
                '    MsgBox(ex.Message)
                'End Try
                'Try
                '    'wb.ActiveSheet.Selection.ShapeRange.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                '    wb.ActiveSheet.shapes(wb.ActiveSheet.shapes.count).Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                'Catch ex As Exception
                '    MsgBox(ex.Message)
                'End Try


                ''wb.ActiveSheet.paste()



                tbl.Cell(CountA + ctHdRows + 1, 1).Select()
                '.Selection.PasteSpecial(Link:=False, DataType:=Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteEnhancedMetafile, Placement:=Microsoft.Office.Interop.Word.WdOLEPlacement.wdInLine, DisplayAsIcon:=False)
                '.Selection.Paste()
                '.Selection.PasteSpecial(Link:=False, DataType:=Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteMetafilePicture, Placement:=Microsoft.Office.Interop.Word.WdOLEPlacement.wdInLine, DisplayAsIcon:=False)


                Try
                    .Selection.Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
                    .Selection.PasteSpecial(Link:=False, DataType:=Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteOLEObject, Placement:=Microsoft.Office.Interop.Word.WdOLEPlacement.wdInLine, DisplayAsIcon:=False)

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                'end old code

                'new code
                'tbl.Cell(var2, 1).Select()
                '.Selection.SelectRow()
                '.Selection.MoveDown(Unit:=wdLine, Count:=17, Extend:=wdExtend)
                '.Selection.Copy()
                '.Selection.Delete(Unit:=wdCharacter, Count:=1)

                '.Selection.Tables(1).Cell(3, 1).Select()
                '.Selection.SelectRow()
                '.Selection.MoveDown(Unit:=wdLine, Count:=17, Extend:=wdExtend)

                '.Selection.Cells.Merge()
                '.Selection.PasteSpecial(Link:=False, DataType:=wdPasteEnhancedMetafile, Placement:=wdInLine, DisplayAsIcon:=False)



            Next

            frmH.lblProgress.Text = strM1
            frmH.lblProgress.Refresh()

            'delete rest of table
            'intRows = tbl.Rows.Count - intPageCount - ctHdRows

            tbl.Cell(ctHdRows + 1 + intPageCount + 1, 1).Select()

            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)
            .Selection.Rows.Delete()

            'the delete action will move cursor outside of table
            'put it back in
            intRows = tbl.Rows.Count
            tbl.Cell(intRows, 1).Select()


            ''Pause(0.25)
            'Dim intR As Int16
            'intR = .Selection.Rows.Count
            'Do Until intR >= intRows
            '    intR = .Selection.Rows.Count
            '    If intR = 0 Then
            '        Do Until intR <> 0
            '            .Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            '            intR = .Selection.Rows.Count
            '        Loop
            '    End If
            '    '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=intRows - 1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            '    intR = .Selection.Rows.Count
            '    If intR >= intRows Then
            '        Exit Do
            '    End If
            '    .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=intRows - 1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

            '    'wdd.visible = True
            'Loop
            ''.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=intRows - 1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            ''Pause(0.25)

        End With

    End Sub

End Module
