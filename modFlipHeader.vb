Option Compare Text

Imports Microsoft.Office.Interop.Word
Imports Word = Microsoft.Office.Interop.Word

Module modFlipHeader


    Public boolFlipHeaderAuto As Boolean = False
    Public boolHeaderIsText As Boolean = True
    Public boolFooterIsText As Boolean = True

    Sub FlipHeader(wd As Word.Application)

        Dim strM As String
        Try
            Call aDo(wd)
        Catch ex As Exception
            strM = "There seems to be a problem Flipping Headers:" & ChrW(10) & ChrW(10)
            strM = strM & ex.Message

            If boolDisableWarnings Then
            Else
                MsgBox(strM, MsgBoxStyle.Information, "Problem Flipping Headers...")
            End If

        End Try


    End Sub

    Sub DeleteShapes(wd As Word.Application)

        'not in use

        Dim shp As Microsoft.Office.Interop.Word.Shape ' Shape
        Dim sec As Word.Section ' Section

        For Each sec In wd.ActiveDocument.Sections

            For Each shp In sec.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Shapes

                shp.Delete()

            Next

            For Each shp In sec.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage).Shapes

                shp.Delete()

            Next
        Next



    End Sub

    Sub aDo(wd As Word.Application)

        Dim intS As Integer
        Dim shp As Word.Shape
        Dim Ishp As InlineShape
        Dim Ishp1 As InlineShape
        Dim tbxName As String
        Dim var1, var2
        Dim rng1 As Word.Range
        Dim sec As Section

        Dim dT As Single
        Dim dB As Single
        Dim dL As Single
        Dim dR As Single
        Dim dPW As Single
        Dim dPH As Single 'page height
        Dim dF As Single 'footer
        Dim dH As Single 'header

        Dim pT As Single
        Dim pB As Single
        Dim pL As Single
        Dim pR As Single
        Dim pPH As Single
        Dim pPW As Single
        Dim pF As Single 'footer
        Dim pH As Single 'header

        Dim a As Single
        Dim b As Single
        Dim c As Single
        Dim d As Single

        Dim oHPos As Single
        Dim oVPos As Single
        Dim pWid As Single
        Dim pHt As Single

        Dim intSections As Integer
        Dim intSection As Integer

        Dim boolText As Boolean

        Dim shpHAP As Single
        Dim shpVAP As Single

        Dim boolV As Boolean = wd.Visible


        boolText = False

        wd.Selection.HomeKey(Unit:=Word.WdUnits.wdStory)

        wd.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader 'or Word.WdSeekView.wdSeekCurrentPageFooter

        '    ActiveWindow.ActivePane.View.NextHeaderFooter
        '    WordBasic.GoToFooter

        frmH.lblProgress.Text = "Flipping landscape headers/footers to left/right margins..."
        frmH.lblProgress.Refresh()
        frmH.pb1.Maximum = wd.ActiveDocument.Sections.Count
        frmH.pb1.Value = 1
        frmH.pb1.Refresh()

        intSection = 0
        For Each sec In wd.ActiveDocument.Sections

            intSection = intSection + 1

            frmH.lblProgress.Text = "Flipping landscape headers/footers to left/right margins..."
            frmH.lblProgress.Refresh()
            frmH.pb1.Value = intSection
            frmH.pb1.Refresh()

            sec.Range.Select()

            'This needs work
            'If two L in a row and link is true, second one doesn't process
            'If the 'linktoprevious' in IF below is left out, margins become progressively worse.
            If sec.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape And sec.Headers.Item(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False Then

                wd.Visible = True

                'do header

                wd.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader 'or Word.WdSeekView.wdSeekCurrentPageFooter

                'sec.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary)
                'sec.Headers(wdHeaderFooterFirstPage)

                dT = sec.PageSetup.TopMargin
                dB = sec.PageSetup.BottomMargin
                dL = sec.PageSetup.LeftMargin
                dR = sec.PageSetup.RightMargin
                dPW = sec.PageSetup.PageWidth
                dPH = sec.PageSetup.PageHeight
                dH = sec.PageSetup.HeaderDistance
                dF = sec.PageSetup.FooterDistance

                ''top margin becomes left margin + header distance + shpH
                'pL = dT + dH
                ''bottom margin becomes right margin + footer distance + shpH
                'pR = dB + dF
                ''left margin becomes becomes bottom margin 
                'pB = dL
                ''right margin becomes top margin
                'pT = dR

                'sec.PageSetup.TopMargin = pT
                'sec.PageSetup.BottomMargin = pB
                'sec.PageSetup.LeftMargin = pL
                'sec.PageSetup.RightMargin = pR

                ''reset margins because the p variables get used for something else later
                'dT = pT
                'dB = pB
                'dL = pL
                'dR = pR

                'do pictures first

                'pictures may be inlineshapes or shapes
                'must do both

                For Each shp In sec.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ShapeRange


                    shp.Select()
                    var1 = "a"
                    With wd.Selection

                        .ShapeRange.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                        .ShapeRange.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                        .ShapeRange.RelativeHorizontalSize = Word.WdRelativeHorizontalSize.wdRelativeHorizontalSizePage
                        .ShapeRange.RelativeVerticalSize = Word.WdRelativeVerticalSize.wdRelativeVerticalSizePage

                    End With


                    With wd.Selection

                        'in order to do next text section
                        'must cut/paste into header
                        Try
                            shp.Rotation = 90.0#
                        Catch ex As Exception
                            'MsgBox(ex.Message)
                        End Try

                    End With

                Next

                For Each Ishp In sec.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.InlineShapes


                    Ishp.Select()
                    var1 = "a"

                    shp = Ishp.ConvertToShape
                    With wd.Selection

                        .ShapeRange.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                        .ShapeRange.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                        .ShapeRange.RelativeHorizontalSize = Word.WdRelativeHorizontalSize.wdRelativeHorizontalSizePage
                        .ShapeRange.RelativeVerticalSize = Word.WdRelativeVerticalSize.wdRelativeVerticalSizePage

                    End With


                    With wd.Selection

                        'in order to do next text section
                        'must cut/paste into header
                        Try
                            Ishp.ConvertToShape.IncrementRotation(90.0#)
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try

                    End With

                Next

                'top margin  + header distance + shpH becomes left margin
                Try
                    pL = dT + dH ' + shp.Height
                Catch ex As Exception
                    pL = dT + dH
                End Try

                'bottom margin  + footer distance + shpH becomes right margin
                Try
                    pR = dB + dF ' + shp.Height
                Catch ex As Exception
                    pR = dB + dF
                End Try

                'left margin becomes becomes bottom margin 
                pB = dL
                'right margin becomes top margin
                pT = dR

                sec.PageSetup.TopMargin = pT
                sec.PageSetup.BottomMargin = pB
                sec.PageSetup.LeftMargin = pL
                sec.PageSetup.RightMargin = pR

                'reset margins because the p variables get used for something else later
                dT = pT
                dB = pB
                dL = pL
                dR = pR

                'now do text

                'calculate textbox left
                pL = dPW - dR

                'calculate textbox top
                pT = dT ' dPH - dT

                'calculate height
                pPH = dPH - dT - dB

                'calculate width
                pPW = dR ' dPW - dL - dR

                wd.Selection.WholeStory()
                wd.Selection.Cut()

                With sec.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary)

                    'shp = .Shapes.AddTextbox(Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, aLeft, aTop, aWidth, aHeight, )
                    shp = .Shapes.AddTextbox(Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 1, 1, 72, 432, .Range) ' Anchor.Paragraphs.First.Range
                    'shp = .Shapes.AddTextbox(Word.WdTextOrientation.wdTextOrientationHorizontal, 1, 1, 72, 432, .Range)
                    tbxName = shp.Name

                End With

                '****
                wd.Selection.HeaderFooter.Shapes(tbxName).Select()
                wd.Selection.WholeStory()
                wd.Selection.Delete(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                '    Selection.ShapeRange.Height = 432#
                '    Selection.ShapeRange.Width = 72#

                With wd.Selection

                    .WholeStory()
                    .Delete(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                    .PasteAndFormat(Word.WdRecoveryType.wdPasteDefault)
                    .Delete(Unit:=Word.WdUnits.wdCharacter, Count:=1)

                    .Orientation = Word.WdTextOrientation.wdTextOrientationDownward

                    .ShapeRange.TextFrame.AutoSize = False

                    .ShapeRange.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                    .ShapeRange.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                    .ShapeRange.RelativeHorizontalSize = Word.WdRelativeHorizontalSize.wdRelativeHorizontalSizePage
                    .ShapeRange.RelativeVerticalSize = Word.WdRelativeVerticalSize.wdRelativeVerticalSizePage

                    .ShapeRange.Line.Visible = False ' Office.Core.MsoTriState.msoFalse
                    .ShapeRange.LockAspectRatio = False ' Office.Core.MsoTriState.msoFalse

                    .ShapeRange.Line.Visible = False


                    .ShapeRange.Height = pPH ' 432# '85.7
                    'width needs to be margin
                    .ShapeRange.Width = pPW
                    '            .ShapeRange.Width = 72#
                    .ShapeRange.Left = pL ' 558#
                    .ShapeRange.Top = pT ' -10  '113.75
                    .ShapeRange.TextFrame.MarginLeft = 7.2
                    .ShapeRange.TextFrame.MarginRight = 7.2
                    .ShapeRange.TextFrame.MarginTop = 3.6
                    .ShapeRange.TextFrame.MarginBottom = 3.6
                    .ShapeRange.WrapFormat.DistanceTop = 0 ' InchesToPoints(0)
                    .ShapeRange.WrapFormat.DistanceBottom = 0 ' InchesToPoints(0)
                    .ShapeRange.WrapFormat.DistanceLeft = 9.36 ' InchesToPoints(0.13)
                    .ShapeRange.WrapFormat.DistanceRight = 9.36 ' InchesToPoints(0.13)

                    'do all this to send text to back

                    .ShapeRange.TopRelative = Word.WdShapePositionRelative.wdShapePositionRelativeNone
                    .ShapeRange.WidthRelative = Word.WdShapeSizeRelative.wdShapeSizeRelativeNone
                    .ShapeRange.HeightRelative = Word.WdShapeSizeRelative.wdShapeSizeRelativeNone
                    .ShapeRange.LockAnchor = False
                    .ShapeRange.LayoutInCell = True
                    .ShapeRange.WrapFormat.AllowOverlap = True
                    .ShapeRange.WrapFormat.Side = Word.WdWrapSideType.wdWrapBoth

                    .ShapeRange.WrapFormat.Type = 3
                    .ShapeRange.ZOrder(5)
                    .ShapeRange.TextFrame.AutoSize = False
                    .ShapeRange.TextFrame.WordWrap = True
                    .ShapeRange.TextFrame.VerticalAnchor = Office.Core.MsoVerticalAnchor.msoAnchorTop


                    .ShapeRange.WrapFormat.AllowOverlap = True
                    .ShapeRange.WrapFormat.Side = Word.WdWrapSideType.wdWrapBoth
                    '    .ShapeRange.WrapFormat.DistanceTop = InchesToPoints(0)
                    '    .ShapeRange.WrapFormat.DistanceBottom = InchesToPoints(0)
                    '    .ShapeRange.WrapFormat.DistanceLeft = InchesToPoints(0.13)
                    '    .ShapeRange.WrapFormat.DistanceRight = InchesToPoints(0.13)
                    .ShapeRange.WrapFormat.Type = 3
                    .ShapeRange.ZOrder(5)
                    .ShapeRange.TextFrame.AutoSize = True
                    .ShapeRange.TextFrame.WordWrap = True

                    'pPW = .ShapeRange.Width 'use for footer text box


                    '***

                End With



                'now do footer

                wd.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter 'or wdSeekCurrentPageFooter

                '        'legend
                '        dT = sec.PageSetup.TopMargin
                '        dB = sec.PageSetup.BottomMargin
                '        dL = sec.PageSetup.LeftMargin
                '        dR = sec.PageSetup.RightMargin
                '        dPW = sec.PageSetup.PageWidth
                '        dPH = sec.PageSetup.PageHeight

                'do pictures first

                For Each Ishp In sec.Footers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.InlineShapes

                    Ishp.Select()
                    var1 = "a"

                    shp = Ishp.ConvertToShape
                    With wd.Selection

                        .ShapeRange.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                        .ShapeRange.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                        .ShapeRange.RelativeHorizontalSize = Word.WdRelativeHorizontalSize.wdRelativeHorizontalSizePage
                        .ShapeRange.RelativeVerticalSize = Word.WdRelativeVerticalSize.wdRelativeVerticalSizePage

                    End With

                    oHPos = shp.Left
                    oVPos = shp.Top
                    pWid = shp.Width
                    pHt = shp.Height

                    'create new pL and pT
                    pL = (dPH - (oVPos + pHt)) 'should be ~50 :  599-(527-25)

                    'calculate  top
                    pT = dT ' dPH - dT

                    pT = ((oHPos / dPW) * dPH)

                    With wd.Selection

                        'in order to do next text section
                        'must cut/paste into header
                        'Ishp.ConvertToShape.IncrementRotation(90.0#)
                        Try
                            Ishp.ConvertToShape.IncrementRotation(90.0#)
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try

                    End With


                Next

                'now do text

                'calculate textbox left
                pL = 0 '(dPW - dR) * 0.5
                pL = dL - pPW 'pPW comes from header text box

                'calculate textbox top
                pT = dT ' dPH - dT

                'calculate height
                pPH = dPH - dT - dB

                'calculate width
                pPW = dL ' dPW - dL - dR

                wd.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter
                wd.Selection.WholeStory()
                wd.Selection.Cut()

                With sec.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary)

                    'shp = .Shapes.AddTextbox(Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, aLeft, aTop, aWidth, aHeight, )
                    shp = .Shapes.AddTextbox(Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 1, 1, 72, 432, .Range) ' Anchor.Paragraphs.First.Range

                    tbxName = shp.Name

                End With

                wd.Selection.HeaderFooter.Shapes(tbxName).Select()
                wd.Selection.WholeStory()

                With wd.Selection

                    .WholeStory()
                    .Delete(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                    .PasteAndFormat(Word.WdRecoveryType.wdPasteDefault)

                    .Orientation = Word.WdTextOrientation.wdTextOrientationDownward
                    .ShapeRange.TextFrame.AutoSize = False

                    .ShapeRange.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                    .ShapeRange.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                    .ShapeRange.RelativeHorizontalSize = Word.WdRelativeHorizontalSize.wdRelativeHorizontalSizePage
                    .ShapeRange.RelativeVerticalSize = Word.WdRelativeVerticalSize.wdRelativeVerticalSizePage

                    .ShapeRange.Line.Visible = False ' Office.Core.MsoTriState.msoFalse
                    .ShapeRange.LockAspectRatio = False ' Office.Core.MsoTriState.msoFalse

                    'do all this to send text to back

                    .ShapeRange.TextFrame.MarginLeft = 7.2
                    .ShapeRange.TextFrame.MarginRight = 7.2
                    .ShapeRange.TextFrame.MarginTop = 3.6
                    .ShapeRange.TextFrame.MarginBottom = 3.6
                    .ShapeRange.WrapFormat.DistanceTop = 0 ' InchesToPoints(0)
                    .ShapeRange.WrapFormat.DistanceBottom = 0 ' InchesToPoints(0)
                    .ShapeRange.WrapFormat.DistanceLeft = 9.36 ' InchesToPoints(0.13)
                    .ShapeRange.WrapFormat.DistanceRight = 9.36 ' InchesToPoints(0.13)

                    .ShapeRange.TopRelative = Word.WdShapePositionRelative.wdShapePositionRelativeNone
                    .ShapeRange.WidthRelative = Word.WdShapeSizeRelative.wdShapeSizeRelativeNone
                    .ShapeRange.HeightRelative = Word.WdShapeSizeRelative.wdShapeSizeRelativeNone

                    .ShapeRange.LockAnchor = False
                    .ShapeRange.LayoutInCell = True
                    .ShapeRange.WrapFormat.AllowOverlap = True
                    .ShapeRange.WrapFormat.Side = Word.WdWrapSideType.wdWrapBoth

                    .ShapeRange.WrapFormat.Type = 3
                    .ShapeRange.ZOrder(5)
                    .ShapeRange.TextFrame.AutoSize = False
                    .ShapeRange.TextFrame.WordWrap = True
                    .ShapeRange.TextFrame.VerticalAnchor = Office.Core.MsoVerticalAnchor.msoAnchorTop

                    .ShapeRange.WrapFormat.AllowOverlap = True
                    .ShapeRange.WrapFormat.Side = Word.WdWrapSideType.wdWrapBoth
                    '    .ShapeRange.WrapFormat.DistanceTop = InchesToPoints(0)
                    '    .ShapeRange.WrapFormat.DistanceBottom = InchesToPoints(0)
                    '    .ShapeRange.WrapFormat.DistanceLeft = InchesToPoints(0.13)
                    '    .ShapeRange.WrapFormat.DistanceRight = InchesToPoints(0.13)
                    .ShapeRange.WrapFormat.Type = 3

                    .ShapeRange.TextFrame.AutoSize = True
                    .ShapeRange.TextFrame.WordWrap = True

                    .ShapeRange.ZOrder(5)
                    .ShapeRange.Height = pPH ' 432# '85.7
                    'width
                    .ShapeRange.Width = pPW
                    '.ShapeRange.Left = pL ' 558#
                    .ShapeRange.Left = 0
                    .ShapeRange.Top = pT ' -10  '113.75


                End With

                'seems way to complicated.
                'just switch header and footer

                'legend
                'dT = sec.PageSetup.TopMargin
                'dB = sec.PageSetup.BottomMargin
                'dL = sec.PageSetup.LeftMargin
                'dR = sec.PageSetup.RightMargin
                'dPW = sec.PageSetup.PageWidth
                'dPH = sec.PageSetup.PageHeight

                ''top margin becomes left margin + header distance
                'pL = dT + dH
                ''bottom margin becomes right margin
                'pR = dB
                ''left margin becomes becomes bottom margin + foother distance
                'pB = dL + dF
                ''right margin becomes top margin
                'pT = dR

                'sec.PageSetup.TopMargin = pT
                'sec.PageSetup.BottomMargin = pB
                'sec.PageSetup.LeftMargin = pL
                'sec.PageSetup.RightMargin = pR

            End If


            Try
                wd.ActiveWindow.ActivePane.View.NextHeaderFooter()
            Catch ex As Exception

            End Try

        Next

        wd.Visible = boolV

        wd.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument ' wdSeekMainDocument
        wd.Selection.HomeKey(Unit:=Word.WdUnits.wdStory)

        frmH.lblProgress.Text = "Finished Flipping landscape headers/footers to left/right margins..."
        frmH.lblProgress.Refresh()
        frmH.pb1.Value = wd.ActiveDocument.Sections.Count
        frmH.pb1.Refresh()

    End Sub

End Module
