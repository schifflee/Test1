Public Class frmAddRowsChoice

    Public idCT As Int32 'id_tblConfigReportTables

    Public frm As Form
    Public boolCancel As Boolean = True
    Public intRBSel As Short = 0

    '20180803 LEE:
    Public idRT As Int32 'idRT = NZ(dgvT("ID_TBLREPORTTABLE", intRowT).Value, -1)

    Public boolrbULOQVis As Boolean = True

    'idT = dgvT("ID_TBLCONFIGREPORTTABLES", intRowT).Value
    '       idRT = NZ(dgvT("ID_TBLREPORTTABLE", intRowT).Value, -1)


    Private Sub frmAddRowsChoice_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim strM As String

        strM = "This table requires additional assignments for the sample set."
        strM = strM & ChrW(10) & ChrW(10) & "Please make a selection among the options below."
        strM = strM & ChrW(10) & ChrW(10) & "If you wish to make the assignments later, check the 'Ignore' checkbox."

        Me.lblH.Text = strM

        Call InitStuff()


    End Sub

    Function GetGB() As GroupBox


        'Legend
        'Case 13, 14, 15, 17, 22, 23, 29, 32, 35

        Dim strM As String = ""

        Select Case idCT

            Case Is = 1 'Summary of Analytical Runs
            Case Is = 2 'Summary of Regression Constants
            Case Is = 3 'Summary of Back-Calculated Calibration Std Conc
            Case Is = 4 'Summary of Interpolated QC Std Conc
            Case Is = 5 'Summary of Samples
            Case Is = 6 'Summary of Reassayed Samples
            Case Is = 7 'Summary of Repeat Samples
            Case Is = 11 'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
            Case Is = 12 'Summary of Interpolated Dilution QC Concentrations
            Case Is = 13 'Summary of Combined Recovery

                GetGB = Me.gbRecovery
                strM = "Recovery"
                GetGB.Text = strM

                Me.rbQC.Top = 24
                Me.rbRS.Top = 51
                Me.rbPES.Visible = False

            Case Is = 14 'Summary of True Recovery

                GetGB = Me.gbRecovery
                strM = "Recovery"
                GetGB.Text = strM

                Me.rbQC.Top = 24
                Me.rbPES.Top = 51
                Me.rbRS.Visible = False

            Case Is = 15 'Summary of Suppression/Enhancement

                GetGB = Me.gbRecovery
                strM = "Recovery or Matrix Factor"
                GetGB.Text = strM

                Me.rbRS.Top = 24
                Me.rbPES.Top = 51
                Me.rbQC.Visible = False

            Case Is = 17 'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments

                GetGB = Me.gbULots

            Case Is = 18 'Summary of [Period Temp] Stability in Matrix
            Case Is = 19 'Summary of Freeze/Thaw [#Cycles] Stability in Matrix
            Case Is = 21 '[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
            Case Is = 22 '[Period Temp] Stock Solution Stability Assessment

                GetGB = Me.gbStockSoln

            Case Is = 23 '[Period Temp] Spiking Solution Stability Assessment

                GetGB = Me.gbSpikingSoln

            Case Is = 29 '[Period Temp] Long-Term QC Std Storage Stability

                GetGB = Me.gbLongTermQC

            Case Is = 30 'Incurred Samples
            Case Is = 31 'Ad Hoc QC Stability Table
            Case Is = 32 'Ad Hoc QC Stability Comparison Table

                GetGB = Me.gbAdHocStabComp

            Case Is = 33 'System Suitability Table v1
            Case Is = 34 'Selectivity in Individual Lots Table v1
            Case Is = 35 'Carryover in Individual Lots Table v1

                GetGB = Me.gbCarryOver

                '20181011 LEE:
                'If user has excluded ULOQ, must hide gbULOQ
                If boolIncludePSAE Then
                    Me.rbULOQ.Visible = False
                    boolrbULOQVis = False
                Else
                    Me.rbULOQ.Visible = True
                    boolrbULOQVis = True
                End If

            Case Is = 36 'Method Trial Back-Calculated Calibration Std Conc v1
            Case Is = 37 'Method Trial Control and Fortified QC Samples v1
            Case Is = 38 'Method Trial Incurred Blinded Samples v1


        End Select

    End Function

    Sub InitStuff()

        Me.Width = 619

        Call HideGroupBoxes()

        Me.rbQC.Left = 23
        Me.rbQC.Top = 51

        Dim gb As GroupBox
        Dim ctrl As Control
        Dim str1 As String
        Dim str2 As String

        Dim BorderWidth As Int32 = (Me.Width - Me.ClientSize.Width) / 2
        Dim TitlebarHeight As Int32 = Me.Height - Me.ClientSize.Height - 2 * BorderWidth


        Dim a, b, c
        Dim w, wh, bh


        wh = Me.Width / 2

        'place checkbox
        a = Me.lblH.Top + Me.lblH.Height + 10

        Me.chkIgnore.Top = a

        bh = Me.chkIgnore.Width / 2
        c = wh - bh
        Me.chkIgnore.Left = c

        'make each gb the same width
        Dim wb
        wb = Me.gbRecovery.Width
        For Each ctrl In Me.Controls
            str1 = Mid(ctrl.Name, 1, 2)
            If StrComp(str1, "gb", CompareMethod.Text) = 0 Then
                ctrl.Width = wb
            End If
        Next

        'place each gb top
        a = Me.chkIgnore.Top + Me.chkIgnore.Height + 10
        For Each ctrl In Me.Controls
            str1 = Mid(ctrl.Name, 1, 2)
            If StrComp(str1, "gb", CompareMethod.Text) = 0 Then
                ctrl.Top = a
            End If
        Next

        gb = GetGB()

        bh = gb.Width / 2
        c = wh - bh
        gb.Left = c

        'place cmds
        a = gb.Top + gb.Height + 20

        Me.cmdOK.Top = a
        Me.cmdCancel.Top = a

        Me.cmdOK.Left = wh - (Me.cmdOK.Width + 20)
        Me.cmdCancel.Left = wh + 20

        gb.Visible = True
   

        'set form height

        a = Me.cmdOK.Top + Me.cmdOK.Height + 40 + TitlebarHeight

        Me.Height = a

        'now place form in center
        Dim mainScreen As Screen = Screen.FromPoint(Me.Location)
        Dim X As Integer = (mainScreen.WorkingArea.Width - Me.Width) / 2 + mainScreen.WorkingArea.Left
        Dim Y As Integer = (mainScreen.WorkingArea.Height - Me.Height) / 2 + mainScreen.WorkingArea.Top

        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New System.Drawing.Point(X, Y)

        Try
            Me.txtAdHocStabComp.Focus()
        Catch ex As Exception

        End Try

        '20180803 LEE:
        'Table 17 Unique lots needs to be evaluated for Matrix Effect

        Dim intRowT As Int16
        Dim strF As String
        Dim var1, var2, var3
        Dim Count1 As Short
        Dim Count2 As Short

        Try


            Dim rowsTP() As DataRow
            Dim intTP As Short

            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idRT
            rowsTP = tblTableProperties.Select(strF)

            If rowsTP.Length = 0 Then
                GoTo end1
            End If

            var1 = NZ(rowsTP(0).Item("BOOLDOINDREC"), 0)
            If var1 = 0 Then
            Else
                'needs matrix effect
                gb = Me.gbULots
                Dim rb As RadioButton

                For Count1 = 1 To 10
                    str1 = "rbLot" & Count1
                    rb = gb.Controls(str1)
                    If Count1 = 1 Then
                        str2 = "Solvent"
                    Else
                        str2 = "Lot " & Count1 - 1
                    End If
                    rb.Text = str2
                    'Me.Controls(str1).Text = str2
                    var1 = var1
                Next
            End If

end1:

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try



    End Sub

    Sub HideGroupBoxes()

        Dim gb As GroupBox
        Dim ctrl As Control
        Dim str1 As String
        Dim str2 As String
        For Each ctrl In Me.Controls
            str1 = Mid(ctrl.Name, 1, 2)
            If StrComp(str1, "gb", CompareMethod.Text) = 0 Then
                ctrl.Visible = False
            End If
        Next

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub rbLot1_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot1.CheckedChanged

        intRBSel = 0

    End Sub

    Private Sub rbLot2_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot2.CheckedChanged

        intRBSel = 1

    End Sub

    Private Sub rbLot3_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot3.CheckedChanged

        intRBSel = 2

    End Sub

    Private Sub rbLot4_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot4.CheckedChanged

        intRBSel = 3

    End Sub

    Private Sub rbLot5_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot5.CheckedChanged

        intRBSel = 4

    End Sub

    Private Sub rbLot6_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot6.CheckedChanged

        intRBSel = 5

    End Sub

    Private Sub rbLot7_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot7.CheckedChanged

        intRBSel = 6

    End Sub

    Private Sub rbLot8_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot8.CheckedChanged

        intRBSel = 7

    End Sub

    Private Sub rbLot9_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot9.CheckedChanged

        intRBSel = 8

    End Sub

    Private Sub rbLot10_CheckedChanged(sender As Object, e As EventArgs) Handles rbLot10.CheckedChanged

        intRBSel = 9

    End Sub

    Private Sub rbLLOQ_CheckedChanged(sender As Object, e As EventArgs) Handles rbLLOQ.CheckedChanged

        intRBSel = 0

    End Sub

    Private Sub rbULOQ_CheckedChanged(sender As Object, e As EventArgs) Handles rbULOQ.CheckedChanged

        intRBSel = 1

    End Sub

    Private Sub rbBlank_CheckedChanged(sender As Object, e As EventArgs) Handles rbBlank.CheckedChanged

        intRBSel = 2

    End Sub

    Private Sub rbOldStockSoln_CheckedChanged(sender As Object, e As EventArgs) Handles rbOldStockSoln.CheckedChanged

        intRBSel = 0

    End Sub

    Private Sub rbNewStockSoln_CheckedChanged(sender As Object, e As EventArgs) Handles rbNewStockSoln.CheckedChanged

        intRBSel = 1

    End Sub

    Private Sub rbOldSpikingSoln_CheckedChanged(sender As Object, e As EventArgs) Handles rbOldSpikingSoln.CheckedChanged

        intRBSel = 0

    End Sub

    Private Sub rbNewSpikingSoln_CheckedChanged(sender As Object, e As EventArgs) Handles rbNewSpikingSoln.CheckedChanged

        intRBSel = 1

    End Sub

    Private Sub rbLTQCT0_CheckedChanged(sender As Object, e As EventArgs) Handles rbLTQCT0.CheckedChanged

        intRBSel = 0

    End Sub

    Private Sub rbLTQCLTA_CheckedChanged(sender As Object, e As EventArgs) Handles rbLTQCLTA.CheckedChanged

        intRBSel = 1

    End Sub
End Class