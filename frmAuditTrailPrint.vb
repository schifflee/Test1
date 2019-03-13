Option Compare Text

Public Class frmAuditTrailPrint


    Private checkPrint As Integer
    Private pagenum As Integer
    Private boolSomePages As Boolean = False
    Private intSP As Int32 = 1
    Private intEP As Int32 = 999999
    Private intTP As Int32 = 999999
    Private strDt As String
    Private strTime As String
    Private LineCounter As Int64 = 0
    Private intPageL As Int64 = 0
    Public intSelPage As Int32 = 0
    Private intAP As Int32 = 0 'actual page

    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        checkPrint = 0
        pagenum = 0
        LineCounter = 0
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim prFont As New Font("Verdana", 10, GraphicsUnit.Point)


        '*****

        Dim Y As Integer
        Dim str1 As String
        Dim boolNP As Boolean = False 'next page
        Dim lp As Short = 37 'number lines on a page
        Dim ls As Short = 0 'linecounter start

        'Note: 1 page = 38 lines

        Dim stringFormat As New StringFormat()

        'Dim tabs As Single() = {150, 100, 100, 100}
        Dim tabs As Single() = {300}

        stringFormat.SetTabStops(0, tabs)

        Do While pagenum < intSP - 1 And LineCounter < Me.rtfPrint.Lines.Length - 1

            LineCounter += 1
            ls = ls + 1

            str1 = Me.rtfPrint.Lines(LineCounter).ToString

            If ls = lp Then
                pagenum = pagenum + 1
                ls = 0
            End If

        Loop

        Y = 72
        checkPrint = intPageL
        intPageL = 0

        ls = 0
        LineCounter = LineCounter - 1

        'keep this next line for future reference
        ''checkPrint = Me.rtfPrint.Print(checkPrint, Me.rtfPrint.TextLength, e)

        Do While pagenum < intEP And LineCounter < Me.rtfPrint.Lines.Length - 1

            LineCounter += 1
            ls = ls + 1
            str1 = Me.rtfPrint.Lines(LineCounter).ToString

            e.Graphics.DrawString(Me.rtfPrint.Lines(LineCounter), Me.rtfPrint.Font, Brushes.Black, 72, Y, stringFormat)

            If ls = lp Then
                'e.Graphics.DrawString("Page " & pagenum & " of " & intTP, prFont, Brushes.Black, 700, 50)
                intAP = intAP + 1
                e.Graphics.DrawString("Page " & intAP & " of " & intTP, prFont, Brushes.Black, 700, 50)
                pagenum += 1
                ls = 0
                If pagenum < intEP Then
                    e.HasMorePages = True
                Else
                    e.HasMorePages = False
                End If

                Exit Do
            End If

            Y += Me.rtfPrint.Font.Height

        Loop

    End Sub

    Sub PrintThis()

        Dim minPage
        Dim maxPage
        Dim var1

        Dim dt As Date = Now
        strDt = FormatDateTime(dt, DateFormat.LongDate)
        strTime = FormatDateTime(dt, DateFormat.LongTime)

        Me.PrintDialog1.AllowSomePages = True
        Me.PrintDialog1.AllowSelection = False
        Me.PrintDialog1.AllowCurrentPage = True
        'Me.PrintDialog1.PrinterSettings.MinimumPage = 1
        'Me.PrintDialog1.PrinterSettings.MaximumPage = 10


        Me.PrintDialog1.Document = Me.PrintDocument1

        If Me.PrintDialog1.ShowDialog() = DialogResult.OK Then

            Me.PrintDocument1.PrinterSettings = Me.PrintDialog1.PrinterSettings

            If Me.PrintDocument1.PrinterSettings.PrintRange = Printing.PrintRange.SomePages Then
                intSP = Me.PrintDocument1.PrinterSettings.FromPage
                intEP = Me.PrintDocument1.PrinterSettings.ToPage
                intTP = intEP - intSP + 1
            ElseIf Me.PrintDocument1.PrinterSettings.PrintRange = Printing.PrintRange.CurrentPage Then
                intSP = intSelPage
                intEP = intSelPage
                intTP = 1
            Else
                intSP = 1
                intEP = 9999999
            End If


            'Me.PrintDocument1.PrinterSettings.MinimumPage = 1
            'Me.PrintDocument1.PrinterSettings.MaximumPage = 10

            Me.PrintDocument1.Print()

        End If

    End Sub

    Sub FormatThis()

        Me.rtfPrint.AcceptsTab = True
        Me.rtfPrint.SelectionIndent = 10 ' 100
        Me.rtfPrint.SelectionRightIndent = 25
        Me.rtfPrint.SelectionHangingIndent = 300

        'Me.rtfPrint.SelectionTabs = New Integer() {100, 80, 120, 160}
        Me.rtfPrint.SelectionTabs = New Integer() {300}

        'clear rtf
        Me.rtfPrint.Text = ""
        'now load rtf
        Me.rtfPrint.WordWrap = True

    End Sub

    Private Sub frmAuditTrailPrint_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class