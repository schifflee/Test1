Option Compare Text

Public Class frmProgress_01
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Sub crap()



    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProgress_01))
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblProgress
        '
        Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress.ForeColor = System.Drawing.Color.Lime
        Me.lblProgress.Location = New System.Drawing.Point(8, 16)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(416, 80)
        Me.lblProgress.TabIndex = 0
        Me.lblProgress.Text = "Label1"
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmProgress_01
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(432, 109)
        Me.Controls.Add(Me.lblProgress)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmProgress_01"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Progress..."
        Me.ResumeLayout(False)

    End Sub

#End Region


    'Sub Prac()

    '    Dim fso, fo, fi
    '    Dim bool
    '    Dim Count1
    '    Dim strPath1
    '    Dim strPath2
    '    Dim strPath3

    '    'MsgBox "Start Prod Dir"

    '    strPath1 = "C:\LABIntegrity"
    '    strPath2 = strPath1 & "\StudyDoc"
    '    strPath3 = "Start"
    '    fso = CreateObject("Scripting.FileSystemObject")
    '    For Count1 = 1 To 12
    '        strPath3 = ""
    '        Select Case Count1
    '            Case 1
    '                strPath3 = strPath1
    '            Case 2
    '                strPath3 = strPath2
    '            Case 3
    '                strPath3 = strPath2 & "\ArchivedMDBs"
    '            Case 4
    '                strPath3 = strPath2 & "\DOC"
    '            Case 5
    '                strPath3 = strPath2 & "\Ini"
    '            Case 6
    '                strPath3 = strPath2 & "\MDBDatabase"
    '            Case 7
    '                'strPath3 =  strPath2 & "\ReportTemplates"
    '            Case 8
    '                strPath3 = strPath2 & "\Temp"
    '            Case 9
    '                strPath3 = strPath2 & "\XML"
    '            Case 10
    '                strPath3 = strPath2 & "\Test"
    '            Case 11
    '                strPath3 = strPath2 & "\Manuals"
    '            Case 12
    '                strPath3 = strPath2 & "\Logs"

    '        End Select

    '        If Len(strPath3) = 0 Then
    '        Else
    '            bool = fso.FolderExists(strPath3)
    '            If bool = True Then
    '            Else
    '                fso.CreateFolder(strPath3)
    '            End If
    '        End If

    '    Next

    '    'now copy files to the new directories
    '    Dim tardir
    '    Dim strPath4
    '    Dim bool1
    '    Dim strVersion
    '    Dim var1
    '    Dim var2
    '    Dim var3

    '    'Help for CustomActionData:
    '    'http://msdn.microsoft.com/en-us/library/vstudio/9cdb5eda(v=vs.100).aspx
    '    'USE [TARGETDIR]
    '    'tardir = Session.Property("CustomActionData")
    '    If fso.folderexists("C:\Program Files (x86)\LABIntegrity\") Then
    '        tardir = "C:\Program Files (x86)\LABIntegrity\StudyDoc\"
    '    Else
    '        tardir = "C:\Program Files\LABIntegrity\StudyDoc\"
    '    End If


    '    'Hmmm. tardir is coming up null
    '    'msgbox(tardir)
    '    'tardir = "C:\"

    '    For Count1 = 1 To 8
    '        bool = False
    '        bool1 = False

    '        strPath3 = ""
    '        strPath4 = ""
    '        Select Case Count1
    '            Case 1
    '                strPath3 = strPath2 & "\MDBDatabase\StudyDoc_01.mdb"
    '                strPath4 = tardir & "PackagedComponents\DatabaseInstallation\StudyDoc_01.mdb" 'note: tardir already ends in backslash
    '            Case 2
    '                strPath3 = strPath2 & "\Ini\StudyDoc.ini"
    '                strPath4 = tardir & "PackagedComponents\StudyDoc.ini"
    '                'case 3
    '                'strpath3 = strPath2 & "\ReportTemplates\SampleAnalysisTemplate_04.doc"
    '                'strpath4 = tardir & "TempDocs\SampleAnalysisTemplate_04.doc"
    '            Case 3
    '                strPath3 = strPath2 & "\Manuals\StudyDoc_AdminManual_020207.doc"
    '                strPath4 = tardir & "PackagedComponents\TempDocs\StudyDoc_AdminManual_020207.doc"
    '            Case 4
    '                strPath3 = strPath2 & "\Manuals\StudyDoc_UsersGuide_020207.pdf"
    '                strPath4 = tardir & "PackagedComponents\TempDocs\StudyDoc_UsersGuide_020207.pdf"
    '            Case 5
    '                strPath3 = strPath2 & "\Manuals\StudyDoc_Updates_02.doc"
    '                strPath4 = tardir & "PackagedComponents\TempDocs\StudyDoc_Updates_01.doc"
    '            Case 6
    '                strPath3 = strPath2 & "\Manuals\StudyDoc_Updates_02.pdf"
    '                strPath4 = tardir & "PackagedComponents\TempDocs\StudyDoc_Updates_01.pdf"
    '            Case 7
    '                strPath3 = strPath2 & "\ArchivedMDBs\MethodVal.MDB"
    '                strPath4 = tardir & "PackagedComponents\DatabaseInstallation\MethodVal.MDB"
    '            Case 8
    '                strPath3 = strPath2 & "\ArchivedMDBs\SampleAnalysis.MDB"
    '                strPath4 = tardir & "PackagedComponents\DatabaseInstallation\SampleAnalysis.MDB"

    '        End Select

    '        bool = fso.FileExists(strPath3)
    '        bool1 = fso.FileExists(strPath4)

    '        'msgbox("bool: " & strpath3 & ", bool1: " & strpath4)

    '        If bool1 = True Then
    '            If bool = True Then
    '            Else
    '                fso.CopyFile(strPath4, strPath3, False)
    '            End If
    '        End If

    '    Next

    '    'msgbox "Done copying files to GuWu folders"

    '    Dim oShell
    '    Dim strP
    '    oShell = CreateObject("WScript.Shell")

    '    'strP = "C:\LabIntegrity\StudyDoc\Bin\officeviewer.ocx"

    '    'oShell.run "regsvr32.exe /s """ & strP & """"
    '    'Set oShell = Nothing

    '    fso = Nothing

    '    'MsgBox "Done Prod Dirs"

    '    'msgbox strP


    '    'oShell.run "RegSvr32 outlook_export.dll /s"
    '    'if not try
    '    '' oShell.run """RegSvr32 outlook_export.dll /s"""

    '    'wshell.run "cscript.exe """ & wscript.scriptfullname & """"

    '    '''since I moved to Windows 7, GuWu refuses to register OfficeViewer.ocx. Must do so manually:
    '    On Error Resume Next
    '    Dim strM As String
    '    oShell("regsvr32.exe C:\Program Files\LABIntegrity\StudyDoc\officeviewer.ocx")
    '    If Err.Number = 0 Then
    '        MsgBox("LABIntegrity 32bit OfficeViewer.ocx registration successful")
    '    Else 'try 64bit
    '        MsgBox("LABIntegrity 32bit Err registering OfficeViewer: " & Err.Number & ": " & Err.Description)
    '        Err.Clear()
    '        On Error Resume Next
    '        oShell("regsvr32.exe C:\Program Files (x86)\LABIntegrity\StudyDoc\officeviewer.ocx")
    '        If Err.Number = 0 Then
    '            MsgBox("LABIntegrity 64bit OfficeViewer.ocx registration successful")
    '        Else
    '            MsgBox("LABIntegrity 64bit Err registering OfficeViewer: " & Err.Number & ": " & Err.Description)
    '        End If
    '    End If


    'End Sub

End Class
