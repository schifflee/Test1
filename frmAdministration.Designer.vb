<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdministration
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAdministration))
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lbxTab1 = New System.Windows.Forms.ListBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.pb1 = New System.Windows.Forms.ProgressBar()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lblcbxModules = New System.Windows.Forms.Label()
        Me.cbxModules = New System.Windows.Forms.ComboBox()
        Me.pan1 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.gbUA = New System.Windows.Forms.GroupBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.gbWatsonAccount = New System.Windows.Forms.GroupBox()
        Me.cmdClearWatson = New System.Windows.Forms.Button()
        Me.ID_TBLWATSONACCOUNT = New System.Windows.Forms.TextBox()
        Me.cbxWatsonAccount = New System.Windows.Forms.ComboBox()
        Me.gbWindowsAuth = New System.Windows.Forms.GroupBox()
        Me.lblNetworkAccount = New System.Windows.Forms.Label()
        Me.cmdTestAccount = New System.Windows.Forms.Button()
        Me.gbLDAP = New System.Windows.Forms.GroupBox()
        Me.rbADVAPI32 = New System.Windows.Forms.RadioButton()
        Me.rbLDAPNon = New System.Windows.Forms.RadioButton()
        Me.rbLDAP = New System.Windows.Forms.RadioButton()
        Me.panLDAP = New System.Windows.Forms.Panel()
        Me.lblLDAPaaa = New System.Windows.Forms.Label()
        Me.CHARLDAP = New System.Windows.Forms.TextBox()
        Me.lblLDAP = New System.Windows.Forms.Label()
        Me.lblLDAPeg = New System.Windows.Forms.Label()
        Me.cmdCopyLDAP = New System.Windows.Forms.Button()
        Me.lblLDAPClear = New System.Windows.Forms.Label()
        Me.cmdClearNet = New System.Windows.Forms.Button()
        Me.cmdGetUserName = New System.Windows.Forms.Button()
        Me.CHARNETWORKACCOUNT = New System.Windows.Forms.TextBox()
        Me.gbSetPerm = New System.Windows.Forms.GroupBox()
        Me.cbxPermissionsGroup = New System.Windows.Forms.ComboBox()
        Me.gbxPassword = New System.Windows.Forms.GroupBox()
        Me.chkAccountIsLockedOut = New System.Windows.Forms.CheckBox()
        Me.chkPasswordNeverExpires = New System.Windows.Forms.CheckBox()
        Me.chkUserCannotChangePassword = New System.Windows.Forms.CheckBox()
        Me.chkChangePasswordAtNextLogon = New System.Windows.Forms.CheckBox()
        Me.cmdEnterPassword = New System.Windows.Forms.Button()
        Me.cmdResetUserAccounts = New System.Windows.Forms.Button()
        Me.cmdAddUserID = New System.Windows.Forms.Button()
        Me.cmdAddUser = New System.Windows.Forms.Button()
        Me.gbGlobalParams = New System.Windows.Forms.GroupBox()
        Me.rbShowInactiveUserIDs = New System.Windows.Forms.RadioButton()
        Me.rbShowActiveUserIDs = New System.Windows.Forms.RadioButton()
        Me.rbShowAllUserIDs = New System.Windows.Forms.RadioButton()
        Me.gbUserShow = New System.Windows.Forms.GroupBox()
        Me.rbShowInactiveUsers = New System.Windows.Forms.RadioButton()
        Me.rbShowActiveUsers = New System.Windows.Forms.RadioButton()
        Me.rbShowAllUsers = New System.Windows.Forms.RadioButton()
        Me.dgvUserAttributes = New System.Windows.Forms.DataGridView()
        Me.dgvUsers = New System.Windows.Forms.DataGridView()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmdSelectAllPermissions = New System.Windows.Forms.Button()
        Me.cmdDeselectAllPermissions = New System.Windows.Forms.Button()
        Me.lblcbxModulesPers = New System.Windows.Forms.Label()
        Me.lvPermissionsAdmin = New System.Windows.Forms.ListView()
        Me.lvPermissions = New System.Windows.Forms.ListView()
        Me.pan2 = New System.Windows.Forms.Panel()
        Me.dgvDropdownboxTitle = New System.Windows.Forms.DataGridView()
        Me.cmdOrderDropdownbox = New System.Windows.Forms.Button()
        Me.dgvDropdownboxContents = New System.Windows.Forms.DataGridView()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmdAddDropdownbox = New System.Windows.Forms.Button()
        Me.cmdResetDropdownbox = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pan3 = New System.Windows.Forms.Panel()
        Me.gbxlblCorporateAdderesses = New System.Windows.Forms.GroupBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.gbDropDown = New System.Windows.Forms.GroupBox()
        Me.rbShowInactiveAddresses = New System.Windows.Forms.RadioButton()
        Me.rbShowActiveAddresses = New System.Windows.Forms.RadioButton()
        Me.rbShowAllAddresses = New System.Windows.Forms.RadioButton()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cmdAddCorporateAddress = New System.Windows.Forms.Button()
        Me.dgvCorporateAddresses = New System.Windows.Forms.DataGridView()
        Me.dgvNickNames = New System.Windows.Forms.DataGridView()
        Me.cmdResetCorporateAddressses = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pan4 = New System.Windows.Forms.Panel()
        Me.gbSTD = New System.Windows.Forms.GroupBox()
        Me.lblTExpl = New System.Windows.Forms.Label()
        Me.dgvTemplates = New System.Windows.Forms.DataGridView()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cmdAddTemplate = New System.Windows.Forms.Button()
        Me.gbStudyTemplates = New System.Windows.Forms.GroupBox()
        Me.rbShowInactiveTemplates = New System.Windows.Forms.RadioButton()
        Me.rbShowActiveTemplates = New System.Windows.Forms.RadioButton()
        Me.rbShowAllTemplates = New System.Windows.Forms.RadioButton()
        Me.dgvTemplateAttributes = New System.Windows.Forms.DataGridView()
        Me.cmdResetDefineReports = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.pan5 = New System.Windows.Forms.Panel()
        Me.gbxlabelGlobalParameters = New System.Windows.Forms.GroupBox()
        Me.lblIntegrity = New System.Windows.Forms.Label()
        Me.panGP = New System.Windows.Forms.Panel()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.dgvGlobal = New System.Windows.Forms.DataGridView()
        Me.lblGlobalValues = New System.Windows.Forms.Label()
        Me.lbxGlobal = New System.Windows.Forms.ListBox()
        Me.cmdResetGlobal = New System.Windows.Forms.Button()
        Me.cmdBrowseGlobal = New System.Windows.Forms.Button()
        Me.lblGlobalParameters = New System.Windows.Forms.Label()
        Me.pan6 = New System.Windows.Forms.Panel()
        Me.cmdRefreshHook = New System.Windows.Forms.Button()
        Me.cmdResetHooks = New System.Windows.Forms.Button()
        Me.dgvHooks = New System.Windows.Forms.DataGridView()
        Me.cmdAddHook = New System.Windows.Forms.Button()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.lblRestricted = New System.Windows.Forms.Label()
        Me.chkEditMode = New System.Windows.Forms.CheckBox()
        Me.cmdIncreaseFont = New System.Windows.Forms.Button()
        Me.cmdSymbol = New System.Windows.Forms.Button()
        Me.pan7 = New System.Windows.Forms.Panel()
        Me.cmdTestESig = New System.Windows.Forms.Button()
        Me.cmdRemoveRFC = New System.Windows.Forms.Button()
        Me.cmdRemoveMOS = New System.Windows.Forms.Button()
        Me.cmdAddRFC = New System.Windows.Forms.Button()
        Me.cmdAddMOS = New System.Windows.Forms.Button()
        Me.lblReasonForChange = New System.Windows.Forms.Label()
        Me.lblMeaningOfSig = New System.Windows.Forms.Label()
        Me.dgvRFC = New System.Windows.Forms.DataGridView()
        Me.dgvMOS = New System.Windows.Forms.DataGridView()
        Me.gbAuditTrail = New System.Windows.Forms.GroupBox()
        Me.gbReasonForChange = New System.Windows.Forms.GroupBox()
        Me.panRFCOptions = New System.Windows.Forms.Panel()
        Me.chkReasonFreeForm = New System.Windows.Forms.CheckBox()
        Me.chkReasonForChange = New System.Windows.Forms.CheckBox()
        Me.rbAuditTrailOff = New System.Windows.Forms.RadioButton()
        Me.gbESig = New System.Windows.Forms.GroupBox()
        Me.panESigOptions = New System.Windows.Forms.Panel()
        Me.chkSigFreeForm = New System.Windows.Forms.CheckBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.gbUserIDType = New System.Windows.Forms.GroupBox()
        Me.rbUserIDChoice = New System.Windows.Forms.RadioButton()
        Me.rbOnlyLoggedOn = New System.Windows.Forms.RadioButton()
        Me.chkMeaningOfSign = New System.Windows.Forms.CheckBox()
        Me.rbESigOff = New System.Windows.Forms.RadioButton()
        Me.rbESigOn = New System.Windows.Forms.RadioButton()
        Me.rbAuditTrailOn = New System.Windows.Forms.RadioButton()
        Me.lblCC = New System.Windows.Forms.Label()
        Me.lblOpen = New System.Windows.Forms.Label()
        Me.panFC = New System.Windows.Forms.Panel()
        Me.lblFieldCodes = New System.Windows.Forms.Label()
        Me.panFCcmd = New System.Windows.Forms.Panel()
        Me.cmdRemoveFC = New System.Windows.Forms.Button()
        Me.cmdResetFC = New System.Windows.Forms.Button()
        Me.cmdAddFC = New System.Windows.Forms.Button()
        Me.dgvFC = New System.Windows.Forms.DataGridView()
        Me.cmdDecreaseFont = New System.Windows.Forms.Button()
        Me.pan7b = New System.Windows.Forms.Panel()
        Me.pan7a = New System.Windows.Forms.Panel()
        Me.cmdRemovePM = New System.Windows.Forms.Button()
        Me.cmdAddPM = New System.Windows.Forms.Button()
        Me.pan8 = New System.Windows.Forms.Panel()
        Me.lblPM = New System.Windows.Forms.Label()
        Me.panPM = New System.Windows.Forms.Panel()
        Me.lvPermissionsFinalReport = New System.Windows.Forms.ListView()
        Me.lvPermissionsReportTemplate = New System.Windows.Forms.ListView()
        Me.dgvPermissions = New System.Windows.Forms.DataGridView()
        Me.lblDo = New System.Windows.Forms.Label()
        Me.lblBase = New System.Windows.Forms.Label()
        Me.lblPermissions = New System.Windows.Forms.Label()
        Me.lbllbx1 = New System.Windows.Forms.Label()
        Me.lblS = New System.Windows.Forms.Label()
        Me.cbxPermBase = New System.Windows.Forms.ComboBox()
        Me.lbx1 = New System.Windows.Forms.ListBox()
        Me.pan1.SuspendLayout()
        Me.gbUA.SuspendLayout()
        Me.gbWatsonAccount.SuspendLayout()
        Me.gbWindowsAuth.SuspendLayout()
        Me.gbLDAP.SuspendLayout()
        Me.panLDAP.SuspendLayout()
        Me.gbSetPerm.SuspendLayout()
        Me.gbxPassword.SuspendLayout()
        Me.gbGlobalParams.SuspendLayout()
        Me.gbUserShow.SuspendLayout()
        CType(Me.dgvUserAttributes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvUsers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan2.SuspendLayout()
        CType(Me.dgvDropdownboxTitle, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDropdownboxContents, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan3.SuspendLayout()
        Me.gbxlblCorporateAdderesses.SuspendLayout()
        Me.gbDropDown.SuspendLayout()
        CType(Me.dgvCorporateAddresses, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvNickNames, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan4.SuspendLayout()
        Me.gbSTD.SuspendLayout()
        CType(Me.dgvTemplates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbStudyTemplates.SuspendLayout()
        CType(Me.dgvTemplateAttributes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan5.SuspendLayout()
        Me.gbxlabelGlobalParameters.SuspendLayout()
        Me.panGP.SuspendLayout()
        CType(Me.dgvGlobal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan6.SuspendLayout()
        CType(Me.dgvHooks, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan7.SuspendLayout()
        CType(Me.dgvRFC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvMOS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbAuditTrail.SuspendLayout()
        Me.gbReasonForChange.SuspendLayout()
        Me.panRFCOptions.SuspendLayout()
        Me.gbESig.SuspendLayout()
        Me.panESigOptions.SuspendLayout()
        Me.gbUserIDType.SuspendLayout()
        Me.panFC.SuspendLayout()
        Me.panFCcmd.SuspendLayout()
        CType(Me.dgvFC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pan7b.SuspendLayout()
        Me.pan7a.SuspendLayout()
        Me.pan8.SuspendLayout()
        Me.panPM.SuspendLayout()
        CType(Me.dgvPermissions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.Color.Gray
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdSave.ForeColor = System.Drawing.Color.ForestGreen
        Me.cmdSave.Location = New System.Drawing.Point(898, 8)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(79, 33)
        Me.cmdSave.TabIndex = 91
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdEdit
        '
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdEdit.CausesValidation = False
        Me.cmdEdit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdEdit.ForeColor = System.Drawing.Color.Blue
        Me.cmdEdit.Location = New System.Drawing.Point(817, 8)
        Me.cmdEdit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(79, 33)
        Me.cmdEdit.TabIndex = 89
        Me.cmdEdit.Text = "&Edit"
        Me.cmdEdit.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.CausesValidation = False
        Me.cmdExit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdExit.ForeColor = System.Drawing.Color.Red
        Me.cmdExit.Location = New System.Drawing.Point(1062, 8)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(79, 33)
        Me.cmdExit.TabIndex = 88
        Me.cmdExit.Text = "G&o Back"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(2, 25)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(195, 20)
        Me.Label4.TabIndex = 87
        Me.Label4.Text = "Table of Contents"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lbxTab1
        '
        Me.lbxTab1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbxTab1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbxTab1.ItemHeight = 17
        Me.lbxTab1.Location = New System.Drawing.Point(2, 51)
        Me.lbxTab1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lbxTab1.Name = "lbxTab1"
        Me.lbxTab1.Size = New System.Drawing.Size(194, 818)
        Me.lbxTab1.TabIndex = 86
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gray
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(980, 8)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(79, 33)
        Me.cmdCancel.TabIndex = 90
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'pb1
        '
        Me.pb1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.pb1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pb1.Location = New System.Drawing.Point(66, 3)
        Me.pb1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(365, 34)
        Me.pb1.TabIndex = 102
        Me.pb1.Visible = False
        '
        'lblProgress
        '
        Me.lblProgress.BackColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle))
        Me.lblProgress.ForeColor = System.Drawing.Color.White
        Me.lblProgress.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblProgress.Location = New System.Drawing.Point(890, 65)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(251, 52)
        Me.lblProgress.TabIndex = 103
        Me.lblProgress.Text = "Label35"
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblProgress.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(2, 3)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(115, 47)
        Me.Button1.TabIndex = 105
        Me.Button1.Text = "Test"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'lblcbxModules
        '
        Me.lblcbxModules.AutoSize = True
        Me.lblcbxModules.BackColor = System.Drawing.Color.Transparent
        Me.lblcbxModules.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.lblcbxModules.ForeColor = System.Drawing.Color.Blue
        Me.lblcbxModules.Location = New System.Drawing.Point(426, 12)
        Me.lblcbxModules.Name = "lblcbxModules"
        Me.lblcbxModules.Size = New System.Drawing.Size(119, 17)
        Me.lblcbxModules.TabIndex = 118
        Me.lblcbxModules.Text = "Choose a Module:"
        '
        'cbxModules
        '
        Me.cbxModules.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxModules.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxModules.FormattingEnabled = True
        Me.cbxModules.Location = New System.Drawing.Point(547, 8)
        Me.cbxModules.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxModules.Name = "cbxModules"
        Me.cbxModules.Size = New System.Drawing.Size(213, 29)
        Me.cbxModules.TabIndex = 117
        '
        'pan1
        '
        Me.pan1.AutoScroll = True
        Me.pan1.BackColor = System.Drawing.Color.White
        Me.pan1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan1.Controls.Add(Me.Label3)
        Me.pan1.Controls.Add(Me.gbUA)
        Me.pan1.Controls.Add(Me.gbWatsonAccount)
        Me.pan1.Controls.Add(Me.gbWindowsAuth)
        Me.pan1.Controls.Add(Me.gbSetPerm)
        Me.pan1.Controls.Add(Me.gbxPassword)
        Me.pan1.Controls.Add(Me.cmdEnterPassword)
        Me.pan1.Controls.Add(Me.cmdResetUserAccounts)
        Me.pan1.Controls.Add(Me.cmdAddUserID)
        Me.pan1.Controls.Add(Me.cmdAddUser)
        Me.pan1.Controls.Add(Me.gbGlobalParams)
        Me.pan1.Controls.Add(Me.gbUserShow)
        Me.pan1.Controls.Add(Me.dgvUserAttributes)
        Me.pan1.Controls.Add(Me.dgvUsers)
        Me.pan1.Controls.Add(Me.Label7)
        Me.pan1.Controls.Add(Me.Label6)
        Me.pan1.Location = New System.Drawing.Point(215, 619)
        Me.pan1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan1.Name = "pan1"
        Me.pan1.Size = New System.Drawing.Size(762, 297)
        Me.pan1.TabIndex = 120
        Me.pan1.Visible = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(0, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(917, 21)
        Me.Label3.TabIndex = 120
        Me.Label3.Text = "User Accounts"
        '
        'gbUA
        '
        Me.gbUA.AutoSize = True
        Me.gbUA.Controls.Add(Me.Label18)
        Me.gbUA.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbUA.Location = New System.Drawing.Point(374, 21)
        Me.gbUA.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbUA.Name = "gbUA"
        Me.gbUA.Padding = New System.Windows.Forms.Padding(0)
        Me.gbUA.Size = New System.Drawing.Size(256, 29)
        Me.gbUA.TabIndex = 130
        Me.gbUA.TabStop = False
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label18.Location = New System.Drawing.Point(3, 7)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(242, 17)
        Me.Label18.TabIndex = 137
        Me.Label18.Text = "* A:  Select = Active, De-select = Inactive"
        '
        'gbWatsonAccount
        '
        Me.gbWatsonAccount.BackColor = System.Drawing.Color.Transparent
        Me.gbWatsonAccount.Controls.Add(Me.cmdClearWatson)
        Me.gbWatsonAccount.Controls.Add(Me.ID_TBLWATSONACCOUNT)
        Me.gbWatsonAccount.Controls.Add(Me.cbxWatsonAccount)
        Me.gbWatsonAccount.Enabled = False
        Me.gbWatsonAccount.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbWatsonAccount.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.gbWatsonAccount.Location = New System.Drawing.Point(374, 478)
        Me.gbWatsonAccount.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbWatsonAccount.Name = "gbWatsonAccount"
        Me.gbWatsonAccount.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbWatsonAccount.Size = New System.Drawing.Size(542, 66)
        Me.gbWatsonAccount.TabIndex = 141
        Me.gbWatsonAccount.TabStop = False
        Me.gbWatsonAccount.Text = "Configure Watson Account (Optional)"
        '
        'cmdClearWatson
        '
        Me.cmdClearWatson.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdClearWatson.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdClearWatson.ForeColor = System.Drawing.Color.Chocolate
        Me.cmdClearWatson.Location = New System.Drawing.Point(308, 21)
        Me.cmdClearWatson.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdClearWatson.Name = "cmdClearWatson"
        Me.cmdClearWatson.Size = New System.Drawing.Size(140, 33)
        Me.cmdClearWatson.TabIndex = 130
        Me.cmdClearWatson.Text = "Clear Selection..."
        Me.cmdClearWatson.UseVisualStyleBackColor = False
        '
        'ID_TBLWATSONACCOUNT
        '
        Me.ID_TBLWATSONACCOUNT.Location = New System.Drawing.Point(388, 25)
        Me.ID_TBLWATSONACCOUNT.Name = "ID_TBLWATSONACCOUNT"
        Me.ID_TBLWATSONACCOUNT.Size = New System.Drawing.Size(148, 25)
        Me.ID_TBLWATSONACCOUNT.TabIndex = 3
        Me.ID_TBLWATSONACCOUNT.Visible = False
        '
        'cbxWatsonAccount
        '
        Me.cbxWatsonAccount.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cbxWatsonAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxWatsonAccount.FormattingEnabled = True
        Me.cbxWatsonAccount.IntegralHeight = False
        Me.cbxWatsonAccount.Location = New System.Drawing.Point(14, 25)
        Me.cbxWatsonAccount.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxWatsonAccount.Name = "cbxWatsonAccount"
        Me.cbxWatsonAccount.Size = New System.Drawing.Size(280, 25)
        Me.cbxWatsonAccount.TabIndex = 0
        '
        'gbWindowsAuth
        '
        Me.gbWindowsAuth.BackColor = System.Drawing.Color.Transparent
        Me.gbWindowsAuth.Controls.Add(Me.lblNetworkAccount)
        Me.gbWindowsAuth.Controls.Add(Me.cmdTestAccount)
        Me.gbWindowsAuth.Controls.Add(Me.gbLDAP)
        Me.gbWindowsAuth.Controls.Add(Me.panLDAP)
        Me.gbWindowsAuth.Controls.Add(Me.lblLDAPClear)
        Me.gbWindowsAuth.Controls.Add(Me.cmdClearNet)
        Me.gbWindowsAuth.Controls.Add(Me.cmdGetUserName)
        Me.gbWindowsAuth.Controls.Add(Me.CHARNETWORKACCOUNT)
        Me.gbWindowsAuth.Enabled = False
        Me.gbWindowsAuth.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbWindowsAuth.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.gbWindowsAuth.Location = New System.Drawing.Point(374, 559)
        Me.gbWindowsAuth.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbWindowsAuth.Name = "gbWindowsAuth"
        Me.gbWindowsAuth.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbWindowsAuth.Size = New System.Drawing.Size(542, 241)
        Me.gbWindowsAuth.TabIndex = 140
        Me.gbWindowsAuth.TabStop = False
        Me.gbWindowsAuth.Text = "Configure Windows Authentication Network Account (Optional)"
        '
        'lblNetworkAccount
        '
        Me.lblNetworkAccount.AutoSize = True
        Me.lblNetworkAccount.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNetworkAccount.Location = New System.Drawing.Point(14, 194)
        Me.lblNetworkAccount.Name = "lblNetworkAccount"
        Me.lblNetworkAccount.Size = New System.Drawing.Size(100, 15)
        Me.lblNetworkAccount.TabIndex = 148
        Me.lblNetworkAccount.Text = "Network Account"
        '
        'cmdTestAccount
        '
        Me.cmdTestAccount.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdTestAccount.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTestAccount.ForeColor = System.Drawing.Color.Chocolate
        Me.cmdTestAccount.Location = New System.Drawing.Point(198, 156)
        Me.cmdTestAccount.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdTestAccount.Name = "cmdTestAccount"
        Me.cmdTestAccount.Size = New System.Drawing.Size(173, 33)
        Me.cmdTestAccount.TabIndex = 147
        Me.cmdTestAccount.Text = "&Test Network Account..."
        Me.cmdTestAccount.UseVisualStyleBackColor = False
        '
        'gbLDAP
        '
        Me.gbLDAP.Controls.Add(Me.rbADVAPI32)
        Me.gbLDAP.Controls.Add(Me.rbLDAPNon)
        Me.gbLDAP.Controls.Add(Me.rbLDAP)
        Me.gbLDAP.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbLDAP.ForeColor = System.Drawing.Color.Chocolate
        Me.gbLDAP.Location = New System.Drawing.Point(295, 20)
        Me.gbLDAP.Name = "gbLDAP"
        Me.gbLDAP.Size = New System.Drawing.Size(241, 45)
        Me.gbLDAP.TabIndex = 135
        Me.gbLDAP.TabStop = False
        Me.gbLDAP.Text = "Global Windows Authentication Type"
        '
        'rbADVAPI32
        '
        Me.rbADVAPI32.AutoSize = True
        Me.rbADVAPI32.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbADVAPI32.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbADVAPI32.Location = New System.Drawing.Point(157, 17)
        Me.rbADVAPI32.Name = "rbADVAPI32"
        Me.rbADVAPI32.Size = New System.Drawing.Size(78, 19)
        Me.rbADVAPI32.TabIndex = 2
        Me.rbADVAPI32.Text = "ADVAPI32"
        Me.rbADVAPI32.UseVisualStyleBackColor = True
        '
        'rbLDAPNon
        '
        Me.rbLDAPNon.AutoSize = True
        Me.rbLDAPNon.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbLDAPNon.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbLDAPNon.Location = New System.Drawing.Point(67, 17)
        Me.rbLDAPNon.Name = "rbLDAPNon"
        Me.rbLDAPNon.Size = New System.Drawing.Size(82, 19)
        Me.rbLDAPNon.TabIndex = 1
        Me.rbLDAPNon.Text = "Non-LDAP"
        Me.rbLDAPNon.UseVisualStyleBackColor = True
        '
        'rbLDAP
        '
        Me.rbLDAP.AutoSize = True
        Me.rbLDAP.Checked = True
        Me.rbLDAP.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbLDAP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbLDAP.Location = New System.Drawing.Point(5, 17)
        Me.rbLDAP.Name = "rbLDAP"
        Me.rbLDAP.Size = New System.Drawing.Size(54, 19)
        Me.rbLDAP.TabIndex = 0
        Me.rbLDAP.TabStop = True
        Me.rbLDAP.Text = "LDAP"
        Me.rbLDAP.UseVisualStyleBackColor = True
        '
        'panLDAP
        '
        Me.panLDAP.Controls.Add(Me.lblLDAPaaa)
        Me.panLDAP.Controls.Add(Me.CHARLDAP)
        Me.panLDAP.Controls.Add(Me.lblLDAP)
        Me.panLDAP.Controls.Add(Me.lblLDAPeg)
        Me.panLDAP.Controls.Add(Me.cmdCopyLDAP)
        Me.panLDAP.Location = New System.Drawing.Point(14, 68)
        Me.panLDAP.Name = "panLDAP"
        Me.panLDAP.Size = New System.Drawing.Size(516, 87)
        Me.panLDAP.TabIndex = 146
        '
        'lblLDAPaaa
        '
        Me.lblLDAPaaa.AutoSize = True
        Me.lblLDAPaaa.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLDAPaaa.Location = New System.Drawing.Point(28, 39)
        Me.lblLDAPaaa.Name = "lblLDAPaaa"
        Me.lblLDAPaaa.Size = New System.Drawing.Size(439, 17)
        Me.lblLDAPaaa.TabIndex = 132
        Me.lblLDAPaaa.Text = "Use ""nltest /dclist:[Domain]"" at a cmd prompt to generate LDAP server list"
        '
        'CHARLDAP
        '
        Me.CHARLDAP.Location = New System.Drawing.Point(0, 58)
        Me.CHARLDAP.Name = "CHARLDAP"
        Me.CHARLDAP.Size = New System.Drawing.Size(509, 25)
        Me.CHARLDAP.TabIndex = 5
        '
        'lblLDAP
        '
        Me.lblLDAP.AutoSize = True
        Me.lblLDAP.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLDAP.Location = New System.Drawing.Point(0, 0)
        Me.lblLDAP.Name = "lblLDAP"
        Me.lblLDAP.Size = New System.Drawing.Size(228, 17)
        Me.lblLDAP.TabIndex = 6
        Me.lblLDAP.Text = "Enter the LDAP Server address below:"
        '
        'lblLDAPeg
        '
        Me.lblLDAPeg.AutoSize = True
        Me.lblLDAPeg.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLDAPeg.Location = New System.Drawing.Point(28, 19)
        Me.lblLDAPeg.Name = "lblLDAPeg"
        Me.lblLDAPeg.Size = New System.Drawing.Size(135, 17)
        Me.lblLDAPeg.TabIndex = 131
        Me.lblLDAPeg.Text = "Example:  LI.LIinc.local"
        '
        'cmdCopyLDAP
        '
        Me.cmdCopyLDAP.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCopyLDAP.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdCopyLDAP.ForeColor = System.Drawing.Color.Chocolate
        Me.cmdCopyLDAP.Location = New System.Drawing.Point(281, 0)
        Me.cmdCopyLDAP.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCopyLDAP.Name = "cmdCopyLDAP"
        Me.cmdCopyLDAP.Size = New System.Drawing.Size(228, 33)
        Me.cmdCopyLDAP.TabIndex = 129
        Me.cmdCopyLDAP.Text = "Copy &LDAP from existing..."
        Me.cmdCopyLDAP.UseVisualStyleBackColor = False
        '
        'lblLDAPClear
        '
        Me.lblLDAPClear.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLDAPClear.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblLDAPClear.Location = New System.Drawing.Point(14, 23)
        Me.lblLDAPClear.Name = "lblLDAPClear"
        Me.lblLDAPClear.Size = New System.Drawing.Size(298, 39)
        Me.lblLDAPClear.TabIndex = 134
        Me.lblLDAPClear.Text = "To disable Windows Authentication for this" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "account, click the 'Clear Account...'" & _
    " button"
        Me.lblLDAPClear.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdClearNet
        '
        Me.cmdClearNet.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdClearNet.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdClearNet.ForeColor = System.Drawing.Color.Chocolate
        Me.cmdClearNet.Location = New System.Drawing.Point(382, 156)
        Me.cmdClearNet.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdClearNet.Name = "cmdClearNet"
        Me.cmdClearNet.Size = New System.Drawing.Size(140, 33)
        Me.cmdClearNet.TabIndex = 133
        Me.cmdClearNet.Text = "Clear Account..."
        Me.cmdClearNet.UseVisualStyleBackColor = False
        '
        'cmdGetUserName
        '
        Me.cmdGetUserName.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdGetUserName.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetUserName.ForeColor = System.Drawing.Color.Chocolate
        Me.cmdGetUserName.Location = New System.Drawing.Point(14, 156)
        Me.cmdGetUserName.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdGetUserName.Name = "cmdGetUserName"
        Me.cmdGetUserName.Size = New System.Drawing.Size(173, 33)
        Me.cmdGetUserName.TabIndex = 132
        Me.cmdGetUserName.Text = "&Get Network Account..."
        Me.cmdGetUserName.UseVisualStyleBackColor = False
        '
        'CHARNETWORKACCOUNT
        '
        Me.CHARNETWORKACCOUNT.Location = New System.Drawing.Point(14, 211)
        Me.CHARNETWORKACCOUNT.Name = "CHARNETWORKACCOUNT"
        Me.CHARNETWORKACCOUNT.Size = New System.Drawing.Size(509, 25)
        Me.CHARNETWORKACCOUNT.TabIndex = 4
        '
        'gbSetPerm
        '
        Me.gbSetPerm.BackColor = System.Drawing.Color.Transparent
        Me.gbSetPerm.Controls.Add(Me.cbxPermissionsGroup)
        Me.gbSetPerm.Enabled = False
        Me.gbSetPerm.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSetPerm.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.gbSetPerm.Location = New System.Drawing.Point(374, 408)
        Me.gbSetPerm.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSetPerm.Name = "gbSetPerm"
        Me.gbSetPerm.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSetPerm.Size = New System.Drawing.Size(542, 58)
        Me.gbSetPerm.TabIndex = 139
        Me.gbSetPerm.TabStop = False
        Me.gbSetPerm.Text = "Configure User ID Permissions Group"
        '
        'cbxPermissionsGroup
        '
        Me.cbxPermissionsGroup.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cbxPermissionsGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxPermissionsGroup.FormattingEnabled = True
        Me.cbxPermissionsGroup.IntegralHeight = False
        Me.cbxPermissionsGroup.Location = New System.Drawing.Point(14, 23)
        Me.cbxPermissionsGroup.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxPermissionsGroup.Name = "cbxPermissionsGroup"
        Me.cbxPermissionsGroup.Size = New System.Drawing.Size(280, 25)
        Me.cbxPermissionsGroup.TabIndex = 0
        '
        'gbxPassword
        '
        Me.gbxPassword.BackColor = System.Drawing.Color.Transparent
        Me.gbxPassword.Controls.Add(Me.chkAccountIsLockedOut)
        Me.gbxPassword.Controls.Add(Me.chkPasswordNeverExpires)
        Me.gbxPassword.Controls.Add(Me.chkUserCannotChangePassword)
        Me.gbxPassword.Controls.Add(Me.chkChangePasswordAtNextLogon)
        Me.gbxPassword.Enabled = False
        Me.gbxPassword.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxPassword.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.gbxPassword.Location = New System.Drawing.Point(374, 327)
        Me.gbxPassword.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxPassword.Name = "gbxPassword"
        Me.gbxPassword.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxPassword.Size = New System.Drawing.Size(542, 70)
        Me.gbxPassword.TabIndex = 136
        Me.gbxPassword.TabStop = False
        Me.gbxPassword.Text = "Configure User ID Password Attributes"
        '
        'chkAccountIsLockedOut
        '
        Me.chkAccountIsLockedOut.AutoSize = True
        Me.chkAccountIsLockedOut.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAccountIsLockedOut.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAccountIsLockedOut.Location = New System.Drawing.Point(322, 45)
        Me.chkAccountIsLockedOut.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkAccountIsLockedOut.Name = "chkAccountIsLockedOut"
        Me.chkAccountIsLockedOut.Size = New System.Drawing.Size(125, 17)
        Me.chkAccountIsLockedOut.TabIndex = 3
        Me.chkAccountIsLockedOut.Text = "User ID is locked out"
        Me.chkAccountIsLockedOut.UseVisualStyleBackColor = True
        '
        'chkPasswordNeverExpires
        '
        Me.chkPasswordNeverExpires.AutoSize = True
        Me.chkPasswordNeverExpires.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPasswordNeverExpires.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPasswordNeverExpires.Location = New System.Drawing.Point(322, 24)
        Me.chkPasswordNeverExpires.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkPasswordNeverExpires.Name = "chkPasswordNeverExpires"
        Me.chkPasswordNeverExpires.Size = New System.Drawing.Size(138, 17)
        Me.chkPasswordNeverExpires.TabIndex = 2
        Me.chkPasswordNeverExpires.Text = "Password never expires"
        Me.chkPasswordNeverExpires.UseVisualStyleBackColor = True
        '
        'chkUserCannotChangePassword
        '
        Me.chkUserCannotChangePassword.AutoSize = True
        Me.chkUserCannotChangePassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUserCannotChangePassword.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkUserCannotChangePassword.Location = New System.Drawing.Point(16, 45)
        Me.chkUserCannotChangePassword.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkUserCannotChangePassword.Name = "chkUserCannotChangePassword"
        Me.chkUserCannotChangePassword.Size = New System.Drawing.Size(171, 17)
        Me.chkUserCannotChangePassword.TabIndex = 1
        Me.chkUserCannotChangePassword.Text = "User cannot change password"
        Me.chkUserCannotChangePassword.UseVisualStyleBackColor = True
        '
        'chkChangePasswordAtNextLogon
        '
        Me.chkChangePasswordAtNextLogon.AutoSize = True
        Me.chkChangePasswordAtNextLogon.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkChangePasswordAtNextLogon.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkChangePasswordAtNextLogon.Location = New System.Drawing.Point(16, 24)
        Me.chkChangePasswordAtNextLogon.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkChangePasswordAtNextLogon.Name = "chkChangePasswordAtNextLogon"
        Me.chkChangePasswordAtNextLogon.Size = New System.Drawing.Size(224, 17)
        Me.chkChangePasswordAtNextLogon.TabIndex = 0
        Me.chkChangePasswordAtNextLogon.Text = "User must change password at next logon"
        Me.chkChangePasswordAtNextLogon.UseVisualStyleBackColor = True
        '
        'cmdEnterPassword
        '
        Me.cmdEnterPassword.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cmdEnterPassword.CausesValidation = False
        Me.cmdEnterPassword.Enabled = False
        Me.cmdEnterPassword.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEnterPassword.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdEnterPassword.Location = New System.Drawing.Point(374, 92)
        Me.cmdEnterPassword.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdEnterPassword.Name = "cmdEnterPassword"
        Me.cmdEnterPassword.Size = New System.Drawing.Size(176, 33)
        Me.cmdEnterPassword.TabIndex = 132
        Me.cmdEnterPassword.Text = "Change &Pswd"
        Me.cmdEnterPassword.UseVisualStyleBackColor = False
        '
        'cmdResetUserAccounts
        '
        Me.cmdResetUserAccounts.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdResetUserAccounts.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdResetUserAccounts.CausesValidation = False
        Me.cmdResetUserAccounts.Enabled = False
        Me.cmdResetUserAccounts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetUserAccounts.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdResetUserAccounts.Location = New System.Drawing.Point(2224, 23)
        Me.cmdResetUserAccounts.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdResetUserAccounts.Name = "cmdResetUserAccounts"
        Me.cmdResetUserAccounts.Size = New System.Drawing.Size(82, 33)
        Me.cmdResetUserAccounts.TabIndex = 131
        Me.cmdResetUserAccounts.Text = "Reset"
        Me.cmdResetUserAccounts.UseVisualStyleBackColor = False
        '
        'cmdAddUserID
        '
        Me.cmdAddUserID.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddUserID.Enabled = False
        Me.cmdAddUserID.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddUserID.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddUserID.Location = New System.Drawing.Point(374, 56)
        Me.cmdAddUserID.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddUserID.Name = "cmdAddUserID"
        Me.cmdAddUserID.Size = New System.Drawing.Size(176, 33)
        Me.cmdAddUserID.TabIndex = 130
        Me.cmdAddUserID.Text = "&Add New UserID"
        Me.cmdAddUserID.UseVisualStyleBackColor = True
        '
        'cmdAddUser
        '
        Me.cmdAddUser.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddUser.Enabled = False
        Me.cmdAddUser.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddUser.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddUser.Location = New System.Drawing.Point(9, 31)
        Me.cmdAddUser.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddUser.Name = "cmdAddUser"
        Me.cmdAddUser.Size = New System.Drawing.Size(174, 33)
        Me.cmdAddUser.TabIndex = 127
        Me.cmdAddUser.Text = "&Add New User"
        Me.cmdAddUser.UseVisualStyleBackColor = True
        '
        'gbGlobalParams
        '
        Me.gbGlobalParams.BackColor = System.Drawing.Color.Transparent
        Me.gbGlobalParams.Controls.Add(Me.rbShowInactiveUserIDs)
        Me.gbGlobalParams.Controls.Add(Me.rbShowActiveUserIDs)
        Me.gbGlobalParams.Controls.Add(Me.rbShowAllUserIDs)
        Me.gbGlobalParams.Location = New System.Drawing.Point(579, 82)
        Me.gbGlobalParams.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbGlobalParams.Name = "gbGlobalParams"
        Me.gbGlobalParams.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbGlobalParams.Size = New System.Drawing.Size(338, 65)
        Me.gbGlobalParams.TabIndex = 129
        Me.gbGlobalParams.TabStop = False
        '
        'rbShowInactiveUserIDs
        '
        Me.rbShowInactiveUserIDs.AutoSize = True
        Me.rbShowInactiveUserIDs.Location = New System.Drawing.Point(164, 14)
        Me.rbShowInactiveUserIDs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowInactiveUserIDs.Name = "rbShowInactiveUserIDs"
        Me.rbShowInactiveUserIDs.Size = New System.Drawing.Size(157, 21)
        Me.rbShowInactiveUserIDs.TabIndex = 2
        Me.rbShowInactiveUserIDs.TabStop = True
        Me.rbShowInactiveUserIDs.Text = "Show Inactive User IDs"
        Me.rbShowInactiveUserIDs.UseVisualStyleBackColor = True
        '
        'rbShowActiveUserIDs
        '
        Me.rbShowActiveUserIDs.AutoSize = True
        Me.rbShowActiveUserIDs.Location = New System.Drawing.Point(7, 39)
        Me.rbShowActiveUserIDs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowActiveUserIDs.Name = "rbShowActiveUserIDs"
        Me.rbShowActiveUserIDs.Size = New System.Drawing.Size(148, 21)
        Me.rbShowActiveUserIDs.TabIndex = 1
        Me.rbShowActiveUserIDs.TabStop = True
        Me.rbShowActiveUserIDs.Text = "Show Active User IDs"
        Me.rbShowActiveUserIDs.UseVisualStyleBackColor = True
        '
        'rbShowAllUserIDs
        '
        Me.rbShowAllUserIDs.AutoSize = True
        Me.rbShowAllUserIDs.Location = New System.Drawing.Point(7, 14)
        Me.rbShowAllUserIDs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowAllUserIDs.Name = "rbShowAllUserIDs"
        Me.rbShowAllUserIDs.Size = New System.Drawing.Size(128, 21)
        Me.rbShowAllUserIDs.TabIndex = 0
        Me.rbShowAllUserIDs.TabStop = True
        Me.rbShowAllUserIDs.Text = "Show All User IDs"
        Me.rbShowAllUserIDs.UseVisualStyleBackColor = True
        '
        'gbUserShow
        '
        Me.gbUserShow.BackColor = System.Drawing.Color.Transparent
        Me.gbUserShow.Controls.Add(Me.rbShowInactiveUsers)
        Me.gbUserShow.Controls.Add(Me.rbShowActiveUsers)
        Me.gbUserShow.Controls.Add(Me.rbShowAllUsers)
        Me.gbUserShow.Location = New System.Drawing.Point(9, 61)
        Me.gbUserShow.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbUserShow.Name = "gbUserShow"
        Me.gbUserShow.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbUserShow.Size = New System.Drawing.Size(261, 65)
        Me.gbUserShow.TabIndex = 128
        Me.gbUserShow.TabStop = False
        '
        'rbShowInactiveUsers
        '
        Me.rbShowInactiveUsers.AutoSize = True
        Me.rbShowInactiveUsers.Location = New System.Drawing.Point(7, 38)
        Me.rbShowInactiveUsers.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowInactiveUsers.Name = "rbShowInactiveUsers"
        Me.rbShowInactiveUsers.Size = New System.Drawing.Size(141, 21)
        Me.rbShowInactiveUsers.TabIndex = 2
        Me.rbShowInactiveUsers.TabStop = True
        Me.rbShowInactiveUsers.Text = "Show Inactive Users"
        Me.rbShowInactiveUsers.UseVisualStyleBackColor = True
        '
        'rbShowActiveUsers
        '
        Me.rbShowActiveUsers.AutoSize = True
        Me.rbShowActiveUsers.Location = New System.Drawing.Point(126, 14)
        Me.rbShowActiveUsers.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowActiveUsers.Name = "rbShowActiveUsers"
        Me.rbShowActiveUsers.Size = New System.Drawing.Size(132, 21)
        Me.rbShowActiveUsers.TabIndex = 1
        Me.rbShowActiveUsers.TabStop = True
        Me.rbShowActiveUsers.Text = "Show Active Users"
        Me.rbShowActiveUsers.UseVisualStyleBackColor = True
        '
        'rbShowAllUsers
        '
        Me.rbShowAllUsers.AutoSize = True
        Me.rbShowAllUsers.Location = New System.Drawing.Point(7, 14)
        Me.rbShowAllUsers.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowAllUsers.Name = "rbShowAllUsers"
        Me.rbShowAllUsers.Size = New System.Drawing.Size(112, 21)
        Me.rbShowAllUsers.TabIndex = 0
        Me.rbShowAllUsers.TabStop = True
        Me.rbShowAllUsers.Text = "Show All Users"
        Me.rbShowAllUsers.UseVisualStyleBackColor = True
        '
        'dgvUserAttributes
        '
        Me.dgvUserAttributes.AllowUserToAddRows = False
        Me.dgvUserAttributes.AllowUserToDeleteRows = False
        Me.dgvUserAttributes.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvUserAttributes.BackgroundColor = System.Drawing.Color.White
        Me.dgvUserAttributes.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvUserAttributes.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvUserAttributes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvUserAttributes.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvUserAttributes.Location = New System.Drawing.Point(374, 150)
        Me.dgvUserAttributes.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvUserAttributes.MultiSelect = False
        Me.dgvUserAttributes.Name = "dgvUserAttributes"
        Me.dgvUserAttributes.ReadOnly = True
        Me.dgvUserAttributes.RowHeadersWidth = 25
        Me.dgvUserAttributes.Size = New System.Drawing.Size(542, 170)
        Me.dgvUserAttributes.TabIndex = 125
        '
        'dgvUsers
        '
        Me.dgvUsers.AllowUserToAddRows = False
        Me.dgvUsers.AllowUserToDeleteRows = False
        Me.dgvUsers.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvUsers.BackgroundColor = System.Drawing.Color.White
        Me.dgvUsers.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvUsers.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvUsers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvUsers.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgvUsers.Location = New System.Drawing.Point(9, 150)
        Me.dgvUsers.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvUsers.MultiSelect = False
        Me.dgvUsers.Name = "dgvUsers"
        Me.dgvUsers.ReadOnly = True
        Me.dgvUsers.RowHeadersWidth = 25
        Me.dgvUsers.Size = New System.Drawing.Size(359, 654)
        Me.dgvUsers.TabIndex = 124
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(374, 132)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(124, 17)
        Me.Label7.TabIndex = 122
        Me.Label7.Text = "Configure User IDs"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(9, 132)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(81, 17)
        Me.Label6.TabIndex = 121
        Me.Label6.Text = "User Names"
        '
        'cmdSelectAllPermissions
        '
        Me.cmdSelectAllPermissions.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSelectAllPermissions.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSelectAllPermissions.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdSelectAllPermissions.Location = New System.Drawing.Point(495, 90)
        Me.cmdSelectAllPermissions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSelectAllPermissions.Name = "cmdSelectAllPermissions"
        Me.cmdSelectAllPermissions.Size = New System.Drawing.Size(92, 33)
        Me.cmdSelectAllPermissions.TabIndex = 143
        Me.cmdSelectAllPermissions.Text = "&Select All"
        Me.cmdSelectAllPermissions.UseVisualStyleBackColor = False
        '
        'cmdDeselectAllPermissions
        '
        Me.cmdDeselectAllPermissions.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdDeselectAllPermissions.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeselectAllPermissions.ForeColor = System.Drawing.Color.Red
        Me.cmdDeselectAllPermissions.Location = New System.Drawing.Point(594, 90)
        Me.cmdDeselectAllPermissions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdDeselectAllPermissions.Name = "cmdDeselectAllPermissions"
        Me.cmdDeselectAllPermissions.Size = New System.Drawing.Size(92, 33)
        Me.cmdDeselectAllPermissions.TabIndex = 142
        Me.cmdDeselectAllPermissions.Text = "&Deselect All"
        Me.cmdDeselectAllPermissions.UseVisualStyleBackColor = False
        '
        'lblcbxModulesPers
        '
        Me.lblcbxModulesPers.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblcbxModulesPers.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblcbxModulesPers.ForeColor = System.Drawing.Color.White
        Me.lblcbxModulesPers.Location = New System.Drawing.Point(206, 94)
        Me.lblcbxModulesPers.Name = "lblcbxModulesPers"
        Me.lblcbxModulesPers.Size = New System.Drawing.Size(164, 17)
        Me.lblcbxModulesPers.TabIndex = 140
        Me.lblcbxModulesPers.Text = "Choose a Module"
        Me.lblcbxModulesPers.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lvPermissionsAdmin
        '
        Me.lvPermissionsAdmin.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvPermissionsAdmin.Enabled = False
        Me.lvPermissionsAdmin.Location = New System.Drawing.Point(171, 4)
        Me.lvPermissionsAdmin.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lvPermissionsAdmin.MultiSelect = False
        Me.lvPermissionsAdmin.Name = "lvPermissionsAdmin"
        Me.lvPermissionsAdmin.Size = New System.Drawing.Size(105, 53)
        Me.lvPermissionsAdmin.TabIndex = 141
        Me.lvPermissionsAdmin.UseCompatibleStateImageBehavior = False
        '
        'lvPermissions
        '
        Me.lvPermissions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvPermissions.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvPermissions.Enabled = False
        Me.lvPermissions.Location = New System.Drawing.Point(377, 148)
        Me.lvPermissions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lvPermissions.MultiSelect = False
        Me.lvPermissions.Name = "lvPermissions"
        Me.lvPermissions.Size = New System.Drawing.Size(1081, 36)
        Me.lvPermissions.TabIndex = 133
        Me.lvPermissions.UseCompatibleStateImageBehavior = False
        '
        'pan2
        '
        Me.pan2.AutoScroll = True
        Me.pan2.BackColor = System.Drawing.Color.White
        Me.pan2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan2.Controls.Add(Me.dgvDropdownboxTitle)
        Me.pan2.Controls.Add(Me.cmdOrderDropdownbox)
        Me.pan2.Controls.Add(Me.dgvDropdownboxContents)
        Me.pan2.Controls.Add(Me.Label16)
        Me.pan2.Controls.Add(Me.Label12)
        Me.pan2.Controls.Add(Me.cmdAddDropdownbox)
        Me.pan2.Controls.Add(Me.cmdResetDropdownbox)
        Me.pan2.Controls.Add(Me.Label1)
        Me.pan2.Location = New System.Drawing.Point(687, 106)
        Me.pan2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan2.Name = "pan2"
        Me.pan2.Size = New System.Drawing.Size(384, 295)
        Me.pan2.TabIndex = 121
        Me.pan2.Visible = False
        '
        'dgvDropdownboxTitle
        '
        Me.dgvDropdownboxTitle.AllowUserToAddRows = False
        Me.dgvDropdownboxTitle.AllowUserToDeleteRows = False
        Me.dgvDropdownboxTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvDropdownboxTitle.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvDropdownboxTitle.BackgroundColor = System.Drawing.Color.White
        Me.dgvDropdownboxTitle.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDropdownboxTitle.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.dgvDropdownboxTitle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDropdownboxTitle.Location = New System.Drawing.Point(12, 171)
        Me.dgvDropdownboxTitle.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvDropdownboxTitle.Name = "dgvDropdownboxTitle"
        Me.dgvDropdownboxTitle.ReadOnly = True
        Me.dgvDropdownboxTitle.RowHeadersWidth = 25
        Me.dgvDropdownboxTitle.Size = New System.Drawing.Size(252, 105)
        Me.dgvDropdownboxTitle.TabIndex = 142
        '
        'cmdOrderDropdownbox
        '
        Me.cmdOrderDropdownbox.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOrderDropdownbox.BackgroundImage = CType(resources.GetObject("cmdOrderDropdownbox.BackgroundImage"), System.Drawing.Image)
        Me.cmdOrderDropdownbox.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdOrderDropdownbox.Enabled = False
        Me.cmdOrderDropdownbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOrderDropdownbox.Location = New System.Drawing.Point(337, 131)
        Me.cmdOrderDropdownbox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdOrderDropdownbox.Name = "cmdOrderDropdownbox"
        Me.cmdOrderDropdownbox.Size = New System.Drawing.Size(55, 34)
        Me.cmdOrderDropdownbox.TabIndex = 141
        Me.cmdOrderDropdownbox.Text = "Order"
        Me.cmdOrderDropdownbox.UseVisualStyleBackColor = False
        '
        'dgvDropdownboxContents
        '
        Me.dgvDropdownboxContents.AllowUserToAddRows = False
        Me.dgvDropdownboxContents.AllowUserToDeleteRows = False
        Me.dgvDropdownboxContents.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDropdownboxContents.BackgroundColor = System.Drawing.Color.White
        Me.dgvDropdownboxContents.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDropdownboxContents.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvDropdownboxContents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDropdownboxContents.Location = New System.Drawing.Point(271, 171)
        Me.dgvDropdownboxContents.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvDropdownboxContents.Name = "dgvDropdownboxContents"
        Me.dgvDropdownboxContents.ReadOnly = True
        Me.dgvDropdownboxContents.RowHeadersWidth = 25
        Me.dgvDropdownboxContents.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDropdownboxContents.Size = New System.Drawing.Size(248, 105)
        Me.dgvDropdownboxContents.TabIndex = 140
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label16.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.White
        Me.Label16.Location = New System.Drawing.Point(12, 153)
        Me.Label16.Margin = New System.Windows.Forms.Padding(0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(132, 17)
        Me.Label16.TabIndex = 139
        Me.Label16.Text = "Dropdown Box Title"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label12.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(271, 153)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(159, 17)
        Me.Label12.TabIndex = 138
        Me.Label12.Text = "Dropdown Box Contents"
        '
        'cmdAddDropdownbox
        '
        Me.cmdAddDropdownbox.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddDropdownbox.Enabled = False
        Me.cmdAddDropdownbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddDropdownbox.ForeColor = System.Drawing.Color.Blue
        Me.cmdAddDropdownbox.Location = New System.Drawing.Point(271, 112)
        Me.cmdAddDropdownbox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddDropdownbox.Name = "cmdAddDropdownbox"
        Me.cmdAddDropdownbox.Size = New System.Drawing.Size(56, 33)
        Me.cmdAddDropdownbox.TabIndex = 137
        Me.cmdAddDropdownbox.Text = "&Add"
        Me.cmdAddDropdownbox.UseVisualStyleBackColor = False
        '
        'cmdResetDropdownbox
        '
        Me.cmdResetDropdownbox.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdResetDropdownbox.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdResetDropdownbox.CausesValidation = False
        Me.cmdResetDropdownbox.Enabled = False
        Me.cmdResetDropdownbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetDropdownbox.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdResetDropdownbox.Location = New System.Drawing.Point(15299, 23)
        Me.cmdResetDropdownbox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdResetDropdownbox.Name = "cmdResetDropdownbox"
        Me.cmdResetDropdownbox.Size = New System.Drawing.Size(82, 33)
        Me.cmdResetDropdownbox.TabIndex = 136
        Me.cmdResetDropdownbox.Text = "Reset"
        Me.cmdResetDropdownbox.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(430, 21)
        Me.Label1.TabIndex = 135
        Me.Label1.Text = "Dropdownbox Configuration"
        '
        'pan3
        '
        Me.pan3.AutoScroll = True
        Me.pan3.BackColor = System.Drawing.Color.White
        Me.pan3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan3.Controls.Add(Me.gbxlblCorporateAdderesses)
        Me.pan3.Controls.Add(Me.gbDropDown)
        Me.pan3.Controls.Add(Me.Label15)
        Me.pan3.Controls.Add(Me.Label14)
        Me.pan3.Controls.Add(Me.cmdAddCorporateAddress)
        Me.pan3.Controls.Add(Me.dgvCorporateAddresses)
        Me.pan3.Controls.Add(Me.dgvNickNames)
        Me.pan3.Controls.Add(Me.cmdResetCorporateAddressses)
        Me.pan3.Controls.Add(Me.Label2)
        Me.pan3.Location = New System.Drawing.Point(429, 516)
        Me.pan3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan3.Name = "pan3"
        Me.pan3.Size = New System.Drawing.Size(724, 472)
        Me.pan3.TabIndex = 122
        Me.pan3.Visible = False
        '
        'gbxlblCorporateAdderesses
        '
        Me.gbxlblCorporateAdderesses.AutoSize = True
        Me.gbxlblCorporateAdderesses.Controls.Add(Me.Label11)
        Me.gbxlblCorporateAdderesses.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblCorporateAdderesses.Location = New System.Drawing.Point(413, 142)
        Me.gbxlblCorporateAdderesses.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxlblCorporateAdderesses.Name = "gbxlblCorporateAdderesses"
        Me.gbxlblCorporateAdderesses.Padding = New System.Windows.Forms.Padding(0)
        Me.gbxlblCorporateAdderesses.Size = New System.Drawing.Size(295, 29)
        Me.gbxlblCorporateAdderesses.TabIndex = 128
        Me.gbxlblCorporateAdderesses.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label11.Location = New System.Drawing.Point(3, 5)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(247, 13)
        Me.Label11.TabIndex = 124
        Me.Label11.Text = "* A = Include this label item in the Report Title."
        '
        'gbDropDown
        '
        Me.gbDropDown.BackColor = System.Drawing.Color.Transparent
        Me.gbDropDown.Controls.Add(Me.rbShowInactiveAddresses)
        Me.gbDropDown.Controls.Add(Me.rbShowActiveAddresses)
        Me.gbDropDown.Controls.Add(Me.rbShowAllAddresses)
        Me.gbDropDown.Location = New System.Drawing.Point(10, 23)
        Me.gbDropDown.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbDropDown.Name = "gbDropDown"
        Me.gbDropDown.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbDropDown.Size = New System.Drawing.Size(359, 69)
        Me.gbDropDown.TabIndex = 125
        Me.gbDropDown.TabStop = False
        '
        'rbShowInactiveAddresses
        '
        Me.rbShowInactiveAddresses.AutoSize = True
        Me.rbShowInactiveAddresses.Location = New System.Drawing.Point(183, 14)
        Me.rbShowInactiveAddresses.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowInactiveAddresses.Name = "rbShowInactiveAddresses"
        Me.rbShowInactiveAddresses.Size = New System.Drawing.Size(169, 21)
        Me.rbShowInactiveAddresses.TabIndex = 2
        Me.rbShowInactiveAddresses.TabStop = True
        Me.rbShowInactiveAddresses.Text = "Show Inactive Addresses"
        Me.rbShowInactiveAddresses.UseVisualStyleBackColor = True
        '
        'rbShowActiveAddresses
        '
        Me.rbShowActiveAddresses.AutoSize = True
        Me.rbShowActiveAddresses.Location = New System.Drawing.Point(7, 39)
        Me.rbShowActiveAddresses.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowActiveAddresses.Name = "rbShowActiveAddresses"
        Me.rbShowActiveAddresses.Size = New System.Drawing.Size(160, 21)
        Me.rbShowActiveAddresses.TabIndex = 1
        Me.rbShowActiveAddresses.TabStop = True
        Me.rbShowActiveAddresses.Text = "Show Active Addresses"
        Me.rbShowActiveAddresses.UseVisualStyleBackColor = True
        '
        'rbShowAllAddresses
        '
        Me.rbShowAllAddresses.AutoSize = True
        Me.rbShowAllAddresses.BackColor = System.Drawing.Color.Transparent
        Me.rbShowAllAddresses.Location = New System.Drawing.Point(7, 14)
        Me.rbShowAllAddresses.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowAllAddresses.Name = "rbShowAllAddresses"
        Me.rbShowAllAddresses.Size = New System.Drawing.Size(140, 21)
        Me.rbShowAllAddresses.TabIndex = 0
        Me.rbShowAllAddresses.TabStop = True
        Me.rbShowAllAddresses.Text = "Show All Addresses"
        Me.rbShowAllAddresses.UseVisualStyleBackColor = False
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label15.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(292, 155)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(70, 17)
        Me.Label15.TabIndex = 127
        Me.Label15.Text = "Addresses"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label14.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(10, 155)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(77, 17)
        Me.Label14.TabIndex = 126
        Me.Label14.Text = "NickNames"
        '
        'cmdAddCorporateAddress
        '
        Me.cmdAddCorporateAddress.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddCorporateAddress.Enabled = False
        Me.cmdAddCorporateAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddCorporateAddress.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddCorporateAddress.Location = New System.Drawing.Point(117, 140)
        Me.cmdAddCorporateAddress.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddCorporateAddress.Name = "cmdAddCorporateAddress"
        Me.cmdAddCorporateAddress.Size = New System.Drawing.Size(56, 33)
        Me.cmdAddCorporateAddress.TabIndex = 123
        Me.cmdAddCorporateAddress.Text = "&Add"
        Me.cmdAddCorporateAddress.UseVisualStyleBackColor = False
        '
        'dgvCorporateAddresses
        '
        Me.dgvCorporateAddresses.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCorporateAddresses.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvCorporateAddresses.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvCorporateAddresses.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCorporateAddresses.Location = New System.Drawing.Point(292, 173)
        Me.dgvCorporateAddresses.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvCorporateAddresses.MultiSelect = False
        Me.dgvCorporateAddresses.Name = "dgvCorporateAddresses"
        Me.dgvCorporateAddresses.ReadOnly = True
        Me.dgvCorporateAddresses.Size = New System.Drawing.Size(416, 297)
        Me.dgvCorporateAddresses.TabIndex = 122
        '
        'dgvNickNames
        '
        Me.dgvNickNames.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvNickNames.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvNickNames.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle8
        Me.dgvNickNames.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvNickNames.Location = New System.Drawing.Point(10, 173)
        Me.dgvNickNames.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvNickNames.MultiSelect = False
        Me.dgvNickNames.Name = "dgvNickNames"
        Me.dgvNickNames.ReadOnly = True
        Me.dgvNickNames.Size = New System.Drawing.Size(273, 297)
        Me.dgvNickNames.TabIndex = 121
        '
        'cmdResetCorporateAddressses
        '
        Me.cmdResetCorporateAddressses.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdResetCorporateAddressses.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdResetCorporateAddressses.CausesValidation = False
        Me.cmdResetCorporateAddressses.Enabled = False
        Me.cmdResetCorporateAddressses.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetCorporateAddressses.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdResetCorporateAddressses.Location = New System.Drawing.Point(15656, 23)
        Me.cmdResetCorporateAddressses.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdResetCorporateAddressses.Name = "cmdResetCorporateAddressses"
        Me.cmdResetCorporateAddressses.Size = New System.Drawing.Size(82, 33)
        Me.cmdResetCorporateAddressses.TabIndex = 120
        Me.cmdResetCorporateAddressses.Text = "Reset"
        Me.cmdResetCorporateAddressses.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(0, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(722, 21)
        Me.Label2.TabIndex = 119
        Me.Label2.Text = "Corporate Addresses"
        '
        'pan4
        '
        Me.pan4.AutoScroll = True
        Me.pan4.BackColor = System.Drawing.Color.White
        Me.pan4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan4.Controls.Add(Me.gbSTD)
        Me.pan4.Controls.Add(Me.dgvTemplates)
        Me.pan4.Controls.Add(Me.Label10)
        Me.pan4.Controls.Add(Me.Label9)
        Me.pan4.Controls.Add(Me.cmdAddTemplate)
        Me.pan4.Controls.Add(Me.gbStudyTemplates)
        Me.pan4.Controls.Add(Me.dgvTemplateAttributes)
        Me.pan4.Controls.Add(Me.cmdResetDefineReports)
        Me.pan4.Controls.Add(Me.Label5)
        Me.pan4.Location = New System.Drawing.Point(1045, 554)
        Me.pan4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan4.Name = "pan4"
        Me.pan4.Size = New System.Drawing.Size(575, 288)
        Me.pan4.TabIndex = 122
        Me.pan4.Visible = False
        '
        'gbSTD
        '
        Me.gbSTD.AutoSize = True
        Me.gbSTD.Controls.Add(Me.lblTExpl)
        Me.gbSTD.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSTD.Location = New System.Drawing.Point(397, 64)
        Me.gbSTD.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSTD.Name = "gbSTD"
        Me.gbSTD.Padding = New System.Windows.Forms.Padding(0)
        Me.gbSTD.Size = New System.Drawing.Size(359, 60)
        Me.gbSTD.TabIndex = 129
        Me.gbSTD.TabStop = False
        '
        'lblTExpl
        '
        Me.lblTExpl.BackColor = System.Drawing.Color.Transparent
        Me.lblTExpl.Font = New System.Drawing.Font("Segoe UI", 8.25!)
        Me.lblTExpl.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblTExpl.Location = New System.Drawing.Point(2, 4)
        Me.lblTExpl.Margin = New System.Windows.Forms.Padding(0)
        Me.lblTExpl.Name = "lblTExpl"
        Me.lblTExpl.Size = New System.Drawing.Size(357, 50)
        Me.lblTExpl.TabIndex = 127
        Me.lblTExpl.Text = "Label to be entered in code" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "here."
        '
        'dgvTemplates
        '
        Me.dgvTemplates.AllowUserToAddRows = False
        Me.dgvTemplates.AllowUserToDeleteRows = False
        Me.dgvTemplates.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvTemplates.BackgroundColor = System.Drawing.Color.White
        Me.dgvTemplates.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTemplates.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle9
        Me.dgvTemplates.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTemplates.Location = New System.Drawing.Point(10, 173)
        Me.dgvTemplates.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvTemplates.Name = "dgvTemplates"
        Me.dgvTemplates.ReadOnly = True
        Me.dgvTemplates.RowHeadersWidth = 25
        Me.dgvTemplates.Size = New System.Drawing.Size(378, 609)
        Me.dgvTemplates.TabIndex = 128
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(397, 153)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(142, 19)
        Me.Label10.TabIndex = 126
        Me.Label10.Text = "Template Attributes"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(10, 153)
        Me.Label9.Margin = New System.Windows.Forms.Padding(0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(78, 19)
        Me.Label9.TabIndex = 125
        Me.Label9.Text = "Templates"
        '
        'cmdAddTemplate
        '
        Me.cmdAddTemplate.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddTemplate.Enabled = False
        Me.cmdAddTemplate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddTemplate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddTemplate.Location = New System.Drawing.Point(105, 140)
        Me.cmdAddTemplate.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddTemplate.Name = "cmdAddTemplate"
        Me.cmdAddTemplate.Size = New System.Drawing.Size(56, 33)
        Me.cmdAddTemplate.TabIndex = 124
        Me.cmdAddTemplate.Text = "&Add"
        Me.cmdAddTemplate.UseVisualStyleBackColor = False
        '
        'gbStudyTemplates
        '
        Me.gbStudyTemplates.BackColor = System.Drawing.Color.Transparent
        Me.gbStudyTemplates.Controls.Add(Me.rbShowInactiveTemplates)
        Me.gbStudyTemplates.Controls.Add(Me.rbShowActiveTemplates)
        Me.gbStudyTemplates.Controls.Add(Me.rbShowAllTemplates)
        Me.gbStudyTemplates.Location = New System.Drawing.Point(10, 55)
        Me.gbStudyTemplates.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStudyTemplates.Name = "gbStudyTemplates"
        Me.gbStudyTemplates.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStudyTemplates.Size = New System.Drawing.Size(359, 69)
        Me.gbStudyTemplates.TabIndex = 123
        Me.gbStudyTemplates.TabStop = False
        '
        'rbShowInactiveTemplates
        '
        Me.rbShowInactiveTemplates.AutoSize = True
        Me.rbShowInactiveTemplates.Location = New System.Drawing.Point(183, 14)
        Me.rbShowInactiveTemplates.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowInactiveTemplates.Name = "rbShowInactiveTemplates"
        Me.rbShowInactiveTemplates.Size = New System.Drawing.Size(168, 21)
        Me.rbShowInactiveTemplates.TabIndex = 2
        Me.rbShowInactiveTemplates.TabStop = True
        Me.rbShowInactiveTemplates.Text = "Show Inactive Templates"
        Me.rbShowInactiveTemplates.UseVisualStyleBackColor = True
        '
        'rbShowActiveTemplates
        '
        Me.rbShowActiveTemplates.AutoSize = True
        Me.rbShowActiveTemplates.Location = New System.Drawing.Point(7, 39)
        Me.rbShowActiveTemplates.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowActiveTemplates.Name = "rbShowActiveTemplates"
        Me.rbShowActiveTemplates.Size = New System.Drawing.Size(159, 21)
        Me.rbShowActiveTemplates.TabIndex = 1
        Me.rbShowActiveTemplates.TabStop = True
        Me.rbShowActiveTemplates.Text = "Show Active Templates"
        Me.rbShowActiveTemplates.UseVisualStyleBackColor = True
        '
        'rbShowAllTemplates
        '
        Me.rbShowAllTemplates.AutoSize = True
        Me.rbShowAllTemplates.Location = New System.Drawing.Point(7, 14)
        Me.rbShowAllTemplates.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowAllTemplates.Name = "rbShowAllTemplates"
        Me.rbShowAllTemplates.Size = New System.Drawing.Size(139, 21)
        Me.rbShowAllTemplates.TabIndex = 0
        Me.rbShowAllTemplates.TabStop = True
        Me.rbShowAllTemplates.Text = "Show All Templates"
        Me.rbShowAllTemplates.UseVisualStyleBackColor = True
        '
        'dgvTemplateAttributes
        '
        Me.dgvTemplateAttributes.AllowUserToAddRows = False
        Me.dgvTemplateAttributes.AllowUserToDeleteRows = False
        Me.dgvTemplateAttributes.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvTemplateAttributes.BackgroundColor = System.Drawing.Color.White
        Me.dgvTemplateAttributes.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle10.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTemplateAttributes.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle10
        Me.dgvTemplateAttributes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTemplateAttributes.Location = New System.Drawing.Point(397, 173)
        Me.dgvTemplateAttributes.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvTemplateAttributes.MultiSelect = False
        Me.dgvTemplateAttributes.Name = "dgvTemplateAttributes"
        Me.dgvTemplateAttributes.ReadOnly = True
        Me.dgvTemplateAttributes.RowHeadersWidth = 25
        Me.dgvTemplateAttributes.Size = New System.Drawing.Size(359, 609)
        Me.dgvTemplateAttributes.TabIndex = 122
        '
        'cmdResetDefineReports
        '
        Me.cmdResetDefineReports.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdResetDefineReports.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdResetDefineReports.CausesValidation = False
        Me.cmdResetDefineReports.Enabled = False
        Me.cmdResetDefineReports.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetDefineReports.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdResetDefineReports.Location = New System.Drawing.Point(7915, 23)
        Me.cmdResetDefineReports.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdResetDefineReports.Name = "cmdResetDefineReports"
        Me.cmdResetDefineReports.Size = New System.Drawing.Size(82, 33)
        Me.cmdResetDefineReports.TabIndex = 121
        Me.cmdResetDefineReports.Text = "Reset"
        Me.cmdResetDefineReports.UseVisualStyleBackColor = False
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(0, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(756, 21)
        Me.Label5.TabIndex = 120
        Me.Label5.Text = "Study Template Definitions"
        '
        'pan5
        '
        Me.pan5.AutoScroll = True
        Me.pan5.BackColor = System.Drawing.Color.White
        Me.pan5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan5.Controls.Add(Me.gbxlabelGlobalParameters)
        Me.pan5.Controls.Add(Me.panGP)
        Me.pan5.Controls.Add(Me.cmdResetGlobal)
        Me.pan5.Controls.Add(Me.cmdBrowseGlobal)
        Me.pan5.Controls.Add(Me.lblGlobalParameters)
        Me.pan5.Location = New System.Drawing.Point(222, 65)
        Me.pan5.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan5.Name = "pan5"
        Me.pan5.Size = New System.Drawing.Size(418, 430)
        Me.pan5.TabIndex = 122
        Me.pan5.Visible = False
        '
        'gbxlabelGlobalParameters
        '
        Me.gbxlabelGlobalParameters.AutoSize = True
        Me.gbxlabelGlobalParameters.Controls.Add(Me.lblIntegrity)
        Me.gbxlabelGlobalParameters.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlabelGlobalParameters.Location = New System.Drawing.Point(215, 24)
        Me.gbxlabelGlobalParameters.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxlabelGlobalParameters.Name = "gbxlabelGlobalParameters"
        Me.gbxlabelGlobalParameters.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxlabelGlobalParameters.Size = New System.Drawing.Size(544, 196)
        Me.gbxlabelGlobalParameters.TabIndex = 131
        Me.gbxlabelGlobalParameters.TabStop = False
        Me.gbxlabelGlobalParameters.Visible = False
        '
        'lblIntegrity
        '
        Me.lblIntegrity.BackColor = System.Drawing.Color.Transparent
        Me.lblIntegrity.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIntegrity.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblIntegrity.Location = New System.Drawing.Point(6, 11)
        Me.lblIntegrity.Name = "lblIntegrity"
        Me.lblIntegrity.Size = New System.Drawing.Size(530, 176)
        Me.lblIntegrity.TabIndex = 130
        Me.lblIntegrity.Text = "Entered by code"
        '
        'panGP
        '
        Me.panGP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panGP.Controls.Add(Me.Label17)
        Me.panGP.Controls.Add(Me.dgvGlobal)
        Me.panGP.Controls.Add(Me.lblGlobalValues)
        Me.panGP.Controls.Add(Me.lbxGlobal)
        Me.panGP.Location = New System.Drawing.Point(9, 223)
        Me.panGP.Name = "panGP"
        Me.panGP.Size = New System.Drawing.Size(2114, 186)
        Me.panGP.TabIndex = 132
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label17.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label17.ForeColor = System.Drawing.Color.White
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(73, 17)
        Me.Label17.TabIndex = 126
        Me.Label17.Text = "Categories"
        '
        'dgvGlobal
        '
        Me.dgvGlobal.AllowUserToAddRows = False
        Me.dgvGlobal.AllowUserToDeleteRows = False
        Me.dgvGlobal.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvGlobal.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvGlobal.BackgroundColor = System.Drawing.Color.White
        Me.dgvGlobal.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGlobal.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle11
        Me.dgvGlobal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGlobal.Location = New System.Drawing.Point(206, 21)
        Me.dgvGlobal.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvGlobal.MultiSelect = False
        Me.dgvGlobal.Name = "dgvGlobal"
        Me.dgvGlobal.ReadOnly = True
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle12.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle12.Padding = New System.Windows.Forms.Padding(0, 5, 0, 5)
        DataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGlobal.RowHeadersDefaultCellStyle = DataGridViewCellStyle12
        Me.dgvGlobal.RowHeadersWidth = 25
        Me.dgvGlobal.Size = New System.Drawing.Size(1222, 155)
        Me.dgvGlobal.TabIndex = 125
        '
        'lblGlobalValues
        '
        Me.lblGlobalValues.AutoSize = True
        Me.lblGlobalValues.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblGlobalValues.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGlobalValues.ForeColor = System.Drawing.Color.White
        Me.lblGlobalValues.Location = New System.Drawing.Point(206, 0)
        Me.lblGlobalValues.Name = "lblGlobalValues"
        Me.lblGlobalValues.Size = New System.Drawing.Size(49, 17)
        Me.lblGlobalValues.TabIndex = 127
        Me.lblGlobalValues.Text = "Values"
        '
        'lbxGlobal
        '
        Me.lbxGlobal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbxGlobal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbxGlobal.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbxGlobal.ItemHeight = 17
        Me.lbxGlobal.Location = New System.Drawing.Point(0, 21)
        Me.lbxGlobal.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lbxGlobal.Name = "lbxGlobal"
        Me.lbxGlobal.Size = New System.Drawing.Size(194, 155)
        Me.lbxGlobal.TabIndex = 124
        '
        'cmdResetGlobal
        '
        Me.cmdResetGlobal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdResetGlobal.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdResetGlobal.CausesValidation = False
        Me.cmdResetGlobal.Enabled = False
        Me.cmdResetGlobal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetGlobal.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdResetGlobal.Location = New System.Drawing.Point(5133, 23)
        Me.cmdResetGlobal.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdResetGlobal.Name = "cmdResetGlobal"
        Me.cmdResetGlobal.Size = New System.Drawing.Size(82, 33)
        Me.cmdResetGlobal.TabIndex = 129
        Me.cmdResetGlobal.Text = "Reset"
        Me.cmdResetGlobal.UseVisualStyleBackColor = False
        '
        'cmdBrowseGlobal
        '
        Me.cmdBrowseGlobal.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdBrowseGlobal.Enabled = False
        Me.cmdBrowseGlobal.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseGlobal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdBrowseGlobal.Location = New System.Drawing.Point(80, 28)
        Me.cmdBrowseGlobal.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdBrowseGlobal.Name = "cmdBrowseGlobal"
        Me.cmdBrowseGlobal.Size = New System.Drawing.Size(98, 33)
        Me.cmdBrowseGlobal.TabIndex = 128
        Me.cmdBrowseGlobal.Text = "&Browse..."
        Me.cmdBrowseGlobal.UseVisualStyleBackColor = False
        '
        'lblGlobalParameters
        '
        Me.lblGlobalParameters.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblGlobalParameters.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblGlobalParameters.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblGlobalParameters.ForeColor = System.Drawing.Color.White
        Me.lblGlobalParameters.Location = New System.Drawing.Point(0, 0)
        Me.lblGlobalParameters.Name = "lblGlobalParameters"
        Me.lblGlobalParameters.Size = New System.Drawing.Size(759, 21)
        Me.lblGlobalParameters.TabIndex = 123
        Me.lblGlobalParameters.Text = "Global Parameters"
        '
        'pan6
        '
        Me.pan6.BackColor = System.Drawing.Color.White
        Me.pan6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan6.Controls.Add(Me.cmdRefreshHook)
        Me.pan6.Controls.Add(Me.cmdResetHooks)
        Me.pan6.Controls.Add(Me.dgvHooks)
        Me.pan6.Controls.Add(Me.cmdAddHook)
        Me.pan6.Controls.Add(Me.Label19)
        Me.pan6.Location = New System.Drawing.Point(774, 190)
        Me.pan6.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan6.Name = "pan6"
        Me.pan6.Size = New System.Drawing.Size(347, 168)
        Me.pan6.TabIndex = 123
        Me.pan6.Visible = False
        '
        'cmdRefreshHook
        '
        Me.cmdRefreshHook.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRefreshHook.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRefreshHook.CausesValidation = False
        Me.cmdRefreshHook.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRefreshHook.ForeColor = System.Drawing.Color.Chocolate
        Me.cmdRefreshHook.Location = New System.Drawing.Point(249, 64)
        Me.cmdRefreshHook.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdRefreshHook.Name = "cmdRefreshHook"
        Me.cmdRefreshHook.Size = New System.Drawing.Size(93, 54)
        Me.cmdRefreshHook.TabIndex = 131
        Me.cmdRefreshHook.Text = "Refresh &Hook"
        Me.cmdRefreshHook.UseVisualStyleBackColor = False
        '
        'cmdResetHooks
        '
        Me.cmdResetHooks.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdResetHooks.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdResetHooks.CausesValidation = False
        Me.cmdResetHooks.Enabled = False
        Me.cmdResetHooks.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetHooks.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdResetHooks.Location = New System.Drawing.Point(260, 23)
        Me.cmdResetHooks.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdResetHooks.Name = "cmdResetHooks"
        Me.cmdResetHooks.Size = New System.Drawing.Size(82, 33)
        Me.cmdResetHooks.TabIndex = 130
        Me.cmdResetHooks.Text = "&Reset"
        Me.cmdResetHooks.UseVisualStyleBackColor = False
        '
        'dgvHooks
        '
        Me.dgvHooks.AllowUserToAddRows = False
        Me.dgvHooks.AllowUserToDeleteRows = False
        Me.dgvHooks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvHooks.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvHooks.BackgroundColor = System.Drawing.Color.White
        Me.dgvHooks.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle13.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvHooks.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle13
        Me.dgvHooks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvHooks.Location = New System.Drawing.Point(12, 145)
        Me.dgvHooks.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvHooks.Name = "dgvHooks"
        Me.dgvHooks.ReadOnly = True
        Me.dgvHooks.RowHeadersWidth = 25
        Me.dgvHooks.Size = New System.Drawing.Size(330, 17)
        Me.dgvHooks.TabIndex = 129
        '
        'cmdAddHook
        '
        Me.cmdAddHook.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddHook.Enabled = False
        Me.cmdAddHook.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddHook.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddHook.Location = New System.Drawing.Point(12, 105)
        Me.cmdAddHook.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddHook.Name = "cmdAddHook"
        Me.cmdAddHook.Size = New System.Drawing.Size(76, 33)
        Me.cmdAddHook.TabIndex = 128
        Me.cmdAddHook.Text = "&Add"
        Me.cmdAddHook.UseVisualStyleBackColor = False
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label19.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label19.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label19.ForeColor = System.Drawing.Color.White
        Me.Label19.Location = New System.Drawing.Point(0, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(345, 21)
        Me.Label19.TabIndex = 127
        Me.Label19.Text = "Hook Configuration"
        '
        'lblRestricted
        '
        Me.lblRestricted.AutoSize = True
        Me.lblRestricted.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblRestricted.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRestricted.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblRestricted.Location = New System.Drawing.Point(801, 42)
        Me.lblRestricted.Name = "lblRestricted"
        Me.lblRestricted.Size = New System.Drawing.Size(291, 17)
        Me.lblRestricted.TabIndex = 124
        Me.lblRestricted.Text = "User does not have Edit permissions in this window"
        Me.lblRestricted.Visible = False
        '
        'chkEditMode
        '
        Me.chkEditMode.AutoSize = True
        Me.chkEditMode.Location = New System.Drawing.Point(686, 44)
        Me.chkEditMode.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkEditMode.Name = "chkEditMode"
        Me.chkEditMode.Size = New System.Drawing.Size(88, 21)
        Me.chkEditMode.TabIndex = 125
        Me.chkEditMode.Text = "Edit Mode"
        Me.chkEditMode.UseVisualStyleBackColor = True
        Me.chkEditMode.Visible = False
        '
        'cmdIncreaseFont
        '
        Me.cmdIncreaseFont.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdIncreaseFont.CausesValidation = False
        Me.cmdIncreaseFont.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdIncreaseFont.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdIncreaseFont.Location = New System.Drawing.Point(204, 10)
        Me.cmdIncreaseFont.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdIncreaseFont.Name = "cmdIncreaseFont"
        Me.cmdIncreaseFont.Size = New System.Drawing.Size(75, 38)
        Me.cmdIncreaseFont.TabIndex = 126
        Me.cmdIncreaseFont.Text = "&Font +"
        Me.cmdIncreaseFont.UseVisualStyleBackColor = False
        '
        'cmdSymbol
        '
        Me.cmdSymbol.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSymbol.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSymbol.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdSymbol.Location = New System.Drawing.Point(5120, 8)
        Me.cmdSymbol.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSymbol.Name = "cmdSymbol"
        Me.cmdSymbol.Size = New System.Drawing.Size(150, 33)
        Me.cmdSymbol.TabIndex = 145
        Me.cmdSymbol.Text = "Show Symbol Copy"
        Me.cmdSymbol.UseVisualStyleBackColor = True
        '
        'pan7
        '
        Me.pan7.BackColor = System.Drawing.Color.White
        Me.pan7.Controls.Add(Me.cmdTestESig)
        Me.pan7.Controls.Add(Me.cmdRemoveRFC)
        Me.pan7.Controls.Add(Me.cmdRemoveMOS)
        Me.pan7.Controls.Add(Me.cmdAddRFC)
        Me.pan7.Controls.Add(Me.cmdAddMOS)
        Me.pan7.Controls.Add(Me.lblReasonForChange)
        Me.pan7.Controls.Add(Me.lblMeaningOfSig)
        Me.pan7.Controls.Add(Me.dgvRFC)
        Me.pan7.Controls.Add(Me.dgvMOS)
        Me.pan7.Controls.Add(Me.gbAuditTrail)
        Me.pan7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pan7.Location = New System.Drawing.Point(0, 0)
        Me.pan7.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan7.Name = "pan7"
        Me.pan7.Size = New System.Drawing.Size(561, 334)
        Me.pan7.TabIndex = 128
        '
        'cmdTestESig
        '
        Me.cmdTestESig.BackColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdTestESig.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTestESig.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdTestESig.Location = New System.Drawing.Point(215, 43)
        Me.cmdTestESig.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdTestESig.Name = "cmdTestESig"
        Me.cmdTestESig.Size = New System.Drawing.Size(124, 29)
        Me.cmdTestESig.TabIndex = 130
        Me.cmdTestESig.Text = "Test ESig Display"
        Me.cmdTestESig.UseVisualStyleBackColor = True
        '
        'cmdRemoveRFC
        '
        Me.cmdRemoveRFC.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRemoveRFC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemoveRFC.ForeColor = System.Drawing.Color.Red
        Me.cmdRemoveRFC.Location = New System.Drawing.Point(558, 335)
        Me.cmdRemoveRFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdRemoveRFC.Name = "cmdRemoveRFC"
        Me.cmdRemoveRFC.Size = New System.Drawing.Size(97, 33)
        Me.cmdRemoveRFC.TabIndex = 140
        Me.cmdRemoveRFC.Text = "&Remove"
        Me.cmdRemoveRFC.UseVisualStyleBackColor = False
        '
        'cmdRemoveMOS
        '
        Me.cmdRemoveMOS.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRemoveMOS.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemoveMOS.ForeColor = System.Drawing.Color.Red
        Me.cmdRemoveMOS.Location = New System.Drawing.Point(570, 42)
        Me.cmdRemoveMOS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdRemoveMOS.Name = "cmdRemoveMOS"
        Me.cmdRemoveMOS.Size = New System.Drawing.Size(97, 33)
        Me.cmdRemoveMOS.TabIndex = 139
        Me.cmdRemoveMOS.Text = "&Remove"
        Me.cmdRemoveMOS.UseVisualStyleBackColor = False
        '
        'cmdAddRFC
        '
        Me.cmdAddRFC.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddRFC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddRFC.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddRFC.Location = New System.Drawing.Point(495, 335)
        Me.cmdAddRFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddRFC.Name = "cmdAddRFC"
        Me.cmdAddRFC.Size = New System.Drawing.Size(56, 33)
        Me.cmdAddRFC.TabIndex = 138
        Me.cmdAddRFC.Text = "&Add"
        Me.cmdAddRFC.UseVisualStyleBackColor = False
        '
        'cmdAddMOS
        '
        Me.cmdAddMOS.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddMOS.CausesValidation = False
        Me.cmdAddMOS.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddMOS.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddMOS.Location = New System.Drawing.Point(507, 42)
        Me.cmdAddMOS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddMOS.Name = "cmdAddMOS"
        Me.cmdAddMOS.Size = New System.Drawing.Size(56, 33)
        Me.cmdAddMOS.TabIndex = 137
        Me.cmdAddMOS.Text = "&Add"
        Me.cmdAddMOS.UseVisualStyleBackColor = False
        '
        'lblReasonForChange
        '
        Me.lblReasonForChange.AutoSize = True
        Me.lblReasonForChange.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblReasonForChange.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReasonForChange.ForeColor = System.Drawing.Color.White
        Me.lblReasonForChange.Location = New System.Drawing.Point(355, 351)
        Me.lblReasonForChange.Name = "lblReasonForChange"
        Me.lblReasonForChange.Size = New System.Drawing.Size(126, 17)
        Me.lblReasonForChange.TabIndex = 5
        Me.lblReasonForChange.Text = "Reason For Change"
        '
        'lblMeaningOfSig
        '
        Me.lblMeaningOfSig.AutoSize = True
        Me.lblMeaningOfSig.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblMeaningOfSig.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMeaningOfSig.ForeColor = System.Drawing.Color.White
        Me.lblMeaningOfSig.Location = New System.Drawing.Point(355, 58)
        Me.lblMeaningOfSig.Name = "lblMeaningOfSig"
        Me.lblMeaningOfSig.Size = New System.Drawing.Size(142, 17)
        Me.lblMeaningOfSig.TabIndex = 4
        Me.lblMeaningOfSig.Text = "Meaning of Signature"
        '
        'dgvRFC
        '
        Me.dgvRFC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvRFC.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle14.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle14.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle14.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle14.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle14.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvRFC.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle14
        Me.dgvRFC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRFC.Location = New System.Drawing.Point(355, 369)
        Me.dgvRFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvRFC.Name = "dgvRFC"
        Me.dgvRFC.Size = New System.Drawing.Size(189, 234)
        Me.dgvRFC.TabIndex = 3
        '
        'dgvMOS
        '
        Me.dgvMOS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMOS.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle15.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle15.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle15.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle15.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle15.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvMOS.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle15
        Me.dgvMOS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMOS.Location = New System.Drawing.Point(355, 76)
        Me.dgvMOS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvMOS.Name = "dgvMOS"
        Me.dgvMOS.Size = New System.Drawing.Size(189, 234)
        Me.dgvMOS.TabIndex = 2
        '
        'gbAuditTrail
        '
        Me.gbAuditTrail.Controls.Add(Me.gbReasonForChange)
        Me.gbAuditTrail.Controls.Add(Me.rbAuditTrailOff)
        Me.gbAuditTrail.Controls.Add(Me.gbESig)
        Me.gbAuditTrail.Controls.Add(Me.rbAuditTrailOn)
        Me.gbAuditTrail.Location = New System.Drawing.Point(19, 67)
        Me.gbAuditTrail.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbAuditTrail.Name = "gbAuditTrail"
        Me.gbAuditTrail.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbAuditTrail.Size = New System.Drawing.Size(320, 417)
        Me.gbAuditTrail.TabIndex = 1
        Me.gbAuditTrail.TabStop = False
        Me.gbAuditTrail.Text = "Audit Trail Options"
        '
        'gbReasonForChange
        '
        Me.gbReasonForChange.Controls.Add(Me.panRFCOptions)
        Me.gbReasonForChange.Location = New System.Drawing.Point(21, 61)
        Me.gbReasonForChange.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbReasonForChange.Name = "gbReasonForChange"
        Me.gbReasonForChange.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbReasonForChange.Size = New System.Drawing.Size(274, 95)
        Me.gbReasonForChange.TabIndex = 133
        Me.gbReasonForChange.TabStop = False
        Me.gbReasonForChange.Text = "Reason for Change Options"
        '
        'panRFCOptions
        '
        Me.panRFCOptions.Controls.Add(Me.chkReasonFreeForm)
        Me.panRFCOptions.Controls.Add(Me.chkReasonForChange)
        Me.panRFCOptions.Location = New System.Drawing.Point(41, 25)
        Me.panRFCOptions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panRFCOptions.Name = "panRFCOptions"
        Me.panRFCOptions.Size = New System.Drawing.Size(222, 64)
        Me.panRFCOptions.TabIndex = 3
        '
        'chkReasonFreeForm
        '
        Me.chkReasonFreeForm.AutoSize = True
        Me.chkReasonFreeForm.Location = New System.Drawing.Point(27, 34)
        Me.chkReasonFreeForm.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkReasonFreeForm.Name = "chkReasonFreeForm"
        Me.chkReasonFreeForm.Size = New System.Drawing.Size(198, 21)
        Me.chkReasonFreeForm.TabIndex = 4
        Me.chkReasonFreeForm.Text = "Restrict to dropdown choices"
        Me.chkReasonFreeForm.UseVisualStyleBackColor = True
        '
        'chkReasonForChange
        '
        Me.chkReasonForChange.AutoSize = True
        Me.chkReasonForChange.Location = New System.Drawing.Point(0, 4)
        Me.chkReasonForChange.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkReasonForChange.Name = "chkReasonForChange"
        Me.chkReasonForChange.Size = New System.Drawing.Size(141, 21)
        Me.chkReasonForChange.TabIndex = 2
        Me.chkReasonForChange.Text = "Reason For Change"
        Me.chkReasonForChange.UseVisualStyleBackColor = True
        '
        'rbAuditTrailOff
        '
        Me.rbAuditTrailOff.AutoSize = True
        Me.rbAuditTrailOff.Checked = True
        Me.rbAuditTrailOff.Location = New System.Drawing.Point(83, 31)
        Me.rbAuditTrailOff.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbAuditTrailOff.Name = "rbAuditTrailOff"
        Me.rbAuditTrailOff.Size = New System.Drawing.Size(44, 21)
        Me.rbAuditTrailOff.TabIndex = 2
        Me.rbAuditTrailOff.TabStop = True
        Me.rbAuditTrailOff.Text = "Off"
        Me.rbAuditTrailOff.UseVisualStyleBackColor = True
        '
        'gbESig
        '
        Me.gbESig.Controls.Add(Me.panESigOptions)
        Me.gbESig.Controls.Add(Me.rbESigOff)
        Me.gbESig.Controls.Add(Me.rbESigOn)
        Me.gbESig.Location = New System.Drawing.Point(21, 165)
        Me.gbESig.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbESig.Name = "gbESig"
        Me.gbESig.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbESig.Size = New System.Drawing.Size(274, 234)
        Me.gbESig.TabIndex = 0
        Me.gbESig.TabStop = False
        Me.gbESig.Text = "Electronic Signature Options"
        '
        'panESigOptions
        '
        Me.panESigOptions.Controls.Add(Me.chkSigFreeForm)
        Me.panESigOptions.Controls.Add(Me.Label21)
        Me.panESigOptions.Controls.Add(Me.gbUserIDType)
        Me.panESigOptions.Controls.Add(Me.chkMeaningOfSign)
        Me.panESigOptions.Location = New System.Drawing.Point(17, 98)
        Me.panESigOptions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panESigOptions.Name = "panESigOptions"
        Me.panESigOptions.Size = New System.Drawing.Size(250, 119)
        Me.panESigOptions.TabIndex = 2
        '
        'chkSigFreeForm
        '
        Me.chkSigFreeForm.AutoSize = True
        Me.chkSigFreeForm.Location = New System.Drawing.Point(49, 80)
        Me.chkSigFreeForm.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkSigFreeForm.Name = "chkSigFreeForm"
        Me.chkSigFreeForm.Size = New System.Drawing.Size(198, 21)
        Me.chkSigFreeForm.TabIndex = 3
        Me.chkSigFreeForm.Text = "Restrict to dropdown choices"
        Me.chkSigFreeForm.UseVisualStyleBackColor = True
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(16, 22)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(131, 17)
        Me.Label21.TabIndex = 2
        Me.Label21.Text = "Include the following:"
        '
        'gbUserIDType
        '
        Me.gbUserIDType.Controls.Add(Me.rbUserIDChoice)
        Me.gbUserIDType.Controls.Add(Me.rbOnlyLoggedOn)
        Me.gbUserIDType.Location = New System.Drawing.Point(138, 16)
        Me.gbUserIDType.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbUserIDType.Name = "gbUserIDType"
        Me.gbUserIDType.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbUserIDType.Size = New System.Drawing.Size(97, 39)
        Me.gbUserIDType.TabIndex = 2
        Me.gbUserIDType.TabStop = False
        Me.gbUserIDType.Text = "User ID Option"
        Me.gbUserIDType.Visible = False
        '
        'rbUserIDChoice
        '
        Me.rbUserIDChoice.AutoSize = True
        Me.rbUserIDChoice.Location = New System.Drawing.Point(7, 55)
        Me.rbUserIDChoice.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbUserIDChoice.Name = "rbUserIDChoice"
        Me.rbUserIDChoice.Size = New System.Drawing.Size(153, 21)
        Me.rbUserIDChoice.TabIndex = 2
        Me.rbUserIDChoice.Text = "Allow Choice of Users"
        Me.rbUserIDChoice.UseVisualStyleBackColor = True
        '
        'rbOnlyLoggedOn
        '
        Me.rbOnlyLoggedOn.AutoSize = True
        Me.rbOnlyLoggedOn.Checked = True
        Me.rbOnlyLoggedOn.Location = New System.Drawing.Point(7, 25)
        Me.rbOnlyLoggedOn.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbOnlyLoggedOn.Name = "rbOnlyLoggedOn"
        Me.rbOnlyLoggedOn.Size = New System.Drawing.Size(153, 21)
        Me.rbOnlyLoggedOn.TabIndex = 1
        Me.rbOnlyLoggedOn.TabStop = True
        Me.rbOnlyLoggedOn.Text = "Only Logged On User"
        Me.rbOnlyLoggedOn.UseVisualStyleBackColor = True
        '
        'chkMeaningOfSign
        '
        Me.chkMeaningOfSign.AutoSize = True
        Me.chkMeaningOfSign.Location = New System.Drawing.Point(27, 50)
        Me.chkMeaningOfSign.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkMeaningOfSign.Name = "chkMeaningOfSign"
        Me.chkMeaningOfSign.Size = New System.Drawing.Size(155, 21)
        Me.chkMeaningOfSign.TabIndex = 1
        Me.chkMeaningOfSign.Text = "Meaning Of Signature"
        Me.chkMeaningOfSign.UseVisualStyleBackColor = True
        '
        'rbESigOff
        '
        Me.rbESigOff.AutoSize = True
        Me.rbESigOff.Checked = True
        Me.rbESigOff.Location = New System.Drawing.Point(41, 60)
        Me.rbESigOff.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbESigOff.Name = "rbESigOff"
        Me.rbESigOff.Size = New System.Drawing.Size(196, 21)
        Me.rbESigOff.TabIndex = 1
        Me.rbESigOff.TabStop = True
        Me.rbESigOff.Text = "Off (Record Silent Audit Trail)"
        Me.rbESigOff.UseVisualStyleBackColor = True
        '
        'rbESigOn
        '
        Me.rbESigOn.AutoSize = True
        Me.rbESigOn.Location = New System.Drawing.Point(41, 30)
        Me.rbESigOn.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbESigOn.Name = "rbESigOn"
        Me.rbESigOn.Size = New System.Drawing.Size(177, 21)
        Me.rbESigOn.TabIndex = 0
        Me.rbESigOn.Text = "On (Require ESig prompt)"
        Me.rbESigOn.UseVisualStyleBackColor = True
        '
        'rbAuditTrailOn
        '
        Me.rbAuditTrailOn.AutoSize = True
        Me.rbAuditTrailOn.Location = New System.Drawing.Point(21, 31)
        Me.rbAuditTrailOn.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbAuditTrailOn.Name = "rbAuditTrailOn"
        Me.rbAuditTrailOn.Size = New System.Drawing.Size(43, 21)
        Me.rbAuditTrailOn.TabIndex = 1
        Me.rbAuditTrailOn.Text = "On"
        Me.rbAuditTrailOn.UseVisualStyleBackColor = True
        '
        'lblCC
        '
        Me.lblCC.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblCC.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblCC.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblCC.ForeColor = System.Drawing.Color.White
        Me.lblCC.Location = New System.Drawing.Point(0, 0)
        Me.lblCC.Name = "lblCC"
        Me.lblCC.Size = New System.Drawing.Size(561, 21)
        Me.lblCC.TabIndex = 136
        Me.lblCC.Text = "Compliance Configuration"
        '
        'lblOpen
        '
        Me.lblOpen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOpen.Location = New System.Drawing.Point(1313, 3)
        Me.lblOpen.Name = "lblOpen"
        Me.lblOpen.Size = New System.Drawing.Size(117, 69)
        Me.lblOpen.TabIndex = 129
        '
        'panFC
        '
        Me.panFC.AutoScroll = True
        Me.panFC.BackColor = System.Drawing.Color.White
        Me.panFC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panFC.Controls.Add(Me.lblFieldCodes)
        Me.panFC.Controls.Add(Me.panFCcmd)
        Me.panFC.Controls.Add(Me.dgvFC)
        Me.panFC.Location = New System.Drawing.Point(1179, 35)
        Me.panFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panFC.Name = "panFC"
        Me.panFC.Size = New System.Drawing.Size(611, 202)
        Me.panFC.TabIndex = 130
        Me.panFC.Visible = False
        '
        'lblFieldCodes
        '
        Me.lblFieldCodes.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblFieldCodes.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblFieldCodes.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFieldCodes.ForeColor = System.Drawing.Color.White
        Me.lblFieldCodes.Location = New System.Drawing.Point(0, 0)
        Me.lblFieldCodes.Name = "lblFieldCodes"
        Me.lblFieldCodes.Size = New System.Drawing.Size(609, 21)
        Me.lblFieldCodes.TabIndex = 137
        Me.lblFieldCodes.Text = "Custom Field Codes"
        '
        'panFCcmd
        '
        Me.panFCcmd.CausesValidation = False
        Me.panFCcmd.Controls.Add(Me.cmdRemoveFC)
        Me.panFCcmd.Controls.Add(Me.cmdResetFC)
        Me.panFCcmd.Controls.Add(Me.cmdAddFC)
        Me.panFCcmd.Location = New System.Drawing.Point(7, 25)
        Me.panFCcmd.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panFCcmd.Name = "panFCcmd"
        Me.panFCcmd.Size = New System.Drawing.Size(285, 43)
        Me.panFCcmd.TabIndex = 3
        '
        'cmdRemoveFC
        '
        Me.cmdRemoveFC.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRemoveFC.CausesValidation = False
        Me.cmdRemoveFC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemoveFC.ForeColor = System.Drawing.Color.Red
        Me.cmdRemoveFC.Location = New System.Drawing.Point(88, 4)
        Me.cmdRemoveFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdRemoveFC.Name = "cmdRemoveFC"
        Me.cmdRemoveFC.Size = New System.Drawing.Size(97, 33)
        Me.cmdRemoveFC.TabIndex = 140
        Me.cmdRemoveFC.Text = "&Remove"
        Me.cmdRemoveFC.UseVisualStyleBackColor = False
        '
        'cmdResetFC
        '
        Me.cmdResetFC.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdResetFC.CausesValidation = False
        Me.cmdResetFC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetFC.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdResetFC.Location = New System.Drawing.Point(194, 4)
        Me.cmdResetFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdResetFC.Name = "cmdResetFC"
        Me.cmdResetFC.Size = New System.Drawing.Size(85, 33)
        Me.cmdResetFC.TabIndex = 92
        Me.cmdResetFC.Text = "Reset"
        Me.cmdResetFC.UseVisualStyleBackColor = False
        '
        'cmdAddFC
        '
        Me.cmdAddFC.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddFC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddFC.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddFC.Location = New System.Drawing.Point(0, 4)
        Me.cmdAddFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddFC.Name = "cmdAddFC"
        Me.cmdAddFC.Size = New System.Drawing.Size(79, 33)
        Me.cmdAddFC.TabIndex = 8
        Me.cmdAddFC.Text = "&Add"
        Me.cmdAddFC.UseVisualStyleBackColor = False
        '
        'dgvFC
        '
        Me.dgvFC.AllowUserToDeleteRows = False
        Me.dgvFC.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvFC.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle16.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFC.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle16
        Me.dgvFC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFC.Location = New System.Drawing.Point(7, 70)
        Me.dgvFC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvFC.Name = "dgvFC"
        Me.dgvFC.ReadOnly = True
        Me.dgvFC.Size = New System.Drawing.Size(594, 126)
        Me.dgvFC.TabIndex = 2
        '
        'cmdDecreaseFont
        '
        Me.cmdDecreaseFont.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdDecreaseFont.CausesValidation = False
        Me.cmdDecreaseFont.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDecreaseFont.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdDecreaseFont.Location = New System.Drawing.Point(280, 10)
        Me.cmdDecreaseFont.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdDecreaseFont.Name = "cmdDecreaseFont"
        Me.cmdDecreaseFont.Size = New System.Drawing.Size(75, 38)
        Me.cmdDecreaseFont.TabIndex = 131
        Me.cmdDecreaseFont.Text = "&Font -"
        Me.cmdDecreaseFont.UseVisualStyleBackColor = False
        '
        'pan7b
        '
        Me.pan7b.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan7b.Controls.Add(Me.pan7a)
        Me.pan7b.Controls.Add(Me.lblCC)
        Me.pan7b.Location = New System.Drawing.Point(1153, 194)
        Me.pan7b.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan7b.Name = "pan7b"
        Me.pan7b.Size = New System.Drawing.Size(563, 357)
        Me.pan7b.TabIndex = 132
        '
        'pan7a
        '
        Me.pan7a.Controls.Add(Me.pan7)
        Me.pan7a.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pan7a.Location = New System.Drawing.Point(0, 21)
        Me.pan7a.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan7a.Name = "pan7a"
        Me.pan7a.Size = New System.Drawing.Size(561, 334)
        Me.pan7a.TabIndex = 137
        '
        'cmdRemovePM
        '
        Me.cmdRemovePM.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRemovePM.CausesValidation = False
        Me.cmdRemovePM.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemovePM.ForeColor = System.Drawing.Color.Red
        Me.cmdRemovePM.Location = New System.Drawing.Point(5, 8)
        Me.cmdRemovePM.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdRemovePM.Name = "cmdRemovePM"
        Me.cmdRemovePM.Size = New System.Drawing.Size(195, 33)
        Me.cmdRemovePM.TabIndex = 141
        Me.cmdRemovePM.Text = "&Delete Group..."
        Me.cmdRemovePM.UseVisualStyleBackColor = False
        '
        'cmdAddPM
        '
        Me.cmdAddPM.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddPM.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddPM.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddPM.Location = New System.Drawing.Point(5, 48)
        Me.cmdAddPM.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAddPM.Name = "cmdAddPM"
        Me.cmdAddPM.Size = New System.Drawing.Size(195, 33)
        Me.cmdAddPM.TabIndex = 124
        Me.cmdAddPM.Text = "&Add New Group"
        Me.cmdAddPM.UseVisualStyleBackColor = False
        '
        'pan8
        '
        Me.pan8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pan8.AutoScroll = True
        Me.pan8.BackColor = System.Drawing.Color.White
        Me.pan8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan8.Controls.Add(Me.lblPM)
        Me.pan8.Controls.Add(Me.panPM)
        Me.pan8.Location = New System.Drawing.Point(660, 281)
        Me.pan8.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.pan8.Name = "pan8"
        Me.pan8.Size = New System.Drawing.Size(702, 248)
        Me.pan8.TabIndex = 134
        Me.pan8.Visible = False
        '
        'lblPM
        '
        Me.lblPM.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblPM.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblPM.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblPM.ForeColor = System.Drawing.Color.White
        Me.lblPM.Location = New System.Drawing.Point(0, 0)
        Me.lblPM.Name = "lblPM"
        Me.lblPM.Size = New System.Drawing.Size(700, 21)
        Me.lblPM.TabIndex = 134
        Me.lblPM.Text = "Permissions Manager"
        '
        'panPM
        '
        Me.panPM.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panPM.AutoScroll = True
        Me.panPM.Controls.Add(Me.lvPermissionsFinalReport)
        Me.panPM.Controls.Add(Me.lvPermissionsReportTemplate)
        Me.panPM.Controls.Add(Me.lvPermissionsAdmin)
        Me.panPM.Controls.Add(Me.dgvPermissions)
        Me.panPM.Controls.Add(Me.lblDo)
        Me.panPM.Controls.Add(Me.lblBase)
        Me.panPM.Controls.Add(Me.lblPermissions)
        Me.panPM.Controls.Add(Me.cmdAddPM)
        Me.panPM.Controls.Add(Me.lbllbx1)
        Me.panPM.Controls.Add(Me.lblS)
        Me.panPM.Controls.Add(Me.cmdSelectAllPermissions)
        Me.panPM.Controls.Add(Me.cmdRemovePM)
        Me.panPM.Controls.Add(Me.cmdDeselectAllPermissions)
        Me.panPM.Controls.Add(Me.lblcbxModulesPers)
        Me.panPM.Controls.Add(Me.cbxPermBase)
        Me.panPM.Controls.Add(Me.lbx1)
        Me.panPM.Controls.Add(Me.lvPermissions)
        Me.panPM.Location = New System.Drawing.Point(3, 29)
        Me.panPM.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panPM.Name = "panPM"
        Me.panPM.Size = New System.Drawing.Size(682, 208)
        Me.panPM.TabIndex = 150
        '
        'lvPermissionsFinalReport
        '
        Me.lvPermissionsFinalReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvPermissionsFinalReport.Enabled = False
        Me.lvPermissionsFinalReport.Location = New System.Drawing.Point(348, 86)
        Me.lvPermissionsFinalReport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lvPermissionsFinalReport.MultiSelect = False
        Me.lvPermissionsFinalReport.Name = "lvPermissionsFinalReport"
        Me.lvPermissionsFinalReport.Size = New System.Drawing.Size(105, 53)
        Me.lvPermissionsFinalReport.TabIndex = 153
        Me.lvPermissionsFinalReport.UseCompatibleStateImageBehavior = False
        '
        'lvPermissionsReportTemplate
        '
        Me.lvPermissionsReportTemplate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvPermissionsReportTemplate.Enabled = False
        Me.lvPermissionsReportTemplate.Location = New System.Drawing.Point(340, 78)
        Me.lvPermissionsReportTemplate.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lvPermissionsReportTemplate.MultiSelect = False
        Me.lvPermissionsReportTemplate.Name = "lvPermissionsReportTemplate"
        Me.lvPermissionsReportTemplate.Size = New System.Drawing.Size(105, 53)
        Me.lvPermissionsReportTemplate.TabIndex = 152
        Me.lvPermissionsReportTemplate.UseCompatibleStateImageBehavior = False
        '
        'dgvPermissions
        '
        Me.dgvPermissions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvPermissions.BackgroundColor = System.Drawing.Color.White
        Me.dgvPermissions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPermissions.Location = New System.Drawing.Point(8, 112)
        Me.dgvPermissions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvPermissions.MultiSelect = False
        Me.dgvPermissions.Name = "dgvPermissions"
        Me.dgvPermissions.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvPermissions.Size = New System.Drawing.Size(191, 72)
        Me.dgvPermissions.TabIndex = 151
        '
        'lblDo
        '
        Me.lblDo.BackColor = System.Drawing.Color.Transparent
        Me.lblDo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDo.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDo.ForeColor = System.Drawing.Color.Blue
        Me.lblDo.Location = New System.Drawing.Point(215, 8)
        Me.lblDo.Name = "lblDo"
        Me.lblDo.Size = New System.Drawing.Size(193, 73)
        Me.lblDo.TabIndex = 150
        Me.lblDo.Text = "Click Save or Cancel to finish adding group"
        Me.lblDo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblDo.Visible = False
        '
        'lblBase
        '
        Me.lblBase.BackColor = System.Drawing.Color.Transparent
        Me.lblBase.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBase.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblBase.Location = New System.Drawing.Point(491, 9)
        Me.lblBase.Name = "lblBase"
        Me.lblBase.Size = New System.Drawing.Size(226, 41)
        Me.lblBase.TabIndex = 148
        Me.lblBase.Text = "Make selections based on existing Permissions Group below"
        Me.lblBase.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblPermissions
        '
        Me.lblPermissions.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPermissions.AutoSize = True
        Me.lblPermissions.ForeColor = System.Drawing.Color.Blue
        Me.lblPermissions.Location = New System.Drawing.Point(1500, 127)
        Me.lblPermissions.Name = "lblPermissions"
        Me.lblPermissions.Size = New System.Drawing.Size(380, 17)
        Me.lblPermissions.TabIndex = 149
        Me.lblPermissions.Text = "User ID Permissions (Check = Allow Editing in that Tab/Window)"
        Me.lblPermissions.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lbllbx1
        '
        Me.lbllbx1.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lbllbx1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbllbx1.ForeColor = System.Drawing.Color.White
        Me.lbllbx1.Location = New System.Drawing.Point(8, 94)
        Me.lbllbx1.Margin = New System.Windows.Forms.Padding(0)
        Me.lbllbx1.Name = "lbllbx1"
        Me.lbllbx1.Size = New System.Drawing.Size(191, 17)
        Me.lbllbx1.TabIndex = 149
        Me.lbllbx1.Text = "Choose a Permissions Group"
        Me.lbllbx1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblS
        '
        Me.lblS.AutoSize = True
        Me.lblS.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblS.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblS.ForeColor = System.Drawing.Color.White
        Me.lblS.Location = New System.Drawing.Point(377, 130)
        Me.lblS.Name = "lblS"
        Me.lblS.Size = New System.Drawing.Size(106, 17)
        Me.lblS.TabIndex = 147
        Me.lblS.Text = "Make selections"
        '
        'cbxPermBase
        '
        Me.cbxPermBase.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cbxPermBase.FormattingEnabled = True
        Me.cbxPermBase.Location = New System.Drawing.Point(491, 55)
        Me.cbxPermBase.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxPermBase.Name = "cbxPermBase"
        Me.cbxPermBase.Size = New System.Drawing.Size(226, 25)
        Me.cbxPermBase.TabIndex = 147
        '
        'lbx1
        '
        Me.lbx1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbx1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbx1.FormattingEnabled = True
        Me.lbx1.ItemHeight = 17
        Me.lbx1.Location = New System.Drawing.Point(206, 112)
        Me.lbx1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lbx1.Name = "lbx1"
        Me.lbx1.Size = New System.Drawing.Size(164, 72)
        Me.lbx1.TabIndex = 146
        '
        'frmAdministration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1752, 1054)
        Me.ControlBox = False
        Me.Controls.Add(Me.pan4)
        Me.Controls.Add(Me.pan3)
        Me.Controls.Add(Me.pan2)
        Me.Controls.Add(Me.pan1)
        Me.Controls.Add(Me.pan6)
        Me.Controls.Add(Me.panFC)
        Me.Controls.Add(Me.cmdIncreaseFont)
        Me.Controls.Add(Me.pan7b)
        Me.Controls.Add(Me.cmdSymbol)
        Me.Controls.Add(Me.pan8)
        Me.Controls.Add(Me.pan5)
        Me.Controls.Add(Me.cmdDecreaseFont)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.pb1)
        Me.Controls.Add(Me.lblOpen)
        Me.Controls.Add(Me.chkEditMode)
        Me.Controls.Add(Me.lblRestricted)
        Me.Controls.Add(Me.lblcbxModules)
        Me.Controls.Add(Me.cbxModules)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lbxTab1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MinimizeBox = False
        Me.Name = "frmAdministration"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "  LABIntegrity StudyDoc Administration"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pan1.ResumeLayout(False)
        Me.pan1.PerformLayout()
        Me.gbUA.ResumeLayout(False)
        Me.gbUA.PerformLayout()
        Me.gbWatsonAccount.ResumeLayout(False)
        Me.gbWatsonAccount.PerformLayout()
        Me.gbWindowsAuth.ResumeLayout(False)
        Me.gbWindowsAuth.PerformLayout()
        Me.gbLDAP.ResumeLayout(False)
        Me.gbLDAP.PerformLayout()
        Me.panLDAP.ResumeLayout(False)
        Me.panLDAP.PerformLayout()
        Me.gbSetPerm.ResumeLayout(False)
        Me.gbxPassword.ResumeLayout(False)
        Me.gbxPassword.PerformLayout()
        Me.gbGlobalParams.ResumeLayout(False)
        Me.gbGlobalParams.PerformLayout()
        Me.gbUserShow.ResumeLayout(False)
        Me.gbUserShow.PerformLayout()
        CType(Me.dgvUserAttributes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvUsers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan2.ResumeLayout(False)
        Me.pan2.PerformLayout()
        CType(Me.dgvDropdownboxTitle, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvDropdownboxContents, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan3.ResumeLayout(False)
        Me.pan3.PerformLayout()
        Me.gbxlblCorporateAdderesses.ResumeLayout(False)
        Me.gbxlblCorporateAdderesses.PerformLayout()
        Me.gbDropDown.ResumeLayout(False)
        Me.gbDropDown.PerformLayout()
        CType(Me.dgvCorporateAddresses, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvNickNames, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan4.ResumeLayout(False)
        Me.pan4.PerformLayout()
        Me.gbSTD.ResumeLayout(False)
        CType(Me.dgvTemplates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbStudyTemplates.ResumeLayout(False)
        Me.gbStudyTemplates.PerformLayout()
        CType(Me.dgvTemplateAttributes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan5.ResumeLayout(False)
        Me.pan5.PerformLayout()
        Me.gbxlabelGlobalParameters.ResumeLayout(False)
        Me.panGP.ResumeLayout(False)
        Me.panGP.PerformLayout()
        CType(Me.dgvGlobal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan6.ResumeLayout(False)
        CType(Me.dgvHooks, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan7.ResumeLayout(False)
        Me.pan7.PerformLayout()
        CType(Me.dgvRFC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvMOS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbAuditTrail.ResumeLayout(False)
        Me.gbAuditTrail.PerformLayout()
        Me.gbReasonForChange.ResumeLayout(False)
        Me.panRFCOptions.ResumeLayout(False)
        Me.panRFCOptions.PerformLayout()
        Me.gbESig.ResumeLayout(False)
        Me.gbESig.PerformLayout()
        Me.panESigOptions.ResumeLayout(False)
        Me.panESigOptions.PerformLayout()
        Me.gbUserIDType.ResumeLayout(False)
        Me.gbUserIDType.PerformLayout()
        Me.panFC.ResumeLayout(False)
        Me.panFCcmd.ResumeLayout(False)
        CType(Me.dgvFC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pan7b.ResumeLayout(False)
        Me.pan7a.ResumeLayout(False)
        Me.pan8.ResumeLayout(False)
        Me.panPM.ResumeLayout(False)
        Me.panPM.PerformLayout()
        CType(Me.dgvPermissions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lbxTab1 As System.Windows.Forms.ListBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lblcbxModules As System.Windows.Forms.Label
    Friend WithEvents cbxModules As System.Windows.Forms.ComboBox
    Friend WithEvents pan1 As System.Windows.Forms.Panel
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents gbxPassword As System.Windows.Forms.GroupBox
    Friend WithEvents chkAccountIsLockedOut As System.Windows.Forms.CheckBox
    Friend WithEvents chkPasswordNeverExpires As System.Windows.Forms.CheckBox
    Friend WithEvents chkUserCannotChangePassword As System.Windows.Forms.CheckBox
    Friend WithEvents chkChangePasswordAtNextLogon As System.Windows.Forms.CheckBox
    Friend WithEvents lvPermissions As System.Windows.Forms.ListView
    Friend WithEvents cmdEnterPassword As System.Windows.Forms.Button
    Friend WithEvents cmdResetUserAccounts As System.Windows.Forms.Button
    Friend WithEvents cmdAddUserID As System.Windows.Forms.Button
    Friend WithEvents cmdAddUser As System.Windows.Forms.Button
    Friend WithEvents gbGlobalParams As System.Windows.Forms.GroupBox
    Friend WithEvents rbShowInactiveUserIDs As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowActiveUserIDs As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowAllUserIDs As System.Windows.Forms.RadioButton
    Friend WithEvents gbUserShow As System.Windows.Forms.GroupBox
    Friend WithEvents rbShowInactiveUsers As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowActiveUsers As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowAllUsers As System.Windows.Forms.RadioButton
    Friend WithEvents dgvUserAttributes As System.Windows.Forms.DataGridView
    Friend WithEvents dgvUsers As System.Windows.Forms.DataGridView
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents pan2 As System.Windows.Forms.Panel
    Friend WithEvents dgvDropdownboxTitle As System.Windows.Forms.DataGridView
    Friend WithEvents cmdOrderDropdownbox As System.Windows.Forms.Button
    Friend WithEvents dgvDropdownboxContents As System.Windows.Forms.DataGridView
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cmdAddDropdownbox As System.Windows.Forms.Button
    Friend WithEvents cmdResetDropdownbox As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pan3 As System.Windows.Forms.Panel
    Friend WithEvents gbDropDown As System.Windows.Forms.GroupBox
    Friend WithEvents rbShowInactiveAddresses As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowActiveAddresses As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowAllAddresses As System.Windows.Forms.RadioButton
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cmdAddCorporateAddress As System.Windows.Forms.Button
    Friend WithEvents dgvCorporateAddresses As System.Windows.Forms.DataGridView
    Friend WithEvents dgvNickNames As System.Windows.Forms.DataGridView
    Friend WithEvents cmdResetCorporateAddressses As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pan4 As System.Windows.Forms.Panel
    Friend WithEvents dgvTemplates As System.Windows.Forms.DataGridView
    Friend WithEvents lblTExpl As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmdAddTemplate As System.Windows.Forms.Button
    Friend WithEvents gbStudyTemplates As System.Windows.Forms.GroupBox
    Friend WithEvents rbShowInactiveTemplates As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowActiveTemplates As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowAllTemplates As System.Windows.Forms.RadioButton
    Friend WithEvents dgvTemplateAttributes As System.Windows.Forms.DataGridView
    Friend WithEvents cmdResetDefineReports As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents pan5 As System.Windows.Forms.Panel
    Friend WithEvents cmdResetGlobal As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseGlobal As System.Windows.Forms.Button
    Friend WithEvents lblGlobalValues As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents dgvGlobal As System.Windows.Forms.DataGridView
    Friend WithEvents lbxGlobal As System.Windows.Forms.ListBox
    Friend WithEvents lblGlobalParameters As System.Windows.Forms.Label
    Friend WithEvents lblIntegrity As System.Windows.Forms.Label
    Friend WithEvents pan6 As System.Windows.Forms.Panel
    Friend WithEvents cmdRefreshHook As System.Windows.Forms.Button
    Friend WithEvents cmdResetHooks As System.Windows.Forms.Button
    Friend WithEvents dgvHooks As System.Windows.Forms.DataGridView
    Friend WithEvents cmdAddHook As System.Windows.Forms.Button
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents lblcbxModulesPers As System.Windows.Forms.Label
    Friend WithEvents lvPermissionsAdmin As System.Windows.Forms.ListView
    Friend WithEvents cmdDeselectAllPermissions As System.Windows.Forms.Button
    Friend WithEvents cmdSelectAllPermissions As System.Windows.Forms.Button
    Friend WithEvents lblRestricted As System.Windows.Forms.Label
    Friend WithEvents chkEditMode As System.Windows.Forms.CheckBox
    Friend WithEvents cmdIncreaseFont As System.Windows.Forms.Button
    Friend WithEvents pan7 As System.Windows.Forms.Panel
    Friend WithEvents lblCC As System.Windows.Forms.Label
    Friend WithEvents lblReasonForChange As System.Windows.Forms.Label
    Friend WithEvents lblMeaningOfSig As System.Windows.Forms.Label
    Friend WithEvents dgvRFC As System.Windows.Forms.DataGridView
    Friend WithEvents dgvMOS As System.Windows.Forms.DataGridView
    Friend WithEvents gbAuditTrail As System.Windows.Forms.GroupBox
    Friend WithEvents rbAuditTrailOff As System.Windows.Forms.RadioButton
    Friend WithEvents gbESig As System.Windows.Forms.GroupBox
    Friend WithEvents panESigOptions As System.Windows.Forms.Panel
    Friend WithEvents chkReasonFreeForm As System.Windows.Forms.CheckBox
    Friend WithEvents chkSigFreeForm As System.Windows.Forms.CheckBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents gbUserIDType As System.Windows.Forms.GroupBox
    Friend WithEvents rbUserIDChoice As System.Windows.Forms.RadioButton
    Friend WithEvents rbOnlyLoggedOn As System.Windows.Forms.RadioButton
    Friend WithEvents chkReasonForChange As System.Windows.Forms.CheckBox
    Friend WithEvents chkMeaningOfSign As System.Windows.Forms.CheckBox
    Friend WithEvents rbESigOff As System.Windows.Forms.RadioButton
    Friend WithEvents rbESigOn As System.Windows.Forms.RadioButton
    Friend WithEvents rbAuditTrailOn As System.Windows.Forms.RadioButton
    Friend WithEvents cmdRemoveRFC As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveMOS As System.Windows.Forms.Button
    Friend WithEvents cmdAddRFC As System.Windows.Forms.Button
    Friend WithEvents cmdAddMOS As System.Windows.Forms.Button
    Friend WithEvents lblOpen As System.Windows.Forms.Label
    Friend WithEvents cmdTestESig As System.Windows.Forms.Button
    Friend WithEvents panRFCOptions As System.Windows.Forms.Panel
    Friend WithEvents panFC As System.Windows.Forms.Panel
    Friend WithEvents lblFieldCodes As System.Windows.Forms.Label
    Friend WithEvents panFCcmd As System.Windows.Forms.Panel
    Friend WithEvents cmdAddFC As System.Windows.Forms.Button
    Friend WithEvents dgvFC As System.Windows.Forms.DataGridView
    Friend WithEvents cmdResetFC As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveFC As System.Windows.Forms.Button
    Friend WithEvents cmdDecreaseFont As System.Windows.Forms.Button
    Friend WithEvents pan7b As System.Windows.Forms.Panel
    Friend WithEvents pan7a As System.Windows.Forms.Panel
    Friend WithEvents gbReasonForChange As System.Windows.Forms.GroupBox
    Friend WithEvents pan8 As System.Windows.Forms.Panel
    Friend WithEvents cmdSymbol As System.Windows.Forms.Button
    Friend WithEvents cmdRemovePM As System.Windows.Forms.Button
    Friend WithEvents cmdAddPM As System.Windows.Forms.Button
    Friend WithEvents lblPM As System.Windows.Forms.Label
    Friend WithEvents gbSetPerm As System.Windows.Forms.GroupBox
    Friend WithEvents cbxPermissionsGroup As System.Windows.Forms.ComboBox
    Friend WithEvents lbx1 As System.Windows.Forms.ListBox
    Friend WithEvents panPM As System.Windows.Forms.Panel
    Friend WithEvents lbllbx1 As System.Windows.Forms.Label
    Friend WithEvents cbxPermBase As System.Windows.Forms.ComboBox
    Friend WithEvents lblBase As System.Windows.Forms.Label
    Friend WithEvents lblS As System.Windows.Forms.Label
    Friend WithEvents lblPermissions As System.Windows.Forms.Label
    Friend WithEvents lblDo As System.Windows.Forms.Label
    Friend WithEvents dgvPermissions As System.Windows.Forms.DataGridView
    Friend WithEvents gbWatsonAccount As System.Windows.Forms.GroupBox
    Friend WithEvents cbxWatsonAccount As System.Windows.Forms.ComboBox
    Friend WithEvents gbWindowsAuth As System.Windows.Forms.GroupBox
    Friend WithEvents gbxlblCorporateAdderesses As System.Windows.Forms.GroupBox
    Friend WithEvents gbUA As System.Windows.Forms.GroupBox
    Friend WithEvents gbSTD As System.Windows.Forms.GroupBox
    Friend WithEvents gbxlabelGlobalParameters As System.Windows.Forms.GroupBox
    Friend WithEvents panGP As System.Windows.Forms.Panel
    Friend WithEvents lvPermissionsFinalReport As System.Windows.Forms.ListView
    Friend WithEvents lvPermissionsReportTemplate As System.Windows.Forms.ListView
    Friend WithEvents ID_TBLWATSONACCOUNT As System.Windows.Forms.TextBox
    Friend WithEvents CHARNETWORKACCOUNT As System.Windows.Forms.TextBox
    Friend WithEvents cmdCopyLDAP As System.Windows.Forms.Button
    Friend WithEvents lblLDAP As System.Windows.Forms.Label
    Friend WithEvents CHARLDAP As System.Windows.Forms.TextBox
    Friend WithEvents lblLDAPeg As System.Windows.Forms.Label
    Friend WithEvents cmdGetUserName As System.Windows.Forms.Button
    Friend WithEvents cmdClearWatson As System.Windows.Forms.Button
    Friend WithEvents cmdClearNet As System.Windows.Forms.Button
    Friend WithEvents lblLDAPClear As System.Windows.Forms.Label
    Friend WithEvents gbLDAP As System.Windows.Forms.GroupBox
    Friend WithEvents rbLDAPNon As System.Windows.Forms.RadioButton
    Friend WithEvents rbLDAP As System.Windows.Forms.RadioButton
    Friend WithEvents panLDAP As System.Windows.Forms.Panel
    Friend WithEvents lblLDAPaaa As System.Windows.Forms.Label
    Friend WithEvents rbADVAPI32 As System.Windows.Forms.RadioButton
    Friend WithEvents cmdTestAccount As System.Windows.Forms.Button
    Friend WithEvents lblNetworkAccount As System.Windows.Forms.Label
End Class
