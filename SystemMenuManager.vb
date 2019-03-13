''' -----------------------------------------------------------------------------
''' <summary>
'''     Disables and removes items in the system menu of a form.
''' </summary>
''' -----------------------------------------------------------------------------
Public NotInheritable Class SystemMenuManager

#Region " Types "

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     The states a menu item can have.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Enum MenuItemState
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     The menu item is present and can be selected.
        ''' </summary>
        ''' <remarks>
        '''     This is the default state.
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Enabled = MF_ENABLED
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     The menu item is present but can not be selected.
        ''' </summary>
        ''' -----------------------------------------------------------------------------
        Disabled = MF_DISABLED
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     The menu item is present but greyed-out and cannot be selected.
        ''' </summary>
        ''' -----------------------------------------------------------------------------
        Greyed = MF_GRAYED
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     The menu item is not present.
        ''' </summary>
        ''' -----------------------------------------------------------------------------
        Removed
    End Enum

#End Region 'Types

#Region " Constants "

    Private Const SC_MOVE As Integer = &HF010   'The Move menu item.
    Private Const SC_CLOSE As Integer = &HF060  'The Close menu item.

    Private Const MF_BYCOMMAND As Integer = &H0 'The menu item is specified by command.

    Private Const MF_ENABLED As Integer = &H0   'Enable the menu item.
    Private Const MF_DISABLED As Integer = &H2  'Disable the menu item.
    Private Const MF_GRAYED As Integer = &H1    'Grey out the menu item.

#End Region 'Constants

#Region " Variables "

    Private WithEvents m_Form As Form       'The form for which the system menu will be managed.
    Private m_CloseState As MenuItemState   'The state of the Close menu item.
    Private m_MenuHandle As IntPtr          'The handle to the form's system menu.

#End Region 'Variables

#Region " APIs "

    ' -----------------------------------------------------------------------------
    ' <summary>
    '     Gets the handle of the form's system menu.
    ' </summary>
    ' <param name="hWnd">
    '     The handle of the form for which to get the system menu.
    ' </param>
    ' <param name="bRevert">
    '     Indicates whether the menu should be reset to its original state.
    ' </param>
    ' <returns>
    '     The handle of the system menu of the specified form.
    ' </returns>
    ' -----------------------------------------------------------------------------
    Private Declare Auto Function GetSystemMenu Lib "user32" (ByVal hWnd As IntPtr, ByVal bRevert As Boolean) As IntPtr

    ' -----------------------------------------------------------------------------
    ' <summary>
    '     Sets the state of the specified menu item.
    ' </summary>
    ' <param name="hMenu">
    '     The handle of the menu containing the item.
    ' </param>
    ' <param name="wIDEnableItem">
    '     The menu item for which to set the state.
    ' </param>
    ' <param name="wEnable">
    '     The way in which the wIDEnableItem argument identifies the menu item and
    '     the new state of the item.
    ' </param>
    ' <returns>
    '     The previous state of the menu item if it exists;
    '     -1 otherwise.
    ' </returns>
    ' -----------------------------------------------------------------------------
    Private Declare Auto Function EnableMenuItem Lib "user32" (ByVal hMenu As IntPtr, _
                                                               ByVal wIDEnableItem As Integer, _
                                                               ByVal wEnable As Integer) As Integer

    ' -----------------------------------------------------------------------------
    ' <summary>
    '     Deletes the specified menu item.
    ' </summary>
    ' <param name="hMenu">
    '     The handle of the menu from which to delete the item.
    ' </param>
    ' <param name="uPosition">
    '     The menu item to delete.
    ' </param>
    ' <param name="uFlags">
    '     Indicates how the uPosition indentifies the menu item.
    ' </param>
    ' <returns>
    '     Non-zero if the function succeeds;
    '     zero otherwise.
    ' </returns>
    ' -----------------------------------------------------------------------------
    Private Declare Auto Function DeleteMenu Lib "user32" (ByVal hMenu As IntPtr, _
                                                           ByVal uPosition As Integer, _
                                                           ByVal uFlags As Integer) As Boolean

#End Region 'APIs

#Region " Constructors "

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Creates a new SystemMenuManager object for the specified form and
    '''     optionally removes the Move item from the system menu.
    ''' </summary>
    ''' <param name="form">
    '''     The form for which to manage the system menu.
    ''' </param>
    ''' <param name="movePresent">
    '''     Indicates whether the Move menu item is present or not.
    ''' </param>
    ''' <remarks>
    '''     The Close menu item is unaffected.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal form As Form, ByVal movePresent As Boolean)
        Me.New(form, _
               movePresent, _
               MenuItemState.Enabled)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Creates a new SystemMenuManager object for the specified form and
    '''     optionally changes the state of the Close menu item.
    ''' </summary>
    ''' <param name="form">
    '''     The form for which to manage the system menu.
    ''' </param>
    ''' <param name="closeState">
    '''     The state of the Close menu item.
    ''' </param>
    ''' <remarks>
    '''     The Move menu item is unaffected.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal form As Form, ByVal closeState As MenuItemState)
        Me.New(form, _
               True, _
               closeState)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Creates a new SystemMenuManager object for the specified form and
    '''     optionally changes the state of the Move and Close menu items.
    ''' </summary>
    ''' <param name="form">
    '''     The form for which to manage the system menu.
    ''' </param>
    ''' <param name="movePresent">
    '''     Indicates whether the Move menu item is present or not.
    ''' </param>
    ''' <param name="closeState">
    '''     The state of the Close menu item.
    ''' </param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal form As Form, _
                   ByVal movePresent As Boolean, _
                   ByVal closeState As MenuItemState)
        Me.m_Form = form
        Me.m_CloseState = closeState

        'Get the handle to the form's system menu.
        Me.m_MenuHandle = Me.GetSystemMenu(Me.m_Form.Handle, False)

        If Not movePresent Then
            'Remove the Move menu item.
            Me.DeleteMenu(Me.m_MenuHandle, _
                          Me.SC_MOVE, _
                          Me.MF_BYCOMMAND)
        End If

        If Me.m_CloseState = MenuItemState.Removed Then
            'Remove the Close menu item.
            Me.DeleteMenu(Me.m_MenuHandle, _
                          Me.SC_CLOSE, _
                          Me.MF_BYCOMMAND)
        Else
            Me.RefreshCloseItem()
        End If

        If Me.m_CloseState <> MenuItemState.Enabled Then
            'Set the Keypreview to True so that the Alt+F4 key combination can be detected.
            Me.m_Form.KeyPreview = True
        End If
    End Sub

#End Region 'Constructors

#Region " Methods "

    ' -----------------------------------------------------------------------------
    ' <summary>
    '     Refreshes the state of the Close menu item.
    ' </summary>
    ' <remarks>
    '     Action is only taken if the state is Disabled or Greyed.
    ' </remarks>
    ' -----------------------------------------------------------------------------
    Private Sub RefreshCloseItem()
        If Me.m_CloseState = MenuItemState.Disabled OrElse Me.m_CloseState = MenuItemState.Greyed Then
            'Set the Close menu item state.
            Me.EnableMenuItem(Me.m_MenuHandle, _
                              Me.SC_CLOSE, _
                              Me.MF_BYCOMMAND Or Me.m_CloseState)
        End If
    End Sub

#End Region 'Methods

#Region " Event Handlers "

    Private Sub managedForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles m_Form.Load
        'The Close menu item must have it's state refresehed if it is present and not enabled.
        Me.RefreshCloseItem()
    End Sub

    Private Sub managedForm_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles m_Form.Resize
        'The Close menu item must have it's state refresehed if it is present and not enabled.
        Me.RefreshCloseItem()
    End Sub

    Private Sub m_Form_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles m_Form.KeyDown
        If e.KeyCode = Keys.F4 AndAlso _
           e.Alt AndAlso _
           Me.m_CloseState <> MenuItemState.Enabled Then
            'Disable the Alt+F4 key combination.
            e.Handled = True
        End If
    End Sub

#End Region 'Event Handlers

End Class
