Imports System
Imports System.Reflection
Imports System.Windows.Forms

'some datagrids flicker and redraw slowly
'this will speedup the grid redraw
'http://bitmatic.com/c/fixing-a-slow-scrolling-datagridview

Module modExtensionMethods

    Public Sub DoubleBufferedControl(ByVal ctrl As Control, ByVal setting As Boolean)

        Dim ctrlType As Type = ctrl.[GetType]()
        Dim pi As PropertyInfo = ctrlType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        pi.SetValue(ctrl, setting, Nothing)

    End Sub

    Public Sub SetButton(ByVal ctrl As Control, ByVal setting As Boolean)

        Dim ctrlType As Type = ctrl.[GetType]()
        Dim pi As PropertyInfo = ctrlType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        pi.SetValue(ctrl, setting, Nothing)

    End Sub

End Module
