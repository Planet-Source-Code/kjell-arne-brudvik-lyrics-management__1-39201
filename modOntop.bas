Attribute VB_Name = "modOntop"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Function AlwaysOnTop(frmTop As Form, SetOnTop As Boolean)

    '// Check if we set it ontop, or remove it from ontop.
    If SetOnTop = True Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If

    SetWindowPos frmTop.hwnd, lFlag, frmTop.Left / Screen.TwipsPerPixelX, _
    frmTop.Top / Screen.TwipsPerPixelY, frmTop.Width / Screen.TwipsPerPixelX, _
    frmTop.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
End Function

