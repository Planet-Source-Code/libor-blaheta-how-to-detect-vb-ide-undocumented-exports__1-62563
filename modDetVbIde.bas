Attribute VB_Name = "modDetVbIde"
Option Explicit

'
' --- How to detect VB IDE - undocumented exports ---
'
' When you run your VB app in VB IDE, the app exports 3 functions.
' Their names are
'
' _VB_CALLBACK_REGISTER_@8
' _VB_CALLBACK_REVOKE_@8
' _VB_CALLBACK_GETHWNDMAIN_@4
'
' Don't ask me what those functions do. I don't know. :-)
'
' So, if we detect our app exports one of those functions, we know the app runs in VB IDE.
' Simple, isn't it?
'
' by Libor Blaheta
'

Private Const exp1 As String = "_VB_CALLBACK_REGISTER_@8"
Private Const exp2 As String = "_VB_CALLBACK_REVOKE_@8"
Private Const exp3 As String = "_VB_CALLBACK_GETHWNDMAIN_@4"

Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Function VbIde() As Boolean

    VbIde = IIf(GetProcAddress(App.hInstance, exp1) = 0, False, True)

End Function

Public Sub Main()
    MsgBox "In VB IDE - " & VbIde, vbInformation, "Detect VB IDE"
End Sub
