Attribute VB_Name = "Max_e_Min_Tela"
Option Explicit

Declare PtrSafe Function FindWindowA& Lib "User32" (ByVal lpClassName$, ByVal lpWindowName$)
Declare PtrSafe Function GetWindowLongA& Lib "User32" (ByVal hWnd&, ByVal nIndex&)
Declare PtrSafe Function SetWindowLongA& Lib "User32" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)


' Déclaration des constantes
Public Const GWL_STYLE As Long = -16
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_FULLSIZING = &H70000

'Attention, envoyer après changement du caption de l'UF
Public Sub InitMaxMin(mCaption As String, Optional Max As Boolean = True, Optional Min As Boolean = True _
        , Optional Sizing As Boolean = True)
Dim hWnd As Long
    hWnd = FindWindowA(vbNullString, mCaption)
    If Max Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MAXIMIZEBOX
    If Min Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MINIMIZEBOX
    If Sizing Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_FULLSIZING
End Sub

Sub CallWebPage(url)
ActiveWorkbook.FollowHyperlink Address:=url, NewWindow:=True, AddHistory:=True

End Sub
Sub CallWebPage2(url)
ActiveWorkbook.FollowHyperlink Address:=url, NewWindow:=True, AddHistory:=True

End Sub



