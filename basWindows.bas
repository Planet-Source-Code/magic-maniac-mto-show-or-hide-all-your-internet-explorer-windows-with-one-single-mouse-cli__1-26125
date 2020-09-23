Attribute VB_Name = "basWindows"
'# Show/Hide IE
'#
'# Coded by:
'# MAGiC MANiAC^mTo - mto@kabelfoon.nl
'# MORTAL OBSESSiON - http://welcome.to/mto

Option Explicit

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type

Public Type IE_STATE_SAVE
  hWnd As Long
  wp As WINDOWPLACEMENT
End Type

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassNameA Lib "user32" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private arrayIESS() As IE_STATE_SAVE
Global Const SW_HIDE = 0

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
  Dim l As Long
  Dim sClsNm As String
  Dim ss As IE_STATE_SAVE
  Dim wp As WINDOWPLACEMENT
  wp.Length = Len(wp)
  sClsNm = GetClassName(hWnd)
  If sClsNm = "IEFrame" Or sClsNm = "CabinetWClass" Then
    GetWindowPlacement hWnd, wp
    ss.hWnd = hWnd
    ss.wp = wp
    l = UBound(arrayIESS)
    arrayIESS(l) = ss
    ReDim Preserve arrayIESS(l + 1)
    wp.showCmd = SW_HIDE
    SetWindowPlacement hWnd, wp
  End If
  EnumWindowsProc = True
End Function

Public Sub MinimizeIE()
  ReDim arrayIESS(0)
  EnumWindows AddressOf EnumWindowsProc, 0
End Sub

Public Sub RestoreIE()
  On Error GoTo exit_RestoreIE
  Dim ieSS As IE_STATE_SAVE
  Dim l As Long
  For l = UBound(arrayIESS) To LBound(arrayIESS) Step -1
    ieSS = arrayIESS(l)
    With ieSS
      If .hWnd > 0 Then
        SetWindowPlacement .hWnd, .wp
      End If
    End With
  Next
exit_RestoreIE:
  Exit Sub
End Sub

Private Function GetClassName(ByVal hWnd As Long) As String
  Dim lngReturn As Long
  Dim strReturn As String
  strReturn = Space(255)
  lngReturn = GetClassNameA(hWnd, strReturn, Len(strReturn))
  GetClassName = Left(strReturn, lngReturn)
End Function

