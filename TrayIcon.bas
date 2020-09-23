Attribute VB_Name = "modTrayIcon"
'# Show/Hide IE
'#
'# Coded by:
'# MAGiC MANiAC^mTo - mto@kabelfoon.nl
'# MORTAL OBSESSiON - http://welcome.to/mto

Option Explicit

Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uId As Long
  uFlags As Long
  uCallBackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203   'Left Double-Click
Public Const WM_LBUTTONDOWN = &H201     'Left Button Down
Public Const WM_LBUTTONUP = &H202       'Left Button Up
Public Const WM_RBUTTONDBLCLK = &H206   'Right Double-Click
Public Const WM_RBUTTONDOWN = &H204     'Right Button Down
Public Const WM_RBUTTONUP = &H205       'Right Button Up
    
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public nid As NOTIFYICONDATA

Dim IEVisible As Boolean

Public Sub AddTrayIcon(hWnd As Long, hIcon As Long, sTxt As String)
  nid.cbSize = Len(nid)
  nid.hWnd = hWnd
  nid.uId = vbNull
  nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  nid.uCallBackMessage = WM_MOUSEMOVE
  nid.hIcon = hIcon
  nid.szTip = sTxt & vbNullChar
  Shell_NotifyIcon NIM_ADD, nid
End Sub

Public Sub RemoveTrayIcon()
  Shell_NotifyIcon NIM_DELETE, nid
End Sub

Public Sub TrayEvent(Frm As Form, fXCoord As Single)
  Dim lMsg As Long
  lMsg = fXCoord / Screen.TwipsPerPixelX
  If lMsg = WM_RBUTTONDOWN Then
    If frmMain.AboutVisible Then
      frmAbout.Show
    Else
      frmAbout.Hide
    End If
    frmMain.AboutVisible = Not frmMain.AboutVisible
  End If
  If lMsg = WM_LBUTTONDOWN Then
    If IEVisible Then
      RestoreIE
      ChangeTrayIcon Frm.Image1(3).Picture, "Hide Internet Explorer"
    Else
      MinimizeIE
      ChangeTrayIcon Frm.Image1(2).Picture, "Show Internet Explorer"
    End If
    IEVisible = Not IEVisible
  End If
End Sub

Public Sub ChangeTrayIcon(hIcon As Long, sTxt As String)
  nid.hIcon = hIcon
  nid.szTip = sTxt & vbNullChar
  Shell_NotifyIcon NIM_MODIFY, nid
End Sub

