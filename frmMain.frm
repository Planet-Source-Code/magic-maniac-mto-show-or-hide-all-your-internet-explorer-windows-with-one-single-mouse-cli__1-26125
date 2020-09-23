VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   330
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   330
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   855
      Picture         =   "frmMain.frx":014A
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   585
      Picture         =   "frmMain.frx":0294
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   315
      Picture         =   "frmMain.frx":03DE
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   45
      Picture         =   "frmMain.frx":0528
      Top             =   45
      Width           =   240
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mPopHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# Show/Hide IE
'#
'# Coded by:
'# MAGiC MANiAC^mTo - mto@kabelfoon.nl
'# MORTAL OBSESSiON - http://welcome.to/mto

Option Explicit

Public AboutVisible As Boolean

Private Sub Form_Load()
  If App.PrevInstance Then 'application already running...
    Unload Me
    End
  End If
  AboutVisible = False
  
  Me.Visible = False
  AddTrayIcon Me.hWnd, Me.Image1(3).Picture, "Hide Internet Explorer"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveTrayIcon
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TrayEvent Me, X
End Sub

