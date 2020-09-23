VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Pong Clock - Menu"
   ClientHeight    =   1320
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   2190
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   2190
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Pong Clock"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnTop 
         Caption         =   "Always on top"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''==========================='''
''' Pong Clock                '''
''' Paul Bahlawan Dec 2005    '''
''' (pop-up menu)             '''
'''==========================='''
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub mnuAbout_Click()
    MsgBox " Pong Clock v0." & App.Revision & vbCrLf & " 2005 Paul Bahlawan     ", vbOKOnly + vbInformation, "Pong Clock - About"
End Sub

Private Sub mnuExit_Click()
    Unload frmPongClock
    Unload Me
End Sub

Private Sub mnuOnTop_Click()
    If mnuOnTop.Checked = True Then
        mnuOnTop.Checked = False
        SetWindowPos frmPongClock.hWnd, -2, 0, 0, 0, 0, &H1 Or &H2
    Else
        mnuOnTop.Checked = True
        SetWindowPos frmPongClock.hWnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
End Sub
