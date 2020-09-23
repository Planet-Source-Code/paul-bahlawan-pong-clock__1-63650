VERSION 5.00
Begin VB.Form frmPongClock 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "pongclock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrGameTick 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "frmPongClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''======================================='''
''' Pong Clock                            '''
''' Paul Bahlawan Dec 2005                '''
''' (Main)                                '''
'''======================================='''
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Dim hour As Long
Dim min As Long
Dim cngHour As Boolean
Dim cngMin As Boolean
Dim DrWidth As Long
Dim num(9) As String
Dim BallX As Long
Dim BallY As Single
Dim Dx As Long
Dim Dy As Single
Dim LPaddle As Long
Dim RPaddle As Long
Dim Speed As Long
Dim delay As Long


Private Sub Form_Load()
    Randomize
    Me.Show
    Me.ScaleMode = vbPixels
    Me.BackColor = vbBlack
    
    num(0) = "0020202424040400"
    num(1) = "2024"
    num(2) = "00202022220202040424"
    num(3) = "2024002002220424"
    num(4) = "000202222024"
    num(5) = "20000002022222242404"
    num(6) = "0004042424222202"
    num(7) = "00202024"
    num(8) = "00202024240404000222"
    num(9) = "2420200000020222"
    
    LPaddle = 50
    RPaddle = 50
    Do
        Dx = (Int(Rnd * 3) - 1)
    Loop While Dx = 0
    
    Draw_Court
    delay = 15
    tmrGameTick.Enabled = True
End Sub


Private Sub Draw_Court()
Dim X As Long
Dim Y As Long
Dim hw As Long
    Me.DrawWidth = 1 '**
    X = Me.ScaleWidth
    Y = Me.ScaleHeight
    hw = 1 + Int(Y / 60)
    DrWidth = hw * 2
    Me.AutoRedraw = True
    Me.Cls
    Me.ForeColor = vbWhite
    Me.Line (X / 2 - hw / 2, 0)-(X / 2 + hw / 2, Y), &H808080, BF
    Me.Line (0, 0)-(X, 3 * Y / 100), , BF
    Me.Line (0, Y)-(X, 97 * Y / 100), , BF
    Me.AutoRedraw = False
End Sub


Private Sub Draw_Time()
Dim h As String
Dim m As String
Dim cnt As Long
Dim tmp As Long
    h = Format(Time(), "hh")
    m = Format(Time(), "nn")
    
    'Store current time so later we can check if it has changed
    hour = CLng(h)
    min = CLng(m)
    
    Me.AutoRedraw = True
    Me.DrawWidth = DrWidth / 2 '**
    tmp = Me.ScaleWidth / 2
     
    'Clear old "score"
    Me.Line (tmp - (DrWidth * 9), DrWidth * 3)-(tmp - DrWidth * 2, DrWidth * 7), vbBlack, BF
    Me.Line (tmp + (DrWidth * 2), DrWidth * 3)-(tmp + DrWidth * 9, DrWidth * 7), vbBlack, BF
    
    'Show new "score"
    For cnt = 1 To 2
        Draw_Digit Val(Mid(h, cnt, 1)), tmp - (DrWidth * 11) + (cnt * DrWidth * 3), DrWidth * 3
        Draw_Digit Val(Mid(m, cnt, 1)), tmp + (cnt * DrWidth * 3), DrWidth * 3
    Next
    Me.AutoRedraw = False
End Sub


'Draw a digit from data in num()
Private Sub Draw_Digit(Digit As Long, X As Long, Y As Long)
Dim i As Long
    For i = 1 To Len(num(Digit)) Step 4
        Me.Line (X + Val(Mid(num(Digit), i, 1)) * DrWidth, Y + Val(Mid(num(Digit), i + 1, 1)) * DrWidth)-(X + Val(Mid(num(Digit), i + 2, 1)) * DrWidth, Y + Val(Mid(num(Digit), i + 3, 1)) * DrWidth)
    Next
End Sub


'Pop up menu with left click
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu frmMenu.mnuFile
    End If
End Sub


'Allow dragging the window
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub


'Re-draw the screen if resized
Private Sub Form_Resize()
    Draw_Court
    Draw_Time
End Sub


'All the 'game' action is done here
Private Sub tmrGameTick_Timer()
    If delay And BallY > 0 Then
        delay = delay - 1
        Exit Sub
    End If
    
    'Set up a New_Ball?
    If BallY = 0 Then
        BallX = 50
        BallY = 50
        Do
            Dy = Int(Rnd * 50) / 10 - 2
        Loop While Dy = 0
        Me.Refresh
        Draw_Time
        Draw_BP
        Speed = 50
        tmrGameTick.Interval = Speed
        cngHour = False
        cngMin = False
        delay = 15
        Exit Sub
    End If
    
    'Goal scored!
    If BallX < -2 Or BallX > 102 Then
        BallY = 0
        Exit Sub
    End If
    
    'Set a flag if minute or hour has changed
    If Not cngHour And Not cngMin Then
        If CLng(Format(Time(), "hh")) <> hour Then
            cngHour = True
        Else
            If CLng(Format(Time(), "nn")) <> min Then
                cngMin = True
            End If
        End If
    End If
    
    'Bounce ball off wall (top or bottom)
    If (BallY < 4 And Dy < 0) Or (BallY > 96 And Dy > 0) Then
        Dy = -Dy
    End If
    
    'Bounce ball off left paddle
    If BallX < 5 And Not cngMin Then
        Dx = -Dx
        Dy = Dy + (Int(Rnd * 10) / 10) - 0.5
        Speed = Speed - 2
    End If
    
    'Bounce ball off right paddle
    If BallX > 95 And Not cngHour Then
        Dx = -Dx
        Dy = Dy + (Int(Rnd * 10) / 10) - 0.5
        Speed = Speed - 2
    End If
    
    'Calculate ball's new position
    If Dy < -3.5 Then Dy = -3.5
    If Dy > 3.5 Then Dy = 3.5
    BallX = BallX + Dx
    BallY = BallY + Dy
    
    'Calculate paddle's new position
    If Dx < 0 Then
        'Left paddle
        If cngMin And BallX < 10 Then 'move paddle outa the way of the ball
            If Dy <= 0 Then
                LPaddle = LPaddle + 1
            Else
                LPaddle = LPaddle - 1
            End If
        Else                          'move paddle to block ball
            If Abs(LPaddle - BallY) > 4 Then
                If LPaddle > BallY Then
                    LPaddle = LPaddle - 4
                Else
                    LPaddle = LPaddle + 4
                End If
            Else
                LPaddle = BallY
            End If
        End If
        If LPaddle < 0 Then LPaddle = 0
        If LPaddle > 100 Then LPaddle = 100
    
    Else
        'Right paddle
        If cngHour And BallX > 90 Then 'move paddle outa the way of the ball
            If Dy <= 0 Then
                RPaddle = RPaddle + 1
            Else
                RPaddle = RPaddle - 1
            End If
        Else                          'move paddle to block ball
            If Abs(RPaddle - BallY) > 4 Then
                If RPaddle > BallY Then
                    RPaddle = RPaddle - 4
                Else
                    RPaddle = RPaddle + 4
                End If
            Else
                RPaddle = BallY
            End If
        End If
        If RPaddle < 0 Then RPaddle = 0
        If RPaddle > 100 Then RPaddle = 100
        
    End If
        
    Draw_BP
        
    'Speed up as game progresses
    tmrGameTick.Interval = Speed
End Sub

'Draw Ball and Paddles
Private Sub Draw_BP()
Dim tmp As Long
Dim X As Long
Dim Y As Long
Dim cy As Long
    Me.DrawWidth = 1 '**
    Me.AutoRedraw = True '(this also causes the 'old' ball to be erased)
    
'Draw paddles too
    'left...
    tmp = DrWidth / 2
    X = 4 * ScaleWidth / 100
    Y = LPaddle * ScaleHeight / 100
    Me.Line (X - tmp, Y - DrWidth)-(X, Y + DrWidth), , BF
    
    cy = 97 * Me.ScaleHeight / 100 - 1
    If Y + DrWidth < cy Then
        Me.Line (X - tmp, cy)-(X, Y + DrWidth), vbBlack, BF
    End If
    cy = 3 * Me.ScaleHeight / 100 + 1
    If Y - DrWidth > cy Then
        Me.Line (X - tmp, Y - DrWidth)-(X, cy), vbBlack, BF
    End If

    'right...
    X = 96 * ScaleWidth / 100
    Y = RPaddle * ScaleHeight / 100
    Me.Line (X, Y - DrWidth)-(X + tmp, Y + DrWidth), , BF
    
    cy = 97 * Me.ScaleHeight / 100 - 1
    If Y + DrWidth < cy Then
        Me.Line (X, cy)-(X + tmp, Y + DrWidth), vbBlack, BF
    End If
    cy = 3 * Me.ScaleHeight / 100 + 1
    If Y - DrWidth > cy Then
        Me.Line (X, Y - DrWidth)-(X + tmp, cy), vbBlack, BF
    End If
    Me.AutoRedraw = False

'Draw ball
    tmp = DrWidth / 4
    X = BallX * ScaleWidth / 100
    Y = CLng(BallY) * ScaleHeight / 100
    Me.Line (X - tmp, Y - tmp)-(X + tmp, Y + tmp), , BF
End Sub
