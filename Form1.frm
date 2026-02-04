VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "3D Chess Pawn"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10530
   ForeColor       =   &H80000005&
   LinkTopic       =   "Form1"
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   702
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRotate 
      Interval        =   30
      Left            =   8520
      Top             =   5520
   End
   Begin VB.PictureBox picOpenGL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   1080
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   960
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = "3D Chess Pawn - Full Rotation (Mouse & Keyboard)"
    Me.ScaleMode = vbPixels

    picOpenGL.ScaleMode = vbPixels
    picOpenGL.AutoRedraw = False
    picOpenGL.BackColor = vbBlack
    picOpenGL.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

    InitGL picOpenGL.hDC, picOpenGL.ScaleWidth, picOpenGL.ScaleHeight
    DrawScene picOpenGL.hDC

    ' penting agar Form_KeyDown menerima tombol panah
    Me.KeyPreview = True
End Sub

' ===== Mouse rotate =====
Private Sub picOpenGL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        isDragging = True
        lastMouseX = X
        lastMouseY = Y
    End If
End Sub

Private Sub picOpenGL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isDragging Then
        
        'rotasi 2
        'Gerak mouse kiri–kanan mengubah rotY (putaran horizontal kamera).
        'Gerak atas–bawah mengubah rotX (putaran vertikal objek).
        'rotX dibatasi agar pion tidak terbalik 180°.
        
        rotY = rotY + (X - lastMouseX) * 0.8   ' kiri-kanan ? Y
        rotX = rotX + (Y - lastMouseY) * 0.8   ' atas-bawah ? X
        
        'rotY mengubah posisi kamera (kiri–kanan).
        'rotX memutar pion (atas–bawah).

        ' clamp optional
        If rotX > 89 Then rotX = 89
        If rotX < -89 Then rotX = -89

        lastMouseX = X
        lastMouseY = Y

        DrawScene picOpenGL.hDC
    End If
End Sub

Private Sub picOpenGL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isDragging = False
End Sub

' ===== Keyboard rotate =====
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        
        'rotasi 3
        'Tombol panah juga bisa mengatur arah rotasi.
        'Me.KeyPreview = True pada Form_Load membuat form tetap bisa menangkap event keyboard.
        
        Case vbKeyLeft:  rotY = rotY - 5
        Case vbKeyRight: rotY = rotY + 5
        Case vbKeyUp:    rotX = rotX - 5
        Case vbKeyDown:  rotX = rotX + 5
        Case vbKeyR:     rotX = 0: rotY = 0   ' reset (opsional)
    End Select
    DrawScene picOpenGL.hDC
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        picOpenGL.Move 0, 0, ScaleWidth, ScaleHeight
        ResizeGLScene picOpenGL.ScaleWidth, picOpenGL.ScaleHeight
        DrawScene picOpenGL.hDC
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillGL picOpenGL.hDC
End Sub

