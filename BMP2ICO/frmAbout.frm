VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                       About Author"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4575
      Begin VB.PictureBox picScroll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1545
         Left            =   120
         ScaleHeight     =   99
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   283
         TabIndex        =   3
         Top             =   240
         Width           =   4305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KyriakosNicola@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   960
         TabIndex        =   4
         Top             =   1920
         Width           =   2580
      End
   End
   Begin VB.Label lblBye 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Me !"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   1680
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3960
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Made By . . ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   600
      Picture         =   "frmAbout.frx":030A
      Top             =   480
      Width           =   3690
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'the actual text to scroll. This could also be loaded from a text file
Const ScrollText As String = "-:Easy Bmp 2 Ico:-" & vbCrLf & _
                             vbCrLf & vbCrLf & _
                             "Hello There!" & vbCrLf & _
                             "Thanks for downloading this fragment of code..." & vbCrLf & _
                             "I hope you like it and vote for me!" & _
                             vbCrLf & vbCrLf & _
                             vbCrLf & "If you have any good ideas and you want to share them with me PLEASE e-mail me!!!" & _
                             vbCrLf & "" & _
                             vbCrLf & vbCrLf & _
                             ""
                             'Without the last three lines a part
                             'of the txt doesn't show up.
                             'If you know why and how can it be fixed
                             'please e-mail me |kyriakosnicola@yahoo.com|
Dim EndingFlag As Boolean
Private Sub Form_Activate()
RunMain
End Sub

Private Sub RunMain()
Dim LastFrameTime As Long
Const IntervalTime As Long = 40
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long 'Upper left point of drawing rect
Dim RectHeight As Long

'show the form
frmAbout.Refresh

'Get the size of the drawing rectangle by suppying the DT_CALCRECT constant
rt = DrawText(picScroll.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then 'err
    MsgBox "Error scrolling text", vbExclamation
    EndingFlag = True
Else
    DrawingRect.Top = picScroll.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = picScroll.ScaleWidth
    'Store the height of The rect
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
End If


Do While Not EndingFlag
    
    If GetTickCount() - LastFrameTime > IntervalTime Then
                    
        picScroll.Cls
        
        DrawText picScroll.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
        'update the coordinates of the rectangle
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        
        'control the scolling and reset if it goes out of bounds
        If DrawingRect.Top < -(RectHeight) Then 'time to reset
            DrawingRect.Top = picScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
        End If
        
        picScroll.Refresh
        
        LastFrameTime = GetTickCount()
        
    End If
    
    DoEvents
Loop

Unload Me
Set frmAbout = Nothing

End Sub

Private Sub Form_Load()
picScroll.FontSize = 14
picScroll.ForeColor = vbGreen
End Sub

Private Sub Form_Unload(Cancel As Integer)

    EndingFlag = True 'ends flag
   
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBye.ForeColor = &HC0&
lblBye.BorderStyle = 0
End Sub

Private Sub lblBye_Click()
Unload Me
EndingFlag = True
End Sub

Private Sub lblBye_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBye.ForeColor = vbRed
lblBye.BorderStyle = 1
End Sub
