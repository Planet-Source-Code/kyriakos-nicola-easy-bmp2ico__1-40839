VERSION 5.00
Begin VB.Form frmbmp2icoHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Browse for Bitmap Image (*.bmp)"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   120
         Picture         =   "frmbmp2icoHelp.frx":0000
         Top             =   1440
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   360
         Picture         =   "frmbmp2icoHelp.frx":03B6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2160
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   120
         Picture         =   "frmbmp2icoHelp.frx":13EF8
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Change Converted Application's Path . . ."
         ForeColor       =   &H000000C0&
         Height          =   555
         Left            =   600
         TabIndex        =   2
         Top             =   1800
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmbmp2icoHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
