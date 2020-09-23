VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBmp2Ico 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                                           Easy  bmp2ico"
   ClientHeight    =   2010
   ClientLeft      =   1845
   ClientTop       =   1395
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBmp2Ico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   2520
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtIco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox txtBmp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   360
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   4560
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Close"
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
      Left            =   3240
      MouseIcon       =   "frmBmp2Ico.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1440
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   3645
      Left            =   3240
      Picture         =   "frmBmp2Ico.frx":11D4
      Stretch         =   -1  'True
      Top             =   -1440
      Width           =   4395
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Convert (bmp2ico)"
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
      Left            =   240
      MouseIcon       =   "frmBmp2Ico.frx":491A6
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1440
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "ICO file :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "BMP file :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmBmp2Ico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Please ignore some spelling mistakes ,if i have any,
'cause my English are not so much
'
'If you have any good ideas how to make this program better or
'you want to make me some questions just e-mail me : kyriakosnicola@yhoo.com
'By the way thanks for downloading this fragment of code...

Private Sub cmdBrowse_Click()
With cd
    .DialogTitle = "Open"
    .Filter = "Bitmap Images(*.bmp)|*.bmp"
    .ShowOpen
End With

txtBmp.Text = cd.FileName
If txtBmp.Text <> "" Then
    txtIco.Text = App.Path & "\Untitled.ico"
Else
    txtIco.Text = ""
End If
End Sub

Private Sub cmdSave_Click()
With cd
    .DialogTitle = "Save"
    .Filter = "Ico(*.ico)|*.ico"
    .FileName = "untitled.ico"
    .ShowSave
End With

txtIco.Text = cd.FileName

End Sub

Private Sub Command4_Click()
frmbmp2icoHelp.Show 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HC0& 'just for graphic effects
Label3.ForeColor = &HC0& '    <<          <<
Label2.BorderStyle = 0   '     <<          <<
Label3.BorderStyle = 0   '      <<          <<
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HC0& 'just for graphic effects
Label3.ForeColor = &HC0& '    <<          <<
Label2.BorderStyle = 0   '     <<          <<
Label3.BorderStyle = 0   '      <<          <<
End Sub

Private Sub Label2_Click()
If txtBmp.Text <> "" And txtIco.Text <> "" Then
    ' Load the picture into the ImageList.
    ImageList1.ListImages.Add , , LoadPicture(txtBmp.Text)
    ' Save the icon file.
    SavePicture ImageList1.ListImages(1).ExtractIcon, txtIco.Text
    MsgBox "BMP converted succesfuly to ICO", vbInformation, "Finshed"
Else
    MsgBox "Please verify that you have entered a bmp to convert and/or a path to save as ico", vbInformation, "Error"
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbRed 'just for graphic effects
Label2.BorderStyle = 1   '    <<          <<
End Sub

Private Sub Label3_Click()
Unload Me 'make a guess! :-)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbRed 'just for graphic effects
Label3.BorderStyle = 1   '    <<          <<
End Sub

Private Sub Label4_Click()
frmAbout.Show 1 'duh, it loads the frmAbout form
End Sub

Private Sub Timer1_Timer()
If txtIco.Text = "" Then
    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If
End Sub
