VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   ". : :  C O D E R - 5 1   : : . "
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5640
      Top             =   600
   End
   Begin MSComctlLib.ProgressBar Bar 
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ProgressBar Bar2 
      Height          =   3495
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   6165
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
   End
   Begin VB.Label Label1 
      Caption         =   "2 Simple Progress Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   -2640
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   2175
      Left            =   3240
      Top             =   -1560
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   1320
      Top             =   480
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   3735
      Left            =   600
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer
Dim num2 As Integer
Private Sub Form_Load()
 num = 1
 num2 = 100
End Sub
Private Sub Timer1_Timer()
   If num = 100 Then
   Timer2.Enabled = False
           MsgBox "EXIT", 0, "2 Simple Progress Bar"
    End If
If num = 100 Then
    Unload Me
    Exit Sub
End If
num = num + 1
Bar.Value = num
End Sub

Private Sub Timer2_Timer()
If num2 = 1 Then
    'Unload Me
    End If
num2 = num2 - 1
Bar2.Value = num2

End Sub
