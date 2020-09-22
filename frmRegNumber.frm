VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Register"
   ClientHeight    =   1470
   ClientLeft      =   5445
   ClientTop       =   7800
   ClientWidth     =   2835
   LinkTopic       =   "Form2"
   ScaleHeight     =   1470
   ScaleWidth      =   2835
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   1845
      TabIndex        =   3
      Top             =   1035
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Verify"
      Height          =   255
      Left            =   135
      TabIndex        =   2
      Top             =   1065
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      MaxLength       =   5
      TabIndex        =   1
      Top             =   600
      Width           =   2505
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Type ""sam"" as your serial number"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   285
      Width           =   2370
   End
   Begin VB.Label Label1 
      Caption         =   "Please Enter Your Serial Number"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "sam" Then
SaveSetting appName, secName, "reg", "Þ"
Unload Me
Form3.Show
Else
MsgBox "Invalid Serial Number", , "Error"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
