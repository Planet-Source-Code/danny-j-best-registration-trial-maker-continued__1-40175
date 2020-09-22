VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4950
   ClientLeft      =   5805
   ClientTop       =   5385
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   4950
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel Registration"
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   510
      Left            =   2475
      TabIndex        =   0
      Top             =   210
      Width           =   2040
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This is our software's Main MDI / SDI Form."
      Height          =   3195
      Left            =   195
      TabIndex        =   2
      Top             =   1275
      Width           =   4230
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*                         Trial Registration (Continued)                  *
'*     Look here, I'm not copying someone's work i'm just                  *
'*     continuing what the prevoius programer made.                        *
'*     Don't give the credit to me. All I did is make it more secure       *
'***************************************************************************
Option Explicit     'Genaral

Private Sub Command1_Click()
Unload Me       'Unloads the form
End Sub

Private Sub Command2_Click()
If GetAllSettings(appName, secName) Then    'Get the reg settings
Command2.Enabled = True                 'To prevent errors, disable command2
Else                                    'Else
Command2.Enabled = False                'Else containment
DeleteSetting appName, secName, "st"    'Delete st
DeleteSetting appName, secName, "start" 'Delete start
DeleteSetting appName, secName, "now"   'Delete now
DeleteSetting appName, secName, "reg"   'Delete reg
DeleteSetting appName, secName, "alt"   'Delete alt
End Sub

Private Sub Form_Load()

End Sub
