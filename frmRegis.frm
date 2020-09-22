VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Main"
   ClientHeight    =   3690
   ClientLeft      =   4335
   ClientTop       =   4935
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5115
   Begin VB.CommandButton Command4 
      Caption         =   "Register"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Try It !"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "UNREGISTERED PROGRAM"
      Height          =   975
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "You have used 0 days of your 30 day Trial"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
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
Dim abd As Integer
Dim Registered As Boolean
Dim jj As Integer
Dim st, en As Date

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Unload Me           'Unload The form
End Sub

Private Sub Command3_Click()
Unload Me           'Unload the form
Form3.Show          'Show form3
End Sub

Private Sub Command4_Click()
Unload Me           'Unload Form
Form2.Show          'Show form2
End Sub

Private Sub Form_Load()
appName = "Microsoft"   'Dim
secName = "zx12Win"     'Dim


If GetSetting(appName, secName, "st") <> "×" Then   'GetSetting
SaveSetting appName, secName, "st", "×"             'GetSetting
SaveSetting appName, secName, "start", Date         'GetSetting
SaveSetting appName, secName, "now", Date           'GetSetting
SaveSetting appName, secName, "reg", "1"            'GetSetting
SaveSetting appName, secName, "alt", "Ö"            'GetSetting
End If                                              'End if

If GetSetting(appName, secName, "reg") = "Þ" Then   'GetSetting
Unload Me                                           'Unload the form
MsgBox "Software registered"                        'You don't really need that
Form3.Show                                          'Show form3
Else                                                'Else

st = GetSetting(appName, secName, "start")          'GetSetting
en = GetSetting(appName, secName, "now")            'GetSetting
abd = DateDiff("d", st, Date)                       'GetSetting
jj = DateDiff("d", en, Date)                        'GetSetting

If abd >= 0 And jj >= 0 And GetSetting(appName, secName, "alt") = "Ö" Then 'Set the trial
Label1.Caption = "Your " & (30 - abd) & " day(s) left for the try"         'Days Remaining
Else                                                                       'else
SaveSetting appName, secName, "alt", "1"                                   'SaveSetting
MsgBox "[The Program] has detected a date alter.", vbCritical, "Date Alter Detected"    'Date Alter Detected. Bad boy...
Unload Me                                           'Unload the form
Form2.Show                                          'Show form2
End If                                              'End If

End If                                              'End If (again)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GetSetting(appName, secName, "alt") = "Ö" Then   'GetAlter
    Dim tt As String                                'Dim tt
    tt = Date                                       'What tt is
    SaveSetting appName, secName, "now", tt = "©®" & MM / DD / YY   'SaveSetting (Encrypted)
End If
End Sub
