VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   1440
   ClientLeft      =   4935
   ClientTop       =   2685
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   Begin VB.CommandButton copyfiles 
      Caption         =   "Start copying"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Select drive to copy files to:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim folderdrive As String

Private Sub copyfiles_Click()
fso.CopyFolder dirname, folderdrive, False
End Sub

Private Sub Drive1_Change()
folderdrive = Drive1
End Sub
