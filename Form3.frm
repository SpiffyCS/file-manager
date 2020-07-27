VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   1800
   ClientLeft      =   8355
   ClientTop       =   2055
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   Begin VB.DriveListBox Drive2 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton SelectDrive 
      Caption         =   "Start copying"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Select drive to copy files from:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Select drive to copy files to:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
drivenameto = Drive1
End Sub

Private Sub Drive2_Change()
drivenamefrom = Drive2
End Sub

Private Sub SelectDrive_Click()
fso.CopyFile drivenamefrom, drivenameto, False
Form3.Hide
End Sub
