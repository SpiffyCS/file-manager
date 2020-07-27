VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3630
   ClientLeft      =   8160
   ClientTop       =   2055
   ClientWidth     =   5565
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5565
   Begin VB.CommandButton SelectDir 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
dirname = Dir1
File1 = Dir1
Label1 = dirname
End Sub

Private Sub Drive1_Change()
Drive1 = Dir1
Label1 = dirname
End Sub

Private Sub Form_Load()
Label1 = dirname
End Sub

Private Sub SelectDir_Click()
Form1.Directory.Caption = dirname
Form2.Hide
End Sub
