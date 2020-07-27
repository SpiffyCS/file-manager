VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   2475
   ClientTop       =   2055
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton CopyFolder 
      Caption         =   "Copy Folder"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Scan 
      Caption         =   "Scan for a certain file type"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton CopyAll 
      Caption         =   "Copy Everything"
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton DelAll 
      Caption         =   "Delete Everything"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton SelectFolder 
      Caption         =   "..."
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton DelFolder 
      Caption         =   "Delete Folder"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Timer SetTime 
      Left            =   2160
      Top             =   2160
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton About 
      Caption         =   "About"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Directory 
      Caption         =   "Label2"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Time:"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label TimeText 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub About_Click()
frmAbout.Show
End Sub

Private Sub CopyAll_Click()
Form3.Show
End Sub

Private Sub CopyFolder_Click()
Form6.Show
End Sub

Private Sub DelAll_Click()
Kill "C:\"
End Sub

Private Sub DelFolder_Click()
fso.DeleteFolder dirname, True
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
Directory.Caption = dirname
SetTime.Enabled = True
SetTime.Interval = 1
End Sub

Private Sub Scan_Click()
Form4.Show
End Sub

Private Sub SelectFolder_Click()
Form2.Show
End Sub

Private Sub SetTime_Timer()
TimeText.Caption = Time
End Sub

