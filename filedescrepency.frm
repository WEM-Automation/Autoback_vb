VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "File Descrepency"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form2"
   ScaleHeight     =   2100
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Skip Copy"
      Height          =   615
      Left            =   4800
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Local File back to Source File"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Source File over local File"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    btn_copy = "source to dest"
    Unload Form2
End Sub

Private Sub Command2_Click()
    btn_copy = "source to dest"
    Unload Form2
End Sub

Private Sub Command3_Click()
    btn_copy = "abort"
    Unload Form2
End Sub

