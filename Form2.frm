VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      Height          =   975
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MODA"
      Height          =   1215
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MEDIANA"
      Height          =   1215
      Index           =   3
      Left            =   5040
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MEDIA"
      Height          =   1215
      Index           =   4
      Left            =   2640
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
MODA.Show

End Sub

Private Sub Command2_Click(Index As Integer)
End
MEDIANA.Show

End Sub

Private Sub Command4_Click(Index As Integer)
End

End Sub
