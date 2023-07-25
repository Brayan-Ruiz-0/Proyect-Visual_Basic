VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MEDIA 
   Caption         =   "Form3"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15015
   LinkTopic       =   "Form3"
   ScaleHeight     =   7635
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6135
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10821
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SALIR"
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MENU"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULAR"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   2895
   End
End
Attribute VB_Name = "MEDIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form2.Show

End Sub

Private Sub Command3_Click()
End

End Sub

