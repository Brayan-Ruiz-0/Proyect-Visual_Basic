VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MODA 
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   4695
      Left            =   3840
      TabIndex        =   9
      Top             =   1800
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8281
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   6720
      TabIndex        =   8
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000D&
      Caption         =   "Command4"
      Height          =   735
      Left            =   360
      MaskColor       =   &H008080FF&
      TabIndex        =   6
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MODA"
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MENU"
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SALIR"
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label MODA 
      Caption         =   "MODA"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   4
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "MODA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim MODA, ctm, ctt, ct2, l, n, v, h As Integer

ctm = 0
ct2 = 0

For j = 0 To ct Step 1
For i = 0 To at Step 1
Text3.Text = Text3.Text & "  " & m(i, j)
If ct2 < a + 1 Then

v = 0

For l = 0 To ct Step 1
For n = 0 To at Step 1

If ctt < a + 1 Then



If m(i, j) = mm(n, l) Then

v = v + 1

mm(n, l) = 0
End If

ctt = ctt + 1
Else

i = a + 1
j = a + 1
End If


Next

Next


If v > h Then
h = v
MODA = m(i, j)
End If
End If

ct2 = ct2 + 1
List1.AddItem (Text3.Text)
Text3.Text = ""
Next
Next






Text1.Text = MODA
Text2.Text = h


End Sub

Private Sub Command2_Click()
Form2.Show

End Sub

Private Sub Command3_Click()
End

End Sub

