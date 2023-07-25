VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4095
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.TextBox text1 
      Height          =   735
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox text2 
      Height          =   735
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox text3 
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MUESTRA"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      Height          =   735
      Left            =   8160
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MEDIDAS DE TENDENCIA "
      Height          =   735
      Left            =   6120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "MEDIDAS DE TENDENCIA CENTRAL"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim b, c, d, i, j, ctr, xy As Integer
ct = 1
a = Val(Text1.Text)
If a > 15 Then
i = 8
While i < a
If a Mod i = 0 Then
b = a / i
Text2.Text = b
Text3.Text = i
ct = i + 3
at = b + 3

mscuadro.Cols = ct
mscuadro.Rows = at

MEDIA.MSFlexGrid1.Cols = ct
MEDIA.MSFlexGrid1.Rows = at

MODA.MSFlexGrid2.Cols = ct
MODA.MSFlexGrid2.Rows = at




i = i + a
End If
i = i + 1
Wend


Else

i = 2

While i < a
If a Mod i = 0 Then
b = a / i
Text2.Text = b
Text3.Text = i
ct = i + 3
at = b + 3
mscuadro.Cols = ct
mscuadro.Rows = at

MEDIA.MSFlexGrid1.Cols = ct
MEDIA.MSFlexGrid1.Rows = at

MODA.MSFlexGrid2.Cols = ct
MODA.MSFlexGrid2.Rows = at

i = i + a
End If
i = i + 1
Wend
End If
ctr = 1
For j = 0 To ct - 1
For i = 0 To at - 1
If ctr < a + 1 Then

m(i, j) = Val(InputBox("digite el valor de la muestra", "No", ct))

mscuadro.TextMatrix(i, j) = m(i, j)

mm(i, j) = m(i, j)


MEDIA.MSFlexGrid1.TextMatrix(i, j) = m(i, j)

MODA.MSFlexGrid2.TextMatrix(i, j) = m(i, j)





ctr = ctr + 1
Else
i = i + a
j = j + a
End If

Next
Next

End Sub


Private Sub Command2_Click()
End

End Sub

Private Sub Command3_Click()
Form2.Show

End Sub

Private Sub Form_Load()
Dim i, j As Integer


MsgBox ("BIENVENIDOS A CONTINUACION DIGITE EL NOMBRE DE LA POBLACION")
msg = InputBox("digite el nombre")
tt = MsgBox("digite la muestra", vbdefaultbuttton1 = 0, "poblacion: ")
msg = InputBox(msg, "De la poblacion ", "Escriba la cantidad de la muestra estadistica ")
Text1.Text = msg





End Sub


Private Sub C_Change()

End Sub


End Sub
