VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7200
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Color de fondo del objeto"
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   5760
      Width           =   2055
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   4920
      List            =   "Form1.frx":0016
      TabIndex        =   7
      Text            =   "Color de letra"
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0049
      Left            =   360
      List            =   "Form1.frx":004B
      TabIndex        =   6
      Text            =   "Tipos de fuente"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   5880
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "SUBRAYADO"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "CURSIVA"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "NEGRITA"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Text            =   "Tamaño de fuente"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1560
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Text1.Font = Combo1.Text
End Sub

Private Sub Combo2_Click()
Text1.Font = Combo2.Text
Text1.FontSize = Combo2.Text
End Sub

Private Sub Combo3_Click()
If Combo3.Text = "Amarillo" Then
Text1.ForeColor = vbYellow
End If
If Combo3.Text = "Azul" Then
Text1.ForeColor = vbBlue
End If
If Combo3.Text = "Rojo" Then
Text1.ForeColor = vbRed
End If
If Combo3.Text = "Verde" Then
Text1.ForeColor = vbGreen
End If
If Combo3.Text = "Naranja" Then
Text1.ForeColor = &H80FF&
End If
If Combo3.Text = "Celeste" Then
Text1.ForeColor = &HFFFF00
End If
End Sub

Private Sub Command1_Click()
Text1.FontBold = clean
Text1.FontItalic = clean
Text1.FontUnderline = clean
Combo1.Text = clean
Combo2.Text = clean
End Sub

Private Sub Command2_Click()
CD1.ShowColor
Text1.BackColor = CD1.Color
End Sub

Private Sub Form_Load()
For x = 0 To Screen.FontCount - 1
Combo1.AddItem Screen.Fonts(x)
Next x
For x = 5 To 72
Combo2.AddItem x
Next x
Combo1.Text = Tahoma
Combo2.Text = 9
End Sub

Private Sub Option1_Click()
Text1.FontBold = True
End Sub

Private Sub Option2_Click()
Text1.FontItalic = True
End Sub

Private Sub Option3_Click()
Text1.FontUnderline = True
End Sub
