VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menambahkan Baris Baru ke TextBox"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim NewText As String
    With Text1
    'Ganti teks 'My New Text' dengan teks yang Anda
    'inginkan ditambah
        NewText = "My New Text"
        .SelStart = Len(.Text)
        .SelText = vbNewLine & NewText
    End With
End Sub

Private Sub Form_Load()
    Text1.Text = "My Initial Text"
End Sub

