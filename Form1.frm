VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   10800
      TabIndex        =   22
      Text            =   "Combo2"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton selesai 
      Caption         =   "Selesai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   21
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   360
      Left            =   6720
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   6720
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   6000
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   6720
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   6720
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   6720
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Beri Tanda Jika Kawin"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jenis Kelamin"
      Height          =   1455
      Left            =   9360
      TabIndex        =   3
      Top             =   2640
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "Wanita"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pria"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   6720
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   6720
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   6720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label10 
      Caption         =   "Jumlah Anak"
      Height          =   255
      Left            =   9360
      TabIndex        =   23
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Gaji Bersih"
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Pajak"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Tunjangan Anak"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Tunjangan Kawin"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Gaji Pokok"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Status Perkawinan"
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Bagian"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Pegawai"
      Height          =   615
      Left            =   4920
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Pegawai"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
 If Check1.Value = 1 Then
 Text4.Text = 0.1 * Val(Text3)
 If Val(Combo2) <= 3 Then
    Text5.Text = Val(Combo2) * 0.1 * Val(Text3)
    Else
    Text5.Text = 3 * 0.1 * Val(Text3)
    End If
 Else
 Text4.Text = 0
 Text5.Text = 0
 End If
End Sub

Private Sub Combo1_Click()
 If Combo1.Text = "Akuntansi" Then
 Text3.Text = "750000"
 ElseIf Combo1.Text = "Administrasi Umum" Then
 Text3.Text = "500000"
 ElseIf Combo1.Text = "Produksi" Then
 Text3.Text = "600000"
 Else
 Text3.Text = "500000"
 End If
 
 If Check1.Value = 1 Then
 Text4.Text = 0.1 * Val(Text3)
 If Val(Combo2) <= 3 Then
    Text5.Text = Val(Combo2) * 0.1 * Val(Text3)
    Else
    Text5.Text = 3 * 0.1 * Val(Text3)
    End If
 Else
 Text4.Text = 0
 Text5.Text = 0
 End If
End Sub

Private Sub Form_Activate()
 'Combo 1
 Combo1.AddItem "Akuntansi"
 Combo1.AddItem "Administrasi Umum"
 Combo1.AddItem "Produksi"
 Combo1.AddItem "Pengamanan"
 
 'Combo 2
 Combo2.AddItem "1"
 Combo2.AddItem "2"
 Combo2.AddItem "3"
 Combo2.AddItem "4"
 Combo2.AddItem "5"
 Combo2.AddItem "6"
 Combo2.AddItem "7"
 Combo2.AddItem "8"
 Combo2.AddItem "9"
 Combo2.AddItem "10"
End Sub

Private Sub selesai_Click()
 End
End Sub

Private Sub Text3_Change()
Text6.Text = 0.15 * Val(Text3) + Val(Text4) + Val(Text5)
Text7.Text = Val(Text3) + Val(Text4) + Val(Text5) - Val(Text6)
End Sub
