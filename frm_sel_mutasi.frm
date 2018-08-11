VERSION 5.00
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_sel_mutasi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleksi"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_sel_mutasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5775
         Begin VB.TextBox Txt_Alamat 
            Height          =   320
            Left            =   1560
            TabIndex        =   2
            Top             =   840
            Width           =   3975
         End
         Begin MSMask.MaskEdBox Tgl_Masuk1 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Tgl_Masuk2 
            Height          =   315
            Left            =   3480
            TabIndex        =   4
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Order"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   210
            Index           =   13
            Left            =   1440
            TabIndex        =   8
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/D"
            Height          =   210
            Index           =   19
            Left            =   3120
            TabIndex        =   7
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   210
            Index           =   11
            Left            =   1440
            TabIndex        =   6
            Top             =   840
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Brg"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   690
         End
      End
      Begin IsButton_Ard.isButton Cmd_Lihat 
         Height          =   495
         Left            =   3720
         TabIndex        =   10
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Icon            =   "frm_sel_mutasi.frx":27C92
         Style           =   8
         Caption         =   "&Tampil"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin IsButton_Ard.isButton Cmd_Keluar 
         Height          =   495
         Left            =   4800
         TabIndex        =   11
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Icon            =   "frm_sel_mutasi.frx":27CAE
         Style           =   8
         Caption         =   "&Keluar"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
   End
End
Attribute VB_Name = "frm_sel_mutasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_Foto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Lihat.SetFocus
End Sub

Private Sub Cmd_Keluar_Click()

    Unload Me

End Sub

Private Sub Cmd_Lihat_Click()
    
    Dim sql As String
    
    If Tgl_Masuk1.Text = "__/__/____" Or Tgl_Masuk2.Text = "__/__/____" Then
        MsgBox "Periode harus diisi"
        Exit Sub
    End If
    
    If periksa_tanggal(Tgl_Masuk1.Text) = False Then
        MsgBox "Format tgl salah"
        Tgl_Masuk1.SetFocus
        Exit Sub
    End If
    
    If periksa_tanggal(Tgl_Masuk2.Text) = False Then
        MsgBox "Format tgl salah"
        Tgl_Masuk2.SetFocus
        Exit Sub
    End If
    
    sql = "select * from VIEW_Hist_Brg where tgl >='" & Format(Trim(Tgl_Masuk1.Text), "yyyy/mm/dd") & "' and tgl <='" & Format(Trim(Tgl_Masuk2.Text), "yyyy/mm/dd") & "'"
    
    If Txt_Alamat.Text <> "" Then
        sql = sql & " and namabrg like '%" & Trim(Txt_Alamat.Text) & "%'"
    End If
    
    Mysq = sql
    
    macem2 = Trim(Tgl_Masuk1.Text)
    macem2_lagi = Trim(Tgl_Masuk2.Text)
    
    frm_lap_mutasi_brg.Show
    
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Tgl_Masuk1.SetFocus
End Sub

Private Sub Form_Load()
    
Dim status As String
status = Buka_Koneksi
If status = "-2147467259" Then
    Dim konfirm As Integer
        konfirm = CInt(MsgBox("Koneksi terputus ....", vbOKOnly + vbInformation, "Informasi"))
        
        End
        Exit Sub
End If
    
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = 250
    End With


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If
    
 
End Sub


Private Sub Tgl_Masuk1_GotFocus()
    Call Focus_(Tgl_Masuk1)
End Sub

Private Sub Tgl_Masuk1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Tgl_Masuk2.SetFocus
End Sub

Private Sub Tgl_Masuk2_GotFocus()
    Call Focus_(Tgl_Masuk2)
End Sub

Private Sub Tgl_Masuk2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Alamat.SetFocus
    End If
        
End Sub

Private Sub Txt_Alamat_GotFocus()
    Call Focus_(Txt_Alamat)
End Sub

Private Sub Txt_Alamat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Lihat.SetFocus
End Sub


