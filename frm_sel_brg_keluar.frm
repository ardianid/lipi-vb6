VERSION 5.00
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_sel_brg_keluar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleksi"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_sel_brg_keluar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3945
      ScaleWidth      =   5985
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   27
         Top             =   3840
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox Text1 
            Height          =   320
            Left            =   1320
            TabIndex        =   28
            Text            =   "Semua"
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cabang"
            Height          =   210
            Index           =   4
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   210
            Index           =   3
            Left            =   1080
            TabIndex        =   29
            Top             =   360
            Width           =   60
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   5775
         Begin VB.TextBox Txt_Alamat 
            Height          =   320
            Left            =   1560
            TabIndex        =   8
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox Txt_Telp 
            Height          =   320
            Left            =   1560
            TabIndex        =   7
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox Txt_Nama 
            Height          =   320
            Left            =   1560
            TabIndex        =   6
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox Txt_Kode 
            Height          =   320
            Left            =   1560
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
         Begin MSMask.MaskEdBox Tgl_Lhr1 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Tgl_Lhr2 
            Height          =   315
            Left            =   3480
            TabIndex        =   10
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Tgl_Masuk1 
            Height          =   315
            Left            =   2880
            TabIndex        =   11
            Top             =   2520
            Visible         =   0   'False
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
            Left            =   4800
            TabIndex        =   12
            Top             =   2520
            Visible         =   0   'False
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
            Caption         =   "Nama Cust"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat Cust"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   25
            Top             =   1800
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   210
            Index           =   11
            Left            =   1440
            TabIndex        =   24
            Top             =   1440
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   210
            Index           =   12
            Left            =   1440
            TabIndex        =   23
            Top             =   1800
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/D"
            Height          =   210
            Index           =   19
            Left            =   4440
            TabIndex        =   22
            Top             =   2520
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/D"
            Height          =   210
            Index           =   18
            Left            =   3120
            TabIndex        =   21
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   210
            Index           =   13
            Left            =   2760
            TabIndex        =   20
            Top             =   2520
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   210
            Index           =   10
            Left            =   1440
            TabIndex        =   19
            Top             =   1080
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   210
            Index           =   9
            Left            =   1440
            TabIndex        =   18
            Top             =   720
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   210
            Index           =   8
            Left            =   1440
            TabIndex        =   17
            Top             =   360
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Order"
            Height          =   195
            Index           =   2
            Left            =   1560
            TabIndex        =   16
            Top             =   2520
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Trans"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Brg"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Bukti"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   585
         End
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   5775
         Begin VB.OptionButton Opt_Kriteria 
            Caption         =   "&Berdasarkan Kriteria"
            Height          =   255
            Left            =   2520
            TabIndex        =   3
            Top             =   360
            Width           =   2175
         End
         Begin VB.OptionButton Opt_Semua 
            Caption         =   "&Semua"
            Height          =   255
            Left            =   360
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
      End
      Begin IsButton_Ard.isButton Cmd_Lihat 
         Height          =   495
         Left            =   3720
         TabIndex        =   31
         Top             =   3360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Icon            =   "frm_sel_brg_keluar.frx":27C92
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
         TabIndex        =   32
         Top             =   3360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Icon            =   "frm_sel_brg_keluar.frx":27CAE
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
Attribute VB_Name = "frm_sel_brg_keluar"
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
    
    If Opt_Semua.Value = True Then
    
    sql = "select  * from VIEW_BrgKeluar"
    
'        If UCase(Text1.Text) <> UCase("Semua") Then
'            sql = sql & " where kode_counter in (select kode_counter from view_counter_user where nama_counter like '%" & Trim(Text1.Text) & "%' and id_user=" & Flag_tempat & ")"
'        Else
'            sql = sql & " where kode_counter in (select kode_counter from view_counter_user where id_user=" & Flag_tempat & ")"
'        End If
    
    sql = sql & " order by nobukti,tgl asc"
    
    Else
    
    If Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or Tgl_Lhr1.Text <> "__/__/____" Or Tgl_Lhr2.Text <> "__/__/____" Or Txt_Alamat.Text <> "" Or _
        Txt_Telp.Text <> "" Or Tgl_Masuk1.Text <> "__/__/____" Or Tgl_Masuk2.Text <> "__/__/____" Then
        
        sql = "select * from VIEW_BrgKeluar where"
        
'        If UCase(Text1.Text) <> UCase("semua") Then
'            sql = sql & " where kode_counter in (select kode_counter from view_counter_user where nama_counter like '%" & Trim(Text1.Text) & "%' and id_user=" & Flag_tempat & ")"
'        Else
'            sql = sql & " where kode_counter in (select kode_counter from view_counter_user where id_user=" & Flag_tempat & ")"
'        End If
        
        If Txt_Kode.Text <> "" Then
            sql = sql & " nobukti like '%" & Trim(Txt_Kode.Text) & "%'"
        End If
        
        If Txt_Nama.Text <> "" And Txt_Kode.Text = "" Then
            sql = sql & " namabrg like '%" & Trim(Txt_Nama.Text) & "%'"
        End If
        
        If Txt_Nama.Text <> "" And Txt_Kode.Text <> "" Then
            sql = sql & " and namabrg like '%" & Trim(Txt_Nama.Text) & "%'"
        End If
        
        If Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
            sql = sql & " tgl >='" & Format(Trim(Tgl_Lhr1.Text), "yyyy/mm/dd") & "' and Tgl <='" & Format(Trim(Tgl_Lhr2.Text), "yyyy/mm/dd") & "'"
        End If
        
        If Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____" And (Txt_Nama.Text <> "" And Txt_Kode.Text <> "") Then
            sql = sql & " and tgl >='" & Format(Trim(Tgl_Lhr1.Text), "yyyy/mm/dd") & "' and Tgl <='" & Format(Trim(Tgl_Lhr2.Text), "yyyy/mm/dd") & "'"
        End If
        
        If Txt_Alamat.Text <> "" And Tgl_Lhr1.Text = "__/__/____" And Tgl_Lhr2.Text = "__/__/____" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
            sql = sql & " nama_cust like '%" & Trim(Txt_Alamat.Text) & "%'" ' or Alamat_2 like '%" & Trim(Txt_Alamat.Text) & "%' or Alamat_3 like '%" & Trim(Txt_Alamat.Text) & "%'"
        End If

        If Txt_Alamat.Text <> "" And ((Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____") Or Txt_Nama.Text <> "" Or Txt_Kode.Text <> "") Then
            sql = sql & " and nama_cust like '%" & Trim(Txt_Alamat.Text) & "%'" ' or Alamat_2 like '%" & Trim(Txt_Alamat.Text) & "%' or Alamat_3 like '%" & Trim(Txt_Alamat.Text) & "%'"
        End If
        
        If Txt_Telp.Text <> "" And Txt_Alamat.Text = "" And Tgl_Lhr1.Text = "__/__/____" And Tgl_Lhr2.Text = "__/__/____" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
            sql = sql & " alamat_cust like '%" & Trim(Txt_Telp.Text) & "%'" ' or No_telp_Hp like '%" & Trim(Txt_Telp.Text) & "%'"
        End If

        If Txt_Telp.Text <> "" And (Txt_Alamat.Text <> "" Or (Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____") Or Txt_Nama.Text <> "" Or Txt_Kode.Text <> "") Then
            sql = sql & " and alamat_cust like '%" & Trim(Txt_Telp.Text) & "%'" ' or No_telp_Hp like '%" & Trim(Txt_Telp.Text) & "%'"
        End If

'        If Tgl_Masuk1.Text <> "__/__/____" And Tgl_Masuk2.Text <> "__/__/____" And Txt_Telp.Text = "" And Txt_Alamat.Text = "" And Tgl_Lhr1.Text = "__/__/____" And Tgl_Lhr2.Text = "__/__/____" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
'            sql = sql & " tgl_order >= '" & Format(Trim(Tgl_Masuk1.Text), "yyyy/mm/dd") & "' and tgl_order <='" & Format(Trim(Tgl_Masuk2.Text), "yyyy/mm/dd") & "'"
'        End If
'
'        If Tgl_Masuk1.Text <> "__/__/____" And Tgl_Masuk2.Text <> "__/__/____" And (Txt_Telp.Text <> "" Or Txt_Alamat.Text <> "" Or (Tgl_Lhr1.Text <> "__/__/____" And Tgl_Lhr2.Text <> "__/__/____") Or Txt_Nama.Text <> "" Or Txt_Kode.Text <> "") Then
'            sql = sql & " and tgl_order >= '" & Format(Trim(Tgl_Masuk1.Text), "yyyy/mm/dd") & "' and tgl_order <='" & Format(Trim(Tgl_Masuk2.Text), "yyyy/mm/dd") & "'"
'        End If
        
        sql = sql & " order by nobukti,tgl asc"
        
        
    Else
        
        Dim konfirm As Integer
            konfirm = CInt(MsgBox("Kriteria pencarian harus diisi", vbOKOnly + vbInformation, "Informasi"))
        
        Exit Sub
    End If
    
    End If
    
    
    Mysq = sql
    
'    Load Frm_Lap_Karyawan
        frm_lap_barang_keluar.Show
    
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
    
    Opt_Semua.Value = True
    
    Text1.Text = "Semua"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If
    
End Sub

Private Sub Opt_Kriteria_Click()
    
    If Opt_Kriteria.Value = True Then Frame2.Enabled = True
    
End Sub

Private Sub Opt_Kriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Kode.SetFocus
End Sub

Private Sub Opt_Semua_Click()
    If Opt_Semua.Value = True Then
        Frame2.Enabled = False
    
    Dim a As Object
        For Each a In Me
            If TypeOf a Is TextBox Then
                a.Text = ""
            End If
            
            If TypeOf a Is MaskEdBox Then a.Text = "__/__/____"
        Next
        
        Set a = Nothing
    
    End If
End Sub

Private Sub Opt_Semua_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Lihat.SetFocus
End Sub

Private Sub Text1_GotFocus()
    Call Focus_(Text1)
End Sub

Private Sub Text1_LostFocus()
    
    If Text1.Text = "" Then Text1.Text = "Semua"
    
End Sub

Private Sub Tgl_Lhr1_GotFocus()
    Call Focus_(Tgl_Lhr1)
End Sub

Private Sub Tgl_Lhr1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Tgl_Lhr2.SetFocus
End Sub

Private Sub Tgl_Lhr2_GotFocus()
    Call Focus_(Tgl_Lhr2)
End Sub

Private Sub Tgl_Lhr2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Alamat.SetFocus
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
    If KeyCode = 13 Then Txt_Alamat.SetFocus
End Sub

Private Sub Txt_Alamat_GotFocus()
    Call Focus_(Txt_Alamat)
End Sub

Private Sub Txt_Alamat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Telp.SetFocus
    End If
End Sub


Private Sub Txt_Kode_GotFocus()
    Call Focus_(Txt_Kode)
End Sub

Private Sub Txt_Kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Nama.SetFocus
End Sub

Private Sub Txt_Nama_GotFocus()
    Call Focus_(Txt_Nama)
End Sub

Private Sub Txt_Nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Tgl_Lhr1.SetFocus
End Sub



Private Sub Txt_Telp_GotFocus()
    Call Focus_(Txt_Telp)
End Sub

Private Sub Txt_Telp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Lihat.Enabled = True Then Cmd_Lihat.SetFocus
    End If
End Sub








