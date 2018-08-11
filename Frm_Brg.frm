VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form Frm_Brg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Barang"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Brg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Cari 
      Height          =   2295
      Left            =   -5160
      TabIndex        =   35
      Top             =   4560
      Visible         =   0   'False
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   4048
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Brg.frx":27C92
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Brg.frx":27CAE
      Childs          =   "Frm_Brg.frx":27D5A
      Begin VB.CommandButton Cmd_Cancel 
         Caption         =   "&Keluar"
         Height          =   405
         Left            =   4680
         TabIndex        =   37
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Txt_Cr_Kode 
         Height          =   320
         Left            =   2160
         TabIndex        =   38
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Cmd_OK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   3720
         TabIndex        =   36
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Txt_Cr_Nama 
         Height          =   320
         Left            =   2160
         TabIndex        =   39
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox Txt_Cr_Nopol 
         Height          =   320
         Left            =   2160
         TabIndex        =   40
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   47
         Top             =   120
         Width           =   960
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   5400
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   46
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   45
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   2
         Left            =   2040
         TabIndex        =   44
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   3
         Left            =   2040
         TabIndex        =   43
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ket"
         Height          =   240
         Index           =   4
         Left            =   360
         TabIndex        =   42
         Top             =   2400
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   240
         Index           =   5
         Left            =   2040
         TabIndex        =   41
         Top             =   2400
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   5895
      Left            =   -5160
      TabIndex        =   48
      Top             =   4080
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   10398
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Brg.frx":27D76
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Brg.frx":27D92
      Childs          =   "Frm_Brg.frx":27E3E
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   50
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   51
         Top             =   600
         Width           =   2535
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   2
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   5175
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   4695
         Left            =   240
         OleObjectBlob   =   "Frm_Brg.frx":27E5A
         TabIndex        =   52
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         Height          =   195
         Index           =   14
         Left            =   2280
         TabIndex        =   55
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   210
         Index           =   15
         Left            =   240
         TabIndex        =   54
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN TYPE BARANG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   240
         TabIndex        =   53
         Top             =   120
         Width           =   2385
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   10575
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4545
         ScaleWidth      =   10305
         TabIndex        =   20
         Top             =   3960
         Width           =   10335
         Begin TrueOleDBGrid60.TDBGrid Grid_Status 
            Height          =   4335
            Left            =   120
            OleObjectBlob   =   "Frm_Brg.frx":2A7B2
            TabIndex        =   21
            Top             =   120
            Width           =   10095
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   120
         ScaleHeight     =   2985
         ScaleWidth      =   10305
         TabIndex        =   1
         Top             =   120
         Width           =   10335
         Begin VB.Frame Frame2 
            Caption         =   "Barang"
            Height          =   1815
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   10095
            Begin VB.CheckBox cek_stock 
               Alignment       =   1  'Right Justify
               Caption         =   "&Stock"
               Enabled         =   0   'False
               Height          =   195
               Left            =   2880
               TabIndex        =   56
               Top             =   360
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.TextBox txt_satuan 
               Height          =   315
               Left            =   1080
               TabIndex        =   12
               Top             =   1080
               Width           =   1695
            End
            Begin VB.TextBox txt_estimasi 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   1440
               TabIndex        =   11
               Text            =   "0"
               Top             =   2040
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txt_jenis_kend 
               Height          =   315
               Left            =   1080
               TabIndex        =   10
               Top             =   720
               Width           =   6015
            End
            Begin VB.TextBox txt_kode_kend 
               Height          =   315
               Left            =   1080
               TabIndex        =   9
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txt_ket 
               Height          =   360
               Left            =   1440
               TabIndex        =   13
               Top             =   1920
               Visible         =   0   'False
               Width           =   3495
            End
            Begin TDBNumber6Ctl.TDBNumber tdb_harga 
               Height          =   315
               Left            =   1080
               TabIndex        =   58
               Top             =   1440
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               Calculator      =   "Frm_Brg.frx":2FB7A
               Caption         =   "Frm_Brg.frx":2FB9A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Frm_Brg.frx":2FC06
               Keys            =   "Frm_Brg.frx":2FC24
               Spin            =   "Frm_Brg.frx":2FC6E
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999
               MinValue        =   -999999999
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   1
               Value           =   0
               MaxValueVT      =   1028849669
               MinValueVT      =   1598423045
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Harga :"
               Height          =   195
               Index           =   2
               Left            =   435
               TabIndex        =   57
               Top             =   1440
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               Height          =   240
               Index           =   9
               Left            =   1320
               TabIndex        =   19
               Top             =   2040
               Visible         =   0   'False
               Width           =   75
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estimasi Pakai"
               Height          =   195
               Index           =   8
               Left            =   240
               TabIndex        =   18
               Top             =   2040
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama :"
               Height          =   195
               Index           =   6
               Left            =   480
               TabIndex        =   17
               Top             =   720
               Width           =   510
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode :"
               Height          =   195
               Index           =   4
               Left            =   480
               TabIndex        =   16
               Top             =   360
               Width           =   465
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Satuan :"
               Height          =   195
               Index           =   10
               Left            =   360
               TabIndex        =   15
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ket"
               Height          =   240
               Index           =   12
               Left            =   240
               TabIndex        =   14
               Top             =   1920
               Visible         =   0   'False
               Width           =   270
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Type Barang"
            Height          =   1095
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   10095
            Begin VB.CommandButton cmd_browse_ex 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2400
               TabIndex        =   4
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txt_kode_ex 
               Height          =   315
               Left            =   1080
               TabIndex        =   3
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lbl_nama_ex 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1080
               TabIndex        =   7
               Top             =   720
               Width           =   6015
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama :"
               Height          =   195
               Index           =   1
               Left            =   480
               TabIndex        =   6
               Top             =   720
               Width           =   510
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode :"
               Height          =   195
               Index           =   0
               Left            =   480
               TabIndex        =   5
               Top             =   360
               Width           =   465
            End
         End
      End
      Begin VB.Frame v 
         Height          =   735
         Left            =   360
         TabIndex        =   22
         Top             =   3120
         Width           =   9615
         Begin VB.CommandButton cmd_keluar 
            Caption         =   "&Keluar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Cari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7200
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_hapus 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6360
            TabIndex        =   30
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_rubah 
            Caption         =   "&Rubah"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5520
            TabIndex        =   29
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_navigasi 
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2040
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmd_navigasi 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   25
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmd_navigasi 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   840
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmd_tambah 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   28
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Frame2"
            Height          =   855
            Left            =   3480
            TabIndex        =   27
            Top             =   0
            Width           =   15
         End
         Begin VB.CommandButton cmd_navigasi 
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmd_simpan 
            Caption         =   "&Simpan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_batal 
            Caption         =   "&Batal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5520
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "Frm_Brg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim Moving As Boolean
Dim yold, xold As Long
Dim tampilkan As Boolean

Private Sub IsiSemua()
    
    Dim sql As String
        sql = "select * from VIEW_Barang order by Nama_Jenis,Nama,Kode desc"
        
        Set Rs_Nav = New ADODB.Recordset
            Rs_Nav.Open sql, kon, adOpenKeyset
        
        Set Grid_Status.DataSource = Rs_Nav
            Grid_Status.Refresh
    
End Sub


Private Sub cek_stock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub Cmd_Batal_Click()

    rubah = False
    tampilkan = True
    
    Cmd_Simpan.Visible = False
    Cmd_Tambah.Visible = True
    Cmd_Tambah.Enabled = True
    Cmd_Batal.Visible = False
    Cmd_Rubah.Visible = True
    Cmd_Rubah.Enabled = True
    Cmd_Hapus.Enabled = True
    Cmd_Cari.Enabled = True
    Cmd_Keluar.Enabled = True
    'frame_nav.Enabled = False
    
    Cmd_Simpan.Enabled = True
        
    Dim n As Object
        For Each n In Me
            If TypeOf n Is TextBox Then
                If Left(n.Name, 6) <> "Txt_Cr" Then
                    n.Enabled = False
                End If
            End If
            
            If TypeOf n Is TDBNumber Then n.Enabled = False
            
            If TypeOf n Is TDBContainer3D Then
                n.Visible = False
            End If
            
            If TypeOf n Is CommandButton Then
                If n.Caption = "..." Then n.Enabled = False
            End If
            
            
        Next
    Set n = Nothing
    
    Cmd_Tambah.SetFocus
    
    txt_estimasi.Text = 0
    cek_stock.Enabled = False
    
    If Rs_Nav.State = adStateOpen Then
        If Rs_Nav.RecordCount > 0 Then Rs_Nav.MoveLast
    End If
    
End Sub

Private Sub Cmd_Cancel_Click()
    
    Cmd_Tambah.Enabled = True
    Cmd_Rubah.Visible = True
    Cmd_Batal.Visible = False
    Cmd_Hapus.Enabled = True
    Cmd_Cari.Enabled = True
    Cmd_Keluar.Enabled = True
    
    TDB_Cari.Visible = False
End Sub

Private Sub cmd_browse_ex_Click()
    
    With TDB_Daftar
        
        .Left = Picture3.Left + Picture1.Left + Frame1.Left + cmd_browse_ex.Left + cmd_browse_ex.Width / 2 - .Width / 2
        .Top = Picture3.Top + Picture1.Top + Frame1.Top + cmd_browse_ex.Top + cmd_browse_ex.Height + 15
        
        If .Visible = False Then
        
        Txt_Cr_Daftar(0).Text = ""
        Txt_Cr_Daftar(1).Text = ""
        txt_cr_daftar_KeyUp 0, 0, 0
        
        .Visible = True
        
        Txt_Cr_Daftar(0).SetFocus
        
        Else
            .Visible = False
        End If
    End With
End Sub

Private Sub Cmd_Cari_Click()
        
    If Rs_Nav.RecordCount <= 0 Then Exit Sub
            
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
'    Cmd_Cari.Enabled = False
    Cmd_Keluar.Enabled = False
        
    With TDB_Cari
        
        .Left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
        
        If .Visible = False Then
            Txt_Cr_Kode.Text = ""
            Txt_Cr_Nama.Text = ""
            .Visible = True
            Txt_Cr_Kode.SetFocus
        Else
            .Visible = False
        End If
        
    End With
    
End Sub

Private Sub Cmd_Hapus_Click()
On Error GoTo err_handler
    
    If Rs_Nav.RecordCount <= 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    
    If MsgBox("Yakin akan menghapus data barang " & txt_kode_kend.Text, vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then
        On Error GoTo 0
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "delete from Tb_Barang where Kode ='" & Trim(txt_kode_kend.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    Dim konfirm As Integer
'        Konfirm = CInt(MsgBox("Data telah dihapus", vbOKOnly + vbInformation, "Informasi"))
    
    IsiSemua
    
    On Error GoTo 0
    Exit Sub

err_handler:
    
    'Dim Konfirm As Integer
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Informasi"))
            Err.Clear

End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Navigasi_Click(Index As Integer)

On Error Resume Next

With Rs_Nav
    Select Case Index
        Case 0
            .MoveLast
        Case 1
            
            If .EOF Then .MoveLast
                
                .MoveNext
                
            If .EOF Then .MoveLast
            
        Case 2
            
            If .BOF Then .MoveFirst
                
                .MovePrevious
                
            If .BOF Then .MoveFirst
            
        Case 3
            
            .MoveFirst
            
    End Select
End With

Set Grid_Status.DataSource = Rs_Nav
    Grid_Status.Refresh

End Sub

Private Sub Cmd_Ok_Click()
On Error Resume Next

    With Rs_Nav
        .MoveFirst
        
        If Txt_Cr_Kode.Text <> "" Then
            .Find "Kode like '%" & Trim(Txt_Cr_Kode.Text) & "%'"
        ElseIf Txt_Cr_Nama.Text <> "" And Txt_Cr_Kode.Text = "" Then
            .Find "Nama_Jasa like '%" & Trim(Txt_Cr_Nama.Text) & "%'"
        ElseIf Txt_Cr_Nopol.Text <> "" And Txt_Cr_Kode.Text = "" And Txt_Cr_Nama.Text = "" Then
            .Find "Ket like '%" & Trim(Txt_Cr_Nama.Text) & "%'"
        End If
        
    End With
    
    Set Grid_Status.DataSource = Rs_Nav
        Grid_Status.Refresh
        
End Sub

Private Sub Cmd_Rubah_Click()
    
    If Rs_Nav.RecordCount <= 0 Then Exit Sub
    
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Cari.Enabled = False
    Cmd_Keluar.Enabled = False

    rubah = True
    txt_jenis_kend.Enabled = True
    tdb_harga.Enabled = True
    txt_satuan.Enabled = True
    txt_ket.Enabled = True
        
    txt_jenis_kend.SetFocus
    
End Sub

Private Sub cmd_simpan_Click()
On Error GoTo err_handler

Dim konfirm As Integer
    If txt_kode_kend.Text = "" Then
        konfirm = CInt(MsgBox("Kode barang tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))

        txt_kode_kend.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
    If txt_kode_ex.Text = "" Then
        konfirm = CInt(MsgBox("Kode jenis barang tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        txt_kode_ex.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
       
     Dim stock As String
        If cek_stock.Value = vbChecked Then
            stock = 1
        Else
            stock = 0
        End If
    
    Dim harga_br As Double
        If tdb_harga.ValueIsNull Then
            harga_br = 0
        Else
            harga_br = Replace(Trim(tdb_harga.Value), ",", "")
        End If
    
    Dim sql, sql1 As String
    Dim rs As Recordset
    Dim rs1 As Recordset
    
    If rubah = False Then
        
    sql = "insert into Tb_Barang (Kode,Kode_Jenis,Nama,Harga,Satuan,Stock) values('" & Trim(txt_kode_kend.Text) & "','" & Trim(txt_kode_ex.Text) & "','" & Trim(txt_jenis_kend.Text) & "'," & harga_br & ",'" & Trim(txt_satuan.Text) & "','" & stock & "')"
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    If stock = 1 Then
        

                        
        Dim sql2 As String
        Dim rs2 As Recordset
            sql2 = "select * from Tb_Jml_Stock where kode='" & Trim(txt_kode_kend.Text) & "'" ' and kode_counter='" & !kode & "'"
            Set rs2 = New ADODB.Recordset
                rs2.Open sql2, kon, adOpenKeyset
            If rs2.EOF Then
                
                Dim sqlstock As String
                Dim rsstock As Recordset
                    sqlstock = "insert into Tb_Jml_Stock (kode,jml_Baik,Jml_Rusak,Jml_Lengkap,Jml_Kurang)"
                    sqlstock = sqlstock & " values('" & Trim(txt_kode_kend.Text) & "',0,0,0,0)"
                    
                    Set rsstock = New ADODB.Recordset
                        rsstock.Open sqlstock, kon
                    
                
            End If
            
            Set rs2 = Nothing
        
    End If
    
    konfirm = CInt(MsgBox("Data telah disimpan", vbOKOnly + vbInformation, "Informasi"))
    
    IsiSemua
    
    txt_kode_kend.Text = ""
    txt_jenis_kend.Text = ""
    txt_ket.Text = ""
    tdb_harga.Value = Null
    
    
    txt_kode_kend.SetFocus
    
    Else
        
        sql = "update Tb_Barang set Nama='" & Trim(txt_jenis_kend.Text) & "',Harga=" & harga_br & ",Satuan='" & Trim(txt_satuan.Text) & "' where Kode='" & Trim(txt_kode_kend.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
            
        konfirm = CInt(MsgBox("Data telah dirubah", vbOKOnly + vbInformation, "Informasi"))
        
        IsiSemua
        Cmd_Batal_Click
        
    End If
    
    
    
    On Error GoTo 0
    Exit Sub
    
err_handler:
    
    konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Informaton"))
        Err.Clear
    
End Sub

Private Sub Cmd_Tambah_Click()

rubah = False
tampilkan = False
    
    cmd_browse_ex.Enabled = True
    txt_kode_ex.Enabled = True
    txt_kode_ex.Text = ""
    
    lbl_nama_ex.Caption = ""
    txt_kode_kend.Text = ""
    txt_jenis_kend.Text = ""
    txt_satuan.Text = ""
    tdb_harga.Value = Null
    
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Cari.Enabled = False
    Cmd_Keluar.Enabled = False
        
    txt_kode_ex.SetFocus

End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Cmd_Tambah.SetFocus
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
    .Left = Utama.Width / 2 - .Width / 2
    .Top = 150
End With

tampilkan = True

IsiSemua

rubah = False
txt_kode_ex.Enabled = False
cmd_browse_ex.Enabled = False
txt_kode_kend.Enabled = False
txt_jenis_kend.Enabled = False
txt_ket.Enabled = False
txt_estimasi.Enabled = False
txt_satuan.Enabled = False
cek_stock.Enabled = False
tdb_harga.Enabled = False

End Sub


Private Sub Form_Unload(Cancel As Integer)
        
    If Cmd_Keluar.Enabled = False Then
        Cancel = True
    Else
        Cancel = False
        
        If kon.State = adStateOpen Then
            
            kon.Close
            Set kon = Nothing
        End If
                
    End If
    
    
End Sub

Private Sub grid_daftar_DblClick()

If Grid_Daftar.Row < 0 Then Exit Sub

txt_kode_ex.Text = Grid_Daftar.Columns(0).Text
lbl_nama_ex.Caption = Grid_Daftar.Columns(1).Text

TDB_Daftar.Visible = False

txt_kode_kend.Enabled = True
txt_jenis_kend.Enabled = True
'tdb_harga.Enabled = True
'tdb_komisi.Enabled = True
txt_ket.Enabled = True
tdb_harga.Enabled = True
txt_satuan.Enabled = True
'cek_stock.Enabled = True

txt_kode_kend.SetFocus

End Sub

Private Sub grid_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_daftar_DblClick
    If KeyCode = vbKeyEscape Then TDB_Daftar.Visible = False: txt_kode_ex.SetFocus
End Sub

Private Sub Grid_Status_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If tampilkan = False Then Exit Sub

  With Rs_Nav
    If .RecordCount = 0 Then
        
        txt_kode_ex.Text = ""
        lbl_nama_ex.Caption = ""
        txt_kode_kend.Text = ""
        txt_jenis_kend.Text = ""
        tdb_harga.Value = Null
        txt_satuan.Text = ""
        txt_ket.Text = ""
        
    Else
                
        txt_kode_ex.Text = IIf(Not IsNull(!Kode_Jenis), !Kode_Jenis, "")
        lbl_nama_ex.Caption = IIf(Not IsNull(!nama_jenis), !nama_jenis, "")
        txt_kode_kend.Text = !kode
        txt_jenis_kend.Text = IIf(Not IsNull(!nama), !nama, "")
        tdb_harga.Value = IIf(Not IsNull(!harga), !harga, Null)
        txt_satuan.Text = IIf(Not IsNull(!satuan), !satuan, 0)
        
        Dim sto As String
            sto = IIf(Not IsNull(!stock), !stock, 0)
            
            If sto = 0 Then
                cek_stock.Value = vbUnchecked
            Else
                cek_stock.Value = vbChecked
            End If
    End If
    End With
End Sub

Private Sub Label2_Click(Index As Integer)

End Sub

Private Sub TDB_Cari_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = True
If Moving = True Then
   yold = y
   xold = x
End If
End Sub

Private Sub TDB_Cari_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Moving = True Then
   TDB_Cari.Top = TDB_Cari.Top - (yold - y)
   TDB_Cari.Left = TDB_Cari.Left - (xold - x)
End If

End Sub

Private Sub TDB_Cari_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = False
End Sub

Private Sub tdb_harga_GotFocus()
    Call Focus_(tdb_harga)
End Sub

Private Sub tdb_harga_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Simpan.Enabled = True Then Cmd_Simpan.SetFocus
    End If
End Sub

Private Sub Txt_Cr_Kode_Change()
    Txt_Cr_Nama.Text = ""
    Txt_Cr_Nopol.Text = ""
End Sub

Private Sub Txt_Cr_Kode_GotFocus()
    Call Focus_(Txt_Cr_Kode)
End Sub

Private Sub Txt_Cr_Kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Cr_Nama.SetFocus
End Sub

Private Sub Txt_Cr_Nama_Change()
    Txt_Cr_Kode.Text = ""
    Txt_Cr_Nopol.Text = ""
End Sub

Private Sub Txt_Cr_Nama_GotFocus()
    Call Focus_(Txt_Cr_Nama)
End Sub

Private Sub Txt_Cr_Nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_OK.SetFocus
End Sub

Private Sub txt_Pend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub TDB_Daftar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = True
If Moving = True Then
   yold = y
   xold = x
End If
End Sub

Private Sub TDB_Daftar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Moving = True Then
   TDB_Daftar.Top = TDB_Daftar.Top - (yold - y)
   TDB_Daftar.Left = TDB_Daftar.Left - (xold - x)
End If

End Sub

Private Sub TDB_Daftar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = False
End Sub

Private Sub txt_cr_daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then TDB_Daftar.Visible = False: txt_kode_ex.SetFocus
End Sub

Private Sub txt_cr_daftar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "select Kode,Jenis from Tb_Type_Brg"
    
    Select Case Index
        Case 0
            sql = sql & " where Kode like '%" & Trim(Txt_Cr_Daftar(0).Text) & "%'"
        Case 1
            sql = sql & " where Jenis like '%" & Trim(Txt_Cr_Daftar(1).Text) & "%'"
    End Select
    
    sql = sql & " order by Kode asc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set Grid_Daftar.DataSource = rs
        Grid_Daftar.Refresh
    
End Sub

Private Sub Txt_Cr_Nopol_Change()
    Txt_Cr_Kode.Text = ""
    Txt_Cr_Nama.Text = ""
End Sub

Private Sub Txt_Cr_Nopol_GotFocus()
    Call Focus_(Txt_Cr_Nopol)
End Sub

Private Sub Txt_Cr_Nopol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_OK.SetFocus
End Sub

Private Sub txt_estimasi_GotFocus()
    Call Focus_(txt_estimasi)
End Sub

Private Sub txt_estimasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_satuan.SetFocus
End Sub

Private Sub txt_estimasi_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_jenis_kend_GotFocus()
    Call Focus_(txt_jenis_kend)
End Sub

Private Sub txt_jenis_kend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_satuan.SetFocus
End Sub

Private Sub txt_ket_GotFocus()
    Call Focus_(txt_ket)
End Sub

Private Sub txt_ket_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If cek_stock.Enabled = True Then
            cek_stock.SetFocus
        Else
            Cmd_Simpan.SetFocus
        End If
    End If
End Sub

Private Sub txt_kode_ex_GotFocus()
    Call Focus_(txt_kode_ex)
End Sub

Private Sub txt_kode_ex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_kode_ex_LostFocus
    If KeyCode = vbKeyF3 Then txt_kode_kend.Text = "": cmd_browse_ex_Click
End Sub

Private Sub txt_kode_ex_LostFocus()

    If txt_kode_ex.Text = "" Then Exit Sub
    
    Dim comd As Command
    Set comd = New ADODB.Command
    With comd
        .ActiveConnection = kon
        .CommandText = "cari_jenis_brg"
        .CommandType = adCmdStoredProc
        .Parameters("@kode").Value = Trim(txt_kode_ex.Text)
        
        .Execute
        
    End With
    
    Dim rs As Recordset
        Set rs = New ADODB.Recordset
            rs.Open comd
            
            Dim konfirm As Integer
            
            With rs
                If Not .EOF Then
                    
                    lbl_nama_ex.Caption = IIf(Not IsNull(!Jenis), !Jenis, "")
                    
                    txt_kode_kend.Enabled = True
                    txt_jenis_kend.Enabled = True
'                    tdb_harga.Enabled = True
'                    tdb_komisi.Enabled = True
                    txt_ket.Enabled = True
                    tdb_harga.Enabled = True
                    txt_satuan.Enabled = True
                    cek_stock.Enabled = True
                    
                    txt_kode_kend.SetFocus
                    
                Else
                    
                    konfirm = CInt(MsgBox("Kode jenis barang yang anda masukkan tidak ditemukan", vbOKOnly + vbInformation, "informasi"))
                    
                    txt_kode_ex.Text = ""
                    lbl_nama_ex.Caption = ""
                    
                    txt_kode_ex.SetFocus
                    
                End If
            End With
        
        Set comd.ActiveConnection = Nothing


End Sub

Private Sub txt_kode_kend_GotFocus()
    Call Focus_(txt_kode_kend)
End Sub

Private Sub txt_kode_kend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_jenis_kend.SetFocus
End Sub

Private Sub txt_kode_kend_LostFocus()

    If txt_kode_kend.Text = "" Then Exit Sub
    
    Dim comd As Command
    Set comd = New ADODB.Command
    With comd
        .ActiveConnection = kon
        .CommandText = "cari_kode_brg"
        .CommandType = adCmdStoredProc
        .Parameters("@kode").Value = Trim(txt_kode_kend.Text)
        
        .Execute
            
        If .Parameters("@ada") = 1 Then
            Dim konfirm As Integer
                konfirm = CInt(MsgBox("Kode barang yang anda masukkan sudah ada ...", vbOKOnly + vbInformation, "Informasi"))
                
                txt_kode_kend.Text = ""
                txt_kode_kend.SetFocus
        End If
            
    End With
    Set comd.ActiveConnection = Nothing

End Sub

Private Sub txt_satuan_GotFocus()
    Call Focus_(txt_satuan)
End Sub

Private Sub txt_satuan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdb_harga.SetFocus
End Sub
