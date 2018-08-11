VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form BrgMasuk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Barang Masuk"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BrgMasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   3975
      Left            =   -840
      TabIndex        =   59
      Top             =   3600
      Visible         =   0   'False
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   7011
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "BrgMasuk.frx":27C92
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "BrgMasuk.frx":27CAE
      Childs          =   "BrgMasuk.frx":27D5A
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   61
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame Frame6 
         Height          =   135
         Index           =   2
         Left            =   240
         TabIndex        =   60
         Top             =   360
         Width           =   6495
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   2895
         Left            =   240
         OleObjectBlob   =   "BrgMasuk.frx":27D76
         TabIndex        =   65
         Top             =   960
         Width           =   6495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         Height          =   195
         Index           =   40
         Left            =   360
         TabIndex        =   29
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl"
         Height          =   195
         Index           =   41
         Left            =   3480
         TabIndex        =   43
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   55
         Top             =   120
         Width           =   945
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Brg 
      Height          =   3615
      Left            =   3240
      TabIndex        =   41
      Top             =   6600
      Visible         =   0   'False
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "BrgMasuk.frx":2C68A
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "BrgMasuk.frx":2C6A6
      Childs          =   "BrgMasuk.frx":2C752
      Begin VB.TextBox TxtCr_Brg 
         Height          =   300
         Index           =   1
         Left            =   2880
         TabIndex        =   48
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox TxtCr_Brg 
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   47
         Top             =   600
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   46
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   45
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   44
         Top             =   600
         Width           =   405
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TdbAdd 
      Height          =   5535
      Left            =   2400
      TabIndex        =   33
      Top             =   5160
      Visible         =   0   'False
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   9763
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "BrgMasuk.frx":2C76E
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "BrgMasuk.frx":2C78A
      Childs          =   "BrgMasuk.frx":2C836
      Begin TrueOleDBGrid60.TDBGrid Gridadd 
         Height          =   3735
         Left            =   240
         OleObjectBlob   =   "BrgMasuk.frx":2C852
         TabIndex        =   64
         Top             =   1080
         Width           =   6495
      End
      Begin VB.CommandButton CmdFinish 
         Caption         =   "&FINISH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   40
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   39
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox TSatuan 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TNamaBrg 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl :"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   37
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti :"
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   35
         Top             =   600
         Width           =   690
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   6720
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B A R A N G  M A S U K ( B E R D A S A R K A N  O R D E R )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   120
         Width           =   4590
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Cbang 
      Height          =   3615
      Left            =   3000
      TabIndex        =   53
      Top             =   6720
      Visible         =   0   'False
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "BrgMasuk.frx":31396
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "BrgMasuk.frx":313B2
      Childs          =   "BrgMasuk.frx":3145E
      Begin VB.TextBox TxtCrCbang 
         Height          =   285
         Left            =   960
         TabIndex        =   58
         Top             =   600
         Width           =   2895
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   57
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   56
         Top             =   120
         Width           =   945
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDBSupp 
      Height          =   3615
      Left            =   1560
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "BrgMasuk.frx":3147A
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "BrgMasuk.frx":31496
      Childs          =   "BrgMasuk.frx":31542
      Begin VB.TextBox Txt_Cr_Supp 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   28
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Txt_Cr_Supp 
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   27
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   5295
      End
      Begin TrueOleDBGrid60.TDBGrid GridSupp 
         Height          =   2415
         Left            =   240
         OleObjectBlob   =   "BrgMasuk.frx":3155E
         TabIndex        =   63
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl"
         Height          =   195
         Index           =   14
         Left            =   3360
         TabIndex        =   32
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   240
         TabIndex        =   30
         Top             =   120
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6495
      ScaleWidth      =   10335
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.TextBox txtnama 
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   1200
         Width           =   8655
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5760
         TabIndex        =   15
         Top             =   5640
         Width           =   4455
         Begin VB.CommandButton Cmd_Keluar 
            Caption         =   "&Keluar"
            Height          =   495
            Left            =   3480
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Daftar 
            Caption         =   "&Cari"
            Height          =   495
            Left            =   2640
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Hapus 
            Caption         =   "&Hapus"
            Height          =   495
            Left            =   1800
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Rubah 
            Caption         =   "&Rubah"
            Height          =   495
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Tambah 
            Caption         =   "&Tambah"
            Height          =   495
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Simpan 
            Caption         =   "&Simpan"
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Batal 
            Caption         =   "&Batal"
            Height          =   495
            Left            =   960
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "+"
            Height          =   495
            Left            =   1800
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdDel 
            Caption         =   "-"
            Height          =   495
            Left            =   2640
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame_Nav 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   5640
         Width           =   2175
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   1560
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   1080
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   600
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton CmdSupp 
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
         Height          =   275
         Left            =   4800
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox TNamaSupp 
         Height          =   300
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TKodeSupp 
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox TBukti 
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTgl 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   49217537
         CurrentDate     =   39467
      End
      Begin TrueOleDBGrid60.TDBGrid GridBrg 
         Height          =   3855
         Left            =   0
         OleObjectBlob   =   "BrgMasuk.frx":344DF
         TabIndex        =   62
         Top             =   1680
         Width           =   10215
      End
      Begin VB.TextBox TKodeCbang 
         Height          =   300
         Left            =   3600
         TabIndex        =   50
         Top             =   3240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TNamaCbang 
         Height          =   300
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   3240
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CommandButton CmdCbang 
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
         Height          =   275
         Left            =   8400
         TabIndex        =   52
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pemohon :"
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   66
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Brg :"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bukti :"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl :"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang :"
         Height          =   195
         Index           =   6
         Left            =   2775
         TabIndex        =   49
         Top             =   3240
         Visible         =   0   'False
         Width           =   660
      End
   End
End
Attribute VB_Name = "BrgMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim Moving As Boolean
Dim yold, xold As Long
Dim ArrBrg As New XArrayDB
Dim TotalOld As Double
Dim arradd As New XArrayDB

Private Sub isi_semua(ByVal rec As Recordset)
    With rec
        
        If .EOF Then .MoveLast
        If .BOF Then .MoveFirst
        
        TBukti.Text = IIf(Not IsNull(!nobukti), !nobukti, "")
        DTgl.Value = !tgl
        TKodeSupp.Text = IIf(Not IsNull(!bukti_order), !bukti_order, "")
        TNamaSupp.Text = IIf(Not IsNull(!tgl_order), !tgl_order, "")
        txtnama.Text = IIf(Not IsNull(!atas_nama), !atas_nama, "")
'        TKodeCbang.Text = IIf(Not IsNull(!kodecounter), !kodecounter, "")
'        TNamaCbang.Text = IIf(Not IsNull(!nama_counter), !nama_counter, "")
        
        IsiGridDetail Trim(TBukti.Text)
        
'        GridBrg.Columns(5).FooterText = IIf(Not IsNull(!total), !total, 0)
        
    End With
    
End Sub

Private Sub IsiGridDetail(ByVal bukti As String)
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from Tb_BrgMasuk_Detail where nobukti='" & bukti & "'"
    
    Dim a As Long
    Dim kode, nama, satuan As String
    Dim qty, harga, jml As Double
    Dim ida, ido As Integer
    Dim kondisi As String
    
    a = 1
    ArrBrg.ReDim 0, 0, 0, 0
    ArrBrg.ReDim 1, 1, 1, GridBrg.Columns.Count
        GridBrg.ReBind
        GridBrg.Refresh
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        If Not .EOF Then
            
            Do While Not .EOF
                ArrBrg.ReDim 1, a, 0, GridBrg.Columns.Count
                    GridBrg.ReBind
                    GridBrg.Refresh
                
                kode = IIf(Not IsNull(!kodebrg), !kodebrg, "")
                nama = IIf(Not IsNull(!namabrg), !namabrg, "")
                satuan = IIf(Not IsNull(!satuan), !satuan, "")
                qty = IIf(Not IsNull(!jml), !jml, 0)
                harga = 0 'If(Not IsNull(!harga), !harga, 0)
                jml = 0 '0IIf(Not IsNull(!jumlah), !jumlah, 0)
                ida = !id
                ido = !idorder
                kondisi = IIf(Not IsNull(!kondisi), !kondisi, "")
                
                ArrBrg(a, 0) = kode
                ArrBrg(a, 1) = nama
                ArrBrg(a, 2) = qty
                ArrBrg(a, 3) = satuan
                ArrBrg(a, 4) = harga
                ArrBrg(a, 5) = jml
                ArrBrg(a, 6) = ida
                ArrBrg(a, 7) = ido
                
                If Trim(kondisi) = "Baik" Then
                    ArrBrg(a, 8) = "Baik"
                ElseIf Trim(kondisi) = "Rusak" Then
                    ArrBrg(a, 8) = "Rusak"
                ElseIf Trim(kondisi) = "Lengkap" Then
                    ArrBrg(a, 8) = "Lengkap"
                ElseIf Trim(kondisi) = "Kurang" Then
                    ArrBrg(a, 8) = "Kurang"
                End If
                
                GridBrg.MoveLast
                DoEvents
                
            a = a + 1
            .MoveNext
            Loop
            
            GridBrg.ReBind
            GridBrg.Refresh
            
            GridBrg.MoveLast
            
        End If
    End With
    
End Sub
    

Private Sub Cmd_Batal_Click()

    Frame_Nav.Enabled = True
    rubah = False
             
        Cmd_Tambah.Visible = True
        Cmd_Tambah.Enabled = True
        Cmd_Simpan.Visible = False
        Cmd_Rubah.Visible = True
        Cmd_Rubah.Enabled = True
        Cmd_Hapus.Visible = True
        Cmd_Hapus.Enabled = True
        Cmd_Daftar.Visible = True
        Cmd_Daftar.Enabled = True
        Cmd_Keluar.Enabled = True
        
        CmdAdd.Visible = False
        CmdDel.Visible = False
        
        CmdAdd.Enabled = True
        CmdDel.Enabled = True
      
    TBukti.Enabled = False
    DTgl.Enabled = False
    TKodeSupp.Enabled = False
    CmdSupp.Enabled = False
    TKodeCbang.Enabled = False
    CmdCbang.Enabled = False
    
    TDBSupp.Visible = False
    TdbAdd.Visible = False
    TDB_Daftar.Visible = False
    
    Cmd_Batal.Visible = False
    Cmd_Batal.Enabled = True
    
    Cmd_Tambah.SetFocus
    
    txt_cr_daftar_KeyUp 0, 0, 0

    Cmd_Navigasi_Click 0
   
End Sub

Private Sub Cmd_Daftar_Click()

Frame_Nav.Enabled = False
With TDB_Daftar

If .Visible = False Then
    
    .Left = Me.Width / 2 - .Width / 2
    .Top = Me.Height / 2 - .Height / 2
    
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
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

Private Sub Cmd_Hapus_Click()
On Error GoTo err_handler

    If TBukti.Text = "" Then Exit Sub
    
    If MsgBox("Yakin akan menghapus data ini ?", vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then Exit Sub
    
    kon.BeginTrans
    
    Dim a As Long
        For a = 1 To ArrBrg.UpperBound(1)
            
            Dim comd1 As Command
            Set comd1 = New ADODB.Command
        With comd1
            .ActiveConnection = kon
            .CommandText = "kurangi_stock"
            .CommandType = adCmdStoredProc
            
            If ArrBrg(a, 8) = "Baik" Then
                .Parameters("@jml_stock").Value = ArrBrg(a, 2)
                .Parameters("@jml_rusak").Value = 0
                .Parameters("@jml_lengkap").Value = 0
                .Parameters("@jml_kurang").Value = 0
            ElseIf ArrBrg(a, 8) = "Rusak" Then
                .Parameters("@jml_stock").Value = 0
                .Parameters("@jml_rusak").Value = ArrBrg(a, 2)
                .Parameters("@jml_lengkap").Value = 0
                .Parameters("@jml_kurang").Value = 0
            ElseIf ArrBrg(a, 8) = "Lengkap" Then
                .Parameters("@jml_stock").Value = 0
                .Parameters("@jml_rusak").Value = 0
                .Parameters("@jml_lengkap").Value = ArrBrg(a, 2)
                .Parameters("@jml_kurang").Value = 0
            ElseIf ArrBrg(a, 8) = "Kurang" Then
                .Parameters("@jml_stock").Value = 0
                .Parameters("@jml_rusak").Value = 0
                .Parameters("@jml_lengkap").Value = 0
                .Parameters("@jml_kurang").Value = ArrBrg(a, 2)
            End If
            
            .Parameters("@kode_brg").Value = ArrBrg(a, 0)
            .Execute
        End With

            
        Next
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "delete from Tb_BrgMasuk where nobukti='" & Trim(TBukti.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    Dim sql1 As String
    Dim rs1 As Recordset
    
    sql1 = "delete from Tb_BrgMasuk_Detail where nobukti='" & Trim(TBukti.Text) & "'"
        Set rs1 = New ADODB.Recordset
            rs1.Open sql1, kon
    
    kon.CommitTrans
    Cmd_Batal_Click
    On Error GoTo 0
    Exit Sub
    
err_handler:
        
        kon.RollbackTrans
        
        MsgBox Error$
    
End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Navigasi_Click(Index As Integer)
On Error Resume Next

With Rs_Nav
Select Case Index
    Case 0
        .MoveFirst
    Case 1
        
        If .BOF Then .MoveFirst
        
        .MovePrevious
        
        If .BOF Then .MoveFirst
        
    Case 2
        
        If .EOF Then .MoveLast
        
        .MoveNext
        
        If .EOF Then .MoveLast
        
    Case 3
        
        .MoveLast
        
End Select
End With

isi_semua Rs_Nav


End Sub

Private Sub Cmd_Rubah_Click()
    
    If TBukti.Text <> "" Then
        
        DTgl.Enabled = True
        
        TotalOld = GridBrg.Columns(5).FooterText
        
        Cmd_Tambah.Visible = False
        Cmd_Simpan.Visible = True
        Cmd_Rubah.Visible = False
        Cmd_Batal.Visible = True
        Cmd_Hapus.Visible = False
        CmdAdd.Visible = True
        Cmd_Daftar.Visible = False
        CmdDel.Visible = True
        
        rubah = True
        DTgl.SetFocus
        
    End If
    
End Sub

Private Sub cmd_simpan_Click()
On Error GoTo err_handler

    If TBukti.Text = "" Or _
        TKodeSupp.Text = "" Then Exit Sub
    If ArrBrg.UpperBound(1) = 1 And ArrBrg(1, 1) = Empty Then Exit Sub
    
    kon.BeginTrans
    
    If rubah = False Then
        
        Dim sql As String
        Dim rs As Recordset
        
        sql = "select nobukti from Tb_BrgMasuk where nobukti='" & Trim(TBukti.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
        With rs
            If Not .EOF Then
                MsgBox "No bukti sudah ada"
                    kon.RollbackTrans
                    TBukti.SetFocus
                    Exit Sub
            End If
        End With
        
        Set rs = Nothing
        
        EvSimpan
        
    Else
        
        EvUpdate
        
    End If
    
    kon.CommitTrans
    
    Cmd_Batal_Click
    
    On Error GoTo 0
    Exit Sub
    
err_handler:
        
    kon.RollbackTrans
    MsgBox Error$
    
End Sub

Private Sub EvUpdate()
    
    Dim sql As String
    Dim rs As Recordset
        
    sql = "update Tb_BrgMasuk set tgl='" & Format(Trim(DTgl.Value), "yyyy/mm/dd") & "'"
    sql = sql & " where NoBukti='" & Trim(TBukti.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    Dim a As Long
    For a = 1 To ArrBrg.UpperBound(1)
    
    If ArrBrg(a, 6) = "" Then
        
        Dim sql1 As String
        Dim rs1 As Recordset
            sql1 = "insert into Tb_BrgMasuk_Detail (IdOrder,NoBukti,KodeBrg,NamaBrg,Jml,Satuan,Kondisi)"
            sql1 = sql1 & " values(" & ArrBrg(a, 7) & ",'" & Trim(TBukti.Text) & "','" & ArrBrg(a, 0) & "','" & ArrBrg(a, 1) & "'," & ArrBrg(a, 2) & ",'" & ArrBrg(a, 3) & "','" & ArrBrg(a, 8) & "' )"
            
            
        Set rs1 = New ADODB.Recordset
            rs1.Open sql1, kon
    
    Dim comd1 As Command
    Set comd1 = New ADODB.Command
        With comd1
            .ActiveConnection = kon
            .CommandText = "tambah_stock_update"
            .CommandType = adCmdStoredProc
            
            If ArrBrg(a, 8) = "Baik" Then
                .Parameters("@jml_stock").Value = ArrBrg(a, 2)
                .Parameters("@jml_rusak").Value = 0
                .Parameters("@jml_lengkap").Value = 0
                .Parameters("@jml_kurang").Value = 0
            ElseIf ArrBrg(a, 8) = "Rusak" Then
                .Parameters("@jml_stock").Value = 0
                .Parameters("@jml_rusak").Value = ArrBrg(a, 2)
                .Parameters("@jml_lengkap").Value = 0
                .Parameters("@jml_kurang").Value = 0
            ElseIf ArrBrg(a, 8) = "Lengkap" Then
                .Parameters("@jml_stock").Value = 0
                .Parameters("@jml_rusak").Value = 0
                .Parameters("@jml_lengkap").Value = ArrBrg(a, 2)
                .Parameters("@jml_kurang").Value = 0
            ElseIf ArrBrg(a, 8) = "Kurang" Then
                .Parameters("@jml_stock").Value = 0
                .Parameters("@jml_rusak").Value = 0
                .Parameters("@jml_lengkap").Value = 0
                .Parameters("@jml_kurang").Value = ArrBrg(a, 2)
            End If
            
            .Parameters("@kode_brg").Value = ArrBrg(a, 0)
            .Execute
        End With
    
    End If
    
    Next
    
    
End Sub

Private Sub EvSimpan()
    
    Dim sql As String
    Dim rs As Recordset
        
        sql = "insert into Tb_BrgMasuk (nobukti,tgl,bukti_order)"
        sql = sql & " values('" & Trim(TBukti.Text) & "','" & Format(Trim(DTgl.Value), "yyyy/mm/dd") & "','" & Trim(TKodeSupp.Text) & "')"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
        
        Dim a As Long
        For a = 1 To ArrBrg.UpperBound(1)
            
          If ArrBrg(a, 6) = "" Then
            
                    Dim sql1 As String
            Dim rs1 As Recordset
                    sql1 = "insert into Tb_BrgMasuk_Detail (IdOrder,NoBukti,KodeBrg,NamaBrg,Jml,Satuan,Kondisi)"
            sql1 = sql1 & " values(" & ArrBrg(a, 7) & ",'" & Trim(TBukti.Text) & "','" & ArrBrg(a, 0) & "','" & ArrBrg(a, 1) & "'," & ArrBrg(a, 2) & ",'" & ArrBrg(a, 3) & "','" & ArrBrg(a, 8) & "')"
            
            
            Set rs1 = New ADODB.Recordset
            rs1.Open sql1, kon

            Dim comd1 As Command
            Set comd1 = New ADODB.Command
                With comd1
                    .ActiveConnection = kon
                    .CommandText = "tambah_stock_update"
                    .CommandType = adCmdStoredProc
                    
                    If ArrBrg(a, 8) = "Baik" Then
                        .Parameters("@jml_stock").Value = ArrBrg(a, 2)
                        .Parameters("@jml_rusak").Value = 0
                        .Parameters("@jml_lengkap").Value = 0
                        .Parameters("@jml_kurang").Value = 0
                    ElseIf ArrBrg(a, 8) = "Rusak" Then
                        .Parameters("@jml_stock").Value = 0
                        .Parameters("@jml_rusak").Value = ArrBrg(a, 2)
                        .Parameters("@jml_lengkap").Value = 0
                        .Parameters("@jml_kurang").Value = 0
                    ElseIf ArrBrg(a, 8) = "Lengkap" Then
                        .Parameters("@jml_stock").Value = 0
                        .Parameters("@jml_rusak").Value = 0
                        .Parameters("@jml_lengkap").Value = ArrBrg(a, 2)
                        .Parameters("@jml_kurang").Value = 0
                    ElseIf ArrBrg(a, 8) = "Kurang" Then
                        .Parameters("@jml_stock").Value = 0
                        .Parameters("@jml_rusak").Value = 0
                        .Parameters("@jml_lengkap").Value = 0
                        .Parameters("@jml_kurang").Value = ArrBrg(a, 2)
                    End If
                    
                    .Parameters("@kode_brg").Value = ArrBrg(a, 0)
                    .Execute
                End With
            End If
            
        Next
    
End Sub

Private Sub Cmd_Tambah_Click()

    rubah = False
    
    Frame_Nav.Enabled = False
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
     Cmd_Rubah.Visible = False
     Cmd_Batal.Visible = True
     Cmd_Hapus.Visible = False
     Cmd_Daftar.Visible = False
     Cmd_Keluar.Enabled = False
     
     CmdAdd.Visible = True
     CmdDel.Visible = True
         
    ArrBrg.ReDim 0, 0, 0, 0
    ArrBrg.ReDim 1, 1, 1, GridBrg.Columns.Count
        GridBrg.ReBind
        GridBrg.Refresh
        
'        GridBrg.Columns(5).FooterText = 0
        
     TBukti.Text = ""
     TBukti.Enabled = True
     TBukti.SetFocus

End Sub

Private Sub cmdadd_Click()
    
    If TBukti.Text = "" _
        Or TKodeSupp.Text = "" _
            Then Exit Sub
    
    With TdbAdd
        
        If .Visible = False Then
            
            .Left = Me.Width / 2 - .Width / 2
            .Top = Me.Height / 2 - .Height / 2
            
            TNamaBrg.Text = Trim(TKodeSupp.Text)
            TSatuan.Text = Trim(TNamaSupp.Text)
            
            .Visible = True
            
            isi_barangorder
            
            Gridadd.SetFocus
            
        Else
            .Visible = False
        End If
        
    End With
    
End Sub

Private Sub isi_barangorder()
    
    arradd.ReDim 0, 0, 0, 0
    arradd.ReDim 1, 1, 1, 1
        Gridadd.ReBind
        Gridadd.Refresh
    
    Dim sql As String
    Dim rs As Recordset
        
        sql = "select id,kodebrg,namabrg,jml,satuan from Tb_Order_Detail where nobukti='" & Trim(TNamaBrg.Text) & "'"
        sql = sql & " and id not in (select idorder from Tb_BrgMasuk_Detail)"
        
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Dim a As Long
    Dim kode, nama, id, jml, satuan As String
        a = 1
        
    With rs
        
        Do While Not .EOF
            arradd.ReDim 1, a, 0, Gridadd.Columns.Count
                Gridadd.ReBind
                Gridadd.Refresh
            
            kode = IIf(Not IsNull(!kodebrg), !kodebrg, "")
            nama = IIf(Not IsNull(!namabrg), !namabrg, "")
            id = !id
            jml = IIf(Not IsNull(!jml), !jml, 0)
            satuan = IIf(Not IsNull(!satuan), !satuan, "")
            
            arradd(a, 0) = kode
            arradd(a, 1) = nama
            arradd(a, 2) = satuan
            arradd(a, 3) = jml
            arradd(a, 4) = 0
            arradd(a, 5) = id
            arradd(a, 6) = "Baik"
            
        a = a + 1
        .MoveNext
        Loop
        
        Gridadd.ReBind
        Gridadd.Refresh
        
    End With
    
End Sub

Private Sub cmddel_Click()
    
    If ArrBrg.UpperBound(1) = 1 And ArrBrg(1, 1) = Empty Then Exit Sub
    If ArrBrg.UpperBound(1) = 1 Then Exit Sub
    
    If ArrBrg(GridBrg.Bookmark, 6) <> "" Then
    
        Dim sql As String
        Dim rs As Recordset
        
        Dim comd1 As Command
        Set comd1 = New ADODB.Command
        With comd1
            .ActiveConnection = kon
            .CommandText = "kurangi_stock"
            .CommandType = adCmdStoredProc
            
            If ArrBrg(GridBrg.Bookmark, 8) = "Baik" Then
                .Parameters("@jml_stock").Value = ArrBrg(GridBrg.Bookmark, 2)
                .Parameters("@jml_rusak").Value = 0
                .Parameters("@jml_lengkap").Value = 0
                .Parameters("@jml_kurang").Value = 0
            ElseIf ArrBrg(GridBrg.Bookmark, 8) = "Rusak" Then
                .Parameters("@jml_stock").Value = 0
                .Parameters("@jml_rusak").Value = ArrBrg(GridBrg.Bookmark, 2)
                .Parameters("@jml_lengkap").Value = 0
                .Parameters("@jml_kurang").Value = 0
            ElseIf ArrBrg(GridBrg.Bookmark, 8) = "Lengkap" Then
                .Parameters("@jml_stock").Value = 0
                .Parameters("@jml_rusak").Value = 0
                .Parameters("@jml_lengkap").Value = ArrBrg(GridBrg.Bookmark, 2)
                .Parameters("@jml_kurang").Value = 0
            ElseIf ArrBrg(GridBrg.Bookmark, 8) = "Kurang" Then
                .Parameters("@jml_stock").Value = 0
                .Parameters("@jml_rusak").Value = 0
                .Parameters("@jml_lengkap").Value = 0
                .Parameters("@jml_kurang").Value = ArrBrg(GridBrg.Bookmark, 2)
            End If
            
            .Parameters("@kode_brg").Value = ArrBrg(GridBrg.Bookmark, 0)
'            .Parameters("@kode_cbang").Value = Trim(TKodeCbang.Text)
            .Execute
        End With

        sql = "delete from Tb_BrgMasuk_Detail where id=" & ArrBrg(GridBrg.Bookmark, 6)
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
    End If
    
'    GridBrg.Columns(5).FooterText = CDbl(GridBrg.Columns(5).FooterText) - CDbl(ArrBrg(GridBrg.Bookmark, 5))
    
    If ArrBrg.UpperBound(1) = 1 Then
        
        ArrBrg.ReDim 0, 0, 0, 0
        ArrBrg.ReDim 1, 1, 1, GridBrg.Columns.Count
    
    Else
        GridBrg.Delete
    End If
    
            GridBrg.ReBind
            GridBrg.Refresh
    
     Cmd_Batal.Enabled = False
    
End Sub

Private Sub CmdFinish_Click()
    TdbAdd.Visible = False
End Sub

Private Sub CmdOk_Click()
    
    
    If arradd.UpperBound(1) = 1 And arradd(1, 1) = Empty Then Exit Sub
    
    Gridadd.MoveLast
    Gridadd.MoveFirst
    
    Dim x As Long
        For x = 1 To arradd.UpperBound(1)
        If arradd(x, 4) <> 0 Then
        If IsNumeric(arradd(x, 4)) = True Then
            If PeriksaBrgAdd(arradd(x, 0)) = True Then
                MsgBox "Barang yang akan ditambahkan sudah ada"
               Gridadd.Bookmark = x
               Gridadd.SetFocus
            Else
                
                Dim a As Long
                If ArrBrg(1, 1) = Empty And ArrBrg.UpperBound(1) = 1 Then
                    a = 1
                Else
                    a = ArrBrg.UpperBound(1) + 1
                End If
                
                ArrBrg.ReDim 1, a, 0, GridBrg.Columns.Count
                    GridBrg.ReBind
                    GridBrg.Refresh
                
                ArrBrg(a, 0) = arradd(x, 0)
                ArrBrg(a, 1) = arradd(x, 1)
                ArrBrg(a, 2) = arradd(x, 4)
                ArrBrg(a, 3) = arradd(x, 2)
                ArrBrg(a, 4) = ""
                ArrBrg(a, 5) = ""
                ArrBrg(a, 6) = ""
                ArrBrg(a, 7) = arradd(x, 5)
                ArrBrg(a, 8) = arradd(x, 6)
                
                GridBrg.ReBind
                GridBrg.Refresh
                
                GridBrg.MoveLast
                
        '        GridBrg.Columns(5).FooterText = CDbl(GridBrg.Columns(5).FooterText) + totJml
                End If
            Else
                
                MsgBox "Jml masuk harus number"
                Gridadd.Bookmark = x
                Gridadd.Col = 4
                Gridadd.SetFocus
                
            
            End If
            End If
        
        Next
        
End Sub

Private Function PeriksaBrgAdd(ByVal kode As String) As Boolean
    
    If ArrBrg.UpperBound(1) = 1 And ArrBrg(1, 1) = Empty Then
        PeriksaBrgAdd = False
        Exit Function
    End If
    
    Dim a As Long
    Dim hasil As Boolean
        hasil = False
    
    For a = 1 To ArrBrg.UpperBound(1)
        If ArrBrg(a, 0) = kode Then
            hasil = True
            Exit For
        End If
    Next
    
    PeriksaBrgAdd = hasil
    
End Function

Private Sub CmdSupp_Click()
    
    With TDBSupp
        If .Visible = False Then
            
            .Left = Picture1.Left + CmdSupp.Left + CmdSupp.Width / 2 - .Width / 2
            .Top = Picture1.Top + CmdSupp.Top + CmdSupp.Height + 15
            
            Txt_Cr_Supp(0).Text = ""
            Txt_Cr_Supp(1).Text = ""
            
            txt_cr_supp_KeyUp 0, 0, 0
            
            .Visible = True
            
            Txt_Cr_Supp(0).SetFocus
            
        Else
            .Visible = False
        End If
    End With
    
End Sub

Private Sub DTgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If TKodeSupp.Enabled = True Then
            TKodeSupp.SetFocus
        Else
            CmdAdd.SetFocus
        End If
    End If
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

rubah = False

With Me
    .Left = Screen.Width / 2 - .Width / 2
    .Top = 250
End With

TBukti.Enabled = False
DTgl.Enabled = False
TKodeSupp.Enabled = False
CmdSupp.Enabled = False
TKodeCbang.Enabled = False
CmdCbang.Enabled = False

GridBrg.Array = ArrBrg
    
    ArrBrg.ReDim 0, 0, 0, 0
    ArrBrg.ReDim 1, 1, 1, GridBrg.Columns.Count
        GridBrg.ReBind
        GridBrg.Refresh

Gridadd.Array = arradd
    
    arradd.ReDim 0, 0, 0, 0
    arradd.ReDim 1, 1, 1, Gridadd.Columns.Count
        Gridadd.ReBind
        Gridadd.Refresh

txt_cr_daftar_KeyUp 0, 0, 0
Cmd_Navigasi_Click 0

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
    
    Dim nobuk As String
        nobuk = Grid_Daftar.Columns(0).Text
    
    Rs_Nav.MoveFirst
    
    Rs_Nav.Find "nobukti='" & nobuk & "'"

    isi_semua Rs_Nav
    
    TDB_Daftar.Visible = False
    Frame_Nav.Enabled = True
    Cmd_Navigasi(0).SetFocus
    Cmd_Rubah.Enabled = True
    Cmd_Rubah.Visible = True
    Cmd_Batal.Visible = False
    Cmd_Batal.Enabled = True
    Cmd_Tambah.Enabled = True
    Cmd_Hapus.Enabled = True
    Cmd_Daftar.Enabled = True
    Cmd_Keluar.Enabled = True
    
End Sub

Private Sub grid_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_daftar_DblClick
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub



Private Sub Gridadd_AfterColUpdate(ByVal ColIndex As Integer)
    
    If ColIndex = 4 Or ColIndex = 6 Then
        
        arradd(Gridadd.Bookmark, ColIndex) = Gridadd.Columns(ColIndex).Text
        
        DoEvents
        
    End If
    
End Sub

Private Sub GridSupp_DblClick()
    
    If GridSupp.Row < 0 Then Exit Sub
    
        TKodeSupp.Text = GridSupp.Columns(0).Text
        TNamaSupp.Text = GridSupp.Columns(1).Text
        txtnama.Text = GridSupp.Columns(2).Text
    
    TDBSupp.Visible = False
    CmdAdd.SetFocus
    
End Sub

Private Sub GridSupp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then GridSupp_DblClick
    If KeyCode = vbKeyEscape Then CmdSupp_Click
End Sub

Private Sub TBukti_GotFocus()
    Call Focus_(TBukti)
End Sub


Private Sub TBukti_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TBukti_LostFocus
    
End Sub

Private Sub TBukti_LostFocus()
    
    If TBukti.Text = "" Then Exit Sub
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select nobukti from Tb_BrgMasuk where nobukti='" & Trim(TBukti.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        If Not .EOF Then
            MsgBox "No bukti sudah ada"
            TBukti.SetFocus
            
            DTgl.Enabled = False
            TKodeSupp.Enabled = False
            CmdSupp.Enabled = False
            TKodeCbang.Enabled = False
            CmdCbang.Enabled = False
            Cmd_Simpan.Enabled = False
            
        Else
            
            DTgl.Enabled = True
            TKodeSupp.Enabled = True
            CmdSupp.Enabled = True
            TKodeCbang.Enabled = True
            CmdCbang.Enabled = True
            Cmd_Simpan.Enabled = True
            
            TKodeSupp.Text = ""
            TNamaSupp.Text = ""
            TKodeCbang.Text = ""
            TNamaCbang.Text = ""
            txtnama.Text = ""
            
            DTgl.SetFocus
            
        End If
    End With
    
End Sub


Private Sub TDB_Brg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = True
If Moving = True Then
   yold = y
   xold = x
End If
End Sub

Private Sub TDB_Brg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Moving = True Then
   TDB_Brg.Top = TDB_Brg.Top - (yold - y)
   TDB_Brg.Left = TDB_Brg.Left - (xold - x)
End If

End Sub

Private Sub TDB_Brg_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = False
End Sub

Private Sub TDB_Cbang_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = True
If Moving = True Then
   yold = y
   xold = x
End If
End Sub

Private Sub TDB_Cbang_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Moving = True Then
   TDB_Cbang.Top = TDB_Cbang.Top - (yold - y)
   TDB_Cbang.Left = TDB_Cbang.Left - (xold - x)
End If

End Sub

Private Sub TDB_Cbang_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = False
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

Private Sub TDBSupp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = True
If Moving = True Then
   yold = y
   xold = x
End If
End Sub

Private Sub TDBSupp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Moving = True Then
   TDBSupp.Top = TDBSupp.Top - (yold - y)
   TDBSupp.Left = TDBSupp.Left - (xold - x)
End If

End Sub

Private Sub TDBSupp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = False
End Sub

Private Sub TKodeCbang_LostFocus()
    
    If TKodeCbang.Text = "" Then Exit Sub
    
    Dim sql As String
    Dim rs As Recordset
        
        sql = "select kode_counter,nama_counter from view_counter_user where id_user=" & Flag_tempat
        sql = sql & " and kode_counter='" & Trim(TKodeCbang.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If Not .EOF Then
                
                TNamaCbang.Text = IIf(Not IsNull(!nama_counter), !nama_counter, "")
                
            Else
                MsgBox "Data cabang tidak ditemukan"
                TNamaCbang.Text = ""
                TKodeCbang.SetFocus
            End If
        End With
        
        
        
    
End Sub

Private Sub TKodeSupp_GotFocus()
    Call Focus_(TKodeSupp)
End Sub


Private Sub TKodeSupp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then CmdSupp_Click
    If KeyCode = 13 Then CmdAdd.SetFocus
End Sub

Private Sub TKodeSupp_LostFocus()
    
    If TKodeSupp.Text = "" Then Exit Sub
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "select nobukti,tgl,atas_nama from tb_order where nobukti='" & Trim(TKodeSupp.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        If Not .EOF Then
            
            TNamaSupp.Text = IIf(Not IsNull(!tgl), !tgl, "")
            txtnama.Text = IIf(Not IsNull(!atas_nama), !atas_nama, "")
        
        Else
            
            MsgBox "No bukti order barang tidak ditemukan"
            TNamaSupp.Text = ""
            txtnama.Text = ""
            TKodeSupp.SetFocus
            
        End If
    End With
    
End Sub


Private Sub txt_cr_daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub txt_cr_daftar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
           
    Dim sql As String
        sql = "select top 100 * from VIEW_BrgMasuk " ' where kodecounter in (select kode_counter from VIEW_Counter_User where id_user=" & Flag_tempat & ")"
        
    If Txt_Cr_Daftar(0).Text <> "" Or Txt_Cr_Daftar(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where nobukti like '%" & Trim(Txt_Cr_Daftar(0).Text) & "%'"
            Case 1
            
            If Len(Txt_Cr_Daftar(1).Text) = 10 Then
                sql = sql & " where tgl='" & Format(Txt_Cr_Daftar(1).Text, "yyyy/mm/dd") & "'"
            End If
            
        End Select
    End If
    
    sql = sql & " order by tgl desc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Set Grid_Daftar.DataSource = Rs_Nav
        Grid_Daftar.Refresh

End Sub

Private Sub txt_cr_supp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then GridSupp.SetFocus
    If KeyCode = vbKeyEscape Then CmdSupp_Click
End Sub

Private Sub txt_cr_supp_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sql As String
Dim rs As Recordset

    sql = "select * from Tb_order"
        
    If Txt_Cr_Supp(0).Text <> "" Or Txt_Cr_Supp(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where nobukti like  '%" & Trim(Txt_Cr_Supp(0).Text) & "%'"
            Case 1
            If Len(Txt_Cr_Supp(Index).Text) = 10 Then
                sql = sql & " where tgl ='" & Format(Trim(Txt_Cr_Supp(Index).Text), "yyyy/mm/dd") & "'"
            End If
        End Select
    End If
    
    sql = sql & " order by nobukti,tgl desc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set GridSupp.DataSource = rs
        GridSupp.Refresh
    
End Sub

