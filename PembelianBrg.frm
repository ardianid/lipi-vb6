VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form PembelianBrg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "O R D E R  B A R A N G"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PembelianBrg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   3975
      Left            =   1320
      TabIndex        =   67
      Top             =   2400
      Visible         =   0   'False
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   7011
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "PembelianBrg.frx":27C92
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "PembelianBrg.frx":27CAE
      Childs          =   "PembelianBrg.frx":27D5A
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   70
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   69
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame Frame6 
         Height          =   135
         Index           =   2
         Left            =   240
         TabIndex        =   68
         Top             =   360
         Width           =   6495
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   2895
         Left            =   240
         OleObjectBlob   =   "PembelianBrg.frx":27D76
         TabIndex        =   71
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
         Top             =   120
         Width           =   945
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Brg 
      Height          =   3615
      Left            =   -4560
      TabIndex        =   49
      Top             =   2040
      Visible         =   0   'False
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "PembelianBrg.frx":2B93E
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "PembelianBrg.frx":2B95A
      Childs          =   "PembelianBrg.frx":2BA06
      Begin VB.TextBox TxtCr_Brg 
         Height          =   300
         Index           =   1
         Left            =   2880
         TabIndex        =   56
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtCr_Brg 
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   55
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
         TabIndex        =   50
         Top             =   360
         Width           =   4815
      End
      Begin TrueOleDBGrid60.TDBGrid GridCr_Brg 
         Height          =   2415
         Left            =   240
         OleObjectBlob   =   "PembelianBrg.frx":2BA22
         TabIndex        =   51
         Top             =   960
         Width           =   4815
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   600
         Width           =   405
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TdbAdd 
      Height          =   2535
      Left            =   -6600
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   4471
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "PembelianBrg.frx":2E9A1
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "PembelianBrg.frx":2E9BD
      Childs          =   "PembelianBrg.frx":2EA69
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
         TabIndex        =   48
         Top             =   1920
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
         TabIndex        =   47
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TSatuan 
         Height          =   300
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   960
         Width           =   4695
      End
      Begin VB.CommandButton CmdBrg 
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
         Left            =   5880
         TabIndex        =   38
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox TNamaBrg 
         Height          =   300
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox TKodeBrg 
         Height          =   300
         Left            =   1080
         TabIndex        =   36
         Top             =   600
         Width           =   1095
      End
      Begin TDBNumber6Ctl.TDBNumber TDB_Qty 
         Height          =   285
         Left            =   1080
         TabIndex        =   44
         Top             =   1320
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   503
         Calculator      =   "PembelianBrg.frx":2EA85
         Caption         =   "PembelianBrg.frx":2EAA5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PembelianBrg.frx":2EB11
         Keys            =   "PembelianBrg.frx":2EB2F
         Spin            =   "PembelianBrg.frx":2EB79
         AlignHorizontal =   1
         AlignVertical   =   0
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
      Begin TDBNumber6Ctl.TDBNumber TDB_Hrg 
         Height          =   285
         Left            =   1080
         TabIndex        =   45
         Top             =   1680
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   503
         Calculator      =   "PembelianBrg.frx":2EBA1
         Caption         =   "PembelianBrg.frx":2EBC1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PembelianBrg.frx":2EC2D
         Keys            =   "PembelianBrg.frx":2EC4B
         Spin            =   "PembelianBrg.frx":2EC95
         AlignHorizontal =   1
         AlignVertical   =   0
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
         ValueVT         =   2089877505
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber TDB_Jml 
         Height          =   285
         Left            =   1080
         TabIndex        =   46
         Top             =   2040
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   503
         Calculator      =   "PembelianBrg.frx":2ECBD
         Caption         =   "PembelianBrg.frx":2ECDD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PembelianBrg.frx":2ED49
         Keys            =   "PembelianBrg.frx":2ED67
         Spin            =   "PembelianBrg.frx":2EDB1
         AlignHorizontal =   1
         AlignVertical   =   0
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   2089877505
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah :"
         Height          =   195
         Index           =   4
         Left            =   375
         TabIndex        =   42
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga :"
         Height          =   195
         Index           =   3
         Left            =   435
         TabIndex        =   41
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty :"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   40
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan :"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barang :"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   600
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   6840
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P E N A M B A H A N B A R A N G"
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
         Width           =   2505
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Cbang 
      Height          =   3615
      Left            =   7320
      TabIndex        =   61
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "PembelianBrg.frx":2EDD9
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "PembelianBrg.frx":2EDF5
      Childs          =   "PembelianBrg.frx":2EEA1
      Begin VB.TextBox TxtCrCbang 
         Height          =   285
         Left            =   960
         TabIndex        =   66
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
         TabIndex        =   62
         Top             =   360
         Width           =   4455
      End
      Begin TrueOleDBGrid60.TDBGrid GridCbang 
         Height          =   2415
         Left            =   240
         OleObjectBlob   =   "PembelianBrg.frx":2EEBD
         TabIndex        =   63
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   65
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
         TabIndex        =   64
         Top             =   120
         Width           =   945
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDBSupp 
      Height          =   3615
      Left            =   -4680
      TabIndex        =   25
      Top             =   2400
      Visible         =   0   'False
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "PembelianBrg.frx":3195F
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "PembelianBrg.frx":3197B
      Childs          =   "PembelianBrg.frx":31A27
      Begin VB.TextBox Txt_Cr_Supp 
         Height          =   315
         Index           =   0
         Left            =   840
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Txt_Cr_Supp 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   27
         Top             =   600
         Width           =   2655
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
         OleObjectBlob   =   "PembelianBrg.frx":31A43
         TabIndex        =   29
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   14
         Left            =   2280
         TabIndex        =   32
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   360
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
      Left            =   240
      ScaleHeight     =   6495
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.TextBox txtnama 
         Height          =   300
         Left            =   1320
         MaxLength       =   75
         TabIndex        =   76
         Top             =   840
         Width           =   7335
      End
      Begin VB.TextBox txt_ket 
         Height          =   300
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   75
         Top             =   1200
         Width           =   7335
      End
      Begin TrueOleDBGrid60.TDBGrid GridBrg 
         Height          =   3855
         Left            =   120
         OleObjectBlob   =   "PembelianBrg.frx":349CC
         TabIndex        =   9
         Top             =   1680
         Width           =   9135
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
         Left            =   13320
         TabIndex        =   60
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TNamaCbang 
         Height          =   300
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   3960
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox TKodeCbang 
         Height          =   300
         Left            =   8520
         TabIndex        =   58
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
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
         Left            =   4800
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
         Left            =   13320
         TabIndex        =   8
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TNamaSupp 
         Height          =   300
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3600
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox TKodeSupp 
         Height          =   300
         Left            =   8520
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
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
         Format          =   3866625
         CurrentDate     =   39467
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pemohon :"
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   77
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang :"
         Height          =   195
         Index           =   6
         Left            =   7695
         TabIndex        =   57
         Top             =   3960
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ket :"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   1200
         Width           =   345
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
   End
End
Attribute VB_Name = "PembelianBrg"
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

Private Sub isi_semua(ByVal rec As Recordset)
    With rec
        
        If .EOF Then .MoveLast
        If .BOF Then .MoveFirst
        
        TBukti.Text = IIf(Not IsNull(!nobukti), !nobukti, "")
        DTgl.Value = !tgl
        txtnama.Text = IIf(Not IsNull(!atas_nama), !atas_nama, "")
        txt_ket.Text = IIf(Not IsNull(!ket), !ket, "")
'        TKodeSupp.Text = IIf(Not IsNull(!KodeSupplier), !KodeSupplier, "")
'        TNamaSupp.Text = IIf(Not IsNull(!Nama_Toko), !Nama_Toko, "")
'        TKodeCbang.Text = IIf(Not IsNull(!kodecounter), !kodecounter, "")
'        TNamaCbang.Text = IIf(Not IsNull(!nama_counter), !nama_counter, "")
        
        IsiGridDetail Trim(TBukti.Text)
        
        GridBrg.Columns(5).FooterText = IIf(Not IsNull(!total), !total, 0)
        
    End With
    
End Sub

Private Sub IsiGridDetail(ByVal bukti As String)
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from Tb_Order_Detail where nobukti='" & bukti & "'"
    
    Dim a As Long
    Dim kode, nama, satuan As String
    Dim qty, harga, jml As Double
    Dim ida As Integer
    
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
                harga = IIf(Not IsNull(!harga), !harga, 0)
                jml = IIf(Not IsNull(!jumlah), !jumlah, 0)
                ida = !id
                
                ArrBrg(a, 0) = kode
                ArrBrg(a, 1) = nama
                ArrBrg(a, 2) = qty
                ArrBrg(a, 3) = satuan
                ArrBrg(a, 4) = harga
                ArrBrg(a, 5) = jml
                ArrBrg(a, 6) = ida
                
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

Private Sub JmlAdd()
    
    Dim qty As Double
        If TDB_Qty.ValueIsNull Then
            qty = 0
        Else
            qty = Replace(Trim(TDB_Qty.Value), ",", "")
        End If
    
    Dim hrg As Double
        If TDB_Hrg.ValueIsNull Then
            hrg = 0
        Else
            hrg = Replace(Trim(TDB_Hrg.Value), ",", "")
        End If
    
    Dim totJml As Double
        totJml = qty * hrg
    
    If totJml = 0 Then
        TDB_Jml.Value = Null
    Else
        TDB_Jml.Value = totJml
    End If
    
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
    txt_ket.Enabled = False
    txtnama.Enabled = False
    
    TDBSupp.Visible = False
    TdbAdd.Visible = False
    TDB_Daftar.Visible = False
    
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
    
'    Dim a As Long
'        For a = 1 To ArrBrg.UpperBound(1)
'
'            Dim comd1 As Command
'            Set comd1 = New ADODB.Command
'        With comd1
'            .ActiveConnection = kon
'            .CommandText = "kurangi_stock"
'            .CommandType = adCmdStoredProc
'            .Parameters("@jml_stock").Value = ArrBrg(a, 2)
'            .Parameters("@kode_brg").Value = ArrBrg(a, 0)
'            .Parameters("@kode_cbang").Value = Trim(TKodeCbang.Text)
'            .Execute
'        End With
'
'
'        Next
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "delete from Tb_Order where nobukti='" & Trim(TBukti.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    Dim sql1 As String
    Dim rs1 As Recordset
    
    sql1 = "delete from Tb_Order_detail where nobukti='" & Trim(TBukti.Text) & "'"
        Set rs1 = New ADODB.Recordset
            rs1.Open sql1, kon
    
    
'    Dim sql2 As String
'        Dim rs2 As Recordset
'
'            sql2 = "update tb_supplier set hutang=hutang -" & CDbl(GridBrg.Columns(5).FooterText)
'            sql2 = sql2 & " where kode='" & Trim(TKodeSupp.Text) & "'"
'
'            Set rs2 = New ADODB.Recordset
'                rs2.Open sql2, kon
    
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
        txtnama.Enabled = True
        txt_ket.Enabled = True
        
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

    If TBukti.Text = "" Then Exit Sub
'        TKodeSupp.Text = "" Or _
'            TKodeCbang.Text = "" Then Exit Sub
    If ArrBrg.UpperBound(1) = 1 And ArrBrg(1, 1) = Empty Then Exit Sub
    
    kon.BeginTrans
    
    If rubah = False Then
        
        Dim sql As String
        Dim rs As Recordset
        
        sql = "select nobukti from Tb_Order where nobukti='" & Trim(TBukti.Text) & "'"
    
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
        
    sql = "update Tb_Order set tgl='" & Format(Trim(DTgl.Value), "yyyy/mm/dd") & "',atas_nama='" & Trim(txtnama.Text) & "',ket='" & Trim(txt_ket.Text) & "', total = " & GridBrg.Columns(5).FooterText
    sql = sql & " where NoBukti='" & Trim(TBukti.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
'    Dim sql1 As String
'        Dim rs1 As Recordset
'
'            sql1 = "update tb_supplier set hutang=hutang -" & TotalOld  'CDbl(GridBrg.Columns(5).FooterText)
'            sql1 = sql1 & " where kode='" & Trim(TKodeSupp.Text) & "'"
'
'            Set rs1 = New ADODB.Recordset
'                rs1.Open sql1, kon
    
'    Dim sql2 As String
'    Dim rs2 As Recordset
'
'        sql2 = "update tb_supplier set hutang=hutang +" & CDbl(GridBrg.Columns(5).FooterText)
'            sql2 = sql2 & " where kode='" & Trim(TKodeSupp.Text) & "'"
'
'            Set rs2 = New ADODB.Recordset
'                rs2.Open sql2, kon
        
    
    Dim a As Long
    For a = 1 To ArrBrg.UpperBound(1)
    
    If ArrBrg(a, 6) = "" Then
    Dim comd As Command
    Set comd = New ADODB.Command
        With comd
            .ActiveConnection = kon
            .CommandText = "tambah_beli_detail"
            .CommandType = adCmdStoredProc
            .Parameters("@bukti").Value = Trim(TBukti.Text)
            .Parameters("@kodebrg").Value = ArrBrg(a, 0)
            .Parameters("@jml").Value = ArrBrg(a, 2)
            .Parameters("@satuan").Value = ArrBrg(a, 3)
            .Parameters("@harga").Value = ArrBrg(a, 4)
            .Parameters("@jumlah").Value = ArrBrg(a, 5)
            .Parameters("@nama").Value = ArrBrg(a, 1)
            .Execute
        End With
    
'    Dim comd1 As Command
'    Set comd1 = New ADODB.Command
'        With comd1
'            .ActiveConnection = kon
'            .CommandText = "tambah_stock_update"
'            .CommandType = adCmdStoredProc
'            .Parameters("@jml_stock").Value = ArrBrg(a, 2)
'            .Parameters("@kode_brg").Value = ArrBrg(a, 0)
'            .Parameters("@kode_cbang").Value = Trim(TKodeCbang.Text)
'            .Execute
'        End With
    
    End If
    
    Next
    
    
End Sub

Private Sub EvSimpan()
    
    Dim sql As String
    Dim rs As Recordset
        
        sql = "insert into Tb_Order (nobukti,tgl,Ket,total,atas_nama)"
        sql = sql & " values('" & Trim(TBukti.Text) & "','" & Format(Trim(DTgl.Value), "yyyy/mm/dd") & "','" & Trim(txt_ket.Text) & "'," & GridBrg.Columns(5).FooterText & ",'" & Trim(txtnama.Text) & "' )"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
'        Dim sql1 As String
'        Dim rs1 As Recordset
'
'            sql1 = "update tb_supplier set hutang=hutang +" & CDbl(GridBrg.Columns(5).FooterText)
'            sql1 = sql1 & " where kode='" & Trim(TKodeSupp.Text) & "'"
'
'            Set rs1 = New ADODB.Recordset
'                rs1.Open sql1, kon
        
        Dim a As Long
        For a = 1 To ArrBrg.UpperBound(1)
            
          If ArrBrg(a, 6) = "" Then
            Dim comd As Command
            Set comd = New ADODB.Command
                With comd
                    .ActiveConnection = kon
                    .CommandText = "tambah_beli_detail"
                    .CommandType = adCmdStoredProc
                    .Parameters("@bukti").Value = Trim(TBukti.Text)
                    .Parameters("@kodebrg").Value = ArrBrg(a, 0)
                    .Parameters("@jml").Value = ArrBrg(a, 2)
                    .Parameters("@satuan").Value = ArrBrg(a, 3)
                    .Parameters("@harga").Value = ArrBrg(a, 4)
                    .Parameters("@jumlah").Value = ArrBrg(a, 5)
                    .Parameters("@nama").Value = ArrBrg(a, 1)
                    .Execute
                End With
            
'            Dim comd1 As Command
'            Set comd1 = New ADODB.Command
'                With comd1
'                    .ActiveConnection = kon
'                    .CommandText = "tambah_stock_update"
'                    .CommandType = adCmdStoredProc
'                    .Parameters("@jml_stock").Value = ArrBrg(a, 2)
'                    .Parameters("@kode_brg").Value = ArrBrg(a, 0)
'                    .Parameters("@kode_cbang").Value = Trim(TKodeCbang.Text)
'                    .Execute
'                End With
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
        
        GridBrg.Columns(5).FooterText = 0
        
     TBukti.Text = ""
     TBukti.Enabled = True
     TBukti.SetFocus

End Sub

Private Sub cmdadd_Click()
    
    If TBukti.Text = "" Then Exit Sub
'        Or TKodeSupp.Text = "" _
'           Or TKodeCbang.Text = "" Then Exit Sub
    
    With TdbAdd
        
        If .Visible = False Then
            
            .Left = Me.Width / 2 - .Width / 2
            .Top = Me.Height / 2 - .Height / 2
            
            TKodeBrg.Text = ""
            TNamaBrg.Text = ""
            TSatuan.Text = ""
            TDB_Qty.Value = Null
            TDB_Hrg.Value = Null
            TDB_Jml.Value = Null
            
            .Visible = True
            
            TKodeBrg.SetFocus
            
        Else
            .Visible = False
        End If
        
    End With
    
End Sub

Private Sub CmdBrg_Click()

    With TDB_Brg
        If .Visible = False Then
            
            .Left = (TdbAdd.Left + CmdBrg.Left + CmdBrg.Width / 2 - .Width / 2) - 350
            .Top = TdbAdd.Top + CmdBrg.Top + CmdBrg.Height + 15
            
            TxtCr_Brg(0).Text = ""
            TxtCr_Brg(1).Text = ""
            
            TxtCr_Brg_KeyUp 0, 0, 0
            
            .Visible = True
            
            TxtCr_Brg(0).SetFocus
            
        Else
            .Visible = False
        End If
    End With


End Sub

Private Sub CmdCbang_Click()

    With TDB_Cbang
        If .Visible = False Then
            
            .Left = Picture1.Left + CmdCbang.Left + CmdCbang.Width / 2 - .Width / 2
            .Top = Picture1.Top + CmdCbang.Top + CmdCbang.Height + 15
            
            TxtCrCbang.Text = ""
            
            TxtCrCbang_KeyUp 0, 0
            
            .Visible = True
            
            TxtCrCbang.SetFocus
            
        Else
            .Visible = False
        End If
    End With


End Sub

Private Sub cmddel_Click()
    
    If ArrBrg.UpperBound(1) = 1 And ArrBrg(1, 1) = Empty Then Exit Sub
    If ArrBrg.UpperBound(1) = 1 Then Exit Sub
    
    If ArrBrg(GridBrg.Bookmark, 6) <> "" Then
    
        Dim sql As String
        Dim rs As Recordset
        
        sql = "delete from Tb_Order_Detail where id=" & ArrBrg(GridBrg.Bookmark, 6)
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
'        Dim comd1 As Command
'        Set comd1 = New ADODB.Command
'        With comd1
'            .ActiveConnection = kon
'            .CommandText = "kurangi_stock"
'            .CommandType = adCmdStoredProc
'            .Parameters("@jml_stock").Value = ArrBrg(GridBrg.Bookmark, 2)
'            .Parameters("@kode_brg").Value = ArrBrg(GridBrg.Bookmark, 0)
'            .Parameters("@kode_cbang").Value = Trim(TKodeCbang.Text)
'            .Execute
'        End With

        
    End If
    
    GridBrg.Columns(5).FooterText = CDbl(GridBrg.Columns(5).FooterText) - CDbl(ArrBrg(GridBrg.Bookmark, 5))
    
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
    
    If TKodeBrg.Text = "" Then Exit Sub
    
    Dim qty As Double
        If TDB_Qty.ValueIsNull Then
            qty = 0
        Else
            qty = Replace(Trim(TDB_Qty.Value), ",", "")
        End If
    
    Dim hrg As Double
        If TDB_Hrg.ValueIsNull Then
            hrg = 0
        Else
            hrg = Replace(Trim(TDB_Hrg.Value), ",", "")
        End If
    
    If qty = 0 Or hrg = 0 Then Exit Sub
        
    Dim totJml As Double
        If TDB_Jml.ValueIsNull Then
            totJml = 0
        Else
            totJml = Replace(Trim(TDB_Jml.Value), ",", "")
        End If
    
    If PeriksaBrgAdd(Trim(TKodeBrg.Text)) = True Then
        MsgBox "Barang yang akan ditambahkan sudah ada"
        TKodeBrg.SetFocus
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
        
        ArrBrg(a, 0) = TKodeBrg.Text
        ArrBrg(a, 1) = TNamaBrg.Text
        ArrBrg(a, 2) = qty
        ArrBrg(a, 3) = TSatuan.Text
        ArrBrg(a, 4) = hrg
        ArrBrg(a, 5) = totJml
        ArrBrg(a, 6) = ""
        
        GridBrg.ReBind
        GridBrg.Refresh
        
        GridBrg.MoveLast
        
        GridBrg.Columns(5).FooterText = CDbl(GridBrg.Columns(5).FooterText) + totJml
        
            TKodeBrg.Text = ""
            TNamaBrg.Text = ""
            TSatuan.Text = ""
            TDB_Qty.Value = Null
            TDB_Hrg.Value = Null
            TDB_Jml.Value = Null
        
        TKodeBrg.SetFocus
        
    End If
    
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
        txtnama.SetFocus
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
    .Top = 350
End With

TBukti.Enabled = False
DTgl.Enabled = False
TKodeSupp.Enabled = False
CmdSupp.Enabled = False
TKodeCbang.Enabled = False
CmdCbang.Enabled = False
txtnama.Enabled = False

GridBrg.Array = ArrBrg
    
    ArrBrg.ReDim 0, 0, 0, 0
    ArrBrg.ReDim 1, 1, 1, GridBrg.Columns.Count
        GridBrg.ReBind
        GridBrg.Refresh

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
Private Sub GridCbang_DblClick()
    
    If GridCbang.Row < 0 Then Exit Sub
    
    TKodeCbang.Text = GridCbang.Columns(0).Text
    TNamaCbang.Text = GridCbang.Columns(1).Text
    
    TDB_Cbang.Visible = False
    TKodeCbang.SetFocus
    
End Sub

Private Sub GridCbang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then GridCbang_DblClick
    If KeyCode = vbKeyEscape Then CmdCbang_Click
End Sub

Private Sub GridCr_Brg_DblClick()
    
    If GridCr_Brg.Row < 0 Then Exit Sub
    
        TKodeBrg.Text = GridCr_Brg.Columns(0).Text
'        TNamaBrg.Text = GridCr_Brg.Columns(1).Text
'        TSatuan.Text = GridCr_Brg.Columns(2).Text
    
    TKodeBrg_LostFocus
    
    TDB_Brg.Visible = False
    TDB_Qty.SetFocus
    
End Sub

Private Sub GridCr_Brg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then GridCr_Brg_DblClick
    If KeyCode = vbKeyEscape Then CmdBrg_Click
End Sub

Private Sub GridSupp_DblClick()
    
    If GridSupp.Row < 0 Then Exit Sub
    
        TKodeSupp.Text = GridSupp.Columns(0).Text
        TNamaSupp.Text = GridSupp.Columns(1).Text
    
    TDBSupp.Visible = False
    TKodeSupp.SetFocus
    
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
        sql = "select nobukti from Tb_Order where nobukti='" & Trim(TBukti.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        If Not .EOF Then
            MsgBox "No bukti sudah ada"
            TBukti.SetFocus
            
            txt_ket.Enabled = False
            DTgl.Enabled = False
            TKodeSupp.Enabled = False
            CmdSupp.Enabled = False
            TKodeCbang.Enabled = False
            CmdCbang.Enabled = False
            Cmd_Simpan.Enabled = False
            txtnama.Enabled = False
            
        Else
            
            DTgl.Enabled = True
            TKodeSupp.Enabled = True
            CmdSupp.Enabled = True
            TKodeCbang.Enabled = True
            CmdCbang.Enabled = True
            Cmd_Simpan.Enabled = True
            txt_ket.Enabled = True
            txtnama.Enabled = True
            
            TKodeSupp.Text = ""
            TNamaSupp.Text = ""
            TKodeCbang.Text = ""
            TNamaCbang.Text = ""
            txt_ket.Text = ""
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
Private Sub TDB_Hrg_Change()
    JmlAdd
End Sub

Private Sub TDB_Hrg_GotFocus()
    Call Focus_(TDB_Hrg)
End Sub

Private Sub TDB_Qty_Change()
    JmlAdd
End Sub

Private Sub TDB_Qty_GotFocus()
    Call Focus_(TDB_Qty)
End Sub

Private Sub TDB_Qty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then CmdOk.SetFocus
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


Private Sub TKodeBrg_GotFocus()
    Call Focus_(TKodeBrg)
End Sub

Private Sub TKodeBrg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TDB_Qty.SetFocus
    If KeyCode = vbKeyF3 Then CmdBrg_Click
End Sub

Private Sub TKodeBrg_LostFocus()
    
    If TKodeBrg.Text = "" Then Exit Sub
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from tb_barang where kode='" & Trim(TKodeBrg.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If Not .EOF Then
                TNamaBrg.Text = IIf(Not IsNull(!nama), !nama, "")
                TDB_Hrg.Value = IIf(Not IsNull(!harga), !harga, Null)
                TSatuan.Text = IIf(Not IsNull(!satuan), !satuan, "")
            Else
                MsgBox "Barang yang anda masukkan tidak ditemukan"
                TNamaBrg.Text = ""
                TSatuan.Text = ""
            End If
        End With
        
    
End Sub

Private Sub TKodeCbang_GotFocus()
    Call Focus_(TKodeCbang)
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
End Sub

Private Sub TKodeSupp_LostFocus()
    
    If TKodeSupp.Text = "" Then Exit Sub
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "select kode,nama_toko from tb_supplier where kode='" & Trim(TKodeSupp.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        If Not .EOF Then
            
            TNamaSupp.Text = IIf(Not IsNull(!Nama_Toko), !Nama_Toko, "")
        
        Else
            
            MsgBox "Supplier tidak ditemukan"
            TNamaSupp.Text = ""
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
        sql = "select top 100 * from Tb_Order " 'where kodecounter in (select kode_counter from VIEW_Counter_User where id_user=" & Flag_tempat & ")"
        
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

    sql = "select * from Tb_Supplier"
        
    If Txt_Cr_Supp(0).Text <> "" Or Txt_Cr_Supp(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where Kode like  '%" & Trim(Txt_Cr_Supp(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Toko like '%" & Trim(Txt_Cr_Supp(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by Kode,Nama_Toko asc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set GridSupp.DataSource = rs
        GridSupp.Refresh
    
End Sub

Private Sub txt_ket_GotFocus()
    Call Focus_(txt_ket)
End Sub

Private Sub txt_ket_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If CmdAdd.Enabled = True Then CmdAdd.SetFocus
    End If
End Sub

Private Sub TxtCr_Brg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then GridCr_Brg.SetFocus
    If KeyCode = vbKeyEscape Then CmdBrg_Click
End Sub

Private Sub TxtCr_Brg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sql As String
Dim rs As Recordset

    sql = "select * from Tb_barang"
        
    If TxtCr_Brg(0).Text <> "" Or TxtCr_Brg(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where Kode like  '%" & Trim(TxtCr_Brg(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama like '%" & Trim(TxtCr_Brg(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by Kode,Nama asc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set GridCr_Brg.DataSource = rs
        GridCr_Brg.Refresh
    
End Sub

Private Sub TxtCrCbang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then GridCbang.SetFocus
    If KeyCode = vbKeyEscape Then CmdCbang_Click
End Sub

Private Sub TxtCrCbang_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sql As String
Dim rs As Recordset

    sql = "select * from view_counter_user where id_user=" & Flag_tempat
        
    If TxtCrCbang.Text <> "" Then
            sql = sql & " and  nama_counter like  '%" & Trim(TxtCrCbang.Text) & "%'"
    End If
    
    sql = sql & " order by nama_counter asc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set GridCbang.DataSource = rs
        GridCbang.Refresh
    
End Sub

Private Sub txtnama_GotFocus()
    Call Focus_(txtnama)
End Sub

Private Sub txtnama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_ket.SetFocus
    End If
End Sub
