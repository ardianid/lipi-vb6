VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Stock_Awal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Awal"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Stock_Awal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D tdb_cabang 
      Height          =   5895
      Left            =   1320
      TabIndex        =   35
      Top             =   8640
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   10398
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Stock_Awal.frx":27C92
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Stock_Awal.frx":27CAE
      Childs          =   "Frm_Stock_Awal.frx":27D5A
      Begin VB.TextBox txt_cr_cabang 
         Height          =   360
         Left            =   1200
         TabIndex        =   37
         Top             =   600
         Width           =   4215
      End
      Begin IsButton_Ard.isButton cmd_ok_cabang 
         Height          =   375
         Left            =   3360
         TabIndex        =   36
         Top             =   5880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Icon            =   "Frm_Stock_Awal.frx":27D76
         Style           =   1
         Caption         =   "OK"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin TrueOleDBGrid60.TDBGrid grid_cabang 
         Height          =   4575
         Left            =   240
         OleObjectBlob   =   "Frm_Stock_Awal.frx":27D92
         TabIndex        =   38
         Top             =   1080
         Width           =   5175
      End
      Begin IsButton_Ard.isButton isButton1 
         Height          =   375
         Left            =   4440
         TabIndex        =   39
         Top             =   5880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Icon            =   "Frm_Stock_Awal.frx":2AD1A
         Style           =   1
         Caption         =   "CANCEL"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang"
         Height          =   240
         Index           =   21
         Left            =   360
         TabIndex        =   41
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
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
         Index           =   18
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1065
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   5400
         Y1              =   480
         Y2              =   480
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_barang 
      Height          =   6015
      Left            =   -5880
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   10610
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Stock_Awal.frx":2AD36
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Stock_Awal.frx":2AD52
      Childs          =   "Frm_Stock_Awal.frx":2ADFE
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   3840
         TabIndex        =   20
         Top             =   600
         Width           =   2055
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   5655
      End
      Begin TrueOleDBGrid60.TDBGrid grd_barang 
         Height          =   4815
         Left            =   240
         OleObjectBlob   =   "Frm_Stock_Awal.frx":2AE1A
         TabIndex        =   17
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN DATA BARANG"
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
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   180
         Width           =   2190
      End
   End
   Begin VB.PictureBox j 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   -5520
      ScaleHeight     =   5625
      ScaleWidth      =   5865
      TabIndex        =   8
      Top             =   8400
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmd_x 
         Caption         =   "x"
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
         Left            =   5400
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5865
         TabIndex        =   9
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7935
      ScaleWidth      =   11175
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin TrueOleDBGrid60.TDBGrid grd_daftar 
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "Frm_Stock_Awal.frx":2DC61
         TabIndex        =   3
         Top             =   1200
         Width           =   10815
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Perhatian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8160
         TabIndex        =   30
         Top             =   120
         Width           =   2895
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Form ini digunakan hanya untuk memasukkan stock awal dari barang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   615
            Index           =   6
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   22
         Top             =   6960
         Visible         =   0   'False
         Width           =   7815
         Begin VB.CommandButton Cmd_Cari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6840
            TabIndex        =   29
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Txt_Cr_Nama 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Left            =   4560
            TabIndex        =   26
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox Txt_Cr_Kode 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Left            =   1560
            TabIndex        =   23
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   3120
            TabIndex        =   28
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   4440
            TabIndex        =   27
            Top             =   360
            Width           =   60
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1440
            TabIndex        =   24
            Top             =   360
            Width           =   60
         End
      End
      Begin VB.CommandButton Cmd_Keluar 
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   21
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8280
         TabIndex        =   7
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Frame Frame2 
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
         TabIndex        =   1
         Top             =   120
         Width           =   7935
         Begin VB.CommandButton cmd_browse_cabang 
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
            Left            =   5160
            TabIndex        =   34
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txt_cabang 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Left            =   3000
            TabIndex        =   33
            Top             =   840
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton Cmd_Browse_Brg 
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
            Left            =   3480
            TabIndex        =   18
            Top             =   360
            Width           =   375
         End
         Begin MSComCtl2.DTPicker dtp_tgl 
            Height          =   330
            Left            =   600
            TabIndex        =   11
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            Format          =   49676289
            CurrentDate     =   39211
         End
         Begin VB.TextBox txt_kode 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Left            =   1320
            TabIndex        =   6
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmd_tampil 
            Caption         =   "Tampil"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6480
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cabang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   7
            Left            =   2160
            TabIndex        =   32
            Top             =   840
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   5
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Visible         =   0   'False
            Width           =   270
         End
      End
   End
End
Attribute VB_Name = "Frm_Stock_Awal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim arr_barang As New XArrayDB
Dim sql As String, kode_b As String

Dim Moving As Boolean
Dim yold, xold As Long
Dim kode_cabang As String

Private Sub cmd_browse_brg_Click()

With pic_barang
    
    If .Visible = False Then
        
        .Left = Picture1.Left + Frame2.Left + Cmd_Browse_Brg.Left + Cmd_Browse_Brg.Width / 2 - .Width / 2
        .Top = Picture1.Top + Frame2.Top + Cmd_Browse_Brg.Top + Cmd_Browse_Brg.Height + 15
        
        Txt_Kode.Text = ""
        txt(0).Text = ""
        txt(1).Text = ""
        pic_barang.Visible = True
        txt(0).SetFocus
    Else
        .Visible = False
    End If
    
End With

End Sub

Private Sub cmd_browse_cabang_Click()

With tdb_cabang
    
    If .Visible = False Then
        
        .Left = Picture1.Left + Frame2.Left + cmd_browse_cabang.Left + cmd_browse_cabang.Width / 2 - .Width / 2
        .Top = Picture1.Top + Frame2.Top + cmd_browse_cabang.Top + cmd_browse_cabang.Height + 15
        
        txt_cr_cabang.Text = ""
        
        txt_cr_cabang_KeyUp 0, 0
        
        .Visible = True
        
        txt_cr_cabang.SetFocus
        
    Else
        .Visible = False
    End If
    
End With


End Sub

Private Sub Cmd_Cari_Click()
        
     If arr_daftar.UpperBound(1) <= 0 Then Exit Sub
        If Txt_Cr_Kode.Text = "" And Txt_Cr_Nama.Text = "" Then Exit Sub
        
     Dim RowFound As Long
     If Txt_Cr_Kode.Text <> "" Then
     
        RowFound = arr_daftar.Find(arr_daftar.LowerBound(1), 2, CStr(Txt_Cr_Kode.Text), XORDER_ASCEND, XCOMP_GE, XTYPE_STRING)
    
     ElseIf Txt_Cr_Nama.Text <> "" Then
           
        RowFound = arr_daftar.Find(arr_daftar.LowerBound(1), 3, CStr(Txt_Cr_Nama.Text), XORDER_ASCEND, XCOMP_GE, XTYPE_STRING)
    
    End If
    
        If RowFound >= 0 Then
            grd_daftar.Bookmark = RowFound
            grd_daftar.Col = 3
           ' Grid_Tambah.Columns(2).Value = vbChecked
'            Grid_Tambah.SetFocus
        End If
    
End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
        
End Sub

Private Sub cmd_simpan_Click()

On Error GoTo err_simpan

'    Dim sql1, sql2 As String
'    Dim rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim a As Long
    Dim comd As Command
    Dim comd1 As Command
    Dim jangan As Integer
    
    Dim rs1 As Recordset
    
    If MsgBox("Yakin semua data yang anda masukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        On Error GoTo 0
        Exit Sub
    End If
        
        
        kon.BeginTrans
        
        grd_barang.MoveNext
        If grd_barang.EOF Then grd_barang.MovePrevious
        grd_daftar.MoveFirst
        
        For a = 1 To arr_daftar.UpperBound(1)
            If arr_daftar(a, 4) <> Empty _
                Or arr_daftar(a, 5) <> Empty _
                    Or arr_daftar(a, 6) <> Empty _
                        Or arr_daftar(a, 7) <> Empty Then
                
                Set comd1 = New ADODB.Command
                With comd1
                    .ActiveConnection = kon
                    .CommandText = "Tambah_Stock"
                    .CommandType = adCmdStoredProc
                    .Parameters("@id_brg").Value = arr_daftar(a, 2)
'                    .Parameters("@brg_i").Value = 0 'arr_daftar(a, 4)
'                    .Parameters("@brg_o").Value = 0
'                    .Parameters("@tg").Value = Format(Trim(dtp_tgl.Value), "yyyy/mm/dd")
'                    .Parameters("@Ke").Value = "Stock awal"
                    .Parameters("@brg_m").Value = CDbl(arr_daftar(a, 4))
                    .Parameters("@brg_rusak").Value = CDbl(arr_daftar(a, 5))
                    .Parameters("@brg_lengkap").Value = CDbl(arr_daftar(a, 6))
                    .Parameters("@brg_kurang").Value = CDbl(arr_daftar(a, 7))
'                    .Parameters("@kode_counter").Value = rs1!kode_counter
                    
                .Execute
                
                End With
                
                comd1.ActiveConnection = Nothing
            
            End If
        Next a
        
        jangan = CInt(MsgBox("Data telah disimpan", vbOKOnly + vbInformation, "Informasi"))
        kon.CommitTrans
        kosong_daftar
        Txt_Kode.Text = "Semua"
        Txt_Kode.SetFocus
        
        On Error GoTo 0
        Exit Sub
        
err_simpan:
    kon.RollbackTrans
    Dim psn As Integer
        psn = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
        
End Sub

Private Sub cmd_tampil_Click()
    isi
End Sub

Private Sub cmd_x_Click()
    pic_barang.Visible = False
    Txt_Kode.SetFocus
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

grd_daftar.Array = arr_daftar

grd_barang.Array = arr_barang

With pic_barang
    .Left = 3360
    .Top = 875
End With

Txt_Kode.Text = ""
Txt_Kode.Text = "Semua"
txt_cabang.Text = "Semua"

kosong_daftar

dtp_tgl.Value = Format(Date, "dd/mm/yyyy")

isi_barang
'txt_cr_cabang_KeyUp 0, 0

'' akses command ''

'    hak_akses_percommand CStr(Me.Name)
'
'    cmd_simpan.Enabled = c_tambah

'' stop here ''

End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub kosong_barang()
    arr_barang.ReDim 0, 0, 0, 0
    grd_barang.ReBind
    grd_barang.Refresh
End Sub

Private Sub isi_barang()

On Error GoTo er_handler

    Dim rs_barang As New ADODB.Recordset
    Dim comd As Command
    
    kosong_barang
        
        Set comd = New ADODB.Command
        With comd
            .ActiveConnection = kon
            .CommandText = "lht_brg_peny_stock"
            .CommandType = adCmdStoredProc
            .Parameters("@kriteria").Value = 0
        End With
        
        Set rs_barang = comd.Execute
'            rs_barang.CursorType = adOpenKeyset
        'rs_barang.Open sql1, kon, adOpenKeyset
            If Not rs_barang.EOF Then
                
'                rs_barang.MoveLast
'                rs_barang.MoveFirst
'
                lanjut_barang rs_barang
            End If
        rs_barang.Close
        comd.ActiveConnection = Nothing
        
        On Error GoTo 0
        Exit Sub
        
er_handler:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information")
            Err.Clear
        
End Sub

Private Sub lanjut_barang(rs_barang As Recordset)
    Dim nama_counter, kode_barang, nama_barang As String
    Dim a As Long
            
            a = 1
                Do While Not rs_barang.EOF
                    arr_barang.ReDim 1, a, 0, 3
                    grd_barang.ReBind
                    grd_barang.Refresh
                        
'                        If Not IsNull(rs_barang("Jenis_Barang")) Then
'                            nama_counter = rs_barang("Jenis_Barang")
'                        Else
'                            nama_counter = ""
'                        End If
                        
                        If Not IsNull(rs_barang("Kode")) Then
                            kode_barang = rs_barang("Kode")
                        Else
                            kode_barang = ""
                        End If
                        
                        If Not IsNull(rs_barang("Nama")) Then
                            nama_barang = rs_barang("Nama")
                        Else
                            nama_barang = ""
                        End If
                        
                     arr_barang(a, 0) = ""
                     arr_barang(a, 1) = kode_barang
                     arr_barang(a, 2) = nama_barang
                     
                     a = a + 1
                     rs_barang.MoveNext
                     Loop
                     grd_barang.ReBind
                     grd_barang.Refresh
End Sub
Private Sub isi()

On Error GoTo er_isi

Dim rs_daftar As New ADODB.Recordset
Dim comd As Command

    kosong_daftar
    
    If Txt_Kode.Text = "" Then
        Txt_Kode.Text = "Semua"
    End If
    
'    If txt_cabang.Text = "" Or UCase(txt_cabang.Text) = UCase("Semua") Then
'        MsgBox "Cabang harus diisi"
'        Exit Sub
'    End If
    
    Set comd = New ADODB.Command
    
    With comd
        .ActiveConnection = kon
        .CommandText = "lht_peny_stock"
        .CommandType = adCmdStoredProc
    
    If Txt_Kode.Text = "Semua" Then
        .Parameters("@ada_kode").Value = 0
        .Parameters("@kode_brg").Value = ""
    ElseIf Txt_Kode.Text <> "Semua" And Txt_Kode.Text <> "" Then
        .Parameters("@ada_kode").Value = 1
        .Parameters("@kode_brg").Value = Trim(Txt_Kode.Text)
    End If
        
'        .Parameters("@kode_counter").Value = kode_cabang
        
    End With
    
    Set rs_daftar = comd.Execute
        If Not rs_daftar.EOF Then
            
'            rs_daftar.MoveLast
'            rs_daftar.MoveFirst
            
            isi_daftar rs_daftar
            
        End If
   rs_daftar.Close
   comd.ActiveConnection = Nothing
    
   On Error GoTo 0
   Exit Sub
   
er_isi:
   Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information")
            Err.Clear
End Sub

Private Sub isi_daftar(rs_daftar As Recordset)
    
    Dim id_barang, kode_counter, nama_counter, kode_barang, nama_barang, stock As String
    Dim a As Long
        
        a = 1
            Do While Not rs_daftar.EOF
                
                arr_daftar.ReDim 1, a, 0, 10
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
'                    id_barang = rs_daftar("id_barang")
                    
                    If Not IsNull(rs_daftar("Kode_Jenis")) Then
                        kode_counter = rs_daftar("Kode_Jenis")
                    Else
                        kode_counter = ""
                    End If
                    
                    If Not IsNull(rs_daftar("Nama_Jenis")) Then
                        nama_counter = rs_daftar("Nama_Jenis")
                    Else
                        nama_counter = ""
                    End If
                    
                    If Not IsNull(rs_daftar("kode")) Then
                        kode_barang = rs_daftar("kode")
                    Else
                        kode_barang = ""
                    End If
                    
                    If Not IsNull(rs_daftar("nama")) Then
                        nama_barang = rs_daftar("nama")
                    Else
                        nama_barang = ""
                    End If
                    
'                    If Not IsNull(rs_daftar("jml_stock")) Then
'                        stock = rs_daftar("jml_stock")
'                    Else
'                        stock = ""
'                    End If
                    
'               arr_daftar(a, 0) = id_barang
               arr_daftar(a, 0) = kode_counter
               arr_daftar(a, 1) = nama_counter
               arr_daftar(a, 2) = kode_barang
               arr_daftar(a, 3) = nama_barang
               arr_daftar(a, 4) = 0
               arr_daftar(a, 5) = 0
               arr_daftar(a, 6) = 0
               arr_daftar(a, 7) = 0
'               arr_daftar(a, 9) = "-"
               
            a = a + 1
            rs_daftar.MoveNext
            Loop
            
            If arr_daftar.UpperBound(1) = 1 Then
                arr_daftar.ReDim 1, arr_daftar.UpperBound(1) + 1, 0, 10
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
               arr_daftar(a, 0) = ""
               arr_daftar(a, 1) = ""
               arr_daftar(a, 2) = ""
               arr_daftar(a, 3) = ""
               arr_daftar(a, 4) = ""
               arr_daftar(a, 5) = ""
               arr_daftar(a, 6) = ""
               arr_daftar(a, 7) = ""
'               arr_daftar(a, 8) = 0
'               arr_daftar(a, 9) = "-"
            End If
                
            grd_daftar.ReBind
            grd_daftar.Refresh
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    
    End If

End Sub

Private Sub grd_barang_Click()
    On Error Resume Next
        If arr_barang.UpperBound(1) > 0 Then
            kode_b = arr_barang(grd_barang.Bookmark, 1)
        End If
End Sub

Private Sub grd_barang_DblClick()

If arr_barang.UpperBound(1) > 0 Then
    Txt_Kode.Text = kode_b
    pic_barang.Visible = False
    Txt_Kode.SetFocus
End If
    
End Sub

Private Sub grd_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_barang_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
    End If
End Sub

Private Sub grd_barang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_barang_Click
End Sub

Private Sub grd_daftar_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'    On Error GoTo er_c
'
'    If ColIndex = 6 Then
'
'        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
'
'
'        If CDbl(arr_daftar(grd_daftar.Bookmark, 5)) > CDbl(arr_daftar(grd_daftar.Bookmark, ColIndex)) Then
'
'            arr_daftar(grd_daftar.Bookmark, 7) = 0
'
'            arr_daftar(grd_daftar.Bookmark, 8) = CDbl(arr_daftar(grd_daftar.Bookmark, 5)) - CDbl(arr_daftar(grd_daftar.Bookmark, ColIndex))
'
'        Else
'
'            arr_daftar(grd_daftar.Bookmark, 8) = 0
'
'            arr_daftar(grd_daftar.Bookmark, 7) = CDbl(arr_daftar(grd_daftar.Bookmark, ColIndex)) - CDbl(arr_daftar(grd_daftar.Bookmark, 5))
'        End If
'        Exit Sub
'
'    End If
'
'    If ColIndex = 9 Then
'        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
'    End If
'
'
'
'    Exit Sub
'
'er_c:
'
'    Dim psn
'        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information")
'        Err.Clear
End Sub

Private Sub grd_daftar_HeadClick(ByVal ColIndex As Integer)

'On Error GoTo er_h
'
'    Dim sql2 As String
'    Dim rs_daftar As New ADODB.Recordset
'
'
'    If sql = "" Then
'        Exit Sub
'    End If
'
'
'    If arr_daftar.UpperBound(1) = 0 Then
'        Exit Sub
'    End If
'
'    sql2 = ""
'    sql2 = sql2 & sql
'
'        Select Case ColIndex
'
'            Case 2
'
'                sql2 = sql2 & ",nama_counter"
'
'            Case 3
'
'                sql2 = sql2 & ",kode_barang"
'
'            Case 4
'
'                sql2 = sql2 & ",nama_barang"
'
'            Case 5
'
'                sql2 = sql2 & ",jml_stock"
'
'        End Select
'
'        rs_daftar.Open sql2, kon, adOpenKeyset
'            If Not rs_daftar.EOF Then
'
'                rs_daftar.MoveLast
'                rs_daftar.MoveFirst
'
'                isi_daftar rs_daftar
'            End If
'        rs_daftar.Close
'
'    Exit Sub
'
'er_h:
'    Dim psn
'            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
'            Err.Clear
End Sub


Private Sub pic_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pic_barang.Visible = False
        Txt_Kode.Visible = False
    End If
End Sub

Private Sub grid_cabang_DblClick()
    
    If grid_cabang.Row < 0 Then Exit Sub
    
    txt_cabang.Text = grid_cabang.Columns(1).Text
    kode_cabang = grid_cabang.Columns(0).Text
    
    tdb_cabang.Visible = False: txt_cabang.SetFocus
    
End Sub

Private Sub grid_cabang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_cabang_DblClick
    If KeyCode = vbKeyEscape Then tdb_cabang.Visible = False: txt_cabang.SetFocus
End Sub

Private Sub pic_barang_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = True
If Moving = True Then
   yold = y
   xold = x
End If
End Sub

Private Sub pic_barang_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Moving = True Then
   pic_barang.Top = pic_barang.Top - (yold - y)
   pic_barang.Left = pic_barang.Left - (xold - x)
End If

End Sub

Private Sub pic_barang_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = False
End Sub

Private Sub tdb_cabang_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = True
If Moving = True Then
   yold = y
   xold = x
End If
End Sub

Private Sub tdb_cabang_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Moving = True Then
   tdb_cabang.Top = tdb_cabang.Top - (yold - y)
   tdb_cabang.Left = tdb_cabang.Left - (xold - x)
End If

End Sub

Private Sub tdb_cabang_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Moving = False
End Sub

Private Sub txt_cabang_GotFocus()
    Call Focus_(txt_cabang)
End Sub

Private Sub txt_cabang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_tampil.SetFocus
    If KeyCode = vbKeyF3 Then cmd_browse_cabang_Click
End Sub

Private Sub txt_cabang_LostFocus()

    kode_cabang = ""
    
    If txt_cabang.Text = "" Then txt_cabang.Text = "Semua"
    
    If txt_cabang.Text <> "semua" Then
    
    Dim rs As Recordset
    Dim sql As String
        
        sql = "select kode_counter from VIEW_Counter_User where nama_counter='" & Trim(txt_cabang.Text) & "' and id_user=" & Flag_tempat
        
        Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
        
        If Not rs.EOF Then
                
                kode_cabang = IIf(Not IsNull(rs!kode_counter), rs!kode_counter, "")
        
        Else
        
            Dim konfirm As Integer
                konfirm = CInt(MsgBox("Cabang yang anda masukkan tidak ditemukan", vbOKOnly + vbInformation, "Informasi"))
                
                txt_cabang.SetFocus
                
                
        End If
        
        Set rs = Nothing
    
    End If
    
End Sub

Private Sub txt_cr_cabang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_cabang.SetFocus
    If KeyCode = vbKeyEscape Then tdb_cabang.Visible = False: txt_cabang.SetFocus
End Sub

Private Sub txt_cr_cabang_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sql As String
    Dim rs As Recordset
                    
        sql = "select * from view_counter_user where id_user=" & Flag_tempat
        
        If txt_cr_cabang.Text <> "" Then sql = sql & " and nama_counter like '%" & Trim(txt_cr_cabang.Text) & "%'"
        
        sql = sql & " order by kode_counter asc"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        Set grid_cabang.DataSource = rs
            grid_cabang.Refresh

End Sub

Private Sub Txt_Cr_Kode_Change()
    If Txt_Cr_Nama.Text <> "" Then Txt_Cr_Nama.Text = ""
End Sub

Private Sub Txt_Cr_Kode_GotFocus()
    Call Focus_(Txt_Cr_Kode)
End Sub

Private Sub Txt_Cr_Kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Cr_Nama.SetFocus
End Sub

Private Sub Txt_Cr_Nama_Change()
    If Txt_Cr_Kode.Text <> "" Then Txt_Cr_Kode.Text = ""
End Sub

Private Sub Txt_Cr_Nama_GotFocus()
    Call Focus_(Txt_Cr_Nama)
End Sub

Private Sub Txt_Cr_Nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_tampil.SetFocus
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            txt(0).SelStart = 0
            txt(0).SelLength = Len(txt(0))
        Case 1
            txt(1).SelStart = 0
            txt(1).SelLength = Len(txt(1))
    End Select
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
        Txt_Kode.SetFocus
    End If
    
    If KeyCode = 13 Then
        grd_barang_DblClick
    End If
    
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo er_u

    Dim rs_barang As New ADODB.Recordset
    Dim comd As Command
    
' If arr_barang.UpperBound(1) > 0 Then
 
        kosong_barang
        
        Set comd = New ADODB.Command
            
        comd.ActiveConnection = kon
        comd.CommandText = "lht_brg_peny_stock"
        comd.CommandType = adCmdStoredProc
        
    If txt(0).Text <> "" Or txt(1).Text <> "" Then
        comd.Parameters("@kriteria").Value = 1
        Select Case Index
            Case 0
                comd.Parameters("@kode_sel").Value = Trim(txt(0).Text)
            Case 1
                comd.Parameters("@nama_sel").Value = Trim(txt(1).Text)
        End Select
    Else
        comd.Parameters("@kriteria").Value = 0
    End If
    
         Set rs_barang = comd.Execute
'            rs_barang.CursorType = adOpenKeyset
        'rs_barang.Open sql1, kon, adOpenKeyset
            If Not rs_barang.EOF Then
                
'                rs_barang.MoveLast
'                rs_barang.MoveFirst
                
                lanjut_barang rs_barang
            End If
        rs_barang.Close
        comd.ActiveConnection = Nothing
'End If
      
Exit Sub

er_u:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information")
            Err.Clear
                       
End Sub

Private Sub Txt_Kode_GotFocus()
    Txt_Kode.SelStart = 0
    Txt_Kode.SelLength = Len(Txt_Kode)
End Sub

Private Sub Txt_Kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        
        Txt_Kode.Text = ""
        txt(0).Text = ""
        txt(1).Text = ""
        pic_barang.Visible = True
        txt(0).SetFocus
        
    End If
    If KeyCode = 13 Then cmd_tampil.SetFocus
    
End Sub
