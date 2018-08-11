VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form Frm_Mast_Type_Brg 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Type Barang"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Mast_Cr_Byr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Cari 
      Height          =   2055
      Left            =   -4320
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   3625
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Mast_Cr_Byr.frx":27C92
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Mast_Cr_Byr.frx":27CAE
      Childs          =   "Frm_Mast_Cr_Byr.frx":27D5A
      Begin VB.CommandButton Cmd_Cancel 
         Caption         =   "&Keluar"
         Height          =   400
         Left            =   4680
         TabIndex        =   26
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Cmd_OK 
         Caption         =   "&OK"
         Height          =   400
         Left            =   3840
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Txt_Cr_Nama 
         Height          =   320
         Left            =   1920
         TabIndex        =   24
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox Txt_Cr_Kode 
         Height          =   320
         Left            =   1920
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   3
         Left            =   1800
         TabIndex        =   31
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   2
         Left            =   1800
         TabIndex        =   30
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   29
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   600
         Width           =   420
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   5400
         Y1              =   360
         Y2              =   360
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
         TabIndex        =   27
         Top             =   120
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3945
      ScaleWidth      =   7425
      TabIndex        =   20
      Top             =   2520
      Width           =   7455
      Begin TrueOleDBGrid60.TDBGrid Grid_Status 
         Height          =   3735
         Left            =   120
         OleObjectBlob   =   "Frm_Mast_Cr_Byr.frx":27D76
         TabIndex        =   21
         Top             =   120
         Width           =   7215
      End
   End
   Begin VB.Frame v 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   7455
      Begin VB.CommandButton cmd_keluar 
         Caption         =   "&Keluar"
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
         Left            =   6360
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Cari 
         Caption         =   "&Cari"
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
         Left            =   5520
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmd_hapus 
         Caption         =   "&Hapus"
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
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmd_rubah 
         Caption         =   "&Rubah"
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
         Left            =   3840
         TabIndex        =   14
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
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   495
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
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   495
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
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_tambah 
         Caption         =   "&Tambah"
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
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frame2"
         Height          =   855
         Left            =   2880
         TabIndex        =   12
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
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "&Simpan"
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
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmd_batal 
         Caption         =   "&Batal"
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
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   7425
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txt_ket 
         Height          =   320
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox txt_nama 
         Height          =   320
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txt_kode 
         Height          =   320
         Left            =   720
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ket :"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis :"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode :"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   120
         Width           =   465
      End
   End
End
Attribute VB_Name = "Frm_Mast_Type_Brg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim Moving As Boolean
Dim yold, xold As Long
'Dim idnya As Double

Private Sub IsiSemua()
    
    Dim sql As String
        sql = "select * from Tb_Type_Brg order by Kode desc"
        
        Set Rs_Nav = New ADODB.Recordset
            Rs_Nav.Open sql, kon, adOpenKeyset
        
        Set Grid_Status.DataSource = Rs_Nav
            Grid_Status.Refresh
    
End Sub

Private Sub Cbo_Kondisi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_ket.SetFocus
End Sub

Private Sub Cmd_Batal_Click()

    rubah = False
    
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
            
'            If TypeOf n Is DTPicker Then n.Enabled = False
            
            If TypeOf n Is TDBContainer3D Then
                n.Visible = False
            End If
            
'            If TypeOf n Is ComboBox Then n.Enabled = False
            
        Next
    Set n = Nothing
    
    Cmd_Tambah.SetFocus
        
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
    
    If MsgBox("Yakin akan menghapus Jenis " & Txt_Nama.Text, vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then
        On Error GoTo 0
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "delete from Tb_Type_Brg where Kode ='" & Trim(Txt_Kode.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
'    Dim konfirm As Integer
'        konfirm = CInt(MsgBox("Data telah dihapus", vbOKOnly + vbInformation, "Informasi"))
    
    IsiSemua
    
    On Error GoTo 0
    Exit Sub

err_handler:
    
    Dim konfirm As Integer
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
'        If Len(Txt_Cr_Kode.Text) = 10 Then
            .Find "Kode like '%" & Trim(Txt_Cr_Kode.Text) & "%'"
'        End If
        ElseIf Txt_Cr_Nama.Text <> "" And Txt_Cr_Kode.Text = "" Then
            .Find "Jenis like '%" & Trim(Txt_Cr_Nama.Text) & "%'"
'        ElseIf Txt_Cr_Nama.Text <> "" And Txt_Cr_Kode.Text <> "" Then
'            .Find "Tgl='" & Format(Trim(Txt_Cr_Kode.Text), "yyyy/mm/dd") & "' and Pendidikan like '%" & Trim(Txt_Cr_Nama.Text) & "%'"
        End If
        
    End With
    
    Set Grid_Status.DataSource = Rs_Nav
        Grid_Status.Refresh
    
   'TDB_Cari.Visible = False
    
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
    Txt_Nama.Enabled = True
    txt_ket.Enabled = True
    
    Txt_Nama.SetFocus
    
End Sub

Private Sub cmd_simpan_Click()
On Error GoTo err_handler

Dim konfirm As Integer
    If Txt_Kode.Text = "" Then
        konfirm = CInt(MsgBox("Kode tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))

        Txt_Kode.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
    If Txt_Nama.Text = "" Then
        konfirm = CInt(MsgBox("Jenis Type Barang tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        Txt_Nama.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
'    Dim harga As Double
'    If TDB_harga.ValueIsNull Then
'        harga = 0
'    Else
'        harga = Replace(Trim(TDB_harga.Value), ",", "")
'    End If
'
'    If harga = 0 Then
'        konfirm = CInt(MsgBox("Harga perjenis customer tidak boleh 0", vbOKOnly + vbInformation, "Informasi"))
'
'        TDB_harga.SetFocus
'        On Error GoTo 0
'        Exit Sub
'    End If
    
    Dim sql, sql1 As String
    Dim rs As Recordset
    Dim rs1 As Recordset
    
    If rubah = False Then
        
'        sql1 = "select Kode from Tb_Pendidikan where Kode='" & Trim(Txt_Kode.Text) & "'"
'
'        Set rs1 = New ADODB.Recordset
'            rs1.Open sql1, kon
'
'        With rs1
'            If Not .EOF Then
'                konfirm = CInt(MsgBox("Kode pendidikan yang anda masukkan sudah ada", vbOKOnly + vbInformation, "Informasi"))
'
'                Txt_Kode.SetFocus
'                On Error GoTo 0
'                Exit Sub
'            Else
                
                sql = "insert into Tb_Type_Brg (Kode,Jenis,Ket) values('" & Trim(Txt_Kode.Text) & "','" & Trim(Txt_Nama.Text) & "','" & Trim(txt_ket.Text) & "')"
                Set rs = New ADODB.Recordset
                    rs.Open sql, kon
                
                konfirm = CInt(MsgBox("Data telah disimpan", vbOKOnly + vbInformation, "Informasi"))
                
'
'            End If
'        End With
    
    Else
        
        sql = "update Tb_Type_Brg set Jenis='" & Trim(Txt_Nama.Text) & "',Ket='" & Trim(txt_ket.Text) & "' where Kode='" & Trim(Txt_Kode.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
            
        konfirm = CInt(MsgBox("Data telah dirubah", vbOKOnly + vbInformation, "Informasi"))
        
    End If
    
    IsiSemua
    
    
    Cmd_Batal_Click
    On Error GoTo 0
    Exit Sub
    
err_handler:
    
    konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Informaton"))
        Err.Clear
    
End Sub

Private Sub Cmd_Tambah_Click()

rubah = False
'    Dim n As Object
'        For Each n In Me
'            If TypeOf n Is TextBox Then
'                If Left(n.Name, 6) <> "Txt_Cr" Then
'                    n.Text = ""
'                End If
'            End If
'        Next
'    Set n = Nothing
    
    Txt_Kode.Enabled = True
    Txt_Kode.Text = ""
    
    
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Cari.Enabled = False
    Cmd_Keluar.Enabled = False
    
'    cmd_simpan.Enabled = False
    
    Txt_Kode.SetFocus

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
    .Left = Screen.Width / 2 - .Width / 2
    .Top = 120
End With

IsiSemua

'' akses command ''

'    hak_akses_percommand CStr(Me.Name)
'
'    cmd_tambah.Enabled = c_tambah
'    cmd_rubah.Enabled = c_rubah
'    cmd_hapus.Enabled = c_hapus

'' stop here ''

rubah = False
Txt_Kode.Enabled = False
Txt_Nama.Enabled = False
txt_ket.Enabled = False

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

Private Sub Grid_Status_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Rs_Nav.RecordCount = 0 Then
        Txt_Kode.Text = ""
        Txt_Nama.Text = ""
        txt_ket.Text = ""
        
    Else
                
        Txt_Kode.Text = Rs_Nav!kode
        Txt_Nama.Text = IIf(Not IsNull(Rs_Nav!Jenis), Rs_Nav!Jenis, "")
        
        txt_ket.Text = IIf(Not IsNull(Rs_Nav!ket), Rs_Nav!ket, "")
        
    End If
    
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


Private Sub Txt_Cr_Kode_GotFocus()
    Call Focus_(Txt_Cr_Kode)
End Sub

Private Sub Txt_Cr_Kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Cr_Nama.SetFocus
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

Private Sub txt_ket_GotFocus()
    Call Focus_(txt_ket)
End Sub

Private Sub txt_ket_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub Txt_Kode_GotFocus()
    Call Focus_(Txt_Kode)
End Sub

Private Sub Txt_Kode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
    
    Dim konfirm As Integer
        If Txt_Kode.Text = "" Then
            konfirm = CInt(MsgBox("Kode cara bayar tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
            Exit Sub
        End If
    
    
    If Rs_Nav.RecordCount = 0 Then
        
        Txt_Nama.Text = ""
        txt_ket.Text = ""
        Txt_Nama.Enabled = True
        txt_ket.Enabled = True
        
        Txt_Nama.SetFocus
        
        Exit Sub
    End If
    
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select Kode from Tb_Type_Brg where Kode='" & Trim(Txt_Kode.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        With rs
        If Not .EOF Then
            konfirm = CInt(MsgBox("Kode yang anda masukkan sudah ada", vbOKOnly + vbInformation, "Informasi"))
        Else
            
            Txt_Nama.Text = ""
            txt_ket.Text = ""
            
            Txt_Nama.Enabled = True
            txt_ket.Enabled = True
            
            Txt_Nama.SetFocus
            
        End If
        End With
    
    End If
    
End Sub


Private Sub Txt_Nama_GotFocus()
    Call Focus_(Txt_Nama)
End Sub

Private Sub Txt_Nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_ket.SetFocus
End Sub
