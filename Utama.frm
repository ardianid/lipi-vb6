VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm Utama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "L I P I"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7275
   Icon            =   "Utama.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Utama.frx":27C92
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":2064D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":206DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":20768A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":207F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":20883E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":209118
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":2099F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":20A2CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":20ABA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "login2"
            Object.ToolTipText     =   "Login"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "logof2"
            Object.ToolTipText     =   "Log Off"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "rubahpwd_t"
            Object.ToolTipText     =   "Change Password"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PembelianBrg_T"
            Object.ToolTipText     =   "Order Barang"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BrgMasuk_T"
            Object.ToolTipText     =   "Barang Masuk"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BrgKeluar_T"
            Object.ToolTipText     =   "Barang Keluar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5670
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "27/08/2008"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "8:50"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu fL 
      Caption         =   "&File"
      Begin VB.Menu login 
         Caption         =   "&Login"
      End
      Begin VB.Menu logof 
         Caption         =   "Log &Off"
      End
      Begin VB.Menu grs1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu user 
      Caption         =   "&User"
      Begin VB.Menu User_Baru_M 
         Caption         =   "&Tambah User"
      End
      Begin VB.Menu Form_Hak_Akses_M 
         Caption         =   "&Seting Hak Akses"
      End
      Begin VB.Menu grspw 
         Caption         =   "-"
      End
      Begin VB.Menu rubahpwd_M 
         Caption         =   "&Rubah Password"
      End
   End
   Begin VB.Menu mast 
      Caption         =   "&Master"
      Begin VB.Menu Karyawan_M 
         Caption         =   "&Karyawan"
      End
      Begin VB.Menu Frm_Mast_Type_Brg_M 
         Caption         =   "&Type Barang"
      End
      Begin VB.Menu Frm_Brg_M 
         Caption         =   "&Barang"
      End
   End
   Begin VB.Menu trans 
      Caption         =   "&Transaksi"
      Begin VB.Menu PembelianBrg_M 
         Caption         =   "&Order Barang"
      End
      Begin VB.Menu pemisahtrans 
         Caption         =   "-"
      End
      Begin VB.Menu BrgMasuk_M 
         Caption         =   "&Barang Masuk"
      End
      Begin VB.Menu BrgKeluar_M 
         Caption         =   "Barang &Keluar"
      End
   End
   Begin VB.Menu lap 
      Caption         =   "&Laporan"
      Begin VB.Menu Frm_sel_Karyawan_M 
         Caption         =   "&Karyawan"
      End
      Begin VB.Menu frm_sel_stock_M 
         Caption         =   "&Stock Barang"
      End
      Begin VB.Menu frm_sel_orderbrg_M 
         Caption         =   "&Order Barang"
      End
      Begin VB.Menu frm_sel_brg_masuk_M 
         Caption         =   "Barang &Masuk"
      End
      Begin VB.Menu frm_sel_brg_keluar_M 
         Caption         =   "Barang &Keluar"
      End
      Begin VB.Menu frm_sel_mutasi_M 
         Caption         =   "Mu&tasi Barang"
      End
   End
End
Attribute VB_Name = "Utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim status As String

Public Sub SetAktifMenu(ByVal sql As String)
    
    Dim obj As Object
     Dim a As Long
            
    Dim rec As Recordset
        Set rec = New ADODB.Recordset
            rec.Open sql, kon, adOpenKeyset

    With rec
        If Not .EOF Then
        Do While Not .EOF

               Dim nama_f
               Dim namatol
                    nama_f = !nama_form
                    namatol = nama_f
                    nama_f = nama_f & "_M"
                    namatol = namatol & "_T"
                    
               For Each obj In Me
               
                If TypeOf obj Is Toolbar Then
                Else
                If obj.Name = nama_f Then
                    obj.Enabled = True
                    Exit For
                End If
                End If
                
               Next
                
               
               For a = 1 To 10
                    If UCase(Toolbar1.Buttons.Item(a).Key) = UCase(namatol) Then
                        Toolbar1.Buttons.Item(a).Enabled = True
                        Exit For
                    End If
               Next
                
        .MoveNext
        Loop
        End If

    End With

    rubahpwd_M.Enabled = True
    Toolbar1.Buttons.Item(2).Enabled = True
    Toolbar1.Buttons.Item(1).Enabled = False
    Toolbar1.Buttons.Item(4).Enabled = True
    
End Sub

Private Sub BrgKeluar_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("BrgKeluar") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = BrgKeluar
        Frm.Show
        
    Else
        
        If Cek_akses_Form("BrgKeluar") = False Then Exit Sub
        
        Set Frm = BrgKeluar
        Frm.Show
    End If

End Sub

Private Sub BrgMasuk_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("BrgMasuk") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = BrgMasuk
        Frm.Show
        
    Else
        
        If Cek_akses_Form("BrgMasuk") = False Then Exit Sub
        
        Set Frm = BrgMasuk
        Frm.Show
    End If

End Sub

Private Sub exit_Click()
    End
End Sub

Public Sub enable_menu(ByVal sett As Boolean)
   
   Dim a As Object
   Dim x As Long
        For Each a In Me
        
            If TypeOf a Is Toolbar Then
            Else
            If (UCase(Right(a.Name, 1)) = UCase("M")) Then
                a.Enabled = sett
            End If
            End If
            
        Next
   
        For x = 1 To 10
            If UCase(Right(Toolbar1.Buttons.Item(x).Key, 1)) = UCase("t") Then
                 Toolbar1.Buttons.Item(x).Enabled = sett
            End If
        Next
        
'   adduser_S.Enabled = sett
'   setingakses_S.Enabled = sett
'   rubahpwd_S.Enabled = sett
'
'   kary_S.Enabled = sett
'   anggota_S.Enabled = sett
'   hargaperkilo_S.Enabled = sett
'   simpananwajib_S.Enabled = sett
'
'   timbang_S.Enabled = sett
'   timbang_btl.Enabled = sett
'
'
'   lapkary_S.Enabled = sett
'   lapanggota_S.Enabled = sett
'   laptimbang.Enabled = sett
    
End Sub


Private Sub Form_Hak_Akses_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Form_Hak_Akses") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Form_Hak_Akses
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Form_Hak_Akses") = False Then Exit Sub
        
        Set Frm = Form_Hak_Akses
        Frm.Show
    End If

End Sub

Private Sub Frm_Brg_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_Brg") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_Brg
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_Brg") = False Then Exit Sub
        
        Set Frm = Frm_Brg
        Frm.Show
    End If


End Sub

Private Sub Frm_Mast_Type_Brg_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_Mast_Type_Brg") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_Mast_Type_Brg
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_Mast_Type_Brg") = False Then Exit Sub
        
        Set Frm = Frm_Mast_Type_Brg
        Frm.Show
    End If


End Sub

Private Sub Frm_Peny_Stock_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_Peny_Stock") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_Peny_Stock
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_Peny_Stock") = False Then Exit Sub
        
        Set Frm = Frm_Peny_Stock
        Frm.Show
    End If


End Sub

Private Sub frm_sel_brg_keluar_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_brg_keluar") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_brg_keluar
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_brg_keluar") = False Then Exit Sub
        
        Set Frm = frm_sel_brg_keluar
        Frm.Show
    End If

End Sub

Private Sub frm_sel_brg_masuk_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_brg_masuk") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_brg_masuk
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_brg_masuk") = False Then Exit Sub
        
        Set Frm = frm_sel_brg_masuk
        Frm.Show
    End If

End Sub

Private Sub Frm_sel_Karyawan_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_sel_Karyawan") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_sel_Karyawan
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_sel_Karyawan") = False Then Exit Sub
        
        Set Frm = Frm_sel_Karyawan
        Frm.Show
    End If


End Sub

Private Sub frm_sel_mutasi_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_mutasi") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_mutasi
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_mutasi") = False Then Exit Sub
        
        Set Frm = frm_sel_mutasi
        Frm.Show
    End If


End Sub

Private Sub frm_sel_orderbrg_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_orderbrg") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_orderbrg
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_orderbrg") = False Then Exit Sub
        
        Set Frm = frm_sel_orderbrg
        Frm.Show
    End If


End Sub

Private Sub frm_sel_stock_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_stock") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_stock
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_stock") = False Then Exit Sub
        
        Set Frm = frm_sel_stock
        Frm.Show
    End If


End Sub

Private Sub Frm_Stock_Awal_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_Stock_Awal") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_Stock_Awal
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_Stock_Awal") = False Then Exit Sub
        
        Set Frm = Frm_Stock_Awal
        Frm.Show
    End If


End Sub

Private Sub Karyawan_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Karyawan") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Karyawan
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Karyawan") = False Then Exit Sub
        
        Set Frm = Karyawan
        Frm.Show
    End If


End Sub

Private Sub login_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
    End If
    
    enable_menu False
    StatusBar1.Panels(1).Text = "User Actived :"
    U_Masuk.Show

End Sub

Private Sub logof_Click()
    
    If kon.State = adStateClosed Then
            Buka_Koneksi
    End If
    
    If Not (Frm Is Nothing) Then
        Unload Frm
    End If
    
    enable_menu False
    StatusBar1.Panels(1).Text = "User Actived :"
    U_Masuk.Show
    
End Sub

Private Sub MDIForm_Load()
    
    enable_menu False

 status = Buka_Koneksi
 If status = "-2147467259" Then
    
            Dim konfirm As Integer
            Dim Informasi As String
                Informasi = "Koneksi terhadap server tidak berhasil :"
                Informasi = Informasi & vbCrLf & "1. Pastikan server telah hidup dan SQL Server telah dijalankan pada server,atau"
                Informasi = Informasi & vbCrLf & "2. Apabila masih terjadi kegagalan koneksi,periksa nama komputer server,Pastikan nama komputer server tidak berubah"
                Informasi = Informasi & vbCrLf & vbCrLf & "apakah anda ingin menyeting ulang koneksi nama komputer server ?"
        
                konfirm = CInt(MsgBox(Informasi, vbYesNo + vbQuestion, "Konfimasi"))
        
                If konfirm = vbYes Then
        
                    Load Frm_New_Seting
                    Frm_New_Seting.Show
        
                    Unload Me
                    Exit Sub
                Else
                    Unload Me
                    End
                    Exit Sub
                End If

 End If

'    Dim btas As Double
'        btas = Batas
'
'    If btas = 100 Then
'        End
'        Exit Sub
'    Else
'        btas = btas + 1
'        SaveSetting "bts", "bts", "bts", btas
'    End If


    logof.Enabled = False
    Toolbar1.Buttons.Item(2).Enabled = False
    
End Sub



Private Sub PembelianBrg_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("PembelianBrg") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = PembelianBrg
        Frm.Show
        
    Else
        
        If Cek_akses_Form("PembelianBrg") = False Then Exit Sub
        
        Set Frm = PembelianBrg
        Frm.Show
    End If


End Sub

Private Sub rubahpwd_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
                
        Set Frm = Nothing
        Set Frm = Frm_Rubah_Pwd
        Frm.Show
        
    Else
                
        Set Frm = Frm_Rubah_Pwd
        Frm.Show
    End If


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1
            login_Click
        Case 2
            logof_Click
        Case 4
            rubahpwd_M_Click
        Case 6
            PembelianBrg_M_Click
        Case 8
            BrgMasuk_M_Click
        Case 9
            BrgKeluar_M_Click
        Case 11
            exit_Click
    End Select
    
End Sub

Private Sub User_Baru_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("User_Baru") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = User_Baru
        Frm.Show
        
    Else
        
        If Cek_akses_Form("User_Baru") = False Then Exit Sub
        
        Set Frm = User_Baru
        Frm.Show
    End If
    
End Sub
