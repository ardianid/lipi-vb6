VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} r_mutasi_brg 
   ClientHeight    =   9600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   OleObjectBlob   =   "r_mutasi_brg.dsx":0000
End
Attribute VB_Name = "r_mutasi_brg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Report_Initialize()
    Text20.SetText 0
    Text21.SetText 0
    Text22.SetText 0
    Text23.SetText 0
End Sub

Private Sub Section11_Format(ByVal pFormattingInfo As Object)
    
    If Field2.Value = Empty Then Exit Sub
    
    Dim jml_awal As Double
    Dim nil_awal As Double
        jml_awal = 0
        nil_awal = 0
        
    Dim comd As Command
        Set comd = New ADODB.Command
        With comd
            .ActiveConnection = kon
            .CommandText = "Quant_Sbl"
            .CommandType = adCmdStoredProc
            .Parameters("@tgl").Value = macem2
            .Parameters("@kode").Value = Field2.Value
            .Execute
            
            If Not IsNull(.Parameters("@jml")) Then
                jml_awal = .Parameters("@jml")
                Text20.SetText (Format(.Parameters("@jml"), "###,###,###"))
                
            Else
                Text20.SetText 0
            End If
            
        End With
    
    Set comd.ActiveConnection = Nothing
    
    Dim comd1 As Command
        Set comd1 = New ADODB.Command
        With comd1
            .ActiveConnection = kon
            .CommandText = "Quant_Sbl_Nilai"
            .CommandType = adCmdStoredProc
            .Parameters("@tgl").Value = macem2
            .Parameters("@kode").Value = Field2.Value
            .Execute
            
            If Not IsNull(.Parameters("@jml")) Then
                nil_awal = .Parameters("@jml")
                Text21.SetText (Format(.Parameters("@jml"), "###,###,###"))
            Else
                Text21.SetText 0
            End If
            
        End With
    
    Set comd1.ActiveConnection = Nothing
    
    Dim jml As Double
        If Field8.Value = Empty Then
            jml = 0
        Else
            jml = Field8.Value
        End If
        
    Dim nilai As Double
        If Field9.Value = Empty Then
            nilai = 0
        Else
            nilai = Field9.Value
        End If
    
    Dim jml_akhir As Double
    Dim nil_akhir As Double
        
        jml_akhir = jml_awal + jml
        nil_akhir = nil_awal + nilai
    
    If jml_akhir = 0 Then
        Text22.SetText 0
    Else
        Text22.SetText Format(jml_akhir, "###,###,###")
    End If
        
    If nil_akhir = 0 Then
        Text23.SetText 0
    Else
        Text23.SetText Format(nil_akhir, "###,###,###")
    End If
        
End Sub

Private Sub Section6_Format(ByVal pFormattingInfo As Object)
    t1.SetText (macem2)
    t2.SetText (macem2_lagi)
End Sub
