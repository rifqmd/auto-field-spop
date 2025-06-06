' Version: 1.4
Attribute VB_Name = "Module_SPOP_MultiSheet"
Sub SPOP_MultiSheet()
    Dim wsData As Worksheet
    Dim wsTemplate1 As Worksheet, wsTemplate2, wsTemplate3 As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim dataCluster As String, dataBlok As String, dataKelurahan As String, dataLuasTanah As String, dataLuasBangunan As String
    Dim dataNama As String, dataJumlahLantai As String
    
    ' Set worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsTemplate1 = ThisWorkbook.Sheets("SPOP (1)")
    Set wsTemplate2 = ThisWorkbook.Sheets("SPOP (2)")
    Set wsTemplate3 = ThisWorkbook.Sheets("LSPOP")
    
    ' Cari baris terakhir di sheet Data
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Loop melalui setiap baris di sheet Data
    For i = 2 To lastRow ' Mulai dari baris 2 (abaikan header)
        ' Salin template SPOP (1) ke sheet baru
        wsTemplate1.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Dim newSheet1 As Worksheet
        Set newSheet1 = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        newSheet1.Name = "SPOP1_" & i - 1 ' Nama sheet SPOP1_1, SPOP1_2, dst.
        
        ' Salin template SPOP (2) ke sheet baru
        wsTemplate2.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Dim newSheet2 As Worksheet
        Set newSheet2 = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        newSheet2.Name = "SPOP2_" & i - 1
        
        ' Salin template LSPOP ke sheet baru
        wsTemplate3.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Dim newSheet3 As Worksheet
        Set newSheet3 = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        newSheet3.Name = "LSPOP_" & i - 1
        
        ' Ambil data dari tabel
        dataNama = CStr(wsData.Cells(i, 2).Value)      ' Kolom "Nama"
        dataCluster = CStr(wsData.Cells(i, 3).Value)   ' Kolom "Cluster" (Nama Jalan)
        dataBlok = CStr(wsData.Cells(i, 4).Value)      ' Kolom "Blok"
        dataKelurahan = CStr(wsData.Cells(i, 7).Value) ' Kolom "Kelurahan"
        dataLuasTanah = CStr(wsData.Cells(i, 5).Value) ' Kolom "Luas Tanah"
        dataLuasBangunan = CStr(wsData.Cells(i, 6).Value)
        dataJumlahLantai = CStr(wsData.Cells(i, 8).Value)
        
        ' ===== ISI DATA UNTUK SPOP (1) =====
        ' Isi NAMA JALAN (Cluster) per karakter di B29, C29, D29, dst.
        For j = 1 To Len(dataCluster)
            newSheet1.Cells(29, 2 + j - 1).Value = Mid(dataCluster, j, 1)
        Next j
        
        ' Isi BLOK per karakter di AF29, AG29, AH29, dst.
        For j = 1 To Len(dataBlok)
            newSheet1.Cells(29, 32 + j - 1).Value = Mid(dataBlok, j, 1) ' AF = Kolom 32
        Next j
        
        ' Isi KELURAHAN per karakter di B33, C33, D33, dst.
        For j = 1 To Len(dataKelurahan)
            newSheet1.Cells(33, 2 + j - 1).Value = Mid(dataKelurahan, j, 1)
        Next j
        
        ' Isi LUAS TANAH per karakter di J60, K60, L60, dst.
        For j = 1 To Len(dataLuasTanah)
            newSheet1.Cells(60, 10 + j - 1).Value = Mid(dataLuasTanah, j, 1) ' J = Kolom 10
        Next j
        
        ' ===== ISI DATA UNTUK LSPOP =====
        For j = 1 To Len(dataLuasBangunan)
            newSheet3.Cells(32, 12 + j - 1).Value = Mid(dataLuasBangunan, j, 1)
        Next j
        
        For j = 1 To Len(dataJumlahLantai)
            newSheet3.Cells(32, 37 + j - 1).Value = Mid(dataJumlahLantai, j, 1)
        Next j
        
        ' === TANPA PEMISAH KARAKTER
        newSheet3.Range("AR1").Value = dataBlok
        
    Next i
    
    MsgBox "Proses selesai! " & (lastRow - 1) & " SPOP dan LSPOP telah dibuat.", vbInformation
End Sub
