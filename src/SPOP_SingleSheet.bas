Attribute VB_Name = "Module_SPOP_SingleSheet"
Sub SPOP_SingleSheet()
    Dim wsData As Worksheet
    Dim wsTemplate As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim dataCluster As String, dataBlok As String, dataKelurahan As String, dataLuasTanah As String
    
    ' Set worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsTemplate = ThisWorkbook.Sheets("SPOP (1)")
    
    ' Cari baris terakhir di sheet Data
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Loop melalui setiap baris di sheet Data
    For i = 2 To lastRow ' Mulai dari baris 2 (abaikan header)
        ' Salin template ke sheet baru
        wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Dim newSheet As Worksheet
        Set newSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        newSheet.Name = "SPOP_" & i - 1 ' Nama sheet SPOP_1, SPOP_2, dst.
        
        ' Ambil data dari tabel
        dataCluster = CStr(wsData.Cells(i, 3).Value)  ' Kolom "Cluster" (Nama Jalan)
        dataBlok = CStr(wsData.Cells(i, 4).Value)     ' Kolom "Blok"
        dataKelurahan = CStr(wsData.Cells(i, 7).Value) ' Kolom "Kelurahan"
        dataLuasTanah = CStr(wsData.Cells(i, 5).Value) ' Kolom "Luas Tanah"
        
        ' Isi NAMA JALAN (Cluster) per karakter di B29, C29, D29, dst.
        For j = 1 To Len(dataCluster)
            newSheet.Cells(29, 2 + j - 1).Value = Mid(dataCluster, j, 1)
        Next j
        
        ' Isi BLOK per karakter di AF29, AG29, AH29, dst.
        For j = 1 To Len(dataBlok)
            newSheet.Cells(29, 32 + j - 1).Value = Mid(dataBlok, j, 1) ' AF = Kolom 32
        Next j
        
        ' Isi KELURAHAN per karakter di B33, C33, D33, dst.
        For j = 1 To Len(dataKelurahan)
            newSheet.Cells(33, 2 + j - 1).Value = Mid(dataKelurahan, j, 1)
        Next j
        
        ' Isi LUAS TANAH per karakter di J60, K60, L60, dst.
        For j = 1 To Len(dataLuasTanah)
            newSheet.Cells(60, 10 + j - 1).Value = Mid(dataLuasTanah, j, 1) ' J = Kolom 10
        Next j
    Next i
    
    MsgBox "Proses selesai! " & (lastRow - 1) & " SPOP telah dibuat.", vbInformation
End Sub
