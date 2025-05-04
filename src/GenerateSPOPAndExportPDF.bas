Attribute VB_Name = "Module_GenerateSPOPAndExportPDF"
Sub GenerateSPOPAndExportPDF()
    Dim wsData As Worksheet
    Dim wsTemplate1 As Worksheet, wsTemplate2 As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim dataCluster As String, dataBlok As String, dataKelurahan As String, dataLuasTanah As String
    Dim dataNama As String
    Dim pdfName As String
    Dim folderPath As String
    Dim tempSheet As Worksheet
    
    ' Set worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsTemplate1 = ThisWorkbook.Sheets("SPOP (1)")
    Set wsTemplate2 = ThisWorkbook.Sheets("SPOP (2)")
    
    ' Set folder tujuan (ganti dengan path folder Anda)
    folderPath = "/Users/rifqi/Downloads/"
    If Right(folderPath, 1) <> "/" Then folderPath = folderPath & "/"
    
    ' Buat folder jika belum ada
    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
    
    ' Cari baris terakhir di sheet Data
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Loop melalui setiap baris di sheet Data
    For i = 2 To lastRow
        ' Ambil data dari tabel
        dataBlok = CStr(wsData.Cells(i, 4).Value)      ' Kolom "Blok"
        dataNama = CStr(wsData.Cells(i, 2).Value)      ' Kolom "Nama"
        dataCluster = CStr(wsData.Cells(i, 3).Value)   ' Kolom "Cluster"
        dataKelurahan = CStr(wsData.Cells(i, 7).Value) ' Kolom "Kelurahan"
        dataLuasTanah = CStr(wsData.Cells(i, 5).Value) ' Kolom "Luas Tanah"
        
        ' ===== BUAT SHEET SEMENTARA UNTUK PDF =====
        ' Copy SPOP (1) ke sheet baru
        wsTemplate1.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set tempSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        tempSheet.Name = "Temp_SPOP1"
        
        ' Isi data ke SPOP (1)
        With tempSheet
            ' Isi NAMA JALAN (Cluster) per karakter di B29, C29, dst.
            For j = 1 To Len(dataCluster)
                .Cells(29, 2 + j - 1).Value = Mid(dataCluster, j, 1)
            Next j
            
            ' Isi BLOK per karakter di AF29, AG29, dst.
            For j = 1 To Len(dataBlok)
                .Cells(29, 32 + j - 1).Value = Mid(dataBlok, j, 1)
            Next j
            
            ' Isi KELURAHAN per karakter di B33, C33, dst.
            For j = 1 To Len(dataKelurahan)
                .Cells(33, 2 + j - 1).Value = Mid(dataKelurahan, j, 1)
            Next j
            
            ' Isi LUAS TANAH per karakter di J60, K60, dst.
            For j = 1 To Len(dataLuasTanah)
                .Cells(60, 10 + j - 1).Value = Mid(dataLuasTanah, j, 1)
            Next j
        End With
        
        ' Copy SPOP (2) ke sheet baru
        wsTemplate2.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set tempSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        tempSheet.Name = "Temp_SPOP2"
        
        ' Isi NAMA di B13 (SPOP 2)
        tempSheet.Range("B13").Value = dataNama
        
        ' ===== GABUNGKAN DAN EXPORT KE PDF =====
        ' Nama file PDF berdasarkan Blok (contoh: "SPOP_BLOK-A1.pdf")
        pdfName = "SPOP_" & Replace(dataBlok, "/", "-") & ".pdf" ' Ganti "/" dengan "-" agar valid
        
        ' Export kedua sheet ke satu PDF
        Sheets(Array("Temp_SPOP1", "Temp_SPOP2")).Select
        ActiveSheet.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=folderPath & pdfName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False
        
        ' Hapus sheet sementara
        Application.DisplayAlerts = False
        Sheets("Temp_SPOP1").Delete
        Sheets("Temp_SPOP2").Delete
        Application.DisplayAlerts = True
    Next i
    
    MsgBox "Proses selesai! " & (lastRow - 1) & " file PDF telah disimpan di: " & folderPath, vbInformation
End Sub

