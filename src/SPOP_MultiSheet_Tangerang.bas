' SPOP_MultiSheet_Tangerang.bas
' This code is part of a VBA project to generate multiple SPOP and LSPOP sheets

' Version: 1.5
' Author: Rifqi Muhammad
' License: GPL-3.0

Sub SPOP_Tangerang()
    Dim wsData As Worksheet
    Dim wsTemplate1 As Worksheet, wsTemplate2, wsTemplate3, wsTemplate4 As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim dataCluster As String, dataBlok As String, dataKelurahan As String, dataLuasTanah As String, dataLuasBangunan As String
    Dim dataNama As String, dataJumlahLantai As String
    
    ' Set worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsTemplate1 = ThisWorkbook.Sheets("SPOP (1)")
    Set wsTemplate2 = ThisWorkbook.Sheets("SPOP (2)")
    Set wsTemplate3 = ThisWorkbook.Sheets("LSPOP (1)")
    Set wsTemplate4 = ThisWorkbook.Sheets("LSPOP (2)")
    
    ' Cari baris terakhir di sheet Data
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Loop melalui setiap baris di sheet Data
    For i = 2 To lastRow ' Mulai dari baris 2 (abaikan header)
        wsTemplate1.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Dim newSheet1 As Worksheet
        Set newSheet1 = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        newSheet1.Name = "SPOP1_" & i - 1
        
        wsTemplate2.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Dim newSheet2 As Worksheet
        Set newSheet2 = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        newSheet2.Name = "SPOP2_" & i - 1
        
        wsTemplate3.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Dim newSheet3 As Worksheet
        Set newSheet3 = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        newSheet3.Name = "LSPOP1_" & i - 1
        
        wsTemplate4.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Dim newSheet4 As Worksheet
        Set newSheet4 = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        newSheet3.Name = "LSPOP2_" & i - 1
        
        ' Ambil data dari tabel
        dataNama = CStr(wsData.Cells(i, 2).Value)
        dataCluster = CStr(wsData.Cells(i, 3).Value)
        dataBlok = CStr(wsData.Cells(i, 4).Value)
        dataKelurahan = CStr(wsData.Cells(i, 7).Value)
        dataLuasTanah = CStr(wsData.Cells(i, 5).Value)
        dataLuasBangunan = CStr(wsData.Cells(i, 6).Value)
        dataJumlahLantai = CStr(wsData.Cells(i, 8).Value)
        
        ' ===== ISI DATA UNTUK SPOP (1) =====
        ' Isi NAMA JALAN (Cluster) per karakter di B29, C29, D29, dst.
        For j = 1 To Len(dataCluster)
            newSheet1.Cells(31, 2 + j - 1).Value = Mid(dataCluster, j, 1)
        Next j
        
        ' Isi BLOK per karakter di AF29, AG29, AH29, dst.
        For j = 1 To Len(dataBlok)
            newSheet1.Cells(31, 32 + j - 1).Value = Mid(dataBlok, j, 1)
        Next j
        
        ' Isi KELURAHAN per karakter di B33, C33, D33, dst.
        For j = 1 To Len(dataKelurahan)
            newSheet1.Cells(36, 2 + j - 1).Value = Mid(dataKelurahan, j, 1)
        Next j
        
        ' Isi LUAS TANAH dari belakang (kanan ke kiri)
        digitLength = Len(dataLuasTanah)
        startCol = 10 + 8 ' Kolom J (10) + 8 kolom ke kanan (anggap maks 5 digit)
        
        For j = 1 To digitLength
            newSheet1.Cells(69, startCol - j + 1).Value = Mid(dataLuasTanah, digitLength - j + 1, 1)
        Next j
        
        ' ===== ISI DATA UNTUK LSPOP =====
        ' Isi LUAS BANGUNAN dari belakang (kanan ke kiri)
        digitLength = Len(dataLuasBangunan)
        startCol = 12 + 9
        
        For j = 1 To digitLength
            newSheet3.Cells(30, startCol - j + 1).Value = Mid(dataLuasBangunan, digitLength - j + 1, 1)
        Next j
        
        ' Isi JUMLAH LANTAI dari belakang (kanan ke kiri)
        digitLength = Len(dataJumlahLantai)
        startCol = 37 + 1
        
        For j = 1 To digitLength
            newSheet3.Cells(30, startCol - j + 1).Value = Mid(dataJumlahLantai, digitLength - j + 1, 1)
        Next j
        
        ' === TANPA PEMISAH KARAKTER
        newSheet3.Range("AR1").Value = dataBlok
        
    Next i
    
    MsgBox "Proses selesai! " & (lastRow - 1) & " SPOP dan LSPOP telah dibuat.", vbInformation
End Sub
