Attribute VB_Name = "Module_SPOP_SingleSheet"
Sub SPOP_SingleSheet()
    Dim wsData As Worksheet
    Dim wsTemplate1 As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim alamatOP As String, dataBlok As String, dataKelurahan As String, dataLuasTanah As String
    
    ' Set worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsTemplate1 = ThisWorkbook.Sheets("SPOP (1)")
    
    ' Cari baris terakhir di sheet Data
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Loop melalui setiap baris di sheet Data
    For i = 2 To lastRow ' Mulai dari baris 2 (abaikan header)
        ' Salin template SPOP (1) ke sheet baru
        wsTemplate1.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Dim newSheet1 As Worksheet
        Set newSheet1 = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        newSheet1.Name = "SPOP1_" & i - 1 ' Nama sheet SPOP1_1, SPOP1_2, dst.
        
        ' Ambil data dari tabel
        dataNama = CStr(wsData.Cells(i, 2).Value)
        alamatOP = CStr(wsData.Cells(i, 3).Value)
        dataBlok = CStr(wsData.Cells(i, 4).Value)
        dataKelurahan = CStr(wsData.Cells(i, 7).Value)
        dataLuasTanah = CStr(wsData.Cells(i, 5).Value)
        
        ' ===== ISI DATA UNTUK SPOP (1) =====
        For j = 1 To Len(alamatOP)
            newSheet1.Cells(29, 2 + j - 1).Value = Mid(alamatOP, j, 1)
        Next j
        
        For j = 1 To Len(dataBlok)
            newSheet1.Cells(29, 32 + j - 1).Value = Mid(dataBlok, j, 1)
        Next j
        
        For j = 1 To Len(dataKelurahan)
            newSheet1.Cells(33, 2 + j - 1).Value = Mid(dataKelurahan, j, 1)
        Next j
        
        For j = 1 To Len(dataLuasTanah)
            newSheet1.Cells(60, 10 + j - 1).Value = Mid(dataLuasTanah, j, 1) ' J = Kolom 10
        Next j
        
    Next i
    
    MsgBox "Proses selesai! " & (lastRow - 1) & " SPOP telah dibuat.", vbInformation
End Sub
