Attribute VB_Name = "OnlyCreateDrafts"
Option Explicit

' ================================================================
' TEST MODU: True -> Ýlk 5 geçerli taslaðý oluþturur
' ================================================================
Private Const TEST_MODE As Boolean = FALSE
Private Const TEST_MAX_ORDERS As Long = 5

' --- SABÝTLER ---
Private Const CONFIG_SHEET As String = "Ayarlar"
Private Const LOG_SHEET As String = "Log"
Private Const REPORT_SHEET As String = "MailRaporu"
Private Const MISSING_SHEET As String = "EksikMail"
Private Const HEADER_ROW As Long = 1
Private Const DEFAULT_START_ROW As Long = 2
Private Const DRAFT_MARK As String = "+"
Private Const NOT_FOUND_MARK As String = "MAIL BULUNAMADI"
Private Const CC_RECIPIENTS As String = ""

' --- VARSAYILAN SÜTUNLAR (Ayarlar sayfasýna göre güncellendi) ---
Private Const DEF_COL_ORDER_NO As String = "F"
Private Const DEF_COL_KALEM_NO As String = "G"
Private Const DEF_COL_VENDOR As String = "L"
Private Const DEF_COL_MALZEME As String = "D"
Private Const DEF_COL_URETIM_YERI As String = "B"
Private Const DEF_COL_UZUN_METIN As String = "E"
Private Const DEF_COL_BIRIM As String = "I"
Private Const DEF_COL_BAKIYE As String = "H"
Private Const DEF_COL_STATUS As String = "N"
Private Const DEF_COL_REPLY As String = "O"
Private Const DEF_COL_REPLY_DATE As String = "P"
Private Const DEF_COL_REPLY_FROM As String = "Q"
Private Const DEF_COL_MATCH_METHOD As String = "R"
Private Const DEF_COL_ATTACH As String = "S"
Private Const DEF_COL_VENDOR_NAME As String = "A"
Private Const DEF_COL_VENDOR_EMAIL As String = "B"

Private gCfg As Object
Private gLogWs As Worksheet
Private gLogRow As Long
Private gReportWs As Worksheet
Private gMissingWs As Worksheet
Private gSignatureHtml As String
Private gSignatureFetched As Boolean

' ================================================================
' BOÞ HÜCRE OKUMA KORUMASI
' ================================================================
Private Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Then
        SafeStr = ""
    ElseIf IsError(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim$(CStr(v))
    End If
End Function

Private Function T(ByVal s As String) As String
    Dim r As String: r = s
    r = Replace(r, "~c", ChrW(231)): r = Replace(r, "~C", ChrW(199))
    r = Replace(r, "~g", ChrW(287)): r = Replace(r, "~G", ChrW(286))
    r = Replace(r, "~i", ChrW(305)): r = Replace(r, "~I", ChrW(304))
    r = Replace(r, "~o", ChrW(246)): r = Replace(r, "~O", ChrW(214))
    r = Replace(r, "~s", ChrW(351)): r = Replace(r, "~S", ChrW(350))
    r = Replace(r, "~u", ChrW(252)): r = Replace(r, "~U", ChrW(220))
    T = r
End Function

Private Function IsTestMode() As Boolean
    Dim ws As Worksheet: On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET): On Error GoTo 0
    If ws Is Nothing Then IsTestMode = TEST_MODE: Exit Function
    Dim v As String: v = UCase$(SafeStr(ws.Range("B26").Value))
    If v = "EVET" Or v = "TRUE" Or v = "1" Then
        IsTestMode = True
    ElseIf v = "HAYIR" Or v = "FALSE" Or v = "0" Then
        IsTestMode = False
    Else
        IsTestMode = TEST_MODE
    End If
End Function

Public Sub TaslaklariOlustur()
    Dim cfg As Object
    On Error GoTo FatalError
    Set cfg = GetConfig(): cfg.RemoveAll
    If Not LoadOrCollectSettings(cfg) Then Exit Sub
    DoCreateDrafts cfg
    Exit Sub
FatalError:
    MsgBox "Kritik Hata (Satýr " & Erl & "):" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Lütfen Excel sayfalarýnýzýn ve sütun baþlýklarýnýzýn doðru olduðundan emin olun.", vbCritical, T("Açýk Sipariþ Mail Otomasyonu")
    WriteLog "FATAL", "", "", "", "TaslaklariOlustur", Err.Description
End Sub

Private Function LoadOrCollectSettings(ByRef cfg As Object) As Boolean
    Dim answer As VbMsgBoxResult
    If LoadSettingsFromSheet(cfg) Then
        answer = MsgBox(T("Kayýtlý ayarlar yüklendi:") & vbCrLf & vbCrLf & _
                        GetSettingsSummary(cfg) & vbCrLf & vbCrLf & _
                        T("Bu ayarlarla devam edilsin mi?"), _
                        vbYesNoCancel + vbQuestion, T("Açýk Sipariþ Mail Otomasyonu"))
        If answer = vbCancel Then Exit Function
        If answer = vbYes Then LoadOrCollectSettings = True: Exit Function
    End If
    cfg.RemoveAll
    If Not CollectSettings(cfg) Then Exit Function
    SaveSettingsToSheet cfg
    LoadOrCollectSettings = True
End Function

Private Function LoadSettingsFromSheet(ByRef cfg As Object) As Boolean
    Dim ws As Worksheet, s As String
    Dim orderSheets As Collection, wsOrder As Worksheet, wsVendor As Worksheet
    Dim i As Long, parts() As String, sname As String, vName As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    On Error GoTo CleanFail

    s = SafeStr(ws.Range("B1").Value)
    If Len(s) = 0 Then Exit Function
    parts = Split(s, ",")
    Set orderSheets = New Collection
    For i = LBound(parts) To UBound(parts)
        sname = Trim$(parts(i))
        If Len(sname) > 0 Then
            Set wsOrder = Nothing
            On Error Resume Next
            Set wsOrder = ThisWorkbook.Worksheets(sname)
            On Error GoTo CleanFail
            If Not wsOrder Is Nothing Then orderSheets.Add wsOrder
        End If
    Next i
    If orderSheets.Count = 0 Then Exit Function

    vName = SafeStr(ws.Range("B2").Value)
    If Len(vName) = 0 Then Exit Function
    On Error Resume Next
    Set wsVendor = ThisWorkbook.Worksheets(vName)
    On Error GoTo CleanFail
    If wsVendor Is Nothing Then Exit Function

    Set cfg("OrderSheets") = orderSheets
    Set cfg("VendorWs") = wsVendor
    cfg("StartRow") = CLng(Val(ws.Range("B3").Value))
    If cfg("StartRow") < 2 Then cfg("StartRow") = DEFAULT_START_ROW
    cfg("InboxFolder") = SafeStr(ws.Range("B4").Value)
    
    ' AYARLAR SAYFASINDAN SÜTUN HARFLERÝNÝ OKU (DÜZELTME BURADA)
    On Error Resume Next
    cfg("OrderNoCol") = ColLetterToNum(SafeStr(ws.Range("B7").Value))
    cfg("KalemNoCol") = ColLetterToNum(SafeStr(ws.Range("B8").Value))
    cfg("VendorCol") = ColLetterToNum(SafeStr(ws.Range("B9").Value))
    cfg("MalzemeCol") = ColLetterToNum(SafeStr(ws.Range("B10").Value))
    cfg("UretimYeriCol") = ColLetterToNum(SafeStr(ws.Range("B11").Value))
    cfg("UzunMetinCol") = ColLetterToNum(SafeStr(ws.Range("B12").Value))
    cfg("BirimCol") = ColLetterToNum(SafeStr(ws.Range("B13").Value))
    cfg("BakiyeCol") = ColLetterToNum(SafeStr(ws.Range("B14").Value))
    cfg("StatusCol") = ColLetterToNum(SafeStr(ws.Range("B15").Value))
    cfg("ReplyCol") = ColLetterToNum(SafeStr(ws.Range("B16").Value))
    cfg("ReplyDateCol") = ColLetterToNum(SafeStr(ws.Range("B17").Value))
    cfg("ReplyFromCol") = ColLetterToNum(SafeStr(ws.Range("B18").Value))
    cfg("MatchMethodCol") = ColLetterToNum(SafeStr(ws.Range("B19").Value))
    cfg("AttachCol") = ColLetterToNum(SafeStr(ws.Range("B20").Value))
    cfg("VendorNameCol") = ColLetterToNum(SafeStr(ws.Range("B23").Value))
    cfg("VendorEmailCol") = ColLetterToNum(SafeStr(ws.Range("B24").Value))
    On Error GoTo CleanFail
    
    ' AutoFix ile Excel'in gerçek baþlýklarýný öðren ve yanlýþ harfleri düzelt!
    AutoFixAllColumns cfg, orderSheets(1), wsVendor
    LoadSettingsFromSheet = True
    Exit Function
CleanFail:
    LoadSettingsFromSheet = False
End Function

Private Sub AutoFixAllColumns(ByRef cfg As Object, ByVal ws As Worksheet, ByVal vws As Worksheet)
    Dim c As Long, r As Long, cv As String
    Dim emailCol As Long, nameCol As Long, lastC As Long
    Dim res As Long
    
    ' Sipariþ Sayfasý Sütunlarýný Baþlýklardan Bul (DETAY_UY vb. eklendi)
    res = FindColumnByHeaders(ws, Array("Satýnalma belgesi", "Satin alma", "Sipariþ No", "Siparis No", "Sipariþ Numarasý", "Order No", "SIPARIS_NO"), HEADER_ROW)
    If res > 0 Then cfg("OrderNoCol") = res
    
    res = FindColumnByHeaders(ws, Array("Kalem", "Kalem No", "KalemNo", "Item No", "Kalem Numarasý", "SAT_KALEM_NO"), HEADER_ROW)
    If res > 0 Then cfg("KalemNoCol") = res
    
    res = FindColumnByHeaders(ws, Array("Ad 1", "Ad1", "Tedarikçi Ünvaný", "Satýcý Ýsmi", "Firma Adý", "Tedarikçi Adý", "Müþteri Adý", "TEDARIKCI_ADI"), HEADER_ROW)
    If res > 0 Then cfg("VendorCol") = res
    
    res = FindColumnByHeaders(ws, Array("Malzeme", "Malzeme Kodu", "Material", "Stok Kodu", "Ürün Kodu", "DETAY MALZEME", "DETAY_MALZEME"), HEADER_ROW)
    If res > 0 Then cfg("MalzemeCol") = res
    
    res = FindColumnByHeaders(ws, Array("Üretim yeri", "Uretim yeri", "Üretim Yeri", "UretimYeri", "DETAY_UY", "DETAY UY"), HEADER_ROW)
    If res > 0 Then cfg("UretimYeriCol") = res
    
    res = FindColumnByHeaders(ws, Array("Malzeme Uzun Metni", "Uzun Metin", "Malzeme Tanýmý", "Taným", "DETAY_MALZ_METNI", "DETAY MALZ_METNI"), HEADER_ROW)
    If res > 0 Then cfg("UzunMetinCol") = res
    
    res = FindColumnByHeaders(ws, Array("Temel ölçü birimi", "Temel Olcu Birimi", "Birim", "Ölçü Birimi", "TEMEL_OLCU_BIRIM"), HEADER_ROW)
    If res > 0 Then cfg("BirimCol") = res
    
    res = FindColumnByHeaders(ws, Array("Bakiye MG Miktarý", "Bakiye Miktar", "Bakiye", "Kalan Miktar", "Miktar", "DETAY_MIKTAR"), HEADER_ROW)
    If res > 0 Then cfg("BakiyeCol") = res
    
    res = FindColumnByHeaders(ws, Array("Mail Atýldý", "Mail Atildi", "Mail Durumu", "Durum", "Status"), HEADER_ROW)
    If res > 0 Then cfg("StatusCol") = res
    
    res = FindColumnByHeaders(ws, Array("Gelen Yanýt", "Gelen Yanit", "Yanýt", "Cevap", "Reply"), HEADER_ROW)
    If res > 0 Then cfg("ReplyCol") = res
    
    res = FindColumnByHeaders(ws, Array("Yanýt Tarihi", "Yanit Tarihi", "Cevap Tarihi"), HEADER_ROW)
    If res > 0 Then cfg("ReplyDateCol") = res
    
    res = FindColumnByHeaders(ws, Array("Yanýtlayan E-posta", "Yanýtlayan", "Yanitlayan", "Reply From"), HEADER_ROW)
    If res > 0 Then cfg("ReplyFromCol") = res
    
    res = FindColumnByHeaders(ws, Array("Eþleþme Yöntemi", "Eslesme", "Eþleþme", "Match Method"), HEADER_ROW)
    If res > 0 Then cfg("MatchMethodCol") = res
    
    res = FindColumnByHeaders(ws, Array("Ekler", "Ek", "Dosyalar", "Attachments"), HEADER_ROW)
    If res > 0 Then cfg("AttachCol") = res

    ' Tedarikçi Sayfasý Sütunlarýný Baþlýklardan Bul
    res = FindColumnByHeaders(vws, Array("Ad 1", "Ad1", "Satýcý ismi", "Satici ismi", "Tedarikçi", "Tedarikci Adý", "Firma Adý", "Firma", "Ünvan", "Unvan", "Ýsim", "Satýcý Unvaný"), HEADER_ROW)
    If res > 0 Then cfg("VendorNameCol") = res
    
    res = FindColumnByHeaders(vws, Array("ana e-mail adresi", "Ana E-mail", "E-posta", "Eposta", "Email", "Mail", "E-Mail Adresi", "Mail Adresi", "ana e-mail"), HEADER_ROW)
    If res > 0 Then cfg("VendorEmailCol") = res

    ' KURTARMA MODU: Eðer Tedarikçi sayfasýnda baþlýklar bulunamadýysa hücrelerin içine bak
    If cfg("VendorNameCol") < 1 Or cfg("VendorEmailCol") < 1 Then
        lastC = vws.Cells(HEADER_ROW, vws.Columns.Count).End(xlToLeft).Column
        
        For c = 1 To lastC
            For r = 2 To 100
                cv = SafeStr(vws.Cells(r, c).Value)
                If InStr(cv, "@") > 0 And InStr(cv, ".") > 0 Then
                    emailCol = c: Exit For
                End If
            Next r
            If emailCol > 0 Then Exit For
        Next c
        
        If emailCol > 0 Then
            For c = 1 To emailCol - 1
                For r = 2 To 100
                    cv = SafeStr(vws.Cells(r, c).Value)
                    If Len(cv) > 3 And InStr(cv, "@") = 0 And Not IsNumeric(cv) Then
                        nameCol = c: Exit For
                    End If
                Next r
                If nameCol > 0 Then Exit For
            Next c
        End If
        
        If nameCol > 0 Then cfg("VendorNameCol") = nameCol
        If emailCol > 0 Then cfg("VendorEmailCol") = emailCol
    End If

    ' Varsayýlanlarý ata (Eðer hala bulunamadýysa)
    If cfg("OrderNoCol") < 1 Then cfg("OrderNoCol") = ColLetterToNum(DEF_COL_ORDER_NO)
    If cfg("KalemNoCol") < 1 Then cfg("KalemNoCol") = ColLetterToNum(DEF_COL_KALEM_NO)
    If cfg("VendorCol") < 1 Then cfg("VendorCol") = ColLetterToNum(DEF_COL_VENDOR)
    If cfg("MalzemeCol") < 1 Then cfg("MalzemeCol") = ColLetterToNum(DEF_COL_MALZEME)
    If cfg("UretimYeriCol") < 1 Then cfg("UretimYeriCol") = ColLetterToNum(DEF_COL_URETIM_YERI)
    If cfg("UzunMetinCol") < 1 Then cfg("UzunMetinCol") = ColLetterToNum(DEF_COL_UZUN_METIN)
    If cfg("BirimCol") < 1 Then cfg("BirimCol") = ColLetterToNum(DEF_COL_BIRIM)
    If cfg("BakiyeCol") < 1 Then cfg("BakiyeCol") = ColLetterToNum(DEF_COL_BAKIYE)
    If cfg("StatusCol") < 1 Then cfg("StatusCol") = ColLetterToNum(DEF_COL_STATUS)
    If cfg("ReplyCol") < 1 Then cfg("ReplyCol") = ColLetterToNum(DEF_COL_REPLY)
    If cfg("ReplyDateCol") < 1 Then cfg("ReplyDateCol") = ColLetterToNum(DEF_COL_REPLY_DATE)
    If cfg("ReplyFromCol") < 1 Then cfg("ReplyFromCol") = ColLetterToNum(DEF_COL_REPLY_FROM)
    If cfg("MatchMethodCol") < 1 Then cfg("MatchMethodCol") = ColLetterToNum(DEF_COL_MATCH_METHOD)
    If cfg("AttachCol") < 1 Then cfg("AttachCol") = ColLetterToNum(DEF_COL_ATTACH)
    If cfg("VendorNameCol") < 1 Then cfg("VendorNameCol") = ColLetterToNum(DEF_COL_VENDOR_NAME)
    If cfg("VendorEmailCol") < 1 Then cfg("VendorEmailCol") = ColLetterToNum(DEF_COL_VENDOR_EMAIL)
End Sub

Private Sub SaveSettingsToSheet(ByRef cfg As Object)
    Dim ws As Worksheet, orderSheets As Collection, s As String, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add: ws.Name = CONFIG_SHEET
    On Error GoTo 0
    ws.Cells.Clear
    ws.Range("A1").Value = T("SÝPARÝÞ SAYFALARI (virgülle ayrýlmýþ)")
    Set orderSheets = cfg("OrderSheets")
    s = "": For i = 1 To orderSheets.Count
        If Len(s) > 0 Then s = s & ","
        s = s & orderSheets(i).Name
    Next i
    ws.Range("B1").Value = s
    ws.Range("A2").Value = T("TEDARÝKÇÝ SAYFASI"): ws.Range("B2").Value = cfg("VendorWs").Name
    ws.Range("A3").Value = T("BAÞLANGIÇ SATIRI"):  ws.Range("B3").Value = cfg("StartRow")
    ws.Range("A4").Value = T("OUTLOOK KLASÖR"):      ws.Range("B4").Value = cfg("InboxFolder")
    ws.Range("A6").Value = T("SÝPARÝÞ SÜTUNLARI (harf)")
    ws.Range("A7").Value = T("Sipariþ No"):          ws.Range("B7").Value = NumToColLetter(cfg("OrderNoCol"))
    ws.Range("A8").Value = T("Kalem No"):             ws.Range("B8").Value = NumToColLetter(cfg("KalemNoCol"))
    ws.Range("A9").Value = T("Tedarikçi"):           ws.Range("B9").Value = NumToColLetter(cfg("VendorCol"))
    ws.Range("A10").Value = T("Malzeme"):             ws.Range("B10").Value = NumToColLetter(cfg("MalzemeCol"))
    ws.Range("A11").Value = T("Üretim Yeri"):        ws.Range("B11").Value = NumToColLetter(cfg("UretimYeriCol"))
    ws.Range("A12").Value = T("Uzun Metin"):          ws.Range("B12").Value = NumToColLetter(cfg("UzunMetinCol"))
    ws.Range("A13").Value = T("Birim"):               ws.Range("B13").Value = NumToColLetter(cfg("BirimCol"))
    ws.Range("A14").Value = T("Bakiye"):              ws.Range("B14").Value = NumToColLetter(cfg("BakiyeCol"))
    ws.Range("A15").Value = T("Mail Atýldý"):       ws.Range("B15").Value = NumToColLetter(cfg("StatusCol"))
    ws.Range("A16").Value = T("Gelen Yanýt"):        ws.Range("B16").Value = NumToColLetter(cfg("ReplyCol"))
    ws.Range("A17").Value = T("Yanýt Tarihi"):       ws.Range("B17").Value = NumToColLetter(cfg("ReplyDateCol"))
    ws.Range("A18").Value = T("Yanýtlayan"):         ws.Range("B18").Value = NumToColLetter(cfg("ReplyFromCol"))
    ws.Range("A19").Value = T("Eþleþme Yöntemi"):  ws.Range("B19").Value = NumToColLetter(cfg("MatchMethodCol"))
    ws.Range("A20").Value = T("Ekler"):               ws.Range("B20").Value = NumToColLetter(cfg("AttachCol"))
    ws.Range("A22").Value = T("TEDARÝKÇÝ SÜTUNLARI")
    ws.Range("A23").Value = T("Satýcý Ýsmi"):      ws.Range("B23").Value = NumToColLetter(cfg("VendorNameCol"))
    ws.Range("A24").Value = T("Mail Adresi"):         ws.Range("B24").Value = NumToColLetter(cfg("VendorEmailCol"))
    ws.Range("A26").Value = T("TEST MODU (EVET/HAYIR)")
    If IsTestMode() Then ws.Range("B26").Value = "EVET" Else ws.Range("B26").Value = "HAYIR"
    ws.Columns("A:A").ColumnWidth = 40: ws.Columns("B:B").ColumnWidth = 25
    On Error Resume Next: ws.Visible = xlSheetHidden: On Error GoTo 0
End Sub

Private Function CollectSettings(ByRef cfg As Object) As Boolean
    Dim orderSheets As Collection, ws As Worksheet
    Dim sheetNames As String, parts() As String
    Dim i As Long, sname As String, answer As VbMsgBoxResult
    Dim vws As Worksheet, startRow As Long, inboxFolder As String, sampleWs As Worksheet
    Set orderSheets = New Collection
    sheetNames = InputBox(T("Mail atýlacak SÝPARÝÞ SAYFALARINI girin") & vbCrLf & _
                          T("(birden fazla için VÝRGÜL ile ayýrýn)") & vbCrLf & vbCrLf & _
                          T("Mevcut sayfalar:") & vbCrLf & ListSheetNames(), _
                          T("Açýk Sipariþ Mail Otomasyonu"), "Sheet1")
    If StrPtr(sheetNames) = 0 Then Exit Function
    sheetNames = Trim$(sheetNames)
    If Len(sheetNames) = 0 Then Exit Function
    parts = Split(sheetNames, ",")
    For i = LBound(parts) To UBound(parts)
        sname = Trim$(parts(i))
        If Len(sname) > 0 Then
            Set ws = Nothing
            On Error Resume Next: Set ws = ThisWorkbook.Worksheets(sname): On Error GoTo 0
            If ws Is Nothing Then
                MsgBox T("Sayfa bulunamadý: ") & sname, vbExclamation: Exit Function
            End If
            orderSheets.Add ws
        End If
    Next i
    If orderSheets.Count = 0 Then Exit Function
    Set vws = AskWorksheet(T("TEDARÝKÇÝ LÝSTESÝ hangi SAYFADA?"), "Sheet3")
    If vws Is Nothing Then Exit Function
    startRow = AskStartRow()
    If startRow = 0 Then Exit Function
    inboxFolder = InputBox(T("Okunacak Outlook klasör adý:") & vbCrLf & _
                           T("Boþ = varsayýlan Gelen Kutusu."), _
                           T("Açýk Sipariþ Mail Otomasyonu"), "")
    If StrPtr(inboxFolder) = 0 Then Exit Function
    inboxFolder = Trim$(inboxFolder)
    Set sampleWs = orderSheets(1)
    
    Set cfg("OrderSheets") = orderSheets
    Set cfg("VendorWs") = vws
    cfg("StartRow") = startRow
    cfg("InboxFolder") = inboxFolder
    
    ' Sütunlarý otomatik tespit et
    AutoFixAllColumns cfg, sampleWs, vws

    answer = MsgBox(T("Algýlanan ayarlar:") & vbCrLf & vbCrLf & GetSettingsSummary(cfg) & vbCrLf & vbCrLf & _
                    T("Devam edilsin mi?"), vbYesNo + vbQuestion)
    If answer <> vbYes Then Exit Function
    CollectSettings = True
End Function

Private Function GetSettingsSummary(ByRef cfg As Object) As String
    Dim s As String, i As Long, orderSheets As Collection
    Set orderSheets = cfg("OrderSheets")
    s = T("SÝPARÝÞ SAYFALARI") & " (" & orderSheets.Count & T(" adet):") & vbCrLf
    For i = 1 To orderSheets.Count
        s = s & "  " & i & ") " & orderSheets(i).Name & vbCrLf
    Next i
    s = s & vbCrLf & T("TEDARÝKÇÝ: ") & cfg("VendorWs").Name & vbCrLf
    s = s & T("BAÞLANGIÇ SATIRI: ") & cfg("StartRow") & vbCrLf
    s = s & T("OUTLOOK KLASÖR: ") & IIf(Len(cfg("InboxFolder")) = 0, T("(varsayýlan)"), cfg("InboxFolder")) & vbCrLf & vbCrLf
    s = s & T("SÜTUNLAR (1. Satýrdan Okundu):") & vbCrLf
    s = s & T("  Tedarikçi: ") & NumToColLetter(cfg("VendorCol")) & vbCrLf
    s = s & T("  Mail Atýldý: ") & NumToColLetter(cfg("StatusCol")) & vbCrLf
    s = s & T("  TEST MODU: ") & IIf(IsTestMode(), "EVET (ilk 5 taslak)", "HAYIR (tam)") & vbCrLf
    GetSettingsSummary = s
End Function

Private Sub DoCreateDrafts(ByRef cfg As Object)
    Dim olApp As Object, orderSheets As Collection, i As Long
    Dim createdCount As Long, skippedCount As Long, failedCount As Long, totalRows As Long
    Dim tMode As Boolean: tMode = IsTestMode()

    If tMode Then
        If MsgBox("TEST MODU ACIK: Ýlk " & TEST_MAX_ORDERS & _
                  " geçerli taslak oluþturulana kadar tarama yapýlacak." & vbCrLf & "Devam edilsin mi?", _
                  vbYesNo + vbQuestion, "Test Modu") <> vbYes Then Exit Sub
    End If

    Set olApp = GetOutlookApp()
    If olApp Is Nothing Then
        MsgBox T("Outlook'a baðlanýlamadý."), vbCritical: Exit Sub
    End If

    EnsureReportSheets
    gSignatureFetched = False: gSignatureHtml = ""
    Set orderSheets = cfg("OrderSheets")
    WriteLog "BASLAT", "", "", "", "TaslaklariOlustur", orderSheets.Count & " sayfa"

    For i = 1 To orderSheets.Count
        ProcessDraftsForSheet cfg, orderSheets(i), olApp, createdCount, skippedCount, failedCount, totalRows, tMode
    Next i

    WriteLog "BITIR", "", "", "", "TaslaklariOlustur", _
             "Taslak: " & createdCount & ", Bulunamadi: " & skippedCount & ", Hata: " & failedCount & ", Satir: " & totalRows

    Dim msg As String
    msg = "Islem tamamlandi." & vbCrLf & vbCrLf & _
          "Oluþturulan taslak: " & createdCount & vbCrLf & _
          "Mail adresi bulunamayan: " & skippedCount & vbCrLf & _
          "Kaydedilemeyen (Hata): " & failedCount & vbCrLf & _
          "Ýþlenen geçerli satýr: " & totalRows & vbCrLf & vbCrLf
    
    If createdCount = 0 Then
        msg = msg & "DÝKKAT: Hiç taslak oluþturulamadý!" & vbCrLf & _
              "- Tüm satýrlarda 'Mail Atýldý' sütunu dolu olabilir." & vbCrLf & _
              "- Tedarikçi adý boþ olabilir." & vbCrLf & _
              "- E-posta adresi bulunamamýþ olabilir. (Detay: EksikMail sayfasý)" & vbCrLf & _
              "- Outlook taslak kaydederken hata vermiþ olabilir. (Detay: Log sayfasý)"
    Else
        msg = msg & "Outlook > Taslaklar klasörüne gidin, kontrol edip gönderin." & vbCrLf & _
              "Detay: 'Log' ve 'MailRaporu' sayfalarýna bakýn."
    End If

    MsgBox T(msg), vbInformation, T("Açýk Sipariþ Mail Otomasyonu")
End Sub

Private Sub ProcessDraftsForSheet(ByRef cfg As Object, ByVal ws As Worksheet, _
                                  ByVal olApp As Object, _
                                  ByRef createdCount As Long, ByRef skippedCount As Long, _
                                  ByRef failedCount As Long, ByRef totalRows As Long, _
                                  ByVal tMode As Boolean)
    Dim vws As Worksheet, vendorCol As Long, statusCol As Long
    Dim vendorNameCol As Long, vendorEmailCol As Long
    Dim startRow As Long, lastRow As Long, vLastRow As Long, r As Long
    Dim groups As Object, vendorNames As Object, vendorRows As Collection, key As Variant
    Dim vendorName As String, emailAddr As String, matchState As String
    Dim mail As Object, html As String, vn As String, keyStr As String
    Dim sVal As String, previewQueue As Collection
    Dim ignoreStatus As Boolean
    
    Set previewQueue = New Collection
    ignoreStatus = False

    On Error GoTo SheetError
    Set vws = cfg("VendorWs")
    vendorCol = cfg("VendorCol"): statusCol = cfg("StatusCol")
    vendorNameCol = cfg("VendorNameCol"): vendorEmailCol = cfg("VendorEmailCol")
    startRow = cfg("StartRow")

    lastRow = ws.Cells(ws.Rows.Count, vendorCol).End(xlUp).Row
    vLastRow = vws.Cells(vws.Rows.Count, vendorNameCol).End(xlUp).Row
    
    If startRow < 2 Then startRow = 2
    If lastRow < startRow Then
        WriteLog "BILGI", ws.Name, "", "", "Taslak", "Veri yok (lastRow < startRow)"
        Exit Sub
    End If

    ' Eðer sütun doluysa kullanýcýya sor
    Dim hasStatusData As Boolean
    hasStatusData = False
    For r = startRow To lastRow
        If SafeStr(ws.Cells(r, statusCol).Value) <> "" Then
            hasStatusData = True
            Exit For
        End If
    Next r
    
    If hasStatusData Then
        If MsgBox("'" & ws.Name & "' sayfasýnda 'Mail Atýldý' sütununda dolu hücreler var." & vbCrLf & _
                  "Bunlar zaten iþlenmiþ demektir. Yeniden taslak oluþturmak ister misiniz?" & vbCrLf & vbCrLf & _
                  "(Evet derseniz bu sütundaki veriler yoksayýlýr ve tüm satýrlar iþlenir)", _
                  vbYesNo + vbQuestion, "Durum Sütunu Dolu") = vbYes Then
            ignoreStatus = True
        End If
    End If

    Set groups = CreateObject("Scripting.Dictionary")
    Set vendorNames = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    vendorNames.CompareMode = vbTextCompare

    For r = startRow To lastRow
        If tMode And (createdCount + skippedCount + failedCount) >= TEST_MAX_ORDERS Then Exit For
        
        vn = SafeStr(ws.Cells(r, vendorCol).Value)
        sVal = SafeStr(ws.Cells(r, statusCol).Value)
        
        ' Eðer kullanýcý yoksay dediyse veya sütun boþsa iþle
        If Len(vn) > 0 And (ignoreStatus Or Len(sVal) = 0) Then
            keyStr = NormalizeName(vn)
            If Not groups.Exists(keyStr) Then
                Set vendorRows = New Collection
                groups.Add keyStr, vendorRows
                vendorNames.Add keyStr, vn
            End If
            groups(keyStr).Add r
            totalRows = totalRows + 1
        End If
    Next r

    If groups.Count = 0 Then Exit Sub

    For Each key In groups.Keys
        If tMode And (createdCount + skippedCount + failedCount) >= TEST_MAX_ORDERS Then Exit For
        
        vendorName = CStr(vendorNames(key))
        Set vendorRows = groups(key)
        
        emailAddr = FindVendorEmail(vws, vLastRow, vendorNameCol, vendorEmailCol, vendorName, matchState)

        If Len(emailAddr) = 0 Or Left$(matchState, 2) <> "OK" Then
            skippedCount = skippedCount + 1
            Dim mark As String: mark = NOT_FOUND_MARK
            If Len(matchState) > 0 And UCase$(matchState) <> UCase$(NOT_FOUND_MARK) Then
                mark = NOT_FOUND_MARK & " (" & matchState & ")"
            End If
            MarkRows ws, vendorRows, statusCol, mark
            WriteLog "ATLANDI", ws.Name, vendorName, "", "Taslak", matchState
            WriteMissingRow ws.Name, vendorName, matchState, vendorRows.Count
        Else
            html = BuildEmailHtml(ws, vendorRows, cfg)
            Set mail = olApp.CreateItem(0)
            On Error Resume Next
            With mail
                .To = emailAddr
                .Subject = T("Açýk sipariþler hakkýnda")
                .BodyFormat = 2
                .HTMLBody = html & FetchSignature(olApp)
                AddCcRecipients mail, CC_RECIPIENTS
            End With
            mail.Save
            If Err.Number <> 0 Then
                WriteLog "HATA", ws.Name, vendorName, emailAddr, "Taslak", Err.Description
                Err.Clear
                failedCount = failedCount + 1
            Else
                MarkRows ws, vendorRows, statusCol, DRAFT_MARK
                createdCount = createdCount + 1
                WriteLog "TASLAK", ws.Name, vendorName, emailAddr, "OK", vendorRows.Count & " satir"
                WriteReportRow ws.Name, vendorName, emailAddr, vendorRows.Count
                If tMode And previewQueue.Count < 5 Then previewQueue.Add mail
            End If
            On Error GoTo SheetError
            Set mail = Nothing
        End If
    Next key

    If tMode Then
        Dim k As Long
        For k = 1 To previewQueue.Count
            On Error Resume Next
            previewQueue(k).Display
            On Error GoTo SheetError
        Next k
    End If
    Exit Sub
SheetError:
    WriteLog "HATA", ws.Name, "", "", "ProcessDraftsForSheet", Err.Description
End Sub

Private Function BuildEmailHtml(ByVal ws As Worksheet, ByVal rows As Collection, _
                                ByRef cfg As Object) As String
    Dim h As String, i As Long, r As Long
    Dim cOrder As Long, cKalem As Long, cMalzeme As Long, cUretim As Long
    Dim cUzun As Long, cBirim As Long, cBakiye As Long
    cOrder = cfg("OrderNoCol"): cKalem = cfg("KalemNoCol"): cMalzeme = cfg("MalzemeCol")
    cUretim = cfg("UretimYeriCol"): cUzun = cfg("UzunMetinCol")
    cBirim = cfg("BirimCol"): cBakiye = cfg("BakiyeCol")

    h = "<html><head><meta charset=""UTF-8""></head>" & _
        "<body style='font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1f1f1f;'>"
    h = h & "<p>" & T("Merhabalar,") & "</p>"
    h = h & "<p>" & T("Aþaðýdaki açýk sipariþlerimiz için termin tarihi nedir? Ne zaman teslim edilecek?") & "</p>"
    h = h & "<p>" & T("Dönüþlerinizi rica ederiz.") & "</p>"
    h = h & "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;font-size:10pt;'>"
    h = h & "<tr style='font-weight:bold;background-color:#f2f2f2;'>"
    h = h & "<td>" & T("ÜRETÝM YERÝ") & "</td><td>" & T("MALZEME") & "</td><td>" & _
            T("MALZEME UZUN METNÝ") & "</td><td>" & T("SÝPARÝÞ NO") & "</td><td>" & _
            T("KALEM NO") & "</td><td>" & T("TEMEL ÖLÇÜ BÝRÝMÝ") & "</td><td>" & _
            T("BAKÝYE MÝKTARI") & "</td><td style='background-color:#fff9c4;'>" & _
            T("TERMÝN TARÝHÝ") & "</td></tr>"

    For i = 1 To rows.Count
        r = CLng(rows(i))
        h = h & "<tr>"
        h = h & "<td>" & HtmlEncode(ws.Cells(r, cUretim).Text) & "</td>"
        h = h & "<td>" & HtmlEncode(ws.Cells(r, cMalzeme).Text) & "</td>"
        h = h & "<td>" & HtmlEncode(ws.Cells(r, cUzun).Text) & "</td>"
        h = h & "<td>" & HtmlEncode(ws.Cells(r, cOrder).Text) & "</td>"
        h = h & "<td>" & HtmlEncode(ws.Cells(r, cKalem).Text) & "</td>"
        h = h & "<td>" & HtmlEncode(ws.Cells(r, cBirim).Text) & "</td>"
        h = h & "<td>" & HtmlEncode(ws.Cells(r, cBakiye).Text) & "</td>"
        h = h & "<td style='background-color:#fff9c4;'>&nbsp;</td>"
        h = h & "</tr>"
    Next i
    h = h & "</table><p>" & T("Saygýlarýmla.") & "</p></body></html>"
    BuildEmailHtml = h
End Function

Private Function FetchSignature(ByVal olApp As Object) As String
    If gSignatureFetched Then FetchSignature = gSignatureHtml: Exit Function
    Dim tmp As Object
    On Error Resume Next
    Set tmp = olApp.CreateItem(0)
    If Not tmp Is Nothing Then
        tmp.Display
        If Err.Number = 0 Then gSignatureHtml = tmp.HTMLBody
        tmp.Close 1
    End If
    Err.Clear
    On Error GoTo 0
    gSignatureFetched = True
    FetchSignature = gSignatureHtml
End Function

Private Sub EnsureReportSheets()
    On Error Resume Next
    Set gReportWs = ThisWorkbook.Worksheets(REPORT_SHEET)
    If gReportWs Is Nothing Then
        Set gReportWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        gReportWs.Name = REPORT_SHEET
    End If
    gReportWs.Cells.Clear
    With gReportWs
        .Range("A1").Value = T("Tarih/Saat"): .Range("B1").Value = T("Sayfa")
        .Range("C1").Value = T("Tedarikçi"):  .Range("D1").Value = T("Mail")
        .Range("E1").Value = T("Satýr Sayýsý"): .Range("F1").Value = T("Durum")
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(197, 217, 241)
        .Columns("A:F").ColumnWidth = 22
    End With

    Set gMissingWs = ThisWorkbook.Worksheets(MISSING_SHEET)
    If gMissingWs Is Nothing Then
        Set gMissingWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        gMissingWs.Name = MISSING_SHEET
    End If
    gMissingWs.Cells.Clear
    With gMissingWs
        .Range("A1").Value = T("Tarih/Saat"): .Range("B1").Value = T("Sayfa")
        .Range("C1").Value = T("Tedarikçi"):  .Range("D1").Value = T("Neden")
        .Range("E1").Value = T("Etkilenen Satýr")
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(255, 199, 206)
        .Columns("A:E").ColumnWidth = 22
    End With
    On Error GoTo 0
End Sub

Private Sub WriteReportRow(ByVal sayfa As String, ByVal tedarikci As String, _
                           ByVal mail As String, ByVal satirSayisi As Long)
    On Error Resume Next
    If gReportWs Is Nothing Then Exit Sub
    Dim r As Long
    r = gReportWs.Cells(gReportWs.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    gReportWs.Cells(r, 1).Value = Format(Now, "dd.mm.yyyy hh:nn:ss")
    gReportWs.Cells(r, 2).Value = sayfa
    gReportWs.Cells(r, 3).Value = tedarikci
    gReportWs.Cells(r, 4).Value = mail
    gReportWs.Cells(r, 5).Value = satirSayisi
    gReportWs.Cells(r, 6).Value = "TASLAK"
    On Error GoTo 0
End Sub

Private Sub WriteMissingRow(ByVal sayfa As String, ByVal tedarikci As String, _
                            ByVal neden As String, ByVal satirSayisi As Long)
    On Error Resume Next
    If gMissingWs Is Nothing Then Exit Sub
    Dim r As Long
    r = gMissingWs.Cells(gMissingWs.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    gMissingWs.Cells(r, 1).Value = Format(Now, "dd.mm.yyyy hh:nn:ss")
    gMissingWs.Cells(r, 2).Value = sayfa
    gMissingWs.Cells(r, 3).Value = tedarikci
    gMissingWs.Cells(r, 4).Value = neden
    gMissingWs.Cells(r, 5).Value = satirSayisi
    On Error GoTo 0
End Sub

Private Function FindVendorEmail(ByVal vws As Worksheet, ByVal lastRow As Long, _
                                 ByVal nameCol As Long, ByVal emailCol As Long, _
                                 ByVal vendorName As String, ByRef matchState As String) As String
    Dim r As Long, src As String, cand As String
    Dim target As String, targetShort As String
    Dim found As String, hits As Long
    
    ' Aradýðýn ismi hata mesajýnda göster ki kullanýcý neyi aradýðýný anlasýn
    matchState = "BULUNAMADI (" & Left$(vendorName, 15) & "...)"
    
    target = NormalizeName(vendorName)
    targetShort = NormalizeVendorName(vendorName)

    If lastRow < 2 Then lastRow = 1000

    ' 1. Aþama: Tam Eþleþme
    found = "": hits = 0
    For r = 2 To lastRow
        src = NormalizeName(SafeStr(vws.Cells(r, nameCol).Value))
        If src = target Then
            cand = SafeStr(vws.Cells(r, emailCol).Value)
            If IsValidEmail(cand) Then
                If Len(found) = 0 Then
                    found = cand: hits = 1
                ElseIf LCase$(found) <> LCase$(cand) Then
                    hits = hits + 1
                End If
            End If
        End If
    Next r
    If hits = 1 Then matchState = "OK": FindVendorEmail = found: Exit Function
    If hits > 1 Then matchState = T("AYNI ÝSÝMDE FARKLI E-POSTALAR"): Exit Function

    ' 2. Aþama: Sonekler Temizlenmiþ Eþleþme
    If Len(targetShort) >= 4 And targetShort <> target Then
        found = "": hits = 0
        For r = 2 To lastRow
            src = NormalizeVendorName(SafeStr(vws.Cells(r, nameCol).Value))
            If src = targetShort Then
                cand = SafeStr(vws.Cells(r, emailCol).Value)
                If IsValidEmail(cand) Then
                    If Len(found) = 0 Then
                        found = cand: hits = 1
                    ElseIf LCase$(found) <> LCase$(cand) Then
                        hits = hits + 1
                    End If
                End If
            End If
        Next r
        If hits = 1 Then matchState = "OK (suffix'siz)": FindVendorEmail = found: Exit Function
    End If

    ' 3. Aþama: Ýçinde Geçme (Kýsmi Eþleþme)
    found = "": hits = 0
    For r = 2 To lastRow
        src = NormalizeVendorName(SafeStr(vws.Cells(r, nameCol).Value))
        If Len(src) > 4 And Len(targetShort) > 4 Then
            If InStr(1, src, targetShort, vbTextCompare) > 0 Or InStr(1, targetShort, src, vbTextCompare) > 0 Then
                cand = SafeStr(vws.Cells(r, emailCol).Value)
                If IsValidEmail(cand) Then
                    If Len(found) = 0 Then
                        found = cand: hits = 1
                    ElseIf LCase$(found) <> LCase$(cand) Then
                        hits = hits + 1
                    End If
                End If
            End If
        End If
    Next r
    If hits = 1 Then matchState = "OK (içerir)": FindVendorEmail = found: Exit Function

    ' 4. Aþama: Ýlk 2 Kelime Eþleþmesi
    Dim twoWords As String: twoWords = GetFirstNthWord(targetShort, 2)
    If Len(twoWords) >= 6 Then
        Dim candRows As Object: Set candRows = CreateObject("Scripting.Dictionary")
        For r = 2 To lastRow
            Dim rowShort As String, rowFirst2 As String
            rowShort = NormalizeVendorName(SafeStr(vws.Cells(r, nameCol).Value))
            rowFirst2 = GetFirstNthWord(rowShort, 2)
            If Len(rowFirst2) >= 6 And rowFirst2 = twoWords Then
                cand = SafeStr(vws.Cells(r, emailCol).Value)
                If IsValidEmail(cand) Then candRows.Add r, cand
            End If
        Next r
        If candRows.Count = 1 Then
            matchState = "OK (2 kelime)": FindVendorEmail = CStr(candRows(candRows.Keys()(0))): Exit Function
        End If
    End If

    ' 5. Aþama: Ýlk 1 Kelime Eþleþmesi
    Dim firstWord As String: firstWord = GetFirstNthWord(targetShort, 1)
    If Len(firstWord) >= 4 Then
        Dim candRows2 As Object: Set candRows2 = CreateObject("Scripting.Dictionary")
        For r = 2 To lastRow
            Dim rowShort2 As String, rowFirst As String
            rowShort2 = NormalizeVendorName(SafeStr(vws.Cells(r, nameCol).Value))
            rowFirst = GetFirstNthWord(rowShort2, 1)
            If Len(rowFirst) >= 4 And rowFirst = firstWord Then
                cand = SafeStr(vws.Cells(r, emailCol).Value)
                If IsValidEmail(cand) Then candRows2.Add r, cand
            End If
        Next r
        If candRows2.Count = 1 Then
            matchState = "OK (1 kelime)"
            FindVendorEmail = CStr(candRows2(candRows2.Keys()(0))): Exit Function
        ElseIf candRows2.Count > 1 Then
            matchState = "BULUNAMADI (Coklu eslesme)"
        End If
    End If
End Function

Private Function GetConfig() As Object
    If gCfg Is Nothing Then
        Set gCfg = CreateObject("Scripting.Dictionary")
        gCfg.CompareMode = vbTextCompare
    End If
    Set GetConfig = gCfg
End Function

Private Function GetOutlookApp() As Object
    Dim app As Object
    On Error Resume Next
    Set app = GetObject(, "Outlook.Application")
    If app Is Nothing Then Set app = CreateObject("Outlook.Application")
    On Error GoTo 0
    Set GetOutlookApp = app
End Function

Private Sub AddCcRecipients(ByVal mail As Object, ByVal ccList As String)
    Dim parts() As String, i As Long, addr As String, rec As Object
    On Error Resume Next
    parts = Split(ccList, ";")
    For i = LBound(parts) To UBound(parts)
        addr = Trim$(parts(i))
        If Len(addr) > 0 Then
            Set rec = mail.Recipients.Add(addr)
            rec.Type = 2
        End If
    Next i
    mail.Recipients.ResolveAll
    On Error GoTo 0
End Sub

Private Sub MarkRows(ByVal ws As Worksheet, ByVal rows As Collection, _
                     ByVal col As Long, ByVal statusText As String)
    Dim i As Long
    For i = 1 To rows.Count
        ws.Cells(CLng(rows(i)), col).Value = statusText
    Next i
End Sub

Private Function GetFirstNthWord(ByVal s As String, ByVal n As Long) As String
    Dim parts() As String, i As Long, result As String, count As Long
    s = Trim$(s)
    s = Replace(s, ".", " "): s = Replace(s, ",", " "): s = Replace(s, ";", " ")
    s = Replace(s, "-", " "): s = Replace(s, "/", " "): s = Replace(s, "&", " ")
    s = Replace(s, "'", " "): s = Replace(s, """", " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    parts = Split(Trim$(s), " ")
    result = "": count = 0
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            If Len(result) > 0 Then result = result & " "
            result = result & parts(i)
            count = count + 1
            If count >= n Then Exit For
        End If
    Next i
    GetFirstNthWord = Trim$(result)
End Function

Private Function NormalizeName(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, ChrW(304), "I"): s = Replace(s, ChrW(305), "I")
    s = Replace(s, ChrW(350), "S"): s = Replace(s, ChrW(351), "S")
    s = Replace(s, ChrW(286), "G"): s = Replace(s, ChrW(287), "G")
    s = Replace(s, ChrW(220), "U"): s = Replace(s, ChrW(252), "U")
    s = Replace(s, ChrW(214), "O"): s = Replace(s, ChrW(246), "O")
    s = Replace(s, ChrW(199), "C"): s = Replace(s, ChrW(231), "C")
    s = Replace(s, ".", " "): s = Replace(s, ",", " "): s = Replace(s, ";", " ")
    s = Replace(s, ":", " "): s = Replace(s, "-", " "): s = Replace(s, "_", " ")
    s = Replace(s, "&", " "): s = Replace(s, "/", " "): s = Replace(s, "'", " ")
    s = Replace(s, """", " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeName = Trim$(s)
End Function

Private Function NormalizeVendorName(ByVal s As String) As String
    Dim n As String: n = NormalizeName(s)
    n = RemoveCompanySuffixes(n)
    NormalizeVendorName = n
End Function

Private Function RemoveCompanySuffixes(ByVal s As String) As String
    Dim arr As Variant, i As Long, j As Long
    Dim parts() As String, result As String
    Dim p As String, keep As Boolean
    Dim arrStr As String
    
    ' "V" harfi de eklendi (SANAYI V gibi kýsaltmalar için)
    arrStr = "TIC|TICARET|LTD|LIMITED|STI|SIRKETI|SIRKET|AS|ANONIM|ANONIMSIRKETI|" & _
             "SAN|SANAYI|SANAYII|PAZ|PAZARLAMA|VE|V|DIS|DISTICARET|ITH|ITHALAT|IHR|IHRACAT|" & _
             "TAAH|TAAHHUT|TAAHUT|MUM|MUMESSILLIK|MUMES|MUH|MUHENDISLIK|MAK|MAKINA|MAKINE|" & _
             "KOLL|KOLLEKTIF|ORT|ORTAKLARI|ORTAKLIGI|GRUP|GROUP|GRUBU|HOLDING|HOLDIN|" & _
             "KARDESLERI|KARDESI|KOMP|KOMPRESSOR|KOMPRESOR|END|ENDUSTRIYEL|ENDUSTRI|" & _
             "OTO|OTOMOTIV|OTOMOTIF|ELEK|ELEKTRIK|ELEKTRONIK|ELEKTRONI|MALZ|MALZEMELERI|MALZEMESI|" & _
             "YAPI|YATIRIM|YATIRIMLARI|URUN|URUNLERI|URUNLER|SIS|SISTEMLERI|SISTEM|" & _
             "TEM|TEMSILCILIK|TEMSILCI|BIL|BILISIM|SAG|SAGLIK|GIDA|TARIM|" & _
             "TEK|TEKNIK|TEKNOLOJI|TEKNOLOJILERI|TUR|TURIZM|INS|INSAAT|INSAT|" & _
             "MET|METAL|METALURJI|KIM|KIMYA|KIMYASAL|PLAST|PLASTIK|" & _
             "AMB|AMBALAJ|LAB|LABORATUVAR|HIR|HIRDAVAT|OFT|OFSET|MATBAA|MATBAACILIK|" & _
             "REK|REKLAM|DAN|DANISMANLIK|PRO|PROJE|PROJELERI"
    arr = Split(arrStr, "|")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    parts = Split(s, " ")
    result = ""
    For i = LBound(parts) To UBound(parts)
        p = Trim$(UCase$(parts(i)))
        p = Replace(p, ".", ""): p = Replace(p, ",", "")
        p = Replace(p, ";", ""): p = Replace(p, ":", "")
        p = Trim$(p)
        If Len(p) > 0 Then
            keep = True
            For j = LBound(arr) To UBound(arr)
                If p = CStr(arr(j)) Then keep = False: Exit For
            Next j
            If Len(p) <= 1 Then keep = False ' Tek harflileri sil (örn: V, A)
            If keep Then
                If Len(result) > 0 Then result = result & " "
                result = result & p
            End If
        End If
    Next i
    If Len(Trim$(result)) = 0 Then result = s
    RemoveCompanySuffixes = Trim$(result)
End Function

Private Function IsValidEmail(ByVal s As String) As Boolean
    s = Trim$(s)
    If InStr(1, s, "@") < 2 Then Exit Function
    If InStrRev(s, ".") < InStr(1, s, "@") + 2 Then Exit Function
    If Right$(s, 1) = "." Then Exit Function
    IsValidEmail = True
End Function

Private Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;"): s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;"): s = Replace(s, """", "&quot;")
    HtmlEncode = s
End Function

Private Function ColLetterToNum(ByVal letter As String) As Long
    Dim i As Long, n As Long, ch As String
    letter = UCase$(Trim$(letter))
    If Len(letter) = 0 Then Exit Function
    For i = 1 To Len(letter)
        ch = Mid$(letter, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
        n = n * 26 + (Asc(ch) - 64)
    Next i
    ColLetterToNum = n
End Function

Private Function NumToColLetter(ByVal colNum As Long) As String
    Dim s As String, n As Long, rr As Long
    n = colNum
    Do While n > 0
        rr = (n - 1) Mod 26
        s = Chr$(65 + rr) & s
        n = (n - 1) \ 26
    Loop
    NumToColLetter = s
End Function

Private Function FindColumnByHeaders(ByVal ws As Worksheet, ByVal headers As Variant, ByVal headerRow As Long) As Long
    Dim lastCol As Long, c As Long, i As Long
    Dim cv As String, cvClean As String, hvClean As String
    On Error Resume Next
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    On Error GoTo 0
    If lastCol < 1 Then lastCol = 1
    
    ' 1. Aþama: Tam Eþleþme
    For c = 1 To lastCol
        cv = SafeStr(ws.Cells(headerRow, c).Value)
        If Len(cv) > 0 Then
            cvClean = CleanForCompare(cv)
            For i = LBound(headers) To UBound(headers)
                hvClean = CleanForCompare(CStr(headers(i)))
                If cvClean = hvClean Then
                    FindColumnByHeaders = c: Exit Function
                End If
            Next i
        End If
    Next c
    
    ' 2. Aþama: Ýçinde Geçme (Contains)
    For c = 1 To lastCol
        cv = SafeStr(ws.Cells(headerRow, c).Value)
        If Len(cv) > 0 Then
            cvClean = CleanForCompare(cv)
            For i = LBound(headers) To UBound(headers)
                hvClean = CleanForCompare(CStr(headers(i)))
                If Len(hvClean) >= 4 Then
                    If InStr(1, cvClean, hvClean, vbTextCompare) > 0 Then
                        FindColumnByHeaders = c: Exit Function
                    End If
                End If
            Next i
        End If
    Next c
End Function

Private Function CleanForCompare(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, " ", ""): s = Replace(s, ".", "")
    s = Replace(s, "-", ""): s = Replace(s, "_", "")
    s = Replace(s, ChrW(304), "I"): s = Replace(s, ChrW(305), "I")
    s = Replace(s, ChrW(350), "S"): s = Replace(s, ChrW(351), "S")
    s = Replace(s, ChrW(286), "G"): s = Replace(s, ChrW(287), "G")
    s = Replace(s, ChrW(220), "U"): s = Replace(s, ChrW(252), "U")
    s = Replace(s, ChrW(214), "O"): s = Replace(s, ChrW(246), "O")
    s = Replace(s, ChrW(199), "C"): s = Replace(s, ChrW(231), "C")
    CleanForCompare = s
End Function

Private Function AskWorksheet(ByVal prompt As String, ByVal defaultName As String) As Worksheet
    Dim s As String, ws As Worksheet
    s = InputBox(prompt & vbCrLf & vbCrLf & T("Mevcut sayfalar:") & vbCrLf & _
        ListSheetNames() & vbCrLf & vbCrLf & T("Boþ býrakýrsanýz '") & defaultName & _
        T("' kullanýlýr."), T("Açýk Sipariþ Mail Otomasyonu"), defaultName)
    If StrPtr(s) = 0 Then Exit Function
    s = Trim$(s)
    If Len(s) = 0 Then s = defaultName
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(s): On Error GoTo 0
    If ws Is Nothing Then
        MsgBox T("Sayfa bulunamadý: ") & s, vbExclamation: Exit Function
    End If
    Set AskWorksheet = ws
End Function

Private Function ListSheetNames() As String
    Dim i As Long, s As String
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & "  - " & ThisWorkbook.Worksheets(i).Name
    Next i
    ListSheetNames = s
End Function

Private Function AskStartRow() As Long
    Dim s As String, v As Long
    s = InputBox(T("Hangi satýrdan baþlansýn?") & vbCrLf & _
                 T("Boþ = ") & DEFAULT_START_ROW, T("Açýk Sipariþ Mail Otomasyonu"), CStr(DEFAULT_START_ROW))
    If StrPtr(s) = 0 Then Exit Function
    s = Trim$(s)
    If Len(s) = 0 Then AskStartRow = DEFAULT_START_ROW: Exit Function
    If Not IsNumeric(s) Then
        MsgBox T("Geçerli bir satýr numarasý girin."): Exit Function
    End If
    v = CLng(s)
    If v < 2 Then
        MsgBox T("Baþlangýç satýrý en az 2 olmalýdýr."): Exit Function
    End If
    AskStartRow = v
End Function

Private Sub WriteLog(ByVal islem As String, ByVal sayfa As String, _
                     ByVal tedarikci As String, ByVal mail As String, _
                     ByVal detay1 As String, ByVal detay2 As String)
    On Error Resume Next
    If gLogWs Is Nothing Then
        Set gLogWs = ThisWorkbook.Worksheets(LOG_SHEET)
        If gLogWs Is Nothing Then
            Set gLogWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            gLogWs.Name = LOG_SHEET
            gLogWs.Range("A1").Value = T("Tarih/Saat"): gLogWs.Range("B1").Value = T("Ýþlem")
            gLogWs.Range("C1").Value = T("Sayfa"): gLogWs.Range("D1").Value = T("Tedarikçi")
            gLogWs.Range("E1").Value = T("Mail"): gLogWs.Range("F1").Value = T("Detay 1")
            gLogWs.Range("G1").Value = T("Detay 2")
            gLogWs.Range("A1:G1").Font.Bold = True
            gLogWs.Range("A1:G1").Interior.Color = RGB(220, 220, 220)
            gLogWs.Columns("A:G").ColumnWidth = 25
        End If
        gLogRow = gLogWs.Cells(gLogWs.Rows.Count, 1).End(xlUp).Row + 1
        If gLogRow < 2 Then gLogRow = 2
    End If
    gLogWs.Cells(gLogRow, 1).Value = Format(Now, "dd.mm.yyyy hh:nn:ss")
    gLogWs.Cells(gLogRow, 2).Value = islem
    gLogWs.Cells(gLogRow, 3).Value = sayfa
    gLogWs.Cells(gLogRow, 4).Value = tedarikci
    gLogWs.Cells(gLogRow, 5).Value = mail
    gLogWs.Cells(gLogRow, 6).Value = detay1
    gLogWs.Cells(gLogRow, 7).Value = detay2
    gLogRow = gLogRow + 1
End Sub