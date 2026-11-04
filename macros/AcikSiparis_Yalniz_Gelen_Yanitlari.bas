Attribute VB_Name = "OnlyReadReplies"
Option Explicit

'============================================================
'  ACIK SIPARIS MAIL OTOMASYONU - GELEN YANITLARI OKU
'  Surum: 6.1 (Tablo Eslestirme Fix ve Ayar Butonlari)
'============================================================

Private Const TEST_MODE As Boolean = False
Private Const TEST_MAX_ORDERS As Long = 5
Private Const CONFIG_SHEET As String = "Ayarlar"
Private Const LOG_SHEET As String = "Log"
Private Const TERMIN_REPORT_SHEET As String = "TerminRaporu"
Private Const PROCESSED_SHEET As String = "IslenenMailler"
Private Const HEADER_ROW As Long = 1
Private Const DEFAULT_START_ROW As Long = 2
Private Const ATTACH_BASE_FOLDER As String = "YanitEkleri"
Private Const MAX_LOOKBACK_DAYS As Long = 180
Private Const MAX_BODY_LEN As Long = 32000

Private Const STATUS_TESLIM   As String = "Teslim Edilecek"
Private Const STATUS_URETIM   As String = "Uretimde"
Private Const STATUS_SEVK     As String = "Sevk Edildi"
Private Const STATUS_STOK_YOK As String = "Stok Yok"
Private Const STATUS_IPTAL    As String = "Iptal"
Private Const STATUS_BELIRSIZ As String = "Belirsiz"

Private Const DEF_COL_ORDER_NO       As String = "F"
Private Const DEF_COL_KALEM_NO       As String = "G"
Private Const DEF_COL_VENDOR         As String = "L"
Private Const DEF_COL_MALZEME        As String = "D"
Private Const DEF_COL_URETIM_YERI    As String = "B"
Private Const DEF_COL_UZUN_METIN     As String = "E"
Private Const DEF_COL_BIRIM          As String = "I"
Private Const DEF_COL_BAKIYE         As String = "H"
Private Const DEF_COL_STATUS         As String = "N"
Private Const DEF_COL_REPLY          As String = "O"
Private Const DEF_COL_REPLY_DATE     As String = "P"
Private Const DEF_COL_REPLY_FROM     As String = "Q"
Private Const DEF_COL_MATCH_METHOD   As String = "R"
Private Const DEF_COL_ATTACH         As String = "S"
Private Const DEF_COL_VENDOR_NAME    As String = "A"
Private Const DEF_COL_VENDOR_EMAIL   As String = "B"
Private Const DEF_COL_DURUM          As String = "M"
Private Const DEF_COL_TERMIN         As String = "T"

Private gCfg As Object
Private gLogWs As Worksheet
Private gLogRow As Long
Private gTermReportWs As Worksheet
Private gProcessedDict As Object
Private gInternalDomains As Object
Private gSiparisYokCount As Long
Private gBelirsizCount As Long

Private gOrigScreenUpdating As Boolean
Private gOrigEnableEvents As Boolean
Private gOrigCalculation As Long
Private gOrigDisplayStatusBar As Boolean
Private gStateSaved As Boolean

'============================================================
'  YENI EKLENEN BUTON MAKROLARI
'============================================================
Public Sub AyarlariAc()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Ayarlar sayfasý bulunamadý. Lütfen önce otomasyonu çalýþtýrýn.", vbExclamation, AppTitle()
        Exit Sub
    End If
    
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub

Public Sub AyarlariKaydet()
    Dim cfg As Object
    Set cfg = GetConfig()
    cfg.RemoveAll
    
    ' AutoFix'i atlayarak kullanýcýnýn manuel girdiði sütun harflerini koruyoruz
    If LoadSettingsFromSheet(cfg, skipAutoFix:=True) Then
        SaveSettingsToSheet cfg
        MsgBox "Ayarlar baþarýyla kaydedildi.", vbInformation, AppTitle()
    Else
        MsgBox "Ayarlar okunamadý veya hatalý. Lütfen sayfadaki bilgileri kontrol edin.", vbCritical, AppTitle()
    End If
End Sub

'============================================================
'  YARDIMCI
'============================================================
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

Private Function AppTitle() As String
    AppTitle = T("A~c~ik Sipari~s Mail Otomasyonu")
End Function

Private Function IsTestMode() As Boolean
    Dim ws As Worksheet: On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET): On Error GoTo 0
    If ws Is Nothing Then IsTestMode = TEST_MODE: Exit Function
    Dim v As String: v = UCase$(Trim$(SafeStr(ws.Range("B26").Value)))
    If v = "EVET" Or v = "TRUE" Or v = "1" Then
        IsTestMode = True
    ElseIf v = "HAYIR" Or v = "FALSE" Or v = "0" Then
        IsTestMode = False
    Else
        IsTestMode = TEST_MODE
    End If
End Function

Private Sub AppStateBegin()
    If gStateSaved Then Exit Sub
    On Error Resume Next
    gOrigScreenUpdating = Application.ScreenUpdating
    gOrigEnableEvents = Application.EnableEvents
    gOrigCalculation = Application.Calculation
    gOrigDisplayStatusBar = Application.DisplayStatusBar
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    gStateSaved = True
    On Error GoTo 0
End Sub

Private Sub AppStateEnd()
    If Not gStateSaved Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = gOrigScreenUpdating
    Application.EnableEvents = gOrigEnableEvents
    Application.Calculation = gOrigCalculation
    Application.DisplayStatusBar = gOrigDisplayStatusBar
    gStateSaved = False
    On Error GoTo 0
End Sub

Private Function StripHtmlArtifacts(ByVal s As String) As String
    If Len(s) = 0 Then StripHtmlArtifacts = "": Exit Function
    Dim re As Object
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True: re.MultiLine = True
    
    re.Pattern = "<style[^>]*>[\s\S]*?</style>"
    s = re.Replace(s, "")
    re.Pattern = "<script[^>]*>[\s\S]*?</script>"
    s = re.Replace(s, "")
    
    re.Pattern = "</td[^>]*>"
    s = re.Replace(s, vbTab)
    re.Pattern = "</th[^>]*>"
    s = re.Replace(s, vbTab)
    re.Pattern = "</tr[^>]*>"
    s = re.Replace(s, vbCrLf)

    re.Pattern = "<[^>]+>"
    s = re.Replace(s, "")
    On Error GoTo 0

    s = Replace(s, "_x000D_", ""): s = Replace(s, "_x000d_", "")
    s = Replace(s, "&nbsp;", " "): s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "<"): s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """"): s = Replace(s, "&#39;", "'")
    
    Do While InStr(s, vbCrLf & vbCrLf & vbCrLf) > 0
        s = Replace(s, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop
    StripHtmlArtifacts = s
End Function

Private Function NormalizeOrderNo(ByVal s As String) As String
    Dim r As String: r = Trim$(Replace(s, " ", ""))
    r = Replace(r, vbTab, ""): r = Replace(r, "-", "")
    Do While Len(r) > 1 And Left$(r, 1) = "0"
        r = Mid$(r, 2)
    Loop
    NormalizeOrderNo = r
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
    If colNum < 1 Then NumToColLetter = "": Exit Function
    n = colNum
    Do While n > 0
        rr = (n - 1) Mod 26
        s = Chr$(65 + rr) & s
        n = (n - 1) \ 26
    Loop
    NumToColLetter = s
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
    s = Replace(s, "@", " "): s = Replace(s, """", " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeName = Trim$(s)
End Function

Private Function GetFirstNthWord(ByVal s As String, ByVal n As Long) As String
    Dim parts() As String, i As Long, result As String, count As Long
    s = NormalizeName(s)
    parts = Split(s, " ")
    result = "": count = 0
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            If Len(result) > 0 Then result = result & " "
            result = result & parts(i): count = count + 1
            If count >= n Then Exit For
        End If
    Next i
    GetFirstNthWord = Trim$(result)
End Function

Private Function LookupColByHeader(ByVal ws As Worksheet, ByVal headers As Variant) As Long
    Dim lastCol As Long, c As Long, i As Long
    Dim cv As String, cvClean As String, hvClean As String
    On Error Resume Next
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    On Error GoTo 0
    If lastCol < 1 Then lastCol = 1
    For c = 1 To lastCol
        cv = SafeStr(ws.Cells(HEADER_ROW, c).Value)
        If Len(cv) > 0 Then
            cvClean = CleanForCompare(cv)
            For i = LBound(headers) To UBound(headers)
                hvClean = CleanForCompare(CStr(headers(i)))
                If cvClean = hvClean Then LookupColByHeader = c: Exit Function
            Next i
        End If
    Next c
    For c = 1 To lastCol
        cv = SafeStr(ws.Cells(HEADER_ROW, c).Value)
        If Len(cv) > 0 Then
            cvClean = CleanForCompare(cv)
            For i = LBound(headers) To UBound(headers)
                hvClean = CleanForCompare(CStr(headers(i)))
                If Len(hvClean) >= 4 Then
                    If InStr(1, cvClean, hvClean, vbTextCompare) > 0 Then
                        LookupColByHeader = c: Exit Function
                    End If
                End If
            Next i
        End If
    Next c
End Function

Public Sub IslenenMailleriTemizle()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PROCESSED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox T("IslenenMailler sayfasi bulunamadi."), vbInformation, AppTitle()
        Exit Sub
    End If
    ws.Cells.Clear
    EnsureProcessedSheet
    Set gProcessedDict = Nothing
    MsgBox T("Islenen mailler listesi temizlendi. Tum mailler yeniden islenecek."), vbInformation, AppTitle()
End Sub

Public Sub GelenYanitlariOku()
    Dim cfg As Object
    Set gProcessedDict = Nothing
    Set gInternalDomains = Nothing
    gSiparisYokCount = 0
    gBelirsizCount = 0
    gStateSaved = False

    On Error GoTo FatalError
    Set cfg = GetConfig()
    cfg.RemoveAll
    If Not LoadOrCollectSettings(cfg) Then Exit Sub
    If Not ValidateSchema(cfg) Then Exit Sub
    GeminiPreflight
    DoReadReplies cfg
    Exit Sub
FatalError:
    AppStateEnd
    MsgBox T("Kritik Hata (Satir ") & Erl & "):" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           T("Lutfen Excel sayfalarini ve sutun basliklarini kontrol edin."), vbCritical, AppTitle()
    WriteLog "FATAL", "", "", "", "GelenYanitlariOku", Err.Description
End Sub

Private Function ValidateSchema(ByRef cfg As Object) As Boolean
    Dim msgs As String
    If cfg("VendorNameCol") = cfg("VendorEmailCol") And cfg("VendorNameCol") > 0 Then
        msgs = msgs & "- " & T("Tedarikci ADI ve MAIL sutunlari ayni: ") & _
               NumToColLetter(cfg("VendorNameCol")) & vbCrLf
    End If
    If Len(msgs) > 0 Then
        If MsgBox(T("Asagidaki sorunlar tespit edildi:") & vbCrLf & vbCrLf & msgs & vbCrLf & _
                  T("Devam edilsin mi?"), vbYesNo + vbExclamation, AppTitle()) <> vbYes Then
            ValidateSchema = False
            Exit Function
        End If
    End If
    ValidateSchema = True
End Function

Private Function LoadOrCollectSettings(ByRef cfg As Object) As Boolean
    Dim answer As VbMsgBoxResult
    If LoadSettingsFromSheet(cfg) Then
        SaveSettingsToSheet cfg
        answer = MsgBox(T("Kay~itl~i ayarlar y~uklendi:") & vbCrLf & vbCrLf & _
                        GetSettingsSummary(cfg) & vbCrLf & vbCrLf & _
                        T("Bu ayarlarla devam edilsin mi?"), _
                        vbYesNoCancel + vbQuestion, AppTitle())
        If answer = vbCancel Then Exit Function
        If answer = vbYes Then LoadOrCollectSettings = True: Exit Function
    End If
    cfg.RemoveAll
    If Not CollectSettings(cfg) Then Exit Function
    SaveSettingsToSheet cfg
    LoadOrCollectSettings = True
End Function

Private Function LoadSettingsFromSheet(ByRef cfg As Object, Optional ByVal skipAutoFix As Boolean = False) As Boolean
    Dim ws As Worksheet, s As String
    Dim orderSheets As Collection, wsOrder As Worksheet, wsVendor As Worksheet
    Dim i As Long, parts() As String, sname As String, vName As String
    Dim wasAutoFixed As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET): On Error GoTo 0
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
            If wsOrder Is Nothing Then
                MsgBox T("Ayarlar sayfasinda belirtilen SIPARIS SAYFASI bulunamadi: ") & sname, vbExclamation, AppTitle()
                Exit Function
            End If
            orderSheets.Add wsOrder
        End If
    Next i
    If orderSheets.Count = 0 Then Exit Function

    vName = SafeStr(ws.Range("B2").Value)
    If Len(vName) = 0 Then Exit Function
    On Error Resume Next
    Set wsVendor = ThisWorkbook.Worksheets(vName)
    On Error GoTo CleanFail
    If wsVendor Is Nothing Then
        MsgBox T("Ayarlar sayfasinda belirtilen TEDARIKCI SAYFASI bulunamadi: ") & vName, vbExclamation, AppTitle()
        Exit Function
    End If

    Set cfg("OrderSheets") = orderSheets
    Set cfg("VendorWs") = wsVendor
    cfg("StartRow") = CLng(Val(ws.Range("B3").Value))
    If cfg("StartRow") < 2 Then cfg("StartRow") = DEFAULT_START_ROW
    cfg("InboxFolder") = SafeStr(ws.Range("B4").Value)
    cfg("InternalDomains") = SafeStr(ws.Range("B5").Value)
    cfg("UpdateAllLines") = UCase$(SafeStr(ws.Range("B6").Value))
    If cfg("UpdateAllLines") <> "EVET" Then cfg("UpdateAllLines") = "HAYIR"
    cfg("TerminCol") = 0

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
    cfg("DurumCol") = ColLetterToNum(SafeStr(ws.Range("B27").Value))
    Dim b28 As String: b28 = SafeStr(ws.Range("B28").Value)
    If Len(b28) > 0 Then cfg("TerminCol") = ColLetterToNum(b28)

    If Not skipAutoFix Then
        wasAutoFixed = AutoFixAllColumns(cfg, orderSheets(1), wsVendor)
        If wasAutoFixed Then
            WriteLog "AYAR_DUZELT", "", "", "", "Ayarlar sayfasi otomatik duzeltildi."
        End If
    End If
    LoadSettingsFromSheet = True
    Exit Function
CleanFail:
    LoadSettingsFromSheet = False
End Function

Private Function AutoFixAllColumns(ByRef cfg As Object, ByVal ws As Worksheet, ByVal vws As Worksheet) As Boolean
    Dim changed As Boolean: changed = False
    Dim newVal As Long

    newVal = LookupColByHeader(ws, Array("Satýnalma belgesi", "Satin alma", "Sipariþ No", "Siparis No", "SIPARIS_NO"))
    If newVal < 1 Then newVal = LookupColByHeader(ws, Array("SIPARIS_NO"))
    If newVal > 0 And cfg("OrderNoCol") <> newVal Then cfg("OrderNoCol") = newVal: changed = True

    newVal = LookupColByHeader(ws, Array("Kalem No", "KalemNo", "Item No", "SAT_KALEM_NO"))
    If newVal < 1 Then newVal = LookupColByHeader(ws, Array("KALEM_NO"))
    If newVal > 0 And cfg("KalemNoCol") <> newVal Then cfg("KalemNoCol") = newVal: changed = True

    newVal = LookupColByHeader(ws, Array("Tedarikçi Adý", "Tedarikçi Ünvaný", "Tedarikçi", "TEDARIKCI_ADI", "Satýcý Ýsmi", "Firma Adý"))
    If newVal < 1 Then newVal = LookupColByHeader(ws, Array("TEDARIKCI"))
    If newVal > 0 And cfg("VendorCol") <> newVal Then cfg("VendorCol") = newVal: changed = True

    newVal = LookupColByHeader(ws, Array("Malzeme", "Malzeme Kodu", "Stok Kodu", "DETAY MALZEME", "DETAY_MALZEME"))
    If newVal > 0 And cfg("MalzemeCol") <> newVal Then cfg("MalzemeCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Üretim yeri", "Uretim yeri", "DETAY_UY"))
    If newVal > 0 And cfg("UretimYeriCol") <> newVal Then cfg("UretimYeriCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Malzeme Uzun Metni", "Uzun Metin", "DETAY_MALZ_METNI", "DETAY_MALZ_AL_METNI"))
    If newVal > 0 And cfg("UzunMetinCol") <> newVal Then cfg("UzunMetinCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Temel ölçü birimi", "Birim", "TEMEL_OLCU_BIRIM"))
    If newVal > 0 And cfg("BirimCol") <> newVal Then cfg("BirimCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Bakiye", "Miktar", "DETAY_MIKTAR"))
    If newVal > 0 And cfg("BakiyeCol") <> newVal Then cfg("BakiyeCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Mail Atýldý", "Mail Atildi", "Mail Durumu"))
    If newVal > 0 And cfg("StatusCol") <> newVal Then cfg("StatusCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Gelen Yanýt", "Gelen Yanit", "Cevap"))
    If newVal > 0 And cfg("ReplyCol") <> newVal Then cfg("ReplyCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Yanýt Tarihi", "Yanit Tarihi"))
    If newVal > 0 And cfg("ReplyDateCol") <> newVal Then cfg("ReplyDateCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Yanýtlayan E-posta", "Yanýtlayan"))
    If newVal > 0 And cfg("ReplyFromCol") <> newVal Then cfg("ReplyFromCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Eþleþme Yöntemi", "Eþleþme"))
    If newVal > 0 And cfg("MatchMethodCol") <> newVal Then cfg("MatchMethodCol") = newVal: changed = True
    newVal = LookupColByHeader(ws, Array("Ekler", "Ek"))
    If newVal > 0 And cfg("AttachCol") <> newVal Then cfg("AttachCol") = newVal: changed = True
    
    newVal = LookupColByHeader(ws, Array("Sipariþ Durumu", "Siparis Durumu", "Durum", "DETAY_STATU"))
    If newVal > 0 Then
        If cfg("DurumCol") <> newVal Then cfg("DurumCol") = newVal: changed = True
    Else
        Dim dC As Long: dC = EnsureOrCreateColumn(ws, T("Sipari~s Durumu"))
        If cfg("DurumCol") <> dC Then cfg("DurumCol") = dC: changed = True
    End If

    newVal = LookupColByHeader(ws, Array("Termin", "Termin Tarihi", "Teslim Tarihi"))
    If newVal > 0 Then
        If cfg("TerminCol") <> newVal Then cfg("TerminCol") = newVal: changed = True
    Else
        Dim tC As Long: tC = EnsureOrCreateColumn(ws, T("Termin"))
        If cfg("TerminCol") <> tC Then cfg("TerminCol") = tC: changed = True
    End If

    newVal = LookupColByHeader(vws, Array( _
        "Satýcý Unvaný", "Satýcý Ünvaný", "Satýcý Ýsmi", "Satýcý", _
        "Tedarikçi Ünvaný", "Tedarikçi Adý", "Tedarikçi", _
        "Firma Adý", "Firma", "Vendor Name", "Ad 1"))
    If newVal > 0 And cfg("VendorNameCol") <> newVal Then cfg("VendorNameCol") = newVal: changed = True
    newVal = LookupColByHeader(vws, Array("ana e-mail adresi", "ana e-mail", "e-posta", "email", "mail adresi"))
    If newVal > 0 And cfg("VendorEmailCol") <> newVal Then cfg("VendorEmailCol") = newVal: changed = True

    If cfg("VendorNameCol") = cfg("VendorEmailCol") Or cfg("VendorNameCol") < 1 Then
        Dim aHeader As String: aHeader = SafeStr(vws.Cells(HEADER_ROW, 1).Value)
        If Len(aHeader) > 0 And InStr(aHeader, "@") = 0 Then
            cfg("VendorNameCol") = 1
            changed = True
        End If
    End If

    If cfg("VendorNameCol") < 1 Then cfg("VendorNameCol") = ColLetterToNum(DEF_COL_VENDOR_NAME)
    If cfg("VendorEmailCol") < 1 Then cfg("VendorEmailCol") = ColLetterToNum(DEF_COL_VENDOR_EMAIL)
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
    If cfg("DurumCol") < 1 Then cfg("DurumCol") = ColLetterToNum(DEF_COL_DURUM)
    If cfg("TerminCol") < 1 Then cfg("TerminCol") = ColLetterToNum(DEF_COL_TERMIN)

    AutoFixAllColumns = changed
End Function

Private Sub SaveSettingsToSheet(ByRef cfg As Object)
    Dim ws As Worksheet, orderSheets As Collection, s As String, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add: ws.Name = CONFIG_SHEET
    On Error GoTo 0
    ws.Cells.Clear
    ws.Range("A1").Value = T("S~IPAR~I~S SAYFALARI"): ws.Range("B1").Value = ""
    Set orderSheets = cfg("OrderSheets"): s = ""
    For i = 1 To orderSheets.Count
        If Len(s) > 0 Then s = s & ","
        s = s & orderSheets(i).Name
    Next i
    ws.Range("B1").Value = s
    ws.Range("A2").Value = T("TEDAR~IK~C~I SAYFASI"): ws.Range("B2").Value = cfg("VendorWs").Name
    ws.Range("A3").Value = T("BA~SLANGI~C SATIRI"):  ws.Range("B3").Value = cfg("StartRow")
    ws.Range("A4").Value = T("OUTLOOK KLAS~OR"):      ws.Range("B4").Value = cfg("InboxFolder")
    ws.Range("A5").Value = T("DAH~IL~I DOMA~INLER"): ws.Range("B5").Value = SafeStr(cfg("InternalDomains"))
    
    ws.Range("A6").Value = T("E~SLE~SMEYENLER~I T~UM SATIRLARA UYGULA (EVET/HAYIR)"): ws.Range("B6").Value = SafeStr(cfg("UpdateAllLines"))
    
    ws.Range("A7").Value = T("Sipari~s No"):          ws.Range("B7").Value = NumToColLetter(cfg("OrderNoCol"))
    ws.Range("A8").Value = T("Kalem No"):             ws.Range("B8").Value = NumToColLetter(cfg("KalemNoCol"))
    ws.Range("A9").Value = T("Tedarik~ci"):           ws.Range("B9").Value = NumToColLetter(cfg("VendorCol"))
    ws.Range("A10").Value = T("Malzeme"):             ws.Range("B10").Value = NumToColLetter(cfg("MalzemeCol"))
    ws.Range("A11").Value = T("~Uretim Yeri"):        ws.Range("B11").Value = NumToColLetter(cfg("UretimYeriCol"))
    ws.Range("A12").Value = T("Uzun Metin"):          ws.Range("B12").Value = NumToColLetter(cfg("UzunMetinCol"))
    ws.Range("A13").Value = T("Birim"):               ws.Range("B13").Value = NumToColLetter(cfg("BirimCol"))
    ws.Range("A14").Value = T("Bakiye"):              ws.Range("B14").Value = NumToColLetter(cfg("BakiyeCol"))
    ws.Range("A15").Value = T("Mail At~ild~i"):       ws.Range("B15").Value = NumToColLetter(cfg("StatusCol"))
    ws.Range("A16").Value = T("Gelen Yan~it"):        ws.Range("B16").Value = NumToColLetter(cfg("ReplyCol"))
    ws.Range("A17").Value = T("Yan~it Tarihi"):       ws.Range("B17").Value = NumToColLetter(cfg("ReplyDateCol"))
    ws.Range("A18").Value = T("Yan~itlayan"):         ws.Range("B18").Value = NumToColLetter(cfg("ReplyFromCol"))
    ws.Range("A19").Value = T("E~sle~sme"):           ws.Range("B19").Value = NumToColLetter(cfg("MatchMethodCol"))
    ws.Range("A20").Value = T("Ekler"):               ws.Range("B20").Value = NumToColLetter(cfg("AttachCol"))
    ws.Range("A23").Value = T("Sat~ic~i ~Ismi"):      ws.Range("B23").Value = NumToColLetter(cfg("VendorNameCol"))
    ws.Range("A24").Value = T("Mail Adresi"):         ws.Range("B24").Value = NumToColLetter(cfg("VendorEmailCol"))
    ws.Range("A26").Value = T("TEST MODU (EVET/HAYIR)")
    If IsTestMode() Then ws.Range("B26").Value = "EVET" Else ws.Range("B26").Value = "HAYIR"
    If cfg("DurumCol") > 0 Then
        ws.Range("A27").Value = T("Durum S~utunu"):   ws.Range("B27").Value = NumToColLetter(cfg("DurumCol"))
    End If
    If cfg("TerminCol") > 0 Then
        ws.Range("A28").Value = T("Termin S~utunu"):  ws.Range("B28").Value = NumToColLetter(cfg("TerminCol"))
    End If
    ws.Columns("A:A").ColumnWidth = 55: ws.Columns("B:B").ColumnWidth = 25
    On Error Resume Next: ws.Visible = xlSheetHidden: On Error GoTo 0
End Sub

Private Function CollectSettings(ByRef cfg As Object) As Boolean
    Dim orderSheets As Collection, ws As Worksheet
    Dim sheetNames As String, parts() As String
    Dim i As Long, sname As String, answer As VbMsgBoxResult
    Dim vws As Worksheet, startRow As Long, inboxFolder As String, sampleWs As Worksheet
    Set orderSheets = New Collection
    sheetNames = InputBox(T("S~IPAR~I~S SAYFALARINI girin (virg~ulle)") & vbCrLf & ListSheetNames(), _
                          T("Otomasyon"), T("S~if~ir Stoklu"))
    If StrPtr(sheetNames) = 0 Then Exit Function
    sheetNames = Trim$(sheetNames)
    If Len(sheetNames) = 0 Then Exit Function
    parts = Split(sheetNames, ",")
    For i = LBound(parts) To UBound(parts)
        sname = Trim$(parts(i))
        If Len(sname) > 0 Then
            Set ws = Nothing
            On Error Resume Next: Set ws = ThisWorkbook.Worksheets(sname): On Error GoTo 0
            If ws Is Nothing Then MsgBox T("Sayfa bulunamad~i: ") & sname, vbExclamation, AppTitle(): Exit Function
            orderSheets.Add ws
        End If
    Next i
    If orderSheets.Count = 0 Then Exit Function
    Set vws = AskWorksheet(T("TEDAR~IK~C~I SAYFASI?"), "Sheet1"): If vws Is Nothing Then Exit Function
    startRow = AskStartRow(): If startRow = 0 Then Exit Function
    inboxFolder = InputBox(T("Outlook klas~or (bo~s = Gelen Kutusu):"), T("Otomasyon"), "")
    If StrPtr(inboxFolder) = 0 Then Exit Function
    Set sampleWs = orderSheets(1)
    Set cfg("OrderSheets") = orderSheets
    Set cfg("VendorWs") = vws
    cfg("StartRow") = startRow
    cfg("InboxFolder") = Trim$(inboxFolder)
    cfg("DurumCol") = 0: cfg("TerminCol") = 0
    cfg("InternalDomains") = ""
    cfg("UpdateAllLines") = "EVET"
    AutoFixAllColumns cfg, sampleWs, vws
    answer = MsgBox(T("Alg~ilanan ayarlar:") & vbCrLf & vbCrLf & GetSettingsSummary(cfg) & vbCrLf & vbCrLf & _
                    T("Devam edilsin mi?"), vbYesNo + vbQuestion, AppTitle())
    If answer <> vbYes Then Exit Function
    CollectSettings = True
End Function

Private Function GetSettingsSummary(ByRef cfg As Object) As String
    Dim s As String, i As Long, orderSheets As Collection
    Set orderSheets = cfg("OrderSheets")
    s = T("SAYFALAR: ") & orderSheets.Count & vbCrLf
    For i = 1 To orderSheets.Count: s = s & "  - " & orderSheets(i).Name & vbCrLf: Next i
    s = s & T("TEDAR~IK~C~I: ") & cfg("VendorWs").Name & vbCrLf
    s = s & T("Tedarik~ciAd~i=") & NumToColLetter(cfg("VendorNameCol")) & _
        ", " & T("Tedarik~ciMail=") & NumToColLetter(cfg("VendorEmailCol")) & vbCrLf
    s = s & T("Siparis=") & NumToColLetter(cfg("OrderNoCol")) & _
        ", " & T("Malzeme=") & NumToColLetter(cfg("MalzemeCol")) & _
        ", " & T("Yanit=") & NumToColLetter(cfg("ReplyCol")) & _
        ", " & T("Yanitlayan=") & NumToColLetter(cfg("ReplyFromCol")) & _
        ", " & T("Eslesme=") & NumToColLetter(cfg("MatchMethodCol")) & vbCrLf
    s = s & T("Durum=") & NumToColLetter(cfg("DurumCol")) & _
        ", " & T("Termin=") & NumToColLetter(cfg("TerminCol")) & vbCrLf
    s = s & T("E~sle~smeyenleri T~um Sat~irlara Uygula: ") & SafeStr(cfg("UpdateAllLines")) & vbCrLf
    s = s & T("TEST MODU: ") & IIf(IsTestMode(), T("EVET (ilk 5)"), T("HAYIR (tam)"))
    GetSettingsSummary = s
End Function

Private Sub EnsureProcessedSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PROCESSED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PROCESSED_SHEET
        ws.Range("A1").Value = "Mail EntryID"
        ws.Range("B1").Value = T("Islem Tarihi")
        ws.Range("C1").Value = T("Tedarik~ci")
        ws.Range("A1:C1").Font.Bold = True
        ws.Range("A1:C1").Interior.Color = RGB(220, 220, 220)
        ws.Columns("A:C").ColumnWidth = 40
        ws.Visible = xlSheetHidden
        On Error GoTo 0
    End If
End Sub

Private Function GetProcessedMails() As Object
    If gProcessedDict Is Nothing Then
        Set gProcessedDict = CreateObject("Scripting.Dictionary")
        gProcessedDict.CompareMode = vbTextCompare
        EnsureProcessedSheet
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(PROCESSED_SHEET)
        On Error GoTo 0
        If Not ws Is Nothing Then
            Dim r As Long, lastR As Long
            lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            For r = 2 To lastR
                Dim key As String: key = SafeStr(ws.Cells(r, 1).Value)
                If Len(key) > 0 And Not gProcessedDict.Exists(key) Then gProcessedDict.Add key, True
            Next r
        End If
    End If
    Set GetProcessedMails = gProcessedDict
End Function

Private Sub MarkMailProcessed(ByVal entryId As String, ByVal vendorName As String)
    If Len(entryId) = 0 Then Exit Sub
    Dim dict As Object: Set dict = GetProcessedMails()
    If dict.Exists(entryId) Then Exit Sub
    dict.Add entryId, True
    EnsureProcessedSheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PROCESSED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim nr As Long: nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nr < 2 Then nr = 2
    ws.Cells(nr, 1).Value = entryId
    ws.Cells(nr, 2).Value = Format(Now, "dd.mm.yyyy hh:nn:ss")
    ws.Cells(nr, 3).Value = vendorName
End Sub

Private Function LoadInternalDomains(ByRef cfg As Object, ByVal olNs As Object) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim cfgVal As String: cfgVal = SafeStr(cfg("InternalDomains"))
    If Len(cfgVal) > 0 Then
        Dim parts() As String, i As Long
        parts = Split(cfgVal, ",")
        For i = LBound(parts) To UBound(parts)
            Dim d As String: d = LCase$(Trim$(parts(i)))
            If Len(d) > 0 Then
                If InStr(d, "@") > 0 Then d = Mid$(d, InStr(d, "@") + 1)
                dict(d) = True
            End If
        Next i
    End If
    dict("assanaluminyum.com") = True
    dict("kibar.com") = True
    Set LoadInternalDomains = dict
End Function

Private Function IsInternalDomain(ByVal domain As String) As Boolean
    If Len(domain) = 0 Or gInternalDomains Is Nothing Then Exit Function
    If gInternalDomains.Exists(LCase$(domain)) Then IsInternalDomain = True
End Function

Private Sub DoReadReplies(ByRef cfg As Object)
    Dim vws As Worksheet
    Dim vendorNameCol As Long, vendorEmailCol As Long
    Dim orderSheets As Collection
    Dim olApp As Object, olNs As Object, inbox As Object, mail As Object
    Dim vendorByEmail As Object, vendorByDomain As Object
    Dim vLastRow As Long, r As Long
    Dim vname As String, vemail As String, vdomain As String
    Dim matchCount As Long, replyCount As Long, scanCount As Long, skipCount As Long
    Dim internalSkipCount As Long, processedSkipCount As Long, dateSkipCount As Long
    Dim i As Long, cutoffDate As Date, mailDate As Date
    Dim senderEmail As String, senderName As String, mailSubject As String, mailBody As String
    Dim senderDomain As String, vendorName As String, matchMethod As String
    Dim newReply As String, attachmentList As String
    Dim sheetUpdated As Long, upd As Long
    Dim tMode As Boolean: tMode = IsTestMode()

    If tMode Then
        If MsgBox(T("TEST MODU ACIK: Ilk ") & TEST_MAX_ORDERS & T(" satir guncellenecek.") & vbCrLf & _
                  T("Devam edilsin mi?"), vbYesNo + vbQuestion, AppTitle()) <> vbYes Then Exit Sub
    End If

    AppStateBegin
    Set vws = cfg("VendorWs")
    vendorNameCol = cfg("VendorNameCol"): vendorEmailCol = cfg("VendorEmailCol")
    Set orderSheets = cfg("OrderSheets")

    WriteLog "AYARLAR", "", "", "", _
        "VendorAd=" & NumToColLetter(vendorNameCol) & _
        " | VendorMail=" & NumToColLetter(vendorEmailCol) & _
        " | Siparis=" & NumToColLetter(cfg("OrderNoCol")) & _
        " | Malzeme=" & NumToColLetter(cfg("MalzemeCol")) & _
        " | Yanit=" & NumToColLetter(cfg("ReplyCol")) & _
        " | Yanitlayan=" & NumToColLetter(cfg("ReplyFromCol")) & _
        " | Eslesme=" & NumToColLetter(cfg("MatchMethodCol")) & _
        " | Durum=" & NumToColLetter(cfg("DurumCol")) & _
        " | Termin=" & NumToColLetter(cfg("TerminCol")) & _
        " | UpdateAll=" & SafeStr(cfg("UpdateAllLines"))

    If vendorNameCol = vendorEmailCol Then
        WriteLog "UYARI", "", "", "", "Tedarikci Adi ve Mail sutunlari ayni (" & _
                 NumToColLetter(vendorNameCol) & ") - buyuk olasilikla yanlis ayar!"
        AppStateEnd
        MsgBox T("KRITIK: Tedarikci Adi ve Mail sutunlari ayni (") & NumToColLetter(vendorNameCol) & ")." & vbCrLf & vbCrLf & _
               T("Ayarlar sayfasinda 'Satici Ismi' (B23) ve 'Mail Adresi' (B24) sutunlarini kontrol edin."), _
               vbCritical, AppTitle()
        Exit Sub
    End If

    vLastRow = vws.Cells(vws.Rows.Count, vendorNameCol).End(xlUp).Row
    Set vendorByEmail = CreateObject("Scripting.Dictionary")
    Set vendorByDomain = CreateObject("Scripting.Dictionary")
    vendorByEmail.CompareMode = vbTextCompare
    vendorByDomain.CompareMode = vbTextCompare

    Dim loadedCount As Long: loadedCount = 0
    For r = 2 To vLastRow
        vname = Trim$(SafeStr(vws.Cells(r, vendorNameCol).Value))
        Dim rawEmail As String: rawEmail = SafeStr(vws.Cells(r, vendorEmailCol).Value)
        If InStr(rawEmail, ";") > 0 Then rawEmail = Split(rawEmail, ";")(0)
        If InStr(rawEmail, ",") > 0 Then rawEmail = Split(rawEmail, ",")(0)
        vemail = LCase$(Trim$(rawEmail))
        If InStr(UCase$(vname), "BLOKE") > 0 Or InStr(UCase$(vname), "BLOCKED") > 0 Then GoTo NextVendorRow
        If InStr(vname, "#") > 0 And (InStr(vname, "#N/A") > 0 Or InStr(vname, "#REF") > 0) Then GoTo NextVendorRow
        If InStr(vemail, "#N/A") > 0 Or InStr(vemail, "#REF") > 0 Then GoTo NextVendorRow
        If Len(vname) > 0 And InStr(vemail, "@") > 0 Then
            If Not vendorByEmail.Exists(vemail) Then vendorByEmail.Add vemail, vname
            vdomain = Mid$(vemail, InStr(vemail, "@") + 1)
            If Len(vdomain) > 0 Then
                If vendorByDomain.Exists(vdomain) Then
                    If vendorByDomain(vdomain) <> "***CAKISMA***" And _
                       LCase$(CStr(vendorByDomain(vdomain))) <> LCase$(vname) Then
                        vendorByDomain(vdomain) = "***CAKISMA***"
                    End If
                Else
                    vendorByDomain.Add vdomain, vname
                End If
            End If
            loadedCount = loadedCount + 1
        End If
NextVendorRow:
    Next r
    WriteLog "TEDARIKCI_YUKLE", "", "", "", "Yuklenen tedarikci: " & loadedCount & _
             " / toplam satir: " & (vLastRow - 1)

    Set olApp = GetOutlookApp()
    If olApp Is Nothing Then
        AppStateEnd
        MsgBox T("Outlook'a ba~glan~ilamad~i."), vbCritical, AppTitle()
        Exit Sub
    End If
    Set olNs = olApp.GetNamespace("MAPI")
    Set inbox = GetInboxFolder(olNs, cfg("InboxFolder"))
    If inbox Is Nothing Then
        AppStateEnd
        MsgBox T("Gelen kutusu bulunamad~i."), vbCritical, AppTitle()
        Exit Sub
    End If

    Set gInternalDomains = LoadInternalDomains(cfg, olNs)
    EnsureTerminReportSheet
    cutoffDate = DateAdd("d", -MAX_LOOKBACK_DAYS, Now)
    WriteLog "BASLAT", "", "", "", "GelenYanitlariOku", _
             "Inbox: " & inbox.Name & " | Toplam: " & inbox.Items.Count

    For Each mail In inbox.Items
        scanCount = scanCount + 1
        Dim entryId As String: entryId = ""
        On Error Resume Next
        entryId = SafeStr(mail.EntryID)
        On Error GoTo 0
        If Len(entryId) > 0 Then
            If GetProcessedMails().Exists(entryId) Then
                processedSkipCount = processedSkipCount + 1
                skipCount = skipCount + 1
                GoTo NextMail
            End If
        End If

        On Error Resume Next
        mailDate = mail.ReceivedTime
        If Err.Number <> 0 Then Err.Clear: dateSkipCount = dateSkipCount + 1: skipCount = skipCount + 1: GoTo NextMail
        If mailDate < cutoffDate Then dateSkipCount = dateSkipCount + 1: skipCount = skipCount + 1: GoTo NextMail
        senderEmail = GetSenderSmtpEx(mail)
        senderName = mail.SenderName
        mailSubject = mail.Subject
        
        mailBody = mail.HTMLBody
        If Err.Number <> 0 Or Len(mailBody) = 0 Then
            Err.Clear
            mailBody = mail.Body
        End If
        If Err.Number <> 0 Then
            WriteLog "HATA", "", "", "", "Mail Okuma", Err.Description
            Err.Clear: GoTo NextMail
        End If
        On Error GoTo 0

        mailBody = StripHtmlArtifacts(mailBody)
        senderDomain = ""
        If InStr(senderEmail, "@") > 0 Then senderDomain = Mid$(senderEmail, InStr(senderEmail, "@") + 1)

        If IsInternalDomain(senderDomain) Then
            internalSkipCount = internalSkipCount + 1
            skipCount = skipCount + 1
            GoTo NextMail
        End If

        vendorName = "": matchMethod = ""
        If Len(senderEmail) > 0 And vendorByEmail.Exists(senderEmail) Then
            vendorName = vendorByEmail(senderEmail): matchMethod = "e-posta"
        End If
        If Len(vendorName) = 0 And Len(senderDomain) > 0 Then
            If vendorByDomain.Exists(senderDomain) Then
                If vendorByDomain(senderDomain) <> "***CAKISMA***" Then
                    vendorName = vendorByDomain(senderDomain): matchMethod = "domain"
                End If
            End If
        End If
        If Len(vendorName) = 0 Then
            If IsSiparisKonusu(mailSubject) Or IsSiparisKonusu(mailBody) Then
                Dim subState As String
                vendorName = FindVendorInSenderName(vws, vLastRow, vendorNameCol, senderName, subState)
                If Len(vendorName) > 0 Then matchMethod = "gonderen-adi (" & subState & ")"
            End If
        End If

        If Len(vendorName) > 0 And matchMethod <> "e-posta" Then
            WriteLog "GEMINI_GONDEREN_BELIRSIZ", "", vendorName, senderEmail, "Tedarikci listesine tam e-posta adresini ekleyin."
            GoTo NextMail
        End If
        If Len(vendorName) > 0 Then
            replyCount = replyCount + 1
            newReply = ExtractNewReply(mailBody)
            If Len(newReply) = 0 Then newReply = "(Bos yanit)"
            attachmentList = SaveAttachments(mail, vendorName)
            sheetUpdated = 0
            For i = 1 To orderSheets.Count
                upd = WriteReplyToSheet(cfg, orderSheets(i), vendorName, newReply, _
                    mailDate, senderEmail, matchMethod, attachmentList, tMode)
                sheetUpdated = sheetUpdated + upd
            Next i
            matchCount = matchCount + sheetUpdated
            If sheetUpdated > 0 And Len(entryId) > 0 Then MarkMailProcessed entryId, vendorName
            WriteLog "YANIT", "", vendorName, senderEmail, matchMethod, _
                     "Konu: " & Left$(mailSubject, 80) & " | Guncellenen: " & sheetUpdated
        Else
            If IsSiparisKonusu(mailSubject) Then
                WriteLog "ESLESMEDI", "", senderName, senderEmail, _
                         "Konu: " & Left$(mailSubject, 80), "Tedarikci bulunamadi"
            End If
        End If
NextMail:
    Next mail

    WriteLog "BITIR", "", "", "", "GelenYanitlariOku", _
             "Taranan: " & scanCount & ", Islenen: " & replyCount & _
             ", Guncellenen: " & matchCount & ", Atlanan: " & skipCount
    If gSiparisYokCount > 0 Then
        WriteLog "OZET", "", "", "", "SIPARIS_YOK (acik siparisi olmayan tedarikci sayisi): " & gSiparisYokCount
    End If
    If gBelirsizCount > 0 Then
        WriteLog "OZET", "", "", "", "BELIRSIZ (e~sle~smeyen): " & gBelirsizCount
    End If

    WriteTerminReport cfg, scanCount, replyCount, matchCount
    AppStateEnd

    MsgBox T("Okuma tamamland~i.") & vbCrLf & vbCrLf & _
           T("Taranan mail: ") & scanCount & vbCrLf & _
           T("E~sle~stirilen yan~it: ") & replyCount & vbCrLf & _
           T("Excel'de g~uncellenen sat~ir: ") & matchCount & vbCrLf & _
           T("Belirsiz (e~sle~smeyen): ") & gBelirsizCount & vbCrLf & _
           T("Atlanan (i~slenmi~s): ") & processedSkipCount & vbCrLf & _
           T("Atlanan (tarih): ") & dateSkipCount & vbCrLf & _
           T("Atlanan (dahili): ") & internalSkipCount & vbCrLf & _
           T("A~c~ik sipari~si olmayan tedarik~ci: ") & gSiparisYokCount, _
           vbInformation, AppTitle()
End Sub

Private Function WriteReplyToSheet(ByRef cfg As Object, ByVal ws As Worksheet, _
                                   ByVal vendorName As String, ByVal newReply As String, _
                                   ByVal mailDate As Date, ByVal senderEmail As String, _
                                   ByVal matchMethod As String, ByVal attachmentList As String, _
                                   ByVal tMode As Boolean) As Long
    ' Fail closed: Gemini errors stop the run instead of using legacy fuzzy matching.
    WriteReplyToSheet = GeminiWriteSheet(cfg, ws, vendorName, newReply, mailDate, _
        senderEmail, matchMethod, attachmentList, tMode)
End Function

Private Function EnsureOrCreateColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    On Error Resume Next
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    On Error GoTo 0
    If lastCol < 1 Then lastCol = 1
    For c = 1 To lastCol
        If UCase$(Trim$(SafeStr(ws.Cells(HEADER_ROW, c).Value))) = UCase$(headerText) Then
            EnsureOrCreateColumn = c: Exit Function
        End If
    Next c
    Dim newCol As Long: newCol = lastCol + 1
    On Error Resume Next
    ws.Cells(HEADER_ROW, newCol).Value = headerText
    ws.Cells(HEADER_ROW, newCol).Font.Bold = True
    ws.Cells(HEADER_ROW, newCol).Interior.Color = RGB(200, 220, 240)
    On Error GoTo 0
    EnsureOrCreateColumn = newCol
End Function

Private Function FindRowByCode(ByVal ws As Worksheet, ByVal code As String, _
                                ByVal materialCol As Long, ByVal orderCol As Long, _
                                ByVal itemCol As Long, ByVal startRow As Long, ByVal lastRow As Long) As Long
    Dim r As Long, cv As String, cvItem As String
    Dim sipNo As String, kalemNo As String
    Dim parts() As String
    
    If InStr(code, "|") > 0 Then
        parts = Split(code, "|")
        sipNo = NormalizeOrderNo(parts(0))
        kalemNo = NormalizeOrderNo(parts(1))
    Else
        sipNo = NormalizeOrderNo(code)
        kalemNo = ""
    End If
    
    If Len(sipNo) < 4 Then Exit Function
    
    If orderCol > 0 And Len(sipNo) >= 6 Then
        For r = startRow To lastRow
            cv = NormalizeOrderNo(CStr(ws.Cells(r, orderCol).Value))
            If cv = sipNo Then
                If Len(kalemNo) > 0 And itemCol > 0 Then
                    cvItem = NormalizeOrderNo(CStr(ws.Cells(r, itemCol).Value))
                    If cvItem = kalemNo Then
                        FindRowByCode = r: Exit Function
                    End If
                Else
                    FindRowByCode = r: Exit Function
                End If
            End If
        Next r
    End If
End Function

Private Function MatchVendor(ByVal normVendor As String, ByVal normRow As String, _
                             ByVal firstVendor As String, ByVal twoVendor As String) As Boolean
    If Len(normRow) = 0 Or Len(normVendor) = 0 Then Exit Function
    If normRow = normVendor Then MatchVendor = True: Exit Function
    If Len(twoVendor) >= 6 Then
        Dim twoRow As String: twoRow = GetFirstNthWord(normRow, 2)
        If Len(twoRow) >= 6 And twoRow = twoVendor Then MatchVendor = True: Exit Function
    End If
    If Len(firstVendor) >= 4 Then
        Dim firstRow As String: firstRow = GetFirstNthWord(normRow, 1)
        If Len(firstRow) >= 4 And firstRow = firstVendor Then MatchVendor = True: Exit Function
    End If
    If Len(normRow) >= 4 And Len(normVendor) >= 4 Then
        If InStr(1, normRow, normVendor, vbTextCompare) > 0 Or _
           InStr(1, normVendor, normRow, vbTextCompare) > 0 Then MatchVendor = True: Exit Function
    End If
End Function

Private Sub WriteReplyToRow(ByVal ws As Worksheet, ByVal rowNum As Long, _
                             ByVal newReply As String, ByVal mailDate As Date, _
                             ByVal senderEmail As String, ByVal matchMethod As String, _
                             ByVal attachmentList As String, _
                             ByVal rBC As Long, ByVal rDC As Long, _
                             ByVal rFC As Long, ByVal rMC As Long, ByVal rAC As Long, _
                             ByVal dCol As Long, ByVal tCol As Long, _
                             ByVal statusCol As Long, ByVal explicitDateText As String)
    Dim existing As String, combined As String
    Dim terminDate As Variant, statusText As String
    If rBC < 1 Or rDC < 1 Then Exit Sub
    existing = SafeStr(ws.Cells(rowNum, rBC).Value)
    combined = AppendReplyHistory(existing, newReply, mailDate)
    If Len(combined) > MAX_BODY_LEN Then combined = Left$(combined, MAX_BODY_LEN)
    ws.Cells(rowNum, rBC).Value = combined
    ws.Cells(rowNum, rDC).Value = mailDate
    If rFC > 0 Then ws.Cells(rowNum, rFC).Value = senderEmail
    If rMC > 0 Then ws.Cells(rowNum, rMC).Value = matchMethod
    If rAC > 0 And Len(attachmentList) > 0 Then ws.Cells(rowNum, rAC).Value = attachmentList

    Dim dateSource As String
    If Len(explicitDateText) > 0 Then dateSource = explicitDateText Else dateSource = newReply
    Err.Clear
    terminDate = ParseTurkishDate(dateSource, mailDate)
    If tCol > 0 Then
        On Error Resume Next
        If Not IsEmpty(terminDate) Then
            If IsDate(terminDate) Then
                ws.Cells(rowNum, tCol).Value = terminDate
                ws.Cells(rowNum, tCol).NumberFormat = "dd.mm.yyyy"
            End If
        End If
        On Error GoTo 0
    End If
    If dCol > 0 Then
        statusText = ClassifyStatus(newReply, Not IsEmpty(terminDate))
        ws.Cells(rowNum, dCol).Value = statusText
    End If
    If statusCol > 0 Then
        Dim curStatus As String
        curStatus = SafeStr(ws.Cells(rowNum, statusCol).Value)
        If InStr(1, curStatus, "+", vbTextCompare) = 0 Then
            If Len(curStatus) = 0 Then
                ws.Cells(rowNum, statusCol).Value = "+"
            Else
                ws.Cells(rowNum, statusCol).Value = "+ " & curStatus
            End If
        End If
    End If
End Sub

Private Function AppendReplyHistory(ByVal existing As String, ByVal newReply As String, _
                                     ByVal mailDate As Date) As String
    Dim entry As String
    entry = "[" & Format(mailDate, "dd.mm.yyyy hh:nn") & "] " & Trim$(newReply)
    If Len(Trim$(existing)) = 0 Then
        AppendReplyHistory = entry
    Else
        AppendReplyHistory = existing & vbCrLf & String(30, "-") & vbCrLf & entry
    End If
End Function

Private Function ParseSapTerminTable(ByVal body As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim clean As String
    clean = Replace(body, vbCrLf, vbLf): clean = Replace(clean, vbCr, vbLf)
    Dim lines() As String: lines = Split(clean, vbLf)
    Dim i As Long, line As String, parts() As String
    
    Dim re As Object
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.Pattern = "(\d{1,4}[./-]\d{1,2}[./-]\d{1,4})" ' Tarih formati (örn: 25.10.2023 veya 2023-10-25)
    
    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) > 0 Then
            parts = Split(line, vbTab) ' StripHtmlArtifacts tablolari Tab ile ayirmisti
            If UBound(parts) >= 3 Then
                Dim sipNo As String, kalemNo As String, tarih As String, p As String
                Dim j As Long
                sipNo = "": kalemNo = "": tarih = ""
                
                For j = 0 To UBound(parts)
                    p = Trim$(Replace(parts(j), """", ""))
                    
                    ' Tarih kontrolü
                    If re.Test(p) Then
                        tarih = p
                    ' Sipariþ No kontrolü (Genellikle 10 haneli ve 04 ile baþlar)
                    ElseIf p Like "##########" Then
                        sipNo = p
                    ' Kalem No kontrolü (Genellikle 5 haneli ve 000 ile baþlar)
                    ElseIf p Like "#####" Then
                        kalemNo = p
                    End If
                Next j
                
                If Len(sipNo) > 0 And Len(tarih) > 0 Then
                    Dim key As String
                    If Len(kalemNo) > 0 Then
                        key = sipNo & "|" & kalemNo
                    Else
                        key = sipNo
                    End If
                    
                    If Not dict.Exists(key) Then dict.Add key, tarih
                End If
            End If
        End If
    Next i
    On Error GoTo 0
    Set ParseSapTerminTable = dict
End Function

Private Function ParseTurkishDate(ByVal text As String, ByVal refDate As Date) As Variant
    Dim t As String: t = LCase$(text)
    t = Replace(t, ChrW(305), "i"): t = Replace(t, ChrW(351), "s")
    t = Replace(t, ChrW(287), "g"): t = Replace(t, ChrW(252), "u")
    t = Replace(t, ChrW(246), "o"): t = Replace(t, ChrW(231), "c")
    t = Replace(t, ChrW(304), "i"): t = Replace(t, ChrW(350), "s")
    t = Replace(t, ChrW(286), "g"): t = Replace(t, ChrW(220), "u")
    t = Replace(t, ChrW(214), "o"): t = Replace(t, ChrW(199), "c")
    On Error Resume Next
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp"): re.Global = False
    
    ' 1) yyyy-mm-dd veya yyyy.mm.dd formatý
    re.Pattern = "(\d{4})[./-](\d{1,2})[./-](\d{1,2})"
    Set m = re.Execute(t)
    If m.Count > 0 Then
        Dim y1 As Long, m1 As Long, d1 As Long
        y1 = CLng(m(0).SubMatches(0)): m1 = CLng(m(0).SubMatches(1)): d1 = CLng(m(0).SubMatches(2))
        If d1 >= 1 And d1 <= 31 And m1 >= 1 And m1 <= 12 Then
            Err.Clear
            Dim dt1 As Date: dt1 = DateSerial(y1, m1, d1)
            If Err.Number = 0 Then
                ParseTurkishDate = dt1: Exit Function
            End If
        End If
    End If
    
    ' 2) dd.mm.yyyy veya dd.mm formatý
    re.Pattern = "(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?"
    Set m = re.Execute(t)
    If m.Count > 0 Then
        Dim dd As Long, mm As Long, yy As Long
        dd = CLng(m(0).SubMatches(0)): mm = CLng(m(0).SubMatches(1))
        If dd >= 1 And dd <= 31 And mm >= 1 And mm <= 12 Then
            If Len(m(0).SubMatches(2)) > 0 Then
                yy = CLng(m(0).SubMatches(2))
                If yy < 100 Then yy = 2000 + yy
            Else
                yy = Year(refDate)
            End If
            Err.Clear
            Dim dt As Date: dt = DateSerial(yy, mm, dd)
            If Err.Number = 0 Then
                If dt < refDate And Len(m(0).SubMatches(2)) = 0 Then
                    dt = DateSerial(yy + 1, mm, dd)
                End If
                ParseTurkishDate = dt: Exit Function
            End If
        End If
    End If
    If InStr(t, "yarin") > 0 Then ParseTurkishDate = refDate + 1: Exit Function
    If InStr(t, "bugun") > 0 Then ParseTurkishDate = refDate: Exit Function
    If InStr(t, "obur gun") > 0 Then ParseTurkishDate = refDate + 2: Exit Function
    Dim dayNames As Variant, i As Long
    dayNames = Array("pazar", "pazartesi", "sali", "carsamba", "persembe", "cuma", "cumartesi")
    For i = 0 To 6
        If InStr(t, dayNames(i)) > 0 Then
            Dim targetDay As Long: targetDay = i + 1
            Dim diff As Long
            diff = (targetDay - Weekday(refDate) + 7) Mod 7
            If diff = 0 Then diff = 7
            If InStr(t, "haftaya") > 0 Or InStr(t, "gelecek hafta") > 0 Then diff = diff + 7
            ParseTurkishDate = refDate + diff: Exit Function
        End If
    Next i
    Set re = CreateObject("VBScript.RegExp"): re.Global = False
    re.Pattern = "(\d+)\s*(is\s*g|g)[un]+"
    Set m = re.Execute(t)
    If m.Count > 0 Then
        Dim nDays As Long: nDays = CLng(m(0).SubMatches(0))
        Dim isBusiness As Boolean: isBusiness = (InStr(LCase$(m(0).Value), "is") > 0)
        If isBusiness Then
            Dim cnt As Long, curDate As Date: curDate = refDate: cnt = 0
            Do While cnt < nDays
                curDate = curDate + 1
                If Weekday(curDate) <> 1 And Weekday(curDate) <> 7 Then cnt = cnt + 1
            Loop
            ParseTurkishDate = curDate
        Else
            ParseTurkishDate = refDate + nDays
        End If
        Exit Function
    End If
    If InStr(t, "gelecek hafta") > 0 Then ParseTurkishDate = refDate + 7: Exit Function
    If InStr(t, "haftaya") > 0 Then ParseTurkishDate = refDate + 7: Exit Function
    Dim months As Variant, mi As Long
    months = Array("ocak", "subat", "mart", "nisan", "mayis", "haziran", _
                   "temmuz", "agustos", "eylul", "ekim", "kasim", "aralik")
    For mi = 0 To 11
        If InStr(t, months(mi)) > 0 Then
            re.Pattern = "(\d{1,2})\s*" & months(mi)
            Set m = re.Execute(t)
            If m.Count > 0 Then
                dd = CLng(m(0).SubMatches(0))
                If dd >= 1 And dd <= 31 Then
                    Err.Clear
                    dt = DateSerial(Year(refDate), mi + 1, dd)
                    If Err.Number = 0 Then
                        If dt < refDate Then dt = DateSerial(Year(refDate) + 1, mi + 1, dd)
                        ParseTurkishDate = dt: Exit Function
                    End If
                End If
            End If
            Err.Clear
            dt = DateSerial(Year(refDate), mi + 2, 0)
            If Err.Number = 0 Then
                If dt < refDate Then dt = DateSerial(Year(refDate) + 1, mi + 2, 0)
                ParseTurkishDate = dt: Exit Function
            End If
        End If
    Next mi
    On Error GoTo 0
End Function

Private Function ClassifyStatus(ByVal replyText As String, ByVal hasDate As Boolean) As String
    Dim t As String: t = LCase$(replyText)
    t = Replace(t, ChrW(305), "i"): t = Replace(t, ChrW(351), "s")
    t = Replace(t, ChrW(287), "g"): t = Replace(t, ChrW(252), "u")
    t = Replace(t, ChrW(246), "o"): t = Replace(t, ChrW(231), "c")
    t = Replace(t, ChrW(304), "i"): t = Replace(t, ChrW(350), "s")
    t = Replace(t, ChrW(286), "g"): t = Replace(t, ChrW(220), "u")
    t = Replace(t, ChrW(214), "o"): t = Replace(t, ChrW(199), "c")
    If InStr(t, "iptal") > 0 Or InStr(t, "veremiyoruz") > 0 Or InStr(t, "yapamiyoruz") > 0 _
       Or InStr(t, "reddettik") > 0 Or InStr(t, "red ettik") > 0 Or InStr(t, "vazgec") > 0 Then
        ClassifyStatus = STATUS_IPTAL: Exit Function
    End If
    If InStr(t, "stok yok") > 0 Or InStr(t, "stokta yok") > 0 Or InStr(t, "stogumuzda yok") > 0 _
       Or InStr(t, "mevcut degil") > 0 Or InStr(t, "tukendi") > 0 Or InStr(t, "bulunmuyor") > 0 Then
        ClassifyStatus = STATUS_STOK_YOK: Exit Function
    End If
    If InStr(t, "sevk edildi") > 0 Or InStr(t, "sevkettik") > 0 Or InStr(t, "kargoya verildi") > 0 _
       Or InStr(t, "yola cikti") > 0 Or InStr(t, "gonderildi") > 0 Or InStr(t, "yolda") > 0 Then
        ClassifyStatus = STATUS_SEVK: Exit Function
    End If
    If InStr(t, "uretimde") > 0 Or InStr(t, "imalatta") > 0 _
       Or InStr(t, "uretime alindi") > 0 Or InStr(t, "uretime girdi") > 0 Or InStr(t, "imal ediliyor") > 0 Then
        ClassifyStatus = STATUS_URETIM: Exit Function
    End If
    If hasDate Then ClassifyStatus = STATUS_TESLIM: Exit Function
    If InStr(t, "yarin") > 0 Or InStr(t, "haftaya") > 0 Or InStr(t, "gun sonra") > 0 _
       Or InStr(t, "is gunu") > 0 Then
        ClassifyStatus = STATUS_TESLIM: Exit Function
    End If
    ClassifyStatus = STATUS_BELIRSIZ
End Function

Private Sub EnsureTerminReportSheet()
    On Error Resume Next
    Set gTermReportWs = ThisWorkbook.Worksheets(TERMIN_REPORT_SHEET)
    If gTermReportWs Is Nothing Then
        Set gTermReportWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        gTermReportWs.Name = TERMIN_REPORT_SHEET
    End If
    On Error GoTo 0
End Sub

Private Sub WriteTerminReport(ByRef cfg As Object, ByVal scanned As Long, _
                               ByVal replied As Long, ByVal updated As Long)
    On Error Resume Next
    EnsureTerminReportSheet
    If gTermReportWs Is Nothing Then Exit Sub
    gTermReportWs.Cells.Clear
    Dim r As Long
    gTermReportWs.Range("A1").Value = T("TERM~IN RAPORU")
    gTermReportWs.Range("A1").Font.Bold = True: gTermReportWs.Range("A1").Font.Size = 14
    gTermReportWs.Range("A2").Value = T("Olu~sturulma: ") & Format(Now, "dd.mm.yyyy hh:nn")
    r = 4
    gTermReportWs.Cells(r, 1).Value = T("OZET"): gTermReportWs.Cells(r, 1).Font.Bold = True
    r = r + 1
    gTermReportWs.Cells(r, 1).Value = T("Taranan mail"): gTermReportWs.Cells(r, 2).Value = scanned: r = r + 1
    gTermReportWs.Cells(r, 1).Value = T("E~sle~stirilen yan~it"): gTermReportWs.Cells(r, 2).Value = replied: r = r + 1
    gTermReportWs.Cells(r, 1).Value = T("G~uncellenen sat~ir"): gTermReportWs.Cells(r, 2).Value = updated: r = r + 1
    gTermReportWs.Cells(r, 1).Value = T("Belirsiz (e~sle~smeyen)"): gTermReportWs.Cells(r, 2).Value = gBelirsizCount: r = r + 2
    Dim statusCounts As Object: Set statusCounts = CreateObject("Scripting.Dictionary")
    statusCounts.Add STATUS_TESLIM, 0: statusCounts.Add STATUS_URETIM, 0
    statusCounts.Add STATUS_SEVK, 0: statusCounts.Add STATUS_STOK_YOK, 0
    statusCounts.Add STATUS_IPTAL, 0: statusCounts.Add STATUS_BELIRSIZ, 0
    Dim orderSheets As Collection: Set orderSheets = cfg("OrderSheets")
    Dim dCol As Long: dCol = CLng(cfg("DurumCol"))
    Dim ws As Worksheet, i As Long, lastR As Long, k As Long, sVal As String
    If dCol > 0 Then
        For i = 1 To orderSheets.Count
            Set ws = orderSheets(i)
            lastR = ws.Cells(ws.Rows.Count, dCol).End(xlUp).Row
            For k = CLng(cfg("StartRow")) To lastR
                sVal = Trim$(CStr(ws.Cells(k, dCol).Value))
                If Len(sVal) > 0 Then
                    If statusCounts.Exists(sVal) Then statusCounts(sVal) = statusCounts(sVal) + 1
                End If
            Next k
        Next i
    End If
    gTermReportWs.Cells(r, 1).Value = T("DURUM DA~GILIMI"): gTermReportWs.Cells(r, 1).Font.Bold = True
    r = r + 1
    Dim key As Variant
    For Each key In statusCounts.Keys
        gTermReportWs.Cells(r, 1).Value = key
        gTermReportWs.Cells(r, 2).Value = statusCounts(key)
        r = r + 1
    Next key
    gTermReportWs.Columns("A:B").ColumnWidth = 30
    On Error GoTo 0
End Sub

Private Function ReplyMentionsOrderLine(ByVal replyText As String, ByVal ws As Worksheet, _
                                        ByVal rowNum As Long, ByVal orderCol As Long, _
                                        ByVal itemCol As Long, ByVal materialCol As Long, _
                                        ByVal longTextCol As Long) As Boolean
    Dim bodyNorm As String, orderNo As String, itemNo As String
    Dim material As String, longText As String, token As String
    bodyNorm = NormalizeName(replyText)
    If Len(bodyNorm) = 0 Then Exit Function
    If orderCol > 0 Then orderNo = Trim$(CStr(ws.Cells(rowNum, orderCol).Value))
    If itemCol > 0 Then itemNo = Trim$(CStr(ws.Cells(rowNum, itemCol).Value))
    If materialCol > 0 Then material = Trim$(CStr(ws.Cells(rowNum, materialCol).Value))
    If longTextCol > 0 Then longText = Trim$(CStr(ws.Cells(rowNum, longTextCol).Value))
    If Len(material) >= 5 Then
        token = NormalizeName(material)
        If Len(token) >= 5 And InStr(1, bodyNorm, token, vbTextCompare) > 0 Then
            ReplyMentionsOrderLine = True: Exit Function
        End If
    End If
    If Len(longText) >= 8 Then
        Dim upperLT As String: upperLT = UCase$(Trim$(longText))
        Dim skipText As Boolean: skipText = False
        If upperLT = "EUR" Or upperLT = "USD" Or upperLT = "TRY" Or _
           upperLT = "ST" Or upperLT = "KG" Or upperLT = "M" Or _
           upperLT = "M2" Or upperLT = "TKM" Or upperLT = "PAK" Then skipText = True
        If Not skipText Then
            token = NormalizeName(longText)
            If Len(token) >= 8 And InStr(1, bodyNorm, token, vbTextCompare) > 0 Then
                ReplyMentionsOrderLine = True: Exit Function
            End If
        End If
    End If
    If Len(orderNo) > 0 Then
        Dim normOrder As String: normOrder = NormalizeOrderNo(orderNo)
        Dim bodyNumbers As Object: Set bodyNumbers = ExtractNumbers(replyText)
        If bodyNumbers.Exists(normOrder) Then
            ReplyMentionsOrderLine = True: Exit Function
        End If
    End If
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    If Len(itemNo) > 0 Then
        re.Pattern = "(kalem|kalem no|item|item no|pozisyon)\s*[:#-]?\s*" & EscapeRegex(itemNo) & "([^0-9]|$)"
        If re.Test(replyText) Then ReplyMentionsOrderLine = True: Exit Function
    End If
    If Len(orderNo) > 0 Then
        re.Pattern = "(siparis|siparis no|satinalma|order|po)\s*[:#-]?\s*" & EscapeRegex(orderNo) & "([^0-9]|$)"
        If re.Test(replyText) Then ReplyMentionsOrderLine = True: Exit Function
    End If
End Function

Private Function ExtractNumbers(ByVal text As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim re As Object, m As Object, i As Long
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = False
    re.Pattern = "\b\d{6,12}\b"
    Set m = re.Execute(text)
    For i = 0 To m.Count - 1
        Dim num As String: num = NormalizeOrderNo(m(i).Value)
        If Not dict.Exists(num) Then dict.Add num, True
    Next i
    Set ExtractNumbers = dict
End Function

Private Function EscapeRegex(ByVal value As String) As String
    Dim chars As Variant, i As Long
    chars = Array("\", ".", "+", "*", "?", "^", "$", "(", ")", "[", "]", "{", "}", "|")
    For i = LBound(chars) To UBound(chars)
        value = Replace(value, CStr(chars(i)), "\" & CStr(chars(i)))
    Next i
    EscapeRegex = value
End Function

Private Function SaveAttachments(ByVal mail As Object, ByVal vendorName As String) As String
    Dim att As Object, savePath As String, basePath As String
    Dim fname As String, result As String, safeVendor As String
    On Error Resume Next
    If mail.Attachments.Count = 0 Then SaveAttachments = "": Exit Function
    On Error GoTo 0
    safeVendor = SanitizeFileName(vendorName)
    basePath = ThisWorkbook.Path & "\" & ATTACH_BASE_FOLDER & "\"
    savePath = basePath & safeVendor & "\"
    If Len(Dir(basePath, vbDirectory)) = 0 Then MkDir basePath
    If Len(Dir(savePath, vbDirectory)) = 0 Then MkDir savePath
    If Len(Dir(savePath, vbDirectory)) = 0 Then
        basePath = Environ("USERPROFILE") & "\Documents\" & ATTACH_BASE_FOLDER & "\"
        savePath = basePath & safeVendor & "\"
        If Len(Dir(basePath, vbDirectory)) = 0 Then MkDir basePath
        If Len(Dir(savePath, vbDirectory)) = 0 Then MkDir savePath
    End If
    result = ""
    For Each att In mail.Attachments
        On Error Resume Next
        fname = Format(Now, "yyyymmdd_hhnnss") & "_" & att.FileName
        Err.Clear
        att.SaveAsFile savePath & fname
        If Err.Number = 0 Then
            If Len(result) > 0 Then result = result & "; "
            result = result & att.FileName
        Else
            WriteLog "EK_HATA", "", vendorName, "", att.FileName, Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next att
    SaveAttachments = result
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad): s = Replace(s, CStr(bad(i)), "_"): Next i
    s = Trim$(s): If Len(s) = 0 Then s = "Bilinmeyen"
    SanitizeFileName = s
End Function

Private Function ExtractNewReply(ByVal body As String) As String
    Dim markers As Variant, i As Long, pos As Long, earliest As Long
    Dim re As Object, m As Object
    body = StripHtmlArtifacts(body)
    earliest = Len(body) + 1
    markers = Array("From: ", T("G~onderilmis: "), T("G~onderen: "), T("Kimden: "), _
        "-----Original Message-----", "-------- Original Message --------", _
        "-------Original Message-------", "________________________________", _
        T("-----Orijinal Mesaj-----"), T("----- Orijinal Mesaj -----"), _
        T("Konu:"), T("Kime:"), T("Tarih:"))
    For i = LBound(markers) To UBound(markers)
        pos = InStr(1, body, CStr(markers(i)), vbTextCompare)
        If pos > 0 And pos < earliest Then earliest = pos
    Next i
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.Pattern = "\d{1,2}[./]\d{1,2}[./]\d{2,4}\s+\d{1,2}:\d{2}\s+tarihinde\s+.+?yazd"
    Set m = re.Execute(body)
    If m.Count > 0 Then
        pos = m(0).FirstIndex + 1
        If pos > 0 And pos < earliest Then earliest = pos
    End If
    pos = InStr(1, body, "tarihinde", vbTextCompare)
    If pos > 0 Then
        Dim yazdiPos As Long
        yazdiPos = InStr(pos, body, "yazd", vbTextCompare)
        If yazdiPos > 0 And (yazdiPos - pos) < 150 Then
            If pos < earliest Then earliest = pos
        End If
    End If
    On Error GoTo 0
    If earliest <= Len(body) Then
        ExtractNewReply = Trim$(Left$(body, earliest - 1))
    Else
        ExtractNewReply = Trim$(body)
    End If
End Function

Private Function FilterRelevantContent(ByVal raw As String) As String
    Dim keywords As Variant, lines() As String
    Dim groups As Object, i As Long, k As Long
    Dim line As String, tip As String, prevTip As String
    Dim buffer As String, result As String, groupIdx As Long
    Dim key As Variant, grp As String, gTip As String, gText As String
    Dim sepPos As Long, matched As Boolean
    raw = StripHtmlArtifacts(raw)
    keywords = GetRelevantKeywords()
    raw = Replace(raw, vbCrLf, vbLf): raw = Replace(raw, vbCr, vbLf)
    lines = Split(raw, vbLf)
    Set groups = CreateObject("Scripting.Dictionary")
    groupIdx = 0: buffer = "": prevTip = ""
    For i = LBound(lines) To UBound(lines)
        line = Trim$(CStr(lines(i)))
        If Len(line) = 0 Then
            If Len(buffer) > 0 Then
                groups.Add groupIdx, prevTip & "|" & buffer
                groupIdx = groupIdx + 1: buffer = "": prevTip = ""
            End If
        Else
            tip = GetLineType(line)
            If Len(buffer) = 0 Then
                buffer = line: prevTip = tip
            ElseIf tip = prevTip Then
                buffer = buffer & " " & line
            Else
                groups.Add groupIdx, prevTip & "|" & buffer
                groupIdx = groupIdx + 1: buffer = line: prevTip = tip
            End If
        End If
    Next i
    If Len(buffer) > 0 Then groups.Add groupIdx, prevTip & "|" & buffer
    result = ""
    For Each key In groups.Keys
        grp = CStr(groups(key))
        sepPos = InStr(grp, "|")
        gTip = Left$(grp, sepPos - 1)
        gText = Mid$(grp, sepPos + 1)
        If gTip = "product" Then
            If HasDateOrNumber(gText) Or HasNumericCode(gText) Then
                If Len(result) > 0 Then result = result & vbCrLf
                result = result & gText
            End If
        ElseIf gTip <> "greeting" And gTip <> "signature" And _
               gTip <> "legal" And gTip <> "phone" And _
               gTip <> "address" And gTip <> "url" And gTip <> "company" Then
            matched = False
            For k = LBound(keywords) To UBound(keywords)
                If InStr(1, gText, CStr(keywords(k)), vbTextCompare) > 0 Then
                    matched = True: Exit For
                End If
            Next k
            If Not matched Then
                If HasDateOrNumber(gText) Or HasMonthOrWeek(gText) Or HasNumericCode(gText) Then matched = True
            End If
            If matched And Len(gText) >= 10 Then
                If Len(result) > 0 Then result = result & vbCrLf
                result = result & gText
            End If
        End If
    Next key
    If Len(result) > MAX_BODY_LEN Then result = Left$(result, MAX_BODY_LEN)
    FilterRelevantContent = result
End Function

Private Function HasNumericCode(ByVal s As String) As Boolean
    Dim re As Object
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = "\b\d{6,12}\b"
    HasNumericCode = re.Test(s)
End Function

Private Function GetLineType(ByVal s As String) As String
    If IsLegalDisclaimer(s) Then
        GetLineType = "legal"
    ElseIf IsCompanyLine(s) Then
        GetLineType = "company"
    ElseIf IsUrlLine(s) Then
        GetLineType = "url"
    ElseIf IsPhoneLine(s) Then
        GetLineType = "phone"
    ElseIf IsAddressLine(s) Then
        GetLineType = "address"
    ElseIf IsGreeting(s) Then
        GetLineType = "greeting"
    ElseIf IsSignatureBlock(s) Then
        GetLineType = "signature"
    ElseIf IsProductLine(s) Then
        GetLineType = "product"
    Else
        GetLineType = "normal"
    End If
End Function

Private Function IsCompanyLine(ByVal s As String) As Boolean
    Dim l As String: l = LCase$(s)
    If Len(s) > 150 Then Exit Function
    If (InStr(l, "ltd. ") > 0 Or InStr(l, "ltd ") > 0) And (InStr(l, "sti") > 0 Or InStr(l, "sirket") > 0) Then
        IsCompanyLine = True: Exit Function
    End If
    If InStr(l, "a.s.") > 0 Or InStr(l, "san. ve tic") > 0 Or InStr(l, "sanayi ve ticaret") > 0 Then
        IsCompanyLine = True
    End If
End Function

Private Function IsLegalDisclaimer(ByVal s As String) As Boolean
    Dim l As String: l = LCase$(s)
    If InStr(l, "kisisel veri") > 0 Or InStr(l, "kvkk") > 0 Or InStr(l, "6698") > 0 Then IsLegalDisclaimer = True: Exit Function
    If InStr(l, "bu e-posta") > 0 Or InStr(l, "bu elektronik posta") > 0 Then IsLegalDisclaimer = True: Exit Function
    If InStr(l, "korunmasi kanunu") > 0 Or InStr(l, "confidentiality") > 0 Then IsLegalDisclaimer = True: Exit Function
    If InStr(l, "disclaimer") > 0 Or (InStr(l, "gizlilik") > 0 And InStr(l, "kanun") > 0) Then IsLegalDisclaimer = True
End Function

Private Function IsProductLine(ByVal s As String) As Boolean
    Dim nUp As Long, nLow As Long, nLetter As Long, nDigit As Long
    Dim hasSemi As Boolean, i As Long, ch As String
    Dim ratio As Double
    If Len(s) < 10 Or Len(s) > 300 Then Exit Function
    If InStr(s, " ") = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "A" And ch <= "Z" Then
            nUp = nUp + 1: nLetter = nLetter + 1
        ElseIf ch >= "a" And ch <= "z" Then
            nLow = nLow + 1: nLetter = nLetter + 1
        ElseIf ch >= "0" And ch <= "9" Then
            nDigit = nDigit + 1
        ElseIf ch = ";" Then
            hasSemi = True
        End If
    Next i
    If nLetter < 5 Then Exit Function
    ratio = nUp / nLetter
    If ratio >= 0.7 And Len(s) >= 15 Then IsProductLine = True: Exit Function
    If ratio >= 0.5 And (nDigit >= 2 Or hasSemi) And Len(s) >= 15 Then IsProductLine = True: Exit Function
    If hasSemi And nDigit >= 3 And Len(s) < 150 Then IsProductLine = True
End Function

Private Function IsGreeting(ByVal s As String) As Boolean
    Dim arr As Variant, i As Long
    arr = Array("merhaba", "merhabalar", "selam", "iyi gunler", "gunaydin", _
                "iyi aksamlar", "umarim iyi", "nasilsiniz", "iyisinizdir")
    For i = LBound(arr) To UBound(arr)
        If InStr(1, s, CStr(arr(i)), vbTextCompare) > 0 Then IsGreeting = True: Exit Function
    Next i
End Function

Private Function IsSignatureBlock(ByVal s As String) As Boolean
    Dim arr As Variant, i As Long
    arr = Array("iyi calismalar", "saygilar", "saygilarimla", "saygilarimizla", _
                "best regards", "kind regards", "regards", _
                "tesekkur ederim", "tesekkurler", "gizlilik", "confidentiality", _
                "tarafindan gonderilmistir", "dilerim", "dileriz")
    For i = LBound(arr) To UBound(arr)
        If InStr(1, s, CStr(arr(i)), vbTextCompare) > 0 Then IsSignatureBlock = True: Exit Function
    Next i
End Function

Private Function IsPhoneLine(ByVal s As String) As Boolean
    Dim l As String, cl As String
    Dim nDigit As Long, nLetter As Long, i As Long, ch As String
    Dim allDig As Boolean
    l = LCase$(s)
    If InStr(l, "tel:") > 0 Or InStr(l, "telefon:") > 0 Or _
       InStr(l, "mobil:") > 0 Or InStr(l, "gsm:") > 0 Or _
       InStr(l, "mobile:") > 0 Or InStr(l, "phone:") > 0 Or _
       InStr(l, "faks:") > 0 Or InStr(l, "fax:") > 0 Then
        If Len(s) < 250 Then IsPhoneLine = True: Exit Function
    End If
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            nDigit = nDigit + 1
        ElseIf (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            nLetter = nLetter + 1
        End If
    Next i
    If nDigit >= 12 And nLetter <= 15 Then IsPhoneLine = True: Exit Function
    cl = s
    cl = Replace(cl, "+", ""): cl = Replace(cl, "(", ""): cl = Replace(cl, ")", "")
    cl = Replace(cl, "-", ""): cl = Replace(cl, " ", ""): cl = Replace(cl, ".", "")
    cl = Replace(cl, "/", ""): cl = Replace(cl, Chr(9), "")
    If Len(cl) >= 7 And Len(cl) <= 20 Then
        allDig = True
        For i = 1 To Len(cl)
            ch = Mid$(cl, i, 1)
            If ch < "0" Or ch > "9" Then allDig = False: Exit For
        Next i
        If allDig Then IsPhoneLine = True
    End If
End Function

Private Function IsAddressLine(ByVal s As String) As Boolean
    Dim l As String, score As Long
    l = LCase$(s)
    If InStr(l, "adres:") > 0 Or InStr(l, "address:") > 0 Or InStr(l, "adresi:") > 0 Then IsAddressLine = True: Exit Function
    If InStr(l, " mh.") > 0 Or InStr(l, " mh ") > 0 Or InStr(l, " mah.") > 0 Or InStr(l, " mahalle") > 0 Then score = score + 2
    If InStr(l, " sk.") > 0 Or InStr(l, " sok.") > 0 Or InStr(l, " sokak") > 0 Then score = score + 2
    If InStr(l, " blok") > 0 Or InStr(l, " blk") > 0 Then score = score + 2
    If InStr(l, " no:") > 0 Or InStr(l, " no.") > 0 Then score = score + 2
    If InStr(l, " daire") > 0 Or InStr(l, " kat") > 0 Then score = score + 2
    If InStr(l, " cad") > 0 Or InStr(l, " cadde") > 0 Then score = score + 1
    If InStr(l, " mevki") > 0 Or InStr(l, " mevkii") > 0 Then score = score + 1
    If InStr(l, " bulvar") > 0 Or InStr(l, " blv.") > 0 Then score = score + 1
    If InStr(l, " organize") > 0 Or InStr(l, " osb") > 0 Then score = score + 1
    If InStr(l, ChrW(252) & "rkiye") > 0 Or InStr(l, "turkiye") > 0 Or InStr(l, "turkey") > 0 Then score = score + 1
    Dim cities As Variant, i As Long
    cities = Array("istanbul", "ankara", "izmir", "kocaeli", "bursa", "antalya", "adana", "konya", "gaziantep", "mersin", "kayseri", "samsun", "denizli", "malatya", "erzurum", "trabzon", "manisa", "sakarya", "hatay", "zonguldak", "balikesir")
    For i = LBound(cities) To UBound(cities)
        If InStr(l, CStr(cities(i))) > 0 Then score = score + 1: Exit For
    Next i
    Dim re As Object
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.Pattern = "\b\d{5}\b"
    If re.Test(l) Then score = score + 1
    On Error GoTo 0
    If score >= 2 Then IsAddressLine = True
End Function

Private Function IsUrlLine(ByVal s As String) As Boolean
    Dim l As String: l = LCase$(s)
    If InStr(l, "http://") > 0 Or InStr(l, "https://") > 0 Then IsUrlLine = True: Exit Function
    If InStr(l, "www.") > 0 Or InStr(l, "urldefense") > 0 Then IsUrlLine = True: Exit Function
    If Len(s) < 100 Then
        If InStr(l, ".com") > 0 Or InStr(l, ".net") > 0 Or InStr(l, ".org") > 0 Or InStr(l, ".tr") > 0 Then IsUrlLine = True
    End If
End Function

Private Function HasDateOrNumber(ByVal s As String) As Boolean
    If CountDigits(s) >= 4 Then HasDateOrNumber = True
End Function

Private Function HasMonthOrWeek(ByVal s As String) As Boolean
    Dim arr As Variant, i As Long
    arr = Array("ocak", "subat", "mart", "nisan", "mayis", "haziran", "temmuz", _
                "agustos", "eylul", "ekim", "kasim", "aralik", _
                "hafta", "is gunu", "gun icinde")
    For i = LBound(arr) To UBound(arr)
        If InStr(1, s, CStr(arr(i)), vbTextCompare) > 0 Then HasMonthOrWeek = True: Exit Function
    Next i
End Function

Private Function CountDigits(ByVal s As String) As Long
    Dim i As Long, ch As String, n As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then n = n + 1
    Next i
    CountDigits = n
End Function

Private Function GetRelevantKeywords() As Variant
    Dim s As String
    s = "revize|revizyon|guncelle|duzelt|degisiklik|degisim|fiyat farki|birim fiyat|" & _
        "liste fiyat|guncel fiyat|indirim|iskonto|zam|" & _
        "termin tarih|teslim tarih|sevk tarih|teslim edil|sevk edil|sevkiyat|" & _
        "onayl|onay ver|kabul edil|iptal|reddedil|red edil|" & _
        "stokta|stok yok|mevcut degil|tukendi|" & _
        "tedarik edil|tedarik sure|uretim tarih|uretil|imalat|" & _
        "odeme|vade|fatura|gecikme|gecikebilir|aksama|aksaklik|" & _
        "sorun|problem|hata var|acil|oncelik|" & _
        "kalite|sertifika|analiz|rapor|donecegiz|" & _
        "bilgi ver|haber ver|teyit|is gunu|hafta sonra|" & _
        "hafta ici|gun icinde|termin|teslim|sevk|" & _
        "bugun|yarin|kargoya|kargolandi"
    GetRelevantKeywords = Split(s, "|")
End Function

Private Function IsSiparisKonusu(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    If InStr(1, s, T("A~c~ik sipari~sler hakk~inda"), vbTextCompare) > 0 Then IsSiparisKonusu = True: Exit Function
    If InStr(1, s, "Acik siparisler", vbTextCompare) > 0 Then IsSiparisKonusu = True: Exit Function
    If InStr(1, s, T("sipari~sler hakk~inda"), vbTextCompare) > 0 Then IsSiparisKonusu = True
End Function

Private Function FindVendorInSenderName(ByVal vws As Worksheet, ByVal lastRow As Long, _
                                        ByVal nameCol As Long, ByVal searchText As String, _
                                        ByRef matchState As String) As String
    Dim r As Long, vname As String, found As String, hits As Long
    Dim searchNorm As String, vNorm As String
    Dim twoWords As String, twoWordsNorm As String
    Dim firstWord As String, firstWordNorm As String
    matchState = "": searchNorm = NormalizeName(searchText)
    If Len(searchNorm) < 3 Then Exit Function
    found = "": hits = 0
    For r = 2 To lastRow
        vname = Trim$(CStr(vws.Cells(r, nameCol).Value))
        If Len(vname) >= 3 Then
            vNorm = NormalizeName(vname)
            If Len(vNorm) >= 3 And InStr(1, searchNorm, vNorm, vbTextCompare) > 0 Then
                If Len(found) = 0 Then
                    found = vname: hits = 1
                ElseIf LCase$(found) <> LCase$(vname) Then
                    hits = hits + 1
                End If
            End If
        End If
    Next r
    If hits = 1 Then matchState = "tam isim": FindVendorInSenderName = found: Exit Function
    found = "": hits = 0
    For r = 2 To lastRow
        vname = Trim$(CStr(vws.Cells(r, nameCol).Value))
        If Len(vname) >= 5 Then
            twoWords = GetFirstNthWord(vname, 2)
            If Len(twoWords) >= 6 Then
                twoWordsNorm = NormalizeName(twoWords)
                If Len(twoWordsNorm) >= 5 And InStr(1, searchNorm, twoWordsNorm, vbTextCompare) > 0 Then
                    If Len(found) = 0 Then
                        found = vname: hits = 1
                    ElseIf LCase$(found) <> LCase$(vname) Then
                        hits = hits + 1
                    End If
                End If
            End If
        End If
    Next r
    If hits = 1 Then matchState = "ilk 2 kelime": FindVendorInSenderName = found: Exit Function
    found = "": hits = 0
    For r = 2 To lastRow
        vname = Trim$(CStr(vws.Cells(r, nameCol).Value))
        If Len(vname) >= 3 Then
            firstWord = GetFirstNthWord(vname, 1)
            If Len(firstWord) >= 4 Then
                firstWordNorm = NormalizeName(firstWord)
                If Len(firstWordNorm) >= 4 And InStr(1, searchNorm, firstWordNorm, vbTextCompare) > 0 Then
                    If Len(found) = 0 Then
                        found = vname: hits = 1
                    ElseIf LCase$(found) <> LCase$(vname) Then
                        hits = hits + 1
                    End If
                End If
            End If
        End If
    Next r
    If hits = 1 Then
        matchState = "ilk 1 kelime": FindVendorInSenderName = found
    ElseIf hits > 1 Then
        matchState = "belirsiz (1 kelime)"
    End If
End Function

Private Function GetOutlookApp() As Object
    Dim app As Object
    On Error Resume Next
    Set app = GetObject(, "Outlook.Application")
    If app Is Nothing Then Set app = CreateObject("Outlook.Application")
    On Error GoTo 0
    Set GetOutlookApp = app
End Function

Private Function GetInboxFolder(ByVal olNs As Object, ByVal folderPath As String) As Object
    Dim f As Object, parts() As String, i As Long
    On Error Resume Next
    If Len(Trim$(folderPath)) = 0 Then
        Set f = olNs.GetDefaultFolder(6)
    Else
        parts = Split(folderPath, "\")
        Set f = olNs.GetDefaultFolder(6)
        For i = LBound(parts) To UBound(parts)
            Set f = f.Folders(CStr(parts(i)))
        Next i
    End If
    On Error GoTo 0
    Set GetInboxFolder = f
End Function

Private Function GetSenderSmtpEx(ByVal mail As Object) As String
    Dim s As String, pa As Object, snd As Object
    GetSenderSmtpEx = ""
    On Error Resume Next
    Set pa = mail.PropertyAccessor
    If Not pa Is Nothing Then s = pa.GetProperty("http://schemas.microsoft.com/mapi/proptags/0x39FE001E")
    On Error GoTo 0
    If Len(s) > 0 And InStr(s, "@") > 0 Then GetSenderSmtpEx = LCase$(Trim$(s)): Exit Function
    s = "": On Error Resume Next: s = mail.SenderEmailAddress: On Error GoTo 0
    If Len(s) > 0 And InStr(s, "@") > 0 And InStr(s, "/") = 0 Then GetSenderSmtpEx = LCase$(Trim$(s)): Exit Function
    s = "": On Error Resume Next
    Set snd = mail.sender
    If Not snd Is Nothing Then s = snd.GetExchangeUser.PrimarySmtpAddress
    On Error GoTo 0
    If Len(s) > 0 And InStr(s, "@") > 0 Then GetSenderSmtpEx = LCase$(Trim$(s)): Exit Function
    s = "": On Error Resume Next
    If mail.ReplyRecipients.Count > 0 Then s = mail.ReplyRecipients(1).Address
    On Error GoTo 0
    If Len(s) > 0 And InStr(s, "@") > 0 And InStr(s, "/") = 0 Then GetSenderSmtpEx = LCase$(Trim$(s))
End Function

Private Function GetConfig() As Object
    If gCfg Is Nothing Then
        Set gCfg = CreateObject("Scripting.Dictionary")
        gCfg.CompareMode = vbTextCompare
    End If
    Set GetConfig = gCfg
End Function

Private Function ListSheetNames() As String
    Dim i As Long, s As String
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & "  - " & ThisWorkbook.Worksheets(i).Name
    Next i
    ListSheetNames = s
End Function

Private Function AskWorksheet(ByVal prompt As String, ByVal defaultName As String) As Worksheet
    Dim s As String, ws As Worksheet
    s = InputBox(prompt & vbCrLf & ListSheetNames() & vbCrLf & T("Bo~s = ") & defaultName, _
                 T("Otomasyon"), defaultName)
    If StrPtr(s) = 0 Then Exit Function
    s = Trim$(s)
    If Len(s) = 0 Then s = defaultName
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(s): On Error GoTo 0
    If ws Is Nothing Then MsgBox T("Sayfa bulunamad~i: ") & s, vbExclamation, AppTitle(): Exit Function
    Set AskWorksheet = ws
End Function

Private Function AskStartRow() As Long
    Dim s As String, v As Long
    s = InputBox(T("Ba~slang~i~c sat~ir? (varsay~ilan ") & DEFAULT_START_ROW & ")", _
                 T("Otomasyon"), CStr(DEFAULT_START_ROW))
    If StrPtr(s) = 0 Then Exit Function
    s = Trim$(s)
    If Len(s) = 0 Then AskStartRow = DEFAULT_START_ROW: Exit Function
    If Not IsNumeric(s) Then MsgBox T("Ge~cersiz."), vbExclamation, AppTitle(): Exit Function
    v = CLng(s): If v < 2 Then v = 2
    AskStartRow = v
End Function

Private Sub WriteLog(ByVal islem As String, ByVal sayfa As String, _
                     ByVal tedarikci As String, ByVal mail As String, _
                     ByVal detay1 As String, Optional ByVal detay2 As String = "")
    On Error Resume Next
    If gLogWs Is Nothing Then
        Set gLogWs = ThisWorkbook.Worksheets(LOG_SHEET)
        If gLogWs Is Nothing Then
            Set gLogWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            gLogWs.Name = LOG_SHEET
            gLogWs.Range("A1").Value = T("Tarih/Saat"): gLogWs.Range("B1").Value = T("~I~slem")
            gLogWs.Range("C1").Value = T("Sayfa"): gLogWs.Range("D1").Value = T("Tedarik~ci")
            gLogWs.Range("E1").Value = T("Mail"): gLogWs.Range("F1").Value = T("Detay 1")
            gLogWs.Range("G1").Value = T("Detay 2")
            gLogWs.Range("A1:G1").Font.Bold = True
            gLogWs.Range("A1:G1").Interior.Color = RGB(220, 220, 220)
            gLogWs.Columns("A:G").ColumnWidth = 25
        End If
    End If
    gLogRow = gLogWs.Cells(gLogWs.Rows.Count, 1).End(xlUp).Row + 1
    If gLogRow < 2 Then gLogRow = 2
    gLogWs.Cells(gLogRow, 1).Value = Format(Now, "dd.mm.yyyy hh:nn:ss")
    gLogWs.Cells(gLogRow, 2).Value = islem
    gLogWs.Cells(gLogRow, 3).Value = sayfa
    gLogWs.Cells(gLogRow, 4).Value = tedarikci
    gLogWs.Cells(gLogRow, 5).Value = mail
    gLogWs.Cells(gLogRow, 6).Value = detay1
    gLogWs.Cells(gLogRow, 7).Value = detay2
    gLogRow = gLogRow + 1
End Sub
' Gemini integration. Model output is data only, never VBA or shell commands.
Private Sub GeminiPreflight()
    Dim shell As Object, root As String, fso As Object, backup As String
    Set shell = CreateObject("WScript.Shell")
    If Len(shell.Environment("PROCESS")("GEMINI_API_KEY")) = 0 Then _
        Err.Raise vbObjectError + 901, , "Windows GEMINI_API_KEY eksik. Ayarlayip Excel'i yeniden acin."
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = GeminiRoot()
    If Not fso.FileExists(root & "\gemini\bridge.py") Then _
        Err.Raise vbObjectError + 902, , "gemini\bridge.py bulunamadi. GEMINI_BRIDGE_DIR ayarini kontrol edin."
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 903, , "Once Excel dosyasini kaydedin."
    backup = ThisWorkbook.Path & "\GeminiYedek_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & _
        fso.GetBaseName(fso.GetTempName) & "." & fso.GetExtensionName(ThisWorkbook.Name)
    ThisWorkbook.SaveCopyAs backup
End Sub

Private Function GeminiRoot() As String
    GeminiRoot = Environ$("GEMINI_BRIDGE_DIR")
    If Len(GeminiRoot) = 0 Then GeminiRoot = ThisWorkbook.Path
End Function

Private Function GeminiJson(ByVal value As String) As String
    Dim i As Long, ch As String, result As String, code As Long
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)
        Select Case code
            Case 34: result = result & Chr$(92) & Chr$(34)
            Case 92: result = result & Chr$(92) & Chr$(92)
            Case 0 To 31: result = result & Chr$(92) & "u" & Right$("0000" & Hex$(code), 4)
            Case Else: result = result & ch
        End Select
    Next i
    GeminiJson = Chr$(34) & result & Chr$(34)
End Function

Private Sub GeminiWriteUtf8(ByVal path As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText value
    stream.SaveToFile path, 2
    stream.Close
End Sub

Private Function GeminiReadUtf8(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile path
    GeminiReadUtf8 = stream.ReadText
    stream.Close
End Function

Private Function GeminiRun(ByVal request As String) As String
    Dim fso As Object, shell As Object, source As String, dest As String
    Dim python As String, command As String, result As String, rc As Long
    Dim errNum As Long, errText As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    source = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName)
    dest = source & ".result"
    GeminiWriteUtf8 source, request
    python = Environ$("GEMINI_PYTHON")
    If Len(python) = 0 Then
        python = "py -3"
    Else
        python = Chr$(34) & python & Chr$(34)
    End If
    command = python & " " & Chr$(34) & GeminiRoot() & "\gemini\bridge.py" & Chr$(34) & _
        " " & Chr$(34) & source & Chr$(34) & " " & Chr$(34) & dest & Chr$(34)
    rc = shell.Run(command, 0, True)
    If Not fso.FileExists(dest) Then Err.Raise vbObjectError + 904, , "Python baslatilamadi. Python 3 ve py launcher kurulumunu kontrol edin."
    result = GeminiReadUtf8(dest)
    If rc <> 0 Or Left$(result, 5) = "ERROR" Then Err.Raise vbObjectError + 905, , result
    GeminiRun = result
    GoTo Cleanup
Failed:
    errNum = Err.Number: errText = Err.Description
Cleanup:
    On Error Resume Next
    If Not fso Is Nothing Then
        If fso.FileExists(source) Then fso.DeleteFile source, True
        If fso.FileExists(dest) Then fso.DeleteFile dest, True
    End If
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "Gemini", errText
End Function

Private Function GeminiWriteSheet(ByRef cfg As Object, ByVal ws As Worksheet, _
    ByVal vendor As String, ByVal reply As String, ByVal received As Date, _
    ByVal sender As String, ByVal method As String, ByVal attachments As String, ByVal testMode As Boolean) As Long
    Dim candidates As String, request As String, output As String, line As Variant, fields As Variant
    Dim r As Long, lastRow As Long, count As Long, target As Long, d As String
    Dim allowed As Object, seen As Object, norm As String
    Set allowed = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    norm = NormalizeName(vendor)
    lastRow = ws.Cells(ws.Rows.Count, CLng(cfg("VendorCol"))).End(xlUp).Row
    For r = CLng(cfg("StartRow")) To lastRow
        ' Exact normalized vendor only; fuzzy sender/vendor matches do not authorize other rows.
        If NormalizeName(SafeStr(ws.Cells(r, CLng(cfg("VendorCol"))).Value)) = norm Then
            If IsDate(ws.Cells(r, CLng(cfg("ReplyDateCol"))).Value) Then
                If CDate(ws.Cells(r, CLng(cfg("ReplyDateCol"))).Value) >= received Then GoTo NextCandidate
            End If
            If allowed.Count >= 500 Then Err.Raise vbObjectError + 906, , "500'den fazla aday satir: islemi daraltin."
            If Len(candidates) > 0 Then candidates = candidates & ","
            candidates = candidates & "{" & GeminiJson("row") & ":" & CStr(r) & "," & _
                GeminiJson("order") & ":" & GeminiJson(SafeStr(ws.Cells(r, CLng(cfg("OrderNoCol"))).Value)) & "," & _
                GeminiJson("item") & ":" & GeminiJson(SafeStr(ws.Cells(r, CLng(cfg("KalemNoCol"))).Value)) & "," & _
                GeminiJson("material") & ":" & GeminiJson(SafeStr(ws.Cells(r, CLng(cfg("MalzemeCol"))).Value)) & "}"
            allowed.Add CStr(r), True
        End If
NextCandidate:
    Next r
    If allowed.Count = 0 Then Exit Function
    request = "{" & GeminiJson("body") & ":" & GeminiJson(reply) & "," & _
        GeminiJson("received_at") & ":" & GeminiJson(Format$(received, "yyyy-mm-dd")) & "," & _
        GeminiJson("candidates") & ":[" & candidates & "]}"
    output = GeminiRun(request)
    If Len(output) = 0 Then
        WriteLog "GEMINI_BELIRSIZ", ws.Name, vendor, sender, "Guvenilir eslesme yok; satir degistirilmedi."
        Exit Function
    End If
    For Each line In Split(output, vbLf)
        fields = Split(CStr(line), vbTab)
        If UBound(fields) <> 3 Then Err.Raise vbObjectError + 907, , "Gemini sonuc bicimi hatali."
        target = CLng(fields(0))
        If Not allowed.Exists(CStr(target)) Or seen.Exists(CStr(target)) Then _
            Err.Raise vbObjectError + 908, , "Gemini satiri dogrulanamadi."
        seen.Add CStr(target), True
        If testMode And count >= TEST_MAX_ORDERS Then Exit For
        ' Bypass old free-text date/status heuristics. Apply validated structured fields directly.
        WriteReplyToRow ws, target, CStr(fields(3)), received, sender, "Gemini / " & method, attachments, _
            CLng(cfg("ReplyCol")), CLng(cfg("ReplyDateCol")), CLng(cfg("ReplyFromCol")), _
            CLng(cfg("MatchMethodCol")), CLng(cfg("AttachCol")), 0, 0, CLng(cfg("StatusCol")), ""
        If CLng(cfg("DurumCol")) > 0 Then ws.Cells(target, CLng(cfg("DurumCol"))).Value = CStr(fields(2))
        d = CStr(fields(1))
        If Len(d) > 0 And CLng(cfg("TerminCol")) > 0 Then
            ws.Cells(target, CLng(cfg("TerminCol"))).Value = DateSerial(CInt(Left$(d, 4)), CInt(Mid$(d, 6, 2)), CInt(Right$(d, 2)))
            ws.Cells(target, CLng(cfg("TerminCol"))).NumberFormat = "dd.mm.yyyy"
        End If
        count = count + 1
    Next line
    GeminiWriteSheet = count
End Function
