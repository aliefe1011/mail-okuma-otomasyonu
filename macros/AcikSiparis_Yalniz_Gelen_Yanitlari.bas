Attribute VB_Name = "OnlyReadReplies"
Option Explicit

'============================================================
'  ACIK SIPARIS MAIL OTOMASYONU - GELEN YANITLARI OKU
'  Surum: 6.2 (Tedarikci ve kalem bazli guvenli eslestirme)
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
    
    re.Pattern = "<br\s*/?>|</p\s*>|</div\s*>|</li\s*>"
    s = re.Replace(s, vbCrLf)
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
    DoReadReplies cfg
    Exit Sub
FatalError:
    AppStateEnd
    MsgBox T("Kritik Hata (Satir ") & Erl & "):" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           T("Lutfen Excel sayfalarini ve sutun basliklarini kontrol edin."), vbCritical, AppTitle()
    WriteLog "FATAL", "", "", "", "GelenYanitlariOku", Err.Description
End Sub

Private Function ValidateSchema(ByRef cfg As Object) As Boolean
    Dim required As Variant, outputs As Variant, inputs As Variant
    Dim key As Variant, other As Variant, used As Object, col As Long, msgs As String
    Set used = CreateObject("Scripting.Dictionary")
    required = Array("OrderNoCol", "KalemNoCol", "VendorCol", "MalzemeCol", "ReplyCol", "ReplyDateCol", "VendorNameCol", "VendorEmailCol")
    For Each key In required
        col = CLng(cfg(CStr(key)))
        If col < 1 Or col > 16384 Then msgs = msgs & CStr(key) & " sutunu gecersiz." & vbCrLf
    Next key
    If cfg("VendorNameCol") = cfg("VendorEmailCol") Then msgs = msgs & "Tedarikci adi ve e-posta sutunu ayni olamaz." & vbCrLf
    inputs = Array("OrderNoCol", "KalemNoCol", "VendorCol", "MalzemeCol", "UretimYeriCol", "UzunMetinCol", "BirimCol", "BakiyeCol")
    outputs = Array("ReplyCol", "ReplyDateCol", "ReplyFromCol", "MatchMethodCol", "AttachCol", "DurumCol", "TerminCol", "StatusCol")
    For Each key In outputs
        col = CLng(cfg(CStr(key)))
        If col > 0 Then
            If col > 16384 Then msgs = msgs & CStr(key) & " sutunu gecersiz." & vbCrLf
            If used.Exists(CStr(col)) Then
                msgs = msgs & CStr(key) & " baska bir sonuc sutunuyla ayni." & vbCrLf
            Else
                used.Add CStr(col), True
            End If
            For Each other In inputs
                If col = CLng(cfg(CStr(other))) Then msgs = msgs & CStr(key) & " kaynak veri sutunuyla ayni: " & CStr(other) & vbCrLf
            Next other
        End If
    Next key
    If Len(msgs) > 0 Then
        MsgBox "Ayarlar duzeltilmeden okuma baslatilmadi:" & vbCrLf & msgs, vbCritical, AppTitle()
        Exit Function
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
    Dim addresses As Variant, address As Variant
    For r = 2 To vLastRow
        vname = Trim$(SafeStr(vws.Cells(r, vendorNameCol).Value))
        If Len(vname) > 0 And InStr(UCase$(vname), "BLOKE") = 0 And InStr(UCase$(vname), "BLOCKED") = 0 Then
            addresses = Split(Replace(Replace(SafeStr(vws.Cells(r, vendorEmailCol).Value), ",", ";"), vbCrLf, ";"), ";")
            For Each address In addresses
                vemail = NormalizeEmail(CStr(address))
                If Len(vemail) > 0 Then
                    AddVendorIdentity vendorByEmail, vemail, vname
                    vdomain = Mid$(vemail, InStr(vemail, "@") + 1)
                    AddVendorIdentity vendorByDomain, vdomain, vname
                    loadedCount = loadedCount + 1
                End If
            Next address
        End If
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

        If Not ReadMailFields(mail, mailDate, senderEmail, senderName, mailSubject, mailBody) Then
            skipCount = skipCount + 1
            GoTo NextMail
        End If
        If mailDate < cutoffDate Then
            dateSkipCount = dateSkipCount + 1: skipCount = skipCount + 1
            GoTo NextMail
        End If
        ' Fields have been read with a local error handler; no Resume Next leaks across mails.
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
            If vendorByEmail(senderEmail) = "***CAKISMA***" Then
                WriteLog "BELIRSIZ_GONDEREN", "", "", senderEmail, "Ayni e-posta birden fazla tedarikciye kayitli."
                GoTo NextMail
            End If
            vendorName = vendorByEmail(senderEmail): matchMethod = "e-posta"
        End If
        If Len(vendorName) = 0 And Len(senderDomain) > 0 And Not IsPublicMailDomain(senderDomain) Then
            If vendorByDomain.Exists(senderDomain) Then
                If vendorByDomain(senderDomain) <> "***CAKISMA***" Then
                    vendorName = vendorByDomain(senderDomain): matchMethod = "domain"
                End If
            End If
        End If

        If Len(vendorName) > 0 Then
            replyCount = replyCount + 1
            newReply = ExtractNewReply(mailBody)
            If Len(Trim$(newReply)) = 0 Then
                WriteLog "BOS_YANIT", "", vendorName, senderEmail, "Yeni yanit bos; satir degistirilmedi."
                GoTo NextMail
            End If
            attachmentList = SaveAttachments(mail, vendorName)
            sheetUpdated = 0
            For i = 1 To orderSheets.Count
                upd = WriteReplyToSheet(cfg, orderSheets(i), vendorName, newReply, _
                    mailDate, senderEmail, matchMethod, attachmentList, tMode, mailSubject)
                sheetUpdated = sheetUpdated + upd
            Next i
            matchCount = matchCount + sheetUpdated
            If sheetUpdated > 0 And Len(entryId) > 0 And Not tMode Then MarkMailProcessed entryId, vendorName
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
    ByVal vendorName As String, ByVal newReply As String, ByVal mailDate As Date, _
    ByVal senderEmail As String, ByVal matchMethod As String, ByVal attachmentList As String, _
    ByVal tMode As Boolean, Optional ByVal mailSubject As String = "") As Long
    Dim matches As Object, key As Variant, r As Long, n As Long
    Dim replyText As String
    Set matches = SelectReplyRows(cfg, ws, vendorName, newReply, mailSubject)
    If matches.Count = 0 Then
        gBelirsizCount = gBelirsizCount + 1
        WriteLog "BELIRSIZ_YANIT", ws.Name, vendorName, senderEmail, _
            "Tedarikci/siparis/kalem kesin eslesmedi; hicbir satir degistirilmedi."
        Exit Function
    End If
    For Each key In matches.Keys
        r = CLng(key)
        If tMode And n >= TEST_MAX_ORDERS Then Exit For
        ' Inbox order is not chronological. Never let an older reply overwrite a newer one.
        If IsDate(ws.Cells(r, CLng(cfg("ReplyDateCol"))).Value) Then
            If CDate(ws.Cells(r, CLng(cfg("ReplyDateCol"))).Value) >= mailDate Then GoTo NextMatch
        End If
        replyText = CStr(matches(key))
        WriteReplyToRow ws, r, replyText, mailDate, senderEmail, matchMethod & " / kalem-dogrulandi", _
            attachmentList, CLng(cfg("ReplyCol")), CLng(cfg("ReplyDateCol")), _
            CLng(cfg("ReplyFromCol")), CLng(cfg("MatchMethodCol")), CLng(cfg("AttachCol")), _
            CLng(cfg("DurumCol")), CLng(cfg("TerminCol")), CLng(cfg("StatusCol")), ""
        n = n + 1
NextMatch:
    Next key
    WriteReplyToSheet = n
End Function

Private Function SelectReplyRows(ByRef cfg As Object, ByVal ws As Worksheet, _
    ByVal vendorName As String, ByVal reply As String, ByVal subject As String) As Object
    Dim result As Object, vendorRows As Collection, allOrders As Object
    Dim r As Long, lastRow As Long, key As Variant, rowKey As Variant
    Dim lines As Variant, oneLine As Variant, line As String, activeRow As String
    Dim hits As Collection, orderNo As String, itemNo As String, material As String
    Dim lineHasOrder As Boolean, subjectHasOrder As Boolean, lineHasItem As Boolean
    Dim orderContext As String, processedReply As String
    Dim orderOK As Boolean, itemOK As Boolean, materialOK As Boolean
    Dim hasAnyIdentifier As Boolean, ambiguous As Boolean, vendorCount As Long
    Dim hasLineMaterial As Boolean, score As Long, bestScore As Long
    Set result = CreateObject("Scripting.Dictionary")
    Set allOrders = CreateObject("Scripting.Dictionary")
    allOrders.CompareMode = vbTextCompare
    Set vendorRows = New Collection
    lastRow = ws.Cells(ws.Rows.Count, CLng(cfg("OrderNoCol"))).End(xlUp).Row
    For r = CLng(cfg("StartRow")) To lastRow
        orderNo = NormalizeOrderNo(SafeStr(ws.Cells(r, CLng(cfg("OrderNoCol"))).Value))
        If Len(orderNo) >= 4 Then allOrders(orderNo) = True
        If Len(orderNo) > 0 And NormalizeName(SafeStr(ws.Cells(r, CLng(cfg("VendorCol"))).Value)) = NormalizeName(vendorName) Then
            vendorRows.Add r
        End If
    Next r
    vendorCount = vendorRows.Count
    If vendorCount = 0 Then Set SelectReplyRows = result: Exit Function
    orderContext = subject
    subjectHasOrder = HasKnownOrder(subject, allOrders) Or HasOrderLabel(subject)
    lines = Split(Replace(Replace(reply, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For Each oneLine In lines
        line = Trim$(CStr(oneLine))
        If IsSignatureStart(line) Then Exit For
        If Len(line) = 0 Then GoTo NextLine
        If Len(processedReply) > 0 Then processedReply = processedReply & vbCrLf
        processedReply = processedReply & line
        lineHasOrder = HasKnownOrder(line, allOrders) Or HasOrderLabel(line)
        If lineHasOrder Then
            orderContext = line
            subjectHasOrder = True
        End If
        lineHasItem = HasItemLabel(line)
        If lineHasOrder Or lineHasItem Then hasAnyIdentifier = True
        Set hits = New Collection
        hasLineMaterial = False: bestScore = 0
        For Each rowKey In vendorRows
            r = CLng(rowKey)
            orderNo = SafeStr(ws.Cells(r, CLng(cfg("OrderNoCol"))).Value)
            itemNo = SafeStr(ws.Cells(r, CLng(cfg("KalemNoCol"))).Value)
            material = SafeStr(ws.Cells(r, CLng(cfg("MalzemeCol"))).Value)
            orderOK = False: itemOK = False: materialOK = False
            If lineHasOrder Then
                orderOK = ContainsCode(line, orderNo)
            ElseIf subjectHasOrder Then
                orderOK = ContainsCode(orderContext, orderNo)
            End If
            If Len(Trim$(material)) >= 5 Then materialOK = ContainsCode(line, material)
            If materialOK Then hasAnyIdentifier = True: hasLineMaterial = True
            itemOK = ContainsItem(line, itemNo)
            ' An explicit order in the line takes priority over the mail subject.
            If lineHasOrder And Not orderOK Then GoTo NextVendorRow
            If Not lineHasOrder And subjectHasOrder And Not orderOK Then GoTo NextVendorRow
            If lineHasItem And Not itemOK Then GoTo NextVendorRow
            score = 0
            If materialOK Then
                score = 3
            ElseIf orderOK And itemOK Then
                score = 3
            ElseIf orderOK And Not lineHasItem And InStr(line, vbTab) = 0 Then
                score = 1
            End If
            If score > 0 Then
                If score > bestScore Then
                    Set hits = New Collection
                    bestScore = score
                End If
                If score = bestScore Then hits.Add CStr(r)
            End If
NextVendorRow:
        Next rowKey
        If hits.Count = 1 Then
            activeRow = CStr(hits(1))
            If Not result.Exists(activeRow) Then
                result.Add activeRow, line
            Else
                result(activeRow) = result(activeRow) & vbCrLf & line
            End If
        ElseIf hits.Count > 1 Then
            If Len(activeRow) > 0 And Not lineHasOrder And Not lineHasItem And Not hasLineMaterial And ExtractNumbers(line).Count = 0 Then
                result(activeRow) = result(activeRow) & vbCrLf & line
            Else
                ' Multiple candidate items: never pick the first, never fan out to every row.
                ambiguous = True: activeRow = ""
            End If
        ElseIf lineHasOrder Or lineHasItem Or ExtractNumbers(line).Count > 0 Then
            activeRow = ""
        ElseIf Len(activeRow) > 0 Then
            result(activeRow) = result(activeRow) & vbCrLf & line
        End If
NextLine:
    Next oneLine
    ' Generic reply is safe only when exactly one candidate exists and no conflicting code is present.
    If result.Count = 0 And vendorCount = 1 And Not hasAnyIdentifier And Not ambiguous And Not subjectHasOrder Then
        If ExtractNumbers(processedReply & " " & subject).Count = 0 And Len(Trim$(processedReply)) > 0 Then
            result.Add CStr(vendorRows(1)), processedReply
        End If
    End If
    Set SelectReplyRows = result
End Function

Private Function HasKnownOrder(ByVal text As String, ByVal orders As Object) As Boolean
    Dim re As Object, token As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[A-Za-z0-9_-]+"
    For Each token In re.Execute(text)
        If orders.Exists(NormalizeOrderNo(CStr(token.Value))) Then HasKnownOrder = True: Exit Function
    Next token
End Function

Private Function ContainsCode(ByVal text As String, ByVal code As String) As Boolean
    Dim re As Object, normalized As String
    normalized = Trim$(code)
    If Len(normalized) = 0 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True: re.Global = False
    If normalized Like String$(Len(normalized), "#") Then
        normalized = NormalizeOrderNo(normalized)
        re.Pattern = "(^|[^A-Za-z0-9])0*" & EscapeRegex(normalized) & "([^A-Za-z0-9]|$)"
    Else
        re.Pattern = "(^|[^A-Za-z0-9])" & EscapeRegex(normalized) & "([^A-Za-z0-9]|$)"
    End If
    ContainsCode = re.Test(text)
End Function

Private Function HasItemLabel(ByVal text As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(^|[^a-z])(kalem|item|pozisyon)(\s*no)?\s*[:#-]?\s*0*[0-9]+([^0-9]|$)"
    HasItemLabel = re.Test(NormalizeName(text))
End Function

Private Function ContainsItem(ByVal text As String, ByVal code As String) As Boolean
    Dim re As Object, parts As Variant, part As Variant, normalized As String
    normalized = NormalizeOrderNo(code)
    If Len(normalized) = 0 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(^|[^a-z])(kalem|item|pozisyon)(\s*no)?\s*[:#-]?\s*0*" & EscapeRegex(normalized) & "([^0-9]|$)"
    If re.Test(NormalizeName(text)) Then ContainsItem = True: Exit Function
    If InStr(text, vbTab) > 0 Then
        parts = Split(text, vbTab)
        For Each part In parts
            If NormalizeOrderNo(Trim$(CStr(part))) = normalized Then ContainsItem = True: Exit Function
        Next part
    End If
End Function

Private Function IsSignatureStart(ByVal text As String) As Boolean
    Dim t As String
    t = NormalizeName(text)
    IsSignatureStart = (t = "SAYGILAR" Or t = "SAYGILARIMLA" Or t = "SAYGILARIMIZLA" Or _
        t = "IYI CALISMALAR" Or t = "BEST REGARDS" Or t = "KIND REGARDS" Or t = "REGARDS")
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
    Dim lines As Variant, oneLine As Variant, line As String, result As String
    Dim re As Object
    body = StripHtmlArtifacts(body)
    lines = Split(Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(on .+wrote:|.+tarihinde .+yazd[ýi]:|_{10,})$"
    For Each oneLine In lines
        line = Trim$(CStr(oneLine))
        If InStr(line, ":") > 0 Then
            If IsReplyHeader(Left$(line, InStr(line, ":") - 1)) Then Exit For
        End If
        If InStr(1, line, "Original Message", vbTextCompare) > 0 And Left$(line, 2) = "--" Then Exit For
        If InStr(1, NormalizeName(line), "ORIJINAL MESAJ", vbTextCompare) > 0 And Left$(line, 2) = "--" Then Exit For
        If re.Test(line) Then Exit For
        If Len(result) > 0 Then result = result & vbCrLf
        result = result & CStr(oneLine)
    Next oneLine
    ExtractNewReply = Trim$(result)
End Function

Private Function IsReplyHeader(ByVal label As String) As Boolean
    Dim t As String
    t = NormalizeName(label)
    IsReplyHeader = (t = "FROM" Or t = "GONDEREN" Or t = "KIMDEN" Or t = "SENT" Or _
        t = "GONDERILMIS" Or t = "GONDERILDI")
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
    Dim f As Object, parts As Variant, i As Long, first As String
    On Error GoTo Failed
    Set f = olNs.GetDefaultFolder(6)
    folderPath = Trim$(Replace(folderPath, "/", "\"))
    Do While Left$(folderPath, 1) = "\": folderPath = Mid$(folderPath, 2): Loop
    If Len(folderPath) = 0 Then Set GetInboxFolder = f: Exit Function
    parts = Split(folderPath, "\")
    first = CStr(parts(0)): i = 0
    If LCase$(first) = LCase$(f.Name) Or LCase$(first) = "inbox" Or NormalizeName(first) = "GELEN KUTUSU" Then
        i = 1
    Else
        Dim store As Object
        For Each store In olNs.Folders
            If LCase$(store.Name) = LCase$(first) Then
                Set f = store: i = 1: Exit For
            End If
        Next store
    End If
    For i = i To UBound(parts)
        If Len(Trim$(CStr(parts(i)))) > 0 Then Set f = f.Folders(CStr(parts(i)))
    Next i
    Set GetInboxFolder = f
    Exit Function
Failed:
    Set GetInboxFolder = Nothing
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
Private Function ReadMailFields(ByVal mail As Object, ByRef received As Date, _
    ByRef senderEmail As String, ByRef senderName As String, ByRef subject As String, _
    ByRef body As String) As Boolean
    On Error GoTo Failed
    received = 0: senderEmail = "": senderName = "": subject = "": body = ""
    If mail.Class <> 43 Then Exit Function
    received = mail.ReceivedTime
    senderEmail = NormalizeEmail(GetSenderSmtpEx(mail))
    senderName = SafeStr(mail.SenderName)
    subject = SafeStr(mail.Subject)
    ' HTML keeps table cells; plain Body is the fallback for plain-text mails.
    body = SafeStr(mail.HTMLBody)
    If Len(body) = 0 Then body = SafeStr(mail.Body)
    ReadMailFields = True
    Exit Function
Failed:
    ReadMailFields = False
End Function

Private Function NormalizeEmail(ByVal value As String) As String
    Dim leftPos As Long, rightPos As Long
    value = LCase$(Trim$(value))
    leftPos = InStr(value, "<"): rightPos = InStr(value, ">")
    If leftPos > 0 And rightPos > leftPos Then value = Mid$(value, leftPos + 1, rightPos - leftPos - 1)
    value = Trim$(Replace(value, "mailto:", ""))
    If InStr(value, "@") < 2 Or InStr(value, " ") > 0 Or InStr(value, "#") > 0 Then Exit Function
    NormalizeEmail = value
End Function

Private Sub AddVendorIdentity(ByVal dict As Object, ByVal identity As String, ByVal vendor As String)
    If dict.Exists(identity) Then
        If NormalizeName(CStr(dict(identity))) <> NormalizeName(vendor) Then dict(identity) = "***CAKISMA***"
    Else
        dict.Add identity, vendor
    End If
End Sub

Private Function IsPublicMailDomain(ByVal domain As String) As Boolean
    Select Case LCase$(domain)
        Case "gmail.com", "googlemail.com", "hotmail.com", "hotmail.com.tr", "outlook.com", _
             "outlook.com.tr", "live.com", "msn.com", "yahoo.com", "yahoo.com.tr", "icloud.com", "yandex.com", "yandex.com.tr"
            IsPublicMailDomain = True
    End Select
End Function

Private Function HasOrderLabel(ByVal text As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(^|[^a-z])(siparis|order|po|satinalma)(\s*(no|numarasi))?\s*[:#-]?\s*[0-9]{4,}([^0-9]|$)"
    HasOrderLabel = re.Test(NormalizeName(text))
End Function

' Run manually in Excel: no Outlook, no API, temporary workbook only.
Public Sub MailOkumaOzTest()
    Dim wb As Workbook, ws As Worksheet, cfg As Object, got As Object
    Dim passed As Long, errText As String, n As Long
    On Error GoTo Failed
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set cfg = CreateObject("Scripting.Dictionary")
    cfg("StartRow") = 2: cfg("OrderNoCol") = 1: cfg("KalemNoCol") = 2
    cfg("VendorCol") = 3: cfg("MalzemeCol") = 4
    cfg("ReplyCol") = 8: cfg("ReplyDateCol") = 9: cfg("ReplyFromCol") = 10
    cfg("MatchMethodCol") = 11: cfg("AttachCol") = 12: cfg("StatusCol") = 13
    cfg("DurumCol") = 14: cfg("TerminCol") = 15
    ws.Cells(2, 1).Value = "4500000001": ws.Cells(2, 2).Value = "00010"
    ws.Cells(2, 3).Value = "ALFA": ws.Cells(2, 4).Value = "MAT10001"
    ws.Cells(3, 1).Value = "4500000001": ws.Cells(3, 2).Value = "00020"
    ws.Cells(3, 3).Value = "ALFA": ws.Cells(3, 4).Value = "MAT20002"
    ws.Cells(4, 1).Value = "4500000002": ws.Cells(4, 2).Value = "10"
    ws.Cells(4, 3).Value = "BETA": ws.Cells(4, 4).Value = "MAT30003"
    ws.Cells(5, 1).Value = "4500000003": ws.Cells(5, 2).Value = "10"
    ws.Cells(5, 3).Value = "ALFA": ws.Cells(5, 4).Value = "MAT40004"

    Set got = SelectReplyRows(cfg, ws, "ALFA", "4500000001 kalem 10 yarin teslim", "")
    ReplyAssert got.Count = 1 And got.Exists("2"), "Siparis ve kalem", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "4500000001 yarin teslim", "")
    ReplyAssert got.Count = 0, "Belirsiz siparis iki kaleme yayilmamali", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "MAT20002 yarin teslim", "")
    ReplyAssert got.Count = 1 And got.Exists("3"), "Malzeme kodu", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "4500000001 MAT20002 yarin teslim", "")
    ReplyAssert got.Count = 1 And got.Exists("3"), "Siparis ve malzeme ayni kalemi secmeli", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "4500000002 kalem 10 yarin teslim", "")
    ReplyAssert got.Count = 0, "Baska tedarikciye yazilmamali", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "Kalem no: 20 yarin teslim", "RE: 4500000001")
    ReplyAssert got.Count = 1 And got.Exists("3"), "Konu siparisi ve yanit kalemi", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", _
        "4500000001" & vbTab & "00010" & vbTab & "11.11.2026" & vbCrLf & _
        "4500000001" & vbTab & "00020" & vbTab & "12.11.2026", "")
    ReplyAssert got.Count = 2, "Tabloda iki ayri kalem", passed
    ReplyAssert InStr(CStr(got("2")), "12.11.2026") = 0 And InStr(CStr(got("3")), "11.11.2026") = 0, "Tarihler kalemler arasinda karismamali", passed
    ws.Cells(3, 4).Value = "MAT10001"
    Set got = SelectReplyRows(cfg, ws, "ALFA", "MAT10001 yarin teslim", "")
    ReplyAssert got.Count = 0, "Tekrarlayan malzeme belirsiz", passed
    ws.Cells(3, 4).Value = "MAT20002"
    Set got = SelectReplyRows(cfg, ws, "ALFA", "Siparis 4500000001" & vbCrLf & "Kalem 20" & vbCrLf & "Yarin teslim", "")
    ReplyAssert got.Count = 1 And got.Exists("3"), "Onceki satirdan siparis baglami", passed
    ReplyAssert InStr(CStr(got("3")), "Yarin teslim") > 0, "Devam satiri kaybolmamali", passed
    ReplyAssert Not ContainsCode("145000000011", "4500000001"), "Kod alt dize olmamali", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "Yarin teslim edecegiz", "")
    ReplyAssert got.Count = 0, "Genel yanit coklu siparise yayilmamali", passed
    Set got = SelectReplyRows(cfg, ws, "BETA", "Yarin teslim edecegiz", "")
    ReplyAssert got.Count = 1 And got.Exists("4"), "Tek adayda genel yanit", passed
    Set got = SelectReplyRows(cfg, ws, "BETA", "Saygilarimla", "")
    ReplyAssert got.Count = 0, "Yalnizca imza olan yanit yazilmamali", passed
    ReplyAssert ExtractNewReply("Termin konusu: 12.11.2026" & vbCrLf & "From: buyer@example.invalid" & vbCrLf & "eski metin") = "Termin konusu: 12.11.2026", "Alintilanan eski mail ayrilmali", passed
    ReplyAssert InStr(StripHtmlArtifacts("<p>bir</p><p>iki</p>"), vbCrLf) > 0, "HTML paragraflari korunmali", passed
    ReplyAssert NormalizeEmail("Ali <ALI@example.invalid>") = "ali@example.invalid", "Gonderen adresi", passed
    ReplyAssert IsPublicMailDomain("gmail.com"), "Ortak domain eslestirilmemeli", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "MAT10001 yarin teslim", "Siparis 4500000002")
    ReplyAssert got.Count = 0, "Celisen konu siparisi", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "Siparis 4999999999 yarin teslim", "")
    ReplyAssert got.Count = 0, "Listede olmayan siparis", passed
    n = WriteReplyToSheet(cfg, ws, "ALFA", "4500000001 kalem 10 11.11.2026 teslim", DateSerial(2026, 11, 4), "test@example.invalid", "test", "", False)
    ReplyAssert n = 1 And Len(SafeStr(ws.Cells(3, 8).Value)) = 0 And Len(SafeStr(ws.Cells(4, 8).Value)) = 0, "Gercek yazma yalnizca hedef satira", passed
    n = WriteReplyToSheet(cfg, ws, "ALFA", "4500000001 kalem 10 10.11.2026 teslim", DateSerial(2026, 11, 3), "test@example.invalid", "test", "", False)
    ReplyAssert n = 0 And InStr(SafeStr(ws.Cells(2, 8).Value), "10.11.2026") = 0, "Eski mail yeni yaniti ezmemeli", passed
    wb.Close SaveChanges:=False
    MsgBox passed & " kontrol gecti. Outlook'a baglanilmadi; asil kitap degistirilmedi.", vbInformation
    Exit Sub
Failed:
    errText = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Oz test basarisiz: " & errText, vbCritical
End Sub

Private Sub ReplyAssert(ByVal condition As Boolean, ByVal title As String, ByRef passed As Long)
    If Not condition Then Err.Raise vbObjectError + 620, "MailOkumaOzTest", title
    passed = passed + 1
End Sub
