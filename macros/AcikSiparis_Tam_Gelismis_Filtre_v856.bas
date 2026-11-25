Attribute VB_Name = "OnlyReadReplies"
Option Explicit

'============================================================
'  ACIK SIPARIS MAIL OTOMASYONU - GELEN YANITLARI OKU
'  Surum: 8.5.6 (Ayni maildeki siparis yanitlarini ayirma)
'============================================================

Private Const APP_VERSION As String = "8.5.6"
Private Const REASON_PAIR_FAILED As String = "SIPARIS_KALEM_CIFTI_DOGRULANAMADI"
Private Const REASON_PAIR_MATCHED As String = "SIPARIS_KALEM_CIFTI"
Private Const TEST_MODE As Boolean = False
Private Const TEST_MAX_ORDERS As Long = 5
Private Const CONFIG_SHEET As String = "Ayarlar"
Private Const LOG_SHEET As String = "Log"
Private Const PREVIEW_LOG_SHEET As String = "FiltreOnizleme"
Private Const TERMIN_REPORT_SHEET As String = "TerminRaporu"
Private Const PROCESSED_SHEET As String = "IslenenMailler"
Private Const HEADER_ROW As Long = 1
Private Const DEFAULT_START_ROW As Long = 2
Private Const ATTACH_BASE_FOLDER As String = "YanitEkleri"
Private Const MAX_LOOKBACK_DAYS As Long = 180
Private Const MAX_BODY_LEN As Long = 32000
Private Const TABLE_END As String = "[TABLE_END]"
Private Const TERM_CONFIRMED As Long = 1
Private Const TERM_WITHDRAWN As Long = 2
Private Const TERM_UNCERTAIN As Long = 3

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

Private gCfg As Object
Private gLogWs As Worksheet
Private gLogRow As Long
Private gTermReportWs As Worksheet
Private gProcessedDict As Object
Private gInternalDomains As Object
Private gSiparisYokCount As Long
Private gBelirsizCount As Long
Private gMailHasVendorRows As Boolean
Private gMailMatchedRows As Long
Private gDeepFolders As Boolean
Private gSelectedRows As Object
Private gRowIndex As Object
Private gSentThreadLinks As Object
Private gOrderOnlySends As Object
Private gReplyTrackingReasons As Object
Private gUnmatchedVendorReplies As Object
Private gRegisteredVendors As Object
Private gVendorDirectoryNotes As Object
Private gSentScanFailed As Boolean
Private gCurrentMail As Object
Private gCurrentAttachments As String
Private gAttachmentsLoaded As Boolean
Private gOrigStatusBar As Variant
Private gOrigCancelKey As Long
Private gScanCutoff As Date

Private gOrigScreenUpdating As Boolean
Private gOrigEnableEvents As Boolean
Private gOrigCalculation As Long
Private gOrigDisplayStatusBar As Boolean
Private gStateSaved As Boolean
Private gPreview As Boolean

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
    gOrigStatusBar = Application.StatusBar
    gOrigCancelKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
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
    Application.StatusBar = gOrigStatusBar
    Application.EnableCancelKey = gOrigCancelKey
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
    
    re.Pattern = "<!--[\s\S]*?-->"
    s = re.Replace(s, "")
    re.Pattern = "<(html|body|div|p|table|tr|td|br)\b"
    If re.Test(s) Then
        re.Pattern = "[ \t]*[\r\n]+[ \t]*"
        s = re.Replace(s, " ")
        s = FlattenHtmlTableCells(s)
    End If
    re.Pattern = "<br\s*/?>|</p\s*>|</div\s*>|</li\s*>|</tr\s*>|<table\b[^>]*>"
    s = re.Replace(s, vbCrLf)

    re.Pattern = "</table\s*>"
    s = re.Replace(s, vbCrLf & TABLE_END & vbCrLf)
    re.Pattern = "<(div|section)\b[^>]*(divRplyFwdMsg|gmail_quote|yahoo_quoted)[^>]*>|<blockquote[^>]*>"
    s = re.Replace(s, vbCrLf & "[QUOTED_MESSAGE]" & vbCrLf)
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
    StripHtmlArtifacts = DecodeHtmlEntities(s)
End Function

Private Function FlattenHtmlTableCells(ByVal html As String) As String
    Static rows As Object, cells As Object, tags As Object
    Dim row As Object, cell As Object, result As String, nextPos As Long, rowText As String
    Dim value As String, firstCell As Boolean
    If rows Is Nothing Then
        Set rows = CreateObject("VBScript.RegExp")
        rows.Global = True: rows.IgnoreCase = True
        rows.Pattern = "<tr\b[^>]*>((?!<tr\b)[\s\S])*?</tr\s*>"
        Set cells = CreateObject("VBScript.RegExp")
        cells.Global = True: cells.IgnoreCase = True
        cells.Pattern = "<t[dh]\b[^>]*>([\s\S]*?)</t[dh]\s*>"
        Set tags = CreateObject("VBScript.RegExp")
        tags.Global = True: tags.IgnoreCase = True
        tags.Pattern = "<br\b[^>]*>|</(p|div|li)\s*>"
    End If
    nextPos = 1
    For Each row In rows.Execute(html)
        rowText = "": firstCell = True
        For Each cell In cells.Execute(CStr(row.Value))
            value = Trim$(tags.Replace(CStr(cell.SubMatches(0)), " "))
            If Not firstCell Then rowText = rowText & vbTab
            rowText = rowText & value
            firstCell = False
        Next cell
        result = result & Mid$(html, nextPos, row.FirstIndex + 1 - nextPos)
        If firstCell Then
            result = result & CStr(row.Value)
        Else
            result = result & vbCrLf & rowText & vbCrLf
        End If
        nextPos = row.FirstIndex + row.Length + 1
    Next row
    FlattenHtmlTableCells = result & Mid$(html, nextPos)
End Function

Private Function NormalizeOrderNo(ByVal s As String) As String
    Dim r As String
    r = Trim$(Replace(Replace(s, ChrW(160), " "), ChrW(8203), ""))
    r = Replace(r, vbTab, ""): r = Replace(r, " ", "")
    r = Replace(r, vbCr, ""): r = Replace(r, vbLf, "")
    If Left$(r, 1) = "'" Then r = Mid$(r, 2)
    r = Replace(r, "-", "")
    If Len(r) > 0 Then
        If r Like String$(Len(r), "#") Then
            Do While Len(r) > 1 And Left$(r, 1) = "0"
                r = Mid$(r, 2)
            Loop
        End If
    End If
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
    s = Replace(Replace(s, ChrW(160), " "), ChrW(8203), "")
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
    s = Replace(Replace(s, ChrW(160), " "), ChrW(8203), "")
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

Public Sub GelenYanitlariDerinOku()
    gDeepFolders = True
    GelenYanitlariOku
    gDeepFolders = False
End Sub

Public Sub GelenYanitlariOku()
    RunReplyScan False
End Sub

Public Sub GelenYanitlariOnizle()
    RunReplyScan True
End Sub

Public Sub GelenSeciliMailiOnizle()
    RunReplyScan True, True
End Sub

Public Sub GelenSeciliMailiIsle()
    RunReplyScan False, True
End Sub

Public Sub TedarikciAdresleriniOnizle()
    RunReplyScan True, vendorCheckOnly:=True
End Sub

Private Sub RunReplyScan(ByVal preview As Boolean, Optional ByVal useSelectedMail As Boolean = False, _
    Optional ByVal vendorCheckOnly As Boolean = False)
    Dim cfg As Object, fatalMessage As String, fatalLine As Long
    On Error GoTo FatalError
    gPreview = preview
    Set gLogWs = Nothing
    Set gRowIndex = CreateObject("Scripting.Dictionary")
    Set gSentThreadLinks = CreateObject("Scripting.Dictionary")
    Set gOrderOnlySends = CreateObject("Scripting.Dictionary")
    Set gReplyTrackingReasons = CreateObject("Scripting.Dictionary")
    Set gUnmatchedVendorReplies = CreateObject("Scripting.Dictionary")
    Set gRegisteredVendors = CreateObject("Scripting.Dictionary")
    gRegisteredVendors.CompareMode = vbTextCompare
    Set gVendorDirectoryNotes = CreateObject("Scripting.Dictionary")
    gVendorDirectoryNotes.CompareMode = vbTextCompare
    gUnmatchedVendorReplies.CompareMode = vbTextCompare
    gSentScanFailed = False
    Set gCurrentMail = Nothing
    Set gSelectedRows = CreateObject("Scripting.Dictionary")
    Set gProcessedDict = Nothing
    Set gInternalDomains = Nothing
    gSiparisYokCount = 0
    gBelirsizCount = 0
    gStateSaved = False

    Set cfg = GetConfig()
    cfg.RemoveAll
    If preview Then
        If Not LoadSettingsFromSheet(cfg) Then
            MsgBox "Onizleme icin gecerli kayitli ayarlar gerekli.", vbExclamation, AppTitle()
            GoTo Finish
        End If
    Else
        If Not LoadOrCollectSettings(cfg) Then GoTo Finish
    End If
    If ValidateSchema(cfg) Then DoReadReplies cfg, useSelectedMail, vendorCheckOnly
Finish:
    Set gCurrentMail = Nothing
    Set gRowIndex = Nothing
    Set gSentThreadLinks = Nothing
    Set gOrderOnlySends = Nothing
    Set gReplyTrackingReasons = Nothing
    Set gUnmatchedVendorReplies = Nothing
    Set gRegisteredVendors = Nothing
    Set gVendorDirectoryNotes = Nothing
    Set gSelectedRows = Nothing
    gDeepFolders = False
    gPreview = False
    Set gLogWs = Nothing
    AppStateEnd
    Exit Sub
FatalError:
    fatalMessage = Err.Description: fatalLine = Erl
    WriteLog "FATAL", "", "", "", "GelenYanitlariOku", fatalMessage
    MsgBox T("Kritik Hata (Satir ") & fatalLine & "):" & vbCrLf & fatalMessage & vbCrLf & vbCrLf & _
           T("Lutfen Excel sayfalarini ve sutun basliklarini kontrol edin."), vbCritical, AppTitle()
    Resume Finish
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
    Dim schemaSheets As Collection, schemaWs As Worksheet
    Set schemaSheets = cfg("OrderSheets")
    For Each schemaWs In schemaSheets
        If schemaWs Is cfg("VendorWs") Then msgs = msgs & "Siparis ve tedarikci sayfasi ayni olamaz." & vbCrLf
        If schemaWs.Name = LOG_SHEET Or schemaWs.Name = PREVIEW_LOG_SHEET Or schemaWs.Name = TERMIN_REPORT_SHEET Or _
            schemaWs.Name = CONFIG_SHEET Or schemaWs.Name = PROCESSED_SHEET Or schemaWs.Name = "YanitGecmisi" Then
            msgs = msgs & schemaWs.Name & " siparis sayfasi olarak secilemez." & vbCrLf
        End If
        If schemaWs.ProtectContents And Not gPreview Then msgs = msgs & schemaWs.Name & " sayfasi korumali." & vbCrLf
        For Each key In outputs
            col = CLng(cfg(CStr(key)))
            If col > 0 Then
                If IsProtectedOrderHeader(SafeStr(schemaWs.Cells(HEADER_ROW, col).Value)) Then
                    msgs = msgs & schemaWs.Name & " / " & NumToColLetter(col) & " kaynak siparis sutunu." & vbCrLf
                End If
            End If
        Next key
    Next schemaWs
    If Len(msgs) > 0 Then
        MsgBox msgs, vbCritical, AppTitle(): Exit Function
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
    cfg("LookbackDays") = ReadSettingLong(ws.Range("B29").Value, MAX_LOOKBACK_DAYS, 1, 3650)
    cfg("ScanSubfolders") = ReadSettingYesNo(ws.Range("B30").Value, "EVET")
    cfg("SaveAttachments") = ReadSettingYesNo(ws.Range("B31").Value, "EVET")
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
    Dim allSheets As Collection
    Set allSheets = cfg("OrderSheets")
    newVal = ResolveOutputColumn(cfg, allSheets, "DurumCol", _
        Array("Tedarikci Yanit Durumu", "Siparis Durumu", "Durum"))
    If cfg("DurumCol") <> newVal Then cfg("DurumCol") = newVal: changed = True
    newVal = ResolveOutputColumn(cfg, allSheets, "TerminCol", _
        Array("Teyitli Termin", "Termin", "Termin Tarihi"))
    If cfg("TerminCol") <> newVal Then cfg("TerminCol") = newVal: changed = True

    AutoFixAllColumns = changed
End Function

Private Function IsProtectedOrderHeader(ByVal header As String) As Boolean
    Dim raw As String
    raw = FoldTurkish(Trim$(header))
    If raw = "siparis_durumu" Or raw = "sat_kalem_no" Or Left$(raw, 6) = "detay_" Then
        IsProtectedOrderHeader = True
        Exit Function
    End If
    Select Case CleanForCompare(header)
        Case "SIPARISNO", "SIPARISNUMARASI", "KALEMNO", "SATKALEMNO", "TEDARIKCIADI", "MALZEMEKODU", _
            "DETAYSTATU", "ISTENENTARIH", "ISTENENTESLIMTARIHI", "TALEPEDILENTESLIMTARIHI", _
            "PLANLANANTESLIMTARIHI", "TESLIMTARIHI"
            IsProtectedOrderHeader = True
    End Select
End Function

Private Function ResolveOutputColumn(ByVal cfg As Object, ByVal sheets As Collection, _
    ByVal outputKey As String, ByVal acceptedHeaders As Variant) As Long
    Dim sheet As Worksheet, first As Worksheet, key As Variant, heading As Variant
    Dim reserved As Object, col As Long, lastCol As Long, firstLastCol As Long, configured As Long
    Dim lastCell As Range
    Set reserved = CreateObject("Scripting.Dictionary")
    For Each key In Array("OrderNoCol", "KalemNoCol", "VendorCol", "MalzemeCol", "UretimYeriCol", _
        "UzunMetinCol", "BirimCol", "BakiyeCol", "StatusCol", "ReplyCol", "ReplyDateCol", "ReplyFromCol", _
        "MatchMethodCol", "AttachCol", "DurumCol", "TerminCol")
        If CStr(key) <> outputKey Then
            col = CLng(cfg(CStr(key)))
            If col > 0 Then reserved(CStr(col)) = True
            If col > lastCol Then lastCol = col
        End If
    Next key
    configured = CLng(cfg(outputKey))
    If configured > 0 And configured <= 16384 Then
        If Not reserved.Exists(CStr(configured)) Then
            If OutputColumnCompatible(sheets, configured, acceptedHeaders) Then
                ResolveOutputColumn = configured
                Exit Function
            End If
        End If
    End If

    Set first = sheets(1)
    firstLastCol = first.Cells(HEADER_ROW, first.Columns.Count).End(xlToLeft).Column
    For Each heading In acceptedHeaders
        For col = 1 To firstLastCol
            If Not reserved.Exists(CStr(col)) Then
                If CleanForCompare(SafeStr(first.Cells(HEADER_ROW, col).Value)) = CleanForCompare(CStr(heading)) Then
                    If OutputColumnCompatible(sheets, col, acceptedHeaders) Then
                        ResolveOutputColumn = col
                        Exit Function
                    End If
                End If
            End If
        Next col
    Next heading
    For Each sheet In sheets
        Set lastCell = sheet.Cells.Find(What:="*", After:=sheet.Cells(1, 1), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
        If Not lastCell Is Nothing Then
            If lastCell.Column > lastCol Then lastCol = lastCell.Column
        End If
    Next sheet
    If lastCol >= 16384 Then Err.Raise vbObjectError + 845, , "Yanit durum/termin sutunu icin bos alan yok."
    ResolveOutputColumn = lastCol + 1
End Function

Private Function OutputColumnCompatible(ByVal sheets As Collection, ByVal col As Long, ByVal acceptedHeaders As Variant) As Boolean
    Dim sheet As Worksheet, header As String, item As Variant, recognized As Boolean
    For Each sheet In sheets
        header = SafeStr(sheet.Cells(HEADER_ROW, col).Value)
        If IsProtectedOrderHeader(header) Then Exit Function
        If Len(header) > 0 Then
            recognized = False
            For Each item In acceptedHeaders
                If CleanForCompare(header) = CleanForCompare(CStr(item)) Then recognized = True: Exit For
            Next item
            If Not recognized Then Exit Function
        ElseIf sheet.Cells(sheet.Rows.Count, col).End(xlUp).Row > HEADER_ROW Then
            Exit Function
        End If
    Next sheet
    OutputColumnCompatible = True
End Function

Private Sub SaveSettingsToSheet(ByRef cfg As Object)
    Dim ws As Worksheet, orderSheets As Collection, s As String, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add: ws.Name = CONFIG_SHEET
    On Error GoTo 0
    Dim retainedTestMode As Boolean: retainedTestMode = IsTestMode()
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
    
    ws.Range("A6").Value = T("Eski ayar (kullan~ilm~iyor; belirsiz e~sle~smeler atlan~ir)"): ws.Range("B6").Value = SafeStr(cfg("UpdateAllLines"))
    
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
    If retainedTestMode Then ws.Range("B26").Value = "EVET" Else ws.Range("B26").Value = "HAYIR"
    If cfg("DurumCol") > 0 Then
        ws.Range("A27").Value = T("Durum S~utunu"):   ws.Range("B27").Value = NumToColLetter(cfg("DurumCol"))
    End If
    If cfg("TerminCol") > 0 Then
        ws.Range("A28").Value = T("Termin S~utunu"):  ws.Range("B28").Value = NumToColLetter(cfg("TerminCol"))
    End If
    ws.Range("A29").Value = T("Arama ge~cmi~si (g~un, 1-3650)")
    ws.Range("B29").Value = ConfigLong(cfg, "LookbackDays", MAX_LOOKBACK_DAYS)
    ws.Range("A30").Value = T("Gelen kutusu alt klas~orlerini tara (EVET/HAYIR)")
    ws.Range("B30").Value = ConfigText(cfg, "ScanSubfolders", "EVET")
    ws.Range("A31").Value = T("E~sle~sen mail eklerini kaydet (EVET/HAYIR)")
    ws.Range("B31").Value = ConfigText(cfg, "SaveAttachments", "EVET")
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
    cfg("UpdateAllLines") = "HAYIR"
    cfg("LookbackDays") = MAX_LOOKBACK_DAYS
    cfg("ScanSubfolders") = "EVET"
    cfg("SaveAttachments") = "EVET"
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
    s = s & T("Belirsiz e~sle~smeler yaz~ilmaz. Arama (g~un): ") & ConfigLong(cfg, "LookbackDays", MAX_LOOKBACK_DAYS) & vbCrLf
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
        parts = Split(Replace(Replace(Replace(cfgVal, ";", ","), vbCr, ","), vbLf, ","), ",")
        For i = LBound(parts) To UBound(parts)
            Dim d As String: d = LCase$(Trim$(parts(i)))
            If Len(d) > 0 Then
                If InStr(d, "@") > 0 Then d = Mid$(d, InStr(d, "@") + 1)
                dict(d) = True
            End If
        Next i
    End If
    Set LoadInternalDomains = dict
End Function

Private Function IsInternalDomain(ByVal domain As String) As Boolean
    If Len(domain) = 0 Or gInternalDomains Is Nothing Then Exit Function
    If gInternalDomains.Exists(LCase$(domain)) Then IsInternalDomain = True
End Function

Private Sub DoReadReplies(ByRef cfg As Object, ByVal useSelectedMail As Boolean, ByVal vendorCheckOnly As Boolean)
    Dim vws As Worksheet
    Dim vendorNameCol As Long, vendorEmailCol As Long
    Dim orderSheets As Collection
    Dim olApp As Object, olNs As Object, inbox As Object, mail As Object
    Dim vendorByEmail As Object, vendorByDomain As Object, activeVendors As Object, emailCollisions As Object
    Dim vLastRow As Long, r As Long
    Dim vname As String, vemail As String, vdomain As String
    Dim matchCount As Long, replyCount As Long, scanCount As Long, skipCount As Long
    Dim internalSkipCount As Long, dateSkipCount As Long, collisionCount As Long, outsideCount As Long
    Dim i As Long, cutoffDate As Date, mailDate As Date
    Dim senderEmail As String, senderName As String, mailSubject As String, mailBody As String
    Dim senderDomain As String, vendorName As String, matchMethod As String
    Dim newReply As String, attachmentList As String
    Dim sheetUpdated As Long, upd As Long
    Dim updatedMailCount As Long, unknownVendorCount As Long
    Dim tMode As Boolean: tMode = IsTestMode()
    If gPreview Then tMode = False

    If tMode Then
        If MsgBox(T("TEST MODU ACIK: Ilk ") & TEST_MAX_ORDERS & T(" satir guncellenecek.") & vbCrLf & _
                  T("Devam edilsin mi?"), vbYesNo + vbQuestion, AppTitle()) <> vbYes Then Exit Sub
    End If

    AppStateBegin
    Set vws = cfg("VendorWs")
    vendorNameCol = cfg("VendorNameCol"): vendorEmailCol = cfg("VendorEmailCol")
    Set orderSheets = cfg("OrderSheets")

    WriteLog "SURUM", "", "", "", APP_VERSION
    WriteLog "AYARLAR", "", "", "", _
        "VendorAd=" & NumToColLetter(vendorNameCol) & _
        " | VendorMail=" & NumToColLetter(vendorEmailCol) & _
        " | Siparis=" & NumToColLetter(cfg("OrderNoCol")) & _
        " | Malzeme=" & NumToColLetter(cfg("MalzemeCol")) & _
        " | Kalem=" & NumToColLetter(cfg("KalemNoCol")) & _
        " | SiparisTedarikci=" & NumToColLetter(cfg("VendorCol")) & _
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
    Set emailCollisions = CreateObject("Scripting.Dictionary")
    emailCollisions.CompareMode = vbTextCompare
    Set activeVendors = ActiveOrderVendors(cfg, orderSheets)
    vendorByEmail.CompareMode = vbTextCompare
    vendorByDomain.CompareMode = vbTextCompare

    Dim loadedCount As Long: loadedCount = 0
    Dim addresses As Variant, address As Variant, previousVendor As String
    Dim collisionVendors As Object, vendorAliases As Object, vendorStates As Object, directoryTitles As Object
    Dim nameKey As String, rawAddresses As String
    Set vendorStates = CreateObject("Scripting.Dictionary")
    Set directoryTitles = CreateObject("Scripting.Dictionary")
    Set vendorAliases = BuildVendorAliases(vws, vLastRow, vendorNameCol, activeVendors, vendorStates, directoryTitles)
    For r = 2 To vLastRow
        vname = Trim$(SafeStr(vws.Cells(r, vendorNameCol).Value))
        nameKey = NormalizeName(vname)
        If Len(nameKey) > 0 And Not VendorIsBlocked(vname) Then
            If vendorAliases.Exists(nameKey) Then vname = CStr(vendorAliases(nameKey))
            rawAddresses = Trim$(Replace(SafeStr(vws.Cells(r, vendorEmailCol).Value), ChrW(160), " "))
            If Len(rawAddresses) > 0 Then
                If CLng(vendorStates(nameKey)) < 2 Then vendorStates(nameKey) = 2
            End If
            addresses = ExtractEmailList(rawAddresses)
            For Each address In addresses
                vemail = NormalizeEmail(CStr(address))
                If Len(vemail) > 0 Then
                    vendorStates(nameKey) = 3
                    gRegisteredVendors(NormalizeName(vname)) = True
                    If vendorByEmail.Exists(vemail) Then
                        previousVendor = CStr(vendorByEmail(vemail))
                        If NormalizeName(previousVendor) <> NormalizeName(vname) Then
                            If Not emailCollisions.Exists(vemail) Then
                                Set collisionVendors = CreateObject("Scripting.Dictionary")
                                collisionVendors.Add NormalizeName(previousVendor), previousVendor
                                emailCollisions.Add vemail, collisionVendors
                            End If
                            Set collisionVendors = emailCollisions(vemail)
                            collisionVendors(NormalizeName(vname)) = vname
                        End If
                    End If
                    AddVendorIdentity vendorByEmail, vemail, vname
                    vdomain = Mid$(vemail, InStr(vemail, "@") + 1)
                    AddVendorIdentity vendorByDomain, vdomain, vname
                    loadedCount = loadedCount + 1
                End If
            Next address
        End If
    Next r
    ReportVendorDirectory vws.Name, activeVendors, vendorStates, vendorAliases, directoryTitles, vendorCheckOnly
    WriteLog "TEDARIKCI_YUKLE", "", "", "", "Yuklenen adres=" & loadedCount & " | Kayit satiri=" & (vLastRow - 1), _
        "Secili tedarikci=" & activeVendors.Count & " | Benzersiz ad eslesmesi=" & vendorAliases.Count & _
        " | Adres/tedarikci uyarisi=" & gVendorDirectoryNotes.Count
    If vendorCheckOnly Then
        AppStateEnd
        MsgBox T("Tedarik~ci adresi kontrol~u tamamland~i.") & vbCrLf & _
               T("Kontrol edilen tedarik~ci: ") & activeVendors.Count & vbCrLf & _
               T("Uyar~i: ") & gVendorDirectoryNotes.Count & vbCrLf & _
               T("Ayr~int~ilar: ") & PREVIEW_LOG_SHEET, vbInformation, AppTitle()
        Exit Sub
    End If

    Dim checkSheet As Worksheet
    For i = 1 To orderSheets.Count
        Set checkSheet = orderSheets(i)
        WriteLog "SUTUN_KONTROL", checkSheet.Name, "", "", _
            "Siparis=" & NumToColLetter(CLng(cfg("OrderNoCol"))) & ":" & SafeStr(checkSheet.Cells(HEADER_ROW, CLng(cfg("OrderNoCol"))).Value) & _
            " | Kalem=" & NumToColLetter(CLng(cfg("KalemNoCol"))) & ":" & SafeStr(checkSheet.Cells(HEADER_ROW, CLng(cfg("KalemNoCol"))).Value), _
            "Tedarikci=" & NumToColLetter(CLng(cfg("VendorCol"))) & ":" & SafeStr(checkSheet.Cells(HEADER_ROW, CLng(cfg("VendorCol"))).Value) & _
            " | Durum=" & NumToColLetter(CLng(cfg("DurumCol")))
    Next i
    Set olApp = GetOutlookApp()
    If olApp Is Nothing Then
        AppStateEnd
        MsgBox T("Outlook'a ba~glan~ilamad~i."), vbCritical, AppTitle()
        Exit Sub
    End If
    Set olNs = olApp.GetNamespace("MAPI")
    If Not useSelectedMail Then
        Set inbox = GetInboxFolder(olNs, cfg("InboxFolder"))
        If inbox Is Nothing Then
            AppStateEnd
            MsgBox T("Gelen kutusu bulunamad~i."), vbCritical, AppTitle()
            Exit Sub
        End If
    End If

    Dim sortedItems As Object, quotedContext As String, filteredSkipCount As Long
    Dim explorer As Object, failedMailCount As Long, mailError As String
    If useSelectedMail Then
        Set explorer = olApp.ActiveExplorer
        If explorer Is Nothing Then
            MsgBox "Outlook'ta bir e-posta secin.", vbExclamation, AppTitle()
            Exit Sub
        End If
        If explorer.Selection.Count <> 1 Then
            MsgBox "Outlook'ta tek bir e-posta secin.", vbExclamation, AppTitle()
            Exit Sub
        End If
        Set mail = explorer.Selection.Item(1)
        Set sortedItems = New Collection
        sortedItems.Add mail
    End If

    Set gInternalDomains = LoadInternalDomains(cfg, olNs)
    If gPreview Then
        WriteLog "ONIZLEME_BASLAT", "", "", "", "Siparis verisi degistirilmez."
    Else
        If Not useSelectedMail Then EnsureTerminReportSheet
        For Each checkSheet In orderSheets
            If Len(SafeStr(checkSheet.Cells(HEADER_ROW, CLng(cfg("DurumCol"))).Value)) = 0 Then _
                checkSheet.Cells(HEADER_ROW, CLng(cfg("DurumCol"))).Value = "Tedarikci Yanit Durumu"
            If Len(SafeStr(checkSheet.Cells(HEADER_ROW, CLng(cfg("TerminCol"))).Value)) = 0 Then _
                checkSheet.Cells(HEADER_ROW, CLng(cfg("TerminCol"))).Value = "Teyitli Termin"
        Next checkSheet
    End If
    cutoffDate = DateAdd("d", -ConfigLong(cfg, "LookbackDays", MAX_LOOKBACK_DAYS), Now)
    gScanCutoff = cutoffDate
    If useSelectedMail Then
        WriteLog "BASLAT_SECILI_MAIL", "", "", "", CStr(IIf(gPreview, "Onizleme", "Yazma"))
    Else
        WriteLog "BASLAT", "", "", "", "GelenYanitlariOku", "Inbox: " & inbox.Name & " | Toplam: " & inbox.Items.Count
        If gDeepFolders Or ConfigText(cfg, "ScanSubfolders", "EVET") = "EVET" Then
            Set sortedItems = CollectDeepMails(inbox)
        Else
            Set sortedItems = RecentMailItems(inbox, cutoffDate, "ReceivedTime")
        End If
    End If
    Dim sentSource As Object, latestSent As Object
    If useSelectedMail Then
        Set sentSource = mail.Parent
    Else
        Set sentSource = inbox
    End If
    Set latestSent = ReadSentOrderLinks(cfg, sentSource, vendorByEmail, vendorByDomain, activeVendors, _
        DateValue(cutoffDate), gSentThreadLinks)
    For Each mail In sortedItems
        scanCount = scanCount + 1
        If scanCount Mod 100 = 0 Then Application.StatusBar = "Mail taraniyor: " & scanCount
        Dim entryId As String: entryId = ""
        On Error Resume Next
        entryId = SafeStr(mail.EntryID)
        On Error GoTo MailFailed
        vendorName = ""

        If Not ReadMailFields(mail, mailDate, senderEmail, senderName, mailSubject) Then
            If useSelectedMail Then WriteLog "MAIL_OKUNAMADI", "", "", "", "Secili oge bir e-posta olmayabilir."
            skipCount = skipCount + 1
            GoTo NextMail
        End If
        If mailDate < cutoffDate And Not useSelectedMail Then
            dateSkipCount = dateSkipCount + 1: skipCount = skipCount + 1
            GoTo NextMail
        End If
        senderDomain = ""
        If InStr(senderEmail, "@") > 0 Then senderDomain = Mid$(senderEmail, InStr(senderEmail, "@") + 1)

        If IsInternalDomain(senderDomain) Then
            If useSelectedMail Then WriteLog "DAHILI_GONDEREN", "", "", senderEmail, mailSubject
            internalSkipCount = internalSkipCount + 1
            skipCount = skipCount + 1
            GoTo NextMail
        End If

        vendorName = "": matchMethod = ""
        If Len(senderEmail) > 0 And vendorByEmail.Exists(senderEmail) Then
            vendorName = vendorByEmail(senderEmail)
            matchMethod = "e-posta"
        End If
        If Len(vendorName) = 0 And Len(senderDomain) > 0 And Not IsPublicMailDomain(senderDomain) Then
            If vendorByDomain.Exists(senderDomain) Then
                If vendorByDomain(senderDomain) <> "***CAKISMA***" Then
                    vendorName = vendorByDomain(senderDomain): matchMethod = "domain"
                End If
            End If
        End If

        If Len(vendorName) > 0 Then
            If vendorName <> "***CAKISMA***" Then
                If Not activeVendors.Exists(NormalizeName(vendorName)) Then
                    If useSelectedMail Then WriteLog "TABLO_DISI_TEDARIKCI", "", vendorName, senderEmail, mailSubject
                    outsideCount = outsideCount + 1
                    skipCount = skipCount + 1
                    GoTo NextMail
                End If
            End If
            If IsAutomaticReply(mail, mailSubject) Then
                If useSelectedMail Then WriteLog "OTOMATIK_YANIT", "", vendorName, senderEmail, mailSubject
                filteredSkipCount = filteredSkipCount + 1: skipCount = skipCount + 1
                GoTo NextMail
            End If
            If Not ReadMessageBody(mail, mailBody) Then
                WriteLog "GOVDE_OKUNAMADI", "", vendorName, senderEmail, "Konu: " & Left$(mailSubject, 160)
                skipCount = skipCount + 1
                GoTo NextMail
            End If
            newReply = CleanReplyContent(ExtractNewReply(mailBody))
            quotedContext = FirstQuotedMessage(mailBody)
            If Len(newReply) = 0 Then
                If ReadPlainMessageBody(mail, mailBody) Then
                    newReply = CleanReplyContent(ExtractNewReply(mailBody))
                    If Len(quotedContext) = 0 Then quotedContext = FirstQuotedMessage(mailBody)
                End If
            End If
            If vendorName = "***CAKISMA***" Then
                Set collisionVendors = emailCollisions(senderEmail)
                vendorName = ResolveVendorCollision(cfg, orderSheets, collisionVendors, newReply, mailSubject, quotedContext, mailDate)
                If Len(vendorName) = 0 Then
                    collisionCount = collisionCount + 1
                    skipCount = skipCount + 1
                    WriteLog "BELIRSIZ_GONDEREN", "", "", senderEmail, _
                        "Kayitli tedarikci=" & collisionVendors.Count & " | Tek siparis hedefi dogrulanamadi", _
                        "Konu: " & Left$(mailSubject, 160)
                    GoTo NextMail
                End If
                matchMethod = "e-posta / siparisle dogrulandi"
            End If
            replyCount = replyCount + 1
            If Len(Trim$(newReply)) = 0 Then
                gUnmatchedVendorReplies(NormalizeName(vendorName)) = True
                WriteLog "BOS_YANIT", "", vendorName, senderEmail, "Yeni yanit bos; satir degistirilmedi.", _
                    "Konu: " & Left$(mailSubject, 160)
                GoTo NextMail
            End If
            Set gCurrentMail = mail
            gCurrentAttachments = "": gAttachmentsLoaded = False
            attachmentList = ""
            sheetUpdated = 0
            gMailHasVendorRows = False: gMailMatchedRows = 0
            For i = 1 To orderSheets.Count
                upd = WriteReplyToSheet(cfg, orderSheets(i), vendorName, newReply, _
                    mailDate, senderEmail, matchMethod, attachmentList, tMode, mailSubject, quotedContext)
                sheetUpdated = sheetUpdated + upd
            Next i
            If sheetUpdated > 0 Then
                updatedMailCount = updatedMailCount + 1
            ElseIf Not gMailHasVendorRows Then
                gSiparisYokCount = gSiparisYokCount + 1
            ElseIf gMailMatchedRows = 0 Then
                gUnmatchedVendorReplies(NormalizeName(vendorName)) = True
                gBelirsizCount = gBelirsizCount + 1
            End If
            matchCount = matchCount + sheetUpdated
            If sheetUpdated > 0 And Len(entryId) > 0 And Not tMode And Not gPreview Then MarkMailProcessed entryId, vendorName
            If sheetUpdated > 0 Then
                WriteLog CStr(IIf(gPreview, "YANIT_ADAYI", "YANIT")), "", vendorName, senderEmail, matchMethod, _
                         "Konu: " & Left$(mailSubject, 120) & " | Satir: " & sheetUpdated
            End If
        Else
            unknownVendorCount = unknownVendorCount + 1: skipCount = skipCount + 1
            If useSelectedMail Or IsSiparisKonusu(mailSubject) Or ExtractNumbers(mailSubject).Count > 0 Then
                WriteLog "ESLESMEDI", "", senderName, senderEmail, _
                         "Konu: " & Left$(mailSubject, 80), "Tedarikci bulunamadi"
            End If
        End If
        GoTo NextMail
MailFailed:
        If Err.Number = 18 Then
            mailError = Err.Description
            On Error GoTo 0
            Err.Raise 18, , mailError
        End If
        mailError = CStr(Err.Number) & " - " & Err.Description
        failedMailCount = failedMailCount + 1
        skipCount = skipCount + 1
        WriteLog "MAIL_HATASI", "", vendorName, senderEmail, mailError, "Konu: " & Left$(mailSubject, 160)
        Resume NextMail
NextMail:
        On Error GoTo 0
    Next mail

    WriteLog "BITIR", "", "", "", _
             "GelenYanitlariOku | Taranan=" & scanCount & " | SeciliTedarikciMaili=" & replyCount & _
             " | UygunMail=" & updatedMailCount & " | UygunSatir=" & matchCount & " | Atlanan=" & skipCount, _
             "SeciliTabloDisi=" & outsideCount & " | TedarikciBilinmiyor=" & unknownVendorCount & _
             " | GonderenBelirsiz=" & collisionCount & " | Otomatik=" & filteredSkipCount & _
             " | Okuma/yazmaHatasi=" & failedMailCount
    gSiparisYokCount = gSiparisYokCount + outsideCount
    If gBelirsizCount > 0 Then
        WriteLog "OZET", "", "", "", "BELIRSIZ (e~sle~smeyen): " & gBelirsizCount
    End If

    If gPreview Then
        AppStateEnd
        MsgBox "Onizleme tamamlandi." & vbCrLf & "Taranan mail: " & scanCount & vbCrLf & _
               "Guncellemeye uygun satir: " & matchCount & vbCrLf & "Ayrintilar: FiltreOnizleme", _
               vbInformation, AppTitle()
        Exit Sub
    End If
    If useSelectedMail Then
        If gSelectedRows.Count > 0 Then ApplySentTracking cfg, latestSent, tMode, gSelectedRows
        AppStateEnd
        MsgBox "Secili e-posta tamamlandi." & vbCrLf & "Guncellenen satir: " & matchCount & _
               vbCrLf & "Ayrintilar: Log", vbInformation, AppTitle()
        Exit Sub
    End If
    ApplySentTracking cfg, latestSent, tMode
    Set gCurrentMail = Nothing
    WriteTerminReport cfg, scanCount, replyCount, matchCount, updatedMailCount
    AppStateEnd

    MsgBox T("Okuma tamamland~i.") & vbCrLf & vbCrLf & _
           T("Taranan mail: ") & scanCount & vbCrLf & _
           "Secili tablodaki tedarikcilerden okunan mail: " & replyCount & vbCrLf & _
           "Satir guncelleyen yanit: " & updatedMailCount & vbCrLf & _
           "Satir guncellemeyen yanit: " & (replyCount - updatedMailCount) & vbCrLf & _
           T("Excel'de g~uncellenen sat~ir: ") & matchCount & vbCrLf & _
           T("Belirsiz (e~sle~smeyen): ") & gBelirsizCount & vbCrLf & _
           "Atlanan (belirsiz gonderen): " & collisionCount & vbCrLf & _
           "Atlanan (otomatik yanit): " & filteredSkipCount & vbCrLf & _
           T("Atlanan (tarih): ") & dateSkipCount & vbCrLf & _
           T("Atlanan (dahili): ") & internalSkipCount & vbCrLf & _
           "Secili tablonun disinda kalan mail: " & gSiparisYokCount, _
           vbInformation, AppTitle()
End Sub

Private Function WriteReplyToSheet(ByRef cfg As Object, ByVal ws As Worksheet, _
    ByVal vendorName As String, ByVal newReply As String, ByVal mailDate As Date, _
    ByVal senderEmail As String, ByVal matchMethod As String, ByVal attachmentList As String, _
    ByVal tMode As Boolean, Optional ByVal mailSubject As String = "", _
    Optional ByVal quotedContext As String = "") As Long
    Dim matches As Object, key As Variant, r As Long, n As Long
    Dim cleanReply As String, reason As String, reasonParts As Variant
    Dim emptyCount As Long, oldCount As Long, supersededCount As Long, rowDetail As String
    Dim termState As Long, terminDate As Variant, statusText As String, decision As String
    Dim stockState As Long, stockDate As Variant, storedDate As Variant, storedReply As String
    Dim repairStockStatus As Boolean, stockRepairCount As Long, testLimitCount As Long
    Dim matchData As Variant, summaryRow As Boolean, noteOnly As Boolean
    Set matches = SelectDeepReplyRows(cfg, ws, vendorName, newReply, mailSubject, quotedContext, reason, mailDate)
    If matches.Count = 0 And InStr(1, reason, REASON_PAIR_FAILED, vbBinaryCompare) <> 1 Then
        If Not gCurrentMail Is Nothing Then
            Set matches = RecoverPlainTextMatches(cfg, ws, vendorName, mailSubject, newReply, quotedContext, mailDate, matches, reason)
            If matches.Count = 0 Then
                If AppendVerifiedThreadMatch(cfg, ws, vendorName, SafeConversationId(gCurrentMail), _
                    newReply, mailSubject, quotedContext, mailDate, matches, reason) Then reason = "GONDERILEN_MAIL_ZINCIRI"
            End If
        End If
    End If
    gMailMatchedRows = gMailMatchedRows + matches.Count
    If matches.Count = 0 Then
        If ws.Parent Is ThisWorkbook Then _
            WriteLog "ESLESME_NEDENI", ws.Name, vendorName, senderEmail, reason, "Konu: " & Left$(mailSubject, 160)
        Exit Function
    End If
    For Each key In matches.Keys
        r = CLng(key)
        If tMode And r >= CLng(cfg("StartRow")) + TEST_MAX_ORDERS Then
            testLimitCount = testLimitCount + 1
            GoTo NextMatch
        End If
        If ws.Parent Is ThisWorkbook Then
            If Not gSelectedRows Is Nothing Then
                If gSelectedRows.Exists(ws.CodeName & ":" & CStr(r)) Then
                    supersededCount = supersededCount + 1
                    GoTo NextMatch
                End If
            End If
        End If
        matchData = matches(key)
        summaryRow = IsArray(matchData)
        noteOnly = False
        If summaryRow Then
            cleanReply = CStr(matchData(0))
            noteOnly = CBool(matchData(1))
        Else
            cleanReply = SummarizeReply(CStr(matchData), mailDate)
        End If
        If Len(cleanReply) = 0 Then
            emptyCount = emptyCount + 1
            RememberReplyTrackingReason ws, r, mailDate, T("Yan~it bulundu fakat yaz~ilabilir sipari~s bilgisi yok.")
            GoTo NextMatch
        End If
        storedDate = ws.Cells(r, CLng(cfg("ReplyDateCol"))).Value
        storedReply = SafeStr(ws.Cells(r, CLng(cfg("ReplyCol"))).Value)
        If IsDate(storedDate) And Len(storedReply) > 0 Then
            If CDate(storedDate) > mailDate Then
                oldCount = oldCount + 1
                GoTo NextMatch
            End If
        End If
        termState = 0: stockState = 0
        If Not noteOnly Then terminDate = ResolveDeliveryDate(cleanReply, mailDate, termState)
        If summaryRow And termState <> TERM_CONFIRMED Then noteOnly = True
        If noteOnly Then
            termState = 0
            statusText = "Siparis notu"
        Else
            statusText = ClassifyStatus(cleanReply, termState = TERM_CONFIRMED)
        End If
        If Not noteOnly And IsStockAvailability(cleanReply) And termState <> TERM_CONFIRMED And statusText <> STATUS_IPTAL And _
           statusText <> STATUS_SEVK And statusText <> "Teslim Edildi" Then
            stockDate = ResolveDeliveryDate(cleanReply, mailDate, stockState, stockOnly:=True)
            If stockState = TERM_CONFIRMED Then
                statusText = "Stok tarihi: " & Format$(stockDate, "dd.mm.yyyy")
            End If
        End If
        If ReplyOutputHasFormula(ws, r, cfg, noteOnly) Then
            RememberReplyTrackingReason ws, r, mailDate, T("Yan~it bulundu fakat hedef h~ucre form~ull~u veya birle~sik.")
            If ws.Parent Is ThisWorkbook Then WriteLog "FORMUL_KORUNDU", ws.Name, vendorName, senderEmail, "Satir=" & r, _
                "Formul veya birlesik hucre | Konu: " & Left$(mailSubject, 140)
            GoTo NextMatch
        End If
        repairStockStatus = False
        If IsDate(storedDate) And Len(storedReply) > 0 Then
            If CDate(storedDate) = mailDate Then
                If stockState = TERM_CONFIRMED Then
                    repairStockStatus = CanRepairStockStatus(ws, r, cfg, storedReply, cleanReply, statusText)
                End If
                If Not repairStockStatus Then
                    oldCount = oldCount + 1
                    GoTo NextMatch
                End If
            End If
        End If
        If repairStockStatus Then
            If Not gPreview Then ws.Cells(r, CLng(cfg("DurumCol"))).Value2 = statusText
            stockRepairCount = stockRepairCount + 1
            GoTo RowAccepted
        End If
        If Not gPreview And Not tMode And (ws.Parent Is ThisWorkbook) And ConfigText(cfg, "SaveAttachments", "EVET") = "EVET" Then
            If Not gCurrentMail Is Nothing Then
                If Not gAttachmentsLoaded Then
                    gCurrentAttachments = SaveAttachments(gCurrentMail, vendorName)
                    gAttachmentsLoaded = True
                End If
                attachmentList = gCurrentAttachments
            End If
        End If
        reasonParts = Split(reason, " | ")
        If Not gPreview Then
            WriteReplyToRow ws, r, cleanReply, mailDate, senderEmail, matchMethod & " / " & CStr(reasonParts(0)), _
                attachmentList, CLng(cfg("ReplyCol")), CLng(cfg("ReplyDateCol")), _
                CLng(cfg("ReplyFromCol")), CLng(cfg("MatchMethodCol")), CLng(cfg("AttachCol")), _
                CLng(cfg("DurumCol")), CLng(cfg("TerminCol")), terminDate, termState, statusText, noteOnly
        End If
RowAccepted:
        If ws.Parent Is ThisWorkbook Then
            If Not gSelectedRows Is Nothing Then gSelectedRows(ws.CodeName & ":" & CStr(r)) = True
            rowDetail = "Satir=" & r & " | Siparis=" & SafeStr(ws.Cells(r, CLng(cfg("OrderNoCol"))).Value) & _
                " | Kalem=" & SafeStr(ws.Cells(r, CLng(cfg("KalemNoCol"))).Value)
            decision = reason & " | MailTarihi=" & Format$(mailDate, "dd.mm.yyyy hh:nn:ss") & _
                " | Durum=" & statusText
            If summaryRow Then decision = decision & " | Yanit=" & Left$(cleanReply, 250)
            If repairStockStatus Then decision = decision & " | Eski kayit: yalnizca stok durumu tamamlandi"
            Select Case termState
                Case TERM_CONFIRMED: decision = decision & " | Termin=" & Format$(terminDate, "dd.mm.yyyy")
                Case TERM_WITHDRAWN: decision = decision & " | Termin=iptal/teyitsiz"
                Case TERM_UNCERTAIN: decision = decision & " | Termin=belirsiz"
                Case Else: decision = decision & " | Yeni termin yok"
            End Select
            WriteLog CStr(IIf(gPreview, "SATIR_ADAYI", "SATIR_YAZILDI")), ws.Name, vendorName, senderEmail, rowDetail, decision
        End If
        n = n + 1
NextMatch:
    Next key
    If ws.Parent Is ThisWorkbook Then
        If emptyCount + oldCount + supersededCount + stockRepairCount + testLimitCount > 0 Then
            WriteLog "YANIT_SATIR_OZETI", ws.Name, vendorName, senderEmail, _
                "Aday=" & matches.Count & " | Uygun=" & n & " | YanitVerisiYok=" & emptyCount & _
                " | KayitDahaYeniVeyaAyni=" & oldCount & " | BuTaramadaYeniYanitVar=" & supersededCount & _
                " | TamamlananStokDurumu=" & stockRepairCount & " | TestSiniri=" & testLimitCount, _
                "Konu: " & Left$(mailSubject, 160)
        End If
    End If
    WriteReplyToSheet = n
End Function

Private Sub RememberReplyTrackingReason(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal mailDate As Date, ByVal reason As String)
    Dim rowKey As String, recorded As Variant
    If gReplyTrackingReasons Is Nothing Then Exit Sub
    If Not (ws.Parent Is ThisWorkbook) Then Exit Sub
    rowKey = ws.CodeName & ":" & CStr(rowNumber)
    If gReplyTrackingReasons.Exists(rowKey) Then
        recorded = gReplyTrackingReasons(rowKey)
        If CDate(recorded(0)) >= mailDate Then Exit Sub
    End If
    gReplyTrackingReasons(rowKey) = Array(mailDate, reason)
End Sub

Private Function RecoverPlainTextMatches(ByRef cfg As Object, ByVal ws As Worksheet, _
    ByVal vendor As String, ByVal subject As String, ByVal reply As String, ByVal quoted As String, _
    ByVal mailDate As Date, ByVal originalMatches As Object, ByRef reason As String) As Object
    Dim body As String, plainReply As String, plainQuote As String, candidate As Object, alternativeReason As String
    Set RecoverPlainTextMatches = originalMatches
    If InStr(1, reason, "CELISIYOR", vbTextCompare) > 0 Then Exit Function
    If InStr(1, reason, "TEDARIKCI_SATIRI_YOK", vbTextCompare) = 1 Then Exit Function
    If Not ReadPlainMessageBody(gCurrentMail, body) Then Exit Function
    plainReply = CleanReplyContent(ExtractNewReply(body))
    plainQuote = FirstQuotedMessage(body)

    ' Prefer the original new text; a plain-text quote can repair a broken HTML thread boundary.
    If Len(plainQuote) > 0 And StrComp(plainQuote, quoted, vbTextCompare) <> 0 Then
        Set candidate = SelectDeepReplyRows(cfg, ws, vendor, reply, subject, plainQuote, alternativeReason, mailDate)
        If candidate.Count > 0 Then
            reason = "DUZ_METIN_ALINTISI | " & alternativeReason
            Set RecoverPlainTextMatches = candidate
            Exit Function
        End If
    End If
    If Not IsUsefulReply(plainReply) Then Exit Function
    If StrComp(NormalizeName(plainReply), NormalizeName(reply), vbTextCompare) = 0 Then Exit Function
    If ReplyHasExplicitCodes(reply) Then Exit Function
    If IsUsefulReply(reply) And Not IsStockAvailability(reply) Then Exit Function
    If IsStockAvailability(reply) And Not IsStockAvailability(plainReply) Then Exit Function
    If Len(plainQuote) = 0 Then plainQuote = quoted
    Set candidate = SelectDeepReplyRows(cfg, ws, vendor, plainReply, subject, plainQuote, alternativeReason, mailDate)
    If candidate.Count > 0 Then
        reason = "DUZ_METIN_YANITI | " & alternativeReason
        Set RecoverPlainTextMatches = candidate
    End If
End Function

Private Function CanRepairStockStatus(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal cfg As Object, _
    ByVal storedReply As String, ByVal cleanReply As String, ByVal statusText As String) As Boolean
    Dim col As Long, previous As String
    col = CLng(cfg("DurumCol"))
    If col < 1 Then Exit Function
    If StrComp(NormalizeName(storedReply), NormalizeName(cleanReply), vbTextCompare) <> 0 Then Exit Function
    previous = SafeStr(ws.Cells(rowNumber, col).Value2)
    If Len(previous) = 0 Or previous = STATUS_BELIRSIZ Then
        CanRepairStockStatus = (previous <> statusText)
    End If
End Function

Private Function ReplyOutputHasFormula(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal cfg As Object, _
    Optional ByVal noteOnly As Boolean = False) As Boolean
    Dim key As Variant, col As Long
    For Each key In Array("ReplyCol", "ReplyDateCol", "ReplyFromCol", "MatchMethodCol", "AttachCol", "DurumCol", "TerminCol", "StatusCol")
        If noteOnly Then
            If key = "DurumCol" Or key = "TerminCol" Or key = "StatusCol" Then GoTo NextFormulaCheck
        End If
        col = CLng(cfg(CStr(key)))
        If col > 0 Then
            If ws.Cells(rowNumber, col).HasFormula Or ws.Cells(rowNumber, col).MergeCells Then
                ReplyOutputHasFormula = True
                Exit Function
            End If
        End If
NextFormulaCheck:
    Next key
End Function

Private Function SelectReplyRows(ByRef cfg As Object, ByVal ws As Worksheet, _
    ByVal vendorName As String, ByVal reply As String, ByVal subject As String, _
    Optional ByRef reason As String = "", Optional ByVal incoming As Boolean = False) As Object
    Dim result As Object, vendorRows As Collection, allOrders As Object
    Dim r As Long, lastRow As Long, key As Variant, rowKey As Variant
    Dim lines As Variant, oneLine As Variant, line As String, activeRow As String
    Dim hits As Collection, orderNo As String, itemNo As String
    Dim description As String, descriptionOK As Boolean, normalizedLine As String
    Dim rowDetails As Variant, rowData As Variant
    Dim lineHasOrder As Boolean, subjectHasOrder As Boolean, lineHasItem As Boolean
    Dim orderContext As String, processedReply As String
    Dim useSubjectItem As Boolean, blockedContinuation As Boolean, hasLineIdentity As Boolean
    Dim lineOrderCount As Long, subjectOrderContext As Boolean, subjectHasItem As Boolean
    Dim subjectRows As Long, subjectCodes As Object
    Dim orderOK As Boolean, itemOK As Boolean, materialOK As Boolean
    Dim hasAnyIdentifier As Boolean, ambiguous As Boolean, vendorCount As Long
    Dim hasLineMaterial As Boolean, score As Long, bestScore As Long
    Dim lineMaterials As Object, materialRow As Variant, knownMaterial As String
    Dim explicitMaterial As Boolean
    Dim tableColumns As Object, rawLine As String, replyLine As String, tableRow As Boolean
    Dim lineItems As Object, activeGroup As Collection
    Dim subjectPairs As Object, linePairs As Object, contextPairs As Object, seenPairs As Object, pairKey As Variant
    Dim pairMatched As Boolean
    Set tableColumns = CreateObject("Scripting.Dictionary")
    Set result = CreateObject("Scripting.Dictionary")
    LoadSheetRowIndex cfg, ws, vendorName, vendorRows, allOrders, rowDetails:=rowDetails
    vendorCount = vendorRows.Count
    If vendorCount = 0 Then
        lastRow = ws.Cells(ws.Rows.Count, CLng(cfg("OrderNoCol"))).End(xlUp).Row
        reason = "TEDARIKCI_SATIRI_YOK | TedarikciSutunu=" & NumToColLetter(CLng(cfg("VendorCol"))) & _
            " | SiparisSutunu=" & NumToColLetter(CLng(cfg("OrderNoCol"))) & " | SonSatir=" & lastRow
        Set SelectReplyRows = result: Exit Function
    End If
    gMailHasVendorRows = True
    subject = NormalizeSasReferences(subject)
    Set subjectCodes = ExtractNumbers(subject)
    Set subjectPairs = ReadOrderItemPairs(subject)
    Set contextPairs = subjectPairs
    Set seenPairs = CreateObject("Scripting.Dictionary")
    For Each pairKey In subjectPairs.Keys
        seenPairs(pairKey) = True
    Next pairKey
    For Each rowKey In vendorRows
        rowData = rowDetails(CStr(rowKey))
        If ContainsCode(subject, CStr(rowData(0))) Then subjectRows = subjectRows + 1
    Next rowKey
    reply = CleanReplyContent(reply)
    orderContext = subject
    subjectHasOrder = (subjectRows > 0 Or SubjectHasOrderReference(subject, allOrders))
    subjectOrderContext = subjectHasOrder
    subjectHasItem = (HasItemLabel(subject) Or HasOrderItemPair(subject))
    lines = Split(Replace(Replace(reply, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For Each oneLine In lines
        rawLine = Trim$(CStr(oneLine))
        If rawLine = TABLE_END Then
            tableColumns.RemoveAll
            activeRow = ""
            Set activeGroup = Nothing
            orderContext = subject
            Set contextPairs = subjectPairs
            subjectHasOrder = (subjectRows > 0 Or SubjectHasOrderReference(subject, allOrders))
            subjectOrderContext = subjectHasOrder
            blockedContinuation = True
            GoTo NextLine
        End If
        line = ExpandTableMatchLine(rawLine, tableColumns)
        If line = "[TABLE_HEADER]" Then
            activeRow = ""
            Set activeGroup = Nothing
            orderContext = subject
            Set contextPairs = subjectPairs
            subjectHasOrder = (subjectRows > 0 Or SubjectHasOrderReference(subject, allOrders))
            subjectOrderContext = subjectHasOrder
            blockedContinuation = True
            GoTo NextLine
        End If
        If IsSignatureStart(line) Then Exit For
        If Len(line) = 0 Then
            If InStr(rawLine, vbTab) > 0 Then
                activeRow = ""
                blockedContinuation = True
            End If
            GoTo NextLine
        End If
        replyLine = rawLine
        line = NormalizeSasReferences(line)
        tableRow = (InStr(rawLine, vbTab) > 0)
        If incoming And tableRow And tableColumns.Count > 0 Then replyLine = ReadTableReply(rawLine, tableColumns)
        If Len(processedReply) > 0 Then processedReply = processedReply & vbCrLf
        processedReply = processedReply & rawLine
        lineOrderCount = CountKnownOrders(line, allOrders)
        lineHasOrder = (lineOrderCount > 0 Or HasOrderLabel(line) Or StartsWithOrderNumber(line))
        If lineHasOrder Then
            orderContext = line
            subjectHasOrder = True
            subjectOrderContext = False
        End If
        Set lineItems = Nothing
        If Not tableRow And InStr(1, line, "kalem", vbTextCompare) > 0 Then
            Set lineItems = ExplicitItemList(line)
            If lineItems.Count < 2 Then Set lineItems = Nothing
        End If
        lineHasItem = (HasItemLabel(line) Or HasOrderItemPair(line) Or Not lineItems Is Nothing)
        explicitMaterial = HasMaterialCodeLabel(line)
        hasLineIdentity = (lineHasOrder Or lineHasItem Or explicitMaterial)
        Set linePairs = ReadOrderItemPairs(line)
        If Not lineItems Is Nothing Then
            ' A labelled list under one order already verifies each of its named items.
            If linePairs.Count = 1 And lineOrderCount = 1 And HasItemLabel(line) And Not HasOrderItemPair(line) Then _
                linePairs.RemoveAll
        End If
        If linePairs.Count = 0 And lineItems Is Nothing Then
            If lineHasItem And Not lineHasOrder And Not tableRow And subjectHasOrder Then
                If Not HasItemLabel(orderContext) And Not HasOrderItemPair(orderContext) Then _
                    Set linePairs = ReadOrderItemPairs(orderContext & vbLf & line)
            ElseIf Not lineHasOrder And Not lineHasItem And Not tableRow Then
                Set linePairs = contextPairs
            End If
        End If
        If lineHasOrder Or (lineHasItem And linePairs.Count > 0) Then Set contextPairs = linePairs
        If linePairs.Count > 0 Then
            For Each pairKey In linePairs.Keys
                If seenPairs.Count >= 16 Then Exit For
                seenPairs(pairKey) = True
            Next pairKey
            Set lineItems = Nothing
            If linePairs.Count > 1 Then
                ambiguous = True
                activeRow = ""
                Set activeGroup = Nothing
                blockedContinuation = True
                GoTo NextLine
            End If
        End If
        If hasLineIdentity Then hasAnyIdentifier = True
        Set lineMaterials = CreateObject("Scripting.Dictionary")
        For Each materialRow In vendorRows
            rowData = rowDetails(CStr(materialRow))
            knownMaterial = CStr(rowData(2))
            If Len(knownMaterial) >= 5 Then
                If ContainsCode(line, knownMaterial) Then lineMaterials(CStr(materialRow)) = True
            End If
        Next materialRow
        If lineMaterials.Count > 0 Then hasLineIdentity = True
        If Not activeGroup Is Nothing Then
            If Not hasLineIdentity And Not tableRow And Not ExplicitWholeOrder(rawLine) Then
                If ExtractNumbers(line).Count = 0 Then
                    For Each key In activeGroup
                        result(CStr(key)) = result(CStr(key)) & vbCrLf & replyLine
                    Next key
                    GoTo NextLine
                End If
            End If
            Set activeGroup = Nothing
        End If
        If blockedContinuation And Not hasLineIdentity And Not ExplicitWholeOrder(rawLine) Then GoTo NextLine
        If lineOrderCount > 1 Or (ExplicitWholeOrder(rawLine) And _
           (CountKnownOrders(orderContext, allOrders) > 1 Or ExtractNumbers(orderContext).Count > 1)) Then
            ambiguous = True
            activeRow = ""
            blockedContinuation = True
            GoTo NextLine
        End If
        useSubjectItem = (subjectOrderContext And Not hasLineIdentity And Not tableRow And Not ExplicitWholeOrder(rawLine))
        Set hits = New Collection
        hasLineMaterial = False: bestScore = 0
        normalizedLine = " " & NormalizeName(line) & " "
        For Each rowKey In vendorRows
            r = CLng(rowKey)
            rowData = rowDetails(CStr(rowKey))
            orderNo = CStr(rowData(0))
            itemNo = CStr(rowData(1))
            orderOK = False: itemOK = False: materialOK = False: descriptionOK = False
            description = CStr(rowData(3))
            If Len(description) >= 20 And InStr(description, " ") > 0 Then _
                descriptionOK = (InStr(normalizedLine, " " & description & " ") > 0)
            If lineHasOrder Then
                orderOK = ContainsCode(line, orderNo)
            ElseIf subjectHasOrder Then
                orderOK = ContainsCode(orderContext, orderNo)
            End If
            materialOK = lineMaterials.Exists(CStr(rowKey))
            If materialOK Or descriptionOK Then hasAnyIdentifier = True: hasLineMaterial = True
            If lineItems Is Nothing Then
                itemOK = ContainsItem(line, itemNo, orderNo)
            Else
                itemOK = lineItems.Exists(NormalizeOrderNo(itemNo))
            End If
            If useSubjectItem And orderOK Then itemOK = ContainsItem(subject, itemNo, orderNo)
            If linePairs.Count > 0 Then
                If Not linePairs.Exists(OrderItemKey(orderNo, itemNo)) Then GoTo NextVendorRow
                orderOK = True: itemOK = True
            Else
                ' An explicit order in the line takes priority over the mail subject.
                If lineHasOrder And Not orderOK Then GoTo NextVendorRow
                If Not lineHasOrder And subjectHasOrder And Not orderOK Then GoTo NextVendorRow
                If (lineHasItem Or (useSubjectItem And subjectHasItem)) And Not itemOK Then GoTo NextVendorRow
            End If
            ' A named material is a hard constraint, even if order/item labels suggest another row.
            If lineMaterials.Count > 0 And Not materialOK Then GoTo NextVendorRow
            If explicitMaterial And Not materialOK Then GoTo NextVendorRow
            score = 0
            If materialOK Then
                score = 3
            ElseIf orderOK And itemOK Then
                score = 4
            ElseIf descriptionOK Then
                score = 2
            ElseIf orderOK And Not lineHasItem And Not tableRow Then
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
        If Not lineItems Is Nothing Then
            activeRow = ""
            If ExactItemGroup(hits, rowDetails, lineItems.Count) Then
                Set activeGroup = hits
                blockedContinuation = False
                For Each key In activeGroup
                    If result.Exists(CStr(key)) Then
                        result(CStr(key)) = result(CStr(key)) & vbCrLf & replyLine
                    Else
                        result.Add CStr(key), replyLine
                    End If
                Next key
            Else
                ambiguous = True
                blockedContinuation = True
            End If
            GoTo NextLine
        End If
        If hits.Count = 1 And linePairs.Count > 0 Then
            rowData = rowDetails(CStr(hits(1)))
            knownMaterial = ""
            If lineMaterials.Exists(CStr(hits(1))) Then knownMaterial = CStr(rowData(2))
            pairKey = OrderItemKey(CStr(rowData(0)), CStr(rowData(1)))
            If PairExistsOnAnotherSheet(cfg, ws, vendorName, CStr(pairKey), knownMaterial) Then
                Set hits = New Collection
                ambiguous = True
                activeRow = ""
                blockedContinuation = True
                GoTo NextLine
            End If
        End If
        If hits.Count = 1 Then
            If linePairs.Count > 0 Then pairMatched = True
            activeRow = CStr(hits(1))
            blockedContinuation = False
            If Not result.Exists(activeRow) Then
                result.Add activeRow, replyLine
            Else
                result(activeRow) = result(activeRow) & vbCrLf & replyLine
            End If
        ElseIf hits.Count > 1 Then
            If ExplicitWholeOrder(rawLine) And subjectHasOrder And Not lineHasItem And Not hasLineMaterial And _
               linePairs.Count = 0 And CountKnownOrders(orderContext, allOrders) = 1 And ExtractNumbers(orderContext).Count <= 1 Then
                For Each key In hits
                    If Not result.Exists(CStr(key)) Then
                        result.Add CStr(key), replyLine
                    Else
                        result(CStr(key)) = result(CStr(key)) & vbCrLf & replyLine
                    End If
                Next key
                activeRow = ""
            ElseIf Len(activeRow) > 0 And Not hasLineIdentity And Not hasLineMaterial And Not tableRow And ExtractNumbers(line).Count = 0 Then
                result(activeRow) = result(activeRow) & vbCrLf & replyLine
            Else
                ' Multiple candidate items: never pick the first, never fan out to every row.
                ambiguous = True: activeRow = ""
                blockedContinuation = True
            End If
        ElseIf tableRow Or hasLineIdentity Or ExtractNumbers(line).Count > 0 Then
            activeRow = ""
            blockedContinuation = True
        ElseIf Len(activeRow) > 0 Then
            result(activeRow) = result(activeRow) & vbCrLf & replyLine
        End If
NextLine:
    Next oneLine
    ' Incoming replies require a target from the new reply, subject or quoted order.
    If Not incoming And result.Count = 0 And vendorCount = 1 And Not hasAnyIdentifier And Not ambiguous And Not subjectHasOrder And seenPairs.Count = 0 Then
        If ExtractNumbers(processedReply & " " & subject).Count = 0 And Len(Trim$(processedReply)) > 0 Then
            result.Add CStr(vendorRows(1)), processedReply
        End If
    End If
    If result.Count > 0 Then
        If pairMatched Then
            reason = REASON_PAIR_MATCHED
        Else
            reason = "ESLESTI"
        End If
    ElseIf seenPairs.Count > 0 Then
        reason = REASON_PAIR_FAILED & " | Cift=" & Left$(Join(seenPairs.Keys, ", "), 320)
        If ambiguous Then reason = reason & " | BIRDEN_FAZLA_KALEM_ADAYI"
    ElseIf SubjectHasOrderReference(subject, allOrders) And subjectRows = 0 Then
        reason = "KONU_SIPARISI_TEDARIKCIDE_YOK"
    ElseIf ambiguous Then
        reason = "BIRDEN_FAZLA_KALEM_ADAYI"
    ElseIf Not hasAnyIdentifier And Not subjectHasOrder Then
        reason = "YANITTA_SIPARIS_KALEM_MALZEME_YOK"
    Else
        reason = "KOD_VEYA_KALEM_ESLESMEDI"
    End If
    reason = reason & " | TedarikciSatiri=" & vendorCount & " | KonuKodSayisi=" & subjectCodes.Count & _
        " | KonuSiparisSatiri=" & subjectRows & " | SonucSatiri=" & result.Count
    Set SelectReplyRows = result
End Function

Private Function OrderItemKey(ByVal orderNo As String, ByVal itemNo As String) As String
    OrderItemKey = NormalizeOrderNo(orderNo) & "/" & NormalizeOrderNo(itemNo)
End Function

Private Function ReadOrderItemPairs(ByVal text As String) As Object
    Static pattern As Object
    Dim pairs As Object, match As Object
    Set pairs = CreateObject("Scripting.Dictionary")
    If pattern Is Nothing Then
        Set pattern = CreateObject("VBScript.RegExp")
        pattern.Global = True: pattern.IgnoreCase = True
        pattern.Pattern = "\b([0-9]{6,16})(?:\s*[/\-]\s*|\s*[·|;,:/\-]?\s*(?:sat[ _]+)?(?:kalem|item|pozisyon)" & _
            "(?:[ _]*(?:no|numarasi|number))?\s*[.:#=\-]*\s*[\(\[]?\s*)([0-9]{1,5})\b(?![./-][0-9])"
    End If
    For Each match In pattern.Execute(FoldTurkish(text))
        pairs(OrderItemKey(CStr(match.SubMatches(0)), CStr(match.SubMatches(1)))) = True
    Next match
    Set ReadOrderItemPairs = pairs
End Function

Private Function PairExistsOnAnotherSheet(ByVal cfg As Object, ByVal currentSheet As Worksheet, _
    ByVal vendor As String, ByVal pair As String, ByVal material As String) As Boolean
    Dim sheets As Collection, ws As Worksheet, rows As Collection, orders As Object
    Dim details As Variant, row As Variant, values As Variant, otherMaterial As String
    If Not cfg.Exists("OrderSheets") Then Exit Function
    If Not (currentSheet.Parent Is ThisWorkbook) Then Exit Function
    Set sheets = cfg("OrderSheets")
    If sheets.Count < 2 Then Exit Function
    For Each ws In sheets
        If Not (ws Is currentSheet) Then
            LoadSheetRowIndex cfg, ws, vendor, rows, orders, rowDetails:=details
            For Each row In rows
                values = details(CStr(row))
                If OrderItemKey(CStr(values(0)), CStr(values(1))) = pair Then
                    otherMaterial = CStr(values(2))
                    If Len(material) = 0 Or Len(otherMaterial) = 0 Or StrComp(material, otherMaterial, vbTextCompare) = 0 Then
                        PairExistsOnAnotherSheet = True
                        Exit Function
                    End If
                End If
            Next row
        End If
    Next ws
End Function

Private Function ExactItemGroup(ByVal hits As Collection, ByVal details As Object, ByVal wantedCount As Long) As Boolean
    Dim key As Variant, rowData As Variant, currentOrder As String, expectedOrder As String
    Dim itemNo As String, items As Object
    If hits.Count <> wantedCount Then Exit Function
    Set items = CreateObject("Scripting.Dictionary")
    For Each key In hits
        rowData = details(CStr(key))
        currentOrder = NormalizeOrderNo(CStr(rowData(0)))
        itemNo = NormalizeOrderNo(CStr(rowData(1)))
        If Len(currentOrder) = 0 Or Len(itemNo) = 0 Then Exit Function
        If Len(expectedOrder) = 0 Then expectedOrder = currentOrder
        If expectedOrder <> currentOrder Then Exit Function
        items(itemNo) = True
    Next key
    ExactItemGroup = (items.Count = wantedCount)
End Function

Private Function SubjectHasOrderReference(ByVal subject As String, ByVal orders As Object) As Boolean
    Static re As Object
    If HasOrderLabel(subject) Or HasKnownOrder(subject, orders) Then
        SubjectHasOrderReference = True
        Exit Function
    End If
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^[0-9]{6,12}(\s*[,;/]\s*[0-9]{6,12})*$"
    End If
    SubjectHasOrderReference = re.Test(NormalizeReplySubject(subject))
End Function

Private Function NormalizeSasReferences(ByVal text As String) As String
    Static re As Object
    If InStr(1, text, "sas", vbTextCompare) = 0 Then
        NormalizeSasReferences = text
        Exit Function
    End If
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = True
        re.Pattern = "\bSAS\s*([:#-]\s*)?([0-9]{6,16})\b"
    End If
    NormalizeSasReferences = re.Replace(text, "SAS $2")
End Function

Private Function HasKnownOrder(ByVal text As String, ByVal orders As Object) As Boolean
    HasKnownOrder = (CountKnownOrders(text, orders, True) > 0)
End Function

Private Function CountKnownOrders(ByVal text As String, ByVal orders As Object, _
    Optional ByVal stopAfterFirst As Boolean = False) As Long
    Static tokens As Object, paired As Object, dates As Object
    Dim token As Object, matches As Object, found As Object, code As String, raw As String
    If Len(text) = 0 Or orders.Count = 0 Then Exit Function
    If tokens Is Nothing Then
        Set tokens = CreateObject("VBScript.RegExp")
        tokens.Global = True
        tokens.Pattern = "[A-Za-z0-9_]+(-[A-Za-z0-9_]+)*"
        Set paired = CreateObject("VBScript.RegExp")
        paired.Pattern = "^([0-9]{6,16})-[0-9]{1,5}$"
        Set dates = CreateObject("VBScript.RegExp")
        dates.Global = True
        dates.Pattern = "\b[0-9]{1,4}[./-][0-9]{1,2}[./-][0-9]{1,4}\b"
    End If
    If Not stopAfterFirst Then
        Set found = CreateObject("Scripting.Dictionary")
        found.CompareMode = vbTextCompare
    End If
    For Each token In tokens.Execute(dates.Replace(text, " "))
        raw = CStr(token.Value)
        code = NormalizeOrderNo(raw)
        Set matches = paired.Execute(raw)
        If matches.Count > 0 Then
            If orders.Exists(NormalizeOrderNo(CStr(matches(0).SubMatches(0)))) Then
                code = NormalizeOrderNo(CStr(matches(0).SubMatches(0)))
            End If
        End If
        If orders.Exists(code) Then
            If stopAfterFirst Then CountKnownOrders = 1: Exit Function
            found(code) = True
        End If
    Next token
    If Not stopAfterFirst Then CountKnownOrders = found.Count
End Function

Private Function StartsWithOrderNumber(ByVal text As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^\s*[0-9]{6,12}(?![0-9]|\s*(adet|pcs|kg|ton|litre|metre|tl|eur|usd)\b)(\s|/|-|,|:|$)"
    End If
    StartsWithOrderNumber = re.Test(text)
End Function

Private Function ContainsCode(ByVal text As String, ByVal code As String) As Boolean
    Static re As Object
    Dim normalized As String
    normalized = Trim$(Replace(Replace(code, ChrW(160), ""), ChrW(8203), ""))
    If Left$(normalized, 1) = "'" Then normalized = Mid$(normalized, 2)
    If Len(normalized) = 0 Then Exit Function
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
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
    re.Pattern = "(^|[^a-z])(kalem|item|pozisyon)(\s*(no|numarasi))?\s*[.:#-]*\s*\s*0*[0-9]+([^0-9]|$)|" & _
        "(^|[^0-9])[0-9]{1,5}\s+(nolu|no lu|numarali)\s+(kalem|item|pozisyon)([^a-z]|$)"
    HasItemLabel = re.Test(ItemSearchText(text))
End Function

Private Function HasOrderItemPair(ByVal text As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "\b[0-9]{6,12}\s*[/\-]\s*[0-9]{1,5}\b(?![./-][0-9])"
    End If
    HasOrderItemPair = re.Test(text)
End Function

Private Function ContainsItem(ByVal text As String, ByVal code As String, Optional ByVal orderNo As String = "") As Boolean
    Static re As Object
    Dim parts As Variant, normalized As String
    normalized = NormalizeOrderNo(code)
    If Len(normalized) = 0 Then Exit Function
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(^|[^a-z])(kalem|item|pozisyon)(\s*(no|numarasi))?\s*[.:#-]*\s*\s*0*" & EscapeRegex(normalized) & "(?![0-9]|[./-][0-9])([^0-9]|$)"
    If re.Test(ItemSearchText(text)) Then ContainsItem = True: Exit Function
    re.Pattern = "(^|[^0-9])0*" & EscapeRegex(normalized) & "\s+(nolu|no lu|numarali)\s+(kalem|item|pozisyon)([^a-z]|$)"
    If re.Test(ItemSearchText(text)) Then ContainsItem = True: Exit Function
    If Len(orderNo) > 0 Then
        re.Pattern = "(^|[^0-9])0*" & EscapeRegex(NormalizeOrderNo(orderNo)) & "\s*[/\-]\s*0*" & EscapeRegex(normalized) & "(?![0-9]|[./-][0-9])([^0-9]|$)"
        If re.Test(text) Then ContainsItem = True: Exit Function
    End If
    ' Preserve the standard exported order / zero-padded item table layout.
    If Len(orderNo) > 0 And InStr(text, vbTab) > 0 Then
        parts = Split(text, vbTab)
        Dim cellIndex As Long, cellText As String
        For cellIndex = LBound(parts) To UBound(parts) - 1
            If NormalizeOrderNo(Trim$(CStr(parts(cellIndex)))) = NormalizeOrderNo(orderNo) Then
                cellText = Trim$(CStr(parts(cellIndex + 1)))
                If Len(cellText) >= 5 And Left$(cellText, 1) = "0" Then
                    If cellText Like String$(Len(cellText), "#") Then
                        If NormalizeOrderNo(cellText) = normalized Then ContainsItem = True: Exit Function
                    End If
                End If
            End If
        Next cellIndex
    End If
    ' A bare numeric table cell can be quantity, price or a date component.
    ' Require a labelled item or order/item pair instead of guessing its meaning.
End Function

Private Function IsSignatureStart(ByVal text As String) As Boolean
    Dim t As String
    t = NormalizeName(text)
    IsSignatureStart = (t = "SAYGILAR" Or t = "SAYGILARIMLA" Or t = "SAYGILARIMIZLA" Or _
        t = "IYI CALISMALAR" Or t = "BEST REGARDS" Or t = "KIND REGARDS" Or t = "REGARDS" Or _
        t = "SAYGILARIMIZLA BEST REGARDS" Or t = "SAYGILARIMLA BEST REGARDS" Or t = "IYI GUNLER")
End Function

Private Function EnsureOrCreateColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If CleanForCompare(SafeStr(ws.Cells(HEADER_ROW, c).Value)) = CleanForCompare(headerText) Then
            EnsureOrCreateColumn = c: Exit Function
        End If
    Next c
    If lastCol >= ws.Columns.Count Then Err.Raise vbObjectError + 801, , "Yeni sonuc sutunu icin yer yok."
    c = lastCol + 1
    ws.Cells(HEADER_ROW, c).Value = headerText
    ws.Cells(HEADER_ROW, c).Font.Bold = True
    ws.Cells(HEADER_ROW, c).Interior.Color = RGB(200, 220, 240)
    EnsureOrCreateColumn = c
End Function





Private Sub WriteReplyToRow(ByVal ws As Worksheet, ByVal rowNum As Long, _
                             ByVal newReply As String, ByVal mailDate As Date, _
                             ByVal senderEmail As String, ByVal matchMethod As String, _
                             ByVal attachmentList As String, _
                             ByVal rBC As Long, ByVal rDC As Long, _
                             ByVal rFC As Long, ByVal rMC As Long, ByVal rAC As Long, _
                             ByVal dCol As Long, ByVal tCol As Long, _
                             ByVal terminDate As Variant, _
                             ByVal termState As Long, ByVal statusText As String, Optional ByVal noteOnly As Boolean = False)
    Dim displayReply As String
    If rBC < 1 Or rDC < 1 Then Exit Sub
    ArchivePreviousReply ws, rowNum, rBC, rDC, rFC, rAC, dCol, tCol, mailDate
    If noteOnly Then dCol = 0: tCol = 0
    displayReply = newReply
    If Len(displayReply) > 600 Then displayReply = Left$(displayReply, 597) & "..."
    With ws.Cells(rowNum, rBC)
        .NumberFormat = "@"
        .Value2 = displayReply
    End With
    If rFC > 0 Then ws.Cells(rowNum, rFC).Value = senderEmail
    If rMC > 0 Then ws.Cells(rowNum, rMC).Value = matchMethod
    If rAC > 0 Then ws.Cells(rowNum, rAC).Value = attachmentList

    If tCol > 0 Then
        If termState = TERM_CONFIRMED Then
            ws.Cells(rowNum, tCol).Value = terminDate
            ws.Cells(rowNum, tCol).NumberFormat = "dd.mm.yyyy"
        ElseIf termState = TERM_WITHDRAWN Then
            ws.Cells(rowNum, tCol).ClearContents
        End If
    End If
    If dCol > 0 Then
        If statusText <> STATUS_BELIRSIZ Or termState <> 0 Or Len(SafeStr(ws.Cells(rowNum, dCol).Value)) = 0 Then
            ws.Cells(rowNum, dCol).Value = statusText
        End If
    End If
    If ws.Parent Is ThisWorkbook And termState = TERM_UNCERTAIN Then
        WriteLog "TERMIN_BELIRSIZ", ws.Name, "", senderEmail, "Satir=" & rowNum, _
            "Birden fazla tarih, tahmin veya tarih araligi"
    End If
    ' Commit the reply timestamp after the result columns.
    ws.Cells(rowNum, rDC).Value = mailDate
End Sub

Private Sub MarkMailSent(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal statusCol As Long)
    Dim cell As Range, existing As String
    If statusCol < 1 Then Exit Sub
    Set cell = ws.Cells(rowNum, statusCol)
    If cell.HasFormula Or cell.MergeCells Or IsError(cell.Value2) Then Exit Sub
    existing = SafeStr(cell.Value2)
    If Left$(existing, 1) = "+" Then Exit Sub
    If Len(existing) = 0 Then
        cell.Value2 = "+"
    Else
        cell.Value2 = "+ " & existing
    End If
End Sub





Private Function ResolveDeliveryDate(ByVal reply As String, ByVal mailDate As Date, ByRef termState As Long, _
    Optional ByVal stockOnly As Boolean = False) As Variant
    Static operational As Object, revised As Object, uncertain As Object, dateToken As Object, onlyTiming As Object
    Static dimensions As Object, clauseBreak As Object
    Dim part As Variant, line As String, norm As String, value As Variant, chosen As Variant
    Dim clauses As String
    Dim tokens As Object, latestExplicit As Boolean, isExplicit As Boolean, sawUncertainty As Boolean
    Dim status As String, marker As Object, isStock As Boolean, stockAbsent As Boolean
    If operational Is Nothing Then
        Set operational = CreateObject("VBScript.RegExp")
        operational.IgnoreCase = True
        operational.Pattern = "\b(termin|teslim[a-z]*|sevk[a-z]*|varis|delivery|deliver[a-z]*)\b"
        Set revised = CreateObject("VBScript.RegExp")
        revised.IgnoreCase = True: revised.Global = True
        revised.Pattern = "\b((yeni|guncel|revize|teyitli|onayli)\s+(termin(i| tarihi)?|teslim(at)?( tarihi)?)|" & _
            "yerine|new delivery date|revised delivery date|confirmed delivery date)\b"
        Set uncertain = CreateObject("VBScript.RegExp")
        uncertain.IgnoreCase = True
        uncertain.Pattern = "\b(tahmin[a-z]*|yaklasik|muhtemel[a-z]*|olabilir|olursa|ihtimal[a-z]*|veya|yahut|" & _
            "arasinda|araliginda|bekleniyor|netlesince|planlaniyor|estimated|tentative|around|approximately|" & _
            "uygun\s+(mu|mudur)|mumkun\s+(mu|mudur)|teyit\s+bekliyoruz|teyidinizi|en\s+gec|en\s+erken|" & _
            "son\s+tarih|by\s+the\s+latest)\b|\?|\b[0-9]+\s*(is\s+)?(gun(u)?|hafta(si)?)\s+icinde\b|" & _
            "\b\d{1,2}\s*[-–]\s*\d{1,2}\s+(is\s+)?(gun|hafta|ocak|subat|mart|nisan|mayis|haziran|" & _
            "temmuz|agustos|eylul|ekim|kasim|aralik)\b"
        Set dateToken = CreateObject("VBScript.RegExp")
        dateToken.Global = True: dateToken.IgnoreCase = True
        dateToken.Pattern = "\b([0-9]{4}-[0-9]{1,2}-[0-9]{1,2}|[0-9]{1,2}[./-][0-9]{1,2}([./-][0-9]{2,4})?|" & _
            "[0-9]{1,2}\s+(ocak|subat|mart|nisan|mayis|haziran|temmuz|agustos|eylul|ekim|kasim|aralik)(\s+[0-9]{4})?|" & _
            "bugun|yarin|obur gun|pazartesi|sali|carsamba|persembe|cuma|cumartesi|pazar|[0-9]{1,4}\s*(is\s*)?gun(u)?)\b"
        Set dimensions = CreateObject("VBScript.RegExp")
        dimensions.Global = True: dimensions.IgnoreCase = True
        dimensions.Pattern = "\b[0-9]{1,2}[./-][0-9]{1,2}\s*(mm|cm|kg|inc|inch|adet|pcs)\b"
        Set clauseBreak = CreateObject("VBScript.RegExp")
        clauseBreak.Global = True: clauseBreak.IgnoreCase = True
        clauseBreak.Pattern = ",\s*(?=(stok|teslim|sevk|termin|ancak|ama|fakat)\b)|" & _
            "\s+(ve|ancak|ama|fakat)\s+(?=(teslim tarihi|sevk tarihi|stok tarihi|termin)\b)"
        Set onlyTiming = CreateObject("VBScript.RegExp")
        onlyTiming.IgnoreCase = True
        onlyTiming.Pattern = "^(\d{1,4}[./-]\d{1,2}([./-]\d{1,4})?[.]?|\d{1,2}\s+(ocak|subat|mart|nisan|mayis|" & _
            "haziran|temmuz|agustos|eylul|ekim|kasim|aralik)(\s+\d{4})?|bugun|yarin|obur gun|" & _
            "((haftaya|bu|gelecek hafta)\s+)?(pazartesi|sali|carsamba|persembe|cuma|cumartesi|pazar)|" & _
            "\d{1,4}\s*(is\s*)?gun(u)?(\s+(sonra|icinde))?)$"
    End If
    termState = 0
    clauses = Replace(Replace(Replace(reply, vbCrLf, " / "), " / ", vbLf), ";", vbLf)
    clauses = clauseBreak.Replace(FoldTurkish(clauses), vbLf)
    For Each part In Split(clauses, vbLf)
        line = Trim$(WithoutLinks(CStr(part)))
        norm = dimensions.Replace(FoldTurkish(line), "")
        isStock = IsStockAvailability(norm)
        If stockOnly And Not isStock Then GoTo NextDatePart
        If isStock And Not stockOnly Then GoTo NextDatePart
        Set marker = revised.Execute(norm)
        isExplicit = (marker.Count > 0)
        If isExplicit Then
            norm = Trim$(Mid$(norm, marker(marker.Count - 1).FirstIndex + marker(marker.Count - 1).Length + 1))
        End If
        status = ClassifyStatus(norm, False)
        If status = STATUS_STOK_YOK And Not operational.Test(norm) Then
            stockAbsent = True
            GoTo NextDatePart
        End If
        If status = STATUS_IPTAL Or ExplicitlyWithdrawsDate(norm) Or _
            (isExplicit And ExplicitlyWithdrawsDate("termin " & norm)) Then
            If Not latestExplicit Or isExplicit Then
                chosen = Empty
                termState = TERM_WITHDRAWN
                sawUncertainty = False
                If isExplicit Then latestExplicit = True
            End If
            GoTo NextDatePart
        End If
        If status = "Teslim Edildi" Or status = STATUS_SEVK Then GoTo NextDatePart
        If Not stockOnly And Not isExplicit And Not operational.Test(norm) And Not onlyTiming.Test(norm) Then GoTo NextDatePart
        If latestExplicit And Not isExplicit Then GoTo NextDatePart
        If isExplicit Then
            If Not latestExplicit Then
                chosen = Empty
                termState = 0
                sawUncertainty = False
            End If
            latestExplicit = True
        End If
        Set tokens = dateToken.Execute(norm)
        If uncertain.Test(norm) Or tokens.Count > 1 Then
            sawUncertainty = True
            GoTo NextDatePart
        End If
        If tokens.Count = 0 Then
            If isExplicit Then sawUncertainty = True
            GoTo NextDatePart
        End If
        value = ParseTurkishDate(norm, mailDate)
        If IsEmpty(value) Then
            sawUncertainty = True
            GoTo NextDatePart
        End If
        If IsEmpty(chosen) Then
            chosen = value
        ElseIf CDate(chosen) <> CDate(value) Then
            sawUncertainty = True
        End If
NextDatePart:
    Next part
    If sawUncertainty Then
        termState = TERM_UNCERTAIN
    ElseIf Not IsEmpty(chosen) Then
        termState = TERM_CONFIRMED
        ResolveDeliveryDate = chosen
    ElseIf stockAbsent And termState = 0 And Not stockOnly Then
        termState = TERM_WITHDRAWN
    End If
End Function

Private Function IsStockAvailability(ByVal text As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Pattern = "\b(stokta|stokumuzda|stoklarimizda|stoklarinda|stoklarda)\s+" & _
            "([0-9]{1,4}[./-][0-9]{1,2}([./-][0-9]{2,4})?([ '](da|de|ta|te)|\s+tarihinde)?\s+)?" & _
            "(olacak(tir|lar)?|bulunacak(tir|lar)?|yer alacak(tir)?)\b|" & _
            "\b(stoga|stogumuza|stoklarimiza)\s+(girecek(tir|ler)?|gelecek(tir|ler)?|alinacak(tir)?)\b|" & _
            "(^|\s/\s|[\r\n])\s*(stok tarihi|stok giris tarihi|stoga giris tarihi)\s*:\s*(\d|bugun\b|yarin\b)"
    End If
    IsStockAvailability = re.Test(FoldTurkish(text))
End Function

Private Function FoldTurkish(ByVal text As String) As String
    text = Replace(text, ChrW(304), "I")
    text = LCase$(text)
    text = Replace(text, ChrW(305), "i")
    text = Replace(text, ChrW(351), "s")
    text = Replace(text, ChrW(287), "g")
    text = Replace(text, ChrW(252), "u")
    text = Replace(text, ChrW(246), "o")
    text = Replace(text, ChrW(231), "c")
    FoldTurkish = text
End Function

Private Function ExplicitlyWithdrawsDate(ByVal text As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Pattern = "\b(termin(i|imiz| tarihi)?|teslim(at)?( tarihi)?|sevk(i| tarihi)?)\b.{0,65}" & _
            "\b(belli degil|net degil|netlesmedi|veremiyoruz|tarih yok|teyit edilmedi|iptal edildi|" & _
            "edileme[a-z]*|edilmeyecek|yapilamay[a-z]*|gerceklesmeyecek|mumkun degil|gecersiz|belirsiz)\b"
    End If
    ExplicitlyWithdrawsDate = re.Test(FoldTurkish(text))
End Function

Private Function ParseTurkishDate(ByVal text As String, ByVal refDate As Date) As Variant
    refDate = DateValue(refDate)
    Dim t As String: t = FoldTurkish(text)
    On Error Resume Next
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp"): re.Global = False
    
    ' 1) yyyy-mm-dd veya yyyy.mm.dd formatý
    re.Pattern = "\b(\d{4})[./-](\d{1,2})[./-](\d{1,2})\b"
    Set m = re.Execute(t)
    If m.Count > 0 Then
        Dim y1 As Long, m1 As Long, d1 As Long
        y1 = CLng(m(0).SubMatches(0)): m1 = CLng(m(0).SubMatches(1)): d1 = CLng(m(0).SubMatches(2))
        If d1 >= 1 And d1 <= 31 And m1 >= 1 And m1 <= 12 Then
            Err.Clear
            Dim dt1 As Date: dt1 = DateSerial(y1, m1, d1)
            If Err.Number = 0 And Year(dt1) = y1 And Month(dt1) = m1 And Day(dt1) = d1 Then
                ParseTurkishDate = dt1: Exit Function
            End If
        End If
        Exit Function
    End If
    
    ' 2) dd.mm.yyyy veya dd.mm formatý
    re.Pattern = "\b(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?\b"
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
            If Err.Number = 0 And Month(dt) = mm And Day(dt) = dd Then
                If DateDiff("d", dt, refDate) > 180 And Len(m(0).SubMatches(2)) = 0 Then
                    dt = DateSerial(yy + 1, mm, dd)
                End If
                ParseTurkishDate = dt: Exit Function
            End If
        End If
        Exit Function
    End If
    re.Pattern = "\byarin\b"
    If re.Test(t) Then ParseTurkishDate = refDate + 1: Exit Function
    re.Pattern = "\bbugun\b"
    If re.Test(t) Then ParseTurkishDate = refDate: Exit Function
    re.Pattern = "\bobur gun\b"
    If re.Test(t) Then ParseTurkishDate = refDate + 2: Exit Function
    Dim dayNames As Variant, i As Long
    dayNames = Array("pazar", "pazartesi", "sali", "carsamba", "persembe", "cuma", "cumartesi")
    For i = 0 To 6
        re.Pattern = "(^|[^a-z])" & dayNames(i) & "([^a-z]|$)"
        If re.Test(t) Then
            Dim targetDay As Long: targetDay = i + 1
            Dim diff As Long
            diff = (targetDay - Weekday(refDate) + 7) Mod 7
            If diff = 0 Then diff = 7
            If InStr(t, "haftaya") > 0 Or InStr(t, "gelecek hafta") > 0 Then _
                diff = 8 - Weekday(refDate, vbMonday) + ((targetDay + 5) Mod 7)
            ParseTurkishDate = refDate + diff: Exit Function
        End If
    Next i
    Set re = CreateObject("VBScript.RegExp"): re.Global = False
    re.Pattern = "\b([0-9]{1,4})\s*(is\s+)?gun(u)?\b"
    Set m = re.Execute(t)
    If m.Count > 0 Then
        Dim nDays As Long: nDays = CLng(m(0).SubMatches(0))
        If nDays > 3650 Then Exit Function
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
    If InStr(t, "gelecek hafta") > 0 Or InStr(t, "haftaya") > 0 Then Exit Function
    Dim months As Variant, mi As Long
    months = Array("ocak", "subat", "mart", "nisan", "mayis", "haziran", _
                   "temmuz", "agustos", "eylul", "ekim", "kasim", "aralik")
    For mi = 0 To 11
        If InStr(t, months(mi)) > 0 Then
            re.Pattern = "(\d{1,2})\s*" & months(mi) & "(?:\s+(\d{4}))?"
            Set m = re.Execute(t)
            If m.Count > 0 Then
                dd = CLng(m(0).SubMatches(0))
                If dd >= 1 And dd <= 31 Then
                    Err.Clear
                    yy = Year(refDate)
                    If Len(m(0).SubMatches(1)) > 0 Then yy = CLng(m(0).SubMatches(1))
                    dt = DateSerial(yy, mi + 1, dd)
                    If Err.Number = 0 And Year(dt) = yy And Month(dt) = mi + 1 And Day(dt) = dd Then
                        If DateDiff("d", dt, refDate) > 180 And Len(m(0).SubMatches(1)) = 0 Then dt = DateSerial(yy + 1, mi + 1, dd)
                        ParseTurkishDate = dt: Exit Function
                    End If
                End If
            End If
            If m.Count > 0 Then Exit Function
        End If
    Next mi
    On Error GoTo 0
End Function

Private Function ClassifyStatus(ByVal replyText As String, ByVal hasDate As Boolean) As String
    Dim text As String, re As Object
    text = FoldTurkish(NormalizeName(replyText))
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True
    re.Pattern = "iptal (edilmedi|edilmeyecek|degil|yok)|iptal edilmeyecektir"
    text = re.Replace(text, "")
    re.Pattern = "\b(iptal (edildi|edilmistir|edilecek|ettik|oldu)|vazgecildi|siparis reddedildi|siparisi reddettik)\b|^iptal$"
    If re.Test(text) Then ClassifyStatus = STATUS_IPTAL: Exit Function
    re.Pattern = "(stok(ta|umuzda|larimizda)? yok|stok mevcut degil|tukendi|stok bulunmuyor)"
    If Not hasDate And re.Test(text) Then ClassifyStatus = STATUS_STOK_YOK: Exit Function
    re.Pattern = "(teslim edildi|teslim edilmistir|teslim ettik|teslim alinmistir)"
    If re.Test(text) Then ClassifyStatus = "Teslim Edildi": Exit Function
    re.Pattern = "\b(sevk edildi|sevk edilmistir|sevk ?ettik|kargoya verildi|yola cikti|yolda|" & _
        "(malzeme|urun|siparis|paket) gonderildi)\b"
    If re.Test(text) Then ClassifyStatus = STATUS_SEVK: Exit Function
    re.Pattern = "(uretimde|imalatta|uretime alindi|uretime girdi|imal ediliyor)"
    If re.Test(text) Then ClassifyStatus = STATUS_URETIM: Exit Function
    If hasDate Then ClassifyStatus = STATUS_TESLIM: Exit Function
    re.Pattern = "(termin|tarih|teslim).*(belli degil|belirsiz|veremiyoruz|netlesmedi|teyit edilmedi)"
    If re.Test(text) Then ClassifyStatus = STATUS_BELIRSIZ: Exit Function
    If ExplicitlyWithdrawsDate(text) Then ClassifyStatus = STATUS_BELIRSIZ: Exit Function
    re.Pattern = "\b(teslim edilecek(tir)?|teslim edecegiz|sevk edilecek(tir)?|sevk edecegiz)\b"
    If re.Test(text) Then ClassifyStatus = STATUS_TESLIM: Exit Function
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
                               ByVal replied As Long, ByVal updated As Long, Optional ByVal updatedMails As Long = 0)
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
    gTermReportWs.Cells(r, 1).Value = "Secili tablodaki tedarikcilerden okunan mail": gTermReportWs.Cells(r, 2).Value = replied: r = r + 1
    gTermReportWs.Cells(r, 1).Value = "Satir guncelleyen yanit": gTermReportWs.Cells(r, 2).Value = updatedMails: r = r + 1
    gTermReportWs.Cells(r, 1).Value = "Satir guncellemeyen yanit": gTermReportWs.Cells(r, 2).Value = replied - updatedMails: r = r + 1
    gTermReportWs.Cells(r, 1).Value = "Secili tablonun disinda kalan mail": gTermReportWs.Cells(r, 2).Value = gSiparisYokCount: r = r + 1
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
    Dim att As Object, fs As Object, root As String, vendorFolder As String, mailFolder As String
    Dim result As String, target As String, stamp As String, identity As String
    Dim position As Long, hidden As Boolean, contentId As String, extension As String, failure As String
    On Error GoTo Failed
    If mail.Attachments.Count = 0 Then Exit Function
    Set fs = CreateObject("Scripting.FileSystemObject")
    root = ThisWorkbook.Path
    If Len(root) = 0 Or InStr(root, "://") > 0 Then root = Environ$("USERPROFILE") & "\Documents"
    root = root & "\" & ATTACH_BASE_FOLDER
    If Not fs.FolderExists(root) Then fs.CreateFolder root
    vendorFolder = root & "\" & Left$(SanitizeFileName(vendorName), 65)
    If Not fs.FolderExists(vendorFolder) Then fs.CreateFolder vendorFolder
    identity = SafeStr(mail.EntryID)
    stamp = Format$(mail.ReceivedTime, "yyyymmdd_hhnnss") & "_" & Right$(SanitizeFileName(identity), 32)
    mailFolder = vendorFolder & "\" & stamp
    If Not fs.FolderExists(mailFolder) Then fs.CreateFolder mailFolder
    For Each att In mail.Attachments
        position = position + 1: hidden = False: contentId = ""
        On Error Resume Next
        hidden = CBool(att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptags/0x7FFE000B"))
        contentId = SafeStr(att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptags/0x3712001E"))
        On Error GoTo Failed
        extension = LCase$(fs.GetExtensionName(SafeStr(att.FileName)))
        If hidden Then GoTo NextAttachment
        If Len(contentId) > 0 Then
            If extension = "png" Or extension = "jpg" Or extension = "jpeg" Or extension = "gif" Then GoTo NextAttachment
        End If
        target = mailFolder & "\" & Format$(position, "000") & "_" & Left$(SanitizeFileName(SafeStr(att.FileName)), 90)
        On Error Resume Next
        Err.Clear
        If Not fs.FileExists(target) Then att.SaveAsFile target
        failure = Err.Description
        On Error GoTo Failed
        If Len(failure) = 0 Then
            If Len(result) > 0 Then result = result & "; "
            result = result & target
        Else
            WriteLog "EK_HATA", "", vendorName, "", SafeStr(att.FileName), failure
        End If
NextAttachment:
    Next att
    SaveAttachments = Left$(result, MAX_BODY_LEN)
    Exit Function
Failed:
    failure = Err.Description
    WriteLog "EK_KLASOR_HATA", "", vendorName, "", failure
    SaveAttachments = Left$(result, MAX_BODY_LEN)
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
        If Left$(line, 1) = ">" Then Exit For
        If IsQuotedBoundary(line) Then Exit For
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
        t = "GONDERILMIS" Or t = "GONDERILDI" Or t = "GONDERME TARIHI" Or t = "SENT ON" Or _
        t = "DE" Or t = "VON" Or t = "EXPEDITEUR")
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
    s = " " & LCase$(NormalizeName(s)) & " "
    arr = Array("ocak", "subat", "mart", "nisan", "mayis", "haziran", "temmuz", _
                "agustos", "eylul", "ekim", "kasim", "aralik", _
                "hafta", "is gunu", "gun icinde")
    For i = LBound(arr) To UBound(arr)
        If InStr(1, s, " " & CStr(arr(i)) & " ", vbTextCompare) > 0 Then HasMonthOrWeek = True: Exit Function
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
    Dim logName As String
    logName = CStr(IIf(gPreview, PREVIEW_LOG_SHEET, LOG_SHEET))
    On Error Resume Next
    If gLogWs Is Nothing Then
        Set gLogWs = ThisWorkbook.Worksheets(logName)
        If gLogWs Is Nothing Then
            Set gLogWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            gLogWs.Name = logName
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
    gLogWs.Range(gLogWs.Cells(gLogRow, 2), gLogWs.Cells(gLogRow, 7)).NumberFormat = "@"
    gLogWs.Cells(gLogRow, 2).Value = islem
    gLogWs.Cells(gLogRow, 3).Value = sayfa
    gLogWs.Cells(gLogRow, 4).Value = tedarikci
    gLogWs.Cells(gLogRow, 5).Value = mail
    gLogWs.Cells(gLogRow, 6).Value = detay1
    gLogWs.Cells(gLogRow, 7).Value = detay2
    gLogRow = gLogRow + 1
End Sub
Private Function ReadMailFields(ByVal mail As Object, ByRef received As Date, _
    ByRef senderEmail As String, ByRef senderName As String, ByRef subject As String) As Boolean
    On Error GoTo Failed
    received = 0: senderEmail = "": senderName = "": subject = ""
    If mail.Class <> 43 Then Exit Function
    received = mail.ReceivedTime
    senderEmail = NormalizeEmail(GetSenderSmtpEx(mail))
    senderName = SafeStr(mail.SenderName)
    subject = SafeStr(mail.Subject)
    ReadMailFields = True
    Exit Function
Failed:
    ReadMailFields = False
End Function

Private Function ReadMessageBody(ByVal mail As Object, ByRef body As String) As Boolean
    body = ""
    On Error Resume Next
    body = SafeStr(mail.HTMLBody)
    If Len(body) = 0 Then
        Err.Clear
        body = SafeStr(mail.Body)
    End If
    ReadMessageBody = (Err.Number = 0 Or Len(body) > 0)
    On Error GoTo 0
End Function

Private Function ReadPlainMessageBody(ByVal mail As Object, ByRef body As String) As Boolean
    On Error Resume Next
    body = SafeStr(mail.Body)
    ReadPlainMessageBody = Len(body) > 0
    On Error GoTo 0
End Function

Private Function NormalizeEmail(ByVal value As String) As String
    Dim leftPos As Long, rightPos As Long, re As Object
    value = LCase$(Trim$(Replace(value, ChrW(160), " ")))
    leftPos = InStr(value, "<"): rightPos = InStr(value, ">")
    If leftPos > 0 And rightPos > leftPos Then value = Mid$(value, leftPos + 1, rightPos - leftPos - 1)
    value = Trim$(Replace(value, "mailto:", ""))
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$"
    If re.Test(value) Then NormalizeEmail = value
End Function

Private Function VendorIsBlocked(ByVal name As String) As Boolean
    Dim normalized As String
    normalized = NormalizeName(name)
    VendorIsBlocked = (InStr(normalized, "BLOKE") > 0 Or InStr(normalized, "BLOCKED") > 0)
End Function

Private Function VendorTitleWithoutLegalForm(ByVal name As String) As String
    Static suffix As Object
    Dim normalized As String, shortened As String
    normalized = FoldTurkish(NormalizeName(name))
    If suffix Is Nothing Then
        Set suffix = CreateObject("VBScript.RegExp")
        suffix.Pattern = "\s+(a\s*s|anonim\s+sirketi|l\s*t\s*d(\s+(sti|sirketi))?|limit(ed|et)(\s+(sti|sirketi))?)$"
    End If
    Do
        shortened = Trim$(suffix.Replace(normalized, ""))
        If shortened = normalized Then Exit Do
        normalized = shortened
    Loop
    VendorTitleWithoutLegalForm = normalized
End Function

Private Function VendorLegalForm(ByVal name As String) As String
    Static anonymousCompany As Object, limitedCompany As Object
    Dim norm As String
    If anonymousCompany Is Nothing Then
        Set anonymousCompany = CreateObject("VBScript.RegExp")
        anonymousCompany.Pattern = "\s+(a\s*s|anonim\s+sirketi)$"
        Set limitedCompany = CreateObject("VBScript.RegExp")
        limitedCompany.Pattern = "\s+(l\s*t\s*d(\s+(sti|sirketi))?|limit(ed|et)(\s+(sti|sirketi))?)$"
    End If
    norm = FoldTurkish(NormalizeName(name))
    If anonymousCompany.Test(norm) Then VendorLegalForm = "AS"
    If limitedCompany.Test(norm) Then VendorLegalForm = "LTD"
End Function

Private Function BuildVendorAliases(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal nameCol As Long, _
    ByVal activeVendors As Object, ByVal vendorStates As Object, ByVal directoryTitles As Object) As Object
    Dim row As Long, rawName As String, key As String
    For row = 2 To lastRow
        rawName = SafeStr(ws.Cells(row, nameCol).Value)
        key = NormalizeName(rawName)
        If Len(key) > 0 Then
            If Not vendorStates.Exists(key) Then
                vendorStates.Add key, 0
                directoryTitles.Add key, rawName
            End If
            If Not VendorIsBlocked(rawName) Then
                If CLng(vendorStates(key)) = 0 Then directoryTitles(key) = rawName
                vendorStates(key) = 1
            End If
        End If
    Next row
    Set BuildVendorAliases = ResolveVendorAliases(activeVendors, vendorStates, gVendorDirectoryNotes)
End Function

Private Function ResolveVendorAliases(ByVal active As Object, ByVal directory As Object, ByVal notes As Object) As Object
    Dim resolved As Object
    Set resolved = CreateObject("Scripting.Dictionary")
    AddVendorTitleAliases active, directory, notes, resolved
    AddTruncatedVendorAliases active, directory, notes, resolved
    Set ResolveVendorAliases = resolved
End Function

Private Sub AddVendorTitleAliases(ByVal active As Object, ByVal directory As Object, _
    ByVal notes As Object, ByVal resolved As Object)
    Dim groups As Object, proposed As Object, claims As Object
    Dim bucket As Collection, entry As Variant, item As Variant
    Dim baseName As String, suffix As String, form As String
    Dim orderName As String, masterName As String, hits As Long
    Dim exactSuffixName As String, exactSuffixHits As Long
    Set groups = CreateObject("Scripting.Dictionary")
    Set proposed = CreateObject("Scripting.Dictionary")
    Set claims = CreateObject("Scripting.Dictionary")

    For Each item In directory.Keys
        SplitVendorTitle CStr(item), baseName, suffix, form
        If Len(baseName) >= 4 Then
            If Not groups.Exists(baseName) Then
                Set bucket = New Collection
                groups.Add baseName, bucket
            End If
            Set bucket = groups(baseName)
            bucket.Add Array(CStr(item), suffix, form)
        End If
    Next item

    For Each item In active.Keys
        orderName = CStr(item)
        If directory.Exists(orderName) Then GoTo NextOrderTitle
        SplitVendorTitle orderName, baseName, suffix, form
        If Not groups.Exists(baseName) Then GoTo NextOrderTitle
        Set bucket = groups(baseName)
        hits = 0: exactSuffixHits = 0
        For Each entry In bucket
            If Len(form) > 0 And Len(CStr(entry(2))) > 0 And form <> CStr(entry(2)) Then GoTo NextTitleCandidate
            If bucket.Count > 1 Then
                If Not VendorSuffixMatches(suffix, CStr(entry(1))) Then GoTo NextTitleCandidate
            End If
            hits = hits + 1
            masterName = CStr(entry(0))
            If Len(suffix) > 0 And suffix = CStr(entry(1)) Then
                Select Case Mid$(suffix, InStrRev(suffix, " ") + 1)
                    Case "T", "TI", "A"
                        ' Kesik ek tek basina tam unvan sayilmaz.
                    Case Else
                        exactSuffixHits = exactSuffixHits + 1
                        exactSuffixName = masterName
                End Select
            End If
NextTitleCandidate:
        Next entry
        If exactSuffixHits = 1 Then
            hits = 1
            masterName = exactSuffixName
        End If
        Select Case hits
            Case 0
                notes(orderName) = T("Firma ad~i bulundu fakat unvan ekleri veya ~sirket t~ur~u e~sle~smedi.")
            Case 1
                proposed.Add orderName, masterName
                AddVendorIdentity claims, masterName, orderName
            Case Else
                notes(orderName) = T("Ayn~i firma ad~ina sahip birden fazla kay~it unvan ekleriyle e~sle~siyor.")
        End Select
NextOrderTitle:
    Next item

    For Each item In proposed.Keys
        masterName = CStr(proposed(item))
        If active.Exists(masterName) Or CStr(claims(masterName)) = "***CAKISMA***" Then
            notes(CStr(item)) = T("Tedarik~ci kayd~i birden fazla sipari~s ad~iyla e~sle~siyor.")
        Else
            resolved.Add masterName, CStr(item)
        End If
    Next item
End Sub

Private Sub SplitVendorTitle(ByVal title As String, ByRef baseName As String, _
    ByRef suffix As String, ByRef form As String)
    Dim words() As String, i As Long, code As String
    Dim readingSuffix As Boolean
    baseName = VendorTitleWithoutLegalForm(title)
    form = VendorLegalForm(title)
    suffix = ""
    If Len(baseName) = 0 Then Exit Sub
    words = Split(baseName, " ")
    baseName = words(0)
    For i = 1 To UBound(words)
        code = VendorActivityCode(words(i))
        If i = UBound(words) And Len(form) = 0 Then
            Select Case words(i)
                Case "T", "TI", "A": code = words(i)
            End Select
        End If
        If Len(code) > 0 Then readingSuffix = True
        If readingSuffix Then
            If Len(code) = 0 Then code = words(i)
            If Len(suffix) > 0 Then suffix = suffix & " "
            suffix = suffix & code
        Else
            baseName = baseName & " " & words(i)
        End If
    Next i
    If Len(form) > 0 Then
        If Len(suffix) > 0 Then suffix = suffix & " "
        suffix = suffix & form
    ElseIf suffix = "A" Or Right$(suffix, 2) = " A" Then
        form = "AS"
    End If
End Sub

Private Function VendorActivityCode(ByVal word As String) As String
    Select Case word
        Case "ELK", "ELEK", "ELEKTRIK": VendorActivityCode = "ELEK"
        Case "ELKT", "ELKTR", "ELEKTRONIK": VendorActivityCode = "ELEKTRONIK"
        Case "MAK", "MAKINA", "MAKINE": VendorActivityCode = "MAK"
        Case "INS", "INSAAT": VendorActivityCode = "INS"
        Case "TURZ", "TURIZM": VendorActivityCode = "TURZ"
        Case "SAN", "SANAYI", "SANAYII": VendorActivityCode = "SAN"
        Case "TIC", "TICARET", "TICARETI": VendorActivityCode = "TIC"
        Case "TEKS", "TEKSTIL": VendorActivityCode = "TEKSTIL"
        Case "TAS", "TASIMACILIK": VendorActivityCode = "TASIMACILIK"
        Case "NAK", "NAKL", "NAKLIYE", "NAKLIYAT": VendorActivityCode = "NAKLIYAT"
        Case "GID", "GIDA": VendorActivityCode = "GIDA"
        Case "PAZ", "PAZARLAMA": VendorActivityCode = "PAZARLAMA"
        Case "MALZ", "MALZEME", "MALZEMELERI": VendorActivityCode = "MALZEME"
        Case "HIZ", "HIZMET", "HIZMETLERI": VendorActivityCode = "HIZMET"
        Case "PLAS", "PLASTIK": VendorActivityCode = "PLASTIK"
        Case "KIM", "KIMYA": VendorActivityCode = "KIMYA"
    End Select
End Function

Private Function VendorSuffixMatches(ByVal orderSuffix As String, ByVal directorySuffix As String) As Boolean
    Dim requested() As String, available() As String, i As Long
    Dim expected As String, actual As String
    If Len(orderSuffix) = 0 Then VendorSuffixMatches = True: Exit Function
    If Len(directorySuffix) = 0 Then Exit Function
    requested = Split(orderSuffix, " ")
    available = Split(directorySuffix, " ")
    If UBound(available) < UBound(requested) Then Exit Function
    For i = 0 To UBound(requested)
        expected = requested(i)
        actual = available(i)
        If expected <> actual Then
            If i < UBound(requested) Then Exit Function
            Select Case expected
                Case "T", "TI"
                    If Left$(actual, Len(expected)) <> expected Then Exit Function
                Case "A"
                    If actual <> "AS" Then Exit Function
                Case "AS", "LTD"
                    Exit Function
                Case Else
                    If Len(expected) < 2 Or VendorActivityCode(expected) <> "" Then Exit Function
                    If expected Like "*[!A-Z]*" Then Exit Function
                    If Left$(actual, Len(expected)) <> expected Then Exit Function
            End Select
        End If
    Next i
    VendorSuffixMatches = True
End Function

Private Sub AddTruncatedVendorAliases(ByVal active As Object, ByVal directory As Object, _
    ByVal notes As Object, ByVal resolved As Object)
    Const ambiguousName As String = "***CAKISMA***"
    Dim matched As Object, candidates As Object, claims As Object
    Dim orderName As Variant, directoryName As Variant, masterName As String
    Dim orderKey As String, orderForm As String, directoryForm As String
    Dim prefixLength As Long, nextCharacter As String
    Set matched = CreateObject("Scripting.Dictionary")
    Set candidates = CreateObject("Scripting.Dictionary")
    Set claims = CreateObject("Scripting.Dictionary")
    For Each directoryName In resolved.Keys
        matched(CStr(resolved(directoryName))) = True
    Next directoryName
    For Each orderName In active.Keys
        If directory.Exists(orderName) Or matched.Exists(orderName) Or notes.Exists(orderName) Then GoTo NextOrderVendor
        orderKey = CStr(orderName)
        prefixLength = Len(orderKey)
        ' Kisa/genel adlar yerine ilk kelimesi tam olan ve en az ikinci kelimeye uzanan adlari tara.
        If prefixLength < 8 Or InStr(orderKey, " ") = 0 Then GoTo NextOrderVendor
        orderForm = VendorLegalForm(orderKey)
        For Each directoryName In directory.Keys
            If Len(CStr(directoryName)) <= prefixLength Then GoTo NextDirectoryVendor
            If Left$(CStr(directoryName), prefixLength) <> orderKey Then GoTo NextDirectoryVendor
            nextCharacter = Mid$(CStr(directoryName), prefixLength + 1, 1)
            ' Rakamla biten kodlarin parcali eslesmesini kabul etme.
            If nextCharacter <> " " Then
                If Not (nextCharacter Like "[A-Z]") Or Not (Right$(orderKey, 1) Like "[A-Z]") Then GoTo NextDirectoryVendor
            End If
            directoryForm = VendorLegalForm(CStr(directoryName))
            If Len(orderForm) > 0 And Len(directoryForm) > 0 And orderForm <> directoryForm Then GoTo NextDirectoryVendor
            AddVendorIdentity candidates, orderKey, CStr(directoryName)
NextDirectoryVendor:
        Next directoryName
        If candidates.Exists(orderKey) Then
            masterName = CStr(candidates(orderKey))
            If masterName <> ambiguousName Then AddVendorIdentity claims, masterName, orderKey
        End If
NextOrderVendor:
    Next orderName
    For Each orderName In candidates.Keys
        masterName = CStr(candidates(orderName))
        If masterName = ambiguousName Then
            notes(orderName) = T("Tedarik~ci ad~i birden fazla kay~itla e~sle~siyor. Tedarik~ci listesini kontrol edin.")
        ElseIf active.Exists(masterName) Or resolved.Exists(masterName) Or CStr(claims(masterName)) = ambiguousName Then
            notes(orderName) = T("Tedarik~ci kayd~i birden fazla sipari~s ad~iyla e~sle~siyor.")
        Else
            resolved(masterName) = CStr(orderName)
        End If
    Next orderName
End Sub

Private Sub ReportVendorDirectory(ByVal sheetName As String, ByVal active As Object, ByVal states As Object, _
    ByVal aliases As Object, ByVal titles As Object, ByVal reportExact As Boolean)
    Dim byActive As Object, item As Variant, orderName As String, masterName As String
    Dim masterTitle As String, method As String, note As String
    Dim orderBase As String, masterBase As String, suffix As String, form As String
    Set byActive = CreateObject("Scripting.Dictionary")
    For Each item In aliases.Keys
        byActive(CStr(aliases(item))) = CStr(item)
    Next item

    For Each item In active.Keys
        orderName = CStr(item)
        If gVendorDirectoryNotes.Exists(orderName) Then
            WriteLog "TEDARIKCI_ADRES_UYARISI", sheetName, orderName, "", CStr(gVendorDirectoryNotes(orderName))
            GoTo NextVendorReport
        End If
        masterName = orderName
        method = "TAM_AD"
        If byActive.Exists(orderName) Then
            masterName = CStr(byActive(orderName))
            SplitVendorTitle orderName, orderBase, suffix, form
            SplitVendorTitle masterName, masterBase, suffix, form
            If orderBase = masterBase Then
                method = "FIRMA_ADI_VE_UNVAN_EKLERI"
            Else
                method = "KESIK_AD"
            End If
        End If

        masterTitle = ""
        If states.Exists(masterName) Then
            masterTitle = CStr(titles(masterName))
            note = VendorDirectoryIssue(CLng(states(masterName)))
        Else
            note = VendorDirectoryIssue(-1)
        End If

        If Len(note) > 0 Then
            gVendorDirectoryNotes(orderName) = note
            WriteLog "TEDARIKCI_ADRES_UYARISI", sheetName, orderName, "", note, masterTitle
        ElseIf method <> "TAM_AD" Or reportExact Then
            WriteLog "TEDARIKCI_AD_ESLESMESI", sheetName, orderName, "", masterTitle, method & " | Gecerli e-posta var"
        End If
NextVendorReport:
    Next item
End Sub

Private Function VendorDirectoryIssue(ByVal state As Long) As String
    Select Case state
        Case 0: VendorDirectoryIssue = T("Tedarik~ci kayd~i blokeli.")
        Case 1: VendorDirectoryIssue = T("Tedarik~ci e-posta h~ucresi bo~s.")
        Case 2: VendorDirectoryIssue = T("Tedarik~ci e-posta h~ucresinde ge~cerli adres yok.")
        Case -1: VendorDirectoryIssue = T("Tedarik~ci ad~i listede e~sle~smedi. Sipari~s ve tedarik~ci tablosunu kontrol edin.")
    End Select
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
    Dim passed As Long, errText As String, n As Long, reason As String
    Dim originalVendorFlag As Boolean: originalVendorFlag = gMailHasVendorRows
    Dim originalPreview As Boolean: originalPreview = gPreview
    Dim context As String, objects() As Object, times() As Date, j As Long
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

    Dim sampleStockReply As String, stockQuote As String, sampleStockDate As Variant, sampleStockState As Long
    sampleStockReply = T("~Ur~unler 01.12.2026 tarihinde stoklar~im~izda olacakt~ir.") & vbCrLf & vbCrLf & "Bilginize"
    stockQuote = "4500000001 kalem 10" & vbCrLf & "4500000001 kalem 20"
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", sampleStockReply, "YNT: 4500000001", "", reason)
    ReplyAssert got.Count = 2 And got.Exists("2") And got.Exists("3"), "Konudan iki stok kalemi", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", sampleStockReply, "YNT: Termin", stockQuote, reason)
    ReplyAssert got.Count = 2 And got.Exists("2") And got.Exists("3"), "Alintidan iki stok kalemi", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", sampleStockReply, "YNT: Termin", "", reason)
    ReplyAssert got.Count = 0, "Siparis kaniti olmayan stok bildirimi dagitilmamali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Bazi " & sampleStockReply, "YNT: 4500000001", "", reason)
    ReplyAssert got.Count = 0, "Kismi stok bildirimi iki kaleme dagitilmamali", passed
    sampleStockDate = ResolveDeliveryDate(sampleStockReply, DateSerial(2026, 11, 23), sampleStockState, stockOnly:=True)
    ReplyAssert sampleStockState = TERM_CONFIRMED And sampleStockDate = DateSerial(2026, 12, 1), "Tedarikci stok tarihi", passed
    sampleStockDate = ResolveDeliveryDate(sampleStockReply, DateSerial(2026, 11, 23), sampleStockState)
    ReplyAssert sampleStockState = 0 And IsEmpty(sampleStockDate), "Stok tarihi teslim termini sayilmamali", passed
    n = WriteReplyToSheet(cfg, ws, "ALFA", sampleStockReply, DateSerial(2026, 11, 1), _
        "test@example.invalid", "test", "", False, "4500000001")
    ReplyAssert n = 2 And ws.Cells(2, 14).Value2 = "Stok tarihi: 01.12.2026" And _
        ws.Cells(3, 14).Value2 = "Stok tarihi: 01.12.2026", "Iki kaleme stok tarihi yazilmali", passed
    ws.Range("H2:O3").ClearContents

    Dim wrapped As String, combinedMail As String
    wrapped = T("~Iyi g~unler") & vbCrLf & T("~Ur~unler") & vbCrLf & "01.12.2026 tarihinde" & vbCrLf & _
        T("stoklar~im~izda olacakt~ir.") & vbCrLf & "Bilginize"
    ReplyAssert InStr(CleanReplyContent(wrapped), T("~Ur~unler 01.12.2026 tarihinde stoklar~im~izda olacakt~ir.")) = 1, _
        "Selam ve satira bolunmus stok yaniti", passed
    combinedMail = "<p>" & sampleStockReply & "</p><div id=""divRplyFwdMsg""><p>4500000001 kalem 10</p>" & _
        "<p>4500000001 kalem 20</p></div>"
    ReplyAssert InStr(ExtractNewReply(combinedMail), "4500000001") = 0 And _
        InStr(FirstQuotedMessage(combinedMail), "kalem 20") > 0, "Outlook alinti blogu ayrilmali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", sampleStockReply, "YNT: 4500000001", "Siparis 4500000001", reason)
    ReplyAssert got.Count = 2, "Alintida sadece ayni siparis varsa iki kalem secilmeli", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Kalem 10 ve 20, 01.12.2026 tarihinde stokta olacak.", _
        "YNT: 4500000001", "", reason)
    ReplyAssert got.Count = 2, "Yanitta acikca listelenen 10 ve 20 numarali kalemler", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", sampleStockReply, "YNT: 4500000001", "4500000003 kalem 10", reason)
    ReplyAssert got.Count = 0, "Celisen siparis alintisina iki kalem yazilmamali", passed
    ws.Cells(2, 8).Value2 = SummarizeReply(sampleStockReply)
    ws.Cells(2, 9).Value = DateSerial(2026, 11, 1)
    ws.Cells(2, 14).Value2 = STATUS_BELIRSIZ
    n = WriteReplyToSheet(cfg, ws, "ALFA", sampleStockReply, DateSerial(2026, 11, 1), _
        "test@example.invalid", "test", "", False, "4500000001")
    ReplyAssert n = 2 And ws.Cells(2, 14).Value2 = "Stok tarihi: 01.12.2026", _
        "Ayni mailden daha once eksik kaydedilen stok durumu tamamlanmali", passed
    ws.Range("H2:O3").ClearContents

    Dim listed As Object, stockTable As String
    Set listed = ExplicitItemList("Kalemler: 0010, 20 ve 0030, 01.12.2026 tarihinde stokta olacak.")
    ReplyAssert listed.Count = 3 And listed.Exists("10") And listed.Exists("20") And listed.Exists("30"), _
        "Uc kalemlik liste ve yanindaki tarih ayri okunmali", passed
    ws.Cells(6, 1).Value2 = "4500000001"
    ws.Cells(6, 2).Value2 = "00030"
    ws.Cells(6, 3).Value2 = "ALFA"
    ws.Cells(6, 4).Value2 = "MAT60006"
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Kalemler: 10, 20 ve 30, 01.12.2026 tarihinde stokta olacak.", _
        "YNT: 4500000001", "", reason)
    ReplyAssert got.Count = 3 And got.Exists("2") And got.Exists("3") And got.Exists("6"), _
        "Acikca listelenen uc kaleme stok yaniti uygulanmali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", sampleStockReply, "YNT: 4500000001", "", reason)
    ReplyAssert got.Count = 0, "Kalemleri belirsiz mesaj uc kaleme dagitilmamali", passed
    ws.Range("A6:D6").ClearContents
    stockTable = "Siparis No" & vbTab & "Kalem No" & vbTab & "Stok Tarihi" & vbCrLf & _
                 "4500000001" & vbTab & "10" & vbTab & "01.12.2026"
    n = WriteReplyToSheet(cfg, ws, "ALFA", stockTable, DateSerial(2026, 11, 1), _
        "test@example.invalid", "test", "", False, "4500000001")
    ReplyAssert n = 1 And ws.Cells(2, 14).Value2 = "Stok tarihi: 01.12.2026" And _
        IsEmpty(ws.Cells(2, 15).Value), "Tablodaki stok tarihi teslim termini olmamali", passed
    ws.Range("H2:O3").ClearContents

    Dim oneRow As Variant, manyRows As Variant, indexedRows As Collection, indexedOrders As Object, indexedDetails As Variant
    oneRow = OrderColumnValues(ws, 2, 2, 1)
    manyRows = OrderColumnValues(ws, 2, 3, 1)
    ReplyAssert CStr(oneRow(1, 1)) = "4500000001" And CStr(manyRows(2, 1)) = "4500000001", _
        "Tek ve cok satirli Excel dizileri", passed
    LoadSheetRowIndex cfg, ws, "ALFA", indexedRows, indexedOrders, rowDetails:=indexedDetails
    ReplyAssert indexedRows.Count = 3 And indexedDetails.Exists("2") And indexedOrders.Exists("4500000001"), _
        "Satir endeksinde tedarikci ve siparis birlikte dogrulanmali", passed

    Set got = SelectReplyRows(cfg, ws, "ALFA", "4500000001 kalem 10 yarin teslim", "")
    ReplyAssert got.Count = 1 And got.Exists("2"), "Siparis ve kalem", passed
    TestOrderItemPairs cfg, ws, passed
    TestOrderSummary cfg, ws, passed
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
    ReplyAssert NormalizeEmail("Ornek <ORNEK@example.invalid>") = "ornek@example.invalid", "Gonderen adresi", passed
    ReplyAssert IsPublicMailDomain("gmail.com"), "Ortak domain eslestirilmemeli", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "MAT10001 yarin teslim", "Siparis 4500000002")
    ReplyAssert got.Count = 0, "Celisen konu siparisi", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "Siparis 4999999999 yarin teslim", "")
    ReplyAssert got.Count = 0, "Listede olmayan siparis", passed
    n = WriteReplyToSheet(cfg, ws, "ALFA", "4500000001 kalem 10 11.11.2026 teslim", DateSerial(2026, 11, 4), "test@example.invalid", "test", "", False)
    ReplyAssert n = 1 And Len(SafeStr(ws.Cells(3, 8).Value)) = 0 And Len(SafeStr(ws.Cells(4, 8).Value)) = 0, "Gercek yazma yalnizca hedef satira", passed
    n = WriteReplyToSheet(cfg, ws, "ALFA", "4500000001 kalem 10 10.11.2026 teslim", DateSerial(2026, 11, 3), "test@example.invalid", "test", "", False)
    ReplyAssert n = 0 And InStr(SafeStr(ws.Cells(2, 8).Value), "10.11.2026") = 0, "Eski mail yeni yaniti ezmemeli", passed
    ReplyAssert NormalizeReplySubject("RE: YNT: [External] Siparis 4500000001") = "siparis 4500000001", "Konu onekleri temizlenmeli", passed
    ReplyAssert NormalizeSasReferences("RE: SA Siparisi : SAS450000103") = "RE: SA Siparisi : SAS 450000103", _
        "Yapisik SAS siparis numarasi ayrilmali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Yarin teslim edilecek.", "RE: SAS4500000003", "", reason)
    ReplyAssert got.Count = 1 And got.Exists("5"), "SAS konulu tek kalem eslesmeli", passed
    TestSentThreadMatch cfg, ws, passed
    TestTrackingExplanations cfg, ws, passed
    TestVendorAliases passed
    TestVendorTitleAliases passed
    Dim sendSample As Date
    sendSample = DateSerial(2026, 11, 23) + TimeSerial(14, 15, 0)
    ReplyAssert FormatSentFollowUp(sendSample, 0) = T("Mail at~ild~i: 23.11.2026 14:15 | Geri d~on~u~s bekleniyor."), _
        "Dunku gonderim tarihi yanit olmasa da gorunmeli", passed
    ReplyAssert FormatSentFollowUp(sendSample, DateSerial(2026, 11, 24) + TimeSerial(9, 10, 0)) = _
        T("Mail at~ild~i: 23.11.2026 14:15 | Geri d~on~u~s al~ind~i: 24.11.2026 09:10"), _
        "Gonderim ve yanit tarihleri birlikte gorunmeli", passed
    ReplyAssert FormatSentFollowUp(sendSample, DateSerial(2026, 11, 22)) = FormatSentFollowUp(sendSample, 0), _
        "Eski yanit son gonderimi yanitlanmis gostermemeli", passed
    ReplyAssert ReplyTopicKey("ALFA", "RE: Siparis 4500000001", "1") = ReplyTopicKey("ALFA", "Siparis 4500000001", "2"), "Ayni konu tek grup olmali", passed
    ReplyAssert ReplyTopicKey("ALFA", "Siparis 4500000001", "1") <> ReplyTopicKey("BETA", "Siparis 4500000001", "2"), "Farkli tedarikciler birlesmemeli", passed
    ReplyAssert ReplyTopicKey("ALFA", "Siparis 4500000001", "1") <> ReplyTopicKey("ALFA", "Siparis 4500000002", "2"), "Farkli siparis konulari birlesmemeli", passed
    ReplyAssert SummarizeReply("Merhaba" & vbCrLf & "Yarin teslim edilecek" & vbCrLf & "Yarin teslim edilecek" & vbCrLf & "Saygilarimla") = "Yarin teslim edilecek", "Ozet selam imza ve tekrari cikarmali", passed
    ReplyAssert InStr(SummarizeReply("Yarin teslim edilemez, stok yok."), "edilemez") > 0, "Ozet olumsuzlugu korumali", passed
    ws.Cells(2, 13).Value = "+ Mail atildi"
    n = WriteReplyToSheet(cfg, ws, "ALFA", "4500000001 kalem 10 stok yok", DateSerial(2026, 11, 5), "test@example.invalid", "test", "", False)
    ReplyAssert n = 1 And InStr(SafeStr(ws.Cells(2, 8).Value), "stok yok") > 0, "Arti olan satir yeni maille guncellenmeli", passed
    ReplyAssert InStr(SafeStr(ws.Cells(2, 8).Value), "11.11.2026") = 0 And IsEmpty(ws.Cells(2, 15).Value), "Eski ozet ve termin temizlenmeli", passed
    ReplyAssert ws.Cells(2, 13).Value = "+ Mail atildi", "Mail atildi isareti korunmali", passed
    ws.Cells(4, 13).ClearContents
    MarkMailSent ws, 4, 13
    ReplyAssert ws.Cells(4, 13).Value2 = "+", "Gonderim isareti bos hucreye yazilmali", passed
    ws.Cells(4, 13).Value2 = "Takip"
    MarkMailSent ws, 4, 13
    MarkMailSent ws, 4, 13
    ReplyAssert ws.Cells(4, 13).Value2 = "+ Takip", "Mevcut metin korunmali ve arti tekrarlanmamali", passed
    ws.Cells(4, 13).Formula = "=1+1"
    MarkMailSent ws, 4, 13
    ReplyAssert ws.Cells(4, 13).HasFormula, "Gonderim durumundaki formul korunmali", passed
    ws.Cells(4, 13).ClearContents
    ReplyAssert wb.Worksheets("YanitGecmisi").Cells(2, 4).Value <> "", "Onceki yanit arsivlenmeli", passed
    ReplyAssert InStr(SummarizeReply("Merhaba, yarin teslim edilecek."), "teslim") > 0, "Selamla ayni satirdaki teslim bilgisi korunmali", passed
    ReplyAssert InStr(SummarizeReply("Kargo takip: https://example.invalid/123"), "https://") > 0, "Kargo baglantisi korunmali", passed
    ReplyAssert InStr(SummarizeReply("4500000001" & vbTab & "00010" & vbTab & "11.11.2026"), "11.11.2026") > 0, "Sayisal tablo telefon sanilmamali", passed
    ReplyAssert ParseTurkishDate("pazartesi", DateSerial(2026, 11, 4)) = DateSerial(2026, 11, 9), "Pazartesi pazarla karismamali", passed
    ReplyAssert ParseTurkishDate("haftaya pazartesi", DateSerial(2026, 11, 4)) = DateSerial(2026, 11, 9), "Gelecek hafta fazladan yedi gun eklenmemeli", passed
    ReplyAssert IsEmpty(ParseTurkishDate("haftaya teslim", DateSerial(2026, 11, 4))), "Belirsiz haftadan kesin tarih uydurulmamali", passed
    ReplyAssert ParseTurkishDate("04.11", DateSerial(2026, 11, 4) + TimeSerial(15, 0, 0)) = DateSerial(2026, 11, 4), "Ayni gun gelecek yila atlamamali", passed
    ReplyAssert IsEmpty(ParseTurkishDate("31.02.2026", DateSerial(2026, 11, 4))), "Gecersiz tarih baska aya kaymamali", passed
    ReplyAssert IsEmpty(ParseTurkishDate("2026-02-30", DateSerial(2026, 11, 4))), "Gecersiz ISO tarih yeniden yorumlanmamali", passed
    ReplyAssert ParseTurkishDate("12 kasim 2027", DateSerial(2026, 11, 4)) = DateSerial(2027, 11, 12), "Yazili ayda acik yil korunmali", passed
    Set got = SelectReplyRows(cfg, ws, "YOK", "yarin teslim", "", reason)
    ReplyAssert InStr(reason, "TEDARIKCI_SATIRI_YOK") = 1, "Tedarikci satiri yok nedeni ayri olmali", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "yarin teslim", "Siparis 4999999999", reason)
    ReplyAssert InStr(reason, "KONU_SIPARISI_TEDARIKCIDE_YOK") = 1, "Siparis yok nedeni ayri olmali", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "4500000001 yarin teslim", "", reason)
    ReplyAssert InStr(reason, "BIRDEN_FAZLA_KALEM_ADAYI") = 1, "Kalem belirsizligi ayri olmali", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "yarin teslim", "Acik siparisler hakkinda", reason)
    ReplyAssert InStr(reason, "YANITTA_SIPARIS_KALEM_MALZEME_YOK") = 1, "Genel yanit nedeni ayri olmali", passed
    context = "4500000001 kalem 10 11.11.2026 teslim"
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Yarin teslim edilecek", "Acik siparisler hakkinda", context, reason)
    ReplyAssert got.Count = 1 And got.Exists("2"), "Alintida tek kalem kisa yaniti baglamali", passed
    ReplyAssert InStr(CStr(got("2")), "11.11.2026") = 0, "Eski alintinin tarihi yeni yanita karismamali", passed
    context = "4500000001" & vbTab & "00010" & vbTab & "11.11.2026" & vbCrLf & _
        "4500000001" & vbTab & "00020" & vbTab & "12.11.2026"
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Kalem 20 yarin teslim", "Acik siparisler hakkinda", context, reason)
    ReplyAssert got.Count = 1 And got.Exists("3"), "Alintidan siparis yeni yanittan kalem", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Yarin teslim", "Acik siparisler hakkinda", context, reason)
    ReplyAssert got.Count = 0, "Alintida iki kalem genel yanita yetmez", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Tum kalemler yarin teslim", "Acik siparisler hakkinda", context, reason)
    ReplyAssert got.Count = 2, "Acik tum kalemler yalnizca alintidaki hedeflere uygulanmali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Siparis 4999999999 yarin teslim", "Acik siparisler hakkinda", context, reason)
    ReplyAssert got.Count = 0, "Yeni yanittaki celisen kodu alinti ezmemeli", passed
    Set got = SelectDeepReplyRows(cfg, ws, "BETA", "Yarin teslim", "Acik siparisler hakkinda", context, reason)
    ReplyAssert got.Count = 0, "Tek aday olsa bile alintidaki baska siparise yazilmamali", passed
    cfg("UzunMetinCol") = 5
    ws.Cells(2, 5).Value = "Hidrolik silindir baglanti aparati"
    Set got = SelectReplyRows(cfg, ws, "ALFA", "Hidrolik silindir baglanti aparati uretimde", "")
    ReplyAssert got.Count = 1 And got.Exists("2"), "Tam uzun aciklama benzersizse eslesmeli", passed
    ws.Cells(3, 5).Value = ws.Cells(2, 5).Value
    Set got = SelectReplyRows(cfg, ws, "ALFA", "Hidrolik silindir baglanti aparati uretimde", "")
    ReplyAssert got.Count = 0, "Ortak aciklama iki kaleme yayilmamali", passed
    ReplyAssert ReplyTopicKey("ALFA", "Acik siparisler hakkinda", "1", "C1") <> _
        ReplyTopicKey("ALFA", "Acik siparisler hakkinda", "2", "C2"), "Genel baslikta farkli konusmalar ayri kalmali", passed
    ReplyAssert ReplyTopicKey("ALFA", "RE: Acik siparisler hakkinda", "1", "C1") = _
        ReplyTopicKey("ALFA", "Acik siparisler hakkinda", "2", "C1"), "Ayni konusmanin tekrar yaniti birlesmeli", passed
    ReplyAssert InStr(FirstQuotedMessage("Yarin teslim" & vbCrLf & "From: first@example.invalid" & vbCrLf & _
        "4500000001" & vbCrLf & "From: second@example.invalid" & vbCrLf & "4500000002"), "4500000002") = 0, "Sadece ilk alinti mesaji kullanilmali", passed
    ReDim objects(1 To 3): ReDim times(1 To 3)
    For j = 1 To 3
        Set objects(j) = CreateObject("Scripting.Dictionary")
        objects(j).Add "id", j
    Next j
    times(1) = DateSerial(2026, 11, 1): times(2) = DateSerial(2026, 11, 3): times(3) = DateSerial(2026, 11, 2)
    SortMailObjects objects, times, 1, 3
    ReplyAssert objects(1)("id") = 2 And objects(3)("id") = 1, "Alt klasorler birlikte yeniden eskiye siralanmali", passed
    ReplyAssert ContainsItem("20 nolu kalem yarin teslim", "00020"), "Ters yazilan kalem etiketi", passed
    ReplyAssert ContainsItem("Kalem (00020) yarin teslim", "20"), "Parantezli kalem etiketi", passed
    ReplyAssert ContainsItem("4500000001/00020 yarin teslim", "20", "4500000001"), "Siparis bolu kalem bicimi", passed
    ReplyAssert Not ContainsItem("4500000001 11/20 teslim", "20", "4500000001"), "Tarih kalem numarasi sanilmamali", passed
    ReplyAssert Not HasItemLabel("10 kalem hazir"), "Kalem adedi kalem numarasi degil", passed
    ReplyAssert Not HasItemLabel("Kalem 10.12.2026 teslim"), "Kalem etiketi yanindaki tarih ayrilmali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Kalemler 10 ve 20" & vbCrLf & "Termin 01.12.2026", _
        "YNT: 4500000001", "", reason)
    ReplyAssert got.Count = 2 And InStr(CStr(got("2")), "01.12.2026") > 0 And InStr(CStr(got("3")), "01.12.2026") > 0, _
        "Iki kalemlik acik gruba alt satirdaki termin uygulanmali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Kalemler 10 ve 9999 teslim edilecek.", _
        "YNT: 4500000001", "", reason)
    ReplyAssert got.Count = 0, "Eksik kalem listesi parcali uygulanmamali", passed
    context = "4500000003 kalem 10 yarin teslim"
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Yarin teslim", "RE: 4500000001", context, reason)
    ReplyAssert got.Count = 0, "Konu ile celisen alintiyla kalem secilmemeli", passed
    ReplyAssert ExtractNewReply("Yarin teslim" & vbCrLf & "> 4500000001 kalem 10") = "Yarin teslim", "Buyuktur isaretli eski alinti ayrilmali", passed
    ReplyAssert InStr(FirstQuotedMessage("Yarin teslim" & vbCrLf & "> 4500000001 kalem 10"), "4500000001") > 0, "Basliksiz alinti baglam icin okunmali", passed
    ReplyAssert Not IsUsefulReply("Merhaba"), "Selamlama eski yaniti gizlememeli", passed
    ReplyAssert Not IsUsefulReply("Tesekkurler"), "Tesekkur eski yaniti gizlememeli", passed
    ReplyAssert Not IsUsefulReply("4500000001"), "Yalniz siparis kodu yanit sayilmamali", passed
    ReplyAssert IsUsefulReply("Yarin teslim edilecek"), "Teslim yaniti korunmali", passed
    ReplyAssert IsUsefulReply("12.11.2026"), "Yalniz termin tarihi korunmali", passed
    ReplyAssert ExplicitWholeOrder("Tum kalemler yarin teslim"), "Tum kalemler buyuk kucuk harf", passed
    ReplyAssert Not ExplicitWholeOrder("Tum kalemler degil, sadece hazir olanlar teslim"), "Kismi yanit tum kalemlere uygulanmamali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Tum kalemler yarin teslim", _
        "RE: 4500000001 ve 4500000003", "", reason)
    ReplyAssert got.Count = 0, "Birden fazla siparis icin toplu varsayim yapilmamali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Yarin teslim", "RE: 4500000001-10", _
        "4500000001 kalem 20", reason)
    ReplyAssert got.Count = 0 And InStr(reason, "CELISIYOR") > 0, "Konu ve alintidaki kalem farki", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "4500000001-20 yarin teslim", "", "", reason)
    ReplyAssert got.Count = 1 And got.Exists("3"), "Tireli siparis ve kalem", passed
    ReplyAssert InStr(FirstQuotedMessage("Yarin teslim" & vbCrLf & "> 4500000001 kalem 10" & vbCrLf & _
        ">> 4500000003 kalem 20"), "4500000003") = 0, "Daha eski ic ice alinti kullanilmamali", passed
    ReplyAssert CleanReplyContent("Merhaba" & vbCrLf & "Yarin teslim" & vbCrLf & "Saygilarimizla" & vbCrLf & "Telefon 0000000000") = "Yarin teslim", "Imza ve selamlama ayiklanmali", passed
    ReplyAssert Not ContainsItem("Miktar" & vbTab & "20" & vbTab & "12.11.2026", "20"), "Tablodaki miktar kalem sanilmamali", passed
    ReplyAssert ContainsItem("Kalem No.: 00020 yarin teslim", "20"), "Noktali kalem etiketi", passed
    ReplyAssert Not HasMonthOrWeek("market fiyati"), "Mart kelime icinde tarih sanilmamali", passed
    ReplyAssert HasMonthOrWeek("12 Kasim teslim"), "Ay adi korunmali", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "4500000001 kalem 10 MAT20002 yarin teslim", "", reason)
    ReplyAssert got.Count = 0, "Celisen malzeme ve kalem yazilmamali", passed
    context = T("Ornek <ornek@example.invalid> ~sunlar~i yazd~i (2 Kas 2026 14:16):") & vbCrLf & _
        "URETIM YERI" & vbTab & "MALZEME" & vbTab & "SIPARIS NO" & vbTab & "KALEM NO" & vbCrLf & _
        "Tesis A" & vbTab & "MAT10001" & vbTab & "04500000001" & vbTab & "00010"
    Dim sampleMail As String, quotePart As String, replyPart As String
    sampleMail = T("Merhabalar, 4 Kas~im'da teslim edilecektir.") & vbCrLf & _
        T("iPhone'umdan g~onderildi") & vbCrLf & context
    quotePart = FirstQuotedMessage(sampleMail)
    replyPart = CleanReplyContent(ExtractNewReply(sampleMail))
    ReplyAssert InStr(replyPart, "04500000001") = 0, "Alinti tablosu yeni yanita karismamali", passed
    ReplyAssert InStr(LCase$(replyPart), "iphone") = 0, "Turkce iPhone imzasi temizlenmeli", passed
    ReplyAssert InStr(quotePart, "04500000001") > 0, "Apple Mail alinti tablosu bulunmali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", replyPart, T("Re: A~c~ik sipari~sler hakk~inda"), quotePart, reason)
    ReplyAssert got.Count = 1, "Ekran goruntusu biciminde tek kalem bulunmali", passed
    ReplyAssert got.Exists("2"), "Sifirli alinti dogru siparis satirina gitmeli", passed
    ReplyAssert ContainsCode("04500000001", "4500000001"), "Basta sifirli mail", passed
    ReplyAssert ContainsCode("4500000001", "04500000001"), "Basta sifirsiz mail", passed
    ReplyAssert Not ContainsCode("14500000001", "4500000001"), "Farkli siparis kismi eslesmemeli", passed
    ReplyAssert NormalizeOrderNo("0004500000001") = "4500000001", "Kanonik siparis", passed
    ReplyAssert ParseTurkishDate(T("4 Kas~im'da teslim edilecektir"), DateSerial(2026, 11, 2)) = DateSerial(2026, 11, 4), "Turkce ay ve ek", passed
    ReplyAssert ParseTurkishDate(T("4 Kas~im'da teslim edildi"), DateSerial(2026, 11, 5)) = DateSerial(2026, 11, 4), "Dunku tarih gelecek yila atlamamali", passed
    Dim columns As Object, expanded As String
    Set columns = CreateObject("Scripting.Dictionary")
    expanded = ExpandTableMatchLine("Siparis No" & vbTab & "Miktar" & vbTab & "Kalem No" & vbTab & "Termin", columns)
    ReplyAssert expanded = "[TABLE_HEADER]", "Tablo basligi algilanmali", passed
    expanded = ExpandTableMatchLine("04500000001" & vbTab & "750" & vbTab & "20" & vbTab & "12.11.2026", columns)
    ReplyAssert ContainsItem(expanded, "00020", "4500000001"), "Baslikli tabloda sifirsiz kalem okunmali", passed
    ReplyAssert Not ContainsItem(expanded, "750", "4500000001"), "Miktar sutunu kaleme karismamali", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "Siparis No" & vbTab & "Kalem No" & vbTab & "Termin" & vbCrLf & _
        "04500000001" & vbTab & "20" & vbTab & "12.11.2026", "", reason)
    ReplyAssert got.Count = 1 And got.Exists("3"), "Sutun basliklari ile tam satir eslestirmesi", passed
    ReplyAssert ClassifyStatus("Iptal edilmedi, yarin teslim edilecek", True) = STATUS_TESLIM, "Olumsuz iptal ifadesi", passed
    ReplyAssert ClassifyStatus("Termin veremiyoruz", False) = STATUS_BELIRSIZ, "Termin belirsizligi iptal degildir", passed
    ReplyAssert ClassifyStatus("Teslim edildi", True) = "Teslim Edildi", "Gerceklesen teslim", passed
    ReplyAssert NormalizeEmail("Numune <numune@example.invalid>") = "selma@example.invalid", "Gorunen adli eposta", passed
    ReplyAssert NormalizeEmail("abc@@example.invalid") = "", "Gecersiz eposta reddedilmeli", passed
    ReplyAssert DecodeHtmlEntities("Kas&#305;m") = T("Kas~im"), "Sayisal HTML karakteri", passed
    ReplyAssert HasMaterialCodeLabel("Malzeme kodu MAT99999"), "Acik malzeme kodu", passed
    Set got = SelectReplyRows(cfg, ws, "ALFA", "4500000001 malzeme MAT99999 yarin teslim", "", reason)
    ReplyAssert got.Count = 0, "Bilinmeyen acik malzeme koduyla yazilmamali", passed
    ReplyAssert ReadSettingLong("hata", 180, 1, 3650) = 180, "Gecersiz arama ayari", passed
    ReplyAssert ReadSettingYesNo("HAYIR", "EVET") = "HAYIR", "Alt klasor ayari", passed
    ReplyAssert Not IsUsefulReply("Siparis listesi ektedir." & vbCrLf & "Uretim Departmani"), _
        "Liste bildirimi ve departman imzasi yazilmamali", passed
    ReplyAssert Not IsUsefulReply("Dikkat: Bu e-posta dis kaynakli. Siparis eklerini kontrol edin."), _
        "Dis kaynak uyarisi veri olarak yazilmamali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "BETA", "Yarin teslim", "Bilgi", "", reason)
    ReplyAssert got.Count = 0, "Tedarikci tek basina hedef sayilmamali", passed
    sampleMail = "Siparis No" & vbTab & "Miktar" & vbTab & "Kalem No" & vbTab & _
        "Istenen Tarih" & vbTab & "Yeni Termin" & vbCrLf & _
        "4500000001" & vbTab & "00010" & vbTab & "20" & vbTab & "07.11.2026" & vbTab & "13.11.2026"
    n = WriteReplyToSheet(cfg, ws, "ALFA", sampleMail, DateSerial(2026, 11, 6), _
        "test@example.invalid", "test", "", False)
    ReplyAssert n = 1 And ws.Cells(3, 8).Value = "Termin: 13.11.2026", _
        "Sadece yanit sutunu yazilmali; miktar kalemle karismamali", passed
    ReplyAssert ws.Cells(3, 15).Value = DateSerial(2026, 11, 13), "Eski istenen tarih termin olmamali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Kontrol ediyoruz", "RE: 4500000001", _
        "Onceki mesaja bakiniz", reason)
    ReplyAssert got.Count = 0 And InStr(reason, "ALINTI_KONU_SIPARISI_CELISIYOR") = 0, _
        "Siparis icermeyen alinti celiski sayilmamali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Yarin teslim", "RE: 4999999999", _
        "4500000001 kalem 10", reason)
    ReplyAssert got.Count = 0 And InStr(reason, "ALINTI_KONU_SIPARISI_CELISIYOR") > 0, _
        "Farkli siparisli alinti hedef yapilmamali", passed
    sampleMail = "<table><tr><th><p>SIPARIS_NO</p></th><th><p>SAT_KALEM_NO</p></th>" & _
        "<th><p>Istenen Tarih</p></th><th><p>Guncel Termin</p></th></tr>" & vbCrLf & _
        "<tr><td><p>4500000001</p></td><td><p>20</p></td><td><p>07.11.2026</p></td>" & _
        "<td><p>14.11.2026</p></td></tr></table>"
    n = WriteReplyToSheet(cfg, ws, "ALFA", CleanReplyContent(ExtractNewReply(sampleMail)), _
        DateSerial(2026, 11, 7), "test@example.invalid", "test", "", False)
    ReplyAssert n = 1 And ws.Cells(3, 8).Value = "Termin: 14.11.2026", _
        "Gercek basliklar ve HTML paragrafli hucreler", passed
    Dim termState As Long, termValue As Variant, refDate As Date
    refDate = DateSerial(2026, 11, 20)
    termValue = ResolveDeliveryDate("Eski termin 21.11.2026, yeni termin 25.11.2026", refDate, termState)
    ReplyAssert termState = TERM_CONFIRMED And termValue = DateSerial(2026, 11, 25), "Revize termin secilmeli", passed
    termValue = ResolveDeliveryDate("Termin 23.11.2026 veya 26.11.2026", refDate, termState)
    ReplyAssert termState = TERM_UNCERTAIN And IsEmpty(termValue), "Alternatif tarihler kesinlestirilmemeli", passed
    termValue = ResolveDeliveryDate("Termin 27.11.2026 uygun mudur?", refDate, termState)
    ReplyAssert termState = TERM_UNCERTAIN And IsEmpty(termValue), "Tarih sorusu kesin termin sayilmamali", passed
    termValue = ResolveDeliveryDate("5 is gunu icinde teslim edilecek", refDate, termState)
    ReplyAssert termState = TERM_UNCERTAIN And IsEmpty(termValue), "Teslim araligi kesin gun sayilmamali", passed
    termValue = ResolveDeliveryDate("Tahmini termin 27.11.2026", refDate, termState)
    ReplyAssert termState = TERM_UNCERTAIN And IsEmpty(termValue), "Tahmini tarih kesin termin olmamali", passed
    termValue = ResolveDeliveryDate("12-15 Aralik arasi teslim", refDate, termState)
    ReplyAssert termState = TERM_UNCERTAIN And IsEmpty(termValue), "Araliktan tek gun secilmemeli", passed
    termValue = ResolveDeliveryDate("Yarin teslim edilemez", refDate, termState)
    ReplyAssert termState = TERM_WITHDRAWN And IsEmpty(termValue), "Olumsuz teslim ifadesi", passed
    termValue = ResolveDeliveryDate("Termin 15.12.2026; stokta yok", refDate, termState)
    ReplyAssert termState = TERM_CONFIRMED And termValue = DateSerial(2026, 12, 15), "Stok yok yeni teslim tarihini silmemeli", passed
    termValue = ResolveDeliveryDate("Urunler 01.12.2026 stokta olacak; teslim tarihi 15.12.2026", refDate, termState)
    ReplyAssert termState = TERM_CONFIRMED And termValue = DateSerial(2026, 12, 15), "Stok tarihi ile teslim tarihi ayri okunmali", passed
    termValue = ResolveDeliveryDate("Fatura 24.11.2026 tarihinde gonderildi", refDate, termState)
    ReplyAssert termState = 0 And IsEmpty(termValue), "Fatura tarihi ayiklanmali", passed
    ReplyAssert Not IsUsefulReply("Fatura ve dekont gonderildi"), "Idari mail satir yazmamali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Yarin teslim", "RE: 4500000001/20", "", reason)
    ReplyAssert got.Count = 1 And got.Exists("3"), "Konu siparis/kalem cifti", passed
    Set got = SelectDeepReplyRows(cfg, ws, "BETA", "Yarin teslim", "4500000002/99", "", reason)
    ReplyAssert got.Count = 0, "Konudaki bilinmeyen kalem tek adaya yazilmamali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "Kalem 99" & vbCrLf & "Yarin teslim", "4500000003", "", reason)
    ReplyAssert got.Count = 0, "Eslesmeyen kalem sonraki satiri siparise baglamamali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "4500000001 ve 4500000003 kalem 10 yarin teslim", "", "", reason)
    ReplyAssert got.Count = 0, "Cok siparisli satirda kalem tahmin edilmemeli", passed
    n = WriteReplyToSheet(cfg, ws, "ALFA", "Kalem 20 kargo takip: 000123456", DateSerial(2026, 11, 8), _
        "test@example.invalid", "test", "", False, "4500000001")
    ReplyAssert n = 1 And ws.Cells(3, 15).Value = DateSerial(2026, 11, 14), _
        "Takip bilgisi kayitli termini silmemeli", passed
    Dim testSheets As New Collection, testVendors As Object, testActive As Object
    testSheets.Add ws
    Set testVendors = CreateObject("Scripting.Dictionary")
    testVendors.Add "ALFA", "ALFA": testVendors.Add "BETA", "BETA"
    Set testActive = ActiveOrderVendors(cfg, testSheets)
    ReplyAssert testActive.Count = 2 And testActive.Exists("ALFA"), "Secili tablodaki tedarikciler", passed
    ReplyAssert ResolveVendorCollision(cfg, testSheets, testVendors, "Yarin teslim", "") = "", _
        "Ortak posta adresi tek basina yeterli degil", passed
    ReplyAssert ResolveVendorCollision(cfg, testSheets, testVendors, "Yarin teslim", "4500000002") = "BETA", _
        "Ortak adres ve tek dogrulanmis siparis", passed
    ReplyAssert IsProtectedOrderHeader("DETAY_STATU") And IsProtectedOrderHeader("SIPARIS_DURUMU"), _
        "Ham SAP durumlari korunmali", passed
    ReplyAssert IsProtectedOrderHeader("Istenen Teslim Tarihi") And Not IsProtectedOrderHeader("Teyitli Termin"), _
        "Kaynak tarih ve yanit tarihi ayrilmali", passed
    ws.Cells(1, 14).Value = "DETAY_STATU"
    ws.Cells(1, 16).Value = "Tedarikci Yanit Durumu"
    ReplyAssert ResolveOutputColumn(cfg, testSheets, "DurumCol", Array("Tedarikci Yanit Durumu")) = 16, _
        "Ham durum yerine mevcut yanit sutunu kullanilmali", passed
    ws.Cells(1, 16).ClearContents
    ReplyAssert ResolveOutputColumn(cfg, testSheets, "DurumCol", Array("Tedarikci Yanit Durumu")) > 15, _
        "Yanita ait olmayan sutun uzerine yazilmamali", passed
    gPreview = True
    Dim oldReply As String, oldReplyDate As Variant, oldTermin As Variant
    oldReply = SafeStr(ws.Cells(3, 8).Value)
    oldReplyDate = ws.Cells(3, 9).Value
    oldTermin = ws.Cells(3, 15).Value
    n = WriteReplyToSheet(cfg, ws, "ALFA", "Kalem 20 termin 19.11.2026", DateSerial(2026, 11, 9), _
        "test@example.invalid", "test", "", False, "4500000001")
    gPreview = originalPreview
    ReplyAssert n = 1 And ws.Cells(3, 8).Value = oldReply And ws.Cells(3, 9).Value = oldReplyDate And ws.Cells(3, 15).Value = oldTermin, _
        "Onizleme aday bulmali ve sonuc hucrelerini degistirmemeli", passed
    ws.Cells(3, 15).Formula = "=1+1"
    n = WriteReplyToSheet(cfg, ws, "ALFA", "Kalem 20 termin 19.11.2026", DateSerial(2026, 11, 9), _
        "test@example.invalid", "test", "", False, "4500000001")
    ReplyAssert n = 0 And ws.Cells(3, 15).HasFormula, "Sonuc sutunundaki formul korunmali", passed
    ReplyAssert Not StartsWithOrderNumber("100000 adet teslim edilecek"), "Miktar siparis numarasi sayilmamali", passed
    ReplyAssert CleanReplyContent("Fatura 4999999999 gonderildi" & vbCrLf & "Yarin teslim") = "Yarin teslim", _
        "Idari satir eslesmeden once ayiklanmali", passed
    wb.Close SaveChanges:=False
    gMailHasVendorRows = originalVendorFlag
    MsgBox passed & " kontrol gecti. Outlook'a baglanilmadi; asil kitap degistirilmedi.", vbInformation
    Exit Sub
Failed:
    errText = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    gMailHasVendorRows = originalVendorFlag
    gPreview = originalPreview
    MsgBox "Oz test basarisiz: " & errText, vbCritical
End Sub

Private Sub TestSentThreadMatch(ByRef cfg As Object, ByVal ws As Worksheet, ByRef passed As Long)
    Dim previousThreads As Object, targets As Object, result As Object, key As String
    Dim sentAt As Date, receivedAt As Date, errCode As Long, errText As String
    Set previousThreads = gSentThreadLinks
    On Error GoTo Failed
    Set gSentThreadLinks = CreateObject("Scripting.Dictionary")
    Set targets = CreateObject("Scripting.Dictionary")
    Set result = CreateObject("Scripting.Dictionary")
    key = SentThreadKey("ALFA", "test-thread")
    sentAt = DateSerial(2026, 11, 23)
    receivedAt = sentAt + 1
    targets("row2") = Array(ws.CodeName, 2)
    IndexSentThread gSentThreadLinks, key, sentAt, targets
    ReplyAssert AppendVerifiedThreadMatch(cfg, ws, "ALFA", "test-thread", "Yarin teslim", "", "", receivedAt, result, ""), _
        "Ayni tedarikci ve Outlook zinciri", passed
    ReplyAssert result.Count = 1 And result.Exists("2"), "Yalnizca dogrulanmis satir", passed
    result.RemoveAll
    ReplyAssert Not AppendVerifiedThreadMatch(cfg, ws, "BETA", "test-thread", "Yarin teslim", "", "", receivedAt, result, ""), _
        "Baska tedarikciye uygulanmamali", passed
    ReplyAssert Not AppendVerifiedThreadMatch(cfg, ws, "ALFA", "test-thread", "Yarin teslim", "", "", sentAt, result, ""), _
        "Yanit gonderimden sonra olmali", passed
    ReplyAssert Not AppendVerifiedThreadMatch(cfg, ws, "ALFA", "test-thread", "Yarin teslim", "", _
        "4500000003 kalem 10", receivedAt, result, ""), "Alintida siparis varsa zincire geri donulmemeli", passed
    targets.RemoveAll
    targets("row3") = Array(ws.CodeName, 3)
    IndexSentThread gSentThreadLinks, key, sentAt - 1, targets
    ReplyAssert Not AppendVerifiedThreadMatch(cfg, ws, "ALFA", "test-thread", "Yarin teslim", "", "", receivedAt, result, ""), _
        "Birden fazla satir bulunan zincir otomatik eslesmemeli", passed
    Set gSentThreadLinks = previousThreads
    Exit Sub
Failed:
    errCode = Err.Number: errText = Err.Description
    Set gSentThreadLinks = previousThreads
    Err.Raise errCode, "TestSentThreadMatch", errText
End Sub

Private Sub TestTrackingExplanations(ByRef cfg As Object, ByVal ws As Worksheet, ByRef passed As Long)
    Dim prevOrders As Object, prevRegistered As Object, prevReasons As Object, prevUnmatched As Object
    Dim prevFailure As Boolean, key As String, reason As String, actual As String, whenSent As Date
    Dim failureCode As Long, failureText As String
    Set prevOrders = gOrderOnlySends: Set prevRegistered = gRegisteredVendors
    Set prevReasons = gReplyTrackingReasons: Set prevUnmatched = gUnmatchedVendorReplies
    prevFailure = gSentScanFailed
    On Error GoTo Failed
    Set gOrderOnlySends = CreateObject("Scripting.Dictionary")
    Set gRegisteredVendors = CreateObject("Scripting.Dictionary")
    Set gReplyTrackingReasons = CreateObject("Scripting.Dictionary")
    Set gUnmatchedVendorReplies = CreateObject("Scripting.Dictionary")
    gSentScanFailed = False
    key = TrackingOrderKey(ws, "ALFA", "4500000001")
    whenSent = DateSerial(2026, 11, 23) + TimeSerial(14, 15, 0)
    actual = UnverifiedSentFollowUp(ws, cfg, 3, "4500000001", "ALFA", 0, "")
    ReplyAssert InStr(actual, T("kay~itl~i e-posta adresi bulunamad~i")) > 0, "Adres eksigi satirda aciklanmali", passed
    gRegisteredVendors("ALFA") = True
    actual = UnverifiedSentFollowUp(ws, cfg, 3, "4500000001", "ALFA", 0, "")
    ReplyAssert InStr(actual, T("bu kaleme ait g~onderim bulunamad~i")) > 0, "Gonderim ve yanit yoksa neden gorunmeli", passed
    gOrderOnlySends(key) = whenSent
    actual = UnverifiedSentFollowUp(ws, cfg, 3, "4500000001", "ALFA", 0, "")
    ReplyAssert InStr(actual, "23.11.2026 14:15") > 0 And InStr(actual, T("do~grulanamad~i")) > 0, "Siparis ve kalem kaniti ayrilmali", passed
    gReplyTrackingReasons(ws.CodeName & ":3") = Array(whenSent + 1, "Yanit icerigi yetersiz.")
    reason = ReplyTrackingReason(ws.CodeName & ":3", "ALFA", whenSent, 0)
    ReplyAssert reason = "Yanit icerigi yetersiz.", "Dogrulanan kalemde atlama nedeni gorunmeli", passed
    gSentScanFailed = True
    actual = UnverifiedSentFollowUp(ws, cfg, 3, "4500000001", "ALFA", 0, "")
    ReplyAssert InStr(actual, "okunamad") > 0, "Gonderilenler hatasi eslesme yok diye gizlenmemeli", passed
RestoreTracking:
    Set gOrderOnlySends = prevOrders: Set gRegisteredVendors = prevRegistered
    Set gReplyTrackingReasons = prevReasons: Set gUnmatchedVendorReplies = prevUnmatched
    gSentScanFailed = prevFailure
    If failureCode <> 0 Then
        On Error GoTo 0
        Err.Raise failureCode, "TestTrackingExplanations", failureText
    End If
    Exit Sub
Failed:
    failureCode = Err.Number: failureText = Err.Description
    Resume RestoreTracking
End Sub

Private Sub TestOrderSummary(ByVal cfg As Object, ByVal ws As Worksheet, ByRef passed As Long)
    Dim quote As String, body As String, reason As String, got As Object, data As Variant
    Dim mailDate As Date, n As Long
    mailDate = DateSerial(2026, 11, 25)
    quote = "Siparis No" & vbTab & "Kalem No" & vbTab & "Miktar" & vbCrLf & _
        "04500000001" & vbTab & "00010" & vbTab & "120" & vbCrLf & _
        "4500000003" & vbTab & "10" & vbTab & "160" & vbCrLf & _
        "4500000001" & vbTab & "00020" & vbTab & "240" & vbCrLf & TABLE_END
    body = T("04500000001 numaral~i a~c~ik sipari~sinizde, 10 ve 11 beden depoya teslim edilmi~stir. ") & _
        T("Kalan a~c~ik sipari~sleriniz 3-4 hafta i~cinde haz~ir olacak.") & vbCrLf & _
        T("Stoklar~im~iza ula~s~inca sevk edece~giz.")
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", body, "RE: Acik siparisler hakkinda", quote, reason, mailDate)
    ReplyAssert got.Count = 3 And got.Exists("2") And got.Exists("3") And got.Exists("5"), _
        "Alintidaki siparisler kendi kalemlerine ayrilmali", passed
    data = got("2")
    ReplyAssert InStr(data(0), "10 ve 11 beden") > 0 And CBool(data(1)), "Teslim edilen bedenler ortak not olmali", passed
    data = got("5")
    ReplyAssert InStr(data(0), "3-4 hafta") > 0 And InStr(data(0), "sevk") > 0 And CBool(data(1)), _
        "Goreli sure ve sevk bilgisi diger siparise eklenmeli", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", body, "", "", reason, mailDate)
    ReplyAssert got.Count = 0 And reason = "SIPARIS_OZETI_TABLO_DOGRULANAMADI", "Alintisiz kalan kapsam atlanmali", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", body, "", quote & vbCrLf & quote, reason, mailDate)
    ReplyAssert got.Count = 0, "Birden fazla alinti tablosu arasinda tahmin yapilmamali", passed
    ws.Cells(2, 14).Value2 = "Bekliyor"
    ws.Cells(5, 14).Value2 = "Uretimde"
    ws.Cells(2, 15).Value = mailDate + 1
    n = WriteReplyToSheet(cfg, ws, "ALFA", body, mailDate, "test@example.invalid", "test", "", False, "", quote)
    ReplyAssert n = 3 And ws.Cells(2, 14).Value2 = "Bekliyor" And ws.Cells(5, 14).Value2 = "Uretimde", _
        "Siparis notu mevcut kalem durumunu korumali", passed
    ReplyAssert CDate(ws.Cells(2, 15).Value) = mailDate + 1 And IsEmpty(ws.Cells(5, 15).Value), _
        "Sure araligi termin yazmamali veya eski termini silmemeli", passed
    ws.Range("H2:O3").ClearContents
    ws.Range("H5:O5").ClearContents
End Sub

Private Sub TestOrderItemPairs(ByRef cfg As Object, ByVal ws As Worksheet, ByRef passed As Long)
    Dim pairs As Object, got As Object, reason As String, body As String, oldRows As Variant
    Set pairs = ReadOrderItemPairs("SAS: 450000101 · Kalem: 20")
    ReplyAssert pairs.Count = 1 And pairs.Exists("450000101/20"), "Etiketli siparis ve kalem cifti", passed
    Set pairs = ReadOrderItemPairs("0450000101 / 00020; 0450000102-00020")
    ReplyAssert pairs.Count = 2 And pairs.Exists("450000101/20") And pairs.Exists("450000102/20"), _
        "Ayni kalem kodu farkli siparislerde ayri kalmali", passed
    Set pairs = ReadOrderItemPairs("450000101 - 25.11.2026 tarihinde teslim")
    ReplyAssert pairs.Count = 0, "Teslim tarihi kalem numarasi olmamali", passed

    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "SAS: 04500000001 · Kalem No: 00020 yarin teslim", "", "", reason)
    ReplyAssert got.Count = 1 And got.Exists("3") And InStr(reason, "SIPARIS_KALEM_CIFTI") = 1, _
        "Etiketli cift tek satira baglanmali", passed
    body = "SAS: 04500000001" & vbCrLf & "Kalem: 00020" & vbCrLf & "Yarin teslim"
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", body, "", "", reason)
    ReplyAssert got.Count = 1 And got.Exists("3"), "Ayri satirlardaki etiketli cift", passed
    body = "Kalem No" & vbTab & "Siparis No" & vbTab & "Termin" & vbCrLf & _
           "20" & vbTab & "04500000001" & vbTab & "12.11.2026"
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", body, "", "", reason)
    ReplyAssert got.Count = 1 And got.Exists("3"), "Sutun sirasi degisse de ayni tablo cifti", passed
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "04500000001 / 99999 yarin teslim", "", _
        "4500000001 kalem 20", reason)
    ReplyAssert got.Count = 0 And InStr(reason, REASON_PAIR_FAILED) = 1, _
        "Gecersiz yeni cift eski alintiya dusmemeli", passed

    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "4500000001 kalem 10 ve 20, yarin teslim", "", "", reason)
    ReplyAssert got.Count = 2 And got.Exists("2") And got.Exists("3"), "Acik kalem listesi korunmali", passed

    oldRows = ws.Range("A6:D7").Value2
    ws.Cells(6, 1).Value2 = "0450000101"
    ws.Cells(7, 1).Value2 = "0450000102"
    ws.Range("B6:B7").Value2 = "00020"
    ws.Range("C6:C7").Value2 = "ALFA"
    ws.Range("D6:D7").Value2 = "XX00000001"
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "0450000101/00020 yarin teslim", "", "", reason)
    ws.Range("A6:D7").Value2 = oldRows
    ReplyAssert got.Count = 1 And got.Exists("6"), "Kalemi ve malzemesi ayni olan farkli siparisler karismamali", passed

    ws.Cells(6, 1).Value2 = "4500000001"
    ws.Cells(6, 2).Value2 = "20"
    ws.Cells(6, 3).Value2 = "ALFA"
    ws.Cells(6, 4).Value2 = "MAT60006"
    Set got = SelectDeepReplyRows(cfg, ws, "ALFA", "4500000001/20 yarin teslim", "", "", reason)
    ws.Range("A6:D7").Value2 = oldRows
    ReplyAssert got.Count = 0 And InStr(reason, "BIRDEN_FAZLA_KALEM_ADAYI") > 0, _
        "Ayni cift iki Excel satirinda ise belirsiz kalmali", passed
End Sub

Private Sub TestVendorAliases(ByRef passed As Long)
    Dim active As Object, directory As Object, notes As Object, resolved As Object, emails As Object, domains As Object
    Set active = CreateObject("Scripting.Dictionary")
    Set directory = CreateObject("Scripting.Dictionary")
    Set notes = CreateObject("Scripting.Dictionary")
    active("ALFA MAKINA") = True
    directory("ALFA MAKINA LTD STI") = 3
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 1 And resolved("ALFA MAKINA LTD STI") = "ALFA MAKINA", _
        "Tekil unvan farki eslesmeli", passed
    directory("ALFA MAKINA A S") = 3
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 0 And notes.Exists("ALFA MAKINA"), "Unvan cakisirsa tahmin edilmemeli", passed
    notes.RemoveAll
    directory("ALFA MAKINA") = 1
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 0 And notes.Count = 0, "Tam ad eslesmesinin onceligi korunmali", passed
    active.RemoveAll: directory.RemoveAll: notes.RemoveAll
    active("ALFA MAKINA A S") = True
    directory("ALFA MAKINA LTD STI") = 3
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 0 And notes.Exists("ALFA MAKINA A S"), _
        "Farkli sirket turleri otomatik birlestirilmemeli", passed
    active.RemoveAll: directory.RemoveAll: notes.RemoveAll
    Dim shortName As String, fullName As String
    shortName = NormalizeName("ÖRNEKFÝRMA ELEK.MAK.ÝNÞ.TURZ.SAN.T")
    fullName = NormalizeName("ÖRNEKFÝRMA ELEK.MAK.ÝNÞ.TURZ.SAN.TÝC.A.")
    active(shortName) = True
    directory(fullName) = 3
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 1 And resolved(fullName) = shortName, "Son kelimesi kesilen tedarikci adi", passed
    directory(fullName & " S") = 3
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 0 And notes.Exists(shortName), "Birden fazla unvan eslesmesi", passed
    directory.Remove fullName & " S"
    notes.RemoveAll
    active(fullName) = True
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 0 And notes.Exists(shortName), "Tam siparis adi kisaltilmis addan once gelmeli", passed
    active.Remove fullName
    notes.RemoveAll
    active(NormalizeName("ÖRNEKFÝRMA ELEK.MAK.ÝNÞ.TURZ.SAN.TÝ")) = True
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 0 And notes.Count = 2, "Ayni kayda uyan iki kisaltilmis ad", passed
    ReplyAssert Len(VendorDirectoryIssue(3)) = 0, "Gecerli adres icin uyari olmamali", passed
    ReplyAssert VendorDirectoryIssue(1) <> VendorDirectoryIssue(2), "Bos ve gecersiz adres ayri raporlanmali", passed
    ReplyAssert Len(VendorDirectoryIssue(0)) > 0 And Len(VendorDirectoryIssue(-1)) > 0, _
        "Blokeli ve bulunamayan kayit icin uyari verilmeli", passed
    Set emails = CreateObject("Scripting.Dictionary")
    Set domains = CreateObject("Scripting.Dictionary")
    domains("example.invalid") = "ALFA"
    ReplyAssert FindSentRecipientVendor("contact@example.invalid", emails, domains) = "ALFA", "Tekil kurumsal alan adi", passed
    emails("contact@example.invalid") = "***CAKISMA***"
    ReplyAssert FindSentRecipientVendor("contact@example.invalid", emails, domains) = "", "Adres cakisimi alan adina dusmemeli", passed
End Sub

Private Sub TestVendorTitleAliases(ByRef passed As Long)
    Dim active As Object, directory As Object, notes As Object, resolved As Object
    Dim orderName As String, masterName As String, otherName As String
    Set active = CreateObject("Scripting.Dictionary")
    Set directory = CreateObject("Scripting.Dictionary")
    Set notes = CreateObject("Scripting.Dictionary")

    masterName = NormalizeName("ÖRNEKFÝRMA ELEKTRÝK MAKÝNE ÝNÞAAT TURÝZM SANAYÝ TÝCARET A.Þ.")
    otherName = NormalizeName("ÖRNEKFÝRMA ELEKTRÝK MAKÝNE ÝNÞAAT TURÝZM SANAYÝ TEKSTÝL A.Þ.")
    directory(masterName) = 3
    active("ORNEKFIRMA") = True
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 1 And resolved(masterName) = "ORNEKFIRMA", "Tekil firma adi", passed

    directory(otherName) = 3
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 0 And notes.Exists("ORNEKFIRMA"), "Eksiz yinelenen firma adi", passed
    active.RemoveAll: notes.RemoveAll
    orderName = NormalizeName("ÖRNEKFÝRMA ELEK.MAK.ÝNÞ.TURZ.SAN.TÝ")
    active(orderName) = True
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 1 And resolved(masterName) = orderName, "Son unvan ekiyle secim", passed

    active.RemoveAll: notes.RemoveAll
    orderName = NormalizeName("ÖRNEKFÝRMA ELEK.MAK.ÝNÞ.TURZ.SAN.T")
    active(orderName) = True
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 0 And notes.Exists(orderName), "T ile baslayan iki unvan", passed

    directory.Remove otherName
    directory(NormalizeName("ÖRNEKFÝRMA ÝNÞAAT MAKÝNE ELEKTRÝK TURÝZM SANAYÝ TÝCARET A.Þ.")) = 3
    notes.RemoveAll
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 1 And resolved(masterName) = orderName, "Ilk unvan ekiyle secim", passed

    directory(NormalizeName("ÖRNEKFÝRMA ELEKTRÝK MAKÝNE ÝNÞAAT TURÝZM SANAYÝ TÝCARET LTD. ÞTÝ.")) = 3
    active.RemoveAll: notes.RemoveAll
    orderName = NormalizeName("ÖRNEKFÝRMA ELEK.MAK.ÝNÞ.TURZ.SAN.TÝC.A.")
    active(orderName) = True
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 1 And resolved(masterName) = orderName, "Son sirket eki A ve LTD ayrilmali", passed
    active.RemoveAll: directory.RemoveAll: notes.RemoveAll
    active("ALFA INS") = True
    directory("ALFA INSAAT") = 3
    directory("ALFA INSAAT TURIZM") = 3
    Set resolved = ResolveVendorAliases(active, directory, notes)
    ReplyAssert resolved.Count = 1 And resolved("ALFA INSAAT") = "ALFA INS", "Tam unvan eki onceligi", passed
    ReplyAssert VendorSuffixMatches("MAK KES", "MAK KESICI AS"), "Kesik son unvan kelimesi", passed
    ReplyAssert Not VendorSuffixMatches("MAK 12", "MAK 123 AS"), "Sayisal unvan kodu kesik sayilmamali", passed
End Sub

Private Sub ReplyAssert(ByVal condition As Boolean, ByVal title As String, ByRef passed As Long)
    If Not condition Then Err.Raise vbObjectError + 620, "MailOkumaOzTest", title
    passed = passed + 1
End Sub

Private Function NormalizeReplySubject(ByVal subject As String) As String
    Dim re As Object, before As String
    subject = Trim$(Replace(subject, ChrW(160), " "))
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True: re.Global = False
    re.Pattern = "^\s*(\[(external|dis|spam)\]\s*|(re|fw|fwd|aw|wg|sv|ynt|ilt|yanit|yanýt)(\[[0-9]+\])?\s*:\s*)"
    Do
        before = subject
        subject = Trim$(re.Replace(subject, ""))
    Loop Until before = subject
    ' Preserve punctuation and order/item numbers; do not merge similar but distinct subjects.
    NormalizeReplySubject = LCase$(subject)
End Function

Private Function ReplyTopicKey(ByVal vendor As String, ByVal subject As String, ByVal entryId As String, _
    Optional ByVal conversationId As String = "") As String
    Dim normalized As String
    normalized = NormalizeReplySubject(subject)
    If Len(normalized) = 0 Then
        normalized = "ENTRY:" & entryId
    ElseIf ExtractNumbers(normalized).Count = 0 And Len(conversationId) > 0 Then
        normalized = normalized & " | CONVERSATION:" & conversationId
    End If
    ReplyTopicKey = CStr(Len(NormalizeName(vendor))) & ":" & NormalizeName(vendor) & ":" & normalized
End Function

Private Function IsAutomaticReply(ByVal mail As Object, ByVal subject As String) As Boolean
    Dim t As String, headers As String, re As Object
    t = NormalizeName(NormalizeReplySubject(subject))
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(automatic reply|auto reply|out of office|otomatik yanit|otomatik cevap|undeliverable|delivery status notification|iletilemedi)( |$)"
    If re.Test(t) Then IsAutomaticReply = True: Exit Function
    ' Optional transport headers; absence is not an error for ordinary mail.
    On Error Resume Next
    headers = SafeStr(mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptags/0x007D001E"))
    On Error GoTo 0
    re.MultiLine = True
    re.Pattern = "^Auto-Submitted:[ \t]*(auto-replied|auto-generated)([ \t;]|$)"
    IsAutomaticReply = re.Test(Replace(headers, vbCrLf, vbLf))
End Function

Private Function SummarizeReply(ByVal reply As String, Optional ByVal mailDate As Date = 0) As String
    Dim lines As Variant, part As Variant, cells As Variant, cell As Variant
    Dim line As String, norm As String, selected As String, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    If mailDate = 0 Then mailDate = Date
    reply = CleanReplyContent(ExtractNewReply(reply))
    lines = Split(Replace(reply, vbCrLf, vbLf), vbLf)
    For Each part In lines
        cells = Split(CStr(part), vbTab)
        For Each cell In cells
            line = Trim$(CStr(cell))
            If IsOrderUpdate(line, mailDate) Then
                norm = UCase$(line)
                If Not seen.Exists(norm) Then
                    seen.Add norm, True
                    If Len(selected) > 0 Then selected = selected & " / "
                    selected = selected & line
                End If
            End If
        Next cell
    Next part
    SummarizeReply = selected
End Function


Private Function IsOrderUpdate(ByVal line As String, ByVal mailDate As Date) As Boolean
    Static detail As Object, generic As Object, timing As Object, footer As Object
    Dim norm As String, plain As String, parsedDate As Variant
    If Len(line) = 0 Then Exit Function
    If IsReplyNoise(line) Then Exit Function
    plain = Trim$(WithoutLinks(line))
    norm = NormalizeName(plain)
    If Len(norm) = 0 Then Exit Function
    If detail Is Nothing Then
        Set detail = CreateObject("VBScript.RegExp")
        detail.IgnoreCase = True
        detail.Pattern = "\b(teslim|termin|sevk[a-z]*|kargo[a-z]*|uretim[a-z]*|imalat[a-z]*|stok[a-z]*|iptal|gecik[a-z]*|" & _
            "onaylan[a-z]*|onaylad[a-z]*|teyit (edil[a-z]*|ett[a-z]*)|hazir[a-z]*|tukendi|reddedil[a-z]*|fiyat farki|" & _
            "yola cikti|yolda|kabul edildi|gonderildi|gonderilecek|adet|miktar|eksik|tedarik ed[a-z]*)\b"
        Set generic = CreateObject("VBScript.RegExp")
        generic.IgnoreCase = True
        generic.Pattern = "^(siparis( no)?|kalem( no)?|malzeme( kodu)?|termin( tarihi)?|durum|aciklama|not|teslim(at)?|" & _
            "sevk|stok|miktar|adet|uretim|kargo|kargo takip|bilgi|tarih|teyit|onay|hazirlik|bekliyoruz|termin n a|termin na)$"
        Set timing = CreateObject("VBScript.RegExp")
        timing.IgnoreCase = True
        timing.Pattern = "^(bugun|yarin|obur gun|haftaya|gelecek hafta|bu hafta|hafta ici|hafta sonu|" & _
            "pazartesi|sali|carsamba|persembe|cuma|cumartesi|pazar)$|\b[0-9]+\s*(is\s*)?(gun|hafta)\s*(sonra|icinde)?\s*$"
        Set footer = CreateObject("VBScript.RegExp")
        footer.IgnoreCase = True
        footer.Pattern = "^(siparis(ler)? (hakkinda|listesi|tablosu)|siparis (hatti|iletisim|sorumlusu)|acik siparisler( hakkinda| listesi)?|" & _
            "(stok|termin|teslimat?|siparis|sevkiyat|uretim|kargo) (yonetimi|departmani|bolumu|birimi|ekibi|uzmani|muduru|sorumlusu|operasyonlari)( |$))"
    End If
    If generic.Test(norm) Or footer.Test(norm) Or IsAdministrativeLine(norm) Then Exit Function
    If detail.Test(norm) Or timing.Test(norm) Then
        IsOrderUpdate = True
        Exit Function
    End If
    ' A date by itself is useful when the order line is already known.
    ' Free-form administrative dates and date-like fragments inside codes are skipped.
    Static dateOnly As Object
    If dateOnly Is Nothing Then
        Set dateOnly = CreateObject("VBScript.RegExp")
        dateOnly.IgnoreCase = True
        dateOnly.Pattern = "^[0-9]{1,4}[./-][0-9]{1,2}([./-][0-9]{1,4})?[.]?$|" & _
            "^[0-9]{1,2} (ocak|subat|mart|nisan|mayis|haziran|temmuz|agustos|eylul|ekim|kasim|aralik)( [0-9]{4})?$"
    End If
    If dateOnly.Test(plain) Or dateOnly.Test(norm) Then
        parsedDate = ParseTurkishDate(plain, mailDate)
        IsOrderUpdate = Not IsEmpty(parsedDate)
    End If
End Function

Private Function WithoutLinks(ByVal line As String) As String
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = True
        re.Pattern = "(https?://|www\.)[^\s<>]+"
    End If
    WithoutLinks = re.Replace(line, "")
End Function

Private Sub ArchivePreviousReply(ByVal ws As Worksheet, ByVal rowNum As Long, _
    ByVal replyCol As Long, ByVal dateCol As Long, ByVal senderCol As Long, _
    ByVal attachCol As Long, ByVal statusCol As Long, ByVal terminCol As Long, ByVal replacedAt As Date)
    Dim history As Worksheet, r As Long, i As Long, headers As Variant
    If Len(SafeStr(ws.Cells(rowNum, replyCol).Value)) = 0 Then Exit Sub
    On Error Resume Next
    Set history = ws.Parent.Worksheets("YanitGecmisi")
    On Error GoTo 0
    If history Is Nothing Then
        Set history = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
        history.Name = "YanitGecmisi"
        headers = Array("Arsiv Zamani", "Sayfa", "Satir", "Onceki Yanit", "Onceki Yanit Tarihi", "Onceki Gonderen", "Onceki Durum", "Onceki Termin", "Onceki Ekler", "Yeni Mail Tarihi")
        For i = LBound(headers) To UBound(headers)
            history.Cells(1, i + 1).Value = headers(i)
        Next i
        history.Rows(1).Font.Bold = True
    End If
    If SafeStr(history.Cells(1, 1).Value) <> "Arsiv Zamani" Or SafeStr(history.Cells(1, 4).Value) <> "Onceki Yanit" Then
        Err.Raise vbObjectError + 621, , "YanitGecmisi adinda farkli bir sayfa var. Sayfayi yedekleyip yeniden adlandirin."
    End If
    r = history.Cells(history.Rows.Count, 1).End(xlUp).Row + 1
    history.Cells(r, 1).Value = Now
    history.Cells(r, 2).NumberFormat = "@": history.Cells(r, 2).Value = ws.Name
    history.Cells(r, 3).Value = rowNum
    history.Cells(r, 4).NumberFormat = "@": history.Cells(r, 4).Value = SafeStr(ws.Cells(rowNum, replyCol).Value)
    history.Cells(r, 5).Value = ws.Cells(rowNum, dateCol).Value
    If senderCol > 0 Then
        history.Cells(r, 6).NumberFormat = "@": history.Cells(r, 6).Value = SafeStr(ws.Cells(rowNum, senderCol).Value)
    End If
    If statusCol > 0 Then
        history.Cells(r, 7).NumberFormat = "@": history.Cells(r, 7).Value = SafeStr(ws.Cells(rowNum, statusCol).Value)
    End If
    If terminCol > 0 Then history.Cells(r, 8).Value = ws.Cells(rowNum, terminCol).Value
    If attachCol > 0 Then
        history.Cells(r, 9).NumberFormat = "@": history.Cells(r, 9).Value = SafeStr(ws.Cells(rowNum, attachCol).Value)
    End If
    history.Cells(r, 10).Value = replacedAt
    history.Cells(r, 1).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    history.Cells(r, 5).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    history.Cells(r, 8).NumberFormat = "dd.mm.yyyy"
    history.Cells(r, 10).NumberFormat = "dd.mm.yyyy hh:mm:ss"
End Sub

' The quoted request may identify a row; it never supplies the new status/date/summary.
Private Function SelectDeepReplyRows(ByRef cfg As Object, ByVal ws As Worksheet, _
    ByVal vendor As String, ByVal reply As String, ByVal subject As String, _
    ByVal quoted As String, ByRef reason As String, Optional ByVal mailDate As Date = 0) As Object
    Dim direct As Object, contextRows As Object, contextual As Object, orders As Object, allOrders As Object
    Dim r As Long, lastRow As Long, code As String, orderText As String, key As Variant
    Dim originalReason As String, subjectAgrees As Boolean, subjectHasOrder As Boolean, quoteHasOrder As Boolean
    Dim checkQuotedItems As Boolean
    subject = NormalizeSasReferences(subject)
    quoted = NormalizeSasReferences(quoted)
    Set direct = MatchOrderSummary(cfg, ws, vendor, subject, reply, quoted, reason)
    If Not direct Is Nothing Then
        Set SelectDeepReplyRows = direct
        Exit Function
    End If
    Set direct = SelectReplyRows(cfg, ws, vendor, reply, subject, reason, True)
    Set SelectDeepReplyRows = direct
    If InStr(1, reason, "TEDARIKCI_SATIRI_YOK", vbBinaryCompare) = 1 Then Exit Function
    If InStr(1, reason, REASON_PAIR_FAILED, vbBinaryCompare) = 1 Then Exit Function
    If InStr(1, reason, REASON_PAIR_MATCHED, vbBinaryCompare) <> 1 Then
        If TryMatchStockItems(cfg, ws, vendor, reply, subject, quoted, direct, mailDate) Then
            reason = "STOK_KALEMLERI_ESLESTI | Kalem=" & direct.Count
            Exit Function
        End If
    End If
    If direct.Count = 0 And IsStockAvailability(reply) Then reason = reason & " | STOK_IFADESI_VAR"
    If Len(quoted) = 0 Then Exit Function
    originalReason = reason
    Set orders = CreateObject("Scripting.Dictionary")
    Dim contextVendorRows As Collection, contextRow As Variant
    LoadSheetRowIndex cfg, ws, vendor, contextVendorRows, allOrders
    subjectHasOrder = SubjectHasOrderReference(subject, allOrders)
    quoteHasOrder = HasKnownOrder(quoted, allOrders) Or HasOrderLabel(quoted)
    For Each contextRow In contextVendorRows
        code = SafeStr(ws.Cells(CLng(contextRow), CLng(cfg("OrderNoCol"))).Value)
        If Len(code) >= 4 And ContainsCode(quoted, code) Then orders(NormalizeOrderNo(code)) = True
    Next contextRow
    If direct.Count > 0 Then
        If Not ReplyHasExplicitCodes(reply) And quoteHasOrder Then
            If Not subjectHasOrder Or orders.Count > 0 Or HasKnownOrder(quoted, allOrders) Then
                checkQuotedItems = (HasItemLabel(subject) Or HasOrderItemPair(subject))
                If checkQuotedItems Then Set contextRows = SelectReplyRows(cfg, ws, vendor, quoted, "", reason, True)
                For Each key In direct.Keys
                    code = SafeStr(ws.Cells(CLng(key), CLng(cfg("OrderNoCol"))).Value)
                    If Not ContainsCode(quoted, code) Then
                        direct.Remove key
                    ElseIf checkQuotedItems Then
                        If contextRows.Count > 0 Then
                            If Not contextRows.Exists(CStr(key)) Then direct.Remove key
                        End If
                    End If
                Next key
                If direct.Count = 0 Then
                    reason = "ALINTI_SIPARIS_VEYA_KALEM_CELISIYOR"
                    Set SelectDeepReplyRows = direct: Exit Function
                End If
                reason = originalReason
            End If
        End If
        Set SelectDeepReplyRows = direct: Exit Function
    End If
    ' If the reply names an order, material or item that conflicts, never overwrite that evidence.
    If orders.Count = 1 And Not subjectHasOrder Then
        For Each key In orders.Keys: orderText = CStr(key): Next key
        Set contextual = SelectReplyRows(cfg, ws, vendor, reply, orderText, reason, True)
        If contextual.Count > 0 And ExplicitWholeOrder(reply) Then
            Set contextRows = SelectReplyRows(cfg, ws, vendor, quoted, "", reason, True)
            For Each key In contextual.Keys
                If Not contextRows.Exists(CStr(key)) Then contextual.Remove key
            Next key
        End If
        If contextual.Count > 0 Then
            reason = "ALINTIDAN_SIPARIS_BAGLAMI | " & reason
            Set SelectDeepReplyRows = contextual: Exit Function
        End If
    End If
    If orders.Count = 0 Then
        reason = originalReason & " | ALINTIDA_SECILI_SIPARIS_YOK"
        Exit Function
    End If
    If subjectHasOrder Then
        For Each key In orders.Keys
            If ContainsCode(subject, CStr(key)) Then subjectAgrees = True
        Next key
        If Not subjectAgrees Then
            reason = originalReason & " | ALINTI_KONU_SIPARISI_CELISIYOR"
            Exit Function
        End If
    End If
    Set contextRows = SelectReplyRows(cfg, ws, vendor, quoted, "", reason, True)
    For Each key In contextRows.Keys
        code = SafeStr(ws.Cells(CLng(key), CLng(cfg("OrderNoCol"))).Value)
        If Not orders.Exists(NormalizeOrderNo(code)) Or (subjectHasOrder And Not ContainsCode(subject, code)) Then
            contextRows.Remove key
        End If
    Next key
    If Not ReplyHasExplicitCodes(reply) And IsUsefulReply(reply) Then
        If contextRows.Count = 1 And (orders.Count = 1 Or subjectHasOrder) Then
            For Each key In contextRows.Keys
                direct.Add CStr(key), reply
            Next key
            reason = "ALINTIDA_TEK_HEDEF | Aday=" & contextRows.Count
            Exit Function
        End If
    End If
    reason = originalReason & " | AlintiSiparisi=" & orders.Count & " | AlintiHedefi=" & contextRows.Count
End Function

' A partial delivery statement is a note for the named order. The remaining
' statement only applies to orders from the first quoted order table.
Private Function MatchOrderSummary(ByVal cfg As Object, ByVal ws As Worksheet, _
    ByVal vendor As String, ByVal subject As String, ByVal reply As String, ByVal quoted As String, _
    ByRef reason As String) As Object
    Dim namedOrder As String, deliveredNote As String, remainingNote As String
    Dim notesOnly As Boolean, pairs As Object, targets As Object, result As Object
    Dim subjectPairs As Object, remainingOrders As Object, target As Variant, pair As Variant
    Dim orderNo As String, rowKey As String

    If Not SplitOrderSummary(reply, namedOrder, deliveredNote, remainingNote, notesOnly) Then Exit Function
    Set subjectPairs = ReadOrderItemPairs(subject)
    If subjectPairs.Count > 0 Then Exit Function

    Set result = CreateObject("Scripting.Dictionary")
    Set MatchOrderSummary = result
    reason = "SIPARIS_OZETI_TABLO_DOGRULANAMADI"
    Set pairs = QuotedOrderSummaryPairs(quoted)
    If pairs Is Nothing Then Exit Function

    Set remainingOrders = CreateObject("Scripting.Dictionary")
    Dim namedInTable As Boolean
    For Each pair In pairs.Keys
        orderNo = Split(CStr(pair), "/")(0)
        If orderNo = namedOrder Then
            namedInTable = True
        Else
            remainingOrders(orderNo) = True
        End If
    Next pair
    If Not namedInTable Or remainingOrders.Count = 0 Then Exit Function

    Set targets = VerifiedOrderSummaryTargets(cfg, ws, vendor, pairs)
    If targets Is Nothing Then Exit Function
    For Each pair In targets.Keys
        target = targets(pair)
        If CStr(target(0)) = ws.CodeName Then
            rowKey = CStr(target(1))
            If Split(CStr(pair), "/")(0) = namedOrder Then
                result.Add rowKey, Array(deliveredNote, True)
            Else
                result.Add rowKey, Array(remainingNote, notesOnly)
            End If
        End If
    Next pair
    If result.Count > 0 Then reason = "SIPARIS_OZETI | AcikSiparis=" & namedOrder & _
        " | KalanSiparisler=" & Join(remainingOrders.Keys, ", ")
End Function

Private Function SplitOrderSummary(ByVal reply As String, ByRef namedOrder As String, _
    ByRef deliveredNote As String, ByRef remainingNote As String, ByRef notesOnly As Boolean) As Boolean
    Static remaining As Object, delivered As Object, orderNumber As Object, relativeTerm As Object, orderHeading As Object
    Dim sections As Object, marker As Object, identifiers As Object, pairs As Object
    Dim firstPart As String, folded As String, startOfRest As Long, startOfNote As Long
    If remaining Is Nothing Then
        Set remaining = CreateObject("VBScript.RegExp")
        remaining.Global = True
        remaining.Pattern = "(^|[.!?;\r\n]\s*)((?:geri\s+)?kalan|diger)\s+(?:acik\s+)?siparis(?:ler[a-z]*)?\b"
        Set delivered = CreateObject("VBScript.RegExp")
        delivered.Pattern = "\bteslim\s+edil(?:mistir|di)\b"
        Set orderNumber = CreateObject("VBScript.RegExp")
        orderNumber.Global = True
        orderNumber.Pattern = "\b[0-9]{6,16}\b"
        Set relativeTerm = CreateObject("VBScript.RegExp")
        relativeTerm.Pattern = "\b\d{1,2}\s*(?:[-–]\s*\d{1,2}\s*)?hafta[a-z]*\b|\bisleme\s+alindigi\s+gunden\b"
        Set orderHeading = CreateObject("VBScript.RegExp")
        orderHeading.Pattern = "^\d{6,16}\s+(?:no(?:lu)?|numarali)\s+(?:acik\s+)?siparis[a-z]*\s*[,:\-]?\s*"
    End If
    folded = FoldTurkish(reply)
    Set sections = remaining.Execute(folded)
    If sections.Count <> 1 Then Exit Function
    Set marker = sections(0)
    startOfRest = marker.FirstIndex + Len(CStr(marker.SubMatches(0)))
    firstPart = Trim$(Left$(reply, startOfRest))
    Set identifiers = orderNumber.Execute(firstPart)
    If identifiers.Count <> 1 Then Exit Function
    If Not delivered.Test(FoldTurkish(firstPart)) Then Exit Function
    Set pairs = ReadOrderItemPairs(reply)
    If pairs.Count > 0 Or HasItemLabel(reply) Then Exit Function

    namedOrder = NormalizeOrderNo(CStr(identifiers(0).Value))
    startOfNote = identifiers(0).FirstIndex + 1
    deliveredNote = Mid$(firstPart, startOfNote)
    Set sections = orderHeading.Execute(FoldTurkish(deliveredNote))
    If sections.Count = 1 Then deliveredNote = Mid$(deliveredNote, sections(0).Length + 1)
    deliveredNote = Trim$(deliveredNote)
    remainingNote = Trim$(Mid$(reply, startOfRest + 1))
    If Len(deliveredNote) = 0 Or orderNumber.Test(remainingNote) Then Exit Function
    notesOnly = relativeTerm.Test(FoldTurkish(remainingNote))
    If Not notesOnly And Not IsUsefulReply(remainingNote) Then Exit Function
    remainingNote = Replace(Replace(remainingNote, vbCr, ""), vbLf, " ")
    Do While InStr(remainingNote, "  ") > 0
        remainingNote = Replace(remainingNote, "  ", " ")
    Loop
    SplitOrderSummary = True
End Function

Private Function QuotedOrderSummaryPairs(ByVal quoted As String) As Object
    Dim columns As Object, listed As Object, found As Object, line As Variant, pair As Variant
    Dim expanded As String, tableActive As Boolean, tableFinished As Boolean
    If Len(quoted) = 0 Or Len(quoted) >= 24000 Then Exit Function
    Set columns = CreateObject("Scripting.Dictionary")
    Set listed = CreateObject("Scripting.Dictionary")
    For Each line In Split(Replace(Replace(quoted, vbCrLf, vbLf), vbCr, vbLf), vbLf)
        If CStr(line) = TABLE_END Then
            If tableActive Then tableFinished = True
            tableActive = False
            GoTo NextQuotedSummaryLine
        End If
        expanded = ExpandTableMatchLine(CStr(line), columns)
        If expanded = "[TABLE_HEADER]" Then
            If columns.Exists("Order") And columns.Exists("Item") Then
                If tableActive Or tableFinished Then Exit Function
                tableActive = True
            ElseIf tableActive Then
                tableActive = False
                tableFinished = True
            End If
        ElseIf tableActive And InStr(CStr(line), vbTab) > 0 Then
            Set found = ReadOrderItemPairs(expanded)
            If found.Count <> 1 Then Exit Function
            For Each pair In found.Keys
                listed(pair) = True
            Next pair
        End If
NextQuotedSummaryLine:
    Next line
    If listed.Count >= 2 Then Set QuotedOrderSummaryPairs = listed
End Function

Private Function VerifiedOrderSummaryTargets(ByVal cfg As Object, ByVal currentSheet As Worksheet, _
    ByVal vendor As String, ByVal pairs As Object) As Object
    Dim sheets As Collection, ws As Worksheet, rows As Collection, orders As Object
    Dim targets As Object, details As Variant, values As Variant, row As Variant, pair As String
    If cfg.Exists("OrderSheets") Then
        Set sheets = cfg("OrderSheets")
    Else
        Set sheets = New Collection
        sheets.Add currentSheet
    End If
    Set targets = CreateObject("Scripting.Dictionary")
    For Each ws In sheets
        LoadSheetRowIndex cfg, ws, vendor, rows, orders, rowDetails:=details
        If ws Is currentSheet Then
            If rows.Count > 0 Then gMailHasVendorRows = True
        End If
        For Each row In rows
            values = details(CStr(row))
            pair = OrderItemKey(CStr(values(0)), CStr(values(1)))
            If pairs.Exists(pair) Then
                If targets.Exists(pair) Then Exit Function
                targets.Add pair, Array(ws.CodeName, CStr(row))
            End If
        Next row
    Next ws
    If targets.Count = pairs.Count Then Set VerifiedOrderSummaryTargets = targets
End Function

Private Function TryMatchStockItems(ByRef cfg As Object, ByVal ws As Worksheet, _
    ByVal vendor As String, ByVal reply As String, ByVal subject As String, _
    ByVal quoted As String, ByVal result As Object, ByVal mailDate As Date) As Boolean
    Static plural As Object, partial As Object, countPhrase As Object
    Dim body As String, stockDate As Variant, stockState As Long, orderRows As Collection
    Dim allOrders As Object, candidates As Object, candidate As Variant, orderNo As String
    Dim matchedOrder As String, itemNo As String, items As Object, subjectNamesOrder As Boolean
    Dim reason As String, scope As String
    If Not IsStockAvailability(reply) Then Exit Function
    If result.Count > 0 Then Exit Function
    If ExplicitItemList(reply).Count > 1 Then Exit Function
    body = SummarizeReply(reply)
    If Not IsStockAvailability(body) Or HasItemLabel(subject) Or HasOrderItemPair(subject) Then Exit Function
    If plural Is Nothing Then
        Set plural = CreateObject("VBScript.RegExp")
        plural.Pattern = "\b(urunler[a-z]*|malzemeler[a-z]*|her iki|iki kalem|2 kalem|tum kalemler|butun kalemler)\b"
        Set partial = CreateObject("VBScript.RegExp")
        partial.Pattern = "\b(bazi|bazilari|kismi|bir kismi|yalniz[a-z]*|sadece|haric|disinda|digerleri|birisi|biri)\b"
        Set countPhrase = CreateObject("VBScript.RegExp")
        countPhrase.Global = True
        countPhrase.Pattern = "\b(her\s+iki|iki|2)\s+kalem\b(?!\s*(no|numara|[0-9]))"
    End If
    If partial.Test(FoldTurkish(NormalizeName(reply))) Then Exit Function
    If Not plural.Test(FoldTurkish(NormalizeName(reply))) Then Exit Function
    If ReplyHasExplicitCodes(countPhrase.Replace(FoldTurkish(reply), "")) Then Exit Function
    If mailDate = 0 Then mailDate = Date
    stockDate = ResolveDeliveryDate(body, mailDate, stockState, stockOnly:=True)
    If IsEmpty(stockDate) Or stockState <> TERM_CONFIRMED Then Exit Function

    LoadSheetRowIndex cfg, ws, vendor, orderRows, allOrders
    subjectNamesOrder = SubjectHasOrderReference(subject, allOrders)
    If CountKnownOrders(subject, allOrders) > 1 Or CountKnownOrders(quoted, allOrders) > 1 Then Exit Function
    Set candidates = SelectReplyRows(cfg, ws, vendor, quoted, subject, reason, True)
    If candidates.Count = 0 Then
        If HasItemLabel(quoted) Or HasOrderItemPair(quoted) Or InStr(quoted, vbTab) > 0 Then Exit Function
        If subjectNamesOrder Then
            scope = subject
        ElseIf CountKnownOrders(quoted, allOrders) = 1 Then
            scope = quoted
        Else
            Exit Function
        End If
        For Each candidate In orderRows
            orderNo = SafeStr(ws.Cells(CLng(candidate), CLng(cfg("OrderNoCol"))).Value)
            If ContainsCode(scope, orderNo) Then candidates(CStr(candidate)) = True
        Next candidate
    End If
    If candidates.Count <> 2 Then Exit Function
    Set items = CreateObject("Scripting.Dictionary")
    For Each candidate In candidates.Keys
        orderNo = NormalizeOrderNo(SafeStr(ws.Cells(CLng(candidate), CLng(cfg("OrderNoCol"))).Value))
        itemNo = NormalizeOrderNo(SafeStr(ws.Cells(CLng(candidate), CLng(cfg("KalemNoCol"))).Value))
        If Len(orderNo) = 0 Or Len(itemNo) = 0 Then Exit Function
        If subjectNamesOrder And Not ContainsCode(subject, orderNo) Then Exit Function
        If HasKnownOrder(quoted, allOrders) And Not ContainsCode(quoted, orderNo) Then Exit Function
        If Len(matchedOrder) = 0 Then matchedOrder = orderNo
        If matchedOrder <> orderNo Then Exit Function
        items(itemNo) = True
    Next candidate
    If items.Count <> candidates.Count Then Exit Function
    For Each candidate In candidates.Keys
        result.Add CStr(candidate), reply
    Next candidate
    TryMatchStockItems = True
End Function

Private Function ExplicitItemList(ByVal reply As String) As Object
    Static leading As Object, trailing As Object, number As Object
    Dim matches As Object, item As Object, found As Object, norm As String
    Set found = CreateObject("Scripting.Dictionary")
    Set ExplicitItemList = found
    norm = FoldTurkish(reply)
    If leading Is Nothing Then
        Set leading = CreateObject("VBScript.RegExp")
        leading.Global = True
        leading.Pattern = "\bkalem(ler)?( no(lari)?| numara(lari)?)?[ \t]*[:#]?[ \t]*" & _
            "(\d{1,5}(?![\d./-])([ \t]*(ve|ile|,|;)[ \t]*\d{1,5}(?![\d./-]))+)\b"
        Set trailing = CreateObject("VBScript.RegExp")
        trailing.Global = True
        trailing.Pattern = "\b(\d{1,5}(?![\d./-])([ \t]*(ve|ile|,|;)[ \t]*\d{1,5}(?![\d./-]))+)" & _
            "[ \t]*(no lu|nolu|numarali)[ \t]*kalem(ler)?\b"
        Set number = CreateObject("VBScript.RegExp")
        number.Global = True
        number.Pattern = "\d{1,5}"
    End If
    Set matches = leading.Execute(norm)
    If matches.Count > 1 Then Exit Function
    If matches.Count = 1 Then
        norm = CStr(matches(0).SubMatches(4))
    Else
        Set matches = trailing.Execute(norm)
        If matches.Count <> 1 Then Exit Function
        norm = CStr(matches(0).SubMatches(0))
    End If
    For Each item In number.Execute(norm)
        found(NormalizeOrderNo(CStr(item.Value))) = True
    Next item
    If found.Count < 2 Then found.RemoveAll
End Function

Private Function ReplyHasExplicitCodes(ByVal text As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(^|[^a-z0-9])([a-z]+[-_/]?[0-9]{3,}|[0-9]{6,})([^a-z0-9]|$)"
    ReplyHasExplicitCodes = HasOrderLabel(text) Or HasItemLabel(text) Or re.Test(text)
End Function

Private Function ExplicitWholeOrder(ByVal text As String) As Boolean
    Static whole As Object, partial As Object
    Dim norm As String
    If whole Is Nothing Then
        Set whole = CreateObject("VBScript.RegExp")
        whole.Pattern = "\b((tum|butun)\s+kalemler(i|in|e|de)?|siparisin\s+tamami)\b"
        Set partial = CreateObject("VBScript.RegExp")
        partial.Pattern = "\b((tum|butun)\s+kalemler(i|in|e|de)?\s+degil|siparisin\s+tamami\s+degil|" & _
            "haric(inde)?|sadece|yalniz[a-z]*|bir\s+kismi|bazilari|bazi\s+kalemler)\b"
    End If
    norm = FoldTurkish(NormalizeName(text))
    ExplicitWholeOrder = whole.Test(norm) And Not partial.Test(norm)
End Function

Private Function FirstQuotedMessage(ByVal body As String) As String
    Dim lines As Variant, part As Variant, line As String, label As String
    Dim started As Boolean, result As String, primaryHeader As Boolean
    Dim quoteDepth As Long, firstQuoteDepth As Long
    lines = Split(Replace(Replace(StripHtmlArtifacts(body), vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For Each part In lines
        line = Trim$(CStr(part)): primaryHeader = False
        If Left$(line, 1) = ">" Then
            quoteDepth = 0
            Do While Left$(line, 1) = ">"
                quoteDepth = quoteDepth + 1
                line = Trim$(Mid$(line, 2))
            Loop
            If firstQuoteDepth = 0 Then
                If quoteDepth > 1 Then Exit For
                firstQuoteDepth = quoteDepth
            End If
            If quoteDepth > firstQuoteDepth Then Exit For
            started = True
        End If
        If InStr(line, ":") > 0 Then
            label = NormalizeName(Left$(line, InStr(line, ":") - 1))
            primaryHeader = (label = "FROM" Or label = "KIMDEN" Or label = "GONDEREN" Or _
                label = "VON" Or label = "DE" Or label = "EXPEDITEUR")
        End If
        primaryHeader = primaryHeader Or IsQuotedBoundary(line)
        If primaryHeader Then
            If started And Len(result) > 0 Then Exit For
            started = True
            GoTo NextQuotedLine
        End If
        If started And Len(line) > 0 Then
            result = result & line & vbCrLf
            If Len(result) > 24000 Then Exit For
        End If
NextQuotedLine:
    Next part
    FirstQuotedMessage = Left$(result, 24000)
End Function

Private Function SafeConversationId(ByVal mail As Object) As String
    On Error Resume Next
    SafeConversationId = SafeStr(mail.ConversationID)
    On Error GoTo 0
End Function

' Read-only traversal; all folders are collected before any Excel rows are changed.
Private Function CollectDeepMails(ByVal root As Object) As Collection
    Dim objects() As Object, dates() As Date, count As Long, capacity As Long
    Dim result As Collection, visited As Object, i As Long
    Set result = New Collection
    Set visited = CreateObject("Scripting.Dictionary")
    capacity = 256: ReDim objects(1 To capacity): ReDim dates(1 To capacity)
    CollectFolder root, objects, dates, count, capacity, visited, 0
    If count > 1 Then SortMailObjects objects, dates, 1, count
    For i = 1 To count: result.Add objects(i): Next i
    WriteLog "DERIN_TARAMA", root.Name, "", "", "Klasor=" & visited.Count & " | OutlookOgesi=" & count
    Set CollectDeepMails = result
End Function

Private Sub CollectFolder(ByVal folder As Object, ByRef objects() As Object, ByRef dates() As Date, _
    ByRef count As Long, ByRef capacity As Long, ByVal visited As Object, ByVal depth As Long)
    Dim item As Object, child As Object, identity As String, timestamp As Date
    Dim message As String
    On Error GoTo Failed
    If depth > 64 Then Err.Raise vbObjectError + 630, , "Klasor derinligi 64 sinirini asti."
    identity = CStr(folder.StoreID) & ":" & CStr(folder.EntryID)
    If visited.Exists(identity) Then Exit Sub
    visited.Add identity, True
    Dim recentItems As Object
    Set recentItems = RecentMailItems(folder, gScanCutoff, "ReceivedTime")
    For Each item In recentItems
        timestamp = SafeReceivedTime(item)
        If timestamp < gScanCutoff Then GoTo NextFolderItem
        count = count + 1
        If count > capacity Then
            capacity = capacity + 256
            ReDim Preserve objects(1 To capacity): ReDim Preserve dates(1 To capacity)
        End If
        Set objects(count) = item: dates(count) = timestamp
NextFolderItem:
    Next item
    For Each child In folder.Folders
        CollectFolder child, objects, dates, count, capacity, visited, depth + 1
    Next child
    Exit Sub
Failed:
    message = Err.Description
    Err.Raise vbObjectError + 631, "CollectFolder", "Alt klasor taramasi tamamlanamadi: " & message
End Sub

Private Function SafeReceivedTime(ByVal item As Object) As Date
    On Error Resume Next
    SafeReceivedTime = item.ReceivedTime
    On Error GoTo 0
End Function

Private Sub SortMailObjects(ByRef objects() As Object, ByRef dates() As Date, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Date, tmpDate As Date, tmp As Object
    i = lo: j = hi: pivot = dates((lo + hi) \ 2)
    Do While i <= j
        Do While dates(i) > pivot: i = i + 1: Loop
        Do While dates(j) < pivot: j = j - 1: Loop
        If i <= j Then
            tmpDate = dates(i): dates(i) = dates(j): dates(j) = tmpDate
            Set tmp = objects(i): Set objects(i) = objects(j): Set objects(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortMailObjects objects, dates, lo, j
    If i < hi Then SortMailObjects objects, dates, i, hi
End Sub

Private Function ItemSearchText(ByVal text As String) As String
    Static dates As Object
    If dates Is Nothing Then
        Set dates = CreateObject("VBScript.RegExp")
        dates.Global = True
        dates.Pattern = "\b[0-9]{1,4}[./-][0-9]{1,2}([./-][0-9]{2,4})?\b"
    End If
    text = dates.Replace(text, " DATEVALUE ")
    text = Replace(Replace(text, "(", " "), ")", " ")
    text = Replace(Replace(text, "[", " "), "]", " ")
    ItemSearchText = NormalizeName(text)
End Function

' True only for operational content; signatures, acknowledgements and code-only rows do not hide older answers.
Private Function IsUsefulReply(ByVal text As String) As Boolean
    IsUsefulReply = (Len(SummarizeReply(text)) > 0)
End Function

' Read sent items from the same mailbox as the selected Inbox. No mail is sent by this module.
Private Function ReadSentOrderLinks(ByRef cfg As Object, ByVal inbox As Object, _
    ByVal vendorEmails As Object, ByVal vendorDomains As Object, ByVal activeVendors As Object, _
    ByVal cutoff As Date, ByVal threads As Object) As Object
    Dim sent As Object, items As Object, mail As Object, recipient As Object, ae As Object
    Dim vendors As Object, latest As Object, vendor As Variant, matches As Object, key As Variant, targets As Object
    Dim ws As Worksheet, sheets As Collection, i As Long, r As Long
    Dim sentAt As Date, subject As String, body As String, address As String, reason As String
    Dim rowKey As String, orderNo As String, threadId As String, failure As String
    Dim mailText As String, indexedRows As Collection, indexedOrders As Object, rowDetails As Variant
    Dim recipientVendor As String
    Set latest = CreateObject("Scripting.Dictionary")
    Set ReadSentOrderLinks = latest
    On Error GoTo Failed
    Set sent = inbox.Store.GetDefaultFolder(5)
    Set items = RecentMailItems(sent, cutoff, "SentOn")
    Set sheets = cfg("OrderSheets")
    For Each mail In items
        If Not ReadSentFields(mail, sentAt, subject) Then GoTo NextSent
        If sentAt < cutoff Then GoTo NextSent
        Set vendors = CreateObject("Scripting.Dictionary")
        For Each recipient In mail.Recipients
            address = "": Set ae = Nothing
            On Error Resume Next
            Set ae = recipient.AddressEntry
            If Not ae Is Nothing Then
                If ae.Type = "EX" Then address = ae.GetExchangeUser.PrimarySmtpAddress
                If Len(address) = 0 Then address = ae.Address
            End If
            address = NormalizeEmail(address)
            If Len(address) = 0 Then address = NormalizeEmail(SafeStr(recipient.PropertyAccessor.GetProperty( _
                "http://schemas.microsoft.com/mapi/proptags/0x39FE001E")))
            If Len(address) = 0 Then address = NormalizeEmail(SafeStr(recipient.Address))
            On Error GoTo Failed
            recipientVendor = FindSentRecipientVendor(address, vendorEmails, vendorDomains)
            If Len(recipientVendor) > 0 Then
                If activeVendors.Exists(NormalizeName(recipientVendor)) Then vendors(recipientVendor) = True
            End If
        Next recipient
        If vendors.Count = 0 Then GoTo NextSent
        If Not ReadMessageBody(mail, body) Then GoTo NextSent
        body = NormalizeSasReferences(ExtractNewReply(body))
        subject = NormalizeSasReferences(subject)
        mailText = subject & vbCrLf & body
        threadId = SafeConversationId(mail)
        For Each vendor In vendors.Keys
            Set targets = CreateObject("Scripting.Dictionary")
            For i = 1 To sheets.Count
                Set ws = sheets(i)
                Set matches = SelectReplyRows(cfg, ws, CStr(vendor), body, subject, reason)
                For Each key In matches.Keys
                    r = CLng(key): rowKey = ws.CodeName & ":" & CStr(r)
                    orderNo = SafeStr(ws.Cells(r, CLng(cfg("OrderNoCol"))).Value)
                    ' Require an explicit order in the new sent text or subject, never quoted history.
                    If ContainsCode(mailText, orderNo) Then
                        If Not latest.Exists(rowKey) Then latest.Add rowKey, sentAt
                        targets(rowKey) = Array(ws.CodeName, r)
                    End If
                Next key
                LoadSheetRowIndex cfg, ws, CStr(vendor), indexedRows, indexedOrders, rowDetails:=rowDetails
                RecordOrderOnlySends ws, CStr(vendor), mailText, sentAt, indexedRows, rowDetails, matches
            Next i
            If Len(threadId) > 0 And (targets.Count > 0 Or IsSiparisKonusu(subject) Or ExtractNumbers(subject).Count > 0) Then
                If targets.Count = 1 And HasSeveralSentOrderRows(body) Then targets.RemoveAll
                IndexSentThread threads, SentThreadKey(CStr(vendor), threadId), sentAt, targets
            End If
        Next vendor
NextSent:
    Next mail
    Exit Function
Failed:
    If Err.Number = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    failure = Err.Description
    gSentScanFailed = True
    threads.RemoveAll
    latest.RemoveAll
    gOrderOnlySends.RemoveAll
    WriteLog "GONDERILENLER_OKUNAMADI", "", "", "", failure
End Function

Private Function FindSentRecipientVendor(ByVal address As String, ByVal emails As Object, ByVal domains As Object) As String
    Dim domain As String, name As String, delimiter As Long
    If emails.Exists(address) Then
        name = CStr(emails(address))
    Else
        delimiter = InStr(address, "@")
        If delimiter <= 1 Then Exit Function
        domain = Mid$(address, delimiter + 1)
        If IsPublicMailDomain(domain) Then Exit Function
        If domains.Exists(domain) Then name = CStr(domains(domain))
    End If
    If name <> "***CAKISMA***" Then FindSentRecipientVendor = name
End Function

Private Function TrackingOrderKey(ByVal ws As Worksheet, ByVal vendor As String, ByVal orderNo As String) As String
    vendor = NormalizeName(vendor)
    TrackingOrderKey = ws.CodeName & ":" & Len(vendor) & ":" & vendor & ":" & NormalizeOrderNo(orderNo)
End Function

Private Sub RecordOrderOnlySends(ByVal ws As Worksheet, ByVal vendor As String, ByVal mailText As String, _
    ByVal sentAt As Date, ByVal rows As Collection, ByVal rowDetails As Object, ByVal matches As Object)
    Dim matchedOrders As Object, checkedOrders As Object, row As Variant, data As Variant, orderNo As String, key As String
    If gOrderOnlySends Is Nothing Then Exit Sub
    Set matchedOrders = CreateObject("Scripting.Dictionary")
    Set checkedOrders = CreateObject("Scripting.Dictionary")
    For Each row In matches.Keys
        data = rowDetails(CStr(row))
        matchedOrders(NormalizeOrderNo(CStr(data(0)))) = True
    Next row
    For Each row In rows
        data = rowDetails(CStr(row))
        orderNo = NormalizeOrderNo(CStr(data(0)))
        If Not matchedOrders.Exists(orderNo) And Not checkedOrders.Exists(orderNo) Then
            checkedOrders(orderNo) = True
            If ContainsCode(mailText, CStr(data(0))) Then
                key = TrackingOrderKey(ws, vendor, orderNo)
                If Not gOrderOnlySends.Exists(key) Then gOrderOnlySends.Add key, sentAt
            End If
        End If
    Next row
End Sub

Private Sub ApplySentTracking(ByRef cfg As Object, ByVal latest As Object, ByVal testMode As Boolean, _
    Optional ByVal selectedRows As Object = Nothing)
    Dim ws As Worksheet, sheets As Collection, i As Long, r As Long, lastRow As Long, n As Long
    Dim rowKey As String, dateCol As Long, infoCol As Long, replyAt As Date, sentAt As Date
    Dim statusCol As Long, statusText As String, failure As String
    Dim orderNo As String, vendor As String, replyReason As String, written As Long
    Dim scannedSentAt As Date
    On Error GoTo Failed
    Set sheets = cfg("OrderSheets")
    statusCol = CLng(cfg("StatusCol"))
    For i = 1 To sheets.Count
        Set ws = sheets(i)
        dateCol = EnsureOrCreateColumn(ws, "Son Mail Gonderim Tarihi")
        infoCol = EnsureOrCreateColumn(ws, "Mail Takip Aciklamasi")
        lastRow = ws.Cells(ws.Rows.Count, CLng(cfg("OrderNoCol"))).End(xlUp).Row
        n = 0
        For r = CLng(cfg("StartRow")) To lastRow
            If testMode And n >= TEST_MAX_ORDERS Then Exit For
            rowKey = ws.CodeName & ":" & CStr(r)
            If Not selectedRows Is Nothing Then
                If Not selectedRows.Exists(rowKey) Then GoTo NextTrackingRow
            End If
            orderNo = SafeStr(ws.Cells(r, CLng(cfg("OrderNoCol"))).Value)
            If Len(orderNo) = 0 Then GoTo NextTrackingRow
            vendor = SafeStr(ws.Cells(r, CLng(cfg("VendorCol"))).Value)
            sentAt = 0: replyAt = 0
            If IsDate(ws.Cells(r, dateCol).Value) Then sentAt = CDate(ws.Cells(r, dateCol).Value)
            If latest.Exists(rowKey) Then
                scannedSentAt = CDate(latest(rowKey))
                If scannedSentAt > sentAt Then
                    sentAt = scannedSentAt
                    If Not ws.Cells(r, dateCol).HasFormula And Not ws.Cells(r, dateCol).MergeCells Then
                        ws.Cells(r, dateCol).Value = sentAt
                        ws.Cells(r, dateCol).NumberFormat = "dd.mm.yyyy hh:mm"
                    End If
                End If
            End If
            If Len(SafeStr(ws.Cells(r, CLng(cfg("ReplyCol"))).Value)) > 0 Then
                If IsDate(ws.Cells(r, CLng(cfg("ReplyDateCol"))).Value) Then _
                    replyAt = CDate(ws.Cells(r, CLng(cfg("ReplyDateCol"))).Value)
            End If
            replyReason = ReplyTrackingReason(rowKey, vendor, sentAt, replyAt)
            If sentAt > 0 Then
                MarkMailSent ws, r, statusCol
                statusText = FormatSentFollowUp(sentAt, replyAt)
                If Len(replyReason) > 0 Then statusText = statusText & " | " & replyReason
            Else
                statusText = UnverifiedSentFollowUp(ws, cfg, r, orderNo, vendor, replyAt, replyReason)
            End If
            If Not ws.Cells(r, infoCol).HasFormula And Not ws.Cells(r, infoCol).MergeCells Then
                If SafeStr(ws.Cells(r, infoCol).Value) <> statusText Then
                    ws.Cells(r, infoCol).Value2 = statusText
                    written = written + 1
                End If
            End If
NextTrackingRow:
            n = n + 1
        Next r
    Next i
    WriteLog "GONDERIM_TAKIP", "", "", "", "Eslesen satir=" & latest.Count, "Takip aciklamasi yazilan satir=" & written
    Exit Sub
Failed:
    failure = Err.Description
    WriteLog "GONDERIM_TAKIP_HATA", "", "", "", failure
    MsgBox T("Gelen yan~itlar okundu; g~onderilmi~s mail takibi tamamlanamad~i: ") & failure, vbExclamation, AppTitle()
End Sub

Private Function ReplyTrackingReason(ByVal rowKey As String, ByVal vendor As String, ByVal sentAt As Date, ByVal replyAt As Date) As String
    Dim note As Variant
    If gReplyTrackingReasons.Exists(rowKey) Then
        note = gReplyTrackingReasons(rowKey)
        If CDate(note(0)) >= sentAt And CDate(note(0)) > replyAt Then
            ReplyTrackingReason = CStr(note(1))
            Exit Function
        End If
    End If
    If replyAt <= 0 And gUnmatchedVendorReplies.Exists(NormalizeName(vendor)) Then
        ReplyTrackingReason = T("Tedarik~ciden gelen e-postalar bu kalemle e~sle~stirilemedi.")
    End If
End Function

Private Function UnverifiedSentFollowUp(ByVal ws As Worksheet, ByVal cfg As Object, ByVal row As Long, _
    ByVal orderNo As String, ByVal vendor As String, ByVal replyAt As Date, ByVal replyReason As String) As String
    Dim statusText As String, key As String, statusCol As Long
    key = TrackingOrderKey(ws, vendor, orderNo)
    statusCol = CLng(cfg("StatusCol"))
    If gSentScanFailed Then
        statusText = T("G~onderilen ~O~geler okunamad~i. Log sayfas~ina bak~in.")
    ElseIf Len(NormalizeName(vendor)) = 0 Then
        statusText = T("Tedarik~ci bilgisi bo~s.")
    ElseIf Not gRegisteredVendors.Exists(NormalizeName(vendor)) Then
        statusText = T("Tedarik~ci i~cin kay~itl~i e-posta adresi bulunamad~i.")
        If Not gVendorDirectoryNotes Is Nothing Then
            If gVendorDirectoryNotes.Exists(NormalizeName(vendor)) Then
                statusText = CStr(gVendorDirectoryNotes(NormalizeName(vendor)))
            End If
        End If
    ElseIf gOrderOnlySends.Exists(key) Then
        statusText = T("Sipari~se ait g~onderim: ") & Format$(CDate(gOrderOnlySends(key)), "dd.mm.yyyy hh:nn") & _
            T(" | Bu kalemin g~onderimi do~grulanamad~i.")
    Else
        statusText = T("Taranan aral~ikta bu kaleme ait g~onderim bulunamad~i.")
        If statusCol > 0 Then
            If Left$(SafeStr(ws.Cells(row, statusCol).Value), 1) = "+" Then
                statusText = T("Mevcut + i~sareti var, g~onderim tarihi do~grulanamad~i.")
            End If
        End If
    End If
    If replyAt > 0 Then
        statusText = statusText & T(" | Kay~itl~i yan~it: ") & Format$(replyAt, "dd.mm.yyyy hh:nn")
    ElseIf Len(replyReason) > 0 Then
        statusText = statusText & " | " & replyReason
    Else
        statusText = statusText & T(" | Bu kalem i~cin kay~itl~i yan~it yok.")
    End If
    UnverifiedSentFollowUp = statusText
End Function

Private Function FormatSentFollowUp(ByVal sentAt As Date, ByVal replyAt As Date) As String
    Dim description As String
    description = T("Mail at~ild~i: ") & Format$(sentAt, "dd.mm.yyyy hh:nn")
    If replyAt >= sentAt Then
        description = description & T(" | Geri d~on~u~s al~ind~i: ") & Format$(replyAt, "dd.mm.yyyy hh:nn")
    Else
        description = description & T(" | Geri d~on~u~s bekleniyor.")
    End If
    FormatSentFollowUp = description
End Function

Private Function HasSeveralSentOrderRows(ByVal body As String) As Boolean
    Dim line As Variant, count As Long
    For Each line In Split(Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf), vbLf)
        If StartsWithOrderNumber(Trim$(CStr(line))) Or HasOrderLabel(CStr(line)) Then
            count = count + 1
            If count > 1 Then
                HasSeveralSentOrderRows = True
                Exit Function
            End If
        End If
    Next line
End Function

Private Function SentThreadKey(ByVal vendor As String, ByVal threadId As String) As String
    vendor = NormalizeName(vendor)
    SentThreadKey = Len(vendor) & ":" & vendor & threadId
End Function

Private Sub IndexSentThread(ByVal threads As Object, ByVal key As String, ByVal sentAt As Date, ByVal targets As Object)
    Dim previous As Variant, entry As Variant, target As Variant, rowKey As Variant
    entry = Array(sentAt, "", 0&)
    If targets.Count = 1 Then
        For Each rowKey In targets.Keys
            target = targets(rowKey)
            entry = Array(sentAt, CStr(target(0)), CLng(target(1)))
        Next rowKey
    End If
    If threads.Exists(key) Then
        previous = threads(key)
        If CDate(previous(0)) > sentAt Then sentAt = CDate(previous(0))
        If CStr(previous(1)) <> CStr(entry(1)) Or CLng(previous(2)) <> CLng(entry(2)) Then
            entry = Array(sentAt, "", 0&)
        Else
            entry(0) = sentAt
        End If
    End If
    threads(key) = entry
End Sub

Private Function AppendVerifiedThreadMatch(ByRef cfg As Object, ByVal ws As Worksheet, ByVal vendor As String, _
    ByVal threadId As String, ByVal reply As String, ByVal subject As String, ByVal quoted As String, _
    ByVal receivedAt As Date, ByVal result As Object, ByVal reason As String) As Boolean
    Static partial As Object
    Dim key As String, entry As Variant, rows As Collection, allOrders As Object, row As Variant
    If Len(threadId) = 0 Or gSentThreadLinks Is Nothing Then Exit Function
    If InStr(1, reason, "CELISIYOR", vbTextCompare) > 0 Then Exit Function
    If InStr(1, reason, "KONU_SIPARISI_TEDARIKCIDE_YOK", vbTextCompare) > 0 Then Exit Function
    If ReplyHasExplicitCodes(reply) Or Not IsUsefulReply(reply) Then Exit Function
    key = SentThreadKey(vendor, threadId)
    If Not gSentThreadLinks.Exists(key) Then Exit Function
    entry = gSentThreadLinks(key)
    If CDate(entry(0)) >= receivedAt Or CStr(entry(1)) <> ws.CodeName Or CLng(entry(2)) < 1 Then Exit Function
    LoadSheetRowIndex cfg, ws, vendor, rows, allOrders
    If SubjectHasOrderReference(NormalizeSasReferences(subject), allOrders) Then Exit Function
    If HasKnownOrder(NormalizeSasReferences(quoted), allOrders) Or HasOrderLabel(quoted) Or HasItemLabel(quoted) Then Exit Function
    If partial Is Nothing Then
        Set partial = CreateObject("VBScript.RegExp")
        partial.Pattern = "\b(bazi|bazilari|kismi|haric(inde)?|sadece|yalniz[a-z]*|digerleri)\b"
    End If
    If partial.Test(FoldTurkish(NormalizeName(reply))) Then Exit Function
    For Each row In rows
        If CLng(row) = CLng(entry(2)) Then
            result.Add CStr(row), reply
            AppendVerifiedThreadMatch = True
            Exit Function
        End If
    Next row
End Function

Private Function ReadSentFields(ByVal mail As Object, ByRef sentAt As Date, _
    ByRef subject As String) As Boolean
    On Error GoTo Failed
    If mail.Class <> 43 Then Exit Function
    If Not mail.Sent Then Exit Function
    sentAt = mail.SentOn
    subject = SafeStr(mail.Subject)
    ReadSentFields = (sentAt > 0)
Failed:
End Function

' Clean without flattening table cells or removing order identifiers needed for matching.
Private Function CleanReplyContent(ByVal text As String) As String
    Dim lines As Variant, part As Variant, line As String, norm As String, result As String
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(8203), "")
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For Each part In lines
        line = Trim$(CStr(part))
        If Len(line) = 0 Then GoTo NextCleanLine
        norm = NormalizeName(line)
        If IsSignatureStart(line) Then
            If Len(result) = 0 And (norm = "IYI GUNLER" Or norm = "IYI CALISMALAR") Then GoTo NextCleanLine
            Exit For
        End If
        If IsMobileSignature(line) Or IsQuotedBoundary(line) Then Exit For
        If Left$(line, 1) = ">" Or Len(norm) = 0 Then GoTo NextCleanLine
        If InStr(line, vbTab) = 0 Then
            If IsReplyNoise(line) Or IsAdministrativeLine(norm) Then GoTo NextCleanLine
        End If
        ' Preserve repeated lines: identical statuses may belong to different orders.
        If Len(result) > 0 Then result = result & vbCrLf
        result = result & line
NextCleanLine:
    Next part
    CleanReplyContent = Left$(JoinWrappedStockReply(result), MAX_BODY_LEN)
End Function

Private Function JoinWrappedStockReply(ByVal body As String) As String
    Static target As Object, availability As Object
    If InStr(body, vbCrLf) = 0 Or InStr(1, body, "stok", vbTextCompare) = 0 Then
        JoinWrappedStockReply = body
        Exit Function
    End If
    If target Is Nothing Then
        Set target = CreateObject("VBScript.RegExp")
        target.Global = True: target.IgnoreCase = True
        target.Pattern = "(\b(urunler|urunlerimiz|malzemeler|malzemelerimiz|iki kalem|2 kalem|her iki urun)" & _
            "[ \t]*[,;:]?)[ ]*\r\n[ ]*(?=(\d{1,4}[./-]|yarin|bugun|[0-9]{1,2} ))"
        Set availability = CreateObject("VBScript.RegExp")
        availability.Global = True: availability.IgnoreCase = True
        availability.Pattern = "(\b(tarihinde|tarihi itibariyle|itibariyle|itibaren)|[0-9]{1,2}[./-][0-9]{1,2}([./-][0-9]{2,4})?)" & _
            "[ ]*\r\n[ ]*(?=stok)"
    End If
    ' Run the replacements on the original text, using folded positions only to preserve Turkce spelling.
    body = ReplaceStockLineBreaks(body, target)
    JoinWrappedStockReply = ReplaceStockLineBreaks(body, availability)
End Function

Private Function ReplaceStockLineBreaks(ByVal body As String, ByVal pattern As Object) As String
    Dim matches As Object, part As Object, i As Long, start As Long, width As Long
    Set matches = pattern.Execute(FoldTurkish(body))
    For i = matches.Count - 1 To 0 Step -1
        Set part = matches(i)
        start = part.FirstIndex + Len(CStr(part.SubMatches(0))) + 1
        width = part.Length - Len(CStr(part.SubMatches(0)))
        body = Left$(body, start - 1) & " " & Mid$(body, start + width)
    Next i
    ReplaceStockLineBreaks = body
End Function

Private Function IsAdministrativeLine(ByVal norm As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Pattern = "^(\w+\s+)?(fatura|e fatura|odeme|dekont|teklif|mutabakat|iban|vade|" & _
            "banka bilgileri|vergi levhasi|e arsiv|invoice|payment)\b"
    End If
    IsAdministrativeLine = re.Test(norm)
End Function

Private Function IsReplyNoise(ByVal line As String) As Boolean
    Static re As Object
    Dim norm As String
    norm = NormalizeName(line)
    If Len(norm) = 0 Or IsSignatureStart(line) Or IsMobileSignature(line) Then
        IsReplyNoise = True
        Exit Function
    End If
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Pattern = "^\W*(cid:|https?://|www\.|mailto:|\[?image[0-9]*\.(png|jpe?g|gif))|" & _
            "^\W*(adres|adresi|address|tel|telefon|mobil|gsm|mobile|phone|faks|fax|web|e[ -]?posta|e[ -]?mail)\s*:|" & _
            "^\W*(di[sk] (kaynakli|gonderici|e posta|mail)|external (email|e mail|sender|warning)|caution external|" & _
            "uyari bu|dikkat bu|bu e posta|bu elektronik posta|this e mail|this email|do not click|outlook for ios|" & _
            "outlook for android|sent from my|lutfen bu e postayi|bu mesaji yazdirmadan|gizlilik bildirimi|" & _
            "kisisel verilerin korunmasi|kvkk aydinlatma|confidentiality notice|disclaimer)( |$)|" & _
            "^(merhaba|merhabalar|selam|selamlar|gunaydin|iyi aksamlar|tesekkurler|tesekkur ederim|bilginize|" & _
            "thank you|thanks|hello|hi|ok|tamam)$"
    End If
    IsReplyNoise = re.Test(line) Or re.Test(norm)
End Function

' Apple Mail / Gmail / Outlook localized reply boundaries.
Private Function IsQuotedBoundary(ByVal line As String) As Boolean
    Dim norm As String, re As Object
    norm = LCase$(NormalizeName(line))
    If Trim$(line) = "[QUOTED_MESSAGE]" Then IsQuotedBoundary = True: Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^on .+ wrote\s*:?$|.+ tarihinde .+ yazdi\s*:?$|.+ sunlari yazdi(\s*\(.+\))?\s*:?$|.+ wrote\s*\(.+\)\s*:?$"
    If re.Test(norm) Then IsQuotedBoundary = True: Exit Function
    If norm = "ORIGINAL MESSAGE" Or norm = "ORIJINAL MESAJ" Or norm = "FORWARDED MESSAGE" Or _
       norm = "ILETILEN ILETI" Or norm = "ILETILEN MESAJ" Then IsQuotedBoundary = True: Exit Function
    If Left$(Trim$(line), 2) = "--" Then
        IsQuotedBoundary = (InStr(norm, "original message") > 0 Or _
            InStr(norm, "orijinal mesaj") > 0 Or InStr(norm, "forwarded message") > 0 Or _
            InStr(norm, "iletilen ileti") > 0)
    End If
End Function

Private Function IsMobileSignature(ByVal line As String) As Boolean
    Dim norm As String, re As Object
    norm = NormalizeName(line)
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(iphone|ipad|android|samsung).*(gonderildi|gonderilen|sent)|^sent from my (iphone|ipad|android)|^(ios|android) icin outlook"
    IsMobileSignature = re.Test(norm)
End Function

Private Function ConfigText(ByVal cfg As Object, ByVal key As String, ByVal fallback As String) As String
    ConfigText = fallback
    If cfg.Exists(key) Then
        If Len(SafeStr(cfg(key))) > 0 Then ConfigText = SafeStr(cfg(key))
    End If
End Function

Private Function ConfigLong(ByVal cfg As Object, ByVal key As String, ByVal fallback As Long) As Long
    ConfigLong = fallback
    If cfg.Exists(key) Then ConfigLong = ReadSettingLong(cfg(key), fallback, 1, 3650)
End Function

Private Function ReadSettingLong(ByVal value As Variant, ByVal fallback As Long, _
    ByVal minimum As Long, ByVal maximum As Long) As Long
    Dim number As Double
    ReadSettingLong = fallback
    If Not IsNumeric(value) Then Exit Function
    On Error GoTo Failed
    number = CDbl(value)
    If number < minimum Or number > maximum Or number <> Fix(number) Then Exit Function
    ReadSettingLong = CLng(number)
Failed:
End Function

Private Function ReadSettingYesNo(ByVal value As Variant, ByVal fallback As String) As String
    Dim normalized As String
    normalized = NormalizeName(SafeStr(value))
    Select Case normalized
        Case "EVET", "TRUE", "1": ReadSettingYesNo = "EVET"
        Case "HAYIR", "FALSE", "0": ReadSettingYesNo = "HAYIR"
        Case Else: ReadSettingYesNo = fallback
    End Select
End Function

' Restrict narrows Outlook enumeration; per-message dates are still checked by callers.
Private Function RecentMailItems(ByVal folder As Object, ByVal cutoff As Date, _
    ByVal dateField As String) As Object
    Dim items As Object, selected As Object, filter As String, failure As String
    If dateField <> "SentOn" And dateField <> "ReceivedTime" Then _
        Err.Raise vbObjectError + 802, , "Desteklenmeyen tarih alani."
    Set items = folder.Items
    filter = "[" & dateField & "] >= '" & Format$(cutoff, "ddddd h:nn AMPM") & "'"
    On Error Resume Next
    Set selected = items.Restrict(filter)
    failure = Err.Description
    On Error GoTo 0
    If selected Is Nothing Then
        Set selected = items
        WriteLog "TARIH_FILTRESI_YEDEK", folder.Name, "", "", failure
    End If
    selected.Sort "[" & dateField & "]", True
    Set RecentMailItems = selected
End Function

' The index lives for one run. Self-tests on another workbook always rebuild it.
Private Sub LoadSheetRowIndex(ByVal cfg As Object, ByVal ws As Worksheet, ByVal vendor As String, _
    ByRef vendorRows As Collection, ByRef allOrders As Object, Optional ByRef vendorIndex As Variant, _
    Optional ByRef rowDetails As Variant)
    Dim index As Object, vendors As Object, rows As Collection, details As Object
    Dim cacheKey As String, name As String, orderNo As String, r As Long, firstRow As Long, lastRow As Long, offset As Long
    Dim useCache As Boolean, rawOrder As String, description As String
    Dim orderValues As Variant, vendorValues As Variant, itemValues As Variant, materialValues As Variant, descriptions As Variant
    Dim descriptionCol As Long
    cacheKey = ws.CodeName
    useCache = (ws.Parent Is ThisWorkbook)
    Set vendorRows = New Collection
    If gRowIndex Is Nothing Then Set gRowIndex = CreateObject("Scripting.Dictionary")
    If useCache And gRowIndex.Exists(cacheKey) Then
        Set index = gRowIndex(cacheKey)
    Else
        Set index = CreateObject("Scripting.Dictionary")
        Set vendors = CreateObject("Scripting.Dictionary")
        vendors.CompareMode = vbTextCompare
        Set allOrders = CreateObject("Scripting.Dictionary")
        allOrders.CompareMode = vbTextCompare
        Set details = CreateObject("Scripting.Dictionary")
        firstRow = CLng(cfg("StartRow"))
        lastRow = ws.Cells(ws.Rows.Count, CLng(cfg("OrderNoCol"))).End(xlUp).Row
        If lastRow >= firstRow Then
            orderValues = OrderColumnValues(ws, firstRow, lastRow, CLng(cfg("OrderNoCol")))
            vendorValues = OrderColumnValues(ws, firstRow, lastRow, CLng(cfg("VendorCol")))
            itemValues = OrderColumnValues(ws, firstRow, lastRow, CLng(cfg("KalemNoCol")))
            materialValues = OrderColumnValues(ws, firstRow, lastRow, CLng(cfg("MalzemeCol")))
            If cfg.Exists("UzunMetinCol") Then descriptionCol = CLng(cfg("UzunMetinCol"))
            If descriptionCol > 0 Then descriptions = OrderColumnValues(ws, firstRow, lastRow, descriptionCol)
            For r = firstRow To lastRow
                offset = r - firstRow + 1
                rawOrder = SafeStr(orderValues(offset, 1))
                orderNo = NormalizeOrderNo(rawOrder)
                If Len(orderNo) > 0 Then
                    If Len(orderNo) >= 4 Then allOrders(orderNo) = True
                    name = NormalizeName(SafeStr(vendorValues(offset, 1)))
                    If Not vendors.Exists(name) Then
                        Set rows = New Collection
                        vendors.Add name, rows
                    End If
                    Set rows = vendors(name): rows.Add r
                    description = ""
                    If descriptionCol > 0 Then description = NormalizeName(SafeStr(descriptions(offset, 1)))
                    details.Add CStr(r), Array(rawOrder, SafeStr(itemValues(offset, 1)), _
                        SafeStr(materialValues(offset, 1)), description)
                End If
            Next r
        End If
        index.Add "RowDetails", details
        index.Add "Vendors", vendors
        index.Add "Orders", allOrders
        If useCache Then gRowIndex.Add cacheKey, index
    End If
    Set allOrders = index("Orders")
    Set vendors = index("Vendors")
    If Not IsMissing(vendorIndex) Then Set vendorIndex = vendors
    If Not IsMissing(rowDetails) Then Set rowDetails = index("RowDetails")
    name = NormalizeName(vendor)
    If vendors.Exists(name) Then Set vendorRows = vendors(name)
End Sub

Private Function OrderColumnValues(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
    ByVal col As Long) As Variant
    Dim singleValue(1 To 1, 1 To 1) As Variant
    If firstRow = lastRow Then
        singleValue(1, 1) = ws.Cells(firstRow, col).Value2
        OrderColumnValues = singleValue
    Else
        OrderColumnValues = ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col)).Value2
    End If
End Function

Private Function ActiveOrderVendors(ByVal cfg As Object, ByVal sheets As Collection) As Object
    Dim active As Object, sheet As Worksheet, rows As Collection, orders As Object, vendors As Variant, key As Variant
    Set active = CreateObject("Scripting.Dictionary")
    active.CompareMode = vbTextCompare
    For Each sheet In sheets
        LoadSheetRowIndex cfg, sheet, "", rows, orders, vendors
        For Each key In vendors.Keys
            If Len(CStr(key)) > 0 Then active(key) = True
        Next key
    Next sheet
    Set ActiveOrderVendors = active
End Function

Private Function ResolveVendorCollision(ByVal cfg As Object, ByVal sheets As Collection, _
    ByVal candidates As Object, ByVal reply As String, ByVal subject As String, _
    Optional ByVal quoted As String = "", Optional ByVal mailDate As Date = 0) As String
    Dim candidate As Variant, vendor As String, resolved As String, ws As Worksheet, matched As Object
    Dim key As Variant, code As String, reason As String
    subject = NormalizeSasReferences(subject)
    quoted = NormalizeSasReferences(quoted)
    For Each candidate In candidates.Items
        vendor = CStr(candidate)
        For Each ws In sheets
            Set matched = SelectDeepReplyRows(cfg, ws, vendor, reply, subject, quoted, reason, mailDate)
            For Each key In matched.Keys
                code = SafeStr(ws.Cells(CLng(key), CLng(cfg("OrderNoCol"))).Value)
                If Len(code) >= 6 And (ContainsCode(subject, code) Or ContainsCode(NormalizeSasReferences(reply), code) Or ContainsCode(quoted, code)) Then
                    If Len(resolved) > 0 Then
                        If StrComp(NormalizeName(resolved), NormalizeName(vendor), vbTextCompare) <> 0 Then Exit Function
                    End If
                    resolved = vendor
                    Exit For
                End If
            Next key
        Next ws
    Next candidate
    ResolveVendorCollision = resolved
End Function

' Use the actual table headings to keep quantity, dates and identifiers in their own fields.
Private Function ExpandTableMatchLine(ByVal line As String, ByRef columns As Object) As String
    Dim cells As Variant, candidate As Object, i As Long, heading As String, key As Variant
    Dim enriched As String, cell As String
    ExpandTableMatchLine = line
    If InStr(line, vbTab) = 0 Then Exit Function
    cells = Split(line, vbTab)
    Set candidate = CreateObject("Scripting.Dictionary")
    For i = LBound(cells) To UBound(cells)
        heading = CleanForCompare(CStr(cells(i)))
        Select Case heading
            Case "SIPARISNO", "SIPARISNUMARASI", "SIPARIS", "ORDERNO", "ORDERNUMBER", "PONO", "PO", "SATINALMASIPARISI", "SASNO", "SATINALMASIPARISNO"
                candidate("Order") = i
            Case "KALEM", "KALEMNO", "KALEMNUMARASI", "ITEM", "ITEMNO", "ITEMNUMBER", "POZISYON", "SATKALEMNO", "SATKALEM", "SASKALEMNO", _
                "SIPARISKALEMNO", "SATINALMAKALEMNO"
                candidate("Item") = i
            Case "MALZEME", "MALZEMENO", "MALZEMEKODU", "MATERIAL", "MATERIALNO", "URUNKODU"
                candidate("Material") = i
            Case "TERMIN", "TERMINTARIHI", "TESLIMTARIHI", "DELIVERYDATE"
                candidate("Delivery") = i
            Case "YENITERMIN", "YENITERMINTARIHI", "TEYITLITERMIN", "TEYITTERMINI", "ONAYLITERMIN", _
                "CONFIRMEDDELIVERYDATE", "PROMISEDDATE", "TEDARIKCITERMINI", "TEDARIKCITERMINTARIHI", _
                "TEYITLITESLIMTARIHI", "YENITESLIMTARIHI", "GUNCELTERMIN", "GUNCELTERMINTARIHI", _
                "TEYITTERMINTARIHI", "TEYITTARIHI", "REVIZETERMIN"
                candidate("ConfirmedDelivery") = i
            Case "STOKTARIHI", "STOKGIRISTARIHI", "STOKAGIRISTARIHI", "STOKTAOLACAGITARIH", "STOCKDATE"
                candidate("StockDate") = i
            Case "DURUM", "STATUS", "TEDARIKCIDURUMU", "TERMINDURUMU", "TESLIMATDURUMU"
                candidate("Status") = i
            Case "YANIT", "TEDARIKCIYANITI", "TEDARIKCICEVABI", "TEDARIKCIACIKLAMASI", "NOT", "NOTLAR", _
                "ACIKLAMA", "CEVAP", "REPLY", "COMMENTS", "COMMENT"
                candidate("Note") = i
            Case "TEYITLIMIKTAR", "ONAYLANANMIKTAR", "SEVKMIKTARI", "TESLIMMIKTARI", "CONFIRMEDQUANTITY"
                candidate("Quantity") = i
            Case "KARGOTAKIP", "KARGOTAKIPNO", "TAKIPNUMARASI", "TRACKINGNUMBER", "TRACKINGNO"
                candidate("Tracking") = i
        End Select
    Next i
    If candidate.Exists("Order") Or candidate.Exists("Item") Or candidate.Exists("Material") Then
        Set columns = candidate
        ExpandTableMatchLine = "[TABLE_HEADER]": Exit Function
    End If
    If columns.Count = 0 Then Exit Function
    For Each key In Array("Order", "Item", "Material")
        If Not columns.Exists(key) Then GoTo NextTableIdentifier
        i = CLng(columns(key))
        If i <= UBound(cells) Then
            cell = Trim$(CStr(cells(i)))
            If Len(cell) > 0 Then
                Select Case CStr(key)
                    Case "Order": enriched = enriched & vbCrLf & "siparis no " & cell
                    Case "Item": enriched = enriched & vbCrLf & "kalem no " & cell
                    Case "Material": enriched = enriched & vbCrLf & "malzeme " & cell
                End Select
            End If
        End If
NextTableIdentifier:
    Next key
    ExpandTableMatchLine = Mid$(enriched, 3)
End Function


Private Function ReadTableReply(ByVal line As String, ByVal columns As Object) As String
    Dim cells As Variant, field As Variant, position As Long, value As String
    cells = Split(line, vbTab)
    For Each field In Array("ConfirmedDelivery", "Delivery", "StockDate", "Status", "Note", "Quantity", "Tracking")
        If CStr(field) = "Delivery" And columns.Exists("ConfirmedDelivery") Then GoTo NextTableField
        If columns.Exists(CStr(field)) Then
            position = CLng(columns(CStr(field)))
            If position <= UBound(cells) Then
                value = Trim$(CStr(cells(position)))
                If Len(value) > 0 And value <> "-" Then
                    Select Case CleanForCompare(value)
                        Case "", "NA", "N/A", "NULL", "YOK", "BOS", "00000000": GoTo NextTableField
                    End Select
                    Select Case CStr(field)
                        Case "Delivery", "ConfirmedDelivery": value = "Termin: " & value
                        Case "StockDate": value = "Stok tarihi: " & value
                        Case "Quantity": value = "Miktar: " & value
                        Case "Tracking": value = "Kargo takip: " & value
                    End Select
                    If Len(ReadTableReply) > 0 Then ReadTableReply = ReadTableReply & vbCrLf
                    ReadTableReply = ReadTableReply & value
                End If
            End If
        End If
NextTableField:
    Next field
End Function

Private Function HasMaterialCodeLabel(ByVal text As String) As Boolean
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(^|[^a-z])(malzeme|material|urun)(\s*(no|kodu|code))?\s*[.:#-]*\s*([a-z0-9_-]*[0-9][a-z0-9_-]*)"
    Set matches = re.Execute(NormalizeName(text))
    If matches.Count > 0 Then HasMaterialCodeLabel = (Len(CStr(matches(0).SubMatches(4))) >= 5)
End Function

Private Function DecodeHtmlEntities(ByVal text As String) As String
    Dim re As Object, matches As Object, token As Object, i As Long, code As Long, value As String
    text = Replace(text, "&apos;", "'")
    text = Replace(text, "&rsquo;", "'")
    text = Replace(text, "&lsquo;", "'")
    text = Replace(text, "&ndash;", "-")
    text = Replace(text, "&mdash;", "-")
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True
    re.Pattern = "&#(x[0-9a-f]{1,4}|[0-9]{1,5});"
    Set matches = re.Execute(text)
    For i = matches.Count - 1 To 0 Step -1
        Set token = matches(i): value = CStr(token.SubMatches(0))
        If LCase$(Left$(value, 1)) = "x" Then code = CLng("&H" & Mid$(value, 2)) Else code = CLng(value)
        If code > 0 And code <= 65535 And (code < 55296 Or code > 57343) Then
            text = Left$(text, token.FirstIndex) & ChrW(code) & Mid$(text, token.FirstIndex + token.Length + 1)
        End If
    Next i
    DecodeHtmlEntities = text
End Function


Private Function ExtractEmailList(ByVal text As String) As Variant
    Dim re As Object, match As Object, addresses As Object, address As String
    Set addresses = CreateObject("Scripting.Dictionary")
    addresses.CompareMode = vbTextCompare
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True: re.Global = True
    re.Pattern = "[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+"
    For Each match In re.Execute(text)
        address = NormalizeEmail(CStr(match.Value))
        If Len(address) > 0 Then addresses(address) = True
    Next match
    ExtractEmailList = addresses.Keys
End Function
