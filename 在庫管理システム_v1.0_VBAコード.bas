Attribute VB_Name = "Module_Inventory"
Option Explicit

' ================================================================
' 在庫管理システム v1.0 - VBAモジュール
' ================================================================

' ── シート名定数 ─────────────────────────────────────────────────
Const SH_MASTER   As String = "商品マスター"
Const SH_INVENTORY As String = "在庫管理"
Const SH_PURCHASE  As String = "仕入れ管理"
Const SH_LOG       As String = "在庫異動ログ"
Const SH_ALERT     As String = "発注アラート"

' ── 列番号：商品マスター ─────────────────────────────────────────
Const PM_ID       As Integer = 1   ' A internal_id
Const PM_NAME     As Integer = 2   ' B product_name
Const PM_STATUS   As Integer = 15  ' O status
Const PM_SUPPLIER As Integer = 12  ' L supplier_name
Const PM_CREATED  As Integer = 17  ' Q created_at
Const PM_UPDATED  As Integer = 18  ' R updated_at

' ── 列番号：在庫管理 ─────────────────────────────────────────────
Const INV_ID       As Integer = 1   ' A inventory_id  (formula)
Const INV_INTID    As Integer = 2   ' B internal_id
Const INV_CHANNEL  As Integer = 3   ' C channel
Const INV_NAME     As Integer = 4   ' D product_name  (formula)
Const INV_STOCK    As Integer = 5   ' E stock_quantity
Const INV_RESERVED As Integer = 6   ' F reserved_quantity
Const INV_AVAIL    As Integer = 7   ' G available_quantity (formula)
Const INV_REORDER  As Integer = 8   ' H reorder_point
Const INV_RQ       As Integer = 9   ' I reorder_quantity
Const INV_LOC      As Integer = 10  ' J location
Const INV_STOCKED  As Integer = 11  ' K last_stocked_at
Const INV_SOLD     As Integer = 12  ' L last_sold_at
Const INV_UPDATED  As Integer = 13  ' M updated_at
Const INV_ALERT    As Integer = 14  ' N アラート (formula)

' ── 列番号：仕入れ管理 ──────────────────────────────────────────
Const PO_ID       As Integer = 1   ' A purchase_id
Const PO_INTID    As Integer = 2   ' B internal_id
Const PO_NAME     As Integer = 3   ' C product_name  (formula)
Const PO_ODATE    As Integer = 4   ' D order_date
Const PO_ADATE    As Integer = 5   ' E arrival_date
Const PO_SUPPLIER As Integer = 6   ' F supplier_name
Const PO_ORDNO    As Integer = 7   ' G supplier_order_no
Const PO_QTY_ORD  As Integer = 8   ' H quantity_ordered
Const PO_QTY_RCV  As Integer = 9   ' I quantity_received
Const PO_UCOST    As Integer = 10  ' J unit_cost
Const PO_TAX      As Integer = 11  ' K tax_rate
Const PO_TOTAL    As Integer = 12  ' L total_cost (formula: J*I)
Const PO_SHIP     As Integer = 13  ' M shipping_cost
Const PO_DEST     As Integer = 14  ' N destination
Const PO_STATUS   As Integer = 15  ' O status
Const PO_INVOICE  As Integer = 16  ' P invoice_no
Const PO_PAYDATE  As Integer = 17  ' Q payment_date
Const PO_NOTES    As Integer = 18  ' R notes
Const PO_REFLECTED As Integer = 19 ' S reflected_at (二重反映防止)

' ── 列番号：在庫異動ログ ─────────────────────────────────────────
Const LOG_ID      As Integer = 1   ' A log_id
Const LOG_DATE    As Integer = 2   ' B log_date
Const LOG_INTID   As Integer = 3   ' C internal_id
Const LOG_NAME    As Integer = 4   ' D product_name
Const LOG_CHANNEL As Integer = 5   ' E channel
Const LOG_TYPE    As Integer = 6   ' F transaction_type
Const LOG_BEFORE  As Integer = 7   ' G quantity_before
Const LOG_CHANGE  As Integer = 8   ' H quantity_change
Const LOG_AFTER   As Integer = 9   ' I quantity_after
Const LOG_REF     As Integer = 10  ' J reference_id
Const LOG_NOTES   As Integer = 11  ' K notes

' ================================================================
' ユーティリティ: 在庫異動ログへ1行追記
' ================================================================
Private Sub WriteLog(intId As String, prodName As String, ch As String, _
                     txType As String, before As Long, change As Long, _
                     after As Long, refId As String, note As String)
    Dim wsLog As Worksheet
    Dim logRow As Long

    ' ログシートの保護を一時解除
    Set wsLog = ThisWorkbook.Sheets(SH_LOG)
    wsLog.Unprotect Password:="log2024"

    logRow = wsLog.Cells(wsLog.Rows.Count, LOG_ID).End(xlUp).Row + 1

    With wsLog
        .Cells(logRow, LOG_ID).Value      = logRow - 1
        .Cells(logRow, LOG_DATE).Value    = Now()
        .Cells(logRow, LOG_INTID).Value   = intId
        .Cells(logRow, LOG_NAME).Value    = prodName
        .Cells(logRow, LOG_CHANNEL).Value = ch
        .Cells(logRow, LOG_TYPE).Value    = txType
        .Cells(logRow, LOG_BEFORE).Value  = before
        .Cells(logRow, LOG_CHANGE).Value  = change
        .Cells(logRow, LOG_AFTER).Value   = after
        .Cells(logRow, LOG_REF).Value     = refId
        .Cells(logRow, LOG_NOTES).Value   = note
    End With

    wsLog.Protect Password:="log2024"
End Sub

' ================================================================
' ① 商品登録マクロ
' 使い方：商品マスターで登録した行を選択して実行
' ================================================================
Sub RegisterProduct()
    Dim wsM As Worksheet, wsI As Worksheet
    Dim masterRow As Long
    Dim intId As String, prodName As String
    Dim ch As Variant, channels As Variant
    Dim lastInvRow As Long

    Set wsM = ThisWorkbook.Sheets(SH_MASTER)
    Set wsI = ThisWorkbook.Sheets(SH_INVENTORY)

    masterRow = ActiveCell.Row
    If masterRow <= 1 Then
        MsgBox "商品マスターで登録行を選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    intId    = Trim(wsM.Cells(masterRow, PM_ID).Value)
    prodName = Trim(wsM.Cells(masterRow, PM_NAME).Value)

    If intId = "" Then
        MsgBox "内部管理ID（A列）を入力してから実行してください。", vbExclamation
        Exit Sub
    End If

    ' 重複チェック
    Dim invRow As Long
    lastInvRow = wsI.Cells(wsI.Rows.Count, INV_INTID).End(xlUp).Row
    For invRow = 2 To lastInvRow
        If wsI.Cells(invRow, INV_INTID).Value = intId Then
            If MsgBox("[" & intId & "] の在庫行が既に存在します。追加しますか？", _
                      vbYesNo + vbExclamation) = vbNo Then Exit Sub
            Exit For
        End If
    Next invRow

    ' 登録日時を記録
    wsM.Cells(masterRow, PM_CREATED).Value = Now()
    wsM.Cells(masterRow, PM_UPDATED).Value = Now()
    If Trim(wsM.Cells(masterRow, PM_STATUS).Value) = "" Then
        wsM.Cells(masterRow, PM_STATUS).Value = "active"
    End If

    ' 在庫管理に3チャネル追加
    channels = Array("self", "fba", "rakuten")
    For Each ch In channels
        lastInvRow = wsI.Cells(wsI.Rows.Count, INV_INTID).End(xlUp).Row + 1
        wsI.Cells(lastInvRow, INV_INTID).Value   = intId
        wsI.Cells(lastInvRow, INV_CHANNEL).Value = ch
        wsI.Cells(lastInvRow, INV_STOCK).Value   = 0
        wsI.Cells(lastInvRow, INV_RESERVED).Value = 0
        wsI.Cells(lastInvRow, INV_UPDATED).Value = Now()
    Next ch

    MsgBox "登録完了: [" & intId & "] の在庫行を self / fba / rakuten の3チャネル分追加しました。", vbInformation
    ThisWorkbook.Sheets(SH_INVENTORY).Activate
End Sub

' ================================================================
' ② 仕入れ確定マクロ
' 使い方：仕入れ管理で対象行を選択して実行
' ================================================================
Sub ConfirmPurchase()
    Dim wsPO  As Worksheet, wsI As Worksheet
    Dim poRow As Long, invRow As Long
    Dim intId As String, dest As String
    Dim qtyRcv As Long, ucost As Long, purchId As String

    Set wsPO = ThisWorkbook.Sheets(SH_PURCHASE)
    Set wsI  = ThisWorkbook.Sheets(SH_INVENTORY)

    poRow = ActiveCell.Row
    If poRow <= 1 Then
        MsgBox "仕入れ管理で対象行を選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    ' 二重反映チェック
    If Trim(wsPO.Cells(poRow, PO_REFLECTED).Value) <> "" Then
        MsgBox "この行はすでに在庫反映済みです。" & Chr(10) & _
               "反映日時: " & wsPO.Cells(poRow, PO_REFLECTED).Value, vbExclamation
        Exit Sub
    End If

    intId   = Trim(wsPO.Cells(poRow, PO_INTID).Value)
    dest    = Trim(wsPO.Cells(poRow, PO_DEST).Value)
    qtyRcv  = Val(wsPO.Cells(poRow, PO_QTY_RCV).Value)
    ucost   = Val(wsPO.Cells(poRow, PO_UCOST).Value)
    purchId = Trim(wsPO.Cells(poRow, PO_ID).Value)

    If intId = ""   Then MsgBox "internal_id が空です。",           vbExclamation : Exit Sub
    If qtyRcv <= 0  Then MsgBox "入荷数（I列）が 0 以下です。",     vbExclamation : Exit Sub
    If dest = ""    Then MsgBox "入庫先（N列）が未設定です。",       vbExclamation : Exit Sub

    Dim prodName As String
    prodName = Trim(wsPO.Cells(poRow, PO_NAME).Value)

    ' 確認
    If MsgBox("【仕入れ確定】" & Chr(10) & _
              "商品ID: " & intId & Chr(10) & _
              "入荷数: " & qtyRcv & "　入庫先: " & dest & Chr(10) & Chr(10) & _
              "在庫を加算し、異動ログを記録します。よろしいですか？", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ' チャネル展開
    Dim channels() As String
    Select Case LCase(Trim(dest))
        Case "self":  channels = Split("self", ",")
        Case "fba":   channels = Split("fba", ",")
        Case "both":  channels = Split("self,fba", ",")
        Case Else:    channels = Split(dest, ",")
    End Select

    Dim ch As Variant
    Dim foundAny As Boolean
    foundAny = False
    Dim lastInvRow As Long
    lastInvRow = wsI.Cells(wsI.Rows.Count, INV_INTID).End(xlUp).Row

    For Each ch In channels
        ch = Trim(ch)
        For invRow = 2 To lastInvRow
            If wsI.Cells(invRow, INV_INTID).Value = intId And _
               wsI.Cells(invRow, INV_CHANNEL).Value = ch Then

                Dim prev As Long
                prev = Val(wsI.Cells(invRow, INV_STOCK).Value)
                wsI.Cells(invRow, INV_STOCK).Value   = prev + qtyRcv
                wsI.Cells(invRow, INV_STOCKED).Value = Now()
                wsI.Cells(invRow, INV_UPDATED).Value = Now()

                Call WriteLog(intId, prodName, ch, "purchase_in", _
                              prev, qtyRcv, prev + qtyRcv, purchId, "仕入れ確定マクロ")
                foundAny = True
                Exit For
            End If
        Next invRow

        If Not foundAny Then
            MsgBox "[" & intId & "-" & ch & "] が在庫管理に見つかりません。" & Chr(10) & _
                   "先に商品登録マクロを実行してください。", vbExclamation
        End If
    Next ch

    ' 仕入れ管理を更新
    wsPO.Cells(poRow, PO_STATUS).Value    = "received"
    wsPO.Cells(poRow, PO_REFLECTED).Value = Now()

    MsgBox "在庫反映完了。ステータスを received に更新しました。", vbInformation
End Sub

' ================================================================
' ③ 在庫修正マクロ
' 使い方：在庫管理で修正対象行を選択して実行
' ================================================================
Sub AdjustInventory()
    Dim wsI   As Worksheet
    Dim invRow As Long
    Dim intId As String, ch As String, prodName As String
    Dim curStock As Long, newStock As Long

    Set wsI = ThisWorkbook.Sheets(SH_INVENTORY)

    invRow = ActiveCell.Row
    If invRow <= 1 Then
        MsgBox "在庫管理で対象行を選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    intId    = Trim(wsI.Cells(invRow, INV_INTID).Value)
    ch       = Trim(wsI.Cells(invRow, INV_CHANNEL).Value)
    curStock = Val(wsI.Cells(invRow, INV_STOCK).Value)
    prodName = Trim(wsI.Cells(invRow, INV_NAME).Value)

    If intId = "" Then MsgBox "内部管理IDが空です。", vbExclamation : Exit Sub

    ' 種別選択
    Dim txType As String
    Dim typeNum As String
    typeNum = InputBox("修正種別を選択してください（番号で入力）:" & Chr(10) & Chr(10) & _
                       "1: adjustment  （棚卸し調整）" & Chr(10) & _
                       "2: disposal    （廃棄・処分）" & Chr(10) & _
                       "3: return_in   （返品入庫）" & Chr(10) & _
                       "4: sale_out    （販売出庫）" & Chr(10) & _
                       "5: fba_transfer（FBA転送）", _
                       "在庫修正 - 種別選択")
    If typeNum = "" Then Exit Sub
    Select Case typeNum
        Case "1": txType = "adjustment"
        Case "2": txType = "disposal"
        Case "3": txType = "return_in"
        Case "4": txType = "sale_out"
        Case "5": txType = "fba_transfer"
        Case Else: MsgBox "1〜5を入力してください。", vbExclamation : Exit Sub
    End Select

    ' 修正後在庫数
    Dim newStockStr As String
    newStockStr = InputBox("現在庫数: " & curStock & Chr(10) & Chr(10) & _
                           "修正後の在庫数を入力してください:", _
                           "在庫修正 - " & intId & " / " & ch)
    If newStockStr = "" Then Exit Sub
    If Not IsNumeric(newStockStr) Then MsgBox "数値を入力してください。", vbExclamation : Exit Sub
    newStock = CLng(newStockStr)
    If newStock < 0 Then MsgBox "0以上の値を入力してください。", vbExclamation : Exit Sub

    ' 備考
    Dim noteStr As String
    noteStr = InputBox("備考（省略可）:", "在庫修正 - 備考")
    If noteStr = "" Then noteStr = txType & " による手動調整"

    ' 確認
    Dim diff As Long
    diff = newStock - curStock
    If MsgBox("【在庫修正確認】" & Chr(10) & _
              "ID: " & intId & " / " & ch & Chr(10) & _
              "種別: " & txType & Chr(10) & _
              "現在庫: " & curStock & "  →  修正後: " & newStock & _
              "  （差分: " & IIf(diff >= 0, "+" & diff, diff) & "）" & Chr(10) & Chr(10) & _
              "修正して在庫異動ログを記録しますか？", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ' 更新
    wsI.Cells(invRow, INV_STOCK).Value   = newStock
    wsI.Cells(invRow, INV_UPDATED).Value = Now()

    Call WriteLog(intId, prodName, ch, txType, curStock, diff, newStock, "", noteStr)

    MsgBox "在庫修正完了。異動ログへの記録も完了しました。", vbInformation
End Sub

' ================================================================
' ④ 発注アラート更新マクロ
' ================================================================
Sub UpdateReorderAlert()
    Dim wsI    As Worksheet, wsAlert As Worksheet, wsM As Worksheet
    Dim invRow As Long, alertRow As Long, lastInvRow As Long

    Set wsI    = ThisWorkbook.Sheets(SH_INVENTORY)
    Set wsAlert = ThisWorkbook.Sheets(SH_ALERT)
    Set wsM    = ThisWorkbook.Sheets(SH_MASTER)

    ' アラートシートをクリア（3行目以降のデータ）
    Dim lastAlert As Long
    lastAlert = wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Row
    If lastAlert >= 3 Then wsAlert.Range("A3:I" & lastAlert).ClearContents

    lastInvRow = wsI.Cells(wsI.Rows.Count, INV_INTID).End(xlUp).Row
    alertRow = 3

    For invRow = 2 To lastInvRow
        Dim intId As String
        Dim avail As Long, reorder As Long
        intId   = Trim(wsI.Cells(invRow, INV_INTID).Value)
        avail   = Val(wsI.Cells(invRow, INV_AVAIL).Value)
        reorder = Val(wsI.Cells(invRow, INV_REORDER).Value)

        If intId = ""  Then GoTo NextInvRow
        If reorder = 0 Then GoTo NextInvRow  ' 発注点未設定はスキップ
        If avail > reorder Then GoTo NextInvRow

        Dim supp As String
        Dim sv As Variant
        sv = Application.VLookup(intId, wsM.Range("A:L"), PM_SUPPLIER, False)
        supp = IIf(IsError(sv), "", sv)

        With wsAlert
            .Cells(alertRow, 1).Value = intId
            .Cells(alertRow, 2).Value = wsI.Cells(invRow, INV_NAME).Value
            .Cells(alertRow, 3).Value = wsI.Cells(invRow, INV_CHANNEL).Value
            .Cells(alertRow, 4).Value = avail
            .Cells(alertRow, 5).Value = reorder
            .Cells(alertRow, 6).Value = reorder - avail
            .Cells(alertRow, 7).Value = Val(wsI.Cells(invRow, INV_RQ).Value)
            .Cells(alertRow, 8).Value = supp
            .Cells(alertRow, 9).Value = Now()
        End With
        alertRow = alertRow + 1

NextInvRow:
    Next invRow

    Dim cnt As Long
    cnt = alertRow - 3
    If cnt = 0 Then
        MsgBox "発注が必要な商品はありません。", vbInformation
    Else
        MsgBox cnt & " 件の発注アラートがあります。", vbInformation
    End If
    wsAlert.Activate
End Sub

' ================================================================
' ⑤ 内部管理ID採番
' ================================================================
Function NextInternalId() As String
    Dim wsM As Worksheet
    Dim i As Long, lastRow As Long, maxNum As Long
    Set wsM = ThisWorkbook.Sheets(SH_MASTER)
    lastRow = wsM.Cells(wsM.Rows.Count, PM_ID).End(xlUp).Row
    maxNum = 0
    For i = 2 To lastRow
        Dim v As String
        v = Trim(wsM.Cells(i, PM_ID).Value)
        If Left(v, 1) = "P" And IsNumeric(Mid(v, 2)) Then
            Dim n As Long
            n = CLng(Mid(v, 2))
            If n > maxNum Then maxNum = n
        End If
    Next i
    NextInternalId = "P" & Format(maxNum + 1, "000000")
End Function

Sub ShowNextId()
    MsgBox "次の内部管理ID: " & NextInternalId(), vbInformation, "内部管理ID採番"
End Sub
