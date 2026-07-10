Attribute VB_Name = "Module_InventoryOps"
Option Explicit

' =============================================================
'  在庫運用マクロ — 棚卸確定処理
'  対象ファイル: 在庫管理テーブル_v1.0.xlsm
'  設計方針: GS(Apps Script)移植を考慮した列定数方式
' =============================================================

' ----- シート名定数 -----
Private Const SH_STOCK   As String = "在庫管理テーブル"
Private Const SH_INPUT   As String = "棚卸入力シート"
Private Const SH_LOG     As String = "在庫異動ログ"
Private Const SH_ZONE    As String = "保管区分マスター"

' ----- 在庫管理テーブル 列番号 -----
Private Const ST_SID      As Long = 1   ' 在庫ID
Private Const ST_PID      As Long = 2   ' 内部管理ID
Private Const ST_JAN      As Long = 3   ' JAN
Private Const ST_NAME     As Long = 4   ' 商品名
Private Const ST_ZONE     As Long = 5   ' 保管区分
Private Const ST_QTY      As Long = 6   ' 在庫数
Private Const ST_ALLOC    As Long = 7   ' 引当数
Private Const ST_AVAIL    As Long = 8   ' 販売可能数
Private Const ST_REORDER  As Long = 9   ' 発注点
Private Const ST_ORDERQTY As Long = 10  ' 発注数
Private Const ST_STATUS   As Long = 11  ' 在庫ステータス
Private Const ST_LASTINV  As Long = 12  ' 最終棚卸日
Private Const ST_INVNOTE  As Long = 13  ' 棚卸備考
Private Const ST_UPDATED  As Long = 14  ' 最終更新日
Private Const ST_UPDATER  As Long = 15  ' 最終更新者
Private Const ST_NOTE     As Long = 16  ' 備考

' ----- 棚卸入力シート 列番号 -----
Private Const SI_BATCHID  As Long = 1   ' 棚卸バッチID
Private Const SI_DETAILID As Long = 2   ' 棚卸明細ID
Private Const SI_PID      As Long = 3   ' 内部管理ID
Private Const SI_JAN      As Long = 4   ' JAN
Private Const SI_NAME     As Long = 5   ' 商品名
Private Const SI_ZONE     As Long = 6   ' 保管区分
Private Const SI_SYSQTY   As Long = 7   ' システム在庫数
Private Const SI_REALQTY  As Long = 8   ' 実棚数
Private Const SI_DIFF     As Long = 9   ' 差異数量
Private Const SI_REASON   As Long = 10  ' 差異理由
Private Const SI_STAFF    As Long = 11  ' 棚卸担当者
Private Const SI_DATE     As Long = 12  ' 棚卸日
Private Const SI_STATUS   As Long = 13  ' 反映ステータス
Private Const SI_REFLTIME As Long = 14  ' 反映日時
Private Const SI_LOGREF   As Long = 15  ' 参照ログID

' ----- 在庫異動ログ 列番号 -----
Private Const LG_LOGID    As Long = 1   ' ログID
Private Const LG_DATETIME As Long = 2   ' 処理日時
Private Const LG_PID      As Long = 3   ' 内部管理ID
Private Const LG_NAME     As Long = 4   ' 商品名
Private Const LG_ZONE     As Long = 5   ' 保管区分
Private Const LG_TYPE     As Long = 6   ' 処理区分
Private Const LG_QTY_BEF  As Long = 7   ' 変更前在庫
Private Const LG_DELTA    As Long = 8   ' 増減数
Private Const LG_QTY_AFT  As Long = 9   ' 変更後在庫
Private Const LG_REASON   As Long = 10  ' 差異理由
Private Const LG_UPDATER  As Long = 11  ' 更新者
Private Const LG_REF      As Long = 12  ' 参照ID（棚卸明細ID）
Private Const LG_BATCHID  As Long = 13  ' 棚卸バッチID
Private Const LG_NOTE     As Long = 14  ' 備考

' =============================================================
'  棚卸確定メイン（「棚卸確定」ボタンに割り当て）
' =============================================================
Public Sub ConfirmInventory()

    Dim wsStock As Worksheet
    Dim wsInput As Worksheet
    Dim wsLog   As Worksheet
    Set wsStock = ThisWorkbook.Sheets(SH_STOCK)
    Set wsInput = ThisWorkbook.Sheets(SH_INPUT)
    Set wsLog   = ThisWorkbook.Sheets(SH_LOG)

    ' ---- 処理対象行を収集（反映ステータス = "未反映" のみ）----
    Dim inputLast As Long
    inputLast = wsInput.Cells(wsInput.Rows.Count, SI_PID).End(xlUp).Row

    If inputLast < 3 Then   ' 行2はガイド行
        MsgBox "棚卸入力データがありません。", vbInformation
        Exit Sub
    End If

    Dim targetRows() As Long
    Dim tCount       As Long
    tCount = 0
    ReDim targetRows(1 To inputLast)

    Dim i As Long
    For i = 3 To inputLast   ' 行3から（行1=ヘッダー, 行2=ガイド）
        If CStr(wsInput.Cells(i, SI_STATUS).Value) = "未反映" Then
            tCount = tCount + 1
            targetRows(tCount) = i
        End If
    Next i

    If tCount = 0 Then
        MsgBox "確定対象の行がありません（全行が確定済または空白です）。", vbInformation
        Exit Sub
    End If

    ' ---- STEP 1: バリデーション（エラーなら即中断・何も変更しない）----
    Dim errMsg As String
    errMsg = ""

    For i = 1 To tCount
        Dim r    As Long:   r    = targetRows(i)
        Dim pid  As String: pid  = Trim(CStr(wsInput.Cells(r, SI_PID).Value))
        Dim zone As String: zone = Trim(CStr(wsInput.Cells(r, SI_ZONE).Value))

        If pid = "" Then _
            errMsg = errMsg & "行" & r & ": 内部管理IDが未入力です。" & vbCrLf

        If wsInput.Cells(r, SI_REALQTY).Value = "" Then _
            errMsg = errMsg & "行" & r & ": 実棚数が未入力です（0と空白は区別されます）。" & vbCrLf

        If Trim(CStr(wsInput.Cells(r, SI_STAFF).Value)) = "" Then _
            errMsg = errMsg & "行" & r & ": 棚卸担当者が未入力です。" & vbCrLf

        If wsInput.Cells(r, SI_DATE).Value = "" Then _
            errMsg = errMsg & "行" & r & ": 棚卸日が未入力です。" & vbCrLf

        If pid <> "" And zone <> "" Then
            If FindStockRow(wsStock, pid, zone) = 0 Then
                errMsg = errMsg & "行" & r & ": 内部管理ID=" & pid & _
                         " 保管区分=" & zone & " が在庫管理テーブルに存在しません。" & vbCrLf
            End If
        End If
    Next i

    If errMsg <> "" Then
        MsgBox "入力エラーが発生しました。処理を中断します。" & vbCrLf & vbCrLf & errMsg, vbExclamation
        Exit Sub
    End If

    ' ---- STEP 2: スナップショット取得（ロールバック用）----
    Dim snapStockRow() As Long
    Dim snapStockData() As Variant
    ReDim snapStockRow(1 To tCount)
    ReDim snapStockData(1 To tCount, 1 To 16)

    For i = 1 To tCount
        Dim stR As Long
        stR = FindStockRow(wsStock, _
              Trim(CStr(wsInput.Cells(targetRows(i), SI_PID).Value)), _
              Trim(CStr(wsInput.Cells(targetRows(i), SI_ZONE).Value)))
        snapStockRow(i) = stR
        Dim col As Long
        For col = 1 To 16
            snapStockData(i, col) = wsStock.Cells(stR, col).Value
        Next col
    Next i

    ' ログ処理開始前の最終行を記録
    Dim logLastBefore As Long
    logLastBefore = wsLog.Cells(wsLog.Rows.Count, LG_LOGID).End(xlUp).Row

    ' 棚卸バッチID採番（同日の既存バッチIDを参照して+1）
    Dim batchId As String
    batchId = GenerateBatchId(wsInput)

    ' 次のログID番号（既存最大+1。欠番は埋めない）
    Dim nextLogNum As Long
    nextLogNum = GetNextLogNum(wsLog)

    ' ---- STEP 3 & 4: 在庫更新 → ログ追記 → 反映ステータス更新 ----
    Dim logRowCur  As Long: logRowCur = logLastBefore + 1
    Dim detailSeq  As Long: detailSeq = 1
    Dim procTime   As Date: procTime  = Now()

    ' ログシート保護解除
    wsLog.Unprotect Password:=""

    On Error GoTo ErrHandler

    For i = 1 To tCount
        Dim inputRow As Long: inputRow = targetRows(i)
        Dim stRow    As Long: stRow    = snapStockRow(i)

        Dim pidVal    As String: pidVal    = Trim(CStr(wsInput.Cells(inputRow, SI_PID).Value))
        Dim nameVal   As String: nameVal   = Trim(CStr(wsInput.Cells(inputRow, SI_NAME).Value))
        Dim zoneVal   As String: zoneVal   = Trim(CStr(wsInput.Cells(inputRow, SI_ZONE).Value))
        Dim realQty   As Long:   realQty   = CLng(wsInput.Cells(inputRow, SI_REALQTY).Value)
        Dim sysQty    As Long:   sysQty    = CLng(snapStockData(i, ST_QTY))
        Dim staffVal  As String: staffVal  = Trim(CStr(wsInput.Cells(inputRow, SI_STAFF).Value))
        Dim invDate   As Variant: invDate  = wsInput.Cells(inputRow, SI_DATE).Value
        Dim reasonVal As String: reasonVal = Trim(CStr(wsInput.Cells(inputRow, SI_REASON).Value))
        Dim allocQty  As Long:   allocQty  = 0
        If IsNumeric(snapStockData(i, ST_ALLOC)) Then allocQty = CLng(snapStockData(i, ST_ALLOC))

        ' 棚卸明細ID: バッチID + 連番4桁
        Dim detailId As String
        detailId = batchId & "-" & Format(detailSeq, "0000")
        detailSeq = detailSeq + 1

        ' 処理区分: 最終棚卸日が空 = 初回棚卸、それ以外 = 棚卸調整
        Dim procType As String
        If IsEmpty(snapStockData(i, ST_LASTINV)) Or snapStockData(i, ST_LASTINV) = "" Then
            procType = "初回棚卸"
        Else
            procType = "棚卸調整"
        End If

        ' 販売可能数と在庫ステータスを再計算
        Dim avail     As Long:   avail     = realQty - allocQty
        Dim newStatus As String: newStatus = CalcStatus(avail, snapStockData(i, ST_REORDER))

        ' 在庫管理テーブルを更新
        wsStock.Cells(stRow, ST_QTY).Value     = realQty
        wsStock.Cells(stRow, ST_AVAIL).Value   = avail
        wsStock.Cells(stRow, ST_STATUS).Value  = newStatus
        wsStock.Cells(stRow, ST_LASTINV).Value = invDate
        wsStock.Cells(stRow, ST_INVNOTE).Value = reasonVal
        wsStock.Cells(stRow, ST_UPDATED).Value = procTime
        wsStock.Cells(stRow, ST_UPDATER).Value = staffVal

        ' ログIDを確定（L000001形式）
        Dim logId As String: logId = "L" & Format(nextLogNum, "000000")
        nextLogNum = nextLogNum + 1

        ' 在庫異動ログへ追記
        With wsLog
            .Cells(logRowCur, LG_LOGID).Value    = logId
            .Cells(logRowCur, LG_DATETIME).Value = procTime
            .Cells(logRowCur, LG_PID).Value      = pidVal
            .Cells(logRowCur, LG_NAME).Value     = nameVal
            .Cells(logRowCur, LG_ZONE).Value     = zoneVal
            .Cells(logRowCur, LG_TYPE).Value     = procType
            .Cells(logRowCur, LG_QTY_BEF).Value  = sysQty
            .Cells(logRowCur, LG_DELTA).Value    = realQty - sysQty
            .Cells(logRowCur, LG_QTY_AFT).Value  = realQty
            .Cells(logRowCur, LG_REASON).Value   = reasonVal
            .Cells(logRowCur, LG_UPDATER).Value  = staffVal
            .Cells(logRowCur, LG_REF).Value      = detailId
            .Cells(logRowCur, LG_BATCHID).Value  = batchId
            .Cells(logRowCur, LG_NOTE).Value     = ""
        End With
        logRowCur = logRowCur + 1

        ' 棚卸入力シートを更新（STEP 4: 全処理成功後に反映）
        wsInput.Cells(inputRow, SI_BATCHID).Value  = batchId
        wsInput.Cells(inputRow, SI_DETAILID).Value = detailId
        wsInput.Cells(inputRow, SI_STATUS).Value   = "確定済"
        wsInput.Cells(inputRow, SI_REFLTIME).Value = procTime
        wsInput.Cells(inputRow, SI_LOGREF).Value   = logId
    Next i

    ' ログシート再保護
    wsLog.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=""

    MsgBox "棚卸確定が完了しました。" & vbCrLf & vbCrLf & _
           "更新件数　　: " & tCount & " 件" & vbCrLf & _
           "棚卸バッチID: " & batchId, vbInformation
    Exit Sub

' =============================================================
'  エラーハンドラ ＆ ロールバック
' =============================================================
ErrHandler:
    Dim errNum  As Long:   errNum  = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    On Error Resume Next

    Dim rbErr As String: rbErr = ""

    ' 1. 在庫管理テーブルを変更前の値へ復元
    Dim j As Long, c As Long
    For j = 1 To tCount
        If snapStockRow(j) > 0 Then
            For c = 1 To 16
                wsStock.Cells(snapStockRow(j), c).Value = snapStockData(j, c)
            Next c
            If Err.Number <> 0 Then
                rbErr = rbErr & "在庫管理テーブル 行" & snapStockRow(j) & _
                        " (PID=" & wsInput.Cells(targetRows(j), SI_PID).Value & ") の復元失敗。" & vbCrLf
                Err.Clear
            End If
        End If
    Next j

    ' 2. 在庫異動ログの追記行を削除
    If logRowCur > logLastBefore + 1 Then
        wsLog.Rows(logLastBefore + 1 & ":" & logRowCur - 1).Delete
        If Err.Number <> 0 Then
            rbErr = rbErr & "在庫異動ログ 行" & (logLastBefore + 1) & "〜" & (logRowCur - 1) & " の削除失敗。" & vbCrLf
            Err.Clear
        End If
    End If

    ' 3. 棚卸入力シートを未反映状態に戻す
    For j = 1 To tCount
        wsInput.Cells(targetRows(j), SI_STATUS).Value   = "未反映"
        wsInput.Cells(targetRows(j), SI_REFLTIME).Value = ""
        wsInput.Cells(targetRows(j), SI_LOGREF).Value   = ""
        wsInput.Cells(targetRows(j), SI_BATCHID).Value  = ""
        wsInput.Cells(targetRows(j), SI_DETAILID).Value = ""
        If Err.Number <> 0 Then
            rbErr = rbErr & "棚卸入力シート 行" & targetRows(j) & " の復元失敗。" & vbCrLf
            Err.Clear
        End If
    Next j

    ' ログシート再保護（エラー時も必ず実行）
    wsLog.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=""

    ' 結果メッセージ
    Dim msg As String
    msg = "エラーが発生しました。処理を中断しロールバックしました。" & vbCrLf & vbCrLf & _
          "エラー番号: " & errNum & vbCrLf & _
          "エラー内容: " & errDesc & vbCrLf

    If rbErr <> "" Then
        msg = msg & vbCrLf & _
              "【警告】以下の復元に失敗しました。手動確認が必要です:" & vbCrLf & rbErr
    Else
        msg = msg & vbCrLf & "ロールバックは正常に完了しました。"
    End If

    MsgBox msg, vbCritical
End Sub

' =============================================================
'  ユーティリティ関数
' =============================================================

' 在庫管理テーブルから PID + 保管区分 に一致する行番号を返す（0 = 未発見）
Private Function FindStockRow(ws As Worksheet, pid As String, zone As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, ST_PID).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, ST_PID).Value)) = pid And _
           Trim(CStr(ws.Cells(i, ST_ZONE).Value)) = zone Then
            FindStockRow = i
            Exit Function
        End If
    Next i
    FindStockRow = 0
End Function

' 販売可能数から在庫ステータスを計算
' 判定順: 引数 avail が既に「在庫数 − 引当数」の計算済み値
Private Function CalcStatus(avail As Long, reorderPt As Variant) As String
    If avail <= 0 Then
        CalcStatus = "欠品"
    ElseIf Not IsEmpty(reorderPt) And IsNumeric(reorderPt) And avail <= CLng(reorderPt) Then
        CalcStatus = "要発注"
    Else
        CalcStatus = "在庫あり"
    End If
End Function

' 棚卸バッチID採番: STK-YYYYMMDD-NNN（同日の最大連番+1）
Private Function GenerateBatchId(wsInput As Worksheet) As String
    Dim todayStr As String: todayStr = Format(Date, "YYYYMMDD")
    Dim prefix   As String: prefix   = "STK-" & todayStr & "-"
    Dim maxSeq   As Long:   maxSeq   = 0

    Dim lastRow As Long
    lastRow = wsInput.Cells(wsInput.Rows.Count, SI_BATCHID).End(xlUp).Row
    Dim i As Long
    For i = 3 To lastRow   ' 行1=ヘッダー, 行2=ガイド
        Dim val As String: val = CStr(wsInput.Cells(i, SI_BATCHID).Value)
        If Left(val, Len(prefix)) = prefix Then
            Dim seq As Long
            If IsNumeric(Mid(val, Len(prefix) + 1)) Then
                seq = CLng(Mid(val, Len(prefix) + 1))
                If seq > maxSeq Then maxSeq = seq
            End If
        End If
    Next i

    GenerateBatchId = prefix & Format(maxSeq + 1, "000")
End Function

' ログIDの次の番号を返す（既存最大値+1。欠番は埋めない）
Private Function GetNextLogNum(wsLog As Worksheet) As Long
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, LG_LOGID).End(xlUp).Row
    If lastRow < 2 Then
        GetNextLogNum = 1
        Exit Function
    End If
    Dim maxNum As Long: maxNum = 0
    Dim i As Long
    For i = 2 To lastRow
        Dim val As String: val = CStr(wsLog.Cells(i, LG_LOGID).Value)
        If Left(val, 1) = "L" And IsNumeric(Mid(val, 2)) Then
            Dim num As Long: num = CLng(Mid(val, 2))
            If num > maxNum Then maxNum = num
        End If
    Next i
    GetNextLogNum = maxNum + 1
End Function
