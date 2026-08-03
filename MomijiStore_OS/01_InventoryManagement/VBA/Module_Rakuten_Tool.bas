Option Explicit

' ============================================================
' 定数定義
' ============================================================
Const TC_JAN As Long = 1
Const TC_MANAGE_NO As Long = 2
Const TC_ITEM_NO As Long = 3
Const TC_SKU_NO As Long = 4
Const TC_NAME As Long = 5
Const TC_PRICE As Long = 6
Const TC_COST As Long = 7
Const TC_STOCK As Long = 8
Const TC_TOTAL As Long = 9
Const TC_JAN_CHECK As Long = 10
Const TC_IMPORT_DT As Long = 11

Const PM_JAN As Long = 1
Const PM_MANAGE_NO As Long = 2
Const PM_NAME As Long = 3
Const PM_TYPE As Long = 4
Const PM_COST As Long = 5
Const PM_UPDATED As Long = 6
Const PM_NOTE As Long = 7

Const AG_JAN As Long = 1
Const AG_MANAGE_NO As Long = 2
Const AG_SKU_NO As Long = 3
Const AG_NAME As Long = 4
Const AG_STOCK As Long = 5
Const AG_COST As Long = 6
Const AG_COST_SRC As Long = 7
Const AG_AMOUNT As Long = 8
Const AG_PRICE As Long = 9
Const AG_SELL_AMT As Long = 10
Const AG_STATUS As Long = 11
Const AG_ITEM_NO As Long = 12

Const CK_SKU As Long = 1
Const CK_JAN As Long = 2
Const CK_MANAGE_NO As Long = 3
Const CK_NAME As Long = 4
Const CK_REASON As Long = 5
Const CK_STOCK As Long = 6
Const CK_COST As Long = 7
Const CK_NOTE As Long = 8

' 列マッピング用（CSV読込時に設定）
Private colJAN As Long
Private colManageNo As Long
Private colItemNo As Long
Private colSkuNo As Long
Private colName As Long
Private colPrice As Long
Private colCost As Long
Private colStock As Long
Private colJanCheck As Long

' ============================================================
' ① 楽天CSV読込
' ============================================================
Sub CsvImport()
    ' 固定フォルダ運用（Mac対応）: Import フォルダの決まったファイルを読み込む
    Dim filePath As String
    Dim wsTake As Worksheet

    filePath = ThisWorkbook.Path & Application.PathSeparator & "Import" & _
               Application.PathSeparator & "楽天在庫リスト_import.xlsx"

    If Dir(filePath) = "" Then
        MsgBox "Importフォルダに楽天在庫リスト_import.xlsxを保存してください。", vbExclamation, "入力ファイルなし"
        Exit Sub
    End If

    Set wsTake = ThisWorkbook.Sheets("楽天CSV取込")

    If wsTake.Cells(3, TC_SKU_NO).Value <> "" Or wsTake.Cells(3, TC_MANAGE_NO).Value <> "" Then
        If MsgBox("既存の取込データを上書きします。よろしいですか？", vbYesNo + vbQuestion, "取込データの上書き確認") = vbNo Then Exit Sub
    End If

    wsTake.Range("A3:K100000").ClearContents

    Call ImportFromExcel(filePath, wsTake)

    MsgBox "読込が完了しました。" & Chr(10) & "続けて「集計実行」ボタンを押してください。", vbInformation, "読込完了"
End Sub

' ---- CSV形式から読込（Mac対応: Open For Input 使用）----
Private Sub ImportFromCSV(filePath As String, wsTake As Worksheet)
    Dim fileNum As Integer
    Dim content As String
    Dim tmpLine As String
    Dim lines() As String
    Dim headerLine As String
    Dim headers() As String
    Dim importDt As String
    Dim dataRow As Long
    Dim i As Long
    Dim lineStr As String
    Dim cols() As String
    Dim skuVal As String
    Dim pStr As String
    Dim costStr As String
    Dim sStr As String

    fileNum = FreeFile()
    content = ""
    On Error Resume Next
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, tmpLine
        content = content & tmpLine & Chr(10)
    Loop
    Close #fileNum
    On Error GoTo 0

    If Len(content) = 0 Then
        MsgBox "CSVの内容を読み取れませんでした。", vbExclamation
        Exit Sub
    End If

    ' BOM除去
    If Left(content, 1) = Chr(65279) Then content = Mid(content, 2)

    If InStr(content, Chr(10)) > 0 Then
        lines = Split(content, Chr(10))
    Else
        lines = Split(content, Chr(13))
    End If

    If UBound(lines) < 1 Then
        MsgBox "CSVのデータが1行以下です。", vbExclamation
        Exit Sub
    End If

    headerLine = Replace(lines(0), Chr(13), "")
    headers = SplitCSVLine(headerLine)
    If Not MapColumns(headers) Then Exit Sub

    importDt = Format(Now, "yyyy/mm/dd hh:mm:ss")
    dataRow = 3
    For i = 1 To UBound(lines)
        lineStr = Replace(Trim(lines(i)), Chr(13), "")
        If lineStr = "" Then GoTo NextLine

        cols = SplitCSVLine(lineStr)

        skuVal = ""
        If colSkuNo > 0 Then skuVal = Trim(SafeGetArr(cols, colSkuNo))
        If skuVal = "" And colManageNo > 0 Then skuVal = Trim(SafeGetArr(cols, colManageNo))
        If skuVal = "" Then GoTo NextLine

        wsTake.Cells(dataRow, TC_JAN).Value = SafeGetArr(cols, colJAN)
        wsTake.Cells(dataRow, TC_MANAGE_NO).Value = SafeGetArr(cols, colManageNo)
        wsTake.Cells(dataRow, TC_ITEM_NO).Value = SafeGetArr(cols, colItemNo)
        wsTake.Cells(dataRow, TC_SKU_NO).Value = SafeGetArr(cols, colSkuNo)
        wsTake.Cells(dataRow, TC_NAME).Value = SafeGetArr(cols, colName)

        pStr = CleanNum(SafeGetArr(cols, colPrice))
        If IsNumeric(pStr) Then
            wsTake.Cells(dataRow, TC_PRICE).Value = CDbl(pStr)
        Else
            wsTake.Cells(dataRow, TC_PRICE).Value = SafeGetArr(cols, colPrice)
        End If

        costStr = CleanNum(SafeGetArr(cols, colCost))
        If IsNumeric(costStr) And CDbl(costStr) > 0 Then
            wsTake.Cells(dataRow, TC_COST).Value = CDbl(costStr)
        End If

        sStr = CleanNum(SafeGetArr(cols, colStock))
        If IsNumeric(sStr) Then
            wsTake.Cells(dataRow, TC_STOCK).Value = CLng(sStr)
        Else
            wsTake.Cells(dataRow, TC_STOCK).Value = SafeGetArr(cols, colStock)
        End If

        wsTake.Cells(dataRow, TC_JAN_CHECK).Value = SafeGetArr(cols, colJanCheck)
        wsTake.Cells(dataRow, TC_IMPORT_DT).Value = importDt
        dataRow = dataRow + 1
NextLine:
    Next i
End Sub

' ---- Excel形式から読込 ----
Private Sub ImportFromExcel(filePath As String, wsTake As Worksheet)
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim lastCol As Long
    Dim headers() As String
    Dim j As Long
    Dim importDt As String
    Dim lastRow As Long
    Dim dataRow As Long
    Dim i As Long
    Dim skuVal As String
    Dim pVal As Variant
    Dim cVal As Variant
    Dim sVal As Variant

    Application.ScreenUpdating = False
    Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True)
    Set wsSrc = wbSrc.Sheets(1)

    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    ReDim headers(lastCol - 1)
    For j = 1 To lastCol
        headers(j - 1) = CStr(wsSrc.Cells(1, j).Value)
    Next j

    If Not MapColumns(headers) Then
        wbSrc.Close False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    importDt = Format(Now, "yyyy/mm/dd hh:mm:ss")
    ' JAN列は空欄があり得るため、SKU→管理番号→商品名の優先順で最終行を検出する
    Dim lastRowCol As Long
    If colSkuNo > 0 Then
        lastRowCol = colSkuNo
    ElseIf colManageNo > 0 Then
        lastRowCol = colManageNo
    Else
        lastRowCol = colName
    End If
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, lastRowCol).End(xlUp).Row

    dataRow = 3
    For i = 2 To lastRow
        skuVal = ""
        If colSkuNo > 0 Then skuVal = Trim(CStr(wsSrc.Cells(i, colSkuNo).Value))
        If skuVal = "" And colManageNo > 0 Then skuVal = Trim(CStr(wsSrc.Cells(i, colManageNo).Value))
        If skuVal = "" Then GoTo NextExcelLine

        wsTake.Cells(dataRow, TC_JAN).Value = SafeGetCell(wsSrc, i, colJAN)
        wsTake.Cells(dataRow, TC_MANAGE_NO).Value = SafeGetCell(wsSrc, i, colManageNo)
        wsTake.Cells(dataRow, TC_ITEM_NO).Value = SafeGetCell(wsSrc, i, colItemNo)
        wsTake.Cells(dataRow, TC_SKU_NO).Value = SafeGetCell(wsSrc, i, colSkuNo)
        wsTake.Cells(dataRow, TC_NAME).Value = SafeGetCell(wsSrc, i, colName)

        pVal = SafeGetCell(wsSrc, i, colPrice)
        If IsNumeric(pVal) Then
            wsTake.Cells(dataRow, TC_PRICE).Value = CDbl(pVal)
        Else
            wsTake.Cells(dataRow, TC_PRICE).Value = pVal
        End If

        cVal = SafeGetCell(wsSrc, i, colCost)
        If IsNumeric(cVal) And CDbl(cVal) > 0 Then
            wsTake.Cells(dataRow, TC_COST).Value = CDbl(cVal)
        End If

        sVal = SafeGetCell(wsSrc, i, colStock)
        If IsNumeric(sVal) Then
            wsTake.Cells(dataRow, TC_STOCK).Value = CLng(sVal)
        Else
            wsTake.Cells(dataRow, TC_STOCK).Value = sVal
        End If

        wsTake.Cells(dataRow, TC_JAN_CHECK).Value = SafeGetCell(wsSrc, i, colJanCheck)
        wsTake.Cells(dataRow, TC_IMPORT_DT).Value = importDt
        dataRow = dataRow + 1
NextExcelLine:
    Next i

    wbSrc.Close False
    Application.ScreenUpdating = True
End Sub

' ---- 列名マッピング ----
Private Function MapColumns(headers() As String) As Boolean
    Dim i As Long
    Dim h As String

    colJAN = 0: colManageNo = 0: colItemNo = 0: colSkuNo = 0
    colName = 0: colPrice = 0: colCost = 0: colStock = 0: colJanCheck = 0

    For i = 0 To UBound(headers)
        ' 空白除去と括弧の全角統一で表記ゆれ（半角括弧・空白混入）を吸収する
        h = Trim(headers(i))
        h = Replace(h, " ", "")
        h = Replace(h, "　", "")
        h = Replace(h, "(", "（")
        h = Replace(h, ")", "）")
        Select Case h
            Case "JAN", "JANコード", "jan": colJAN = i + 1
            Case "楽天商品管理番号（ASIN）", "楽天商品管理番号", "商品管理番号", "item_id", "管理番号": colManageNo = i + 1
            Case "商品番号": colItemNo = i + 1
            Case "SKU管理番号", "sku_id", "SKU": colSkuNo = i + 1
            Case "商品名", "商品タイトル", "item_name": colName = i + 1
            Case "販売価格", "price": colPrice = i + 1
            Case "単価", "原価", "仕入れ価格", "仕入価格", "cost": colCost = i + 1
            Case "在庫数", "inventory", "在庫": colStock = i + 1
            Case "JAN検証結果": colJanCheck = i + 1
        End Select
    Next i

    If colSkuNo = 0 And colManageNo = 0 Then
        MsgBox "CSVに「SKU管理番号」または「楽天商品管理番号」列が見つかりません。" & Chr(10) & "CSVファイルを確認してください。", vbExclamation, "列名エラー"
        MapColumns = False: Exit Function
    End If
    If colStock = 0 Then
        MsgBox "CSVに「在庫数」列が見つかりません。", vbExclamation, "列名エラー"
        MapColumns = False: Exit Function
    End If
    MapColumns = True
End Function

' ============================================================
' ② 集計実行
' ============================================================
Sub RunAggregation()
    Dim wsTake As Worksheet, wsMaster As Worksheet
    Dim wsAgg As Worksheet, wsCheck As Worksheet
    Dim lastRow As Long
    Dim dicSku As Collection
    Dim aggRow As Long, chkRow As Long
    Dim totalAmount As Double, totalStock As Double
    Dim skuWithStock As Long, skuRegistered As Long
    Dim skuUnregistered As Long, totalSellAmount As Double
    Dim i As Long
    Dim skuNo As String, janVal As String, manageNo As String
    Dim itemNo As String, itemName As String
    Dim stockVal As Variant, priceVal As Variant, csvCost As Variant
    Dim hasIssue As Boolean, issueReason As String
    Dim stockNum As Double
    Dim csvCostNum As Double, csvCostValid As Boolean
    Dim dupKey As String, isDup As Boolean, tmpVal As Variant
    Dim adoptedCost As Variant, costSource As String
    Dim isSet As Boolean, isAst As Boolean
    Dim mSrc As String, mVal As Variant
    Dim uJan As String, uQty As Long
    Dim uSrc As String, uCost As Variant
    Dim aSrc As String, aVal As Variant
    Dim nSrc As String, nVal As Variant
    Dim invAmount As Variant, sellAmount As Variant
    Dim statusLabel As String
    Dim aggDt As String

    Set wsTake = ThisWorkbook.Sheets("楽天CSV取込")
    Set wsMaster = ThisWorkbook.Sheets("原価マスター")
    Set wsAgg = ThisWorkbook.Sheets("在庫金額集計")
    Set wsCheck = ThisWorkbook.Sheets("要確認一覧")

    If wsTake.Cells(3, TC_SKU_NO).Value = "" And wsTake.Cells(3, TC_MANAGE_NO).Value = "" Then
        MsgBox "取込データがありません。先に「楽天CSV読込」ボタンを押してください。", vbExclamation, "データなし"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    wsAgg.Range("A10:L100000").ClearContents
    wsAgg.Range("A10:L100000").Interior.ColorIndex = xlNone
    wsCheck.Range("A3:H100000").ClearContents
    wsCheck.Range("A3:H100000").Interior.ColorIndex = xlNone

    lastRow = wsTake.Cells(wsTake.Rows.Count, TC_SKU_NO).End(xlUp).Row
    If lastRow < 3 Then lastRow = wsTake.Cells(wsTake.Rows.Count, TC_MANAGE_NO).End(xlUp).Row

    Set dicSku = New Collection

    aggRow = 10
    chkRow = 3
    totalAmount = 0: totalStock = 0
    skuWithStock = 0: skuRegistered = 0
    skuUnregistered = 0: totalSellAmount = 0

    For i = 3 To lastRow
        skuNo = Trim(CStr(wsTake.Cells(i, TC_SKU_NO).Value))
        janVal = Trim(CStr(wsTake.Cells(i, TC_JAN).Value))
        manageNo = Trim(CStr(wsTake.Cells(i, TC_MANAGE_NO).Value))
        itemNo = Trim(CStr(wsTake.Cells(i, TC_ITEM_NO).Value))
        itemName = Trim(CStr(wsTake.Cells(i, TC_NAME).Value))
        stockVal = wsTake.Cells(i, TC_STOCK).Value
        priceVal = wsTake.Cells(i, TC_PRICE).Value
        csvCost = wsTake.Cells(i, TC_COST).Value

        If skuNo = "" And manageNo = "" Then GoTo NextRow

        hasIssue = False
        issueReason = ""

        If janVal = "" And manageNo = "" Then
            hasIssue = True: issueReason = issueReason & "ID未設定 / "
        End If

        stockNum = 0
        If Not IsNumeric(stockVal) Then
            hasIssue = True: issueReason = issueReason & "在庫数エラー(数値でない) / "
        Else
            stockNum = CDbl(stockVal)
            If stockNum < 0 Then hasIssue = True: issueReason = issueReason & "在庫数異常(負数) / "
        End If

        csvCostNum = 0
        csvCostValid = False
        If CStr(csvCost) <> "" And Not IsNull(csvCost) Then
            If IsNumeric(csvCost) Then
                csvCostNum = CDbl(csvCost)
                If csvCostNum > 0 Then
                    csvCostValid = True
                ElseIf csvCostNum < 0 Then
                    hasIssue = True: issueReason = issueReason & "原価異常(負数) / "
                End If
            Else
                hasIssue = True: issueReason = issueReason & "原価エラー(数値でない) / "
            End If
        End If

        ' 楽天では同一SKU番号が別商品(管理番号違い)に使い回されるため、
        ' 管理番号×SKUのペアで重複判定する(SKU単独だと別商品を誤って除外する)
        dupKey = manageNo & "|" & skuNo
        isDup = False
        On Error Resume Next
        tmpVal = dicSku(dupKey)
        If Err.Number = 0 Then isDup = True
        Err.Clear: On Error GoTo 0
        If isDup Then
            hasIssue = True: issueReason = issueReason & "重複取込疑い / "
        Else
            dicSku.Add dupKey, dupKey
        End If

        isSet = IsSetSku(skuNo)
        isAst = IsAstSku(skuNo)

        ' 重複行・負数在庫行は要確認一覧のみに記載し、通常集計（明細・合計）から除外する
        If isDup Or stockNum < 0 Then GoTo WriteCheckOnly

        adoptedCost = ""
        costSource = "未登録"

        If csvCostValid Then
            adoptedCost = csvCostNum
            costSource = "CSV"
        ElseIf isSet Then
            mVal = FindCostInMaster(wsMaster, skuNo, manageNo, mSrc)
            If mVal <> "" Then
                adoptedCost = mVal: costSource = mSrc
            Else
                uJan = ExtractSetJan(skuNo)
                uQty = ExtractSetQty(skuNo)
                If uJan <> "" And uQty > 0 Then
                    uCost = FindCostInMaster(wsMaster, uJan, "", uSrc)
                    If uCost <> "" And IsNumeric(uCost) Then
                        adoptedCost = CDbl(uCost) * uQty
                        costSource = "セット推定"
                        hasIssue = True
                        If InStr(issueReason, "セット推定") = 0 Then issueReason = issueReason & "セット推定（要確認） / "
                    End If
                End If
            End If
        ElseIf isAst Then
            aVal = FindCostByManageNo(wsMaster, manageNo, aSrc)
            If aVal <> "" Then adoptedCost = aVal: costSource = aSrc
        Else
            nVal = FindCostInMaster(wsMaster, janVal, manageNo, nSrc)
            If nVal <> "" Then adoptedCost = nVal: costSource = nSrc
        End If

        If adoptedCost = "" Then
            costSource = "未登録"
            hasIssue = True
            If InStr(issueReason, "原価未登録") = 0 Then issueReason = issueReason & "原価未登録 / "
        End If

        invAmount = ""
        If adoptedCost <> "" And IsNumeric(adoptedCost) Then invAmount = stockNum * CDbl(adoptedCost)

        sellAmount = ""
        If IsNumeric(priceVal) Then sellAmount = stockNum * CDbl(priceVal)

        wsAgg.Cells(aggRow, AG_JAN).Value = janVal
        wsAgg.Cells(aggRow, AG_MANAGE_NO).Value = manageNo
        wsAgg.Cells(aggRow, AG_SKU_NO).Value = skuNo
        wsAgg.Cells(aggRow, AG_NAME).Value = itemName
        wsAgg.Cells(aggRow, AG_STOCK).Value = stockNum
        If adoptedCost <> "" Then wsAgg.Cells(aggRow, AG_COST).Value = adoptedCost
        wsAgg.Cells(aggRow, AG_COST_SRC).Value = costSource
        If invAmount <> "" Then wsAgg.Cells(aggRow, AG_AMOUNT).Value = invAmount
        If IsNumeric(priceVal) Then wsAgg.Cells(aggRow, AG_PRICE).Value = CDbl(priceVal)
        If sellAmount <> "" Then wsAgg.Cells(aggRow, AG_SELL_AMT).Value = sellAmount

        statusLabel = ""
        Select Case costSource
            Case "CSV": statusLabel = "登録済（CSV）"
            Case "原価マスター（JAN）", "原価マスター（管理番号）": statusLabel = "登録済"
            Case "セット推定": statusLabel = "推定値"
            Case Else: statusLabel = "未登録"
        End Select
        wsAgg.Cells(aggRow, AG_STATUS).Value = statusLabel
        wsAgg.Cells(aggRow, AG_ITEM_NO).Value = itemNo

        totalStock = totalStock + stockNum
        If stockNum > 0 Then skuWithStock = skuWithStock + 1
        If invAmount <> "" And IsNumeric(invAmount) Then
            totalAmount = totalAmount + CDbl(invAmount)
            skuRegistered = skuRegistered + 1
        End If
        If costSource = "未登録" Then skuUnregistered = skuUnregistered + 1
        If sellAmount <> "" And IsNumeric(sellAmount) Then totalSellAmount = totalSellAmount + CDbl(sellAmount)

        aggRow = aggRow + 1

WriteCheckOnly:
        If hasIssue Then
            If Len(issueReason) >= 3 Then issueReason = Left(issueReason, Len(issueReason) - 3)
            wsCheck.Cells(chkRow, CK_SKU).Value = skuNo
            wsCheck.Cells(chkRow, CK_JAN).Value = janVal
            wsCheck.Cells(chkRow, CK_MANAGE_NO).Value = manageNo
            wsCheck.Cells(chkRow, CK_NAME).Value = itemName
            wsCheck.Cells(chkRow, CK_REASON).Value = issueReason
            wsCheck.Cells(chkRow, CK_STOCK).Value = stockVal
            wsCheck.Cells(chkRow, CK_COST).Value = csvCost
            wsCheck.Cells(chkRow, CK_NOTE).Value = IIf(isSet, "セット商品", IIf(isAst, "アソート商品", ""))
            chkRow = chkRow + 1
        End If
NextRow:
    Next i

    ' ソート（在庫金額 降順）
    If aggRow > 11 Then
        wsAgg.Range("A10:L" & (aggRow - 1)).Sort _
            Key1:=wsAgg.Cells(10, AG_AMOUNT), Order1:=xlDescending, Header:=xlNo
    End If

    ' KPI更新
    aggDt = Format(Now, "yyyy/mm/dd hh:mm:ss")
    wsAgg.Cells(3, 2).Value = totalAmount
    wsAgg.Cells(3, 6).Value = totalSellAmount
    wsAgg.Cells(4, 2).Value = totalStock
    wsAgg.Cells(4, 6).Value = skuWithStock
    wsAgg.Cells(5, 2).Value = skuRegistered
    wsAgg.Cells(5, 6).Value = skuUnregistered
    wsAgg.Cells(6, 2).Value = aggDt

    ' 書式適用
    Call FormatAggSheet(wsAgg, aggRow)
    Call FormatCheckSheet(wsCheck, chkRow)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    wsAgg.Activate

    MsgBox "集計完了。" & Chr(10) & Chr(10) & _
           "■ 総在庫金額    : " & Format(totalAmount, "#,##0") & " 円" & Chr(10) & _
           "■ 総在庫数量    : " & Format(totalStock, "#,##0") & " 個" & Chr(10) & _
           "■ 原価未登録    : " & skuUnregistered & " SKU" & Chr(10) & Chr(10) & _
           IIf(skuUnregistered > 0, "「要確認一覧」シートを確認してください。", "要確認事項はありません。"), _
           vbInformation, "集計完了"
End Sub

' ============================================================
' 原価マスター検索
' ============================================================
Private Function FindCostInMaster(wsMaster As Worksheet, janVal As String, manageNo As String, ByRef srcOut As String) As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim cv As Variant

    FindCostInMaster = "": srcOut = ""
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, PM_COST).End(xlUp).Row
    If lastRow < 3 Then Exit Function

    If janVal <> "" Then
        For i = 3 To lastRow
            If Trim(CStr(wsMaster.Cells(i, PM_JAN).Value)) = janVal Then
                cv = wsMaster.Cells(i, PM_COST).Value
                If IsNumeric(cv) And CDbl(cv) > 0 Then
                    FindCostInMaster = CDbl(cv): srcOut = "原価マスター（JAN）": Exit Function
                End If
            End If
        Next i
    End If

    If manageNo <> "" Then
        For i = 3 To lastRow
            If Trim(CStr(wsMaster.Cells(i, PM_MANAGE_NO).Value)) = manageNo Then
                cv = wsMaster.Cells(i, PM_COST).Value
                If IsNumeric(cv) And CDbl(cv) > 0 Then
                    FindCostInMaster = CDbl(cv): srcOut = "原価マスター（管理番号）": Exit Function
                End If
            End If
        Next i
    End If
End Function

Private Function FindCostByManageNo(wsMaster As Worksheet, manageNo As String, ByRef srcOut As String) As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim cv As Variant

    FindCostByManageNo = "": srcOut = ""
    If manageNo = "" Then Exit Function
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, PM_MANAGE_NO).End(xlUp).Row
    If lastRow < 3 Then Exit Function
    For i = 3 To lastRow
        If Trim(CStr(wsMaster.Cells(i, PM_MANAGE_NO).Value)) = manageNo Then
            cv = wsMaster.Cells(i, PM_COST).Value
            If IsNumeric(cv) And CDbl(cv) > 0 Then
                FindCostByManageNo = CDbl(cv): srcOut = "原価マスター（管理番号）": Exit Function
            End If
        End If
    Next i
End Function

' ============================================================
' セット・アソート商品ヘルパー
' ============================================================
Private Function IsSetSku(skuNo As String) As Boolean
    IsSetSku = (Right(UCase(skuNo), 3) = "SET" And InStr(skuNo, "-") > 0)
End Function

Private Function IsAstSku(skuNo As String) As Boolean
    IsAstSku = (Left(UCase(skuNo), 4) = "AST-")
End Function

Private Function ExtractSetJan(skuNo As String) As String
    Dim pos As Long
    pos = InStrRev(skuNo, "-")
    ExtractSetJan = IIf(pos > 1, Left(skuNo, pos - 1), "")
End Function

Private Function ExtractSetQty(skuNo As String) As Long
    Dim pos As Long
    Dim s As String
    ExtractSetQty = 0
    pos = InStrRev(skuNo, "-")
    If pos = 0 Then Exit Function
    s = Replace(UCase(Mid(skuNo, pos + 1)), "SET", "")
    If IsNumeric(s) And CLng(s) > 0 Then ExtractSetQty = CLng(s)
End Function

' ============================================================
' CSV行分割（ダブルクォート対応）
' ============================================================
Private Function SplitCSVLine(lineStr As String) As String()
    Dim result() As String
    Dim idx As Long
    Dim pos As Long
    Dim inQ As Boolean
    Dim fld As String
    Dim c As String

    ReDim result(0)
    idx = 0
    pos = 1
    inQ = False
    fld = ""
    Do While pos <= Len(lineStr)
        c = Mid(lineStr, pos, 1)
        If inQ Then
            If c = Chr(34) Then
                If pos < Len(lineStr) And Mid(lineStr, pos + 1, 1) = Chr(34) Then
                    fld = fld & Chr(34): pos = pos + 1
                Else
                    inQ = False
                End If
            Else
                fld = fld & c
            End If
        Else
            If c = Chr(34) Then
                inQ = True
            ElseIf c = "," Then
                ReDim Preserve result(idx)
                result(idx) = fld: idx = idx + 1: fld = ""
            Else
                fld = fld & c
            End If
        End If
        pos = pos + 1
    Loop
    ReDim Preserve result(idx)
    result(idx) = fld
    SplitCSVLine = result
End Function

' ============================================================
' ユーティリティ
' ============================================================
Private Function SafeGetArr(arr() As String, col As Long) As String
    SafeGetArr = ""
    If col <= 0 Or col - 1 > UBound(arr) Then Exit Function
    SafeGetArr = Trim(arr(col - 1))
End Function

Private Function SafeGetCell(ws As Worksheet, r As Long, col As Long) As Variant
    ' IIf は両辺を常に評価するため col=0 でもCellsが実行されエラーになる → If/Else で回避
    If col > 0 Then
        SafeGetCell = ws.Cells(r, col).Value
    Else
        SafeGetCell = ""
    End If
End Function

Private Function CleanNum(s As String) As String
    CleanNum = Replace(Replace(Replace(Replace(s, ",", ""), "\", ""), Chr(165), ""), " ", "")
End Function

' ============================================================
' 書式適用
' ============================================================
Private Sub FormatAggSheet(wsAgg As Worksheet, lastRow As Long)
    Dim i As Long
    Dim src As String
    Dim rowClr As Long

    If lastRow <= 10 Then Exit Sub
    For i = 10 To lastRow - 1
        src = CStr(wsAgg.Cells(i, AG_COST_SRC).Value)
        Select Case src
            Case "CSV": rowClr = RGB(226, 239, 218)
            Case "原価マスター（JAN）", "原価マスター（管理番号）": rowClr = RGB(214, 228, 240)
            Case "セット推定": rowClr = RGB(253, 233, 217)
            Case "未登録": rowClr = RGB(252, 228, 214)
            Case Else: rowClr = IIf(i Mod 2 = 0, RGB(242, 242, 242), RGB(255, 255, 255))
        End Select
        wsAgg.Range(wsAgg.Cells(i, 1), wsAgg.Cells(i, 12)).Interior.Color = rowClr
        wsAgg.Cells(i, AG_COST).NumberFormat = "#,##0"
        wsAgg.Cells(i, AG_AMOUNT).NumberFormat = "#,##0"
        wsAgg.Cells(i, AG_PRICE).NumberFormat = "#,##0"
        wsAgg.Cells(i, AG_SELL_AMT).NumberFormat = "#,##0"
    Next i
End Sub

Private Sub FormatCheckSheet(wsCheck As Worksheet, lastRow As Long)
    Dim i As Long
    Dim reason As String
    Dim rowClr As Long

    If lastRow <= 3 Then Exit Sub
    For i = 3 To lastRow - 1
        reason = CStr(wsCheck.Cells(i, CK_REASON).Value)
        If InStr(reason, "重複") > 0 Then
            rowClr = RGB(255, 235, 156)
        ElseIf InStr(reason, "未登録") > 0 Then
            rowClr = RGB(252, 228, 214)
        ElseIf InStr(reason, "セット推定") > 0 Then
            rowClr = RGB(253, 233, 217)
        ElseIf InStr(reason, "負数") > 0 Or InStr(reason, "異常") > 0 Then
            rowClr = RGB(255, 199, 206)
        Else
            rowClr = RGB(255, 255, 204)
        End If
        wsCheck.Range(wsCheck.Cells(i, 1), wsCheck.Cells(i, 8)).Interior.Color = rowClr
    Next i
End Sub

' ============================================================
' ③ データクリア
' ============================================================
Sub ClearData()
    Dim wsAgg As Worksheet
    Dim wsCheck As Worksheet

    If MsgBox("取込データと集計結果をクリアします。（原価マスターはクリアされません）", _
              vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

    ThisWorkbook.Sheets("楽天CSV取込").Range("A3:K100000").ClearContents

    Set wsAgg = ThisWorkbook.Sheets("在庫金額集計")
    wsAgg.Range("A10:L100000").ClearContents
    wsAgg.Range("A10:L100000").Interior.ColorIndex = xlNone
    wsAgg.Cells(3, 2).Value = "": wsAgg.Cells(3, 6).Value = ""
    wsAgg.Cells(4, 2).Value = "": wsAgg.Cells(4, 6).Value = ""
    wsAgg.Cells(5, 2).Value = "": wsAgg.Cells(5, 6).Value = ""
    wsAgg.Cells(6, 2).Value = ""

    Set wsCheck = ThisWorkbook.Sheets("要確認一覧")
    wsCheck.Range("A3:H100000").ClearContents
    wsCheck.Range("A3:H100000").Interior.ColorIndex = xlNone

    MsgBox "クリアしました。", vbInformation, "完了"
End Sub

' ============================================================
' ④ Excel保存
' ============================================================
Sub SaveExcel()
    ' 固定フォルダ運用（Mac対応）: Output フォルダへ日時付きファイル名で保存する
    Dim savePath As String
    Dim outDir As String
    Dim wbNew As Workbook

    outDir = ThisWorkbook.Path & Application.PathSeparator & "Output"
    ' Mac のサンドボックスでは未許可フォルダに対して Dir() が "" を返すため、
    ' 存在チェックせず MkDir を試み、既存時のエラー75は無視する
    On Error Resume Next
    MkDir outDir
    On Error GoTo 0

    savePath = outDir & Application.PathSeparator & _
        "楽天在庫金額集計結果_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"

    Set wbNew = Workbooks.Add
    ThisWorkbook.Sheets("在庫金額集計").Copy Before:=wbNew.Sheets(1)
    Application.DisplayAlerts = False
    wbNew.SaveAs savePath, xlOpenXMLWorkbook
    wbNew.Close False
    Application.DisplayAlerts = True
    MsgBox "Excelを保存しました。" & Chr(10) & savePath, vbInformation, "保存完了"
End Sub

' ============================================================
' ⑤ PDF保存
' ============================================================
Sub SavePDF()
    ' 固定フォルダ運用（Mac対応）: Output フォルダへ日時付きファイル名で保存する
    Dim savePath As String
    Dim outDir As String

    outDir = ThisWorkbook.Path & Application.PathSeparator & "Output"
    ' Mac のサンドボックスでは未許可フォルダに対して Dir() が "" を返すため、
    ' 存在チェックせず MkDir を試み、既存時のエラー75は無視する
    On Error Resume Next
    MkDir outDir
    On Error GoTo 0

    savePath = outDir & Application.PathSeparator & _
        "楽天在庫金額集計結果_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"

    ThisWorkbook.Sheets("在庫金額集計").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=savePath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    MsgBox "PDFを保存しました。" & Chr(10) & savePath, vbInformation, "保存完了"
End Sub


