Attribute VB_Name = "Module_Inventory"
Option Explicit

Const SH_MASTER    As String = "���i�}�X�^�["
Const SH_INVENTORY As String = "�݌ɊǗ�"
Const SH_PURCHASE  As String = "�d����Ǘ�"
Const SH_LOG       As String = "�݌Ɉٓ����O"
Const SH_ALERT     As String = "�����A���[�g"

Const PM_ID       As Integer = 1
Const PM_NAME     As Integer = 2
Const PM_STATUS   As Integer = 15
Const PM_SUPPLIER As Integer = 12
Const PM_CREATED  As Integer = 17
Const PM_UPDATED  As Integer = 18

Const INV_INTID    As Integer = 2
Const INV_CHANNEL  As Integer = 3
Const INV_NAME     As Integer = 4
Const INV_STOCK    As Integer = 5
Const INV_RESERVED As Integer = 6
Const INV_REORDER  As Integer = 8
Const INV_RQ       As Integer = 9
Const INV_STOCKED  As Integer = 11
Const INV_UPDATED  As Integer = 13

Const PO_ID        As Integer = 1
Const PO_INTID     As Integer = 2
Const PO_NAME      As Integer = 3
Const PO_QTY_ORD   As Integer = 8
Const PO_QTY_RCV   As Integer = 9
Const PO_DEST      As Integer = 14
Const PO_STATUS    As Integer = 15
Const PO_REFLECTED As Integer = 19

Const LOG_ID      As Integer = 1
Const LOG_DATE    As Integer = 2
Const LOG_INTID   As Integer = 3
Const LOG_NAME    As Integer = 4
Const LOG_CHANNEL As Integer = 5
Const LOG_TYPE    As Integer = 6
Const LOG_BEFORE  As Integer = 7
Const LOG_CHANGE  As Integer = 8
Const LOG_AFTER   As Integer = 9
Const LOG_REF     As Integer = 10
Const LOG_NOTES   As Integer = 11

Private Sub WriteLog(intId As String, prodName As String, ch As String, _
                     txType As String, before As Long, change As Long, _
                     after As Long, refId As String, note As String)
    Dim wsLog As Worksheet
    Dim logRow As Long
    Set wsLog = ThisWorkbook.Sheets(SH_LOG)
    wsLog.Unprotect Password:="log2024"
    logRow = 2
    Do While wsLog.Cells(logRow, LOG_ID).Value <> ""
        logRow = logRow + 1
    Loop
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

Sub RegisterProduct()
    Dim wsM As Worksheet, wsI As Worksheet
    Dim masterRow As Long, lastInvRow As Long
    Dim intId As String
    Set wsM = ThisWorkbook.Sheets(SH_MASTER)
    Set wsI = ThisWorkbook.Sheets(SH_INVENTORY)
    masterRow = ActiveCell.Row
    If masterRow <= 1 Then
        MsgBox "���i�}�X�^�[�̓o�^�s��I�����Ă�����s���Ă��������B", vbExclamation
        Exit Sub
    End If
    intId = Trim(wsM.Cells(masterRow, PM_ID).Value)
    If intId = "" Then
        MsgBox "�����Ǘ�ID�iA��j����͂��Ă��������B", vbExclamation
        Exit Sub
    End If
    Dim invRow As Long
    lastInvRow = wsI.Cells(wsI.Rows.Count, INV_INTID).End(xlUp).Row
    For invRow = 2 To lastInvRow
        If wsI.Cells(invRow, INV_INTID).Value = intId Then
            If MsgBox("[" & intId & "] �̍݌ɍs�����ɑ��݂��܂��B�ǉ����܂����H", vbYesNo + vbExclamation) = vbNo Then
                Exit Sub
            End If
            Exit For
        End If
    Next invRow
    wsM.Cells(masterRow, PM_CREATED).Value = Now()
    wsM.Cells(masterRow, PM_UPDATED).Value = Now()
    If Trim(wsM.Cells(masterRow, PM_STATUS).Value) = "" Then
        wsM.Cells(masterRow, PM_STATUS).Value = "active"
    End If
    Dim channels As Variant
    channels = Array("self", "fba", "rakuten")
    Dim ch As Variant
    For Each ch In channels
        lastInvRow = wsI.Cells(wsI.Rows.Count, INV_INTID).End(xlUp).Row + 1
        wsI.Cells(lastInvRow, INV_INTID).Value    = intId
        wsI.Cells(lastInvRow, INV_CHANNEL).Value  = ch
        wsI.Cells(lastInvRow, INV_STOCK).Value    = 0
        wsI.Cells(lastInvRow, INV_RESERVED).Value = 0
        wsI.Cells(lastInvRow, INV_UPDATED).Value  = Now()
    Next ch
    MsgBox "�o�^����: [" & intId & "] �� self / fba / rakuten ��3�`���l���Œǉ����܂����B", vbInformation
    ThisWorkbook.Sheets(SH_INVENTORY).Activate
End Sub

Sub ConfirmPurchase()
    Dim wsPO As Worksheet, wsI As Worksheet
    Dim poRow As Long, invRow As Long, lastInvRow As Long
    Dim intId As String, dest As String, prodName As String
    Dim qtyRcv As Long, purchId As String
    Set wsPO = ThisWorkbook.Sheets(SH_PURCHASE)
    Set wsI  = ThisWorkbook.Sheets(SH_INVENTORY)
    poRow = ActiveCell.Row
    If poRow <= 1 Then
        MsgBox "�d����Ǘ��̑Ώۍs��I�����Ă�����s���Ă��������B", vbExclamation
        Exit Sub
    End If
    If Trim(CStr(wsPO.Cells(poRow, PO_REFLECTED).Value)) <> "" Then
        MsgBox "���̍s�͂��łɍ݌ɔ��f�ς݂ł��B" & Chr(10) & _
               "���f����: " & wsPO.Cells(poRow, PO_REFLECTED).Value, vbExclamation
        Exit Sub
    End If
    intId   = Trim(wsPO.Cells(poRow, PO_INTID).Value)
    dest    = Trim(wsPO.Cells(poRow, PO_DEST).Value)
    qtyRcv  = Val(wsPO.Cells(poRow, PO_QTY_RCV).Value)
    purchId = Trim(wsPO.Cells(poRow, PO_ID).Value)
    prodName = Trim(wsPO.Cells(poRow, PO_NAME).Value)
    If intId = ""  Then MsgBox "internal_id ����ł��B",       vbExclamation : Exit Sub
    If qtyRcv <= 0 Then MsgBox "���א��iI��j�� 0 �ȉ��ł��B", vbExclamation : Exit Sub
    If dest = ""   Then MsgBox "���ɐ�iN��j�����ݒ�ł��B",  vbExclamation : Exit Sub
    If MsgBox("�y�d����m��z" & Chr(10) & _
              "���iID: " & intId & Chr(10) & _
              "���א�: " & qtyRcv & "  ���ɐ�: " & dest & Chr(10) & Chr(10) & _
              "�݌ɂ����Z���A�ٓ����O���L�^���܂��B��낵���ł����H", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub
    lastInvRow = wsI.Cells(wsI.Rows.Count, INV_INTID).End(xlUp).Row
    Dim found As Boolean
    found = False
    For invRow = 2 To lastInvRow
        If wsI.Cells(invRow, INV_INTID).Value = intId And _
           wsI.Cells(invRow, INV_CHANNEL).Value = dest Then
            Dim prev As Long
            prev = Val(wsI.Cells(invRow, INV_STOCK).Value)
            wsI.Cells(invRow, INV_STOCK).Value   = prev + qtyRcv
            wsI.Cells(invRow, INV_STOCKED).Value = Now()
            wsI.Cells(invRow, INV_UPDATED).Value = Now()
            Call WriteLog(intId, prodName, dest, "purchase_in", _
                          prev, qtyRcv, prev + qtyRcv, purchId, "�d����m��}�N��")
            found = True
            Exit For
        End If
    Next invRow
    If Not found Then
        MsgBox "[" & intId & "-" & dest & "] ���݌ɊǗ��Ɍ�����܂���B" & Chr(10) & _
               "��ɏ��i�o�^�}�N���iRegisterProduct�j�����s���Ă��������B", vbExclamation
        Exit Sub
    End If
    wsPO.Cells(poRow, PO_STATUS).Value    = "received"
    wsPO.Cells(poRow, PO_REFLECTED).Value = Now()
    MsgBox "�݌ɔ��f�����B�X�e�[�^�X�� received �ɍX�V���܂����B", vbInformation
End Sub

Sub AdjustInventory()
    Dim wsI As Worksheet
    Dim invRow As Long
    Dim intId As String, ch As String, prodName As String
    Dim curStock As Long, newStock As Long
    Set wsI = ThisWorkbook.Sheets(SH_INVENTORY)
    invRow = ActiveCell.Row
    If invRow <= 1 Then
        MsgBox "�݌ɊǗ��̑Ώۍs��I�����Ă�����s���Ă��������B", vbExclamation
        Exit Sub
    End If
    intId    = Trim(wsI.Cells(invRow, INV_INTID).Value)
    ch       = Trim(wsI.Cells(invRow, INV_CHANNEL).Value)
    curStock = Val(wsI.Cells(invRow, INV_STOCK).Value)
    prodName = Trim(wsI.Cells(invRow, INV_NAME).Value)
    If intId = "" Then MsgBox "�����Ǘ�ID����ł��B", vbExclamation : Exit Sub
    Dim typeNum As String
    typeNum = InputBox("�C����ʂ���́i1�`5�j:" & Chr(10) & _
        "1: adjustment�i�I�����j" & Chr(10) & _
        "2: disposal�i�p���j"     & Chr(10) & _
        "3: return_in�i�ԕi�j"    & Chr(10) & _
        "4: sale_out�i�o�Ɂj"     & Chr(10) & _
        "5: fba_transfer�iFBA�]���j", "�݌ɏC��")
    Dim txType As String
    Select Case typeNum
        Case "1": txType = "adjustment"
        Case "2": txType = "disposal"
        Case "3": txType = "return_in"
        Case "4": txType = "sale_out"
        Case "5": txType = "fba_transfer"
        Case Else: Exit Sub
    End Select
    Dim newStockStr As String
    newStockStr = InputBox("���݌ɐ�: " & curStock & Chr(10) & "�C����̍݌ɐ������:", "�݌ɏC��")
    If newStockStr = "" Or Not IsNumeric(newStockStr) Then Exit Sub
    newStock = CLng(newStockStr)
    If newStock < 0 Then MsgBox "0�ȏ�̒l����͂��Ă��������B", vbExclamation : Exit Sub
    Dim diff As Long
    diff = newStock - curStock
    Dim noteStr As String
    noteStr = InputBox("���l�i�ȗ��j:", "�݌ɏC�����l")
    If noteStr = "" Then noteStr = txType & " �ɂ��蓮����"
    Dim msg As String
    msg = "�m��: " & intId & "/" & ch & "  " & curStock & " �� " & newStock
    If diff >= 0 Then msg = msg & " (+" & diff & ")" Else msg = msg & " (" & diff & ")"
    If MsgBox(msg, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    wsI.Cells(invRow, INV_STOCK).Value   = newStock
    wsI.Cells(invRow, INV_UPDATED).Value = Now()
    Call WriteLog(intId, prodName, ch, txType, curStock, diff, newStock, "", noteStr)
    MsgBox "�݌ɏC�������B�ٓ����O�ւ̋L�^���������܂����B", vbInformation
End Sub

Sub UpdateReorderAlert()
    Dim wsI As Worksheet, wsAlert As Worksheet, wsM As Worksheet
    Dim invRow As Long, alertRow As Long, lastInvRow As Long
    Set wsI     = ThisWorkbook.Sheets(SH_INVENTORY)
    Set wsAlert = ThisWorkbook.Sheets(SH_ALERT)
    Set wsM     = ThisWorkbook.Sheets(SH_MASTER)
    Dim lastAlert As Long
    lastAlert = wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Row
    If lastAlert >= 3 Then wsAlert.Range("A3:I" & lastAlert).ClearContents
    lastInvRow = wsI.Cells(wsI.Rows.Count, INV_INTID).End(xlUp).Row
    alertRow = 3
    For invRow = 2 To lastInvRow
        Dim intId As String, avail As Long, reorder As Long
        intId   = Trim(wsI.Cells(invRow, INV_INTID).Value)
        avail   = Val(wsI.Cells(invRow, 7).Value)
        reorder = Val(wsI.Cells(invRow, INV_REORDER).Value)
        If intId = "" Then GoTo NextRow
        If reorder = 0 Then GoTo NextRow
        If avail > reorder Then GoTo NextRow
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
NextRow:
    Next invRow
    Dim cnt As Long
    cnt = alertRow - 3
    If cnt = 0 Then
        MsgBox "�������K�v�ȏ��i�͂���܂���B", vbInformation
    Else
        MsgBox cnt & " ���̔����A���[�g������܂��B", vbInformation
    End If
    wsAlert.Activate
End Sub

Function NextInternalId() As String
    Dim wsM As Worksheet, i As Long, lastRow As Long, maxNum As Long
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
    MsgBox "���̓����Ǘ�ID: " & NextInternalId(), vbInformation, "�����Ǘ�ID�̔�"
End Sub