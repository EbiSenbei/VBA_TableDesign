Attribute VB_Name = "GetTableData"


Const SCHMA_NAME As String = "" 'スキーマ名
Const COL_START_ROW As Long = 7 'カラム情報は7行目からスタート
Const TBL_START_ROW As Long = 5 'テーブル情報は5行目からスタート
Const SHEET_TABLE_LIST As String = "テーブル一覧表" 'テーブル一覧表のシート名

'項目情報の構造体
Type typeTable
    lngNo           As Long     'No
    strLogicalName  As String   '論理名
    strPhysicsName  As String   '物理名
    strSchema       As String   'スキーマ名
    strHistoryFlag  As String   '履歴作成フラグ(要/否)
    strKind         As String   'テーブル種類
    strOverview     As String   'テーブル内容
End Type

'項目情報の構造体
Type typeColumn
    lngNo           As Long     'No
    strLogicalName  As String   '論理名
    strPhysicsName  As String   '物理名
    strDataType     As String   'データ型
    lngLength       As Long     'データ桁数
    lngDecimal      As Long     '小数桁数
    strRequiredFlag As String   '必須区分
    strPrimaryKey   As String   '主キー
    strDefalutData  As String   'デフォルト値
    strRemarks      As String   '備考
End Type

Public strSchemaName As String  'スキーマ名

'テーブル・カラム情報を取得
Public Sub getTableData(ByRef tTbl As typeTable, ByRef arrColumn() As typeColumn)
On Error GoTo Err0
    'アクティブなシートを取得する。
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name

    '--------------------------------------------------------------------------------------
    'アクティブなシートからテーブル情報を取得する。
    
    ' Dim tTbl As typeTable
    tTbl.strLogicalName = Trim(wsActSheet.Range("A4").Value)  'テーブル名
    tTbl.strPhysicsName = Trim(wsActSheet.Range("C4").Value)  'テーブル名(英字)
    tTbl.strSchema = SCHMA_NAME                               'スキーマ名
    tTbl.strHistoryFlag = Trim(wsActSheet.Range("I2").Value)  '履歴作成フラグ(要/否)
    tTbl.strOverview = Trim(wsActSheet.Range("D4").Value)     'テーブル内容

    '--------------------------------------------------------------------------------------
    'アクティブなシートからカラム情報を取得する
    
    '対象件数(カラム数)を取得する。
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    wkRows = wsActSheet.Cells.Rows.count
    lngMaxLine = wsActSheet.Cells(wkRows, 1).End(xlUp).Row
    lngLineCount = lngMaxLine - COL_START_ROW + 1
    
    ' Dim arrColumn() As typeColumn
    ReDim arrColumn(lngLineCount + 1)
    Dim tColBuf As typeColumn
    Dim strBuf As String

    'カラム情報を一行ずつ取得する
    Dim i As Long
    Dim cnt As Long
    cnt = 0
    For i = COL_START_ROW To lngMaxLine + 1
    
        '[No]列に取消し線がある場合、出力対象外とする。
        If (isStrikethrough(wsActSheet.Range("A" + Format(i))) = True) Then
            '[No]列に取消し線がある場合、
            '要素数を減少させる。
            lngLineCount = lngLineCount - 1
        
        Else
            '[No]列に取消し線がない場合、
            'シートから値を取得する。
            'No
            tColBuf.lngNo = wsActSheet.Range("A" + Format(i)).Value
            '論理名
            tColBuf.strLogicalName _
                = removeStrikethrough(wsActSheet.Range("B" + Format(i)))
            '物理名
            tColBuf.strPhysicsName _
                = removeStrikethrough(wsActSheet.Range("C" + Format(i)))
            'データ型
            tColBuf.strDataType _
                = removeStrikethrough(wsActSheet.Range("D" + Format(i)))
            'データ桁数
            strBuf = removeStrikethrough(wsActSheet.Range("E" + Format(i)))
            If (IsNumeric(strBuf) = True) Then
                tColBuf.lngLength = Val(strBuf)
            Else
                tColBuf.lngLength = 0
            End If
            
            '小数桁数
            strBuf = removeStrikethrough(wsActSheet.Range("F" + Format(i)))
            If (IsNumeric(strBuf) = True) Then
                tColBuf.lngDecimal = Val(strBuf)
            Else
                tColBuf.lngDecimal = 0
            End If
            '必須区分
            tColBuf.strRequiredFlag = wsActSheet.Range("G" + Format(i)).Value
            '主キー
            tColBuf.strPrimaryKey = wsActSheet.Range("H" + Format(i)).Value
            'デフォルト値
            tColBuf.strDefalutData = wsActSheet.Range("I" + Format(i)).Value
            '備考
            tColBuf.strRemarks = wsActSheet.Range("K" + Format(i)).Value
    
            '配列に値をセットする。
            arrColumn(cnt).lngNo = tColBuf.lngNo                   'No
            arrColumn(cnt).strLogicalName = Trim(tColBuf.strLogicalName) '論理名
            arrColumn(cnt).strPhysicsName = Trim(tColBuf.strPhysicsName) '物理名
            arrColumn(cnt).strDataType = Trim(tColBuf.strDataType)       'データ型
            arrColumn(cnt).lngLength = tColBuf.lngLength           'データ桁数
            arrColumn(cnt).lngDecimal = tColBuf.lngDecimal         '小数桁数
            arrColumn(cnt).strRequiredFlag = Trim(tColBuf.strRequiredFlag) '必須区分
            arrColumn(cnt).strPrimaryKey = Trim(tColBuf.strPrimaryKey)   '主キー
            arrColumn(cnt).strDefalutData = Trim(tColBuf.strDefalutData) 'デフォルト値
            arrColumn(cnt).strRemarks = Trim(tColBuf.strRemarks)         '備考
            cnt = cnt + 1
        
        End If

    Next i

    '配列サイズの再設定（取消し線でスキップ分を減らす）
    ReDim Preserve arrColumn(cnt)

    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
End Sub
