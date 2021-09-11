Attribute VB_Name = "MakePythonModule"


Const SCHMA_NAME As String = "" 'スキーマ名
Const COL_START_ROW As Long = 7 'カラム情報は7行目からスタート
Const TBL_START_ROW As Long = 5 'テーブル情報は5行目からスタート
Const SHEET_TABLE_LIST As String = "テーブル一覧表" 'テーブル一覧表のシート名

Public strSchemaName As String  'スキーマ名

'テーブル一覧表シートのテーブルスクリプトを出力。
Public Sub makePythonFile_All()
On Error GoTo Err0
    '警告を出ないように設定
    Application.DisplayAlerts = False
    
    Dim strReportText As String '出力結果をテキストで出力
    
    'アクティブなシートを取得する。
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    
    '-----------------------------------------------------------------
    '初期化
    
    'アクティブシートの切り替え
    ActiveWorkbook.Worksheets(SHEET_TABLE_LIST).Activate
        
    '対象件数(カラム数)を取得する。
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    wkRows = wsActSheet.Cells.Rows.count
    lngMaxLine = wsActSheet.Cells(wkRows, 1).End(xlUp).Row
    lngLineCount = lngMaxLine - TBL_START_ROW + 1
    
    
    '-----------------------------------------------------------------
    'テーブル一覧情報を取得
    
    'テーブル情報リストを宣言
    Dim arrTable() As typeTable
    ReDim arrTable(lngLineCount + 1)
    Dim tTblBuf As typeTable
    Dim strBuf As String
    
    'カラム情報を一行ずつ取得する
    Dim i As Long
    Dim cnt As Long
    cnt = 0
    For i = TBL_START_ROW To lngMaxLine
    
        '[No]列に取消し線がある場合、出力対象外とする。
        If (isStrikethrough(wsActSheet.Range("A" + Format(i))) = True) Then
            '[No]列に取消し線がある場合、
            '要素数を減少させる。
            lngLineCount = lngLineCount - 1
        
        Else
            '[No]列に取消し線がない場合、
            'テーブル情報を取得する。
            
            'No
            strBuf = wsActSheet.Range("A" + Format(i)).Value
            tTblBuf.lngNo = Val(strBuf)
            '論理名
            tTblBuf.strLogicalName _
                = removeStrikethrough(wsActSheet.Range("C" + Format(i)))
            '物理名
            tTblBuf.strPhysicsName _
                = removeStrikethrough(wsActSheet.Range("K" + Format(i)))
                
            '配列に値をセットする。
            arrTable(cnt).lngNo = tTblBuf.lngNo
            arrTable(cnt).strLogicalName = tTblBuf.strLogicalName
            arrTable(cnt).strPhysicsName = tTblBuf.strPhysicsName
            cnt = cnt + 1
        End If
    Next i
    
    '-----------------------------------------------------------------
    'テーブル一覧よりスクリプトを作成
    Dim strSheetName As String
    For i = 0 To cnt - 1
        strSheetName = arrTable(i).strLogicalName
        
        'アクティブシートの切り替え
        ActiveWorkbook.Worksheets(strSheetName).Activate
        
        'アクティブシートのテーブルスクリプトを出力。
        Call makePythonFile
    
    
    Next i
    
    
    '元のシートに戻す
    ActiveWorkbook.Worksheets(strActSheetName).Activate
        
    MsgBox ("出力　完了")
    
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True

End Sub


'アクティブシートのテーブルスクリプトを出力。
Public Sub makePythonFile_Sheet()
On Error GoTo Err0
   '警告を出ないように設定
    Application.DisplayAlerts = False

    'アクティブシートのテーブルスクリプトを出力。
    Call makePythonFile
    
    '警告を出るように設定を戻す
    Application.DisplayAlerts = True
   
    MsgBox (strBuf + "出力　完了")
 
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub

'アクティブシートのテーブルスクリプトを出力。
Private Sub makePythonFile()
On Error GoTo Err0

    Dim tTbl As getTableData.typeTable           'テーブル情報
    Dim arrColumn() As getTableData.typeColumn   'カラム情報
    Dim strBuf As String
    strBuf = ""
    
    '-------------------------------------------------------------------------------
'    'アクティブのシートからテーブルとカラム情報を取得
    Call getTableData.getTableData(tTbl, arrColumn)
'
'    'テーブル情報を要素クラスのPythonファイルを作成
    Call MakePythonEntityModule.outputPythonEntity(tTbl, arrColumn)
    Call MakePythonDaoModule.outputPythonDao(tTbl, arrColumn)
    '-------------------------------------------------------------------------------

    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub
