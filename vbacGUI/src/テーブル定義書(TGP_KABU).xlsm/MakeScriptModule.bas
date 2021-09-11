Attribute VB_Name = "MakeScriptModule"


Const SCHMA_NAME As String = "" 'スキーマ名
Const COL_START_ROW As Long = 7 'カラム情報は7行目からスタート
Const TBL_START_ROW As Long = 5 'テーブル情報は5行目からスタート
Const SHEET_TABLE_LIST As String = "テーブル一覧表" 'テーブル一覧表のシート名


Public strSchemaName As String  'スキーマ名

'各テーブル定義からテーブル一覧を作成
Public Sub makeTableList()

    On Error GoTo Err0
    
    '警告を出ないように設定
    Application.DisplayAlerts = False
    
    Dim lngLineCnt As Long
    lngLineCnt = TBL_START_ROW
    
    'アクティブシートの切り替え
    ActiveWorkbook.Worksheets(SHEET_TABLE_LIST).Activate
    
    'クリアする
    Range("A5:AW100").ClearContents
        
    
    Dim tTblBuf As getTableData.typeTable    'テーブル定義情報
    For Each Ws In Worksheets
        If Ws.Name <> "来歴" And Ws.Name <> SHEET_TABLE_LIST And Ws.Name <> "Sheet1" Then
                   
            'テーブル情報の初期化
            tTblBuf.lngNo = 0     'No
            tTblBuf.strLogicalName = ""   '論理名
            tTblBuf.strPhysicsName = ""   '物理名
            tTblBuf.strSchema = ""   'スキーマ名
            tTblBuf.strHistoryFlag = ""   '履歴作成フラグ(要/否)
            tTblBuf.strKind = ""   '備考
            
            'テーブル情報の取得
            tTblBuf.lngNo = lngLineCnt - TBL_START_ROW + 1 'No
            tTblBuf.strLogicalName = Ws.Range("A4").Value  '論理名
            tTblBuf.strPhysicsName = Ws.Range("C4").Value  '物理名
            tTblBuf.strKind = Ws.Range("J2").Value  'テーブル種類
            tTblBuf.strOverview = Ws.Range("D4").Value 'テーブル内容
            
            'テーブル一覧にテーブル情報をセット
            Worksheets(SHEET_TABLE_LIST).Range("A" + Format(lngLineCnt)).Value = tTblBuf.lngNo  'No
            Worksheets(SHEET_TABLE_LIST).Range("C" + Format(lngLineCnt)).Value = tTblBuf.strLogicalName '論理名
            Worksheets(SHEET_TABLE_LIST).Range("K" + Format(lngLineCnt)).Value = tTblBuf.strPhysicsName '物理名
            Worksheets(SHEET_TABLE_LIST).Range("Q" + Format(lngLineCnt)).Value = tTblBuf.strOverview    'テーブル内容
            Worksheets(SHEET_TABLE_LIST).Range("AT" + Format(lngLineCnt)).Value = tTblBuf.strKind    'テーブル種類
               
            lngLineCnt = lngLineCnt + 1
            
        End If
    Next Ws

    Exit Sub
    
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    

End Sub


'テーブル一覧表シートのテーブルスクリプトを出力。
Public Sub makeScript_CreateTable_All()
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
    Dim arrTable() As getTableData.typeTable
    ReDim arrTable(lngLineCount + 1)
    Dim tTblBuf As getTableData.typeTable
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
        
        If (strSheetName = "") Then
            'アクティブシートの切り替え
            ActiveWorkbook.Worksheets(strSheetName).Activate
            
            'アクティブシートのテーブルスクリプトを出力。
            Call makeScript_CreateTable
         End If
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
Public Sub makeScript_CreateTable_Sheet()
On Error GoTo Err0
   '警告を出ないように設定
    Application.DisplayAlerts = False

    'アクティブシートのテーブルスクリプトを出力。
    Call makeScript_CreateTable
    
    '警告を出るように設定を戻す
    Application.DisplayAlerts = True
   
    MsgBox (strBuf + "出力　完了")
 
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub

'アクティブシートのテーブルスクリプトを出力。
Private Sub makeScript_CreateTable()
On Error GoTo Err0

    Dim tTbl As getTableData.typeTable           'テーブル情報
    Dim arrColumn() As getTableData.typeColumn   'カラム情報
    Dim strBuf As String
    strBuf = ""
    
    '-------------------------------------------------------------------------------
    'アクティブのシートからテーブルとカラム情報を取得
    Call getTableData.getTableData(tTbl, arrColumn)
    
    'テーブル情報をスクリプトに出力
    Call outputScriput_CreateTable(tTbl, arrColumn)
    strBuf = strBuf + "CreateTable_" + tTbl.strPhysicsName + "(" + tTbl.strLogicalName + ").sql" + vbCrLf
    '-------------------------------------------------------------------------------

    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub

'テーブル情報をスクリプトに出力
Private Sub outputScriput_CreateTable(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn)
On Error GoTo Err0
    'アクティブなシートを取得する。
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    
    '対象件数(カラム数)を取得する。
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    '--------------------------------------------------------------------------------------
    'スクリプトを組み立てる。
   
    Dim strSql As String
    strSql = ""
    
    'DROP TABLE文を仕込む
    ' --------------------------------------------------------------------------------------
    ' Oracle
'    strSql = strSql + "BEGIN " + vbCrLf
'    If (tTbl.strSchema = "") Then
'        strSql = strSql + "   EXECUTE IMMEDIATE 'DROP TABLE " + tTbl.strPhysicsName + "';" + vbCrLf
'    Else
'        strSql = strSql + "   EXECUTE IMMEDIATE 'DROP TABLE " + tTbl.strSchema + "." + tTbl.strPhysicsName + "';" + vbCrLf
'    End If
'    strSql = strSql + "EXCEPTION " + vbCrLf
'    strSql = strSql + "   WHEN OTHERS THEN" + vbCrLf
'    strSql = strSql + "      IF SQLCODE != -942 THEN " + vbCrLf
'    strSql = strSql + "         RAISE;" + vbCrLf
'    strSql = strSql + "      END IF;" + vbCrLf
'    strSql = strSql + "END;" + vbCrLf
'    strSql = strSql + "/" + vbCrLf
    ' ----------------------------------------------------------------------------------------
    'SQLServer
    strSql = strSql + "IF OBJECT_ID(N'" + tTbl.strPhysicsName + "', N'U') IS NOT NULL " + vbCrLf
    strSql = strSql + "DROP TABLE " + tTbl.strPhysicsName + vbCrLf
    strSql = strSql + "GO " + vbCrLf + vbCrLf
    
    ' ----------------------------------------------------------------------------------------
    
'
   'テーブル名を宣言
    If (tTbl.strSchema = "") Then
        strSql = strSql + "CREATE TABLE [dbo].[" + tTbl.strPhysicsName + "]" + vbCrLf
    Else
        strSql = strSql + "CREATE TABLE [dbo].[" + tTbl.strSchema + "." + tTbl.strPhysicsName + "]" + vbCrLf
    End If
    strSql = strSql + "(" + vbCrLf

    'カラム物理名と桁数を宣言
    Dim strSqlLine As String
    
    For i = 0 To lngLineCount - 1
        '初期化 #0-3[4]
        strSqlLine = "    "     '
        strBuf = ""

        '物理名 #4-34[31]
        strBuf = "[" + arrColumn(i).strPhysicsName + "]"
        strSqlLine = strSqlLine + strBuf + String(31 - Len(strBuf), " ")

        'データ型(桁数,小数桁数) #35-XX[-]
        If (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Or (arrColumn(i).strDataType = "int") Or (arrColumn(i).strDataType = "float") Then
            strBuf = "[" + arrColumn(i).strDataType + "]"
        ElseIf (arrColumn(i).strDataType = "NUMBER") Then
            strBuf = "[" + arrColumn(i).strDataType + "](" + CStr(arrColumn(i).lngLength) + "," + CStr(arrColumn(i).lngDecimal) + ")"
        ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
            strBuf = "[" + arrColumn(i).strDataType + "](" + CStr(arrColumn(i).lngLength) + ")"
        Else
            On Error GoTo Err0
        End If
        strSqlLine = strSqlLine + strBuf
        
        'デフォルト値
        If (arrColumn(i).strDefalutData <> "") Then
            strBuf = " DEFAULT " + arrColumn(i).strDefalutData
            strSqlLine = strSqlLine + strBuf
        End If

        '必須区分
        If (arrColumn(i).strRequiredFlag <> "") Then
            strBuf = " NOT NULL"
            strSqlLine = strSqlLine + strBuf
        End If
        
        
        ' 'カンマ
        ' If (i <> (lngLineCount - 1)) Then
        '     strSqlLine = strSqlLine + ","
        '     strSql = strSql + strSqlLine + vbCrLf
        ' Else
        '     strSql = strSql + strSqlLine + vbCrLf
        ' End If

        'カンマ
        strSqlLine = strSqlLine + ","
        strSql = strSql + strSqlLine + vbCrLf

    Next i

    '主キー
    strBuf = "    CONSTRAINT [PK_" + tTbl.strPhysicsName + "] "
    strSqlLine = strBuf

    strBuf = "PRIMARY KEY CLUSTERED " + vbCrLf + "(" + vbCrLf
    For i = 0 To lngLineCount - 1
        If (arrColumn(i).strPrimaryKey <> "") Then
            strBuf = strBuf + "       " + arrColumn(i).strPhysicsName + " ASC, " + vbCrLf
        End If
    Next i
    strBuf = Left(strBuf, (Len(strBuf) - 4)) + vbCrLf + "" '余計なカンマを削除
    strBuf = strBuf + "    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]" + vbCrLf
    strBuf = strBuf + ") ON [PRIMARY] " + vbCrLf
    strBuf = strBuf + "GO " + vbCrLf
    
    strSqlLine = strSqlLine + strBuf + vbCrLf
'    strSqlLine = strSqlLine + "        ENABLE" + vbCrLf
'    strSqlLine = strSqlLine + ")"
    strSql = strSql + strSqlLine + vbCrLf
    
'    'テーブルコメント
'    strSql = strSql + "/" + vbCrLf
'    If (tTbl.strSchema = "") Then
'        strSql = strSql + "COMMENT ON TABLE " + tTbl.strPhysicsName + " IS '" + tTbl.strLogicalName + "'" + vbCrLf
'    Else
'        strSql = strSql + "COMMENT ON TABLE " + tTbl.strSchema + "." + tTbl.strPhysicsName + " IS '" + tTbl.strLogicalName + "'" + vbCrLf
'    End If
    
    'カラムコメント
    ' --------------------------------------------------------------------------------------
    ' Oracle
'    strSql = strSql + "/" + vbCrLf
'    For i = 0 To lngLineCount - 1
'        If (tTbl.strSchema = "") Then
'            strBuf = "COMMENT ON COLUMN " + tTbl.strPhysicsName + "." + arrColumn(i).strPhysicsName _
'                + " IS '" + arrColumn(i).strLogicalName + "'" + vbCrLf
'        Else
'            strBuf = "COMMENT ON COLUMN " + tTbl.strSchema + "." + tTbl.strPhysicsName + "." + arrColumn(i).strPhysicsName _
'                + " IS '" + arrColumn(i).strLogicalName + "'" + vbCrLf
'        End If
'        strBuf = strBuf + "/" + vbCrLf
'        strSql = strSql + strBuf
    ' --------------------------------------------------------------------------------------
    ' SQLServer
    For i = 0 To lngLineCount - 1
        strSql = strSql + "EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'" _
            + arrColumn(i).strLogicalName + "' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'" _
            + tTbl.strPhysicsName + "', @level2type=N'COLUMN',@level2name=N'" + arrColumn(i).strPhysicsName + "'" + vbCrLf _
            + "Go" + vbCrLf
    
        
    Next i
    
   '--------------------------------------------------------------------------------------
    'スクリプトの出力
    Dim datFile As String
    datFile = ActiveWorkbook.Path + "\CreateTable_" + tTbl.strPhysicsName + "(" + tTbl.strLogicalName + ").sql"
    Open datFile For Output As #1

    Print #1, strSql

    Close #1
    
'    MsgBox (datFile + "に書き出しました")
 
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub




