Attribute VB_Name = "MakePythonEntityModule"


Const SCHMA_NAME As String = "" 'スキーマ名
Const COL_START_ROW As Long = 7 'カラム情報は7行目からスタート
Const TBL_START_ROW As Long = 5 'テーブル情報は5行目からスタート
Const SHEET_TABLE_LIST As String = "テーブル一覧表" 'テーブル一覧表のシート名


Public strSchemaName As String  'スキーマ名

'テーブル情報を元に要素(Entity)クラスのpythonファイルを出力
Public Sub outputPythonEntity(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn)
On Error GoTo Err0
    'アクティブなシートを取得する。
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Entity"  'クラス名
    
    '対象件数(カラム数)を取得する。
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    '--------------------------------------------------------------------------------------
    'スクリプトを組み立てる。
       Dim strText As String
    strText = ""
    

    ' --------------------------------------------------------------------------------------
    ' ヘッダ情報
    strText = strText + "import copy" + vbCrLf
    strText = strText + "import datetime" + vbCrLf
    strText = strText + "# ------------------------------------------------------------------" + vbCrLf
    strText = strText + "# 定数" + vbCrLf
    strText = strText + "DB_DRIBER: str = ""{ODBC Driver 13 for SQL Server}""" + vbCrLf
    strText = strText + "" + vbCrLf
    
    strText = strText + "# -------------------------------------------------------------------" + vbCrLf
    strText = strText + "# クラス（要素情報）" + vbCrLf
    strText = strText + "# 参照する" + tTbl.strLogicalName + "情報(" + tTbl.strPhysicsName + ")" + vbCrLf
    strText = strText + "class " + strClassName + ":" + vbCrLf

    ' ----------------------------------------------------------------------------------------
    strText = strText + "    # クラス変数" + vbCrLf
    For i = 0 To lngLineCount - 1
        '初期化 #0-3[4]
        strTextLine = "    "     '
        strBuf = ""

        '変数名 #4-34[31]
        If (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
            strBuf = "date" + arrColumn(i).strPhysicsName + "= datetime.time(0,0,0)"
        ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
            strBuf = "int" + arrColumn(i).strPhysicsName + "= 0"
        ElseIf (arrColumn(i).strDataType = "float") Then
            strBuf = "flt" + arrColumn(i).strPhysicsName + "= 0.00"
        ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
            strBuf = "str" + arrColumn(i).strPhysicsName + "= """""
        Else
            On Error GoTo Err0
        End If
        strText = strText + strTextLine + strBuf + String(41 - Len(strBuf), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
    Next i
    strText = strText + vbCrLf

   ' ----------------------------------------------------------------------------------------
    strText = strText + "    def __init__(self):" + vbCrLf
    strText = strText + "        pass" + vbCrLf
    strText = strText + vbCrLf
    
   ' ----------------------------------------------------------------------------------------
    strText = strText + "    def print(self):" + vbCrLf
    strText = strText + "        print(""/*-" + strClassName + "-------------------------*/"")" + vbCrLf
    For i = 0 To lngLineCount - 1
        strTextLine = "        print("""
        strBuf = ""

        '変数名 #8-34[31]
        If (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
            strBuf = "date" + arrColumn(i).strPhysicsName
        ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
            strBuf = "int" + arrColumn(i).strPhysicsName
        ElseIf (arrColumn(i).strDataType = "float") Then
            strBuf = "flt" + arrColumn(i).strPhysicsName
        ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
            strBuf = "str" + arrColumn(i).strPhysicsName
        Else
            On Error GoTo Err0
        End If
        strText = strText + strTextLine + strBuf + String(21 - Len(strBuf), " ") + " ="" + str(self." + strBuf + "))  #" + arrColumn(i).strLogicalName + vbCrLf


    Next i
    strText = strText + vbCrLf
        
   '--------------------------------------------------------------------------------------
    'スクリプトの出力
    Dim datFile As String
    datFile = ActiveWorkbook.Path + "\" + tTbl.strPhysicsName + "_Entity.py"
'    Open datFile For Output As #1
'
'    Print #1, strText
'
'    Close #1
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText strText
        .SaveToFile datFile, 2
        .Close
    End With
'    MsgBox (datFile + "に書き出しました")
 
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub

