Attribute VB_Name = "MakePythonDaoModule"


Const SCHMA_NAME As String = "" 'スキーマ名
Const COL_START_ROW As Long = 7 'カラム情報は7行目からスタート
Const TBL_START_ROW As Long = 5 'テーブル情報は5行目からスタート
Const SHEET_TABLE_LIST As String = "テーブル一覧表" 'テーブル一覧表のシート名
Const INDENT_SPACE As Long = 100 '論理名コメントを記載するための半角スペース

Public strSchemaName As String  'スキーマ名

'テーブル情報を元にDB処理(Dao)クラスのpythonファイルを出力
Public Sub outputPythonDao(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn)
On Error GoTo Err0
    'アクティブなシートを取得する。
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  'クラス名
    
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
    strText = strText + "import datetime" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + "from Dao.SqlConDao import SqlConDao" + vbCrLf
    strText = strText + "from Entity." + tTbl.strPhysicsName + "_Entity import " + tTbl.strPhysicsName + "_Entity" + vbCrLf
    strText = strText + vbCrLf

    strText = strText + "# " + tTbl.strLogicalName + "の処理 " + vbCrLf
    strText = strText + "class " + strClassName + ": " + vbCrLf
    strText = strText + "    CLASS_NAME: str = """ + strClassName + """ " + vbCrLf
    strText = strText + "    sqlCon: SqlConDao " + vbCrLf
    strText = strText + vbCrLf
    ' コンストラクタ
    strText = strText + "    def __init__(self):" + vbCrLf
    strText = strText + "        self.sqlCon = SqlConDao() " + vbCrLf
    strText = strText + "        pass " + vbCrLf
    strText = strText + vbCrLf
    ' デスストラクタ
    strText = strText + "    def __del__(self):" + vbCrLf
    strText = strText + "        if hasattr(self, ""conn""):" + vbCrLf
    strText = strText + "            self.conn.close() " + vbCrLf
    strText = strText + vbCrLf

    strText = strText + "    # private関数 ---------------------------------------------------------------------------------" + vbCrLf
    ' 登録処理-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Insert(tTbl, arrColumn)

    ' 更新処理-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Update(tTbl, arrColumn)

    strText = strText + "    # public関数 ---------------------------------------------------------------------------------" + vbCrLf
    ' 登録/更新まとめ処理-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Push(tTbl, arrColumn)

    ' 削除処理-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Delete(tTbl, arrColumn)
    
    ' 参照処理-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Select(tTbl, arrColumn)
    
        
   '--------------------------------------------------------------------------------------
    'スクリプトの出力
    Dim datFile As String
    datFile = ActiveWorkbook.Path + "\" + tTbl.strPhysicsName + "_Dao.py"
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

'テーブル情報を元にDB処理(Dao)クラス(python)のInsert文の処理の文字列を作成
Private Function outputPythonDao_Insert(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  'クラス名
    
    '対象件数(カラム数)を取得する。
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""

    ' 登録処理----------------------------------------------------------------------------------------
    strText = strText + "    # " + tTbl.strLogicalName + "への登録処理 " + vbCrLf
    strText = strText + "    def __insert" + tTbl.strPhysicsName + "(self, entityData: " + tTbl.strPhysicsName + "_Entity," + vbCrLf
    strText = strText + "                      sysDateTime: str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')):" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + "        try: " + vbCrLf
    strText = strText + "            sql: str = """"" + vbCrLf
    'Insert文の項目部分
    strText = strText + "            sql += ""INSERT INTO " + tTbl.strPhysicsName + "("" " + vbCrLf
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strText = strText + "            sql += ""      " + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        Else
            strText = strText + "            sql += ""     ," + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        End If
    Next i
    strText = strText + "            sql += "")SELECT"" + ""\n""" + vbCrLf
    
    'Insert文のデータ部分
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strTextLine = "            sql += ""     """
        Else
            strTextLine = "            sql += ""    ,"""
        End If
        'メンバー変数
        
        If (arrColumn(i).strPhysicsName = "UP_DT") Or (arrColumn(i).strPhysicsName = "MAKE_DT") Then
            'MAKE_DT/UP_DTの場合は、[sysDateTime]をセット
            strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
        ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
            'SHORI_KBNかつ、新規登録の場合は、1(新規登録)をセット
             strTextLine = strTextLine + " + ""'1'"" "
        ElseIf (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
            strTextLine = strTextLine + " + "" + self.sqlCon.sanitizeSQL(entityData.entityData." + "date" + arrColumn(i).strPhysicsName + ")"
        ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
            strTextLine = strTextLine + " + "" + self.sqlCon.sanitizeSQL(entityData.entityData." + "int" + arrColumn(i).strPhysicsName + ")"
        ElseIf (arrColumn(i).strDataType = "float") Then
            strTextLine = strTextLine + " + "" + self.sqlCon.sanitizeSQL(entityData.entityData." + "flt" + arrColumn(i).strPhysicsName + ")"
        ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
            strTextLine = strTextLine + " + "" + self.sqlCon.sanitizeSQL(entityData.entityData." + "str" + arrColumn(i).strPhysicsName + ")"
        Else
            On Error GoTo Err0
        End If
        strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
      
    Next i
    
    
    'Insert文の項目部分
    Dim tableName_J As String
    tableName_J = Left(tTbl.strPhysicsName, 1) + "J" + Right(tTbl.strPhysicsName, Len(tTbl.strPhysicsName) - 1)
    
    strText = strText + "" + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "            sql += ""INSERT INTO " + tableName_J + "("" " + vbCrLf
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strText = strText + "            sql += ""     " + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        Else
            strText = strText + "            sql += ""    ," + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        End If
    Next i
    strText = strText + "            sql += "")SELECT"" + ""\n""" + vbCrLf
    
    'Insert文のデータ部分
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strTextLine = "            sql += ""     """
        Else
            strTextLine = "            sql += ""    ,"""
        End If
        'メンバー変数
        
        If (arrColumn(i).strPhysicsName = "UP_DT") Or (arrColumn(i).strPhysicsName = "MAKE_DT") Then
            'MAKE_DT/UP_DTの場合は、[sysDateTime]をセット
            strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
        ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
            'SHORI_KBNかつ、新規登録の場合は、1(新規登録)をセット
             strTextLine = strTextLine + " + ""'1'"" "
        ElseIf (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
            strTextLine = strTextLine + " + "" + self.sqlCon.sanitizeSQL(entityData.entityData." + "date" + arrColumn(i).strPhysicsName + ")"
        ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
            strTextLine = strTextLine + " + "" + self.sqlCon.sanitizeSQL(entityData.entityData." + "int" + arrColumn(i).strPhysicsName + ")"
        ElseIf (arrColumn(i).strDataType = "float") Then
            strTextLine = strTextLine + " + "" + self.sqlCon.sanitizeSQL(entityData.entityData." + "flt" + arrColumn(i).strPhysicsName + ")"
        ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
            strTextLine = strTextLine + " + "" + self.sqlCon.sanitizeSQL(entityData.entityData." + "str" + arrColumn(i).strPhysicsName + ")"
        Else
            On Error GoTo Err0
        End If
        
        strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
      
    Next i
    
    '後処理
    strText = strText + "" + vbCrLf
    strText = strText + "            self.sqlCon.executeOnlySql(sql)" + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "            return True" + vbCrLf
    strText = strText + "        except Exception as e:" + vbCrLf
    strText = strText + "            print(""Error:"" + str(e))" + vbCrLf
    strText = strText + "            raise e " + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "" + vbCrLf


    outputPythonDao_Insert = strText
    Exit Function
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Function


'テーブル情報を元にDB処理(Dao)クラス(python)のUpdate文の処理の文字列を作成
Private Function outputPythonDao_Update(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  'クラス名
    
    '対象件数(カラム数)を取得する。
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""

    ' 更新処理-----------------------------------------------------------------------------------------
    strText = strText + "    # " + tTbl.strLogicalName + "への更新処理 " + vbCrLf
    strText = strText + "    def __update" + tTbl.strPhysicsName + "(self, entityData: " + tTbl.strPhysicsName + "_Entity," + vbCrLf
    strText = strText + "                      sysDateTime: str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')):" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + "        try: " + vbCrLf
    strText = strText + "            sql: str = """"" + vbCrLf
    'Update文のデータ部分
    strText = strText + "            sql += ""UPDATE " + tTbl.strPhysicsName + """ + ""\n"" " + vbCrLf
    strText = strText + "            sql += ""SET """ + vbCrLf
    
    Dim count As Integer
    
    'SET句部分
    count = 0
    For i = 0 To lngLineCount - 1
        If arrColumn(i).strPrimaryKey = "" Then
        '主キーの項目以外を一斉更新。
            strTextLine = ""
            If count = 0 Then
                strTextLine = strTextLine + "            sql += "" " + arrColumn(i).strPhysicsName + "="""
            Else
                strTextLine = strTextLine + "            sql += ""," + arrColumn(i).strPhysicsName + "="""
            End If
            
            
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'MAKE_DT/UP_DTの場合は、[sysDateTime]をセット
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBNかつ、新規登録の場合は、2(訂正)をセット
                 strTextLine = strTextLine + " + ""'2'"" "
            ElseIf (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
                strTextLine = strTextLine + " + ""'"" " + " + str(entityData." + "date" + arrColumn(i).strPhysicsName + ") " + (" + ""'"" ")
            ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
                strTextLine = strTextLine + " + str(entityData." + "int" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "float") Then
                strTextLine = strTextLine + " + str(entityData." + "flt" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
                strTextLine = strTextLine + " + ""N'"" " + " + str(entityData." + "str" + arrColumn(i).strPhysicsName + ") " + (" + ""'"" ")
            Else
                On Error GoTo Err0
            End If
            
            strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
            count = count + 1
        End If
    Next i
    
    'WHERE句部分
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
    
        If arrColumn(i).strPrimaryKey <> "" Then
        '主キーの項目以外を一斉更新。
            strTextLine = ""
            strTextLine = strTextLine + "            sql += "" AND " + arrColumn(i).strPhysicsName + "="""
            
            
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'MAKE_DT/UP_DTの場合は、[sysDateTime]をセット
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBNかつ、新規登録の場合は、2(訂正)をセット
                 strTextLine = strTextLine + " + ""'2'"" "
            ElseIf (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
                strTextLine = strTextLine + " + ""'"" " + " + str(entityData." + "date" + arrColumn(i).strPhysicsName + ") " + (" + ""'"" ")
            ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
                strTextLine = strTextLine + " + str(entityData." + "int" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "float") Then
                strTextLine = strTextLine + " + str(entityData." + "flt" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
                strTextLine = strTextLine + " + ""N'"" " + " + str(entityData." + "str" + arrColumn(i).strPhysicsName + ") " + (" + ""'"" ")
            Else
                On Error GoTo Err0
            End If
            
            strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
            count = count + 1
        End If
    Next i
    
    'Insert文の項目部分
    tableName_J = Left(tTbl.strPhysicsName, 1) + "J" + Right(tTbl.strPhysicsName, Len(tTbl.strPhysicsName) - 1)
    
    strText = strText + "" + vbCrLf
    strText = strText + "            sql += ""INSERT INTO " + tableName_J + "("" " + vbCrLf
    '項目指定部分
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strText = strText + "            sql += ""   " + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        Else
            strText = strText + "            sql += ""  ," + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        End If
    Next i
    strText = strText + "            sql += "" )SELECT " + "\n""" + vbCrLf
    '項目指定部分
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strText = strText + "            sql += ""   " + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        Else
            strText = strText + "            sql += ""  ," + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        End If
    Next i
    strText = strText + "            sql += ""FROM " + tTbl.strPhysicsName + " \n""" + vbCrLf
    'WHERE句部分
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
    
        If arrColumn(i).strPrimaryKey <> "" Then
        '主キーの項目以外を一斉更新。
            strTextLine = ""
            strTextLine = strTextLine + "            sql += "" AND " + arrColumn(i).strPhysicsName + "="""
                  
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'MAKE_DT/UP_DTの場合は、[sysDateTime]をセット
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBNかつ、新規登録の場合は、2(訂正)をセット
                 strTextLine = strTextLine + " + ""'2'"" "
            ElseIf (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
                strTextLine = strTextLine + " + ""'"" " + "+ str(entityData." + "date" + arrColumn(i).strPhysicsName + ")" + (" + ""'"" ")
            ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
                strTextLine = strTextLine + " + str(entityData." + "int" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "float") Then
                strTextLine = strTextLine + " + str(entityData." + "flt" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
                strTextLine = strTextLine + " + ""N'"" " + "+ str(entityData." + "str" + arrColumn(i).strPhysicsName + ")" + (" + ""'"" ")
            Else
                On Error GoTo Err0
            End If
            
            strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
            count = count + 1
        End If
    Next i
    
    
    '後処理
    strText = strText + "" + vbCrLf
    strText = strText + "            self.sqlCon.executeOnlySql(sql)" + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "            return True" + vbCrLf
    strText = strText + "        except Exception as e:" + vbCrLf
    strText = strText + "            print(""Error:"" + str(e))" + vbCrLf
    strText = strText + "            raise e " + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "" + vbCrLf


    outputPythonDao_Update = strText
    Exit Function
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Function

'テーブル情報を元にDB処理(Dao)クラス(python)の登録/更新まとめ処理
Private Function outputPythonDao_Push(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  'クラス名
    
    '対象件数(カラム数)を取得する。
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""

    strText = strText + "    def push" + tTbl.strPhysicsName + "(self, entityData: " + tTbl.strPhysicsName + "_Entity," + vbCrLf
    strText = strText + "                    sysDateTime: str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')):" + vbCrLf
    strText = strText + "        try:" + vbCrLf
    strText = strText + "            # 登録済みか確認し、もし登録済みであれば更新(Update)、未登録なら登録(Insert)する。" + vbCrLf
    strText = strText + "            entityTemp: " + tTbl.strPhysicsName + "_Entity = " + tTbl.strPhysicsName + "_Entity()" + vbCrLf
    strText = strText + "            entityTemp = self.select" + tTbl.strPhysicsName + "("

    Dim count As Integer
    count = 0
    For i = 0 To lngLineCount - 1
        If arrColumn(i).strPrimaryKey <> "" Then
            If count = 0 Then
                strText = strText + "entityData."
            Else
                strText = strText + ",entityData."
            End If
            
           If (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
                strText = strText + "date" + arrColumn(i).strPhysicsName
            ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
                strText = strText + "int" + arrColumn(i).strPhysicsName
            ElseIf (arrColumn(i).strDataType = "float") Then
                 strText = strText + "flt" + arrColumn(i).strPhysicsName
            ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
                 strText = strText + "str" + arrColumn(i).strPhysicsName
            Else
                On Error GoTo Err0
            End If
                      
            count = count + 1
        End If
    Next i
    strText = strText + ")" + vbCrLf
    strText = strText + "            if entityTemp is None:" + vbCrLf
    strText = strText + "                self.__insert" + tTbl.strPhysicsName + "(entityData,sysDateTime)" + vbCrLf
    strText = strText + "            else:" + vbCrLf
    strText = strText + "                entityData.dateMAKE_DT = entityTemp.dateMAKE_DT" + vbCrLf
    strText = strText + "                self.__update" + tTbl.strPhysicsName + "(entityData,sysDateTime)" + vbCrLf
    strText = strText + "            return True" + vbCrLf
    strText = strText + "        except Exception as e:" + vbCrLf
    strText = strText + "            print(self.CLASS_NAME + ""Error:"" + str(e))" + vbCrLf
    strText = strText + "            raise e" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + vbCrLf

    outputPythonDao_Push = strText
    Exit Function
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Function

'テーブル情報を元にDB処理(Dao)クラス(python)のDelete文の処理の文字列を作成
Private Function outputPythonDao_Delete(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  'クラス名
    
    '対象件数(カラム数)を取得する。
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""


    ' 削除処理-----------------------------------------------------------------------------------------
    strText = strText + "    # " + tTbl.strLogicalName + "への削除処理 " + vbCrLf
    strText = strText + "    def delete" + tTbl.strPhysicsName + "(self, entityData: " + tTbl.strPhysicsName + "_Entity," + vbCrLf
    strText = strText + "                      sysDateTime: str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')):" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + "        try: " + vbCrLf
    strText = strText + "            sql: str = """"" + vbCrLf
    
    Dim count As Integer
    'Insert文の項目部分
    tableName_J = Left(tTbl.strPhysicsName, 1) + "J" + Right(tTbl.strPhysicsName, Len(tTbl.strPhysicsName) - 1)
    
    strText = strText + "" + vbCrLf
    strText = strText + "            sql += ""INSERT INTO " + tableName_J + "("" " + vbCrLf
    '項目指定部分
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strText = strText + "            sql += ""     " + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        Else
            strText = strText + "            sql += ""    ," + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        End If
    Next i
    'SELECT句
    strText = strText + "            sql += "")SELECT"" + ""\n""" + "  " + vbCrLf
    '項目指定部分
    For i = 0 To lngLineCount - 1
        strTextLine = ""
        '主キーの項目を指定。
        If i = 0 Then
            strTextLine = strTextLine + "            sql += ""     "
        Else
            strTextLine = strTextLine + "            sql += ""    ,"
        End If
        
        
        If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'UP_DTの場合は、[sysDateTime]をセット
                strTextLine = strTextLine + "'""  + sysDateTime  + ""' "
        ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBNかつ、新規登録の場合は、9(削除)をセット
                 strTextLine = strTextLine + " '9' "
        Else
                  strTextLine = strTextLine + arrColumn(i).strPhysicsName
        
        End If
         strText = strText + strTextLine + """  " + ("+ "" \n"" ") + vbCrLf
    Next i
    strText = strText + "            sql += ""FROM " + tTbl.strPhysicsName + " \n """ + vbCrLf
    'WHERE句部分
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
    
        If arrColumn(i).strPrimaryKey <> "" Then
        '主キーの項目を指定。
            strTextLine = ""
            strTextLine = strTextLine + "            sql += "" AND " + arrColumn(i).strPhysicsName + "="""
                  
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'UP_DTの場合は、[sysDateTime]をセット
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBNかつ、新規登録の場合は、9(削除)をセット
                 strTextLine = strTextLine + " + ""'9'"" "
            ElseIf (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
                strTextLine = strTextLine + " + ""'"" " + " + str(entityData." + "date" + arrColumn(i).strPhysicsName + ") " + (" + ""'"" ")
            ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
                strTextLine = strTextLine + " + str(entityData." + "int" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "float") Then
                strTextLine = strTextLine + " + str(entityData." + "flt" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
                strTextLine = strTextLine + " + ""N'"" " + "+ str(entityData." + "str" + arrColumn(i).strPhysicsName + ")" + (" + ""'"" ")
            Else
                On Error GoTo Err0
            End If
            
            strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
            count = count + 1
        End If
    Next i
    
    '削除処理-----------------------------------------------------------------------------------
    strText = strText + "" + vbCrLf
    strText = strText + "            sql += ""DELETE FROM " + tTbl.strPhysicsName + " \n """ + vbCrLf
    'WHERE句部分
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
        If arrColumn(i).strPrimaryKey <> "" Then
        '主キーの項目を指定。
            strTextLine = ""
            strTextLine = strTextLine + "            sql += "" AND " + arrColumn(i).strPhysicsName + "="""
                  
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'UP_DTの場合は、[sysDateTime]をセット
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBNかつ、新規登録の場合は、9(削除)をセット
                 strTextLine = strTextLine + " + ""'9'"" "
            ElseIf (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
                strTextLine = strTextLine + " + ""'"" " + " + str(entityData." + "date" + arrColumn(i).strPhysicsName + ") " + (" + ""'"" ")
            ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
                strTextLine = strTextLine + " + str(entityData." + "int" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "float") Then
                strTextLine = strTextLine + " + str(entityData." + "flt" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
                strTextLine = strTextLine + " + ""N'"" " + " + str(entityData." + "str" + arrColumn(i).strPhysicsName + ") " + (" + ""'"" ")
            Else
                On Error GoTo Err0
            End If
            
            strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
            count = count + 1
        End If
    Next i
    
    '後処理
    strText = strText + "" + vbCrLf
    strText = strText + "            self.sqlCon.executeOnlySql(sql)" + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "            return True" + vbCrLf
    strText = strText + "        except Exception as e:" + vbCrLf
    strText = strText + "            print(""Error:"" + str(e))" + vbCrLf
    strText = strText + "            raise e " + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "" + vbCrLf


    outputPythonDao_Delete = strText
    Exit Function
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Function

'テーブル情報を元にDB処理(Dao)クラス(python)のSelect文の処理の文字列を作成
Private Function outputPythonDao_Select(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    'アクティブなシートを取得する。
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  'クラス名
    
    '対象件数(カラム数)を取得する。
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""

    ' 参照処理-----------------------------------------------------------------------------------------
    strText = strText + "    # " + tTbl.strLogicalName + "への参照処理 " + vbCrLf
    strText = strText + "    def select" + tTbl.strPhysicsName + "(self"
    ' 関数の引数部分を設定
    For i = 0 To lngLineCount - 1
        strTextLine = ""
        If arrColumn(i).strPrimaryKey <> "" Then
            
            If (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
                strTextLine = strTextLine + ",date" + arrColumn(i).strPhysicsName + ":" + "str"
            ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
                strTextLine = strTextLine + ",int" + arrColumn(i).strPhysicsName + ":" + "int"
            ElseIf (arrColumn(i).strDataType = "float") Then
               strTextLine = strTextLine + ",flt" + arrColumn(i).strPhysicsName + ":" + "float"
            ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
               strTextLine = strTextLine + ",str" + arrColumn(i).strPhysicsName + ":" + "str"
            Else
                On Error GoTo Err0
            End If
        End If
        
        strText = strText + strTextLine
    Next i
    strText = strText + "):" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + "        try: " + vbCrLf
    strText = strText + "            sql: str = """"" + vbCrLf
    
    ' SELECT文-----------------------------------------------------------------------------------------
    ' SELECT句
    strText = strText + "            sql += ""SELECT"" + ""\n""" + "  " + vbCrLf
    '項目指定部分
    For i = 0 To lngLineCount - 1
        strTextLine = ""
        '主キーの項目を指定。
        If i = 0 Then
            strTextLine = strTextLine + "            sql += ""     "
        Else
            strTextLine = strTextLine + "            sql += ""    ,"
        End If
        ' 項目名の物理名を記述
        strTextLine = strTextLine + arrColumn(i).strPhysicsName
        strTextLine = strTextLine + """ + "" \n"" "
        ' 項目名の論理名を記述
        If Len(strTextLine) < INDENT_SPACE Then
            strText = strText + strTextLine + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
        Else
            strText = strText + strTextLine + " #" + arrColumn(i).strLogicalName + vbCrLf
        End If
        ' strText = strText + strTextLine + """ + "" \n"" " + vbCrLf
        ' strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
    Next i
    ' FROM句
    strText = strText + "            sql += ""FROM " + tTbl.strPhysicsName + " \n """ + vbCrLf
    ' WHERE句部分
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
    
        If arrColumn(i).strPrimaryKey <> "" Then
        '主キーの項目を指定。
            strTextLine = ""
            strTextLine = strTextLine + "            sql += "" AND " + arrColumn(i).strPhysicsName + "="""
                  
            If (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
                strTextLine = strTextLine + " + ""'"" " + "+ str(" + "date" + arrColumn(i).strPhysicsName + ")" + (" + ""'"" ")
            ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
                strTextLine = strTextLine + " + int(" + "int" + arrColumn(i).strPhysicsName + ")"
            ElseIf (arrColumn(i).strDataType = "float") Then
                strTextLine = strTextLine + " + " + "flt" + arrColumn(i).strPhysicsName + ""
            ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
                strTextLine = strTextLine + " + ""N'"" " + "+ str(" + "str" + arrColumn(i).strPhysicsName + ")" + (" + ""'"" ")
            Else
                On Error GoTo Err0
            End If
            
            strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
            count = count + 1
        End If
    Next i
    
    'SQLを実行
    strText = strText + "" + vbCrLf
    strText = strText + "            # SQLを実行" + vbCrLf
    strText = strText + "            df = self.sqlCon.executeSql(sql)" + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "            # SQL実行結果を取得" + vbCrLf
    strText = strText + "            entityResult: " + tTbl.strPhysicsName + "_Entity" + " = " + tTbl.strPhysicsName + "_Entity()" + vbCrLf
    strText = strText + "            i = 0 " + vbCrLf
    strText = strText + "            if len(df) > 0: " + vbCrLf
    'SQLを実行結果を取得
    For i = 0 To lngLineCount - 1
        strTextLine = ""
        strTextLine = strTextLine + "                entityResult."
        If (arrColumn(i).strDataType = "DATE") Or (arrColumn(i).strDataType = "datetime") Then
            strTextLine = strTextLine + "date" + arrColumn(i).strPhysicsName + " = df.loc[0,""" + arrColumn(i).strPhysicsName + """]"
        ElseIf (arrColumn(i).strDataType = "NUMBER") Or (arrColumn(i).strDataType = "int") Then
            strTextLine = strTextLine + "int" + arrColumn(i).strPhysicsName + " = df.loc[0,""" + arrColumn(i).strPhysicsName + """]"
        ElseIf (arrColumn(i).strDataType = "float") Then
            strTextLine = strTextLine + "flt" + arrColumn(i).strPhysicsName + " = df.loc[0,""" + arrColumn(i).strPhysicsName + """]"
        ElseIf (arrColumn(i).strDataType = "VARCHAR2") Or (arrColumn(i).strDataType = "nvarchar") Or (arrColumn(i).strDataType = "varchar") Then
            strTextLine = strTextLine + "str" + arrColumn(i).strPhysicsName + " = df.loc[0,""" + arrColumn(i).strPhysicsName + """]"
        Else
            On Error GoTo Err0
        End If
        strText = strText + strTextLine + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
        
    Next i
    strText = strText + "            else: " + vbCrLf
    strText = strText + "                entityResult = None" + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "            return entityResult" + vbCrLf
    strText = strText + "        except Exception as e:" + vbCrLf
    strText = strText + "            print(""Error:"" + str(e))" + vbCrLf
    strText = strText + "            raise e " + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "" + vbCrLf
    outputPythonDao_Select = strText
    Exit Function
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Function

