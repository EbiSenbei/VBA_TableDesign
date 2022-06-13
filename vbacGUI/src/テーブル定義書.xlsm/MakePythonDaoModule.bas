Attribute VB_Name = "MakePythonDaoModule"


Const SCHMA_NAME As String = "" '�X�L�[�}��
Const COL_START_ROW As Long = 7 '�J��������7�s�ڂ���X�^�[�g
Const TBL_START_ROW As Long = 5 '�e�[�u������5�s�ڂ���X�^�[�g
Const SHEET_TABLE_LIST As String = "�e�[�u���ꗗ�\" '�e�[�u���ꗗ�\�̃V�[�g��
Const INDENT_SPACE As Long = 100 '�_�����R�����g���L�ڂ��邽�߂̔��p�X�y�[�X

Public strSchemaName As String  '�X�L�[�}��

'�e�[�u����������DB����(Dao)�N���X��python�t�@�C�����o��
Public Sub outputPythonDao(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn)
On Error GoTo Err0
    '�A�N�e�B�u�ȃV�[�g���擾����B
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  '�N���X��
    
    '�Ώی���(�J������)���擾����B
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    '--------------------------------------------------------------------------------------
    '�X�N���v�g��g�ݗ��Ă�B
    Dim strText As String
    strText = ""
    
    ' --------------------------------------------------------------------------------------
    ' �w�b�_���
    strText = strText + "import datetime" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + "from Dao.SqlConDao import SqlConDao" + vbCrLf
    strText = strText + "from Entity." + tTbl.strPhysicsName + "_Entity import " + tTbl.strPhysicsName + "_Entity" + vbCrLf
    strText = strText + vbCrLf

    strText = strText + "# " + tTbl.strLogicalName + "�̏��� " + vbCrLf
    strText = strText + "class " + strClassName + ": " + vbCrLf
    strText = strText + "    CLASS_NAME: str = """ + strClassName + """ " + vbCrLf
    strText = strText + "    sqlCon: SqlConDao " + vbCrLf
    strText = strText + vbCrLf
    ' �R���X�g���N�^
    strText = strText + "    def __init__(self):" + vbCrLf
    strText = strText + "        self.sqlCon = SqlConDao() " + vbCrLf
    strText = strText + "        pass " + vbCrLf
    strText = strText + vbCrLf
    ' �f�X�X�g���N�^
    strText = strText + "    def __del__(self):" + vbCrLf
    strText = strText + "        if hasattr(self, ""conn""):" + vbCrLf
    strText = strText + "            self.conn.close() " + vbCrLf
    strText = strText + vbCrLf

    strText = strText + "    # private�֐� ---------------------------------------------------------------------------------" + vbCrLf
    ' �o�^����-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Insert(tTbl, arrColumn)

    ' �X�V����-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Update(tTbl, arrColumn)

    strText = strText + "    # public�֐� ---------------------------------------------------------------------------------" + vbCrLf
    ' �o�^/�X�V�܂Ƃߏ���-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Push(tTbl, arrColumn)

    ' �폜����-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Delete(tTbl, arrColumn)
    
    ' �Q�Ə���-----------------------------------------------------------------------------------------
    strText = strText + outputPythonDao_Select(tTbl, arrColumn)
    
        
   '--------------------------------------------------------------------------------------
    '�X�N���v�g�̏o��
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
'    MsgBox (datFile + "�ɏ����o���܂���")
 
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub

'�e�[�u����������DB����(Dao)�N���X(python)��Insert���̏����̕�������쐬
Private Function outputPythonDao_Insert(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  '�N���X��
    
    '�Ώی���(�J������)���擾����B
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""

    ' �o�^����----------------------------------------------------------------------------------------
    strText = strText + "    # " + tTbl.strLogicalName + "�ւ̓o�^���� " + vbCrLf
    strText = strText + "    def __insert" + tTbl.strPhysicsName + "(self, entityData: " + tTbl.strPhysicsName + "_Entity," + vbCrLf
    strText = strText + "                      sysDateTime: str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')):" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + "        try: " + vbCrLf
    strText = strText + "            sql: str = """"" + vbCrLf
    'Insert���̍��ڕ���
    strText = strText + "            sql += ""INSERT INTO " + tTbl.strPhysicsName + "("" " + vbCrLf
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strText = strText + "            sql += ""      " + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        Else
            strText = strText + "            sql += ""     ," + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        End If
    Next i
    strText = strText + "            sql += "")SELECT"" + ""\n""" + vbCrLf
    
    'Insert���̃f�[�^����
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strTextLine = "            sql += ""     """
        Else
            strTextLine = "            sql += ""    ,"""
        End If
        '�����o�[�ϐ�
        
        If (arrColumn(i).strPhysicsName = "UP_DT") Or (arrColumn(i).strPhysicsName = "MAKE_DT") Then
            'MAKE_DT/UP_DT�̏ꍇ�́A[sysDateTime]���Z�b�g
            strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
        ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
            'SHORI_KBN���A�V�K�o�^�̏ꍇ�́A1(�V�K�o�^)���Z�b�g
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
    
    
    'Insert���̍��ڕ���
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
    
    'Insert���̃f�[�^����
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strTextLine = "            sql += ""     """
        Else
            strTextLine = "            sql += ""    ,"""
        End If
        '�����o�[�ϐ�
        
        If (arrColumn(i).strPhysicsName = "UP_DT") Or (arrColumn(i).strPhysicsName = "MAKE_DT") Then
            'MAKE_DT/UP_DT�̏ꍇ�́A[sysDateTime]���Z�b�g
            strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
        ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
            'SHORI_KBN���A�V�K�o�^�̏ꍇ�́A1(�V�K�o�^)���Z�b�g
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
    
    '�㏈��
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


'�e�[�u����������DB����(Dao)�N���X(python)��Update���̏����̕�������쐬
Private Function outputPythonDao_Update(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  '�N���X��
    
    '�Ώی���(�J������)���擾����B
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""

    ' �X�V����-----------------------------------------------------------------------------------------
    strText = strText + "    # " + tTbl.strLogicalName + "�ւ̍X�V���� " + vbCrLf
    strText = strText + "    def __update" + tTbl.strPhysicsName + "(self, entityData: " + tTbl.strPhysicsName + "_Entity," + vbCrLf
    strText = strText + "                      sysDateTime: str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')):" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + "        try: " + vbCrLf
    strText = strText + "            sql: str = """"" + vbCrLf
    'Update���̃f�[�^����
    strText = strText + "            sql += ""UPDATE " + tTbl.strPhysicsName + """ + ""\n"" " + vbCrLf
    strText = strText + "            sql += ""SET """ + vbCrLf
    
    Dim count As Integer
    
    'SET�啔��
    count = 0
    For i = 0 To lngLineCount - 1
        If arrColumn(i).strPrimaryKey = "" Then
        '��L�[�̍��ڈȊO����čX�V�B
            strTextLine = ""
            If count = 0 Then
                strTextLine = strTextLine + "            sql += "" " + arrColumn(i).strPhysicsName + "="""
            Else
                strTextLine = strTextLine + "            sql += ""," + arrColumn(i).strPhysicsName + "="""
            End If
            
            
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'MAKE_DT/UP_DT�̏ꍇ�́A[sysDateTime]���Z�b�g
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBN���A�V�K�o�^�̏ꍇ�́A2(����)���Z�b�g
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
    
    'WHERE�啔��
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
    
        If arrColumn(i).strPrimaryKey <> "" Then
        '��L�[�̍��ڈȊO����čX�V�B
            strTextLine = ""
            strTextLine = strTextLine + "            sql += "" AND " + arrColumn(i).strPhysicsName + "="""
            
            
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'MAKE_DT/UP_DT�̏ꍇ�́A[sysDateTime]���Z�b�g
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBN���A�V�K�o�^�̏ꍇ�́A2(����)���Z�b�g
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
    
    'Insert���̍��ڕ���
    tableName_J = Left(tTbl.strPhysicsName, 1) + "J" + Right(tTbl.strPhysicsName, Len(tTbl.strPhysicsName) - 1)
    
    strText = strText + "" + vbCrLf
    strText = strText + "            sql += ""INSERT INTO " + tableName_J + "("" " + vbCrLf
    '���ڎw�蕔��
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strText = strText + "            sql += ""   " + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        Else
            strText = strText + "            sql += ""  ," + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        End If
    Next i
    strText = strText + "            sql += "" )SELECT " + "\n""" + vbCrLf
    '���ڎw�蕔��
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strText = strText + "            sql += ""   " + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        Else
            strText = strText + "            sql += ""  ," + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        End If
    Next i
    strText = strText + "            sql += ""FROM " + tTbl.strPhysicsName + " \n""" + vbCrLf
    'WHERE�啔��
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
    
        If arrColumn(i).strPrimaryKey <> "" Then
        '��L�[�̍��ڈȊO����čX�V�B
            strTextLine = ""
            strTextLine = strTextLine + "            sql += "" AND " + arrColumn(i).strPhysicsName + "="""
                  
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'MAKE_DT/UP_DT�̏ꍇ�́A[sysDateTime]���Z�b�g
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBN���A�V�K�o�^�̏ꍇ�́A2(����)���Z�b�g
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
    
    
    '�㏈��
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

'�e�[�u����������DB����(Dao)�N���X(python)�̓o�^/�X�V�܂Ƃߏ���
Private Function outputPythonDao_Push(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  '�N���X��
    
    '�Ώی���(�J������)���擾����B
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""

    strText = strText + "    def push" + tTbl.strPhysicsName + "(self, entityData: " + tTbl.strPhysicsName + "_Entity," + vbCrLf
    strText = strText + "                    sysDateTime: str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')):" + vbCrLf
    strText = strText + "        try:" + vbCrLf
    strText = strText + "            # �o�^�ς݂��m�F���A�����o�^�ς݂ł���΍X�V(Update)�A���o�^�Ȃ�o�^(Insert)����B" + vbCrLf
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

'�e�[�u����������DB����(Dao)�N���X(python)��Delete���̏����̕�������쐬
Private Function outputPythonDao_Delete(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  '�N���X��
    
    '�Ώی���(�J������)���擾����B
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""


    ' �폜����-----------------------------------------------------------------------------------------
    strText = strText + "    # " + tTbl.strLogicalName + "�ւ̍폜���� " + vbCrLf
    strText = strText + "    def delete" + tTbl.strPhysicsName + "(self, entityData: " + tTbl.strPhysicsName + "_Entity," + vbCrLf
    strText = strText + "                      sysDateTime: str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')):" + vbCrLf
    strText = strText + vbCrLf
    strText = strText + "        try: " + vbCrLf
    strText = strText + "            sql: str = """"" + vbCrLf
    
    Dim count As Integer
    'Insert���̍��ڕ���
    tableName_J = Left(tTbl.strPhysicsName, 1) + "J" + Right(tTbl.strPhysicsName, Len(tTbl.strPhysicsName) - 1)
    
    strText = strText + "" + vbCrLf
    strText = strText + "            sql += ""INSERT INTO " + tableName_J + "("" " + vbCrLf
    '���ڎw�蕔��
    For i = 0 To lngLineCount - 1
        If i = 0 Then
            strText = strText + "            sql += ""     " + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        Else
            strText = strText + "            sql += ""    ," + arrColumn(i).strPhysicsName + """ + ""\n"" " + vbCrLf
        End If
    Next i
    'SELECT��
    strText = strText + "            sql += "")SELECT"" + ""\n""" + "  " + vbCrLf
    '���ڎw�蕔��
    For i = 0 To lngLineCount - 1
        strTextLine = ""
        '��L�[�̍��ڂ��w��B
        If i = 0 Then
            strTextLine = strTextLine + "            sql += ""     "
        Else
            strTextLine = strTextLine + "            sql += ""    ,"
        End If
        
        
        If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'UP_DT�̏ꍇ�́A[sysDateTime]���Z�b�g
                strTextLine = strTextLine + "'""  + sysDateTime  + ""' "
        ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBN���A�V�K�o�^�̏ꍇ�́A9(�폜)���Z�b�g
                 strTextLine = strTextLine + " '9' "
        Else
                  strTextLine = strTextLine + arrColumn(i).strPhysicsName
        
        End If
         strText = strText + strTextLine + """  " + ("+ "" \n"" ") + vbCrLf
    Next i
    strText = strText + "            sql += ""FROM " + tTbl.strPhysicsName + " \n """ + vbCrLf
    'WHERE�啔��
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
    
        If arrColumn(i).strPrimaryKey <> "" Then
        '��L�[�̍��ڂ��w��B
            strTextLine = ""
            strTextLine = strTextLine + "            sql += "" AND " + arrColumn(i).strPhysicsName + "="""
                  
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'UP_DT�̏ꍇ�́A[sysDateTime]���Z�b�g
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBN���A�V�K�o�^�̏ꍇ�́A9(�폜)���Z�b�g
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
    
    '�폜����-----------------------------------------------------------------------------------
    strText = strText + "" + vbCrLf
    strText = strText + "            sql += ""DELETE FROM " + tTbl.strPhysicsName + " \n """ + vbCrLf
    'WHERE�啔��
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
        If arrColumn(i).strPrimaryKey <> "" Then
        '��L�[�̍��ڂ��w��B
            strTextLine = ""
            strTextLine = strTextLine + "            sql += "" AND " + arrColumn(i).strPhysicsName + "="""
                  
            If (arrColumn(i).strPhysicsName = "UP_DT") Then
                'UP_DT�̏ꍇ�́A[sysDateTime]���Z�b�g
                strTextLine = strTextLine + " + ""'""  + sysDateTime  + ""'"" "
            ElseIf arrColumn(i).strPhysicsName = "SHORI_KBN" Then
                'SHORI_KBN���A�V�K�o�^�̏ꍇ�́A9(�폜)���Z�b�g
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
    
    '�㏈��
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

'�e�[�u����������DB����(Dao)�N���X(python)��Select���̏����̕�������쐬
Private Function outputPythonDao_Select(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn) As String
On Error GoTo Err0

    '�A�N�e�B�u�ȃV�[�g���擾����B
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Dao"  '�N���X��
    
    '�Ώی���(�J������)���擾����B
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    Dim strText As String
    strText = ""

    ' �Q�Ə���-----------------------------------------------------------------------------------------
    strText = strText + "    # " + tTbl.strLogicalName + "�ւ̎Q�Ə��� " + vbCrLf
    strText = strText + "    def select" + tTbl.strPhysicsName + "(self"
    ' �֐��̈���������ݒ�
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
    
    ' SELECT��-----------------------------------------------------------------------------------------
    ' SELECT��
    strText = strText + "            sql += ""SELECT"" + ""\n""" + "  " + vbCrLf
    '���ڎw�蕔��
    For i = 0 To lngLineCount - 1
        strTextLine = ""
        '��L�[�̍��ڂ��w��B
        If i = 0 Then
            strTextLine = strTextLine + "            sql += ""     "
        Else
            strTextLine = strTextLine + "            sql += ""    ,"
        End If
        ' ���ږ��̕��������L�q
        strTextLine = strTextLine + arrColumn(i).strPhysicsName
        strTextLine = strTextLine + """ + "" \n"" "
        ' ���ږ��̘_�������L�q
        If Len(strTextLine) < INDENT_SPACE Then
            strText = strText + strTextLine + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
        Else
            strText = strText + strTextLine + " #" + arrColumn(i).strLogicalName + vbCrLf
        End If
        ' strText = strText + strTextLine + """ + "" \n"" " + vbCrLf
        ' strText = strText + strTextLine + " + ""\n"" " + String(INDENT_SPACE - Len(strTextLine), " ") + " #" + arrColumn(i).strLogicalName + vbCrLf
    Next i
    ' FROM��
    strText = strText + "            sql += ""FROM " + tTbl.strPhysicsName + " \n """ + vbCrLf
    ' WHERE�啔��
    strText = strText + "            sql += ""WHERE 1 = 1 \n """ + vbCrLf
    count = 0
    For i = 0 To lngLineCount - 1
    
        If arrColumn(i).strPrimaryKey <> "" Then
        '��L�[�̍��ڂ��w��B
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
    
    'SQL�����s
    strText = strText + "" + vbCrLf
    strText = strText + "            # SQL�����s" + vbCrLf
    strText = strText + "            df = self.sqlCon.executeSql(sql)" + vbCrLf
    strText = strText + "" + vbCrLf
    strText = strText + "            # SQL���s���ʂ��擾" + vbCrLf
    strText = strText + "            entityResult: " + tTbl.strPhysicsName + "_Entity" + " = " + tTbl.strPhysicsName + "_Entity()" + vbCrLf
    strText = strText + "            i = 0 " + vbCrLf
    strText = strText + "            if len(df) > 0: " + vbCrLf
    'SQL�����s���ʂ��擾
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

