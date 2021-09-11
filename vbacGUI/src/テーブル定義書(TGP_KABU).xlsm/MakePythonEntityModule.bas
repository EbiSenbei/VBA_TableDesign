Attribute VB_Name = "MakePythonEntityModule"


Const SCHMA_NAME As String = "" '�X�L�[�}��
Const COL_START_ROW As Long = 7 '�J��������7�s�ڂ���X�^�[�g
Const TBL_START_ROW As Long = 5 '�e�[�u������5�s�ڂ���X�^�[�g
Const SHEET_TABLE_LIST As String = "�e�[�u���ꗗ�\" '�e�[�u���ꗗ�\�̃V�[�g��


Public strSchemaName As String  '�X�L�[�}��

'�e�[�u���������ɗv�f(Entity)�N���X��python�t�@�C�����o��
Public Sub outputPythonEntity(ByRef tTbl As getTableData.typeTable, ByRef arrColumn() As getTableData.typeColumn)
On Error GoTo Err0
    '�A�N�e�B�u�ȃV�[�g���擾����B
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    Dim strClassName As String
    strClassName = tTbl.strPhysicsName + "_Entity"  '�N���X��
    
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
    strText = strText + "import copy" + vbCrLf
    strText = strText + "import datetime" + vbCrLf
    strText = strText + "# ------------------------------------------------------------------" + vbCrLf
    strText = strText + "# �萔" + vbCrLf
    strText = strText + "DB_DRIBER: str = ""{ODBC Driver 13 for SQL Server}""" + vbCrLf
    strText = strText + "" + vbCrLf
    
    strText = strText + "# -------------------------------------------------------------------" + vbCrLf
    strText = strText + "# �N���X�i�v�f���j" + vbCrLf
    strText = strText + "# �Q�Ƃ���" + tTbl.strLogicalName + "���(" + tTbl.strPhysicsName + ")" + vbCrLf
    strText = strText + "class " + strClassName + ":" + vbCrLf

    ' ----------------------------------------------------------------------------------------
    strText = strText + "    # �N���X�ϐ�" + vbCrLf
    For i = 0 To lngLineCount - 1
        '������ #0-3[4]
        strTextLine = "    "     '
        strBuf = ""

        '�ϐ��� #4-34[31]
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

        '�ϐ��� #8-34[31]
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
    '�X�N���v�g�̏o��
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
'    MsgBox (datFile + "�ɏ����o���܂���")
 
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub

