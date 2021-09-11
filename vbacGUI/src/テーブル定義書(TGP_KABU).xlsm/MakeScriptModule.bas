Attribute VB_Name = "MakeScriptModule"


Const SCHMA_NAME As String = "" '�X�L�[�}��
Const COL_START_ROW As Long = 7 '�J��������7�s�ڂ���X�^�[�g
Const TBL_START_ROW As Long = 5 '�e�[�u������5�s�ڂ���X�^�[�g
Const SHEET_TABLE_LIST As String = "�e�[�u���ꗗ�\" '�e�[�u���ꗗ�\�̃V�[�g��

'���ڏ��̍\����
Type typeTable
    lngNo           As Long     'No
    strLogicalName  As String   '�_����
    strPhysicsName  As String   '������
    strSchema       As String   '�X�L�[�}��
    strHistoryFlag  As String   '�����쐬�t���O(�v/��)
    strKind         As String   '�e�[�u�����
End Type

'���ڏ��̍\����
Type typeColumn
    lngNo           As Long     'No
    strLogicalName  As String   '�_����
    strPhysicsName  As String   '������
    strDataType     As String   '�f�[�^�^
    lngLength       As Long     '�f�[�^����
    lngDecimal      As Long     '��������
    strRequiredFlag As String   '�K�{�敪
    strPrimaryKey   As String   '��L�[
    strDefalutData  As String   '�f�t�H���g�l
    strRemarks      As String   '���l
End Type

Public strSchemaName As String  '�X�L�[�}��

'�e�e�[�u����`����e�[�u���ꗗ���쐬
Public Sub makeTableList()

    On Error GoTo Err0
    
    '�x�����o�Ȃ��悤�ɐݒ�
    Application.DisplayAlerts = False
    
    Dim lngLineCnt As Long
    lngLineCnt = TBL_START_ROW
    
    '�A�N�e�B�u�V�[�g�̐؂�ւ�
    ActiveWorkbook.Worksheets(SHEET_TABLE_LIST).Activate
    
    Dim tTblBuf As typeTable    '�e�[�u����`���
    For Each Ws In Worksheets
        If Ws.Name <> "����" And Ws.Name <> SHEET_TABLE_LIST And Ws.Name <> "Sheet1" Then
                   
            '�e�[�u�����̏�����
            tTblBuf.lngNo = 0     'No
            tTblBuf.strLogicalName = ""   '�_����
            tTblBuf.strPhysicsName = ""   '������
            tTblBuf.strSchema = ""   '�X�L�[�}��
            tTblBuf.strHistoryFlag = ""   '�����쐬�t���O(�v/��)
            tTblBuf.strKind = ""   '���l
            
            '�e�[�u�����̎擾
            tTblBuf.lngNo = lngLineCnt - TBL_START_ROW + 1 'No
            tTblBuf.strLogicalName = Ws.Range("A4").Value  '�_����
            tTblBuf.strPhysicsName = Ws.Range("C4").Value  '������
            tTblBuf.strKind = Ws.Range("N1").Value  '�e�[�u�����
            
            '�e�[�u���ꗗ�Ƀe�[�u�������Z�b�g
            Worksheets(SHEET_TABLE_LIST).Range("A" + Format(lngLineCnt)).Value = tTblBuf.lngNo  'No
            Worksheets(SHEET_TABLE_LIST).Range("C" + Format(lngLineCnt)).Value = tTblBuf.strLogicalName '�_����
            Worksheets(SHEET_TABLE_LIST).Range("K" + Format(lngLineCnt)).Value = tTblBuf.strPhysicsName '������
            Worksheets(SHEET_TABLE_LIST).Range("Y" + Format(lngLineCnt)).Value = tTblBuf.strKind    '�e�[�u�����
               
            lngLineCnt = lngLineCnt + 1
            
        End If
    Next Ws

    Exit Sub
    
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    

End Sub


'�e�[�u���ꗗ�\�V�[�g�̃e�[�u���X�N���v�g���o�́B
Public Sub makeScript_CreateTable_All()
On Error GoTo Err0
    '�x�����o�Ȃ��悤�ɐݒ�
    Application.DisplayAlerts = False
    
    Dim strReportText As String '�o�͌��ʂ��e�L�X�g�ŏo��
    
    '�A�N�e�B�u�ȃV�[�g���擾����B
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    
    '-----------------------------------------------------------------
    '������
    
    '�A�N�e�B�u�V�[�g�̐؂�ւ�
    ActiveWorkbook.Worksheets(SHEET_TABLE_LIST).Activate
        
    '�Ώی���(�J������)���擾����B
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    wkRows = wsActSheet.Cells.Rows.count
    lngMaxLine = wsActSheet.Cells(wkRows, 1).End(xlUp).Row
    lngLineCount = lngMaxLine - TBL_START_ROW + 1
    
    
    '-----------------------------------------------------------------
    '�e�[�u���ꗗ�����擾
    
    '�e�[�u����񃊃X�g��錾
    Dim arrTable() As typeTable
    ReDim arrTable(lngLineCount + 1)
    Dim tTblBuf As typeTable
    Dim strBuf As String
    
    '�J����������s���擾����
    Dim i As Long
    Dim cnt As Long
    cnt = 0
    For i = TBL_START_ROW To lngMaxLine
    
        '[No]��Ɏ������������ꍇ�A�o�͑ΏۊO�Ƃ���B
        If (isStrikethrough(wsActSheet.Range("A" + Format(i))) = True) Then
            '[No]��Ɏ������������ꍇ�A
            '�v�f��������������B
            lngLineCount = lngLineCount - 1
        
        Else
            '[No]��Ɏ���������Ȃ��ꍇ�A
            '�e�[�u�������擾����B
            
            'No
            strBuf = wsActSheet.Range("A" + Format(i)).Value
            tTblBuf.lngNo = Val(strBuf)
            '�_����
            tTblBuf.strLogicalName _
                = removeStrikethrough(wsActSheet.Range("C" + Format(i)))
            '������
            tTblBuf.strPhysicsName _
                = removeStrikethrough(wsActSheet.Range("K" + Format(i)))
                
            '�z��ɒl���Z�b�g����B
            arrTable(cnt).lngNo = tTblBuf.lngNo
            arrTable(cnt).strLogicalName = tTblBuf.strLogicalName
            arrTable(cnt).strPhysicsName = tTblBuf.strPhysicsName
            cnt = cnt + 1
        End If
    Next i
    
    '-----------------------------------------------------------------
    '�e�[�u���ꗗ���X�N���v�g���쐬
    Dim strSheetName As String
    For i = 0 To cnt - 1
        strSheetName = arrTable(i).strLogicalName
        
        '�A�N�e�B�u�V�[�g�̐؂�ւ�
        ActiveWorkbook.Worksheets(strSheetName).Activate
        
        '�A�N�e�B�u�V�[�g�̃e�[�u���X�N���v�g���o�́B
        Call makeScript_CreateTable
    
    
    Next i
    
    
    '���̃V�[�g�ɖ߂�
    ActiveWorkbook.Worksheets(strActSheetName).Activate
        
    MsgBox ("�o�́@����")
    
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True

End Sub


'�A�N�e�B�u�V�[�g�̃e�[�u���X�N���v�g���o�́B
Public Sub makeScript_CreateTable_Sheet()
On Error GoTo Err0
   '�x�����o�Ȃ��悤�ɐݒ�
    Application.DisplayAlerts = False

    '�A�N�e�B�u�V�[�g�̃e�[�u���X�N���v�g���o�́B
    Call makeScript_CreateTable
    
    '�x�����o��悤�ɐݒ��߂�
    Application.DisplayAlerts = True
   
    MsgBox (strBuf + "�o�́@����")
 
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub

'�A�N�e�B�u�V�[�g�̃e�[�u���X�N���v�g���o�́B
Private Sub makeScript_CreateTable()
On Error GoTo Err0

    Dim tTbl As typeTable           '�e�[�u�����
    Dim arrColumn() As typeColumn   '�J�������
    Dim strBuf As String
    strBuf = ""
    
    '-------------------------------------------------------------------------------
    '�A�N�e�B�u�̃V�[�g����e�[�u���ƃJ���������擾
    Call getTableData(tTbl, arrColumn)
    
    '�e�[�u�������X�N���v�g�ɏo��
    Call outputScriput_CreateTable(tTbl, arrColumn)
    strBuf = strBuf + "CreateTable_" + tTbl.strPhysicsName + "(" + tTbl.strLogicalName + ").sql" + vbCrLf

    '��������K�{�敪��='�v'�̏ꍇ
    If (tTbl.strHistoryFlag = "�v") Then
        '�e�[�u�����(����)���X�N���v�g�ɏo��
        Call outputScriput_CreateTable_R(tTbl, arrColumn)
        strBuf = strBuf + "CreateTable_" + tTbl.strPhysicsName + "_R(����_" + tTbl.strLogicalName + ").sql" + vbCrLf
    End If
    '-------------------------------------------------------------------------------

    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub


'�e�[�u���E�J���������擾
Private Sub getTableData(ByRef tTbl As typeTable, ByRef arrColumn() As typeColumn)
On Error GoTo Err0
    '�A�N�e�B�u�ȃV�[�g���擾����B
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name

    '--------------------------------------------------------------------------------------
    '�A�N�e�B�u�ȃV�[�g����e�[�u�������擾����B
    
    ' Dim tTbl As typeTable
    tTbl.strLogicalName = Trim(wsActSheet.Range("A4").Value)  '�e�[�u����
    tTbl.strPhysicsName = Trim(wsActSheet.Range("C4").Value)  '�e�[�u����(�p��)
    tTbl.strSchema = SCHMA_NAME                               '�X�L�[�}��
    tTbl.strHistoryFlag = Trim(wsActSheet.Range("I2").Value)  '�����쐬�t���O(�v/��)

    '--------------------------------------------------------------------------------------
    '�A�N�e�B�u�ȃV�[�g����J���������擾����
    
    '�Ώی���(�J������)���擾����B
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    wkRows = wsActSheet.Cells.Rows.count
    lngMaxLine = wsActSheet.Cells(wkRows, 1).End(xlUp).Row
    lngLineCount = lngMaxLine - COL_START_ROW + 1
    
    ' Dim arrColumn() As typeColumn
    ReDim arrColumn(lngLineCount + 1)
    Dim tColBuf As typeColumn
    Dim strBuf As String

    '�J����������s���擾����
    Dim i As Long
    Dim cnt As Long
    cnt = 0
    For i = COL_START_ROW To lngMaxLine + 1
    
        '[No]��Ɏ������������ꍇ�A�o�͑ΏۊO�Ƃ���B
        If (isStrikethrough(wsActSheet.Range("A" + Format(i))) = True) Then
            '[No]��Ɏ������������ꍇ�A
            '�v�f��������������B
            lngLineCount = lngLineCount - 1
        
        Else
            '[No]��Ɏ���������Ȃ��ꍇ�A
            '�V�[�g����l���擾����B
'            'No
'            tColBuf.lngNo = wsActSheet.Range("A" + Format(i)).Value
'            '�_����
'            tColBuf.strLogicalName = wsActSheet.Range("B" + Format(i)).Value
'            '������
'            tColBuf.strPhysicsName = wsActSheet.Range("C" + Format(i)).Value
'            '�f�[�^�^
'            tColBuf.strDataType = wsActSheet.Range("D" + Format(i)).Value
'            '�f�[�^����
'            tColBuf.lngLength = wsActSheet.Range("E" + Format(i)).Value
'            '��������
'            tColBuf.lngDecimal = wsActSheet.Range("F" + Format(i)).Value
'            '�K�{�敪
'            tColBuf.strRequiredFlag = wsActSheet.Range("G" + Format(i)).Value
'            '��L�[
'            tColBuf.strPrimaryKey = wsActSheet.Range("H" + Format(i)).Value
'            '�f�t�H���g�l
'            tColBuf.strDefalutData = wsActSheet.Range("I" + Format(i)).Value
'            '���l
'            tColBuf.strRemarks = wsActSheet.Range("K" + Format(i)).Value

            'No
            tColBuf.lngNo = wsActSheet.Range("A" + Format(i)).Value
            '�_����
            tColBuf.strLogicalName _
                = removeStrikethrough(wsActSheet.Range("B" + Format(i)))
            '������
            tColBuf.strPhysicsName _
                = removeStrikethrough(wsActSheet.Range("C" + Format(i)))
            '�f�[�^�^
            tColBuf.strDataType _
                = removeStrikethrough(wsActSheet.Range("D" + Format(i)))
            '�f�[�^����
            strBuf = removeStrikethrough(wsActSheet.Range("E" + Format(i)))
            If (IsNumeric(strBuf) = True) Then
                tColBuf.lngLength = Val(strBuf)
            Else
                tColBuf.lngLength = 0
            End If
            
            '��������
            strBuf = removeStrikethrough(wsActSheet.Range("F" + Format(i)))
            If (IsNumeric(strBuf) = True) Then
                tColBuf.lngDecimal = Val(strBuf)
            Else
                tColBuf.lngDecimal = 0
            End If
            '�K�{�敪
            tColBuf.strRequiredFlag = wsActSheet.Range("G" + Format(i)).Value
            '��L�[
            tColBuf.strPrimaryKey = wsActSheet.Range("H" + Format(i)).Value
            '�f�t�H���g�l
            tColBuf.strDefalutData = wsActSheet.Range("I" + Format(i)).Value
            '���l
            tColBuf.strRemarks = wsActSheet.Range("K" + Format(i)).Value
    
            '�z��ɒl���Z�b�g����B
            arrColumn(cnt).lngNo = tColBuf.lngNo                   'No
            arrColumn(cnt).strLogicalName = Trim(tColBuf.strLogicalName) '�_����
            arrColumn(cnt).strPhysicsName = Trim(tColBuf.strPhysicsName) '������
            arrColumn(cnt).strDataType = Trim(tColBuf.strDataType)       '�f�[�^�^
            arrColumn(cnt).lngLength = tColBuf.lngLength           '�f�[�^����
            arrColumn(cnt).lngDecimal = tColBuf.lngDecimal         '��������
            arrColumn(cnt).strRequiredFlag = Trim(tColBuf.strRequiredFlag) '�K�{�敪
            arrColumn(cnt).strPrimaryKey = Trim(tColBuf.strPrimaryKey)   '��L�[
            arrColumn(cnt).strDefalutData = Trim(tColBuf.strDefalutData) '�f�t�H���g�l
            arrColumn(cnt).strRemarks = Trim(tColBuf.strRemarks)         '���l
            cnt = cnt + 1
        
        End If

    Next i

    '�z��T�C�Y�̍Đݒ�i��������ŃX�L�b�v�������炷�j
    ReDim Preserve arrColumn(cnt)

    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
End Sub

'�e�[�u�������X�N���v�g�ɏo��
Private Sub outputScriput_CreateTable(ByRef tTbl As typeTable, ByRef arrColumn() As typeColumn)
On Error GoTo Err0
    '�A�N�e�B�u�ȃV�[�g���擾����B
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name
    
    '�Ώی���(�J������)���擾����B
    Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    '--------------------------------------------------------------------------------------
    '�X�N���v�g��g�ݗ��Ă�B
   
    Dim strSql As String
    strSql = ""
    
    'DROP TABLE�����d����
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
   '�e�[�u������錾
    If (tTbl.strSchema = "") Then
        strSql = strSql + "CREATE TABLE [dbo].[" + tTbl.strPhysicsName + "]" + vbCrLf
    Else
        strSql = strSql + "CREATE TABLE [dbo].[" + tTbl.strSchema + "." + tTbl.strPhysicsName + "]" + vbCrLf
    End If
    strSql = strSql + "(" + vbCrLf

    '�J�����������ƌ�����錾
    Dim strSqlLine As String
    
    For i = 0 To lngLineCount - 1
        '������ #0-3[4]
        strSqlLine = "    "     '
        strBuf = ""

        '������ #4-34[31]
        strBuf = "[" + arrColumn(i).strPhysicsName + "]"
        strSqlLine = strSqlLine + strBuf + String(31 - Len(strBuf), " ")

        '�f�[�^�^(����,��������) #35-XX[-]
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
        
        '�f�t�H���g�l
        If (arrColumn(i).strDefalutData <> "") Then
            strBuf = " DEFAULT " + arrColumn(i).strDefalutData
            strSqlLine = strSqlLine + strBuf
        End If

        '�K�{�敪
        If (arrColumn(i).strRequiredFlag <> "") Then
            strBuf = " NOT NULL"
            strSqlLine = strSqlLine + strBuf
        End If
        
        
        ' '�J���}
        ' If (i <> (lngLineCount - 1)) Then
        '     strSqlLine = strSqlLine + ","
        '     strSql = strSql + strSqlLine + vbCrLf
        ' Else
        '     strSql = strSql + strSqlLine + vbCrLf
        ' End If

        '�J���}
        strSqlLine = strSqlLine + ","
        strSql = strSql + strSqlLine + vbCrLf

    Next i

    '��L�[
    strBuf = "    CONSTRAINT [PK_" + tTbl.strPhysicsName + "] "
    strSqlLine = strBuf

    strBuf = "PRIMARY KEY CLUSTERED " + vbCrLf + "(" + vbCrLf
    For i = 0 To lngLineCount - 1
        If (arrColumn(i).strPrimaryKey <> "") Then
            strBuf = strBuf + "       " + arrColumn(i).strPhysicsName + " ASC, " + vbCrLf
        End If
    Next i
    strBuf = Left(strBuf, (Len(strBuf) - 4)) + vbCrLf + "" '�]�v�ȃJ���}���폜
    strBuf = strBuf + "    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]" + vbCrLf
    strBuf = strBuf + ") ON [PRIMARY] " + vbCrLf
    strBuf = strBuf + "GO " + vbCrLf
    
    strSqlLine = strSqlLine + strBuf + vbCrLf
'    strSqlLine = strSqlLine + "        ENABLE" + vbCrLf
'    strSqlLine = strSqlLine + ")"
    strSql = strSql + strSqlLine + vbCrLf
    
'    '�e�[�u���R�����g
'    strSql = strSql + "/" + vbCrLf
'    If (tTbl.strSchema = "") Then
'        strSql = strSql + "COMMENT ON TABLE " + tTbl.strPhysicsName + " IS '" + tTbl.strLogicalName + "'" + vbCrLf
'    Else
'        strSql = strSql + "COMMENT ON TABLE " + tTbl.strSchema + "." + tTbl.strPhysicsName + " IS '" + tTbl.strLogicalName + "'" + vbCrLf
'    End If
    
    '�J�����R�����g
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
    '�X�N���v�g�̏o��
    Dim datFile As String
    datFile = ActiveWorkbook.Path + "\CreateTable_" + tTbl.strPhysicsName + "(" + tTbl.strLogicalName + ").sql"
    Open datFile For Output As #1

    Print #1, strSql

    Close #1
    
'    MsgBox (datFile + "�ɏ����o���܂���")
 
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub


'�e�[�u�����(����)���X�N���v�g�ɏo��
Private Sub outputScriput_CreateTable_R(ByRef tTbl As typeTable, ByRef arrColumn() As typeColumn)
On Error GoTo Err0
    '�A�N�e�B�u�ȃV�[�g���擾����B
    Dim wsActSheet As Worksheet
    Set wsActSheet = ActiveSheet
    Dim strActSheetName As String
    strActSheetName = ActiveSheet.Name

    strSchemaName = SCHMA_NAME  '�Œ�̃X�L�[�}�����Z�b�g(�K�v�ɉ����ėv�ύX)
    
    '�Ώی���(�J������)���擾����B
   Dim lngMaxLine As Long
    Dim lngLineCount As Long
    lngLineCount = UBound(arrColumn) - 1
    lngMaxLine = lngLineCount + COL_START_ROW - 1
    
    '�����e�[�u���p�̏����쐬
    Dim tTbl_R As typeTable
    tTbl_R.lngNo = tTbl.lngNo
    tTbl_R.strLogicalName = "����_" + tTbl.strLogicalName
    tTbl_R.strPhysicsName = tTbl.strPhysicsName + "_R"
    tTbl_R.strSchema = tTbl.strSchema
    tTbl_R.strHistoryFlag = tTbl.strHistoryFlag
    
    '--------------------------------------------------------------------------------------
    '�X�N���v�g��g�ݗ��Ă�B
   
    Dim strSql As String
    strSql = ""
    
    'DROP TABLE�����d����
    strSql = strSql + "BEGIN " + vbCrLf
    'T_JUCYU_D';
    If (tTbl.strSchema = "") Then
        strSql = strSql + "   EXECUTE IMMEDIATE 'DROP TABLE " + tTbl.strPhysicsName + "';" + vbCrLf
    Else
        strSql = strSql + "   EXECUTE IMMEDIATE 'DROP TABLE " + tTbl.strSchema + "." + tTbl.strPhysicsName + "';" + vbCrLf
    End If
    strSql = strSql + "EXCEPTION " + vbCrLf
    strSql = strSql + "   WHEN OTHERS THEN" + vbCrLf
    strSql = strSql + "      IF SQLCODE != -942 THEN " + vbCrLf
    strSql = strSql + "         RAISE;" + vbCrLf
    strSql = strSql + "      END IF;" + vbCrLf
    strSql = strSql + "END;" + vbCrLf
    strSql = strSql + "/" + vbCrLf
    '�e�[�u������錾
    If (tTbl_R.strSchema = "") Then
        strSql = strSql + "CREATE TABLE " + tTbl_R.strPhysicsName + vbCrLf
    Else
        strSql = strSql + "CREATE TABLE " + tTbl_R.strSchema + "." + tTbl_R.strPhysicsName + vbCrLf
    End If
    strSql = strSql + "(" + vbCrLf

    '�J�����������ƌ�����錾
    Dim strSqlLine As String
    
    For i = 0 To lngLineCount - 1
        '������ #0-3[4]
        strSqlLine = "    "     '
        strBuf = ""

        '������ #4-34[31]
        strBuf = arrColumn(i).strPhysicsName
        strSqlLine = strSqlLine + strBuf + String(31 - Len(strBuf), " ")

        '�f�[�^�^(����,��������) #35-XX[-]
        If (arrColumn(i).strDataType = "DATE") Then
            strBuf = arrColumn(i).strDataType
        ElseIf (arrColumn(i).strDataType = "NUMBER") Then
            strBuf = arrColumn(i).strDataType + "(" + CStr(arrColumn(i).lngLength) + "," + CStr(arrColumn(i).lngDecimal) + ")"
        ElseIf (arrColumn(i).strDataType = "VARCHAR2") Then
            strBuf = arrColumn(i).strDataType + "(" + CStr(arrColumn(i).lngLength) + ")"
        Else
            On Error GoTo Err0
        End If
        strSqlLine = strSqlLine + strBuf
        
'        '�f�t�H���g�l
'        If (arrColumn(i).strDefalutData <> "") Then
'            strBuf = " DEFAULT " + arrColumn(i).strDefalutData
'            strSqlLine = strSqlLine + strBuf
'        End If

        '�K�{�敪
        If (arrColumn(i).strRequiredFlag <> "") Then
            strBuf = " NOT NULL"
            strSqlLine = strSqlLine + strBuf
        End If
        
        '�J���}
        If (i <> (lngLineCount - 1)) Then
            strSqlLine = strSqlLine + ","
            strSql = strSql + strSqlLine + vbCrLf
        Else
            strSql = strSql + strSqlLine + vbCrLf
        End If

        '�񎟍s�̑}��
        ' �������e�[�u����DEL_KBN�̉��ɑ}������\��B
        ' ��FSv2.3�ȑO�́A�񎟍s�̈ʒu�̓o���o�����AFSv2.4�œ��ꂳ���c�c�����H
        If (arrColumn(i).strPhysicsName = "DEL_KBN") Then
            strSql = strSql + "    KAIJI                          NUMBER(18,0)," + vbCrLf
        End If
    Next i

    ' '��L�[
    ' strBuf = "    CONSTRAINT IDX_" + Mid(tTbl_R.strPhysicsName, 3, (Len(tTbl_R.strPhysicsName))) + "_PK "
    ' strSqlLine = strBuf

    ' strBuf = "PRIMARY KEY ("
    ' For i = 0 To lngLineCount - 1
    '     If (arrColumn(i).strPrimaryKey <> "") Then
    '         strBuf = strBuf + arrColumn(i).strPhysicsName + ", "
    '     End If
    ' Next i
    ' strBuf = Left(strBuf, (Len(strBuf) - 2)) + ") USING INDEX" '�]�v�ȃJ���}���폜
    ' strSqlLine = strSqlLine + strBuf + vbCrLf
    ' strSqlLine = strSqlLine + "        ENABLE" + vbCrLf
    ' strSqlLine = strSqlLine + ")"
    ' strSql = strSql + strSqlLine + vbCrLf
    
    
    strSql = strSql + ")" + vbCrLf
    
    '�e�[�u���R�����g
    strSql = strSql + "/" + vbCrLf
    If (tTbl.strSchema = "") Then
        strSql = strSql + "COMMENT ON TABLE " + tTbl_R.strPhysicsName + " IS '" + tTbl_R.strLogicalName + "'" + vbCrLf
    Else
        strSql = strSql + "COMMENT ON TABLE " + tTbl_R.strSchema + "." + tTbl_R.strPhysicsName + " IS '" + tTbl_R.strLogicalName + "'" + vbCrLf
    End If
    
    '�J�����R�����g
    strSql = strSql + "/" + vbCrLf
    For i = 0 To lngLineCount - 1
        If (tTbl.strSchema = "") Then
            strBuf = "COMMENT ON COLUMN " + tTbl_R.strPhysicsName + "." + arrColumn(i).strPhysicsName _
                + " IS '" + arrColumn(i).strLogicalName + "'" + vbCrLf
        Else
            strBuf = "COMMENT ON COLUMN " + tTbl_R.strSchema + "." + tTbl_R.strPhysicsName + "." + arrColumn(i).strPhysicsName _
                + " IS '" + arrColumn(i).strLogicalName + "'" + vbCrLf
        End If
        strBuf = strBuf + "/" + vbCrLf
        strSql = strSql + strBuf
    
        '�񎟍s�̑}��
        ' �������e�[�u����DEL_KBN�̉��ɑ}������\��B
        ' ��FSv2.3�ȑO�́A�񎟍s�̈ʒu�̓o���o�����AFSv2.4�œ��ꂳ���c�c�����H
        If (arrColumn(i).strPhysicsName = "DEL_KBN") Then
            If (tTbl.strSchema = "") Then
                strSql = strSql + "COMMENT ON COLUMN " + tTbl_R.strPhysicsName + ".KAIJI IS '��'" + vbCrLf
            Else
                strSql = strSql + "COMMENT ON COLUMN " + tTbl_R.strSchema + "." + tTbl_R.strPhysicsName + ".KAIJI IS '��'" + vbCrLf
            End If
            strSql = strSql + "/" + vbCrLf
        End If

    Next i
    
   '--------------------------------------------------------------------------------------
    '�X�N���v�g�̏o��
    Dim datFile As String
    datFile = ActiveWorkbook.Path + "\CreateTable_" + tTbl_R.strPhysicsName + "(" + tTbl_R.strLogicalName + ").sql"
    Open datFile For Output As #1

    Print #1, strSql

    Close #1
    
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub


