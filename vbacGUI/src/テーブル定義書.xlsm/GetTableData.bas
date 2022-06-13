Attribute VB_Name = "GetTableData"


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
    strOverview     As String   '�e�[�u�����e
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

'�e�[�u���E�J���������擾
Public Sub getTableData(ByRef tTbl As typeTable, ByRef arrColumn() As typeColumn)
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
    tTbl.strOverview = Trim(wsActSheet.Range("D4").Value)     '�e�[�u�����e

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
