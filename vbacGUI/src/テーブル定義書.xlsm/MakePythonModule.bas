Attribute VB_Name = "MakePythonModule"


Const SCHMA_NAME As String = "" '�X�L�[�}��
Const COL_START_ROW As Long = 7 '�J��������7�s�ڂ���X�^�[�g
Const TBL_START_ROW As Long = 5 '�e�[�u������5�s�ڂ���X�^�[�g
Const SHEET_TABLE_LIST As String = "�e�[�u���ꗗ�\" '�e�[�u���ꗗ�\�̃V�[�g��

Public strSchemaName As String  '�X�L�[�}��

'�e�[�u���ꗗ�\�V�[�g�̃e�[�u���X�N���v�g���o�́B
Public Sub makePythonFile_All()
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
        Call makePythonFile
    
    
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
Public Sub makePythonFile_Sheet()
On Error GoTo Err0
   '�x�����o�Ȃ��悤�ɐݒ�
    Application.DisplayAlerts = False

    '�A�N�e�B�u�V�[�g�̃e�[�u���X�N���v�g���o�́B
    Call makePythonFile
    
    '�x�����o��悤�ɐݒ��߂�
    Application.DisplayAlerts = True
   
    MsgBox (strBuf + "�o�́@����")
 
    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub

'�A�N�e�B�u�V�[�g�̃e�[�u���X�N���v�g���o�́B
Private Sub makePythonFile()
On Error GoTo Err0

    Dim tTbl As getTableData.typeTable           '�e�[�u�����
    Dim arrColumn() As getTableData.typeColumn   '�J�������
    Dim strBuf As String
    strBuf = ""
    
    '-------------------------------------------------------------------------------
'    '�A�N�e�B�u�̃V�[�g����e�[�u���ƃJ���������擾
    Call getTableData.getTableData(tTbl, arrColumn)
'
'    '�e�[�u������v�f�N���X��Python�t�@�C�����쐬
    Call MakePythonEntityModule.outputPythonEntity(tTbl, arrColumn)
    Call MakePythonDaoModule.outputPythonDao(tTbl, arrColumn)
    '-------------------------------------------------------------------------------

    Exit Sub
Err0:
    MsgBox Error
    Application.ScreenUpdating = True
    
End Sub
