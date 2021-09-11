Attribute VB_Name = "CommonModule"

'�Z�����̕����Ɏ�����������݂��邩��������
'�@�߂�l�Ftrue=�����������Afalse=��������Ȃ�
Public Function isStrikethrough(ByRef cell As Object) As Boolean

    Dim strTemp As String
    Dim i As Long
    
    Dim resulfFlag As Boolean
    resulfFlag = False
    
    Select Case cell.Font.Strikethrough
            '�Z���S�̂Ɏ�����̏ꍇ
            Case True
                'Debug.Print "Case True"
                resulfFlag = True
                
            '�Ȃɂ��Ȃ��̏ꍇ
            Case False
                'Debug.Print "Case False"
                resulfFlag = False
                
            '�Z�����̈ꕔ���Ɏ�����̏ꍇ
            Case Else
                'Debug.Print "Case Else"
                '������̂��Ă��Ȃ��������������Ă����ϐ��̏�����
                strTemp = ""
                
               '�Z�����̕�����𖖔�����擪�Ɍ������Ē��ׂĂ���
                For i = Len(cell.Value) To 1 Step -1
                    '�Ђƕ��������肵�A����������Ă����temp�Ɍ���
                    If cell.Characters(Start:=i, Length:=1).Font.Strikethrough = False Then
                        '�������璲�ׂĂ���̂Ő擪�Ɍ������Ă���
                        strTemp = Mid(cell.Value, i, 1) & strTemp
                    Else
                        resulfFlag = True
                    End If
                Next i
                
        End Select
        
        isStrikethrough = resulfFlag

End Function

'�Z�����̕����Ɏ�����������݂��邩��������
'�@�߂�l�Ftrue=�����������Afalse=��������Ȃ�
Public Function removeStrikethrough(ByRef cell As Object) As String

    Dim strTemp As String
    Dim i As Long
    
    Dim strResult As String
    strResult = ""
    
    Select Case cell.Font.Strikethrough
            '�Z���S�̂Ɏ�����̏ꍇ
            Case True
                'Debug.Print "Case True"
                strResult = ""
                
            '�Ȃɂ��Ȃ��̏ꍇ
            Case False
                'Debug.Print "Case False"
                strResult = cell.Value
                
            '�Z�����̈ꕔ���Ɏ�����̏ꍇ
            Case Else
                'Debug.Print "Case Else"
                '������̂��Ă��Ȃ��������������Ă����ϐ��̏�����
                strTemp = ""
                
               '�Z�����̕�����𖖔�����擪�Ɍ������Ē��ׂĂ���
                For i = Len(cell.Value) To 1 Step -1
                    '�Ђƕ��������肵�A����������Ă����temp�Ɍ���
                    If cell.Characters(Start:=i, Length:=1).Font.Strikethrough = False Then
                        '�������璲�ׂĂ���̂Ő擪�Ɍ������Ă���
                        strTemp = Mid(cell.Value, i, 1) & strTemp
                    Else
                        resulfFlag = True
                    End If
                Next i
                strResult = strTemp
                
        End Select
        
        removeStrikethrough = strResult

End Function





