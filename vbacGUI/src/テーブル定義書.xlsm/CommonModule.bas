Attribute VB_Name = "CommonModule"

'セル内の文字に取消し線が存在するかﾁｪｯｸする
'　戻り値：true=取消し線あり、false=取消し線なし
Public Function isStrikethrough(ByRef cell As Object) As Boolean

    Dim strTemp As String
    Dim i As Long
    
    Dim resulfFlag As Boolean
    resulfFlag = False
    
    Select Case cell.Font.Strikethrough
            'セル全体に取消線の場合
            Case True
                'Debug.Print "Case True"
                resulfFlag = True
                
            'なにもなしの場合
            Case False
                'Debug.Print "Case False"
                resulfFlag = False
                
            'セル内の一部分に取消線の場合
            Case Else
                'Debug.Print "Case Else"
                '取消線のついていない文字を結合していく変数の初期化
                strTemp = ""
                
               'セル内の文字列を末尾から先頭に向かって調べていく
                For i = Len(cell.Value) To 1 Step -1
                    'ひと文字ずつ判定し、取消線がついていればtempに結合
                    If cell.Characters(Start:=i, Length:=1).Font.Strikethrough = False Then
                        '末尾から調べているので先頭に結合していく
                        strTemp = Mid(cell.Value, i, 1) & strTemp
                    Else
                        resulfFlag = True
                    End If
                Next i
                
        End Select
        
        isStrikethrough = resulfFlag

End Function

'セル内の文字に取消し線が存在するかﾁｪｯｸする
'　戻り値：true=取消し線あり、false=取消し線なし
Public Function removeStrikethrough(ByRef cell As Object) As String

    Dim strTemp As String
    Dim i As Long
    
    Dim strResult As String
    strResult = ""
    
    Select Case cell.Font.Strikethrough
            'セル全体に取消線の場合
            Case True
                'Debug.Print "Case True"
                strResult = ""
                
            'なにもなしの場合
            Case False
                'Debug.Print "Case False"
                strResult = cell.Value
                
            'セル内の一部分に取消線の場合
            Case Else
                'Debug.Print "Case Else"
                '取消線のついていない文字を結合していく変数の初期化
                strTemp = ""
                
               'セル内の文字列を末尾から先頭に向かって調べていく
                For i = Len(cell.Value) To 1 Step -1
                    'ひと文字ずつ判定し、取消線がついていればtempに結合
                    If cell.Characters(Start:=i, Length:=1).Font.Strikethrough = False Then
                        '末尾から調べているので先頭に結合していく
                        strTemp = Mid(cell.Value, i, 1) & strTemp
                    Else
                        resulfFlag = True
                    End If
                Next i
                strResult = strTemp
                
        End Select
        
        removeStrikethrough = strResult

End Function





