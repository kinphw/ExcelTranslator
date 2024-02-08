Attribute VB_Name = "ET_TranslateObj"
'namespace=vba-files\main

Sub ET_Exe_Obj(control As IRibbonControl)

    On Error GoTo Err_CHK
    
    Dim sel As Variant
    Set sel = Selection
    Dim str As String
    Dim tmpStr As String
    
    'Rectangle 객체인 경우
    If (TypeName(sel) = "Rectangle") Then
        'MsgBox ("Rectangle 개체입니다.")
        
        '번역대상 => 셀 넣을 값으로 재활용
        
        '번역결과
        
        '타겟셀 (넣을 셀)
        'Dim tgtObj As Variant
        
        '1. 번역값 생성
        str = sel.Text
        str = repAP(str) '추가. 230222
        tmpStr = GTdo(str)
        
        str = str + vbCrLf + tmpStr
        
        '2. 결과 투입
        sel.Text = str
        
    '메모인 경우 Not (sel.Comment Is Nothing)
    ElseIf Not (sel.Comment Is Nothing) Then
        'MsgBox ("메모가 있습니다. 메모를 번역합니다.")
        
    '    '1. 번역값 생성
        str = sel.Comment.Text
        str = repAP(str) '추가. 230222
        'str = WorksheetFunction.EncodeURL(str)
        tmpStr = GTdo(str)
        tmpStr = vbCrLf + tmpStr
    
        'str = str + vbCrLf + tmpStr
    '
    '    '2. 결과 투입
    '    'expression.Text(Text,Start,Overwrite) ' Text는 Attribute가 아니라 Method임
        'Call sel.Comment.Text(str)
        Call sel.Comment.Text(tmpStr, Start:=Len(str), Overwrite:=False)
    
    Else
        MsgBox ("Rectangle 개체가 아니거나 메모가 없습니다. 종료합니다.")
        Exit Sub
    End If
    
    'Application.ScreenUpdating = True
    MsgBox "Done"
    
Err_CHK:
        If Err.Number <> 0 Then
            MsgBox "오류발생. 구문을 종료합니다."
        End If
End Sub
