Attribute VB_Name = "ET_TranslateObj"
'namespace=vba-files\main

Sub ET_Exe_Obj(control As IRibbonControl)

    On Error GoTo Err_CHK
    
    Dim sel As Variant
    Set sel = Selection
    Dim str As String
    Dim tmpStr As String
    
    'Rectangle ��ü�� ���
    If (TypeName(sel) = "Rectangle") Then
        'MsgBox ("Rectangle ��ü�Դϴ�.")
        
        '������� => �� ���� ������ ��Ȱ��
        
        '�������
        
        'Ÿ�ټ� (���� ��)
        'Dim tgtObj As Variant
        
        '1. ������ ����
        str = sel.Text
        str = repAP(str) '�߰�. 230222
        tmpStr = GTdo(str)
        
        str = str + vbCrLf + tmpStr
        
        '2. ��� ����
        sel.Text = str
        
    '�޸��� ��� Not (sel.Comment Is Nothing)
    ElseIf Not (sel.Comment Is Nothing) Then
        'MsgBox ("�޸� �ֽ��ϴ�. �޸� �����մϴ�.")
        
    '    '1. ������ ����
        str = sel.Comment.Text
        str = repAP(str) '�߰�. 230222
        'str = WorksheetFunction.EncodeURL(str)
        tmpStr = GTdo(str)
        tmpStr = vbCrLf + tmpStr
    
        'str = str + vbCrLf + tmpStr
    '
    '    '2. ��� ����
    '    'expression.Text(Text,Start,Overwrite) ' Text�� Attribute�� �ƴ϶� Method��
        'Call sel.Comment.Text(str)
        Call sel.Comment.Text(tmpStr, Start:=Len(str), Overwrite:=False)
    
    Else
        MsgBox ("Rectangle ��ü�� �ƴϰų� �޸� �����ϴ�. �����մϴ�.")
        Exit Sub
    End If
    
    'Application.ScreenUpdating = True
    MsgBox "Done"
    
Err_CHK:
        If Err.Number <> 0 Then
            MsgBox "�����߻�. ������ �����մϴ�."
        End If
End Sub
