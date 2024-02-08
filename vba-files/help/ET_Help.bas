Attribute VB_Name = "ET_Help"
'namespace=vba-files\help\

Sub ET_Help(control As IRibbonControl)

    Call MsgBox( _
        "번역대상 언어가 있는 셀을 선택하고 버튼을 누르면 번역됩니다." _
        + vbCrLf _
        + "여러 셀도 선택 가능합니다." _
        + vbCrLf _
        + "버튼에 따라 해당 셀/우측 셀/하단 셀에 번역결과를 출력합니다." _
        + vbCrLf _
        + "셀이 아닌 도형개체의 언어를 번역할 수도 있습니다." _
        + vbCrLf + vbCrLf _
        + "문의처 : kinphw@naver.com(박병) " _
        + vbCrLf _
        + "버전 : 0.2.0", , "엑셀번역기")
    
End Sub
    
