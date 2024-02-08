Attribute VB_Name = "ET_SetInfo"
'namespace=vba-files\set\

Option Explicit

Private rbnET As IRibbonUI

Public btnLangIndex As Long
Private btnLangLabels(2) As String

Public btnDirIndex As Long
Private btnDirLabels(3) As String
Private btnDirImages(3) As String

Public Sub ET_onLoad(ribbonUI As IRibbonUI)

    Set rbnET = ribbonUI
    
    btnLangIndex = 0
    
    btnLangLabels(0) = "¿µÇÑ"
    btnLangLabels(1) = "ÇÑ¿µ"

    btnDirIndex = 0
    
    btnDirLabels(0) = "ÇöÀç¼¿"
    btnDirLabels(1) = "¿ìÃø¼¿"
    btnDirLabels(2) = "¾Æ·¡¼¿"

End Sub

Public Sub GetLabelLang(ctrl As IRibbonControl, ByRef value)
    value = btnLangLabels(btnLangIndex)
End Sub

Public Sub GetLabelDir(ctrl As IRibbonControl, ByRef value)
    value = btnDirLabels(btnDirIndex)
End Sub

Public Sub SetLang1(ctrl As IRibbonControl)
    btnLangIndex = 0
    Call rbnET.Invalidate
End Sub

Public Sub SetLang2(ctrl As IRibbonControl)
    btnLangIndex = 1
    Call rbnET.Invalidate
End Sub

Public Sub SetDir1(ctrl As IRibbonControl)
    btnDirIndex = 0
    Call rbnET.Invalidate
End Sub

Public Sub SetDir2(ctrl As IRibbonControl)
    btnDirIndex = 1
    Call rbnET.Invalidate
End Sub

Public Sub SetDir3(ctrl As IRibbonControl)
    btnDirIndex = 2
    Call rbnET.Invalidate
End Sub
