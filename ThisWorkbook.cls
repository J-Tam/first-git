VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()

    Application.OnKey "{F1}", ""
    Application.OnKey "^k", "frmDiffDataShow"
    debug.print "test"

End Sub

''[Ctrl]+[K]のショートカット割り当てをExcel標準に戻す
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.OnKey "^k"
End Sub


