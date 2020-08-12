Attribute VB_Name = "chartopen"
Option Explicit

Public Sub chartopen()
    
    Call frm_LinReg_Wks.OK_Btn_Click
    'Unload frm_LinReg_Wks
    frm_RegChart.Show

End Sub
