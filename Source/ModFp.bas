Attribute VB_Name = "ModFp"
Public objFrmMdi As frmMdiFp

Public objFrmComp As FrmComp
Public objFrmPr As frmPerson
Public objFrmSrch As frmSrch
Public objFrmType As frmTypesfp
Public objFrmPro As FrmProperty
Public objFrmlogin As frmLogin

Public LoginSucceeded As Boolean
Public Cn As ADODB.Connection

Public Sub MAIN()
    Set Cn = New ADODB.Connection
    Cn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Fp"
    Cn.Open
    
    Set objFrmMdi = New frmMdiFp
    Set objFrmComp = New FrmComp
    Set objFrmPr = New frmPerson
    Set objFrmSrch = New frmSrch
    Set objFrmType = New frmTypesfp
    Set objFrmPro = New FrmProperty
    Set objFrmlogin = New frmLogin
    
    objFrmMdi.Show
End Sub


