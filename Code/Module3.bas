Attribute VB_Name = "Module3"
Dim ERS, Retrytime As New ADODB.Recordset
Public EXCEL

Function ExcelControl(ByVal FN, ByVal strAddress, ByVal Marcro)
'Dim Excel1 As EXCEL.application
On Error GoTo ErrRecord

Dim EXCEL As Object
Set EXCEL = CreateObject("excel.Application")
With EXCEL
    .Workbooks.Open strAddress & FN
    .Visible = True
    .EnableCancelKey = xlErrorHandler
    Form2.RunS = 0
     Form2.RunM = 0
         
On Error GoTo ErrRecord
  
   DoEvents
    Form2.Timer3.Enabled = True

    .Run (Marcro) 'Marcro
    DoEvents
    Form2.Timer3.Enabled = True
Form2.Timer3.Enabled = False
.Application.DisplayAlerts = False

.ActiveWorkbook.Close True
.Application.DisplayAlerts = True
.Quit
End With

Set EXCEL = Nothing


Set ERS = New ADODB.Recordset

If Dir(strAddress & "KReSultforVB6.txt") <> "" Then
X = GetFileInfo(strAddress & "KReSultforVB6.txt")
X = Split(X, vbCrLf)

If X(0) <> "True" Then
ExcelControl = "False"

ERS.Open "select * from ���~�T��", Form2.CN, adOpenKeyset, adLockOptimistic

ERS.AddNew
ERS("�ɮצW��") = Form2.RS("�W��")
ERS("�o�ͮɶ�") = Format(Now, "YYYY-MM-DD HH:NN")
ERS("���~�X") = X(0)
If UBound(X) = 0 Then
ERS("���~����") = "������o���~��T" 'X(1)
Else
ERS("���~����") = X(1)
End If
ERS.Update
ERS.Close

Else
ExcelControl = "True"
End If
On Error Resume Next
Kill strAddress & "KReSultforVB6.txt"
On Error GoTo 0

Else
ExcelControl = "�䤣��Ұʵ��G�ɮ�"
End If
ErrRecord:
If EXCEL Is Nothing <> True Then
EXCEL.Application.DisplayAlerts = False

EXCEL.ActiveWorkbook.Close True
EXCEL.Application.DisplayAlerts = True
EXCEL.Quit
Set EXCEL = Nothing
End If
If Err <> 0 Then
ExcelControl = "�D�ɵo�Ϳ��~,�������᭫��"
ERS.Open "select * from ���~�T��", Form2.CN, adOpenKeyset, adLockOptimistic
ERS.AddNew
ERS("�ɮצW��") = Form2.RS("�W��")
ERS("�o�ͮɶ�") = Format(Now, "YYYY-MM-DD HH:NN")
ERS("���~�X") = Err.Number
ERS("���~����") = Err.Description
ERS.Update

ERS.Close
End If


End Function
