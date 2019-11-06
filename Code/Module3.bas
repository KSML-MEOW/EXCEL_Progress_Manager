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

ERS.Open "select * from 錯誤訊息", Form2.CN, adOpenKeyset, adLockOptimistic

ERS.AddNew
ERS("檔案名稱") = Form2.RS("名稱")
ERS("發生時間") = Format(Now, "YYYY-MM-DD HH:NN")
ERS("錯誤碼") = X(0)
If UBound(X) = 0 Then
ERS("錯誤註解") = "未能取得錯誤資訊" 'X(1)
Else
ERS("錯誤註解") = X(1)
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
ExcelControl = "找不到啟動結果檔案"
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
ExcelControl = "主檔發生錯誤,五分鐘後重試"
ERS.Open "select * from 錯誤訊息", Form2.CN, adOpenKeyset, adLockOptimistic
ERS.AddNew
ERS("檔案名稱") = Form2.RS("名稱")
ERS("發生時間") = Format(Now, "YYYY-MM-DD HH:NN")
ERS("錯誤碼") = Err.Number
ERS("錯誤註解") = Err.Description
ERS.Update

ERS.Close
End If


End Function
