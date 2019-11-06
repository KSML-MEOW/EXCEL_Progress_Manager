VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "設定"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   5130
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "刪除"
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   3960
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   450
      Left            =   3720
      Pattern         =   "*.txt"
      TabIndex        =   11
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   390
      Left            =   1440
      TabIndex        =   10
      Top             =   2505
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   4320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上傳"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "間隔(分):"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "固定開啟時間:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "檔案位置"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "執行巨集"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "名稱:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public s, M, H, TT
Dim sRS, Irs As New ADODB.Recordset
Dim sCN As New ADODB.Connection
Dim DRS As New ADODB.Recordset
Dim NewRs, SELC

Private Sub Command1_Click()

Set DRS = New ADODB.Recordset
DRS.Open "DELETE FROM Record WHERE 名稱 = '" & Text1.Text & "'", sCN, adOpenKeyset, adLockOptimistic

Set DRS = Nothing
'Call Form2.ReFreshRS

Unload Me
'If Combo1.Text <> "記錄檔" Then
'Data = GetFileInfo(App.Path & "\Record\" & Combo1.Text & ".txt")
'Data = Split(Data, vbCrLf)

'Text1.Text = Combo1.Text
'Text2.Text = Data(0)
'Text3.Text = Data(1)
'Text4.Text = Data(2)
'Text5.Text = Data(3)
'Text6.Text = Data(4)
'End If



End Sub



Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" And (Text6.Text = "" And Text5.Text = "") Then
MsgBox "請填入完整資訊"
Exit Sub
End If
If NewRs = 1 Then
sRS.AddNew
sRS("識別碼") = Form2.MaxNum + 1
End If

If Text5.Text <> "" Then
sRS("下次開啟時間") = DateValue(Format(Date + 1, "YYYY/MM/DD")) + TimeValue(Text5.Text)
Else
sRS("下次開啟時間") = Now '+ TimeSerial(0, CInt(Text6.Text), 0)
End If


    sRS("名稱") = Text1.Text
     sRS("執行巨集") = Text2.Text
     sRS("檔案位置") = Text3.Text
     'sRS("圖片目錄") = Text4.Text
     sRS("開啟起始時間") = Text5.Text
     sRS("開啟時間間隔") = Text6.Text
 
    sRS("錯誤次數") = 0
    
    sRS.Update
    
    
    Set Irs = New ADODB.Recordset
 
    Irs.Open "select * from UploadInf ", sCN, adOpenKeyset, adLockOptimistic
    Irs.Filter = "上傳名稱 = '" & SELC & "'"
    If Irs.RecordCount = 0 Then
    
 Irs.AddNew
 Irs("上傳名稱") = SELC
 Irs("下次上傳時間") = Now '+ TimeSerial(0, CInt(Text6.Text), 0)
Else
 
 Irs("下次上傳時間") = Now '+ TimeSerial(0, CInt(Text6.Text), 0)
End If
 
Irs.Update
'Set RS = New ADODB.Recordset
' RS.Open "select 名稱, 交談室編號, 訊息內容, 上傳起始時間, 上傳時間間隔,下次上傳時間  from Record", sCN, adOpenKeyset, adLockReadOnly
'Set Form2.MSHFlexGrid1.DataSource = RS
'RS.Close


Form2.MSHFlexGrid1.Refresh
Unload Form1





'Timer1.Enabled = True
'Open (App.Path & "\Record\" & Text1.Text & ".txt") For Output As #27
'Print #27, Text2.Text & vbCrLf & _
                Text3.Text & vbCrLf & _
                Text4.Text & vbCrLf & _
                Text5.Text & vbCrLf & _
                Text6.Text & vbCrLf
'                Close #27
'                File1.Path = App.Path & "\Record"
'                Combo1.Clear
'If File1.ListCount <> Empty Then
'For i = 0 To File1.ListCount - 1

'File1.ListIndex = i
'Combo1.AddItem Split(File1.filename, ".")(0), i

'Next
'End If
'        Combo1.ListIndex = -1
'       Combo1.Text = "記錄檔"
       TT = 1
End Sub

Private Sub Form_Load()

Set sRS = New ADODB.Recordset
Set sCN = New ADODB.Connection

ACDBPass = App.Path

dbName = ACDBPass & "\Record\" & "DataRecord.accdb"
bsql = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & dbName & _
            ";Jet OLEDB:Database Password=170145056789"
sCN.Open bsql


sRS.Open "select *  from Record", sCN, adOpenKeyset, adLockOptimistic

If Form2.SELC = "" Then
NewRs = 1
Else
SELC = Form2.SELC
NewRs = 0
sRS.Filter = "名稱 = '" & Form2.SELC & "'"
    Text1.Text = Form2.SELC
    Text2.Text = sRS("執行巨集")
    Text3.Text = sRS("檔案位置")
    'Text4.Text = sRS("圖片目錄")
    Text5.Text = sRS("開啟起始時間")
    Text6.Text = sRS("開啟時間間隔")
     
    
End If



'Label8.Caption = Format(Date, "YYYY-MM-DD")
'If Dir(App.Path & "\Record", vbDirectory) = "" Then MkDir (App.Path & "\Record")
'File1.Path = App.Path & "\Record"
'If File1.ListCount <> Empty Then
'For i = 0 To File1.ListCount - 1

'File1.ListIndex = i
'Combo1.AddItem Split(File1.filename, ".")(0), i

'Next
'End If




End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set sRS = Nothing
Set Irs = Nothing
Set sCN = Nothing
Set NewRs = Nothing
Set SELC = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)

Form2.RS.Open "select * from Record", Form2.CN, adOpenKeyset, adLockOptimistic
Set Form2.MSHFlexGrid1.DataSource = Form2.RS
Form2.Show
Form2.Timer1.Enabled = True
End Sub

