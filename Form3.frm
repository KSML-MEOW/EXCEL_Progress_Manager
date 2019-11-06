VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form3 
   Caption         =   "錯誤記錄"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   LinkTopic       =   "Form3"
   ScaleHeight     =   4980
   ScaleWidth      =   15360
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "記錄刪除"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6376
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "程式名稱:"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ERS As New ADODB.Recordset

Dim DRS As New ADODB.Recordset

Private Sub Command1_Click()
Set DRS = New ADODB.Recordset
DRS.Open "DELETE FROM 錯誤訊息 WHERE 檔案名稱 = '" & Text1.Text & "'", Form2.CN, adOpenKeyset, adLockOptimistic

Set DRS = Nothing
Set DRS = New ADODB.Recordset
Form2.CN.Execute "Update [Record] SET [錯誤次數] = 0 where [名稱] = '" & Text1.Text & "'"
'DRS.Open "Update  [Record] set [錯誤次數] = '0' where [檔案名稱] = '" & Text1.Text & "'", Form2.CN, adOpenKeyset, adLockOptimistic

Set DRS = Nothing



'Call Form2.ReFreshRS

Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = Form2.SELC

Set ERS = New ADODB.Recordset
ERS.Open "select * from 錯誤訊息 Where 檔案名稱 = '" & Form2.SELC & "'", Form2.CN, adOpenKeyset, adLockReadOnly

If ERS.RecordCount = 0 Then

Else

With MSHFlexGrid1
Set .DataSource = ERS
.FixedCols = 0
.ColWidth(0) = 0
.ColWidth(1) = 2500
.ColWidth(2) = 1500
.ColWidth(3) = 1000
.ColWidth(4) = 12500
End With

End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
Set DRS = Nothing
Set ERS = Nothing
End Sub
