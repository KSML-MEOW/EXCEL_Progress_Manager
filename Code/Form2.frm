VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "�`��"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form2"
   ScaleHeight     =   5520
   ScaleWidth      =   14985
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14520
      Top             =   3600
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   14280
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   1815
      Left            =   14640
      TabIndex        =   4
      Top             =   1440
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refesh"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   14520
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   14520
      Top             =   3960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�s�W"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   7858
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   16750899
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RS, Irs As New ADODB.Recordset

Public CN As New ADODB.Connection
Public SELC, MaxNum, M
Public RunS, RunM
Dim NRowSel, NColSel


Private Sub Command1_Click()
Timer1.Enabled = False
If RS.RecordCount <> 0 Then

RS.MoveLast

MaxNum = RS("�ѧO�X")
SELC = ""
End If
RS.Close
Form1.Show
End Sub




Private Sub Command2_Click()
Call ReFreshRS
End Sub

Private Sub Form_Initialize()
Set RS = New ADODB.Recordset
Set CN = New ADODB.Connection
M = 0

ACDBPass = App.Path

dbName = ACDBPass & "\Record\" & "DataRecord.accdb"
bsql = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & dbName & _
            ";Jet OLEDB:Database Password=170145056789"
CN.Open bsql


RS.Open "select * from Record", CN, adOpenKeyset, adLockOptimistic

With MSHFlexGrid1
Set .DataSource = RS
.ColWidth(0) = 0
.ColWidth(1) = 2500
.ColWidth(2) = 1800
.ColWidth(3) = 3500
.ColWidth(4) = 0
.ColWidth(5) = 1500
.ColWidth(6) = 1500
.ColWidth(7) = 2000
.ColWidth(8) = 1000
End With



End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MSHFlexGrid1_Click()
'SelCellrow = MSHFlexGrid1.Row
''SelCellcol = MSHFlexGrid1.col
'
'
''MSHFlexGrid1.Row = NRowSel
'Call GridColorClear
'Call GridColor(SelCellrow)
'
'
''NRowSel = SelCellrow
''NColSel = SelCellcol

End Sub

Sub GridColorClear()
Ocol = MSHFlexGrid1.Col
ORow = MSHFlexGrid1.Row
For j = 1 To MSHFlexGrid1.Rows - 1
MSHFlexGrid1.Row = j
For i = 1 To MSHFlexGrid1.Cols - 1
MSHFlexGrid1.Col = i
MSHFlexGrid1.CellBackColor = &H80000005
Next
Next
 MSHFlexGrid1.Row = ORow
MSHFlexGrid1.Col = Ocol
End Sub
Sub GridColor(ByVal SelCellrow)
Ocol = MSHFlexGrid1.Col
ColorChange = 0
MSHFlexGrid1.Row = SelCellrow

For i = 0 To MSHFlexGrid1.Cols - 1
MSHFlexGrid1.Col = i
MSHFlexGrid1.CellBackColor = &H8000000D

Next
MSHFlexGrid1.Col = Ocol
End Sub

 Sub MSHFlexGrid1_DblClick()
 If MSHFlexGrid1.Col < 7 Then
 
 RS.Close
 Timer1.Enabled = False
MSHFlexGrid1.Col = 1
SELC = MSHFlexGrid1.Text
Timer1.Enabled = False
Form1.Show
Else
MSHFlexGrid1.Col = 1
SELC = MSHFlexGrid1.Text
Form3.Show
SELC = ""
End If



End Sub


Private Sub MSHFlexGrid1_SelChange()
SelCellrow = MSHFlexGrid1.Row
'SelCellcol = MSHFlexGrid1.col


'MSHFlexGrid1.Row = NRowSel
Call GridColorClear
Call GridColor(SelCellrow)
End Sub

Private Sub Timer1_Timer()
M = M + 1
ErrMessage = 0
RS.MoveFirst
RowNow = 1
Do Until RS.EOF
DoEvents
    If Now >= RS("�U���}�Үɶ�") And RS("���~����") < 5 Then
        If RS("�}�Ұ_�l�ɶ�") <> "" Or IsNull(RS("�}�Үɶ����j")) Then
                tmpTime = DateValue(Format(DateAdd("d", 1, Date), "YYYY/MM/DD")) + TimeValue(RS("�}�Ұ_�l�ɶ�"))

    Else
    tmpTime = Now + TimeSerial(0, RS("�}�Үɶ����j") - 1, 60 - Second(Now))
    End If
        RS("�}�Ҧ��\") = "���椤"
        RS.Update
         Shell "cmd /c" & App.Path & "\TimeClose.exe"
        X = ExcelControl(RS("�W��"), RS("�ɮצ�m"), RS("���楨��"))
'Do Until Dir(strAddress & "KReSultforVB6.txt") <> ""
'M = M + 1
'If M > 2 Then
'
'
'End If
'
'
'
'Loop
    X = Split(X, vbCrLf)
    RS("�}�Ҧ��\") = CStr(Trim(X(0)))
   If RS("�}�Ҧ��\") <> "True" And RS("���~����") < 5 Then
           M = 5
            tmpTime = Now + TimeSerial(0, 5 - 1, 60 - Second(Now))
            RS("���~����") = RS("���~����") + 1
End If
RS("�U���}�Үɶ�") = tmpTime
    RS.Update

     End If
        RS.MoveNext
        RowNow = RowNow + 1
Loop
Call ReFreshRS
If M >= 5 Then
For i = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(i, 7) <> "True" Then
    ErrProcess = ErrProcess & MSHFlexGrid1.TextMatrix(i, 1) & ", "
    ErrMessage = 1
    MSHFlexGrid1.Row = i
    SelCellrow = MSHFlexGrid1.Row
Call GridColor(SelCellrow)
End If
Next

If ErrMessage <> 0 Then
ErrProcess = Left(ErrProcess, Len(ErrProcess) - 2)
Call UPload(ErrProcess)
Call GridColorClear
End If
M = 0
End If

End Sub

Private Sub Timer2_Timer()

'Call GridColor(MSHFlexGrid1.Row)
DoEvents
Label1.Caption = Now

End Sub
Sub ReFreshRS()
Set RS = New ADODB.Recordset
RS.Open "select * from Record", CN, adOpenKeyset, adLockOptimistic

With MSHFlexGrid1
Set .DataSource = RS
.ColWidth(0) = 0
.ColWidth(1) = 2500
.ColWidth(2) = 1800
.ColWidth(3) = 3500
.ColWidth(4) = 0
.ColWidth(5) = 1500
.ColWidth(6) = 1500
.ColWidth(7) = 2000
.ColWidth(8) = 1000
End With
End Sub

Private Sub Timer3_Timer()
RunS = RunS + 1
If RunS >= 60 Then
RunS = 0
RunM = RunM + 1

End If
Debug.Print RunM & " " & RumS
If RumS > 30 Then MsgBox "�W��"
End Sub

Private Sub UpDown1_DownClick()
If MSHFlexGrid1.Row = RS.RecordCount Then
MsgBox "�w�g�O�̫�@��!"
Else

Set rsMove = New ADODB.Recordset

MSHFlexGrid1.Col = 0
MoveFrom = MSHFlexGrid1.Text
rsMove.Open "select * from Record ", CN, adOpenKeyset, adLockReadOnly
rsMove.Filter = "�ѧO�X = '" & MoveFrom + 1 & "'"
RS.Filter = "�ѧO�X = '" & MoveFrom & "'"
tmpName = RS("�W��")
tmpProcess = RS("���楨��")
tmpAddress = RS("�ɮצ�m")
tmpOpenTime = RS("�}�Ұ_�l�ɶ�")
tmpInterval = RS("�}�Үɶ����j")
tmpNextTime = RS("�U���}�Үɶ�")
tmpSucess = RS("�}�Ҧ��\")
tmpErrorTime = RS("���~����")

RS("�W��") = rsMove("�W��")
RS("���楨��") = rsMove("���楨��")
RS("�ɮצ�m") = rsMove("�ɮצ�m")
RS("�}�Ұ_�l�ɶ�") = rsMove("�}�Ұ_�l�ɶ�")
RS("�}�Үɶ����j") = rsMove("�}�Үɶ����j")
RS("�U���}�Үɶ�") = rsMove("�U���}�Үɶ�")
RS("�}�Ҧ��\") = rsMove("�}�Ҧ��\")
RS("���~����") = rsMove("���~����")

rsMove.Close
RS.Filter = "�ѧO�X = '" & MoveFrom + 1 & "'"

RS("�W��") = tmpName
RS("���楨��") = tmpProcess
RS("�ɮצ�m") = tmpAddress
RS("�}�Ұ_�l�ɶ�") = tmpOpenTime
RS("�}�Үɶ����j") = tmpInterval
RS("�U���}�Үɶ�") = tmpNextTime
RS("�}�Ҧ��\") = tmpSucess
RS("���~����") = tmpErrorTime


RS.Update


Call ReFreshRS
Call GridColorClear
MSHFlexGrid1.Row = MSHFlexGrid1.Row + 1
X = MSHFlexGrid1.Row

Call GridColor(X)
End If

End Sub

Private Sub UpDown1_UpClick()
If MSHFlexGrid1.Row = 1 Then
MsgBox "�w�g�O�Ĥ@��!"
Else

Set rsMove = New ADODB.Recordset

MSHFlexGrid1.Col = 0
MoveFrom = MSHFlexGrid1.Text
rsMove.Open "select * from Record ", CN, adOpenKeyset, adLockReadOnly
rsMove.Filter = "�ѧO�X = '" & MoveFrom - 1 & "'"
RS.Filter = "�ѧO�X = '" & MoveFrom & "'"
tmpName = RS("�W��")
tmpProcess = RS("���楨��")
tmpAddress = RS("�ɮצ�m")
tmpOpenTime = RS("�}�Ұ_�l�ɶ�")
tmpInterval = RS("�}�Үɶ����j")
tmpNextTime = RS("�U���}�Үɶ�")
tmpSucess = RS("�}�Ҧ��\")
tmpErrorTime = RS("���~����")

RS("�W��") = rsMove("�W��")
RS("���楨��") = rsMove("���楨��")
RS("�ɮצ�m") = rsMove("�ɮצ�m")
RS("�}�Ұ_�l�ɶ�") = rsMove("�}�Ұ_�l�ɶ�")
RS("�}�Үɶ����j") = rsMove("�}�Үɶ����j")
RS("�U���}�Үɶ�") = rsMove("�U���}�Үɶ�")
RS("�}�Ҧ��\") = rsMove("�}�Ҧ��\")
RS("���~����") = rsMove("���~����")

rsMove.Close
RS.Filter = "�ѧO�X = '" & MoveFrom - 1 & "'"

RS("�W��") = tmpName
RS("���楨��") = tmpProcess
RS("�ɮצ�m") = tmpAddress
RS("�}�Ұ_�l�ɶ�") = tmpOpenTime
RS("�}�Үɶ����j") = tmpInterval
RS("�U���}�Үɶ�") = tmpNextTime
RS("�}�Ҧ��\") = tmpSucess
RS("���~����") = tmpErrorTime


RS.Update


Call ReFreshRS
Call GridColorClear
MSHFlexGrid1.Row = MSHFlexGrid1.Row - 1

'MSHFlexGrid1.BackColorSel

X = MSHFlexGrid1.Row

Call GridColor(X)

End If

End Sub
