VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   7440
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "強制關閉"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   120
      Top             =   1800
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "已執行"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "分"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "限時:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "執行中程式:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S, M
Public sRS As New ADODB.Recordset
Public sCN As New ADODB.Connection

Private Sub Command1_Click()
Shell ("taskkill /f /im EXCEL.exe") '& Text1.Text)
End Sub

Private Sub Form_Load()
Set sRS = New ADODB.Recordset
Set sCN = New ADODB.Connection

ACDBPass = App.Path

dbName = ACDBPass & "\Record\" & "DataRecord.accdb"
bsql = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=" & dbName & _
            ";Jet OLEDB:Database Password=170145056789"
sCN.Open bsql
Timer1.Enabled = True
Timer2.Enabled = False
'sRS.Open "select *  from Record", sCN, adOpenKeyset, adLockReadOnly
M = 0
S = 0
   Call CheckProcess

End Sub

Private Sub Timer1_Timer()

  '  Call CheckProcess



End Sub

Private Sub Timer2_Timer()
S = S + 1
If S >= 60 Then
    M = M + 1
    S = 0
End If
Me.Label5.Caption = M & "分 " & S & "秒"



End Sub
Sub CheckProcess()
    sRS.Open "select *  from Record where [開啟成功] = '執行中'", sCN, adOpenKeyset, adLockReadOnly
    If sRS.RecordCount <> 0 Then
        Text1.Text = sRS("名稱").Value
        Label6.Caption = sRS("限時").Value
        
        
        
        
        Timer2.Enabled = True
        
        
    
    
    
    End If
    
    
    
    
End Sub

