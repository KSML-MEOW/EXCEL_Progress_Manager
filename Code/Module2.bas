Attribute VB_Name = "Module2"
Public Function GetFileInfo(szPath As String) As String
1        Dim sfileLen As Long
2        Dim FileBuffer() As Byte
3        Dim szRplyData As String
    Dim fileNum As Integer
    Dim s As String
    
5        On Error GoTo hErr
6        sfileLen = 0
    
    
    
    
      
12       If Len(Dir(szPath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly)) = 0 Then
                         
14          ' TraceOut "", Error
15           Exit Function
      
17       End If
    
19       '  If Mid(Trim(strName), 1, 2) = "T5" Or Mid$(Trim(strName), 1, 2) = "T3" Then
20       sfileLen = FileLen(szPath)
        
22       If sfileLen < 1 Then
23           sfileLen = 0
24          'TraceOut "File Len <1(" & szPath & "),Exit funciton", Error
25           Exit Function
26       End If

28       ReDim FileBuffer(sfileLen - 1) As Byte
        
30       fileNum = FreeFile()
31       'On Error GoTo GTFError
32       Open szPath For Binary As fileNum
33       Get fileNum, , FileBuffer
34       Close fileNum
35       s = StrConv(FileBuffer, vbUnicode)
        
37       If Right(s, 2) = vbCrLf Then
38           szRplyData = Left(s, Len(s) - 2)
             
40           GetFileInfo = szRplyData
       
42           Exit Function
       
44       Else
45           szRplyData = s
46           GetFileInfo = szRplyData
47           Exit Function
        
            
            
51       End If
    
53 hErr:

55
56       Err.Clear
 
58       GetFileInfo = ""
 
60       Exit Function

End Function
Public Sub Delay(D_Long As Date)
Dim DelayTime

DelayTime = DateAdd("s", D_Long, Now)
While DelayTime > Now
DoEvents '讓 windows 去處理其他事
Wend
End Sub
