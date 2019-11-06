Attribute VB_Name = "Module1"
Private Declare Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
    ByVal hwnd As Long, _
    ByVal lpText As String, _
    ByVal lpCaption As String, _
    ByVal uType As Long, _
    ByVal wlange As Long, _
    ByVal dwTimeout As Long) As Long
Public tmpTime

 Private Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



  Sub UPload(ByVal ErrProcess)   ' FOR OFFICE
    Dim Name_arr()
    Dim Account, APIkey, CHAT1, CHAT2, Message, CONType
        Dim sPathAndFile As String
 Dim Iurl, MyName, filename, XX, IData, Respo, url, Data, AlarmID, AlarmMessage, Pic As String
 Dim T, N, ssTabPage
On Error Resume Next
    '-------------------------------帳號、API---------------
Account = "api_lcd5"
    APIkey = "FBA5C1C4-C736-0729-5599-D0941022A680"

    '-------------------------------------------------------------
    '*****************聊天室編號**********************
    CHAT1 = 66040



    

    ' ----Message
    Message = Format(Now, "YYYY-MM-DD HH:MM:DD") & "  " & "程式錯誤: " & ErrProcess
    '------------
    '************************************************
'
'IntR = SetWindowPos(Form2.hwnd, -1, 0, 0, 0, 0, 3)
'    ' Get the whole form inclusing borders, caption,...'
'    Set Form2.Picture1.Picture = CaptureForm(Form2)
'
'    If Dir(App.Path & "\" & Format(Date, "YYYYMMDD"), vbDirectory) = "" Then MkDir App.Path & "\" & Format(Date, "YYYYMMDD")
'
'
'    sPathAndFile = App.Path & "\" & Format(Date, "YYYYMMDD") & "\" & "Error.jpg"
''MsgBox App.Path
'
'   'SavePicture Picture1, sPathAndFile
'Call BMPtoJPG.SaveJPG(Form2.Picture1.Picture, sPathAndFile, 50)
'    '---------------------------------------------------------------------------'--------------------------------------------------截圖
'
'IntR = SetWindowPos(Form2.hwnd, -2, 0, 0, 0, 0, 3)    '關閉視窗最上層
'
'
'
'    '-----------------------------------------傳圖--------------------
'
'    Iurl = "http://mapp.innolux.com/teamplus_innolux/API/IMService.ashx?ask=uploadFile"
'
'    MyName = Dir(App.Path & "\" & Format(Date, "YYYYMMDD") & "\" & "*.jpg")
'    j = 0
'    Do While MyName <> ""
'
'        DoEvents
'
'
'
'
'
'        filename = Replace(MyName, App.Path & "\", "")
'        XX = convertImageToBase64(App.Path & "\" & Format(Date, "YYYYMMDD") & "\" & MyName)
'
'
'        XX = Replace(XX, "+", "%2B")
'
'        'XX = Replace(XX, "=", "%3D")
'        XX = Replace(XX, "/", "%2F")
'
'
'
'
'
'
'
'
'        IData = "account=" & Account & "&api_key=" & APIkey & "&file_type=" & Split(filename, ".")(1) & "&data_binary=" & XX
'
'
'
'
'        'Debug.Print data
'
'
'
'        Respo = MakeWebRequest(Iurl, IData)
'
'
'        T = Split(Respo, ":")
'        T = Split(T(2), ",")
'        T = Split(T(0), """")
'        If T(0) = 0 Then
'            'MsgBox Respo
'
'
'        Else
'
'           ' MsgBox "圖片上傳失敗"
'            End
'        End If
'
'
'
'        N = Split(Respo, "FileName"":""")
'        'MsgBox RESPO
'
'        ReDim Preserve Name_arr(j)
'        Name_arr(j) = Split(Replace(N(1), """", ""), ",")(0)
'
'        j = j + 1
'        MyName = Dir   ' 尋找下一個目錄
'    Loop
    '--------------------------------------------------------------------------------------






    '-----------------------------訊息送出-----------------------------------

    url = "http://mapp.innolux.com/teamplus_innolux/API/IMService.ashx?ask=sendChatMessage"



    CONType = 1        ' 1:純文字  2:圖片


    Select Case CONType
    Case 1
        'Message 訊息
        Data = "account=" & Account & "&api_key=" & APIkey & "&chat_sn=" & CHAT1 & _
        "&content_type=" & CONType & "&msg_content=" & Message

        Respo = MakeWebRequest(url, Data)

    Case 2                                                                                                        'PIC 回傳的檔案名稱



        Data = "account=" & Account & "&api_key=" & APIkey & "&chat_sn=" & CHAT1 & _
        "&content_type=" & 1 & "&msg_content=" & Message

        Respo = MakeWebRequest(url, Data)

        For i = 0 To UBound(Name_arr)
            Data = "account=" & Account & "&api_key=" & APIkey & "&chat_sn=" & CHAT1 & _
            "&content_type=" & CONType & "&msg_content=" & Name_arr(i) & "&file_show_name=" & URLEncode(Message) & Format(i, "00") & ".jpg"

            'Debug.Print data
            Respo = MakeWebRequest(url, Data)


        Next
    End Select


 



    Respo = Split(Respo, ":")
    'Respo = Split(Respo(1), ",")
    Respo = Left(Respo(3), 4)
    If Respo = "true" Then
        'MsgBox Respo
'MessageBoxTimeout ByVal 0&, "上傳完畢(生產日報)", "注意", ByVal 48&, ByVal 0&, 3000

Else
'MsgBox "上傳失敗"
'MessageBoxTimeout ByVal 0&, "上傳失敗(生產日報)", "注意", ByVal 48&, ByVal 0&, 3000
End
    End If
 
    MessageBoxTimeout ByVal 0&, "上傳完畢", "注意", ByVal 48&, ByVal 0&, 3000


  

    Kill App.Path & "\" & Format(Date, "YYYYMMDD") & "\*.*"
    RmDir (App.Path & "\" & Format(Date, "YYYYMMDD"))
On Error GoTo 0

End Sub

Public Function MakeWebRequest(url, post_data) As String
    ' make sure to include the Microsoft WinHTTP Services in the project
    ' tools -> references -> Microsoft WinHTTP Services, version 5.1
    ' http://www.808.dk/?code-simplewinhttprequest
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa384106(v=vs.85).aspx
    ' http://www.neilstuff.com/winhttp/

    ' create the request object
    'Set req = CreateObject("MSXML2.ServerXMLHTTP")
   On Error Resume Next
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")

    req.SetTimeouts 60000, 60000, 60000, 60000

    '
    req.Open "POST", url, False


        req.SetRequestHeader "Content-type", "application/x-www-form-urlencoded"

X:

    req.Send post_data

    If Err = -2147012894 Then
    req.Send post_data
Err = 0
GoTo X
End If
    ' read response and return
    MakeWebRequest = req.ResponseText

End Function
Public Function convertImageToBase64(filePath)
  Dim inputStream
  Set inputStream = CreateObject("ADODB.Stream")
  inputStream.Open
  inputStream.Type = 1  ' adTypeBinary
  inputStream.LoadFromFile filePath
  Dim bytes: bytes = inputStream.Read
  Dim dom: Set dom = CreateObject("Microsoft.XMLDOM")
  Dim elem: Set elem = dom.createElement("tmp")
  elem.DataType = "bin.base64"
  elem.nodeTypedValue = bytes
    'convertImageToBase64 = elem.Text
     convertImageToBase64 = Replace(elem.Text, vbLf, "")
' convertImageToBase64 = "data:image/png;base64," & Replace(elem.Text, vbLf, "")
End Function

Public Function URLEncode( _
   ByVal StringVal, _
   Optional SpaceAsPlus As Boolean = False _
) As String
  Dim bytes() As Byte, b As Byte, i As Long, space As String

  If SpaceAsPlus Then space = "+" Else space = "%20"

  If Len(StringVal) > 0 Then
    With New ADODB.Stream
      .Mode = adModeReadWrite
      .Type = adTypeText
      .Charset = "UTF-8"
      .Open
      .WriteText StringVal
      .Position = 0
      .Type = adTypeBinary
      .Position = 3 ' skip BOM
      bytes = .Read
    End With

    ReDim result(UBound(bytes)) As String

    For i = UBound(bytes) To 0 Step -1
      b = bytes(i)
      Select Case b
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Chr(b)
        Case 32
          result(i) = space
        Case 0 To 15
          result(i) = "%0" & Hex(b)
        Case Else
          result(i) = "%" & Hex(b)
      End Select
    Next i

    URLEncode = Join(result, "")
  End If
End Function


