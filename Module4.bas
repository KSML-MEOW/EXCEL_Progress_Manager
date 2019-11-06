Attribute VB_Name = "Module4"
Option Base 0
'
' This module contains several routines for capturing windows into a
' picture.  All routines have palette support.
'
' CreateBitmapPicture   - Creates a picture object from a bitmap and palette.
' CaptureWindow         - Captures any window given a window handle.
' CaptureActiveWindow   - Captures the active window on the desktop.
' CaptureForm           - Captures the entire form.
' CaptureClient         - Captures the client area of a form.
' CaptureScreen         - Captures the entire screen.
' PrintPictureToFitPage - prints any picture as big as possible on the page.
'
Private Type PALETTEENTRY
    peRed   As Byte
    peGreen As Byte
    peBlue  As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors
End Type

Private Type GUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type PicBmp
    Size As Long
    Type As Long
        hBmp As Long
        hPal As Long
        Reserved As Long
    End Type
    
    Private Const RASTERCAPS As Long = 38
    Private Const RC_PALETTE As Long = &H100
    Private Const SIZEPALETTE As Long = 104
    '
    ' DC = Device Context
    '
    ' Creates a bitmap compatible with the device associated
    ' with the specified DC.
    Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
    ByVal hDC As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long) As Long
    
    ' Retrieves device-specific information about a specified device.
    Private Declare Function GetDeviceCaps Lib "GDI32" ( _
    ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
    
    ' Retrieves a range of palette entries from the system palette
    ' associated with the specified DC.
    Private Declare Function GetSystemPaletteEntries Lib "GDI32" ( _
    ByVal hDC As Long, ByVal wStartIndex As Long, _
    ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) _
    As Long
    
    ' Creates a memory DC compatible with the specified device.
    Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
    ByVal hDC As Long) As Long
    
    ' Creates a logical color palette.
    Private Declare Function CreatePalette Lib "GDI32" ( _
    lpLogPalette As LOGPALETTE) As Long
    
    ' Selects the specified logical palette into a DC.
    Private Declare Function SelectPalette Lib "GDI32" ( _
    ByVal hDC As Long, ByVal hPalette As Long, _
    ByVal bForceBackground As Long) As Long
    
    ' Maps palette entries from the current logical
    ' palette to the system palette.
    Private Declare Function RealizePalette Lib "GDI32" ( _
    ByVal hDC As Long) As Long
    
    ' Selects an object into the specified DC. The new
    ' object replaces the previous object of the same type.
    ' Returned is the handle of the replaced object.
    Private Declare Function SelectObject Lib "GDI32" ( _
    ByVal hDC As Long, ByVal hObject As Long) As Long
    
    ' Performs a bit-block transfer of color data corresponding to
    ' a rectangle of pixels from the source DC into a destination DC.
    Private Declare Function BitBlt Lib "GDI32" ( _
    ByVal hDCDest As Long, ByVal XDest As Long, _
    ByVal YDest As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hDCSrc As Long, _
    ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
    As Long
    
    ' Retrieves the DC for the entire window, including title bar,
    ' menus, and scroll bars. A window DC permits painting anywhere
    ' in a window, because the origin of the DC is the upper-left
    ' corner of the window instead of the client area.
    Private Declare Function GetWindowDC Lib "USER32" ( _
    ByVal hwnd As Long) As Long
    
    ' Retrieves a handle to a display DC for the Client area of
    ' a specified window or for the entire screen.  You can use
    ' the returned handle in subsequent GDI functions to draw in
    ' the DC.
    Private Declare Function GetDC Lib "USER32" ( _
    ByVal hwnd As Long) As Long
    
    ' Releases a DC, freeing it for use by other applications.
    ' The effect of the ReleaseDC function depends on the type
    ' of DC.  It frees only common and window DCs.  It has no
    ' effect on class or private DCs.
    Private Declare Function ReleaseDC Lib "USER32" ( _
    ByVal hwnd As Long, ByVal hDC As Long) As Long
    
    ' Deletes the specified DC.
    Private Declare Function DeleteDC Lib "GDI32" ( _
    ByVal hDC As Long) As Long
    
    ' Retrieves the dimensions of the bounding rectangle of the
    ' specified window. The dimensions are given in screen
    ' coordinates that are relative to the upper-left corner
    ' of the screen.
    Private Declare Function GetWindowRect Lib "USER32" ( _
    ByVal hwnd As Long, lpRect As RECT) As Long
    
    ' Returns a handle to the Desktop window.  The desktop
    ' window covers the entire screen and is the area on top
    ' of which all icons and other windows are painted.
    Private Declare Function GetDesktopWindow Lib "USER32" () As Long
    
    ' Returns a handle to the foreground window (the window
    ' the user is currently working). The system assigns a
    ' slightly higher priority to the thread that creates the
    ' foreground window than it does to other threads.
    Private Declare Function GetForegroundWindow Lib "USER32" () As Long
    
    ' Creates a new picture object initialized according to a PICTDESC
    ' structure, which can be NULL, to create an uninitialized object if
    ' the caller wishes to have the picture initialize itself through
    ' IPersistStream::Load.  The fOwn parameter indicates whether the
    ' picture is to own the GDI picture handle for the picture it contains,
    ' so that the picture object will destroy its picture when the object
    ' itself is destroyed.  The function returns an interface pointer to the
    ' new picture object specified by the caller in the riid parameter.
    ' A QueryInterface is built into this call.  The caller is responsible
    ' for calling Release through the interface pointer returned - phew!
    Private Declare Function OleCreatePictureIndirect _
    Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
    ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    'Sub cmdForm_Click()
    ' Dim sPathAndFile As String
    ' Get the whole form inclusing borders, caption,...'
    '    Set Picture1.Picture = CaptureForm(frmMain)
    '    sPathAndFile = App.Path & "\" & "CapturedForm.jpg"
    '    SavePicture Picture1, sPathAndFile
    'End Sub
    Public Function CaptureForm(frm As Form) As Picture
        '
        ' Capture the entire form.
        '
        With frm
            Set CaptureForm = CaptureWindow(.hwnd, False, 0, 0, _
            .ScaleX(.Width, vbTwips, vbPixels), _
            .ScaleY(.Height, vbTwips, vbPixels))
        End With
    End Function
    
    Public Function CaptureWindow(ByVal hWndSrc As Long, _
        ByVal bClient As Boolean, ByVal LeftSrc As Long, _
        ByVal TopSrc As Long, ByVal WidthSrc As Long, _
        ByVal HeightSrc As Long) As Picture
        
        Dim hDCMemory       As Long
        Dim hBmp            As Long
        Dim hBmpPrev        As Long
        Dim r               As Long
        Dim hDCSrc          As Long
        Dim hPal            As Long
        Dim hPalPrev        As Long
        Dim RasterCapsScrn  As Long
        Dim HasPaletteScrn  As Long
        Dim PaletteSizeScrn As Long
        Dim LogPal          As LOGPALETTE
        '
        ' Get the proper Device Context (DC) depending on the value of bClient.
        '
        If bClient Then
            hDCSrc = GetDC(hWndSrc)       'Get DC for Client area.
        Else
            hDCSrc = GetWindowDC(hWndSrc) 'Get DC for entire window.
        End If
        '
        ' Create a memory DC for the copy process.
        '
        hDCMemory = CreateCompatibleDC(hDCSrc)
        '
        ' Create a bitmap and place it in the memory DC.
        '
        hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
        hBmpPrev = SelectObject(hDCMemory, hBmp)
        '
        ' Get the screen properties.
        '
        RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)   'Raster capabilities
        HasPaletteScrn = RasterCapsScrn And RC_PALETTE       'Palette support
        PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) 'Palette size
        '
        ' If the screen has a palette make a copy and realize it.
        '
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            '
            ' Create a copy of the system palette.
            '
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            '
            ' Select the new palette into the memory DC and realize it.
            '
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
        End If
        '
        ' Copy the on-screen image into the memory DC.
        '
        r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
        LeftSrc, TopSrc, vbSrcCopy)
        '
        ' Remove the new copy of the on-screen image.
        '
        hBmp = SelectObject(hDCMemory, hBmpPrev)
        '
        ' If the screen has a palette get back the
        ' palette that was selected in previously.
        '
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
        End If
        '
        ' Release the DC resources back to the system.
        '
        r = DeleteDC(hDCMemory)
        r = ReleaseDC(hWndSrc, hDCSrc)
        '
        ' Create a picture object from the bitmap
        ' and palette handles.
        '
        Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    End Function
    
    Public Function CreateBitmapPicture(ByVal hBmp As Long, _
        ByVal hPal As Long) As Picture
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '
        ' CreateBitmapPicture
        '    - Creates a bitmap type Picture object from a bitmap and palette.
        '
        ' hBmp
        '    - Handle to a bitmap
        '
        ' hPal
        '    - Handle to a Palette
        '    - Can be null if the bitmap doesn't use a palette
        '
        ' Returns
        '    - Returns a Picture object containing the bitmap
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '
        Dim r   As Long
        Dim Pic As PicBmp
        '
        ' IPicture requires a reference to "Standard OLE Types"
        '
        Dim IPic          As IPicture
        Dim IID_IDispatch As GUID
        '
        ' Fill in with IDispatch Interface ID
        '
        With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        '
        ' Fill Pic with the necessary parts.
        '
        With Pic
            .Size = Len(Pic)          ' Length of structure
            .Type = vbPicTypeBitmap   ' Type of Picture (bitmap)
            .hBmp = hBmp              ' Handle to bitmap
            .hPal = hPal              ' Handle to palette (may be null)
        End With
        '
        ' Create the Picture object.
        r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
        '
        ' Return the new Picture object.
        '
        Set CreateBitmapPicture = IPic
    End Function
    
    
    
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
    
    


