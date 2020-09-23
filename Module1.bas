Attribute VB_Name = "Module1"
'************************ColorBox Vertion 1.1************************
'Functions module; Color algorithms
'Author : Saifudheen A.A.
'E-mail : keraleeyan@msn.com.
'This ColorPicker was developed for my Paint programme
'and also this is not completed yet.
' Suggestions, Votes all are welcome.
'********************************************************************

Type POINTAPI
    x As Long
    y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type BITMAPINFOHEADER       ' 40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmicolors(15) As Long
End Type


Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long


Public SelectBox As RECT
Public MainBox As RECT
Public Preset() As RECT



Public SelectedPos As POINTAPI
Public SelectedMainPos As Single
Public cPaletteIndex As Integer

Public svdColor() As Long


Public Const BitsPixel = 12
Public Const Planes = 14
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
'Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Public Declare Function UpdateColors Lib "gdi32" (ByVal hDC As Long) As Long
'Public Declare Function PaletteIndex Lib "gdi32" (ByVal Index As Long) As Long
'Public Declare Function PaletteRGB Lib "gdi32" (ByVal Red As Integer, Green As Integer, Blue As Integer) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Public Const BDR_RAISEDINNER = &H4

Public Const BDR_RAISEDOUTER = &H1

Public Const BF_RIGHT = &H4

Public Const BF_TOP = &H2

Public Const BF_LEFT = &H1

Public Const BF_BOTTOM = &H8

Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long


Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

Public Const DIB_PAL_COLORS = 1 '  color table in palette indices

Dim lpbmINFO As BITMAPINFO
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long


Public Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Public Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry() As PALETTEENTRY
End Type

Public lpBI As BITMAPINFO
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long








Sub LoadVariantsHue(Red As Integer, Green As Integer, Blue As Integer)
  
    'On Error Resume Next
    Dim x As Integer, y As Integer
    Dim sDc As Long
    Dim K1 As Double, K2 As Double, K3 As Double
    Form1.DrawMode = 6
    Form1.Circle (SelectedPos.x, SelectedPos.y), 5  'Erases Previous Circle
    K1 = Red / 255
    K2 = Green / 255
    K3 = Blue / 255
    With Form1
        .DrawWidth = 1
        .DrawMode = 13
        
        Dim M1    As Double, M2     As Double, M3     As Double
        Dim J1    As Double, J2     As Double, J3     As Double
        Dim YMax As Byte
        Dim shdBitmap(0 To 196608) As Byte  '256 ^ 2 * 3
        Dim l As Long
        Dim bpos As Long
        Dim count As Long
        bpos = 0
        count = 0
        
        With lpBI.bmiHeader
            .biHeight = 256
            .biWidth = 256
        End With
        
        On Error Resume Next
        For y = 255 To 0 Step -1
                 M1 = Red - y * K1
                 M2 = Green - y * K2
                 M3 = Blue - y * K3
                 YMax = 255 - y
                 J1 = (YMax - M1) / 255
                 J2 = (YMax - M2) / 255
                 J3 = (YMax - M3) / 255
            For x = 255 To 0 Step -1
                shdBitmap(bpos) = M3 + x * J3    'Blue
                shdBitmap(bpos + 1) = M2 + x * J2    'Green
                shdBitmap(bpos + 2) = M1 + x * J1     'Red
                bpos = bpos + 3
            Next x
        Next y
        
        Dim bDc As Long
        Dim bm As Long
        Dim Ret As Long
        Dim hbmpOld  As Long
        bDc = CreateCompatibleDC(GetDC(0))
        bm = CreateDIBSection(bDc, lpBI, 0, ByVal 0&, ByVal 0&, ByVal 0&)
        Ret = SetDIBits(bDc, bm, 0, 256, shdBitmap(0), lpBI, 1)
        hbmpOld = SelectObject(bDc, bm)
        Ret = BitBlt(.hDC, SelectBox.Left, SelectBox.Top, 256, 256, bDc, 0, 0, vbSrcCopy)
        Ret = SelectObject(bDc, hbmpOld)
        Ret = DeleteDC(bDc)
        Ret = DeleteObject(bm)
    End With
    Form1.DrawSelFrame
    Form1.DrawMode = 6
    Form1.Circle (SelectedPos.x, SelectedPos.y), 5        'Refresh Circle
End Sub


Sub LoadMainBrightness(Red As Integer, Green As Integer, Blue As Integer)
    Dim r As Integer
    Dim g As Integer
    Dim B As Integer
    Dim OldP As POINTAPI
    Dim Cr As Long
    Dim sDc As Long
    sDc = Form1.hDC
    Form1.DrawMode = 13
    Form1.DrawWidth = 1
    For y = 0 To 255
        r = Red - Red * y / 255
        g = Green - Green * y / 255
        B = Blue - Blue * y / 255
        Form1.ForeColor = RGB(r, g, B)
        MoveToEx sDc, MainBox.Left, MainBox.Top + y, OldP
        LineTo sDc, MainBox.Right, MainBox.Top + y
    Next y
    Form1.DrawMainFrame
End Sub


Sub LoadVariantsBrightness()
Dim OldP As POINTAPI
Dim V As Integer
On Error Resume Next
Dim H, M As Single
Dim A As Integer, B As Integer, C As Integer, D As Integer, E As Integer, F As Integer
Dim sDc As Long
Dim Color As Long
Dim Red As Integer, Green As Integer, Blue As Integer
H = SelectBox.Bottom - SelectBox.Top
M = H / 6
A = M
B = 2 * M
C = 3 * M
D = 4 * M
E = 5 * M
F = 6 * M

Dim bBitmap(0 To 256 ^ 2 * 3) As Byte '256 ^ 2 * 3


Form1.DrawMode = 6
Form1.DrawWidth = 1
Form1.Circle (SelectedPos.x, SelectedPos.y), 5  'Erases Previous Circle
With Form1
    .DrawMode = 13
    sDc = .hDC
End With

With lpBI.bmiHeader
    .biHeight = 256
    .biWidth = 256
End With

    Dim Maa As Double, Mcc As Double, Mee As Double
    Dim MV As Double
    Dim Kc As Integer
    Dim YPos As Long
    Maa = 255 + 6 * A  'These are the common Terms taken out from the For Loop for efficciency
    Mcc = 255 + 6 * C  ' ""
    Mee = 255 + 6 * E  ' ""
    Dim pos As Long
    pos = 0
    
Dim x  As Integer, y As Integer
For y = 255 To 0 Step -1
        MV = 1 - y / 255 ' ""
    '1
        For x = 0 To A
            V = x * 6
            Kc = V * MV + y
            bBitmap(pos + 2) = 255
            bBitmap(pos + 1) = y
            bBitmap(pos + 0) = Kc
            pos = pos + 3
        Next x
    '2
        For x = A + 1 To B
            V = Maa - 6 * x ' 255 - (X - A) * 6
            Kc = V * MV + y
            bBitmap(pos + 2) = Kc
            bBitmap(pos + 1) = y
            bBitmap(pos + 0) = 255
            pos = pos + 3
        Next x
     '3
        For x = B + 1 To C
            V = (x - B - 1) * 6
            Kc = V * MV + y
            bBitmap(pos + 2) = y
            bBitmap(pos + 1) = Kc
            bBitmap(pos + 0) = 255
            pos = pos + 3
        Next x
     '4
        For x = C + 1 To D
            V = Mcc - 6 * x
            Kc = V * MV + y
            bBitmap(pos + 2) = y
            bBitmap(pos + 1) = 255
            bBitmap(pos + 0) = Kc
            pos = pos + 3
        Next x
    '5
        For x = D + 1 To E
            V = (x - D - 1) * 6
            Kc = V * MV + y
            bBitmap(pos + 2) = Kc
            bBitmap(pos + 1) = 255
            bBitmap(pos + 0) = y
            pos = pos + 3
        Next x
    '6
        For x = E + 1 To F
            V = Mee - 6 * x
            Kc = V * MV + y
            bBitmap(pos + 2) = 255
            bBitmap(pos + 1) = Kc
            bBitmap(pos + 0) = y
            pos = pos + 3
        Next x
       
Next y

        Dim bDc As Long
        Dim bm As Long
        Dim Ret As Long
        Dim hbmpOld  As Long
        bDc = CreateCompatibleDC(0)
        bm = CreateDIBSection(bDc, lpBI, 0, ByVal 0&, ByVal 0&, ByVal 0&)
        Ret = SetDIBits(bDc, bm, 0, 256, bBitmap(0), lpBI, 1)
        hbmpOld = SelectObject(bDc, bm)
        Ret = StretchBlt(Form1.hDC, SelectBox.Right, SelectBox.Top, -256, 256, bDc, 0, 0, 256, 256, vbSrcCopy)
        'Ret = BitBlt(Form1.hDC, SelectBox.Left, SelectBox.Top, 256, 256, bDc, 0, 0, vbSrcCopy)
        Ret = SelectObject(bDc, hbmpOld)
        Ret = DeleteDC(bDc)
        Ret = DeleteObject(bm)
    Form1.DrawSelFrame
    Form1.DrawMode = 6
    Form1.Circle (SelectedPos.x, SelectedPos.y), 5 'Refresh Circle
    
End Sub


Sub LoadVariantsSaturation()
Dim OldP As POINTAPI
Dim V As Integer
On Error Resume Next
Dim H, M As Single
Dim x As Integer, y As Integer
Dim A As Integer, B As Integer, C As Integer, D As Integer, E As Integer, F As Integer
Dim sDc As Long
Dim Color As Long
Dim Red As Integer, Green As Integer, Blue As Integer
Dim bBitmap(0 To 256 ^ 2 * 3) As Byte '256 ^ 2 * 3
Dim cpos As Long
cpos = 0
H = SelectBox.Bottom - SelectBox.Top
M = H / 6
A = M
B = 2 * M
C = 3 * M
D = 4 * M
E = 5 * M
F = 6 * M

Form1.DrawMode = 6
Form1.Circle (SelectedPos.x, SelectedPos.y), 5  'Erases Previous Circle

With Form1
    .DrawWidth = 1
    .DrawMode = 13
    sDc = .hDC
End With
    Dim Maa As Double, Mcc As Double, Mee As Double
    Dim MV As Double
    Dim Kc As Integer
    Dim YPos As Long
    Maa = 255 + 6 * A  'These are the common Terms taken out from the For Loop for efficiency
    Mcc = 255 + 6 * C  ' ""
    Mee = 255 + 6 * E  ' ""
    
For y = 255 To 0 Step -1
        MV = 1 - y / 255 ' ""
        YPos = SelectBox.Top + y
    '1
        For x = 0 To A
            V = x * 6
            Kc = V * MV
            bBitmap(pos + 2) = 255 - y
            bBitmap(pos + 1) = 0
            bBitmap(pos + 0) = Kc
            pos = pos + 3
        Next x
    '2
        For x = A + 1 To B
            V = Maa - 6 * x
            Kc = V * MV
            bBitmap(pos + 2) = Kc
            bBitmap(pos + 1) = 0
            bBitmap(pos + 0) = 255 - y
            pos = pos + 3
        Next x
     '3
        For x = B + 1 To C
            V = (x - B - 1) * 6
            Kc = V * MV
            bBitmap(pos + 2) = 0
            bBitmap(pos + 1) = Kc
            bBitmap(pos + 0) = 255 - y
            pos = pos + 3
        Next x
     '4
        For x = C + 1 To D
            V = Mcc - 6 * x
            Kc = V * MV
            bBitmap(pos + 2) = 0
            bBitmap(pos + 1) = 255 - y
            bBitmap(pos + 0) = Kc
            pos = pos + 3
        Next x
    '5
        For x = D + 1 To E
            V = (x - D - 1) * 6
            Kc = V * MV
            bBitmap(pos + 2) = Kc
            bBitmap(pos + 1) = 255 - y
            bBitmap(pos + 0) = 0
            pos = pos + 3
        Next x
    '6
        For x = E + 1 To F
            V = Mee - 6 * x
            Kc = V * MV
            bBitmap(pos + 2) = 255 - y
            bBitmap(pos + 1) = Kc
            bBitmap(pos + 0) = 0
            pos = pos + 3
        Next x
       
Next y

        Dim bDc As Long
        Dim bm As Long
        Dim Ret As Long
        Dim hbmpOld  As Long
        bDc = CreateCompatibleDC(0)
        bm = CreateDIBSection(bDc, lpBI, 0, ByVal 0&, ByVal 0&, ByVal 0&)
        Ret = SetDIBits(bDc, bm, 0, 256, bBitmap(0), lpBI, 1)
        hbmpOld = SelectObject(bDc, bm)
        Ret = StretchBlt(Form1.hDC, SelectBox.Right, SelectBox.Top, -256, 256, bDc, 0, 0, 256, 256, vbSrcCopy)
        'Ret = BitBlt(Form1.hDC, SelectBox.Left, SelectBox.Top, 256, 256, bDc, 0, 0, vbSrcCopy)
        Ret = SelectObject(bDc, hbmpOld)
        Ret = DeleteDC(bDc)
        Ret = DeleteObject(bm)
    Form1.DrawSelFrame
    Form1.DrawMode = 6
    Form1.Circle (SelectedPos.x, SelectedPos.y), 5 'Refresh Circle

End Sub
Sub LoadMainSaturation(Red As Integer, Green As Integer, Blue As Integer)
    Dim r As Integer
    Dim g As Integer
    Dim B As Integer
    Dim OldP As POINTAPI
    Dim sDc As Long
    sDc = Form1.hDC
    Form1.DrawMode = 13
    Form1.DrawWidth = 1
    For y = 0 To 255
        r = Red * (1 - y / 255) + y
        g = Green * (1 - y / 255) + y
        B = Blue * (1 - y / 255) + y
        Form1.ForeColor = RGB(r, g, B)
        MoveToEx sDc, MainBox.Left, MainBox.Top + y, OldP
        LineTo sDc, MainBox.Right, MainBox.Top + y
    Next y
    Form1.DrawMainFrame
End Sub



Sub GetRGB(ByRef cl As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
    Dim C As Long
    C = cl
    Red = C Mod &H100
    C = C \ &H100
    Green = C Mod &H100
    C = C \ &H100
    Blue = C Mod &H100
End Sub

Sub DrawSlider(ByVal position As Integer)
    Form1.DrawMode = 6
    'Form1.DrawWidth = 5
    'Form1.Line (MainBox.Right + 7, position)-(MainBox.Right + 8, position)
    Form1.DrawWidth = 2
    Form1.Line (MainBox.Right + 2, position)-(MainBox.Right + 5, position)
    Form1.Line (MainBox.Left - 2, position)-(MainBox.Left - 5, position)
    Form1.DrawWidth = 1
End Sub

Sub LoadSafePalette()
Form1.FillStyle = 0
Form1.DrawMode = 13
Form1.DrawWidth = 1
On Error Resume Next
Dim i, j, k As Integer
Dim l As Long
Dim count As Integer
Dim Plt As Long
Dim Ret As Long
Dim br As Long

For i = 0 To &HFF Step &H33
    For j = 0 To &HFF Step &H33
        For k = 0 To &HFF Step &H33
            count = count + 1
            l = RGB(i, j, k)
            Form1.FillColor = l
            Form1.ForeColor = vbBlack
            Rectangle Form1.hDC, Preset(count).Left, Preset(count).Top, Preset(count).Right, Preset(count).Bottom
        Next k
    Next j
Next i
For i = 217 To 224
    Form1.FillColor = 0
    Rectangle Form1.hDC, Preset(i).Left, Preset(i).Top, Preset(i).Right, Preset(i).Bottom
Next i
Form1.DrawSafePicker cPaletteIndex, False

End Sub

Sub LoadCustomColors()
    Dim FileHandle As Integer
    Dim i As Integer
    Dim strColor As String
    On Error Resume Next
    FileHandle = FreeFile()
    ReDim svdColor(0 To 224)
    Open App.Path & "/usercolors.cps" For Input As #FileHandle
    i = 0
    Form1.Cls
    Form1.FillStyle = 0
    Form1.DrawMode = 13
    Form1.DrawWidth = 1
    For i = 0 To 224
        Line Input #FileHandle, strColor
        svdColor(i) = Val(strColor)
        Form1.ForeColor = vbBlack 'svdColor(i)
        Form1.FillColor = svdColor(i)
        Rectangle Form1.hDC, Preset(i).Left, Preset(i).Top, Preset(i).Right, Preset(i).Bottom
    Next i
    Close #FileHandle
    Form1.DrawSafePicker cPaletteIndex, False
    Form1.Refresh
End Sub

Sub SaveCustomColors()
    Dim FileHandle As Integer
    Dim i As Integer
    On Error Resume Next
    FileHandle = FreeFile()
    Open App.Path & "/usercolors.cps" For Output As #FileHandle
    For i = 0 To 224
        Print #FileHandle, svdColor(i)
    Next i
    Close #FileHandle
  
End Sub
