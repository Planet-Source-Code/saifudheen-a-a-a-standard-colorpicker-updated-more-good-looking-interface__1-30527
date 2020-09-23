VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Box"
   ClientHeight    =   4155
   ClientLeft      =   3960
   ClientTop       =   3570
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "O&K"
      Height          =   405
      Left            =   5550
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   2640
      Width           =   825
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4530
      TabIndex        =   26
      Top             =   3060
      Width           =   975
      Begin VB.Label lblADDColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Preset"
      Height          =   405
      Left            =   5550
      TabIndex        =   7
      Top             =   3120
      Width           =   825
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   " << Add    "
      Height          =   405
      Left            =   4560
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbPreset 
      Height          =   315
      Left            =   4650
      TabIndex        =   24
      Text            =   "Safe Colors"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   4620
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   30
      Width           =   1695
      Begin VB.Label Label3 
         Caption         =   "Old"
         Height          =   195
         Left            =   1050
         TabIndex        =   6
         Top             =   150
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   855
         TabIndex        =   5
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "New"
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   150
         Width           =   465
      End
      Begin VB.Label lblSelColor 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   4590
      TabIndex        =   8
      Top             =   960
      Width           =   2055
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1380
         TabIndex        =   17
         Text            =   "0"
         ToolTipText     =   "Blue"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1380
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "Green"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1380
         TabIndex        =   15
         Text            =   "255"
         ToolTipText     =   "Red"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Left            =   450
         MousePointer    =   4  'Icon
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "Hue"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtS 
         Height          =   285
         Left            =   450
         MousePointer    =   4  'Icon
         TabIndex        =   13
         Text            =   "100"
         ToolTipText     =   "Saturation"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtB 
         Height          =   285
         Left            =   450
         MousePointer    =   4  'Icon
         TabIndex        =   12
         Text            =   "100"
         ToolTipText     =   "Brightness"
         Top             =   1080
         Width           =   375
      End
      Begin VB.OptionButton optH 
         Caption         =   "H:"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   "Hue"
         Top             =   510
         Value           =   -1  'True
         Width           =   465
      End
      Begin VB.OptionButton optS 
         Caption         =   "S:"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "Saturation"
         Top             =   810
         Width           =   465
      End
      Begin VB.OptionButton optB 
         Caption         =   "B:"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "Brightness"
         Top             =   1110
         Width           =   495
      End
      Begin VB.Label lblR 
         Caption         =   "R:"
         Height          =   225
         Left            =   1200
         TabIndex        =   23
         ToolTipText     =   "Red"
         Top             =   510
         Width           =   195
      End
      Begin VB.Label lblG 
         Caption         =   "G:"
         Height          =   225
         Left            =   1200
         TabIndex        =   22
         ToolTipText     =   "Green"
         Top             =   810
         Width           =   225
      End
      Begin VB.Label lblB 
         Caption         =   "B:"
         Height          =   225
         Left            =   1200
         TabIndex        =   21
         ToolTipText     =   "Blue"
         Top             =   1110
         Width           =   195
      End
      Begin VB.Label lblS 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   20
         Top             =   810
         Width           =   165
      End
      Begin VB.Label lblBB 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   1110
         Width           =   195
      End
      Begin VB.Label lblH 
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   840
         TabIndex        =   18
         Top             =   480
         Width           =   105
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   5550
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   3600
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'************************ColorBox Vertion 1.1************************
'Main dialog
'Author : Saifudheen A.A.
'E-mail : keraleeyan@msn.com.
'This ColorPicker was developed for my Paint programme
'and also this is not completed yet.
' Suggestions, Votes all are welcome.
'********************************************************************


Dim MainBoxHit  As Boolean
Dim SelectBoxHit As Boolean

Enum sMode
    Picker = 0
    Custom = 1
    InetExplorer = 2
    NetsCape = 3
    SafeColor = 4
End Enum
Dim Mode As sMode


Private Sub LoadHueShades()
Me.Show
Dim OldP As POINTAPI
Dim V As Integer
On Error Resume Next
Dim H As Single, M As Single
Dim A As Single, B As Single, C As Single, D As Single, E As Single, F As Single
Dim Ratio As Single
Dim sDc As Long
H = SelectBox.Bottom - SelectBox.Top
M = H / 6
A = M
B = 2 * M
C = 3 * M
D = 4 * M
E = 5 * M
F = 6 * M
Dim sBitmap(0 To 16 * 256 * 3) As Byte            '256 ^ 2 * 3
Dim cpos  As Long
With lpBI.bmiHeader
    .biHeight = 256
    .biWidth = 15
End With

cpos = 0
With Me '1
    .DrawWidth = 1
    .DrawMode = 13
    sDc = .hDC
    For y = 0 To Int(A)
        For j = 1 To 16
            sBitmap(cpos + 2) = 255
            sBitmap(cpos + 1) = 0
            sBitmap(cpos + 0) = y * 6
            cpos = cpos + 3
        Next j
    Next y
'2
            
    For y = Int(A) + 1 To Int(B)
        V = 255 - (y - A) * 6
        For j = 1 To 16
            sBitmap(cpos + 2) = V
            sBitmap(cpos + 1) = 0
            sBitmap(cpos + 0) = 255
            cpos = cpos + 3
        Next j
        
    Next y
 '3
     
    For y = Int(B) + 1 To Int(C)
        V = (y - B) * 6
        For j = 1 To 16
            sBitmap(cpos + 2) = 0
            sBitmap(cpos + 1) = V
            sBitmap(cpos + 0) = 255
            cpos = cpos + 3
        Next j
        
    Next y
 '4
    For y = Int(C) + 1 To Int(D)
        V = 255 - (y - C) * 6
        For j = 1 To 16
            sBitmap(cpos + 2) = 0
            sBitmap(cpos + 1) = 255
            sBitmap(cpos + 0) = V
            cpos = cpos + 3
        Next j
    Next y
'5
    For y = Int(D) + 1 To Int(E)
        V = (y - D) * 6
        For j = 1 To 16
            sBitmap(cpos + 2) = V
            sBitmap(cpos + 1) = 255
            sBitmap(cpos + 0) = 0
            cpos = cpos + 3
        Next j
        
    Next y
'6
    For y = Int(E) + 1 To Int(F)
        V = 255 - (y - E) * 6
        For j = 1 To 16
            sBitmap(cpos + 2) = 255
            sBitmap(cpos + 1) = V
            sBitmap(cpos + 0) = 0
            cpos = cpos + 3
        Next j
    Next y
End With
        Dim bDc As Long
        Dim bm As Long
        Dim Ret As Long
        Dim hbmpOld  As Long
        
        bDc = CreateCompatibleDC(0)
        bm = CreateDIBSection(bDc, lpBI, 0, ByVal 0&, ByVal 0&, ByVal 0&)
        Ret = SetDIBits(bDc, bm, 0, 256, sBitmap(0), lpBI, 0)
        hbmpOld = SelectObject(bDc, bm)
        Ret = StretchBlt(Form1.hDC, MainBox.Left, MainBox.Bottom, 15, -256, bDc, 0, 0, 15, 256, vbSrcCopy)
        'Ret = BitBlt(Form1.hDC, MainBox.Left, MainBox.Top, 15, 256, bDc, 0, 0, vbSrcCopy)
        Ret = SelectObject(bDc, hbmpOld)
        Ret = DeleteDC(bDc)
        Ret = DeleteObject(bm)
        DrawMainFrame

End Sub


Private Sub cmbColorType_Click()
    If cmbColorType.ListIndex = 1 Then
        Text4.Visible = True
        lblK.Visible = True
    Else
        Text4.Visible = False
        lblK.Visible = False
    End If
    Select Case cmbColorType.ListIndex
    Case 0
        lblR.Caption = "R": lblG.Caption = "G": lblB.Caption = "B"
    Case 1
        lblR.Caption = "C": lblG.Caption = "M": lblB.Caption = "Y"
    Case 2
        lblR.Caption = "C": lblG.Caption = "M": lblB.Caption = "Y"
    Case 3
        lblR.Caption = "H": lblG.Caption = "S": lblB.Caption = "B"
    End Select
    
End Sub

Private Sub cmbPreset_Click()
Me.Cls
Select Case cmbPreset.ListIndex
Case 0
    LoadCustomColors
    cmdADD.Visible = True
    lblADDColor.Visible = True
    Mode = Custom
Case 1
    cmdADD.Visible = False
    lblADDColor.Visible = False
    Mode = InetExplorer
Case 2
    cmdADD.Visible = False
    lblADDColor.Visible = False
    Mode = NetsCape
Case 3
    LoadSafePalette
    cmdADD.Visible = False
    lblADDColor.Visible = False
    Mode = SafeColor
End Select
End Sub

Private Sub cmdADD_Click()
    svdColor(cPaletteIndex) = lblADDColor.BackColor
    DrawSafePicker cPaletteIndex, True 'Erase
    Form1.DrawWidth = 1
    Form1.DrawMode = 13
    Form1.FillStyle = 0
    Form1.ForeColor = 0
    Form1.FillColor = svdColor(cPaletteIndex)
    
    Rectangle Form1.hDC, Preset(cPaletteIndex).Left, Preset(cPaletteIndex).Top, Preset(cPaletteIndex).Right, Preset(cPaletteIndex).Bottom
    SaveCustomColors
    DrawSafePicker cPaletteIndex, False
    Form1.Refresh
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Sub LoadColorDialog(ByVal oldColor As Long)
    lblSelColor.BackColor = oldColor
End Sub



Private Sub GetColorMode()
    Dim NumColors As Long
    Dim hDC As Long
    Dim x As Long
    Dim PL As Long
    Dim BP As Long
    hDC = CreateDC("DISPLAY", "", "", "")
    PL = GetDeviceCaps(hDC, Planes)
    BP = GetDeviceCaps(hDC, BitsPixel)
    NumColors = 2 ^ CLng(PL * BP)
    x = DeleteDC(hDC)
    Text1.Text = Str$(NumColors)

End Sub

Private Sub Command3_Click()
Me.Cls
If Mode = Picker Then
    Frame2.Visible = False
    Mode = Custom
    LoadCustomColors
    Command3.Caption = "&Picker"
    cmbPreset.Visible = True
    cmbPreset.ListIndex = 0
    lblADDColor.BackColor = lblSelColor.BackColor
    cmdADD.Visible = True
    lblADDColor.Visible = True

Else
    Me.Cls
    cmbPreset.Visible = False
    Command3.Caption = "&Preset"
    DrawPicker
    Select Case True
    Case optH
         optH_Click
    Case optS
        optS_Click
    Case optB
        optB_Click
    End Select
   
    DrawSlider SelectedMainPos
    Mode = Picker
    Frame2.Visible = True
    cmdADD.Visible = False
    lblADDColor.Visible = False
End If


End Sub



Private Function GetSafeColor(Index As Integer, r As Integer, g As Integer, B As Integer, HexVal As String) As Long
Dim i As Long, j As Long, k As Long
Dim count As Integer
Dim strR As String, strG As String, strB As String

For i = 0 To &HFF Step &H33
    For j = 0 To &HFF Step &H33
        For k = 0 To &HFF Step &H33
            count = count + 1
            If count = Index Then
                r = i: g = j: B = k
                GetSafeColor = RGB(i, j, k)
                GetHexVal r, g, B, HexVal
                Exit Function
            End If
           
        Next k
    Next j
Next i

End Function

Sub GetHexVal(Red As Integer, Green As Integer, Blue As Integer, strHex As String)
    Dim strR As String, strG As String, strB As String
    strR = Trim(Hex(Red))
        If Len(strR) = 1 Then strR = "0" & strR
    strG = Trim(Hex(Green))
        If Len(strG) = 1 Then strR = "0" & strR
    strB = Trim(Hex(Blue))
        If Len(strB) = 1 Then strR = "0" & strR
        strHex = strR & strG & strB

End Sub






 
Private Sub Form_Load()
    
    ' Set these Parameters on Basis of where  should Select Box and  Main Box Should appear
    Dim OldP As POINTAPI
    ReDim Preset(1 To 224)  'RECT structure
    Preset(1).Left = 10
    Preset(1).Top = 10
    Preset(1).Right = 25
    Preset(1).Bottom = 25
    Mode = Picker 'Initialize Mode as  normal Picker
    With lpBI.bmiHeader
        .biBitCount = 24
        .biCompression = 0&
        .biPlanes = 1
        .biSize = Len(lpBI.bmiHeader)
    End With

    
    '// Setting the position of RECTS for Safe Colors
    For i = 2 To 224
        If i Mod 16 = 0 Then
            Preset(i).Top = Preset(i - 1).Top
            Preset(i).Left = Preset(i - 1).Right + 3
            Preset(i).Bottom = Preset(i).Top + 15
            Preset(i).Right = Preset(i).Left + 15
            If i = 224 Then GoTo Jump
            i = i + 1
            Preset(i).Top = Preset(i - 1).Bottom + 3
            Preset(i).Left = Preset(1).Left
        Else
            Preset(i).Top = Preset(i - 1).Top
            Preset(i).Left = Preset(i - 1).Right + 3
        End If
Jump:
        Preset(i).Bottom = Preset(i).Top + 15
        Preset(i).Right = Preset(i).Left + 15
    Next i

    
    k = 255
    SelectBox.Left = 10
    SelectBox.Top = 10
    SelectBox.Right = SelectBox.Left + k
    SelectBox.Bottom = SelectBox.Top + k
    MainBox.Left = SelectBox.Right + 12
    MainBox.Top = SelectBox.Top
    MainBox.Right = MainBox.Left + 15
    MainBox.Bottom = SelectBox.Bottom
     
    With cmbPreset
    .AddItem "Custom..."
    .AddItem "Microsoft® Internet Explorer"
    .AddItem "Netscape Navigator (TM)"
    .AddItem "Safe Palette (216)"
    End With
     
     
    LoadHueShades
    SelectedMainPos = MainBox.Bottom
    SelectedPos.x = SelectBox.Right
    SelectedPos.y = SelectBox.Top
    Call DrawSlider(SelectedMainPos)
    DrawPicker
    LoadVariantsHue 255, 0, 0
    cPaletteIndex = 1
    Me.ForeColor = vbBlack

    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If Mode = Picker Then
            If x >= SelectBox.Left And x <= SelectBox.Right And y >= SelectBox.Top And y <= SelectBox.Bottom Then
                '// In SelectBox Boundary
                SelectBoxHit = True
                'Me.MousePointer = vbCustom
               Call MouseOnSelectBox(x, y)
            End If
            If x >= MainBox.Left And x < MainBox.Right + 11 And y >= MainBox.Top - 2 And y < MainBox.Bottom + 3 Then
                '// In MainBox Boundary
                MainBoxHit = True
                If y > MainBox.Bottom Then y = MainBox.Bottom
                If y < MainBox.Top Then y = MainBox.Top
                Call MouseOnMainBox(x, y)
            End If
        Else
            HandlePresetValues x, y, Mode
        End If
        
        
    End If
End Sub

Sub DrawSafePicker(Index As Integer, Clear As Boolean)
    Dim r As RECT
    Dim l As Long
    Me.FillStyle = 1
    Me.DrawMode = 13
    Me.DrawWidth = 3
    If Clear Then
        Me.ForeColor = Me.BackColor
        Rectangle Me.hDC, Preset(Index).Left - 2, Preset(Index).Top - 2, Preset(Index).Right + 2, Preset(Index).Bottom + 2
    Else
        r.Left = Preset(Index).Left - 3
        r.Top = Preset(Index).Top - 3
        r.Right = Preset(Index).Right + 3
        r.Bottom = Preset(Index).Bottom + 3
        Call DrawEdge(Form1.hDC, r, BDR_SUNKENINNER Or BDR_SUNKENOUTER, BF_RECT Or BF_SOFT)
    End If
    
End Sub

Private Sub MouseOnSelectBox(x As Single, y As Single)
            Dim cl As Long
            Dim r As Integer
            Dim g As Integer
            Dim B As Integer
            DrawPicker
            SelectedPos.x = x
            SelectedPos.y = y
            DrawPicker
            cl = Me.Point(x, y)
            GetRGB cl, r, g, B
            If optS.Value Then
                LoadMainSaturation r, g, B
                txtB.Text = Int((255 - SelectedPos.y + SelectBox.Top) * 100 / 255) 'Brightness Level
                txtH.Text = Int((SelectedPos.x - SelectBox.Left) * 360 / 255) 'Hue Level
            End If
            
            If optB.Value Then
                LoadMainBrightness r, g, B
                txtS.Text = Int((SelectBox.Bottom - SelectedPos.y) * 100 / 255)    ' Saturation Level
                txtH.Text = Int((SelectedPos.x - SelectBox.Left) * 360 / 255)   'Hue Level
            End If

            If optH.Value Then
                txtS.Text = Int((SelectedPos.x - SelectBox.Left) * 100 / 255)   ' Saturation Level
                txtB.Text = Int((255 - SelectedPos.y + SelectBox.Top) * 100 / 255) 'Brightness Level
            Else
                cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
                GetRGB cl, r, g, B
            End If
            Text1.Text = r
            Text2.Text = g
            Text3.Text = B
            lblSelColor.BackColor = cl
            'Me.Refresh

End Sub

Private Sub MouseOnMainBox(x As Single, y As Single)
        Dim cl As Long
        Dim r As Integer
        Dim g As Integer
        Dim B As Integer

            DrawSlider SelectedMainPos
            DrawSlider y
            SelectedMainPos = y
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            GetRGB cl, r, g, B
            
            If optH.Value Then
                txtH.Text = Int((255 - y + SelectBox.Top) * 360 / 255)
                GetColorFromHSB Val(txtS.Text), Val(txtB.Text)
            Else
                Text1.Text = r
                Text2.Text = g
                Text3.Text = B
                lblSelColor.BackColor = cl
            End If
            If optS.Value Then
                txtS.Text = Int((255 - y + SelectBox.Top) * 100 / 255)
                'GetColorFromHSB Val(txtS.Text), Val(txtB.Text)
            End If
            If optB.Value Then
                txtB.Text = Int((255 - y + SelectBox.Top) * 100 / 255)
                'GetColorFromHSB Val(txtS.Text), Val(txtB.Text)
            End If

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If Mode = Picker Then
            If SelectBoxHit Then
                If x < SelectBox.Left Then x = SelectBox.Left
                If x > SelectBox.Right Then x = SelectBox.Right
                If y < SelectBox.Top Then y = SelectBox.Top
                If y > SelectBox.Bottom Then y = SelectBox.Bottom
                '// In SelectBox Region
                Call MouseOnSelectBox(x, y)
            End If
            
            If MainBoxHit Then
                '// In MainBox region
                x = MainBox.Left + 2
                If y > MainBox.Bottom Then y = MainBox.Bottom
                If y < MainBox.Top Then y = MainBox.Top
                Call MouseOnMainBox(x, y)
            End If
        Else
            HandlePresetValues x, y, Mode
        End If
    End If

End Sub

Private Sub HandlePresetValues(x As Single, y As Single, pMode As sMode)
            Dim i As Integer
            Dim r As Integer, g As Integer, B As Integer
            Dim HexV As String
            For i = 1 To 224
                If x > Preset(i).Left And x < Preset(i).Right And y > Preset(i).Top And y < Preset(i).Bottom Then
                    DrawSafePicker cPaletteIndex, True
                    DrawSafePicker i, False
                    Me.Refresh
                    cPaletteIndex = i
                    Select Case pMode
                    Case 1
                        lblSelColor.BackColor = svdColor(i)
                        GetRGB svdColor(i), r, g, B
                        GetHexVal r, g, B, HexV
                    Case 2
                    Case 3
                    Case 4
                        lblSelColor.BackColor = GetSafeColor(i, r, g, B, HexV)
                    End Select
                    
                    PrintRGBHEX r, g, B, HexV
                    Exit Sub
                End If
            Next i

End Sub

Private Sub PrintRGBHEX(r As Integer, g As Integer, B As Integer, HexV As String)
    Me.FillColor = Me.BackColor
    Me.ForeColor = Me.BackColor
    Me.DrawMode = 13
    Me.Line (MainBox.Right + 20, 80)-(MainBox.Right + 20 + 80, 80 + 50), Me.BackColor, BF
    Me.CurrentX = MainBox.Right + 20
    Me.CurrentY = 80
    Me.ForeColor = 0
    Me.Print "R: " & r
    Me.CurrentX = MainBox.Right + 20
    Me.Print "G: " & g
    Me.CurrentX = MainBox.Right + 20
    Me.Print "B: " & B
    Me.CurrentX = MainBox.Right + 20
    Me.Print "Hex: #" & HexV

End Sub
Private Sub GetColorFromHSB(ByVal Sat As Integer, ByVal br As Integer)
    '//
    '// This Function Evaluates the Resulting Color Value While Sliding the Hue Shades From Insisting Brightness and Saturation Values
    Dim r As Integer, g As Integer, B As Integer
    Dim Red As Integer, Green As Integer, Blue As Integer
    Dim cl As Long
    Dim x As Long, y As Long
    x = Sat * 255 / 100
    cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
    GetRGB cl, Red, Green, Blue
    r = (Red + ((255 - Red) * (100 - Sat)) / 100) * br / 100
    g = (Green + ((255 - Green) * (100 - Sat)) / 100) * br / 100
    B = (Blue + ((255 - Blue) * (100 - Sat)) / 100) * br / 100
    lblSelColor.BackColor = RGB(r, g, B)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    SelectBoxHit = False
    If MainBoxHit And optH.Value Then
        Dim r As Integer, g As Integer, B As Integer
        Dim cl As Long
        cl = GetPixel(Me.hDC, MainBox.Left + 3, SelectedMainPos)
        GetRGB cl, r, g, B
        LoadVariantsHue r, g, B
        cl = Me.Point(SelectedPos.x, SelectedPos.y)
        GetRGB cl, r, g, B
        Text1.Text = r
        Text2.Text = g
        Text3.Text = B
        lblSelColor.BackColor = Me.Point(SelectedPos.x, SelectedPos.y)
        
    End If
    MainBoxHit = False
End Sub





Private Sub Label2_Click()
    Dim r As Integer, g As Integer, B As Integer
    Dim HexV As String
    GetRGB Label2.BackColor, r, g, B
    GetHexVal r, g, B, HexV
    PrintRGBHEX r, g, B, HexV

End Sub

Private Sub lblADDColor_Click()
    Dim r As Integer, g As Integer, B As Integer
    Dim HexV As String
    GetRGB lblADDColor.BackColor, r, g, B
    GetHexVal r, g, B, HexV
    PrintRGBHEX r, g, B, HexV
End Sub

Private Sub lblSelColor_Click()
    Dim r As Integer, g As Integer, B As Integer
    Dim HexV As String
    GetRGB lblSelColor.BackColor, r, g, B
    GetHexVal r, g, B, HexV
    PrintRGBHEX r, g, B, HexV

End Sub

Private Sub optB_Click()
    Dim r As Integer, g As Integer, B As Integer
    Dim C As Long
    LoadVariantsBrightness  '// Loading the Shades at SelectBox
    DrawPicker 'Erase picker
    SelectedPos.x = SelectBox.Left + Val(txtH.Text) * 255 / 360  'Depending on Hue
    SelectedPos.y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100 'Depending on Saturation
    DrawPicker
    C = Me.Point(SelectedPos.x, SelectedPos.y)
    GetRGB C, r, g, B
    LoadMainBrightness r, g, B

End Sub

Private Sub optH_Click()
    Dim cl As Long
    Dim r As Integer, g As Integer, B As Integer
    LoadHueShades
    DrawSlider SelectedMainPos
    SelectedMainPos = Int(MainBox.Bottom - Val(txtH.Text) * 255 / 360)
    DrawSlider SelectedMainPos
    cl = Me.Point(MainBox.Left + 3, SelectedMainPos)
    GetRGB cl, r, g, B
    LoadVariantsHue r, g, B
    DrawPicker
    SelectedPos.x = SelectBox.Left + Val(txtS.Text) * 255 / 100
    SelectedPos.y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100
    DrawPicker
    cl = Me.Point(SelectedPos.x, SelectedPos.y)
    lblSelColor.BackColor = cl
End Sub

Private Sub optS_Click()
    Dim r As Integer, g As Integer, B As Integer
    Dim C As Long
    LoadVariantsSaturation  '// Loading the Shades at SelectBox
    DrawPicker 'Erase picker
    SelectedPos.x = SelectBox.Left + Val(txtH.Text) * 255 / 360  'Depending on Hue
    SelectedPos.y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100 'Depending on Brightness
    DrawPicker
    C = Me.Point(SelectedPos.x, SelectedPos.y)
    GetRGB C, r, g, B
    LoadMainSaturation r, g, B
    Me.Refresh
End Sub

Private Sub txtB_KeyUp(KeyCode As Integer, Shift As Integer)
    If optH.Value Then
        DrawPicker  'Erase Previous picker
        SelectedPos.y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100
        DrawPicker
        UpdateColorValues
    End If

End Sub
Sub DrawPicker()
    Me.DrawMode = 6
    Me.FillStyle = 1
    Me.DrawWidth = 1
    Me.DrawStyle = 0
    Me.Circle (SelectedPos.x, SelectedPos.y), 5
End Sub
Sub DrawSelFrame()
    Dim SelFrame As RECT
    SelFrame.Left = SelectBox.Left - 1
    SelFrame.Top = SelectBox.Top - 1
    SelFrame.Right = SelectBox.Right + 3
    SelFrame.Bottom = SelectBox.Bottom + 3
    DrawEdge Me.hDC, SelFrame, BDR_SUNKENINNER, BF_RECT
    Me.Refresh
End Sub

Sub DrawMainFrame()
    Dim MainFrame As RECT
     
    MainFrame.Left = MainBox.Left - 1
    MainFrame.Top = MainBox.Top - 1
    MainFrame.Right = MainBox.Right + 1
    MainFrame.Bottom = MainBox.Bottom + 3
    DrawEdge Me.hDC, MainFrame, BDR_SUNKENINNER, BF_RECT
    Me.Refresh
End Sub



Private Sub txtH_Change() 'On Error Resume Next

End Sub
Private Sub UpdateColorValues()
        Dim r As Integer, g As Integer, B As Integer
        Dim cl As Long
        cl = Me.Point(SelectedPos.x, SelectedPos.y)
        lblSelColor.BackColor = cl
        GetRGB cl, r, g, B
        Text1.Text = r
        Text2.Text = g
        Text3.Text = B

End Sub


Private Sub txtH_KeyPress(KeyAscii As Integer)
    CheckValidity txtH, KeyAscii
End Sub

Private Sub txtH_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtH.Text) > 360 Or Val(txtH.Text) < 0 Or Int(Val(txtH.Text)) <> Val(txtH.Text) Then
    MsgBox "Enter an Integer Value Between 0 and 360"
End If

If Val(txtH.Text) < 0 Then txtH.Text = 0
    'txtH.Text = Int(Val(txtH.Text))
    If optH.Value And MainBoxHit = False Then
        Dim r As Integer, g As Integer, B As Integer
        Dim cl As Long

        DrawSlider SelectedMainPos
        SelectedMainPos = MainBox.Bottom - Val(txtH.Text) * 255 / 360
        DrawSlider SelectedMainPos
        cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
        GetRGB cl, r, g, B
        LoadVariantsHue r, g, B
        UpdateColorValues
     End If

End Sub

Private Sub txtH_LostFocus()
If Val(txtH.Text) > 360 Or Val(txtH.Text) < 0 Or Int(Val(txtH.Text)) <> Val(txtH.Text) Then
    MsgBox "Enter an Integer Value Between 0 and 360"
    txtH.SetFocus
End If
End Sub


Private Function CheckValidity(txtBox As TextBox, ascii As Integer) As Boolean
End Function

Private Sub txtS_KeyUp(KeyCode As Integer, Shift As Integer)
    If optH.Value Then
        DrawPicker
        SelectedPos.x = SelectBox.Left + Val(txtS.Text) * 255 / 100
        DrawPicker
        UpdateColorValues
    End If
End Sub
