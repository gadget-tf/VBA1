'Option Explicit

Private Type LOGICAL_FONT
    height As Long
    width As Long
    escapement As Long
    orientation As Long
    weight As Long
    italic As Byte
    underline As Byte
    strikeOut As Byte
    charSet As Byte
    outPrecision As Byte
    clipPrecision As Byte
    quality As Byte
    pitchAndFamily As Byte
    faceName As String * 32
End Type

Private Type FONT_SIZE
    cx As Long
    cy As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hgdiobj As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, _
    ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, _
    ByVal fnWeight As Long, ByVal IfdwItalic As Long, ByVal fdwUnderline As Long, _
    ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, _
    ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
    ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGICAL_FONT) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As FONT_SIZE) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private Const LOGPIXELSY As Long = 90
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_SCRIPT = 64
Private Const DT_CALCRECT = &H400

Sub Test3()
    Dim sz As FONT_SIZE
    Dim fontName As String
    Dim fontSize As Long
    Dim text As String
    Dim fnt As New StdFont
    Dim result As Long
    
    With Worksheets("Sheet4").Range("A1")
        text = .Value
        fontSize = .font.Size
        fontName = .font.Name
    End With
    
    fnt.Name = fontName
    fnt.Size = fontSize
    
    sz = GetLabelSize(Mid(text, 1, 1), fnt)
    
    result = MeasureTextWidth(text, fnt.Name, sz.cy)
    
    
    Worksheets("Sheet4").Range("A2").Value = result
    Worksheets("Sheet4").Range("A3").Value = sz.cy
    
End Sub

Function MeasureTextWidth(text As String, FONT_NAME As String, ByVal font_height As Long) As Long
    Dim hWholeScreenDC As Long: hWholeScreenDC = GetDC(0&)
    Dim hVirtualDC As Long: hVirtualDC = CreateCompatibleDC(hWholeScreenDC)

    Dim hFont As Long: hFont _
        = CreateFont(font_height, 0, 0, 0, FW_NORMAL, _
            0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
            CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
            DEFAULT_PITCH Or FF_SCRIPT, FONT_NAME)
    Dim DrawAreaRectangle As RECT

    Call SelectObject(hVirtualDC, hFont)

    Call DrawText(hVirtualDC, text, -1, DrawAreaRectangle, DT_CALCRECT)

    Call DeleteObject(hFont)
    Call DeleteObject(hVirtualDC)
    Call ReleaseDC(0&, hWholeScreenDC)

    MeasureTextWidth = DrawAreaRectangle.Right - DrawAreaRectangle.Left

End Function

Function GetLabelSize(text As String, font As StdFont) As FONT_SIZE
    Dim tempDC As Long
    Dim tempBMP As Long
    Dim hFnt As Long
    Dim LF As LOGICAL_FONT
    Dim textSize As FONT_SIZE

    tempDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0)
    tempBMP = CreateCompatibleBitmap(tempDC, 1, 1)

    Call DeleteObject(SelectObject(tempDC, tempBMP))

    LF.faceName = font.Name & Chr$(0)
    LF.height = -MulDiv(font.Size, GetDeviceCaps(GetDC(0), 90), 72) 'LOGPIXELSY
    LF.italic = font.italic
    LF.underline = font.underline

    If font.Bold Then LF.weight = 800 Else LF.weight = 400

    hFnt = CreateFontIndirect(LF)

    Call DeleteObject(SelectObject(tempDC, hFnt))

    Call GetTextExtentPoint32(tempDC, text, Len(text), textSize)

    Call DeleteObject(hFnt)
    Call DeleteObject(tempBMP)
    Call DeleteDC(tempDC)

    GetLabelSize = textSize

End Function
