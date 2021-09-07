Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hgdiobj As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, _
    ByVal nWidth As Long, _
    ByVal nEscapement As Long, _
    ByVal nOrientation As Long, _
    ByVal fnWeight As Long, _
    ByVal IfdwItalic As Long, _
    ByVal fdwUnderline As Long, _
    ByVal fdwStrikeOut As Long, _
    ByVal fdwCharSet As Long, _
    ByVal fdwOutputPrecision As Long, _
    ByVal fdwClipPrecision As Long, _
    ByVal fdwQuality As Long, _
    ByVal fdwPitchAndFamily As Long, _
    ByVal lpszFace As String) As Long
    
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
    ByVal lpStr As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal wFormat As Long) As Long

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_SCRIPT = 64
Private Const DT_CALCRECT = &H400

Function MeasureTextWidth( _
        target_text As String, _
        FONT_NAME As String, _
        Optional font_height As Long = 10) As Long
    
    Dim hWholeScreenDC As Long: hWholeScreenDC _
        = GetDC(0&)
    
    Dim hVirtualDC As Long: hVirtualDC _
        = CreateCompatibleDC(hWholeScreenDC)

    Dim hFont As Long: hFont _
        = CreateFont(font_height, 0, 0, 0, FW_NORMAL, _
            0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
            CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
            DEFAULT_PITCH Or FF_SCRIPT, FONT_NAME)
            
    Call SelectObject(hVirtualDC, hFont)
    
    Dim DrawAreaRectangle As RECT
    Call DrawText(hVirtualDC, target_text, -1, DrawAreaRectangle, DT_CALCRECT)
    
    Call DeleteObject(hFont)
    Call DeleteObject(hVirtualDC)
    Call ReleaseDC(0&, hWholeScreenDC)
    MeasureTextWidth = DrawAreaRectangle.Right - DrawAreaRectangle.Left
End Function

Sub 幅を揃えて出力()
    Const 基準テキスト = "固定幅のフォントは"
    Const 対象テキスト = "固定幅のフォントは単純に同じ文字数で改行すれば綺麗な矩形になるが、可変幅のフォントでは単純に同じ文字数で改行してもガタガタになる。"
    Const FONT_NAME = "ＭＳ Ｐゴシック"
    
    Dim 基準テキストの長さ As Long
    基準テキストの長さ = MeasureTextWidth(基準テキスト, FONT_NAME)
    
    Dim tmpText As String
    tmpText = ""
    Dim i As Long
    For i = 1 To Len(対象テキスト)
        If MeasureTextWidth(tmpText + Mid(対象テキスト, i, 1), FONT_NAME) > 基準テキストの長さ Then
            Debug.Print tmpText
            tmpText = Mid(対象テキスト, i, 1)
        Else
            tmpText = tmpText & Mid(対象テキスト, i, 1)
        End If
    Next
    Debug.Print tmpText
End Sub
