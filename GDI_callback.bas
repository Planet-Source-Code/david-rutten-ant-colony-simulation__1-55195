Attribute VB_Name = "GDI_callback"
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RectAPI, ByVal lpString As String, ByVal nCount As Long, Optional lpDx As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As PointAPI) As Long
Declare Function CreateHatchBrush Lib "gdi32" (ByVal vHatchType As Long, ByVal crColor As Long) As Long
Declare Function CoCreateGuid Lib "ole32" (ID As Any) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const Pi = 3.1415926535

Public Enum BannerOption
    STARTUP = 0
    MISTAKE = 1
    DEATH = 2
    COMPLETE = 3
End Enum

Public Type Ant
    X As Long
    Y As Long
    Age As Long
    Cargo As Long
    Direction As Double
    ID As String
End Type

Public Type Quad
    IsHome As Boolean
    FoodAmount As Long
    DefaultScent As Long
    FoodScent As Long
    i As Long
    j As Long
    ID As String
End Type

Public Type Values
    AntAge As Long
    MaxCargo As Long
    AntCount As Long
    ColFood As Long
    Birth As Long
    Transit As Long
    CycleTime As Long
    HomePoint As Point
    TerraExtend As Long
    GridSize As Long
    ColonySize As Long
    BioMatter As Double
    IterationRatio As Long
    RenderGrid As Boolean
    RenderScent As Boolean
    RenderLabels As Boolean
    RenderAnts As Boolean
    RenderFood As Boolean
End Type

Public Type PointAPI
    X As Long
    Y As Long
End Type

Public Type RectAPI
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Function CreateGUID() As String
    Dim ID(15) As Byte
    Dim Cnt As Long, GUID As String
    If CoCreateGuid(ID(0)) = 0 Then
        For Cnt = 0 To 15
            CreateGUID = CreateGUID + IIf(ID(Cnt) < 16, "0", "") + Hex$(ID(Cnt))
        Next Cnt
        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
    Else
        MsgBox "Error while creating GUID!"
    End If
End Function
