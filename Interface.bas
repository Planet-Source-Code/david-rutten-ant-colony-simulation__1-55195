Attribute VB_Name = "Interface"
Option Explicit

Public Offset(1) As Long
Public RedrawFrequency As Long

Public Function RefreshViewport() As Boolean
    If Settings.RenderGrid Then DrawTerraRaster
    If Settings.RenderScent Then DrawTerraGrid
    If Settings.RenderFood Then DrawFood
    If Settings.RenderLabels Then DrawTerraLabels
    If Settings.RenderAnts Then DrawAnts
End Function

Public Function DistributeForm() As Boolean
    On Error GoTo ErrorTrap
    Dim Canvas(1) As Long
    Canvas(0) = Main.ScaleWidth
    Canvas(1) = Main.ScaleHeight - 25
    
    'Resize the statusbar panes
    Main.Status.Panels(1).Width = (Canvas(0) - 20) \ 5
    Main.Status.Panels(2).Width = (Canvas(0) - 20) \ 5
    Main.Status.Panels(3).Width = (Canvas(0) - 20) \ 5
    Main.Status.Panels(4).Width = (Canvas(0) - 20) \ 5
    Main.Status.Panels(5).Width = (Canvas(0) - 20) \ 5
    'Position the control container frame
    Main.frmControls.Top = 0
    Main.frmControls.Left = Canvas(0) - Main.frmControls.Width
    'Position and resize the analysis frame container
    Main.frmGraph.Left = Main.frmControls.Left
    Main.frmGraph.Top = Main.frmControls.Height
    Main.frmGraph.Height = Region(Canvas(1) - Main.frmControls.Height, 1, 1000000#)
    'Position and resize the vertical scrollbar
    Main.vScroll.Top = 0
    Main.vScroll.Left = Main.frmControls.Left - Main.vScroll.Width - 10
    Main.vScroll.Height = Region(Canvas(1) - Main.hScroll.Height, 1, 1000000#)
    'Position and resize the horizontal scrollbar
    Main.hScroll.Left = 0
    Main.hScroll.Width = Region(Main.vScroll.Left, 1, 1000000#)
    Main.hScroll.Top = Canvas(1) - Main.hScroll.Height
    'Position the zoom button
    Main.ZoomExtends.Top = Main.hScroll.Top
    Main.ZoomExtends.Left = Main.vScroll.Left
    'Resize the terrarium viewport
    Main.Terra.Top = 0
    Main.Terra.Left = 0
    Main.Terra.Width = Region(Main.vScroll.Left, 1, 1000000#)
    Main.Terra.Height = Region(Main.hScroll.Top, 1, 1000000#)

    DistributeForm = True
    Exit Function
ErrorTrap:
    DistributeForm = False
End Function

Public Function FindColourFromQuad(ByRef QuadIn As Quad) As Long
    FindColourFromQuad = vbWhite
    Dim rgbZero(2) As Double
    Dim rgbFull(2) As Double
    Dim dblFactor As Double
    Dim r As Byte, g As Byte, b As Byte
    
    'Default background colour
    rgbZero(0) = 230
    rgbZero(1) = 230
    rgbZero(2) = 230
    
    If QuadIn.FoodScent > 0 Then
    'Define fully saturated colour
        rgbFull(0) = 240
        rgbFull(1) = 100
        rgbFull(2) = 190
        dblFactor = Region(QuadIn.FoodScent / 100, 0, 1)
    ElseIf QuadIn.DefaultScent > 0 Then
    'Define fully saturated colour
        rgbFull(0) = 150
        rgbFull(1) = 150
        rgbFull(2) = 150
        dblFactor = Region(QuadIn.DefaultScent / 100, 0, 1)
    Else
        rgbFull(0) = 230
        rgbFull(1) = 230
        rgbFull(2) = 230
        dblFactor = 0
    End If
    'Calculate colour based on saturation gradient
    r = rgbZero(0) + ((rgbFull(0) - rgbZero(0)) * dblFactor)
    g = rgbZero(1) + ((rgbFull(1) - rgbZero(1)) * dblFactor)
    b = rgbZero(2) + ((rgbFull(2) - rgbZero(2)) * dblFactor)
    FindColourFromQuad = RGB(r, g, b)
End Function

Public Function DrawAnts(Optional ByVal lngSize As Long = 2) As Boolean
    Dim i As Long
    If Settings.AntCount = 0 Then Exit Function
    
    For i = 0 To UBound(Ants)
        Call Render.AntCreature(Main.Terra, Ants(i), lngSize)
    Next
    DrawAnts = True
End Function

Public Function DrawTerraRaster(Optional ByVal rgbRasterColour As Long = 13158600) As Boolean
    Dim i As Long, j As Long
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    
    Y1 = -Offset(1)
    Y2 = Y1 + Settings.GridSize * Settings.TerraExtend
    Main.Terra.DrawStyle = 0
    For i = 0 To UBound(TerraGrid, 1)
        X1 = i * Settings.GridSize - Offset(0)
        X2 = X1
        Main.Terra.Line (X1, Y1)-(X2, Y2), rgbRasterColour
    Next
    
    X1 = -Offset(0)
    X2 = X1 + Settings.GridSize * Settings.TerraExtend
    For j = 0 To UBound(TerraGrid, 2)
        Y1 = j * Settings.GridSize - Offset(1)
        Y2 = Y1
        Main.Terra.Line (X1, Y1)-(X2, Y2), rgbRasterColour
    Next
    DrawTerraRaster = True
End Function

Public Function DrawTerraGrid(Optional ByVal blnIncludeFills As Boolean = True, _
                              Optional ByVal blnIncludeEmptyGrids As Boolean = False, _
                              Optional ByVal blnIncludeTarget As Boolean = True) As Boolean
    Dim i As Long, j As Long
    Dim X As Long, Y As Long
    Dim GridProps As Quad
    Dim rgbOutline As Long
    Dim rgbFill As Long
    Dim xMid As Long, yMid As Long
    
    For i = 0 To UBound(TerraGrid, 1)
    For j = 0 To UBound(TerraGrid, 2)
        GridProps = TerraGrid(i, j)
        If blnIncludeEmptyGrids Or _
           GridProps.DefaultScent <> 0 Or _
           GridProps.FoodScent <> 0 Then
            X = i * Settings.GridSize - Offset(0)
            Y = j * Settings.GridSize - Offset(1)
            If blnIncludeFills Then
                rgbFill = FindColourFromQuad(GridProps)
            Else
                rgbFill = -1
            End If

            Call Render.GridElement(Main.Terra, X, Y, Settings.GridSize, Settings.GridSize, rgbOutline, rgbFill)
        End If
    Next
    Next
    
    If blnIncludeTarget Then
        xMid = Int(Main.Terra.ScaleWidth / 2)
        yMid = Int(Main.Terra.ScaleHeight / 2)
        Main.Terra.Line (xMid, yMid - 5)-(xMid, yMid + 6), 0
        Main.Terra.Line (xMid - 5, yMid)-(xMid + 6, yMid), 0
    End If
    
    Call Render.UpdateDevice(Main.Terra)
    DrawTerraGrid = True
End Function

Public Function DrawTerraLabels() As Boolean
    Dim i As Long, j As Long
    Dim X As Long, Y As Long
    Dim GridProps As Quad
    Dim strLabel As String
    
    For i = 0 To UBound(TerraGrid, 1)
    For j = 0 To UBound(TerraGrid, 2)
        GridProps = TerraGrid(i, j)
        X = i * Settings.GridSize - Offset(0)
        Y = j * Settings.GridSize - Offset(1)
        
        If GridProps.FoodAmount > 0 Then
            strLabel = CStr(GridProps.FoodAmount)
            Call Render.GridLabel(Main.Terra, X, Y, strLabel, 0, 12)
            GoTo SkipLoop
        End If
        
        If GridProps.FoodScent > 0 Then
            strLabel = CStr(GridProps.FoodScent)
            Call Render.GridLabel(Main.Terra, X, Y, strLabel, 0, 12)
        ElseIf GridProps.DefaultScent > 0 Then
            strLabel = CStr(GridProps.DefaultScent)
            Call Render.GridLabel(Main.Terra, X, Y, strLabel, 0, 12)
        End If
SkipLoop:
    Next
    Next
    
    Call Render.UpdateDevice(Main.Terra)
    DrawTerraLabels = True
End Function

Public Function DrawFood() As Boolean
    Dim i As Long, j As Long
    Dim X As Long, Y As Long
    Dim GridProps As Quad
    Dim rgbFill As Long
    Dim strLabel As String
    
    Main.Terra.FillStyle = 0
    For i = 0 To UBound(TerraGrid, 1)
    For j = 0 To UBound(TerraGrid, 2)
        GridProps = TerraGrid(i, j)
        If GridProps.IsHome Then
            Main.Terra.DrawStyle = 0
            Main.Terra.FillStyle = 0
            Main.Terra.FillColor = RGB(190, 90, 30)
            X = i * Settings.GridSize - Offset(0)
            Y = j * Settings.GridSize - Offset(1)
            Main.Terra.Circle (X + Settings.GridSize / 2, Y + Settings.GridSize / 2), Settings.GridSize \ 3, 0
        ElseIf GridProps.FoodAmount > 0 Then
            Main.Terra.DrawStyle = 0
            Main.Terra.FillStyle = 0
            Main.Terra.FillColor = RGB(40, 200, 10)
            X = i * Settings.GridSize - Offset(0)
            Y = j * Settings.GridSize - Offset(1)
            Main.Terra.Circle (X + Settings.GridSize / 2, Y + Settings.GridSize / 2), _
                              Region(GridProps.FoodAmount / 2, 2, Settings.GridSize / 2.2), 0
        Else
            'no action required
        End If
    Next
    Next
    
    Call Render.UpdateDevice(Main.Terra)
    DrawFood = True
End Function

Public Function DrawVerticalSlider() As Boolean
    Dim TerraSize As Long
    Dim BarSize As Long
    Dim BarTop As Long
    
    TerraSize = Settings.TerraExtend * Settings.GridSize
    BarTop = (Main.vScroll.ScaleHeight / TerraSize) * Offset(1)
    BarTop = Region(BarTop, 1, Main.vScroll.ScaleHeight - 14)
    BarSize = (Main.vScroll.ScaleHeight / TerraSize) * Main.vScroll.ScaleHeight
    BarSize = Region(BarSize, 13, Main.vScroll.ScaleHeight - BarTop - 2)
    
    Main.vScroll.BackColor = RGB(230, 230, 230)
    Main.vScroll.Cls
    Call Render.GridElement(Main.vScroll, 1, Region(BarTop, 1, 1000000000000#), 13, BarSize, 0, RGB(200, 200, 200))
    Call Render.UpdateDevice(Main.vScroll)
    DrawVerticalSlider = True
End Function

Public Function DrawHorizontalSlider() As Boolean
    Dim TerraSize As Long
    Dim BarSize As Long
    Dim Barleft As Long
    
    TerraSize = Settings.TerraExtend * Settings.GridSize
    Barleft = (Main.hScroll.ScaleWidth / TerraSize) * Offset(0)
    Barleft = Region(Barleft, 1, Main.hScroll.ScaleWidth - 14)
    BarSize = (Main.hScroll.ScaleWidth / TerraSize) * Main.hScroll.ScaleWidth
    BarSize = Region(BarSize, 13, Main.hScroll.ScaleWidth - Barleft - 2)
    
    Main.hScroll.BackColor = RGB(230, 230, 230)
    Main.hScroll.Cls
    Call Render.GridElement(Main.hScroll, Region(Barleft, 1, 1000000000000#), 1, BarSize, 13, 0, RGB(200, 200, 200))
    Call Render.UpdateDevice(Main.hScroll)
    DrawHorizontalSlider = True
End Function

Public Function Settings2Interface() As Boolean
    Main.txtGridSize.Text = CStr(Settings.GridSize)
    Main.txtExtend.Text = CStr(Settings.TerraExtend)
    Main.txtAntAge.Text = CStr(Settings.AntAge)
    Main.txtColonySize.Text = CStr(Settings.ColonySize)
    Main.txtFoodDensity.Text = CStr(Settings.BioMatter)
    Main.txtMaxCargo.Text = CStr(Settings.MaxCargo)
    Main.txtBirth.Text = CStr(Settings.Birth)
    Main.txtIterationRatio.Text = CStr(Settings.IterationRatio)
    Settings2Interface = True
End Function

Public Function Interface2Settings() As Boolean
    Settings.GridSize = Region(CInt(Val(Main.txtGridSize.Text)), 5, 100)
    Settings.TerraExtend = Region(CInt(Val(Main.txtExtend.Text)), 10, 200)
    Settings.AntAge = CInt(Val(Main.txtAntAge.Text))
    Settings.ColonySize = Region(CInt(Val(Main.txtColonySize.Text)), 1, 1000)
    Settings.BioMatter = Region(Val(Main.txtFoodDensity.Text), 0, 1)
    Settings.MaxCargo = Region(CInt(Val(Main.txtMaxCargo.Text)), 1, 9)
    Settings.Birth = Region(CInt(Val(Main.txtBirth.Text)), 1, 100)
    Settings.IterationRatio = Region(CInt(Val(Main.txtIterationRatio)), 1, 500)
    Main.txtBirth.Text = CStr(Settings.Birth)
    Interface2Settings = True
    Call Settings2Interface
End Function

Public Function CenterViewport() As Boolean
    Offset(0) = (Settings.TerraExtend * Settings.GridSize - Main.Terra.ScaleWidth) / 2
    Offset(1) = (Settings.TerraExtend * Settings.GridSize - Main.Terra.ScaleHeight) / 2
    Main.Terra.Cls
    Call RefreshViewport
    Call DrawVerticalSlider
    Call DrawHorizontalSlider
    CenterViewport = True
End Function
