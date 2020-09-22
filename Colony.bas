Attribute VB_Name = "Colony"
Option Explicit
    
Public Settings As Values
Public Ants() As Ant
Public TerraGrid() As Quad
Public tabCount(99) As Long
Public tabCargo(99) As Long

Public Function SimulateTerrarium() As Boolean
    On Error GoTo ErrorTrap
    Dim frames As Long
    Dim StartTick As Long
    Dim StopTick As Long
    Dim i As Long, C As Long
    
    frames = 1
    Do
        If Main.blnPauseSolution Then Exit Do
        If Main.blnStopSolution Then Exit Do
        StartTick = GetTickCount
        If RedrawFrequency = 1 Then
            Call AnimateAnts(Settings.IterationRatio, True)
        Else
            Call AnimateAnts(Settings.IterationRatio, False)
            If frames / RedrawFrequency = frames \ RedrawFrequency Then
                Main.Terra.Cls
                Call Interface.RefreshViewport
                DoEvents
            End If
        End If
        
        If Main.blnPauseSolution Then Exit Do
        If Main.blnStopSolution Then Exit Do
        If Settings.AntCount = 0 Then
            MsgBox "Your colony became extinct.", vbOKOnly, "Death was inevitable"
            Main.cmdStopSimulation_Click
            Exit Function
        End If
        Call AnimateTerraGrid(1, True)
        StopTick = GetTickCount
        Settings.CycleTime = (Settings.CycleTime * 5 + (StopTick - StartTick)) / 6
        
        If Main.blnPauseSolution Then Exit Do
        If Main.blnStopSolution Then Exit Do
        
        Call UpdateTables
        Call RenderTables
        
        frames = frames + 1
        If frames = 100000 Then Exit Do
        
        C = 0
        Settings.Transit = 0
        For i = 0 To UBound(Ants)
            If Ants(i).Cargo > 0 Then
                C = C + 1
                Settings.Transit = Settings.Transit + Ants(i).Cargo
            End If
        Next
        
        Main.Status.Panels(1).Text = "Number of ants: " & Settings.AntCount
        Main.Status.Panels(2).Text = "Food carriers: " & C
        Main.Status.Panels(3).Text = "Cargo in transit: " & Settings.Transit
        Main.Status.Panels(4).Text = "Cycle time: " & Settings.CycleTime & " ms"
        Main.Status.Panels(5).Text = "Iterations: " & frames
    Loop
    
    Main.Status.Panels(1).Text = "Number of ants: "
    Main.Status.Panels(2).Text = "Food carriers: "
    Main.Status.Panels(3).Text = "Cargo in transit: "
    Main.Status.Panels(4).Text = "Cycle time: "
    Main.Status.Panels(5).Text = "Iterations: "
    
    SimulateTerrarium = True
    Exit Function
ErrorTrap:
    SimulateTerrarium = False
End Function

Public Function InitiateTerrarium() As Boolean
    On Error GoTo ErrorTrap
    Dim EmptyQuad As Quad
    Dim i As Long, j As Long

    ReDim TerraGrid(Settings.TerraExtend, Settings.TerraExtend)
    For i = 0 To Settings.TerraExtend - 1
    For j = 0 To Settings.TerraExtend - 1
        
        EmptyQuad = CreateQuad(i, j)
        TerraGrid(i, j) = EmptyQuad
        If Rnd < Settings.BioMatter Then
            TerraGrid(i, j).FoodAmount = Int(Rnd * 10)
        End If
        
        If i = Settings.TerraExtend \ 2 And j = Settings.TerraExtend \ 2 Then
            TerraGrid(i, j).FoodAmount = 0
            TerraGrid(i, j).IsHome = True
            Settings.HomePoint.X = Settings.GridSize * i + Settings.GridSize \ 2
            Settings.HomePoint.Y = Settings.GridSize * j + Settings.GridSize \ 2
        End If
    Next
    Next
    
    InitiateTerrarium = True
    Exit Function
ErrorTrap:
    InitiateTerrarium = False
End Function

Public Function InitiateAnts(ByVal lngAmount As Long) As Boolean
    Dim i As Long
    Dim newAnt As Ant
    
    For i = 0 To lngAmount - 1
        newAnt = CreateAnt(Settings.HomePoint.X, Settings.HomePoint.Y)
        Call AddAnt(newAnt)
    Next
    InitiateAnts = True
End Function

Public Function AnimateTerraGrid(Optional ByVal lngFrames As Long = 1, _
                                 Optional ByVal blnRedraw As Boolean = False) As Boolean
    On Error GoTo ErrorTrap
    Dim i As Long, j As Long
    Dim frames As Long
    
    For frames = 1 To lngFrames
        For i = 0 To UBound(TerraGrid, 1) - 1
        For j = 0 To UBound(TerraGrid, 2) - 1
            If TerraGrid(i, j).IsHome Then
                'Set properties for the home quad
                TerraGrid(i, j).FoodAmount = 0
                TerraGrid(i, j).FoodScent = 0
                TerraGrid(i, j).DefaultScent = 1000
            Else
                'Reduce scents on quad
                TerraGrid(i, j).DefaultScent = Region(TerraGrid(i, j).DefaultScent * 0.95 - 1, 0, 1000)
                TerraGrid(i, j).FoodScent = Region(TerraGrid(i, j).FoodScent * 0.85 - 1, 0, 1000)
                If TerraGrid(i, j).FoodAmount > 0 Then
                    'Probable growth of food in quad
                    If Rnd < 0.5 Then
                        TerraGrid(i, j).FoodAmount = Region(TerraGrid(i, j).FoodAmount + 1, 0, 25)
                        TerraGrid(i, j).FoodScent = 10
                    End If
                Else
                    'Create new food at random
                    If Rnd < 0.0002 Then
                        TerraGrid(i, j).FoodAmount = 1
                        TerraGrid(i, j).FoodScent = 10
                    End If
                End If
            End If
        Next
        Next
        
        If blnRedraw Then
            Main.Terra.Cls
            Call Interface.RefreshViewport
            DoEvents
        End If
    Next
    
    AnimateTerraGrid = True
    Exit Function
ErrorTrap:
    AnimateTerraGrid = False
End Function

Public Function AnimateAnts(Optional ByVal lngFrames As Long = 1, _
                            Optional ByVal blnRedraw As Boolean = True) As Boolean
    On Error GoTo ErrorTrap
    Dim CurQuad As Quad, NewQuad As Quad
    Dim newAnt As Ant
    Dim NewLoc(1) As Long
    Dim FrameCount As Long, i As Long
    Dim DistanceToHome As Double
    Dim HomeDirection As Double
    
    For FrameCount = 1 To lngFrames
        i = 0
        Do
            'Make sure we don't overstep the Ant array
            If i > UBound(Ants) Then Exit Do
            'Calculate current quad
            CurQuad = QuadBelowAnt(Ants(i))
            
            'Increase ant age and check for terminal age. If ant dies then drop cargo
            Ants(i).Age = Ants(i).Age + CInt(Rnd)
            If Ants(i).Age >= Settings.AntAge And Settings.AntAge > 0 Then
                TerraGrid(CurQuad.i, CurQuad.j).FoodAmount = CurQuad.FoodAmount + Ants(i).Cargo
                Settings.Transit = Settings.Transit - Ants(i).Cargo
                Call RemoveAnt(Ants(i).ID)
                If Settings.AntCount <= 0 Then Exit Function
                GoTo SkipLoop
            End If
            
            'Give birth to new ant if sufficient food has been collected
            If Settings.ColFood >= Settings.Birth Then
                newAnt = CreateAnt(Settings.HomePoint.X, Settings.HomePoint.Y)
                Call AddAnt(newAnt)
                Settings.ColFood = Settings.ColFood - Settings.Birth
            End If
            
            'Add scent to terrarium
            If Ants(i).Cargo > 0 Then
                TerraGrid(CurQuad.i, CurQuad.j).FoodScent = CurQuad.FoodScent + Ants(i).Cargo
            Else
                TerraGrid(CurQuad.i, CurQuad.j).DefaultScent = CurQuad.DefaultScent + 1
            End If
            
            'Calculate walkto point
            NewLoc(0) = Ants(i).X + Sin(Ants(i).Direction) * Region((Settings.MaxCargo - Ants(i).Cargo), 1, 4)
            NewLoc(1) = Ants(i).Y + Cos(Ants(i).Direction) * Region((Settings.MaxCargo - Ants(i).Cargo), 1, 4)
            'Check if the ant oversteps the terrarium border. If so then turn around
            If Not IsPointOnTerrarium(CreatePoint(NewLoc(0), NewLoc(1))) Then
                Ants(i).Direction = Ants(i).Direction + Pi
                GoTo SkipLoop
            End If
            'Create platonic ant
            newAnt = CreateAnt(NewLoc(0), NewLoc(1))
            'Calculate new quad
            NewQuad = QuadBelowAnt(newAnt)
            
            'Check if the ant is home. If so then simply walk straight and transfer cargo
            If CurQuad.IsHome Then
                Settings.ColFood = Settings.ColFood + Ants(i).Cargo
                Ants(i).Cargo = 0
                Ants(i).X = NewLoc(0)
                Ants(i).Y = NewLoc(1)
                GoTo SkipLoop
            End If
            
            'Check if the ant can carry more food and if the current quad offers food
            If Ants(i).Cargo < Settings.MaxCargo And CurQuad.FoodAmount > 0 Then
                Do
                    If Ants(i).Cargo = Settings.MaxCargo Then Exit Do
                    If TerraGrid(CurQuad.i, CurQuad.j).FoodAmount = 0 Then Exit Do
                    Ants(i).Cargo = Ants(i).Cargo + 1
                    TerraGrid(CurQuad.i, CurQuad.j).FoodAmount = TerraGrid(CurQuad.i, CurQuad.j).FoodAmount - 1
                    Settings.Transit = Settings.Transit + 1
                Loop
                'Change ant direction
                Ants(i).Direction = Ants(i).Direction + Pi
                'Calculate new walkto point
                NewLoc(0) = Ants(i).X + Sin(Ants(i).Direction) * Region((Settings.MaxCargo - Ants(i).Cargo), 1, 4)
                NewLoc(1) = Ants(i).Y + Cos(Ants(i).Direction) * Region((Settings.MaxCargo - Ants(i).Cargo), 1, 4)
                'Check if the ant oversteps the terrarium border. If so then turn around
                If Not IsPointOnTerrarium(CreatePoint(NewLoc(0), NewLoc(1))) Then
                    Ants(i).Direction = Ants(i).Direction + Pi
                    GoTo SkipLoop
                End If
                'Create platonic ant
                newAnt = CreateAnt(NewLoc(0), NewLoc(1))
                'Calculate new quad
                NewQuad = QuadBelowAnt(newAnt)
            End If
            
            'Check if the ant goes home. If so then transfer cargo and turn around
            If NewQuad.IsHome Then
                Ants(i).Direction = Ants(i).Direction + Pi
                Settings.ColFood = Settings.ColFood + Ants(i).Cargo
                Settings.Transit = Settings.Transit - Ants(i).Cargo
                Ants(i).Cargo = 0
                GoTo SkipLoop
            End If
            
            If Ants(i).Cargo = 0 Then
                'Make sure the ant doesn't leave the food track
                If CurQuad.FoodScent > 0 And NewQuad.FoodScent = 0 Then
                    Ants(i).Direction = Ants(i).Direction + Random(Pi / 3)
                Else
                    'Move the ant to the new location
                    Ants(i).X = NewLoc(0)
                    Ants(i).Y = NewLoc(1)
                    Ants(i).Direction = Ants(i).Direction + Random(Pi / 15)
                End If
            Else
                'Determine if home is nearby
                DistanceToHome = (Ants(i).X - Settings.HomePoint.X) ^ 2 + _
                                 (Ants(i).Y - Settings.HomePoint.Y) ^ 2
                DistanceToHome = Sqr(DistanceToHome)
                If DistanceToHome < Settings.GridSize * 2 Then
                    'Walk straight home
                    HomeDirection = DetermineHomeDirection(Ants(i))
                    Ants(i).Direction = (HomeDirection + 5 * Ants(i).Direction) / 6
                    Ants(i).Direction = HomeDirection
                    'Calculate new walkto point
                    Ants(i).X = Ants(i).X + Sin(Ants(i).Direction) * 2
                    Ants(i).Y = Ants(i).Y + Cos(Ants(i).Direction) * 2
                    GoTo SkipLoop
                End If
                
                'Make sure the ant walks into a quad with stronger scent
                'If (CurQuad.DefaultScent) > 0 And NewQuad.DefaultScent = 0 Then
                If CurQuad.DefaultScent - 10 > NewQuad.DefaultScent Then
                    Ants(i).Direction = Ants(i).Direction + Random(Pi / 3)
                Else
                    'Move the ant to the new location
                    Ants(i).X = NewLoc(0)
                    Ants(i).Y = NewLoc(1)
                    Ants(i).Direction = Ants(i).Direction + Random(Pi / 15)
                End If
            End If
SkipLoop:
            'Jump to above label to avoid code in the loop
            i = i + 1
        Loop
        
        If blnRedraw Then
            Main.Terra.Cls
            Call Interface.RefreshViewport
            DoEvents
        End If
    Next
    AnimateAnts = True
    Exit Function
ErrorTrap:
    AnimateAnts = False
End Function

Public Function InitiateTables() As Boolean
    Dim i As Long
    For i = 0 To UBound(tabCount)
        tabCount(i) = 0
    Next
    For i = 0 To UBound(tabCargo)
        tabCargo(i) = 0
    Next
    InitiateTables = True
End Function

Public Function UpdateTables() As Boolean
    Dim i As Long
    For i = 0 To UBound(tabCount) - 1
        tabCount(i) = tabCount(i + 1)
    Next
    tabCount(UBound(tabCount)) = Region((Settings.AntCount / Settings.ColonySize) * 20, 4, 10000000000#)
    
    For i = 0 To UBound(tabCargo) - 1
        tabCargo(i) = tabCargo(i + 1)
    Next
    tabCargo(UBound(tabCargo)) = Region(80 * (Settings.Transit / (Settings.AntCount * Settings.MaxCargo)), 4, 10000000000#)
    UpdateTables = True
End Function

Public Function RenderTables() As Boolean
    Dim i As Long
    
    Main.GraphCount.Cls
    For i = 0 To UBound(tabCount)
        Main.GraphCount.Line (i + 45, 80)-(i + 45, 80 - tabCount(i)), RGB(100, 180, 30)
        Main.GraphCount.Line (i + 45, 80)-(i + 45, 80 - tabCount(i) / 10), RGB(50, 90, 15)
    Next
    
    Main.GraphCargo.Cls
    For i = 0 To UBound(tabCargo)
        Main.GraphCargo.Line (i + 45, 80)-(i + 45, 80 - tabCargo(i)), RGB(200, 0, 50)
    Next
    RenderTables = True
End Function

Public Function DetermineHomeDirection(ByRef Critter As Ant) As Double
    On Error GoTo ErrorTrap
    Dim dx As Double, dy As Double
    dx = Settings.HomePoint.X - Critter.X + 0.001 'add a subpixel value to avoid division by zero
    dy = Settings.HomePoint.Y - Critter.Y
    If dx < -Settings.GridSize \ 3 Then DetermineHomeDirection = Pi * 1.5: Exit Function
    If dx > Settings.GridSize \ 3 Then DetermineHomeDirection = Pi * 0.5: Exit Function
    If dy < 0 Then DetermineHomeDirection = Pi: Exit Function
    If dy > 0 Then DetermineHomeDirection = 0: Exit Function
ErrorTrap:
    If dx = 0 And dy = 0 Then DetermineHomeDirection = Critter.Direction
End Function

Public Function Random(Optional ByVal dblRange As Double = 1)
    Random = Rnd * 2 - 1
    Random = Random * dblRange
End Function

Public Function IsQuad(ByVal i As Long, ByVal j As Long) As Boolean
    IsQuad = False
    If i < 0 Or j < 0 Then Exit Function
    If i > Settings.TerraExtend Or j > Settings.TerraExtend Then Exit Function
    IsQuad = True
End Function

Public Function IsPointOnTerrarium(ByRef TestPoint As Point) As Boolean
    IsPointOnTerrarium = True
    If TestPoint.X < 1 Or TestPoint.Y < 1 Then IsPointOnTerrarium = False
    If TestPoint.X > Settings.GridSize * Settings.TerraExtend - 1 Then IsPointOnTerrarium = False
    If TestPoint.Y > Settings.GridSize * Settings.TerraExtend - 1 Then IsPointOnTerrarium = False
End Function

Public Function QuadBelowAnt(ByRef WhichAnt As Ant) As Quad
    Dim i As Long, j As Long
    i = WhichAnt.X \ Settings.GridSize
    j = WhichAnt.Y \ Settings.GridSize
    If Not IsQuad(i, j) Then
        QuadBelowAnt = CreateQuad(-1, -1, -1, -1, -1, False)
    Else
        QuadBelowAnt = TerraGrid(i, j)
    End If
End Function

Public Function CreateQuad(ByVal indexI As Long, _
                           ByVal indexJ As Long, _
                           Optional ByVal lngDefaultScent As Long = 0, _
                           Optional ByVal lngFoodScent As Long = 0, _
                           Optional ByVal lngFoodAmount As Long = 0, _
                           Optional ByVal blnIsHome As Boolean = False) As Quad
    CreateQuad.DefaultScent = lngDefaultScent
    CreateQuad.FoodScent = lngFoodScent
    CreateQuad.FoodAmount = lngFoodAmount
    CreateQuad.IsHome = blnIsHome
    CreateQuad.ID = CreateGUID
    CreateQuad.i = indexI
    CreateQuad.j = indexJ
End Function

Public Function IsArrayNull(ByVal arrIn) As Boolean
    On Error GoTo ErrorTrap
    Dim ub As Long
    ub = UBound(arrIn)
    IsArrayNull = False
    Exit Function
ErrorTrap:
    IsArrayNull = True
End Function

Public Function CreateAnt(X, Y) As Ant
    CreateAnt.X = X
    CreateAnt.Y = Y
    CreateAnt.Age = 0
    CreateAnt.Cargo = 0
    CreateAnt.Direction = Rnd * 2 * Pi
    CreateAnt.ID = CreateGUID
End Function

Public Function NukeAnts() As Boolean
    Erase Ants
    Settings.AntCount = 0
    NukeAnts = True
End Function

Public Function AddAnt(ByRef FamilyMember As Ant) As Boolean
    AddAnt = False
    On Error GoTo ErrorTrap
    If Settings.AntCount = 0 Then
        ReDim Ants(0)
        Ants(0) = FamilyMember
    Else
        ReDim Preserve Ants(UBound(Ants) + 1)
        Ants(UBound(Ants)) = FamilyMember
    End If
    
    Settings.AntCount = Settings.AntCount + 1
    AddAnt = True
    Exit Function
ErrorTrap:
    Exit Function
End Function

Public Function RemoveAnt(ByVal GUID As String) As Boolean
    RemoveAnt = False
    On Error GoTo ErrorTrap
    Dim i As Long, j As Long
    If Settings.AntCount = 0 Then Exit Function
    Settings.AntCount = Settings.AntCount - 1
    If Settings.AntCount <= 0 Then Exit Function
    
    For i = 0 To UBound(Ants)
        If Ants(i).ID = GUID Then
            If i = UBound(Ants) Then
                Exit For
            Else
                For j = i To UBound(Ants) - 1
                    Ants(j) = Ants(j + 1)
                Next
                Exit For
            End If
        End If
    Next
    
    ReDim Preserve Ants(UBound(Ants) - 1)
    RemoveAnt = True
    Exit Function
ErrorTrap:
    Exit Function
End Function
