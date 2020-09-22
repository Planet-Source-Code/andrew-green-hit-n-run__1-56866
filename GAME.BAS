Attribute VB_Name = "Game"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const FrameDelay As Integer = 40
Public Type GCharecter
    X As Long
    Y As Long
    Angle As Integer
    Alive As Boolean
End Type
Public Player As GCharecter
Public Person() As GCharecter
Public Picture(3) As DirectDrawSurface7
Public GOps As Boolean
Public AIT As String
Public AIRnd As Boolean
Public GHelp As Boolean
Public Function BeginGLoop()
Dim lastTick As Long
Randomize Timer
ReDim Person(5)

Set Picture(0) = CreateSurfaceFromFile("Pics/Car.gif")
Set Picture(1) = CreateSurfaceFromFile("Pics/Alive.bmp")
Set Picture(2) = CreateSurfaceFromFile("Pics/Dead.bmp")

AIT = "None"
AIRnd = True

Player.X = 640 / 2
Player.Y = 480 / 2
For X = 0 To UBound(Person())
    Person(X).X = Rnd() * 640
    Person(X).Y = Rnd() * 480
    Person(X).Alive = True
Next X



Do
    DoEvents
    
'----------Keys----------
    If GetAsyncKeyState(27) Then 'Esc Key:EXIT
        For X = 0 To 2
            Set Picture(X) = Nothing
        Next X
        EndDX
        EGLoop = True
        End
    End If
    If GetAsyncKeyState(37) And (GetAsyncKeyState(38) Or GetAsyncKeyState(40)) Then 'Left Key:Rotate left
        Player.Angle = Player.Angle - 3
    End If
    If GetAsyncKeyState(39) And (GetAsyncKeyState(38) Or GetAsyncKeyState(40)) Then 'Right Key:Rotate right
        Player.Angle = Player.Angle + 3
    End If
    If GetAsyncKeyState(38) Then 'Up Key: Move forward
        Player.X = Player.X + Cos(Player.Angle * 3.14159265358979 / 180) * 15
        Player.Y = Player.Y + Sin(Player.Angle * 3.14159265358979 / 180) * 15
    End If
    If GetAsyncKeyState(40) Then 'Down Key:Move back
        Player.X = Player.X - Cos(Player.Angle * 3.14159265358979 / 180) * 5
        Player.Y = Player.Y - Sin(Player.Angle * 3.14159265358979 / 180) * 5
    End If
    If GetAsyncKeyState(107) Then 'Addition key:Add more people
        ReDim Preserve Person(UBound(Person()) + 1)
        Person(UBound(Person())).X = Rnd() * 640
        Person(UBound(Person())).Y = Rnd() * 480
        Person(UBound(Person())).Alive = True
    End If
    If GetAsyncKeyState(46) Then ClearBodies 'Del Key:Clear all bodies

'----------Keys----------

    LoopScreen Player
    Clr
    Player.Angle = Player.Angle Mod 360
   
    For X = 0 To UBound(Person())
        DoEvents
        If Person(X).Alive = True Then
            PersonAI Person(X)
            If CircularCollision(Player.X, Player.Y, Person(X).X, Person(X).Y, 15, 10) Then Person(X).Alive = False
        End If
        BltRotate BackBuffer, Picture(Person(X).Alive + 2), Person(X).Angle, Person(X).X, Person(X).Y, 0
    Next
    BltRotate BackBuffer, Picture(0), Player.Angle, Player.X, Player.Y, 0
    Do While (GetTickCount - lastTick) < FrameDelay
        DoEvents
    Loop
    PText "F1 = Help", 40, 0, RGB(0, 255, 0)
    PText "F2 = Options", 120, 0, RGB(0, 255, 0)
    If GHelp Then
        PText "Hit'n Run! by GREENie", 50, 12, RGB(0, 255, 0)
        PText Int(1000 / (GetTickCount - lastTick)) & " FPS", 50, 24, RGB(0, 255, 0)
        PText "+ = Add pedestrian", 50, 36, RGB(0, 255, 0)
        PText "Del = Clear dead pedestrians", 50, 48, RGB(0, 255, 0)
        PText "Esc = Escape", 50, 60, RGB(0, 255, 0)
    ElseIf GOps Then
        PText "1-AI Type = " & AIT, 130, 12, RGB(0, 255, 0)
        PText "2-Random Movement", 130, 24, RGB(0, 255, 0)
        PText "= " & AIRnd, 130, 36, RGB(0, 255, 0)
    End If
    Blt
    lastTick = GetTickCount
Loop

End Function
Private Function LoopScreen(GChr As GCharecter)
'Player apears on opposite side of screen on X axis
If GChr.X < 0 Then
    GChr.X = 640
ElseIf GChr.X > 640 Then
    GChr.X = 0
End If
'Player apears on opposite side of screen on Y axis
If GChr.Y < 0 Then
    GChr.Y = 480
ElseIf GChr.Y > 480 Then
    GChr.Y = 0
End If
End Function

Private Function PersonAI(GChr As GCharecter)
Dim Speed As Long
Randomize Timer
If (GChr.X - Player.X) <> 0 And AIT = "Suicidal" Then
    If (GChr.X - Player.X) < 0 Then
        GChr.Angle = Atn((GChr.Y - Player.Y) / (GChr.X - Player.X)) * 180 / 3.14159265358979
    Else
        GChr.Angle = 180 + Atn((GChr.Y - Player.Y) / (GChr.X - Player.X)) * 180 / 3.14159265358979
    End If
ElseIf (GChr.X - Player.X) <> 0 And AIT = "Scared" Then
    If CircularCollision(GChr.X, GChr.Y, Player.X, Player.Y, 200) Then
        If (GChr.X - Player.X) > 0 Then
            GChr.Angle = Atn((GChr.Y - Player.Y) / (GChr.X - Player.X)) * 180 / 3.14159265358979
        Else
            GChr.Angle = 180 + Atn((GChr.Y - Player.Y) / (GChr.X - Player.X)) * 180 / 3.14159265358979
        End If
    End If
End If
If AIRnd Then
    GChr.Angle = GChr.Angle + (Rnd() * 90 - 45)
End If
Speed = 5
GChr.X = GChr.X + Cos(GChr.Angle * 3.14159265358979 / 180) * Speed
GChr.Y = GChr.Y + Sin(GChr.Angle * 3.14159265358979 / 180) * Speed
LoopScreen GChr
End Function

Private Function ClearBodies()
Dim TChrs() As GCharecter
ReDim TChrs(0)
For X = 0 To UBound(Person())
    If Person(X).Alive Then
        If TChrs(0).Alive = True Then
            ReDim Preserve TChrs(UBound(TChrs()) + 1)
        End If
        TChrs(UBound(TChrs())) = Person(X)
    End If
Next X
Person() = TChrs()
End Function
