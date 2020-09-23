Attribute VB_Name = "gradient"
Function GRAD(frm As Form, COLOR As String, Direction As String, c1 As Long, c2 As Long) As Boolean


    '
    '
    ' Easy Gradients In VB - Code by: TjB R.a.d @ www.alcoholabusecen
    '     ter.com
    ' use the code where and how you want just stop by and say "HI!"
    '
    '
    '
    ' Simple Color - Red,Green,Blue
    ' Direction - Fill Direction "Up" or "Down"
    ' c1 - Color Modifier 1
    ' c2 - color Modifier 2
    '
    ' By messing with c1 and c2 you can get some very cool results...
    '     .
    ' here are some of my favorites......
    '
    ' COLOR c1c2
    ' ---------------------------------
    ' RED15761
    ' BLUE301949
    ' BLUE279161
    ' BLUE634208
    ' BLUE187584
    ' BLUE81458
    ' BLUE23544
    ' BLUE543157
    ' green 939655
    ' green 108784
    ' green 19211
    ' green 74106
    '
    GRAD = True
    COLOR = UCase(COLOR)
    Direction = UCase(Direction)
    If Direction = "UP" Then X1 = 255: X2 = 0: X3 = -0.5
    If Direction = "DOWN" Then X1 = 0: X2 = 255: X3 = 0.5
    If Direction <> "UP" And Direction <> "DOWN" Then GRAD = False: GoTo ed
    MDS = frm.DrawStyle
    MDW = frm.DrawWidth
    MSM = frm.ScaleMode
    MSH = frm.ScaleHeight
    frm.DrawStyle = vbInsideSolid
    frm.DrawWidth = 2
    frm.ScaleMode = vbPixels
    frm.ScaleHeight = 256


    For x = X1 To X2 Step X3


        Select Case COLOR
            
            Case "RED"
            frm.Line (0, x)-(frm.Width, x + 1), RGB(255 - x, c1, c2), B
            Case "GREEN"
            frm.Line (0, x)-(frm.Width, x + 1), RGB(c1, 255 - x, c2), B
            Case "BLUE"
            frm.Line (0, x)-(frm.Width, x + 1), RGB(c1, c2, 255 - x), B
            Case Else
            GRAD = False
        End Select

Next x


frm.DrawStyle = MDS
frm.DrawWidth = MDW
frm.ScaleHeight = MSH
frm.ScaleMode = MSM
ed:
End Function



