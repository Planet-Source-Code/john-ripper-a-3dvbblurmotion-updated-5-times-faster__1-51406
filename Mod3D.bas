Attribute VB_Name = "Mod3D"
Option Explicit

Public MainColor   As Byte
Public PaletteName As String
Public FullScreen  As Boolean
Public MeshName    As String

Public pictBuff()  As Byte
Public saBuff      As SAFEARRAY2D
Public bmpBuff     As BitMap

Const XOrg         As Long = 0
Const YOrg         As Long = 0
Const ZOrg         As Long = 260

Const NumSinVal    As Long = 1024

Public NumPoints   As Long
Public NumFaces    As Long

Public XCenter     As Long
Public YCenter     As Long

Public XScreen     As Long
Public YScreen     As Long

Public Type Point3D
    X   As Long
    Y   As Long
    Z   As Long
    Aux As Long
End Type

Public Points()     As Point3D
Public TempPoints() As Point3D

Public Type Face3D
    A  As Long
    B  As Long
    C  As Long
    Z  As Long
    AB As Long
    BC As Long
    CA As Long
End Type
Public Faces() As Face3D

Public CosTable(1025) As Long
Public SinTable(1025) As Long

Public Const PI As Single = 3.141592654

Public Xangle      As Long
Public Yangle      As Long
Public Zangle      As Long

Public SpeedXangle As Long
Public SpeedYangle As Long
Public SpeedZangle As Long

Public rHeight     As Long
Public rWidth      As Long



Public Sub ReadFileMesh(ByVal FileMesh As String, arrPoints() As Point3D, arrFaces() As Face3D, Optional ByVal ReadFaces As Boolean = False)
    
  Dim dataFile    As String
  Dim nF          As Integer
  Dim i           As Integer
  Dim FilePoints  As Long
  Dim FileFaces   As Long
  Dim FlagFaces   As Boolean
  Dim CounterFile As Long
  Dim Pos1        As Long
  Dim Pos2        As Long
  Dim Pos3        As Long
  Dim Pos4        As Long
  Dim Pos5        As Long

    nF = FreeFile
    Open FileMesh For Input As #nF
    
    'read "header"
    For i = 1 To 8
        Line Input #nF, dataFile
    Next i

    Line Input #nF, dataFile
    Pos1 = InStr(1, dataFile, "=")
    FilePoints = Mid(dataFile, Pos1 + 1)
    NumPoints = FilePoints
    ReDim arrPoints(FilePoints)
    ReDim TempPoints(FilePoints)
    
    Line Input #nF, dataFile
    FlagFaces = True
    If (InStr(1, dataFile, "Not Available") <> 0) Then
        FlagFaces = False
      Else
        Pos1 = InStr(1, dataFile, "=")
        FileFaces = Mid(dataFile, Pos1 + 1)
        NumFaces = FileFaces
    End If
    
    Line Input #nF, dataFile        '""
    Line Input #nF, dataFile        '"--------------------------POINTS-------------------------"
    
    CounterFile = 0
    Do Until CounterFile = FilePoints + 1
        Line Input #nF, dataFile 'X!Y@Z format
        Pos1 = InStr(1, dataFile, "!")
        Pos2 = InStr(1, dataFile, "@")
        Pos3 = InStr(1, dataFile, "*")
        
        arrPoints(CounterFile).X = Mid(dataFile, 1, Pos1 - 1)
        arrPoints(CounterFile).Y = Mid(dataFile, Pos1 + 1, Pos2 - Pos1 - 1)
        If (Pos3 = 0) Then
            arrPoints(CounterFile).Z = Mid(dataFile, Pos2 + 1)
          Else
            arrPoints(CounterFile).Z = Mid(dataFile, Pos2 + 1, Pos3 - Pos2 - 1)
            arrPoints(CounterFile).Aux = Mid(dataFile, Pos3 + 1)
        End If
        CounterFile = CounterFile + 1
    Loop
    
    If (ReadFaces And FlagFaces) Then
        ReDim arrFaces(FileFaces)
        
        Line Input #nF, dataFile    '--------------------------FACES--------------------------
    
        CounterFile = 0
        Do Until CounterFile = FileFaces + 1
            Line Input #nF, dataFile    'A!B@C format
            Pos1 = InStr(1, dataFile, "!")
            Pos2 = InStr(1, dataFile, "@")
            Pos3 = InStr(1, dataFile, "*")
            Pos4 = InStr(1, dataFile, "%")
            Pos5 = InStr(1, dataFile, "(")
            
            arrFaces(CounterFile).A = Mid(dataFile, 1, Pos1 - 1)
            arrFaces(CounterFile).B = Mid(dataFile, Pos1 + 1, Pos2 - Pos1 - 1)
            arrFaces(CounterFile).C = Mid(dataFile, Pos2 + 1, Pos3 - Pos2 - 1)
            arrFaces(CounterFile).Z = 0
            arrFaces(CounterFile).AB = Mid(dataFile, Pos3 + 1, Pos4 - Pos3 - 1)
            arrFaces(CounterFile).BC = Mid(dataFile, Pos4 + 1, Pos5 - Pos4 - 1)
            arrFaces(CounterFile).CA = Mid(dataFile, Pos5 + 1)
            CounterFile = CounterFile + 1
        Loop
    End If
    
    Close #nF
End Sub



Public Sub MakeCosTable()

  Dim CntVal As Long
  Dim CntAng As Single
  Dim IncDeg As Single
  
    IncDeg = 2 * PI / NumSinVal
    CntAng = IncDeg
    CntVal = 0
    
    Do Until CntVal > 1024
        CosTable(CntVal) = CInt((255 * Cos(CntAng)))
        CntAng = CntAng + IncDeg
        CntVal = CntVal + 1
    Loop
End Sub

Public Sub MakeSinTable()

  Dim CntVal As Long
  Dim CntAng As Single
  Dim IncDeg As Single

    IncDeg = 2 * PI / NumSinVal
    CntAng = IncDeg
    CntVal = 0
    
    Do Until CntVal > 1024
        SinTable(CntVal) = CInt((255 * Sin(CntAng)))
        CntAng = CntAng + IncDeg
        CntVal = CntVal + 1
    Loop
End Sub

Public Sub Calc3DRotations(ByVal SinX As Long, ByVal CosX As Long, ByVal SinY As Long, ByVal CosY As Long, ByVal SinZ As Long, ByVal CosZ As Long, _
                           OrgPoints() As Point3D, DesPoints() As Point3D, _
                           ByVal NumPoints As Long)

  Dim x1  As Long
  Dim y1  As Long
  Dim z1  As Long
  Dim cnt As Long

    For cnt = 0 To NumPoints
        
'     X1 := (cos(YAngle) * X  - sin(YAngle) * Z)
        x1 = (CosY * OrgPoints(cnt).X - SinY * OrgPoints(cnt).Z) \ 256
'     Z1 := (sin(YAngle) * X  + cos(YAngle) * Z)
        z1 = (SinY * OrgPoints(cnt).X + CosY * OrgPoints(cnt).Z) \ 256
'     X  := (cos(ZAngle) * X1 + sin(ZAngle) * Y)
        DesPoints(cnt).X = (CosZ * x1 + SinZ * OrgPoints(cnt).Y) \ 256
'     Y1 := (cos(ZAngle) * Y  - sin(ZAngle) * X1)
        y1 = (CosZ * OrgPoints(cnt).Y - SinZ * x1) \ 256
'     Z  := (cos(XAngle) * Z1 - sin(XAngle) * Y1)
        DesPoints(cnt).Z = (CosX * z1 - SinX * y1) \ 256
'     Y  := (sin(XAngle)) * Z1 + cos(XAngle) * Y1)
        DesPoints(cnt).Y = (SinX * z1 + CosX * y1) \ 256
      
        DesPoints(cnt).Aux = OrgPoints(cnt).Aux
   Next cnt
End Sub

Public Sub Proyect3D(ByVal XScreen As Long, ByVal YScreen As Long, ByVal NumPoints As Long, _
                     OrgPoints() As Point3D, DesPoints() As Point3D)

  Dim cnt As Long

    For cnt = 0 To NumPoints
        
        With OrgPoints(cnt)
            DesPoints(cnt).X = XScreen + ((XOrg * .Z - .X * ZOrg) / (.Z - ZOrg))
            DesPoints(cnt).Y = YScreen + ((YOrg * .Z - .Y * ZOrg) / (.Z - ZOrg))
        End With
    Next cnt
End Sub


Public Sub QuickSortZFaces(ByVal NumPoints As Long, Points2qS() As Point3D, Faces2qS() As Face3D)
  
  Dim cnt As Long
    
    For cnt = 0 To NumPoints
     
        With Faces2qS(cnt)
            .Z = (Points2qS(.A).Z + Points2qS(.B).Z + Points2qS(.C).Z) \ 3
        End With
    Next cnt
    
    Call QuickSortFaces(Faces2qS, 0, NumPoints)
End Sub

Private Sub QuickSortFaces(vntArr() As Face3D, ByVal lngLeft As Long, ByVal lngRight As Long)

  Dim i          As Long
  Dim j          As Long
  Dim lngMid     As Long
  Dim vntTestVal As Variant
  Dim vntTemp    As Face3D
    
    If (lngLeft < lngRight) Then
        
        lngMid = (lngLeft + lngRight) \ 2
        vntTestVal = vntArr(lngMid).Z
        i = lngLeft
        j = lngRight
        
        Do
            Do While vntArr(i).Z < vntTestVal
                i = i + 1
            Loop
            Do While vntArr(j).Z > vntTestVal
                j = j - 1
            Loop
            If (i <= j) Then
                vntTemp = vntArr(j)
                vntArr(j) = vntArr(i)
                vntArr(i) = vntTemp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j

        If (j <= lngMid) Then
            Call QuickSortFaces(vntArr, lngLeft, j)
            Call QuickSortFaces(vntArr, i, lngRight)
          Else
            Call QuickSortFaces(vntArr, i, lngRight)
            Call QuickSortFaces(vntArr, lngLeft, j)
        End If
    End If
End Sub

Public Function FaceVisible(x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long) As Boolean
'Simple escalar product:
'Return TRUE if face is visible
'if FaceVisible=False NOT PAINT anything. Increase speed!

  Dim A As Long
  Dim B As Long

    A = (x2 - x1) * (y3 - y1)
    B = (x3 - x1) * (y2 - y1)
    
    FaceVisible = (A - B >= 0)
End Function

Public Sub Render()

  Dim i   As Long
  Dim xx1 As Long
  Dim xx2 As Long
  Dim yy1 As Long
  Dim yy2 As Long
  Dim px  As Long
  Dim py  As Long

    CopyMemory ByVal VarPtrArray(pictBuff), VarPtr(saBuff), 4
    
    Call Calc3DRotations(SinTable(Xangle), CosTable(Xangle), SinTable(Yangle), CosTable(Yangle), SinTable(Zangle), CosTable(Zangle), Points, TempPoints, UBound(Points))
    Call Proyect3D(XCenter, YCenter, UBound(Points), TempPoints, TempPoints)
    Call QuickSortZFaces(UBound(Faces), TempPoints, Faces)

    For i = 0 To UBound(Faces)
    
        If (FaceVisible(TempPoints(Faces(i).A).X, TempPoints(Faces(i).A).Y, TempPoints(Faces(i).B).X, TempPoints(Faces(i).B).Y, TempPoints(Faces(i).C).X, TempPoints(Faces(i).C).Y)) Then

            If (Faces(i).AB = 1) Then
                xx1 = TempPoints(Faces(i).A).X
                yy1 = TempPoints(Faces(i).A).Y

                xx2 = TempPoints(Faces(i).B).X
                yy2 = TempPoints(Faces(i).B).Y
                DrawLine pictBuff, xx1, yy1, xx2, yy2, MainColor
            End If
            
            If (Faces(i).BC = 1) Then
                xx1 = TempPoints(Faces(i).B).X
                yy1 = TempPoints(Faces(i).B).Y

                xx2 = TempPoints(Faces(i).C).X
                yy2 = TempPoints(Faces(i).C).Y
                DrawLine pictBuff, xx1, yy1, xx2, yy2, MainColor
            End If
            
            If (Faces(i).CA = 1) Then
                xx1 = TempPoints(Faces(i).C).X
                yy1 = TempPoints(Faces(i).C).Y

                xx2 = TempPoints(Faces(i).A).X
                yy2 = TempPoints(Faces(i).A).Y
                DrawLine pictBuff, xx1, yy1, xx2, yy2, MainColor
            End If
        End If
    Next i

    Xangle = Xangle + SpeedXangle
    If (Xangle > 1024) Then Xangle = Xangle - 1024 Else If (Xangle < 0) Then Xangle = 0
    
    Yangle = Yangle + SpeedYangle
    If (Yangle > 1024) Then Yangle = Yangle - 1024 Else If (Yangle < 0) Then Yangle = 0
    
    Zangle = Zangle + SpeedZangle
    If (Zangle > 1024) Then Zangle = Zangle - 1024 Else If (Zangle < 0) Then Zangle = 0
    
    'Blur begins here:
    For py = 1 To 238
        For px = 1 To 318
            pictBuff(px, py) = (CLng(pictBuff(px + 1, py)) + pictBuff(px - 1, py) + pictBuff(px, py + 1) + pictBuff(px, py - 1)) \ 4
        Next px
    Next py
    'and ends, here. hehehehe
    
    CopyMemory ByVal VarPtrArray(pictBuff), 0&, 4
    frmMain.PicBuff.Refresh
End Sub

'Public Sub PintaLinea(BitMap() As Byte, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal color As Long)
'
'  Dim X As Long
'  Dim Y As Long
'
'    If Abs(y2 - y1) <= Abs(x2 - x1) Then
'
'        'Por cada X calcular Y
'        For X = x1 To x2
'            Y = y1 + (y2 - y1) * (X - x1) / (x2 - x1)
'            BitMap(X, Y) = color
'        Next
'      Else
'        'Por cara Y calcular X
'        For Y = y1 To y2
'            X = x1 + (x2 - x1) * (Y - y1) / (y2 - y1)
'             BitMap(X, Y) = color
'        Next
'    End If
' End Sub

Public Sub DrawLine(BitMap() As Byte, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal color As Byte)
'// Bresenham's algorithm
'// http://homepage.smc.edu/kennedy_john/BRESENL.PDF

  Dim xCur       As Long
  Dim yCur       As Long
  Dim xInc       As Long
  Dim yInc       As Long
  Dim dx         As Long
  Dim dy         As Long
  Dim TwoDx      As Long
  Dim TwoDy      As Long
  Dim TwoDxAcErr As Long
  Dim TwoDyAcErr As Long

    dx = x2 - x1
    dy = y2 - y1
    TwoDx = dx + dx
    TwoDy = dy + dy
    xCur = x1
    yCur = y1
    xInc = 1
    yInc = 1

    If (dx < 0) Then
        xInc = -1
        dx = -dx
        TwoDx = -TwoDx
    End If
    If (dy < 0) Then
        yInc = -1
        dy = -dy
        TwoDy = -TwoDy
    End If

    If (x1 < 1 Or x1 > 318 Or y1 < 1 Or y1 > 238) Then Exit Sub '***
    BitMap(x1, y1) = color

    If (dx <> 0 Or dy <> 0) Then
        If (dy <= dx) Then
            
            TwoDxAcErr = 0
            Do
                xCur = xCur + xInc
                TwoDxAcErr = TwoDxAcErr + TwoDy
                If (TwoDxAcErr > dx) Then
                    yCur = yCur + yInc
                    TwoDxAcErr = TwoDxAcErr - TwoDx
                End If
                If (xCur < 1 Or xCur > 318 Or yCur < 1 Or yCur > 238) Then Exit Sub '***
                BitMap(xCur, yCur) = color
            Loop Until (xCur = x2)
          
          Else
            TwoDyAcErr = 0
            Do
                yCur = yCur + yInc
                TwoDyAcErr = TwoDyAcErr + TwoDx
                If (TwoDyAcErr > dy) Then
                    xCur = xCur + xInc
                    TwoDyAcErr = TwoDyAcErr - TwoDy
                End If
                If (xCur < 1 Or xCur > 318 Or yCur < 1 Or yCur > 238) Then Exit Sub '***
                BitMap(xCur, yCur) = color
            Loop Until (yCur = y2)
        End If
    End If
End Sub
