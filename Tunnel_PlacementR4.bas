Attribute VB_Name = "Tunnel_PlacementR4"
'Topic : Placement of Tunnel Segment Program
'By : Suben Mukem
'Update : 11.07.2023

 Const Pi As Single = 3.141592654

'Convert Degrees to Radian.

Private Function DegtoRad(d)

    DegtoRad = d * (Pi / 180)

 End Function

'Convert Radian to Degrees.

 Private Function RadtoDeg(r)

    RadtoDeg = r * (180 / Pi)

 End Function
 
 'Compute Distance and Azimuth from 2 Points.

Private Function PvDirecDistAz(EStart, NStart, EEnd, NEnd, DA)

    dE = EEnd - EStart: dN = NEnd - NStart
    Distance = Sqr(dE ^ 2 + dN ^ 2)
    
    If dN <> 0 Then Q = RadtoDeg(Atn(dE / dN))
      If dN = 0 Then
        If dE > 0 Then
          Azi = 90
        ElseIf dE < 0 Then
          Azi = 270
        Else
          Azi = False
      End If
      
    ElseIf dN > 0 Then
      If dE > 0 Then
          Azi = Q
      ElseIf dE < 0 Then
          Azi = 360 + Q
      End If
      
    ElseIf dN < 0 Then
          Azi = 180 + Q
    End If
    
    Select Case UCase$(DA)
      Case "D"
          PvDirecDistAz = Distance
      Case "A"
          PvDirecDistAz = Azi
    End Select

End Function 'DirecDistAz
 
'Compute Northing and Easting by Local Coordinate (Y, X) , Coordinate of Center and Azimuth.

Private Function PvCoorYXtoNE(ECL, NCL, AZCL, Y, X, EN)

    Ei = ECL + Y * Sin(DegtoRad(AZCL)) + X * Sin(DegtoRad(90 + AZCL))
    Ni = NCL + Y * Cos(DegtoRad(AZCL)) + X * Cos(DegtoRad(90 + AZCL))
    
    Select Case UCase$(EN)
     Case "E"
             PvCoorYXtoNE = Ei
     Case "N"
             PvCoorYXtoNE = Ni
  End Select
  
End Function 'Coordinate Y,X to N, E

'Compute Zenith angle from 2 Points.

Sub ZenithAng()
    'Count Total Tunnel Segment
    ActiveSheet.Select
    Range("A4").Select
    num = Range(Selection, Selection.End(xlDown)).Count
    
    Dim RN() As Variant
    Dim CH() As Variant
    Dim EL() As Variant
    
    ReDim RN(num)
    ReDim CH(num)
    ReDim EL(num)
    
    For i = 0 To num - 1
        ActiveSheet.Select
        Range("A4").Select
        RN(i) = ActiveCell.Offset(i, 1)
        CH(i) = ActiveCell.Offset(i, 2)
        EL(i) = ActiveCell.Offset(i, 5)
    Next
    
    'First Row Zenith Angle was conputed and printed
    ZA0 = PvDirecDistAz(CH(0), EL(0), CH(1), EL(1), "A")
    ActiveCell.Offset(0, 7).Value = ZA0
    
    'Zenith Angle was computed and printed
    For i = 1 To num - 1
        ActiveSheet.Select
        Range("A4").Select
        ZAi = PvDirecDistAz(CH(i - 1), EL(i - 1), CH(i), EL(i), "A")
        ActiveCell.Offset(i, 7).Value = ZAi
    Next
    
    MsgBox "Zenith angle was successfully computed!"
        
End Sub
 
'Placement of Tunnel Segment Plan (PTS) by Insert Attribute Block

Sub PTSPlanByInsertBlocks2D()
    
    Dim acadApp As AcadApplication
    Dim acadDoc As AcadDocument
    Dim acadBlock As AcadBlockReference
    Dim acadArr As Variant
    Dim InsertPoint(0 To 2) As Double
    Dim RingName As String
    Dim BlockName As String
    Dim BlockScaleX As Double
    Dim BlockScaleY As Double
    Dim BlockScaleZ As Double
    Dim RotationAngle As Double
    
    'Count Total Tunnel Segment
    ActiveSheet.Select
    Range("A4").Select
    num = Range(Selection, Selection.End(xlDown)).Count
    
    MsgBox "Total Tunnel Segment :" & " " & num
    
    '-------------------------------------Set Working with AutoCAD-------------------------------------'
    
    'Check if AutoCAD application is open. If is not opened create a new instance and make it visible.
    On Error Resume Next
        Set acadApp = GetObject(, "Autocad.application")
    
    If acadApp Is Nothing Then
        Set acadApp = CreateObject("autocad.application")
        acadApp.Visible = True
    End If
    
    'Check (again) if there is an AutoCAD object.
    On Error Resume Next
        Set acadDoc = acadApp.ActiveDocument
    On Error GoTo 0
    
    'If there is no active drawing create a new one.
    If acadDoc Is Nothing Then
        Set acadDoc = acadApp.Documents.Add
    End If
    
    '-------------------------------------Insert Block and Attribute (Text)-------------------------------------'
    
    For i = 0 To num - 1
        ActiveSheet.Select
        Range("A4").Select
        
        'Set the Ring name, Block name and InsertPoint(E,N,Z).
        RingName = ActiveCell.Offset(i, 1)
        BlockName = ActiveCell.Offset(i, 8)
        InsertPoint(0) = ActiveCell.Offset(i, 3) 'Easting
        InsertPoint(1) = ActiveCell.Offset(i, 4) 'Northing
        InsertPoint(2) = 0 'Z
        
        'Set initialize the optional parameters.
        BlockScaleX = 1
        BlockScaleY = 1
        BlockScaleZ = 1
        RotationAngle = 90 - ActiveCell.Offset(i, 6) 'Azimuth
    
        'Inset Block and Attributes.
        Set acadBlock = acadDoc.ModelSpace.Insertblock(InsertPoint, BlockName, BlockScaleX, BlockScaleY, BlockScaleZ, DegtoRad(RotationAngle))
        acadArr = acadBlock.GetAttributes
        acadArr(0).TextString = RingName
    Next
    
    'Zoom in to the drawing area.
    acadApp.ZoomExtents
    
    'Release the objects.
    Set acadBlock = Nothing
    Set acadDoc = Nothing
    Set acadApp = Nothing
        
    MsgBox "Placement of Tunnel Segment Plan (2D) was successfully created!"
    
End Sub


'Placement of Tunnel Segment Plan(PTS) by Polyline

Sub PTSPlanByPolyline2D()
    
    Dim acadApp As AcadApplication
    Dim acadDoc As AcadDocument
    Dim acadLine As acadLine
    Dim acadPol As AcadLWPolyline
    Dim LayerObj As AcadLayer
    Dim layerName(0 To 1) As String
    Dim layerColor(0 To 1) As Integer
    Dim layerLineType(0 To 1) As String
    Dim layerLineweight(0 To 1) As Integer
    Dim acadText As acadText
    Dim TextStyle As AcadTextStyle
    Dim InsertionPoint(0 To 2) As Double
    Dim ZeroPoint(0 To 2) As Double 'Text origin coodinate
    
    '-------------------------------------Index Value and Count number-------------------------------------'
    
    'Daimiter of Tunnel Segment to compute coordinate of tunnel offset
    ActiveSheet.Select
    Range("J5").Select
    TunnelDai = ActiveCell.Offset(0, 1)
        
    'Layers Properties
    ActiveSheet.Select
    Range("J9").Select
    LayerNum = Range(Selection, Selection.End(xlDown)).Count
    
    For i = 0 To LayerNum - 1
        ActiveSheet.Select
        Range("J9").Select
        layerName(i) = ActiveCell.Offset(i, 0)
        layerColor(i) = ActiveCell.Offset(i, 1)
        layerLineType(i) = ActiveCell.Offset(i, 2)
        layerLineweight(i) = ActiveCell.Offset(i, 3) * 100
        'Debug.Print layerName(i), layerColor(i), layerLineType(i), layerLineweight(i)
    Next
         
    'Text Properties
    ActiveSheet.Select
    Range("J14").Select
    TextFont = ActiveCell.Offset(0, 0)
    TextAlign = ActiveCell.Offset(0, 1)
    TextHeight = ActiveCell.Offset(0, 2)
    TextWFactor = ActiveCell.Offset(0, 3)
         
    'Count Total Tunnel Segment, Tunnel Segment Data
    ActiveSheet.Select
    Range("A4").Select
    num = Range(Selection, Selection.End(xlDown)).Count
    
    MsgBox "Total Tunnel Segment :" & " " & num - 1
    
    'Index coodinate of center (ECL, NCL) and compute coordinate of left (ELT, NLT), coordinate of right (ERT, NRT)
    Dim RingName As Variant
    Dim ECL As Variant
    Dim NCL As Variant
    Dim AZCL As Variant
    Dim ELT As Variant
    Dim NLT As Variant
    Dim ERT As Variant
    Dim NRT As Variant
    
    ReDim RingName(num)
    ReDim ECL(num)
    ReDim NCL(num)
    ReDim AZCL(num)
    ReDim ELT(num)
    ReDim NLT(num)
    ReDim ERT(num)
    ReDim NRT(num)
    
    For i = 0 To num - 1
        ActiveSheet.Select
        Range("A4").Select
        RingName(i) = ActiveCell.Offset(i, 1)
        ECL(i) = ActiveCell.Offset(i, 3)
        NCL(i) = ActiveCell.Offset(i, 4)
        AZCL(i) = ActiveCell.Offset(i, 6)
        ELT(i) = PvCoorYXtoNE(ECL(i), NCL(i), AZCL(i), 0, (TunnelDai / 2) * -1, "E")
        NLT(i) = PvCoorYXtoNE(ECL(i), NCL(i), AZCL(i), 0, (TunnelDai / 2) * -1, "N")
        ERT(i) = PvCoorYXtoNE(ECL(i), NCL(i), AZCL(i), 0, (TunnelDai / 2), "E")
        NRT(i) = PvCoorYXtoNE(ECL(i), NCL(i), AZCL(i), 0, (TunnelDai / 2), "N")
    Next
    'Debug.Print RingName(0), ECL(0), NCL(0), AZCL(0), ELT(0), NLT(0), ERT(0), NRT(0)
    
    '-------------------------------------Set Working with AutoCAD-------------------------------------'
    
    'Check if AutoCAD is open.
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    On Error GoTo 0

    'If AutoCAD is not opened create a new instance and make it visible.
    If acadApp Is Nothing Then
        Set acadApp = New AcadApplication
        acadApp.Visible = True
    End If

    'Check if there is an active drawing.
    On Error Resume Next
    Set acadDoc = acadApp.ActiveDocument
    On Error GoTo 0

    'No active drawing found. Create a new one.
    If acadDoc Is Nothing Then
        Set acadDoc = acadApp.Documents.Add
        acadApp.Visible = True
    End If
    
    '-------------------------------------Line and PolyLine-------------------------------------'
    
    'Set Layers and Layer Properties
    For i = 0 To LayerNum - 1
        On Error Resume Next
            Set LayerObj = acadDoc.Layers.Add(layerName(i))
            LayerObj.color = layerColor(i)
            LayerObj.LineType = layerLineType(i)
            LayerObj.LineWeight = layerLineweight(i)
        On Error GoTo 0
    Next
    
    'Select Layer for Line and PolyLine
    Set LayerObj = acadDoc.Layers(layerName(0))
    acadDoc.ActiveLayer = LayerObj
    
    'Sample Line of Tunnel Segment
    Dim SampleLineBP(0 To 2) As Double
    Dim SampleLineEP(0 To 2) As Double

    For i = 0 To num - 1
        'Set Sample Line of Tunnel Segment by 2 Points
        SampleLineBP(0) = ELT(i)
        SampleLineBP(1) = NLT(i)
        SampleLineEP(0) = ERT(i)
        SampleLineEP(1) = NRT(i)
        
        'Draw the Line (Sample Line) to AutoCAD
        If acadDoc.ActiveSpace = acModelSpace Then
            Set acadLine = acadDoc.ModelSpace.AddLine(SampleLineBP, SampleLineEP)
        Else
            Set acadLine = acadDoc.PaperSpace.AddLine(SampleLineBP, SampleLineEP)
        End If
    Next
    
    'Set LT Polyline of Tunnel Segment
    'Joint 2 arrays (ELT, NLT) to 1 array
    Dim LTPLine() As Double
    ReDim LTPLine((num - 1) * 2 + 1)
    
    For i = 0 To num - 1
        LTPLine(i * 2) = ELT(i)
        LTPLine(i * 2 + 1) = NLT(i)
    Next
    
    'Draw the polyline (LT) to AutoCAD
    If acadDoc.ActiveSpace = acModelSpace Then
        Set acadPol = acadDoc.ModelSpace.AddLightWeightPolyline(LTPLine)
    Else
        Set acadPol = acadDoc.PaperSpace.AddLightWeightPolyline(LTPLine)
    End If
    
    'Set RT Polyline of Tunnel Segment
    Dim RTPLine() As Double
    ReDim RTPLine((num - 1) * 2 + 1)
    
    For i = 0 To num - 1
        'Joint 2 arrays (ERT, NRT) to 1 array
        RTPLine(i * 2) = ERT(i)
        RTPLine(i * 2 + 1) = NRT(i)
    Next
    
    'Draw the polyline (RT) to AutoCAD
    If acadDoc.ActiveSpace = acModelSpace Then
        Set acadPol = acadDoc.ModelSpace.AddLightWeightPolyline(RTPLine)
    Else
        Set acadPol = acadDoc.PaperSpace.AddLightWeightPolyline(RTPLine)
    End If
    
    '-------------------------------------Text-------------------------------------'
    
    'Select Layer for Line and PolyLine
    Set LayerObj = acadDoc.Layers(layerName(1))
    acadDoc.ActiveLayer = LayerObj
    
    'Text properties
    Set TextStyle = acadDoc.TextStyles.Add(TextFont)
    
    Select Case TextFont
        Case "Angsana": TextFontFile = "angsana.shx"
        Case "Cordia": TextFontFile = "cordia.shx"
        Case "Romans": TextFontFile = "romans.shx"
        Case "Simplx": TextFontFile = "simplx.shx"
    End Select
    
    TextStyle.fontFile = TextFont
    acadDoc.ActiveTextStyle = TextStyle
    
    'Text Alignment Case
    Select Case TextAlign
        Case "Top Left": TextAlign = acAlignmentTopLeft
        Case "Top Center": TextAlign = acAlignmentTopCenter
        Case "Top Right": TextAlign = acAlignmentTopRight
        Case "Middle Left": TextAlign = acAlignmentMiddleLeft
        Case "Middle Center": TextAlign = acAlignmentMiddleCenter
        Case "Middle Right": TextAlign = acAlignmentMiddleRight
        Case "Bottom Left": TextAlign = acAlignmentBottomLeft
        Case "Bottom Center": TextAlign = acAlignmentBottomCenter
        Case "Bottom Right": TextAlign = acAlignmentBottomRight
    End Select
    
    For i = 1 To num - 1
        'The point at the beginning of 3 axis.
        ZeroPoint(0) = 0
        ZeroPoint(1) = 0
        ZeroPoint(2) = 0
        
        'Set the insertion point and Rotation angle
        InsertionPoint(0) = (ECL(i) + ECL(i - 1)) / 2
        InsertionPoint(1) = (NCL(i) + NCL(i - 1)) / 2
        InsertionPoint(2) = 0 'Z
        RotationAngle = AZCL(i)
        
        'Add the text to AutoCAD.
        Set acadText = acadDoc.ModelSpace.AddText(RingName(i), InsertionPoint, TextHeight)
        acadText.Alignment = TextAlign
        acadText.Move ZeroPoint, InsertionPoint
        acadText.Rotate InsertionPoint, DegtoRad(RotationAngle) * -1
        acadText.ScaleFactor = TextWFactor
        
    Next
    
    acadApp.ZoomExtents
    
    'Release the objects.
    Set acadDoc = Nothing
    Set acadApp = Nothing
    
    MsgBox "Placement of Tunnel Segment Plan (2D) was successfully created!"

End Sub


'Placement of Tunnel Segment Profile (PTS) by Insert Attribute Block

Sub PTSProfileByInsertBlocks2D()
    
    Dim acadApp As AcadApplication
    Dim acadDoc As AcadDocument
    Dim acadBlock As AcadBlockReference
    Dim acadArr As Variant
    Dim InsertPoint(0 To 2) As Double
    Dim UcsOrigin(0 To 1) As Double
    Dim RingName As String
    Dim BlockName As String
    Dim BlockScaleX As Double
    Dim BlockScaleY As Double
    Dim BlockScaleZ As Double
    Dim RotationAngle As Double
    
    'Count Total Tunnel Segment
    ActiveSheet.Select
    Range("A4").Select
    num = Range(Selection, Selection.End(xlDown)).Count
    
    MsgBox "Total Tunnel Segment :" & " " & num
    
    '-------------------------------------Set Working with AutoCAD-------------------------------------'
    
    'Check if AutoCAD application is open. If is not opened create a new instance and make it visible.
    On Error Resume Next
        Set acadApp = GetObject(, "Autocad.application")
    
    If acadApp Is Nothing Then
        Set acadApp = CreateObject("autocad.application")
        acadApp.Visible = True
    End If
    
    'Check (again) if there is an AutoCAD object.
    On Error Resume Next
        Set acadDoc = acadApp.ActiveDocument
    On Error GoTo 0
    
    'If there is no active drawing create a new one.
    If acadDoc Is Nothing Then
        Set acadDoc = acadApp.Documents.Add
    End If
    
    '-------------------------------------Insert Block and Attribute (Text)-------------------------------------'
    
    'UCS Origin Coordinate for Plotting Profile to AutoCAD
    ActiveSheet.Select
    Range("K6").Select
    UcsOrigin(0) = ActiveCell.Offset(0, 1) 'Easting
    UcsOrigin(1) = ActiveCell.Offset(0, 2) 'Northing
    
    For i = 0 To num - 1
        ActiveSheet.Select
        Range("A4").Select
        
        'Set the Ring name, Block name and InsertPoint(CH,EL,Z).
        RingName = ActiveCell.Offset(i, 1)
        BlockName = ActiveCell.Offset(i, 8)
        InsertPoint(0) = ActiveCell.Offset(i, 2) + UcsOrigin(0) 'Chainage
        InsertPoint(1) = ActiveCell.Offset(i, 5) + UcsOrigin(1) 'Elevation
        InsertPoint(2) = 0 'Z
        
        'Set initialize the optional parameters.
        BlockScaleX = 1
        BlockScaleY = 1
        BlockScaleZ = 1
        RotationAngle = 90 - ActiveCell.Offset(i, 7) 'Zenith Angle
    
        'Inset Block and Attributes.
        Set acadBlock = acadDoc.ModelSpace.Insertblock(InsertPoint, BlockName, BlockScaleX, BlockScaleY, BlockScaleZ, DegtoRad(RotationAngle))
        acadArr = acadBlock.GetAttributes
        acadArr(0).TextString = RingName
    Next
    
    'Zoom in to the drawing area.
    acadApp.ZoomExtents
    
    'Release the objects.
    Set acadBlock = Nothing
    Set acadDoc = Nothing
    Set acadApp = Nothing
        
    MsgBox "Placement of Tunnel Segment Profile (2D) was successfully created!"
    
End Sub


'Placement of Tunnel Segment Profile(PTS) by Polyline

Sub PTSProfileByPolyline2D()
    
    Dim acadApp As AcadApplication
    Dim acadDoc As AcadDocument
    Dim acadLine As acadLine
    Dim acadPol As AcadLWPolyline
    Dim LayerObj As AcadLayer
    Dim layerName(0 To 1) As String
    Dim layerColor(0 To 1) As Integer
    Dim layerLineType(0 To 1) As String
    Dim layerLineweight(0 To 1) As Integer
    Dim acadText As acadText
    Dim TextStyle As AcadTextStyle
    Dim InsertionPoint(0 To 2) As Double
    Dim ZeroPoint(0 To 2) As Double 'Text origin coodinate
    Dim UcsOrigin(0 To 1) As Double
    
    '-------------------------------------Index Value and Count number-------------------------------------'
    
    'Daimiter of Tunnel Segment to compute coordinate of tunnel offset
    ActiveSheet.Select
    Range("J5").Select
    TunnelDai = ActiveCell.Offset(0, 1)
        
    'Layers Properties
    ActiveSheet.Select
    Range("J9").Select
    LayerNum = Range(Selection, Selection.End(xlDown)).Count
    
    For i = 0 To LayerNum - 1
        ActiveSheet.Select
        Range("J9").Select
        layerName(i) = ActiveCell.Offset(i, 0)
        layerColor(i) = ActiveCell.Offset(i, 1)
        layerLineType(i) = ActiveCell.Offset(i, 2)
        layerLineweight(i) = ActiveCell.Offset(i, 3) * 100
        'Debug.Print layerName(i), layerColor(i), layerLineType(i), layerLineweight(i)
    Next
         
    'Text Properties
    ActiveSheet.Select
    Range("J14").Select
    TextFont = ActiveCell.Offset(0, 0)
    TextAlign = ActiveCell.Offset(0, 1)
    TextHeight = ActiveCell.Offset(0, 2)
    TextWFactor = ActiveCell.Offset(0, 3)
    
    'UCS Origin Coordinate for Plotting Profile to AutoCAD
    ActiveSheet.Select
    Range("J18").Select
    UcsOrigin(0) = ActiveCell.Offset(0, 1) 'Easting
    UcsOrigin(1) = ActiveCell.Offset(0, 2) 'Northing
         
    'Count Total Tunnel Segment, Tunnel Segment Data
    ActiveSheet.Select
    Range("A4").Select
    num = Range(Selection, Selection.End(xlDown)).Count
    
    MsgBox "Total Tunnel Segment :" & " " & num - 1
    
    'Index coodinate of center (CHCL, ELCL) and compute coordinate of Up (CHUP, ELUP), coordinate of Down (CHDN, ELDN)
    Dim RingName As Variant
    Dim CHCL As Variant
    Dim ELCL As Variant
    Dim ZACL As Variant
    Dim CHUP As Variant
    Dim ELUP As Variant
    Dim CHDN As Variant
    Dim ELDN As Variant
    
    ReDim RingName(num)
    ReDim CHCL(num)
    ReDim ELCL(num)
    ReDim ZACL(num)
    ReDim CHUP(num)
    ReDim ELUP(num)
    ReDim CHDN(num)
    ReDim ELDN(num)
    
    For i = 0 To num - 1
        ActiveSheet.Select
        Range("A4").Select
        RingName(i) = ActiveCell.Offset(i, 1)
        CHCL(i) = ActiveCell.Offset(i, 2) + UcsOrigin(0) 'Chainage of Center
        ELCL(i) = ActiveCell.Offset(i, 5) + UcsOrigin(1) 'Elevation of Center
        ZACL(i) = ActiveCell.Offset(i, 7) 'Zenith Angle of Center
        CHUP(i) = PvCoorYXtoNE(CHCL(i), ELCL(i), ZACL(i), 0, (TunnelDai / 2) * -1, "E") 'Chainage of Up
        ELUP(i) = PvCoorYXtoNE(CHCL(i), ELCL(i), ZACL(i), 0, (TunnelDai / 2) * -1, "N") 'Elevation of Up
        CHDN(i) = PvCoorYXtoNE(CHCL(i), ELCL(i), ZACL(i), 0, (TunnelDai / 2), "E") 'Chainage of Down
        ELDN(i) = PvCoorYXtoNE(CHCL(i), ELCL(i), ZACL(i), 0, (TunnelDai / 2), "N") 'Elevation of Down
    Next
    'Debug.Print RingName(0), CHCL(0), ELCL(0), ZACL(0), CHUP(0), ELUP(0), CHDN(0), ELDN(0)
    
    '-------------------------------------Set Working with AutoCAD-------------------------------------'
    
    'Check if AutoCAD is open.
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    On Error GoTo 0

    'If AutoCAD is not opened create a new instance and make it visible.
    If acadApp Is Nothing Then
        Set acadApp = New AcadApplication
        acadApp.Visible = True
    End If

    'Check if there is an active drawing.
    On Error Resume Next
    Set acadDoc = acadApp.ActiveDocument
    On Error GoTo 0

    'No active drawing found. Create a new one.
    If acadDoc Is Nothing Then
        Set acadDoc = acadApp.Documents.Add
        acadApp.Visible = True
    End If
    
    '-------------------------------------Line and PolyLine-------------------------------------'
    
    'Set Layers and Layer Properties
    For i = 0 To LayerNum - 1
        On Error Resume Next
            Set LayerObj = acadDoc.Layers.Add(layerName(i))
            LayerObj.color = layerColor(i)
            LayerObj.LineType = layerLineType(i)
            LayerObj.LineWeight = layerLineweight(i)
        On Error GoTo 0
    Next
    
    'Select Layer for Line and PolyLine
    Set LayerObj = acadDoc.Layers(layerName(0))
    acadDoc.ActiveLayer = LayerObj
    
    'Sample Line of Tunnel Segment
    Dim SampleLineBP(0 To 2) As Double
    Dim SampleLineEP(0 To 2) As Double

    For i = 0 To num - 1
        'Set Sample Line of Tunnel Segment by 2 Points
        SampleLineBP(0) = CHUP(i)
        SampleLineBP(1) = ELUP(i)
        SampleLineEP(0) = CHDN(i)
        SampleLineEP(1) = ELDN(i)
        
        'Draw the Line (Sample Line) to AutoCAD
        If acadDoc.ActiveSpace = acModelSpace Then
            Set acadLine = acadDoc.ModelSpace.AddLine(SampleLineBP, SampleLineEP)
        Else
            Set acadLine = acadDoc.PaperSpace.AddLine(SampleLineBP, SampleLineEP)
        End If
    Next
    
    'Set UP Polyline of Tunnel Segment
    'Joint 2 arrays (CHUP, ELUP) to 1 array
    Dim UPPLine() As Double
    ReDim UPPLine((num - 1) * 2 + 1)
    
    For i = 0 To num - 1
        UPPLine(i * 2) = CHUP(i)
        UPPLine(i * 2 + 1) = ELUP(i)
    Next
    
    'Draw the polyline (UP) to AutoCAD
    If acadDoc.ActiveSpace = acModelSpace Then
        Set acadPol = acadDoc.ModelSpace.AddLightWeightPolyline(UPPLine)
    Else
        Set acadPol = acadDoc.PaperSpace.AddLightWeightPolyline(UPPLine)
    End If
    
    'Set DOWN Polyline of Tunnel Segment
    Dim DNPLine() As Double
    ReDim DNPLine((num - 1) * 2 + 1)
    
    For i = 0 To num - 1
        'Joint 2 arrays (CHDN, ELDN) to 1 array
        DNPLine(i * 2) = CHDN(i)
        DNPLine(i * 2 + 1) = ELDN(i)
    Next
    
    'Draw the polyline (DN) to AutoCAD
    If acadDoc.ActiveSpace = acModelSpace Then
        Set acadPol = acadDoc.ModelSpace.AddLightWeightPolyline(DNPLine)
    Else
        Set acadPol = acadDoc.PaperSpace.AddLightWeightPolyline(DNPLine)
    End If
    
    '-------------------------------------Text-------------------------------------'
    
    'Select Layer for Line and PolyLine
    Set LayerObj = acadDoc.Layers(layerName(1))
    acadDoc.ActiveLayer = LayerObj
    
    'Text properties
    Set TextStyle = acadDoc.TextStyles.Add(TextFont)
    
    Select Case TextFont
        Case "Angsana": TextFontFile = "angsana.shx"
        Case "Cordia": TextFontFile = "cordia.shx"
        Case "Romans": TextFontFile = "romans.shx"
        Case "Simplx": TextFontFile = "simplx.shx"
    End Select
    
    TextStyle.fontFile = TextFont
    acadDoc.ActiveTextStyle = TextStyle
    
    'Text Alignment Case
    Select Case TextAlign
        Case "Top Left": TextAlign = acAlignmentTopLeft
        Case "Top Center": TextAlign = acAlignmentTopCenter
        Case "Top Right": TextAlign = acAlignmentTopRight
        Case "Middle Left": TextAlign = acAlignmentMiddleLeft
        Case "Middle Center": TextAlign = acAlignmentMiddleCenter
        Case "Middle Right": TextAlign = acAlignmentMiddleRight
        Case "Bottom Left": TextAlign = acAlignmentBottomLeft
        Case "Bottom Center": TextAlign = acAlignmentBottomCenter
        Case "Bottom Right": TextAlign = acAlignmentBottomRight
    End Select
    
    For i = 1 To num - 1
        'The point at the beginning of 3 axis.
        ZeroPoint(0) = 0
        ZeroPoint(1) = 0
        ZeroPoint(2) = 0
        
        'Set the insertion point and Rotation angle
        InsertionPoint(0) = (CHCL(i) + CHCL(i - 1)) / 2
        InsertionPoint(1) = (ELCL(i) + ELCL(i - 1)) / 2
        InsertionPoint(2) = 0 'Z
        RotationAngle = ZACL(i)
        
        'Add the text to AutoCAD.
        Set acadText = acadDoc.ModelSpace.AddText(RingName(i), InsertionPoint, TextHeight)
        acadText.Alignment = TextAlign
        acadText.Move ZeroPoint, InsertionPoint
        acadText.Rotate InsertionPoint, DegtoRad(RotationAngle) * -1
        acadText.ScaleFactor = TextWFactor
        
    Next
    
    acadApp.ZoomExtents
    
    'Release the objects.
    Set acadDoc = Nothing
    Set acadApp = Nothing
    
    MsgBox "Placement of Tunnel Segment Profile (2D) was successfully created!"

End Sub

Sub ClearContent1()
    
    ActiveSheet.Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ActiveSheet.Select
    Range("L6:M6").Select
    Selection.ClearContents
    
    Range("A4").Select
    MsgBox "Contents were completely cleared!"
End Sub

Sub ClearContent2()
    
    ActiveSheet.Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ActiveSheet.Select
    Range("K5").Select
    Selection.ClearContents
       
    ActiveSheet.Select
    Range("J9:M10").Select
    Selection.ClearContents
    
    ActiveSheet.Select
    Range("J14:M14").Select
    Selection.ClearContents
    
    ActiveSheet.Select
    Range("K18:L18").Select
    Selection.ClearContents
    
    Range("A4").Select
    MsgBox "Contents were completely cleared!"
End Sub
