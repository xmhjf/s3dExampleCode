VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:00:39 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          ngl9190 (JG)
'   Creation Date:  Monday, Dec 2 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  27.Sep.2006     RRK TR-84724   If variables of type double are compared, appropriate CmpDblXXX functions are used.
'                                  Changed the type of the Variable "LeftOrRight" from Double to Integer
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double

    'My Dims Objects
    Dim ObjBox As Object
    Dim ObjWheel As Object
    Dim ObjVent As Object
    Dim ObjDome As Object
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    
    'My Dims Vars
    Dim startPos As IJDPosition
    Dim endPos As IJDPosition
    Dim diameter As Double
    Dim loopCounter As Integer
    Dim Coords(0 To 7)  As New AutoMath.DPosition
    Set startPos = New DPosition
    Set endPos = New DPosition
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    
    iOutput = 0

'Insert your code for output 1 (objBox)
    startPos.Set -1.6216, 8.1693, 1.1176
    endPos.Set 1.6216, -8.1693, 4.6228
  
    Set ObjBox = PlaceBox(m_outputColl, startPos, endPos)
  
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing
    
'Insert your code for output 2 to 9 (ObjWheel)
'Gather all the necessary starting coordinates
'Uses Coords(0 to 7) as new AutoMath.DPosition
    Coords(0).Set -0.7937, -7.2549, 0.5334 'SW Outside
    Coords(1).Set 0.7176, -7.2549, 0.5334 'SE Outside
    Coords(2).Set -0.7937, -5.6547, 0.5334  'SW Inside
    Coords(3).Set 0.7176, -5.6547, 0.5334 'SE Inside
    Coords(4).Set -0.7937, 5.6547, 0.5334 'NW Inside
    Coords(5).Set 0.7176, 5.6547, 0.5334 'NE Inside
    Coords(6).Set -0.7937, 7.2549, 0.5334 'NW Inside
    Coords(7).Set 0.7176, 7.2549, 0.5334 'NE Outside
    
'Loop through and build each wheel and add it to the output collection
    For loopCounter = 0 To 7
        endPos.Set Coords(loopCounter).x - 0.0762, Coords(loopCounter).y, Coords(loopCounter).z  'Adjust end point by wheel thickness
        diameter = 1.016
        Set ObjWheel = PlaceCylinder(m_outputColl, Coords(loopCounter), endPos, diameter, True)
    
        ' Set the output
            iOutput = iOutput + 1
            m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjWheel
            Set ObjWheel = Nothing
            
        Dim LeftOrRight As Integer
        LeftOrRight = loopCounter Mod 2
        
        'Build inner cylinder.
        If LeftOrRight <> 0 Then
        'Right Wheel we need to adjust both the start and end points
            Coords(loopCounter).x = Coords(loopCounter).x - 0.0762 'Fix Start Point
            endPos.Set Coords(loopCounter).x - 0.0254, Coords(loopCounter).y, Coords(loopCounter).z  'Adjust end point by wheel thickness
        Else
        'Left wheels the start points are ok need to adjust the end point
            endPos.Set Coords(loopCounter).x + 0.0254, Coords(loopCounter).y, Coords(loopCounter).z  'Adjust end point by wheel thickness
        End If
        diameter = 1.0668
        Set ObjWheel = PlaceCylinder(m_outputColl, Coords(loopCounter), endPos, diameter, True)
        
        ' Set the output
            iOutput = iOutput + 1
            m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjWheel
            Set ObjWheel = Nothing
        
    Next loopCounter
'End Wheel Construction

'Insert your code for output 10 To 17 (objVent)
'These will be used for both the cyclinder and dome
'Point is located at bottom of dome and top of cylinder.
'Uses Coords(0 to 7) as new AutoMath.DPosition
    Coords(0).Set 0#, 6.8961, 4.7244
    Coords(1).Set 0#, -4.3561, 4.7244
    Coords(2).Set 0#, -2.286, 4.7244
    Coords(3).Set 0#, -0.762, 4.7244
    Coords(4).Set 0#, 0.762, 4.7244
    Coords(5).Set 0#, 2.286, 4.7244
    Coords(6).Set 0#, 4.3561, 4.7244
    Coords(7).Set 0#, -6.8961, 4.7244

'   Loops through each instance of a cylinder and places the cylinder
    For loopCounter = 0 To 7
        endPos.Set Coords(loopCounter).x, Coords(loopCounter).y, 4.6228
        diameter = 0.5
        Set ObjVent = PlaceCylinder(m_outputColl, Coords(loopCounter), endPos, diameter, True)
        ' Set the output
            iOutput = iOutput + 1
            m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjVent
            Set ObjVent = Nothing
    Next loopCounter
'End Cylinder Construction

' Insert your code for output 18 To 25 (ObjDomes)
'   Loops through each instance of a dome and places the dome based on
'   previously definded array of coordinates.
    For loopCounter = 0 To 7
        Dim EllipNormX As Double, EllipNormY As Double, EllipNormZ As Double
        Dim MajorX As Double, MajorY As Double, MajorZ As Double
        Dim axis    As New AutoMath.DVector
        Dim PI As Double
        Dim parRatio As Double
        Dim oDomeEllip As IngrGeom3D.EllipticalArc3d
    
        PI = 4 * Atn(1)
    
        axis.Set 0, 0, 1
        MajorX = diameter / 2
        MajorY = 0
        MajorZ = 0
        'The Ratio of minor axis over major axis
        parRatio = (0.0508 / (0.508 / 2)) 'domeHeight / (domeDiameter /2)
        'Normal vector to plane of ellipce)
        EllipNormX = 0
        EllipNormY = -1
        EllipNormZ = 0
    
        Set oDomeEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                                Nothing, Coords(loopCounter).x, _
                                                Coords(loopCounter).y, Coords(loopCounter).z, EllipNormX, _
                                                EllipNormY, EllipNormZ, MajorX, MajorY, MajorZ, _
                                                parRatio, 0, PI / 2)
    
        Set ObjDome = PlaceRevolution(m_outputColl, oDomeEllip, axis, Coords(loopCounter), PI * 2, False)
    
    ' Set the output
        iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDome
        Set ObjDome = Nothing
    
    Next loopCounter
'End Dome Construction

' Insert your code for output 18 To 26 To 45 (ObjPlaneForDump)
    Dim offset(0 To 3) As Double  'used to itterate through to build all four
'Each element is based as a distance off the first
    offset(0) = 0
    offset(1) = 3.1401
    offset(2) = 6.2611
    offset(3) = 9.4012
    For loopCounter = 0 To 3
        Dim Points(0 To 3) As New AutoMath.DPosition
        'Plane 1
        Points(0).Set 1.2192, -5.9198 + offset(loopCounter), 1.1176
        Points(1).Set 0.9144, -5.0054 + offset(loopCounter), 0.508
        Points(2).Set -0.9144, -5.0054 + offset(loopCounter), 0.508
        Points(3).Set -1.2192, -5.9198 + offset(loopCounter), 1.1176
        Call buildPlane(m_outputColl, Points, iOutput, arrayOfOutputs)
        'Plane 2
        Points(0).Set 0.9144, -5.0054 + offset(loopCounter), 0.508
        Points(1).Set 0.9144, -4.3958 + offset(loopCounter), 0.508
        Points(2).Set -0.9144, -4.3958 + offset(loopCounter), 0.508
        Points(3).Set -0.9144, -5.0054 + offset(loopCounter), 0.508
        Call buildPlane(m_outputColl, Points, iOutput, arrayOfOutputs)
        'Plane 3
        Points(0).Set -0.9144, -4.3958 + offset(loopCounter), 0.508
        Points(1).Set -1.2192, -3.4814 + offset(loopCounter), 1.1176
        Points(2).Set 1.2192, -3.4814 + offset(loopCounter), 1.1176
        Points(3).Set 0.9144, -4.3958 + offset(loopCounter), 0.508
        Call buildPlane(m_outputColl, Points, iOutput, arrayOfOutputs)
        'Plane 4
        Points(0).Set 1.2192, -3.4814 + offset(loopCounter), 1.1176
        Points(1).Set 1.2192, -5.9198 + offset(loopCounter), 1.1176
        Points(2).Set 0.9144, -5.0054 + offset(loopCounter), 0.508
        Points(3).Set 0.9144, -4.3958 + offset(loopCounter), 0.508
        Call buildPlane(m_outputColl, Points, iOutput, arrayOfOutputs)
        'Plane 5
        Points(0).Set -1.2192, -3.4814 + offset(loopCounter), 1.1176
        Points(1).Set -1.2192, -5.9198 + offset(loopCounter), 1.1176
        Points(2).Set -0.9144, -5.0054 + offset(loopCounter), 0.508
        Points(3).Set -0.9144, -4.3958 + offset(loopCounter), 0.508
        Call buildPlane(m_outputColl, Points, iOutput, arrayOfOutputs)
        'We don't need the base of the pyramid so leave out of output.
    Next loopCounter

Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

'Used to build the actual plane and add it to the output collection.
'Made generic to possibly add to a library in the future.
'Accepts array(n) of type DPosition and builds an n verticie plane.
'Make a function and remove output if moved to library and return type object
Public Sub buildPlane(ByVal m_outputColl As Object, ByRef passPoints() As AutoMath.DPosition, ByRef iOutput As Double, ByRef arrayOfOutputs() As String)
    
    'Dim and set vars
    Dim newPlane As Object
    Dim arrayPoints() As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim numOfVertices As Integer
    Dim arrayIndex As Integer
    Dim loopCount As Integer
    arrayIndex = 0
    numOfVertices = UBound(passPoints)
    
    'Create array to be passed to Planes3D routine
    ReDim Preserve arrayPoints(((numOfVertices + 1) * 3))
    For loopCount = 0 To numOfVertices
        arrayPoints(arrayIndex) = passPoints(loopCount).x
        arrayIndex = arrayIndex + 1
        arrayPoints(arrayIndex) = passPoints(loopCount).y
        arrayIndex = arrayIndex + 1
        arrayPoints(arrayIndex) = passPoints(loopCount).z
        arrayIndex = arrayIndex + 1
    Next loopCount
    
    Set newPlane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, numOfVertices + 1, arrayPoints)

    '***remove if changed to function***
    'Create output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), newPlane
    Set newPlane = Nothing

    Set geomFactory = Nothing

End Sub

