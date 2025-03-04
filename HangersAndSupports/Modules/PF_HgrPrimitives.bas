Attribute VB_Name = "PF_HgrPrimitives"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   PF_HgrPrimitives.bas
'   ProgID:         PF_HgrPrimitives.bas
'   Author:         Pelican Forge
'   Creation Date:  NA

'   Description:
'       This class contains common functionality that bridges the differences between SP3D and SupportModeler Code
'       It is intended to make conversion from SupMod libraries easier as well as to modularize common functionality
'
'   Change History:
'   dd.mmm.yyyy       who            change description
'   03.May.2006       JRM            Apply fixes from code review (TR 95926)
'   05.May.2006       SS             Apply fixes from code review (TR 95926), Points 7,9,10
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

' Note: Function has been added below to return Pi (GetPi())
' This constant has been left here so we do not break any existing libraries.
' Do not use it in any new development - may.03.2006
Const PI As Double = 3.14159265358979
Private Const MODULE = "PF_HgrPrimitives" 'Used for error messages
Private m_oCrvElems As IJElements
Private objColl As New JObjectCollection
Private m_oContour As ComplexString3d
Private m_extrusion As Integer
Private m_output As String
Private m_numpieces As Integer
Private m_Vertices(0 To 110) As Double
Private m_currentVertex As Integer

' ---------------------------------------------------------------------------
' Name: AddHalfClamp
' Description: Create the graphical representation of a half of a beam clamp
' Example: This function is used in the Anvil library Fig 133.  Each side of the beam clamp is 1 half clamp
'
' Inputs - H - double, GAP - double, FW - Double, SW - double, ST - double, inverted - boolean,
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddHalfClamp(H As Double, GAP As Double, FW As Double, FT As Double, SW As Double, ST As Double, inverted As Boolean, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
    Const METHOD = "AddHalfClamp"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If

    Dim iInverter As Integer
    Dim sRotate As String
    If inverted = False Then
        iInverter = 1
        sRotate = "90"
    Else
        iInverter = -1
        sRotate = "270"
    End If

    AddComposite 10, "EXTRUDED", "", "COMP"
    AddLine -SW / 2, iInverter * (GAP / 2), -H, -SW / 2, iInverter * (GAP / 2 + ST), -H, outputcoll, "", "BOT"
    AddLine -SW / 2, iInverter * (GAP / 2 + ST), -H, -SW / 2, iInverter * (GAP / 2 + ST), -ST, outputcoll, "", "BR"
    AddLine -SW / 2, iInverter * (GAP / 2 + ST), -ST, -SW / 2, iInverter * (FW / 2), -ST, outputcoll, "", "MO"
    AddArc (-90), (90), FT / 2 + ST, "ROTY(90) * ROTX(" & sRotate & ") * " + Loc(-SW / 2, iInverter * (FW / 2), FT / 2), outputcoll, "", "OUT1"
    AddLine -SW / 2, iInverter * (FW / 2), FT + ST, -SW / 2, iInverter * (GAP / 2), FT + ST, outputcoll, "", "TO"
    AddLine -SW / 2, iInverter * (GAP / 2), FT + ST, -SW / 2, iInverter * (GAP / 2), FT, outputcoll, "", "E"
    AddLine -SW / 2, iInverter * (GAP / 2), FT, -SW / 2, iInverter * (FW / 2), FT, outputcoll, "", "TI"
    AddArc (-90), (90), FT / 2, "ROTY(90) * ROTX(" & sRotate & ") * " + Loc(-SW / 2, iInverter * (FW / 2), FT / 2), outputcoll, "", "IN1"
    AddLine -SW / 2, iInverter * (FW / 2), 0, -SW / 2, iInverter * (GAP / 2), 0, outputcoll, "", "MI"
    AddLine -SW / 2, iInverter * (GAP / 2), 0, -SW / 2, iInverter * (GAP / 2), -H, outputcoll, "", "BL"
    AddExtrusion SW, 0, 0, 1, outputcoll, output, name
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddEndPlate
' Description: Create the graphical representation of an end plate.  An ed plate will be a square plate with the Pipe Radius cut out of one end
'
' Inputs - H - double, T - double, W - Double, PIPE_DIA - double, A - double, ALPHA - double, X_OFFSET - double, Z_OFFSET - double
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddEndPlate(H As Double, T As Double, W As Double, PIPE_DIA As Double, A As Double, ALPHA As Double, X_OFFSET As Double, Z_OFFSET As Double, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddEndPlate"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim Beta As Double
    Beta = (180) - ALPHA
    
    AddComposite 4, "EXTRUDED", "", "COMP"
    AddLine X_OFFSET, -W / 2, -Z_OFFSET, X_OFFSET, W / 2, -Z_OFFSET, outputcoll, "", "BOTTOM"
    AddLine X_OFFSET, W / 2, -Z_OFFSET, X_OFFSET, W / 2, -(H + A + Z_OFFSET), outputcoll, "", "RIGHT"
    AddArc -Beta, -ALPHA, PIPE_DIA / 2, "ROTX(90) * ROTY(180) * ROTZ(90) *" + Loc(X_OFFSET, 0, -(H + PIPE_DIA / 2# + Z_OFFSET)), outputcoll, "", "TOP_ARC"
    AddLine X_OFFSET, -W / 2, -(H + A + Z_OFFSET), X_OFFSET, -W / 2#, -Z_OFFSET, outputcoll, "", "LEFT"
    AddExtrusion T, 0, 0, 1, outputcoll, output, name
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddBox
' Description: Create the graphical representation of a box
'
' Inputs - X - double, Y - double, Z - Double, strmatrix - string
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddBox(x As Double, y As Double, z As Double, strmatrix As String, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddBox"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory

    Dim matrix             As IJDT4x4
    GetMatrixFromString strmatrix, matrix

    Dim dblPoints(0 To 14) As Double
    
    Dim numPtsLineStr As Integer
    Dim vector(4) As Double
        
    vector(1) = 0
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(0) = vector(1)
    dblPoints(1) = vector(2)
    dblPoints(2) = vector(3)
    
    vector(1) = x
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(3) = vector(1)
    dblPoints(4) = vector(2)
    dblPoints(5) = vector(3)
    
    vector(1) = x
    vector(2) = y
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(6) = vector(1)
    dblPoints(7) = vector(2)
    dblPoints(8) = vector(3)
    
    vector(1) = 0
    vector(2) = y
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(9) = vector(1)
    dblPoints(10) = vector(2)
    dblPoints(11) = vector(3)
        
    vector(1) = 0
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(12) = vector(1)
    dblPoints(13) = vector(2)
    dblPoints(14) = vector(3)
        
    numPtsLineStr = 5
        
    Dim oLineStr As LineString3d
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                            numPtsLineStr, dblPoints)
'    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
'                            numPtsLineStr, dblPoints)
                            
    vector(1) = 0
    vector(2) = 0
    vector(3) = 1
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    
    'deleted the oContour and Crvelems for DI - 95926, Point No.7 - SS
    
    Dim oBox As Projection3d
    Set oBox = oGeomFactory.Projections3d.CreateByCurve(Nothing, _
                            oLineStr, vector(1) - dblPoints(0), vector(2) - dblPoints(1), vector(3) - dblPoints(2), z, True)
                            
                          
    Set oLineStr = Nothing
        
    outputcoll.AddOutput output, oBox
    
    Set oBox = Nothing
    Set oGeomFactory = Nothing


Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddCylinderWP
' Description: Create the graphical representation of a cylinder.  Specify the start point (3d), end point (3d) and the radius.
'
' Inputs - startx - double, starty - double, startz - Double, endx - double, endy - double, endz - Double, RADIUS - double,
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddCylinderWP(startx As Double, starty As Double, startz As Double, endx As Double, endy As Double, endz As Double, Radius As Double, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddCylinderWP"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim stPoint1   As New AutoMath.DPosition
    Dim enPoint1   As New AutoMath.DPosition
    Dim ldiam1     As Double
    Dim objCylinder1  As Object
    
    stPoint1.Set startx, starty, startz
    enPoint1.Set endx, endy, endz
    
    ldiam1 = Radius * 2
    
    Set objCylinder1 = PlaceCylinder(outputcoll, stPoint1, enPoint1, ldiam1, True)
    outputcoll.AddOutput output, objCylinder1
    
    Set objCylinder1 = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddCylinder
' Description: Create the graphical representation of a cylinder.  Specify the start point (3d matrix), length and the radius
'
' Inputs - LENGTH - double, RADIUS - double, strmatrix - string,
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddCylinder(LENGTH As Double, Radius As Double, strmatrix As String, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddCylinder"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    If LENGTH = 0 Then
        LENGTH = 0.00001
    End If
    
    Dim stPoint1   As New AutoMath.DPosition
    Dim enPoint1   As New AutoMath.DPosition
    Dim ldiam1     As Double
    Dim objCylinder1  As Object
    Dim matrix As IJDT4x4
    GetMatrixFromString strmatrix, matrix

    Dim vector(4) As Double
        
    vector(1) = 0
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    stPoint1.Set vector(1), vector(2), vector(3)
    
    vector(1) = 0
    vector(2) = 0
    vector(3) = LENGTH
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    enPoint1.Set vector(1), vector(2), vector(3)
    
    ldiam1 = Radius * 2
    
    Set objCylinder1 = PlaceCylinder(outputcoll, stPoint1, enPoint1, ldiam1, True)
    outputcoll.AddOutput output, objCylinder1
    
    Set objCylinder1 = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddElbow
' Description: Create the graphical representation of a cylinder elbow.  Specify the cylinder radius, the elbow radius,
'              the sweep angle and the starting position (3d matrix)
'
' Inputs - rodradius - double, elbowradius - double, SweepAngle - double, strmatrix - string,
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddElbow(rodradius As Double, elbowradius As Double, SweepAngle As Double, strmatrix As String, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddElbow"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory
    Dim matrix As IJDT4x4
    Dim elbowCenter(0 To 2) As Double
    Dim elbowNormal(0 To 2) As Double
    Dim circleCenter(0 To 2) As Double
    Dim circleNormal(0 To 2) As Double
    Dim vector(4) As Double

    GetMatrixFromString strmatrix, matrix

    vector(1) = 0
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    elbowCenter(0) = vector(1)
    elbowCenter(1) = vector(2)
    elbowCenter(2) = vector(3)

    vector(1) = 0
    vector(2) = -1
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    elbowNormal(0) = vector(1) - elbowCenter(0)
    elbowNormal(1) = vector(2) - elbowCenter(1)
    elbowNormal(2) = vector(3) - elbowCenter(2)

    vector(1) = elbowradius
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    circleCenter(0) = vector(1)
    circleCenter(1) = vector(2)
    circleCenter(2) = vector(3)

    vector(1) = elbowradius
    vector(2) = 0
    vector(3) = 1
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    circleNormal(0) = vector(1) - circleCenter(0)
    circleNormal(1) = vector(2) - circleCenter(1)
    circleNormal(2) = vector(3) - circleCenter(2)

    ' Create the Bolt Cross section curve to revolve.
    Dim BoltXSection As Circle3d

    Set BoltXSection = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                        circleCenter(0), circleCenter(1), circleCenter(2), _
                                                                        circleNormal(0), circleNormal(1), circleNormal(2), _
                                                                        rodradius)

    'Create the Revolution representing the Semi-Circle Section of the UBolt
    Dim Elbow As Revolution3d
    
    'commented the code and added GetPi() function to the argument to get the PI value - DI-95926
    
'    Set Elbow = oGeomFactory.Revolutions3d.CreateByCurve(Nothing, _
'                                                            BoltXSection, _
'                                                            elbowNormal(0), elbowNormal(1), elbowNormal(2), _
'                                                            elbowCenter(0), elbowCenter(1), elbowCenter(2), _
'                                                            PI * SweepAngle / 180, True)
 
    Set Elbow = oGeomFactory.Revolutions3d.CreateByCurve(Nothing, _
                                                            BoltXSection, _
                                                            elbowNormal(0), elbowNormal(1), elbowNormal(2), _
                                                            elbowCenter(0), elbowCenter(1), elbowCenter(2), _
                                                            GetPi() * SweepAngle / 180, True)
 
    Dim objRemover As IJDObject
    Set objRemover = BoltXSection
    objRemover.Remove
    Set BoltXSection = Nothing
 
    outputcoll.AddOutput output, Elbow

    Set Elbow = Nothing
    Set oGeomFactory = Nothing
 Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddPolygon
' Description: Create the graphical representation of a Polygon.  Specify the number of points and if it will be flat or extruded
'
' Inputs - numPoints - integer, flatOrExtruded - string,
'          Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddPolygon(numPoints As Integer, flatOrExtruded As String, output As String, Optional name As String = "None")
Const METHOD = "AddPolygon"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    If flatOrExtruded = "EXTRUDED" Then
        m_extrusion = 1
    Else
        m_extrusion = 0
    End If
    
    m_output = output
    m_numpieces = numPoints
    Set m_oCrvElems = objColl
    m_currentVertex = 0
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddPoint
' Description: Create the graphical representation of a point.  Specify the location (X, y, z).
'
' Inputs - locX - double, locY - double, locZ - double,
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddPoint(locX As Double, locY As Double, locZ As Double, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddPoint"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory
    
    If output = "" Then
        m_Vertices(m_currentVertex * 3) = locX
        m_Vertices(m_currentVertex * 3 + 1) = locY
        m_Vertices(m_currentVertex * 3 + 2) = locZ
        
        m_currentVertex = m_currentVertex + 1
        
        If m_currentVertex = m_numpieces Then
            Dim oLineStr As LineString3d

            m_Vertices(m_currentVertex * 3) = m_Vertices(0)
            m_Vertices(m_currentVertex * 3 + 1) = m_Vertices(1)
            m_Vertices(m_currentVertex * 3 + 2) = m_Vertices(2)
            
            Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                m_numpieces + 1, m_Vertices)
        
            If m_extrusion = 0 Then
                outputcoll.AddOutput m_output, oLineStr
            Else
                Set m_oCrvElems = objColl
                m_oCrvElems.Add oLineStr
                Set m_oContour = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, m_oCrvElems)
            End If
        
            Set oLineStr = Nothing
        End If
    Else
        ' create a point and add it to the outputs
        Dim oPoint As Point3d
        
        Set oPoint = oGeomFactory.Points3d.CreateByPoint(Nothing, locX, locY, locZ)
        
        outputcoll.AddOutput output, oPoint
        
        Set oPoint = Nothing
    End If
    Set oGeomFactory = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddCone
' Description: Create the graphical representation of a Cone.  Specify the bottom radius, the top radius, the height and a starting point
'
' Inputs - bottomRadius - double, topRadius - double, height - double, strmatrix - string
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddCone(bottomRadius As Double, topRadius As Double, height As Double, strmatrix As String, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddCone"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory
    Dim centerBottom(0 To 2) As Double
    Dim centerTop(0 To 2) As Double
    Dim startBottom(0 To 2) As Double
    Dim startTop(0 To 2) As Double
    
    Dim matrix As IJDT4x4
    GetMatrixFromString strmatrix, matrix
    
    Dim vector(4) As Double
        
    vector(1) = 0
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    centerBottom(0) = vector(1)
    centerBottom(1) = vector(2)
    centerBottom(2) = vector(3)
    
    vector(1) = 0
    vector(2) = 0
    vector(3) = height
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    centerTop(0) = vector(1)
    centerTop(1) = vector(2)
    centerTop(2) = vector(3)
    
    vector(1) = bottomRadius
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    startBottom(0) = vector(1)
    startBottom(1) = vector(2)
    startBottom(2) = vector(3)
    
    vector(1) = topRadius
    vector(2) = 0
    vector(3) = height
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    startTop(0) = vector(1)
    startTop(1) = vector(2)
    startTop(2) = vector(3)
    
    Dim oCone As Cone3d
    
    Set oCone = oGeomFactory.Cones3d.CreateBy4Pts(Nothing, centerBottom(0), centerBottom(1), centerBottom(2), _
        centerTop(0), centerTop(1), centerTop(2), startBottom(0), startBottom(1), startBottom(2), _
        startTop(0), startTop(1), startTop(2), True)

    outputcoll.AddOutput output, oCone
        
    Set oCone = Nothing
    Set oGeomFactory = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddEllipse
' Description: Create the graphical representation of an ellipse.  Specify the two radii, whether it is flat or extruded and a starting point
'
' Inputs - radius1 - double, radius2 - double, flatOrExtruded - string, strmatrix - string
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddEllipse(radius1 As Double, radius2 As Double, flatOrExtruded As String, strmatrix As String, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddEllipse"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory
    Dim ellipseCenter(0 To 2) As Double
    Dim ellipseNormal(0 To 2) As Double
    Dim ellipseMajor(0 To 2) As Double
    
    Dim matrix As IJDT4x4
    GetMatrixFromString strmatrix, matrix
    
    Dim vector(4) As Double
        
    vector(1) = 0
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    ellipseCenter(0) = vector(1)
    ellipseCenter(1) = vector(2)
    ellipseCenter(2) = vector(3)
    
    vector(1) = 0
    vector(2) = 0
    vector(3) = 1
    vector(4) = 1
    RotateVector vector, matrix, vector
    ellipseNormal(0) = vector(1)
    ellipseNormal(1) = vector(2)
    ellipseNormal(2) = vector(3)
    
    vector(1) = radius1
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    RotateVector vector, matrix, vector
    ellipseMajor(0) = vector(1)
    ellipseMajor(1) = vector(2)
    ellipseMajor(2) = vector(3)
    
    Dim oEllipse As Ellipse3d
    Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, ellipseCenter(0), ellipseCenter(1), ellipseCenter(2), _
                    ellipseNormal(0), ellipseNormal(1), ellipseNormal(2), ellipseMajor(0), ellipseMajor(1), ellipseMajor(2), radius2 / radius1)
    
    If flatOrExtruded = "EXTRUDED" Then
        Set m_oCrvElems = objColl
        m_oCrvElems.Add oEllipse
        
        Set m_oContour = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, m_oCrvElems)
    Else
        outputcoll.AddOutput output, oEllipse
    End If
    
    Set oEllipse = Nothing
    Set oGeomFactory = Nothing
 Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddComposite
' Description: Create representation of an Composite.  Specify the number of pieces and whether it is flat or extruded
'
' Inputs - numPieces - Integer, flatOrExtruded - string
'          Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddComposite(numPieces As Integer, flatOrExtruded As String, output As String, Optional name As String = "None")
Const METHOD = "AddComposite"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    If flatOrExtruded = "EXTRUDED" Then
        m_extrusion = 1
    Else
        m_extrusion = 0
    End If
    
    Set m_oCrvElems = objColl
    
    m_output = output
    m_numpieces = numPieces
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddExtrusion
' Description: Create the graphical representation of an extrusion.  Specify the amount to extrude in the x, y and z directions.  Also specify
'              if the extrusion is to be capped
'
' Inputs - extrudeX - double, extrudeY - double, extrudeZ - double, capped - boolean
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddExtrusion(extrudeX As Double, extrudeY As Double, extrudeZ As Double, capped As Boolean, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddExtrusion"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory
    Dim LENGTH As Double

    LENGTH = Sqr(extrudeX * extrudeX + extrudeY * extrudeY + extrudeZ * extrudeZ)
    
    If LENGTH = 0 Then
        GoTo ErrorHandler
    End If
    
    Dim oExtrusion As Projection3d
    Set oExtrusion = oGeomFactory.Projections3d.CreateByCurve(Nothing, _
                            m_oContour, extrudeX / LENGTH, extrudeY / LENGTH, extrudeZ / LENGTH, LENGTH, capped)

    Dim objRemover As IJDObject
    Set objRemover = m_oContour
    objRemover.Remove
    Set m_oContour = Nothing
    
    outputcoll.AddOutput output, oExtrusion
    
    Dim I As Integer

    For I = 1 To m_oCrvElems.Count
        Set objRemover = m_oCrvElems(I)
        objRemover.Remove
    Next
    
    m_oCrvElems.Clear
    Set m_oCrvElems = Nothing
    Set oExtrusion = Nothing
    Set oGeomFactory = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddLine
' Description: Create the graphical representation of a line.  Specify the start x, y, z and end x, y, z
'
' Inputs - startx - double, starty - double, startz - double, endx - double, endy - double, endz - double,
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddLine(startx As Double, starty As Double, startz As Double, endx As Double, endy As Double, endz As Double, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddLine"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory
    Dim dblPoints(0 To 5) As Double
        
    dblPoints(0) = startx
    dblPoints(1) = starty
    dblPoints(2) = startz
    
    dblPoints(3) = endx
    dblPoints(4) = endy
    dblPoints(5) = endz
        
    Dim oLineStr As LineString3d
                                
    ' if this is part of a composite polygon or curve, the outputcoll will be nothing
    If output = "" Then
        Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                            2, dblPoints)
                            
        m_numpieces = m_numpieces - 1
    
        If m_numpieces >= 0 Then
            Set m_oCrvElems = objColl
            m_oCrvElems.Add oLineStr
        End If
    
        Set oLineStr = Nothing
    
        ' if the number of pieces left is zero, finish off the composite
        If m_numpieces = 0 Then
            
            Set m_oContour = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, m_oCrvElems)
            
            Dim I As Integer

            If m_extrusion = 0 Then
                outputcoll.AddOutput m_output, m_oContour
                
                For I = 1 To m_oCrvElems.Count
                    Dim objRemover As IJDObject
                    Set objRemover = m_oCrvElems(I)
                    objRemover.Remove
                Next
                
                m_oCrvElems.Clear
                Set m_oContour = Nothing
                Set m_oCrvElems = Nothing
            End If
    
        End If

        If m_numpieces < -1 Then
            m_numpieces = -1
        End If
    Else
        Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                            2, dblPoints)
        outputcoll.AddOutput output, oLineStr
    
        Set oLineStr = Nothing
    End If
    Set oGeomFactory = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddArc
' Description: Create the graphical representation of an Arc.  Specify the start and end angles, the radius and a starting point
'
' Inputs - startAngle - double, endAngle - double, RADIUS - double, strmatrix - string,
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddArc(startAngle As Double, endAngle As Double, Radius As Double, strmatrix As String, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddArc"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory
    Dim oArc As Arc3d
    Dim matrix As IJDT4x4
    Dim center(0 To 2) As Double
    Dim normal(0 To 2) As Double
    Dim startloc(0 To 2) As Double
    Dim endloc(0 To 2) As Double
    Dim vector(4) As Double
    Dim rotMatrix As IJDT4x4
    Dim ANGLE As Double
    
    GetMatrixFromString strmatrix, matrix

    ' get center (starts at origin)
    vector(1) = 0
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    center(0) = vector(1)
    center(1) = vector(2)
    center(2) = vector(3)

    ' get normal (starts along z)
    vector(1) = 0
    vector(2) = 0
    vector(3) = 1
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    normal(0) = vector(1) - center(0)
    normal(1) = vector(2) - center(1)
    normal(2) = vector(3) - center(2)

    Dim zAxis               As IJDVector
    Set zAxis = New DVector
    zAxis.Set 0#, 0#, 1#

    ' get start point (starts at 1, 0, 0 and rotates around z by startAngle)
    vector(1) = Radius
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    'commented PI constant and added GetPi() function - for DI - 95926
    'ANGLE = PI * startAngle / 180
    ANGLE = GetPi() * startAngle / 180
    Set rotMatrix = matrix.Clone()
    rotMatrix.Rotate ANGLE, zAxis
    MultiplyVectorByMatrix vector, rotMatrix, vector
    startloc(0) = vector(1)
    startloc(1) = vector(2)
    startloc(2) = vector(3)

    ' get end point (starts at 1, 0, 0 and rotates around z by endAngle)
    vector(1) = Radius
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    ANGLE = PI * endAngle / 180
    Set rotMatrix = matrix.Clone()
    rotMatrix.Rotate ANGLE, zAxis
    MultiplyVectorByMatrix vector, rotMatrix, vector
    endloc(0) = vector(1)
    endloc(1) = vector(2)
    endloc(2) = vector(3)
    
    ' if this is part of a composite polygon or curve, the outputcoll will be nothing
    If output = "" Then
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                            center(0), center(1), center(2), _
                            normal(0), normal(1), normal(2), _
                            startloc(0), startloc(1), startloc(2), _
                            endloc(0), endloc(1), endloc(2))
        m_numpieces = m_numpieces - 1

' an arc can be extruded by itself, so there don't have to be numpieces to add it to the ocrvelems
        If m_oCrvElems Is Nothing Then
            Set m_oCrvElems = objColl
            m_numpieces = 0
            m_extrusion = 1
        End If

        m_oCrvElems.Add oArc

        Set oArc = Nothing
        
        ' if the number of pieces left is zero, finish off the composite
        If m_numpieces = 0 Then
            Set m_oContour = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, m_oCrvElems)
            Dim I As Integer
            If m_extrusion = 0 Then
                outputcoll.AddOutput m_output, m_oContour
                For I = 1 To m_oCrvElems.Count
                    Dim objRemover As IJDObject
                    Set objRemover = m_oCrvElems(I)
                    objRemover.Remove
                Next
                m_oCrvElems.Clear
                Set m_oContour = Nothing
                Set m_oCrvElems = Nothing
            End If
        End If
        If m_numpieces < -1 Then
            m_numpieces = -1
        End If
    Else
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                            center(0), center(1), center(2), _
                            normal(0), normal(1), normal(2), _
                            startloc(0), startloc(1), startloc(2), _
                            endloc(0), endloc(1), endloc(2))
        outputcoll.AddOutput output, oArc
        Set oArc = Nothing
    End If
    Set oGeomFactory = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddPort
' Description: Create a port.  Specify the port name, the location (x, y, z), the x direction (x, y, z) and the z direction (x, y, z)
'
' Inputs - portName - string, locX - double, locY - double, locZ - double, xDirX - double, xDirY - double, xDirZ - double,
'           zDirX - double, zDirY - double, zDirZ - double,
'          outputcoll - SP3D Output Collection Object, Output - string, Name - string
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddPort(portName As String, locX As Double, locY As Double, locZ As Double, xDirX As Double, xDirY As Double, xDirZ As Double, zDirX As Double, zDirY As Double, zDirZ As Double, outputcoll As IJDOutputCollection, output As String, ByVal part As Variant, Optional name As String = "None")
Const METHOD = "AddPort"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    ' create the hgr ports
    Dim oHgrPort        As IJHgrPort
    Dim oPortFac        As IJHgrPortFactory
    
    Set oPortFac = New HgrPortFactory
    
    Set oHgrPort = oPortFac.AddHgrPortByName(outputcoll.ResourceManager, part, portName, locX, locY, locZ, xDirX, xDirY, xDirZ, zDirX, zDirY, zDirZ, outputcoll, output)

    Set oHgrPort = Nothing
    Set oPortFac = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: GetMatrixFromString
' Description: Extract a location matrix from a complex matrix string.  Specify the matrix string
'
' Inputs - matrixstring - string, oMatrix - DT4x4,
' Outputs - none
' ---------------------------------------------------------------------------
Public Function GetMatrixFromString(ByVal matrixstring As String, oMatrix As DT4x4)
Const METHOD = "GetMatrixFromString"
    On Error GoTo ErrorHandler
    
    'Dim parsedarray(30) As Double
    Dim parsedarray() As Double ' DI#129348; dynamic memory allocation
    Dim currentSize As Long ' DI#129348
    Dim keyword As String
    Dim copiedstring As String
    Dim index As Integer

    ' remove all spaces from string
    Do While matrixstring <> ""
            If Left(matrixstring, 1) <> " " Then
                    copiedstring = copiedstring + Left(matrixstring, 1)
            End If
            matrixstring = Mid(matrixstring, 2)
    Loop
    
    matrixstring = copiedstring
    
    Dim tempstring As String

    index = 1
    
    ' parse string InStr( string, searchstring )
    currentSize = 128 ' DI#129348
    ReDim parsedarray(currentSize) ' DI#129348
    Do While matrixstring <> ""
            keyword = Left(matrixstring, 4)
            If keyword = "ROTX" Then
                    ' DI#129348 ; check whether enough memory is allocated or not
                    If currentSize < index + 2 Then
                        ReDim Preserve parsedarray(currentSize + 128)
                        currentSize = currentSize + 128
                    End If
                    ' DI#129348 ends here
                    parsedarray(index) = 3
                    index = index + 1
                    tempstring = Mid(matrixstring, 6)
                    tempstring = Left(tempstring, InStr(tempstring, ")") - 1)
                    parsedarray(index) = PI * Val(tempstring) / 180
                    index = index + 1
            End If
            If keyword = "ROTY" Then
                    ' DI#129348 ; check whether enough memory is allocated or not
                    If currentSize < index + 2 Then
                        ReDim Preserve parsedarray(currentSize + 128)
                        currentSize = currentSize + 128
                    End If
                    ' DI#129348 ends here
                    parsedarray(index) = 4
                    index = index + 1
                    tempstring = Mid(matrixstring, 6)
                    tempstring = Left(tempstring, InStr(tempstring, ")") - 1)
                    parsedarray(index) = PI * Val(tempstring) / 180
                    index = index + 1
            End If
            If keyword = "ROTZ" Then
                    ' DI#129348 ; check whether enough memory is allocated or not
                    If currentSize < index + 2 Then
                        ReDim Preserve parsedarray(currentSize + 128)
                        currentSize = currentSize + 128
                    End If
                    ' DI#129348 ends here
                    parsedarray(index) = 5
                    index = index + 1
                    tempstring = Mid(matrixstring, 6)
                    tempstring = Left(tempstring, InStr(tempstring, ")") - 1)
                    parsedarray(index) = PI * Val(tempstring) / 180
                    index = index + 1
            End If
            If keyword = "TRAN" Then
                    ' DI#129348 ; check whether enough memory is allocated or not
                    If currentSize < index + 6 Then
                        ReDim Preserve parsedarray(currentSize + 128)
                        currentSize = currentSize + 128
                    End If
                    ' DI#129348 ends here
                    matrixstring = Mid(matrixstring, 10)
                    tempstring = Left(matrixstring, InStr(matrixstring, ",") - 1)
                    parsedarray(index) = 0
                    parsedarray(index + 1) = Val(tempstring)
                    index = index + 2
                    matrixstring = Mid(matrixstring, InStr(matrixstring, ",") + 1)
                    tempstring = Left(matrixstring, InStr(matrixstring, ",") - 1)
                    parsedarray(index) = 1
                    parsedarray(index + 1) = Val(tempstring)
                    index = index + 2
                    matrixstring = Mid(matrixstring, InStr(matrixstring, ",") + 1)
                    tempstring = Left(matrixstring, InStr(matrixstring, ")") - 1)
                    parsedarray(index) = 2
                    parsedarray(index + 1) = Val(tempstring)
                    index = index + 2
            End If
    
            If InStr(matrixstring, "*") Then
                    matrixstring = Mid(matrixstring, InStr(matrixstring, "*") + 1)
            Else
                    matrixstring = ""
            End If
    Loop
     
    ' create a 4x4 identity matrix
    Dim x As Double, y As Double, z As Double
    Dim I As Integer
    Dim numPairs As Integer
    
    Dim xAxis               As IJDVector
    Set xAxis = New DVector
    xAxis.Set 1#, 0#, 0#
    
    Dim yAxis               As IJDVector
    Set yAxis = New DVector
    yAxis.Set 0#, 1#, 0#
    
    Dim zAxis               As IJDVector
    Set zAxis = New DVector
    zAxis.Set 0#, 0#, 1#
    
'    Dim oMatrix             As IJDT4x4
    Set oMatrix = New DT4x4
       
    oMatrix.LoadIdentity
    
    numPairs = (index - 1) / 2# ' DI#129348 ; index is incr. extra time so decrement and calc pairs
    ' loop through the array, applying each rotation and translation to the matrix
    For I = numPairs To 1 Step -1
            If parsedarray(I * 2 - 1) = 0 Then
                    ' this is an X value, now have a full point for translation (since we're going in reverse through the array)
                    x = parsedarray(I * 2)
                    
                    Dim trans               As IJDVector
                    Set trans = New DVector
                    trans.Set x, y, z
                    
                    oMatrix.Translate trans
            End If
            If parsedarray(I * 2 - 1) = 1 Then
                    ' this is a Y value, store it for later use
                    y = parsedarray(I * 2)
            End If
            If parsedarray(I * 2 - 1) = 2 Then
                    ' this is a Z value, store it for later use
                    z = parsedarray(I * 2)
            End If
            If parsedarray(I * 2 - 1) = 3 Then
                    ' this is an X rotation
                    oMatrix.Rotate parsedarray(I * 2), xAxis
            End If
            If parsedarray(I * 2 - 1) = 4 Then
                    ' this is a Y rotation
                    oMatrix.Rotate parsedarray(I * 2), yAxis
            End If
            If parsedarray(I * 2 - 1) = 5 Then
                    ' this is a Z rotation
                    oMatrix.Rotate parsedarray(I * 2), zAxis
            End If
    Next I
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: MultiplyVectorByMatrix
' Description: Multiply a the specified vector by the specified matrix
'
' Inputs - resultvec() - double array, matrix - DT4x4, vector() - double array,
' Outputs - none
' ---------------------------------------------------------------------------
Public Function MultiplyVectorByMatrix(resultvec() As Double, ByVal matrix As DT4x4, vector() As Double)
Const METHOD = "MultiplyVectorByMatrix"
    On Error GoTo ErrorHandler
    
    Dim I As Integer, J As Integer
    Dim calcvector(4) As Double
        
    For I = 1 To 4
        calcvector(I) = 0
        For J = 1 To 4
            calcvector(I) = calcvector(I) + matrix.IndexValue((J - 1) * 4 + I - 1) * vector(J)
        Next
    Next
    
    For I = 1 To 4
        resultvec(I) = calcvector(I)
    Next
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: RotateVector
' Description: rotate the specified vector by the specified matrix
'
' Inputs - resultvec() - double array, matrix - DT4x4, vector() - double array,
' Outputs - none
' ---------------------------------------------------------------------------
Public Function RotateVector(resultvec() As Double, ByVal matrix As DT4x4, vector() As Double)
Const METHOD = "RotateVector"
    On Error GoTo ErrorHandler
    
    Dim I As Integer, J As Integer
    Dim calcvector(4) As Double
    
    For I = 1 To 4
        calcvector(I) = 0
        For J = 1 To 3
            calcvector(I) = calcvector(I) + matrix.IndexValue((J - 1) * 4 + I - 1) * vector(J)
        Next
    Next
    
    For I = 1 To 4
        resultvec(I) = calcvector(I)
    Next
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: Loc
' Description: Creates a location string from x, y and z coordinates
'
' Inputs - x - double, y - double, z - double
' Outputs - string
' ---------------------------------------------------------------------------
Public Function Loc(x As Double, y As Double, z As Double) As String
Const METHOD = "Loc"
    On Error GoTo ErrorHandler
    
    Loc = "TRANS(DP(" + Str(x) + ", " + Str(y) + ", " + Str(z) + "))"
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: GetPi
' Description: Return an accurate value of Pi
' Apr 27 2006 - Jim Knittig
'
' Inputs - None
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function GetPi() As Double
Const METHOD = "GetPi"
    On Error GoTo ErrorHandler
    
    GetPi = Atn(1) * 4#
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: GetSteelDensityKGPerM
' Description: Return  the steel density in Kilograms per metre so it can be used for weight calculations
' Apr 27 2006 - Jim Knittig
'
' Inputs - None
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function GetSteelDensityKGPerM() As Double
Const METHOD = "GetSteelDensityKGPerM"
    On Error GoTo ErrorHandler
    
    GetSteelDensityKGPerM = 7900
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddRevolution
' Description: Create the graphical representation of an revolution.
'              if the revolution is to be capped
'
' Outputs - none
' ---------------------------------------------------------------------------
Public Function AddRevolution(centerX As Double, centerY As Double, centerZ As Double, _
                              axisX As Double, axisY As Double, axisZ As Double, _
                              ANGLE As Double, capped As Boolean, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddRevolution"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory
    Dim oRevolution As Revolution3d
    
    Set oRevolution = oGeomFactory.Revolutions3d.CreateByCurve(Nothing, _
                                                    m_oContour, axisX, axisY, axisZ, _
                                                    centerX, centerY, centerZ, ANGLE, capped)

    Dim objRemover As IJDObject
    Set objRemover = m_oContour
    objRemover.Remove
    Set m_oContour = Nothing
    
    outputcoll.AddOutput output, oRevolution
    
    Dim I As Integer

    For I = 1 To m_oCrvElems.Count
        Set objRemover = m_oCrvElems(I)
        objRemover.Remove
    Next
    
    m_oCrvElems.Clear
    Set m_oCrvElems = Nothing
    Set oRevolution = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
Public Function AddElbowTr(rodradius As Double, elbowradius As Double, SweepAngle As Double, strmatrix As String, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
Const METHOD = "AddElbowTr"
    On Error GoTo ErrorHandler
    
    If LCase(name) = "none" Then
        name = output
    End If
    
    Dim oGeomFactory As New GeometryFactory
    Dim matrix As IJDT4x4
    Dim elbowCenter(0 To 2) As Double
    Dim elbowNormal(0 To 2) As Double
    Dim circleCenter(0 To 2) As Double
    Dim circleNormal(0 To 2) As Double
    Dim vector(4) As Double

    GetMatrixFromString strmatrix, matrix

    vector(1) = 0
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    elbowCenter(0) = vector(1)
    elbowCenter(1) = vector(2)
    elbowCenter(2) = vector(3)

    vector(1) = 0
    vector(2) = -1
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    elbowNormal(0) = vector(1) - elbowCenter(0)
    elbowNormal(1) = vector(2) - elbowCenter(1)
    elbowNormal(2) = vector(3) - elbowCenter(2)

    vector(1) = elbowradius
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    circleCenter(0) = vector(1)
    circleCenter(1) = vector(2)
    circleCenter(2) = vector(3)

    vector(1) = elbowradius
    vector(2) = 0
    vector(3) = 1
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    circleNormal(0) = vector(1) - circleCenter(0)
    circleNormal(1) = vector(2) - circleCenter(1)
    circleNormal(2) = vector(3) - circleCenter(2)

    ' Create the Bolt Cross section curve to revolve.
    Dim BoltXSection As Circle3d

    Set BoltXSection = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                        circleCenter(0), circleCenter(1), circleCenter(2), _
                                                                        circleNormal(0), circleNormal(1), circleNormal(2), _
                                                                        rodradius)
 

    'Create the Revolution representing the Semi-Circle Section of the UBolt
    Dim Elbow As Revolution3d
 
    Set Elbow = oGeomFactory.Revolutions3d.CreateByCurve(Nothing, _
                                                             BoltXSection, _
                                                             elbowNormal(0), elbowNormal(1), elbowNormal(2), _
                                                            elbowCenter(0), elbowCenter(1), elbowCenter(2), _
                                                            GetPi() * SweepAngle / 180, True)

    Dim objRemover As IJDObject
    Set objRemover = BoltXSection
    objRemover.Remove
    Set BoltXSection = Nothing
   
    'outputcoll.AddOutput output, Elbow

    Set AddElbowTr = Elbow
    
    Set Elbow = Nothing
    Set oGeomFactory = Nothing
 Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
Public Function PlaceCylinderTr(m_outputcoll As IJDOutputCollection, lStartPoint As AutoMath.DPosition, _
                                lEndPoint As AutoMath.DPosition, _
                                lDiameter As Double) As Object

    Const METHOD = "PlaceCylinderTr"
    On Error GoTo ErrorHandler
    
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim dblCylWidth     As Double
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set lStartPoint.x, lStartPoint.y, lStartPoint.z
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set lEndPoint.x - lStartPoint.x, _
                     lEndPoint.y - lStartPoint.y, _
                     lEndPoint.z - lStartPoint.z
    dblCylWidth = circleNormal.LENGTH
    circleNormal.LENGTH = 1

   
  
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        lDiameter / 2)
  
    Set objProjection = geomFactory.Projections3d.CreateByCurve(Nothing, _
                                                        objCircle, _
                                                        circleNormal.x, circleNormal.y, circleNormal.z, _
                                                        dblCylWidth, False)
       ' m_outputcoll.AddOutput "Cylinder", objProjection
           
   
    Set objCircle = Nothing
    
    Set PlaceCylinderTr = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
Public Function PlaceTrCircleByCenter(m_outputcoll As IJDOutputCollection, ByRef centerPoint As AutoMath.DPosition, _
                            ByRef normalVector As AutoMath.DVector, _
                            ByRef Radius As Double) _
                            As IngrGeom3D.Circle3d

    Const METHOD = "PlaceTrCircleByCenter"
    On Error GoTo ErrorHandler
        
    Dim oCircle As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    ' Create Circle object
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            centerPoint.x, centerPoint.y, centerPoint.z, _
                            normalVector.x, normalVector.y, normalVector.z, _
                            Radius)
    
    Set PlaceTrCircleByCenter = oCircle
    Set oCircle = Nothing
    Set geomFactory = Nothing

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function CreChamBoxLine(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                     CenterAxisUpward As Boolean) _
                            As IngrGeom3D.ComplexString3d
                            
    Const METHOD = "CreChamBoxLine"
    On Error GoTo ErrorHandler
                            
    Dim PI As Double
    PI = 4 * Atn(1)
    
                            
     Dim Lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim CP              As New AutoMath.DPosition
    Dim Pt(13)           As New AutoMath.DPosition
  
    
    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
 
    
    Pt(1).Set CP.x + 0.0001 - Width / 2, CP.y + Depth / 2, CP.z
    Pt(2).Set CP.x - 0.0001 + Width / 2, CP.y + Depth / 2, CP.z
    Pt(3).Set CP.x + Width / 2, CP.y + Depth / 2 - 0.0001, CP.z
    Pt(4).Set CP.x + Width / 2, CP.y - Depth / 2 + 0.0001, CP.z
    Pt(5).Set CP.x - 0.0001 + Width / 2, CP.y - Depth / 2, CP.z
    Pt(6).Set CP.x + 0.0001 - Width / 2, CP.y - Depth / 2, CP.z
    Pt(7).Set CP.x - Width / 2, CP.y - Depth / 2 + 0.0001, CP.z
    Pt(8).Set CP.x - Width / 2, CP.y + Depth / 2 - 0.0001, CP.z
    Pt(9).Set CP.x - Width / 2, CP.y + Depth / 2, CP.z
    Pt(10).Set CP.x + Width / 2, CP.y + Depth / 2, CP.z
    Pt(11).Set CP.x + Width / 2, CP.y - Depth / 2, CP.z
    Pt(12).Set CP.x - Width / 2, CP.y - Depth / 2, CP.z
    
If CenterAxisUpward = True Then
    Set oLine = PlaceTrLine(Pt(8), Pt(7))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(7), Pt(6), Pt(12))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(6), Pt(5))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(5), Pt(4), Pt(11))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(4), Pt(3))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(3), Pt(2), Pt(10))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(2), Pt(1))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(1), Pt(8), Pt(9))
    Lines.Add oArc
ElseIf CenterAxisUpward = False Then
    
'    Set oLine = PlaceTrLine(Pt(1), Pt(2))
'    Lines.Add oLine
'    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(3), Pt(10))
'    Lines.Add oArc
'    Set oLine = PlaceTrLine(Pt(3), Pt(4))
'    Lines.Add oLine
'    Set oArc = PlaceTrArcBy3Pts(Pt(4), Pt(5), Pt(11))
'    Lines.Add oArc
'    Set oLine = PlaceTrLine(Pt(5), Pt(6))
'    Lines.Add oLine
'    Set oArc = PlaceTrArcBy3Pts(Pt(6), Pt(7), Pt(12))
'    Lines.Add oArc
'    Set oLine = PlaceTrLine(Pt(7), Pt(8))
'    Lines.Add oLine
'    Set oArc = PlaceTrArcBy3Pts(Pt(8), Pt(1), Pt(9))
'    Lines.Add oArc

    Set oLine = PlaceTrLine(Pt(3), Pt(4))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(4), Pt(5), Pt(11))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(5), Pt(6))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(6), Pt(7), Pt(12))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(7), Pt(8))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(8), Pt(1), Pt(9))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(3), Pt(10))
    Lines.Add oArc

 End If

    Set objCStr = PlaceTrCString(Pt(1), Lines)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransPos As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotation As Double
    Set oDirVector = New AutoMath.DVector
    Set oTransPos = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    
'    If (PlaneofBranch = 0) Then
'        dRotation = 0
'        oDirVector.Set 1, 0, 0
'    ElseIf (PlaneofBranch = PI / 2) Then
'        dRotation = PI / 2
'        oDirVector.Set 0, 0, 1
'    End If
    
    oTransformationMat.Rotate dRotation, oDirVector
    objCStr.Transform oTransformationMat
    
    Set CreChamBoxLine = objCStr
    Set oLine = Nothing
    Set oArc = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.Count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, METHOD, "").Number
    
End Function
