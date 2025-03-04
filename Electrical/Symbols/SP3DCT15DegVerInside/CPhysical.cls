VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         dkl
'   Creation Date:  Monday, Aug 6 2007
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------    -----        ------------------
'   06.Aug.2007     dkl         CR-122817  Created the symbol.
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const NEGLIGIBLE_VALUE = 0.0001
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput As Integer

    Dim parActualWidth As Double
    Dim parActualDepth As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parTangentLength1 = arrayOfInputs(2)
'    parTangentLength2 = arrayOfInputs(3)
              
    iOutput = 0
    Dim dAngle As Double
    dAngle = PI / 12
       
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    Dim parBendRadius As Double
    parBendRadius = oTrayPart.BendRadius
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Dim dHalfDepth As Double
    Dim dHalfWidth As Double
    dHalfDepth = parActualDepth / 2
    dHalfWidth = parActualWidth / 2
    
'   The following Part data Bases are implemented,
'   1) 15 degree Vertical Curved Upward Bend, Symmetrical, specified by Bend Radius and Tangent Length (Default implementation).
'   2) 15 degree Vertical Curved Upward Bend, Asymmetrical, specified by Bend Radius, Tangent Length 1 and Tangent Length 2.
    
    Dim dFacetoCenter1 As Double   'Face to Center along Port 1.
    Dim dFacetoCenter2 As Double   'Face to Center along Port 2.
    Dim dTangentLength1 As Double   'Tangent length along Port 1.
    Dim dTangentLength2 As Double   'Tangent length along Port 2.
       
    Dim lCableTrayPartDataBasis As Integer
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    Select Case lCableTrayPartDataBasis
        
        Case Is <= 1, 127    ' 127 - 15 degree Vertical Curved Upward Bend, Symmetrical, specified by
                             'Bend Radius and Tangent Length (Default implementation).
                             
            ' Retrieve Part properties: TangentLength
            Dim parTangentLength As Double
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dFacetoCenter1 = (parBendRadius + dHalfDepth) * Tan(dAngle / 2) + dTangentLength1
            dFacetoCenter2 = dFacetoCenter1
            
        Case 129     ' 129 - 15 degree Vertical Curved Upward Bend, Asymmetrical, specified by
                     'Bend Radius, Tangent Length 1 and Tangent Length 2.
                     
            dTangentLength1 = arrayOfInputs(2)
            dTangentLength2 = arrayOfInputs(3)
            dFacetoCenter1 = (parBendRadius + dHalfDepth) * Tan(dAngle / 2) + dTangentLength1
            dFacetoCenter2 = (parBendRadius + dHalfDepth) * Tan(dAngle / 2) + dTangentLength2
                      
        Case Else
            GoTo ErrorLabel:
    
    End Select
    
'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(dTangentLength1, 0) Then dTangentLength1 = NEGLIGIBLE_VALUE
    If CmpDblLessThanOrEqualTo(dTangentLength2, 0) Then dTangentLength2 = NEGLIGIBLE_VALUE
    
' Insert your code for output 1(Horizontal Tangent)
    Dim oPort1 As AutoMath.DPosition 'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    oPort1.Set -dFacetoCenter1, 0, 0

    Dim LineStrPoints()  As Double
    ReDim LineStrPoints(0 To 11)
    
    LineStrPoints(0) = oPort1.x
    LineStrPoints(1) = oPort1.y - dHalfWidth
    LineStrPoints(2) = oPort1.z + dHalfDepth
    
    LineStrPoints(3) = oPort1.x
    LineStrPoints(4) = oPort1.y - dHalfWidth
    LineStrPoints(5) = oPort1.z - dHalfDepth
    
    LineStrPoints(6) = oPort1.x
    LineStrPoints(7) = oPort1.y + dHalfWidth
    LineStrPoints(8) = oPort1.z - dHalfDepth
    
    LineStrPoints(9) = oPort1.x
    LineStrPoints(10) = oPort1.y + dHalfWidth
    LineStrPoints(11) = oPort1.z + dHalfDepth
    
    Dim oProjVector As AutoMath.DVector
    Set oProjVector = New AutoMath.DVector
    oProjVector.Set 1, 0, 0
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oLineString As IngrGeom3D.LineString3d
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    
    Dim objHoriTangent As Object
    Set objHoriTangent = PlaceProjection(m_OutputColl, oLineString, oProjVector, _
                                                            dTangentLength1, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHoriTangent
    Set objHoriTangent = Nothing

' Insert your code for output 2(Vertical Bend)
    Dim oLineStrCP As AutoMath.DPosition   'Elbow begins at this point.
    Set oLineStrCP = New AutoMath.DPosition
    oLineStrCP.Set -dFacetoCenter1 + dTangentLength1, 0, 0
               
    LineStrPoints(0) = oLineStrCP.x
    LineStrPoints(1) = oLineStrCP.y - dHalfWidth
    LineStrPoints(2) = oLineStrCP.z + dHalfDepth
    
    LineStrPoints(3) = oLineStrCP.x
    LineStrPoints(4) = oLineStrCP.y - dHalfWidth
    LineStrPoints(5) = oLineStrCP.z - dHalfDepth
    
    LineStrPoints(6) = oLineStrCP.x
    LineStrPoints(7) = oLineStrCP.y + dHalfWidth
    LineStrPoints(8) = oLineStrCP.z - dHalfDepth
    
    LineStrPoints(9) = oLineStrCP.x
    LineStrPoints(10) = oLineStrCP.y + dHalfWidth
    LineStrPoints(11) = oLineStrCP.z + dHalfDepth
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, _
                                                                    LineStrPoints)
    
    oProjVector.Set 0, -1, 0
    Dim oCenterPoint As AutoMath.DPosition 'arc center point
    Set oCenterPoint = New AutoMath.DPosition
    
    oCenterPoint.Set oLineStrCP.x, 0, parBendRadius + dHalfDepth
    
    Dim objVertiCurve As Object
    Set objVertiCurve = PlaceRevolution(m_OutputColl, oLineString, oProjVector, _
                            oCenterPoint, dAngle, False)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiCurve
    Set objVertiCurve = Nothing
    Set oCenterPoint = Nothing
    Set oLineStrCP = Nothing

' Insert your code for output 3(Vertical Tangent)
    Dim oPort2 As AutoMath.DPosition 'Port 2 center point
    Set oPort2 = New AutoMath.DPosition
    oPort2.Set dFacetoCenter2 * Cos(dAngle), 0, dFacetoCenter2 * Sin(dAngle)
    
    LineStrPoints(0) = oPort2.x - dHalfDepth * Sin(dAngle)
    LineStrPoints(1) = oPort2.y - dHalfWidth
    LineStrPoints(2) = oPort2.z + dHalfDepth * Cos(dAngle)
    
    LineStrPoints(3) = oPort2.x + dHalfDepth * Sin(dAngle)
    LineStrPoints(4) = oPort2.y - dHalfWidth
    LineStrPoints(5) = oPort2.z - dHalfDepth * Cos(dAngle)
    
    LineStrPoints(6) = oPort2.x + dHalfDepth * Sin(dAngle)
    LineStrPoints(7) = oPort2.y + dHalfWidth
    LineStrPoints(8) = oPort2.z - dHalfDepth * Cos(dAngle)
    
    LineStrPoints(9) = oPort2.x - dHalfDepth * Sin(dAngle)
    LineStrPoints(10) = oPort2.y + dHalfWidth
    LineStrPoints(11) = oPort2.z + dHalfDepth * Cos(dAngle)
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    oProjVector.Set -Cos(dAngle), 0, -Sin(dAngle)
    
    Dim objInclinedTangent As Object
    Set objInclinedTangent = PlaceProjection(m_OutputColl, oLineString, oProjVector, _
                                                                dTangentLength2, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInclinedTangent
    Set objInclinedTangent = Nothing
    Set oGeomFactory = Nothing
    Set oLineString = Nothing
    Set oProjVector = Nothing
    
' Place Nozzle 1
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
    
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPort1 = Nothing
    Set oPortLocation = Nothing
    
' Place Nozzle 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oDir.Set Cos(dAngle), 0, Sin(dAngle)
    oRadialOrient.Set -Sin(dAngle), 0, Cos(dAngle)
    
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPort2 = Nothing
    Set oPortLocation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
