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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Thursday, Nov 23 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   09.Jul.2003     SymbolTeam(India)  Copyright Information, Header  is added.
'   02.Mar.2005     svsmylav           TR-52998: X co-ordinate of Port 1 is corrected from zero to (depth - cptOffset)
'   08.SEP.2006     KKC                DI-95670  Replace names with initials in all revision history sheets and symbols.
'   02.May.2007     dkl                CR 116914 Implemented part data basis to support Oval,
'                                                Square and Triangluar drilling templates.
'   11.May.2007     dkl                TR-120099 Streamlined all the cases of Part Data Basis to produce
'                                                identical number of outputs.
'   02.Nov.2007     dkl                TR 130675, Additonal outputs added in V7SP2 are added as variable outputs.
'   3.NOV.2007      RRK                CR-123952 Updated the symbol to support part data basis values of 211, 216, 221, 226, 231, 236, 306 and 311
'   6.Dec.2007      RRK                TR-133109 Removed the changes made for TR 130675, as it is not closed.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Dim PI As Double

Private Sub Class_Initialize()
    PI = Atn(1) * 4
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double
    Dim cptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim objFlangeBody As Object
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    
    Dim dFlangeLength As Double
    Dim dFlangeWidth As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
'Assigning the boolean parameter(which will be passed to RetrieveParameters function)
    'an appropriate value based on part data basis
    Dim blnLengthThrHubInclFaceProj As Boolean
    Select Case lPartDataBasis 'Cases where Length-through-hub dimension includes flange face projection
     Case Is <= 1, 210, 215, 220, 225, 230, 235, 305, 310
        blnLengthThrHubInclFaceProj = True
     Case 211, 216, 221, 226, 231, 236, 306, 311 'Cases where Length-through-hub dimension excludes flange face projection
        blnLengthThrHubInclFaceProj = False
    End Select
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        cptOffset, depth, blnLengthThrHubInclFaceProj
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                        cptOffset2, depth2, blnLengthThrHubInclFaceProj
    dFlangeLength = flangeDiam

'This symbol can be associated with the following Part data basis values
'1. Flange, circular drilling template, length though hub includes flange face projection - 210 (The default case)
'2. Flange, circular drilling template, length though hub excludes flange face projection -211
'3. Flange, oval drilling template, length though hub includes flange face projection - 215
'4. Flange, oval drilling template, length though hub excludes flange face projection -216
'5. Flange, square drilling template, length though hub includes flange face projection - 220
'6. Flange, square drilling template, length though hub excludes flange face projection-221
'7. Flange, triangular drilling template, length though hub includes flange face projection - 235.
'8. Flange, triangular drilling template, length though hub excludes flange face projection-236
    
    Select Case lPartDataBasis
        
        Case Is <= 1, 210, 211  'The default state, Weldneck flange, with circular drilling template.
            oStPoint.Set 0, 0, 0
            oEndPoint.Set flangeThick, 0, 0
            
            Set objFlangeBody = PlaceCylinder(m_OutputColl, oStPoint, _
                                                oEndPoint, dFlangeLength, True)
                                                
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFlangeBody
    
        Case 215, 216 'Weldneck flange, with oval drilling template.
            dFlangeWidth = 0.4 * dFlangeLength  'Assumtion Flange Width is 40% of its height
                        
            Dim dPoints() As Double
            ReDim dPoints(0 To 17) 'representing points in the Y-Z plane
            'Point 1
            dPoints(0) = 0      'X
            dPoints(1) = 0      'Y
            dPoints(2) = -dFlangeWidth * 0.5        'Z
            'Point 2
            dPoints(3) = dPoints(0)
            dPoints(4) = dFlangeLength * 0.4
            dPoints(5) = -dFlangeWidth * 0.2
            'Point 3
            dPoints(6) = dPoints(0)
            dPoints(7) = dPoints(4)
            dPoints(8) = -dPoints(5)
            'Point 4
            dPoints(9) = dPoints(0)
            dPoints(10) = dPoints(1)
            dPoints(11) = -dPoints(2)
            'Point 5
            dPoints(12) = dPoints(6)
            dPoints(13) = -dPoints(7)
            dPoints(14) = dPoints(8)
            'Point 6
            dPoints(15) = dPoints(3)
            dPoints(16) = -dPoints(4)
            dPoints(17) = dPoints(5)
                   
            Dim oGeomFactory As IngrGeom3D.GeometryFactory
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            
            Dim oLine As IngrGeom3D.Line3d
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dPoints(0), dPoints(1), dPoints(2), _
                        dPoints(3), dPoints(4), dPoints(5))
            
            Dim objCollection As Collection
            Set objCollection = New Collection
            objCollection.Add oLine
            
            Dim oArc As IngrGeom3D.Arc3d
            'The y Co-ordinate of the center of the arc is computed as dPoints(4) - Y.
            'Where Y = dPoints(8) x Tan(Sin Inverse (A) ), A = (dPoints(11)-dPoints(8)) / dPoints(4)
            'Therefore, Y = dPoints(8) x Tan(Tan Inverse( A / (1-AxA) ))
            Dim dRatio As Double
            dRatio = (dPoints(11) - dPoints(8)) / dPoints(4)
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                        0, dPoints(4) - dPoints(8) * Tan(Atn(dRatio / (1 - dRatio * dRatio))), 0, _
                        dPoints(3), dPoints(4), dPoints(5), _
                        dPoints(6), dPoints(7), dPoints(8))
            objCollection.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dPoints(6), dPoints(7), dPoints(8), _
                        dPoints(9), dPoints(10), dPoints(11))
            objCollection.Add oLine
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dPoints(9), dPoints(10), dPoints(11), _
                        dPoints(12), dPoints(13), dPoints(14))
            objCollection.Add oLine
            
            'This Arc is mirror image of the above arc.
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                        0, -(dPoints(4) - dPoints(8) * Tan(Atn(dRatio / (1 - dRatio * dRatio)))), 0, _
                        dPoints(12), dPoints(13), dPoints(14), _
                        dPoints(15), dPoints(16), dPoints(17))
            objCollection.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dPoints(15), dPoints(16), dPoints(17), _
                        dPoints(0), dPoints(1), dPoints(2))
            objCollection.Add oLine
            
            oStPoint.Set dPoints(0), dPoints(1), dPoints(2)
            Dim objFlangeOutline As IngrGeom3D.ComplexString3d
            Set objFlangeOutline = PlaceTrCString(oStPoint, objCollection)
            
            'Project the Flange Upper Outline along X axis with length Flange body Thickness.
            Set objFlangeBody = oGeomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                    objFlangeOutline, 1, 0, 0, flangeThick, True)
                                    
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFlangeBody
            Set oLine = Nothing
            Set oArc = Nothing
            Set objCollection = Nothing
            Set objFlangeOutline = Nothing
            Set oGeomFactory = Nothing
            
        Case 220, 221 'Weldneck flange, with square drilling template.
            dFlangeWidth = dFlangeLength
            oStPoint.Set 0, -dFlangeLength / 2, -dFlangeWidth / 2
            oEndPoint.Set flangeThick, dFlangeLength / 2, dFlangeWidth / 2
            
            Set objFlangeBody = PlaceBox(m_OutputColl, oStPoint, oEndPoint)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFlangeBody
            
        'Case 225,226 'Weldneck flange, with rectangular drilling template.
            'RetrieveBoltPatternDimensions 1, oPartFclt, dBoltPatternLength, dBoltPatternWidth
                      
        'Case 230,231 'Weldneck flange, with rectangular drilling template, split flange face.
            'RetrieveBoltPatternDimensions 1, oPartFclt, dBoltPatternLength, dBoltPatternWidth
            
        Case 235, 236 'Weldneck flange, with triangular drilling template.
            'The center is at the centroid of the equilateral triangle
            'Side of the triangle is equal to Flange Diameter.
            
            ReDim dPoints(0 To 8)  'representing points in the Y-Z plane
            'Point 1
            dPoints(0) = 0      'X
            'The height of the traingular flange is flange dia x cos (30deg).
            'The centroid divides the total height of the triangle in a way that the height
            'from centroid to the base is one-third of the total height of the triangle.
            dPoints(1) = -dFlangeLength * Cos(PI / 6) / 3 'Y
            dPoints(2) = -dFlangeLength * 0.5        'Z
            'Point 2
            dPoints(3) = dPoints(0)
            dPoints(4) = dPoints(1)
            dPoints(5) = -dPoints(2)
            'Point 3
            dPoints(6) = dPoints(0)
            dPoints(7) = dFlangeLength * Cos(PI / 6) * 2 / 3
            dPoints(8) = 0

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    dPoints(0), dPoints(1), dPoints(2), _
                                    dPoints(3), dPoints(4), dPoints(5))
            Set objCollection = New Collection
            objCollection.Add oLine
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    dPoints(3), dPoints(4), dPoints(5), _
                                    dPoints(6), dPoints(7), dPoints(8))
            objCollection.Add oLine
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    dPoints(6), dPoints(7), dPoints(8), _
                                    dPoints(0), dPoints(1), dPoints(2))
            objCollection.Add oLine
                
            oStPoint.Set dPoints(0), dPoints(1), dPoints(2)
    
            Set objFlangeOutline = PlaceTrCString(oStPoint, objCollection)
            
            'Project the Flange Outline along X axis with length Flange body Thickness.
            Set objFlangeBody = oGeomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                    objFlangeOutline, 1, 0, 0, flangeThick, True)
                                    
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFlangeBody
            Set oLine = Nothing
            Set oGeomFactory = Nothing
            Set objCollection = Nothing
            Set objFlangeOutline = Nothing
            
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
        
    End Select
    
    Set objFlangeBody = Nothing

'Create cylinderical neck joining Flange and Socket.
    
    Dim dCylLength As Double
    dCylLength = parFacetoFace - flangeThick2 - flangeThick
    
    'Avoid failure of cylinder placement by preventing zero length of cylinder.
    If CmpDblLessThanOrEqualTo(dCylLength, 0) Then dCylLength = 0.0001
    
    oStPoint.Set flangeThick, 0, 0
    oEndPoint.Set flangeThick + dCylLength, 0, 0
    
    Dim objCylindericalNeck As Object
    Set objCylindericalNeck = PlaceCylinder(m_OutputColl, oStPoint, _
                                        oEndPoint, pipeDiam, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylindericalNeck
    Set objCylindericalNeck = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing

' Place Nozzle 1
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint       As New AutoMath.DPosition

    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
    oPlacePoint.Set (depth - cptOffset), 0, 0
                                                            
    Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    
    oPlacePoint.Set parFacetoFace - (depth2 - cptOffset2), 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

Private Sub RetrieveBoltPatternDimensions(ByVal index As Integer, _
        ByRef partInput As PartFacelets.IJDPart, ByRef BoltPatternLength As Double, _
         ByRef BoltPatternWidth, Optional ByRef BoltPatternOffset)
    
    Const METHOD = "RetrieveBoltPatternDimensions"
    
    'The attributes BoltPatternLength, BoltPatternWidth and BoltPatternOffset are presently
    'not accesible from the symbol. Once this functionality is available, this sub can
    'be implemented to support Rectangular, Rectangular with split flange face drilling templates.
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
