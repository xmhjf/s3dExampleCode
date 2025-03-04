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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          KKC
'   Creation Date:  Jan 10, 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This reducing coupling symbol is developed from the Victaulic general catalog. Pg 24 Style 750
'   Change History:
'   dd.mmm.yyyy           who                      change description
'   -----------           -----                      ------------------
'   10.JAN.2007            KKC       Created :CR-112137 Provide additional symbols required for Victaulic General Catalog
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

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

    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjBoltEar As Object
   
    Dim parFacetoFace As Double
    Dim parDiscWidth As Double
    Dim parDiscDiamter As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parDiscWidth = arrayOfInputs(3)
    parDiscDiamter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0
    
' Insert your code for output 1(Body)
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    oStPoint.Set -parFacetoFace / 2, 0, 0
    oEnPoint.Set parFacetoFace / 2, 0, 0
 
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                              flangeThick, flangeDiam, cptOffset, depth
                              
    Set ObjBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiscDiamter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Insert your code for output 2(Bolt Ear 1)
    Dim oProjVector As AutoMath.DVector
    Set oProjVector = New AutoMath.DVector
    
    oProjVector.Set 0, 1, 0
    Dim dBoltEarHeight As Double
    Dim dBoltEarLength As Double
    Dim dArcHeight  As Double
    
' Assuming Bolt Ear Height as 0.45 times the Disc Diameter
    dBoltEarHeight = 0.45 * parDiscDiamter
    dBoltEarLength = (parDiscWidth - 1.1 * pipeDiam) / 2
    dArcHeight = 0.3 * dBoltEarLength
    
    Dim oPoint1 As AutoMath.DPosition
    Dim oPoint2 As AutoMath.DPosition
    Dim oPoint3 As AutoMath.DPosition
    Dim oPoint4 As AutoMath.DPosition
    Dim oPoint5 As AutoMath.DPosition
    
    Set oPoint1 = New AutoMath.DPosition
    Set oPoint2 = New AutoMath.DPosition
    Set oPoint3 = New AutoMath.DPosition
    Set oPoint4 = New AutoMath.DPosition
    Set oPoint5 = New AutoMath.DPosition
    
    oPoint1.x = 0.5 * parFacetoFace
    oPoint1.y = -0.5 * dBoltEarHeight
    oPoint1.z = parDiscWidth / 2 - dBoltEarLength
    
    oPoint2.x = 0.5 * parFacetoFace
    oPoint2.y = -0.5 * dBoltEarHeight
    oPoint2.z = parDiscWidth / 2 - dArcHeight
    
    oPoint3.x = -0.5 * parFacetoFace
    oPoint3.y = -0.5 * dBoltEarHeight
    oPoint3.z = parDiscWidth / 2 - dArcHeight
    
    oPoint4.x = -0.5 * parFacetoFace
    oPoint4.y = -0.5 * dBoltEarHeight
    oPoint4.z = parDiscWidth / 2 - dBoltEarLength
    
    oPoint5.x = 0
    oPoint5.y = -0.5 * dBoltEarHeight
    oPoint5.z = parDiscWidth / 2
    
    Dim ObjBottomCurves As Collection
    Set ObjBottomCurves = New Collection
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oLine1 As IngrGeom3D.Line3d
    Set oLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, _
                                                            oPoint2.x, oPoint2.y, oPoint2.z)
    ObjBottomCurves.Add oLine1
    Dim oArc1 As IngrGeom3D.Arc3d
    Set oArc1 = PlaceTrArcBy3Pts(oPoint2, oPoint3, oPoint5)
    
    ObjBottomCurves.Add oArc1
    Dim oLine2 As IngrGeom3D.Line3d
    Set oLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oPoint3.x, oPoint3.y, oPoint3.z, _
                                                            oPoint4.x, oPoint4.y, oPoint4.z)
    ObjBottomCurves.Add oLine2
    Dim oLine3 As IngrGeom3D.Line3d
    Set oLine3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oPoint4.x, oPoint4.y, oPoint4.z, _
                                                            oPoint1.x, oPoint1.y, oPoint1.z)
                                                            
    ObjBottomCurves.Add oLine3
    Dim objEBcurves As IJDObject
    Set objEBcurves = PlaceTrCString(oPoint1, ObjBottomCurves)
    
    Set ObjBoltEar = PlaceProjection(m_OutputColl, objEBcurves, oProjVector, dBoltEarHeight, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBoltEar
    Set ObjBoltEar = Nothing
    
' Insert your code for output 3(Bolt Ear 2)
    oPoint1.x = 0.5 * parFacetoFace
    oPoint1.y = -0.5 * dBoltEarHeight
    oPoint1.z = dBoltEarLength - parDiscWidth / 2
    
    oPoint2.x = 0.5 * parFacetoFace
    oPoint2.y = -0.5 * dBoltEarHeight
    oPoint2.z = dArcHeight - parDiscWidth / 2

    oPoint3.x = -0.5 * parFacetoFace
    oPoint3.y = -0.5 * dBoltEarHeight
    oPoint3.z = dArcHeight - parDiscWidth / 2
    
    oPoint4.x = -0.5 * parFacetoFace
    oPoint4.y = -0.5 * dBoltEarHeight
    oPoint4.z = dBoltEarLength - parDiscWidth / 2
    
    oPoint5.x = 0
    oPoint5.y = -0.5 * dBoltEarHeight
    oPoint5.z = -parDiscWidth / 2
    
    Set ObjBottomCurves = New Collection
    Set oLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, _
                                                            oPoint2.x, oPoint2.y, oPoint2.z)
    ObjBottomCurves.Add oLine1
    Set oArc1 = PlaceTrArcBy3Pts(oPoint2, oPoint3, oPoint5)
    
    ObjBottomCurves.Add oArc1
    Set oLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oPoint3.x, oPoint3.y, oPoint3.z, _
                                                            oPoint4.x, oPoint4.y, oPoint4.z)
    ObjBottomCurves.Add oLine2
    Set oLine3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oPoint4.x, oPoint4.y, oPoint4.z, _
                                                            oPoint1.x, oPoint1.y, oPoint1.z)
                                                            
    ObjBottomCurves.Add oLine3
    Set objEBcurves = PlaceTrCString(oPoint1, ObjBottomCurves)
    Dim ObjBoltEar1 As Object
    
    Set ObjBoltEar1 = PlaceProjection(m_OutputColl, objEBcurves, oProjVector, dBoltEarHeight, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBoltEar1
    Set ObjBoltEar1 = Nothing
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoFace / 2 - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLine1 = Nothing
    Set oLine2 = Nothing
    Set oLine3 = Nothing
    Set oArc1 = Nothing
    Set oPoint1 = Nothing
    Set oPoint2 = Nothing
    Set oPoint3 = Nothing
    Set oPoint4 = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
