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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:44 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008-10 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:    VRK
'   Creation Date: MonDay,June 2 2008
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   02.June.2008     VRK     CR-141836:Provide new electrical equipment symbols for lighting fixtures
'   03.Feb.2010      KKC     TR-157988  TO Do List entries are generated whern placing lighting fixtures
'                            (Modified nozzle code to create with placeholder)
'   25.Mar.2010     RUK     CR-CP-177179  Enhance lighting symbols to provide control point for dimensioning purposes
'                           CR-CP-113328  "can be modified" flag should be True for all best practice equipment symbols
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    PI = Atn(1) * 4

    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Integer
    
    Dim parPoleHeight As Double
    Dim parPoleDiameter1 As Double
    Dim parPoleDiameter2 As Double
    Dim parAngle As Double
    Dim parwattage As Double
    Dim parMountPlateLength As Double
    Dim parMountPlateWidth As Double
    Dim parMountPlateThick As Double
    Dim parExtensionLength As Double
    Dim parFixtureLength As Double
    Dim parFixtureWidth As Double
    Dim parFixtureDepth As Double
    Dim parCPx As Double
    Dim parCPy As Double
    Dim parCPz As Double
    
    Dim dPoleRadius As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oNormal As AutoMath.DVector
    Dim oAxisvec As New AutoMath.DVector

    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCenter = New AutoMath.DPosition
    Set oNormal = New DVector
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPoleHeight = arrayOfInputs(2)        'A
    parPoleDiameter1 = arrayOfInputs(3)     'D1
    parPoleDiameter2 = arrayOfInputs(4)     'D2
    parAngle = arrayOfInputs(5)             'C
    parMountPlateLength = arrayOfInputs(6)  'PL
    parMountPlateWidth = arrayOfInputs(7)   'PW
    parMountPlateThick = arrayOfInputs(8)   'PT
    parExtensionLength = arrayOfInputs(9)   'B
    parFixtureLength = arrayOfInputs(10)    'L
    parFixtureWidth = arrayOfInputs(11)     'W
    parFixtureDepth = arrayOfInputs(12)     'D
    parwattage = arrayOfInputs(13)
    
    If UBound(arrayOfInputs) > 13 Then
        parCPx = arrayOfInputs(14)
        parCPy = arrayOfInputs(15)
        parCPz = arrayOfInputs(16)
    End If

    iOutput = 0
    
    dPoleRadius = parPoleDiameter1 / 2
    
    'check for Angle
    If CmpDblGreaterthanOrEqualTo(parAngle, PI / 2) Then GoTo ErrorLabel
    
    'Create the Default Surface at the origin
     'Create non-persistent circle to use for creating default surface ---
    Dim objPlane As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                              0, 0, 0, _
                             0, 0, -1, dPoleRadius)

    'Create persistent default surface plane - the plane can mate ---
    Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objPlane = Nothing
    Set objCircle = Nothing
    
   'Create Mount Plate
   
    Dim dMountPlatePoints(0 To 14) As Double
    Dim oLineString As New IngrGeom3D.LineString3d
    Dim objMountPlate As Object
    
    dMountPlatePoints(0) = -0.5 * parMountPlateLength
    dMountPlatePoints(1) = -0.5 * parMountPlateWidth
    dMountPlatePoints(2) = 0
    dMountPlatePoints(3) = 0.5 * parMountPlateLength
    dMountPlatePoints(4) = dMountPlatePoints(1)
    dMountPlatePoints(5) = 0
    dMountPlatePoints(6) = dMountPlatePoints(3)
    dMountPlatePoints(7) = dMountPlatePoints(1)
    dMountPlatePoints(8) = parMountPlateThick
    dMountPlatePoints(9) = dMountPlatePoints(0)
    dMountPlatePoints(10) = dMountPlatePoints(1)
    dMountPlatePoints(11) = dMountPlatePoints(8)
    dMountPlatePoints(12) = dMountPlatePoints(0)
    dMountPlatePoints(13) = dMountPlatePoints(1)
    dMountPlatePoints(14) = dMountPlatePoints(2)
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                        5, dMountPlatePoints)
    oAxisvec.Set 0, 1, 0
    Set objMountPlate = PlaceProjection(m_OutputColl, oLineString, oAxisvec, parMountPlateWidth, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMountPlate
    Set objMountPlate = Nothing
    
    'Create Output 3(Pole Body)
    Dim objPole As Object
    oStPoint.Set 0, 0, parMountPlateThick
    oEnPoint.Set 0, 0, oStPoint.z + parPoleHeight
    Set objPole = PlaceCone(m_OutputColl, oStPoint, oEnPoint, 0.5 * parPoleDiameter1, 0.5 * parPoleDiameter2, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPole
    Set objPole = Nothing
    
    'Assuming ArcRadius as 1/10 times PoleHeight
    Dim oCollection As Collection
    Set oCollection = New Collection
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim oTraceStr As New IngrGeom3D.ComplexString3d
    Dim Surfset As IngrGeom3D.IJElements
    Dim oLine As Object
    Dim oArc As Object
    Dim oCircle As New IngrGeom3D.Circle3d
    Dim ObjLightFix As Object
    
    Dim dArcRadius As Double
    dArcRadius = parPoleHeight / 10

    oStPoint.Set 0.5 * parPoleDiameter2, 0, parPoleHeight
    oEnPoint.Set oStPoint.x + dArcRadius * (1 - Cos(PI / 2 - parAngle)), 0, parPoleHeight + dArcRadius * Sin(PI / 2 - parAngle)
    oCenter.Set dPoleRadius + dArcRadius, 0, oStPoint.z
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                     oCenter.x, oCenter.y, oCenter.z, _
                     oStPoint.x, oStPoint.y, oStPoint.z, _
                     oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing
    
    'Line at light
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x + (parExtensionLength - parFixtureLength) * Cos(parAngle), 0, oStPoint.z + (parExtensionLength - parFixtureLength) * Sin(parAngle)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                    oStPoint.x, oStPoint.y, oStPoint.z, _
                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
        
    oCenter.Set 0, 0, parPoleHeight
    oNormal.Set 0, 0, 1
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, 0.5 * parPoleDiameter2)

    oStPoint.Set 0.5 * parPoleDiameter2, 0, parPoleHeight
    Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                  CircularCorner, 0, stnorm, ednorm, False)
    For Each ObjLightFix In Surfset
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLightFix
    Next ObjLightFix

    'Remove References
    Dim iCount As Integer
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    For iCount = 1 To Surfset.Count
        Surfset.Remove 1
    Next iCount
    For iCount = 1 To oCollection.Count
        oTraceStr.RemoveCurve True
    Next iCount
    Set oTraceStr = Nothing
    Set ObjLightFix = Nothing
    Set oCollection = Nothing
    Set Surfset = Nothing
    
    'Create Output for Light
    Dim oCS As IngrGeom3D.ComplexString3d
    Dim oTrace1 As IngrGeom3D.ComplexString3d
    Dim oTrace2 As IngrGeom3D.ComplexString3d
    Dim oEllipArc As New IngrGeom3D.EllipticalArc3d
    Dim dMMRatio As Double
    Dim csObj As IngrGeom3D.IJElements
    Dim trObj As IngrGeom3D.IJElements
    Dim oCscoll As Collection
    Set oCollection = New Collection
    Set oCscoll = New Collection
        
    'Cross Section
    oStPoint.Set 0, 0.5 * parFixtureWidth, 0.4 * parFixtureDepth
    oEnPoint.Set oStPoint.x, 0.5 * parFixtureWidth, -0.4 * parFixtureDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                oStPoint.x, oStPoint.y, oStPoint.z, _
                oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCscoll.Add oLine
    Set oLine = Nothing

    'Arc1
    dMMRatio = (0.1 * parFixtureDepth) / (0.5 * parFixtureWidth)
    oCenter.Set oStPoint.x, 0, -0.4 * parFixtureDepth
    Set oEllipArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    oCenter.x, oCenter.y, oCenter.z, -1, 0, 0, _
                    0, 0.5 * parFixtureWidth, 0, dMMRatio, 0, PI)
    oCscoll.Add oEllipArc
    Set oEllipArc = Nothing

    oStPoint.Set oStPoint.x, -0.5 * parFixtureWidth, -0.4 * parFixtureDepth
    oEnPoint.Set oStPoint.x, -0.5 * parFixtureWidth, 0.4 * parFixtureDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                oStPoint.x, oStPoint.y, oStPoint.z, _
                oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCscoll.Add oLine
    Set oLine = Nothing

    'Arc2
    oCenter.Set oStPoint.x, 0, 0.4 * parFixtureDepth
    Set oEllipArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    oCenter.x, oCenter.y, oCenter.z, -1, 0, 0, _
                    0, -0.5 * parFixtureWidth, 0, dMMRatio, 0, PI)
    oCscoll.Add oEllipArc
    Set oEllipArc = Nothing

    oStPoint.Set oStPoint.x, 0.5 * parFixtureWidth, 0.4 * parFixtureDepth
    Set oCS = PlaceTrCString(oStPoint, oCscoll)
    Set csObj = New JObjectCollection
    csObj.Add oCS

    'Trace 1
    oStPoint.Set -0.5 * parFixtureLength, 0, -0.25 * parFixtureDepth
    oEnPoint.Set -0.25 * parFixtureLength, 0, -0.5 * parFixtureDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                oStPoint.x, oStPoint.y, oStPoint.z, _
                oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set -0.25 * parFixtureLength, 0, -0.5 * parFixtureDepth
    oEnPoint.Set 0.5 * parFixtureLength, 0, -0.5 * parFixtureDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                oStPoint.x, oStPoint.y, oStPoint.z, _
                oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing

    oStPoint.Set -0.5 * parFixtureLength, 0, -0.25 * parFixtureDepth
    Set oTrace1 = PlaceTrCString(oStPoint, oCollection)
    Set trObj = New JObjectCollection

    'Trace 2
    Dim oTopCurveColl As New Collection
    oStPoint.Set -0.5 * parFixtureLength, 0, 0.25 * parFixtureDepth
    oEnPoint.Set 0.25 * parFixtureLength, 0, 0.5 * parFixtureDepth
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                oStPoint.x, oStPoint.y, oStPoint.z, _
                oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oTopCurveColl.Add oLine
    Set oLine = Nothing
    
    dMMRatio = (0.25 * parFixtureLength) / (0.5 * parFixtureDepth)
    oCenter.Set 0.25 * parFixtureLength, 0, 0
    Set oEllipArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    oCenter.x, oCenter.y, oCenter.z, 0, 1, 0, _
                    0, 0, 0.5 * parFixtureDepth, dMMRatio, 0, PI / 2)
    oTopCurveColl.Add oEllipArc
    Set oEllipArc = Nothing
    
    oStPoint.Set -0.5 * parFixtureLength, 0, 0.25 * parFixtureDepth
    Set oTrace2 = PlaceTrCString(oStPoint, oTopCurveColl)
    trObj.Add oTrace1
    trObj.Add oTrace2

    'Create Skinning Surface with 2 Traces and 1 Cross section
    Set Surfset = oGeomFactory.GeometryServices.CreateBySkinning(m_OutputColl.ResourceManager, _
     trObj, csObj, True)
    ' Set the output
    For Each ObjLightFix In Surfset
        Dim oTransMatrix As IJDT4x4
        Set oTransMatrix = New DT4x4
        oTransMatrix.LoadIdentity
        oAxisvec.Set 0, -1, 0
        oTransMatrix.Rotate parAngle, oAxisvec
        oTransMatrix.IndexValue(12) = 0.5 * parPoleDiameter2 + dArcRadius * (1 - Cos(PI / 2 - parAngle)) + (parExtensionLength - parFixtureLength) * Cos(parAngle)
        oTransMatrix.IndexValue(13) = 0
        oTransMatrix.IndexValue(14) = parPoleHeight + dArcRadius * Sin(PI / 2 - parAngle) + (parExtensionLength - parFixtureLength) * Sin(parAngle)
        ObjLightFix.Transform oTransMatrix
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLightFix
    Next ObjLightFix
    Set Surfset = Nothing
    Set ObjLightFix = Nothing
    For iCount = 1 To oCscoll.Count
        oCscoll.Remove 1
    Next iCount
    For iCount = 1 To csObj.Count
        oCS.RemoveCurve True
    Next iCount
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    For iCount = 1 To oTopCurveColl.Count
        oTopCurveColl.Remove 1
    Next iCount
    For iCount = 1 To trObj.Count
        oTrace1.RemoveCurve True
        oTrace2.RemoveCurve True
    Next iCount
    
    ' Insert your code for output 8(Conduit Port 1)
    Dim oConduitPortPoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oConduitPortPoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oDir.Set 0, 0, -1
    oConduitPortPoint.Set 0, 0, 0
    
    Dim ObjConduitPort1 As IJConduitPortOcc
    Set ObjConduitPort1 = CreateConduitNozzlePH(oConduitPortPoint, oDir, m_OutputColl, oPartFclt, 1)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort1
    Set ObjConduitPort1 = Nothing
     
    ' Insert your code for output 9(Cable Port 1)
    Dim oCableTrayPortPoint As AutoMath.DPosition
    Set oCableTrayPortPoint = New AutoMath.DPosition
    
    Dim oRadialOrient As AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    oRadialOrient.Set 0, -1, 0
    
    Dim objCableNozzle As CableNozzle
    Dim iDistribPort As IJDistribPort
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    
    Set objCableNozzle = oNozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                            m_OutputColl.ResourceManager)
                                
    Set iLogicalDistPort = objCableNozzle
    iLogicalDistPort.SetCenterLocation oCableTrayPortPoint
    Set iDistribPort = objCableNozzle
    iDistribPort.SetDirectionVector oDir

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableNozzle
    Set objCableNozzle = Nothing
        
    Set oCscoll = Nothing
    Set oCS = Nothing
    Set trObj = Nothing
    Set csObj = Nothing
    Set oTrace1 = Nothing
    Set oTrace2 = Nothing
    Set oCollection = Nothing
    Set oTopCurveColl = Nothing
    Set oGeomFactory = Nothing
    Set oAxisvec = Nothing
    Set oGeomFactory = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oNormal = Nothing
    Set oCircle = Nothing
    Set oNormal = Nothing
    Set oTransMatrix = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub


