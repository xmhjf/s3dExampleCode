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
'   Creation Date:  FriDay,June 6 2008
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.June.2008    VRK     CR-141836:Provide new electrical equipment symbols for lighting fixtures
'   03.Feb.2010     PRB     TR-CP-157988  TO Do List entries are generated whern placing lighting fixtures
'                            (Modified nozzle code to create with placeholder)
'   25.Mar.2010     RUK     CR-CP-177179  Enhance lighting symbols to provide control point for dimensioning purposes
'                           CR-CP-113328  "can be modified" flag should be True for all best practice equipment symbols
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
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
    Dim parMountPlateLength As Double
    Dim parMountPlateWidth As Double
    Dim parMountPlateThickness As Double
    Dim parFixtureLength As Double
    Dim parFixtureWidth As Double
    Dim parFixtureHeight As Double
    Dim parBoxHeight As Double
    Dim parBoxWidth As Double
    Dim parBoxLength As Double
    Dim parwattage As Double
    Dim parCPx As Double
    Dim parCPy As Double
    Dim parCPz As Double

    Dim dPoleRadius As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim objPole As Object
    Dim objcylinder As Object
    Dim oAxisVec As New AutoMath.DVector
    Dim oTransMatrix As IJDT4x4
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oAxisVec = New AutoMath.DVector
    Set oTransMatrix = New DT4x4

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPoleHeight = arrayOfInputs(2)
    parPoleDiameter1 = arrayOfInputs(3)
    parPoleDiameter2 = arrayOfInputs(4)
    parMountPlateLength = arrayOfInputs(5)
    parMountPlateWidth = arrayOfInputs(6)
    parMountPlateThickness = arrayOfInputs(7)
    parFixtureHeight = arrayOfInputs(8)
    parFixtureWidth = arrayOfInputs(9)
    parFixtureLength = arrayOfInputs(10)
    parBoxHeight = arrayOfInputs(11)
    parBoxWidth = arrayOfInputs(12)
    parBoxLength = arrayOfInputs(13)
    parwattage = arrayOfInputs(14)
    
    If UBound(arrayOfInputs) > 14 Then
        parCPx = arrayOfInputs(15)
        parCPy = arrayOfInputs(16)
        parCPz = arrayOfInputs(17)
    End If

    iOutput = 0
    dPoleRadius = parPoleDiameter1 / 2
    
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
    dMountPlatePoints(8) = parMountPlateThickness
    dMountPlatePoints(9) = dMountPlatePoints(0)
    dMountPlatePoints(10) = dMountPlatePoints(1)
    dMountPlatePoints(11) = dMountPlatePoints(8)
    dMountPlatePoints(12) = dMountPlatePoints(0)
    dMountPlatePoints(13) = dMountPlatePoints(1)
    dMountPlatePoints(14) = dMountPlatePoints(2)
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                        5, dMountPlatePoints)
    oAxisVec.Set 0, 1, 0
    Set objMountPlate = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parMountPlateWidth, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMountPlate
    Set objMountPlate = Nothing
    
    'Vertical Pole 1
    oStPoint.Set 0, 0, parMountPlateThickness
    oEnPoint.Set 0, 0, oStPoint.z + 0.25 * parPoleHeight
    Set objPole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPoleDiameter1, True)

'   Set output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPole
    Set objPole = Nothing
  
    'Vertical Pole 2
    oStPoint.Set 0, 0, parMountPlateThickness + 0.25 * parPoleHeight
    oEnPoint.Set 0, 0, oStPoint.z + 0.25 * parPoleHeight
    Set objPole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, (parPoleDiameter1 + parPoleDiameter2) / 2, True)

'   Set output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPole
    Set objPole = Nothing
    
    'Vertical Pole 3
    oStPoint.Set 0, 0, parMountPlateThickness + 0.5 * parPoleHeight
    oEnPoint.Set 0, 0, oStPoint.z + 0.25 * parPoleHeight
    Set objPole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, (((parPoleDiameter1 + parPoleDiameter2) / 2) + parPoleDiameter2) / 2, True)
'   Set output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPole
    Set objPole = Nothing
    
    'Vertical Pole 4
    oStPoint.Set 0, 0, parMountPlateThickness + 0.75 * parPoleHeight
    oEnPoint.Set 0, 0, oStPoint.z + 0.25 * parPoleHeight
    Set objPole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPoleDiameter2, True)

'   Set output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPole
    Set objPole = Nothing
  
    'Create Output for Light
    Dim oLine As Object
    Dim objCollection As Collection
    Set objCollection = New Collection
    Dim objCurve As Object
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim dCylinderHeight As Double
    Dim oBox As Object
    dCylinderHeight = 0.5 * parBoxHeight

    'Create Box (Output 1)
    oStPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
    oEnPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, 0
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, 0
    oEnPoint.Set 0.5 * parFixtureLength, -0.5 * parFixtureWidth, 0
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set 0.5 * parFixtureLength, -0.5 * parFixtureWidth, 0
    oEnPoint.Set 0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    
    oStPoint.Set 0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
    oEnPoint.Set 0.48 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set 0.48 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
    oEnPoint.Set 0.48 * parFixtureLength, -0.5 * parFixtureWidth, 0.01 * parFixtureHeight
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set 0.48 * parFixtureLength, -0.5 * parFixtureWidth, 0.01 * parFixtureHeight
    oEnPoint.Set -0.48 * parFixtureLength, -0.5 * parFixtureWidth, 0.01 * parFixtureHeight
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set -0.48 * parFixtureLength, -0.5 * parFixtureWidth, 0.01 * parFixtureHeight
    oEnPoint.Set -0.48 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set -0.48 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
    oEnPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
    Set oAxisVec = New DVector
    oAxisVec.Set 0, 1, 0
    Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
    Set objCurve = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parFixtureWidth, True)
    'Set the output
    Set oTransMatrix = New DT4x4
    oTransMatrix.LoadIdentity
    oTransMatrix.IndexValue(12) = 0
    oTransMatrix.IndexValue(13) = 0
    oTransMatrix.IndexValue(14) = parMountPlateThickness + parPoleHeight
    objCurve.Transform oTransMatrix
'   Set output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurve
    Set objCurve = Nothing
    
    Set oAxisVec = Nothing
    Dim iCount As Integer
    For iCount = 1 To objCollection.Count
        objCollection.Remove 1
        oComplexStr.RemoveCurve True
    Next iCount
    Set oComplexStr = Nothing
    Set objCollection = Nothing
    
    'Create the CylindricalJoint
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, 0, -0.1 * dCylinderHeight
    Set objcylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.2 * parFixtureLength, True)
    'Set the Output
    Set oTransMatrix = New DT4x4
    oTransMatrix.LoadIdentity
    oTransMatrix.IndexValue(12) = 0
    oTransMatrix.IndexValue(13) = 0
    oTransMatrix.IndexValue(14) = parMountPlateThickness + parPoleHeight
    objcylinder.Transform oTransMatrix
    
'   Set output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcylinder
    Set objcylinder = Nothing
    'Box
    oStPoint.Set -0.48 * parBoxLength, -0.5 * parBoxWidth, parPoleHeight + 0.5 * parFixtureHeight
    oEnPoint.Set 0.48 * parBoxLength, 0.5 * parBoxWidth, oStPoint.z + parBoxHeight
    Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
'   Set output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
    
     ' Insert your code for output 10(Conduit Port 1)
    Dim oConduitPortPoint As AutoMath.DPosition
    Set oConduitPortPoint = New AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    
    Set oDir = New AutoMath.DVector
    oDir.Set 0, 0, -1
    oConduitPortPoint.Set 0, 0, 0
    
    Dim ObjConduitPort1 As IJConduitPortOcc
    Set ObjConduitPort1 = CreateConduitNozzlePH(oConduitPortPoint, oDir, m_OutputColl, oPartFclt, 1)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort1
    Set ObjConduitPort1 = Nothing
    
    ' Insert your code for output 11(Cable Port 2)
    Dim oCableTrayPortPoint As AutoMath.DPosition
    Set oCableTrayPortPoint = New AutoMath.DPosition
    
    Dim oRadialOrient As AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    
    oRadialOrient.Set 0, -1, 0
    oCableTrayPortPoint.Set 0, 0, 0
    
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
       
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub




