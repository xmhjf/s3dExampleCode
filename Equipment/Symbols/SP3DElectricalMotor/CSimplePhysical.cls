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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/13/2005-6:25:31 PM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MS
'   Creation Date:  Wednesday 6th October, 2004
'   Description:
'    The symbol consists of a box, conduit port.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   11.JAN.2005     MS     CR69200 Added cable ports at the same location of conduit ports.
'   13.Jan.2005     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  22.Apr.2008     VRK           CR-120097  Need to set proper units to the cable nozzle diameter and area
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private m_oCodeListMetadata     As IJDCodeListMetaData
Dim PI As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
    Set m_oGeomHelper = New SymbolServices
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double

    Dim parMotorLength As Double
    Dim parMotorDiameter As Double
    Dim parShaftLength As Double
    Dim parShaftDiameter As Double
    
    Dim ObjConduitPort As IJConduitPortOcc
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parMotorLength = arrayOfInputs(2)
    parMotorDiameter = arrayOfInputs(3)
    parShaftLength = arrayOfInputs(4)
    parShaftDiameter = arrayOfInputs(5)
    
    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
' Insert your code for output 1(Cylindrical Body of Motor)
    'Assuming the dome part to be 0.2 times the Motor Length
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set parMotorLength - 0.5 * parMotorDiameter, 0, 0
    ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parMotorDiameter

' Insert your code for output 2(Dome 1)
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oAxisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double
    Dim DomeMinorRadius As Double
    Dim oMotorArc As IngrGeom3D.EllipticalArc3d
    Dim ObjFrontHead As Object

    DomeMinorRadius = 0.5 * (parMotorDiameter / 2)
    cenX = 0
    cenY = 0
    cenZ = 0

    MajorX = 0
    MajorY = 0
    MajorZ = parMotorDiameter / 2

    mMRatio = DomeMinorRadius / (parMotorDiameter / 2)
    StartAngle = PI
    SweepAngle = PI / 2

    norX = 0
    norY = 1
    norZ = 0

    Set oMotorArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                            StartAngle, SweepAngle)

    oAxisVect.Set 1, 0, 0
    revCenPt.Set 0, 0, 0
    'Revolve it about X-Axiz

    Set ObjFrontHead = PlaceRevolution(m_OutputColl, oMotorArc, oAxisVect, revCenPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontHead
    Set ObjFrontHead = Nothing

' Insert your code for output 3(Dome 2)
    cenX = parMotorLength - 0.5 * parMotorDiameter
    cenY = 0
    cenZ = 0

    MajorX = 0
    MajorY = 0
    MajorZ = parMotorDiameter / 2

    mMRatio = DomeMinorRadius / (parMotorDiameter / 2)
    StartAngle = 1.5 * PI
    SweepAngle = PI / 2

    norX = 0
    norY = -1
    norZ = 0

    Set oMotorArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                            StartAngle, SweepAngle)

    oAxisVect.Set 1, 0, 0
    revCenPt.Set parMotorLength - 0.5 * parMotorDiameter, 0, 0
    'Revolve it about X-Axiz
    Set ObjFrontHead = PlaceRevolution(m_OutputColl, oMotorArc, oAxisVect, revCenPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontHead
    Set ObjFrontHead = Nothing
    Set oGeomFactory = Nothing
    Set oAxisVect = Nothing
    Set revCenPt = Nothing
    
' Insert your code for output 4(Shaft)
'
    stPoint.Set parMotorLength - 0.3 * parMotorDiameter, 0, 0
    enPoint.Set parMotorLength - 0.25 * parMotorDiameter + parShaftLength, 0, 0
'    stPoint.Set 0, 0, 0
'    enPoint.Set parMotorLength - 0.2 * parMotorLength, 0, 0
    ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parShaftDiameter
    
' Insert your code for output 5(Base of Motor)
    Dim objMotorBase As Object
    stPoint.Set 0, -parMotorDiameter / 2, -parMotorDiameter / 2
    enPoint.Set parMotorLength - 0.5 * parMotorDiameter, _
                parMotorDiameter / 2, _
                -parMotorDiameter / 2 - 0.05 * parMotorDiameter
    
    Set objMotorBase = PlaceBox(m_OutputColl, stPoint, enPoint)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMotorBase
    Set objMotorBase = Nothing
    
' Insert your code for output 6(Box for wires and cables)
    Dim objCableBox As Object
    stPoint.Set (parMotorLength - 0.5 * parMotorDiameter) / 2 - 0.1 * parMotorLength, _
                    parMotorDiameter / 2, _
                    -parMotorDiameter / 6
    enPoint.Set (parMotorLength - 0.5 * parMotorDiameter) / 2 + 0.1 * parMotorLength, _
                    1.1 * parMotorDiameter / 2, _
                    parMotorDiameter / 6
    
    Set objCableBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableBox
    Set objCableBox = Nothing
    
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Insert your code for output 7(Conduit Port)
    Dim oDir As IJDVector
    Dim portPoint As AutoMath.DPosition
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 0, 1, 0
   
    portPoint.Set (parMotorLength - 0.5 * parMotorDiameter) / 2, 1.1 * parMotorDiameter / 2, 0
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    Set portPoint = Nothing
    
' Insert your code for output 8(Cable Port)
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim NozzlePHFactory As NozzlePHFactory
    Dim ObjCableNozzle As CableNozzle
    Dim cablePortPoint As AutoMath.DPosition

    Set cablePortPoint = New AutoMath.DPosition
    Set NozzlePHFactory = New NozzlePHFactory
        
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set (parMotorLength - 0.5 * parMotorDiameter) / 2, 1.1 * parMotorDiameter / 2, 0
                                
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation cablePortPoint
    Set iDistribPort = ObjCableNozzle
    iDistribPort.SetDirectionVector oDir
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableNozzle
    Set ObjCableNozzle = Nothing
    Set NozzlePHFactory = Nothing
    Set iLogicalDistPort = Nothing
    Set cablePortPoint = Nothing
    Set iDistribPort = Nothing
    Set oDir = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
