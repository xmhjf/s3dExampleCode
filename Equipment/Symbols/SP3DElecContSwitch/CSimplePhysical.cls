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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/13/2005-6:25:46 PM"
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

Private Sub Class_Initialize()
''
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double

    Dim parElecEquipmentHeight As Double
    Dim parElecEquipmentWidth As Double
    Dim parElecEquipmentLength As Double
    
    Dim ObjProjectedBox As Object
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim ObjSurface2 As IngrGeom3D.Plane3d
    Dim ObjProjectedKnob As Object
    Dim ObjConduitPort As IJConduitPortOcc
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parElecEquipmentHeight = arrayOfInputs(2)
    parElecEquipmentWidth = arrayOfInputs(3)
    parElecEquipmentLength = arrayOfInputs(4)

    iOutput = 0
    
' Insert your code for outputs 1(Default Surface on the side to be supported)
     Dim linePts(0 To 17) As Double
    Dim optLine As IngrGeom3D.LineString3d
    Dim projVect As New AutoMath.DVector
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'Point 1
    linePts(0) = -parElecEquipmentLength / 2
    linePts(1) = 0
    linePts(2) = 0
    'Point2
    linePts(3) = parElecEquipmentLength / 2
    linePts(4) = 0
    linePts(5) = 0
    'Point3
    linePts(6) = parElecEquipmentLength / 2
    linePts(7) = 0
    linePts(8) = parElecEquipmentHeight
    'Point4
    linePts(9) = -parElecEquipmentLength / 2
    linePts(10) = 0
    linePts(11) = parElecEquipmentHeight
    'Point5
    linePts(12) = linePts(0)
    linePts(13) = linePts(1)
    linePts(14) = linePts(2)

    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, linePts)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    
' Insert your code for outputs 2(Projected Body)
    projVect.Set 0, 1, 0
    Set optLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, linePts)
    
    Set ObjProjectedBox = PlaceProjection(m_OutputColl, optLine, projVect, parElecEquipmentWidth, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjProjectedBox
    Set ObjProjectedBox = Nothing

    
' Insert your code for outputs 3(Surface on other side)
    linePts(1) = parElecEquipmentWidth
    linePts(4) = parElecEquipmentWidth
    linePts(7) = parElecEquipmentWidth
    linePts(10) = parElecEquipmentWidth
    
    Set ObjSurface2 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, linePts)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface2
    Set ObjSurface2 = Nothing
    
 ' Insert your code for outputs 4(Knob of Control Switch)
 
    'Point 1
    linePts(0) = -parElecEquipmentLength / 4
    linePts(1) = 0
    linePts(2) = parElecEquipmentHeight / 2
    'Point2
    linePts(3) = 0
    linePts(4) = 0
    linePts(5) = parElecEquipmentHeight / 2 + parElecEquipmentHeight / 6
    'Point3
    linePts(6) = parElecEquipmentLength / 2 + parElecEquipmentLength / 4
    linePts(7) = 0
    linePts(8) = parElecEquipmentHeight / 2 + parElecEquipmentHeight / 6
    'Point4
    linePts(9) = parElecEquipmentLength / 2 + parElecEquipmentLength / 4
    linePts(10) = 0
    linePts(11) = parElecEquipmentHeight / 2 - parElecEquipmentHeight / 6
    'Point5
    linePts(12) = 0
    linePts(13) = 0
    linePts(14) = parElecEquipmentHeight / 2 - parElecEquipmentHeight / 6
    'Point6
    linePts(15) = linePts(0)
    linePts(16) = linePts(1)
    linePts(17) = linePts(2)
    
    projVect.Set 0, -1, 0
    Set optLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, linePts)
    
    Set ObjProjectedKnob = PlaceProjection(m_OutputColl, optLine, projVect, parElecEquipmentWidth / 3, True)
  ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjProjectedKnob
    Set ObjProjectedKnob = Nothing
    Set optLine = Nothing
    Set projVect = Nothing
    Set geomFactory = Nothing
    
' Insert your code for output 5(Conduit Port)
    Dim oDir As IJDVector
    Dim portPoint As AutoMath.DPosition
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 0, 0, 1
    portPoint.Set 0, parElecEquipmentWidth / 2, parElecEquipmentHeight
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    Set portPoint = Nothing

' Insert your code for output 6(Cable Port)
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim NozzlePHFactory As NozzlePHFactory
    Dim ObjCableNozzle As CableNozzle
    Dim cablePortPoint As AutoMath.DPosition
    
    Set cablePortPoint = New AutoMath.DPosition
    Set NozzlePHFactory = New NozzlePHFactory
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set 0, parElecEquipmentWidth / 2, parElecEquipmentHeight
                                
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
