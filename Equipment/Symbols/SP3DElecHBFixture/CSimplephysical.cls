VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplephysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/13/2005-6:25:39 PM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   Csimplephysical.cls
'   Author:         BBA
'   Creation Date:  Wednesday, Oct 13 2004
'   Description:
'   Details are taken from www.lithonia.com site. High Bay Industrial Lighting, Model:TPGE 250S PG16GLE M 120.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   11.JAN.2005     MS     CR69200 Added cable ports at the same location of conduit ports.
'   13.Jan.2005     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   18.Apr.2005     MS TR-75683 Changed port location
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  22.Apr.2008     VRK           CR-120097  Need to set proper units to the cable nozzle diameter and area
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Dim m_oGeomHelper As IJSymbolGeometryHelper
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
    Dim ConduitDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim parElectEquipHeight  As Double
    Dim parElectEquipWidth   As Double
    Dim ostPoint As New AutoMath.DPosition
    Dim oenPoint As New AutoMath.DPosition
    Dim objarc As Object
    Dim ObjExchangerBody  As Object
    Dim iOutput     As Double
    
'origin is at center of the bottom(maximum) diameter of the output 4

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parElectEquipHeight = arrayOfInputs(2)
    parElectEquipWidth = arrayOfInputs(3)
    
 
    
    iOutput = 0

' Insert your code for output 1(TopBox)(height = 7% of the total height & width = 50% of the total width)
    Dim oPos1   As New AutoMath.DPosition
    Dim oPos2   As New AutoMath.DPosition
    Dim oBox    As Object
    oPos1.Set -0.25 * parElectEquipWidth, 0.25 * parElectEquipWidth, 0.93 * parElectEquipHeight
    oPos2.Set 0.25 * parElectEquipWidth, -0.25 * parElectEquipWidth, parElectEquipHeight
    Set oBox = PlaceBox(m_OutputColl, oPos1, oPos2)
  
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
       
' Insert the code for output 2(Middlebox)(height = 16% of the total height & width=75% of the total width)
    oPos1.Set -0.375 * parElectEquipWidth, 0.375 * parElectEquipWidth, 0.77 * parElectEquipHeight
    oPos2.Set 0.375 * parElectEquipWidth, -0.375 * parElectEquipWidth, 0.93 * parElectEquipHeight
    Set oBox = PlaceBox(m_OutputColl, oPos1, oPos2)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
    
' Insert the code for output 3(cylinder)(height = 17% of the total height & diameter = 50% of the width)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim ldiam     As Long
    Dim objCylinder  As Object
    stPoint.Set 0, 0, 0.6 * parElectEquipHeight
    enPoint.Set 0, 0, 0.77 * parElectEquipHeight

    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, 0.5 * parElectEquipWidth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    
 'Insert your code for output 4(lampcover)(height=60% of the total height)
    Dim oPoint1 As New AutoMath.DPosition
    Dim oPoint2 As New AutoMath.DPosition
    Dim oPoint3 As New AutoMath.DPosition
    
    oPoint1.Set 0.25 * parElectEquipWidth, 0, 0.6 * parElectEquipHeight
    oPoint2.Set 0.45 * parElectEquipWidth, 0, 0.3 * parElectEquipHeight
    oPoint3.Set 0.5 * parElectEquipWidth, 0, 0
    
    Dim oGeometryFactory As New IngrGeom3D.GeometryFactory
    Set objarc = oGeometryFactory.Arcs3d.CreateBy3Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, oPoint2.x, oPoint2.y, oPoint2.z, oPoint3.x, oPoint3.y, oPoint3.z)
        
    Dim revVector As New AutoMath.DVector
    Dim revCenterPt As New AutoMath.DPosition
    revVector.Set 0, 0, 1
    revCenterPt.Set 0, 0, 0 'origin
    Set ObjExchangerBody = PlaceRevolution(m_OutputColl, objarc, revVector, revCenterPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    Set objarc = Nothing

' Insert your code for output 5(Conduit Port)(on top and at the center of the ouput 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim oDir As IJDVector
    Dim ObjConduitPort As IJConduitPortOcc
    Dim portPoint As AutoMath.DPosition
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 0, 0, 1
   
    portPoint.Set 0, 0, parElectEquipHeight
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    
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
    cablePortPoint.Set 0, 0, parElectEquipHeight
                                
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
