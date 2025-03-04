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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/13/2005-6:25:39 PM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   Csimplephysical.cls
'   Author:         BBA
'   Creation Date:  Wednesday, Oct 13 2004
'   Description:
'   Details are taken from www.lithonia.com site. Model-Low Bay Industrial Lighting TXV.
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
    
'height along z-axis & width along x and y axis'

'origin is at the intersection of the outputs 4 and 5 and at the center of the ouput 5

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parElectEquipHeight = arrayOfInputs(2)
    parElectEquipWidth = arrayOfInputs(3)
       
    iOutput = 0
    
   

' Insert your code for output 1(TopBox)(height is 15% of the total height & width is 34% of the total width)
    Dim pPos1   As New AutoMath.DPosition
    Dim pPos2   As New AutoMath.DPosition
    Dim oBox    As Object
    pPos1.Set -0.17 * parElectEquipWidth, 0.17 * parElectEquipWidth, 0.4375 * parElectEquipHeight
    pPos2.Set 0.17 * parElectEquipWidth, -0.17 * parElectEquipWidth, 0.5875 * parElectEquipHeight
    Set oBox = PlaceBox(m_OutputColl, pPos1, pPos2)
  
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
    
    
' Insert the code for output 2(Middlebox)(height = 22.5% of the total height & width = 70% of the total width)
    pPos1.Set -0.35 * parElectEquipWidth, 0.35 * parElectEquipWidth, 0.2125 * parElectEquipHeight
    pPos2.Set 0.35 * parElectEquipWidth, -0.35 * parElectEquipWidth, 0.4375 * parElectEquipHeight
    Set oBox = PlaceBox(m_OutputColl, pPos1, pPos2)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
    
' Insert the code for output 3(cylinder)(height = 8.75% of the total height & diameter= 28% of the width)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim objCylinder  As Object
    stPoint.Set 0, 0, 0.2125 * parElectEquipHeight
    enPoint.Set 0, 0, 0.125 * parElectEquipHeight
    
 
' Set the output
    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, 0.28 * parElectEquipWidth, True)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    
    
 'Insert your code for output 4(Top Lampcover)(height = 12.5% of the total height)

    Dim oPoint1 As New AutoMath.DPosition
    Dim oPoint2 As New AutoMath.DPosition
    Dim oPoint3 As New AutoMath.DPosition
    
    oPoint1.Set 0.14 * parElectEquipWidth, 0, 0.125 * parElectEquipHeight
    oPoint2.Set 0.32 * parElectEquipWidth, 0, 0.0625 * parElectEquipHeight
    oPoint3.Set 0.4 * parElectEquipWidth, 0, 0
    
    Dim oGeometryFactory As New IngrGeom3D.GeometryFactory
    Set objarc = oGeometryFactory.Arcs3d.CreateBy3Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, oPoint2.x, oPoint2.y, oPoint2.z, oPoint3.x, oPoint3.y, oPoint3.z)
        
    Dim revVector As New AutoMath.DVector
    Dim revCenterPt As New AutoMath.DPosition
    revVector.Set 0, 0, 1
    revCenterPt.Set 0, 0, 0
    Set ObjExchangerBody = PlaceRevolution(m_OutputColl, objarc, revVector, revCenterPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    Set objarc = Nothing
    

' Insert the code for output 5(Cylindrical Interface)(height=3.75% of the total height & diameter=width)
    stPoint.Set 0, 0, 0   '(origin)
    enPoint.Set 0, 0, -0.0375 * parElectEquipHeight
 
' Set the output
    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parElectEquipWidth, True)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing

'Insert your code for output 6(Bottom Lampcover)(height=37.5% of the total height)
    oPoint1.Set 0.4 * parElectEquipWidth, 0, -0.0375 * parElectEquipHeight
    oPoint2.Set 0.35 * parElectEquipWidth, 0, -0.225 * parElectEquipHeight
    oPoint3.Set 0.25 * parElectEquipWidth, 0, -0.4125 * parElectEquipHeight
    
    Set objarc = oGeometryFactory.Arcs3d.CreateBy3Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, oPoint2.x, oPoint2.y, oPoint2.z, oPoint3.x, oPoint3.y, oPoint3.z)
    revVector.Set 0, 0, 1
    revCenterPt.Set 0, 0, 0
    Set ObjExchangerBody = PlaceRevolution(m_OutputColl, objarc, revVector, revCenterPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    Set objarc = Nothing

' Insert your code for output 7(Conduit Port)(on top and at the center of the ouput 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oDir As IJDVector
    Dim ObjConduitPort As IJConduitPortOcc
    Dim portPoint As AutoMath.DPosition
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 0, 0, 1
   
    portPoint.Set 0, 0, 0.5875 * parElectEquipHeight
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    
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
    cablePortPoint.Set 0, 0, 0.5875 * parElectEquipHeight
                                
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

