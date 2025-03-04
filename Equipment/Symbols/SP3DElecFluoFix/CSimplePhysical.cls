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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/13/2005-6:23:57 PM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   Csimplephysical.cls
'   Author:         BBA
'   Creation Date:  Friday, Oct 15 2004
'   Description:
'   Details are taken from www.lithonia.com site. Model: AFP-Heavy-Duty Industrials, T5HO.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   13.Jan.2005     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim parElectEquipLength   As Double
    Dim ostPoint As New AutoMath.DPosition
    Dim oenPoint As New AutoMath.DPosition
    Dim objarc As Object
    Dim ObjExchangerBody  As Object
    Dim iOutput     As Double
    
    
'Length along x-axis, Width along y-axis, height along z-axis

' Inputs

    Set oPartFclt = arrayOfInputs(1)
    parElectEquipLength = arrayOfInputs(2)
    parElectEquipHeight = arrayOfInputs(3)
    parElectEquipWidth = arrayOfInputs(4)
    
    iOutput = 0

' Insert your code for output 1(TopCover)
    Dim pPos1   As New AutoMath.DPosition
    Dim pPos2   As New AutoMath.DPosition
    Dim oBox    As Object
    pPos1.Set 0, 0.6875 * parElectEquipWidth, 0.9375 * parElectEquipHeight
    pPos2.Set parElectEquipLength, 0.3125 * parElectEquipWidth, parElectEquipHeight
    Set oBox = PlaceBox(m_OutputColl, pPos1, pPos2)
  
' Set the output

    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
      
    
' Insert the code for output 2(SideCoverRight)
    pPos1.Set 0.99583333 * parElectEquipLength, 0.6875 * parElectEquipWidth, 0.375 * parElectEquipHeight
    pPos2.Set parElectEquipLength, 0.3125 * parElectEquipWidth, parElectEquipHeight
    Set oBox = PlaceBox(m_OutputColl, pPos1, pPos2)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
    

' Insert the code for output 3(SideCoverLeft)
    pPos1.Set 0, 0.6875 * parElectEquipWidth, 0.375 * parElectEquipHeight
    pPos2.Set 0.00416667 * parElectEquipLength, 0.3125 * parElectEquipWidth, parElectEquipHeight
    Set oBox = PlaceBox(m_OutputColl, pPos1, pPos2)
    
' Set the output

    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
    

'Insert your code for output 4(curvedFrontLampCover)
    Dim oPoint1 As New AutoMath.DPosition
    Dim oPoint2 As New AutoMath.DPosition
    Dim oPoint3 As New AutoMath.DPosition
    
    oPoint1.Set 0.041666667 * parElectEquipLength, 0.3125 * parElectEquipWidth, parElectEquipHeight
    oPoint2.Set 0.041666667 * parElectEquipLength, 0.15625 * parElectEquipWidth, 0.9166667 * parElectEquipHeight
    oPoint3.Set 0.041666667 * parElectEquipLength, 0, 0
    
    Dim oGeometryFactory As New IngrGeom3D.GeometryFactory
    Set objarc = oGeometryFactory.Arcs3d.CreateBy3Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, oPoint2.x, oPoint2.y, oPoint2.z, oPoint3.x, oPoint3.y, oPoint3.z)
    Dim oDirProj As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 1, 0, 0
    Set ObjExchangerBody = PlaceProjection(m_OutputColl, objarc, oDirProj, 0.9166667 * parElectEquipLength, True)

' Set the output

    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    Set objarc = Nothing
        
'Insert your code for output 5(curvedRearLampCover)
    oPoint1.Set 0.041666667 * parElectEquipLength, 0.6875 * parElectEquipWidth, parElectEquipHeight
    oPoint2.Set 0.041666667 * parElectEquipLength, 0.84375 * parElectEquipWidth, 0.9166667 * parElectEquipHeight
    oPoint3.Set 0.041666667 * parElectEquipLength, parElectEquipWidth, 0
    Set objarc = oGeometryFactory.Arcs3d.CreateBy3Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, oPoint2.x, oPoint2.y, oPoint2.z, oPoint3.x, oPoint3.y, oPoint3.z)
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 1, 0, 0
    Set ObjExchangerBody = PlaceProjection(m_OutputColl, objarc, oDirProj, 0.9166667 * parElectEquipLength, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    Set objarc = Nothing
    
    
' Insert the code for output 6(cylinder)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim objCylinder  As Object
    stPoint.Set 0, 0.375 * parElectEquipWidth, 0.75 * parElectEquipHeight
    enPoint.Set parElectEquipLength, 0.375 * parElectEquipWidth, 0.75 * parElectEquipHeight
    
 
' Set the output
    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, 0.05 * parElectEquipWidth, True)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    
    
' Insert the code for output 7(cylinder)
    stPoint.Set 0, 0.5 * parElectEquipWidth, 0.75 * parElectEquipHeight
    enPoint.Set parElectEquipLength, 0.5 * parElectEquipWidth, 0.75 * parElectEquipHeight
 
' Set the output
    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, 0.05 * parElectEquipWidth, True)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    
    
' Insert the code for output 8(cylinder)
    stPoint.Set 0, 0.625 * parElectEquipWidth, 0.75 * parElectEquipHeight
    enPoint.Set parElectEquipLength, 0.625 * parElectEquipWidth, 0.75 * parElectEquipHeight
    
 
' Set the output
    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, 0.05 * parElectEquipWidth, True)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    
    
' Insert your code for output 9(Conduit Port)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oDir As IJDVector
    Dim ObjConduitPort As IJConduitPortOcc
    Dim portPoint As AutoMath.DPosition
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 0, 0, 1
   
    portPoint.Set 0.5 * parElectEquipLength, 0.5 * parElectEquipWidth, parElectEquipHeight + sptOffset - depth
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
