VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Csimplephysical"
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
'   Creation Date:  Tuesday, Oct 12 2004
'   Description:
'   Details are taken from www.lithonia.com site. Model:TWH-Die-Cast Wall Pack.
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
    Dim parElectEquipLength   As Double '(depth is replaced by lenght)
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition

    Dim iOutput     As Double

'(height along x-axis,depth(or length) along z-axis,width along y-axis)

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parElectEquipLength = arrayOfInputs(2)
    parElectEquipHeight = arrayOfInputs(3)
    parElectEquipWidth = arrayOfInputs(4)
    
    iOutput = 0
    
 ' Insert your code for output 1(curvillinear portion housing the bulb)(height=60% of the total height)
    
 ' ellipse
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oComplexCollection    As Collection
    Set oComplexCollection = New Collection
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
    Dim oWallPackArc As IngrGeom3D.EllipticalArc3d
    
    DomeMinorRadius = 0.6 * parElectEquipHeight
    cenX = 0.4 * parElectEquipHeight
    cenY = 0
    cenZ = 0.2 * parElectEquipLength
    
    MajorX = 0
    MajorY = 0
    MajorZ = 0.8 * parElectEquipLength '(80% od the total depth (or length))

    mMRatio = DomeMinorRadius / MajorZ
    StartAngle = 0
    SweepAngle = PI / 2

    norX = 0
    norY = 1
    norZ = 0

    Set oWallPackArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                            StartAngle, SweepAngle)
                                            
    oComplexCollection.Add oWallPackArc
    'setting the position of the center of ellipse
        
    oStPoint.Set 0.4 * parElectEquipHeight, 0, 0.2 * parElectEquipLength
    oEnPoint.Set parElectEquipHeight, 0, 0.2 * parElectEquipLength
   
    'line1
    Dim oHLine           As IngrGeom3D.Line3d
    Set oHLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oEnPoint.x, oEnPoint.y, oEnPoint.z, oStPoint.x, oStPoint.y, oStPoint.z)
    oComplexCollection.Add oHLine
 
    'line2
    Dim oVLine As IngrGeom3D.Line3d
    oStPoint.Set 0.4 * parElectEquipHeight, 0, 0.2 * parElectEquipLength
    oEnPoint.Set 0.4 * parElectEquipHeight, 0, parElectEquipLength
 
    Set oVLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oComplexCollection.Add oVLine
          
'creating the complex curve
    Dim StartBC   As New AutoMath.DPosition
    
    Dim oDirProj As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 1, 0
    
    Dim objComplexCurve  As IngrGeom3D.ComplexString3d
    
    StartBC.Set 0.4 * parElectEquipHeight, 0, parElectEquipLength

    
    Set objComplexCurve = PlaceTrCString(StartBC, oComplexCollection)
    Dim oEllipBox As Object
    Set oEllipBox = PlaceProjection(m_OutputColl, objComplexCurve, oDirProj, parElectEquipWidth, True)
    Set objComplexCurve = Nothing
    
' Set the output 1(Curved portion of the wallpack)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oEllipBox
    Set oEllipBox = Nothing

' Insert your code for output 2(TopBox) (height=40% of the total height)
    Dim pPos1   As New AutoMath.DPosition
    Dim pPos2   As New AutoMath.DPosition
    Dim oBox    As Object

    pPos1.Set 0, parElectEquipWidth, 0.2 * parElectEquipLength
    pPos2.Set 0.4 * parElectEquipHeight, 0, parElectEquipLength
    Set oBox = PlaceBox(m_OutputColl, pPos1, pPos2)
  
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
    
    ' insert the code for output 3(Bottom Basebox)(height=total height & depth=20% of the total depth(or length))
    pPos1.Set 0, parElectEquipWidth, 0
    pPos2.Set parElectEquipHeight, 0, 0.2 * parElectEquipLength
    Set oBox = PlaceBox(m_OutputColl, pPos1, pPos2)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing

' Insert your code for output 4(Conduit Port)(on top and at the center of the TopBox, along its height)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oDir As IJDVector
    Dim ObjConduitPort As IJConduitPortOcc
    Dim portPoint As AutoMath.DPosition
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set -1, 0, 0
   
    portPoint.Set 0, 0.5 * parElectEquipWidth, 0.5 * parElectEquipLength
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    
' Insert your code for output 5(Cable Port)
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim NozzlePHFactory As NozzlePHFactory
    Dim ObjCableNozzle As CableNozzle
    Dim cablePortPoint As AutoMath.DPosition

    Set cablePortPoint = New AutoMath.DPosition
    Set NozzlePHFactory = New NozzlePHFactory
        
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set 0, 0.5 * parElectEquipWidth, 0.5 * parElectEquipLength
                                
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
