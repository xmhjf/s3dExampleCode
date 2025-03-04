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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:44 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:        MS
'   Creation Date: Tuesday, Feb 03 2004
'   Description:
'    The symbol consists of a box, cable port and conduit port.
'
'   Change History:
'   dd.mmm.yyyy     who           change description
'   -----------    -----          ------------------
'  29.Nov.2004     V6UpgradeSO   Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC           DI-95670  Replace names with initials in all revision history sheets and symbols
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

    Dim parBasicAsmHeight As Double
    Dim parBasicAsmWidth As Double
    Dim parBasicAsmLength As Double
    
    Dim ObjProjectedBox As Object
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim ObjSurface2 As IngrGeom3D.Plane3d
    Dim ObjConduitPort As IJConduitPortOcc
    Dim ObjCableNozzle As CableNozzle

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBasicAsmHeight = arrayOfInputs(2)
    parBasicAsmWidth = arrayOfInputs(3)
    parBasicAsmLength = arrayOfInputs(4)

    iOutput = 0
    
' Insert your code for outputs 1(Default Surface on the side to be supported)
    Dim linePts(0 To 14) As Double
    Dim optLine As IngrGeom3D.LineString3d
    Dim projVect As New AutoMath.DVector
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'Point 1
    linePts(0) = -parBasicAsmLength / 2
    linePts(1) = 0
    linePts(2) = 0
    'Point2
    linePts(3) = parBasicAsmLength / 2
    linePts(4) = 0
    linePts(5) = 0
    'Point3
    linePts(6) = parBasicAsmLength / 2
    linePts(7) = 0
    linePts(8) = parBasicAsmHeight
    'Point4
    linePts(9) = -parBasicAsmLength / 2
    linePts(10) = 0
    linePts(11) = parBasicAsmHeight
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
    
    Set ObjProjectedBox = PlaceProjection(m_OutputColl, optLine, projVect, parBasicAsmWidth, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjProjectedBox
    Set ObjProjectedBox = Nothing
    Set projVect = Nothing
    
' Insert your code for outputs 3(Surface on other side)
    linePts(1) = parBasicAsmWidth
    linePts(4) = parBasicAsmWidth
    linePts(7) = parBasicAsmWidth
    linePts(10) = parBasicAsmWidth
    
    Set ObjSurface2 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, linePts)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface2
    Set ObjSurface2 = Nothing
    Set geomFactory = Nothing

' Insert your code for output 4(Conduit Port)
    Dim oDir As IJDVector
    Dim portPoint As AutoMath.DPosition
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 0, 0, 1
   
    portPoint.Set 0, parBasicAsmWidth / 2, parBasicAsmHeight
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
     
' Insert your code for output 5(Cable Port)
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                            m_OutputColl.ResourceManager)
    portPoint.Set parBasicAsmLength / 2, parBasicAsmWidth / 2, parBasicAsmHeight / 2
                                
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation portPoint
    Set iDistribPort = ObjCableNozzle
    oDir.Set 1, 0, 0
    iDistribPort.SetDirectionVector oDir
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableNozzle
    Set ObjCableNozzle = Nothing
    Set NozzlePHFactory = Nothing
    Set iLogicalDistPort = Nothing
    Set portPoint = Nothing
    Set iDistribPort = Nothing
    Set oDir = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

