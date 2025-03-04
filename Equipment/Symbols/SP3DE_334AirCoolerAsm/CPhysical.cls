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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: VRK
'   Creation Date:  Tuesday, Feb 13 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has twentyeight Outputs:
'      In which Default Surface is created at the box.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
     Set m_oSymGeomHelper = New SymbolServices
     
     Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
   
    Dim iOutput As Double
    Dim ObjCoolerLeftBox  As Object
    Dim ObjCoolerRightBox  As Object
    Dim ObjPlenumBox  As Object
    Dim ObjFan As Object

    Dim parCoolerWidth As Double            'P21
    Dim parCoolerLength As Double           'P22
    Dim parCoolerHeight As Double           'P23
    Dim parPlenumLength As Double           'P24
    Dim parPlenumHeight As Double           'P25
    Dim parInletBlockLength As Double       'P26
    Dim parBlockToptoCoolTop As Double      'P27
    Dim parInletBlockHeight As Double       'P28
    Dim parCoolLeftfromPP As Double         'P29
    Dim parNumberOfFans As Long             'P30
    Dim parFan1CentoPP As Double            'P31
    Dim parFansCentoCen As Double           'P32
    Dim parFanDiameter As Double            'P33
    Dim parFanHeight As Double              'P34
   
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCoolerWidth = arrayOfInputs(2)
    parCoolerLength = arrayOfInputs(3)
    parCoolerHeight = arrayOfInputs(4)
    parPlenumLength = arrayOfInputs(5)
    parPlenumHeight = arrayOfInputs(6)
    parInletBlockLength = arrayOfInputs(7)
    parBlockToptoCoolTop = arrayOfInputs(8)
    parInletBlockHeight = arrayOfInputs(9)
    parCoolLeftfromPP = arrayOfInputs(10)
    parNumberOfFans = arrayOfInputs(11)
    parFan1CentoPP = arrayOfInputs(12)
    parFansCentoCen = arrayOfInputs(13)
    parFanDiameter = arrayOfInputs(14)
    parFanHeight = arrayOfInputs(15)
     
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
       
    'Insert your code for output 1 to 6 (Cooler Body)
    Dim topSurPts(3) As IJDPosition
    Dim botSurPts(3) As IJDPosition
    Dim iCount As Integer
    
    For iCount = 0 To 3
        Set topSurPts(iCount) = New DPosition
        Set botSurPts(iCount) = New DPosition
    Next iCount
  
    'Point 1
    botSurPts(0).x = parCoolLeftfromPP
    botSurPts(0).y = -parCoolerWidth / 2
    botSurPts(0).z = parInletBlockHeight / 2 + parBlockToptoCoolTop - parCoolerHeight
    
    'Point 2
    botSurPts(1).x = parCoolLeftfromPP + parCoolerLength
    botSurPts(1).y = -parCoolerWidth / 2
    botSurPts(1).z = botSurPts(0).z
    
    'Point3
    botSurPts(2).x = parCoolLeftfromPP + parCoolerLength
    botSurPts(2).y = parCoolerWidth / 2
    botSurPts(2).z = botSurPts(0).z
    
    'Point 4
    botSurPts(3).x = parCoolLeftfromPP
    botSurPts(3).y = parCoolerWidth / 2
    botSurPts(3).z = botSurPts(0).z
    
    topSurPts(0).Set botSurPts(0).x, botSurPts(0).y, botSurPts(0).z + parCoolerHeight
    topSurPts(1).Set botSurPts(1).x, botSurPts(1).y, botSurPts(1).z + parCoolerHeight
    topSurPts(2).Set botSurPts(2).x, botSurPts(2).y, botSurPts(2).z + parCoolerHeight
    topSurPts(3).Set botSurPts(3).x, botSurPts(3).y, botSurPts(3).z + parCoolerHeight
    
    Dim ObjCollection As Collection
    
    Set ObjCollection = PlaceTrapezoidWithPlanes(m_OutputColl, topSurPts, botSurPts)
    For iCount = 1 To 6
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollection(iCount)
    Next iCount

    Set ObjCollection = Nothing
    For iCount = 0 To 3
        Set topSurPts(iCount) = Nothing
        Set botSurPts(iCount) = Nothing
    Next iCount
    
    'Insert your code for output 7 (Left Box)
    stPoint.Set -parInletBlockLength + parCoolLeftfromPP, -parCoolerWidth / 2, -parInletBlockHeight / 2
    enPoint.Set stPoint.x + parInletBlockLength, stPoint.y + parCoolerWidth, stPoint.z + parInletBlockHeight
    
    Set ObjCoolerLeftBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCoolerLeftBox
    Set ObjCoolerLeftBox = Nothing
    
    'Insert your code for output 8 (Right Box)
    stPoint.Set parCoolerLength + parCoolLeftfromPP, -parCoolerWidth / 2, -parInletBlockHeight / 2
    enPoint.Set stPoint.x + parInletBlockLength, stPoint.y + parCoolerWidth, stPoint.z + parInletBlockHeight
    
    Set ObjCoolerRightBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCoolerRightBox
    Set ObjCoolerRightBox = Nothing
    
    'Insert your code for output 9(Plenum Box)
    stPoint.Set parCoolLeftfromPP + (parCoolerLength - parPlenumLength) / 2, -parCoolerWidth / 2, _
                                                    parInletBlockHeight / 2 + parBlockToptoCoolTop - parCoolerHeight - parPlenumHeight
    enPoint.Set stPoint.x + parPlenumLength, stPoint.y + parCoolerWidth, stPoint.z + parPlenumHeight
    
    Set ObjPlenumBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlenumBox
    Set ObjPlenumBox = Nothing
    
    'Insert your code for output (Fan)
    Dim fanXOffset As Double
    fanXOffset = 0
    For iCount = 1 To parNumberOfFans
        stPoint.Set parFan1CentoPP + fanXOffset, 0, _
                                                        parInletBlockHeight / 2 + parBlockToptoCoolTop - parCoolerHeight - parPlenumHeight
        enPoint.Set stPoint.x, stPoint.y, stPoint.z - parFanHeight
        
        Set ObjFan = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFanDiameter, True)
        
        'Set the output
        m_OutputColl.AddOutput "Fans_", ObjFan
        Set ObjFan = Nothing
        fanXOffset = fanXOffset + parFansCentoCen
    Next iCount
    
    Set stPoint = Nothing
    Set enPoint = Nothing
  
    'Adding LinePoints and edges on the Baseplate
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim LinePoints(0 To 23) As Double
    Dim objPoint As Object
    Dim line As Object
    
    LinePoints(0) = parCoolLeftfromPP
    LinePoints(1) = -parCoolerWidth / 2
    LinePoints(2) = parInletBlockHeight / 2 + parBlockToptoCoolTop - parCoolerHeight
    
    LinePoints(3) = parCoolLeftfromPP + parCoolerLength
    LinePoints(4) = -parCoolerWidth / 2
    LinePoints(5) = parInletBlockHeight / 2 + parBlockToptoCoolTop - parCoolerHeight
    
    LinePoints(6) = parCoolLeftfromPP + parCoolerLength
    LinePoints(7) = parCoolerWidth / 2
    LinePoints(8) = parInletBlockHeight / 2 + parBlockToptoCoolTop - parCoolerHeight
    
    LinePoints(9) = parCoolLeftfromPP
    LinePoints(10) = parCoolerWidth / 2
    LinePoints(11) = parInletBlockHeight / 2 + parBlockToptoCoolTop - parCoolerHeight
      
    LinePoints(12) = parCoolLeftfromPP
    LinePoints(13) = -parCoolerWidth / 2
    LinePoints(14) = parInletBlockHeight / 2 + parBlockToptoCoolTop
    
    LinePoints(15) = parCoolLeftfromPP + parCoolerLength
    LinePoints(16) = -parCoolerWidth / 2
    LinePoints(17) = parInletBlockHeight / 2 + parBlockToptoCoolTop
    
    LinePoints(18) = parCoolLeftfromPP + parCoolerLength
    LinePoints(19) = parCoolerWidth / 2
    LinePoints(20) = parInletBlockHeight / 2 + parBlockToptoCoolTop
    
    LinePoints(21) = parCoolLeftfromPP
    LinePoints(22) = parCoolerWidth / 2
    LinePoints(23) = parInletBlockHeight / 2 + parBlockToptoCoolTop
    
    'Create salient points on each face of the body
    'Bottom Point
    Dim x, y, z As Double
    x = (LinePoints(0) + LinePoints(6)) / 2
    y = (LinePoints(1) + LinePoints(7)) / 2
    z = (LinePoints(2) + LinePoints(8)) / 2
    
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Top Point
    x = (LinePoints(12) + LinePoints(18)) / 2
    y = (LinePoints(13) + LinePoints(19)) / 2
    z = (LinePoints(14) + LinePoints(20)) / 2
    
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Left Face Point
    x = (LinePoints(0) + LinePoints(15)) / 2
    y = (LinePoints(1) + LinePoints(16)) / 2
    z = (LinePoints(2) + LinePoints(17)) / 2
    
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Right Face Point
    x = (LinePoints(6) + LinePoints(21)) / 2
    y = (LinePoints(7) + LinePoints(22)) / 2
    z = (LinePoints(8) + LinePoints(23)) / 2
    
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Front Face Point
    x = (LinePoints(0) + LinePoints(21)) / 2
    y = (LinePoints(1) + LinePoints(22)) / 2
    z = (LinePoints(2) + LinePoints(23)) / 2
    
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Back Face Point
    x = (LinePoints(3) + LinePoints(18)) / 2
    y = (LinePoints(4) + LinePoints(19)) / 2
    z = (LinePoints(5) + LinePoints(20)) / 2
    
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Create line by two Points:
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(0), LinePoints(1), LinePoints(2), LinePoints(3), LinePoints(4), LinePoints(5))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(3), LinePoints(4), LinePoints(5), LinePoints(6), LinePoints(7), LinePoints(8))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(6), LinePoints(7), LinePoints(8), LinePoints(9), LinePoints(10), LinePoints(11))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(9), LinePoints(10), LinePoints(11), LinePoints(0), LinePoints(1), LinePoints(2))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(0), LinePoints(1), LinePoints(2), LinePoints(12), LinePoints(13), LinePoints(14))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(12), LinePoints(13), LinePoints(14), LinePoints(21), LinePoints(22), LinePoints(23))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(21), LinePoints(22), LinePoints(23), LinePoints(9), LinePoints(10), LinePoints(11))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(18), LinePoints(19), LinePoints(20), LinePoints(6), LinePoints(7), LinePoints(8))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(18), LinePoints(19), LinePoints(20), LinePoints(15), LinePoints(16), LinePoints(17))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(3), LinePoints(4), LinePoints(5), LinePoints(15), LinePoints(16), LinePoints(17))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(15), LinePoints(16), LinePoints(17), LinePoints(12), LinePoints(13), LinePoints(14))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(21), LinePoints(22), LinePoints(23), LinePoints(18), LinePoints(19), LinePoints(20))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


