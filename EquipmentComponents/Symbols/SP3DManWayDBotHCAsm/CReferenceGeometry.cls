VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReferenceGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:08:34 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, May 13 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Davit for Bottom Horizontal Cover, is prepared as per DOW Chemical Company manual.
'   Dimensions are considered for two pipe size viz., 20" and 24" . All assumed values are as per the DOW
'   documents unless otherwise specified.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "ReferenceGeometry:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim parCentertoDavitCover As Double
    Dim parDomeHeight1 As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCentertoDavitCover = arrayOfInputs(2)
    parDomeHeight1 = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'Place a Point3D at origin
    Dim ObjOriginPoint As IngrGeom3D.Point3d
    Dim xValue As Double
    Dim yValue As Double
    Dim zValue As Double
    
    xValue = 0
    yValue = 0
    zValue = 0
    Set ObjOriginPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, xValue, yValue, zValue)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOriginPoint
    
    'Place a Default Surface at the end of Pipe
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double
    
    normalX = 0
    normalY = 0
    normalZ = 1
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                                xValue, yValue, zValue, _
                                                                                normalX, normalY, normalZ, _
                                                                                (pipeDiam) / 2)
    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    
    'Place a Horizontal Line along North
    Dim ObjHoriLine As IngrGeom3D.Line3d
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double
    
    startPoint(0) = 0
    startPoint(1) = 0
    startPoint(2) = 0
    endPoint(0) = 0
    endPoint(1) = pipeDiam / 2
    endPoint(2) = 0
    Set ObjHoriLine = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, startPoint(0), startPoint(1), startPoint(2), _
                                                                                                    endPoint(0), endPoint(1), endPoint(2))
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriLine
    Set ObjHoriLine = Nothing
    
    'Place a vertical Line
    Dim ObjVertLine As IngrGeom3D.Line3d
    startPoint(0) = 0
    startPoint(1) = 0
    startPoint(2) = 0
    endPoint(0) = 0
    endPoint(1) = 0
    endPoint(2) = pipeDiam / 2
    Set ObjVertLine = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, startPoint(0), startPoint(1), startPoint(2), _
                                                                                                    endPoint(0), endPoint(1), endPoint(2))
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertLine
    Set ObjVertLine = Nothing
    
    'Point 3D at control point
    xValue = 0
    yValue = 0
    zValue = parCentertoDavitCover
    Set ObjOriginPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, xValue, yValue, zValue)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOriginPoint
    Set ObjOriginPoint = Nothing
    
    'Default Surface at Control Point
    normalX = 1
    normalY = 0
    normalZ = 0
    
    xValue = 0
    yValue = 0
    zValue = parCentertoDavitCover
    
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                                xValue, yValue, zValue, _
                                                                                normalX, normalY, normalZ, _
                                                                                0.25 * pipeDiam)
    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set oCircle = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub


