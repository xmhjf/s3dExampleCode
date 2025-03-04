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
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:   Thursday, Feb 2 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the Model No. TTF-1R Trap (Type: Thermostatic Steam Trap)
'    in Armstrong's catalog available at URL http://www.armstrong-intl.com
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI       As Double
Dim m_oSymGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "Physical:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
      PI = 4 * Atn(1)
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjDome As Object
    
    Dim parTrapBodyDiameter As Double
    Dim parSTHeight As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parTrapLength1 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parTrapBodyDiameter = arrayOfInputs(2)   'A
    parSTHeight = arrayOfInputs(3)       'B
    parFace1toCenter = arrayOfInputs(4)  'D
    parFace2toCenter = arrayOfInputs(5)  'C
    parTrapLength1 = arrayOfInputs(6)    'H
    parInsulationThickness = arrayOfInputs(7)

    iOutput = 0

    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
'   Assumption: Origin of the symbol is assumed to be at the intersection of Port1 axis & Port2 axis

 ' Insert your code for output 7(Vertical Cylinder)
    Dim dDomeHeight As Double
    Dim dCylinderHeightFromOrigin As Double
    Dim dNozzleLength As Double
     
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

'   Assumption: Nozzle length (of Nozzle 1 and Nozzle 2) is taken to be equal to (H - A) i.e. parTrapLength1 - parTrapBodyDiameter
    dNozzleLength = parTrapLength1 - parTrapBodyDiameter
    If CmpDblGreaterthan(flangeThick, dNozzleLength) Then dNozzleLength = flangeThick 'Helps to protect CreateNozzleWithLength function call
    
'   Assumption: Dome height is taken to be 1/4th of parTrapBodyDiameter (A)
    dDomeHeight = parTrapBodyDiameter / 4
    dCylinderHeightFromOrigin = (parSTHeight - parFace2toCenter) - dDomeHeight
    stPoint.Set 0, dCylinderHeightFromOrigin, 0
    enPoint.Set 0, -parFace2toCenter + dNozzleLength, 0

' Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parTrapBodyDiameter
    
 ' Insert your code for output 8(Dome)
    Dim objEllipse As Object
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim mMRatio As Double

    mMRatio = dDomeHeight / (parTrapBodyDiameter / 2)

    Set objEllipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                         0, dCylinderHeightFromOrigin, 0, _
                                         0, 0, 1, _
                                parTrapBodyDiameter / 2, 0, 0, _
                                mMRatio, PI / 2, PI / 2)
    
    Dim oRevVect As AutoMath.DVector
    Dim oCenterPoint As AutoMath.DPosition
    Set oRevVect = New AutoMath.DVector
    Set oCenterPoint = New AutoMath.DPosition
    oRevVect.Set 0, 1, 0
    oCenterPoint.Set 0, dCylinderHeightFromOrigin, 0
    Set ObjDome = PlaceRevolution(m_OutputColl, objEllipse, oRevVect, oCenterPoint, PI * 2, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDome
    Set ObjDome = Nothing
    Set oRevVect = Nothing
    Set oCenterPoint = Nothing
    Set objEllipse = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
'   'D' is taken as nozzle length to ensure that there is no gap between verical cylinder and nozzle pipe
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace1toCenter)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set 0, -parFace2toCenter - sptOffset + depth, 0
    oDir.Set 0, -1, 0

    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
