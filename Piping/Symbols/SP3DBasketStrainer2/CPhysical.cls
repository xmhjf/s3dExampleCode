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
'   Author:         svsmylav
'   Creation Date:  Tuesday 31, Jan 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the catalog available at URL http://www.strainwell.com/conical-stariners.html

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_oSymGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "Physical:" 'Used for error messages
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim ObjCollarExtn As Object

    Dim parStrDiameter As Double
    Dim parEndDiameter As Double
    Dim parBasketEndDiameter As Double
    Dim parBasketLength As Double
    Dim parCollarThickness As Double
    Dim parArmLength As Double
    Dim parStrWidth As Double
    Dim parCylinderLength As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parStrDiameter = arrayOfInputs(2) 'D1
    parEndDiameter = arrayOfInputs(3) 'D2
    parBasketEndDiameter = arrayOfInputs(4) 'D3
    parBasketLength = arrayOfInputs(5)      'L
    parCollarThickness = arrayOfInputs(6)   'T
    parArmLength = arrayOfInputs(7)
    parStrWidth = arrayOfInputs(8)
    parCylinderLength = arrayOfInputs(9) 'Body at the small end of the basket
    parInsulationThickness = arrayOfInputs(10)

    iOutput = 0

    m_oSymGeomHelper.OutputCollection = m_OutputColl

'   Assumption: Origin of the symbol is assumed to be at half of the parCollarThickness
    
 ' Insert your code for output 2(Strainer)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -parCollarThickness / 2, 0, 0
    enPoint.Set parCollarThickness / 2 + parBasketLength, 0, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCone arrayOfOutputs(iOutput), stPoint, enPoint, parEndDiameter, parBasketEndDiameter

 ' Insert your code for output 3(Strainer End)
    stPoint.Set -parCollarThickness / 2 + parBasketLength, 0, 0
    enPoint.Set stPoint.x + parCylinderLength, stPoint.y, stPoint.z
    
' Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parBasketEndDiameter
 
 ' Insert your code for output 4(Collar)
    stPoint.Set -parCollarThickness / 2, 0, 0
    enPoint.Set parCollarThickness / 2, 0, 0

' Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parStrDiameter

 ' Insert your code for output 5(CollarExtn)
    Dim dCollarMeetHeight As Double
    
    dCollarMeetHeight = Sqr((parStrDiameter / 2) ^ 2 - (parStrWidth / 2) ^ 2)
    
    stPoint.Set -parCollarThickness / 2, dCollarMeetHeight, -parStrWidth / 2
    enPoint.Set parCollarThickness / 2, parStrDiameter / 2 + parArmLength, parStrWidth / 2
    
    Set ObjCollarExtn = PlaceBox(m_OutputColl, stPoint, enPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollarExtn
    Set ObjCollarExtn = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parCollarThickness / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parCollarThickness / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

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
