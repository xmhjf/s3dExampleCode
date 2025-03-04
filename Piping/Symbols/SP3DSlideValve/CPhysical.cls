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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Tuesday, Sept 14 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Slide Valve details were taken from PDS Piping Component Data Reference Manual
'   at Page No - D-21 and SN=V33. This Symbol consist of Physical and Insulation aspects.
'   Physical aspect consist of a Nozzle with Length , Nozzle, branch Cylinder ,a Box and a output for Valve Operator
'   Insulation aspect consist of a Cylinder for main body ,another cylinder for branch and a Box for
'   Projection at top of Branch'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjCylinder As Object
    Dim ObjProjectionBox As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parBranchLength As Double
    Dim parProjectionHeight As Double
    Dim parProjectionLength As Double
    Dim parProjectionWidth As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'A
    parFace2toCenter = arrayOfInputs(3) 'B
    parBranchLength = arrayOfInputs(4)  'J
    parProjectionHeight = arrayOfInputs(5)  'K
    parProjectionLength = arrayOfInputs(6)  'M
    parProjectionWidth = arrayOfInputs(7)   'N
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
 ' Insert your code for output 1(Branch Cylinder)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set 0, 0, 0
    endPoint.Set 0, 0, parBranchLength - parProjectionHeight
    Set ObjCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, pipeDiam1, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
 ' Insert your code for output 2(Projection Box)
    Dim lowerpt As AutoMath.DPosition
    Dim upperpt As AutoMath.DPosition
    
    Set lowerpt = New AutoMath.DPosition
    Set upperpt = New AutoMath.DPosition
    
    lowerpt.Set -parProjectionLength / 2, -parProjectionWidth / 2, parBranchLength - parProjectionHeight
    upperpt.Set parProjectionLength / 2, parProjectionWidth / 2, parBranchLength
    Set ObjProjectionBox = PlaceBox(m_OutputColl, lowerpt, upperpt)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjProjectionBox
    Set ObjProjectionBox = Nothing
    Set lowerpt = Nothing
    Set upperpt = Nothing

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleLength As Double
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    NozzleLength = parFace1toCenter + parFace2toCenter - flangeThick2
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFace2toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


