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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Wednesday, May 29 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: S70AC, Page No. 107 of PDS Piping Component Data Reference Guide.
' SP3DSteamTrap Symbol is created with Ten Outputs of which six insualation aspect outputs.
' Other four are physical aspect outputs created as follows:
'   ObjBody - Using 'PlaceCylinder' function,
'   ObjConnector - Using 'PlaceCylinder' function and
'   two ObjNozzle objects  by using 'CreateNozzleWithLength' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjConnector As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset As Double
    Dim parTrapDiameter As Double
    Dim parTrapLength1 As Double
    Dim parTrapLength2 As Double
    Dim parInsulationThickness  As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parOffset = arrayOfInputs(4)
    parTrapDiameter = arrayOfInputs(5)
    parTrapLength1 = arrayOfInputs(6)
    parTrapLength2 = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0

' Insert your code for output 7(Body of Trap)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim ConnectorDiameter As Double
    
    stPoint.Set 0, -parTrapLength1, parOffset
    enPoint.Set 0, parTrapLength2, parOffset
    Set ObjBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parTrapDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Insert your code for output 8(Connector from ports to Trap)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, parOffset
'   Set ConnectorDiameter equal to 0.2 * minimum of pipeDiam and parTrapDiameter
    ConnectorDiameter = 0.2 * parTrapDiameter
    If parTrapDiameter < pipeDiam Then ConnectorDiameter = 0.2 * pipeDiam
    Set ObjConnector = PlaceCylinder(m_OutputColl, stPoint, enPoint, ConnectorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector
    Set ObjConnector = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
'   oPlacePoint.Set -faceToFace / 2 - sptOffset + depth, 0, 0
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace1toCenter)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

'   oPlacePoint.Set faceToFace / 2 + sptOffset - depth, 0, 0
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace2toCenter)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

