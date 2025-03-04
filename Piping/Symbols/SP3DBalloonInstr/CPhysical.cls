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
'   Author:          BG
'   Creation Date:  Thursday, May 30 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
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
    Dim ObjInstrBody As Object
    Dim ObjStem As Object
    Dim objCone As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset As Double
    Dim parConeHeight As Double
    Dim parConeDiameter1 As Double
    Dim parConeDiameter2 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parOffset = arrayOfInputs(4)    'P4
    parConeHeight = arrayOfInputs(5)    'P5
    parConeDiameter1 = arrayOfInputs(6) 'P6
    parConeDiameter2 = arrayOfInputs(7) 'P7
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0

 ' Insert your code for output 1(Instruments Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim startPoint As AutoMath.DPosition
    Dim endPoint  As AutoMath.DPosition
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set -parFace1toCenter + flangeThick, 0, 0
    endPoint.Set parFace2toCenter - flangeThick, 0, 0
    Set ObjInstrBody = PlaceCylinder(m_OutputColl, startPoint, endPoint, pipeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInstrBody
    Set ObjInstrBody = Nothing
    
 ' Insert your code for output 2(Actuator Stem)
    Dim StemDiameter As Double
    'Stem Diameter is taken as one-fifth of minimum of Pipe and Cone bottom diameter
    If parConeDiameter2 < pipeDiam Then
        StemDiameter = parConeDiameter2 / 5
    Else
        StemDiameter = pipeDiam / 5
    End If
    
    startPoint.Set 0, 0, 0
    endPoint.Set 0, 0, parOffset - parConeHeight
    Set ObjStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, StemDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
 ' Insert your code for output 3(Cone)
    Dim basecenter As AutoMath.DPosition
    Dim topcenter As AutoMath.DPosition
    
    Set basecenter = New AutoMath.DPosition
    Set topcenter = New AutoMath.DPosition
    
    basecenter.Set 0, 0, parOffset - parConeHeight
    topcenter.Set 0, 0, parOffset
    Set objCone = PlaceCone(m_OutputColl, basecenter, topcenter, parConeDiameter2 / 2, _
                                                                parConeDiameter1 / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing
    Set basecenter = Nothing
    Set topcenter = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
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
