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
'   Author:         BBA/MS
'   Creation Date:  Thursday, November 4 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'
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
    Dim ConduitDiam     As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjNozzle1 As GSCADNozzleEntities.IJConduitPortOcc
    Dim ObjNozzle2 As GSCADNozzleEntities.IJConduitPortOcc
    Dim ObjNozzle3 As GSCADNozzleEntities.IJConduitPortOcc
    Dim ObjNozzle4 As GSCADNozzleEntities.IJConduitPortOcc
    
    Dim parFacetoCenter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    
    iOutput = 0
    
' Insert your code for output 1(Cylinder Body 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim stPoint1 As AutoMath.DPosition
    Dim enPoint1 As AutoMath.DPosition
    Dim ObjCrossBody1 As Object
    Set stPoint1 = New AutoMath.DPosition
    Set enPoint1 = New AutoMath.DPosition
    Dim ConduitOD1 As Double
    
    ConduitOD1 = BodyOD(flangeDiam, ConduitDiam)
    
    stPoint1.Set -parFacetoCenter, 0, 0
    enPoint1.Set parFacetoCenter, 0, 0
    
    Set ObjCrossBody1 = PlaceCylinder(m_OutputColl, stPoint1, enPoint1, ConduitOD1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCrossBody1
    Set ObjCrossBody1 = Nothing
    
'Insert your code for output 2(Cylinder Body 2)
    RetrieveParameters 3, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim stPoint2 As AutoMath.DPosition
    Dim enPoint2 As AutoMath.DPosition
    Dim ObjCrossBody2 As Object
    Set stPoint2 = New AutoMath.DPosition
    Set enPoint2 = New AutoMath.DPosition
    Dim ConduitOD3 As Double
    
    ConduitOD3 = BodyOD(flangeDiam, ConduitDiam)
    
    stPoint2.Set 0, -parFacetoCenter, 0
    enPoint2.Set 0, parFacetoCenter, 0
    
    Set ObjCrossBody2 = PlaceCylinder(m_OutputColl, stPoint2, enPoint2, ConduitOD3, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCrossBody2
    Set ObjCrossBody2 = Nothing
 
 ' Insert your code for output 3(Nozzle1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set ObjNozzle1 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 1)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    Set ObjNozzle1 = Nothing
 
 ' Insert your code for output 4(Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oDir.Set 1, 0, 0
    oPlacePoint.Set parFacetoCenter + sptOffset - depth, 0, 0
    
    Set ObjNozzle2 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing

 ' Insert your code for output 5(Nozzle3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oDir.Set 0, 1, 0
    oPlacePoint.Set 0, parFacetoCenter + sptOffset - depth, 0
    
    Set ObjNozzle3 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 3)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle3
    Set ObjNozzle3 = Nothing
   
' Insert your code for output 5(Nozzle4)
    RetrieveParameters 4, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oDir.Set 0, -1, 0
    oPlacePoint.Set 0, -parFacetoCenter - sptOffset + depth, 0
    
    Set ObjNozzle4 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 3)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle4
    Set ObjNozzle4 = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub

