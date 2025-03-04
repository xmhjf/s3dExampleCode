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
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    
    Dim parFacetoFace As Double
    Dim parSealExtLength As Double
    Dim parSealExtDiameter As Double
    Dim parSealSmallExtDiameter As Double
    Dim Angle As Double
    
    Angle = PI / 4
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parSealExtLength = arrayOfInputs(3)
    parSealExtDiameter = arrayOfInputs(4)
    parSealSmallExtDiameter = arrayOfInputs(5)
    
    iOutput = 0
 
    
' Insert your code for output 1(Main Cylinder Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim stPoint1 As AutoMath.DPosition
    Dim enPoint1 As AutoMath.DPosition
    Dim ObjCylBody1 As Object
    Set stPoint1 = New AutoMath.DPosition
    Set enPoint1 = New AutoMath.DPosition
    Dim ConduitOD1 As Double
    
    ConduitOD1 = BodyOD(flangeDiam, ConduitDiam)
    
    stPoint1.Set -parFacetoFace / 2, 0, 0
    enPoint1.Set parFacetoFace / 2, 0, 0
    
    Set ObjCylBody1 = PlaceCylinder(m_OutputColl, stPoint1, enPoint1, ConduitOD1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylBody1
    Set ObjCylBody1 = Nothing
    Set stPoint1 = Nothing
    Set enPoint1 = Nothing
    
'Insert your code for output 2(Seal Cylinder Body)
    Dim stPoint2 As AutoMath.DPosition
    Dim enPoint2 As AutoMath.DPosition
    Set stPoint2 = New AutoMath.DPosition
    Set enPoint2 = New AutoMath.DPosition
    Dim ObjCylBody2 As Object
    
    stPoint2.Set -0.25 * parFacetoFace, 0, 0
    enPoint2.Set -0.25 * parFacetoFace, 0.65 * parSealExtLength, 0
    
    Set ObjCylBody2 = PlaceCylinder(m_OutputColl, stPoint2, enPoint2, parSealExtDiameter, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylBody2
    Set ObjCylBody2 = Nothing
    Set stPoint2 = Nothing
    Set enPoint2 = Nothing
    
 'Insert your code for output 3(Small Seal Cylinder Body)
    Dim stPoint3 As AutoMath.DPosition
    Dim enPoint3 As AutoMath.DPosition
    Set stPoint3 = New AutoMath.DPosition
    Set enPoint3 = New AutoMath.DPosition
    Dim ObjCylBody3 As Object
    
    stPoint3.Set 0, 0, 0
    enPoint3.Set (1.25 * parSealExtLength) * Cos(Angle), (1.25 * parSealExtLength) * Sin(Angle), 0
    
    Set ObjCylBody3 = PlaceCylinder(m_OutputColl, stPoint3, enPoint3, 0.25 * parSealSmallExtDiameter, True)
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylBody3
    Set ObjCylBody3 = Nothing
    Set stPoint3 = Nothing
    Set enPoint3 = Nothing
 
 
' Insert your code for output 4(Nozzle1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oDir.Set -1, 0, 0
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    
    Set ObjNozzle2 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 1)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing
    
' Insert your code for output 5(Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    oDir.Set 1, 0, 0
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    
    Set ObjNozzle2 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
