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
    
    Dim parFace1toExpJoint As Double
    Dim parExpJointLength1 As Double
    Dim parExpJointDiameter As Double
    Dim parExpJointLength2 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toExpJoint = arrayOfInputs(2)
    parExpJointLength1 = arrayOfInputs(3)
    parExpJointDiameter = arrayOfInputs(4)
    parExpJointLength2 = arrayOfInputs(5)
    
    iOutput = 0
     
' Insert your code for output 1(Cylinder Body 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Dim ObjCylBody1 As Object
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    Dim ConduitOD1 As Double
    
    ConduitOD1 = BodyOD(flangeDiam, ConduitDiam)
    
    stPoint.Set 0, 0, 0
    enPoint.Set parFace1toExpJoint, 0, 0
    
    Set ObjCylBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, ConduitOD1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylBody1
    Set ObjCylBody1 = Nothing

    
'Insert your code for output 2(Cylinder Body 2)
    Dim ObjCylBody2 As Object
    stPoint.Set parFace1toExpJoint, 0, 0
    enPoint.Set parFace1toExpJoint + parExpJointLength1, 0, 0
    
    Set ObjCylBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parExpJointDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylBody2
    Set ObjCylBody2 = Nothing
 
'Insert your code for output 3(Cylinder Body 3)
    Dim ObjCylBody3 As Object
    
    stPoint.Set parFace1toExpJoint + parExpJointLength1, 0, 0
    enPoint.Set parFace1toExpJoint + parExpJointLength1 + parExpJointLength2, 0, 0
    
    Set ObjCylBody3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, ConduitOD1, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylBody3
    Set ObjCylBody3 = Nothing
    
'Insert your code for output 4(Cylinder Body 4)
    Dim ObjCylBody4 As Object
    
    stPoint.Set parFace1toExpJoint + parExpJointLength1 + parExpJointLength2, 0, 0
    enPoint.Set (2 * parFace1toExpJoint) + parExpJointLength1 + parExpJointLength2, 0, 0
    
    Set ObjCylBody4 = PlaceCylinder(m_OutputColl, stPoint, enPoint, 0.8 * parExpJointDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylBody4
    Set ObjCylBody4 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Insert your code for output 5(Nozzle1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set ObjNozzle1 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 1)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    Set ObjNozzle1 = Nothing
 
 ' Insert your code for output 6(Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    oDir.Set 1, 0, 0
    
    oPlacePoint.Set (2 * parFace1toExpJoint) + parExpJointLength1 + parExpJointLength2 + sptOffset - depth, 0, 0
    
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


