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
'   Author:          ACM
'   Creation Date:  Wednesday, Oct 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ISSD3),
'   The symbol consists of 10 outputs,8 physical outputs
'   and 2 Insulation Output.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
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
    Dim sptOffset1       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    Dim sptOffset3       As Double
    Dim depth3           As Double
    
    Dim iOutput     As Double
    Dim ObjConicalBody1 As Object
    Dim ObjConicalBody2 As Object
    Dim ObjConicalBody3 As Object
    Dim ObjConicalActBody As Object
    Dim ObjActCylinder As Object
    
    Dim parFace1toCenter As Double ''D2
    Dim parFacetoFace As Double ''D3
    Dim parFace3toCenter As Double ''D4
    Dim parActuatorHeight As Double ''D5
    Dim parActuatorDiameter As Double ''D6
    Dim parActuatorHeight1 As Double ''D19
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parActuatorHeight = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)
    parActuatorHeight1 = arrayOfInputs(7)
    
    iOutput = 0

 ' Insert your code for output 1(ConicalBody1)
   RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
 
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
         
    stPoint.Set -parFace1toCenter + flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    
    Set ObjConicalBody1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConicalBody1
    Set ObjConicalBody1 = Nothing
    
 ' Insert your code for output 2(ConicalBody2)
   
   RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
          
    stPoint.Set parFacetoFace - parFace1toCenter - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    
    Set ObjConicalBody2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
   
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConicalBody2
    Set ObjConicalBody2 = Nothing
    
 ' Insert your code for output 3(ConicalBody3)
   
   RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset3, depth3
          
    stPoint.Set 0, -parFace3toCenter + flangeThick, 0
    enPoint.Set 0, 0, 0
    
    Set ObjConicalBody3 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConicalBody3
    Set ObjConicalBody3 = Nothing
    
 ' Insert your code for output 4(ConicalActBody)
        
    stPoint.Set 0, parActuatorHeight1 + parActuatorHeight, 0
    enPoint.Set 0, parActuatorHeight1, 0

    Set ObjConicalActBody = PlaceCone(m_OutputColl, stPoint, enPoint, parActuatorDiameter / 2, parFace1toCenter * 0.375, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConicalActBody
    Set ObjConicalActBody = Nothing
    
 ' Insert your code for output 5(ActCylinder)
   stPoint.Set 0, 0, 0
   enPoint.Set 0, parActuatorHeight1, 0
    
   Dim CylDia As Double
   CylDia = parFace1toCenter
   Set ObjActCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActCylinder
    Set ObjActCylinder = Nothing
    
' Place Nozzle 1
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    
    oPlacePoint.Set 0, -parFace3toCenter - sptOffset3 + depth3, 0
    oDir.Set 0, -1, 0

    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


