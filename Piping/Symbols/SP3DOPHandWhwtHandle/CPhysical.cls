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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          KKC
'   Creation Date:  Tuesday, Jun 10 2008
'   Description:
'    The symbol is prepared based on 120 deg angle valve
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  10.Jun.2008     KKC      CR-144044  Provide a 120 deg angle valve
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
    Dim iOutput     As Double
    
    Dim parOperatorHeight As Double
    Dim dStemMinimumDiameter As Double
    Dim parOperatorDiameter As Double
    Dim parLeverLength As Double
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parLeverLength = arrayOfInputs(4)
 
'   Initialize
    iOutput = 0
    dStemMinimumDiameter = parOperatorDiameter * 0.2
   '   Check that the stem diameter is within 15mm and 50mm
   If CmpDblLessThan(dStemMinimumDiameter, 0.015) Then dStemMinimumDiameter = 0.015
   If CmpDblGreaterthan(dStemMinimumDiameter, 0.05) Then dStemMinimumDiameter = 0.05

' Insert your code for output 1(Operator Cylinder)
'   Place Cylinder
    Dim ObjOperatorCylinder As Object
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, 0.75 * parOperatorHeight, 0
    Set ObjOperatorCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parOperatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOperatorCylinder
    Set ObjOperatorCylinder = Nothing
    
' Insert your code for output 2(Stem)
'   Place Cylinder
    Dim ObjStem As Object
    oStPoint.Set 0, 0.75 * parOperatorHeight, 0
    oEnPoint.Set 0, 0.9 * parOperatorHeight, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemMinimumDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 3(Handle Cylinder)
'   Place Cylinder
    Dim ObjStemCylinder As Object
    oStPoint.Set 0, 0.81 * parOperatorHeight, 0
    oEnPoint.Set 0, 0.87 * parOperatorHeight, 0
    Set ObjStemCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * dStemMinimumDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemCylinder
    Set ObjStemCylinder = Nothing
    
' Insert your code for output 4(Box to support Handle)
'   Place Cylinder
    Dim ObjHandleHolder As Object
    oStPoint.Set 0, 0.84 * parOperatorHeight, 0.5 * dStemMinimumDiameter
    oEnPoint.Set parLeverLength, 0.87 * parOperatorHeight, -0.5 * dStemMinimumDiameter
    Set ObjHandleHolder = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandleHolder
    Set ObjHandleHolder = Nothing

' Insert your code for the output 5 (Handle)
    Dim ObjHandle As Object
    oStPoint.Set parLeverLength - 0.4 * dStemMinimumDiameter, 0.84 * parOperatorHeight, 0
    oEnPoint.Set parLeverLength - 0.4 * dStemMinimumDiameter, parOperatorHeight, 0
    Set ObjHandle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.8 * dStemMinimumDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandle
    Set ObjHandle = Nothing
    
    Set oEnPoint = Nothing
    Set oStPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

