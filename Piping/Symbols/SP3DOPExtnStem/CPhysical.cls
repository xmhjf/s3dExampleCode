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
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         ACM
'   Creation Date:  Friday, Dec 20 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Operator Extension Stem symbol is prepared based on SP3D Concept design document
''  "Valve Accessories"(valveaccessoriescd.doc).
''  This symbol consists of 2 inputs and 4 outputs all are physical.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   18.Jun.2004     MS     (RI 32889) Changed input parameter name from
'                                           HandWheelDia to HandWheelDiameter(as an existing interface can be used instead of creating a new interface for HandWheelDia)
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
    
    Dim iOutput     As Double
    Dim ObjHandWheelTorus As Object
    Dim ObjSpokeCyl1 As Object
    Dim ObjSpokeCyl2 As Object
    Dim ObjStemRodCyl As Object
    
    Dim parOperatorHeight As Double
    Dim parHandWheelDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parHandWheelDiameter = arrayOfInputs(3)
    
    iOutput = 0

 ' Insert your code for output 1(Hand wheel as Torus)
   ''Assumed Hand Wheel Cross-section Circuar dia as 10mm(thickness)
   Dim dTorusMinorDia As Double
   dTorusMinorDia = 0.01
   
   Dim oTorusCentre As New AutoMath.DPosition
   Dim oTorusAxis As New AutoMath.DVector
      
   oTorusCentre.Set 0, parOperatorHeight, 0
   oTorusAxis.Set 0, 1, 0

   Set ObjHandWheelTorus = PlaceTorus(m_OutputColl, oTorusCentre, oTorusAxis, _
                            (parHandWheelDiameter / 2) - dTorusMinorDia / 2, dTorusMinorDia / 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheelTorus
    Set ObjHandWheelTorus = Nothing
    
 ' Insert your code for output 2(Spoke Cylinder1)
    '' Assumed that spoke diameter is 5mm
   Dim stPoint As New AutoMath.DPosition
   Dim enPoint As New AutoMath.DPosition
   
   Dim dSpokeDia As Double
   
   dSpokeDia = 0.005
   
   stPoint.Set 0, parOperatorHeight, parHandWheelDiameter / 2 - dTorusMinorDia / 2
   enPoint.Set 0, parOperatorHeight, -parHandWheelDiameter / 2 + dTorusMinorDia / 2
   
   Set ObjSpokeCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dSpokeDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSpokeCyl1
    Set ObjSpokeCyl1 = Nothing
    
 ' Insert your code for output 3(Spoke Cylinder2)
   stPoint.Set parHandWheelDiameter / 2 - dTorusMinorDia / 2, parOperatorHeight, 0
   enPoint.Set -parHandWheelDiameter / 2 + dTorusMinorDia / 2, parOperatorHeight, 0
   
   Set ObjSpokeCyl2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dSpokeDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSpokeCyl2
    Set ObjSpokeCyl2 = Nothing
    
 ' Insert your code for output 4(Stem Rod Cylinder)
   Dim dStemDia As Double
   dStemDia = 0.005
   
   stPoint.Set 0, 0, 0
   enPoint.Set 0, parOperatorHeight, 0
   
   Set ObjStemRodCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, dStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemRodCyl
    Set ObjStemRodCyl = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oTorusAxis = Nothing
    Set oTorusCentre = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
