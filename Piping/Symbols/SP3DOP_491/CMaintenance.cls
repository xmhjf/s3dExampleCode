VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          MS
'   Creation Date:  Thursday, September 08 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   08.09.2004      MS  Added Maintenace Envelope to the operator
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double

    Dim parStemMinimumDiameter As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDepth As Double
    Dim parOperatorDiameter As Double
    Dim parActuatorOffset As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorDepth As Double
    
    Dim StemStPt        As New AutoMath.DPosition
    Dim StemEndPt       As New AutoMath.DPosition
    Dim MainCylStPt     As New AutoMath.DPosition
    Dim MainCylEndPt    As New AutoMath.DPosition
    Dim SecCylStPt      As New AutoMath.DPosition
    Dim SecCylEndPt     As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDepth = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)
    parActuatorOffset = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)
    parActuatorDepth = arrayOfInputs(7)

    
'   Initialize
    iOutput = 0

   parStemMinimumDiameter = parOperatorDiameter * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemMinimumDiameter < 0.015 Then parStemMinimumDiameter = 0.015
   If parStemMinimumDiameter > 0.05 Then parStemMinimumDiameter = 0.05
       
' Insert your code for output (Spindle)
    Dim ObjMaintCyl1 As IngrGeom3D.Projection3d
    StemStPt.Set 0, 0, 0
    StemEndPt.Set 0, parOperatorHeight - parOperatorDepth, 0
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, parStemMinimumDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing
    
' Insert your code for output (Main Cylinder)
    Dim ObjMaintCyl2 As IngrGeom3D.Projection3d
    MainCylStPt.Set 0, parOperatorHeight - parOperatorDepth, 0
    MainCylEndPt.Set 0, parOperatorHeight, 0
    Set ObjMaintCyl2 = PlaceCylinder(m_OutputColl, MainCylStPt, MainCylEndPt, parOperatorDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl2
    Set ObjMaintCyl2 = Nothing
    
' Insert your code for output (Second Cylinder)
    Dim ObjMaintCyl3 As IngrGeom3D.Projection3d
    SecCylStPt.Set (parOperatorDiameter + parActuatorDiameter) / 2, _
                    parOperatorHeight - parOperatorDepth, _
                    0
    SecCylEndPt.Set (parOperatorDiameter + parActuatorDiameter) / 2, _
                    parOperatorHeight - parOperatorDepth + parActuatorDepth, _
                    0
    Set ObjMaintCyl3 = PlaceCylinder(m_OutputColl, SecCylStPt, SecCylEndPt, parActuatorDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl3
    Set ObjMaintCyl3 = Nothing
    Set StemStPt = Nothing
    Set StemEndPt = Nothing
    Set MainCylStPt = Nothing
    Set MainCylEndPt = Nothing
    Set SecCylStPt = Nothing
    Set SecCylEndPt = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
