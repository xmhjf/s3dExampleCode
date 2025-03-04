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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CMaintainence.cls
'   Author:         RRK/KKC
'   Creation Date:  Monday, Apr 16 2007
'   Description:
'   This symbol is created for manual hydraulic operator of Linatex type LT pinch valve
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
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

    Dim parOperatorHeight As Double
    Dim parOperatorLength As Double
    Dim parOperatorDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorLength = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)

    iOutput = 0
    
 ' Insert your code for output 5(Box)
    Dim ObjMaintBox As Object
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
'Assumption
    Dim dStemHeight As Double
    dStemHeight = 0.5 * parOperatorHeight
    
    oStPoint.Set parOperatorDiameter / 2, dStemHeight, parOperatorDiameter / 2
    oEnPoint.Set parOperatorLength, parOperatorHeight, -parOperatorDiameter / 2
    
    Set ObjMaintBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintBox
    Set ObjMaintBox = Nothing
    
' Insert your code for output 6(Cylinder)
    Dim ObjMaintCylinder As Object

    Dim dBoxHeight As Double
    Dim dLeverDia As Double
    Dim dLeverLength As Double
    Dim dFrameThickness As Double
    
'Assumptions
    dBoxHeight = 0.25 * parOperatorHeight
    dLeverDia = 0.4 * parOperatorDiameter
    dLeverLength = 0.5 * parOperatorHeight
    dFrameThickness = 0.05 * parOperatorHeight
    
    oStPoint.Set parOperatorLength - dLeverDia, dStemHeight + dFrameThickness + dBoxHeight, 0
    oEnPoint.Set parOperatorLength, dStemHeight + dFrameThickness + dBoxHeight, 0
    
    Set ObjMaintCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * dLeverLength, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCylinder
    
    Set ObjMaintCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
