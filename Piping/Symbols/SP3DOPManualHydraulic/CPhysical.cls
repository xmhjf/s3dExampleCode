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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK/KKC
'   Creation Date:  Monday, Apr 16 2007
'   Description:
'   This code is to create physical aspect of manual hydraulic operator of Linatex type LT
'   pinch valve
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
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
    Dim parOperatorLength As Double
    Dim parOperatorDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorLength = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)
    
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    iOutput = 0

 ' Insert your code for output 1(Spindle)
    Dim objStem As IngrGeom3D.Projection3d
    Dim dStemHeight As Double
    Dim dStemDia As Double
    
'Assumptions
    dStemDia = 0.15 * parOperatorDiameter
    dStemHeight = 0.5 * parOperatorHeight
    
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, dStemHeight, 0
    
    Set objStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStem
    Set objStem = Nothing
    
 ' Insert your code for output 2(Cylinder)
    Dim objCylinder As IngrGeom3D.Projection3d
    
    oStPoint.Set 0, dStemHeight, 0
    oEnPoint.Set 0, parOperatorHeight, 0
    
    Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parOperatorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    
 ' Insert your code for output 3(Frame)
    Dim objFrame As Object
    
    Dim dFrameThickness As Double
    Dim dFrameWidth As Double

'Assumptions
    dFrameThickness = 0.05 * parOperatorHeight
    dFrameWidth = 0.6 * parOperatorDiameter
    
    oStPoint.Set 0, dStemHeight, dFrameWidth / 2
    oEnPoint.Set parOperatorLength, dStemHeight + dFrameThickness, -dFrameWidth / 2
    
    Set objFrame = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFrame
    Set objFrame = Nothing
    
 ' Insert your code for output 4(Box for Control Mechanism)
    Dim objBox As Object
    
    Dim dBoxHeight As Double
    Dim dBoxLength As Double
        
'Assumptions
    dBoxHeight = 0.25 * parOperatorHeight
    dBoxLength = 0.25 * parOperatorLength
    
    oStPoint.Set parOperatorLength - dBoxLength, dStemHeight, dFrameWidth / 2
    oEnPoint.Set parOperatorLength, dStemHeight + dFrameThickness + dBoxHeight, _
                                                                            -dFrameWidth / 2
    
    Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
    Set objBox = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
