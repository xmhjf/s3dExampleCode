VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         dkl
'   Creation Date:  Monday, Apr 23 2007
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      This Actuator symbol is created to made to be associated with Type D and Type P Solenoid Valves
'      as per JIS B8471.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   23.Apr.2007     dkl           CR-113805. Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
End Sub
Private Sub Class_terminate()
    Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)

    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0
   
' Insert your code for output 1, Solenoid-cylinder at the top.
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Assumption: Height of the Solenoid-cylinder is 40% of Operator height.
    Dim dHtOfSolenoidCyl As Double
    dHtOfSolenoidCyl = parOperatorHeight * 0.4
    
    oStPoint.Set 0, parOperatorHeight, 0
    oEnPoint.Set oStPoint.x, oStPoint.y - dHtOfSolenoidCyl, oStPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, _
                                                oEnPoint, parOperatorDiameter

' Insert your code for output 2, Piston-cylinder in the middle.
    'Assumption: 1. Height of the Piston-cylinder is 30% of Operator height.
    '            2. Diameter of the Piston-cylinder is 25% of Operator Diameter.
    Dim dHtOfPistonCyl As Double
    dHtOfPistonCyl = parOperatorHeight * 0.3
    
    oStPoint.Set 0, parOperatorHeight - dHtOfSolenoidCyl, 0
    oEnPoint.Set oStPoint.x, oStPoint.y - dHtOfPistonCyl, oStPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, _
                                                oEnPoint, 0.25 * parOperatorDiameter

' Insert your code for output 3, Stem at the bottom.
    Dim dMinimumStemDiameter As Double
    dMinimumStemDiameter = parOperatorDiameter * 0.1
    'Check that the stem diameter is within 15mm and 50mm
    If CmpDblLessThan(dMinimumStemDiameter, 0.015) Then dMinimumStemDiameter = 0.015
    If CmpDblLessThan(dMinimumStemDiameter, 0.05) Then dMinimumStemDiameter = 0.05
    
    oStPoint.Set 0, parOperatorHeight - dHtOfSolenoidCyl - dHtOfPistonCyl, 0
    oEnPoint.Set oStPoint.x, 0, oStPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, _
                                                oEnPoint, dMinimumStemDiameter

' Insert your code for output 4, Cable-cylinder to the left attached to the Solenoid Cylinder.
    'Assumption: 1. Length of the Cable-cylinder is 10% of Operator Diameter.
    '            2. Diameter of the Cable-cylinder is 15% of Operator Height.
    '            3. The Cable-cylinder center is located at a distance of its diameter
    '               from the top of the Solenoid-Cylinder.
    
    Dim dLenOfCableCyl As Double
    dLenOfCableCyl = parOperatorDiameter * 0.1
    
    oStPoint.Set 0, parOperatorHeight - 0.15 * parOperatorHeight, 0
    oEnPoint.Set -(oStPoint.x + parOperatorDiameter / 2 + dLenOfCableCyl), oStPoint.y, oStPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, _
                                                oEnPoint, 0.15 * parOperatorHeight
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
    
End Sub

