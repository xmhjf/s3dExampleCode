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
'   CMaintenance.cls
'   Author:         dkl
'   Creation Date:  Monday, Apr 23 2007
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who      change description
'   -----------    -----    ------------------
'   23.Apr.2007     dkl      CR-113805. Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages
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

' Insert your code for output 5, Maintenance cylinder covering the Solenoid-cylinder
' and the Cable-cylinder.
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Assumptions: 1. Height of the Solenoid-cylinder is 40% of Operator height.
    '             2. Length of the Cable-cylinder is 10% of Operator Diameter. Hence, the
    'diameter of this cylinder = twice the length of the Cable-cylinder + Operator Diameter.
    Dim dHtOfSolenoidCyl As Double
    dHtOfSolenoidCyl = parOperatorHeight * 0.4
    
    oStPoint.Set 0, parOperatorHeight, 0
    oEnPoint.Set oStPoint.x, oStPoint.y - dHtOfSolenoidCyl, oStPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, _
                                                oEnPoint, 1.2 * parOperatorDiameter
    
' Insert your code for output 6, Maintenance cone covering the Piston-cylinder and the Stem.
    'A value of 0.0000001 is used in X direction to avoid symbol placement failure(gives assertion errors).
    'This should be removed  once it is resolved.
    oStPoint.Set 0.0000001, parOperatorHeight - dHtOfSolenoidCyl, 0
    oEnPoint.Set 0, 0, 0

' Set the output
    iOutput = iOutput + 1

    m_oGeomHelper.CreateCone arrayOfOutputs(iOutput), oStPoint, _
                                                oEnPoint, parOperatorDiameter / 2, 0
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
        
End Sub
