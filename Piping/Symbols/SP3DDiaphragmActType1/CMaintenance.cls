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
''      This is a diaphragm Actuator symbol created to be associated with Pressure Reducing valves.

'   Change History:
'   dd.mmm.yyyy     who      change description
'   -----------    -----    ------------------
'   23.Apr.2007     dkl      CR-118766  Created the symbol.
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
    Dim parBaseHeight As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parBaseHeight = arrayOfInputs(4)
    
    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0

    'The Actuator is located at a distance of Base height from the origin along Y-Axis.
    Dim oCenterPos As AutoMath.DPosition
    Set oCenterPos = New AutoMath.DPosition
    oCenterPos.Set 0, parBaseHeight, 0
    
' Insert your code for output 4, Maintenance cylinder
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim dHeightofActuator As Double 'Height of actuator from base Cylinder bottom to
                                    'the top of the actuator.
    dHeightofActuator = parOperatorHeight - parBaseHeight
    
    'Assumptions: Height of the Maintenance-cylinder is (7/6) of Height of actuator.
       
    oStPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oStPoint.x, oStPoint.y + dHeightofActuator * 7 / 6, oStPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, _
                                                oEnPoint, parOperatorDiameter * 1.1
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
        
End Sub
