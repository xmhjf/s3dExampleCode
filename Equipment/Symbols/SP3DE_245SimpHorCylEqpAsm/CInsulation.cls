VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:01 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RUK
'   Creation Date:  Wednesday, March 7 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'     Body of Vessel
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim ObjInsulatedVessel As Object

    Dim parVesselTantoTan As Double
    Dim parVesselDiameter As Double
    Dim parVesselCenterHeight As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselTantoTan = arrayOfInputs(4)        'P2
    parVesselDiameter = arrayOfInputs(5)        'P3
    parVesselCenterHeight = arrayOfInputs(6)    'P4
    parInsulationThickness = arrayOfInputs(20)

    iOutput = 0
    
    'Origin is taken at left dome center point (i.e. Placement point 1)
    'Vessel is formed by creating a Cylinder and two Elliptical Heads rotating about X-Axis

    'Insert your code for InsulatedVessel
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set parVesselTantoTan, 0, 0

    'Create Insulated Horizontal Drum
    Dim dInsulationDiameter As Double

    dInsulationDiameter = parVesselDiameter + 2 * parInsulationThickness
    Set ObjInsulatedVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVessel
    Set ObjInsulatedVessel = Nothing

    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
