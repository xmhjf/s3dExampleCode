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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:21 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RUK
'   Creation Date:  Wednesday, Mar 14 2007
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

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt           As PartFacelets.IJDPart
    Dim iOutput             As Double
    Dim ObjInsulatedBody    As Object

    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports As Long
    Dim parVesselCenterHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parSupportRadialLocation As Double
    Dim parSupportHeight As Double
    Dim parInsulationThickness As Double

    Dim dInsulationDiameter As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)            'P2
    parVesselCenterHeight = arrayOfInputs(3)        'P5
    parSupportLength = arrayOfInputs(4)             'P6
    parSupportHeight = arrayOfInputs(5)             'P9
    parSupportThickness = arrayOfInputs(6)          'P7
    parNumberOfSupports = arrayOfInputs(7)          'P4
    parSupportAngularLocation = arrayOfInputs(8)    'P3
    parSupportRadialLocation = arrayOfInputs(9)     'P8
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0
    'Origin is taken at Place point1(DP/PP 1)
    'Insert your code for InsulatedBody
    Dim centPoint As New AutoMath.DPosition

    centPoint.Set 0, 0, parVesselCenterHeight
    dInsulationDiameter = parVesselDiameter + 2 * parInsulationThickness
    Set ObjInsulatedBody = PlaceSphere(m_OutputColl, centPoint, dInsulationDiameter / 2)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set centPoint = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

