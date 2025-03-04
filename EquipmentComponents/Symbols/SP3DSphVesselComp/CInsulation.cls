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
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is Spherical Vessel Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E230 Symbol in Page no 293.
'    Symbol is created using the following Outputs:
'    i)  Three standard outputs Consisting of the following:
'       a) One Insulation aspect output,
'       b) One Physical aspect output: A Vessel uses 'PlaceSphere',
'       c) Two ReferenceGeometry aspect outputs: Default surface and a control point.
'    ii) Variable number of Supports (Can be either cylindrical or Cuboid) and
'       default surfaces (for cylindrical supportscase) are computed as per the user input.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object

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

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)     'P2
    parSupportAngularLocation = arrayOfInputs(3) 'P3
    parNumberOfSupports = arrayOfInputs(4)   'P4
    parVesselCenterHeight = arrayOfInputs(5) 'P5
    parSupportLength = arrayOfInputs(6)      'P6
    parSupportThickness = arrayOfInputs(7)   'P7
    parSupportRadialLocation = arrayOfInputs(8) 'P8
    parSupportHeight = arrayOfInputs(9)      'P9
    parInsulationThickness = arrayOfInputs(10)

    iOutput = 0

'   Origin is taken at Place point1(DP/PP 1)

' Insert your code for output 1(InsulatedBody)
    Dim centPoint As New AutoMath.DPosition

    centPoint.Set 0, 0, parVesselCenterHeight
    dInsulationDiameter = parVesselDiameter + 2 * parInsulationThickness
    Set ObjInsulatedBody = PlaceSphere(m_OutputColl, centPoint, dInsulationDiameter / 2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set centPoint = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
