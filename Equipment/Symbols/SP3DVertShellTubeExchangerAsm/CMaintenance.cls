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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:08 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          BG
'   Creation Date:  Friday, Apr 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'The Symbol has three aspects viz., Physical, Insulation and Maintenance.Physical Aspect consist of
'1.Exchanger body, Front End of Exchanger,Rear end of Exchanger, Extension Joint(variable), Supports(variable)
'and four Nozzles at bottom of Rear End of Exchanger, Radial Nozzle at the Exchanger body, along X-axis
'in Rear Exchanger End, along Y-Axis in Rear Exchanger End.
'Insulation Aspect consist of Insulation for Exchanger, Rear end,front end,extension joint and all Nozzles.
'Maintenance Aspect consist of cylinder along z-axis for bundle Pulling
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   06.May.2005     svsmylav        CR-76070: Removed the Nozzle related parameters.
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    
    Dim ObjBundlePulling As Object
    
    Dim parExchangerLength As Double
    Dim parExchangerDiameter As Double
    Dim parBundleFlangeTk As Double
    Dim parBundleFlangeDia As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parExchangerFlangeTk2 As Double
    Dim parExtensionJointPosition As Double
    Dim parExtensionJointThickness As Double
    Dim parExtensionJointDiameter As Double
    Dim parBundlePullingLength As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupport As Double
    Dim parSupportDepth As Double
    Dim parSupportFromOrigin As Double
    Dim parSupportTopWidth As Double
    Dim parSupportBottomWidth As Double
    Dim parSupportThickness As Double
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parRearEndFlangeDia As Double
    Dim parRearEndFlangeTk1 As Double
    Dim parRearEndLength As Double
    Dim parRearEndFlangeTk2 As Double
    Dim parRearEndFlangeTk3 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)
    parExchangerDiameter = arrayOfInputs(3)
    parBundleFlangeTk = arrayOfInputs(4)
    parBundleFlangeDia = arrayOfInputs(5)
    parExchangerFlangeTk1 = arrayOfInputs(6)
    parExchangerFlangeTk2 = arrayOfInputs(7)
    parExtensionJointPosition = arrayOfInputs(8)
    parExtensionJointThickness = arrayOfInputs(9)
    parExtensionJointDiameter = arrayOfInputs(10)
    parBundlePullingLength = arrayOfInputs(11)
    parSupportAngularLocation = arrayOfInputs(12)
    parNumberOfSupport = arrayOfInputs(13)
    parSupportDepth = arrayOfInputs(14)
    parSupportFromOrigin = arrayOfInputs(15)
    parSupportTopWidth = arrayOfInputs(16)
    parSupportBottomWidth = arrayOfInputs(17)
    parSupportThickness = arrayOfInputs(18)
    parFrontEndFlangeDia = arrayOfInputs(19)
    parFrontEndFlangeTk1 = arrayOfInputs(20)
    parFrontEndLength1 = arrayOfInputs(21)
    parFrontEndLength2 = arrayOfInputs(22)
    parFrontEndFlangeTk2 = arrayOfInputs(23)
    parFrontEndFlangeTk3 = arrayOfInputs(24)
    parRearEndFlangeDia = arrayOfInputs(25)
    parRearEndFlangeTk1 = arrayOfInputs(26)
    parRearEndLength = arrayOfInputs(27)
    parRearEndFlangeTk2 = arrayOfInputs(28)
    parRearEndFlangeTk3 = arrayOfInputs(29)
    parInsulationThickness = arrayOfInputs(30)

    iOutput = 0

 ' Insert your code for output 13(Bundle Pulling Cylinder)
    If parBundlePullingLength <= 0 Then parBundlePullingLength = parExchangerLength
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
    stPoint.Set 0, 0, -parFrontEndLength1
    enPoint.Set 0, 0, parBundlePullingLength - parFrontEndLength1
    
    Set ObjBundlePulling = PlaceCylinder(m_OutputColl, stPoint, enPoint, parExchangerDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBundlePulling
    Set ObjBundlePulling = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
