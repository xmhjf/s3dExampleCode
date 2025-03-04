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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:18:35 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 30, Apr 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This symbol is vessel skirt based on "Tower T-181" symbol
' of The DOW Chemical company.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'   27.Aug.2007      MA                TR-117044 Corrected the vessel diameter dimension by ignoring skirt thickness.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub


Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim parVesselDiameter As Double
    Dim parSkirtBottomDiameter As Double
    Dim parSkirtHeight As Double
    Dim parSupportHeight As Double
    Dim parSupportDiameter As Double
    Dim parBasePlateThickness As Double
    Dim parBasePlateDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parSkirtBottomDiameter = arrayOfInputs(3)
    parSkirtHeight = arrayOfInputs(4)
    parSupportHeight = arrayOfInputs(5)
    parSupportDiameter = arrayOfInputs(6)
    parBasePlateThickness = arrayOfInputs(7)
    parBasePlateDiameter = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)

    m_oGeomHelper.OutputCollection = m_OutputColl

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    iOutput = 0
    
' Insert your code for output 5 (Support Insulation)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, parSupportHeight + parInsulationThickness

    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, _
                        parSupportDiameter + 2 * parInsulationThickness
        
' Insert your code for output 6 (Skirt Insulation)
    stPoint.Set 0, 0, parSupportHeight
    enPoint.Set 0, 0, parBasePlateThickness + parSkirtHeight + parInsulationThickness
    
'   Computation of top and bottom diameters with insulation thickness
    Dim InsDiameterTop  As Double
    Dim InsDiameterBottom  As Double
    Dim ConeApexAngle As Double
    
    ConeApexAngle = Atn(Abs(parSkirtBottomDiameter - _
                parVesselDiameter) / (2 * parSkirtHeight))
    
    InsDiameterBottom = parSkirtBottomDiameter + 2 * parInsulationThickness / Cos(ConeApexAngle)
    InsDiameterTop = InsDiameterBottom - 2 * (parSkirtHeight + parInsulationThickness) _
                            * Tan(ConeApexAngle)
                                
    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCone arrayOfOutputs(iOutput), stPoint, enPoint, _
                                InsDiameterBottom, InsDiameterTop
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
