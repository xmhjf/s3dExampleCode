VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReferenceGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:04 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:          MS
'   Creation Date: Tuesday, Feb 24 2004
'   Description:
'    The symbol is prepared from PDS Equipment guide E405. The symbol consists of 9 physical output and 1 Insulation Output.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "ReferenceGeometry:" 'Used for error messages


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
    Dim iOutput     As Double
    
    Dim parBaseLength As Double 'P1
    Dim parBaseWidth1 As Double 'P2
    Dim parBaseWidth2 As Double 'P3
    Dim parBaseThickness As Double 'P4
    Dim parDistBaseLeftToDischarge As Double 'P5
    Dim parDistEqpLeftToDischarge As Double 'P6
    Dim parDistEqpRightToDischarge As Double 'P7
    Dim parEqpWidth1 As Double 'P8
    Dim parEqpWidth2 As Double 'P9
    Dim parEqpDepth1 As Double 'P10
    Dim parEqpDepth2 As Double 'P11
    Dim parRotEqpLength As Double 'P12
    Dim parRotEqpDiameter As Double 'P13
    Dim parDriverLength As Double 'P14
    Dim parDriverWidth1 As Double 'P15
    Dim parDriverWidth2 As Double 'P16
    Dim parDriverHeight As Double 'P17
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBaseLength = arrayOfInputs(2)
    parBaseWidth1 = arrayOfInputs(3)
    parBaseWidth2 = arrayOfInputs(4)
    parBaseThickness = arrayOfInputs(5)
    parDistBaseLeftToDischarge = arrayOfInputs(6)
    parDistEqpLeftToDischarge = arrayOfInputs(7)
    parDistEqpRightToDischarge = arrayOfInputs(8)
    parEqpWidth1 = arrayOfInputs(9)
    parEqpWidth2 = arrayOfInputs(10)
    parEqpDepth1 = arrayOfInputs(11)
    parEqpDepth2 = arrayOfInputs(12)
    parRotEqpLength = arrayOfInputs(13)
    parRotEqpDiameter = arrayOfInputs(14)
    parDriverLength = arrayOfInputs(15)
    parDriverWidth1 = arrayOfInputs(16)
    parDriverWidth2 = arrayOfInputs(17)
    parDriverHeight = arrayOfInputs(18)
    parInsulationThickness = arrayOfInputs(19)
        
    iOutput = 0
    
''Origin is at DP/PP 1(Inside the Equipment Body)
    
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
  
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

    
    Exit Sub
        
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub


