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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:40 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:          BG
'   Creation Date:  Thursday, Feb 6 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using following five Insulation aspect outputs:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim oGBSFactory As IJGeneralBusinessObjectsFactory
    Dim oControlPoint As IJControlPoint
    Dim iOutput As Double
    
    Dim parVesselLength As Double
    Dim parVesselDiameter As Double
    Dim parSupportLength As Double
    Dim parSupportHeight As Double
    Dim parSupportThickness As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parSupportLength = arrayOfInputs(4)
    parSupportHeight = arrayOfInputs(5)
    parSupportThickness = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
    '==========================================
    'Construction of Pump Control Point
    '==========================================
    Set oGBSFactory = New GeneralBusinessObjectsFactory
      
    Set oControlPoint = oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.02, _
                                                        , , False)
    
    oControlPoint.SubType = cpProcessEquipment
    oControlPoint.Type = cpControlPoint

    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oControlPoint
    Set oControlPoint = Nothing
    Set oGBSFactory = Nothing
       
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


