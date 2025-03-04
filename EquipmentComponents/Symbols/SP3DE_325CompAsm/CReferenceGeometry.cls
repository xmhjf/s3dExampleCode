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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:49 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:          MS
'   Creation Date: Tuesday, Feb 17 2004
'   Description:
'    The symbol is prepared from PDS Equipment guide E325. The symbol consists of 11 physical outputs.
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
    
    Dim parBotPlate1toPP As Double 'P1
    Dim parPlate1Height As Double 'P2
    Dim parPlate1Thickness As Double 'P3
    Dim parPlatesWidth As Double 'P4
    Dim parExchangerLength As Double 'P5
    Dim parExchangerHeight As Double 'P6
    Dim parExchangerWidth As Double 'P7
    Dim parPlate2Height As Double 'P8
    Dim parPlate2Thickness As Double 'P9
    Dim parRodsLength As Double 'P10
    Dim parRodsDiameter As Double 'P11
    Dim parBotSidePlatetoRod2Cen As Double 'P12
    Dim parRodsCentoCen As Double 'P13
    Dim parSidePlateThickness As Double 'P14
    Dim parSidePlateWidth As Double 'P15

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBotPlate1toPP = arrayOfInputs(2)
    parPlate1Height = arrayOfInputs(3)
    parPlate1Thickness = arrayOfInputs(4)
    parPlatesWidth = arrayOfInputs(5)
    parExchangerLength = arrayOfInputs(6)
    parExchangerHeight = arrayOfInputs(7)
    parExchangerWidth = arrayOfInputs(8)
    parPlate2Height = arrayOfInputs(9)
    parPlate2Thickness = arrayOfInputs(10)
    parRodsLength = arrayOfInputs(11)
    parRodsDiameter = arrayOfInputs(12)
    parBotSidePlatetoRod2Cen = arrayOfInputs(13)
    parRodsCentoCen = arrayOfInputs(14)
    parSidePlateThickness = arrayOfInputs(15)
    parSidePlateWidth = arrayOfInputs(16)
    
    iOutput = 0
    
'Origin is at DP/PP1
    
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




