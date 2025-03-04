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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:05:33 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:          MS
'   Creation Date:  Wednesday, 25 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Forced Draft Air Cooler Bay (E334) - As per PDS Equipment Modelling Document.
'   Physical Aspect consists of  cooler body, fan .  Cooler Body is made by constructing
'   using planes The fans are made using Place cylinder and Number of fans depends on the number of fans paramter.
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
    
    Dim parCoolerWidth As Double 'P21
    Dim parCoolerLength As Double 'P22
    Dim parCoolerHeight As Double 'P23
    Dim parPlenumLength As Double 'P24
    Dim parPlenumHeight As Double 'P25
    Dim parInletBlockLength As Double 'P26
    Dim parBlockToptoCoolTop As Double 'P27
    Dim parInletBlockHeight As Double 'P28
    Dim parCoolLeftfromPP As Double 'P29
    Dim parNoOfFans As Long 'P30
    Dim parFan1CentoPP As Double 'P31
    Dim parFansCentoCen As Double 'P32
    Dim parFanDiameter As Double 'P33
    Dim parFanHeight As Double 'P34
   
   ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCoolerWidth = arrayOfInputs(2)
    parCoolerLength = arrayOfInputs(3)
    parCoolerHeight = arrayOfInputs(4)
    parPlenumLength = arrayOfInputs(5)
    parPlenumHeight = arrayOfInputs(6)
    parInletBlockLength = arrayOfInputs(7)
    parBlockToptoCoolTop = arrayOfInputs(8)
    parInletBlockHeight = arrayOfInputs(9)
    parCoolLeftfromPP = arrayOfInputs(10)
    parNoOfFans = arrayOfInputs(11)
    parFan1CentoPP = arrayOfInputs(12)
    parFansCentoCen = arrayOfInputs(13)
    parFanDiameter = arrayOfInputs(14)
    parFanHeight = arrayOfInputs(15)
     
    
    iOutput = 0
    
'Origin is at DP
    
    '==========================================
    'Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.02, , , False)
    
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
 
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub


