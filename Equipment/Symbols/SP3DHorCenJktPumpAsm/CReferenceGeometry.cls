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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:33 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, 26 April 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
       
    Dim parDisNozzletoFace      As Double
    Dim parDisNozzletoChFace As Double
    Dim parDischargetoFaceSuction       As Double
    Dim parBaseplatetoSuction       As Double
    Dim parPumptoFaceDisNozzle      As Double
    Dim parDisNozzletoFaceBaseplate         As Double
    Dim parMotorLength      As Double
    Dim parCouplingLength       As Double
    Dim parSupportLength        As Double
    Dim parSupportHeight        As Double
    Dim parSupportWidth         As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double
    
    
' Inputs
     Set oPartFclt = arrayOfInputs(1)
    parDisNozzletoFace = arrayOfInputs(2)
    parDisNozzletoChFace = arrayOfInputs(3)
    parDischargetoFaceSuction = arrayOfInputs(4)
    parBaseplatetoSuction = arrayOfInputs(5)
    parPumptoFaceDisNozzle = arrayOfInputs(6)
    parDisNozzletoFaceBaseplate = arrayOfInputs(7)
    parMotorLength = arrayOfInputs(8)
    parCouplingLength = arrayOfInputs(9)
    parSupportLength = arrayOfInputs(10)
    parSupportHeight = arrayOfInputs(11)
    parSupportWidth = arrayOfInputs(12)
    parXboltHole = arrayOfInputs(13)
    parYboltHole = arrayOfInputs(14)
    
    iOutput = 0
    
' Insert your code for output 32( Point at origin )
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    '==========================================
    'Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, _
                                                                                                                0, 0, 0, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

