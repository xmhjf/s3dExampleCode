VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRefGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Monday, September 03 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'           Control Point is created at the origin
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "RefGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parBaseBotToDischarge As Double
    Dim parBaseBotToSuction As Double
    Dim parHeadLength As Double
    Dim parMotorLength As Double
    Dim parBarrelLength As Double
    Dim parBaseThickness As Double
    Dim parBarrelDiameter As Double
    Dim parHeadDiameter As Double
    Dim parMotorDiameter As Double
    Dim parBasePlateWidth As Double
    Dim parDischProjection As Double
    Dim parSuctionProjection As Double
    Dim parPullEnvolopeLen As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBaseBotToDischarge = arrayOfInputs(2)    'A
    parBaseBotToSuction = arrayOfInputs(3)      'B
    parHeadLength = arrayOfInputs(4)            'C
    parMotorLength = arrayOfInputs(5)           'D
    parBarrelLength = arrayOfInputs(6)          'E
    parBaseThickness = arrayOfInputs(7)         'F
    parBarrelDiameter = arrayOfInputs(8)        'G
    parHeadDiameter = arrayOfInputs(9)          'H
    parMotorDiameter = arrayOfInputs(10)        'J
    parBasePlateWidth = arrayOfInputs(11)       'K
    parDischProjection = arrayOfInputs(12)      'L
    parSuctionProjection = arrayOfInputs(13)    'M
    parPullEnvolopeLen = arrayOfInputs(14)      'P
    parXBoltHole = arrayOfInputs(15)
    parYBoltHole = arrayOfInputs(16)

    iOutput = 0

    'Insert your code for output (Point at origin)
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
