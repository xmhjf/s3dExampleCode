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
'   Creation Date:  Friday, August 31 2007
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
    
    Dim parBaseBotToDischarge   As Double
    Dim parHeadLength   As Double
    Dim parMotorLength  As Double
    Dim parColumnLength   As Double
    Dim parBaseThickness    As Double
    Dim parColumnDiameter   As Double
    Dim parHeadDiameter As Double
    Dim parMotorDiameter    As Double
    Dim parBasePlateWidth   As Double
    Dim parDischProjection  As Double
    Dim parPullEnvolopeLen  As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBaseBotToDischarge = arrayOfInputs(2)    'A
    parHeadLength = arrayOfInputs(3)            'B
    parMotorLength = arrayOfInputs(4)           'C
    parColumnLength = arrayOfInputs(5)          'D
    parBaseThickness = arrayOfInputs(6)         'E
    parColumnDiameter = arrayOfInputs(7)        'F
    parHeadDiameter = arrayOfInputs(8)          'G
    parMotorDiameter = arrayOfInputs(9)         'H
    parBasePlateWidth = arrayOfInputs(10)       'J
    parDischProjection = arrayOfInputs(11)      'K
    parPullEnvolopeLen = arrayOfInputs(12)      'L
    parXBoltHole = arrayOfInputs(13)
    parYBoltHole = arrayOfInputs(14)

    iOutput = 0

    'Insert your code for Control Point (Output 11)
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
