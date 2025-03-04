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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:18:55 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         RUK
'   Creation Date:  Tuesday, April 10 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'           Cylinder at the Left most position
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   05.10.2007      RUK     CP 128584 Added new attributes JunBoxHeight, JunBoxWidth, JunBoxLength
'                               and used to create the Junction Box of Motor.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Maintenance:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parMotorHeight As Double
    Dim parCenterToNozzFace As Double
    Dim parDischargeToSuction As Double
    Dim parImpShaftFrontEnd As Double
    Dim parImpShaftRearEnd As Double
    Dim parDistBetShaftEnds As Double
    Dim parMotorShaftToEnd As Double
    Dim parMotorWidth As Double
    Dim parDismantlingLength As Double
    Dim parBasePlateBottom As Double
    Dim parCenToFouPortStart As Double
    Dim parBasePlateWidth As Double
    Dim parBasePlateLength As Double
    Dim parBasePlateHeight As Double
    Dim parJunBoxHeight As Double
    Dim parJunBoxWidth As Double
    Dim parJunBoxLength As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parMotorHeight = arrayOfInputs(2)               'D
    parCenterToNozzFace = arrayOfInputs(3)          'NH
    parDischargeToSuction = arrayOfInputs(4)        'X
    parImpShaftFrontEnd = arrayOfInputs(5)          'Y
    parImpShaftRearEnd = arrayOfInputs(6)           'W
    parDistBetShaftEnds = arrayOfInputs(7)          'DBSE
    parMotorShaftToEnd = arrayOfInputs(8)           'CMAX
    parMotorWidth = arrayOfInputs(9)                'PMAX
    parDismantlingLength = arrayOfInputs(10)        'EX
    parBasePlateBottom = arrayOfInputs(11)          'HD
    parCenToFouPortStart = arrayOfInputs(12)        'HR
    parBasePlateWidth = arrayOfInputs(13)           'HA
    parBasePlateLength = arrayOfInputs(14)          'HB
    parBasePlateHeight = arrayOfInputs(15)          'BD
    parJunBoxHeight = arrayOfInputs(16)             'JH
    parJunBoxWidth = arrayOfInputs(17)              'JW
    parJunBoxLength = arrayOfInputs(18)             'JL
    parXBoltHole = arrayOfInputs(19)                'BX
    parYBoltHole = arrayOfInputs(20)                'HF
    
    'Origin is at DP1/PP1
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
        
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    oStPoint.Set 0, 0, 0
    oEnPoint.Set -parDismantlingLength, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.5 * 0.7 * parMotorWidth
    
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

