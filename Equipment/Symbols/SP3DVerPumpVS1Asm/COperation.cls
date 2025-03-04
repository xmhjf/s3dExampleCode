VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
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
'   This class module has One Output
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Operation:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

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
    
    'Origin is at PP1
    iOutput = 0
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
        
    'Create Envelope (Output 10)
    Dim objEnvelope As Object
    
    oStPoint.Set 0, 0, parHeadLength + 0.9 * parMotorLength
    oEnPoint.Set 0, 0, oStPoint.z + parPullEnvolopeLen
    Set objEnvelope = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                parMotorDiameter, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEnvelope
    Set objEnvelope = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
