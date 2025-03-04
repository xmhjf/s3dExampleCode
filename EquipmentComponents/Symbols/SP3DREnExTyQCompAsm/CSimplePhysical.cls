VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:31 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:        MS
'   Creation Date: Wednesday, Apr 7 2004
'   Description:
'    This is Exchanger Rear End type Q. Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E319 Symbol in Page no 310.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

    Private m_oSymGeomHelper As IJSymbolGeometryHelper
    Private Const MODULE = "SimplePhysical:" 'Used for error messages
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
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim parFrontEndLength1 As Double
    Dim parExchangerLength As Double
    Dim parRearEndFlangeDia As Double
    Dim parRearEndFlangeTk1 As Double
    Dim parRearEndLength As Double
    Dim parRearEndDiameter As Double
    Dim parChannelDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFrontEndLength1 = arrayOfInputs(2)
    parExchangerLength = arrayOfInputs(3)
    parRearEndFlangeDia = arrayOfInputs(4) ' P40
    parRearEndFlangeTk1 = arrayOfInputs(5) ' P41
    parRearEndLength = arrayOfInputs(6) ' P42
    parRearEndDiameter = arrayOfInputs(7) ' P43
    parChannelDiameter = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
     
    
    iOutput = 0
'   Origin is taken on axis of Exchanger at (parFrontEndLength1 + parExchangerLength) from the left face of the Rear end

    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

' Insert your code for output 1(ExchRearEndRHSFlange)
    stPoint.Set (parFrontEndLength1 + parExchangerLength), 0, 0
    enPoint.Set (parFrontEndLength1 + parExchangerLength) + parRearEndFlangeTk1, 0, 0

' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parRearEndFlangeDia

' Insert your code for output 2(ExRearEndBody)
    stPoint.Set (parFrontEndLength1 + parExchangerLength) + parRearEndFlangeTk1, 0, 0
    'A value of 0.0000001 is used in Z direction to avoid symbol placement failure(gives assertion errors).
    'This should be removed  once it is resolved.
    enPoint.Set (parFrontEndLength1 + parExchangerLength) + parRearEndLength, 0, 0.0000001
                     
' Create Cone using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCone arrayOfOutputs(iOutput), stPoint, enPoint, parChannelDiameter, parRearEndDiameter
    Set stPoint = Nothing
    Set enPoint = Nothing
      
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

