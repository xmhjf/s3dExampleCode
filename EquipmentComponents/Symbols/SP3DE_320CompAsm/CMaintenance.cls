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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:03:05 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          MS
'   Creation Date:  Friday, 27 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      This symbol is prepared based on the Double Pipe Exchanger(E320) as per the PDS Equipment Guide.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjBundlePulling As Object
    
    Dim parBlockEndFromPP As Double
    Dim parPipeEndFromPP As Double
    Dim parPipeCentoCen As Double
    Dim parPipeDiameter As Double
    Dim parBlockLength As Double
    Dim parBlockHeight As Double
    Dim parBlockWidth As Double
    Dim parBundlePullingLength As Double
    Dim parSupport1CenFromPP As Double
    Dim parSupportCentoCen As Double
    Dim parSupportHeight As Double
    Dim parSupportThickness As Double
    Dim parSupportWidth As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBlockEndFromPP = arrayOfInputs(2)
    parPipeEndFromPP = arrayOfInputs(3)
    parPipeCentoCen = arrayOfInputs(4)
    parPipeDiameter = arrayOfInputs(5)
    parBlockLength = arrayOfInputs(6)
    parBlockHeight = arrayOfInputs(7)
    parBlockWidth = arrayOfInputs(8)
    parBundlePullingLength = arrayOfInputs(9)
    parSupport1CenFromPP = arrayOfInputs(10)
    parSupportCentoCen = arrayOfInputs(11)
    parSupportHeight = arrayOfInputs(12)
    parSupportThickness = arrayOfInputs(13)
    parSupportWidth = arrayOfInputs(14)
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    iOutput = 0

''Origin is at DP/PP1

 ' Insert your code for output 6(BundlePulling)
    'P8 default value is P1+P2
    If parBundlePullingLength <= 0 Then parBundlePullingLength = parBlockEndFromPP + parPipeEndFromPP
    
    stPoint.Set parBlockEndFromPP, parBlockWidth / 2, parBlockHeight / 2
    enPoint.Set parBlockEndFromPP + parBundlePullingLength, -parBlockWidth / 2, -parBlockHeight / 2
    
    Set ObjBundlePulling = PlaceBox(m_OutputColl, stPoint, enPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBundlePulling
    Set ObjBundlePulling = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
