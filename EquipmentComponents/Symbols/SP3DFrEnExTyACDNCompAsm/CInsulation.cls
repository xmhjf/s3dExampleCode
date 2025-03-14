VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:05:53 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:        MS
'   Creation Date: Wednesday, Apr 7 2004
'   Description:
'    This is Exchanger Front End type A/C/D/N. Symbol details are taken from PDS Equipment Modeling User's Guide,
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
    Private Const MODULE = "Insultion:" 'Used for error messages
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
    
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parChannelDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFrontEndFlangeDia = arrayOfInputs(2) 'P30
    parFrontEndFlangeTk1 = arrayOfInputs(3) 'P31
    parFrontEndLength1 = arrayOfInputs(4) 'P32
    parFrontEndLength2 = arrayOfInputs(5) 'P33
    parFrontEndFlangeTk2 = arrayOfInputs(6) 'P34
    parFrontEndFlangeTk3 = arrayOfInputs(7) 'P35
    parChannelDiameter = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
         
    iOutput = 0

'   Origin is taken on axis of Exchanger End, at the junction of P32 and P33 dimensions.
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

' Insert your code for output 5(ExchFrontEndRHSFlangeIns)
    'X-coordinate: P32
    stPoint.Set parFrontEndLength1, 0, 0

    'X-coordinate: P32 - P31 - parInsulationThickness
     enPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1 - parInsulationThickness, 0, 0
                
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parFrontEndFlangeDia + 2 * parInsulationThickness

' Insert your code for output 6(ExFrontEndBodyIns)
    'X-coordinate: P32 - P31 - parInsulationThickness
    stPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1 - parInsulationThickness, 0, 0

    'X-coordinate: -P33 + P35 + parInsulationThickness
    enPoint.Set -parFrontEndLength2 + parFrontEndFlangeTk3 + parInsulationThickness, 0, 0

' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parChannelDiameter + 2 * parInsulationThickness
    
' Insert your code for output7(ExFrontEndBodyIFlangeIns)
    'X-coordinate: -P33 + P35 + parInsulationThickness
    stPoint.Set -parFrontEndLength2 + parFrontEndFlangeTk3 + parInsulationThickness, 0, 0

    'X-coordinate: -P33
    enPoint.Set -parFrontEndLength2, 0, 0

 ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parFrontEndFlangeDia + 2 * parInsulationThickness

' Insert your code for output 8(ExchFrontEndLHSFlangeIns)
    'X-coordinate: -P33
    stPoint.Set -parFrontEndLength2, 0, 0

    'X-coordinate: -P33 -P34 - parInsulationThickness
    enPoint.Set -parFrontEndLength2 - parFrontEndFlangeTk2 - parInsulationThickness, 0, 0

 ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parFrontEndFlangeDia + 2 * parInsulationThickness
                             
    Set stPoint = Nothing
    Set enPoint = Nothing
      
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
