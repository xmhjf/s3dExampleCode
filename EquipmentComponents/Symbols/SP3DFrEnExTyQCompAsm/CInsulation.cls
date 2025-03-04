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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:05:40 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:        MS
'   Creation Date: Wednesday, Apr 7 2004
'   Description:
'    This is Exchanger Front End type Q. Symbol details are taken from PDS Equipment Modeling User's Guide,
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
    Dim parFrontEndDiameter As Double
    Dim parChannelDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFrontEndFlangeDia = arrayOfInputs(2) ' P30
    parFrontEndFlangeTk1 = arrayOfInputs(3) ' P31
    parFrontEndLength1 = arrayOfInputs(4) ' P32
    parFrontEndDiameter = arrayOfInputs(5) ' P33
    parChannelDiameter = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
         
    iOutput = 0
'   Origin is taken on axis of Exchanger End, at the end of P32.
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

' Insert your code for output 3(ExchFrontEndRHSFlangeIns)
   'X-coordinate: P32
    stPoint.Set parFrontEndLength1, 0, 0

    'X-coordinate: P32 - P31 - parInsulationThickness
     enPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1 - parInsulationThickness, 0, 0

    Dim ObjExchFrontEndRHSFlangeIns As Object
                
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parFrontEndFlangeDia + 2 * parInsulationThickness
    
' Insert your code for output 4(ExFrontEndBodyIns)
   'X-coordinate: P32 - P31 - parInsulationThickness
    stPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1 - parInsulationThickness, 0, 0

   'X-coordinate: - parInsulationThickness
   'A value of 0.0000001 is used in Z direction to avoid symbol placement failure(gives assertion errors).
   'This should be removed  once it is resolved.
    enPoint.Set -parInsulationThickness, 0, 0.0000001

    Dim dAngle As Double
    Dim dConeBaseRadius As Double
    Dim dConeTopRadius As Double

    'Compute cone half-apex angle
    
    dAngle = Atn((0.5 * (parChannelDiameter - parFrontEndDiameter)) / (parFrontEndLength1 - parFrontEndFlangeTk1))

    dConeBaseRadius = parChannelDiameter / 2 + _
                    parInsulationThickness / Cos(dAngle)

    'Consider insulated cone height for calculating top radius
    dConeTopRadius = dConeBaseRadius - _
                    ((parFrontEndLength1 - parFrontEndFlangeTk1) * Tan(dAngle))

' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCone arrayOfOutputs(iOutput), stPoint, enPoint, 2 * dConeBaseRadius, 2 * dConeTopRadius

    Set stPoint = Nothing
    Set enPoint = Nothing

      
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
