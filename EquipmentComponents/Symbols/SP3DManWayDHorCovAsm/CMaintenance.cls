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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:09:56 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         svsmylav
'   Creation Date:  Monday, March 212003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Vessel part, Davit for Horizontal Cover, is as per DOW chemical company document. This part is
'   constructed for 20" and 24" pipes. Refer DOW document reference G9G-7005-01 for more details.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Maintenance:" 'Used for error messages

Dim PI As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    
    Dim parCentertoDavitCover As Double
    Dim parVesselDiameter As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCentertoDavitCover = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    m_oGeomHelper.OutputCollection = m_OutputColl
    iOutput = 0
    Dim davitSupWidth  As Double
    
    'Place a Cylinder of Diameter twice the Flange diameter.
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Dim envelopeDia As Double
    Dim clearance As Double
    Dim flangeGap      As Double
    Dim vertPinheight As Double
    Dim davitPinDiameter As Double
    Dim torusRadius As Double
    Dim flangeEdgeTodavitCen As Double
    Dim davitRotationAngle As Double
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    'As per DOW equipment davit pin diameter of 40mm for 150 pressure rating
    davitPinDiameter = 0.04
    
    torusRadius = 0.25
    vertPinheight = torusRadius + davitPinDiameter

'As per DOW document davit support width is sum of davit pin diameter, 10mm and 60mm
    flangeEdgeTodavitCen = davitPinDiameter + 0.01
    davitSupWidth = 0.13
    clearance = 0.1
    flangeGap = 0.004
    'Maintenance Envelope is assumed, if the Davit cover is rotated about it's axis by 360deg
    envelopeDia = (flangeDiam + flangeEdgeTodavitCen + clearance) * 2
    davitRotationAngle = PI / 2
    
    startPoint.Set Sin(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen), _
                            Cos(davitRotationAngle) * (flangeDiam / 2 + flangeEdgeTodavitCen), _
                                                                                        flangeGap
    endPoint.Set startPoint.x, startPoint.y, flangeThick + flangeGap + vertPinheight
    
    'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), startPoint, endPoint, envelopeDia
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub


