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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:08:36 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, May 13 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Davit for Bottom Horizontal Cover, is prepared as per DOW Chemical Company manual.
'   Dimensions are considered for two pipe size viz., 20" and 24" . All assumed values are as per the DOW
'   documents unless otherwise specified.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
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

    Dim iOutput     As Double
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim parInsulationThickness As Double
    Dim parCentertoDavitCover As Double
    Dim parDomeHeight1 As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCentertoDavitCover = arrayOfInputs(2)
    parDomeHeight1 = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim flangeGap As Double
    
    'Place a Cylinder of Diameter twice the Flange diameter.
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Dim InsulationDia As Double
    
    flangeGap = 0.004           'Gap between flanges is assumed as 4mm
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
                                            
    InsulationDia = flangeDiam + 2 * parInsulationThickness
    startPoint.Set 0, 0, -flangeGap - flangeThick - parInsulationThickness
    endPoint.Set 0, 0, parCentertoDavitCover - parDomeHeight1 + 0.2 * parDomeHeight1
    
    'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), startPoint, endPoint, InsulationDia
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub



