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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:10:55 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:   Tuesday, April 222003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This symbol is based on "Davit for Vertical Cover" symbol
' of The DOW Chemical company.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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

    Dim parCentertoDavitCover As Double
    Dim parVesselDiameter As Double
    Dim parDavitHinge    As Long
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCentertoDavitCover = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parDavitHinge = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0

    Dim Gap As Double  'Gap between Flange and Manhole Cover
    Gap = 0.004

    Dim NozzleLength As Double
    NozzleLength = parCentertoDavitCover - parVesselDiameter / 2

    'Place a Cylinder of Diameter twice the Flange diameter.
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Dim InsulationDia As Double

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    InsulationDia = flangeDiam + 2 * parInsulationThickness

    Dim IntrusionDist As Double 'Intrusion Distance of Manhole Body into tower
    IntrusionDist = parVesselDiameter / 2 - Sqr(Abs((parVesselDiameter / 2) ^ 2 - (pipeDiam / 2) ^ 2)) _
                    + 0.01   '10mm intrusion in addition to the minimum needed

    stPoint.Set 0, parCentertoDavitCover + Gap + flangeThick + _
                    parInsulationThickness, 0
    enPoint.Set 0, parCentertoDavitCover - NozzleLength - IntrusionDist, 0

    'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, InsulationDia
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub

