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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:22 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Monday, May 26 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Horizontal Drum , is prepared as per DOW Chemical Company manual.
'   All assumed values are as per the DOW documents unless otherwise specified. For more details
'   refer DOW document D-162.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages
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
    
    Dim iOutput     As Double

    Dim parVesselDiameter     As Double
    Dim parVesselTantoTan     As Double
    Dim parInsulationThickness As Double

    'Saddle inputs
    Dim parSupportLength As Double
    Dim parSupportHeight As Double
    Dim parSupportWidth As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parBasePlateThickness As Double
    Dim parBackingPlateThickness As Double
    Dim parBackingPlateWidth As Double
    Dim parStiffnerThickness As Double
    Dim parStiffnerThickness1 As Double
    Dim parStiffnerSpacing As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    '5 for Drum
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parFirstSupportLocation = arrayOfInputs(4)
    parSecondSupportLocation = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)

    '11 Saddle inputs
    parSupportLength = arrayOfInputs(7)
    parSupportHeight = arrayOfInputs(8)
    parSupportWidth = arrayOfInputs(9)
    parBasePlateThickness = arrayOfInputs(10)
    parBackingPlateThickness = arrayOfInputs(11)
    parBackingPlateWidth = arrayOfInputs(12)
    parStiffnerThickness = arrayOfInputs(13)
    parStiffnerThickness1 = arrayOfInputs(14)
    parStiffnerSpacing = arrayOfInputs(15)
    parXboltHole = arrayOfInputs(16)
    parYboltHole = arrayOfInputs(17)

    m_oGeomHelper.OutputCollection = m_OutputColl

    Dim VesselOuterDiameter     As Double
    Dim shellThickness As Double

    shellThickness = 0.0508         'Assumed value for Shell thickness
    VesselOuterDiameter = parVesselDiameter + 2 * shellThickness

    iOutput = 0


    'Place a Cylinder of Diameter twice the Flange diameter.
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Dim InsulationDia As Double

    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition

    InsulationDia = VesselOuterDiameter + 2 * parInsulationThickness
    startPoint.Set -VesselOuterDiameter * 0.25 - parInsulationThickness, 0, 0
    endPoint.Set parVesselTantoTan + VesselOuterDiameter * 0.25 + parInsulationThickness, 0, 0

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

