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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:44 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thusday, June 5 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Horizontal Drum , is prepared as per DOW Chemical Company manual.
'   All assumed values are as per the DOW documents unless otherwise specified. For more details
'   refer DOW document D-182.
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
    Dim parVesselDiameter   As Double
    Dim parVesselTantoTan   As Double
    Dim parCPtoSupport  As Double
    Dim parInsulationThickness  As Double
    
    Dim parSupportOrientation   As Double
    Dim parSupportLength    As Double
    Dim parSupportWidth As Double
    Dim parSupportHeight    As Double
    
    Dim parBackingPlateThickness    As Double
    Dim parBackingPlateDepth    As Double
    
    Dim parLengthBaseplate  As Double
    Dim parWidthBaseplate   As Double
    Dim parBasePlateThickness   As Double
    
    Dim parXboltHole    As Double
    Dim parYboltHole    As Double
    
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parCPtoSupport = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parSupportOrientation = arrayOfInputs(6)
    parSupportLength = arrayOfInputs(7)
    parSupportWidth = arrayOfInputs(8)
    parSupportHeight = arrayOfInputs(9)
    parBackingPlateThickness = arrayOfInputs(10)
    parBackingPlateDepth = arrayOfInputs(11)
    parLengthBaseplate = arrayOfInputs(12)
    parWidthBaseplate = arrayOfInputs(13)
    parBasePlateThickness = arrayOfInputs(14)
    parXboltHole = arrayOfInputs(15)
    parYboltHole = arrayOfInputs(16)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    Dim VesselOuterDiameter     As Double
    Dim shellThickness As Double
    
    shellThickness = 0.0508         'Assumed value for Shell thickness
    VesselOuterDiameter = parVesselDiameter + 2 * shellThickness
    
    iOutput = 0
                                           
    'Place a Cylinder of Diameter twice the Flange diameter.
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Dim InsulationDiameter As Double
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
                                            
    InsulationDiameter = VesselOuterDiameter + 2 * parInsulationThickness
    startPoint.Set 0, 0, -VesselOuterDiameter * 0.25 - parInsulationThickness
    endPoint.Set 0, 0, parVesselTantoTan + VesselOuterDiameter * 0.25 + parInsulationThickness
    
    'Set the Output
    iOutput = iOutput + 1

    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), startPoint, endPoint, InsulationDiameter
    
    Set startPoint = Nothing
    Set endPoint = Nothing

Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

