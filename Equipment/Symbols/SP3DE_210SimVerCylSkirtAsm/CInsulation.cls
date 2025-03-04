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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         VRK
'   Creation Date:  Wednesday,March 14 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:Insulated Vessel
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
    Dim ObjInsVesselTopDome As Object
    Dim ObjInsVesselBottomDome  As Object
    Dim objInsVesselBody As Object
    
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parVesselTantoTan As Double
    Dim parVesselDiameter As Double
    Dim parSkirtTopDiameter As Double
    Dim parSkirtBottomDiameter As Double
    Dim parSkirtHeight As Double
    Dim parSupportReferencePoint As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parInsulationThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parEnd1Type = arrayOfInputs(2)
    parEnd2Type = arrayOfInputs(3)
    parVesselTantoTan = arrayOfInputs(4)            'P2
    parVesselDiameter = arrayOfInputs(5)            'P3
    parSkirtTopDiameter = arrayOfInputs(6)          'P5
    parSkirtBottomDiameter = arrayOfInputs(7)       'P6
    parSkirtHeight = arrayOfInputs(8)               'P4
    parSupportReferencePoint = arrayOfInputs(9)     'P7
    parEndHeadConeTopDiameter = arrayOfInputs(10)   'P51
    parEndHeadConeHeight = arrayOfInputs(11)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(12)     'P51 for TORC and TORS
    parEndHeadDomeradius = arrayOfInputs(13)        'P50 for Dome
    parEndHeadFlangedThick1 = arrayOfInputs(14)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(15)     'P51 for FLGD
    parEndHeadSphericalRadius = arrayOfInputs(16)   'P50
    parInsulationThickness = arrayOfInputs(17)
    
    iOutput = 0
        
    'Origin is at DP2 of Top plane of Skirt.
    'Insert your code for output(Insulation for Vessel Body)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, parVesselTantoTan
    Set objInsVesselBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter + 2 * parInsulationThickness, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsVesselBody
    Set objInsVesselBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
   
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


