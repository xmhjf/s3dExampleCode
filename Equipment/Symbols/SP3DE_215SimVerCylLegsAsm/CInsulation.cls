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
'   Author:    VRK
'   Creation Date:  FriDay,March 23 2007
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
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    Dim ObjVesselCylBodyIns As Object
    
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parVesselTantoTan As Double
    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports  As Long
    Dim parSupportHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parSupportReferencePoint  As Double
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
    parSupportAngularLocation = arrayOfInputs(6)    'P4
    parNumberOfSupports = arrayOfInputs(7)          'P5
    parSupportHeight = arrayOfInputs(8)             'P6
    parSupportLength = arrayOfInputs(9)             'P7
    parSupportThickness = arrayOfInputs(10)         'P8
    parSupportReferencePoint = arrayOfInputs(11)    'P9
    parEndHeadConeTopDiameter = arrayOfInputs(12)   'P51
    parEndHeadConeHeight = arrayOfInputs(13)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(14)     'P51 for TORC and TORS
    parEndHeadDomeradius = arrayOfInputs(15)        'P50 for Dome
    parEndHeadFlangedThick1 = arrayOfInputs(16)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(17)     'P51 for FLGD
    parEndHeadSphericalRadius = arrayOfInputs(18)   'P50
    parInsulationThickness = arrayOfInputs(19)
    
    'Insert your code for Vessel Cylindrical Body Insulation
    Dim oStartPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    
    oStartPoint.Set 0, 0, 0
    oEndPoint.Set 0, 0, parVesselTantoTan
    
    Dim dInsulationDiameter As Double
    dInsulationDiameter = parVesselDiameter + 2 * parInsulationThickness
    
    Set ObjVesselCylBodyIns = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, _
                                                                 dInsulationDiameter, False)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselCylBodyIns
    Set ObjVesselCylBodyIns = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
        
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


