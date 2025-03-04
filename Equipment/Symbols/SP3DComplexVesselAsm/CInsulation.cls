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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:39 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          JTF
'   Creation Date:  Friday, Mar 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjSection1Support As Object
    Dim ObjSection1Dome As Object
    Dim ObjSection1Tank As Object
    Dim ObjSection2Support As Object
    Dim ObjSection2Dome As Object
    Dim ObjSection2Tank As Object
    Dim ObjSection3Support As Object
    Dim ObjSection3Tank As Object
    Dim ObjSection4Support As Object
    Dim ObjSection4Tank As Object
    Dim ObjSection4Dome As Object
    Dim ObjCVMaintenance As Object
    Dim ObjCVOperation As Object
    Dim ObjCVInsulation As Object
    Dim ObjCVInsulation2 As Object
    Dim ObjCVInsulation3 As Object
    Dim ObjCVInsulation4 As Object
    Dim ObjCVInsulation5 As Object
    Dim ObjCVInsulation6 As Object
    Dim ObjCVInsulation7 As Object
    Dim ObjCVInsulation8 As Object

    Dim parSupportHeight1 As Double
    Dim parSupportDiameter1 As Double
    Dim parDomeHeight1 As Double
    Dim parVesselHeight1 As Double
    Dim parVesselDiameter1 As Double
    Dim parSupportHeight2 As Double
    Dim parDomeHeight2 As Double
    Dim parVesselHeight2 As Double
    Dim parVesselDiameter2 As Double
    Dim parSupportHeight3 As Double
    Dim parVesselHeight3 As Double
    Dim parVesselDiameter3 As Double
    Dim parSupportHeight4 As Double
    Dim parVesselHeight4 As Double
    Dim parVesselDiameter4 As Double
    Dim parDomeHeight4 As Double

    Dim parInsulationThickness As Double
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSupportHeight1 = arrayOfInputs(2)
    parSupportDiameter1 = arrayOfInputs(3)
    parDomeHeight1 = arrayOfInputs(4)
    parVesselHeight1 = arrayOfInputs(5)
    parVesselDiameter1 = arrayOfInputs(6)
    parSupportHeight2 = arrayOfInputs(7)
    parDomeHeight2 = arrayOfInputs(8)
    parVesselHeight2 = arrayOfInputs(9)
    parVesselDiameter2 = arrayOfInputs(10)
    parSupportHeight3 = arrayOfInputs(11)
    parVesselHeight3 = arrayOfInputs(12)
    parVesselDiameter3 = arrayOfInputs(13)
    parSupportHeight4 = arrayOfInputs(14)
    parVesselHeight4 = arrayOfInputs(15)
    parVesselDiameter4 = arrayOfInputs(16)
    parDomeHeight4 = arrayOfInputs(17)
    parInsulationThickness = arrayOfInputs(18)
    
    iOutput = 0

' Insert your code for output 14(Insulation for Section1)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    
    stPoint.Set 0, 0, parSupportHeight1
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1

    Set ObjCVInsulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter1 + 2 * parInsulationThickness, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCVInsulation
    Set ObjCVInsulation = Nothing
    
' Insert your code for output 15(Insulation for Section2)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2
    
    Set ObjCVInsulation2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter2 + 2 * parInsulationThickness, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCVInsulation2
    Set ObjCVInsulation2 = Nothing
    
' Insert your code for output 16(Insulation for Section3)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3
    
    Set ObjCVInsulation3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter3 + 2 * parInsulationThickness, True)


' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCVInsulation3
    Set ObjCVInsulation3 = Nothing
 

' Insert your code for output 17(Insulation for Section4)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3 + parSupportHeight4
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3 + parSupportHeight4 + parVesselHeight4

    Set ObjCVInsulation4 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter4 + 2 * parInsulationThickness, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCVInsulation4
    Set ObjCVInsulation4 = Nothing
    
' Addition of Insulation to all other parts of the vessel except the nozzles and Skirt.
' Insert your code for output 18(Support2 Insulation)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2

    Set ObjCVInsulation5 = PlaceCone(m_OutputColl, stPoint, enPoint, parVesselDiameter1 / 2 + parInsulationThickness, parVesselDiameter2 / 2 + parInsulationThickness)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCVInsulation5
    Set ObjCVInsulation5 = Nothing
    
  ' Insert your code for output 19(Support3 Insulation)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3
    
    Set ObjCVInsulation6 = PlaceCone(m_OutputColl, stPoint, enPoint, parVesselDiameter2 / 2 + parInsulationThickness, parVesselDiameter3 / 2 + parInsulationThickness)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCVInsulation6
    Set ObjCVInsulation6 = Nothing


 ' Insert your code for output 20(Support4 Insulation)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3 + parSupportHeight4
    
    Set ObjCVInsulation7 = PlaceCone(m_OutputColl, stPoint, enPoint, parVesselDiameter3 / 2 + parInsulationThickness, parVesselDiameter4 / 2 + parInsulationThickness)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCVInsulation7
    Set ObjCVInsulation7 = Nothing
    
' Insert your code for output 21(Dome4 Insulation)
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim EllipNormX As Double
    Dim EllipNormY As Double
    Dim EllipNormZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim axis    As New AutoMath.DVector
    Dim PI As Double
    Dim dRatio As Double
    Dim oDomeEllip As IngrGeom3D.EllipticalArc3d
    Dim pDomeCenter As New AutoMath.DPosition
    Dim SP3DElem As IJDObject
       
    PI = 4 * Atn(1)
    Dim TotalHeightofAssm As Double
    TotalHeightofAssm = parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3 + parSupportHeight4 + parVesselHeight4 + parDomeHeight4

    MajorX = parVesselDiameter4 / 2 + parInsulationThickness
    MajorY = 0
    MajorZ = 0
    'The Ratio of minor axis over major axis
    dRatio = parDomeHeight4 / (parVesselDiameter4 / 2)
    'Normal vector to plane of ellipce
    EllipNormX = 0
    EllipNormY = -1
    EllipNormZ = 0
    axis.Set 0, 0, 1
    pDomeCenter.Set 0, 0, TotalHeightofAssm - parDomeHeight4
    Set oDomeEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            m_OutputColl.ResourceManager, pDomeCenter.x, _
                                            pDomeCenter.y, pDomeCenter.z, EllipNormX, _
                                            EllipNormY, EllipNormZ, MajorX, MajorY, MajorZ, _
                                            dRatio, 0, PI / 2)
    
    Set ObjCVInsulation8 = PlaceRevolution(m_OutputColl, oDomeEllip, axis, pDomeCenter, PI * 2, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCVInsulation8
    Set ObjCVInsulation8 = Nothing
    
' Delete the elliptical arc that was created for the revolution
    Set SP3DElem = oDomeEllip
    Set oDomeEllip = Nothing
    SP3DElem.Remove
    Set geomFactory = Nothing
    Set pDomeCenter = Nothing
    Set axis = Nothing
Exit Sub
   
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
