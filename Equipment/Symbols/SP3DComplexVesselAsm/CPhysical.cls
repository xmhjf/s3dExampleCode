VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:37 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          JTF
'   Creation Date:  Friday, Mar 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   20.May.2005     MS     CR-76070: Modified the nozzle creation code(as members),added datum points.
'   01.Aug.2006     Kishore(kkkottur)       Replaced 'm_OutputColl.ResourceManager' with "Nothing'.

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

    Dim CenterPos       As New AutoMath.DPosition
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    CenterPos.Set 0, 0, 0
    
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
    
      
    
    iOutput = 0
    
    
 ' useful info
 Dim TotalHeightofAssm As Double
 TotalHeightofAssm = parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3 + parSupportHeight4 + parVesselHeight4 + parDomeHeight4
 
 ' Insert your code for output 1(Support1)

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, parSupportHeight1

    Set ObjSection1Support = PlaceCone(m_OutputColl, stPoint, enPoint, parSupportDiameter1 / 2, parVesselDiameter1 / 2)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection1Support
    Set ObjSection1Support = Nothing
    
 ' Insert your code for output 2(Dome1)

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim EllipNormX As Double, EllipNormY As Double, EllipNormZ As Double
    Dim MajorX As Double, MajorY As Double, MajorZ As Double
    Dim axis    As New AutoMath.DVector
    Dim PI As Double
    Dim parRatio As Double
    Dim oDomeEllip As IngrGeom3D.EllipticalArc3d
    Dim pDomeCenter As New AutoMath.DPosition
    Dim SP3DElem As IJDObject
    Set geomFactory = New IngrGeom3D.GeometryFactory
       
    PI = 4 * Atn(1)
    
    axis.Set 0, 0, 1
    MajorX = parVesselDiameter1 / 2
    MajorY = 0
    MajorZ = 0
'The Ratio of minor axis over major axis
    parRatio = parDomeHeight1 / (parVesselDiameter1 / 2)
'Normal vector to plane of ellipce)
    EllipNormX = 0
    EllipNormY = 1
    EllipNormZ = 0

    
    
    pDomeCenter.Set 0, 0, parSupportHeight1
     
    Set oDomeEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            Nothing, pDomeCenter.x, _
                                            pDomeCenter.y, pDomeCenter.z, EllipNormX, _
                                            EllipNormY, EllipNormZ, MajorX, MajorY, MajorZ, _
                                            parRatio, 0, PI / 2)


    Set ObjSection1Dome = PlaceRevolution(m_OutputColl, oDomeEllip, axis, pDomeCenter, PI * 2, False)


' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection1Dome
    Set ObjSection1Dome = Nothing

'Delete the elliptical arc that was created for the revolution
    Set SP3DElem = oDomeEllip
    Set oDomeEllip = Nothing
    SP3DElem.Remove
    
 ' Insert your code for output 3(Tank1)
 
    
    stPoint.Set 0, 0, parSupportHeight1
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1

    Set ObjSection1Tank = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter1, False)


' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection1Tank
    Set ObjSection1Tank = Nothing
    
 ' Insert your code for output 4(Support2)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2

    Set ObjSection2Support = PlaceCone(m_OutputColl, stPoint, enPoint, parVesselDiameter1 / 2, parVesselDiameter2 / 2)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection2Support
    Set ObjSection2Support = Nothing
    
 ' Insert your code for output 5(Dome2)
 
 
    MajorX = parVesselDiameter2 / 2
    MajorY = 0
    MajorZ = 0
    'The Ratio of minor axis over major axis
    parRatio = parDomeHeight2 / (parVesselDiameter2 / 2)
    'Normal vector to plane of ellipce
    EllipNormX = 0
    EllipNormY = 1
    EllipNormZ = 0

    pDomeCenter.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2
    
    Set oDomeEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            Nothing, pDomeCenter.x, _
                                            pDomeCenter.y, pDomeCenter.z, EllipNormX, _
                                            EllipNormY, EllipNormZ, MajorX, MajorY, MajorZ, _
                                            parRatio, 0, PI / 2)

    
    Set ObjSection2Dome = PlaceRevolution(m_OutputColl, oDomeEllip, axis, pDomeCenter, PI * 2, False)
 
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection2Dome
    Set ObjSection2Dome = Nothing
    
 'Delete the elliptical arc that was created for the revolution
    Set SP3DElem = oDomeEllip
    Set oDomeEllip = Nothing
    SP3DElem.Remove
    
 ' Insert your code for output 6(Tank2)
 
    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2
    
    Set ObjSection2Tank = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter2, False)
 

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection2Tank
    Set ObjSection2Tank = Nothing
    
 ' Insert your code for output 7(Support3)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3
    
    Set ObjSection3Support = PlaceCone(m_OutputColl, stPoint, enPoint, parVesselDiameter2 / 2, parVesselDiameter3 / 2)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection3Support
    Set ObjSection3Support = Nothing
    
 ' Insert your code for output 8(Tank3)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3
    
    Set ObjSection3Tank = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter3, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection3Tank
    Set ObjSection3Tank = Nothing
    
 ' Insert your code for output 9(Support4)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3 + parSupportHeight4
    
    Set ObjSection4Support = PlaceCone(m_OutputColl, stPoint, enPoint, parVesselDiameter3 / 2, parVesselDiameter4 / 2)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection4Support
    Set ObjSection4Support = Nothing
    
 ' Insert your code for output 10(Tank4)

    stPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3 + parSupportHeight4
    enPoint.Set 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3 + parSupportHeight4 + parVesselHeight4

    Set ObjSection4Tank = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter4, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection4Tank
    Set ObjSection4Tank = Nothing
    
 ' Insert your code for output 11(Dome4)
 
 
    MajorX = parVesselDiameter4 / 2
    MajorY = 0
    MajorZ = 0
    'The Ratio of minor axis over major axis
    parRatio = parDomeHeight4 / (parVesselDiameter4 / 2)
    'Normal vector to plane of ellipce
    EllipNormX = 0
    EllipNormY = -1
    EllipNormZ = 0

    pDomeCenter.Set 0, 0, TotalHeightofAssm - parDomeHeight4
    

    Set oDomeEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            Nothing, pDomeCenter.x, _
                                            pDomeCenter.y, pDomeCenter.z, EllipNormX, _
                                            EllipNormY, EllipNormZ, MajorX, MajorY, MajorZ, _
                                            parRatio, 0, PI / 2)

    
    Set ObjSection4Dome = PlaceRevolution(m_OutputColl, oDomeEllip, axis, pDomeCenter, PI * 2, False)
 

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSection4Dome
    Set ObjSection4Dome = Nothing
    
' Delete the elliptical arc that was created for the revolution
    Set SP3DElem = oDomeEllip
    Set oDomeEllip = Nothing
    SP3DElem.Remove

'Connect Point1: places in section 4 at center of tank4
    Dim oConPoint As IngrGeom3D.Point3d
    Set oConPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, 0, (TotalHeightofAssm) - parDomeHeight4 - 0.5 * parVesselHeight4)
                                                    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oConPoint
    Set oConPoint = Nothing

'Connect Point2: places in section 2 at center of tank2
    Set oConPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, 0, parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + 0.5 * parVesselHeight2)
                                                    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oConPoint
    Set oConPoint = Nothing
    
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0, , , True)
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "ComplexVesselControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing


    
    
    
Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub


    


