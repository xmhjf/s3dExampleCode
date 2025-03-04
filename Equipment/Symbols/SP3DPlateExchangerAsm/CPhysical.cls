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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:43 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Wednesday, Apr 24 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is E325 Plate Exchanger symbol.
'    Physical Aspect consist of Exchanger Body, Two plates, One Side Plate, Two Support Rods, Two Supports
'    and Two Nozzles.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk                     CR-76070: Removed the Nozzle related parameters and code to create nozzles.
'  08.SEP.2006     KKC                      DI-95670  Replace names with initials in all revision history sheets and symbols
'   19.Sep.2006     svsmylav                CR-102784: Center of Plate 2 is modified to be same Z-coordinate as DP1 (origin)
'                                           and side plate top is aligned to the top of Plate 2.
'   11.Oct.2006     svsmylav                TR -107522: Corrected Z coordinate of defalt surface 2 to - P13/2 - P12.
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
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
    Dim iOutput     As Double
    Dim ObjPlate1 As Object
    Dim ObjExchangerBody As Object
    Dim ObjPlate2 As Object
    Dim ObjRod1 As Object
    Dim ObjRod2 As Object
    Dim ObjSidePlate As Object
    
    Dim parBotPlate1toPP As Double 'P1
    Dim parPlate1Height As Double 'P2
    Dim parPlate1Thickness As Double 'P3
    Dim parPlatesWidth As Double 'P4
    Dim parExchangerLength As Double 'P5
    Dim parExchangerHeight As Double 'P6
    Dim parExchangerWidth As Double 'P7
    Dim parPlate2Height As Double 'P8
    Dim parPlate2Thickness As Double 'P9
    Dim parRodsLength As Double 'P10
    Dim parRodsDiameter As Double 'P11
    Dim parBotSidePlatetoRod2Cen As Double 'P12
    Dim parRodsCentoCen As Double 'P13
    Dim parSidePlateThickness As Double 'P14
    Dim parSidePlateWidth As Double 'P15
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBotPlate1toPP = arrayOfInputs(2)
    parPlate1Height = arrayOfInputs(3)
    parPlate1Thickness = arrayOfInputs(4)
    parPlatesWidth = arrayOfInputs(5)
    parExchangerLength = arrayOfInputs(6)
    parExchangerHeight = arrayOfInputs(7)
    parExchangerWidth = arrayOfInputs(8)
    parPlate2Height = arrayOfInputs(9)
    parPlate2Thickness = arrayOfInputs(10)
    parRodsLength = arrayOfInputs(11)
    parRodsDiameter = arrayOfInputs(12)
    parBotSidePlatetoRod2Cen = arrayOfInputs(13)
    parRodsCentoCen = arrayOfInputs(14)
    parSidePlateThickness = arrayOfInputs(15)
    parSidePlateWidth = arrayOfInputs(16)

    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    iOutput = 0

 ' Insert your code for output 1(Plate1)
    stPoint.Set 0, parPlatesWidth / 2, parPlate1Height - parBotPlate1toPP
    enPoint.Set parPlate1Thickness, -parPlatesWidth / 2, -parBotPlate1toPP
    
    Set ObjPlate1 = PlaceBox(m_OutputColl, stPoint, enPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate1
    Set ObjPlate1 = Nothing
    
 ' Insert your code for output 2(ExchangerBody)
    stPoint.Set parPlate1Thickness, parExchangerWidth / 2, parExchangerHeight / 2
    enPoint.Set parPlate1Thickness + parExchangerLength, -parExchangerWidth / 2, -parExchangerHeight / 2
    
    Set ObjExchangerBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    
 ' Insert your code for output 3(Plate2)
    Dim ActualBody As Double 'Sum of ExchangerBody, Plate1 and Plate2 lengths
    ActualBody = parPlate1Thickness + parExchangerLength + parPlate2Thickness 'P3+P5+P9
     stPoint.Set parPlate1Thickness + parExchangerLength, _
                    parPlatesWidth / 2, _
                    parPlate2Height / 2
    enPoint.Set ActualBody, -parPlatesWidth / 2, -parPlate2Height / 2
    
    Set ObjPlate2 = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate2
    Set ObjPlate2 = Nothing
    
 ' Insert your code for output 4(Rod1)
    stPoint.Set ActualBody, 0, parRodsCentoCen / 2
    enPoint.Set ActualBody + parRodsLength, 0, parRodsCentoCen / 2
    
    Set ObjRod1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parRodsDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRod1
    Set ObjRod1 = Nothing
    
 ' Insert your code for output 5(Rod2)
    stPoint.Set ActualBody, 0, -parRodsCentoCen / 2
    enPoint.Set ActualBody + parRodsLength, 0, -parRodsCentoCen / 2
    
    Set ObjRod2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parRodsDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRod2
    Set ObjRod2 = Nothing
    
 ' Insert your code for output 6(SidePlate)
     stPoint.Set ActualBody + parRodsLength, _
                    parSidePlateWidth / 2, _
                    parPlate2Height / 2
    enPoint.Set ActualBody + parRodsLength + parSidePlateThickness, _
                -parSidePlateWidth / 2, -parRodsCentoCen / 2 - parBotSidePlatetoRod2Cen
    
    Set ObjSidePlate = PlaceBox(m_OutputColl, stPoint, enPoint)

    Set stPoint = Nothing
    Set enPoint = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSidePlate
    Set ObjSidePlate = Nothing
    
' Create  DefaultSurface1
    Dim ObjDefaultSurface1 As Object
    Dim lineStrPoints(0 To 11)  As Double
    lineStrPoints(0) = 0
    lineStrPoints(1) = parPlatesWidth / 2
    lineStrPoints(2) = -parBotPlate1toPP
        
    lineStrPoints(3) = parPlate1Thickness
    lineStrPoints(4) = parPlatesWidth / 2
    lineStrPoints(5) = -parBotPlate1toPP
            
    lineStrPoints(6) = parPlate1Thickness
    lineStrPoints(7) = -parPlatesWidth / 2
    lineStrPoints(8) = -parBotPlate1toPP
            
    lineStrPoints(9) = 0
    lineStrPoints(10) = -parPlatesWidth / 2
    lineStrPoints(11) = -parBotPlate1toPP
        
    Set ObjDefaultSurface1 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)
        
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface1
    Set ObjDefaultSurface1 = Nothing
    
' Create  DefaultSurface2
    Dim ObjDefaultSurface2 As Object
    
    lineStrPoints(0) = ActualBody + parRodsLength
    lineStrPoints(1) = parSidePlateWidth / 2
    lineStrPoints(2) = -parRodsCentoCen / 2 - parBotSidePlatetoRod2Cen
        
    lineStrPoints(3) = lineStrPoints(0) + parSidePlateThickness
    lineStrPoints(4) = parSidePlateWidth / 2
    lineStrPoints(5) = lineStrPoints(2)
            
    lineStrPoints(6) = lineStrPoints(3)
    lineStrPoints(7) = -parSidePlateWidth / 2
    lineStrPoints(8) = lineStrPoints(2)
            
    lineStrPoints(9) = lineStrPoints(0)
    lineStrPoints(10) = -parSidePlateWidth / 2
    lineStrPoints(11) = lineStrPoints(2)
        
    Set ObjDefaultSurface2 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)
        
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface2
    Set ObjDefaultSurface2 = Nothing
    Set geomFactory = Nothing
    
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0, , , True)
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "PlateExchangerControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub


