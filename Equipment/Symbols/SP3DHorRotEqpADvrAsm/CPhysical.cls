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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:12 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Monday, Apr 22 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is E405 Horizontal Rotating Equipment and Driver symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   23.Jul.2003     svsmylav/ACM            TR 44842:  1.Removed Default surface
'                                                      2. Made pump base which was a box to 6 planes to make bottom as DefaultSurface
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create nozzles.
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
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
    Dim iOutput     As Double
    Dim ObjEquipment As Object
    Dim ObjDriver As Object
    Dim ObjRotEquip As Object
    Dim ObjBase As Object
    
    Dim parBaseLength As Double 'P1
    Dim parBaseWidth1 As Double 'P2
    Dim parBaseWidth2 As Double 'P3
    Dim parBaseThickness As Double 'P4
    Dim parDistBaseLeftToDischarge As Double 'P5
    Dim parDistEqpLeftToDischarge As Double 'P6
    Dim parDistEqpRightToDischarge As Double 'P7
    Dim parEqpWidth1 As Double 'P8
    Dim parEqpWidth2 As Double 'P9
    Dim parEqpDepth1 As Double 'P10
    Dim parEqpDepth2 As Double 'P11
    Dim parRotEqpLength As Double 'P12
    Dim parRotEqpDiameter As Double 'P13
    Dim parDriverLength As Double 'P14
    Dim parDriverWidth1 As Double 'P15
    Dim parDriverWidth2 As Double 'P16
    Dim parDriverHeight As Double 'P17
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBaseLength = arrayOfInputs(2)
    parBaseWidth1 = arrayOfInputs(3)
    parBaseWidth2 = arrayOfInputs(4)
    parBaseThickness = arrayOfInputs(5)
    parDistBaseLeftToDischarge = arrayOfInputs(6)
    parDistEqpLeftToDischarge = arrayOfInputs(7)
    parDistEqpRightToDischarge = arrayOfInputs(8)
    parEqpWidth1 = arrayOfInputs(9)
    parEqpWidth2 = arrayOfInputs(10)
    parEqpDepth1 = arrayOfInputs(11)
    parEqpDepth2 = arrayOfInputs(12)
    parRotEqpLength = arrayOfInputs(13)
    parRotEqpDiameter = arrayOfInputs(14)
    parDriverLength = arrayOfInputs(15)
    parDriverWidth1 = arrayOfInputs(16)
    parDriverWidth2 = arrayOfInputs(17)
    parDriverHeight = arrayOfInputs(18)
    parInsulationThickness = arrayOfInputs(19)
        
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    iOutput = 0

 ' Insert your code for output 4(Equipment)
    stPoint.Set parDistBaseLeftToDischarge - parDistEqpLeftToDischarge, parEqpWidth2, parEqpDepth1
    enPoint.Set parDistBaseLeftToDischarge + parDistEqpRightToDischarge, -parEqpWidth1, -parEqpDepth2
    
    Set ObjEquipment = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEquipment
    Set ObjEquipment = Nothing
    
' Insert your code for output 5(Driver)
    stPoint.Set parDistBaseLeftToDischarge + _
                    parDistEqpRightToDischarge + _
                    parRotEqpLength, parDriverWidth2, parDriverHeight / 2
    enPoint.Set stPoint.x + parDriverLength, -parDriverWidth1, -parDriverHeight / 2
    
    Set ObjDriver = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDriver
    Set ObjDriver = Nothing
    
' Insert your code for output 6(Rotating Equipment)
    stPoint.Set parDistBaseLeftToDischarge + parDistEqpRightToDischarge, 0, 0
    enPoint.Set stPoint.x + parRotEqpLength, 0, 0
    
    Set ObjRotEquip = PlaceCylinder(m_OutputColl, stPoint, enPoint, parRotEqpDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotEquip
    Set ObjRotEquip = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' ' Insert your code for output 7 to 12(Base plate as planes)
'    stPoint.Set 0, parBaseWidth2, -parEqpDepth2
'    enPoint.Set parBaseLength, -parBaseWidth1, -parEqpDepth2 - parBaseThickness
'
'    Set ObjBase = PlaceBox(m_OutputColl, stPoint, enPoint)
    Dim objColl As New Collection
    Dim dBaseCenter As New AutoMath.DPosition

    dBaseCenter.Set parBaseLength / 2, 0, -parEqpDepth2 - parBaseThickness
    Dim oTopPts(0 To 3) As IJDPosition
    Dim oBotPts(0 To 3) As IJDPosition

    Dim iCount As Integer
    For iCount = 0 To 3
        Set oTopPts(iCount) = New DPosition
        Set oBotPts(iCount) = New DPosition
    Next iCount
    
    oTopPts(0).Set dBaseCenter.x - parBaseLength / 2, dBaseCenter.y - parBaseWidth1, dBaseCenter.z + parBaseThickness
    oTopPts(1).Set oTopPts(0).x + parBaseLength, oTopPts(0).y, oTopPts(0).z
    oTopPts(2).Set oTopPts(1).x, oTopPts(1).y + parBaseWidth1 + parBaseWidth2, oTopPts(0).z
    oTopPts(3).Set oTopPts(2).x - parBaseLength, oTopPts(2).y, oTopPts(0).z
        
    oBotPts(0).Set dBaseCenter.x - parBaseLength / 2, dBaseCenter.y - parBaseWidth1, dBaseCenter.z
    oBotPts(1).Set oBotPts(0).x + parBaseLength, oBotPts(0).y, oBotPts(0).z
    oBotPts(2).Set oBotPts(1).x, oBotPts(1).y + parBaseWidth1 + parBaseWidth2, oBotPts(0).z
    oBotPts(3).Set oBotPts(2).x - parBaseLength, oBotPts(2).y, oBotPts(0).z
       
    Set objColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopPts, oBotPts)
    
' Set the output
    For iCount = 1 To objColl.Count
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objColl(iCount)
    Next iCount
    Set objColl = Nothing
    
    For iCount = 0 To 3
        Set oTopPts(iCount) = Nothing
        Set oBotPts(iCount) = Nothing
    Next iCount
   
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0, , , True)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "HorRotEqpADvrControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

    
    Exit Sub
    
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
