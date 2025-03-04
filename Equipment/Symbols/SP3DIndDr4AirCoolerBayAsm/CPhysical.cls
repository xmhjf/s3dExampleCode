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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:13:05 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          SSP
'   Creation Date:  Monday, Aug 18 2003
'   Description:
'   Induced Draft Air Cooler Bay (E332) - As per PDS Equipment Modelling Document.
'   Physical Aspect consists of  cooler body, fan and eight nozzles. There are four cooler bays.
'   Cooler Body is made of planes The fans are made using Place cylinder and Number of fans depends on the number of fans paramter.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.SEP.2007    PK  TR109377  Made changes so that deleted nozzles do not reappear after modification of
'                               equipment properties and also nozzles can be moved.Code for creation of nozzle
'                                place holders has been removed.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
    Dim iOutput     As Double
    Dim ObjDefaultSurface   As IngrGeom3D.Plane3d
    Dim ObjCoolerBodyPlane1 As IngrGeom3D.Plane3d
    Dim ObjCoolerBodyPlane2 As IngrGeom3D.Plane3d
    Dim ObjCoolerBodyPlane3 As IngrGeom3D.Plane3d
    Dim ObjCoolerBodyPlane4 As IngrGeom3D.Plane3d
    Dim ObjCoolerBodyPlane5 As IngrGeom3D.Plane3d
    Dim ObjCoolerLeftBox    As Object
    Dim ObjCoolerRightBox   As Object
    Dim ObjPlenumBox    As Object
    Dim ObjFan  As Object
    
    Dim parCool1toCool2CentoCen As Double 'P1
    Dim parCool2toCool3CentoCen As Double 'P2
    Dim parCool3toCool4CentoCen As Double 'P3
    Dim parCoolerWidth As Double 'P21
    Dim parCoolerLength As Double 'P22
    Dim parCoolerHeight As Double 'P23
    Dim parPlenumLength As Double 'P24
    Dim parPlenumHeight As Double 'P25
    Dim parInletBlockLength As Double 'P26
    Dim parBlockToptoCoolTop As Double 'P27
    Dim parInletBlockHeight As Double 'P28
    Dim parCoolLeftfromPP As Double 'P29
    Dim parNoOfFans As Double 'P30
    Dim parFan1CentoPP As Double 'P31
    Dim parFansCentoCen As Double 'P32
    Dim parFanDiameter As Double 'P33
    Dim parFanHeight As Double 'P34
    Dim parInletNozzLength As Double
    Dim parOutletNozzLength As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCool1toCool2CentoCen = arrayOfInputs(2)
    parCool2toCool3CentoCen = arrayOfInputs(3)
    parCool3toCool4CentoCen = arrayOfInputs(4)
    parCoolerWidth = arrayOfInputs(5)
    parCoolerLength = arrayOfInputs(6)
    parCoolerHeight = arrayOfInputs(7)
    parPlenumLength = arrayOfInputs(8)
    parPlenumHeight = arrayOfInputs(9)
    parInletBlockLength = arrayOfInputs(10)
    parBlockToptoCoolTop = arrayOfInputs(11)
    parInletBlockHeight = arrayOfInputs(12)
    parCoolLeftfromPP = arrayOfInputs(13)
    parNoOfFans = arrayOfInputs(14)
    parFan1CentoPP = arrayOfInputs(15)
    parFansCentoCen = arrayOfInputs(16)
    parFanDiameter = arrayOfInputs(17)
    parFanHeight = arrayOfInputs(18)
    parInletNozzLength = arrayOfInputs(19)
    parOutletNozzLength = arrayOfInputs(20)
    
    iOutput = 0
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim planeProj As Double
    Dim fanProj As Double
    Dim inNozzProj As Double
    Dim outNozzProj As Double
    Dim nozzIndexlong As Long
    Dim nozzIndex As Integer
    Dim iCount As Double
    Dim jCount As Double
    Dim fanXOffset As Double
    Dim oPlacePoint As New AutoMath.DPosition
    Dim oDir        As New AutoMath.DVector
    Dim objNozzle   As IJDNozzle
    
    nozzIndex = 1
    nozzIndexlong = 1
    inNozzProj = parCoolerWidth / 3
    outNozzProj = -parCoolerWidth / 3
    planeProj = -parCoolerWidth / 2
    fanProj = 0
    
For jCount = 1 To 4
    If jCount = 2 Then
        planeProj = -parCool1toCool2CentoCen - parCoolerWidth / 2
        fanProj = -parCool1toCool2CentoCen
        inNozzProj = -parCool1toCool2CentoCen + parCoolerWidth / 3
        outNozzProj = -parCool1toCool2CentoCen - parCoolerWidth / 3
    ElseIf jCount = 3 Then
        planeProj = -parCool1toCool2CentoCen - parCool2toCool3CentoCen - parCoolerWidth / 2
        fanProj = -parCool1toCool2CentoCen - parCool2toCool3CentoCen
        inNozzProj = -parCool1toCool2CentoCen - parCool2toCool3CentoCen + parCoolerWidth / 3
        outNozzProj = -parCool1toCool2CentoCen - parCool2toCool3CentoCen - parCoolerWidth / 3
    ElseIf jCount = 4 Then
        planeProj = -parCool1toCool2CentoCen - parCool2toCool3CentoCen - parCool3toCool4CentoCen - parCoolerWidth / 2
        fanProj = -parCool1toCool2CentoCen - parCool2toCool3CentoCen - parCool3toCool4CentoCen
        inNozzProj = -parCool1toCool2CentoCen - parCool2toCool3CentoCen - parCool3toCool4CentoCen + parCoolerWidth / 3
        outNozzProj = -parCool1toCool2CentoCen - parCool2toCool3CentoCen - parCool3toCool4CentoCen - parCoolerWidth / 3
    End If
    
    Dim topSurPts(3) As IJDPosition
    Dim botSurPts(3) As IJDPosition
    
    For iCount = 0 To 3
        Set topSurPts(iCount) = New DPosition
        Set botSurPts(iCount) = New DPosition
    Next iCount
  
    'Point 1
    botSurPts(0).x = parCoolLeftfromPP
    botSurPts(0).y = planeProj
    botSurPts(0).z = -parCoolerHeight / 2
    'Point 2
    botSurPts(1).x = parCoolLeftfromPP + parCoolerLength
    botSurPts(1).y = planeProj
    botSurPts(1).z = -parCoolerHeight / 2
    
    'Point3
    botSurPts(2).x = parCoolLeftfromPP + parCoolerLength
    botSurPts(2).y = planeProj + parCoolerWidth
    botSurPts(2).z = -parCoolerHeight / 2
    
    'Point 4
    botSurPts(3).x = parCoolLeftfromPP
    botSurPts(3).y = planeProj + parCoolerWidth
    botSurPts(3).z = -parCoolerHeight / 2
    
    topSurPts(0).Set botSurPts(0).x, botSurPts(0).y, botSurPts(0).z + parCoolerHeight
    topSurPts(1).Set botSurPts(1).x, botSurPts(1).y, botSurPts(1).z + parCoolerHeight
    topSurPts(2).Set botSurPts(2).x, botSurPts(2).y, botSurPts(2).z + parCoolerHeight
    topSurPts(3).Set botSurPts(3).x, botSurPts(3).y, botSurPts(3).z + parCoolerHeight
    
    Dim ObjCollection As Collection
    
    Set ObjCollection = PlaceTrapezoidWithPlanes(m_OutputColl, topSurPts, botSurPts)
   
    If jCount = 1 Then
        For iCount = 1 To 6
             iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollection(iCount)
        Next iCount
    Else
            For iCount = 1 To 6
                m_OutputColl.AddOutput "CoolerBody_", ObjCollection(iCount)
            Next iCount
    End If

    Set ObjCollection = Nothing
    For iCount = 0 To 3
        Set topSurPts(iCount) = Nothing
        Set botSurPts(iCount) = Nothing
    Next iCount
  
' Insert your code for output 7 (Left Box)
    stPoint.Set -parInletBlockLength + parCoolLeftfromPP, planeProj, -parInletBlockHeight / 2
    enPoint.Set stPoint.x + parInletBlockLength, stPoint.y + parCoolerWidth, stPoint.z + parInletBlockHeight
    
    Set ObjCoolerLeftBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    If jCount = 1 Then
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCoolerLeftBox
        Set ObjCoolerLeftBox = Nothing
    Else
        m_OutputColl.AddOutput "CoolerBody_", ObjCoolerLeftBox
        Set ObjCoolerLeftBox = Nothing
    End If
' Insert your code for output 8 (Right Box)
    stPoint.Set parCoolerLength + parCoolLeftfromPP, planeProj, -parInletBlockHeight / 2
    enPoint.Set stPoint.x + parInletBlockLength, stPoint.y + parCoolerWidth, stPoint.z + parInletBlockHeight
    
    Set ObjCoolerRightBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    If jCount = 1 Then
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCoolerRightBox
        Set ObjCoolerRightBox = Nothing
    Else
        m_OutputColl.AddOutput "CoolerBody_", ObjCoolerRightBox
        Set ObjCoolerRightBox = Nothing
    End If
' Insert your code for output 9(Plenum Box)
    stPoint.Set parCoolLeftfromPP + (parCoolerLength - parPlenumLength) / 2, planeProj, _
                                            parCoolerHeight / 2
    enPoint.Set stPoint.x + parPlenumLength, stPoint.y + parCoolerWidth, stPoint.z + parPlenumHeight
    
    Set ObjPlenumBox = PlaceBox(m_OutputColl, stPoint, enPoint)
     If jCount = 1 Then
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlenumBox
        Set ObjPlenumBox = Nothing
    Else
        m_OutputColl.AddOutput "CoolerBody_", ObjPlenumBox
        Set ObjPlenumBox = Nothing
    End If
        
' Insert your code for output 10(Fan)
    fanXOffset = 0
    For iCount = 1 To parNoOfFans
        stPoint.Set parFan1CentoPP + fanXOffset, _
                    fanProj, _
                    parInletBlockHeight / 2 + parBlockToptoCoolTop + parPlenumHeight
        enPoint.Set parFan1CentoPP + fanXOffset, _
                    fanProj, _
                    parInletBlockHeight / 2 + parBlockToptoCoolTop + parPlenumHeight + parFanHeight
        
        Set ObjFan = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFanDiameter, True)
    
    ' Set the output
        If iCount = 1 And jCount = 1 Then
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFan
            Set ObjFan = Nothing
        Else
            m_OutputColl.AddOutput "Fans_", ObjFan
            Set ObjFan = Nothing
        End If
        Set stPoint = Nothing
        Set enPoint = Nothing
        fanXOffset = fanXOffset + parFansCentoCen
    Next iCount
Next jCount

'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.02)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "ControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

