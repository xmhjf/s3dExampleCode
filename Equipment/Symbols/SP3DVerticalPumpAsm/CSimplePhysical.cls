VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:18:39 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          HL
'   Creation Date:  Monday, Dec 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    Dim CenterPos As IJDPosition
    Dim strOrigin As String
    
    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim oriOrientation As Orientation
    
    Dim parFaceSuctiontoFaceDischarge As Double
    Dim parPumpShafttoFaceSuction As Double
    Dim parWidthPumpShaftHousing As Double
    Dim parNozzletoBottomPumpBase As Double
    Dim parNozzletoBottomPumpShaft As Double
    Dim parBottomShafttoTopMotor As Double
    Dim parDiameterPumpMotor As Double
    Dim parHeightPumpShaftHousing As Double
    Dim parBottomPumpBasetoBottomShaft As Double
    Dim parShaftDiameter As Double
    Dim parElecBoxPosition As Double

    Dim newCenterX As Double
    Dim newCenterZ As Double
    
    Dim DC1 As Double, HC1 As Double
    Dim DC2 As Double, HC2 As Double
    Dim DC3 As Double, HS2 As Double
    Dim DC5 As Double, HC5 As Double
    Dim HD As Double
    Dim dblTop As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFaceSuctiontoFaceDischarge = arrayOfInputs(2)    'A
    parPumpShafttoFaceSuction = arrayOfInputs(3)        'B
    parWidthPumpShaftHousing = arrayOfInputs(4)         'C
    parNozzletoBottomPumpBase = arrayOfInputs(5)        'D
    parNozzletoBottomPumpShaft = arrayOfInputs(6)       'E
    parBottomShafttoTopMotor = arrayOfInputs(7)         'F
    parDiameterPumpMotor = arrayOfInputs(8)             'G
    parHeightPumpShaftHousing = arrayOfInputs(9)        'H
    parBottomPumpBasetoBottomShaft = arrayOfInputs(10)  'J
    parShaftDiameter = arrayOfInputs(11)                'K
    
    Set CenterPos = New DPosition
    CenterPos.Set 0, 0, 0
    iOutput = 0

'Insert your code for output 1(Body of Pump Box)
    'Center of box is at B, 0, 0
    'box dims are C, C, E*2
    newCenterX = parPumpShafttoFaceSuction
    
    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " U " & CenterPos.z
    Set ObjBody = CreateBox(m_outputColl, strOrigin, Nothing, parWidthPumpShaftHousing, parWidthPumpShaftHousing, parNozzletoBottomPumpShaft * 2, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Insert your code for output 2(Body of Pump Snout1)
    'Center of snout is at B, 0, D/2
    'snout dims are Top Diameter=D, Bottom Diameter=C * 0.7, Length=D, Pointing D
    newCenterZ = parNozzletoBottomPumpBase * 0.5
    Set oriOrientation = New Orientation
    oriOrientation.RotationAboutY = 90
    oriOrientation.ApplyRotations
    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " D " & CenterPos.z + newCenterZ
    Set ObjBody = CreateSnout(m_outputColl, strOrigin, oriOrientation, 0, parNozzletoBottomPumpBase, parWidthPumpShaftHousing * 0.7, parNozzletoBottomPumpBase, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Insert your code for output 3(Body of Pump Cylinder1)
    'Center of cylinder is at B, 0, H/2+E
    'Cylinder dims are Diameter=G * 0.1, Length = H
    DC1 = parDiameterPumpMotor * 0.1
    HC1 = parHeightPumpShaftHousing
    newCenterZ = parHeightPumpShaftHousing * 0.5 + parNozzletoBottomPumpShaft
    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " U " & CenterPos.z + newCenterZ
    Set ObjBody = CreateCylinder(m_outputColl, strOrigin, DC1, HC1, "N 0 U 90")

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Insert your code for output 4(Body of Pump Cylinder2)
    'Center of cylinder is at B, 0, H/2+E+HC1/4
    'Cylinder dims are Diameter=Dia of Cylinder1 * 2.5, Length = Dia of Cylinder1
    DC2 = DC1 * 2.5
    HC2 = DC1
    newCenterZ = newCenterZ + HC1 * 0.25
    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " U " & CenterPos.z + newCenterZ
    Set ObjBody = CreateCylinder(m_outputColl, strOrigin, DC2, HC2, "N 0 U 90")

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Insert your code for output 5(Body of Pump Cylinder3)
    'Center of cylinder is at B, 0, H/2+E+HC1/4+DC1*0.5
    'Cylinder dims are Diameter=Length of Cylinder2 * 1.75, Length = length of Cylinder2
    DC3 = HC2 * 1.75
    newCenterZ = newCenterZ + DC1 * 0.5
    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " U " & CenterPos.z + newCenterZ
    Set ObjBody = CreateCylinder(m_outputColl, strOrigin, DC3, HC2, "N 0 U 90")

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Insert your code for output 6(Body of Pump Cylinder4)
    'Center of cylinder is at B, 0, H/2+E+HC1/4-DC1*0.5
    'Cylinder dims are Diameter=Length of Cylinder2 * 1.75, Length = length of Cylinder2
    newCenterZ = newCenterZ - DC1
    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " U " & CenterPos.z + newCenterZ
    Set ObjBody = CreateCylinder(m_outputColl, strOrigin, DC3, HC2, "N 0 U 90")

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Insert your code for output 7(Body of Pump Snout 2)
    'Center of snout is at B, 0, H/2+E
    'snout dims are Top Diameter=C*.4, Bottom Diameter=C * 0.7, Length=H*.5, Pointing U
    HS2 = parHeightPumpShaftHousing * 0.5
    newCenterZ = parNozzletoBottomPumpShaft + HS2 / 2
    Set oriOrientation = New Orientation
    oriOrientation.RotationAboutY = -90
    oriOrientation.ApplyRotations
    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " U " & CenterPos.z + newCenterZ
    Set ObjBody = CreateSnout(m_outputColl, strOrigin, oriOrientation, 0, parWidthPumpShaftHousing * 0.4, parWidthPumpShaftHousing * 0.7, HS2, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Insert your code for output 8(Body of Pump Cylinder5 - Motor)
    'Center of cylinder is at B, 0, Length/2 + E + H
    'Cylinder dims are Diameter=G*0.97, Length = (F-H) * 0.8
    DC5 = parDiameterPumpMotor * 0.97
    HC5 = (parBottomShafttoTopMotor - parHeightPumpShaftHousing) * 0.8
    newCenterZ = parNozzletoBottomPumpShaft + parHeightPumpShaftHousing + HC5 / 2
    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " U " & CenterPos.z + newCenterZ
    Set ObjBody = CreateCylinder(m_outputColl, strOrigin, DC5, HC5, "N 0 U 90")

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Insert your code for output 9(Body of Pump Motor Dish)
    'Center of dish is at B, 0, E + H + Length of Cylinder 5
    'Dish dims are Diameter=G, Length = (F-H) * 0.2
    HD = (parBottomShafttoTopMotor - parHeightPumpShaftHousing) * 0.2
    newCenterZ = parNozzletoBottomPumpShaft + parHeightPumpShaftHousing + HC5
    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " U " & CenterPos.z + newCenterZ
    Set ObjBody = CreateDish(m_outputColl, strOrigin, "N 0 U 90", parDiameterPumpMotor, HD, 1, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

' Insert your code for output 10(Body of Pump Pyramid)
    'Center of Pyramid is at B, 0, E + H/2
    'Pyramid dims are Top x,y = C-(C-G)*0.5, Bottom x,y = C
    newCenterZ = parNozzletoBottomPumpShaft + parHeightPumpShaftHousing * 0.5
    dblTop = parWidthPumpShaftHousing - (parWidthPumpShaftHousing - parDiameterPumpMotor) / 2

    strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " U " & CenterPos.z + newCenterZ
    Set ObjBody = CreatePyramid(m_outputColl, strOrigin, Nothing, parWidthPumpShaftHousing, parWidthPumpShaftHousing, dblTop, dblTop, 0, 0, parHeightPumpShaftHousing)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'===========================
'Construction of nozzle1 Suction
'===========================
    Dim NozzlePHFactory As New NozzlePHFactory
    Dim oPipeNozzle As PipeNozzle
    Dim iNozzle As IJDNozzle
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim vecDir As IJDVector

    Set oPipeNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(oPartFclt, 1, _
                                            False, m_outputColl.ResourceManager)

    Set iLogicalDistPort = oPipeNozzle
    Set iDistribPort = oPipeNozzle

    iLogicalDistPort.SetCenterLocation CenterPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set vecDir = New DVector
    vecDir.Set 1, 0, 0
    iDistribPort.SetDirectionVector vecDir

    vecDir.Set 0, 1, 0
    iDistribPort.SetRadialOrient vecDir

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = parPumpShafttoFaceSuction
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPipeNozzle
    Set oPipeNozzle = Nothing
    Set iNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing

'========================
'Construction of nozzle2 Discharge
'========================
    Set oPipeNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(oPartFclt, 2, _
                                            False, m_outputColl.ResourceManager)
    Set NozzlePHFactory = Nothing

    Dim oPipePort As IJDPipePort
    Set oPipePort = oPipeNozzle

    CenterPos.Set CenterPos.x - parFaceSuctiontoFaceDischarge, CenterPos.y, CenterPos.z
    Set iLogicalDistPort = oPipeNozzle
    Set iDistribPort = oPipeNozzle
    iLogicalDistPort.SetCenterLocation CenterPos

    Set oPipePort = Nothing

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    vecDir.Set -1, 0, 0
    iDistribPort.SetDirectionVector vecDir

    vecDir.Set 0, 1, 0
    iDistribPort.SetRadialOrient vecDir

    iDistribPort.SetFlowDir DistribFlow_OUT

    Set iNozzle = oPipeNozzle
    iNozzle.Length = parFaceSuctiontoFaceDischarge - parPumpShafttoFaceSuction
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPipeNozzle
    Set oPipeNozzle = Nothing
    Set iNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing

    Set vecDir = Nothing

'Pump Origin set at 0, 0, 0 of symbol
'=====================================================================
'Construction of  Control Point at Pump Origin as sphere zero diameter
'=====================================================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    CenterPos.Set 0, 0, 0
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_outputColl.ResourceManager, CenterPos.x, CenterPos.y, CenterPos.z, 0, , , True)
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing

'Insert your code for output (Body of Pump Cylinder_)
    If parBottomPumpBasetoBottomShaft > 0 And parShaftDiameter > 0 Then
        'Cener of cylinder is at B, 0, D+J/2
        'Cylinder dims are Diameter=K, Length = J
        newCenterZ = parNozzletoBottomPumpBase + parBottomPumpBasetoBottomShaft / 2
        strOrigin = "W " & CenterPos.x + newCenterX & " N " & CenterPos.y & " D " & CenterPos.z + newCenterZ
        Set ObjBody = CreateCylinder(m_outputColl, strOrigin, parShaftDiameter, parBottomPumpBasetoBottomShaft, "N 0 U 90")

        ' Set the output
        m_outputColl.AddOutput "BodyCylinder_", ObjBody
        Set ObjBody = Nothing
    End If
    
    Exit Sub

ErrorLabel:
'    MsgBox (Err.Number & " " & Err.Description & " " & Err.Source)
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub

