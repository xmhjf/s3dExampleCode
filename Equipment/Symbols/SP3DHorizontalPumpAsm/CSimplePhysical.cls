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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:52 AM"
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
'   25.Jul.2003     SSP                              Changing Plane Projection for Base construction to six planes
'                                                                       and  in place of Anchor bolt holes,adding Foundation Port with holes.
'                                                                     (TR-44842  Update Equipment symbols, so as to position the equipment by constraints  )
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages

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
    Dim strDir As String
    
    Dim iOutput     As Double
    Dim ObjBodyCyl As Object
    Dim ObjBodyBaseplate As Object
    Dim oriOrientation As Orientation
    
    Dim parDisNozzletoFace As Double
    Dim parDischargetoFaceSuction As Double
    Dim parBaseplatetoSuction As Double
    Dim parOffsetDisNozzletoPump As Double
    Dim parPumptoFaceDisNozzle As Double
    Dim parDisNozzletoFaceBaseplate As Double
    Dim parWidthBaseplate As Double
    Dim parLengthBaseplate As Double
    Dim parHeightBaseplate As Double
    Dim parFaceBaseplatetoAnchorBolt As Double
    Dim parShafttoAnchorBolt As Double
    Dim parFronttoBackAnchorBolt As Double
    Dim parBaseplatetoMotor As Double

    Dim newCenter As Double
    Dim HC1 As Double, HS1 As Double, HD1 As Double, HC4 As Double
    Dim DC1 As Double, DSB As Double, DST As Double, DC4 As Double
    
    Dim HP1 As Double
    Dim XBP1 As Double, XBP2 As Double
    Dim YTP1 As Double, YTP2 As Double
    Dim SB1 As Double, SBZ As Double
'
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parDisNozzletoFace = arrayOfInputs(2)           'A      - DC
    parDischargetoFaceSuction = arrayOfInputs(3)    'B      - SP
    parBaseplatetoSuction = arrayOfInputs(4)        'D      - D1
    parOffsetDisNozzletoPump = arrayOfInputs(5)     'F      - D0
    parPumptoFaceDisNozzle = arrayOfInputs(6)       'X      - DP
    parDisNozzletoFaceBaseplate = arrayOfInputs(7)  'Z      - BF
    parWidthBaseplate = arrayOfInputs(8)            'HA     - BW
    parLengthBaseplate = arrayOfInputs(9)           'HB     - BL
    parHeightBaseplate = arrayOfInputs(10)          'HG     - BT
    parFaceBaseplatetoAnchorBolt = arrayOfInputs(11) 'C     - DIMC
    parShafttoAnchorBolt = arrayOfInputs(12)        'HE     - HE
    parFronttoBackAnchorBolt = arrayOfInputs(13)    'HF     - HF
    parBaseplatetoMotor = arrayOfInputs(14)         'E      - DIME

    If parDisNozzletoFace <= 0 Then
'        MsgBox "can not have distance less or equal to 0"
        Exit Sub
    End If
    
    Set CenterPos = New DPosition
    CenterPos.Set 0, 0, 0
    iOutput = 0

    'Calculate primitive location from origin of pump
    HC1 = 1.5 * parDischargetoFaceSuction
    DC1 = 1.8 * (parBaseplatetoSuction - parHeightBaseplate)
    
'Insert your code for output 1(Body of Pump Cylinder1)
    newCenter = parDischargetoFaceSuction
'    MsgBox "Y Position is " & newCenter
'    MsgBox "Diameter is " & DC1
'    MsgBox "Length is " & HC1
    strOrigin = "E " & CenterPos.x & " N " & CenterPos.y + newCenter & " U " & CenterPos.z
    Set ObjBodyCyl = CreateCylinder(m_outputColl, strOrigin, DC1, HC1, "N 0 U 0")
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
    Set ObjBodyCyl = Nothing

'Insert your code for output 2(Body of Pump Snout)
    HS1 = parDisNozzletoFace - HC1 / 2 - 0.0508
    newCenter = newCenter + HC1 / 2 + HS1 / 2

    DSB = 0.7 * DC1
    DST = 0.2 * DC1
'    MsgBox "Y Position is " & newCenter
'    MsgBox "Top Diameter is " & DST
'    MsgBox "Bottom Diameter is " & DSB
'    MsgBox "Length is " & HS1

    Set oriOrientation = New Orientation
    oriOrientation.RotationAboutZ = 90
    oriOrientation.ApplyRotations
    strOrigin = "E " & CenterPos.x & " N " & CenterPos.y + newCenter & " U " & CenterPos.z
    Set ObjBodyCyl = CreateSnout(m_outputColl, strOrigin, oriOrientation, 0, DST, DSB, HS1, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
    Set ObjBodyCyl = Nothing

'Insert your code for output 3(Body of Pump Cylinder2 - Pump to Motor shaft)
    newCenter = newCenter + HS1 / 2 + 0.1016
    strOrigin = "E " & CenterPos.x & " N " & CenterPos.y + newCenter & " U " & CenterPos.z
    Set ObjBodyCyl = CreateCylinder(m_outputColl, strOrigin, 0.0254, 0.2032, "N 0 U 0")

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
    Set ObjBodyCyl = Nothing

'Insert your code for output 4(Body of Pump Cylinder3 - Coupling)
    Set ObjBodyCyl = CreateCylinder(m_outputColl, strOrigin, 0.0762, 0.1016, "N 0 U 0")

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
    Set ObjBodyCyl = Nothing

'Insert your code for output 5(Body of Pump Motor Dish 1)
    DC4 = 0.75 * DC1
    HD1 = DC4 / 4
    newCenter = newCenter + 0.1016 + HD1
'    MsgBox "Y position is " & newCenter
'    MsgBox "Height is " & HD1
'    MsgBox "Length is " & DC4
    strOrigin = "E " & CenterPos.x & " N " & CenterPos.y + newCenter & " U " & CenterPos.z
    strDir = "S 180 U 0"
    Set ObjBodyCyl = CreateDish(m_outputColl, strOrigin, strDir, DC4, HD1, 1, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
    Set ObjBodyCyl = Nothing

'Insert your code for output 6(Body of Pump Cylinder4 - Motor)
    HC4 = parLengthBaseplate - parDisNozzletoFaceBaseplate - parDisNozzletoFace - 0.1524 - HD1 - HD1
    newCenter = newCenter + HC4 / 2

    If parBaseplatetoMotor > 0 Then
        newCenter = newCenter + parBaseplatetoMotor / 2
        HC4 = HC4 + parBaseplatetoMotor
    End If
'    MsgBox "Y position is " & newCenter
'    MsgBox "Diameter is " & DC4
'    MsgBox "Height is " & HC4

    strOrigin = "E " & CenterPos.x & " N " & CenterPos.y + newCenter & " U " & CenterPos.z
    Set ObjBodyCyl = CreateCylinder(m_outputColl, strOrigin, DC4, HC4, "N 0 U 0")

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
    Set ObjBodyCyl = Nothing

'Insert your code for output 7(Body of Pump Motor Dish 2)
    newCenter = newCenter + HC4 / 2
'    MsgBox "Y position is " & newCenter
'    MsgBox "Height is " & HD1
'    MsgBox "Length is " & DC4

    strOrigin = "E " & CenterPos.x & " N " & CenterPos.y + newCenter & " U " & CenterPos.z
    strDir = "N 0 U 0"
    Set ObjBodyCyl = CreateDish(m_outputColl, strOrigin, strDir, DC4, HD1, 1, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
    Set ObjBodyCyl = Nothing

' Insert your code for output 8(Body of Pump Motor Pyramid)
    newCenter = newCenter - HC4 / 2
    If parBaseplatetoMotor > 0 Then
        newCenter = newCenter - parBaseplatetoMotor / 2
        HC4 = HC4 - parBaseplatetoMotor
    End If

    HP1 = parBaseplatetoSuction - parHeightBaseplate
    XBP1 = 0.8 * DC1
    XBP2 = 0.8 * DC4
    YTP1 = 0.5 * HC1
    YTP2 = 0.8 * HC4

    strOrigin = "E " & CenterPos.x & " N " & CenterPos.y + newCenter & " D " & CenterPos.z + HP1 / 2
    Set ObjBodyCyl = CreatePyramid(m_outputColl, strOrigin, Nothing, XBP2, YTP2, 0, YTP2, 0, 0, HP1)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
    Set ObjBodyCyl = Nothing

'  Insert your code for output 9(Body of Pump Pyramid)
    strOrigin = "E " & CenterPos.x & " N " & CenterPos.y + parDischargetoFaceSuction & " D " & CenterPos.z + HP1 / 2
    Set ObjBodyCyl = CreatePyramid(m_outputColl, strOrigin, Nothing, XBP1, YTP1, 0, YTP1, 0, 0, HP1)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
    Set ObjBodyCyl = Nothing

'  Insert your code for output 10(Body of Pump Baseplate)
    
'    newCenter = parLengthBaseplate / 2 - parDisNozzletoFaceBaseplate + parDischargetoFaceSuction
'    SBZ = parBaseplatetoSuction - parHeightBaseplate / 2
'    strOrigin = "E " & CenterPos.x & " N " & CenterPos.y + newCenter & " D " & CenterPos.z + SBZ
'    Set ObjBodyCyl = CreateBox(m_outputColl, strOrigin, Nothing, parWidthBaseplate, parLengthBaseplate, parHeightBaseplate, True)
'
'    ' Set the output
'    iOutput = iOutput + 1
'    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCyl
'    Set ObjBodyCyl = Nothing

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
    vecDir.Set 0, -1, 0
    iDistribPort.SetDirectionVector vecDir

    vecDir.Set 1, 0, 0
    iDistribPort.SetRadialOrient vecDir

    iDistribPort.SetFlowDir DistribFlow_IN

    Set iNozzle = oPipeNozzle
    iNozzle.Length = parDischargetoFaceSuction
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

    CenterPos.Set CenterPos.x + parOffsetDisNozzletoPump, CenterPos.y + parDischargetoFaceSuction, CenterPos.z + parPumptoFaceDisNozzle
    Set iLogicalDistPort = oPipeNozzle
    Set iDistribPort = oPipeNozzle
    iLogicalDistPort.SetCenterLocation CenterPos

    Set oPipePort = Nothing

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    vecDir.Set 0, 0, 1
    iDistribPort.SetDirectionVector vecDir

    vecDir.Set 0, 1, 0
    iDistribPort.SetRadialOrient vecDir

    iDistribPort.SetFlowDir DistribFlow_OUT

    Set iNozzle = oPipeNozzle
    iNozzle.Length = parPumptoFaceDisNozzle
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

    'Place Base
    newCenter = parLengthBaseplate / 2 - parDisNozzletoFaceBaseplate + parDischargetoFaceSuction
    SBZ = parBaseplatetoSuction - parHeightBaseplate / 2
    'Use bottom of baseplate's Z value
    strOrigin = "E " & CenterPos.x & " N " & (CenterPos.y + newCenter) & " D " & (CenterPos.z + SBZ + parHeightBaseplate / 2)
    
    Dim BoxOrigin     As IJDPosition
    Dim topSurPts(3) As IJDPosition
    Dim botSurPts(3) As IJDPosition
    Dim iCount As Integer
    
    Set BoxOrigin = convertPositionStringToDPos(strOrigin)

    For iCount = 0 To 3
        Set topSurPts(iCount) = New DPosition
        Set botSurPts(iCount) = New DPosition
    Next iCount
  
    'Point 1
    botSurPts(0).x = BoxOrigin.x - parWidthBaseplate / 2
    botSurPts(0).y = BoxOrigin.y - parLengthBaseplate / 2
    botSurPts(0).z = BoxOrigin.z
    'Point 2
    botSurPts(1).x = BoxOrigin.x + parWidthBaseplate / 2
    botSurPts(1).y = BoxOrigin.y - parLengthBaseplate / 2
    botSurPts(1).z = BoxOrigin.z
    
    'Point3
    botSurPts(2).x = BoxOrigin.x + parWidthBaseplate / 2
    botSurPts(2).y = BoxOrigin.y + parLengthBaseplate / 2
    botSurPts(2).z = BoxOrigin.z
    
    'Point 4
    botSurPts(3).x = BoxOrigin.x - parWidthBaseplate / 2
    botSurPts(3).y = BoxOrigin.y + parLengthBaseplate / 2
    botSurPts(3).z = BoxOrigin.z
    
    topSurPts(0).Set botSurPts(0).x, botSurPts(0).y, botSurPts(0).z + parHeightBaseplate
    topSurPts(1).Set botSurPts(1).x, botSurPts(1).y, botSurPts(1).z + parHeightBaseplate
    topSurPts(2).Set botSurPts(2).x, botSurPts(2).y, botSurPts(2).z + parHeightBaseplate
    topSurPts(3).Set botSurPts(3).x, botSurPts(3).y, botSurPts(3).z + parHeightBaseplate
    
    Dim ObjCollection As Collection
    
    Set ObjCollection = PlaceTrapezoidWithPlanes(m_outputColl, topSurPts, botSurPts)
    For iCount = 1 To 6
         iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCollection(iCount)
    Next iCount

    Set ObjCollection = Nothing
    '==========================================
'Construction of Equipment Foundatiopn Port
'==========================================
      'Coordinate sysytem is at the bottom of support.
    '
    '                        |-------|
    '   Y                  |         |
    '   ^                  |         |
    '   |                  |         |
    '   |                  |         |
    '   |                  |         |
    '   -----> X        |         |
    '   Symbol CS |         |
    '                       |         |
    '                       |   X   |
    '                       |   ^   |
    '                       |   |   |
    '                       |   |   |
    '                       |---|--|----> Y
    '                   Port CS
    Dim ObjFoundationPort As IJEqpFoundationPort
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
   'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = BoxOrigin.z

    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#
    
    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_outputColl.ResourceManager)
    Dim holes() As Variant
    Dim dblXOffset As Double
    Dim dblYOffset As Double
    
    Call ObjFoundationPort.GetHoles(holes())

    dblXOffset = (parWidthBaseplate - parShafttoAnchorBolt * 2) / 2
    dblYOffset = (parLengthBaseplate - parFronttoBackAnchorBolt) / 2
    
    holes(0, 1) = BoxOrigin.y - parLengthBaseplate / 2 + dblYOffset
    holes(0, 2) = BoxOrigin.x - parWidthBaseplate / 2 + dblXOffset
    holes(1, 1) = BoxOrigin.y + parLengthBaseplate / 2 - dblYOffset
    holes(1, 2) = BoxOrigin.x - parWidthBaseplate / 2 + dblXOffset
    holes(2, 1) = BoxOrigin.y + parLengthBaseplate / 2 - dblYOffset
    holes(2, 2) = BoxOrigin.x + parWidthBaseplate / 2 - dblXOffset
    holes(3, 1) = BoxOrigin.y - parLengthBaseplate / 2 + dblYOffset
    holes(3, 2) = BoxOrigin.x + parWidthBaseplate / 2 - dblXOffset
                                                        
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
        
    Call ObjFoundationPort.SetHoles(holes)
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    Exit Sub

ErrorLabel:
    logError (Err.Number & " " & Err.Description & " " & Err.Source)
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub

