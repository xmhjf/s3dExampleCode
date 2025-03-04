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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         VRK
'   Creation Date:  Wednesday, May 30 2007
'   Description:
'       Tangential Tee
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Private Const INCH = 0.0254

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
        
    '{<(InputParamDec)>}
    Dim parHVACShape As Integer
    Dim parWidth As Double
    Dim parBWidth As Double
    Dim parDepth As Double
    Dim parBDepth As Double
    Dim InsulationThickness As Double
        
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHVACShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parDepth = arrayOfInputs(5)
    parBDepth = arrayOfInputs(6)
    InsulationThickness = arrayOfInputs(7)
    
    Dim dHeaderLength As Double
    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    'If this symbol is placed on a round duct then Depth parameter is ZERO.
    'To guard the symbol depth is made equal to width.
    If parDepth = 0 Then parDepth = parWidth
    If parBDepth = 0 Then parBDepth = parBWidth

    '===================================
    'BUILD HVAC NOZZLE ON BASIC ASSEMBLY
    '===================================
    Dim oNozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim oPosition As New AutoMath.DPosition
    Dim odir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double
    Dim Width As Double, Depth As Double
    Dim lCSType As Long
    
    If parHVACShape = Rectangular Then
        dHeaderLength = parBWidth + 2 * INCH
    ElseIf parHVACShape = 4 Then
        dHeaderLength = parBWidth + 2 * INCH
    End If
    
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 3) As Long
    Dim dThickness(1 To 3) As Double
    Dim dFlangeWidth(1 To 3) As Double
    Dim lFlowDir(1 To 3) As Long
    Dim dPortDepth(1 To 3) As Double
    Dim dCptOffSet(1 To 3) As Double
    Dim dNozzLength(1 To 3) As Double
    
    iPortIndex = 1
    'Set HVAC nozzle parameters
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
            dThickness(iPortIndex) = 0.0001
        End If
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + 0.001
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
 
    iPortIndex = 1
    CornerRadius = 0#
    dNozzLength(1) = 0.5 * dHeaderLength
    If parHVACShape = Rectangular Then
        Depth = parDepth
        Width = parWidth
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
    If parHVACShape = 4 Then 'Round
        Depth = parWidth
        Width = parWidth
        lCSType = GSCADNozzleEntities.Round
    End If

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim x As Double, y As Double, Z As Double
    x = -0.5 * dHeaderLength
    y = 0
    Z = 0
    oPosition.Set x, y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation oPosition
    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    odir.Set -1, 0, 0
    iDistribPort.SetDirectionVector odir
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "HvacNozzle1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
    '========================
    'BUILD Body 1 Of Tee
    '========================
    If parHVACShape = Rectangular Then
        Dim ObjBody1 As Object
        Dim ostPoint   As New AutoMath.DPosition
        Dim oEnPoint   As New AutoMath.DPosition
        ostPoint.Set -(dHeaderLength / 2 - 0.004), 0.5 * (parDepth * 1.06), (parWidth + 0.06 * parDepth) / 2
        oEnPoint.Set -(dHeaderLength / 2 + 0.004), -0.5 * parDepth * 1.06, -(parWidth + 0.06 * parDepth) / 2
        Set ObjBody1 = PlaceBox(m_OutputColl, ostPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjBody1", ObjBody1
        Set ObjBody1 = Nothing
    
    ElseIf parHVACShape = 4 Then 'Round
        ostPoint.Set -(dHeaderLength / 2 - 0.004), 0, 0
        oEnPoint.Set -(dHeaderLength / 2 + 0.004), 0, 0
        Set ObjBody1 = PlaceCylinder(m_OutputColl, ostPoint, oEnPoint, parWidth * 1.06, True)
        iOutput = iOutput + 1
        ' Set the output
        m_OutputColl.AddOutput "ObjBody1", ObjBody1
        Set ObjBody1 = Nothing
    End If
    
    '==========================
    'BUILD TakeOff 1 Of Tee
    '===========================
    If parHVACShape = Rectangular Then
        Dim ObjTakeOff1 As Object
        ostPoint.Set -(dHeaderLength / 2 + 2 * INCH), (parDepth + 0.01 * parDepth) / 2, (parWidth + 0.01 * parDepth) / 2
        oEnPoint.Set -dHeaderLength / 2, -(parDepth + 0.01 * parDepth) / 2, -(parWidth + 0.01 * parDepth) / 2
        Set ObjTakeOff1 = PlaceBox(m_OutputColl, ostPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjTakeOff1", ObjTakeOff1
        Set ObjTakeOff1 = Nothing
    
    ElseIf parHVACShape = 4 Then 'Round
        ostPoint.Set -(dHeaderLength / 2 + 2 * INCH), 0, 0
        oEnPoint.Set -dHeaderLength / 2, 0, 0
        Set ObjTakeOff1 = PlaceCylinder(m_OutputColl, ostPoint, oEnPoint, 1.01 * parWidth, False)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjTakeOff1", ObjTakeOff1
        Set ObjTakeOff1 = Nothing
    End If
    
    '========================
    'BUILD Body 2 Of Tee
    '========================
    If parHVACShape = Rectangular Then
        Dim ObjBody2 As Object
        ostPoint.Set (dHeaderLength / 2 - 0.004), (parDepth * 1.06) / 2, (parWidth + 0.06 * parDepth) / 2
        oEnPoint.Set (dHeaderLength / 2 + 0.004), -(1.06 * parDepth) / 2, -(parWidth + 0.06 * parDepth) / 2
        Set ObjBody2 = PlaceBox(m_OutputColl, ostPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjBody2", ObjBody2
        Set ObjBody2 = Nothing
    
    ElseIf parHVACShape = 4 Then 'Round
        ostPoint.Set dHeaderLength / 2 + 0.004, 0, 0
        oEnPoint.Set dHeaderLength / 2 - 0.004, 0, 0
        Set ObjBody2 = PlaceCylinder(m_OutputColl, ostPoint, oEnPoint, parWidth * 1.06, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjBody2", ObjBody2
        Set ObjBody2 = Nothing
  End If
  
    '===========================
    'BUILD TakeOff 2 Of Tee
    '===========================
    If parHVACShape = Rectangular Then
        Dim ObjTakeOff2 As Object
        ostPoint.Set dHeaderLength / 2 + 2 * INCH, (1.01 * parDepth) / 2, (parWidth + 0.01 * parDepth) / 2
        oEnPoint.Set dHeaderLength / 2, -(1.01 * parDepth) / 2, -(parWidth + 0.01 * parDepth) / 2
        Set ObjTakeOff2 = PlaceBox(m_OutputColl, ostPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjTakeOff2", ObjTakeOff2
        Set ObjTakeOff2 = Nothing

    ElseIf parHVACShape = 4 Then 'Round
        ostPoint.Set dHeaderLength / 2, 0, 0
        oEnPoint.Set dHeaderLength / 2 + 2 * INCH, 0, 0
        Set ObjTakeOff2 = PlaceCylinder(m_OutputColl, ostPoint, oEnPoint, 1.01 * parWidth, False)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjTakeOff2", ObjTakeOff2
        Set ObjTakeOff2 = Nothing
    End If
    
    '==================
    'BUILD HVACNOZZLE2
    '==================
    iPortIndex = 2
    dNozzLength(2) = 0.5 * dHeaderLength
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    x = 0.5 * dHeaderLength
    y = 0#
    Z = 0#
    oPosition.Set x, y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation oPosition
    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    odir.Set 1, 0, 0
    iDistribPort.SetDirectionVector odir
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
    '=======================
    'BUILD Body 3 Of Tee
    '========================
    If parHVACShape = Rectangular Then
        Dim ObjBody3 As Object
        ostPoint.Set -(parBWidth / 2 + 0.06 * parBWidth), (parDepth / 2 + 0.06 * parBWidth), _
                                    parWidth / 2 + 1 * INCH - 0.004
        oEnPoint.Set (parBWidth / 2 + 0.06 * parBWidth), (parDepth / 2 - parBDepth - 0.06 * parBWidth), _
                                parWidth / 2 + 1 * INCH + 0.004
        Set ObjBody3 = PlaceBox(m_OutputColl, ostPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjBody3", ObjBody3
        Set ObjBody3 = Nothing
   
    ElseIf parHVACShape = 4 Then 'Round
        Dim YLength As Double
        y = Width / 2 - parBWidth
        x = Sqr((0.5 * Width) ^ 2 - (y * y))
        YLength = x
        ostPoint.Set 0, parWidth / 2 - parBWidth / 2, YLength + 1 * INCH - 0.004
        oEnPoint.Set 0, parWidth / 2 - parBWidth / 2, YLength + 1 * INCH + 0.004
        Set ObjBody3 = PlaceCylinder(m_OutputColl, ostPoint, oEnPoint, 1.06 * parBWidth, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjBody3", ObjBody3
        Set ObjBody3 = Nothing
    End If
    
    '===========================
    'BUILD TakeOff 3 Of Tee
    '===========================
    If parHVACShape = Rectangular Then
        Dim ObjTakeOff3 As Object
        ostPoint.Set -(parBWidth / 2 + 0.01 * parBWidth), (parDepth / 2 + 0.01 * parBWidth), _
                                parWidth / 2 + INCH
        oEnPoint.Set (parBWidth / 2 + 0.01 * parBWidth), (parDepth / 2 - parBDepth - 0.01 * parBWidth), _
                                parWidth / 2 + 3 * INCH
        Set ObjTakeOff3 = PlaceBox(m_OutputColl, ostPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjTakeOff2", ObjTakeOff3
        Set ObjTakeOff3 = Nothing
  
    ElseIf parHVACShape = 4 Then 'Round
        ostPoint.Set 0, parWidth / 2 - parBWidth / 2, YLength + INCH
        oEnPoint.Set 0, parWidth / 2 - parBWidth / 2, YLength + 3 * INCH
        Set ObjTakeOff3 = PlaceCylinder(m_OutputColl, ostPoint, oEnPoint, 1.01 * parBWidth, False)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjTakeOff3", ObjTakeOff3
        Set ObjTakeOff3 = Nothing
    End If
    
    If parHVACShape = Rectangular Then
        Depth = parBDepth
        Width = parBWidth
        lCSType = GSCADNozzleEntities.Rectangular
        dNozzLength(3) = 1 * INCH
        x = 0#
        y = parDepth / 2 - parBDepth / 2
        Z = 0.5 * parWidth + 1 * INCH
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
    If parHVACShape = 4 Then 'Round
        Depth = parBWidth
        Width = parBWidth
        lCSType = GSCADNozzleEntities.Round
        x = 0#
        y = parWidth / 2 - parBWidth / 2
        Z = YLength + 1 * INCH
        dNozzLength(3) = YLength + 1 * INCH
    End If

    '=================
    'BUILD HVACNOZZLE3
    '=================
    iPortIndex = 3
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(3), _
                                            dThickness(3), dFlangeWidth(3), lFlowDir(3), Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", dPortDepth(3), dCptOffSet(3), False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    oPosition.Set x, y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation oPosition

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    odir.Set 0, 0, 1
    iDistribPort.SetDirectionVector odir
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(3)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "HvacNozzle3", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set oNozzleFactory = Nothing
    Set oPosition = Nothing
    Set odir = Nothing
    Set ostPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub


