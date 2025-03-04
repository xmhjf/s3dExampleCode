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
'   Creation Date:  Wednesday, June 06 2007
'   Description:
'       Reducing Tangential Tee 90 Degree
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Private Const INCH = 0.0254
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
        
    '{<(InputParamDec)>}
    Dim parHvacShape As Integer
    Dim parWidth As Double
    Dim parBWidth As Double
    Dim parWidth2 As Double
    Dim parDepth As Double
    Dim parBDepth As Double
    Dim parDepth2 As Double
    Dim parAngle As Double
    Dim InsulationThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHvacShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parDepth = arrayOfInputs(5)
    parBDepth = arrayOfInputs(6)
    parWidth2 = arrayOfInputs(7)
    parDepth2 = arrayOfInputs(8)
    parAngle = arrayOfInputs(9)
    InsulationThickness = arrayOfInputs(10)
    
    Dim dHeaderLength As Double
    Dim dTakeOff As Double
    
    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    If parHvacShape = Rectangular Then
        dHeaderLength = parBWidth + 2 * INCH
    ElseIf parHvacShape = 4 Then
        dHeaderLength = parBWidth + 2 * INCH
    End If
    
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
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim dCornerRadius As Double
    Dim dLength As Double
    Dim dWidth As Double, dDepth As Double
    Dim lCSType As Long

    dTakeOff = 2 * INCH ' 2 Inches
    dLength = 0.5 * dHeaderLength + 0.1 * INCH
    
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 3) As Long
    Dim dThickness(1 To 3) As Double
    Dim dFlangeWidth(1 To 3) As Double
    Dim lFlowDir(1 To 3) As Long
    Dim dPortDepth(1 To 3) As Double
    Dim dCptOffSet(1 To 3) As Double
    Dim dNozzLength(1 To 3) As Double
    
    dCornerRadius = 0#
    dNozzLength(1) = dHeaderLength
    
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
    If parHvacShape = Rectangular Then
        dDepth = parDepth
        dWidth = parWidth
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
    If parHvacShape = 4 Then 'Round
        dDepth = parWidth
        dWidth = parWidth
        lCSType = GSCADNozzleEntities.Round
    End If
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), dWidth, _
                                            dDepth, dCornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim x As Double, y As Double, z As Double
    x = -0.5 * dHeaderLength
    y = 0
    z = 0
    oPosition.Set x, y, z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation oPosition
    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    iDistribPort.SetDirectionVector oDir
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "HvacNozzle1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
    '========================
    'BUILD Strip Of Header1
    '========================
    If parHvacShape = Rectangular Then
        Dim ObjBody1 As Object
        Dim oStPoint   As New AutoMath.DPosition
        Dim oEnPoint   As New AutoMath.DPosition
        oStPoint.Set -(dHeaderLength / 2 - 0.004), 1.06 * parDepth / 2, (parWidth + 0.06 * parDepth) / 2
        oEnPoint.Set -(dHeaderLength / 2 + 0.004), -1.06 * parDepth / 2, -(parWidth + 0.06 * parDepth) / 2
        Set ObjBody1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjBody1", ObjBody1
        Set ObjBody1 = Nothing

    ElseIf parHvacShape = 4 Then 'Round
        oStPoint.Set -(dHeaderLength / 2 - 0.004), 0, 0
        oEnPoint.Set -(dHeaderLength / 2 + 0.004), 0, 0
        Set ObjBody1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth * 1.06, True)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjBody1", ObjBody1
        Set ObjBody1 = Nothing
    End If
  
    '==========================
    'BUILD TakeOff 1 Of Lateral
    '===========================
     If parHvacShape = Rectangular Then
        Dim ObjTakeOff1 As Object
        oStPoint.Set -dHeaderLength / 2, 1.01 * parDepth / 2, (parWidth + 0.01 * parDepth) / 2
        oEnPoint.Set -(dHeaderLength / 2 + 2 * INCH), -1.01 * parDepth / 2, -(parWidth + 0.01 * parDepth) / 2
        Set ObjTakeOff1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjTakeOff1", ObjTakeOff1
        Set ObjTakeOff1 = Nothing
        
    ElseIf parHvacShape = 4 Then 'Round
        oStPoint.Set -dHeaderLength / 2, 0, 0
        oEnPoint.Set -(dHeaderLength / 2 + 2 * INCH), 0, 0
        Set ObjTakeOff1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth, False)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjTakeOff1", ObjTakeOff1
        Set ObjTakeOff1 = Nothing
    End If
    
    '========================
    'BUILD Header 2
    '========================
    Dim objHeader2 As Object
    Dim dHeader2Length As Double
    If parHvacShape = 4 Then 'Round
        dHeader2Length = parWidth - parWidth2
        If CmpDblGreaterthan(dHeader2Length, 12 * INCH) Then
            dHeader2Length = 12 * INCH
        ElseIf CmpDblLessThan(dHeader2Length, 1 * INCH) Then
            dHeader2Length = 1 * INCH
        End If
        oStPoint.Set dHeaderLength / 2, 0, 0
        oEnPoint.Set (dHeaderLength / 2 + dHeader2Length), 0, 0
        Set objHeader2 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, parWidth / 2, parWidth2 / 2, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjCone", objHeader2
        Set objHeader2 = Nothing
    
    ElseIf parHvacShape = Rectangular Then
        dHeader2Length = parDepth - parDepth2
        If CmpDblGreaterthan(dHeader2Length, 12 * INCH) Then
            dHeader2Length = 12 * INCH
        ElseIf CmpDblLessThan(dHeader2Length, 1 * INCH) Then
            dHeader2Length = 1 * INCH
        End If
        oStPoint.Set dHeaderLength / 2, 0, 0
        Set objHeader2 = PlaceTrapezoid(m_OutputColl, oStPoint, parDepth, parWidth, parDepth2, parWidth2, _
                                    dHeader2Length, True, PI / 2, PI / 2, 0)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjCone", objHeader2
        Set objHeader2 = Nothing
    End If
        
    '========================
    'BUILD Strip Of Header 2
    '========================
    Dim objStrip2 As Object
    If parHvacShape = 4 Then 'Round
        oStPoint.Set dHeaderLength / 2 + dHeader2Length - 0.004, 0, 0
        oEnPoint.Set oStPoint.x + 0.008, 0, 0
        Set objStrip2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parWidth2, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength / 2 + dHeader2Length - 0.004, -1.06 * parDepth2 / 2, -(parWidth2 / 2 + 0.06 * parDepth2 / 2)
        oEnPoint.Set oStPoint.x + 0.008, 1.06 * parDepth2 / 2, (parWidth2 / 2 + 0.06 * parDepth2 / 2)
        Set objStrip2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "ObjBody2", objStrip2
    Set objStrip2 = Nothing
     
    '===========================
    'BUILD TakeOff 2 Of Lateral
    '===========================
    Dim dTakeOffLen As Double
    Dim objTakeOff As Object
    dTakeOffLen = 2 * INCH
    If parHvacShape = 4 Then 'Round
        oStPoint.Set dHeaderLength / 2 + dHeader2Length, 0, 0
        oEnPoint.Set oStPoint.x + dTakeOffLen, 0, 0
        Set objTakeOff = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parWidth2, True)
    ElseIf parHvacShape = Rectangular Then
        oStPoint.Set dHeaderLength / 2 + dHeader2Length, -1.01 * parDepth2 / 2, -(parWidth2 / 2 + 0.01 * parDepth2)
        oEnPoint.Set oStPoint.x + dTakeOffLen, 1.01 * parDepth2 / 2, (parWidth2 / 2 + 0.01 * parDepth2 / 2)
        Set objTakeOff = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    End If
    'Set the Output
    iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjTakeOff2", objTakeOff
    Set objTakeOff = Nothing
    
    '==================
    'BUILD HVACNOZZLE2
    '==================
    If parHvacShape = Rectangular Then
        dWidth = parWidth2
        dDepth = parDepth2
        lCSType = GSCADNozzleEntities.Rectangular
         x = (dHeaderLength / 2 + dHeader2Length)
         y = 0#
         z = 0#
        dNozzLength(2) = 0.001
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
    If parHvacShape = 4 Then 'Round
            dWidth = parWidth2
            dDepth = parWidth2
        lCSType = GSCADNozzleEntities.Round
         x = (dHeaderLength / 2 + dHeader2Length)
         y = 0#
         z = 0#
        dNozzLength(2) = 0.001
    End If
    
    iPortIndex = 2
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), dWidth, _
                                            dDepth, dCornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    oPosition.Set x, y, z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation oPosition
    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set 1, 0, 0
    iDistribPort.SetDirectionVector oDir
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
    '=======================
    'BUILD Body 3 Of Lateral
    '========================
    If parAngle = PI / 2 Then

        If parHvacShape = Rectangular Then
            Dim ObjBody3 As Object
            oStPoint.Set (parBWidth + 0.06 * parBDepth) / 2, (parDepth / 2 + 0.06 * parBDepth / 2), 0.5 * parWidth + INCH - 0.004
            oEnPoint.Set -(parBWidth + 0.06 * parBDepth) / 2, (parDepth / 2 - 1.06 * parBDepth), 0.5 * parWidth + INCH + 0.004
            Set ObjBody3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            iOutput = iOutput + 1
            'Set the output
            m_OutputColl.AddOutput "ObjBody3", ObjBody3
            Set ObjBody3 = Nothing
            
        ElseIf parHvacShape = 4 Then 'Round
            Dim YLength As Double
            y = parWidth / 2 - parBWidth
            YLength = Sqr((0.5 * parWidth) ^ 2 - (y * y))
            oStPoint.Set 0, parWidth / 2 - parBWidth / 2, YLength + 1 * INCH - 0.004
            oEnPoint.Set 0, parWidth / 2 - parBWidth / 2, YLength + 1 * INCH + 0.004
            Set ObjBody3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parBWidth, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "ObjBody3", ObjBody3
            Set ObjBody3 = Nothing
        End If
     
    ElseIf parAngle = 3 * PI / 2 Then
         If parHvacShape = Rectangular Then
             oStPoint.Set (parBWidth + 0.06 * parBDepth) / 2, -(parDepth / 2 + 0.06 * parBDepth), INCH + 0.5 * parWidth - 0.004
             oEnPoint.Set -(parBWidth + 0.06 * parBDepth) / 2, -(parDepth / 2 - 1.06 * parBDepth), INCH + 0.5 * parWidth + 0.004
             Set ObjBody3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
             iOutput = iOutput + 1
             'Set the output
             m_OutputColl.AddOutput "ObjBody3", ObjBody3
             Set ObjBody3 = Nothing
        
         ElseIf parHvacShape = 4 Then 'Round
             y = parWidth / 2 - parBWidth
             YLength = Sqr((0.5 * parWidth) ^ 2 - (y * y))
             oStPoint.Set 0, -(parWidth / 2 - parBWidth / 2), YLength + 1 * INCH - 0.004
             oEnPoint.Set 0, -(parWidth / 2 - parBWidth / 2), YLength + 1 * INCH + 0.004
             Set ObjBody3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.06 * parBWidth, True)
             'Set the output
             iOutput = iOutput + 1
             m_OutputColl.AddOutput "ObjBody3", ObjBody3
             Set ObjBody3 = Nothing
         End If
     End If
     
    'Take Off at the Branch
    If parAngle = PI / 2 Then
        If parHvacShape = Rectangular Then
            Dim objTakeOff3 As Object
            oStPoint.Set (parBWidth + 0.01 * parBDepth) / 2, (parDepth + 0.01 * parBDepth) / 2, parWidth / 2 + INCH
            oEnPoint.Set -(parBWidth + 0.01 * parBDepth) / 2, (parDepth / 2 - 1.01 * parBDepth), parWidth / 2 + 3 * INCH
            Set objTakeOff3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            iOutput = iOutput + 1
            'Set the output
            m_OutputColl.AddOutput "TakeOff3", objTakeOff3
            Set ObjBody3 = Nothing
            
        ElseIf parHvacShape = 4 Then 'Round
            y = parWidth / 2 - parBWidth
            YLength = Sqr((0.5 * parWidth) ^ 2 - (y * y))
            oStPoint.Set 0, (parWidth / 2 - parBWidth / 2), YLength + 3 * INCH
            oEnPoint.Set 0, (parWidth / 2 - parBWidth / 2), YLength + 1 * INCH
            Set objTakeOff3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parBWidth, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "TakeOff3", objTakeOff3
            Set objTakeOff3 = Nothing
        End If
     
    ElseIf parAngle = 3 * PI / 2 Then
         If parHvacShape = Rectangular Then
             oStPoint.Set (parBWidth + 0.01 * parBDepth) / 2, -(parDepth + 0.01 * parBDepth) / 2, parWidth / 2 + INCH
             oEnPoint.Set -(parBWidth + 0.01 * parBDepth) / 2, -(parDepth / 2 - 1.01 * parBDepth), parWidth / 2 + 3 * INCH
             Set objTakeOff3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
             iOutput = iOutput + 1
             'Set the output
             m_OutputColl.AddOutput "TakeOff3", objTakeOff3
             Set objTakeOff3 = Nothing
        
         ElseIf parHvacShape = 4 Then 'Round
             y = parWidth / 2 - parBWidth
             YLength = Sqr((0.5 * parWidth) ^ 2 - (y * y))
             oStPoint.Set 0, -(parWidth / 2 - parBWidth / 2), YLength + 3 * INCH
             oEnPoint.Set 0, -(parWidth / 2 - parBWidth / 2), YLength + 1 * INCH
             Set objTakeOff3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * parBWidth, True)
             'Set the output
             iOutput = iOutput + 1
             m_OutputColl.AddOutput "TakeOff3", objTakeOff3
             Set objTakeOff3 = Nothing
         End If
     End If
     
    If parAngle = PI / 2 Then
    
        If parHvacShape = Rectangular Then
             dDepth = parBDepth
             dWidth = parBWidth
             lCSType = GSCADNozzleEntities.Rectangular
             dNozzLength(3) = 0.001
             x = 0#
             y = parDepth / 2 - parBDepth / 2
             z = 0.5 * parWidth + 1 * INCH
             dNozzLength(3) = 0.5 * parWidth + 1 * INCH
         End If
         ' Depth and Width of crosssection will be the same as diameter of the nozzle which
         ' is the Width
         If parHvacShape = 4 Then 'Round
             dDepth = parBWidth
             dWidth = parBWidth
             lCSType = GSCADNozzleEntities.Round
             x = 0#
             y = parWidth / 2 - parBWidth / 2
             z = YLength + 1 * INCH
             dNozzLength(3) = YLength + 1 * INCH
         End If
     ElseIf parAngle = 3 * PI / 2 Then
        
        If parHvacShape = Rectangular Then
             dDepth = parBDepth
             dWidth = parBWidth
             lCSType = GSCADNozzleEntities.Rectangular
             x = 0#
             y = -(parDepth / 2 - parBDepth / 2)
             z = 0.5 * parWidth + 1 * INCH
             dNozzLength(3) = 0.5 * parWidth + 1 * INCH
         End If
         ' Depth and Width of crosssection will be the same as diameter of the nozzle which
         ' is the Width
         If parHvacShape = 4 Then 'Round
             dDepth = parBWidth
             dWidth = parBWidth
             lCSType = GSCADNozzleEntities.Round
             x = 0#
             y = -(parWidth / 2 - parBWidth / 2)
             z = YLength + 1 * INCH
             dNozzLength(3) = YLength + 1 * INCH
         End If
   End If
   
    '=================
    'BUILD HVACNOZZLE3
    '=================
    iPortIndex = 3
    
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(3), _
                                            dThickness(3), dFlangeWidth(3), lFlowDir(3), dWidth, _
                                            dDepth, dCornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", dPortDepth(3), dCptOffSet(3), False, m_OutputColl.ResourceManager)
    
    'Position of the nozzle should be the conenct point of the nozzle
    oPosition.Set x, y, z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation oPosition
    
    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set 0, 0, 1
    iDistribPort.SetDirectionVector oDir
    If parHvacShape = Rectangular Then
        Dim oRadialDir As AutoMath.DVector
        Set oRadialDir = New AutoMath.DVector
        oDir.Set 0, 0, 1
        oRadialDir.Set 0, 1, 0
        iDistribPort.SetDirectionVector oDir
        iDistribPort.SetRadialOrient oRadialDir
    End If
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(3)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "HvacNozzle3", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set oPosition = Nothing
    Set oDir = Nothing
    Set oNozzleFactory = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub




