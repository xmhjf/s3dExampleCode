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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Jan 4 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav                Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   08.SEP.2006      KKC  DI-95670          Replace names with initials in all revision history sheets and symbols
'   11.OCT.2007      PK   CR-128045         Enhance tank drain valve to provide option for
'                                           Saunders tank bottom valve.
'   28.FEB.2008      PK   CR-135703         Update the tank drain valve per following options
'                                            1)Tank drain valve, specified by face-to-center 1, face-to-center 2, and angle
'                                            2)Tank drain valve, specified by tangent length 1, face-to-center 2, and angle
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl
Private PI                  As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput                 As Double
    Dim ObjBodyCone1            As Object
    Dim ObjBodyCone2            As Object
    Dim ObjInsulatedCylinder1   As Object
    Dim ObjInsulatedCylinder2   As Object
    
    Dim parFace1toCenter        As Double
    Dim parFace2toCenter        As Double
    Dim parInsulationThickness  As Double
    
    Dim parAngle                As Double
    Dim parTangentLength1       As Double
    
    'Variables defined by svsmylav
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    Dim parHandwheelAngle   As Double
    
    Dim oDirX           As IJDVector
    Dim oDirY           As IJDVector
    Dim oDirZ           As IJDVector
    Dim oSymbolHelper   As IJSymbolGeometryHelper
    Dim oOperatorPart   As IJDPart
    Dim oOperatorOcc    As IJPartOcc
    Dim OpOrigin        As IJDPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)
    'parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    'parTangentLength1 = arrayOfInputs(6)
    'parAngle = arrayOfInputs(7)
    
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Integer
    
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    If lPartDataBasis <= 1 Or lPartDataBasis = 468 Then
    
        'Inputs
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        parAngle = arrayOfInputs(7)
                
        Dim dAngle As Double
        If lPartDataBasis <= 1 Then
            dAngle = PI / 4
        ElseIf lPartDataBasis = 468 Then
            dAngle = parAngle
        End If
        
        'Insert your code for output 1(Cone Port1 of Body)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        stPoint.Set CenterPos.x - (parFace1toCenter - flangeThick), _
                    CenterPos.y, _
                    CenterPos.z
        enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
        Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
        Set ObjBodyCone1 = Nothing
        
        'Insert your code for output 2(Cone Port2 of Body)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        stPoint.Set CenterPos.x + (parFace2toCenter - flangeThick) * Cos(dAngle), _
                    CenterPos.y - (parFace2toCenter - flangeThick) * Sin(dAngle), _
                    CenterPos.z
        
        enPoint.Set CenterPos.x, _
                    CenterPos.y, _
                    CenterPos.z
        Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
        Set ObjBodyCone2 = Nothing
        
        'Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
        Dim oPlacePoint As AutoMath.DPosition
        Dim oDir        As AutoMath.DVector
        Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
        Dim faceToFace  As Double
    
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector
        
        oPlacePoint.Set CenterPos.x - (parFace1toCenter + sptOffset - depth), _
                        CenterPos.y, _
                        CenterPos.z
        oDir.Set -1, 0, 0
        
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
        oPlacePoint.Set CenterPos.x + (parFace2toCenter + sptOffset - depth) * Cos(dAngle), _
                        CenterPos.y - (parFace2toCenter + sptOffset - depth) * Sin(dAngle), _
                        CenterPos.z
        oDir.Set Cos(dAngle), -Sin(dAngle), 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Insert your code for output 7 (Valve Operator)
        Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
        oSymbolHelper.OutputCollection = m_OutputColl
    
        On Error Resume Next
        
        Set oDirX = New DVector
        Set oDirY = New DVector
        Set oDirZ = New DVector
        
        oDirX.Set Cos(parHandwheelAngle), Sin(parHandwheelAngle), 0
        oDirY.Set 0, 0, 1
        oDirZ.Set -Sin(parHandwheelAngle), Cos(parHandwheelAngle), 0
    
        On Error GoTo ErrorLabel
        Set oPipeComponent = oPartFclt
        If Not oPipeComponent Is Nothing Then
             'Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
             '1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
             '2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
             'symbol should raise error for case 2.
             On Error GoTo HandleOperatorCodeLabel
             Set oOperatorPart = oPipeComponent.GetValveOperatorPart
             If Not oOperatorPart Is Nothing Then
                Set OpOrigin = New DPosition
                OpOrigin.Set 0, 0, 0
                Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
             End If
        End If
        Set oSymbolHelper = Nothing
        Set oOperatorPart = Nothing
        Set oPipeComponent = Nothing
        Set oOperatorOcc = Nothing
    
    ElseIf lPartDataBasis = 469 Then
        
        'Dim parCylinderDiameter As Double
        'Dim parCylinderLength   As Double
        'Dim parFacetoCenter     As Double
        Dim dDiameter       As Double
        'Inputs
        'parCylinderDiameter = arrayOfInputs(6)
        'parCylinderLength = arrayOfInputs(7)
        'parFacetoCenter = arrayOfInputs(8)
        parFace2toCenter = arrayOfInputs(3)
        parTangentLength1 = arrayOfInputs(6)
        parAngle = arrayOfInputs(7)
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        If CmpDblEqual(flangeDiam, 0) Then
            dDiameter = pipeDiam
        Else
            dDiameter = flangeDiam
        End If
        
        'Insert your code for the cylindrical flange
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Dim oFlange      As Object
            
        CenterPos.Set 0, 0, 0
        stPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
        enPoint.Set CenterPos.x - parTangentLength1, CenterPos.y, CenterPos.z
        Set oFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, dDiameter, True)
    
        'Set Output1(Flange)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oFlange
        Set oFlange = Nothing
        
        'Insert your code for body
        Dim oCircle1 As IngrGeom3D.Circle3d
        Dim oCircle2 As IngrGeom3D.Circle3d
        Dim oBody    As Object
        Dim dHeight1 As Double
        
        Set oDir = New AutoMath.DVector
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
        oDir.Set 1, 0, 0
        dHeight1 = dDiameter / 2 * Sin(parAngle)
        Set oCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        CenterPos.x, CenterPos.y, CenterPos.z, oDir.x, oDir.y, oDir.z, _
                        dDiameter / 2)
                        
        oDir.Set Cos(parAngle), Sin(parAngle), 0
        Set oCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        CenterPos.x + dHeight1, CenterPos.y + dDiameter / 2 * _
                        (1 - Cos(parAngle)), CenterPos.z, oDir.x, oDir.y, oDir.z, _
                        3 / 8 * dDiameter)
             
        Set oBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oCircle1, oCircle2, True)
            
        'Set Output2(Body)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
        Set oCircle1 = Nothing
        Set oCircle2 = Nothing
        Set oDir = Nothing
        Set oBody = Nothing
            
        'Insert your code for Operator Base
        Dim oSquare             As IngrGeom3D.ComplexString3d
        Dim oStPoint            As AutoMath.DPosition
        Dim oAxis               As AutoMath.DVector
        Dim lines               As Collection
        Dim oLine               As Object
        Dim oOperatorBase       As Object
        
        Dim iCount              As Double
        Dim dFlangeThickness    As Double
        Dim dFlangeWidth        As Double
        Dim dValue1             As Double
        Dim dValue2             As Double
        Dim dvalue3             As Double
        
        Set oAxis = New AutoMath.DVector
        Set oSquare = New IngrGeom3D.ComplexString3d
        Set oStPoint = New AutoMath.DPosition
        Set lines = New Collection
        
        dFlangeThickness = 0.5 * parTangentLength1
        dFlangeWidth = 3 / 4 * dDiameter
        dValue1 = dFlangeWidth / 2 * Sin(parAngle)
        dValue2 = dDiameter / 2 * (1 - Cos(parAngle))
        dvalue3 = dFlangeWidth / 2 * Cos(parAngle)
        
        'Insert your code for square plane
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, CenterPos.x + dHeight1 + _
                        dValue1, CenterPos.y + dValue2 - dvalue3, CenterPos.z - 0.5 * _
                        dFlangeWidth, CenterPos.x + dHeight1 + dValue1, CenterPos.y + _
                        dValue2 - dvalue3, CenterPos.z + 0.5 * dFlangeWidth)
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, CenterPos.x + dHeight1 + _
                        dValue1, CenterPos.y + dValue2 - dvalue3, CenterPos.z + 0.5 * _
                        dFlangeWidth, CenterPos.x + dHeight1 - dValue1, CenterPos.y + _
                        dValue2 + dvalue3, CenterPos.z + 0.5 * dFlangeWidth)
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, CenterPos.x + dHeight1 - _
                        dValue1, CenterPos.y + dValue2 + dvalue3, CenterPos.z + 0.5 * _
                        dFlangeWidth, CenterPos.x + dHeight1 - dValue1, CenterPos.y + _
                            dValue2 + dvalue3, CenterPos.z - 0.5 * dFlangeWidth)
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, CenterPos.x + dHeight1 - _
                        dValue1, CenterPos.y + dValue2 + dvalue3, CenterPos.z - 0.5 * _
                        dFlangeWidth, CenterPos.x + dHeight1 + dValue1, CenterPos.y + _
                        dValue2 - dvalue3, CenterPos.z - 0.5 * dFlangeWidth)
        lines.Add oLine
        oStPoint.Set CenterPos.x + dHeight1 + dValue1, CenterPos.y + dValue2 - dvalue3, _
                        CenterPos.z - 0.5 * dFlangeWidth
        Set oSquare = PlaceTrCString(oStPoint, lines)
        
        oAxis.Set Cos(parAngle), Sin(parAngle), 0
        Set oOperatorBase = PlaceProjection(m_OutputColl, oSquare, oAxis, dFlangeThickness, _
                                                                                True)
            
        'Set Output3(Operator Base)
        m_OutputColl.AddOutput "OperatorBase", oOperatorBase
        For iCount = 1 To lines.Count
            lines.Remove 1
        Next iCount
        Set oLine = Nothing
        Set lines = Nothing
        Set oOperatorBase = Nothing
        Set oAxis = Nothing
        Set oSquare = Nothing
        Set oStPoint = Nothing
        Set oGeomFactory = Nothing
            
        'Insert your code for Outlet Cylinder
        Dim oOutletCyl As Object
        Dim dValue4    As Double
            
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        dValue4 = 0.5 * dDiameter * (Sin(parAngle) * _
                                    Tan(parAngle) - 1 + Cos(parAngle))
        stPoint.Set CenterPos.x + 2 * dValue4 * Sin(parAngle) * Cos(parAngle), CenterPos.y - _
                                    2 * dValue4 * Cos(parAngle) * Cos(parAngle), CenterPos.z
        enPoint.Set CenterPos.x + (dValue4 * Cos(parAngle) + parFace2toCenter) * Sin(parAngle), CenterPos.y - _
                                    (dValue4 * Cos(parAngle) + parFace2toCenter) * Cos(parAngle), CenterPos.z
    
        Set oOutletCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
            
        'Set Output4(Outlet Cylinder)
        m_OutputColl.AddOutput "OutletCyl", oOutletCyl
        Set oOutletCyl = Nothing
    
        'Place Nozzle1
        Dim oPosition As AutoMath.DPosition
        Dim oVec      As AutoMath.DVector
        Dim oNozzle   As GSCADNozzleEntities.IJDNozzle
            
        Set oPosition = New AutoMath.DPosition
        Set oVec = New AutoMath.DVector
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
        oVec.Set -1, 0, 0
        oPosition.Set CenterPos.x - flangeThick + sptOffset - depth, CenterPos.y, CenterPos.z
        
        Set oNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oVec, oPosition)
        
        'Set Output5(Nozzle1)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
        
        Set oVec = Nothing
        Set oPosition = Nothing
        Set oNozzle = Nothing
        
        'Place Nozzle2
        Set oVec = New AutoMath.DVector
        Set oPosition = New AutoMath.DPosition
        
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
        oVec.Set Sin(parAngle), -Cos(parAngle), 0
        oPosition.Set CenterPos.x + (parFace2toCenter + dValue4 * Cos(parAngle) + (sptOffset - depth)) * Sin(parAngle), _
                      CenterPos.y - (parFace2toCenter + dValue4 * Cos(parAngle) + (sptOffset - depth)) * Cos(parAngle), _
                      CenterPos.z
        
        Set oNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oVec, oPosition)
        
        'Set Output6(Nozzle2)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
        
        Set oVec = Nothing
        Set oPosition = Nothing
        Set oNozzle = Nothing
        
        'Insert your code for valve operator
        Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
        oSymbolHelper.OutputCollection = m_OutputColl
        
        On Error Resume Next
         
        Set oDirX = New AutoMath.DVector
        Set oDirY = New AutoMath.DVector
        Set oDirZ = New AutoMath.DVector
        
        oDirX.Set Sin(parAngle), -Cos(parAngle), 0
        oDirY.Set Cos(parAngle), Sin(parAngle), 0
        oDirZ.Set 0, 0, 1
        
        On Error GoTo ErrorLabel
            
        Dim oIJDInputsArg       As IJDInputsArg
        Dim oIJDEditJDArgument  As IJDEditJDArgument
                    
        Set OpOrigin = New AutoMath.DPosition
            
        OpOrigin.Set dValue4 * Cos(parAngle) * Sin(parAngle), _
                        -dValue4 * Cos(parAngle) * Cos(parAngle), 0
        Set oPipeComponent = oPartFclt
        If Not oPipeComponent Is Nothing Then
            'Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
            '1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
            '2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
            'symbol should raise error for case 2.
            On Error GoTo HandleOperatorCodeLabel
            Set oOperatorPart = oPipeComponent.GetValveOperatorPart
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", _
                                            oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
                Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)
                    
                UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", _
                                            dFlangeWidth, oIJDEditJDArgument
                UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", _
                            "PositionerOffset", 0.5 * dDiameter * _
                            Tan(parAngle) - dValue4 * _
                            Sin(parAngle) + dFlangeThickness, _
                            oIJDEditJDArgument
    
                Set oIJDInputsArg = oOperatorOcc
                oIJDInputsArg.Update
            End If
        End If
             
        Set oSymbolHelper = Nothing
        Set oPipeComponent = Nothing
        Set oOperatorPart = Nothing
        Set oOperatorOcc = Nothing
        Set oIJDInputsArg = Nothing
        Set oIJDEditJDArgument = Nothing
        Set OpOrigin = Nothing
        
    End If
    
HandleOperatorCodeLabel:
    'The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
    If Err.Number = E_NOOPERATORPARTFOUND Then
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
            
 Private Sub UpdateOperatorAttributes(oOccurrence As Object, strInterfaceName As String, strAttribNameOnInterface As String, _
                            vInputParameterValue As Variant, oIJDEditJDArgument As IJDEditJDArgument)
    
    Const METHOD = "UpdateOperatorAttributes"
    On Error GoTo ErrorLabel
    
    Dim oIJDAttributes As IJDAttributes
    Dim oIJDAttr As IJDAttribute
    Dim oAttribCol As IJDAttributesCol
    'Dim vNewValue As Variant
    Dim lInputIndex As Long
    
    Dim oSymbol As IJDSymbol
    Dim oInputs As IJDInputs
    Dim oInput As IJDInput
    Dim lCount As Long
    Dim lCount1 As Long
    
    ''To find the index number of the Attibute to be updated in the operator code.
    Set oSymbol = oOccurrence
    Set oInputs = oSymbol.IJDSymbolDefinition(1).IJDInputs
    lCount = oInputs.Count
    For lCount1 = 1 To lCount
        Set oInput = oInputs.Item(lCount1)
        If StrComp(oInput.name, strAttribNameOnInterface, vbTextCompare) = 0 Then
            lInputIndex = oInput.index
            Exit For
        End If
    Next lCount1
    
    ''If lInputIndex  is ZERO then operator doesnt have the attribute to be updated as it's InputParameter,
    ''so that no need for futher excecution.
    If lInputIndex = 0 Then Exit Sub
        
    'Set value of the user attribute
    Set oIJDAttributes = oOccurrence
    Set oAttribCol = oIJDAttributes.CollectionOfAttributes(strInterfaceName)
    Set oIJDAttr = oAttribCol.Item(strAttribNameOnInterface)
    'vNewValue = dInputParameter
    oIJDAttr.Value = vInputParameterValue
    
    'Provide this value as input argument to the operator symbol
    UpdateSymbolArgument CDbl(vInputParameterValue), lInputIndex, oIJDEditJDArgument
    
    Exit Sub
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
   
End Sub

Private Sub UpdateSymbolArgument(dblValue As Double, lOperatorAttributeArgumentIndex As Long, _
                    oIJDEditJDArgument As IJDEditJDArgument)
                    
    Const METHOD = "UpdateSymbolArgument"
    On Error GoTo ErrorLabel
    
    Dim oPC As IJDParameterContent
    Dim oArgument As IJDArgument
                
    Set oPC = New DParameterContent
    oPC.Type = igValue
    oPC.UomType = 0
    oPC.uomValue = dblValue

    Set oArgument = New DArgument
    oArgument.index = lOperatorAttributeArgumentIndex
    oArgument.Entity = oPC
    
    oIJDEditJDArgument.SetArg oArgument
    
    Set oArgument = Nothing
    Set oPC = Nothing
    
    Exit Sub
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
        
End Sub

Private Function GetIJDEditJDArgument(oOperatorOcc As IJPartOcc) As IJDEditJDArgument

    Const METHOD = "GetIJDEditJDArgument"
    On Error GoTo ErrorLabel

    Dim oIJDEditJDArgument As IJDEditJDArgument
    Dim oIEnumJDArgument As IEnumJDArgument
    Dim oSymbol As IJDSymbol
    Dim oIJDValuesArg As IJDValuesArg
    
    Set oSymbol = oOperatorOcc
    Set oIJDValuesArg = oSymbol.IJDValuesArg
    Set oIEnumJDArgument = oIJDValuesArg.GetValues(igINPUT_ARGUMENTS_SET)
    Set oIJDEditJDArgument = oIEnumJDArgument
    Set GetIJDEditJDArgument = oIJDEditJDArgument
    
    Exit Function
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD

End Function

