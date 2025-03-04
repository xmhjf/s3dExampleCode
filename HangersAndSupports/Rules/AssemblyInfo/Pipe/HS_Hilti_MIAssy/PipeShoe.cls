VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PipeShoe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright   (c) 2011, Intergraph Corporation.  All Rights reserved.
'
'   PipeShoeSupport.cls
'   ProgID:         HS_Hilti_MIAssy.PipeShoe
'   Author:         Chethan
'   Createion Date: Nov 14, 2014
'   Description:    Pipe shoe assembly for Hilti MI library
'
'   Change History:
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'   23-02-2016      Vinay               TR-CP-288050    Unable to place HiltiMIAssy supports using ‘place by Reference’ command
'   31-03-2016      Vinay               TR-CP-290838    Unable to place Hilti MI-PS2/2 Pipe Shoe assemblies
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "PipeShoe"
Private Const ENFORCE_FAILURE = False

Implements IJHgrAssmInfo
Implements IJHgrSymbolBOMServices

'Part Index's
Private SHOE_PART As Integer
Private SLIDE_PLATE As Integer
Private GUIDE_PART As Integer
Private SPACER_PLATE As Integer
Private FIXED_POINT As Integer

'Part Classes/Numbers
Private sShoePart As String
Private sSlidePlatePart As String
Private sGuidePart As String
Private sSpacerPlatePart As String
Private sFixedPoint As String

'General
Private nIndex As Long
Private nSide As Long

'Booleans
Private bIsIncludeSlidePlate As Boolean
Private bIsIncludeGuide As Boolean
Private bIsIncludeFixPoint As Boolean
Private dStandardSpace As Double
Private lShoeH As Long
Private dPipeDia As Double
Private sUnitType As String
Private sQuery As String
Private dFootPlateHeight As Double
Private sSupportNumber As String
Private dShoeWidth As Double


' ---------------------------------------------------------------------------
' Name: IJHgrAssmInfo_GetAssemblyCatalogParts()
' Description: Will Define part class and select the parts.
' Date - Author: November 3, 2008 -
'
' Inputs: pDispInputConfigHlpr As Object
' Outputs: A collection of object implementing IJDPart interface
' ---------------------------------------------------------------------------

Private Function IJHgrAssmInfo_GetAssemblyCatalogParts(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetAssemblyCatalogParts"

    On Error GoTo ErrorHandler
    'Initialize hsHlpr
    Initialize pDispInputConfigHlpr

    'Get the Part Classes / Part Numbers
    sShoePart = GetAttr("ShoePart", , , ENFORCE_FAILURE)
    sSlidePlatePart = GetAttr("SlidePlatePart", , , ENFORCE_FAILURE)
    sGuidePart = GetAttr("GuidePart", , , ENFORCE_FAILURE)
    sFixedPoint = GetAttr("FixedPoint", , , ENFORCE_FAILURE)
    sSpacerPlatePart = GetAttr("SpacerPlatePart", , , ENFORCE_FAILURE)
    bIsIncludeSlidePlate = GetAttr("IncludeSlidePlate", , , ENFORCE_FAILURE)
    bIsIncludeGuide = GetAttr("IncludeGuide", , , ENFORCE_FAILURE)
    bIsIncludeFixPoint = GetAttr("IncludeFixPoint", , , ENFORCE_FAILURE)
    lShoeH = GetAttr("ShoeHeight", , , ENFORCE_FAILURE)
    sSupportNumber = GetAttr("SupportNumber", , , ENFORCE_FAILURE)
    
    'Get the Pipe Dia and its units
    GetNOMPipeDiaAndUnits 1, dPipeDia, sUnitType
    
    Dim oPartOcc As IJPartOccAssembly  'Assembly
    Dim oPart As IJDPart               'Part

    Set oPartOcc = oICH
    Set oPart = oPartOcc.GetPart

    Dim dNDTo As Double
    Dim dNDFrom As Double
    Dim dNDPipeDia As Double
    Dim sPartNumber As String
    Dim sPartSheet As String
    Dim sPartClassName As String

    ' Get the part number and its part class
    sPartNumber = oPart.PartNumber
    sPartClassName = oPart.GetRelatedPartClassName
    sPartSheet = "JCUHAS" & sPartClassName

    If sUnitType = "in" Then
        dPipeDia = GetNOMPipeDiaByUnit(1, NPD_MILLIMETER)
    End If

    'Get the NDFrom and NDTo values
    dNDFrom = MultipleInterfaceDataLookUp("NDFrom", "JHgrSupportDefinition", sPartSheet, "PartNumber", "JDPart", "'" & sPartNumber & "'")
    dNDTo = MultipleInterfaceDataLookUp("NDTo", "JHgrSupportDefinition", sPartSheet, "PartNumber", "JDPart", "'" & sPartNumber & "'")
        
    'Check the pipe size on which axial stops are placing
    If dPipeDia < dNDFrom Or dPipeDia > dNDTo Then
        PF_EventHandler sPartNumber & " can only be placed from " & dNDFrom & "mm to " & dNDTo & "mm pipe sizes.", Err, MODULE, METHOD, False
    End If
    

    'Add Parts
    'Shoe Part
    If sShoePart = "" Or sShoePart = "0" Then
        PF_EventHandler "Pipe Shoe Support must have the shoe part", Err, MODULE, METHOD, False
    Else
        SHOE_PART = AddPart(sShoePart)
    End If

    'Slide Plate Part
    If bIsIncludeSlidePlate Then
        If sSlidePlatePart <> "" Then
           SLIDE_PLATE = AddPart(sSlidePlatePart)
        End If
    End If

    'Guide Part
    'Get the Structure Details
    Dim tStructure As hsSteelMember
    Dim dStructDepth As Double
    
    If Hilti_GetSupportingTypes(1) = "Steel" Then
        tStructure = GetSupportingSectionData(1)
        dStructDepth = tStructure.dDepth
    Else
        dStructDepth = 0.09 ' default value for cases when steel is not selected
    End If
    
    If bIsIncludeGuide Then
        If sGuidePart <> "" Then
            sGuidePart = sGuidePart & dStructDepth * 1000
            GUIDE_PART = AddPart(sGuidePart)
        End If
    End If

    'Fixed Point Part
    If bIsIncludeFixPoint Then
        If sFixedPoint <> "" Then
            sFixedPoint = sFixedPoint & dStructDepth * 1000 & "-M12"
            FIXED_POINT = AddPart(sFixedPoint)
        End If
    End If

    'set the default Height to 85 if it is none
    If lShoeH = -1 Then
        lShoeH = 1
    End If

    'Optional Spacer Plates
    If bIsIncludeGuide Then
        If lShoeH = 1 Then
            sQuery = "SELECT ShoeThick_85 from JUAHgrMIShoe WHERE SupportNumber = '" & sSupportNumber & "' AND(" & dPipeDia & " between NPDMin and NPDMax)"
            dFootPlateHeight = RunDBQuery(sQuery)
        Else
            sQuery = "SELECT ShoeThick_140 from JUAHgrMIShoe WHERE SupportNumber = '" & sSupportNumber & "' AND(" & dPipeDia & " between NPDMin and NPDMax)"
            dFootPlateHeight = RunDBQuery(sQuery)
        End If

        dStandardSpace = Round((dFootPlateHeight * 1000) + 3, 0)

        If dStandardSpace > 12 And dStandardSpace <= 16 Then
            If sSpacerPlatePart <> "" Then
                sSpacerPlatePart = sSpacerPlatePart & "-4MM"
                SPACER_PLATE = AddPart(sSpacerPlatePart)
            End If
        ElseIf dStandardSpace > 16 Then
            If sSpacerPlatePart <> "" Then
                sSpacerPlatePart = sSpacerPlatePart & "-6MM"
                SPACER_PLATE = AddPart(sSpacerPlatePart)
            End If
        Else
            'Don't Include any spacer plate
        End If
    End If

    'Return the Catalog Parts
    Set IJHgrAssmInfo_GetAssemblyCatalogParts = CatalogPartCollection

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
' ---------------------------------------------------------------------------
' Name: IJHgrAssmInfo_GetAssemblyJoints()
' Description: Will specify the joint data to the assembly symbols
' Date - Author: November 3, 2008 -
'
' Inputs: pDispInputConfigHlpr As Object
'         pDispPartOccCollection As Object
' Outputs: A collection of objects implementing the IJHgrJoint interface
' ---------------------------------------------------------------------------
Private Function IJHgrAssmInfo_GetAssemblyJoints(ByVal pDispInputConfigHlpr As Object, ByVal pDispPartOccCollection As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetAssemblyJoints"

    On Error GoTo ErrorHandler:

    'Initialize hsHlpr
    InitializeJoints pDispPartOccCollection

    'Get the Support attributes
    Dim lClampAngle As Long
    Dim dUBoltGap As Double
    Dim dPortOffset As Double
    lClampAngle = GetAttr("ClampAngle", , , ENFORCE_FAILURE)
    dUBoltGap = GetAttr("FixedPointGap", , , ENFORCE_FAILURE)

    'Set the clamp angle to zero if clamp angle is nothing or undefined
    If lClampAngle = -1 Then
        lClampAngle = 1
    End If

    SetPartAttribute SHOE_PART, "ClampRot", lClampAngle, "Long"
    SetPartAttribute SHOE_PART, "ShoeH", lShoeH, "Long"

    'Joints
    Dim JointFactory As New HgrJointFactory

    'Joint to the Shoe Part
    Rigid SHOE_PART, "Route", -1, "Route"

    'Joint to the Slide Plate
    If bIsIncludeSlidePlate Then
        If sSlidePlatePart <> "" Then
            Rigid SLIDE_PLATE, "Port1", SHOE_PART, "Girder"
        End If
    End If

    'Get Shoe Width which is common to both guide and fixed point
    Dim sShoePartProgID As String
    Dim oShoeOcc As IJPartOcc
    Dim oShoePart As IJDPart
    Dim sShoeClass As String
    Dim sQuery As String

    Set oShoeOcc = oPOC.Item(SHOE_PART)
    oShoeOcc.GetPart oShoePart
    sShoeClass = oShoePart.GetRelatedPartClassName
    sQuery = "Select ProgId  from REFDATPartClass Where Name = '" & sShoeClass & "'"
    sShoePartProgID = RunDBQuery(sQuery)

    If sShoePartProgID = "HS_Hilti_MIParts.MI_PS1_1" Then
        If lShoeH = 1 Then
            dShoeWidth = GetPartAttribute(SHOE_PART, "W_85", "Double")
        Else
            dShoeWidth = GetPartAttribute(SHOE_PART, "W_140", "Double")
        End If
    Else
        dShoeWidth = GetPartAttribute(SHOE_PART, "W", "Double")
    End If

    'Warning if both guide and fix point are set to true
    If bIsIncludeGuide And bIsIncludeFixPoint Then
        PF_EventHandler "The support configuration may not be correct. Please Re-Check", Err, MODULE, METHOD, True
    End If

    'Joint to the Guide Part
    If bIsIncludeGuide Then
        If sGuidePart <> "" Then
            Dim dLateralOffset As Double
            dLateralOffset = GetPartAttribute(GUIDE_PART, "D", "Double")

            SetPartAttribute GUIDE_PART, "GuideGap", dShoeWidth, "Double"
            Rigid GUIDE_PART, "StructureInt", -1, "Structure", 0, 0, -dLateralOffset / 2, 180, 0, 0
        End If

        'Include Spacer Plate
        If sSpacerPlatePart <> "" Then
            If dStandardSpace > 12 Then
                Dim dSpacerPlateThick As Double
                Dim dSpacerPlateDepth As Double
                dPortOffset = GetPartAttribute(GUIDE_PART, "PortOffset", "Double", , True)
                dSpacerPlateThick = GetPartAttribute(SPACER_PLATE, "Plate1Thick", "Double", , True)
                dSpacerPlateDepth = GetPartAttribute(SPACER_PLATE, "Plate1Depth", "Double", , True)
                'Set the port offset and Spacer Gap
                SetPartAttribute GUIDE_PART, "PortOffset", dPortOffset + dSpacerPlateThick, "Double"
                SetPartAttribute SPACER_PLATE, "SpacerGap", dShoeWidth + dSpacerPlateDepth, "Double"
                'spacer plate joints
                Rigid SPACER_PLATE, "StructureExt", GUIDE_PART, "StructureExt"
            End If
        End If
    End If

    'Joint to the Fixed Point
    If bIsIncludeFixPoint Then
        If sFixedPoint <> "" Then
            Dim dShoeThick As Double
            Dim dSlidePlateThick As Double
            If sShoePartProgID = "HS_Hilti_MIParts.MI_PS1_1" Then
                If lShoeH = 1 Then
                    dShoeThick = GetPartAttribute(SHOE_PART, "F_85", "Double")
                Else
                    dShoeThick = GetPartAttribute(SHOE_PART, "F_140", "Double")
                End If
            Else
                dShoeThick = GetPartAttribute(SHOE_PART, "F", "Double")
            End If
            'Set the Guide gap
            If dUBoltGap = 0 Then
                dUBoltGap = (dShoeWidth + dShoeThick) / 2
            End If

            SetAttr "FixedPointGap", dUBoltGap, False
            SetPartAttribute FIXED_POINT, "UboltGap", dUBoltGap, "Double"
            dPortOffset = GetPartAttribute(FIXED_POINT, "PortOffset", "Double", , True)

            If bIsIncludeSlidePlate Then
                dSlidePlateThick = GetPartAttribute(SLIDE_PLATE, "Thickness1", "Double", , True)
            End If

            SetPartAttribute FIXED_POINT, "PortOffset", dPortOffset + dShoeThick + dSlidePlateThick, "Double"

            dLateralOffset = GetPartAttribute(FIXED_POINT, "D", "Double")
            Rigid FIXED_POINT, "StructureInt", -1, "Structure", 0, 0, -dLateralOffset / 2, 180, 0, 0
        End If
    End If

    'Add Dimension Points
    oICH.CreateCPAtPoint "PipeCL", -1, "Route", 0, 0, 0, , "Pipe Center", True, , , cpProcessEquipment
    oICH.CreateCPAtPoint "Structure", -1, "Structure", 0, 0, 0, , "Structure", True, , , cpProcessEquipment

    Set IJHgrAssmInfo_GetAssemblyJoints = GetJoints
    Class_Terminate

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Function IJHgrAssmInfo_GetMaxRouteConnectionValue() As Long
    IJHgrAssmInfo_GetMaxRouteConnectionValue = 1
End Function

Private Function IJHgrAssmInfo_GetRouteConnections(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetRouteConnections"
    On Error GoTo ErrorHandler

    Dim RouteConnColl As New Collection
    Dim PartConnInfo(2) As Integer
    Dim nIndex As Integer

    PartConnInfo(1) = 1
    PartConnInfo(2) = 1

    RouteConnColl.Add PartConnInfo

    Set IJHgrAssmInfo_GetRouteConnections = RouteConnColl

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
Private Function IJHgrAssmInfo_GetStructConnections(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetStructConnections"

    On Error GoTo ErrorHandler

    Dim StructConnColl As New Collection
    Dim PartConnInfo(2) As Integer

    PartConnInfo(1) = 1
    PartConnInfo(2) = 1

    StructConnColl.Add PartConnInfo

    Set IJHgrAssmInfo_GetStructConnections = StructConnColl

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Sub IJHgrSymbolBOMServices_EvaluateBOM(ByVal pSupportComp As Object, bstrBOMDesc As String)
    Const METHOD = "IJHgrSymbolBOMServices_EvaluateBOM"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub





