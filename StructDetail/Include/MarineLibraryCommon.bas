Attribute VB_Name = "MarineLibraryCommon"
'*******************************************************************
'
'Copyright (C) 2014-16 Intergraph Corporation. All rights reserved.
'
'File : MarineLibraryCommon.bas
'
'Author : Alligators
'
'Description :
'
'
'History :
'    14/Feb/2014 - svsmylav
'         TR-245505: Added new method 'ForceUpdateOnMemberObjects'.
'
'    11/Aug/2014 - knukala
'         CR-CP-250020: Create AC for ladder support or cage support to ladder rail
'    03/Nov/2014 - MDT/GH
'         CR-CP-250198  Lapped AC for traffic items
'    15/Nov/2016 - svsmylav
'         TR-282194: 'oParentSmartClass' is checked for 'Nothing' in GetSelectorAnswer method,
'                     if it is so then exit the method (to avoid record-exceptions).
'                     Also, unused variables are removed from this file.
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'                   Added Enums and constants and also a new method "AddFeatureEndCutData" and moved one of the existing method to this bas module for the better usage.
'*****************************************************************************
Option Explicit

Public Const ToFilletRoot = 0
Public Const AvoidWeb = 1
Public Const ToFlangeInnerSurface = 2
Public Const ToBoundingEdge = 3

Public Const CUSTOMERID = "SM"
Public Const LINEAR_TOLERANCE_Mbr = 0.00001

Private Const MODULE = "S:\StructDetail\Data\Include\MarineLibraryCommon.bas"
'ConnectedEdgeInfo specifies the edges on the bounding member that the bounded
'member is connected to.  IntersectingEdge is the edge that the bounded member
'is intersecting.  CoplanarEdge is the edge that the bounded member is coplanar
'to. There will be a ConnectedEdgeInfo for each bounded edge in the measurement
'symbol
Public Type ConnectedEdgeInfo
    IntersectingEdge As eBounding_Edge
    CoplanarEdge As eBounding_Edge
End Type

'eBounding_Alias is the cross section alias, used to determine which measurement
'symbol to use.
Public Enum eBounding_Alias
    Web = 1
    WebTopFlangeRight = 2
    WebBuiltUpTopFlangeRight = 3
    WebBottomFlangeRight = 4
    WebBuiltUpBottomFlangeRight = 5
    WebTopAndBottomRightFlanges = 6
    FlangeLeftAndRightBottomWebs = 7
    FlangeLeftAndRightTopWebs = 8
    FlangeLeftAndRightWebs = 9
    Tube = 10
End Enum
'eBounding_Edge has all the possible edges on the bounding member. Above and Below
'are used when the bounded edge is above or below the bounding member. None is used
'when the bounded edge isn't coplanar to any of the bounding edges.
Public Enum eBounding_Edge
    Above = 1
    Top = 514
    Web_Right_Top = 2564
    Top_Flange_Right_Top = 2052
    Top_Flange_Right = 1028
    Top_Flange_Right_Bottom = 772
    Web_Left = 257
    Web_Right = 258
    Bottom_Flange_Right_Top = 771
    Bottom_Flange_Right = 1027
    Bottom_Flange_Right_Bottom = 2051
    Web_Right_Bottom = 2563
    Bottom = 513
    Below = 2
    None = 3
End Enum
'Structure of Member Connection data
Public Type MemberConnectionData
    Matrix As IJDT4x4
    ePortId As Integer
    AxisPort As IJPort
    AxisCurve As IJCurve
    MemberPart As Object
End Type

'Enum of Member Assembly Connection type
Public Enum eACType
    ACType_None = 0
    ACType_Mbr_Generic = 1
    ACType_Axis = 2
    ACType_Split = 3
    ACType_Miter = 4
    ACType_Bounded = 5
    ACType_Stiff_Generic = 6
End Enum

Public Enum eMultiBoundingEdgeIDs

    e_JXSEC_MultipleBounding_5001 = 5001
    e_JXSEC_MultipleBounding_5002 = 5002
    e_JXSEC_MultipleBounding_5003 = 5003
    e_JXSEC_MultipleBounding_5004 = 5004
    e_JXSEC_MultipleBounding_5005 = 5005
    
End Enum

Public Const C_Port_Base = "Base"
Public Const C_Port_Offset = "Offset"
Public Const C_Port_Lateral = "Lateral"
Public Const C_Port_Top = "Top"
Public Const C_Port_Bottom = "Bottom"
Public Const C_Port_WebLeft = "WebLeft"
Public Const C_Port_WebRight = "WebRight"

Const gsStraightNoOffsetWebCuts = "Straight No Offset WebCuts"
Const gsStraightOffsetWebCuts = "Straight Offset WebCuts"
Const gsSnipedNoOffsetWebCuts = "Sniped No Offset WebCuts"
Const gsSnipedOffsetWebCuts = "Sniped Offset WebCuts"
Const gsStraightNoOffsetFlangeCuts = "Straight No Offset FlangeCuts"
Const gsStraightOffsetFlangeCuts = "Straight Offset FlangeCuts"
Const gsSnipedNoOffsetFlangeCuts = "Sniped No Offset FlangeCuts"
Const gsSnipedOffsetFlangeCuts = "Sniped Offset FlangeCuts"

Public Enum EndCutRelativePosition
    Primary = 0
    TopOrLeft = 1
    BottomOrRight = 2
End Enum

Public Enum WebCutDrawingType
    Straight_No_Offset_WebCuts = 0
    Straight_Offset_WebCuts = 1
    Sniped_No_Offset_WebCuts = 2
    Sniped_Offset_WebCuts = 3
End Enum

Public Enum FlangeCutDrawingType
    Straight_No_Offset_FlangeCuts = 0
    Straight_Offset_FlangeCuts = 1
    Sniped_No_Offset_FlangeCuts = 2
    Sniped_Offset_FlangeCuts = 3
End Enum
'


Public Sub GetSelectorAnswer(oOccurrence As Object, strQuestion As String, _
                                strAnswer As Variant, _
                                Optional SmartObject As Object)

    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    
    Dim oParameterLogic As IJDParameterLogic
    Dim oMemberDescription As IJDMemberDescription
    Dim oSelectorLogic As IJDSelectorLogic
    
    Dim oParentSmartClass As IJSmartClass
    
    If oOccurrence Is Nothing Then
        'Need to error out in such cases.
        'For invalid arguments strAnswer is set to ""
        strAnswer = ""
        Exit Sub
    End If
      
    If TypeOf oOccurrence Is IJDMemberDescription Then
        Set oMemberDescription = oOccurrence
        Set oSmartOccurrence = oMemberDescription.CAO
        Set oSmartItem = oSmartOccurrence.ItemObject
        Set oParentSmartClass = oSmartItem.Parent
        
    ElseIf TypeOf oOccurrence Is IJDParameterLogic Then
        Set oParameterLogic = oOccurrence
        Set oSmartItem = oParameterLogic.SmartItem
        Set oSmartOccurrence = oParameterLogic.SmartOccurrence
        Set oParentSmartClass = oSmartItem.Parent
        
    ElseIf TypeOf oOccurrence Is IJSmartOccurrence Then
         Set oSmartOccurrence = oOccurrence
         Set oSmartItem = oSmartOccurrence.ItemObject
         Set oParentSmartClass = oSmartItem.Parent
         
    ElseIf TypeOf oOccurrence Is IJDSelectorLogic Then
         Set oSelectorLogic = oOccurrence
         Set oSmartOccurrence = oSelectorLogic.SmartOccurrence

        On Error Resume Next
         strAnswer = oSelectorLogic.Answer(strQuestion)
        Err.Clear
        On Error GoTo ErrorHandler
        
         Set oParentSmartClass = oSmartOccurrence.RootSelectionObject
         
    ElseIf TypeOf oOccurrence Is IJSmartClass Then
        Set oSmartOccurrence = SmartObject
        Set oParentSmartClass = oOccurrence
    
    End If
    
    If oParentSmartClass Is Nothing Then
        Exit Sub 'Execution cannot proceed any further, so exit
    End If
    
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oSymbolDefinition = oParentSmartClass.SelectionRuleDef
    
    Dim pAttributes As IJDAttributes
    Dim pAttrCol As IJDAttributesCol
    Dim pCmd As IJDCommandDescription
    Dim pAtt As IJDAttribute
    Dim pObject As IJDObject
    Dim pCodeListMD As IJDCodeListMetaData
    
    Set pObject = oSmartOccurrence
    Set pCodeListMD = pObject.ResourceManager
    
    Dim interfaceName As String
    
    'Fix for TR-231923 GetSelectorAnswer() in MarineLibraryCommon is giving exception
    'Should work for .NET and COM. Currently only excersied for .NET as the SelectionRuleDef is null for .NET by design
    If oSymbolDefinition Is Nothing Then
        interfaceName = oParentSmartClass.SelectionRuleInterface
    Else
        On Error Resume Next
        Set pCmd = oSymbolDefinition.IJDUserCommands.GetCommand("AnswersIID")
        If Len(pCmd.Source) < 1 Then
            Exit Sub
        End If
        Err.Clear
        On Error GoTo ErrorHandler:
        interfaceName = pCmd.Source
    End If
    
    Set pAttributes = oSmartOccurrence
    Set pAttrCol = pAttributes.CollectionOfAttributes(interfaceName)
    
    If Not pAttrCol Is Nothing Then
        For Each pAtt In pAttrCol
            If strQuestion = pAtt.AttributeInfo.Name Then
                If pAtt.AttributeInfo.CodeListTableName <> "" Then
                    If pAtt.Value < 65536 Then
                        strAnswer = pAtt.Value
                    ElseIf pAtt.Value >= (65536 * 2) Then ' it's a double code list
                        strAnswer = CDbl(pCodeListMD.ShortStringValue(pAtt.AttributeInfo.CodeListTableName, pAtt.Value))
                    Else ' it is a string code list
                        strAnswer = pCodeListMD.ShortStringValue(pAtt.AttributeInfo.CodeListTableName, pAtt.Value)
                    End If
                Else
                    strAnswer = pAtt.Value
                End If
            Else
                If CStr(strAnswer) = "0" Then
                    On Error Resume Next
                    strAnswer = CStr("")
                    Err.Clear
                    On Error GoTo ErrorHandler
               End If
            End If
        Next pAtt
    End If
    
     If strAnswer = vbEmpty Or Len(CStr(strAnswer)) = 0 Then
        Dim oGrandParentSmartClass As DEFINITIONHELPERSINTFLib.IJSmartClass
         If TypeOf oParentSmartClass Is DEFINITIONHELPERSINTFLib.IJSmartClass Then
            If TypeOf oParentSmartClass.Parent Is DEFINITIONHELPERSINTFLib.IJSmartClass Then
             Set oGrandParentSmartClass = oParentSmartClass.Parent
             GetSelectorAnswer oGrandParentSmartClass, strQuestion, strAnswer, oSmartOccurrence
            End If
        End If
    End If
    Exit Sub

ErrorHandler:
    ' Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub

' ********************************************************************************
' Method:
'   LandingCurve
' Description:
'   Gets the landingcurve for the given Stiffener/ER/Beam
' ********************************************************************************
' This method copied from EndCutRules\Common.bas
' EndCuts should be made to use the MarineLibraryCommon.bas file, which is in a more public location
Public Function GetProfilePartLandingCurve(oProfilePart As Object) As IJWireBody
    On Error GoTo ErrorHandler

    Dim oStructDetailHelper As StructDetailHelper
    Dim oTopoLocate As IJTopologyLocate
    Dim oLandingCrv As IJWireBody

    Set oStructDetailHelper = New StructDetailHelper
    Set oTopoLocate = New TopologyLocate

    'Based on whether the part is derived from system or not,
    'we shall get the landing curve differently for performance

    Dim oIJStructGraph As IJStructGraph
    Set oIJStructGraph = oProfilePart

    If Not oIJStructGraph Is Nothing Then
        Dim oParentSystem As IJSystem
        oStructDetailHelper.IsPartDerivedFromSystem oIJStructGraph, oParentSystem

        'Derived from system?
        If Not oParentSystem Is Nothing Then
            Set oLandingCrv = oTopoLocate.GetProfileParentWireBody(oProfilePart)
        Else
            'Not derived from system
            Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
            Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport

            Set oProfilePartSupport = New GSCADSDPartSupport.ProfilePartSupport
            Set oPartSupport = oProfilePartSupport
            Set oPartSupport.Part = oProfilePart

            ' Get the curve (default direction is SideUnspecified, which gives
            ' a curve through the load point)
            Dim oThicknessDir As IJDVector
            Dim bThicknessCentered As Boolean

            oProfilePartSupport.GetProfilePartLandingCurve oLandingCrv, _
                                                           oThicknessDir, _
                                                           bThicknessCentered

            Set oProfilePartSupport = Nothing
            Set oPartSupport = Nothing
            Set oThicknessDir = Nothing
        End If

        Set oParentSystem = Nothing
    End If

    Set GetProfilePartLandingCurve = oLandingCrv

    Set oLandingCrv = Nothing
    Set oStructDetailHelper = Nothing
    Set oTopoLocate = Nothing

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetProfilePartLandingCurve").Number
End Function

' The following methods were copied from StructDetailObjects.  A common implementation should be used.

Public Function GetDoubleAttribute(ByVal oCrossSection As IJCrossSection, ByVal Name As String) As Double
    Const MT = "GetDoubleAttribute"

    On Error GoTo ErrorHandler

    Dim oAttribute As IJDAttribute
    Dim oAttributeInfo As IJDAttributeInfo
    Dim oAttributesCol As IJDAttributesCol

    Dim vAttributeValue As Variant
    Get_AttributeValue oCrossSection, Name, vAttributeValue
    GetDoubleAttribute = vAttributeValue
    Exit Function

    Set oAttributesCol = oCrossSection.Attributes
    Set oAttribute = oAttributesCol.Item(Name)
    Set oAttributeInfo = oAttribute.AttributeInfo
    If oAttributeInfo.Type = 8 Then  ' double = VT_R8 = SQL_C_DOUBLE = SQL_DOUBLE
        GetDoubleAttribute = oAttribute.Value
    Else
        GoTo ErrorHandler
    End If
    Set oAttribute = Nothing
    Set oAttributeInfo = Nothing
    Set oAttributesCol = Nothing

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
    Debug.Assert False
End Function

Public Sub Get_AttributeValue(oObject As IJDObject, _
                              sAttributeName As String, _
                              vAttributeValue As Variant)
    Const MT = "Get_AttributeValue"
    On Error GoTo ErrorHandler

    Dim iIndex As Long
    Dim nAttributes As Long
    Dim AttributeData() As CustomAttributeData

    '========TR-CP·108807=========
    If oObject Is Nothing Then

        Err.Raise vbObjectError, MT, "Failed to get object in Get_AttributeValue"

    End If
    '=============================

    Get_CustomAttributes oObject, nAttributes, AttributeData
    For iIndex = 1 To nAttributes
        If Trim(LCase(sAttributeName)) = Trim(LCase(AttributeData(iIndex).Name)) Then
            vAttributeValue = AttributeData(iIndex).Value
            Exit Sub
        End If

    Next iIndex

    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub

Public Sub Get_CustomAttributes(oObject As IJDObject, _
                                ByRef nAttributes As Long, _
                                ByRef AttributeData() As CustomAttributeData)
    Const MT = "Congifuration.Get_CustomAttributes"
    On Error GoTo ErrorHandler

    Dim InterfaceID As Variant
    Dim oAttributes As IJDAttributes
    Dim oAttributeMetaData As IJDAttributeMetaData

    Dim iCount As Long
    Dim nInterface As Long
    Dim oAttributeColl As CollectionProxy
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oInterfaceAttributeCollection As IJDInfosCol
    Dim oInterfaceSqlAttributeCollection As IJDInfosCol

    iCount = 0
    nInterface = 0
    Set oAttributes = oObject
    Set oAttributeMetaData = oAttributes

    ' Verify that the Attributes contain at least one
    On Error Resume Next
    
    If oAttributes.Count > 0 Then
        ' Loop thru each Interface in the Attributes Collection
        nInterface = oAttributes.Count
        For Each InterfaceID In oAttributes
            iCount = iCount + 1
            Set oAttributeColl = oAttributes.CollectionOfAttributes(InterfaceID)

            ' verify the current interface Collection is valid
            If Not oAttributeColl Is Nothing Then

                ' veirfy that the current Attribute Interface collection
                ' represents a "User" Attribute interface not System Attribute(??)
                ' ....oAttributeMetaData.InterfaceInfo(InterfaceID).IsHidden
                Set oInterfaceInfo = oAttributeColl.InterfaceInfo
                If oInterfaceInfo.IsHardCoded = False Then

                    ' retrieve collection of Attributes from the InterfaceInfo object
                    ' There appears to be two(2) types of collections
                    '   AttributeCollection is a collection of 'COM' attributes
                    '   SQLAttributeCollection is a collection of 'SQL' attributes
                    ' Note, an attribute can be 'COM' or 'SQL' or both
                    Set oInterfaceAttributeCollection = oInterfaceInfo.AttributeCollection
                    If Not oInterfaceAttributeCollection Is Nothing Then
                        Get_AttributesFromCollection oInterfaceAttributeCollection, _
                                                     oAttributeColl, _
                                                     nAttributes, _
                                                     AttributeData
                        Set oInterfaceAttributeCollection = Nothing
                    End If

                    'NOTE:
                    ' For now only get the Attributes that returned in the AttributeCollection
                    ' attempting to retrieve the SQLAttributeCollection causes Asserts
                    ' because the Attribute can not be gotten by it's Name ??????
                    '$$$                    Set oInterfaceSqlAttributeCollection = oInterfaceInfo.SQLAttributeCollection
                    If Not oInterfaceSqlAttributeCollection Is Nothing Then
                        Get_AttributesFromCollection oInterfaceSqlAttributeCollection, _
                                                     oAttributeColl, _
                                                     nAttributes, _
                                                     AttributeData
                        Set oInterfaceSqlAttributeCollection = Nothing
                    End If

                End If

                Set oInterfaceInfo = Nothing
                Set oAttributeColl = Nothing
            End If


            '$$$ Kludge
            '$$$ for some unknown reason
            '$$$ the For Each InterfaceID In oAttributes loop cycles thru more then
            '$$$ the oAttributes.Count times,
            '$$$ this results in some Attributes being process more then once
            '$$$ to prevent the double Attributes,
            '$$$ exit the For loop when the number of Interfaces processed exceeds the
            '$$$ oAttributes.Count value
            If iCount >= nInterface Then
                Exit For
            End If
        Next    ' interface
    End If

    Err.Clear
    On Error GoTo ErrorHandler

    Set oAttributeMetaData = Nothing
    Set oAttributes = Nothing


    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number

End Sub

Private Sub Get_AttributesFromCollection(oInterfaceInfosCol As IJDInfosCol, _
                                         oCollectionProxy As CollectionProxy, _
                                         ByRef nAttributes As Long, _
                                         ByRef AttributeData() As CustomAttributeData)
    Const MT = "Congifuration.Get_AttributesFromCollection"
    On Error GoTo ErrorHandler

    Dim sName As String
    Dim sUserName As String
    Dim iIndex As Long
    Dim nDispatchId As Long

    Dim oAttribute As IJDAttribute
    Dim oAttributeInfo As IJDAttributeInfo

    iIndex = 0
    
    On Error Resume Next

    For Each oAttributeInfo In oInterfaceInfosCol
        ' verify the current AttirbuteInfo is valid
        iIndex = iIndex + 1
        If Not oAttributeInfo Is Nothing Then

            ' Retreive the current Attribute's Name
            ' use the Attribute's Name to retierve the Attribute object from
            ' the given Attribute Collection
            ' attempting to retrieve the Attribute using the collection Index
            ' sometimes causes an Assert (error) because the underlaying code
            ' uses the given Index number as the Attribute's DispatchID
            ' and the Index value and DispatchID do not always agree
            sName = oAttributeInfo.Name
            Set oAttribute = oCollectionProxy.Item(sName)

            ' verify the current Attirbute is valid
            If Not oAttribute Is Nothing Then

                ' re-dimension the output array
                ' then add the current Attribute to the output array
                nAttributes = nAttributes + 1
                ReDim Preserve AttributeData(nAttributes)

                sUserName = ""
                sUserName = oAttributeInfo.UserName
                If Len(Trim(sUserName)) < 1 Then
                    sUserName = oAttributeInfo.Name
                End If

                nDispatchId = oAttributeInfo.dispid
                AttributeData(nAttributes).Name = oAttributeInfo.Name
                AttributeData(nAttributes).Value = oAttribute.Value
                AttributeData(nAttributes).UnitsType = oAttributeInfo.UnitsType
                AttributeData(nAttributes).UserName = sUserName
                AttributeData(nAttributes).interfaceName = oAttributeInfo.interfaceName

                Set oAttribute = Nothing
            End If

            Set oAttributeInfo = Nothing
        End If
    Next    ' oAttributeInfo

    Err.Clear
    On Error GoTo ErrorHandler
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number

End Sub
' =========================================================================
Public Function GetMajorAxis(ByVal pPRL As IJDParameterLogic) As Double
    Const METHOD = "::GetMajorAxis"

    On Error GoTo ErrorHandler

    Dim oStructFeature As IJStructFeature
    Dim oFeaturePosition As IJDPosition
    Dim oClosestPoint As IJDPosition
    Dim oBoundingPart As Object
    Dim sSectionProperty As String
    
    Set oStructFeature = pPRL.SmartOccurrence

    ' --------------------------------------------------------------------------------
    ' Determine the Struct Feature Type and get the feature position and bounding part
    ' --------------------------------------------------------------------------------
    Dim featureType As StructFeatureTypes
    featureType = oStructFeature.get_StructFeatureType
    Select Case featureType
    Case SF_CornerFeature
        sSectionProperty = "FilletRadius"
        'Get the Corner feature part object
        Dim oCornerFeature As New StructDetailObjects.CornerFeature
        Dim oCornerPart As Object
        
        Set oCornerFeature.object = oStructFeature
        Set oCornerPart = oCornerFeature.GetPartObject
        
        'Get the corner feature location
        oCornerFeature.GetLocationOfCornerFeature oFeaturePosition

        'Get the bounding object from the corner features parent cut
        Dim oDesignChild As IJDesignChild
        Set oDesignChild = pPRL.SmartOccurrence
        
        Dim oCornerParent As Object
        Dim oParentStructFeat As IJStructFeature
        Set oCornerParent = oDesignChild.GetParent()

        If TypeOf oCornerParent Is IJStructFeature Then
            Set oParentStructFeat = oCornerParent
            If oParentStructFeat.get_StructFeatureType = SF_WebCut Then
                Dim oCFWebCut As New StructDetailObjects.WebCut
                Set oCFWebCut.object = oCornerParent
                Set oBoundingPart = oCFWebCut.Bounding
            ElseIf oParentStructFeat.get_StructFeatureType = SF_FlangeCut Then
                Dim oCFFlangeCut As New StructDetailObjects.FlangeCut
                Set oCFFlangeCut.object = oCornerParent
                Set oBoundingPart = oCFFlangeCut.Bounding
            End If
        End If
    Case SF_WebCut
        sSectionProperty = "CornerRadius"
        'Get the Bounded Position and Bounding Part from the Web Cut
        Dim oWebCut As New StructDetailObjects.WebCut
        Set oWebCut.object = oStructFeature
        Set oFeaturePosition = oWebCut.BoundedLocation
        Set oBoundingPart = oWebCut.Bounding
    Case SF_FlangeCut
        sSectionProperty = "CornerRadius"
        'Get the BoundedPosition and Bounding Part from the Flange Cut
        Dim oFlangeCut As New StructDetailObjects.FlangeCut
        Set oFlangeCut.object = oStructFeature
        
        'Get the Web cut from the Flange cut
        Dim oFCWebCut As New StructDetailObjects.WebCut
        Set oFCWebCut.object = oFlangeCut.WebCut
        
        Set oFeaturePosition = oFCWebCut.BoundedLocation
        Set oBoundingPart = oFlangeCut.Bounding
    End Select

    ' ---------------------------------------------
    ' Determine the bounding fillet radius and axis
    ' ---------------------------------------------
    Dim sectionType As String
    Dim radius As Double
    Dim oWireBody As IJWireBody
    Dim oAxis As IJDVector
    Dim oCrossSection As IJCrossSection
    Dim oWireUtil As IJSGOWireBodyUtilities
    Set oWireUtil = New SGOWireBodyUtilities

    If (TypeOf oBoundingPart Is IJStiffener) Or (TypeOf oBoundingPart Is IJBeam) Then   'oPart is a profile or edge reinforcement
        If (TypeOf oBoundingPart Is IJStiffener) Then
            Dim oProfile As New StructDetailObjects.ProfilePart
            Set oProfile.object = oBoundingPart
            sectionType = oProfile.sectionType
            Select Case sectionType
            Case "EA", "UA", "I", "H", "T_XType", "C_SS"
                radius = oProfile.CrossSectionParameter(sSectionProperty)
            Case "ISType", "TSType", "CSType"
                ' ????
            Case Else
                ' ????
            End Select

        Else
            Dim oBeam As New StructDetailObjects.BeamPart
            Set oBeam.object = oBoundingPart
            sectionType = oBeam.sectionType
            Select Case sectionType
            Case "EA", "UA", "I", "H", "T_XType", "C_SS"
                radius = oBeam.CrossSectionParameter(sSectionProperty)
            Case "ISType", "TSType", "CSType"
                ' ????
            Case Else
                ' ????
            End Select
        End If

        Set oWireBody = GetProfilePartLandingCurve(oBoundingPart)
        oWireUtil.GetClosestPointOnWire oWireBody, oFeaturePosition, oClosestPoint, oAxis
    ElseIf TypeOf oBoundingPart Is ISPSMemberPartPrismatic Then

        Dim oMemberPartCommon As ISPSMemberPartCommon
        Dim oSPSCrossSection As ISPSCrossSection

        Set oMemberPartCommon = oBoundingPart
        Set oSPSCrossSection = oMemberPartCommon.CrossSection
        Set oCrossSection = oSPSCrossSection.definition

        Dim oMember As New StructDetailObjects.MemberPart
        Set oMember.object = oBoundingPart

        sectionType = oMember.sectionType

        Select Case sectionType
        Case "W", "M", "HP", "WT", "MT", "C", "MC", "L", "2L"
            Dim flangeThickness As Double
            Dim kdetail As Double
            flangeThickness = GetDoubleAttribute(oCrossSection, "tf")
            kdetail = GetDoubleAttribute(oCrossSection, "kdetail")
            radius = kdetail - flangeThickness
        Case "S", "ST"
            ' ????
        Case Else
            ' ????
        End Select

        Dim isCentered As Boolean
        Dim oThicknessDir As IJDVector
        oMember.LandingCurve oWireBody, oThicknessDir, isCentered

        oWireUtil.GetClosestPointOnWire oWireBody, oFeaturePosition, oClosestPoint, oAxis
    Else
        ' We're only expecting this feature to be the result of an endcut bounded by a profile.
        ' If necessary, the logic can be expanded later.  To guard against crashes, we'll default to
        ' a minor and major radius of 50mm
        radius = 0.05
    End If

    ' -----------------------
    ' Get the sketching plane
    ' -----------------------
    
    Select Case featureType
    Case SF_CornerFeature
        ' Get the face port geometry
        Dim oFacePort As IJPort
        Dim oFacePortGeom As IJSurfaceBody
        Set oFacePort = pPRL.InputObject("LogicalFace")    '(INPUT_PORT1FACE)
        Set oFacePortGeom = oFacePort.Geometry

        ' Get the nearest point on the actual surface
        Dim oModelBodyUtil As IJSGOModelBodyUtilities
        Set oModelBodyUtil = New SGOModelBodyUtilities
    
        Dim dist As Double
        oModelBodyUtil.GetClosestPointOnBody oFacePortGeom, oFeaturePosition, oClosestPoint, dist

        ' --------------------------------------------------
        ' Compute the angle between axis and sketching plane
        ' --------------------------------------------------
        ' Get the face port normal
        Dim oNormal As IJDVector
        oFacePortGeom.GetNormalFromPosition oClosestPoint, oNormal
    
        ' Compute the reference vector
        Dim oReferenceVector As IJDVector
        Set oReferenceVector = oAxis.Cross(oNormal)

        ' Compute angle
        Dim alpha As Double
    
        alpha = oAxis.Angle(oNormal, oReferenceVector)
    
        ' Adjust to 0-180
        Dim dPI As Double
        dPI = Atn(1#) * 4

        If alpha > dPI Then
            alpha = alpha - dPI
        End If
    
        'For some cases Cos(alpha)) becomes negative, so use absolute value
        GetMajorAxis = Abs(radius / Cos(alpha)) * 2
    Case SF_WebCut, SF_FlangeCut
        GetMajorAxis = 2 * radius
    End Select
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function
' =========================================================================
Public Function GetMinorAxis(ByVal pPRL As IJDParameterLogic) As Double
    Const METHOD = "::GetMinorAxis"

    On Error GoTo ErrorHandler
    Dim oStructFeature As IJStructFeature
    Dim oFeaturePosition As IJDPosition
    Dim oClosestPoint As IJDPosition
    Dim oBoundingPart As Object
    Dim sSectionProperty As String
    
    Set oStructFeature = pPRL.SmartOccurrence

    ' --------------------------------------------------------------------------------
    ' Determine the Struct Feature Type and get the feature position and bounding part
    ' --------------------------------------------------------------------------------
    Dim featureType As StructFeatureTypes
    featureType = oStructFeature.get_StructFeatureType
    Select Case featureType
    Case SF_CornerFeature
        sSectionProperty = "FilletRadius"
        'Get the Corner feature part object
        Dim oCornerFeature As New StructDetailObjects.CornerFeature
        Dim oCornerPart As Object
        
        Set oCornerFeature.object = oStructFeature
        Set oCornerPart = oCornerFeature.GetPartObject
        
        'Get the corner feature location
        oCornerFeature.GetLocationOfCornerFeature oFeaturePosition
        
        'Get the bounding object from the corner features parent cut
        Dim oDesignChild As IJDesignChild
        Set oDesignChild = pPRL.SmartOccurrence
        
        Dim oCornerParent As Object
        Dim oParentStructFeat As IJStructFeature
        Set oCornerParent = oDesignChild.GetParent()
        
        If TypeOf oCornerParent Is IJStructFeature Then
            Set oParentStructFeat = oCornerParent
            If oParentStructFeat.get_StructFeatureType = SF_WebCut Then
                Dim oCFWebCut As New StructDetailObjects.WebCut
                Set oCFWebCut.object = oCornerParent
                Set oBoundingPart = oCFWebCut.Bounding
            ElseIf oParentStructFeat.get_StructFeatureType = SF_FlangeCut Then
                Dim oCFFlangeCut As New StructDetailObjects.FlangeCut
                Set oCFFlangeCut.object = oCornerParent
                Set oBoundingPart = oCFFlangeCut.Bounding
            End If
        End If
    Case SF_WebCut
        sSectionProperty = "CornerRadius"
        'Get the Bounded Position and Bounding Part from the Web Cut
        Dim oWebCut As New StructDetailObjects.WebCut
        Set oWebCut.object = oStructFeature
        Set oFeaturePosition = oWebCut.BoundedLocation
        Set oBoundingPart = oWebCut.Bounding
    Case SF_FlangeCut
        sSectionProperty = "CornerRadius"
        'Get the BoundedPosition and Bounding Part from the Flange Cut
        Dim oFlangeCut As New StructDetailObjects.FlangeCut
        Set oFlangeCut.object = oStructFeature
        
        'Get the Web cut from the Flange cut
        Dim oFCWebCut As New StructDetailObjects.WebCut
        Set oFCWebCut.object = oFlangeCut.WebCut
        
        Set oFeaturePosition = oFCWebCut.BoundedLocation
        Set oBoundingPart = oFlangeCut.Bounding
    End Select

    ' ---------------------------------------------
    ' Determine the bounding fillet radius and axis
    ' ---------------------------------------------
    Dim sectionType As String
    Dim radius As Double
    Dim oWireBody As IJWireBody
    Dim oAxis As IJDVector
    Dim oCrossSection As IJCrossSection
    Dim oWireUtil As IJSGOWireBodyUtilities
    Set oWireUtil = New SGOWireBodyUtilities

    If (TypeOf oBoundingPart Is IJStiffener) Or (TypeOf oBoundingPart Is IJBeam) Then   'oPart is a profile or edge reinforcement
        If (TypeOf oBoundingPart Is IJStiffener) Then
            Dim oProfile As New StructDetailObjects.ProfilePart
            Set oProfile.object = oBoundingPart
            sectionType = oProfile.sectionType
            Select Case sectionType
            Case "EA", "UA", "I", "H", "T_XType", "C_SS"
                radius = oProfile.CrossSectionParameter(sSectionProperty)
            Case "ISType", "TSType", "CSType"
                ' ????
            Case Else
                ' ????
            End Select
        Else
            Dim oBeam As New StructDetailObjects.BeamPart
            Set oBeam.object = oBoundingPart
            sectionType = oBeam.sectionType
            Select Case sectionType
            Case "EA", "UA", "I", "H", "T_XType", "C_SS"
                radius = oBeam.CrossSectionParameter(sSectionProperty)
            Case "ISType", "TSType", "CSType"
                ' ????
            Case Else
                ' ????
            End Select
        End If
        
        Set oWireBody = GetProfilePartLandingCurve(oBoundingPart)
        oWireUtil.GetClosestPointOnWire oWireBody, oFeaturePosition, oClosestPoint, oAxis
    ElseIf TypeOf oBoundingPart Is ISPSMemberPartPrismatic Then

        Dim oMemberPartCommon As ISPSMemberPartCommon
        Dim oSPSCrossSection As ISPSCrossSection

        Set oMemberPartCommon = oBoundingPart
        Set oSPSCrossSection = oMemberPartCommon.CrossSection
        Set oCrossSection = oSPSCrossSection.definition

        Dim oMember As New StructDetailObjects.MemberPart
        Set oMember.object = oBoundingPart

        sectionType = oMember.sectionType

        Select Case sectionType
        Case "W", "M", "HP", "WT", "MT", "C", "MC", "L", "2L"
            Dim flangeThickness As Double
            Dim kdetail As Double
            flangeThickness = GetDoubleAttribute(oCrossSection, "tf")
            kdetail = GetDoubleAttribute(oCrossSection, "kdetail")
            radius = kdetail - flangeThickness
        Case "S", "ST"
            ' ????
        Case Else
            ' ????
        End Select

        Dim isCentered As Boolean
        Dim oThicknessDir As IJDVector
        oMember.LandingCurve oWireBody, oThicknessDir, isCentered

        oWireUtil.GetClosestPointOnWire oWireBody, oFeaturePosition, oClosestPoint, oAxis
    Else
        ' We're only expecting this feature to be the result of an endcut bounded by a profile.
        ' If necessary, the logic can be expanded later.  To guard against crashes, we'll default to
        ' a minor and major radius of 50mm
        radius = 0.05
    End If

    GetMinorAxis = 2 * radius

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

Public Function GetSelectorLogicForCustomMethod(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod) As IJDSelectorLogic
    
    ' ---------------------------------------------------------------------
    ' Get Symbol representation so the selector logic object can be created
    ' ---------------------------------------------------------------------
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition

    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Set GetSelectorLogicForCustomMethod = New SelectorLogic
    GetSelectorLogicForCustomMethod.Representation = oSymbolDefinition.IJDRepresentations(1)

End Function

Public Function GetResourceManagerFromObject(ByVal pObject As Object) As Object

    Const MT = "MarineLibraryCommon::GetResourceManagerFromObject"

    On Error GoTo ErrorHandler

    On Error Resume Next
    
    Dim oJDObject As IJDObject
    Dim oResourceMgr As IUnknown
    
    Set oJDObject = pObject
    If Not oJDObject Is Nothing Then
        Set oResourceMgr = oJDObject.ResourceManager
    
        If Not oResourceMgr Is Nothing Then
            Set GetResourceManagerFromObject = oResourceMgr
            Set oResourceMgr = Nothing
        End If
        
        Set oJDObject = Nothing
    End If
        
    Err.Clear
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, MT).Number
End Function

'***********************************************************************************************
'    Function      : CheckSmartItemExists
'
'    Description   : This method helps to determine if given SmartItem exists in Catalog.
'                    It helps to decide whether to select new smart item or use existing
'                    smart item based on availability of new item (incase user did not
'                    bulkload corresponding delta workbook, fall back on old item).
'
'    Parameters    :
'          Input    SmartClass Type, SmartClass Subtype, SmartClass name and SmartItem name
'
'    Return        : True if new SmartItem exists otherwise return False
'
'***********************************************************************************************
Public Function CheckSmartItemExists(eSCType As SmartClassType, eSCSubType As SmartClassSubType, _
                        sClassName As String, sItemName As String) As Boolean
    Const sMETHOD = "CheckSmartItemExists"
    On Error GoTo ErrorHandler
    
    CheckSmartItemExists = False 'Initialize
    
    Dim oCatalogQuery As IJSRDQuery
    Dim oSmartQuery As IJSmartQuery
    Dim oSmartItem As IJSmartItem
    
    Set oCatalogQuery = New SRDQuery
    Set oSmartQuery = oCatalogQuery
    
    ' Query for SmartItem... Check if exist
    On Error Resume Next
    Set oSmartItem = oSmartQuery.GetItem(eSCType, eSCSubType, sClassName, sItemName)
    On Error GoTo ErrorHandler
    
    If Not oSmartItem Is Nothing Then
        CheckSmartItemExists = True
    End If
    
    Set oSmartItem = Nothing
    Set oSmartQuery = Nothing
    Set oCatalogQuery = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
    
End Function

'***********************************************************************
' This method creates an instance of the Slot Mapping Rule
'***********************************************************************
Public Function CreateSlotMappingRuleSymbolInstance() As Object

Const METHOD = "CreateSlotMappingRuleSymbolInstance()"
On Error GoTo ErrorHandler

    Dim oCatalogQuery As IJSRDQuery
    Dim oRule As IJSRDRule
    Dim oRuleQuery As IJSRDRuleQuery
    Dim sSlotMappingRuleProgID As String
     
    Set oCatalogQuery = New SRDQuery
    Set oRuleQuery = oCatalogQuery.GetRulesQuery
    
    Dim sRuleName As String
    Dim oRuleUnk As Object
    
    sRuleName = "SDSlotMappingRule"
    
    ' Check if a Rule Has been bulkloaded into the Catalog
    On Error Resume Next
    
    Set oRuleUnk = oRuleQuery.GetRule(sRuleName)
    
    Err.Clear
    On Error GoTo ErrorHandler
    
    If Not oRuleUnk Is Nothing Then
       ' EndCut Mapping Rule was found
       Set oRule = oRuleUnk
       
       sSlotMappingRuleProgID = oRule.ProgId
    Else
        'Slot Mapping Rule Not Found
        Set CreateSlotMappingRuleSymbolInstance = Nothing
        Exit Function
    End If
    
    Dim strCodeBase As String
    'strCodeBase = Null
    
    Dim oCreateInstanceHelper As New CreateInstanceHelper
    Set CreateSlotMappingRuleSymbolInstance = oCreateInstanceHelper.CreateInstance(sSlotMappingRuleProgID, strCodeBase)
    
    Set oCreateInstanceHelper = Nothing
    Set oCatalogQuery = Nothing
    Set oRuleQuery = Nothing
    Set oRule = Nothing
    Set oRuleUnk = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

'------------------------------------------------------------------------------------------------------------
' METHOD:  GetChamferedPort
'
' DESCRIPTION:  Gets the Plate Port (Base or Offset) on which Chamfer has to be created.
'
' Inputs : pMemberDescription
' Output : oChamferedPort: on which chamfer has to be applied.
'------------------------------------------------------------------------------------------------------------

Public Sub GetChamferedPort(pMemberDescription As IJDMemberDescription, ByRef oChamferedPort As Object)

 On Error GoTo ErrorHandler
 Const sMETHOD = "GetChamferedPort"
 

    Dim oAppConnection  As IJAppConnection
    
    'Get Enum Ports
    Set oAppConnection = pMemberDescription.CAO

    Dim oPenetratingPort As IJPort
    Dim oPenetratedPort As IJPort
    GetPenetratedAndPenetratingPorts oAppConnection, oPenetratedPort, oPenetratingPort
 
    Dim oSlotmappingRule As IJSlotMappingRule
    Set oSlotmappingRule = CreateSlotMappingRuleSymbolInstance()
    
    Dim oBasePort As IJPort
    Dim oMappedPorts As JCmnShp_CollectionAlias
    Set oMappedPorts = New Collection
    
    'Get Mapped Ports
    oSlotmappingRule.GetEmulatedPorts oPenetratingPort.Connectable, oPenetratedPort.Connectable, oBasePort, oMappedPorts
    
    'Get Mapped Top Port
    Set oChamferedPort = oMappedPorts.Item(CStr(JXSEC_TOP))

  Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'------------------------------------------------------------------------------------------------------------
' METHOD:  GetThicknessDiffOfPlateOverMbrFlange
'
' DESCRIPTION:  Gets the Thickness Difference of Plate Part Over Member Flange.
'               +ve value indicates that Plate Part is thicker than Member Flange
'               -ve value indicates that Plate Part is thinner than Member Flange
' Inputs : oMemberPart, oPlatePart
' Output : dThicknessDifference: Thickness difference.
'------------------------------------------------------------------------------------------------------------
Public Sub GetThicknessDiffOfPlateOverMbrFlange(ByVal oMemberPart As Object, ByVal oPlatePart As Object, ByRef dThicknessDifference As Double)

    On Error GoTo ErrorHandler
    Const sMETHOD = "GetThicknessDiffOfPlateOverMbrFlange"
    
    Dim oSlotmappingRule As IJSlotMappingRule
    Set oSlotmappingRule = CreateSlotMappingRuleSymbolInstance()
    Dim oToleranceValue As Double
    oToleranceValue = 0.001
    Dim oBasePort As IJPort
    Dim oMappedPorts As JCmnShp_CollectionAlias
    Set oMappedPorts = New Collection
    
    'Get Mapped Ports
    oSlotmappingRule.GetEmulatedPorts oPlatePart, oMemberPart, oBasePort, oMappedPorts
    
    Dim oPlatePort As IJPort
    
    'Get Mapped Top Port
    Set oPlatePort = oMappedPorts.Item(CStr(JXSEC_TOP))
    Set oBasePort = oMappedPorts.Item(CStr(JXSEC_BOTTOM))

    Dim oModelBody As IJDModelBody
    Dim oTopPort As IJPort
    
    Set oTopPort = GetLateralSubPortBeforeTrim(oMemberPart, JXSEC_TOP)

    Dim oPlateNoraml As IJDVector
    Dim oTopPortNormal As IJDVector
    
    Dim oSDGeoUtil As GSCADStructGeomUtilities.PartInfo
    Set oSDGeoUtil = New GSCADStructGeomUtilities.PartInfo
    
    Dim bApprox As Boolean
    
    Set oPlateNoraml = oSDGeoUtil.GetPortNormal(oPlatePort, bApprox)
    Set oTopPortNormal = oSDGeoUtil.GetPortNormal(oTopPort, bApprox)
    
    'Check the plate Postion w.r.t Member Flange
    If (Abs(oPlateNoraml.Dot(oTopPortNormal))) <> 1 Then Exit Sub

    Dim bIntersectingTopFlange As Boolean
    Dim oSGOModelBodyUtils As GSCADShipGeomOps.SGOModelBodyUtilities
    Set oSGOModelBodyUtils = New GSCADShipGeomOps.SGOModelBodyUtilities
    
    Dim bIntersectingBottomFlange As Boolean
    Dim oBottomPort As IJPort
    
    'Check if Plate is on Top Flange
    bIntersectingTopFlange = oSGOModelBodyUtils.HasIntersectingGeometry(oTopPort, oPlatePart)

       Dim oMemberBody As IJDModelBody
       Dim Position1 As IJDPosition
       Dim Position2 As IJDPosition
       Dim MbrPlateOfsetDist As Double
       Dim MbrPlateBaseDist As Double
       
       'if bIntersectingTopFlange is False, then checks for the tolerance value. If the distance between the plate and member is less than tolerance value,
       'make TopFlangeResultantIntersection as true
    If Not bIntersectingTopFlange Then
        Set oMemberBody = oTopPort.Geometry
        oMemberBody.GetMinimumDistance oPlatePort.Geometry, Position1, Position2, MbrPlateOfsetDist
        oMemberBody.GetMinimumDistance oBasePort.Geometry, Position1, Position2, MbrPlateBaseDist
        If (MbrPlateBaseDist < oToleranceValue) <> (MbrPlateOfsetDist < oToleranceValue) Then
            bIntersectingTopFlange = True
        End If
    End If
   
    If bIntersectingTopFlange Then
        Set oModelBody = GetLateralSubPortBeforeTrim(oMemberPart, JXSEC_TOP_FLANGE_RIGHT_BOTTOM).Geometry
    Else
        Set oBottomPort = GetLateralSubPortBeforeTrim(oMemberPart, JXSEC_BOTTOM)
        
        'Check if Plate is on Bottom Flange
        bIntersectingBottomFlange = oSGOModelBodyUtils.HasIntersectingGeometry(oBottomPort, oPlatePart)
        
       'if bIntersectingBottomFlange is False, then checks for the tolerance value. If the distance between the plate and member is less than tolerance value,
       'make TopFlangeResultantIntersection as true
       
        If Not bIntersectingBottomFlange Then
            Set oMemberBody = oBottomPort.Geometry
            oMemberBody.GetMinimumDistance oPlatePort.Geometry, Position1, Position2, MbrPlateOfsetDist
            oMemberBody.GetMinimumDistance oBasePort.Geometry, Position1, Position2, MbrPlateBaseDist

            If ((MbrPlateBaseDist < oToleranceValue) <> (MbrPlateOfsetDist < oToleranceValue)) Then
                bIntersectingBottomFlange = True
            End If
        End If
        If bIntersectingBottomFlange Then
            Set oModelBody = GetLateralSubPortBeforeTrim(oMemberPart, JXSEC_BOTTOM_FLANGE_RIGHT_TOP).Geometry
        Else
            Exit Sub
        End If
    End If
    
    Dim oPosition1 As IJDPosition
    Dim oPosition2 As IJDPosition
    
    'Get the distance between Plate and Member Flange
    oModelBody.GetMinimumDistance oPlatePort.Geometry, oPosition1, oPosition2, dThicknessDifference
    
    Dim oSD_Member As New StructDetailObjects.MemberPart
    Set oSD_Member.object = oMemberPart
    
    Dim oSD_PlatePart As New StructDetailObjects.PlatePart
    Set oSD_PlatePart.object = oPlatePart
    
    If oSD_Member.flangeThickness > oSD_PlatePart.PlateThickness Then
        dThicknessDifference = -dThicknessDifference
    End If
 Exit Sub
 
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub
'

'------------------------------------------------------------------------------------------------------------
' METHOD:  GetLateralSubPortBeforeTrim
'
' DESCRIPTION:  Gets the Lateral Sub Port of the Member Part(Standard(Rolled) Members) Before Trim/Cut
'               and returns it
'               Method can be enhanced for ProfileParts/Built Up's(Designed Members)
'
' Inputs : oMemberPart : Member Part on which the port exist
'          eSubPort : Enum of the port on the Member part which is to be retrieved(for e.g WEB_LEFT or WEB_RIGHT etc)
'
' Output : A Lateral SubPortBeforeTrim is returned(depending on eSubPort passed, if exists)
'          Or Nothing (if eSubPort passed doesnt exist on member part)
'------------------------------------------------------------------------------------------------------------
Public Function GetLateralSubPortBeforeTrim(oMemberPart As Object, ByVal eSubPort As JXSEC_CODE) As IJPort

    Const METHOD = "GetLateralSubPortBeforeTrim"

    Dim sMsg As String

    Dim lCtxId As Long
    Dim lOprId As Long
    Dim lOptId As Long

    Dim ePortType As JS_TOPOLOGY_PROXY_TYPE
    Dim eFilterType As JS_TOPOLOGY_FILTER_TYPE

    Dim oPort As IJPort
    Dim oPortObject As Object
    Dim oPortElements As IJElements
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Dim oPlatePart As New StructDetailObjects.PlatePart

    Dim oStructGraphConnectable As IJStructGraphConnectable

    Dim oMemberFactory As SPSMembers.SPSMemberFactory
    Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices

    On Error GoTo ErrorHandler

    Set GetLateralSubPortBeforeTrim = Nothing

    If TypeOf oMemberPart Is ISPSMemberPartPrismatic Then 'Standard (rolled) member

            eFilterType = JS_TOPOLOGY_FILTER_LCONNECT_PRT_SUB_LFACES

        ' Verify current Member Part
        If TypeOf oMemberPart Is IJStructGraphConnectable Then
            Set oStructGraphConnectable = oMemberPart
        Else
            sMsg = "Else ... Typeof(MemberPart):" & TypeName(oMemberPart)
            GoTo ErrorHandler
        End If

        ' Retreive list of Ports from the SPS Member Part's Solid Geometry before Member Cut Operation
        oStructGraphConnectable.enumPortsInGraphByTopologyFilter oPortElements, _
                                                                 eFilterType, _
                                                                 StableGeometry, _
                                                                 vbNull
        ' Verify returned List of Port(s) is valid
        If oPortElements Is Nothing Then
            sMsg = "oPortElements Is Nothing"
        ElseIf oPortElements.Count < 1 Then
            sMsg = "oPortElements.Count < 1"
        Else

            For Each oPortObject In oPortElements

                Set oMemberFactory = New SPSMembers.SPSMemberFactory
                Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices

                oMemberConnectionServices.GetStructPortInfo oPortObject, ePortType, _
                                                                                     lCtxId, lOptId, lOprId

                If TypeOf oPortObject Is IJPort Then
                    Set oPort = oPortObject
                    If lOprId = eSubPort Then
                        Set GetLateralSubPortBeforeTrim = oPort
                        Exit For
                    End If

                Else
                    sMsg = "Else... TypeOf oPortObject Is IJPort"
                    GoTo ErrorHandler
                End If

             Next oPortObject

        End If

        Set oPort = Nothing
        Set oPortObject = Nothing
        Set oPortElements = Nothing
        Set oStructGraphConnectable = Nothing
        Set oMemberFactory = Nothing
        Set oMemberConnectionServices = Nothing

    ElseIf TypeOf oMemberPart Is IJProfile Then
        Set oProfilePart.object = oMemberPart
        Set GetLateralSubPortBeforeTrim = oProfilePart.SubPortBeforeTrim(eSubPort)

    ElseIf TypeOf oMemberPart Is IJPlate Then
    'Designed (built-up) member
'        sMsg = METHOD & " not available for built-up members"
'        GoTo ErrorHandler

        Set oPlatePart.object = oMemberPart
        Set GetLateralSubPortBeforeTrim = oPlatePart.BasePortFromOperation(BPT_Lateral, "CreatePlatePart.GeneratePlatePart_AE.1", False)
    End If

    Set oProfilePart = Nothing
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function


' ********************************************************************************
' Method: GetSmartParent
'
' Abstract:
'   Gets Smart Occurrence / Symbol available Selection Rule SmartClass(es)
'
' Inputs:
'
' Outputs:
' ********************************************************************************
Public Sub GetSmartOccurrenceParent(oOccurrenceObject As Object, _
                                    oParentObject As Object)
Const METHOD = "::GetSmartOccurrenceParent"
    On Error GoTo ErrorHandler

    Dim nCount As Long
    Dim oColObject As Object
    Dim oRelationShip As IJDRelationship
    Dim oAssocRelation As IJDAssocRelation
    Dim oRelationShipCol As IJDRelationshipCol
    
    Set oParentObject = Nothing
    If Not TypeOf oOccurrenceObject Is IJSmartOccurrence Then
        Exit Sub
    End If
                
    ' Need to check if given SmartOccurrence
    ' is a member of a CustomAssembly relationship
    ' it if is:
    '   then get the Parent of the SmartOccurrence/SmartItem
    '   to continue up the list
    If TypeOf oOccurrenceObject Is IJDAssocRelation Then
        Set oAssocRelation = oOccurrenceObject
        Set oColObject = oAssocRelation.CollectionRelations("IJFullObject", _
                                                            "toAssembly")
        If oColObject Is Nothing Then
        
        ElseIf TypeOf oColObject Is IJDRelationshipCol Then
            Set oRelationShipCol = oColObject
            nCount = oRelationShipCol.Count
                          
            If nCount > 0 Then
                Set oRelationShip = oRelationShipCol.Item(1)
                Set oParentObject = oRelationShip.Target
            End If
        
        End If
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'***********************************************************************
' METHOD:  InitializeStructFeatureProperties
'
' ARGUMENTS:    oStructFeature As Object
'                   Struct Feature whose properties are being initialized
'
' DESCRIPTION:  Initializes passed struct feature with following properties:
'               NAMING RULE           First available naming rule from
'                                     Molded Forms naming rules
'***********************************************************************

Public Sub InitializeStructFeatureProperties(oStructFeature As Object)

    'Get the default generic naming solver and use it to name the feature
    Dim oNamingUtils2 As IJNamingUtils2
    Set oNamingUtils2 = New GSCADCreateModifyUtilities.StructEntityUtils
    
    Dim oNamingObject As IJDStructEntityNaming
    
    Set oNamingObject = oStructFeature
    oNamingObject.NamingRule = oNamingUtils2.GetDefaultGenericNamingRule()
                
    Set oNamingObject = Nothing
    Set oNamingUtils2 = Nothing
    
End Sub



'*************************************************************************
'Function
'HandleError
'
'Abstract
' called by other subs and fuctions during error. This method logs the error
' and returns success
'
'input
'Module(file) name, method name
'
'Return
'success
'
'Exceptions
'
'***************************************************************************
Public Sub HandleError(sModule As String, sMETHOD As String, Optional sExtraInfo As String = "")
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.AddFromErr Err, sExtraInfo, sMETHOD, sModule
    End If
    Set oEditErrors = Nothing
End Sub


'Function
'Parent_SmartItemName
'
'Description:
'   Given the Smart Occurent Object
'    return the Owning Smart item Name
'
'input
'   oOccurrenceObject
'
'Return
'   sParentItemName
'   oParentObject
'
'Exceptions
'
'***************************************************************************
Public Sub Parent_SmartItemName(oOccurrenceObject As Object, _
                                sParentItemName As String, _
                                Optional oParentObject As Object = Nothing)
    Const METHOD = "MbrEndCutUtilities::Parent_SmartItemName"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim nCount As Long
    Dim oColObject As Object
    Dim oRelationShip As IJDRelationship
    Dim oAssocRelation As IJDAssocRelation
    Dim oRelationShipCol As IJDRelationshipCol
    Dim oParent_SmartOccurrence As IJSmartOccurrence
    
    sParentItemName = ""
    If Not TypeOf oOccurrenceObject Is IJSmartOccurrence Then
        Exit Sub
    End If
                
    ' Need to check if given SmartOccurrence
    ' is a member of a CustomAssembly relationship
    If TypeOf oOccurrenceObject Is IJDAssocRelation Then
        Set oAssocRelation = oOccurrenceObject
        Set oColObject = oAssocRelation.CollectionRelations("IJFullObject", _
                                                            "toAssembly")
        If oColObject Is Nothing Then
        
        ElseIf TypeOf oColObject Is IJDRelationshipCol Then
            Set oRelationShipCol = oColObject
            nCount = oRelationShipCol.Count
                          
            If nCount > 0 Then
                Set oRelationShip = oRelationShipCol.Item(1)
                Set oParentObject = oRelationShip.Target
            End If
        
        End If
    End If
    
    If oParentObject Is Nothing Then
        If TypeOf oOccurrenceObject Is IJDesignChild Then
            Dim oDesignChild As IJDesignChild
            Set oDesignChild = oOccurrenceObject
                        
            If Not oDesignChild.GetParent Is Nothing Then
                If TypeOf oDesignChild.GetParent Is IJAppConnection Then
                    Set oParentObject = oDesignChild.GetParent
                End If
            End If
        End If
    End If
    
    If Not oParentObject Is Nothing Then
        If TypeOf oParentObject Is IJSmartOccurrence Then
            Set oParent_SmartOccurrence = oParentObject
            sParentItemName = oParent_SmartOccurrence.Item
        End If
    Else
        Set oParent_SmartOccurrence = oOccurrenceObject
        sParentItemName = oParent_SmartOccurrence.Item
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*************************************************************************
'Function: 'Set_AttributeValue
'
'Description: Method sets the value to the given attibute only when
'             the Interface and attribute exists on the object
'
'***************************************************************************
Public Sub Set_AttributeValue(oObject As Object, sInterfaceName As String, sAttributeName As String, vValue As Variant)

  Const METHOD = "MarineLibraryCommon::Set_AttributeValue"
  On Error GoTo ErrorHandler

    Dim sMsg As String
    Dim vInterfaceID As Variant
    Dim nAttributes As Long
    Dim iIndex As Long
    Dim oAttributes As IJDAttributes
    Dim AttributeData() As CustomAttributeData
    Dim oAttributeColl As CollectionProxy
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oInterfaceAttributeCollection As IJDInfosCol

    Set oAttributes = oObject

    ' Verify that the Attributes contain at least one
    If oAttributes.Count > 0 Then

        For Each vInterfaceID In oAttributes
            Set oAttributeColl = oAttributes.CollectionOfAttributes(vInterfaceID)

            ' verify the current interface Collection is valid
            If Not oAttributeColl Is Nothing Then

                Set oInterfaceInfo = oAttributeColl.InterfaceInfo
                ' veirfy that the current Attribute Interface collection
                ' represents a "User" Attribute interface not System Attribute(??)
                If Not oInterfaceInfo.IsHardCoded Then

                    If StrComp(Trim$(oInterfaceInfo.UserName), Trim$(sInterfaceName), vbTextCompare) = 0 Then

                    ' retrieve collection of Attributes from the InterfaceInfo object
                    ' There appears to be two(2) types of collections
                    ' AttributeCollection is a collection of 'COM' attributes
                        Set oInterfaceAttributeCollection = oInterfaceInfo.AttributeCollection

                        If Not oInterfaceAttributeCollection Is Nothing Then
                            Get_AttributesFromCollection oInterfaceAttributeCollection, _
                                                         oAttributeColl, _
                                                         nAttributes, _
                                                         AttributeData
                            Set oInterfaceAttributeCollection = Nothing
                        End If

                        'Get the attribute collection and checkif the inputted attributes exists in the collection
                        'Set the given value to the attribute
                        For iIndex = 1 To nAttributes
                            If StrComp(Trim$(sAttributeName), Trim(AttributeData(iIndex).Name), vbTextCompare) = 0 Then
                                oAttributes.CollectionOfAttributes(sInterfaceName).Item(sAttributeName).Value = vValue
                                Exit Sub
                            End If
                        Next iIndex
                    End If
                End If
            End If
        Next
    End If

  Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub


Public Function Has_Attribute(ByVal oSmartOccurance As IJSmartOccurrence, ByVal sAttributeName As String) As Boolean

  Const METHOD = "MarineLibraryCommon::Has_Attribute"
  On Error GoTo ErrorHandler

    Dim sMsg As String
    Dim vInterfaceID As Variant
    Dim nAttributes As Long
    Dim iIndex As Long
    Dim oAttributes As IJDAttributes
    Dim AttributeData() As CustomAttributeData
    Dim oAttributeColl As CollectionProxy
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oInterfaceAttributeCollection As IJDInfosCol
    
    Set oAttributes = oSmartOccurance
    Has_Attribute = False
    
    ' Verify that the Attributes contain at least one
    If oAttributes.Count > 0 Then
        
        For Each vInterfaceID In oAttributes
            Set oAttributeColl = oAttributes.CollectionOfAttributes(vInterfaceID)
            
            ' verify the current interface Collection is valid
            If Not oAttributeColl Is Nothing Then
                
                Set oInterfaceInfo = oAttributeColl.InterfaceInfo
                ' veirfy that the current Attribute Interface collection
                ' represents a "User" Attribute interface not System Attribute(??)
                If Not oInterfaceInfo.IsHardCoded Then

                ' retrieve collection of Attributes from the InterfaceInfo object
                ' There appears to be two(2) types of collections
                ' AttributeCollection is a collection of 'COM' attributes
                    Set oInterfaceAttributeCollection = oInterfaceInfo.AttributeCollection
                    
                    If Not oInterfaceAttributeCollection Is Nothing Then
                        Get_AttributesFromCollection oInterfaceAttributeCollection, oAttributeColl, _
                                                     nAttributes, AttributeData
                        Set oInterfaceAttributeCollection = Nothing
                    End If
                    
                    'Get the attribute collection and checkif the inputted attributes exists in the collection
                    'Set the given value to the attribute
                    For iIndex = 1 To nAttributes
                        If StrComp(Trim$(sAttributeName), Trim(AttributeData(iIndex).Name), vbTextCompare) = 0 Then
                            Has_Attribute = True
                            Exit Function
                        End If
                    Next iIndex
                End If
            End If
        Next
    End If
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function


Public Sub GetPenetratedAndPenetratingPorts(ByVal oAppConnection As IJAppConnection, oPenetratedPort As IJPort, oPenetratingPort As IJPort)

    Dim oElements As IJElements
    oAppConnection.enumPorts oElements

    Set oPenetratingPort = oElements.Item(1)

    If TypeOf oPenetratingPort.Connectable Is IJPlate Then
        Set oPenetratedPort = oElements.Item(2)
    Else
        Set oPenetratingPort = oElements.Item(2)
        Set oPenetratedPort = oElements.Item(1)
    End If

End Sub

'**********************************************************************************************
' Method      : ForceUpdateOnMemberObjects
' Description : Method to forceupdate on input object using IJDMemberObjects
'**********************************************************************************************

Public Sub ForceUpdateOnMemberObjects(oObject As Object)

On Error GoTo ErrorHandler
Const sMETHOD = "ForceUpdateOnMemberObjects"
    
    Dim oStructAssocTools As SP3DStructGenericTools.StructAssocTools
    Set oStructAssocTools = New SP3DStructGenericTools.StructAssocTools
    
    On Error Resume Next
                                   
    oStructAssocTools.UpdateObject oObject, _
                                   "{9FC1AC01-9684-4E11-ABB8-6BDC3F636FE7}" 'IJDMemberObjects
    Err.Clear
    On Error GoTo ErrorHandler

    Set oStructAssocTools = Nothing

    Exit Sub
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Sub

Public Function GetCustomAttribute(oObject As Object, strInterfaceName As String, strAttributeName As String) As Variant

    Const METHOD = "MarineLibraryCommon::GetCustomAttribute"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    If Not TypeOf oObject Is IJDAttributes Then
        Exit Function
    End If
    
    ' ------------------
    ' Get the collection
    ' ------------------
    Dim oAttributes As IJDAttributes
    Set oAttributes = oObject
    
    Dim oAttributesCol As IJDAttributesCol
    On Error Resume Next
    Set oAttributesCol = oAttributes.CollectionOfAttributes(strInterfaceName)
    Err.Clear
    On Error GoTo ErrorHandler
    
    If oAttributesCol Is Nothing Then
        Exit Function
    End If
    
    ' -----------------
    ' Get the attribute
    ' -----------------
    Dim oAttribute As IJDAttribute
    Set oAttribute = oAttributesCol.Item(strAttributeName)
    
    ' -----------------------------------------------
    ' If not missing or empty, retrieve the attribute
    ' -----------------------------------------------
    If oAttribute Is Nothing Then
        Exit Function
    End If
    
    GetCustomAttribute = oAttribute.Value
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

'**********************************************************************************************
' Method      : GetWidthFromStructDetailObjects
' Description : Method to get the webthickness value as distance from web right to left from structdetail objects method.
'**********************************************************************************************
Public Function GetWidthFromStructDetailObjects(oObject As IJDObject) As Double

 Const sMETHOD = "GetWidthFromStructDetailObjects"
 On Error GoTo ErrHandler
 
    Dim oCrossSection As IJCrossSection
    Dim oMemberPartCommon As ISPSMemberPartCommon
    Dim oSPSCrossSection As ISPSCrossSection
    Set oMemberPartCommon = oObject
    Set oSPSCrossSection = oMemberPartCommon.CrossSection
    Set oCrossSection = oSPSCrossSection.definition
    GetWidthFromStructDetailObjects = GetDoubleAttribute(oCrossSection, "Width")

Exit Function

ErrHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

'**********************************************************
' Method Name   : AddFeatureEndCutData
' Inputs        : Feature Occurence Object, End Cut Relative Position, End Cut Data String
' Description   : This method prepares the endcutdata of a given feature in string format as prescribed in the
'                 format shown below;
' End Cut Data-
' String Format : pwfc;pwtc;pwbc|tffc;tflc;tfrc|bffc;bflc;bfrc|swfc;swtc;swbc, where
'pwfc = primary web first cut
'pwtc = primary web top cut
'pwbc = primary web bottom cut
'tffc = top flange first cut
'tflc = top flange left cut
'tfrc = top flange right cut
'bffc = bottom flange first cut
'bflc = bottom flange left cut
'bfrc = bottom flange right cut
'swfc = secondary web first cut
'swtc = secondary web top cut
'swbc = secondary web bottom cut
'*********************************************************

Public Sub AddFeatureEndCutData(oFeatureObj As Object, lEndCutRelativePosition As Long, sEndCutData As String)

    Const sMETHOD = "AddFeatureEndCutData"
    On Error GoTo ErrorHandler
    
    Dim oFeature As IJStructFeature
    Set oFeature = oFeatureObj
    Dim oSDEndCutData As IJSDEndCutData
    Dim oBoundedPort As IJPort
        
    If oFeature.get_StructFeatureType = SF_WebCut Then
        Dim oSDWebCut As New StructDetailObjects.WebCut
        Set oSDWebCut.object = oFeature
        Set oBoundedPort = oSDWebCut.BoundedPort
    ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
        Dim oSDFlangeCut As New StructDetailObjects.FlangeCut
        Set oSDFlangeCut.object = oFeature
        Set oBoundedPort = oSDFlangeCut.BoundedPort
    End If

    If Not TypeOf oBoundedPort.Connectable Is IJProfilePart Then Exit Sub
    
    Dim sCodelistedDrawingType As String
    
    sCodelistedDrawingType = CodelistedDrawingType(oFeatureObj, sEndCutData)
    
    Dim oStructPort As IJStructPort
    Set oStructPort = oBoundedPort
    Dim sPortType As String
    sPortType = Get_PortFaceType(oStructPort)
    Dim sStartEndCutData As String
    Dim sEndEndCutData As String
    Set oSDEndCutData = oBoundedPort.Connectable
    sStartEndCutData = oSDEndCutData.StartEndCutData
    sEndEndCutData = oSDEndCutData.EndEndCutData
    
    Dim sTempString() As String
    
    If sStartEndCutData = vbNullString Then
        sStartEndCutData = " ; ; | ; ; | ; ; | ; ; "
    End If
    
    If sEndEndCutData = vbNullString Then
        sEndEndCutData = " ; ; | ; ; | ; ; | ; ; "
    End If
    
    Dim lFirstDelimiterPos As Long
    Dim lSecDelimiterPos As Long
    
    
    If sPortType = "Base" Then
        sTempString = Split(sStartEndCutData, "|")
    ElseIf sPortType = "Offset" Then
        sTempString = Split(sEndEndCutData, "|")
    Else
        ' not needed
        Exit Sub
    End If
    
    
    Dim sBottomFlange As String
    Dim sFlangeStr As String
    
    If oFeature.get_StructFeatureType = SF_WebCut Then
    
        lFirstDelimiterPos = InStr(1, sTempString(0), ";")
        lSecDelimiterPos = InStrRev(sTempString(0), ";")
        
        Dim oWebString As String
        
        Select Case lEndCutRelativePosition
            Case Primary
                 Mid$(sTempString(0), lFirstDelimiterPos - 1) = sCodelistedDrawingType
            Case TopOrLeft
                Mid$(sTempString(0), lSecDelimiterPos - 1) = sCodelistedDrawingType
            Case BottomOrRight
                Mid$(sTempString(0), lSecDelimiterPos + 1) = sCodelistedDrawingType
        End Select

    ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
        
        Dim oACObject As Object
        GetSmartOccurrenceParent oFeatureObj, oACObject
        ' the reason for checking the selector answer on both TheBottomFlange and BottomFlange is
        ' TheBottomFlange is the selector question defined on all the flange cuts selectors in SMEndcutRules and whereas
        ' BottomFlange is the question defined on all the flange cuts selectors in SMMbrEndcutRules
  
        
        GetSelectorAnswer oFeatureObj, "TheBottomFlange", sBottomFlange
        If sBottomFlange = vbNullString Then
            GetSelectorAnswer oFeatureObj, "BottomFlange", sBottomFlange
        End If
        
        If sBottomFlange = "No" Then
            ' get the top flange string array
            sFlangeStr = sTempString(1)
        Else
            ' get the bottom flange string array
            sFlangeStr = sTempString(2)
        End If
        
        lFirstDelimiterPos = InStr(1, sFlangeStr, ";", vbTextCompare)
        lSecDelimiterPos = InStrRev(sFlangeStr, ";", , vbTextCompare)
        
        Select Case lEndCutRelativePosition
            Case Primary
                 Mid$(sFlangeStr, lFirstDelimiterPos - 1) = sCodelistedDrawingType
            Case TopOrLeft
                Mid$(sFlangeStr, lSecDelimiterPos - 1) = sCodelistedDrawingType
            Case BottomOrRight
                Mid$(sFlangeStr, lSecDelimiterPos + 1) = sCodelistedDrawingType
        End Select
        
        If sBottomFlange = "No" Then
            ' set the top flange string array
             sTempString(1) = sFlangeStr
        Else
            ' set the bottom flange string array
            sTempString(2) = sFlangeStr
        End If
    Else
        ' not supported
    End If
        
    If sPortType = "Base" Then
        oSDEndCutData.StartEndCutData = Join(sTempString, "|")
    ElseIf sPortType = "Offset" Then
        oSDEndCutData.EndEndCutData = Join(sTempString, "|")
    End If
    
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Function CodelistedDrawingType(oFeatureObj As IJStructFeature, EndCutData As String) As String
    
    Const sMETHOD = "CodelistedDrawingType"
    On Error GoTo ErrorHandler
    Dim oCodeListHelper As IMSAttributes.IJDCodeListMetaData
    Set oCodeListHelper = oFeatureObj
    
    If oFeatureObj.get_StructFeatureType = SF_WebCut Then
        Select Case EndCutData
            Case "0"
                CodelistedDrawingType = oCodeListHelper.ValueIDByShortString("WebCutDrawingTypeCodeList", gsStraightNoOffsetWebCuts)
            Case "1"
                CodelistedDrawingType = oCodeListHelper.ValueIDByShortString("WebCutDrawingTypeCodeList", gsStraightOffsetWebCuts)
            Case "2"
                CodelistedDrawingType = oCodeListHelper.ValueIDByShortString("WebCutDrawingTypeCodeList", gsSnipedNoOffsetWebCuts)
            Case "3"
                CodelistedDrawingType = oCodeListHelper.ValueIDByShortString("WebCutDrawingTypeCodeList", gsSnipedOffsetWebCuts)
        End Select
    ElseIf oFeatureObj.get_StructFeatureType = SF_FlangeCut Then
        Select Case EndCutData
            Case "0"
                CodelistedDrawingType = oCodeListHelper.ValueIDByShortString("FlangeCutDrawingTypeCodeList", gsStraightNoOffsetFlangeCuts)
            Case "1"
                CodelistedDrawingType = oCodeListHelper.ValueIDByShortString("FlangeCutDrawingTypeCodeList", gsStraightOffsetFlangeCuts)
            Case "2"
                CodelistedDrawingType = oCodeListHelper.ValueIDByShortString("FlangeCutDrawingTypeCodeList", gsSnipedNoOffsetFlangeCuts)
            Case "3"
                CodelistedDrawingType = oCodeListHelper.ValueIDByShortString("FlangeCutDrawingTypeCodeList", gsSnipedOffsetFlangeCuts)
        End Select
    End If
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function


'***********************************************************************
' METHOD:  Get_PortFaceType
'
' DESCRIPTION:
'       Given an IJStructPort
'       Determine if the Port is 'Base', 'Offset', 'WebLeft', etc.
'
'   inputs:
'       oPort As IJPort
'
'   outputs:
'       GetPortBaseOffsetSide as String
'           "Base"      given Port is "Base" port
'           "Offset"    given Port is "Offset" port
'           "Lateral"   given Port is "Lateral" port
'           "TOP"       given Port is Profile JXSEC_TOP port
'           "BOTTOM"    given Port is Profile JXSEC_BOTTOM port
'           "WEB_LEFT"  given Port is Profile JXSEC_WEB_LEFT port
'           "WEB_RIGHT" given Port is Profile JXSEC_WEB_RIGHT port
'
'***********************************************************************
Public Function Get_PortFaceType(oPortObject As Object) As String
    Const METHOD = "Get_PortFaceType"
    On Error GoTo ErrorHandler

    Dim sMsg As String
    
    Dim lBaseCheck As Long
    Dim lNplusCheck As Long
    Dim lNminusCheck As Long
    Dim lOffsetCheck As Long
    Dim lLateralCheck As Long
    
    Dim lContextID As Long
    Dim lOperatorID As Long
    Dim lOperatationID As Long
    
    Dim ePortType As JS_TOPOLOGY_PROXY_TYPE
    
    Dim sPortSide As String
    Dim eContextid As eUSER_CTX_FLAGS
    
    Dim oPort As IJPort
    Dim oConnectable As IJConnectable
    Dim oStructPort As IJStructPort
    
    Dim eAxisPortIndex As SPSMemberAxisPortIndex
    Dim oSplitAxisPort As ISPSSplitAxisPort
    
    Dim oMemberFactory As SPSMembers.SPSMemberFactory
    Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
    
    If TypeOf oPortObject Is IJPort Then
        Set oPort = oPortObject
        Set oConnectable = oPort.Connectable
    End If
    
    
    ' SP3D Member Object Ports do not implement IJStructPort interface
    sPortSide = ""
    If oConnectable Is Nothing Then
    
    ElseIf TypeOf oPortObject Is ISPSSplitAxisPort Then

        ' SP3D Member Object Axis Port
        Set oSplitAxisPort = oPortObject
        eAxisPortIndex = oSplitAxisPort.PortIndex
    
        If eAxisPortIndex = SPSMemberAxisStart Then
            sPortSide = C_Port_Base
        ElseIf eAxisPortIndex = SPSMemberAxisEnd Then
            sPortSide = C_Port_Offset
        ElseIf eAxisPortIndex = SPSMemberAxisAlong Then
            sPortSide = C_Port_Lateral
        Else
            sPortSide = ""
        End If
        
    ElseIf TypeOf oConnectable Is ISPSMemberPartCommon Then

        ' SP3D Member Object Solid Port
        Set oMemberFactory = New SPSMembers.SPSMemberFactory
        Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices
        
        oMemberConnectionServices.GetStructPortInfo oPortObject, ePortType, _
                                                    lContextID, lOperatationID, lOperatorID
        If lOperatorID = JXSEC_TOP Then
            sPortSide = C_Port_Top
        ElseIf lOperatorID = JXSEC_BOTTOM Then
            sPortSide = C_Port_Bottom
        ElseIf lOperatorID = JXSEC_WEB_LEFT Then
            sPortSide = C_Port_WebLeft
        ElseIf lOperatorID = JXSEC_WEB_RIGHT Then
            sPortSide = C_Port_WebRight
        Else
            lBaseCheck = eContextid And CTX_BASE
            lNplusCheck = eContextid And CTX_NPLUS
            lNminusCheck = eContextid And CTX_NMINUS
            lOffsetCheck = eContextid And CTX_OFFSET
            lLateralCheck = eContextid And CTX_LATERAL
        
            If lBaseCheck <> 0 Then
                sPortSide = C_Port_Base
                
            ElseIf lOffsetCheck <> 0 Then
                sPortSide = C_Port_Offset
                    
            ElseIf lNplusCheck <> 0 Then
                sPortSide = C_Port_Base
                    
            ElseIf lNminusCheck <> 0 Then
                sPortSide = C_Port_Offset
            
            ElseIf lLateralCheck <> 0 Then
                sPortSide = C_Port_Lateral
            Else
                sPortSide = C_Port_Lateral
            End If
        End If
        
    ElseIf TypeOf oPortObject Is IJStructPort Then
        Set oStructPort = oPortObject
        eContextid = oStructPort.ContextID
        lBaseCheck = eContextid And CTX_BASE
        lNplusCheck = eContextid And CTX_NPLUS
        lNminusCheck = eContextid And CTX_NMINUS
        lOffsetCheck = eContextid And CTX_OFFSET
        lLateralCheck = eContextid And CTX_LATERAL
        
        If lBaseCheck <> 0 Then
            sPortSide = C_Port_Base
            
        ElseIf lOffsetCheck <> 0 Then
            sPortSide = C_Port_Offset
                
        ElseIf lNplusCheck <> 0 Then
            sPortSide = C_Port_Base
                
        ElseIf lNminusCheck <> 0 Then
            sPortSide = C_Port_Offset
        
        Else
            ' Connecting Port is not Base or Offset port
            ' check if Profile Top/Bottom/Left/Right port
            lOperatorID = oStructPort.OperatorID
            If lOperatorID = JXSEC_TOP Then
                sPortSide = C_Port_Top
            ElseIf lOperatorID = JXSEC_BOTTOM Then
                sPortSide = C_Port_Bottom
            ElseIf lOperatorID = JXSEC_WEB_LEFT Then
                sPortSide = C_Port_WebLeft
            ElseIf lOperatorID = JXSEC_WEB_RIGHT Then
                sPortSide = C_Port_WebRight
            
            ElseIf lLateralCheck <> 0 Then
                sPortSide = C_Port_Lateral
            
            Else
                sPortSide = C_Port_Lateral
            End If
            
        End If

    End If
    
    Get_PortFaceType = sPortSide
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function




