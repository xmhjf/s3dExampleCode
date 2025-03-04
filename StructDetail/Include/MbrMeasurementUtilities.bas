Attribute VB_Name = "MbrMeasurementUtilities"
'*******************************************************************
'
'Copyright (C) 2007-16 Intergraph Corporation. All rights reserved.
'
'File : MbrMeasurementUtilities.bas
'
'Author : RCM
'
'Description :
'   MbrMeasurementUtilities included all the routines necessary to determine the
'   relative position of the bounded member with respect to the bounding member.
'
'History :
'   25/Aug/2011   - Addedd new methods to create InsetBrace
'   21/Sep/2011   - mpulikol
'           DI-CP-200263  Improve performance by caching measurement symbol results
'   19/Oct/2011   - mpulikol
'           CR-CP-203633 Performance: Increase speed of generic member assembly connections
'    30/Nov/2011 - svsmylav TR-205302: Replaced '.Subport' method call with 'GetLateralSubPortBeforeTrim'.
'    26/Feb/2012 - Alligators 'IsPortIntersectingObj' and 'IsExtendedPortIntersectingBoundingObj' are corrected.
'    6/Jul/2012 - Alligators DM-216590: 'GetSelForMbrBoundedToTube' is modified to handle border cases of
'                  Face-and-OutSide and Outside-and-OutSide to return 'To-Center' selector.
'    11/Oct/2012 - Alligators CR-207934/DM-220895(for v2013)
'                 (i) New method 'GetMeasurements' is prepared to compute dimensions from key points of bounding member
'                     to bounded member (ii) 'GetConnectedEdgeInfo' is updated to use method mensioned in (i), (iii) GetSelFrmBoundedToBoundingOrien
'                     and GetNonPenetratedIntersectedEdge methods are modified to set bUseTopFlange to false where appropriate,
'                    (iv) Added new methods 'AddMeasurementDimInfo', 'TestMeasurements', 'GetMappedPointLocationByIntersection',
'                         'GetStiffenerLoadPointPositionOnMemberSection', 'GetPointMappingFromEdgeMapping',
'                         'TranslateAxisCurve' and 'GetMemberTubeRadius'.
'    26/Oct/2012  - svsmylav
'               TR-219938: 'BoundedHasOutsideMaterial' is modified: if bounding member is tube and for
'               web penetrated case, check for bounded Top/Bottom flange and deduct flange thickness to determine
'               if still there is bounded material outside the bounding.
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'                Created GetEdgeMap,Set_CacheEdgeMapping,Get_CacheEdgeMapping to set and get the cache information of edgemap collection,
'                   section alias.
'    1/Aug/2013 - svsmylav
'               TR-223552: 'GetSketchPlaneForMSSymbol' is modified: with optional argument to persist plane.
'    4/Sep/2013 - svsmylav
'               TR-237205: 'GetMeasurements' method is modified to use 'Height' property on member part instead of 'WebLength'.
'                          (prior to fix, measurements were incorrect for few cross-sections and wrong Axis AC was selected).
'    05/03/2014 -DI-CP-235957  StrDet: Invalid DB Col Name issues (31 Char limit) view gen. on Oracle DB
'    28/Feb/2014 - knukala
'                 DM-249378: 'GetConnectedEdgeInfo' method is modified to use the distance from Flange Top/Bottom to WebRight instead of Flange Thickness for S-cross sections.
' 19/Mar/2014 - CM - TR-250554 : Fix for Access Violation errors. Checking if any object is nothing.
'                                Updated the project using Middle dlls but not Client Dlls
'                                Updated InItEndCutConnectionData() to get proper context port from
'                                connectable object and use it accrodingly to determine proper bounded obj.
'    11/Aug/2014 - knukala
'         CR-CP-250020: Create AC for ladder support or cage support to ladder rail
'    20/Aug/2014 - mchandak
'         CR-CP-250020: Ignoring the plate bounding case in GetConnectedEdgeInfo ,GetMemberBoundingCase and IsWebPenetrated methods
'    24/Sep/2014 - svsmylav
'         TR-CP-261460: Transition Plate ATP Failed for multi-bounded case viz. WT member bounded to WT member and also to a transition plate:
'            In 'IsMultiBoundingCase' method 'GetMultiBoundingEdgeMap' method call failed when eEndCutType parameter 'FlangeCutBottom'
'            (for 'WT' member there is no bottom flange). Fix: added checks before callling the method.
'    03/Nov/2014 - MDT/GH
'         CR-CP-250198  Lapped AC for traffic items
'    03/Dec/14  CSK/CM CR-250022 Connections and free end cuts at ends of ladder rails
'    29/NOV/2014 -GH
'        DI-259276 Updated GetMultiBoundingEdgeMap() method
'    24/Feb/2015 - MDT/GH
'       CR-265236 - GetFlippedPorts() is new method added to get the ports based on the conditions like "FlipPrimaryandSecondary" attribute
'                   and "Flip Primary and Secondary" question answer
' 24/Mar/15  MDT/RPK  TR-269306 Handrail placed on builtup member, cannot create assembly connection, errors,
'                    Added  checks related to design/Builtup member as the bounding member
' 22/April/15  MDT TR-271041 Added IsFlangeMultiBoundingCase() and modified ISMultiBoundingCase()
' 12/May/2015 GH-
'       CR-260982 - Added New method GetMemberDescriptionFromPropertyDescs()
'1/Sept/2015 pkakula
'            TR -270502: Modified GetBoundingCaseForTube() to handle whenever the both bounding and bounded are circular crossections
'27/Dec/2015  -hgunturu
'            DI-282754: Modified GetConnectedEdgeInfo,GetEdgeMap,GetMeasurements Methods
'                       to improve performance of Standard Generic AC
'    6/Jan/2016    svsmylav   DI-273986: Added optional argument 'retValue' to 'ItemExists' and 'KeyExists' methods.
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'                   Moved one of the method to MarineLibraryCommon.bas
'   21/Apr/2016 -  mkonduri
'                           TR-291919: MemberFreeEndCuts are failing when encut type is changed to welded.
'    15/June/2016    knukala   TR-CP-295640: Generic AC is updated when changing the corner feature type.
'                               Changed the logic in GetConnectedEdgeInfo() and GetEdgeMap() to get the correct cached data
'*****************************************************************************
Option Explicit

Private Const MODULE = "StructDetail\Data\Include\MbrMeasurementUtilities"
Private Const TOLERANCE_VALUE = 0.0001   'Tolerence for Double comparison
Private Const EDGE_CASE_TOL = 0.0005 '0.5mm

Public Const C_Port_Base = "Base"
Public Const C_Port_Offset = "Offset"
Public Const C_Port_Lateral = "Lateral"
Public Const C_Port_Top = "Top"
Public Const C_Port_Bottom = "Bottom"
Public Const C_Port_WebLeft = "WebLeft"
Public Const C_Port_WebRight = "WebRight"
Public Const BorderAC_OSOS = 1
Public Const BorderAC_FCOS_TOP = 2
Public Const BorderAC_FCOS_BTM = 3
Public Const BorderAC_ToCenter = 4
Public Const NOT_BorderAC = 5

Private m_oEdgeMappingRuleSymbolInstance As Object

Public Enum eCutFlag
    TopOrBottomCut = 0
    TopCut = 1
    BottomCut = 2
End Enum


' eRelativePointPosition specifies the location of a load point in the measurement
' symbol with respect to a bounded edge.
' It is also used to specify the location of a bounded edge (see eBounded_Edge) relative to
' a specified bounding edge (e.g. mapped bottom flange right)
Private Enum eRelativePointPosition
    Above = 1
    Below = 2
    Coincident = 3
End Enum


'eBounded_Edge has all the possible edges on the bounded member that are used
'by the measurement symbol.
Private Enum eBounded_Edge
    Top = 1
    Bottom = 2
    InsideTopFlange = 3
    InsideBottomFlange = 4
    WebLeft = 5
    WebRight = 6
    FlangeLeft = 7
    FlangeRight = 8
End Enum

Public Enum eMemberBoundingCase

    Unknown = 1
' Center
    Center
' ToEdge
    TopEdge
    BottomEdge
' EdgeAndEdge
    BottomEdgeAndTopEdge
' EdgeAndOS1Edge
    BottomEdgeAndOSTop
    TopEdgeAndOSBottom
' EdgeAndOS2Edge
    BottomEdgeAndOSTopEdge
    TopEdgeAndOSBottomEdge
' FCAndEdge
    FCAndBottomEdge
    FCAndTopEdge
' FCAndOS1Edge
    FCAndOSBottomEdge
    FCAndOSTopEdge
' FCAndOSNoEdge
    FCAndOSBottom
    FCAndOSTop
' OnMember
    OnMemberTop
    OnMemberBottom
' OSAndOS1Edge
    OSBottomAndOSTopEdge
    OSTopAndOSBottomEdge
' OSAndOS2Edge
    OSBottomEdgeAndOSTopEdge
' OSAndOSNoEdge
    OSBottomAndOSTop
' OnTubeMember
    OnTubeMember

End Enum

'*************************************************************************
'Function
'   GetConnectedEdgeInfo
'
'Abstract
'   Given the bounded and bounding members, this function determines the
'   relative position of the bounded member with respect to the bounding
'   member.  The function will return four ConnectedEdgeInfo types, one for
'   each edge of the bounded member in the measurement symbol.  Each ConnectedEdgeInfo
'   specifies the edge on the bounding member that the bounded member edge intersects
'   and whether or not the bounded member edge is coplanar to an edge on the bounding
'   member.
'
'Inputs
'   oACorEC As Object
'       - AppConnection/EndCut object on which the EdgeInfo will be cached
'   oSDO_Bounded As Object
'       - The Bounded Member
'   oSDO_Bounding As Object
'       - The Bounding Member
'   Optional OffsetFromTop As Double
'       - Specifies the offset between the Top edge and the InsideTopFlange construction
'           line on the Web Measurement Symbol.  If not specified, the offset will be set
'           to the flange thickness. Only used for web measurement symbols.
'   Optional OffsetFromBottom As Double
'       - Specifies the offset between the Bottom edge and the InsideBottomFlange construction
'           line on the Web Measurement Symbol.  If not specified, the offset will be set
'           to the flange thickness. Only used for web measurement symbols.
'   Optional bForceRecompute As Boolean = False
'       - If the flag is False, the EdgeInfo will be retrived from Cache if exist.
'         If the flage is True, the EdgeInfo will be computed again
'Return
'   TopOrWL As ConnectedEdgeInfo
'       - Specifies the edges on the bounding member that the Top or WebLeft edge of the
'           bounded member is intersecting & coplanar to.
'   BottomOrWR As ConnectedEdgeInfo
'       - Specifies the edges on the bounding member that the Bottom or WebRight edge of the
'           bounded member is intersecting & coplanar to.
'   InsideTopFlangeOrFL As ConnectedEdgeInfo
'       - Specifies the edges on the bounding member that the InsideTopFlange or FlangeLeft
'           edge of the bounded member is intersecting & coplanar to.
'   InsideBottomFlangeOrFR As ConnectedEdgeInfo
'       - Specifies the edges on the bounding member that the InsideBottomFlange or FlangeRight
'           edge of the bounded member is intersecting & coplanar to.
'Exceptions
'
'***************************************************************************
Public Sub GetConnectedEdgeInfo(oACOrEC As Object, _
                                oBoundedPort As IJPort, _
                                oBoundingPort As IJPort, _
                                ByRef TopOrWL As ConnectedEdgeInfo, _
                                ByRef BottomOrWR As ConnectedEdgeInfo, _
                                ByRef InsideTopFlangeOrFL As ConnectedEdgeInfo, _
                                ByRef InsideBottomFlangeOrFR As ConnectedEdgeInfo, _
                                Optional oMeasurements As Collection, _
                                Optional bPenetratesWeb As Boolean, _
                                Optional OffsetFromTop As Double = -1, _
                                Optional OffsetFromBottom As Double = -1, _
                                Optional bUseTopFlange As Boolean = True, _
                                Optional bForceRecompute As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' ------------------
    ' Get the connection
    ' ------------------
    Dim oACObject As IJAppConnection
    Dim sACItemName As String
    
    ' ---------------
    ' If not colinear
    ' ---------------
    
    Dim bEndToEnd As Boolean
    Dim bColinear As Boolean
    Dim bRightAngle As Boolean
    
    If oBoundingPort Is Nothing Then
        Exit Sub
    ElseIf TypeOf oBoundingPort.Connectable Is IJPlate Then
        Exit Sub
    ElseIf TypeOf oBoundingPort.Connectable Is SPSSlabEntity Then
        Exit Sub
    ElseIf TypeOf oBoundingPort.Connectable Is SPSWallPart Then
        Exit Sub
    ElseIf TypeOf oBoundingPort.Connectable Is ISPSDesignedMember Then
        Exit Sub
    End If
    
    If (Not oBoundedPort Is Nothing) And (Not oBoundingPort Is Nothing) Then
        CheckEndToEndConnection oBoundedPort.Connectable, oBoundingPort.Connectable, bEndToEnd, bColinear, bRightAngle
    End If
    
    If Not bColinear Then
    
        ' ------------------------------------------------------
        ' Try getting the cache data:
        ' When it is axis AC, get the cached data on AC (input can be AC or EC)
        ' When it is Generic AC and bounding collection is one, get the cached data on AC (input can be AC or EC)
        ' When it is Generic AC and bounding collection is more than one, and input is EC, get cached data on EC
        '                                                                     input is AC, making object as nothing to avoid getting cached data on AC
        ' ------------------------------------------------------
        Dim oObjectWithCachedData As Object
        Set oObjectWithCachedData = oACOrEC
        
        Dim eGetACType As eACType

        If Not oACOrEC Is Nothing Then
            If TypeOf oACOrEC Is IJStructFeature Then
                AssemblyConnection_SmartItemName oACOrEC, sACItemName, oACObject
            Else
                Set oACObject = oACOrEC
            End If
            
            eGetACType = GetMbrAssemblyConnectionType(oACObject)
            
            If eGetACType = ACType_Mbr_Generic Or eGetACType = ACType_Stiff_Generic Then
                Dim oReferencesCollection As IJDReferencesCollection
                Dim oEditJDArgument As IJDEditJDArgument
                Dim oBoundingObjectColl As IJElements
                Dim iBoundingObjectsCount As Integer
                
                'Get Bounding Object Collection Count
                Set oReferencesCollection = GetRefCollFromSmartOccurrence(oACObject)
                Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
                Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
                iBoundingObjectsCount = oBoundingObjectColl.Count
                If iBoundingObjectsCount = 1 Then
                    Set oObjectWithCachedData = oACObject
                ElseIf iBoundingObjectsCount > 1 Then
                    If TypeOf oACOrEC Is IJAppConnection Then
                        Set oObjectWithCachedData = Nothing
                    End If
                End If

            ElseIf Not eGetACType = ACType_Mbr_Generic And Not eGetACType = ACType_Stiff_Generic Then
                
                Set oObjectWithCachedData = oACObject
                
            End If
        End If
 
        ' -----------------------------------------------------------------------------------
        ' If force recompute is off, and AC or web cut is passed in, retrieve data from cache
        ' -----------------------------------------------------------------------------------
        Set oMeasurements = New Collection
        If (bForceRecompute = False) And (Not oObjectWithCachedData Is Nothing) Then
            If Get_CacheConnectedEdgeInfo(oObjectWithCachedData, _
                                          TopOrWL, _
                                          BottomOrWR, _
                                          InsideTopFlangeOrFL, _
                                          InsideBottomFlangeOrFR, _
                                          oMeasurements, _
                                          bPenetratesWeb) = True Then
                ' Got the details from Cache.. exit here....
                Exit Sub
            End If
        End If
        
        Dim sCacheString As String ' string used to cache data
        
        ' --------------------------------------
        ' Get the Edge Mapping and Section Alias
        ' --------------------------------------
        Dim sectionAlias As Long
        Dim oEdgeMap As JCmnShp_CollectionAlias
    
        If Not oObjectWithCachedData Is Nothing Then
            Set oEdgeMap = GetEdgeMap(oObjectWithCachedData, oBoundingPort, oBoundedPort, sectionAlias, bPenetratesWeb, bForceRecompute)
        Else
            Set oEdgeMap = GetEdgeMap(oACOrEC, oBoundingPort, oBoundedPort, sectionAlias, bPenetratesWeb, bForceRecompute)
        End If
        
        Dim eBoundingAlias As eBounding_Alias
        eBoundingAlias = GetBoundingAliasSimplified(sectionAlias)
        
        ' ----------------------------------------------------------------------------------------
        ' Compute the measurements as they would be retunred by the deprecated measurement symbols
        ' ----------------------------------------------------------------------------------------
        Dim bUseMeasureSymbol As Boolean
        bUseMeasureSymbol = True
                                
        Dim lStatus As Long
        Dim oBoundedData As MemberConnectionData
        Dim oBoundingData As MemberConnectionData
        Dim sMsg As String
                                
        InitMemberConnectionData oACObject, oBoundedData, oBoundingData, lStatus, sMsg
                        
        Dim compareCodeAndSymbolMeasurements As Boolean
        compareCodeAndSymbolMeasurements = False
        
        Dim oTestMeasure As Collection
        Dim testCacheString As String
    
        Dim sBoundingSectType As String
        Dim sBoundedSectType As String
        sBoundingSectType = UCase(GetSectionType(oBoundingPort))
        sBoundedSectType = UCase(GetSectionType(oBoundedPort))
        
         If (sBoundingSectType = "S" Or sBoundingSectType = "ISTYPE" Or sBoundingSectType = "ST" Or sBoundingSectType = "TSTYPE" Or sBoundingSectType = "C_SS" Or sBoundingSectType = "C_SType") Or _
        (sBoundedSectType = "S" Or sBoundedSectType = "ISTYPE" Or sBoundedSectType = "ST" Or sBoundedSectType = "TSTYPE" Or sBoundedSectType = "C_SS" Or sBoundedSectType = "C_SType") Then
            'For above cross-sections we noticed issue with 'xxxPt17xx' (dimension measurements involving point 17, since point 17 is
            'same as point 15 and point 16 - temporarily bank on old measurement symbol approach to get measurements for the first time -
            'since these dimensions are cached, no performance overhead for second call. However, we filed an item CR-222458 to
            'remove this temporary approach and to fix other known issues with 'S' types.
        Else
            If compareCodeAndSymbolMeasurements Then
                GetMeasurements oObjectWithCachedData, oBoundedPort, oBoundingPort, oBoundedData, sectionAlias, bPenetratesWeb, bUseTopFlange, oEdgeMap, oTestMeasure, testCacheString
            Else
                GetMeasurements oObjectWithCachedData, oBoundedPort, oBoundingPort, oBoundedData, sectionAlias, bPenetratesWeb, bUseTopFlange, oEdgeMap, oMeasurements, sCacheString
                bUseMeasureSymbol = False
            End If
        End If
        ' ---------------------------------------------------
        ' Only when debugging, execute the measurement symbol
        ' ---------------------------------------------------
        If bUseMeasureSymbol Then
        
            Dim MSSymbol As New StructDetailObjects.Measurement
            MSSymbol.Initialize GetResourceMgr()
            
            ' ----------------------------------------------------
            ' Select the measurement symbol based on section alias
            ' ----------------------------------------------------
            'Measurement symbol file prefix 'W_' for Web, 'F_' for Flange
            Dim sMeasurementFilePath As String
            Dim sMSFilePrefix As String
            
            If bPenetratesWeb Then
                sMSFilePrefix = "W_"
            Else
                sMSFilePrefix = "F_"
            End If
               
            Select Case eBoundingAlias
                Case eBounding_Alias.Web
                    sMeasurementFilePath = "MarineLibrary\Measurement\" & sMSFilePrefix & "WebMS.sym"
                Case eBounding_Alias.WebTopFlangeRight
                    sMeasurementFilePath = "MarineLibrary\Measurement\" & sMSFilePrefix & "WebTopFlangeRightMS.sym"
                Case eBounding_Alias.WebBuiltUpTopFlangeRight
                    sMeasurementFilePath = "MarineLibrary\Measurement\" & sMSFilePrefix & "WebBuiltUpTopFlangeRightMS.sym"
                Case eBounding_Alias.WebBottomFlangeRight
                    sMeasurementFilePath = "MarineLibrary\Measurement\" & sMSFilePrefix & "WebBottomFlangeRightMS.sym"
                Case eBounding_Alias.WebBuiltUpBottomFlangeRight
                    sMeasurementFilePath = "MarineLibrary\Measurement\" & sMSFilePrefix & "WebBuiltUpBottomFlangeRightMS.sym"
                Case eBounding_Alias.WebTopAndBottomRightFlanges
                    sMeasurementFilePath = "MarineLibrary\Measurement\" & sMSFilePrefix & "WebTopAndBottomRightFlangesMS.sym"
                Case eBounding_Alias.FlangeLeftAndRightBottomWebs
                    sMeasurementFilePath = "MarineLibrary\Measurement\" & sMSFilePrefix & "FlangeLeftAndRightBottomWebsMS.sym"
                Case eBounding_Alias.FlangeLeftAndRightTopWebs
                    sMeasurementFilePath = "MarineLibrary\Measurement\" & sMSFilePrefix & "FlangeLeftAndRightTopWebsMS.sym"
                Case eBounding_Alias.FlangeLeftAndRightWebs
                    sMeasurementFilePath = "MarineLibrary\Measurement\" & sMSFilePrefix & "FlangeLeftAndRightWebsMS.sym"
                Case 20
                    'Tube/Circular Cross Section
                Case Else
                    'Unknown Section Alias
            End Select
                
            ' -----------------------
            ' Get the sketching plane
            ' -----------------------
            Dim oSDOBoundedMember As New StructDetailObjects.MemberPart
            Dim oSDOBoundedStiffener As New StructDetailObjects.ProfilePart
            
            Dim dBoundedFlangeThickness As Double
            Dim dBoundedWebThickness As Double
            Dim dBoundedWidth As Double
            Dim dBoundedDepth As Double
            Dim sBoundedSectionType As String
    
            If TypeOf oBoundedPort.Connectable Is ISPSMemberPartCommon Then
                Set oSDOBoundedMember.object = oBoundedPort.Connectable
                dBoundedWebThickness = oSDOBoundedMember.webThickness
                dBoundedWidth = oSDOBoundedMember.FlangeLength
                dBoundedDepth = oSDOBoundedMember.WebLength
                sBoundedSectionType = oSDOBoundedMember.sectionType
            ElseIf TypeOf oBoundedPort.Connectable Is IJProfile Then
                Set oSDOBoundedStiffener.object = oBoundedPort.Connectable
                dBoundedWebThickness = oSDOBoundedStiffener.webThickness
                dBoundedWidth = oSDOBoundedStiffener.FlangeLength
                dBoundedDepth = oSDOBoundedStiffener.WebLength
                sBoundedSectionType = oSDOBoundedStiffener.sectionType
            Else
                '!!! Unknown Cases. Error Out or Need to Handle this unknown type
            End If
            
            'The thickness of flange is considered as the distance between Flange top or bottom to Web Right.
            'The flange thickness value that we get from StructDetailObjects doesnot support S- cross sections.
            'The below method GetDistanceFromTopOrBottomToWebRight() is used to fill appropriate flange thickness for S- cross sections also.
        
            Dim dDistfrmFlgToWR As Double
                If HasTopFlange(oBoundedPort.Connectable) Then
                    dDistfrmFlgToWR = GetDistanceFromTopOrBottomToWebRight(oBoundedPort.Connectable, False)
                Else
                    dDistfrmFlgToWR = GetDistanceFromTopOrBottomToWebRight(oBoundedPort.Connectable, True)
                End If
                dBoundedFlangeThickness = dDistfrmFlgToWR
            
            Dim oSketchPlane As IJPlane
            Set oSketchPlane = New Plane3d
            GetSketchPlaneForMSSymbol oBoundedPort, oBoundingPort, bPenetratesWeb, bUseTopFlange, oSketchPlane
                
            If bPenetratesWeb Then
                'Set the 'TopFlangeThickness' and 'BottomFlangeThickness' on the Web Measurement Symbol
                If OffsetFromTop < 0 Then
                    If HasTopFlange(oBoundedPort.Connectable) Then
                        MSSymbol.AddInputParameter "TopFlangeThickness", dBoundedFlangeThickness
                    ElseIf IsTubularMember(oBoundedPort.Connectable) Then
                        MSSymbol.AddInputParameter "TopFlangeThickness", dBoundedWebThickness
                    Else
                        MSSymbol.AddInputParameter "TopFlangeThickness", 0.00001
                    End If
                Else
                    If OffsetFromTop = 0 Then
                        MSSymbol.AddInputParameter "TopFlangeThickness", 0.00001
                    Else
                        MSSymbol.AddInputParameter "TopFlangeThickness", OffsetFromTop
                    End If
                End If
        
                If OffsetFromBottom < 0 Then
                    If HasBottomFlange(oBoundedPort.Connectable) Then
                        MSSymbol.AddInputParameter "BottomFlangeThickness", dBoundedFlangeThickness
                    ElseIf IsTubularMember(oBoundedPort.Connectable) Then
                        MSSymbol.AddInputParameter "BottomFlangeThickness", dBoundedWebThickness
                    Else
                        MSSymbol.AddInputParameter "BottomFlangeThickness", 0.00001
                    End If
                Else
                    If OffsetFromBottom = 0 Then
                        MSSymbol.AddInputParameter "BottomFlangeThickness", 0.00001
                    Else
                        MSSymbol.AddInputParameter "BottomFlangeThickness", OffsetFromBottom
                    End If
                End If
            Else
                'Set the 'WebLeftToFlangeRight' and 'TotalWidth' parameters on the Flange Measurement Symbol
                Select Case sBoundedSectionType
                Case "W", "M", "HP", "S", "H", "I", "ISType"
                    MSSymbol.AddInputParameter "TotalWidth", dBoundedWidth
                    MSSymbol.AddInputParameter "WebLeftToFlangeRight", dBoundedWidth / 2 + dBoundedWebThickness / 2
                Case "L", "EA", "UA"
                    MSSymbol.AddInputParameter "TotalWidth", dBoundedWidth
                    MSSymbol.AddInputParameter "WebLeftToFlangeRight", dBoundedWidth
                Case "C", "MC", "C_S", "C_SS", "CSType"
                    MSSymbol.AddInputParameter "TotalWidth", dBoundedWidth
                    MSSymbol.AddInputParameter "WebLeftToFlangeRight", dBoundedWidth
                Case "T", "MT", "ST", "WT", "BUT", "T_XType", "TSType"
                    MSSymbol.AddInputParameter "TotalWidth", dBoundedWidth
                    MSSymbol.AddInputParameter "WebLeftToFlangeRight", dBoundedWidth / 2 + dBoundedWebThickness / 2
                Case "HSSC", "PIPE"
                    'Circle
                Case "HSSR", "RS", "RECT"
                    'Rectangle
                    MSSymbol.AddInputParameter "TotalWidth", dBoundedWidth
                    MSSymbol.AddInputParameter "WebLeftToFlangeRight", dBoundedWidth
                Case "FB"
                    MSSymbol.AddInputParameter "TotalWidth", dBoundedWebThickness
                    MSSymbol.AddInputParameter "WebLeftToFlangeRight", dBoundedWebThickness
                Case "2L"
                    'Back to Back L
                Case "2C"
                    'Back to Back C
                End Select
            End If
            
            'Compute Measurement Symbol
            If bPenetratesWeb Then
                MSSymbol.ComputeProfileWeb sMeasurementFilePath, oBoundedPort, oBoundingPort, oSketchPlane
            Else
                MSSymbol.ComputeProfileFlange sMeasurementFilePath, oBoundedPort, oBoundingPort, oSketchPlane
            End If
        
            Set oMeasurements = GetDimensionsFromSymbol(MSSymbol, eBoundingAlias, bPenetratesWeb, sCacheString)
    
        End If ' using measurement symbol
    
        ' -------------------------
        ' Get the Intersecting Edges
        ' -------------------------
        If bPenetratesWeb Then
            GetIntersectingEdge eBounded_Edge.Top, eBoundingAlias, oMeasurements, TopOrWL.IntersectingEdge, TopOrWL.CoplanarEdge
            GetIntersectingEdge eBounded_Edge.Bottom, eBoundingAlias, oMeasurements, BottomOrWR.IntersectingEdge, BottomOrWR.CoplanarEdge
            GetIntersectingEdge eBounded_Edge.InsideTopFlange, eBoundingAlias, oMeasurements, InsideTopFlangeOrFL.IntersectingEdge, InsideTopFlangeOrFL.CoplanarEdge
            GetIntersectingEdge eBounded_Edge.InsideBottomFlange, eBoundingAlias, oMeasurements, InsideBottomFlangeOrFR.IntersectingEdge, InsideBottomFlangeOrFR.CoplanarEdge
        Else
            GetIntersectingEdge eBounded_Edge.WebLeft, eBoundingAlias, oMeasurements, TopOrWL.IntersectingEdge, TopOrWL.CoplanarEdge
            GetIntersectingEdge eBounded_Edge.WebRight, eBoundingAlias, oMeasurements, BottomOrWR.IntersectingEdge, BottomOrWR.CoplanarEdge
            GetIntersectingEdge eBounded_Edge.FlangeLeft, eBoundingAlias, oMeasurements, InsideTopFlangeOrFL.IntersectingEdge, InsideTopFlangeOrFL.CoplanarEdge
            GetIntersectingEdge eBounded_Edge.FlangeRight, eBoundingAlias, oMeasurements, InsideBottomFlangeOrFR.IntersectingEdge, InsideBottomFlangeOrFR.CoplanarEdge
        End If
        
        ' --------
        ' Clean up
        ' --------
        Dim oDelPlane As IJDObject
        Set oDelPlane = oSketchPlane
        
        If Not oDelPlane Is Nothing Then
            oDelPlane.Remove
        End If
        
        Set oSketchPlane = Nothing
        Set oDelPlane = Nothing
    
    End If
    
    ' ---------------------
    ' Cache the information
    ' ---------------------
    If Not oObjectWithCachedData Is Nothing Then
        Set_CacheConnectedEdgeInfo oObjectWithCachedData, _
                                   TopOrWL, _
                                   BottomOrWR, _
                                   InsideTopFlangeOrFL, _
                                   InsideBottomFlangeOrFR, _
                                   sCacheString, _
                                   bPenetratesWeb
    End If

    If compareCodeAndSymbolMeasurements Then
        TestMeasurements oTestMeasure, oMeasurements, bPenetratesWeb
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetConnectedEdgeInfo").Number
End Sub

'*************************************************************************
'Function
'   Set_CacheConnectedEdgeInfo
'
'Abstract
'   Cache the below data on IJUAMbrACCacheStorage
'       TopOrWL_CoplanarEdge, TopOrWL_IntersectingEdge
'       BottomOrWR_CoplanarEdge, BottomOrWR_IntersectingEdge
'       InsideTopFlangeOrTFL_CoplanarEdge, InsideTopFlangeOrTFL_IntersectingEdge
'       InsideBottomFlangeOrTFR_CoplanarEdge, InsideBottomFlangeOrTFR_IntersectingEdge
'
'***************************************************************************
Private Sub Set_CacheConnectedEdgeInfo(oACOrEC As Object, _
                                       TopOrWL As ConnectedEdgeInfo, _
                                       BottomOrWR As ConnectedEdgeInfo, _
                                       InsideTopFlangeOrTFL As ConnectedEdgeInfo, _
                                       InsideBottomFlangeOrTFR As ConnectedEdgeInfo, _
                                       sSymbolDimInfo As String, _
                                       bPenetratesWeb As Boolean)
    On Error GoTo ErrorHandler

    Dim oAttributes As IJDAttributes
    Set oAttributes = oACOrEC
    
    Dim oAttributesCol As IJDAttributesCol
    Set oAttributesCol = oAttributes.CollectionOfAttributes("IJUAMbrACCacheStorage")
    
    If Not oAttributesCol Is Nothing Then
        oAttributesCol.Item("TopOrWL_CoplanarEdge").value = TopOrWL.CoplanarEdge
        oAttributesCol.Item("TopOrWL_IntersectingEdge").value = TopOrWL.IntersectingEdge
        
        oAttributesCol.Item("BottomOrWR_CoplanarEdge").value = BottomOrWR.CoplanarEdge
        oAttributesCol.Item("BottomOrWR_IntersectingEdge").value = BottomOrWR.IntersectingEdge
            
        oAttributesCol.Item("InsTopFlgOrTFL_CoplanarEdge").value = InsideTopFlangeOrTFL.CoplanarEdge
        oAttributesCol.Item("InsTopFlgOrTFL_IntrsEdge").value = InsideTopFlangeOrTFL.IntersectingEdge
            
        oAttributesCol.Item("InsBtmFlgOrTFR_CoplanarEdge").value = InsideBottomFlangeOrTFR.CoplanarEdge
        oAttributesCol.Item("InsBtmFlgOrTFR_IntrsEdge").value = InsideBottomFlangeOrTFR.IntersectingEdge
        
        oAttributesCol.Item("IsWebPenetrated").value = bPenetratesWeb
        
        Dim sParsedString As String
        sParsedString = ParseDimString(sSymbolDimInfo)
        If Len(sParsedString) = 0 Then sParsedString = sSymbolDimInfo
        
        oAttributesCol.Item("MeasurementSymbolDimInfo").value = sParsedString
    End If

    Set oAttributesCol = Nothing
    Set oAttributes = Nothing
    
    Exit Sub
ErrorHandler:

End Sub

'*************************************************************************
'Function
'   ParseDimString
'
'Abstract
'   Parse the dimension info string as below to cut short the length before storing into database
'       DimPt -> *
'       DepthAtPt --> @
'       ToTop ---> &
'       ToBottom -> #
'       Inside --> $
'
'***************************************************************************
Private Function ParseDimString(sSymbolDimInfo As String) As String
    Dim sInputString As String
    sInputString = sSymbolDimInfo
    
    sInputString = Replace(sInputString, "DimPt", "*", , , vbTextCompare)
    sInputString = Replace(sInputString, "DepthAtPt", "@", , , vbTextCompare)
    sInputString = Replace(sInputString, "ToTop", "&", , , vbTextCompare)
    sInputString = Replace(sInputString, "ToBottom", "#", , , vbTextCompare)
    sInputString = Replace(sInputString, "Inside", "$", , , vbTextCompare)
    
    ParseDimString = sInputString
    
End Function

'*************************************************************************
'Function
'   DeCodeDimString
'
'Abstract
'   Decode the dimension info string from database. Refer ParseDimString method for details
'       DimPt -> *
'       DepthAtPt --> @
'       ToTop ---> &
'       ToBottom -> #
'       Inside --> $
'
'***************************************************************************
Private Function DeCodeDimString(sInputDimInfo As String) As String
    Dim sInputString As String
    sInputString = sInputDimInfo
    
    sInputString = Replace(sInputString, "*", "DimPt", , , vbTextCompare)
    sInputString = Replace(sInputString, "@", "DepthAtPt", , , vbTextCompare)
    sInputString = Replace(sInputString, "&", "ToTop", , , vbTextCompare)
    sInputString = Replace(sInputString, "#", "ToBottom", , , vbTextCompare)
    sInputString = Replace(sInputString, "$", "Inside", , , vbTextCompare)
    
    DeCodeDimString = sInputString
    
End Function

'*************************************************************************
'Function
'   Get_CacheConnectedEdgeInfo
'
'Abstract
'   Get the below data from Cache on IJUAMbrACCacheStorage
'       TopOrWL_CoplanarEdge, TopOrWL_IntersectingEdge
'       BottomOrWR_CoplanarEdge, BottomOrWR_IntersectingEdge
'       InsideTopFlangeOrTFL_CoplanarEdge, InsideTopFlangeOrTFL_IntersectingEdge
'       InsideBottomFlangeOrTFR_CoplanarEdge, InsideBottomFlangeOrTFR_IntersectingEdge
'
'***************************************************************************
Private Function Get_CacheConnectedEdgeInfo(oACOrEC As Object, _
                                            ByRef TopOrWL As ConnectedEdgeInfo, _
                                            ByRef BottomOrWR As ConnectedEdgeInfo, _
                                            ByRef InsideTopFlangeOrTFL As ConnectedEdgeInfo, _
                                            ByRef InsideBottomFlangeOrTFR As ConnectedEdgeInfo, _
                                            ByRef oMeasurements As Collection, _
                                            ByRef bPenetratesWeb As Boolean) As Boolean
    
    On Error GoTo ErrorHandler
    Get_CacheConnectedEdgeInfo = False

    Dim oAttributes As IJDAttributes
    Set oAttributes = oACOrEC

    Dim oAttributesCol As IJDAttributesCol
    
    On Error Resume Next
    Set oAttributesCol = oAttributes.CollectionOfAttributes("IJUAMbrACCacheStorage")
    Err.Clear
    On Error GoTo ErrorHandler
    
    Dim dValue As Integer
    If Not oAttributesCol Is Nothing Then
        
        dValue = oAttributesCol.Item("TopOrWL_CoplanarEdge").value
        If dValue <= 0 Then GoTo CleanUp
        TopOrWL.CoplanarEdge = dValue
        
        dValue = oAttributesCol.Item("TopOrWL_IntersectingEdge").value
        If dValue <= 0 Then GoTo CleanUp
        TopOrWL.IntersectingEdge = dValue
        
        dValue = oAttributesCol.Item("BottomOrWR_CoplanarEdge").value
        If dValue <= 0 Then GoTo CleanUp
        BottomOrWR.CoplanarEdge = dValue
        
        dValue = oAttributesCol.Item("BottomOrWR_IntersectingEdge").value
        If dValue <= 0 Then GoTo CleanUp
        BottomOrWR.IntersectingEdge = dValue
        
        dValue = oAttributesCol.Item("InsTopFlgOrTFL_CoplanarEdge").value
        If dValue <= 0 Then GoTo CleanUp
        InsideTopFlangeOrTFL.CoplanarEdge = dValue
        
        dValue = oAttributesCol.Item("InsTopFlgOrTFL_IntrsEdge").value
        If dValue <= 0 Then GoTo CleanUp
        InsideTopFlangeOrTFL.IntersectingEdge = dValue
        
        dValue = oAttributesCol.Item("InsBtmFlgOrTFR_CoplanarEdge").value
        If dValue <= 0 Then GoTo CleanUp
        InsideBottomFlangeOrTFR.CoplanarEdge = dValue
        
        dValue = oAttributesCol.Item("InsBtmFlgOrTFR_IntrsEdge").value
        If dValue <= 0 Then GoTo CleanUp
        InsideBottomFlangeOrTFR.IntersectingEdge = dValue
    
        Dim sValue As String
        sValue = oAttributesCol.Item("MeasurementSymbolDimInfo").value
        If Len(sValue) > 0 Then
            Set oMeasurements = New Collection
            Set oMeasurements = GetMeasurementSymDimInfoCollection(sValue)
        End If
    
        bPenetratesWeb = oAttributesCol.Item("IsWebPenetrated").value
        
        Get_CacheConnectedEdgeInfo = True
    End If
    
CleanUp:

    Set oAttributesCol = Nothing
    Set oAttributes = Nothing
    
    Exit Function
ErrorHandler:

End Function

'*************************************************************************
'Function
'   GetDimensionsFromSymbol
'
'Abstract
'   Function to retrieve all the required dimensions from the Measurement symbol.
'   These dimensions are used to determine the relative position of each point with respect
'   to the bounded edges.
'
'Inputs
'   MSSymbol As StructDetailObjects.Measurement
'       - The measurement symbol
'   eBoundingAlias As eBounding_Alias
'       - The cross section alias of the bounding member
'   bPenetratesWeb As Boolean
'       - Whether or not the web of the bounded member is penetrated.
'Return
'   A collection of all dimensions retrieved from the measurement symbols.  The dimensions
'   are stored in the collection with keys equal to the dimension variable name in the
'   measurement symbol.
'Exceptions
'
'***************************************************************************
Private Function GetDimensionsFromSymbol(MSSymbol As StructDetailObjects.Measurement, _
                                         eBoundingAlias As eBounding_Alias, _
                                         bPenetratesWeb As Boolean, _
                                         ByRef outDimInfo As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim oMbrMSDimensions As New Collection
    If bPenetratesWeb Then  'Get Dimensions from the Web Measurement Symbol
    
        'Common Dimensions
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToTop", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToBottom", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToTopInside", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToBottomInside", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt11", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt11", outDimInfo
        
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToTop", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToBottom", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToTopInside", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToBottomInside", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt15", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt15", outDimInfo
        
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt23ToTop", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt23ToBottom", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt23ToTopInside", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt23ToBottomInside", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt23", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt23", outDimInfo
        
        
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToTop", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToBottom", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToTopInside", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToBottomInside", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt3", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt3", outDimInfo
        
        
        If eBoundingAlias = eBounding_Alias.WebTopFlangeRight Or eBoundingAlias = eBounding_Alias.WebBuiltUpTopFlangeRight Or eBoundingAlias = eBounding_Alias.WebTopAndBottomRightFlanges Then
            'Dimensions Common to WebTopFlangeRight, WebBuiltUpTopFlangeRight and WebTopAndBottomRightFlanges
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToTop", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToBottom", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToTopInside", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToBottomInside", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt18", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt18", outDimInfo

            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToTop", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToBottom", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToTopInside", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToBottomInside", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt17", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt17", outDimInfo
        
            If eBoundingAlias = eBounding_Alias.WebBuiltUpTopFlangeRight Then
                'Dimensions only for WebBuiltUpTopFlangeRight
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt14ToTop", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt14ToBottom", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt14ToTopInside", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt14ToBottomInside", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt14", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt14", outDimInfo
                
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt50ToTop", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt50ToBottom", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt50ToTopInside", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt50ToBottomInside", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt50", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt50", outDimInfo
            End If
        End If
        
        If eBoundingAlias = eBounding_Alias.WebBottomFlangeRight Or eBoundingAlias = eBounding_Alias.WebBuiltUpBottomFlangeRight Or eBoundingAlias = eBounding_Alias.WebTopAndBottomRightFlanges Then
            'Dimensions Common to WebBottomFlangeRight, WebBuiltUpBottomFlangeRight, WebTopAndBottomRightFlanges
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToTop", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToBottom", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToTopInside", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToBottomInside", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt20", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt20", outDimInfo

            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToTop", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToBottom", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToTopInside", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToBottomInside", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt21", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt21", outDimInfo
        
        
            If eBoundingAlias = eBounding_Alias.WebBuiltUpBottomFlangeRight Then
                'Dimensions only for WebBuiltUpBottomFlangeRight
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt24ToTop", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt24ToBottom", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt24ToTopInside", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt24ToBottomInside", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt24", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt24", outDimInfo
                    
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt51ToTop", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt51ToBottom", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt51ToTopInside", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt51ToBottomInside", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DepthAtPt51", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "InsideDepthAtPt51", outDimInfo
            End If
        End If
    Else 'Get Dimensions from the Flange Measurement Symbol
    
        'Common Dimensions
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToWL", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToWR", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToFL", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToFR", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt11", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt11", outDimInfo
    
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToWL", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToWR", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToFL", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToFR", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt15", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt15", outDimInfo
        

        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt23ToWL", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt23ToWR", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt23ToFL", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt23ToFR", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt23", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt23", outDimInfo
        
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToWL", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToWR", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToFL", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToFR", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt3", outDimInfo
        AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt3", outDimInfo
        
        If eBoundingAlias = eBounding_Alias.WebTopFlangeRight Or eBoundingAlias = eBounding_Alias.WebBuiltUpTopFlangeRight Or eBoundingAlias = eBounding_Alias.WebTopAndBottomRightFlanges Then
            'Dimensions Common to WebTopFlangeRight,WebBuiltUpTopFlangeRight,WebTopAndBottomRightFlanges
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToWL", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToWR", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToFL", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToFR", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt18", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt18", outDimInfo
    
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToWL", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToWR", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToFL", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToFR", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt17", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt17", outDimInfo
        
            If eBoundingAlias = eBounding_Alias.WebBuiltUpTopFlangeRight Then
                'Dimensions for only WebBuiltUpTopFlangeRight
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt14ToWL", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt14ToWR", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt14ToFL", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt14ToFR", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt14", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt14", outDimInfo
                
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt50ToWL", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt50ToWR", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt50ToFL", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt50ToFR", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt50", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt50", outDimInfo
            End If
        End If
        
        If eBoundingAlias = eBounding_Alias.WebBottomFlangeRight Or eBoundingAlias = eBounding_Alias.WebBuiltUpBottomFlangeRight Or eBoundingAlias = eBounding_Alias.WebTopAndBottomRightFlanges Then
            'Dimensions common to WebBottomFlangeRight, WebBuiltUpBottomFlangeRight, WebTopAndBottomRightFlanges
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToWL", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToWR", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToFL", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToFR", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt20", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt20", outDimInfo
    
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToWL", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToWR", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToFL", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToFR", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt21", outDimInfo
            AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt21", outDimInfo
        
            If eBoundingAlias = eBounding_Alias.WebBuiltUpBottomFlangeRight Then
                'Dimensions for only WebBuiltUpBottomFlangeRight
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt24ToWL", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt24ToWR", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt24ToFL", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt24ToFR", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt24", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt24", outDimInfo
                
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt51ToWL", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt51ToWR", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt51ToFL", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt51ToFR", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WidthAtPt51", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "WebThkAtPt51", outDimInfo
            End If
        End If
    End If
    
    ' -----------------------------
    ' Get point-to-point dimensions
    ' -----------------------------
    Select Case eBoundingAlias
        Case eBounding_Alias.Web, eBounding_Alias.FlangeLeftAndRightBottomWebs, eBounding_Alias.FlangeLeftAndRightWebs, eBounding_Alias.FlangeLeftAndRightTopWebs
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToPt15", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToPt23", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToPt23", outDimInfo
        
        Case eBounding_Alias.WebBottomFlangeRight
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToPt15", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToPt20", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToPt21", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToPt23", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToPt23", outDimInfo
        Case eBounding_Alias.WebBuiltUpBottomFlangeRight
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToPt15", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToPt20", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToPt21", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToPt23", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt23ToPt51", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt24ToPt51", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToPt24", outDimInfo
        Case eBounding_Alias.WebBuiltUpTopFlangeRight
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToPt14", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt14ToPt50", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToPt50", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToPt17", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToPt18", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToPt23", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToPt23", outDimInfo
        Case eBounding_Alias.WebTopAndBottomRightFlanges
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToPt15", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToPt17", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToPt18", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToPt20", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt20ToPt21", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt21ToPt23", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToPt23", outDimInfo
        Case eBounding_Alias.WebTopFlangeRight
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt11ToPt15", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt15ToPt17", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt17ToPt18", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt18ToPt23", outDimInfo
                AddMeasurementSymbolDimInfo oMbrMSDimensions, MSSymbol, "DimPt3ToPt23", outDimInfo

    End Select
    
    Set GetDimensionsFromSymbol = oMbrMSDimensions
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetDimensionsFromSymbol").Number
End Function

'*************************************************************************
'Function
'   AddMeasurementSymbolDimInfo
'
'Abstract
'   Function to add up all the Dim info in the input collection to a String
'   in the below format
'   ex. "DimPt11ToPt15=1.55879004;DimPt15ToPt17=2.469276;DimPt17ToPt18=0.3803475;"
'
'***************************************************************************
Private Sub AddMeasurementSymbolDimInfo(ByRef oMeasurements As Collection, MSSymbol As StructDetailObjects.Measurement, sParameter As String, ByRef outDimInfo As String)

    oMeasurements.Add MSSymbol.GetOutputParameter(sParameter), sParameter
    outDimInfo = outDimInfo & sParameter & "=" & Round(MSSymbol.GetOutputParameter(sParameter), 6) & ";"
    
End Sub

'*************************************************************************
'Function
'   GetMeasurementSymDimInfoCollection
'
'Abstract
'   Function to read the below formatted String and puts each Dim info into
'   Collection item
'   ex. "DimPt11ToPt15=1.55879004;DimPt15ToPt17=2.469276;DimPt17ToPt18=0.3803475;"
'
'***************************************************************************
Private Function GetMeasurementSymDimInfoCollection(sSymDimInfo As String) As Collection

    Dim oMeasurements As Collection
    Set oMeasurements = New Collection
    
    Dim sDecodedString As String
    sDecodedString = DeCodeDimString(sSymDimInfo)
    If Len(sDecodedString) = 0 Then sDecodedString = sSymDimInfo

    
    On Error Resume Next
    Dim arrDimInfo() As String
    arrDimInfo = Split(sDecodedString, ";", , vbTextCompare)
    
    Dim nLowBound As Integer
    Dim nUpBound As Integer
    
    nLowBound = LBound(arrDimInfo)
    nUpBound = UBound(arrDimInfo)
    
    Dim nIndex As Integer: nIndex = 0
    For nIndex = nLowBound To nUpBound
        Dim sDimInfo As String: sDimInfo = ""
        sDimInfo = arrDimInfo(nIndex)
        
        If Len(sDimInfo) > 0 Then
            Dim nEqualPos As Integer: nEqualPos = 0
            nEqualPos = InStr(1, sDimInfo, "=", vbTextCompare)
            
            If nEqualPos > 0 Then
                Dim sDimName As String
                Dim dDimValue As Double
                
                sDimName = VBA.Left$(sDimInfo, nEqualPos - 1)
                dDimValue = VBA.Right$(sDimInfo, Len(sDimInfo) - nEqualPos)
                
                oMeasurements.Add dDimValue, sDimName
            End If
        End If
    Next
    Set GetMeasurementSymDimInfoCollection = oMeasurements
    
End Function

'*************************************************************************
'Function
'   GetRelativePointPosition
'
'Abstract
'   Function to determine the relative position of a point with respect
'   to a bounded edge.
'
'Inputs
'   ByVal PointNumber As Integer
'       - The number of the point (From the Measurement Symbol)
'   ByVal eBoundedEdge As eBounded_Edge
'       - The bounded edge to get the relative position with respect to (IE Top, Bottom)
'   oMSDimensionColl As Collection
'       - The collection of dimensions retrieved from the measurement symbol.
'Return
'   The relative position of the point with respect to the given bounded edge.
'Exceptions
'
'***************************************************************************
Private Function GetRelativePointPosition(ByVal PointNumber As Integer, ByVal eBoundedEdge As eBounded_Edge, oMSDimensionColl As Collection) As eRelativePointPosition
    On Error GoTo ErrorHandler
    
    Dim eRelativePointPos As eRelativePointPosition
    
    'Web Measurement Symbol Dimensions
    Dim dTop As Double 'Distance from point to Top
    Dim dBottom As Double 'Distance from point to Bottom
    Dim dInsideTopFlange As Double 'Distance from point to InsideTopFlange (Construction line)
    Dim dInsideBottomFlange As Double 'Distance from point to InsideBottomFlange (Construction line)
    
    Dim dWL As Double 'Distance from point to Web Left
    Dim dWR As Double 'Distance from point to Web Right
    Dim dTFL As Double 'Distance from point to Top Flange Left
    Dim dTFR As Double 'Distance from point to Top Flange Right
    
    'Dimensions to determine the Width/Depth of the Bounded Member
    Dim dDepth As Double
    Dim dWidth As Double
    Dim dWebLength As Double
    Dim dWebTh As Double
    
    Dim dTolerance As Double
    Dim lRoundOff As Long
    
    dTolerance = 0.000011 'minimum tolerance to be set
    lRoundOff = 6 'RoundOff the values upto 6 digits after decimal
    
    Select Case eBoundedEdge
        Case eBounded_Edge.Top
            dTop = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToTop"), lRoundOff)
            dBottom = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToBottom"), lRoundOff)
            dDepth = Round(oMSDimensionColl.Item("DepthAtPt" & PointNumber), lRoundOff)
            
            If GreaterThanZero(dTop, dTolerance) And GreaterThanZero(dBottom, dTolerance) Then
                If Equal(dTop + dBottom, dDepth) Then
                    eRelativePointPos = eRelativePointPosition.Below
                ElseIf LessThan(dTop, dBottom) Then
                    eRelativePointPos = eRelativePointPosition.Above
                Else
                    eRelativePointPos = eRelativePointPosition.Below
                End If
            ElseIf Equal(dTop, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Coincident
            ElseIf Equal(dBottom, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Below
            Else
                'Unexpected Case (Negative Dimension)
            End If
            
        Case eBounded_Edge.Bottom
            dTop = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToTop"), lRoundOff)
            dBottom = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToBottom"), lRoundOff)
            dDepth = Round(oMSDimensionColl.Item("DepthAtPt" & PointNumber), lRoundOff)

            If GreaterThanZero(dTop, dTolerance) And GreaterThanZero(dBottom, dTolerance) Then
                If Equal(dTop + dBottom, dDepth) Then
                    eRelativePointPos = eRelativePointPosition.Above
                ElseIf LessThan(dTop, dBottom) Then
                    eRelativePointPos = eRelativePointPosition.Above
                Else
                    eRelativePointPos = eRelativePointPosition.Below
                End If
            ElseIf Equal(dTop, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Above
            ElseIf Equal(dBottom, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Coincident
            Else
                'Unexpected Case (Negative Dimension)
            End If

        Case eBounded_Edge.InsideTopFlange
            dInsideTopFlange = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToTopInside"), lRoundOff)
            dInsideBottomFlange = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToBottomInside"), lRoundOff)
            dWebLength = Round(oMSDimensionColl.Item("InsideDepthAtPt" & PointNumber), lRoundOff)

            If GreaterThanZero(dInsideTopFlange, dTolerance) And GreaterThanZero(dInsideBottomFlange, dTolerance) Then
                If Equal(dInsideTopFlange + dInsideBottomFlange, dWebLength) Then
                    eRelativePointPos = eRelativePointPosition.Below
                ElseIf LessThan(dInsideTopFlange, dInsideBottomFlange) Then
                    eRelativePointPos = eRelativePointPosition.Above
                Else
                    eRelativePointPos = eRelativePointPosition.Below
                End If
            ElseIf Equal(dInsideTopFlange, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Coincident
            ElseIf Equal(dInsideBottomFlange, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Below
            Else
                'Unexpected Case (Negative Dimension)
            End If

        Case eBounded_Edge.InsideBottomFlange
            dInsideTopFlange = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToTopInside"), lRoundOff)
            dInsideBottomFlange = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToBottomInside"), lRoundOff)
            dWebLength = Round(oMSDimensionColl.Item("InsideDepthAtPt" & PointNumber), lRoundOff)
            
            If GreaterThanZero(dInsideTopFlange, dTolerance) And GreaterThanZero(dInsideBottomFlange, dTolerance) Then
                If Equal(dInsideTopFlange + dInsideBottomFlange, dWebLength) Then
                    eRelativePointPos = eRelativePointPosition.Above
                ElseIf LessThan(dInsideTopFlange, dInsideBottomFlange) Then
                    eRelativePointPos = eRelativePointPosition.Above
                Else
                    eRelativePointPos = eRelativePointPosition.Below
                End If
            ElseIf Equal(dInsideTopFlange, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Above
            ElseIf Equal(dInsideBottomFlange, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Coincident
            Else
                'Unexpected Case (Negative Dimension)
            End If

        Case eBounded_Edge.WebLeft
            dWL = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToWL"), lRoundOff)
            dWR = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToWR"), lRoundOff)
            dWebTh = Round(oMSDimensionColl.Item("WebThkAtPt" & PointNumber), lRoundOff)
            
            If GreaterThanZero(dWL, dTolerance) And GreaterThanZero(dWR, dTolerance) Then
                If Equal(dWL + dWR, dWebTh) Then
                    eRelativePointPos = eRelativePointPosition.Below
                ElseIf LessThan(dWL, dWR) Then
                    eRelativePointPos = eRelativePointPosition.Above
                Else
                    eRelativePointPos = eRelativePointPosition.Below
                End If
            ElseIf Equal(dWL, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Coincident
            ElseIf Equal(dWR, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Below
            Else
                'Unexpected Case (Negative Dimension)
            End If
        
        Case eBounded_Edge.WebRight
            dWL = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToWL"), lRoundOff)
            dWR = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToWR"), lRoundOff)
            dWebTh = Round(oMSDimensionColl.Item("WebThkAtPt" & PointNumber), lRoundOff)

            If GreaterThanZero(dWL, dTolerance) And GreaterThanZero(dWR, dTolerance) Then
                If Equal(dWL + dWR, dWebTh) Then
                    eRelativePointPos = eRelativePointPosition.Above
                ElseIf LessThan(dWL, dWR) Then
                    eRelativePointPos = eRelativePointPosition.Above
                Else
                    eRelativePointPos = eRelativePointPosition.Below
                End If
            ElseIf Equal(dWL, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Above
            ElseIf Equal(dWR, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Coincident
            Else
                'Unexpected Case (Negative Dimension)
            End If

        Case eBounded_Edge.FlangeLeft
            dTFL = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToFL"), lRoundOff)
            dTFR = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToFR"), lRoundOff)
            dWidth = Round(oMSDimensionColl.Item("WidthAtPt" & PointNumber), lRoundOff)
            
            If GreaterThanZero(dTFL, dTolerance) And GreaterThanZero(dTFR, dTolerance) Then
                If Equal(dTFL + dTFR, dWidth) Then
                    eRelativePointPos = eRelativePointPosition.Below
                ElseIf LessThan(dTFL, dTFR) Then
                    eRelativePointPos = eRelativePointPosition.Above
                Else
                    eRelativePointPos = eRelativePointPosition.Below
                End If
            ElseIf Equal(dTFL, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Coincident
            ElseIf Equal(dTFR, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Below
            Else
                'Unexpected Case (Negative Dimension)
            End If
            
        Case eBounded_Edge.FlangeRight
            dTFL = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToFL"), lRoundOff)
            dTFR = Round(oMSDimensionColl.Item("DimPt" & PointNumber & "ToFR"), lRoundOff)
            dWidth = Round(oMSDimensionColl.Item("WidthAtPt" & PointNumber), lRoundOff)
            
            If GreaterThanZero(dTFL, dTolerance) And GreaterThanZero(dTFR, dTolerance) Then
                If Equal(dTFL + dTFR, dWidth) Then
                    eRelativePointPos = eRelativePointPosition.Above
                ElseIf LessThan(dTFL, dTFR) Then
                    eRelativePointPos = eRelativePointPosition.Above
                Else
                    eRelativePointPos = eRelativePointPosition.Below
                End If
            ElseIf Equal(dTFL, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Above
            ElseIf Equal(dTFR, 0, dTolerance) Then
                eRelativePointPos = eRelativePointPosition.Coincident
            Else
                'Unexpected Case (Negative Dimension)
            End If
    End Select
    
    GetRelativePointPosition = eRelativePointPos
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetRelativePointPosition").Number
End Function

'*************************************************************************
'Function
'   GetIntersectingEdge
'
'Abstract
'   Function to determine the edge that the bounded members edge intersects as
'   well as the edge that the bounded members edge is coplanar too.
'
'Inputs
'   eBoundedEdge As eBounded_Edge
'       - The edge of the bounded member
'   eBoundingAlias As eBounding_Alias
'       - The cross section alias of the bounding member
'   oMSDimensionColl As Collection
'       - The collection of dimensions retrieved from the measurement symbol
'Return
'   ByRef eIntersectedEdge As eBounding_Edge
'       - The edge on the bounding member, which the bounded edge intersects
'   ByRef eCoplanarEdge As eBounding_Edge
'       - The edge on the bounding member, which the bounded edge is coplanar too
'Exceptions
'
'***************************************************************************
Private Sub GetIntersectingEdge(ByVal eBoundedEdge As eBounded_Edge, _
                                eBoundingAlias As eBounding_Alias, _
                                oMSDimensionColl As Collection, _
                                ByRef eIntersectedEdge As eBounding_Edge, _
                                ByRef eCoplanarEdge As eBounding_Edge)
    
    On Error GoTo ErrorHandler

    'By Default, assume no Coplanar Edge
    eCoplanarEdge = eBounding_Edge.None
    
    Dim eRelativePos11 As eRelativePointPosition
    Dim eRelativePos15 As eRelativePointPosition
    Dim eRelativePos23 As eRelativePointPosition
    Dim eRelativePos3 As eRelativePointPosition
    
    eRelativePos11 = GetRelativePointPosition(11, eBoundedEdge, oMSDimensionColl)
    eRelativePos15 = GetRelativePointPosition(15, eBoundedEdge, oMSDimensionColl)
    
    eRelativePos23 = GetRelativePointPosition(23, eBoundedEdge, oMSDimensionColl)
    eRelativePos3 = GetRelativePointPosition(3, eBoundedEdge, oMSDimensionColl)
    
    Select Case eBoundingAlias
        Case eBounding_Alias.Web, eBounding_Alias.FlangeLeftAndRightBottomWebs, eBounding_Alias.FlangeLeftAndRightTopWebs, eBounding_Alias.FlangeLeftAndRightWebs, eBounding_Alias.Tube
            
            If eRelativePos11 = eRelativePointPosition.Below And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Below - 15 Below
                eIntersectedEdge = eBounding_Edge.Above
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Coincident And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Coincident - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Above And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Above - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos15 = eRelativePointPosition.Coincident Then
                '15 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Web_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top
                End Select
                If eRelativePos11 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos15 = eRelativePointPosition.Above And eRelativePos23 = eRelativePointPosition.Below Then
                '15 Above - 23 Below
                eIntersectedEdge = eBounding_Edge.Web_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Coincident Then
                '23 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Web_Right
                End Select
                If eRelativePos3 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Below Then
                '23 Above - 3 Below
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Coincident Then
                '23 Above - 3 Coincident
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Above Then
                '23 Above - 3 Above
                eIntersectedEdge = eBounding_Edge.Below
                eCoplanarEdge = eBounding_Edge.None
            End If
            
        Case eBounding_Alias.WebTopFlangeRight
            If eRelativePos11 = eRelativePointPosition.Below And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Below - 15 Below
                eIntersectedEdge = eBounding_Edge.Above
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Coincident And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Coincident - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Above And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Above - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos15 = eRelativePointPosition.Coincident Then
                '15 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top
                End Select
                If eRelativePos11 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos15 = eRelativePointPosition.Above And GetRelativePointPosition(17, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '15 Above - 17 Below
                eIntersectedEdge = eBounding_Edge.Top_Flange_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(17, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '17 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right
                End Select
                If GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top_Flange_Right_Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf GetRelativePointPosition(17, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '17 Above - 18 Below
                eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '18 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Web_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Bottom
                End Select
            ElseIf GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And eRelativePos23 = eRelativePointPosition.Below Then
                '18 Above - 23 Below
                eIntersectedEdge = eBounding_Edge.Web_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Coincident Then
                '23 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Web_Right
                End Select
                If eRelativePos3 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Below Then
                '23 Above - 3 Below
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Coincident Then
                '23 Above - 3 Coincident
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Above Then
                '23 Above - 3 Above
                eIntersectedEdge = eBounding_Edge.Below
                eCoplanarEdge = eBounding_Edge.None
            End If
            
        Case eBounding_Alias.WebBuiltUpTopFlangeRight
            If eRelativePos11 = eRelativePointPosition.Below And GetRelativePointPosition(14, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '11 Below - 14 Below
                eIntersectedEdge = eBounding_Edge.Above
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Coincident And GetRelativePointPosition(14, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '11 Coincident - 14 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Above And GetRelativePointPosition(14, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '11 Above - 14 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(14, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident And eRelativePos15 = eRelativePointPosition.Below Then
                '14 Coincident - 15 Below
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Web_Right_Top
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top
                End Select
                If eRelativePos11 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf GetRelativePointPosition(14, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And GetRelativePointPosition(50, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below And eRelativePos15 = eRelativePointPosition.Below Then
                '14 Above - 50 Below - 15 Below
                eIntersectedEdge = eBounding_Edge.Web_Right_Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(50, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident And eRelativePos15 = eRelativePointPosition.Below Then
                '50 Coincident - 15 Below
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Top
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Web_Right_Top
                End Select
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(50, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And eRelativePos15 = eRelativePointPosition.Below Then
                '50 Above - 15 Below
                eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos15 = eRelativePointPosition.Coincident Then
                '15 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Top
                End Select
                If GetRelativePointPosition(50, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top_Flange_Right_Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos15 = eRelativePointPosition.Above And GetRelativePointPosition(17, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '15 Above - 17 Below
                eIntersectedEdge = eBounding_Edge.Top_Flange_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(17, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '17 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right
                End Select
                If GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top_Flange_Right_Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf GetRelativePointPosition(17, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '17 Above - 18 Below
                eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '18 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Web_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Bottom
                End Select
            ElseIf GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And eRelativePos23 = eRelativePointPosition.Below Then
                '18 Above - 23 Below
                eIntersectedEdge = eBounding_Edge.Web_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Coincident Then
                '23 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Web_Right
                End Select
                If eRelativePos3 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Below Then
                '23 Above - 3 Below
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Coincident Then
                '23 Above - 3 Coincident
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Above Then
                '23 Above - 3 Above
                eIntersectedEdge = eBounding_Edge.Below
                eCoplanarEdge = eBounding_Edge.None
            End If
            
        Case eBounding_Alias.WebBottomFlangeRight
            If eRelativePos11 = eRelativePointPosition.Below And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Below - 15 Below
                eIntersectedEdge = eBounding_Edge.Above
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Coincident And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Coincident - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Above And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Above - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos15 = eRelativePointPosition.Coincident And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '15 Coincident - 21 Below
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Web_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top
                End Select
                If eRelativePos11 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos15 = eRelativePointPosition.Above And GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '15 Above - 20 Below - 21 Below
                eIntersectedEdge = eBounding_Edge.Web_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '20 Coincident - 21 Below
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Top
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Web_Right
                End Select
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '20 Above - 21 Below
                eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '21 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Top
                End Select
                If GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom_Flange_Right_Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And eRelativePos23 = eRelativePointPosition.Below Then
                '21 Above - 23 Below
                eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Coincident Then
                '23 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right
                End Select
                If eRelativePos3 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Below Then
                '23 Above - 3 Below
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Coincident Then
                '23 Above - 3 Coincident
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Above Then
                '23 Above - 3 Above
                eIntersectedEdge = eBounding_Edge.Below
                eCoplanarEdge = eBounding_Edge.None
            End If
            
        Case eBounding_Alias.WebBuiltUpBottomFlangeRight
            If eRelativePos11 = eRelativePointPosition.Below And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Below - 15 Below
                eIntersectedEdge = eBounding_Edge.Above
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Coincident And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Coincident - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Above And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Above - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos15 = eRelativePointPosition.Coincident And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '15 Coincident - 21 Below
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Web_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top
                End Select
                If eRelativePos11 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos15 = eRelativePointPosition.Above And GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '15 Above - 20 Below - 21 Below
                eIntersectedEdge = eBounding_Edge.Web_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '20 Coincident - 21 Below
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Top
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Web_Right
                End Select
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '20 Above - 21 Below
                eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '21 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Top
                End Select
                If GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom_Flange_Right_Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And eRelativePos23 = eRelativePointPosition.Below Then
                '21 Above - 23 Below
                eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Coincident Then
                '23 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right
                End Select
                If GetRelativePointPosition(51, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom_Flange_Right_Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos23 = eRelativePointPosition.Above And GetRelativePointPosition(51, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '23 Above - 51 Below
                eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(51, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '51 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Web_Right_Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Bottom
                End Select
            ElseIf GetRelativePointPosition(51, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And GetRelativePointPosition(24, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '51 Above - 24 Below
                eIntersectedEdge = eBounding_Edge.Web_Right_Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(24, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '24 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Web_Right_Bottom
                End Select
                If eRelativePos3 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf GetRelativePointPosition(24, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Below Then
                '24 Above - 3 Below
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(24, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Coincident Then
                '24 Above - 3 Coincident
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(24, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Above Then
                '24 Above - 3 Above
                eIntersectedEdge = eBounding_Edge.Below
                eCoplanarEdge = eBounding_Edge.None
            End If
            
        Case eBounding_Alias.WebTopAndBottomRightFlanges
            If eRelativePos11 = eRelativePointPosition.Below And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Below - 15 Below
                eIntersectedEdge = eBounding_Edge.Above
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Coincident And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Coincident - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos11 = eRelativePointPosition.Above And eRelativePos15 = eRelativePointPosition.Below Then
                '11 Above - 15 Below
                eIntersectedEdge = eBounding_Edge.Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos15 = eRelativePointPosition.Coincident Then
                '15 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top
                End Select
                If eRelativePos11 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos15 = eRelativePointPosition.Above And GetRelativePointPosition(17, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '15 Above - 17 Below
                eIntersectedEdge = eBounding_Edge.Top_Flange_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(17, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '17 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right
                End Select
                If GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Top_Flange_Right_Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf GetRelativePointPosition(17, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '17 Above - 18 Below
                eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '18 Coincident - 21 Below
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Web_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Top_Flange_Right_Bottom
                End Select
            ElseIf GetRelativePointPosition(18, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '18 Above - 20 Below - 21 Below
                eIntersectedEdge = eBounding_Edge.Web_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '20 Coincident - 21 Below
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Top
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Web_Right
                End Select
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Below Then
                '20 Above - 21 Below
                eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Top
                eCoplanarEdge = eBounding_Edge.None
            ElseIf GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                '21 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right_Top
                End Select
                If GetRelativePointPosition(20, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom_Flange_Right_Top
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf GetRelativePointPosition(21, eBoundedEdge, oMSDimensionColl) = eRelativePointPosition.Above And eRelativePos23 = eRelativePointPosition.Below Then
                '21 Above - 23 Below
                eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Coincident Then
                '23 Coincident
                Select Case eBoundedEdge
                    Case eBounded_Edge.Top, eBounded_Edge.InsideTopFlange, eBounded_Edge.FlangeLeft, eBounded_Edge.WebLeft
                        eIntersectedEdge = eBounding_Edge.Bottom
                    Case Else
                        eIntersectedEdge = eBounding_Edge.Bottom_Flange_Right
                End Select
                If eRelativePos3 = eRelativePointPosition.Coincident Then
                    eCoplanarEdge = eBounding_Edge.Bottom
                Else
                    eCoplanarEdge = eBounding_Edge.None
                End If
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Below Then
                '23 Above - 3 Below
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Coincident Then
                '23 Above - 3 Coincident
                eIntersectedEdge = eBounding_Edge.Bottom
                eCoplanarEdge = eBounding_Edge.None
            ElseIf eRelativePos23 = eRelativePointPosition.Above And eRelativePos3 = eRelativePointPosition.Above Then
                '23 Above - 3 Above
                eIntersectedEdge = eBounding_Edge.Below
                eCoplanarEdge = eBounding_Edge.None
            End If
    End Select
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetIntersectingEdge").Number
End Sub

'***********************************************************************
' This method tries to retrieve the End Cut Edge mapping Rule ProgID
' from Catalog(if Bulkloaded).
'        If its not bulkloaded we just hard code the ProgID and return
'        it as an output
'***********************************************************************
Public Function EdgeMappingRuleProgID() As String

Const METHOD = "EdgeMappingRuleProgID()"
On Error GoTo ErrorHandler

    Dim oCatalogQuery As IJSRDQuery
    Dim oRule As IJSRDRule
    Dim sClassName As String
    Dim sAssemblyName As String
    Dim oRuleQuery As IJSRDRuleQuery
     
    Set oCatalogQuery = New SRDQuery
    Set oRuleQuery = oCatalogQuery.GetRulesQuery
    
    Dim sRuleName As String
    Dim oRuleUnk As Object
    
    sRuleName = "EndCutBoundingMapRule"
    
    ' Check if a Rule Has been bulkloaded into the Catalog
    On Error Resume Next
    
    Set oRuleUnk = oRuleQuery.GetRule(sRuleName)
    
    Err.Clear
    On Error GoTo ErrorHandler
    
    If Not oRuleUnk Is Nothing Then
       ' EndCut Mapping Rule was found
       Set oRule = oRuleUnk
       
       EdgeMappingRuleProgID = oRule.ProgId
    End If
    
    Set oCatalogQuery = Nothing
    Set oRuleQuery = Nothing
    Set oRule = Nothing
    Set oRuleUnk = Nothing
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

'The user should handle any errors returned by the method.
Public Function CreateSymbolInstance(strProgId As String) As Object
Const METHOD = "CreateSymbolInstance"
On Error Resume Next
    
    Dim strCodeBase As String
    strCodeBase = Null
    
    Dim oCreateInstanceHelper As New CreateInstanceHelper
    Set CreateSymbolInstance = oCreateInstanceHelper.CreateInstance(strProgId, strCodeBase)
    Set oCreateInstanceHelper = Nothing
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function


'The user should handle any errors returned by the method.
Public Function CreateEdgeMappingRuleSymbolInstance() As Object
    Const METHOD = "CreateEdgeMappingRuleSymbolInstance"
    On Error Resume Next
        
    If m_oEdgeMappingRuleSymbolInstance Is Nothing Then
        Set m_oEdgeMappingRuleSymbolInstance = CreateSymbolInstance(EdgeMappingRuleProgID)
    End If
    
    Set CreateEdgeMappingRuleSymbolInstance = m_oEdgeMappingRuleSymbolInstance

    Set m_oEdgeMappingRuleSymbolInstance = Nothing
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
    
End Function

'------------------------------------------------------------------------------------------------------------
' METHOD:  IsPortIntersectingObj
'
' DESCRIPTION : Finds the Intersection between Port and Object(in model) and
'               returns TRUE if Intersection is Found
'               Default answer is set to FALSE
'
' Inputs : oPort
'          oObj
'          Optional oExtendedPort, if provided geometry of this will be used
'          Any valid Port and valid Object from Model
'
' Output : TRUE (If Intersects)
'          FALSE (If No Intersection is found)
'------------------------------------------------------------------------------------------------------------
Public Function IsPortIntersectingObj(oPort As IJPort, oObj As Object, _
    Optional oExtendedPort As Object = Nothing) As Boolean

    Const METHOD = "IsPortIntersectingObj"
 
    IsPortIntersectingObj = False
 
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim oPortModelBody As IJDModelBody
    Dim oPortConnectable As Object
    Dim oMemberPart As New StructDetailObjects.MemberPart
    ' ----------------------
    ' Check for valid Inputs
    ' ----------------------
    If oPort Is Nothing Then
        GoTo ErrorHandler
    ElseIf oPort.Geometry Is Nothing Then
        'Valid Port Geometry doesnt exist(not a valid input)
        sMsg = METHOD & " : Input Port Geometry doesnt exist"
        GoTo ErrorHandler
    ElseIf TypeOf oPort.Geometry Is IJDModelBody Then
        Set oPortModelBody = oPort.Geometry
    End If
    
    If Not oExtendedPort Is Nothing Then
        If TypeOf oPortModelBody Is IJDModelBody Then
            Set oPortModelBody = oExtendedPort 'Use extended port's model body
        Else
             GoTo ErrorHandler
        End If
    End If
    
    ' ------------------------------
    ' Check if given port intersects
    ' ------------------------------
    Dim oGeomOpr As GSCADShipGeomOps.SGOModelBodyUtilities
    Dim oPointOnPort As IJDPosition
    Dim oPointOnObj As IJDPosition
    Dim dDistance As Double
    
    Set oGeomOpr = New SGOModelBodyUtilities
    
    Dim oObjModelBody As IJModelBody, oObjPort As IJPort
    
    If TypeOf oObj Is IJModelBody Then
        Set oObjModelBody = oObj
    ElseIf TypeOf oObj Is IJPort Then
        Set oObjPort = oObj
        Set oObjModelBody = oObjPort.Geometry
    ElseIf TypeOf oObj Is ISPSMemberPartCommon Then
        Set oMemberPart.object = oObj
        Set oObjPort = oMemberPart.BasePortBeforeTrim(BPT_Lateral)
        Set oObjModelBody = oObjPort.Geometry
    Else
        'Unknown case
        GoTo ErrorHandler
    End If

    oGeomOpr.GetClosestPointsBetweenTwoBodies oPortModelBody, _
                                              oObjModelBody, _
                                              oPointOnPort, _
                                              oPointOnObj, _
                                              dDistance
    If (dDistance < 0.000001) Then
        IsPortIntersectingObj = True
        Exit Function
    End If
    
    ' --------------------------------------------------------------
    ' If no intersection and the given port is web left or web right
    ' --------------------------------------------------------------
    Dim oStructPort As IJStructPort
    Set oStructPort = oPort
    
    If oStructPort.SectionID = JXSEC_WEB_LEFT Or oStructPort.SectionID = JXSEC_WEB_RIGHT Then
    
        ' -----------------------------------------------------
        ' Project the port position to the top and bottom ports
        ' -----------------------------------------------------
        Dim oTopPort As IJPort
        Dim oBtmPort As IJPort
        
        Set oTopPort = GetLateralSubPortBeforeTrim(oPort.Connectable, JXSEC_TOP)
        Set oBtmPort = GetLateralSubPortBeforeTrim(oPort.Connectable, JXSEC_BOTTOM)
        
        Dim oTopPos As IJDPosition
        Dim oBtmPos As IJDPosition
        
        oGeomOpr.GetClosestPointOnBody oTopPort.Geometry, oPointOnPort, oTopPos, dDistance
        oGeomOpr.GetClosestPointOnBody oBtmPort.Geometry, oPointOnPort, oBtmPos, dDistance
        
        ' ------------------------------------
        ' Create a wire between the two points
        ' ------------------------------------
        Dim oGeomFactory As IJGeometryFactory
        Dim oLine As IJLine
        Set oGeomFactory = New GeometryFactory
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oTopPos.x, oTopPos.y, oTopPos.z, oBtmPos.x, oBtmPos.y, oBtmPos.z)
        
        Dim oGeomMisc As IJGeometryMisc
        Dim oWire As Object
        
        Set oGeomMisc = New DGeomOpsMisc
        oGeomMisc.CreateModelGeometryFromGType Nothing, oLine, Nothing, oWire
                                    
        ' --------------------------------------------------------
        ' If the wire intersects, consider the web as intersecting
        ' --------------------------------------------------------
        IsPortIntersectingObj = oGeomOpr.HasIntersectingGeometry(oWire, oObj)
    End If
        
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function
'------------------------------------------------------------------------------------------------------------
' METHOD:  IsExtendedPortIntersectingBoundingObj
'
' DESCRIPTION : Finds the Intersection between ExtendedPort(Port which gets extended before
'               Trim is applied) and Object(in model) and
'               returns TRUE if Intersection is Found
'               Default answer is set to FALSE
'
'               This method first sets the default answer to FALSE
'               then Gets the Extended Port of the Input port passed
'               And then finds the Intersection between the Extended Port
'               and Input Bounding Object(argument passed)
'
' Inputs : oPort
'          oObj
'          Any valid Port and valid Object from Model
'
' Output : TRUE (If Intersects)
'          FALSE (If No Intersection is found)
'------------------------------------------------------------------------------------------------------------
Public Function IsExtendedPortIntersectingBoundingObj(oLateralPort As IJPort, oBoundingObj As Object) As Boolean

 Const METHOD = "IsExtendedPortIntersectingBoundingObj"
 Dim sMsg As String
    
 On Error GoTo ErrorHandler
    
    IsExtendedPortIntersectingBoundingObj = False
    
    ' ---------------------
    ' Get the extended port
    ' ---------------------
    Dim oGeomOpr As GSCADShipGeomOps.SGOModelBodyUtilities
    Dim oExtendedPort As Object
        
    Set oGeomOpr = New SGOModelBodyUtilities
    ' Gets the Extended Port
    Set oExtendedPort = GetExtendedPort(oLateralPort)
           
    ' ----------------------------
    ' Check if the port intersects
    ' ----------------------------
    IsExtendedPortIntersectingBoundingObj = IsPortIntersectingObj(oLateralPort, oBoundingObj, oExtendedPort)
    
    Set oGeomOpr = Nothing
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function
'------------------------------------------------------------------------------------------------------------
' METHOD:  GetExtendedPort
'
' DESCRIPTION:  Gets the ExtendedPort(Port which gets extended before
'               Trim is applied) of the Member Part(Standard(Rolled) Members)
'
'               Method can be enhanced for Built Up's(Designed Members)
'
' Inputs : oPort : Any existing Lateral Port of the member part
'
'
' Output : Returns Extend Port of the Argument passed
'          (Returned Output may or maynot be of Type IJPort)
'------------------------------------------------------------------------------------------------------------
Public Function GetExtendedPort(oPort As IJPort) As Object

 Const METHOD = "GetExtendedPort"
 Dim sMsg As String
 
 On Error GoTo ErrorHandler
 
     Dim oConnectable As IJConnectable
     Dim strJXSEC_CODE As String
     Dim oMemberFactory As SPSMembers.SPSMemberFactory
     Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
     Dim oStructProfPart As IJStructProfilePart
     Dim lCtxId As Long
     Dim lOptId As Long
     Dim lOprId As Long
     Dim oExtendedPortElem As IJElements
     Dim CollofFaces() As String
     Dim i As Long
     Dim ePortType As JS_TOPOLOGY_PROXY_TYPE
     
     Set oConnectable = oPort.Connectable
     
     'SP3D Member Object Ports do not implement IJStructPort interface
     'if the port object is other than Member Object then the method needs to be enhanced
     If TypeOf oConnectable Is ISPSMemberPartPrismatic Then
     
            Set oMemberFactory = New SPSMembers.SPSMemberFactory
            Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices
                           
            oMemberConnectionServices.GetStructPortInfo oPort, ePortType, _
                                                        lCtxId, lOptId, lOprId
    ElseIf TypeOf oConnectable Is IJProfile Then
           
        Dim oStructPort As IJStructPort
        Set oStructPort = oPort
        lOprId = oStructPort.OperatorID
        
    Else 'Could be Designed (built-up) member
       sMsg = METHOD & " not available for built-up members"
       GoTo ErrorHandler
    End If
            
            If lOprId = JXSEC_TOP Then
                strJXSEC_CODE = "514"
            ElseIf lOprId = JXSEC_BOTTOM Then
                strJXSEC_CODE = "513"
            ElseIf lOprId = JXSEC_WEB_LEFT Then
                strJXSEC_CODE = "257"
            ElseIf lOprId = JXSEC_WEB_RIGHT Then
                strJXSEC_CODE = "258"
            ElseIf lOprId = JXSEC_TOP_FLANGE_RIGHT Then
                strJXSEC_CODE = "1028"
            ElseIf lOprId = JXSEC_TOP_FLANGE_LEFT Then
                strJXSEC_CODE = "1026"
            ElseIf lOprId = JXSEC_BOTTOM_FLANGE_RIGHT Then
                strJXSEC_CODE = "1027"
            ElseIf lOprId = JXSEC_BOTTOM_FLANGE_LEFT Then
                strJXSEC_CODE = "1025"
            ElseIf lOprId = JXSEC_TOP_FLANGE_LEFT_BOTTOM Then
                strJXSEC_CODE = "770"
            ElseIf lOprId = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
                strJXSEC_CODE = "772"
            ElseIf lOprId = JXSEC_BOTTOM_FLANGE_LEFT_TOP Then
                strJXSEC_CODE = "769"
            ElseIf lOprId = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
                strJXSEC_CODE = "771"
            End If
            
            If TypeOf oConnectable Is IJStructProfilePart Then
                Set oStructProfPart = oConnectable
                oStructProfPart.GetSectionFaces True, oExtendedPortElem, CollofFaces()
                ' Verify returned List of Port(s) is valid
                If oExtendedPortElem Is Nothing Then
                    sMsg = "oExtendedPortElem Is Nothing"
                ElseIf oExtendedPortElem.Count < 1 Then
                    sMsg = "oExtendedPortElem.Count < 1"
                Else
                    For i = 1 To oExtendedPortElem.Count
                        If CollofFaces(i - 1) = strJXSEC_CODE Then
                           Set GetExtendedPort = oExtendedPortElem.Item(i)
                           Exit For
                        End If
                    Next i
                End If
            Else
                sMsg = METHOD & " : Object doesn't support IJStructProfilePart Interface"
                GoTo ErrorHandler
            End If
      
      Set oConnectable = Nothing
      Set oStructProfPart = Nothing
      Set oExtendedPortElem = Nothing
      Set oMemberFactory = Nothing
      Set oMemberConnectionServices = Nothing
      
 Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'   GetBoundingObjectsFromPorts
'
'Abstract
'   Gets all the bounding objects from the collection of selected ports
'
'Inputs
'   oEditJDArgument As IJDEditJDArgument
'       - Ports
'Return
'   A collection of all bounding objects.
'Exceptions
'
'***************************************************************************
Public Function GetBoundingObjectsFromPorts(oEditJDArgument As IJDEditJDArgument) As IJElements

    Dim iPortIndex As Long
    Dim iObjectIndex As Long
    Dim oArgObject As Object
    Dim oPort As IJPort
    Dim oBoundingObj As IJDObject
    Dim oBoundingObjColl As IJElements
    Dim bContainsObject As Boolean
   
    Set oBoundingObjColl = New JObjectCollection

    For iPortIndex = 1 To oEditJDArgument.GetCount
        Set oArgObject = oEditJDArgument.GetEntityByIndex(iPortIndex)
        If TypeOf oArgObject Is IJPort Then
            Set oPort = oArgObject
            Set oBoundingObj = oPort.Connectable

            If oBoundingObjColl.Count = 0 Then
                oBoundingObjColl.Add oBoundingObj
            Else
                bContainsObject = False
                For iObjectIndex = 1 To oBoundingObjColl.Count
                    If oBoundingObjColl.Item(iObjectIndex) Is oBoundingObj Then
                        bContainsObject = True
                        Exit For
                    End If
                Next iObjectIndex
                If bContainsObject = False Then oBoundingObjColl.Add oBoundingObj
            End If
        Else
            'Not a Port - Ignore the argument
        End If
    Next iPortIndex
    
    Set GetBoundingObjectsFromPorts = oBoundingObjColl

End Function

'*************************************************************************
'Function
'   GetPortsFromBoundingObject
'
'Abstract
'   Gets all the selected ports from the bounding object
'
'Inputs
'   oBoundingObject As Object
'       - The Bounding Object
'   oEditJDArgument As IJDEditJDArgument
'       - Selected Ports
'Return
'   The selected ports from the bounding object.  Only returns ports selected
'   by the user.
'Exceptions
'
'***************************************************************************
Public Function GetPortsFromBoundingObject(oBoundingObject As Object, oEditJDArgument As IJDEditJDArgument) As IJElements

    Dim iPortIndex As Long
    Dim oArgObject As Object
    Dim oPort As IJPort
    
    Dim oSelectedPortColl As IJElements
    Set oSelectedPortColl = New JObjectCollection

    For iPortIndex = 1 To oEditJDArgument.GetCount
        Set oArgObject = oEditJDArgument.GetEntityByIndex(iPortIndex)
        If TypeOf oArgObject Is IJPort Then
            Set oPort = oArgObject
            If oPort.Connectable Is oBoundingObject Then
                oSelectedPortColl.Add oPort
            End If
        Else
            'Not a Port - Ignore the argument
        End If
    Next iPortIndex
    
    Set GetPortsFromBoundingObject = oSelectedPortColl

End Function

'*************************************************************************
'Function
'   ReverseEdgeMapping
'
'Abstract
'   Reverses the Edge Mapping and returns the actual edge
'
'Inputs
'   mappedEdge As JXSEC_CODE
'       - The Mapped Edge
'   map As Collection
'       - The Mapped Collection
'Return
'   The actual edge before the mapping
'Exceptions
'
'***************************************************************************
Public Function ReverseMap(mappedEdge As JXSEC_CODE, map As Collection) As JXSEC_CODE

    ReverseMap = JXSEC_UNKNOWN

    Dim keys As New Collection
    keys.Add JXSEC_WEB_LEFT
    keys.Add JXSEC_WEB_RIGHT
    keys.Add JXSEC_BOTTOM
    keys.Add JXSEC_TOP
    keys.Add JXSEC_BOTTOM_FLANGE_LEFT_TOP
    keys.Add JXSEC_TOP_FLANGE_LEFT_BOTTOM
    keys.Add JXSEC_BOTTOM_FLANGE_RIGHT_TOP
    keys.Add JXSEC_TOP_FLANGE_RIGHT_BOTTOM
    keys.Add JXSEC_BOTTOM_FLANGE_LEFT
    keys.Add JXSEC_TOP_FLANGE_LEFT
    keys.Add JXSEC_BOTTOM_FLANGE_RIGHT
    keys.Add JXSEC_TOP_FLANGE_RIGHT
    keys.Add JXSEC_WEB_RIGHT_TOP_CORNER
    keys.Add JXSEC_WEB_RIGHT_BOTTOM_CORNER
    keys.Add JXSEC_WEB_LEFT_BOTTOM_CORNER
    keys.Add JXSEC_WEB_LEFT_TOP_CORNER
    keys.Add JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER
    keys.Add JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER
    keys.Add JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER
    keys.Add JXSEC_BOTTOM_FLANGE_LEFT_TOP_CORNER
    keys.Add JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM_CORNER
    keys.Add JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER
    keys.Add JXSEC_TOP_FLANGE_LEFT_TOP_CORNER
    keys.Add JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM_CORNER
    keys.Add JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM
    keys.Add JXSEC_TOP_FLANGE_LEFT_TOP
    keys.Add JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM
    keys.Add JXSEC_TOP_FLANGE_RIGHT_TOP
    keys.Add JXSEC_WEB_LEFT_BOTTOM_TOP_CORNER
    keys.Add JXSEC_WEB_LEFT_TOP_BOTTOM_CORNER
    keys.Add JXSEC_WEB_RIGHT_BOTTOM_TOP_CORNER
    keys.Add JXSEC_WEB_RIGHT_TOP_BOTTOM_CORNER
    keys.Add JXSEC_WEB_LEFT_BOTTOM
    keys.Add JXSEC_WEB_LEFT_TOP
    keys.Add JXSEC_WEB_RIGHT_BOTTOM
    keys.Add JXSEC_WEB_RIGHT_TOP
    keys.Add JXSEC_WEB_LEFT_BOTTOM_BOTTOM_CORNER
    keys.Add JXSEC_WEB_LEFT_TOP_TOP_CORNER
    keys.Add JXSEC_WEB_RIGHT_BOTTOM_BOTTOM_CORNER
    keys.Add JXSEC_WEB_RIGHT_TOP_TOP_CORNER
    keys.Add JXSEC_LEFT_WEB_TOP
    keys.Add JXSEC_LEFT_WEB_BOTTOM
    keys.Add JXSEC_INNER_WEB_LEFT_TOP
    keys.Add JXSEC_INNER_WEB_LEFT_BOTTOM
    keys.Add JXSEC_RIGHT_WEB_TOP
    keys.Add JXSEC_RIGHT_WEB_BOTTOM
    keys.Add JXSEC_INNER_WEB_RIGHT_TOP
    keys.Add JXSEC_INNER_WEB_RIGHT_BOTTOM
    keys.Add JXSEC_INNER_WEB_RIGHT
    keys.Add JXSEC_INNER_WEB_LEFT
    keys.Add JXSEC_OUTER_TUBE
    
    Dim key As JXSEC_CODE

    Dim i As Long
    For i = 1 To keys.Count
        If ItemExists(keys.Item(i), map) Then
            key = keys.Item(i)
            If map.Item(CStr(key)) = mappedEdge Then
                ReverseMap = key
                Exit For
            End If
        End If
    Next i

End Function

'*************************************************************************
'Function
'   ItemExists
'
'Abstract
'   Determines if the specified edge is in the mapped edge collection.
'
'Inputs
'   Key As JXSEC_CODE
'       - The Mapped Edge
'   MappedEdgeColl As Collection
'       - The Mapped Collection
'Optional output RealEdge As JXSEC_CODE
'       - The Real Edge which is mapped to the input Key
'
'Output (return value): TRUE (If input key exists in the collection), otherwise FALSE
'Exceptions
'
'***************************************************************************
Public Function ItemExists(key As JXSEC_CODE, MappedEdgeColl As Collection, Optional RealEdge As JXSEC_CODE) As Boolean

    ItemExists = KeyExists(CStr(key), MappedEdgeColl, RealEdge)
    
End Function

Public Function KeyExists(key As String, oCollection As Collection, Optional retValue As Variant) As Boolean

    On Error Resume Next
    retValue = oCollection.Item(key)
    KeyExists = (Err <> 5) '5 - Err Code returned when key doesn't exist
    Err.Clear

End Function

Private Function GetResourceMgr() As IJDPOM

    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim pConnMiddle As IJDConnectMiddle
    Dim pAccessMiddle As IJDAccessMiddle
    
    Dim jContext As IJContext
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
 
    Set pConnMiddle = jContext.GetService("ConnectMiddle")
 
    Set pAccessMiddle = pConnMiddle
 
    Dim strModelDB As String
    strModelDB = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set GetResourceMgr = pAccessMiddle.GetResourceManager(strModelDB)
  
      
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set pConnMiddle = Nothing
    Set pAccessMiddle = Nothing
End Function

Public Sub GetSelFrmBoundedToBoundingOrien(oAppConnection As Object, _
                                           oBounded As MemberConnectionData, _
                                           oBounding As MemberConnectionData, _
                                           selString As String, _
                                           Optional oSelectorLogic As IJDSelectorLogic, _
                                           Optional bForceComputeEdgeInfo As Boolean = True)
    On Error GoTo ErrorHandler

    Const METHOD = "GetSelFrmBoundedToBoundingOrien"

    ' ----------------------
    ' Determine bounded type
    ' ----------------------
    If TypeOf oBounded.MemberPart Is ISPSMemberPartCommon Then
        selString = "MbrAxisBy"
    ElseIf TypeOf oBounded.MemberPart Is IJProfile Then
        selString = "StiffEndByMbr"
    Else
        'Using the default string for Members
        selString = "MbrAxisBy"
    End If

    ' ----------------------------------------------------------
    ' Create string based on bounded type and bounding condition
    ' ----------------------------------------------------------
    Dim eCase As eMemberBoundingCase
    eCase = GetMemberBoundingCase(oAppConnection, , , bForceComputeEdgeInfo)
    
    Select Case eCase
        Case Center
            selString = selString & "Center"
        Case TopEdge, BottomEdge
            selString = selString & "Edge"
        Case BottomEdgeAndTopEdge
         If TypeOf oBounded.MemberPart Is IJProfile Then
            selString = selString & "EdgeEdge"
         Else
            selString = selString & "EdgeAndEdge"
         End If
        Case BottomEdgeAndOSTop, TopEdgeAndOSBottom
         If TypeOf oBounded.MemberPart Is IJProfile Then
            selString = selString & "EdgOS1Edg"
         Else
            selString = selString & "EdgeAndOS1Edge"
         End If
        Case BottomEdgeAndOSTopEdge, TopEdgeAndOSBottomEdge
         If TypeOf oBounded.MemberPart Is IJProfile Then
            selString = selString & "EdgOS2Edg"
         Else
            selString = selString & "EdgeAndOS2Edge"
         End If
        Case FCAndBottomEdge, FCAndTopEdge
            selString = selString & "FCAndEdge"
        Case FCAndOSBottomEdge, FCAndOSTopEdge
         If TypeOf oBounded.MemberPart Is IJProfile Then
            selString = selString & "FCOS1Edg"
         Else
            selString = selString & "FCAndOS1Edge"
         End If
        Case FCAndOSBottom, FCAndOSTop
         If TypeOf oBounded.MemberPart Is IJProfile Then
            selString = selString & "FCOSNoEdg"
         Else
            selString = selString & "FCAndOSNoEdge"
         End If
        Case OnMemberTop, OnMemberBottom
            selString = selString & "OnMember"
        Case OSBottomAndOSTopEdge, OSTopAndOSBottomEdge
         If TypeOf oBounded.MemberPart Is IJProfile Then
            selString = selString & "OSOS1Edg"
         Else
            selString = selString & "OSAndOS1Edge"
         End If
        Case OSBottomEdgeAndOSTopEdge
         If TypeOf oBounded.MemberPart Is IJProfile Then
            selString = selString & "OSOS2Edg"
         Else
            selString = selString & "OSAndOS2Edge"
         End If
        Case OSBottomAndOSTop
         If TypeOf oBounded.MemberPart Is IJProfile Then
            selString = selString & "OSOSNoEdg"
         Else
            selString = selString & "OSAndOSNoEdge"
         End If
        Case OnTubeMember
         If TypeOf oBounded.MemberPart Is IJProfile Then
            selString = selString & "OnTubeMbr"
         Else
            selString = selString & "OnTubeMember"
         End If
        Case Else
            selString = vbNullString
    End Select
    
Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Function GetMemberBoundingCase(oACOrEC As Object, Optional bEndToEnd As Boolean, Optional bColinear As Boolean, Optional bForceComputeEdgeInfo As Boolean = True) As eMemberBoundingCase
    On Error GoTo ErrorHandler

    Const METHOD = "GetMemberBoundingCase"

    ' ------------------------------------
    ' Get information about the connection
    ' ------------------------------------
    Dim oAppConnection As IJAppConnection
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim lStatus As Long
    Dim sMsg As String
    Dim bPenetratesWeb As Boolean
    
    If TypeOf oACOrEC Is IJStructFeature Then
        Dim sACItemName As String
        Dim oACObj As Object
        AssemblyConnection_SmartItemName oACOrEC, sACItemName, oAppConnection
    Else
        Set oAppConnection = oACOrEC
    End If
    
    'Get Cached data
    If bForceComputeEdgeInfo = False Then
        On Error Resume Next
        Dim oAttributes As IJDAttributes
        Set oAttributes = oAppConnection
        GetMemberBoundingCase = 0
        GetMemberBoundingCase = oAttributes.CollectionOfAttributes("IJUAMbrACCacheStorage2").Item("BoundingCase").value
        Err.Clear
        On Error GoTo ErrorHandler
        
        If Not GetMemberBoundingCase = 0 Then
            Exit Function
        End If
    End If
    
    GetMemberBoundingCase = Unknown

    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort

    If oBoundingPort Is Nothing Then
        Exit Function
    ElseIf TypeOf oBoundingPort.Connectable Is IJPlate Then
        Exit Function
    ElseIf TypeOf oBoundingPort.Connectable Is SPSSlabEntity Then
        Exit Function
    ElseIf TypeOf oBoundingPort.Connectable Is SPSWallPart Then
        Exit Function
    ElseIf TypeOf oBoundingPort.Connectable Is ISPSDesignedMember Then
        Exit Function
    End If
    
    Dim bRightAngle As Boolean
    
    CheckEndToEndConnection oBoundedPort.Connectable, oBoundingPort.Connectable, bEndToEnd, bColinear, bRightAngle
    
    If bColinear Then
        Exit Function
    End If

    ' -----------------------------------------------------------------
    ' If the bounding member is tubular, use separate method and return
    ' -----------------------------------------------------------------
    If IsTubularMember(oBoundingPort.Connectable) Then
        bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
        GetMemberBoundingCase = GetBoundingCaseForTube(oBoundedPort.Connectable, oBoundingPort.Connectable, bPenetratesWeb, , oACOrEC)
        Exit Function
    End If

    ' -------------------------------------------------------------------
    ' Get the relative positions of the bounded object and bounding edges
    ' -------------------------------------------------------------------
    Dim oTopORWL As ConnectedEdgeInfo
    Dim oBottomOrWR As ConnectedEdgeInfo
    Dim oInsideTopFlgOrFL As ConnectedEdgeInfo
    Dim oInsideBtmFlgOrFR As ConnectedEdgeInfo
    
    GetConnectedEdgeInfo oACOrEC, oBoundedPort, oBoundingPort, oTopORWL, oBottomOrWR, oInsideTopFlgOrFL, oInsideBtmFlgOrFR, , , , , , bForceComputeEdgeInfo

    ' --------------------
    ' Get the edge mapping
    ' --------------------
    ' Do not force recompute.  GetConnectedEdgeInfo (called above) also calls GetEdgeMap
    ' and passes the recompute flag. If we pass bForceComputeEdgeInfo and it is set to True,
    ' we will calculate the edge map twice - once in GetConnectedEdgeInfo and once here.
    Dim ppEdgeMap As JCmnShp_CollectionAlias
    Set ppEdgeMap = GetEdgeMap(oACOrEC, oBoundingPort, _
                               oBoundedPort, , bPenetratesWeb)

    ' -------------------------------------
    ' Determine which bounded flanges exist
    ' -------------------------------------
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR
    
    ' -----------------
    ' If web-penetrated
    ' -----------------
    Dim oEdge1 As ConnectedEdgeInfo
    Dim oEdge2 As ConnectedEdgeInfo
    
    If bPenetratesWeb Then
        ' --------------------------------------------------------------------------------------
        ' First edge is the one intersected by Top, Second edge is the one intersected by Bottom
        ' --------------------------------------------------------------------------------------
        oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
        oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
    ' --------------------
    ' If flange-penetrated
    ' --------------------
    Else
        ' -----------------------------------------------------
        ' If the section has a full flange on the top or bottom
        ' -----------------------------------------------------
        If (bTFL And bTFR) Or (bBFR And bBFL) Then
            ' --------------------------------------------------------------------------------------------------
            ' First edge is the one intersected by FlangeLeft, Second edge is the one intersected by FlangeRight
            ' --------------------------------------------------------------------------------------------------
             oEdge1.IntersectingEdge = oInsideTopFlgOrFL.IntersectingEdge
             oEdge2.IntersectingEdge = oInsideBtmFlgOrFR.IntersectingEdge
        ' ----------------------------------------------------------
        ' If the section has only a left flange on the top or bottom
        ' ----------------------------------------------------------
        ElseIf (bTFL Or bBFL) And Not bBFR And Not bTFR Then
            ' -----------------------------------------------------------------------------------------------
            ' First edge is the one intersected by FlangeLeft, Second edge is the one intersected by WebRight
            ' -----------------------------------------------------------------------------------------------
             oEdge1.IntersectingEdge = oInsideTopFlgOrFL.IntersectingEdge
             oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
        ' -----------------------------------------------------------
        ' If the section has only a right flange on the top or bottom
        ' -----------------------------------------------------------
        ElseIf (bTFR Or bBFR) And Not bTFL And Not bBFL Then
            ' -----------------------------------------------------------------------------------------------
            ' First edge is the one intersected by WebLeft, Second edge is the one intersected by FlangeRight
            ' -----------------------------------------------------------------------------------------------
             oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
             oEdge2.IntersectingEdge = oInsideBtmFlgOrFR.IntersectingEdge
        ' ------------------------------
        ' If the section has not flanges
        ' ------------------------------
        ElseIf Not bTFL And Not bBFL And Not bTFR And Not bBFR Then
            ' --------------------------------------------------------------------------------------------
            ' First edge is the one intersected by WebLeft, Second edge is the one intersected by WebRight
            ' --------------------------------------------------------------------------------------------
             oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
             oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
            
        Else
            '??
        End If
    End If
    
    ' -----------------
    ' If web-penetrated
    ' -----------------
    Dim bCheckCutFeasibility As Boolean
    
    If bPenetratesWeb Then
        
        ' -----------------------------------------------------------
        ' If the outer-most bounded edge is above the bounding object
        ' -----------------------------------------------------------
        If oEdge1.IntersectingEdge = eBounding_Edge.Above Then
            
            ' -----------------------------------------------------------------------------------------------
            ' Check for sufficient material between the inner bounded edge and the top of the bounding object
            ' -----------------------------------------------------------------------------------------------
            If oInsideTopFlgOrFL.IntersectingEdge = eBounding_Edge.Above Then
                'Inside top is above: so check if enough web material is available
                bCheckCutFeasibility = BoundedHasOutsideMaterial(oACOrEC, eCutFlag.TopCut)
            Else
                'Top web cut is not feasible
                bCheckCutFeasibility = False
            End If
            
            ' -------------------------------------------------------------------------------------------------------------------------------
            ' If insufficient, change the intersected edge to flange-right (if bounding is flanged) or web-right (if bounding is not flanged)
            ' -------------------------------------------------------------------------------------------------------------------------------
            If bCheckCutFeasibility = False Then
                If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Then
                    oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right
                Else
                    oEdge1.IntersectingEdge = eBounding_Edge.Web_Right
                End If
            End If
        End If
        
        ' -----------------------------------------------------------
        ' If the outer-most bounded edge is below the bounding object
        ' -----------------------------------------------------------
        If oEdge2.IntersectingEdge = eBounding_Edge.Below Then
        
            ' --------------------------------------------------------------------------------------------------
            ' Check for sufficient material between the inner bounded edge and the bottom of the bounding object
            ' --------------------------------------------------------------------------------------------------
            If oInsideBtmFlgOrFR.IntersectingEdge = eBounding_Edge.Below Then
                'Inside bottom is below: so check if enough web material is available
                bCheckCutFeasibility = BoundedHasOutsideMaterial(oACOrEC, eCutFlag.BottomCut)
            Else
                'Bottom web cut is not feasible
                bCheckCutFeasibility = False
            End If
            
            ' -------------------------------------------------------------------------------------------------------------------------------
            ' If insufficient, change the intersected edge to flange-right (if bounding is flanged) or web-right (if bounding is not flanged)
            ' -------------------------------------------------------------------------------------------------------------------------------
            If bCheckCutFeasibility = False Then
                If ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
                    oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right
                Else
                    oEdge2.IntersectingEdge = eBounding_Edge.Web_Right
                End If
            End If
        End If
    End If

    ' -------------------------------------------------
    ' Identify case based on adjusted intersected edges
    ' -------------------------------------------------
    If oEdge1.IntersectingEdge = eBounding_Edge.Above And oEdge2.IntersectingEdge = eBounding_Edge.Above Then
    
        GetMemberBoundingCase = OnMemberTop
    
    ElseIf (oEdge1.IntersectingEdge = eBounding_Edge.Above And oEdge2.IntersectingEdge = eBounding_Edge.Top) Or _
           (oEdge2.IntersectingEdge = eBounding_Edge.Above And oEdge1.IntersectingEdge = eBounding_Edge.Top) Then
         
         GetMemberBoundingCase = OnMemberTop
    
    ElseIf ((oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            oEdge2.IntersectingEdge = eBounding_Edge.Above) Or _
           ((oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            oEdge1.IntersectingEdge = eBounding_Edge.Above) Or _
           ((oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) Then

       GetMemberBoundingCase = TopEdge
    
    ElseIf oEdge1.IntersectingEdge = eBounding_Edge.Web_Right And oEdge2.IntersectingEdge = eBounding_Edge.Above Or _
           oEdge2.IntersectingEdge = eBounding_Edge.Web_Right And oEdge1.IntersectingEdge = eBounding_Edge.Above Then
        
        If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Then
            GetMemberBoundingCase = FCAndOSTopEdge
        Else
            GetMemberBoundingCase = FCAndOSTop
        End If
    
    ElseIf ((oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top) And _
             oEdge2.IntersectingEdge = eBounding_Edge.Above) Or _
           ((oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top) And _
             oEdge1.IntersectingEdge = eBounding_Edge.Above) Then

        If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = BottomEdgeAndOSTopEdge
        
        ElseIf (ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap)) Or _
                (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) Then
            
            If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Then
                GetMemberBoundingCase = BottomEdgeAndOSTop
            Else
                GetMemberBoundingCase = TopEdgeAndOSBottom
            End If
        
        End If
    
    ElseIf ((oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top) And _
             oEdge2.IntersectingEdge = eBounding_Edge.Top) Or _
           ((oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top) And _
             oEdge1.IntersectingEdge = eBounding_Edge.Top) Then
        
        If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = BottomEdgeAndOSTopEdge
        
        ElseIf (ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap)) Or _
               (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) Then
            
            If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Then
                GetMemberBoundingCase = BottomEdgeAndOSTop
            Else
                GetMemberBoundingCase = TopEdgeAndOSBottom
            End If
        
        End If
    
    ElseIf oEdge1.IntersectingEdge = eBounding_Edge.Below And oEdge2.IntersectingEdge = eBounding_Edge.Above Or _
           oEdge2.IntersectingEdge = eBounding_Edge.Below And oEdge1.IntersectingEdge = eBounding_Edge.Above Then
        
        If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = OSBottomEdgeAndOSTopEdge
        
        ElseIf (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Or ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) And Not _
               (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) Then
            
            If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Then
                GetMemberBoundingCase = OSBottomAndOSTopEdge
            Else
                GetMemberBoundingCase = OSTopAndOSBottomEdge
            End If
        
        ElseIf Not ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = OSBottomAndOSTop
        
        End If
    
    ElseIf oEdge1.IntersectingEdge = eBounding_Edge.Bottom And oEdge2.IntersectingEdge = eBounding_Edge.Top Or _
           oEdge2.IntersectingEdge = eBounding_Edge.Bottom And oEdge1.IntersectingEdge = eBounding_Edge.Top Then
        
        If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = OSBottomEdgeAndOSTopEdge
        
        ElseIf (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Or ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) And Not _
               (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) Then
            
            If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Then
                GetMemberBoundingCase = OSBottomAndOSTopEdge
            Else
                GetMemberBoundingCase = OSTopAndOSBottomEdge
            End If
        
        ElseIf Not ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = OSBottomAndOSTop
        
        End If
    
    ElseIf ((oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            oEdge2.IntersectingEdge = eBounding_Edge.Web_Right) Or _
           ((oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            oEdge1.IntersectingEdge = eBounding_Edge.Web_Right) Then
        
        GetMemberBoundingCase = FCAndTopEdge
    
    ElseIf ((oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            (oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top)) Or _
           ((oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            (oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top)) Then
            
            GetMemberBoundingCase = BottomEdgeAndTopEdge
    
    ElseIf ((oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            oEdge2.IntersectingEdge = eBounding_Edge.Below) Or _
           ((oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            oEdge1.IntersectingEdge = eBounding_Edge.Below) Then

        If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = TopEdgeAndOSBottomEdge
        
        ElseIf (ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap)) Or _
               (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) Then
            
            GetMemberBoundingCase = TopEdgeAndOSBottom
        
        End If
    
    ElseIf ((oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            oEdge2.IntersectingEdge = eBounding_Edge.Bottom) Or _
           ((oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom) And _
            oEdge1.IntersectingEdge = eBounding_Edge.Bottom) Then
        
        If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = TopEdgeAndOSBottomEdge
        
        ElseIf (ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap)) Or _
               (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) Then
            
            GetMemberBoundingCase = TopEdgeAndOSBottom
        
        End If
    
    ElseIf oEdge1.IntersectingEdge = eBounding_Edge.Web_Right And oEdge2.IntersectingEdge = eBounding_Edge.Web_Right Then
        
        GetMemberBoundingCase = Center
    
    ElseIf ((oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top) And _
            oEdge2.IntersectingEdge = eBounding_Edge.Web_Right) Or _
           ((oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top) And _
            oEdge1.IntersectingEdge = eBounding_Edge.Web_Right) Then
        
        GetMemberBoundingCase = FCAndBottomEdge
    
    ElseIf oEdge1.IntersectingEdge = eBounding_Edge.Below And oEdge2.IntersectingEdge = eBounding_Edge.Web_Right Or _
           oEdge2.IntersectingEdge = eBounding_Edge.Below And oEdge1.IntersectingEdge = eBounding_Edge.Web_Right Then

        If ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            GetMemberBoundingCase = FCAndOSBottomEdge
        Else
            GetMemberBoundingCase = FCAndOSBottom
        End If
    
    ElseIf ((oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top) And _
            oEdge2.IntersectingEdge = eBounding_Edge.Below) Or _
           ((oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top) And _
            oEdge1.IntersectingEdge = eBounding_Edge.Below) Or _
           ((oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top) And _
            (oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top)) Then

        GetMemberBoundingCase = BottomEdge
    
    ElseIf oEdge1.IntersectingEdge = eBounding_Edge.Below And oEdge2.IntersectingEdge = eBounding_Edge.Below Then

        GetMemberBoundingCase = OnMemberBottom
    
    ElseIf (oEdge1.IntersectingEdge = eBounding_Edge.Below And oEdge2.IntersectingEdge = eBounding_Edge.Bottom) Or (oEdge2.IntersectingEdge = eBounding_Edge.Below And oEdge1.IntersectingEdge = eBounding_Edge.Bottom) Then

        GetMemberBoundingCase = OnMemberBottom
    
    ElseIf (oEdge1.IntersectingEdge = eBounding_Edge.Web_Right And oEdge2.IntersectingEdge = eBounding_Edge.Bottom) Or _
           (oEdge2.IntersectingEdge = eBounding_Edge.Web_Right And oEdge1.IntersectingEdge = eBounding_Edge.Bottom) Then
        
        If ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            GetMemberBoundingCase = FCAndOSBottomEdge
        Else
            GetMemberBoundingCase = FCAndOSBottom
        End If
    
    ElseIf oEdge1.IntersectingEdge = eBounding_Edge.Web_Right And oEdge2.IntersectingEdge = eBounding_Edge.Top Or _
           oEdge2.IntersectingEdge = eBounding_Edge.Web_Right And oEdge1.IntersectingEdge = eBounding_Edge.Top Then

        If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Then
            GetMemberBoundingCase = FCAndOSTopEdge
        Else
            GetMemberBoundingCase = FCAndOSTop
        End If
    
    ElseIf oEdge1.IntersectingEdge = eBounding_Edge.Above And oEdge2.IntersectingEdge = eBounding_Edge.Bottom Or _
           oEdge1.IntersectingEdge = eBounding_Edge.Top And oEdge2.IntersectingEdge = eBounding_Edge.Below Then
        
        If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = OSBottomEdgeAndOSTopEdge
        
        ElseIf (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Or ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) And Not _
               (ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap)) Then
            
            If ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) Then
                GetMemberBoundingCase = OSBottomAndOSTopEdge
            Else
                GetMemberBoundingCase = OSTopAndOSBottomEdge
            End If
        
        ElseIf Not ItemExists(JXSEC_TOP_FLANGE_RIGHT, ppEdgeMap) And Not ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, ppEdgeMap) Then
            
            GetMemberBoundingCase = OSBottomAndOSTop
        
        End If
    End If
    
    On Error Resume Next
    'Set Attribute value
    Set oAttributes = oAppConnection
    oAttributes.CollectionOfAttributes("IJUAMbrACCacheStorage2").Item("BoundingCase").value = GetMemberBoundingCase
    Err.Clear

Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

' If the input web cut is not penetrated, the method will determine where on the bounding member it intersects
' If the input web is penetrated, the method will determine where the bounded flange intersects.  The user must specify which flange.
Public Sub GetNonPenetratedIntersectedEdge(oACOrEC As Object, _
                                            oBoundingPort As IJPort, _
                                            oBoundedPort As IJPort, _
                                            boundingEdgeId As eBounding_Edge, _
                                            mappedEdge As JXSEC_CODE, _
                                            IsBottomFlange As Boolean)

    Const METHOD = "GetNonPenetratedIntersectedEdge"
    Dim sMsg As String
    sMsg = ""
    
    boundingEdgeId = None
    mappedEdge = JXSEC_UNKNOWN
    
    On Error GoTo ErrorHandler

    ' ---------------------
    ' Get the section alias
    ' ---------------------
    Dim sectionAlias As Long
    Dim oEdgeMap As JCmnShp_CollectionAlias
    Dim bPenetratesWeb As Boolean
    Set oEdgeMap = GetEdgeMap(oACOrEC, oBoundingPort, oBoundedPort, sectionAlias, bPenetratesWeb)

    Dim boundingAlias As eBounding_Alias
    boundingAlias = GetBoundingAliasSimplified(sectionAlias)

    ' ----------------------------------
    ' Get the connected edge information
    ' ----------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
    
    Dim bUseTopFlange As Boolean
    
    bUseTopFlange = Not IsBottomFlange 'If bottom flange is True, set appropriate boolean value for the variable
    GetConnectedEdgeInfo oACOrEC, oBoundedPort, oBoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR, , , , , bUseTopFlange
                         
    Dim outsideEdgeId As eBounding_Edge
    Dim insideEdgeId As eBounding_Edge
    Dim topEdgeId As eBounding_Edge
    Dim bottomEdgeId As eBounding_Edge
    
    If Not bPenetratesWeb Then
        outsideEdgeId = eTopOrWL.IntersectingEdge
        insideEdgeId = eBottomOrWR.IntersectingEdge
        topEdgeId = eTopOrWL.IntersectingEdge
        bottomEdgeId = eBottomOrWR.IntersectingEdge
    ElseIf IsBottomFlange Then
        outsideEdgeId = eBottomOrWR.IntersectingEdge
        insideEdgeId = eInsideBFOrTFR.IntersectingEdge
        topEdgeId = eInsideBFOrTFR.IntersectingEdge
        bottomEdgeId = eBottomOrWR.IntersectingEdge
    Else
        outsideEdgeId = eTopOrWL.IntersectingEdge
        insideEdgeId = eInsideTFOrTFL.IntersectingEdge
        topEdgeId = eTopOrWL.IntersectingEdge
        bottomEdgeId = eInsideTFOrTFL.IntersectingEdge
    End If

    ' --------------------------------------------------------------
    ' If either intersect a bounding edge, use that as bounding port
    ' --------------------------------------------------------------
    If outsideEdgeId = Bottom_Flange_Right Or outsideEdgeId = Top_Flange_Right Then
        boundingEdgeId = outsideEdgeId
    ElseIf insideEdgeId = Bottom_Flange_Right Or insideEdgeId = Top_Flange_Right Then
        boundingEdgeId = insideEdgeId
    End If
    
    ' --------------------------------------------------------------
    ' If the two straddle a bounding edge, use that as bounding port
    ' --------------------------------------------------------------
    ' Skip check if already set
    If boundingEdgeId = None Then
        ' Skip check for top if there is no top edge
        Select Case boundingAlias
            Case WebTopFlangeRight, WebTopAndBottomRightFlanges, WebBuiltUpTopFlangeRight
                ' If topEdgeId is above the top bounding edge, then check the bottom edge to see if it is below it.
                Select Case topEdgeId
                    Case eBounding_Edge.Above, eBounding_Edge.Top, eBounding_Edge.Web_Right_Top, eBounding_Edge.Top_Flange_Right_Top
                        Select Case bottomEdgeId
                            Case eBounding_Edge.Above, eBounding_Edge.Top, eBounding_Edge.Web_Right_Top, _
                                 eBounding_Edge.Top_Flange_Right_Top, eBounding_Edge.Top_Flange_Right
                            Case Else
                                boundingEdgeId = Top_Flange_Right
                        End Select
                End Select
        End Select
    End If
    
    ' Skip check if already set
    If boundingEdgeId = None Then
        ' Skip check for top if there is no top edge
        Select Case boundingAlias
            Case WebBottomFlangeRight, WebTopAndBottomRightFlanges, WebBuiltUpBottomFlangeRight
                ' If bottomEdgeId is below the bottom bounding edge, then check the top edge to see if it is above it.
                Select Case bottomEdgeId
                    Case eBounding_Edge.Below, eBounding_Edge.Bottom, eBounding_Edge.Web_Right_Bottom, eBounding_Edge.Bottom_Flange_Right_Bottom
                        Select Case topEdgeId
                            Case eBounding_Edge.Below, eBounding_Edge.Bottom, eBounding_Edge.Web_Right_Bottom, _
                                 eBounding_Edge.Bottom_Flange_Right_Bottom, eBounding_Edge.Bottom_Flange_Right
                            Case Else
                                boundingEdgeId = Bottom_Flange_Right
                        End Select
                End Select
        End Select
    End If
    
    ' -------------------------------------------------------------
    ' Otherwise, use the surface intersected by the outside surface
    ' -------------------------------------------------------------
    If boundingEdgeId = None Then
        Select Case outsideEdgeId
            Case eBounding_Edge.Above, eBounding_Edge.Below, eBounding_Edge.None
            Case Else
                boundingEdgeId = outsideEdgeId
        End Select
    End If
    
    ' ------------------------------------------------------------------
    ' If the outside surface does not intersect anything, use the inside
    ' ------------------------------------------------------------------
    If boundingEdgeId = None Then
        Select Case insideEdgeId
            Case eBounding_Edge.Above, eBounding_Edge.Below, eBounding_Edge.None
            Case Else
                boundingEdgeId = insideEdgeId
        End Select
    End If
    
    ' ----------------------------------------------------------------
    ' If both inside and outside are above or below, return that value
    ' ----------------------------------------------------------------
    If outsideEdgeId = eBounding_Edge.Above And insideEdgeId = eBounding_Edge.Above Then
        boundingEdgeId = eBounding_Edge.Above
    ElseIf outsideEdgeId = eBounding_Edge.Below And insideEdgeId = eBounding_Edge.Below Then
        boundingEdgeId = eBounding_Edge.Below
    End If
    
    ' --------------------------------------------------------------------
    ' The edge found is a standin for a mapped edge.  Get the mapped edge.
    ' --------------------------------------------------------------------
    If boundingEdgeId <> None And boundingEdgeId <> eBounding_Edge.Above And boundingEdgeId <> eBounding_Edge.Below Then
        mappedEdge = oEdgeMap.Item(CStr(boundingEdgeId))
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Function GetBoundingAliasSimplified(mapRuleAlias As Long) As eBounding_Alias

    GetBoundingAliasSimplified = 0

    Select Case mapRuleAlias
        Case 0, 8, 9, 10, 11, 12, 17
            'Web, WebTopFlangeLeft, WebBuiltUpTopFlangeLeft, WebBottomFlangeLeft
            'WebBuiltUpBottomFlangeLeft, WebTopAndBottomLeftFlanges, TwoWebsTwoFlanges,
            GetBoundingAliasSimplified = eBounding_Alias.Web
        Case 1, 6
            'WebTopFlangeRight, WebTopFlange
            GetBoundingAliasSimplified = eBounding_Alias.WebTopFlangeRight
        Case 2
            'WebBuiltUpTopFlangeRight
            GetBoundingAliasSimplified = eBounding_Alias.WebBuiltUpTopFlangeRight
        Case 3, 7
            'WebBottomFlangeRight, WebBottomFlange
            GetBoundingAliasSimplified = eBounding_Alias.WebBottomFlangeRight
        Case 4
            'WebBuiltUpBottomFlangeRight
            GetBoundingAliasSimplified = eBounding_Alias.WebBuiltUpBottomFlangeRight
        Case 5, 13, 19
            'WebTopAndBottomRightFlanges, WebTopAndBottomFlanges, TwoWebsBetweenFlanges
            GetBoundingAliasSimplified = eBounding_Alias.WebTopAndBottomRightFlanges
        Case 14
            'FlangeLeftAndRightBottomWebs
            GetBoundingAliasSimplified = eBounding_Alias.FlangeLeftAndRightBottomWebs
        Case 15
            'FlangeLeftAndRightTopWebs
            GetBoundingAliasSimplified = eBounding_Alias.FlangeLeftAndRightTopWebs
        Case 16, 18
            'FlangeLeftAndRightWebs, TwoFlangesBetweenWebs
            GetBoundingAliasSimplified = eBounding_Alias.FlangeLeftAndRightWebs
        Case 20
            'Tube/Circular Cross Section
            GetBoundingAliasSimplified = eBounding_Alias.Tube
        Case Else
            'Unknown Section Alias
    End Select
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetBoundingAliasSimplified").Number
End Function

'*********************************************************************************************
' Method      : BoundedHasOutsideMaterial
' Description : This method takes selector logic smart occurence as input and
'               returns determines if the bounded object has material outside the bounding object
'
'  optional input TopOrBottomFlag values can be as follows:
'       1 - if only Top of bounded is to be checked
'       2 - if only Bottom of bounded is to be checked
'  If no value is specified, function checks for either Top or Bottom
'  "Top" is relative to the bounded object in the 2D sketch.  For flange-penetrated cases, this is flange-left
'*********************************************************************************************

Public Function BoundedHasOutsideMaterial(oSmartOccurence As Object, _
                                          Optional TopOrBottomFlag As eCutFlag = eCutFlag.TopOrBottomCut, _
                                          Optional dTol As Double = 0.0005, _
                                          Optional bIsTopFlange As Boolean) As Boolean
    
    BoundedHasOutsideMaterial = False

    ' --------------------------------
    ' Get bounding and bounded objects
    ' --------------------------------
    Dim oAppConnection As IJAppConnection
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim lStatus As Long
    Dim sMsg As String
    Dim bPenetratesWeb As Boolean
    
    If TypeOf oSmartOccurence Is IJStructFeature Then
        Dim sACItemName As String
        Dim oACObj As Object
        
        Parent_SmartItemName oSmartOccurence, sACItemName, oAppConnection
    Else
        Set oAppConnection = oSmartOccurence
    End If
    
    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
                         
                   
    If (IsTubularMember(oBoundingPort.Connectable) And (Not (IsTubularMember(oBoundedPort.Connectable)))) Then
        ' ----------------------------------------------------------------
        ' If the Bounding Object is a Tube, then intersect the Tube Ports
        ' with the Bounded Object Ports to determine if the bounded object
        ' has material outside of the tube.
        ' ----------------------------------------------------------------
        
        ' -----------------
        ' Declare Variables
        ' -----------------
        Dim bTopFlangeLeft As Boolean
        Dim bTopFlangeRight As Boolean
        Dim bBottomFlangeLeft As Boolean
        Dim bBottomFlangeRight As Boolean
    
        Dim oSDO_Bounded As New StructDetailObjects.MemberPart
        Dim oSDO_Bounding As New StructDetailObjects.MemberPart
        
        Dim oBounded As Object
        Dim oBounding As Object
    
        Dim dDistance As Double
        Dim oBoundingOuterPort As IJPort

        Dim oReferencePort As Object
        Dim oBoundedWireBdy(1 To 4) As IJWireBody
        Dim oPort As Object
        Dim pAgtorUnk As IUnknown
        Dim NullObject As Object
        Dim IntersectedObject As Object
        Dim oGenericIntersector As IMSModelGeomOps.DGeomOpsIntersect
        Dim oWLport As Object
        Dim oWRPort As Object
        Dim oModelBodyUtils As IJSGOModelBodyUtilities
        Dim oPointOnBounding As IJDPosition
        Dim oPointOnBounded As IJDPosition

        Dim iCount As Integer
        
        ' -----------------------------
        ' Prepare Bounding Surface Body
        ' Set Outer Port Surface
        ' -----------------------------
        Set oBounded = oBoundedPort.Connectable
        Set oBounding = oBoundingPort.Connectable
        
        Dim oTempBounded As Object
        Dim oTempBounding As Object
        
        Set oTempBounded = oBounded
        Set oTempBounding = oBounding
        
        Dim oBUMember As ISPSDesignedMember
        Dim bBuiltup As Boolean
    
        If Not TypeOf oBounded Is ISPSMemberPartCommon Then
            IsFromBuiltUpMember oBounded, bBuiltup, oBUMember
            If bBuiltup Then
                Set oBounded = oBUMember
            Else
                Exit Function    'Inputs to this method should be of type ispsmemberpartcommon
            End If
        Else
            Set oSDO_Bounded.object = oBounded
        End If
        
        bBuiltup = False
        Set oBUMember = Nothing
    
        If Not TypeOf oBounding Is ISPSMemberPartCommon Then
            IsFromBuiltUpMember oBounding, bBuiltup, oBUMember
            If bBuiltup Then
                Set oBounding = oBUMember
            Else
                Exit Function    'Inputs to this method should be of type ispsmemberpartcommon
            End If
        Else
            Set oSDO_Bounding.object = oBounding
        End If

        If TypeOf oBounding Is ISPSDesignedMember Then
            Set oBoundingOuterPort = GetBaseOffsetOrLateralPortForObject(oTempBounding, BPT_Base)
        ElseIf TypeOf oBounding Is ISPSMemberPartCommon Then
            Set oBoundingOuterPort = GetLateralSubPortBeforeTrim(oSDO_Bounding.object, JXSEC_OUTER_TUBE)
        Else
            Exit Function
        End If

        ' --------------------------------------------------------------
        ' Prepare Bounded Wire Bodies: Need Four Intersection Curves
        ' on bounded member - TopLeft, TopRight, BottomLeft, BottomRight
        ' --------------------------------------------------------------
        
        'Set bounded member SDO wrapper
        Set oWLport = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_WEB_LEFT))
        Set oWRPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_WEB_RIGHT))
        
        'Need to determine which of the Top/Bottom flanges exist
        CrossSection_Flanges oSDO_Bounded.object, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
        
        'Reference port is top port
        Set oReferencePort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_TOP))
        
        '2a: Top-left
        Set oPort = oWLport
        If bTopFlangeLeft Then Set oPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_TOP_FLANGE_LEFT))
        Set oGenericIntersector = New IMSModelGeomOps.DGeomOpsIntersect
        oGenericIntersector.PlaceIntersectionObject NullObject, oReferencePort, oPort, _
                                                        pAgtorUnk, IntersectedObject
        If Not IntersectedObject Is Nothing Then Set oBoundedWireBdy(1) = IntersectedObject
                                                     
        '2b: top-right
        Set oPort = oWRPort
        If bTopFlangeRight Then Set oPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_TOP_FLANGE_RIGHT))
        oGenericIntersector.PlaceIntersectionObject NullObject, oReferencePort, oPort, _
                                                        pAgtorUnk, IntersectedObject
        If Not IntersectedObject Is Nothing Then Set oBoundedWireBdy(2) = IntersectedObject
    
        'Reference port is bottom port
        Set oReferencePort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_BOTTOM))
        
        '2c: bottom-left
        Set oPort = oWLport
        If bBottomFlangeLeft Then Set oPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_BOTTOM_FLANGE_LEFT))
        
        oGenericIntersector.PlaceIntersectionObject NullObject, oReferencePort, oPort, _
                                                        pAgtorUnk, IntersectedObject
        If Not IntersectedObject Is Nothing Then Set oBoundedWireBdy(3) = IntersectedObject
        
        '2d: bottom-right
        Set oPort = oWRPort
        If bBottomFlangeRight Then Set oPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_BOTTOM_FLANGE_RIGHT))
        oGenericIntersector.PlaceIntersectionObject NullObject, oReferencePort, oPort, _
                                                        pAgtorUnk, IntersectedObject
        If Not IntersectedObject Is Nothing Then Set oBoundedWireBdy(4) = IntersectedObject
        
        Set oModelBodyUtils = New SGOModelBodyUtilities

        ' --------------------------------------------------------------
        ' Get the Minimum Distance Between the four boundary curves and
        ' the bounding outer surface
        ' --------------------------------------------------------------
        Dim dDistanceToBounding(1 To 4) As Double
        'Get bounding to bounded vector
        For iCount = 1 To 4
            oModelBodyUtils.GetClosestPointsBetweenTwoBodies oBoundingOuterPort.Geometry, _
                    oBoundedWireBdy(iCount), oPointOnBounding, oPointOnBounded, dDistance
            dDistanceToBounding(iCount) = dDistance
        Next iCount
        
        bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
        
        If bPenetratesWeb Then
            Dim dTopFlgThk As Double
            Dim dBtmFlgThk As Double
            ' ----------------
            ' If for a top cut
            ' ----------------
            dTopFlgThk = 0
            If bTopFlangeLeft Or bTopFlangeRight Then dTopFlgThk = oSDO_Bounded.flangeThickness
                
            dBtmFlgThk = 0
            If bBottomFlangeLeft Or bBottomFlangeRight Then dBtmFlgThk = oSDO_Bounded.flangeThickness
            If TopOrBottomFlag = eCutFlag.TopCut Or TopOrBottomFlag = eCutFlag.TopOrBottomCut Then
                If dDistanceToBounding(1) - dTopFlgThk > dTol Or dDistanceToBounding(2) - dTopFlgThk > dTol Then
                    BoundedHasOutsideMaterial = True
                End If
                
            ' -------------------
            ' If for a bottom cut
            ' -------------------
            ElseIf TopOrBottomFlag = eCutFlag.BottomCut Or TopOrBottomFlag = eCutFlag.TopOrBottomCut Then
                If dDistanceToBounding(3) - dBtmFlgThk > dTol Or dDistanceToBounding(4) - dBtmFlgThk > dTol Then
                    BoundedHasOutsideMaterial = True
                End If
            End If
        Else
            ' ----------------
            ' If for a top cut
            ' ----------------
            If TopOrBottomFlag = eCutFlag.TopCut Or TopOrBottomFlag = eCutFlag.TopOrBottomCut Then
                If dDistanceToBounding(1) > dTol Or dDistanceToBounding(3) > dTol Then
                    BoundedHasOutsideMaterial = True
                End If
                
            ' -------------------
            ' If for a bottom cut
            ' -------------------
            ElseIf TopOrBottomFlag = eCutFlag.BottomCut Or TopOrBottomFlag = eCutFlag.TopOrBottomCut Then
                If dDistanceToBounding(2) > dTol Or dDistanceToBounding(4) > dTol Then
                    BoundedHasOutsideMaterial = True
                End If
            End If
        End If
        

    Else
        ' -----------------------------------------------------------------
        ' Use the Measurement Symbol to determine if the bounded object has
        ' material outside of the bounding object.
        ' -----------------------------------------------------------------
    
        ' ----------------
        ' Get measurements
        ' ----------------
        Dim eTopOrWL As ConnectedEdgeInfo
        Dim eBottomOrWR As ConnectedEdgeInfo
        Dim eInsideTFOrTFL As ConnectedEdgeInfo
        Dim eInsideBFOrTFR As ConnectedEdgeInfo
        Dim oMeasurements As New Collection
    
        GetConnectedEdgeInfo oAppConnection, _
                             oBoundedPort, _
                             oBoundingPort, _
                             eTopOrWL, _
                             eBottomOrWR, _
                             eInsideTFOrTFL, _
                             eInsideBFOrTFR, _
                             oMeasurements, _
                             bPenetratesWeb, , , _
                             bIsTopFlange
            
        Dim dAvailableDistanceForCut As Double
        dAvailableDistanceForCut = -1#
            
        ' --------------------------------------------------
        ' Determine what flanges exist on the bounded object
        ' --------------------------------------------------
        Dim bTFL As Boolean
        Dim bBFL As Boolean
        Dim bTFR As Boolean
        Dim bBFR As Boolean
    
        CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR
        
        Dim bHasLeftFlange As Boolean
        Dim bHasRightFlange As Boolean
        bHasLeftFlange = False
        bHasRightFlange = False
        
        If bTFL Or bBFL Then
            bHasLeftFlange = True
        End If
        
        If bTFR Or bBFR Then
            bHasRightFlange = True
        End If
        
        ' ----------------
        ' If for a top cut
        ' ----------------
        If TopOrBottomFlag = eCutFlag.TopCut Or TopOrBottomFlag = eCutFlag.TopOrBottomCut Then
            ' ----------------------------------------------------------------------------------------------------------
            ' If web-penetrated and both the top and top-inside are outside the part, measure the distance to the inside
            ' ----------------------------------------------------------------------------------------------------------
            ' If there is no top flange, the distance to the inside will equal the distance to the top
            If bPenetratesWeb And (eTopOrWL.IntersectingEdge = eBounding_Edge.Above) And (eInsideTFOrTFL.IntersectingEdge = eBounding_Edge.Above) Then
                dAvailableDistanceForCut = oMeasurements.Item("DimPt11ToTopInside")
            ' ----------------------------------------------------------------------------------------------------
            ' If flange-penetrated with a left flange, if flange goes outside, measure the distance to flange-left
            ' ----------------------------------------------------------------------------------------------------
            ElseIf (Not bPenetratesWeb) And bHasLeftFlange And (eInsideTFOrTFL.IntersectingEdge = Above) Then
                dAvailableDistanceForCut = oMeasurements.Item("DimPt11ToFL")
            ' ---------------------------------------------------------------------------------------------------------------------
            ' If flange-penetrated without a left flange, if both WebLeft and WebRight go outside, measure the distance to WebRight
            ' ---------------------------------------------------------------------------------------------------------------------
            ElseIf (Not bPenetratesWeb) And (Not bHasLeftFlange) And _
                   (eTopOrWL.IntersectingEdge = eBounding_Edge.Above) And (eBottomOrWR.IntersectingEdge = eBounding_Edge.Above) Then
                dAvailableDistanceForCut = oMeasurements.Item("DimPt11ToWR")
            End If
        ' -------------------
        ' If for a bottom cut
        ' -------------------
        ElseIf TopOrBottomFlag = eCutFlag.BottomCut Or TopOrBottomFlag = eCutFlag.TopOrBottomCut Then
            ' ----------------------------------------------------------------------------------------------------------------
            ' If web-penetrated and both the bottom and bottom-inside are outside the part, measure the distance to the inside
            ' ----------------------------------------------------------------------------------------------------------------
            ' If there is no bottom flange, the distance to the inside will equal the distance to the bottom
            If bPenetratesWeb And (eBottomOrWR.IntersectingEdge = eBounding_Edge.Below) And (eInsideBFOrTFR.IntersectingEdge = eBounding_Edge.Below) Then
                dAvailableDistanceForCut = oMeasurements.Item("DimPt23ToBottomInside")
           ' ----------------------------------------------------------------------------------------------------
            ' If flange-penetrated with a right flange, if flange goes outside, measure the distance to flange-right
            ' ----------------------------------------------------------------------------------------------------
            ElseIf (Not bPenetratesWeb) And bHasRightFlange And (eInsideBFOrTFR.IntersectingEdge = Below) Then
                dAvailableDistanceForCut = oMeasurements.Item("DimPt23ToFR")
            ' ---------------------------------------------------------------------------------------------------------------------
            ' If flange-penetrated without a right flange, if both WebLeft and WebRight go outside, measure the distance to WebLeft
            ' ---------------------------------------------------------------------------------------------------------------------
            ElseIf (Not bPenetratesWeb) And (Not bHasRightFlange) And _
                   (eTopOrWL.IntersectingEdge = eBounding_Edge.Below) And (eBottomOrWR.IntersectingEdge = eBounding_Edge.Below) Then
                dAvailableDistanceForCut = oMeasurements.Item("DimPt23ToWL")
            End If
            
        End If
        
    End If
    
    ' -------------------------------------------------------
    ' If the measurement exceeds the tolerance, return "True"
    ' -------------------------------------------------------
    If (dAvailableDistanceForCut) > dTol Then
        BoundedHasOutsideMaterial = True
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "BoundedHasOutsideMaterial").Number
    
End Function

'*********************************************************************************************
' Method      : GetInfPlaneFromPort
' Description : Given port as input, returns infinite plane which contains the port
'
'*********************************************************************************************
Public Function GetInfPlaneFromPort(ByVal oPort As IJPort) As IJPlane
    On Error GoTo ErrorHandler
    
    If Not TypeOf oPort Is IJPlane Then
        GoTo ErrorHandler
    End If
    
    Dim oTmpPlane As IJPlane
    Set oTmpPlane = oPort
    
    Dim dNormX As Double
    Dim dNormY As Double
    Dim dNormZ As Double
    Dim dRootPtX As Double
    Dim dRootPtY As Double
    Dim dRootPtZ As Double
    oTmpPlane.GetRootPoint dRootPtX, dRootPtY, dRootPtZ
    oTmpPlane.GetNormal dNormX, dNormY, dNormZ
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Set GetInfPlaneFromPort = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, dRootPtX, dRootPtY, dRootPtZ, _
                dNormX, dNormY, dNormZ)
                
    Set oGeomFactory = Nothing
    Set oTmpPlane = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetInfPlaneFromPort").Number
    
End Function

'*********************************************************************************************
' Method      : GetDistanceFromBounding
' Description : Gets distance between tube member and bounded member port
' specified by cross section code
' Optionally closest points on bounding and bouded Dpositions are returned.
'*********************************************************************************************

Public Function GetDistanceFromBounding(oBoundingSmartOcc As Object, _
    oBoundedSmartOcc As Object, oCrossSecCode As JXSEC_CODE, _
    Optional oClosestPtOnBounding As IJDPosition, _
    Optional oClosestPtOnBounded As IJDPosition) As Double
    
    On Error GoTo ErrorHandler
    Const METHOD = "GetDistanceFromBounding"
    
    'Declare variables
    Dim oSDO_Bounding As StructDetailObjects.MemberPart
    Dim oBounded As Object
    Dim oBoundingOuterPort As IJPort
    Dim dDistance As Double
    Dim oRequiredPort As Object
            
    'Set bounding outer port surface
    Set oSDO_Bounding = New StructDetailObjects.MemberPart
    
    Dim bBuiltup As Boolean
    Dim oBUMember As ISPSDesignedMember
    
    If TypeOf oBoundingSmartOcc Is ISPSMemberPartCommon Then
        Set oSDO_Bounding.object = oBoundingSmartOcc
    Else
        IsFromBuiltUpMember oBoundingSmartOcc, bBuiltup, oBUMember
        If bBuiltup Then
            Set oSDO_Bounding.object = oBUMember
        End If
    End If
    
    If bBuiltup Then
        Set oBoundingOuterPort = GetBaseOffsetOrLateralPortForObject(oBoundingSmartOcc, BPT_Base)
    Else
        Set oBoundingOuterPort = GetLateralSubPortBeforeTrim(oBoundingSmartOcc, JXSEC_OUTER_TUBE)
    End If
    
    Dim oBoundingSurfBdy As IJSurfaceBody
    Set oBoundingSurfBdy = oBoundingOuterPort.Geometry

    'Set bounded member SDO wrapper
    Set oBounded = oBoundedSmartOcc
    
    'Set input port
    Set oRequiredPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oBounded, oCrossSecCode))
    
    'Below logic gets two intersection curves of the input port with
    ' with either of boundaries
    Dim oSurfBdy As IJSurfaceBody
    Set oSurfBdy = oRequiredPort
    
    Dim oModelBodyUtils As IJSGOModelBodyUtilities
    Set oModelBodyUtils = New SGOModelBodyUtilities
    
    Dim oPoint1 As IJDPosition
    Dim oPOint2 As IJDPosition
    
    oModelBodyUtils.GetClosestPointsBetweenTwoBodies oSurfBdy, _
         oBoundingSurfBdy, oPoint1, oPOint2, dDistance
    GetDistanceFromBounding = dDistance
    
    Set oClosestPtOnBounding = oPOint2
    Set oClosestPtOnBounded = oPoint1
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD

End Function


'*********************************************************************************************
' Method      : GetBoundingCaseForTube
' Description :
'
' Following cases are to be identified and selection string is returned:
' - To Center
' - Face-and-Outside
' - On Member
' - Outside-and-outside
'
'Inputs: To this method should be of type ISPSMemberPartCommon.
'Optional output:
' iACFlag: returns one of the following Const values for Border cases of 'ToCenter' AC:
'BorderAC_OSOS = 1
'BorderAC_FCOS_TOP = 2
'BorderAC_FCOS_BTM = 3
'BorderAC_ToCenter = 4
'NOT_BorderAC = 5

'Important Note: 'bIsWebPenetrated' value must be provided incase 'iACFlag' is needed by caller.
'
'*********************************************************************************************
Public Function GetBoundingCaseForTube(oBounded As Object, _
                                       oBounding As Object, Optional bPenetratesWeb As Boolean, Optional iACFlag As Integer = 0, _
                                       Optional oACOrEC As Object) As eMemberBoundingCase

    
    On Error GoTo ErrorHandler
    
    Const METHOD = "GetBoundingCaseForTube"
        
    ' --------------------------------------
    ' Check if bounding is a built-up member
    ' --------------------------------------

    Dim oSDO_Bounded As New StructDetailObjects.MemberPart
    Dim oSDO_Bounding As New StructDetailObjects.MemberPart
    
    
    Dim oTempBounded As Object
    Dim oTempBounding As Object
    
    Set oTempBounded = oBounded
    Set oTempBounding = oBounding
    
    Dim oBUMember As ISPSDesignedMember
    Dim bBuiltup As Boolean

    If Not TypeOf oTempBounded Is ISPSMemberPartCommon Then
        IsFromBuiltUpMember oTempBounded, bBuiltup, oBUMember
        If bBuiltup Then
            Set oTempBounded = oBUMember
        Else
            Exit Function    'Inputs to this method should be of type ispsmemberpartcommon
        End If
    Else
        Set oSDO_Bounded.object = oTempBounded
    End If
    
    bBuiltup = False
    Set oBUMember = Nothing

    If Not TypeOf oTempBounding Is ISPSMemberPartCommon Then
        IsFromBuiltUpMember oTempBounding, bBuiltup, oBUMember
        If bBuiltup Then
            Set oTempBounding = oBUMember
        Else
            Exit Function    'Inputs to this method should be of type ispsmemberpartcommon
        End If
    Else
        Set oSDO_Bounding.object = oTempBounding
    End If
    
    If IsTubularMember(oTempBounded) Then
    
    ' Logic to identify bounding conditions when bounded is a tubular section
    
        Dim oBoundedData As MemberConnectionData
        Dim oBoundingData As MemberConnectionData
        Dim lStatus As Long
        Dim sMsg As String
        InitMemberConnectionData oACOrEC, oBoundedData, oBoundingData, lStatus, sMsg
        Dim oSketchPlane As IJPlane
        Dim dVdirX As Double, dVdirY As Double, dVdirZ As Double
        Dim oSketchV As IJDVector
        Set oSketchV = New dVector
        'To get sketching plane
        Set oSketchPlane = GetSketchPlaneForTube(oBoundingData.AxisPort, oBoundedData.AxisPort)
        
        oSketchPlane.GetVDirection dVdirX, dVdirY, dVdirZ
        oSketchV.Set dVdirX, dVdirY, dVdirZ
                
        
        Dim oExtendedBddAxis As IJWireBody
        Dim oWireBodyUtilities As New SGOWireBodyUtilities
        Dim oSPSCrossSec As ISPSCrossSection
        Dim oBoundedPart As ISPSMemberPartCommon
        
        Set oBoundedPart = oBoundedData.MemberPart
        Set oSPSCrossSec = oBoundedPart.CrossSection
        
        Dim eSPSPortIndex As SPSMemberAxisPortIndex
        Dim oSPSSplitAxisPort As ISPSSplitAxisPort
        Set oSPSSplitAxisPort = oBoundedData.AxisPort
        eSPSPortIndex = oSPSSplitAxisPort.PortIndex
        
        Dim oPoint As IJPoint
        Dim dx As Double
        Dim dy As Double
        Dim dz As Double
        Set oPoint = oBoundedPart.PointAtEnd(eSPSPortIndex)
        oPoint.GetPoint dx, dy, dz
        Dim oPos As IJDPosition
        Set oPos = New DPosition
        oPos.Set dx, dy, dz
        
        'Translating the axis of the tube to it's centre(For Bounded) if it is not at it's center(i.e, CP 5)
        Dim oBddAxis As IJWireBody
        Dim oCmplx1 As ComplexString3d
        Set oCmplx1 = GetAxisCurveAtTubeCenter(oBoundedData.AxisPort, oPos, oBoundedData)
        Dim GeomOpr As IMSModelGeomOps.DGeomWireFrameBody
        Dim oElemCurves As IJElements
        oCmplx1.GetCurves oElemCurves
        Set GeomOpr = New DGeomWireFrameBody
        Set oBddAxis = GeomOpr.CreateSmartWireBodyFromGTypedCurves(Nothing, oElemCurves)
        oWireBodyUtilities.ExtendWire oBddAxis, 1, 1, oExtendedBddAxis
        
        Dim oExtendedBdgAxis As IJWireBody
        Dim oBoundingPart As ISPSMemberPartCommon
        Set oBoundingPart = oBoundingData.MemberPart
        Set oSPSCrossSec = oBoundingPart.CrossSection
        
        'Translating the axis of the tube to it's centre(For Bounding) if it is not at it's center(i.e, CP 5)
        Dim oBdgAxis As IJWireBody
        Set oCmplx1 = GetAxisCurveAtTubeCenter(oBoundingData.AxisPort, oPos, oBoundingData)
        oCmplx1.GetCurves oElemCurves
        Set oBdgAxis = GeomOpr.CreateSmartWireBodyFromGTypedCurves(Nothing, oElemCurves)
        oWireBodyUtilities.ExtendWire oBdgAxis, 1, 1, oExtendedBdgAxis

        Dim oBdgAxisMBdy As IJDModelBody
        Set oBdgAxisMBdy = oExtendedBdgAxis
        Dim oBddPos As IJDPosition
        Dim oBdgPos As IJDPosition
        Dim dist As Double
        ' To get distance between bounding and bounded axes at their centre
        oBdgAxisMBdy.GetMinimumDistance oExtendedBddAxis, oBdgPos, oBddPos, dist
        Dim distbtnBddtoBdg As Double
        distbtnBddtoBdg = dist
        Dim dBdgRadius As Double
        Dim dBddRadius As Double
        dBddRadius = GetMemberTubeRadius(oBoundedData.MemberPart)
        dBdgRadius = GetMemberTubeRadius(oBoundingData.MemberPart)
        Dim oBddtoBdgVec As IJDVector
        ' Position vector from point on bounded to point bounding
        Set oBddtoBdgVec = oBdgPos.Subtract(oBddPos)
        'To determine the Position vector direction w.r.to sketchplane v direction
        If Sgn(oBddtoBdgVec.Dot(oSketchV)) = -1 Then
            dist = -dist
        End If
        
        'To determine the on member cases
        If Equal(dBddRadius + dBdgRadius, distbtnBddtoBdg) Then
            If Sgn(dist) = -1 Then
            
                GetBoundingCaseForTube = OnMemberTop
            Else
            
                GetBoundingCaseForTube = OnMemberBottom
            End If
        
        
        'To determine the To Centre cases
        ElseIf GreaterThanOrEqualTo(dBdgRadius + dist, dBddRadius) And _
            LessThanOrEqualTo(-dBdgRadius + dist, -dBddRadius) Then
            
            GetBoundingCaseForTube = Center
        
        'To determine the Outside and outside cases
        ElseIf LessThan(dBdgRadius + dist, dBddRadius) And _
            GreaterThan(-dBdgRadius + dist, -dBddRadius) Then
            
            GetBoundingCaseForTube = OSBottomAndOSTop
        
        'To determine the Face and outside cases
        ElseIf Sgn(dist) = -1 Then
         
             GetBoundingCaseForTube = FCAndOSTop
        Else
         
             GetBoundingCaseForTube = FCAndOSBottom
        End If
         
        Exit Function
    
    Else
  
    Dim oBoundingOuterPort As IJPort
    Dim oBoundingInnerPort As IJPort
    
    If TypeOf oTempBounding Is ISPSDesignedMember Then
        Set oBoundingOuterPort = GetBaseOffsetOrLateralPortForObject(oBounding, BPT_Base)
    ElseIf TypeOf oTempBounding Is ISPSMemberPartCommon Then
        Set oBoundingOuterPort = GetLateralSubPortBeforeTrim(oSDO_Bounding.object, JXSEC_OUTER_TUBE)
    Else
        Exit Function
    End If
    
    'For now check and proceed if bounding is hollow tube
    If TypeOf oTempBounding Is ISPSDesignedMember Then
        Set oBoundingInnerPort = GetBaseOffsetOrLateralPortForObject(oBounding, BPT_Offset)
    Else
        Set oBoundingInnerPort = GetLateralSubPortBeforeTrim(oSDO_Bounding.object, JXSEC_INNER_TUBE)
    End If
    
        
    ' -----------------------------------------------
    ' Determine which of the Top/Bottom flanges exist
    ' -----------------------------------------------
    Dim bTopFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bBottomFlangeRight As Boolean
    
    CrossSection_Flanges oSDO_Bounded.object, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
        
    ' ----------------------------------------------------------
    ' Get the four primary ports: WebLeft, WebRight, Top, Bottom
    ' ----------------------------------------------------------
    Dim oWLport As Object
    Dim oWRPort As Object
    Dim oTopPort As Object
    Dim oBtmPort As Object
    
    Set oWLport = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_WEB_LEFT))
    Set oWRPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_WEB_RIGHT))
    Set oTopPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_TOP))
    Set oBtmPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oSDO_Bounded.object, JXSEC_BOTTOM))
    
    ' ------------------------------------------------------------------------------
    ' If there is a Top-Left flange, get intersection of Top and TopFlangeLeft ports
    ' ------------------------------------------------------------------------------
    Dim oGenericIntersector As IMSModelGeomOps.DGeomOpsIntersect
    Dim IntersectedObject As Object
    Dim oBoundedWireBdy(1 To 4) As IJWireBody
    
    Set oGenericIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    
    If bTopFlangeLeft Then
        Dim oTFLPort As Object
        Set oTFLPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oBounded, JXSEC_TOP_FLANGE_LEFT))
        oGenericIntersector.PlaceIntersectionObject Nothing, oTopPort, oTFLPort, Nothing, IntersectedObject
        Set oBoundedWireBdy(1) = IntersectedObject
    ' ----------------------------------------------------
    ' Otherwise, get intersection of Top and WebLeft ports
    ' ----------------------------------------------------
    Else
        oGenericIntersector.PlaceIntersectionObject Nothing, oTopPort, oWLport, Nothing, IntersectedObject
        Set oBoundedWireBdy(1) = IntersectedObject
    End If
                                                     
    ' --------------------------------------------------------------------------------
    ' If there is a Top-Right flange, get intersection of Top and TopFlangeRight ports
    ' --------------------------------------------------------------------------------
    Set IntersectedObject = Nothing
    
    If bTopFlangeRight Then
        Dim oTFRPort As Object
        Set oTFRPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oBounded, JXSEC_TOP_FLANGE_RIGHT))
        oGenericIntersector.PlaceIntersectionObject Nothing, oTopPort, oTFRPort, Nothing, IntersectedObject
        Set oBoundedWireBdy(2) = IntersectedObject
    ' -----------------------------------------------------
    ' Otherwise, get intersection of Top and WebRight ports
    ' -----------------------------------------------------
    Else
        oGenericIntersector.PlaceIntersectionObject Nothing, oTopPort, oWRPort, Nothing, IntersectedObject
        Set oBoundedWireBdy(2) = IntersectedObject
    End If
        
    ' ------------------------------------------------------------------------------------
    ' If there is a Bottom-Left flange, get intersection of Top and BottomFlangeLeft ports
    ' ------------------------------------------------------------------------------------
    If bBottomFlangeLeft Then
        Dim oBFLPort As Object
        Set oBFLPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oBounded, JXSEC_BOTTOM_FLANGE_LEFT))
        oGenericIntersector.PlaceIntersectionObject Nothing, oBtmPort, oBFLPort, Nothing, IntersectedObject
        Set oBoundedWireBdy(3) = IntersectedObject
    ' -------------------------------------------------------
    ' Otherwise, get intersection of Bottom and WebLeft ports
    ' -------------------------------------------------------
    Else
        oGenericIntersector.PlaceIntersectionObject Nothing, oBtmPort, oWLport, Nothing, IntersectedObject
        Set oBoundedWireBdy(3) = IntersectedObject
    End If
    
    ' --------------------------------------------------------------------------------------
    ' If there is a Bottom-Right flange, get intersection of Top and BottomFlangeRight ports
    ' --------------------------------------------------------------------------------------
    If bBottomFlangeRight Then
        Dim oBFRPort As Object
        Set oBFRPort = GetExtendedPort(GetLateralSubPortBeforeTrim(oBounded, JXSEC_BOTTOM_FLANGE_RIGHT))
        oGenericIntersector.PlaceIntersectionObject Nothing, oBtmPort, oBFRPort, Nothing, IntersectedObject
        Set oBoundedWireBdy(4) = IntersectedObject
    Else
        oGenericIntersector.PlaceIntersectionObject Nothing, oBtmPort, oWRPort, Nothing, IntersectedObject
        Set oBoundedWireBdy(4) = IntersectedObject
    End If
    
    ' --------------------------------------------------------------------------------
    ' Get minimum distance between four boundary curves and the bounding outer surface
    ' --------------------------------------------------------------------------------
    Dim oModelBodyUtils As IJSGOModelBodyUtilities
    Dim dDistance As Double
    Dim oPointOnBounding As IJDPosition
    Dim oPointOnBounded As IJDPosition
    Dim oBoundingToBoundedVec(1 To 4) As IJDVector
    Dim iCount As Integer
    Dim bAllEdgeIntersect As Boolean              ' Identification of To-Center case: all values = 0
    Dim bAllCurvesOutsideBounding As Boolean ' Identification of outside-Outside case: all values > tolerance
        
    bAllEdgeIntersect = True
    bAllCurvesOutsideBounding = True

    Set oModelBodyUtils = New SGOModelBodyUtilities

    For iCount = 1 To 4
        oModelBodyUtils.GetClosestPointsBetweenTwoBodies oBoundingOuterPort.Geometry, _
                                                         oBoundedWireBdy(iCount), _
                                                         oPointOnBounding, _
                                                         oPointOnBounded, _
                                                         dDistance
                                                         
        Set oBoundingToBoundedVec(iCount) = oPointOnBounded.Subtract(oPointOnBounding)
        
        If dDistance > EDGE_CASE_TOL Then
            bAllEdgeIntersect = False
        Else
            bAllCurvesOutsideBounding = False
        End If
        
    Next iCount
    
    ' --------------------------------------------------
    ' If all curves intersect, the case is Center (Exit)
    ' --------------------------------------------------
    
    If bAllEdgeIntersect Then
        GetBoundingCaseForTube = Center
        Exit Function
    Else
        If bPenetratesWeb Then
            iACFlag = MbrToTube_IsthisBorderACCase(oBounded, oBounding)
            Select Case iACFlag
            Case BorderAC_OSOS, BorderAC_FCOS_TOP, BorderAC_FCOS_BTM, BorderAC_ToCenter
                GetBoundingCaseForTube = Center
                Exit Function '*** Exit ***
            Case Else
                'Need to continue to check for other AC cases
            End Select
        End If
    End If
    
    
    ' ---------------------------------------------------------------
    ' If all curves are outside, the case is OutsideAndOutside (Exit)
    ' ---------------------------------------------------------------
    Dim iCount_Sec As Integer
    Dim dDotProduct As Double
    
    If bAllCurvesOutsideBounding Then
        GetBoundingCaseForTube = OSBottomAndOSTop
        Exit Function
    ' -------------------------------------------------------------------------------------------------------------------------
    ' If we find one of the four bounding to bounded vectors is directly opposite another, the case is OutsideAndOutside (Exit)
    ' -------------------------------------------------------------------------------------------------------------------------
    Else
        For iCount = 1 To 3
            For iCount_Sec = iCount + 1 To 4
                If oBoundingToBoundedVec(iCount).Length > TOLERANCE_VALUE And oBoundingToBoundedVec(iCount_Sec).Length > TOLERANCE_VALUE Then
                    dDotProduct = oBoundingToBoundedVec(iCount).Dot(oBoundingToBoundedVec(iCount_Sec))
                    If dDotProduct < 0# Then
                        GetBoundingCaseForTube = OSBottomAndOSTop
                        Exit Function
                    End If
                End If
            Next iCount_Sec
        Next iCount
    End If
    
    ' Solid round bar as bounding is not handled at this time
    If oBoundingInnerPort Is Nothing Then
        Exit Function
    End If
    ' -------------------------------------------------------------------------------------------------------------
    ' If the distance to the outer wall is less than the bounding wall thickness, the case is FaceAndOutside (Exit)
    ' -------------------------------------------------------------------------------------------------------------
    Dim dWallThickness As Double
    If bBuiltup Then
        Dim oSDO_PlatePart As StructDetailObjects.PlatePart
        Set oSDO_PlatePart = New StructDetailObjects.PlatePart
        Set oSDO_PlatePart.object = oBounding
        dWallThickness = oSDO_PlatePart.PlateThickness
        Set oSDO_PlatePart = Nothing
    Else
        dWallThickness = oSDO_Bounding.webThickness
    End If
    
    For iCount = 1 To 4
        oModelBodyUtils.GetClosestPointsBetweenTwoBodies oBoundingInnerPort.Geometry, _
                                                         oBoundedWireBdy(iCount), _
                                                         oPointOnBounding, _
                                                         oPointOnBounded, _
                                                         dDistance
        If dWallThickness - dDistance > TOLERANCE_VALUE Then
            If iCount < 3 Then
                GetBoundingCaseForTube = FCAndOSTop
            Else
                GetBoundingCaseForTube = FCAndOSBottom
            End If
            
            Exit Function
        End If
        
        Set oBoundingToBoundedVec(iCount) = oPointOnBounded.Subtract(oPointOnBounding)
    Next iCount
    
    ' ----------------------------------------------------------------------------------------------------------------------
    ' If we find one of the four bounding to bounded vectors is directly opposite another, the case is FaceAndOutside (Exit)
    ' ----------------------------------------------------------------------------------------------------------------------
    For iCount = 1 To 3
        For iCount_Sec = iCount + 1 To 4
            If oBoundingToBoundedVec(iCount).Length > TOLERANCE_VALUE And oBoundingToBoundedVec(iCount_Sec).Length > TOLERANCE_VALUE Then
                dDotProduct = oBoundingToBoundedVec(iCount).Dot(oBoundingToBoundedVec(iCount_Sec))
                If dDotProduct < 0# Then
                    If iCount < 3 Then
                        GetBoundingCaseForTube = FCAndOSTop
                    Else
                        GetBoundingCaseForTube = FCAndOSBottom
                    End If
                    
                    Exit Function
                End If
            End If
        Next iCount_Sec
    Next iCount
    
    ' ----------------------------------------------
    ' If none of the above, the case is OnTubeMember
    ' ----------------------------------------------
    GetBoundingCaseForTube = OnTubeMember
        
    Exit Function
    
    End If
ErrorHandler:
    HandleError MODULE, METHOD

End Function
' ------------------------------------------------------------------------------
' GetAxisCurveAtTubeCenter: given bounding-tube port, end cut position
' and bounding member connection data, returns axis curve transformed to
' the center of the tube.
' ------------------------------------------------------------------------------
Public Function GetAxisCurveAtTubeCenter(oPortBounding As IJPort, oEndCutPos As IJDPosition, _
        oBoundingData As MemberConnectionData) As ComplexString3d
    Const METHOD = "GetAxisCurveAtTubeCenter"
    Dim sMsg As String
    
    'Declare variables
    Dim oMatrix As DT4x4
    Dim oXSecUvec As IJDVector
    Dim oXSecVvec As IJDVector
    
    'Get cardinal point information
    Dim oBoundingPart As ISPSMemberPartCommon
    Dim oSPSCrossSec As ISPSCrossSection
    Dim lCardinalPoint As Long
    
    Dim bBuiltup As Boolean
    Dim oBUMember As ISPSDesignedMember
    
    If TypeOf oPortBounding.Connectable Is ISPSMemberPartCommon Then
        Set oBoundingPart = oPortBounding.Connectable
    Else
        IsFromBuiltUpMember oPortBounding.Connectable, bBuiltup, oBUMember
        If bBuiltup Then
            Set oBoundingPart = oBUMember
        End If
    End If
    
    Set oSPSCrossSec = oBoundingPart.CrossSection
    lCardinalPoint = oSPSCrossSec.CardinalPoint

    'Prepare complex string
    Dim oCmplx As ComplexString3d
    Dim curveElms As IJElements
    Set curveElms = New JObjectCollection
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    
    If TypeOf oBoundingData.AxisCurve Is ComplexString3d Then
        Set oCmplx = oBoundingData.AxisCurve
        oCmplx.GetCurves curveElms
    Else
        curveElms.Add oBoundingData.AxisCurve
    End If
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    Set oCmplx = oGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
    If oSPSCrossSec.CardinalPoint = 5 Or oSPSCrossSec.CardinalPoint = 10 Or oSPSCrossSec.CardinalPoint = 15 Then
        'Cardinal point of bounding is tube center, so no need to transform curve
    Else
        'Need to transform the curve to center of tube
        Dim dU As Double
        Dim dV As Double
        oSPSCrossSec.GetCardinalPointDelta Nothing, oSPSCrossSec.CardinalPoint, 5, dU, dV
        
        Dim oStructPorfile As IJStructProfilePart
        If TypeOf oBoundingPart Is IJStructProfilePart Then
            Set oStructPorfile = oBoundingPart
        Else
            sMsg = "Unknown case" 'Unknown case
            GoTo ErrorHandler
        End If

        Set oMatrix = oStructPorfile.GetCrossSectionMatrixAtPoint(oEndCutPos)

        Set oXSecUvec = New dVector
        Set oXSecVvec = New dVector

        oXSecUvec.Set oMatrix.IndexValue(0), oMatrix.IndexValue(1), oMatrix.IndexValue(2)
        oXSecVvec.Set oMatrix.IndexValue(4), oMatrix.IndexValue(5), oMatrix.IndexValue(6)
        
        'Compute bounding U-vector
        oXSecUvec.Length = dU
    
        'Compute bounding V-vector
        oXSecVvec.Length = dV
                
        'Compute resultant vector
        Dim oResVec As IJDVector
        Set oResVec = oXSecUvec.Add(oXSecVvec)
        
        'Prepare transformation matrix
        oMatrix.LoadIdentity
        oMatrix.Translate oResVec
                
        'Tranform complex string to tube-center
        oCmplx.Transform oMatrix
    End If

    Set GetAxisCurveAtTubeCenter = oCmplx
    
    'Cleanup
    Set oGeometryFactory = Nothing
    Set oBoundingPart = Nothing
    Set oSPSCrossSec = Nothing
    Set curveElms = Nothing
    Set oMatrix = Nothing
    Set oXSecUvec = Nothing
    Set oXSecVvec = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

' ------------------------------------------------------------------------------
' IsBoundingOrthogonalToBoundedFlg: given bounding-data and bounded-data
' checks if bounded flange normal and bounding axis are collinear , end cut position
' and bounding member connection data, returns axis curve transformed to
' the center of the tube.
' ------------------------------------------------------------------------------

Public Function IsBoundingOrthogonalToBoundedFlg(oBoundingData As MemberConnectionData, _
    oBoundedData As MemberConnectionData, ePortId As JXSEC_CODE) As Boolean
    Const METHOD = "::IsBoundingOrthogonalToBoundedFlg"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "IsBoundingOrthogonalToBoundedFlg"

    'Get bounding axis curve to center
    Dim oCrv As IJCurve
    Set oCrv = oBoundingData.AxisCurve
    
    'Determine closest point on top/bottom port
    Dim oBoundedFlgPos As IJDPosition
    Dim oBoundingPos As IJDPosition
    Set oBoundedFlgPos = New DPosition
    Set oBoundingPos = New DPosition
    
    Dim oSDO_BoundedMbr As StructDetailObjects.MemberPart
    Set oSDO_BoundedMbr = New StructDetailObjects.MemberPart
    Set oSDO_BoundedMbr.object = oBoundedData.MemberPart
    Dim oBoundedTopPort As IJPort
    Set oBoundedTopPort = GetLateralSubPortBeforeTrim(oSDO_BoundedMbr.object, ePortId)

    Dim oSurface As IJSurface
    Set oSurface = oBoundedTopPort
    Dim dSrcX As Double, dSrcY As Double, dSrcZ As Double
    Dim dInX As Double, dInY As Double, dInZ As Double
    Dim dDistance As Double
    oCrv.DistanceBetween oSurface, dDistance, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
    Set oBoundedFlgPos = New DPosition
    oBoundedFlgPos.Set dInX, dInY, dInZ
    oBoundingPos.Set dSrcX, dSrcY, dSrcZ
    
    'Notes:-
    ' Matrix.IndexValue(0,1,2)   : U is direction Along Axis
    ' Matrix.IndexValue(4,5,6)   : V is direction normal to Web (from Web Right to Web Left)
    ' Matrix.IndexValue(8,9,10)  : W is direction normal to Flange (from Flange Bottom to Flange Top)
    ' Matrix.IndexValue(12,13,14): Root/Origin Point

    'Get bounding matrix
    oBoundingData.MemberPart.Rotation.GetTransformAtPosition oBoundingPos.x, oBoundingPos.y, oBoundingPos.z, oBoundingData.Matrix, Nothing

    'Get bounding U-direction vector
    Dim oBounding_UDir As IJDVector
    Set oBounding_UDir = New dVector
    oBounding_UDir.Set oBoundingData.Matrix.IndexValue(0), oBoundingData.Matrix.IndexValue(1), oBoundingData.Matrix.IndexValue(2)

    'Get bounded matrix
    oBoundedData.MemberPart.Rotation.GetTransformAtPosition oBoundedFlgPos.x, oBoundedFlgPos.y, oBoundedFlgPos.z, oBoundedData.Matrix, Nothing

    Dim oBounded_WDir As IJDVector
    Set oBounded_WDir = New dVector
    oBounded_WDir.Set oBoundedData.Matrix.IndexValue(8), oBoundedData.Matrix.IndexValue(9), oBoundedData.Matrix.IndexValue(10)
    
    'Compute dot product and check for parallel condition
    If Abs(1# - Abs(oBounding_UDir.Dot(oBounded_WDir))) < 0.001 Then
        IsBoundingOrthogonalToBoundedFlg = True
    Else
        IsBoundingOrthogonalToBoundedFlg = False
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

'*************************************************************************
'Sub
'   GetSketchPlaneForMSSymbol
'
'Abstract
'   Method to compute the web sketching plane
'
'Inputs
'   oBoundedPort As IJPort
'   oBoundingPort As IJPort
'   bPenetratesWeb As Boolean
'   bUseTopFlange As Boolean
'       - The Bounding Member
'   Optional bPersitedPlane As Boolean
'       - Default value is True which persists the sketching plane, user would need to remove it later
'        (kept this as optional and default so that code prior to TR223552 fix is not impacted)
'         User to specify 'False' to create a transient (non-persisted) plane.
'
'Return
'   oSketchPlane As IJPlane
'
'Exceptions
'
'***************************************************************************
Public Sub GetSketchPlaneForMSSymbol(oBoundedPort As IJPort, _
                                     oBoundingPort As IJPort, _
                                     bPenetratesWeb As Boolean, _
                                     bUseTopFlange As Boolean, _
                                     ByRef oSketchPlane As IJPlane, _
                                     Optional bPersitedPlane As Boolean = True)
Const METHOD = "GetSketchPlaneForMSSymbol"
    On Error GoTo ErrorHandler
    
        Dim oBounded As Object
        Set oBounded = oBoundedPort.Connectable
    'Calculate the Sketch Plane for the Measurement Symbol
    Dim oSketchNormal As New dVector
    Dim oSketchPosition As New DPosition
    Dim oSketchU As New dVector
    Dim oTop As IJPort
    Dim oWL As IJPort
    Dim oBottom As IJPort
    Dim oEnd As IJPort
    Dim oTopGeometry As IJSurfaceBody
    Dim oWLGeometry As IJSurfaceBody
    Dim oBottomGeometry As IJSurfaceBody
    Dim oTopNormal As IJDVector
    Dim oBottomNormal As IJDVector
    Dim oWLNormal As IJDVector
    Dim oEndNormal As IJDVector
    Dim oGeometryFactory As IngrGeom3D.IPlanes3d
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    
    If IsTubularMember(oBoundedPort.Connectable) And _
        Get_PortFaceType(oBoundingPort) <> C_Port_Base And _
        Get_PortFaceType(oBoundingPort) <> C_Port_Offset Then
       
        Dim oTubeEndCutSketchPlane As IJPlane
        Dim dRootX As Double, dRootY As Double, dRootZ As Double
        Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
        Dim dUdirX As Double, dUdirY As Double, dUdirZ As Double
        
        Set oTubeEndCutSketchPlane = GetSketchPlaneForTube(oBoundingPort, oBoundedPort)
       
        oTubeEndCutSketchPlane.GetRootPoint dRootX, dRootY, dRootZ
        oTubeEndCutSketchPlane.GetNormal dNormalX, dNormalY, dNormalZ
        oTubeEndCutSketchPlane.GetUDirection dUdirX, dUdirY, dUdirZ
       
        oSketchPosition.Set dRootX, dRootY, dRootZ
        oSketchNormal.Set dNormalX, dNormalY, dNormalZ
        oSketchU.Set dUdirX, dUdirY, dUdirZ
       
    Else
    
        Set oTop = GetLateralSubPortBeforeTrim(oBounded, JXSEC_TOP)
        Set oWL = GetLateralSubPortBeforeTrim(oBounded, JXSEC_WEB_LEFT)
        Set oBottom = GetLateralSubPortBeforeTrim(oBounded, JXSEC_BOTTOM)
        
        Dim SplitAxisPort As ISPSSplitAxisPort
        Dim oBoundedPart As ISPSMemberPartPrismatic
        Dim oBoundedMemberPart As ISPSMemberPartCommon
        Dim MemberAxisPortIndex As SPSMemberAxisPortIndex
        
        If TypeOf oBoundedPort.Connectable Is ISPSMemberPartCommon Then
            Set SplitAxisPort = oBoundedPort
            MemberAxisPortIndex = SplitAxisPort.PortIndex
                        
        Set oBoundedPart = oBoundedPort.Connectable
                Dim oBoundedCurve As IJCurve
        Set oBoundedCurve = oBoundedPart.Axis
        ElseIf TypeOf oBoundedPort.Connectable Is IJProfile Then
            Dim oProfilePort As IJStructPort
            Set oProfilePort = oBoundedPort
            
            'find equivalent Member Axis port, if the bounded Obj is other than Member
            If oProfilePort.ContextID = CTX_BASE Then
                MemberAxisPortIndex = SPSMemberAxisStart
            ElseIf oProfilePort.ContextID = CTX_OFFSET Then
                MemberAxisPortIndex = SPSMemberAxisEnd
            Else
                MemberAxisPortIndex = SPSMemberAxisAlong
            End If
            
            If TypeOf GetProfilePartLandingCurve(oBoundedPort.Connectable) Is IJCurve Then
                Set oBoundedCurve = GetProfilePartLandingCurve(oBoundedPort.Connectable)
            Else
                'Need to get the Axis of Bounded Obj of Type IJCurve!!!
                'Need to handle this case, but currently no such
                'case found
            End If
        End If
        
        Dim startParam As Double
        Dim endParam As Double
        Dim posX As Double
        Dim posY As Double
        Dim posZ As Double
        Dim tanX As Double
        Dim tanY As Double
        Dim tanZ As Double
        Dim curX As Double
        Dim curY As Double
        Dim curZ As Double
        
        oBoundedCurve.ParamRange startParam, endParam
        
        ' ---------------------
        ' Get the end direction
        ' ---------------------
        If MemberAxisPortIndex = SPSMemberAxisStart Then
            ' Negate the u-vector if the port is at the start, so that it points away from the member
            oBoundedCurve.Evaluate startParam, posX, posY, posZ, tanX, tanY, tanZ, curX, curY, curZ
            tanX = -tanX
            tanY = -tanY
            tanZ = -tanZ
        ElseIf MemberAxisPortIndex = SPSMemberAxisEnd Then
            oBoundedCurve.Evaluate endParam, posX, posY, posZ, tanX, tanY, tanZ, curX, curY, curZ
        ElseIf MemberAxisPortIndex = SPSMemberAxisAlong Then
                'For Both Along Case -Split None case
                'Get Bounded Location
            Dim oBoundingCurve As IJCurve
            Dim oBoundingPart As ISPSMemberPartCommon
            Dim dx As Double, dy As Double, dz As Double
            Dim dX2 As Double, dY2 As Double, dZ2 As Double
            Dim dMinDist As Double

            If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then

                Set oBoundingPart = oBoundingPort.Connectable
                Set oBoundingCurve = oBoundingPart.Axis
                            
                oBoundedCurve.DistanceBetween oBoundingCurve, dMinDist, dx, dy, dz, dX2, dY2, dZ2
                
                oBoundedCurve.Parameter dx, dy, dz, endParam
                oBoundedCurve.Evaluate endParam, posX, posY, posZ, tanX, tanY, tanZ, curX, curY, curZ
            End If
        End If
        
        Set oEndNormal = New dVector
        oEndNormal.Set tanX, tanY, tanZ
        
        Set oTopGeometry = oTop.Geometry
        Set oWLGeometry = oWL.Geometry
        Set oBottomGeometry = oBottom.Geometry
        
        Dim oTopPos As IJDPosition
        Dim oWLPos As IJDPosition
        Dim oBottomPos As IJDPosition
        Dim oEndPos As IJDPosition
        Dim oPosition As IJDPosition
        Dim dDistance As Double
        Dim px As Double, py As Double, pz As Double
    
        If TypeOf oBoundedPort.Connectable Is ISPSMemberPartCommon Then
            Set oBoundedMemberPart = oBoundedPort.Connectable
            If MemberAxisPortIndex = SPSMemberAxisStart Or MemberAxisPortIndex = SPSMemberAxisEnd Then
                oBoundedMemberPart.PointAtEnd(MemberAxisPortIndex).GetPoint px, py, pz
            ElseIf MemberAxisPortIndex = SPSMemberAxisAlong Then 'For Both Along case
                px = posX
                py = posY
                pz = posZ
            End If
        Else
            'For Profiles
            px = posX
            py = posY
            pz = posZ
        End If
        
        Set oPosition = New DPosition
        oPosition.Set px, py, pz
    
        Dim oModelBodyUtil As IJSGOModelBodyUtilities
        Set oModelBodyUtil = New SGOModelBodyUtilities
        
        oModelBodyUtil.GetClosestPointOnBody oTopGeometry, oPosition, oTopPos, dDistance
        oTopGeometry.GetNormalFromPosition oTopPos, oTopNormal
        
        oModelBodyUtil.GetClosestPointOnBody oBottomGeometry, oPosition, oBottomPos, dDistance
        oBottomGeometry.GetNormalFromPosition oBottomPos, oBottomNormal
        
        oModelBodyUtil.GetClosestPointOnBody oWLGeometry, oPosition, oWLPos, dDistance
        oWLGeometry.GetNormalFromPosition oWLPos, oWLNormal
            
        ' -------------------------------------------------------------------------------------------------------------------------
        ' The end direction should be to the left looking at the sketch window, which means the sketch U is to the right (opposite)
        ' -------------------------------------------------------------------------------------------------------------------------
        oSketchU.Set -oEndNormal.x, -oEndNormal.y, -oEndNormal.z
        
        Dim SketchPosOffset As New dVector
        
        Dim oSDO_BoundedMember As New StructDetailObjects.MemberPart
        Dim oSDO_BoundedStiffener As New StructDetailObjects.ProfilePart
        Dim dWebThickness As Double
        Dim dFlangeThickness As Double
        
        If TypeOf oBoundedPort.Connectable Is ISPSMemberPartCommon Then
            Set oSDO_BoundedMember.object = oBoundedPort.Connectable
            dWebThickness = oSDO_BoundedMember.webThickness
            dFlangeThickness = oSDO_BoundedMember.flangeThickness
        ElseIf TypeOf oBoundedPort.Connectable Is IJProfile Then
            Set oSDO_BoundedStiffener.object = oBoundedPort.Connectable
            dWebThickness = oSDO_BoundedStiffener.webThickness
            dFlangeThickness = oSDO_BoundedStiffener.flangeThickness
        End If
    
        If bPenetratesWeb Then
            Set oSketchNormal = oSketchU.Cross(oTopNormal)
            SketchPosOffset.Set -oWLNormal.x, -oWLNormal.y, -oWLNormal.z
            SketchPosOffset.Length = dWebThickness / 2
            Set oSketchPosition = oWLPos.Offset(SketchPosOffset)
        Else
            Set oSketchNormal = oSketchU.Cross(oWLNormal)
            If bUseTopFlange Then
                SketchPosOffset.Set -oTopNormal.x, -oTopNormal.y, -oTopNormal.z
                SketchPosOffset.Length = dFlangeThickness / 2
                Set oSketchPosition = oTopPos.Offset(SketchPosOffset)
            Else
                SketchPosOffset.Set -oBottomNormal.x, -oBottomNormal.y, -oBottomNormal.z
                SketchPosOffset.Length = dFlangeThickness / 2
                Set oSketchPosition = oBottomPos.Offset(SketchPosOffset)
            End If
        End If
    End If

    If bPersitedPlane Then
        'Persist the plane by passing resource manager reference: cases involving 'S' members in
        'GetConnectedEdgeInfo' method need persisted plane to resymbolize measurement symbols
        Set oSketchPlane = oGeometryFactory.CreateByPointNormal(GetResourceMgr, _
           oSketchPosition.x, oSketchPosition.y, oSketchPosition.z, _
           oSketchNormal.x, oSketchNormal.y, oSketchNormal.z)
    Else
        'Do not persist the plane - this is preferred approach rather than creating a persisted plane
        'and deleting it later(which will have negative impact on performance).
        Set oSketchPlane = oGeometryFactory.CreateByPointNormal(Nothing, _
           oSketchPosition.x, oSketchPosition.y, oSketchPosition.z, _
           oSketchNormal.x, oSketchNormal.y, oSketchNormal.z)
    End If

    oSketchPlane.SetUDirection oSketchU.x, oSketchU.y, oSketchU.z
    oSketchPlane.SetNormal oSketchNormal.x, oSketchNormal.y, oSketchNormal.z
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub
' IsTopOrBottomInsetBrace()
' B-52232: Method to determine if a Top/Bottom brace is needed for a given AC selection
' For a given relative position of Bounded w.r.t Bounding, using the ConnectedEdge info, we determine if a Top/Bottom brace is needed.
' Returns True if a Top brace is valid.
Public Function IsTopOrBottomInsetBrace(oACRootSelection As String, oEdge1 As ConnectedEdgeInfo, oEdge2 As ConnectedEdgeInfo) As Boolean
Const METHOD = "IsTopOrBottomInsetBrace"
    On Error GoTo ErrorHandler
    'logic to find out if a top/bottom brace is needed is different for different selectors
    
    If oACRootSelection = "" Then
        GoTo ErrorHandler
    End If
    '************************************************************************************************
    '************************************************************************************************
    'Find out if a Top or a Bottom brace is needed
    
    'Get the intersecting edge ports of the Bounding with Top/Bottom of the Bounded. For a Top brace, we are interested in finding out
    'the intersecting edge of the bounding with the Top of the bounded.Same for Bottom brace
    '************************************************************************************************
    Dim oTopIntersectingEdge As JXSEC_CODE, oBtmIntersectingEdge As JXSEC_CODE 'these are the alias edges, remember to be reversemap when using'em.
    
    If InStr(1, oACRootSelection, "MbrAxisByOnMember", vbTextCompare) > 0 Then
        If oEdge1.IntersectingEdge = eBounding_Edge.Above Or oEdge1.IntersectingEdge = eBounding_Edge.Top Then
            IsTopOrBottomInsetBrace = False     'bottom brace
        Else
            IsTopOrBottomInsetBrace = True     'top brace
        End If
    ElseIf InStr(1, oACRootSelection, "MbrAxisByEdge", vbTextCompare) > 0 Then
         If oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or oEdge1.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Bottom Or _
         oEdge2.IntersectingEdge = eBounding_Edge.Below Or oEdge2.IntersectingEdge = eBounding_Edge.Bottom Then
            IsTopOrBottomInsetBrace = True     'top brace
        ElseIf oEdge1.IntersectingEdge = eBounding_Edge.Above Or oEdge1.IntersectingEdge = eBounding_Edge.Top Or _
            oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or oEdge2.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom Then
            IsTopOrBottomInsetBrace = False      'bottom brace
        End If
    ElseIf InStr(1, oACRootSelection, "MbrAxisByFCAndOS1Edge", vbTextCompare) > 0 Or InStr(1, oACRootSelection, "MbrAxisByFCAndOSNoEdge", vbTextCompare) > 0 Or _
            InStr(1, oACRootSelection, "MbrAxisByFCAndEdge", vbTextCompare) > 0 Then
        If oEdge1.IntersectingEdge = eBounding_Edge.Web_Right Then
            IsTopOrBottomInsetBrace = True   'top brace
        ElseIf oEdge2.IntersectingEdge = eBounding_Edge.Web_Right Then
            IsTopOrBottomInsetBrace = False       'bottom brace
        End If
    ElseIf InStr(1, oACRootSelection, "MbrAxisByCenter", vbTextCompare) > 0 Then
        'handled in center def file
    End If
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' GetTopAndBottomIntersectingEdgeIDs()

' B-52232: Method to determine the EdgeIDs of the Intersecting edges of Bounding with Top/Bottom of Bounded
' For a given relative position of Bounded w.r.t Bounding, using the ConnectedEdge info, we determine the Top & Bottom Intersecting edges and determine their IDs.
Public Sub GetIntersectingEdgeID(oEdge As ConnectedEdgeInfo, oIntersectingEdge As JXSEC_CODE)
Const METHOD = "GetIntersectingEdgeID"
    On Error GoTo ErrorHandler
    
    If oEdge.IntersectingEdge = eBounding_Edge.Bottom Then
        oIntersectingEdge = JXSEC_BOTTOM
    ElseIf oEdge.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Then
        oIntersectingEdge = JXSEC_BOTTOM_FLANGE_RIGHT
    ElseIf oEdge.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top Then
        oIntersectingEdge = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
    ElseIf oEdge.IntersectingEdge = eBounding_Edge.Web_Right Then
        oIntersectingEdge = JXSEC_WEB_RIGHT
    ElseIf oEdge.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom Then
        oIntersectingEdge = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
    ElseIf oEdge.IntersectingEdge = eBounding_Edge.Top_Flange_Right Then
        oIntersectingEdge = JXSEC_TOP_FLANGE_RIGHT
    ElseIf oEdge.IntersectingEdge = eBounding_Edge.Top Then
        oIntersectingEdge = JXSEC_TOP
    End If
    
    
Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub
'*****************************************************************************************


'*************************************************************************
' Function:
' AssemblyConnection_SmartItemName
'
' Abstract
' Given a smart occurence object, return the name of the assembly connection item
' (parent of the entire custom assembly)
'
'***************************************************************************
Public Sub AssemblyConnection_SmartItemName(oOccurrenceObject As Object, _
                                            Optional sACItemName As String, _
                                            Optional oACObject As Object = Nothing, _
                                            Optional iHierarchy As Integer = 0)
    
    Const METHOD = "MbrMeasurementUtilities::AssemblyConnection_SmartItemName"

    Dim oChild As Object
    Dim oParent As Object
    Dim sParentName As String
    Dim iLoopCount As Integer
    iLoopCount = 0
    
    Set oChild = oOccurrenceObject
    sACItemName = vbNullString
    
    Do
        Parent_SmartItemName oChild, sParentName, oParent
        iLoopCount = iLoopCount + 1
        
        If oParent Is Nothing Then
            'If for the first iteration itself we dont have any parent AC
            'check if the paased object itself might be AC itself.
            If iLoopCount = 1 And Not (oOccurrenceObject Is Nothing) Then
                If TypeOf oOccurrenceObject Is IJAppConnection Then
                    If TypeOf oOccurrenceObject Is IJSmartOccurrence Then
                        Dim oSO As IJSmartOccurrence
                        Set oSO = oOccurrenceObject
                        sACItemName = oSO.Item
                    End If
                    Set oACObject = oOccurrenceObject
                 End If
            End If
            Exit Do
        ElseIf TypeOf oParent Is IJAppConnection Then
            sACItemName = sParentName
            Set oACObject = oParent
            Exit Do
        ElseIf TypeOf oParent Is IJFreeEndCut Then
            Exit Do
        End If
        
        If iLoopCount = iHierarchy Then
            Exit Do
        End If
        
        Set oChild = oParent
    Loop
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'**********************************************************************************************
' Method      : IsWebPenetrated
' Description : Checks whether the Member is web penerated/ flange penetrated
'               This method Checks whether the test case is web penerated or flange penetrated
'               by returnig a boolean value. If the boolean value is true it is web penetrated case
'               or else it is flange penetrated case
'
'**********************************************************************************************
Public Function IsWebPenetrated(oBoundingPort As IJPort, oBoundedPort As IJPort) As Boolean
    Const MT = "IsWebPenetrated"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim oBoundedPart As Object
    Dim oBoundedMemberPart As ISPSMemberPartCommon
    Dim oBoundingPart As IJStructProfilePart
    Dim oBoundedStart As IJPoint
    Dim oBoundedEnd As IJPoint
    Dim oBoundedPos As IJDPosition
    Dim dStartX As Double
    Dim dStartY As Double
    Dim dStartZ As Double
    Dim dEndX As Double
    Dim dEndY As Double
    Dim dEndZ As Double
    Dim oBoundedMemberPort As ISPSSplitAxisPort
    Dim oBoundingXSectionMatrix As IJDT4x4
    Dim oBoundedXSectionMatrix As IJDT4x4
    Dim oDummyMatrix As IJDT4x4
    Dim oBoundedStructPort As IJStructPort
    Dim bIsBuiltup As Boolean, oBUMember As ISPSDesignedMember
    
    ' by default web is penetrating
    IsWebPenetrated = True
    
    Set oBoundedPart = oBoundedPort.Connectable
    IsFromBuiltUpMember oBoundedPart, bIsBuiltup, oBUMember
    
    If bIsBuiltup Then
        Set oBoundedPart = oBUMember
    End If
    
    If Not oBoundedPart Is Nothing Then
        If TypeOf oBoundedPart Is IJStructProfilePart Then
            ' Valid case
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    If Not oBoundingPort Is Nothing Then
        If Not TypeOf oBoundingPort Is IJPort Then
            Exit Function
        ElseIf oBoundingPort.Connectable Is Nothing Then
            Exit Function
        ElseIf TypeOf oBoundingPort.Connectable Is IJPlate Then
            Dim bIsPlatePartOfBuiltup As Boolean
            IsFromBuiltUpMember oBoundingPort.Connectable, bIsPlatePartOfBuiltup
            If Not bIsPlatePartOfBuiltup Then
                Exit Function
            End If
        ElseIf TypeOf oBoundingPort.Connectable Is SPSSlabEntity Then
            Exit Function
        ElseIf TypeOf oBoundingPort.Connectable Is SPSWallPart Then
            Exit Function
        Else
            ' Valid case
        End If
    Else
        Exit Function
    End If
    
    ' if Bounded is Tube and Bounding is Not Tube, consider as WebPenetrating
    If IsTubularMember(oBoundedPort) And Not IsTubularMember(oBoundingPort) Then
        IsWebPenetrated = True
        Exit Function
    End If
    
    ' if v-vector is more alligned with the bounding x-axis, must be penetrating flange
    Dim oBoundedU As IJDVector
    Dim oBoundedV As IJDVector
    Dim oBoundingW As IJDVector
    Dim oBdedLandingCurve As IJWireBody
    Dim posStart As IJDPosition
    Dim posEnd As IJDPosition
    
    sMsg = "Getting the bounded Position"
    
    Set oBoundedU = New dVector
    Set oBoundedV = New dVector
    Set oBoundingW = New dVector
    Set oBoundedPos = New DPosition
    
    ' Step 1: compute vectors for bounded
    If TypeOf oBoundedPart Is ISPSMemberPartCommon Then
        Set oBoundedMemberPart = oBoundedPart
        Set oBoundedStart = oBoundedMemberPart.PointAtEnd(SPSMemberAxisStart)
        Set oBoundedEnd = oBoundedMemberPart.PointAtEnd(SPSMemberAxisEnd)
        
        If TypeOf oBoundedPort Is ISPSSplitAxisPort Then
            Set oBoundedMemberPort = oBoundedPort
            If oBoundedMemberPort.PortIndex = SPSMemberAxisStart Then
                oBoundedStart.GetPoint dStartX, dStartY, dStartZ
                oBoundedPos.Set dStartX, dStartY, dStartZ
            ElseIf oBoundedMemberPort.PortIndex = SPSMemberAxisEnd Then
                oBoundedEnd.GetPoint dEndX, dEndY, dEndZ
                oBoundedPos.Set dEndX, dEndY, dEndZ
            ElseIf oBoundedMemberPort.PortIndex = SPSMemberAxisAlong Then
            'Both Along ports will be Axis along for Split None case
                Dim oCurve1 As IJCurve
                Dim oCurve2 As IJCurve
                
                Dim oBoundingMemberPart As ISPSMemberPartCommon
                Dim dMinDist As Double
                
                'Get Bounded Location
                If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then
                    Set oBoundingMemberPart = oBoundingPort.Connectable
                    Set oCurve1 = oBoundedMemberPart.Axis
                    Set oCurve2 = oBoundingMemberPart.Axis
                    oCurve1.DistanceBetween oCurve2, dMinDist, dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
                    oBoundedPos.Set dStartX, dStartY, dStartZ
                End If
            End If
        End If

        sMsg = "computing the  vectors for bounded"
    
        oBoundedMemberPart.Rotation.GetTransformAtPosition oBoundedPos.x, oBoundedPos.y, oBoundedPos.z, oBoundedXSectionMatrix, oDummyMatrix
        oBoundedU.Set oBoundedXSectionMatrix.IndexValue(4), oBoundedXSectionMatrix.IndexValue(5), oBoundedXSectionMatrix.IndexValue(6)
        oBoundedV.Set oBoundedXSectionMatrix.IndexValue(8), oBoundedXSectionMatrix.IndexValue(9), oBoundedXSectionMatrix.IndexValue(10)
                  
    ElseIf TypeOf oBoundedPart Is IJProfile Then
        
        Set oBdedLandingCurve = GetProfilePartLandingCurve(oBoundedPart)
        
        oBdedLandingCurve.GetEndPoints posStart, posEnd
        
        If TypeOf oBoundedPort Is IJStructPort Then
            
            Set oBoundedStructPort = oBoundedPort
            
            If oBoundedStructPort.ContextID = CTX_BASE Then
                oBoundedPos.Set posStart.x, posStart.y, posStart.z
            ElseIf oBoundedStructPort.ContextID = CTX_OFFSET Then
                oBoundedPos.Set posEnd.x, posEnd.y, posEnd.z
            Else
                ' We dont expect the code to come here
                'Error Out ot Need to Modify the Logic
            End If
        
        End If
        
        Dim oProfile As IJProfileAttributes
        Dim oProfUVec As IJDVector
        Dim oProfVVec As IJDVector
        Dim oOriginPos As IJDPosition
        Set oProfile = New ProfileUtils
        oProfile.GetProfileOrientationAndLocation oBoundedPart, oBoundedPos, oProfUVec, oProfVVec, oOriginPos
        oBoundedU.Set oProfUVec.x, oProfUVec.y, oProfUVec.z
        oBoundedV.Set oProfVVec.x, oProfVVec.y, oProfVVec.z
    End If
    
    Set oBUMember = Nothing
    bIsBuiltup = False
    
'    Set oBoundingPart = oBoundingPort.Connectable
    IsFromBuiltUpMember oBoundingPort.Connectable, bIsBuiltup, oBUMember
    
    If bIsBuiltup Then
        Set oBoundingPart = oBUMember
    Else
        Set oBoundingPart = oBoundingPort.Connectable
    End If
    
    ' Step 2: compute vectors for bounding
    If TypeOf oBoundingPart Is IJStructProfilePart Then
        'Case for Members and Profiles
'        Set oBoundingPart = oBoundingPort.Connectable
        Set oBoundingXSectionMatrix = New DT4x4
        Set oBoundingXSectionMatrix = oBoundingPart.GetCrossSectionMatrixAtPoint(oBoundedPos)
        'Set BoundingW to the Z Axis of the Bounding cross section
        oBoundingW.Set oBoundingXSectionMatrix.IndexValue(8), oBoundingXSectionMatrix.IndexValue(9), oBoundingXSectionMatrix.IndexValue(10)
    End If
                    
    oBoundedU.Length = 1#
    oBoundedV.Length = 1#
    oBoundingW.Length = 1#

    Dim dot_wU As Double
    Dim dot_wV As Double
    
    sMsg = "Getting the dot product"
    
    dot_wU = Round(Abs(oBoundingW.Dot(oBoundedU)), 4)
    dot_wV = Round(Abs(oBoundingW.Dot(oBoundedV)), 4)

    If GreaterThan(dot_wV, dot_wU) Then
        IsWebPenetrated = False
    Else
        IsWebPenetrated = True
    End If
    
    Set oBoundingXSectionMatrix = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT, sMsg).Number
    
End Function

' |                                                       ' |
' |----------------------------                           ' |
' |----------------------------                           ' |
' |              ^                                        ' +------------------+
' |              |                                        '                    |
' |        Inside clearance                               '            Outside |
' |              |                                        '            Overlap |
' |              v                                        '               |    |
' +------------------+ -----                              '               v    |-------------------------------
'                    |    ^                               ' -------------------+
'                    |    |                               '          |         |
'                    |Inside Overlap                      '          |  ------ +-------------------------------
'                    |    |                               '          |    ^        ^
'                    |    v                               '          |    |        |
' -------------------+ -----                              '          |             |
'          |        ^                                     '          |      Outside clearance
'          |        |                                     '          |             |
'          | Outside clearance                            '          |             |
'          |        |                                     '          |             v
'          |        v                                     '          |---------------------------------------
'          |--------------------------------------        '          |
'          +---------------------------------------       '          |
'                                                         '          +---------------------------------------
Private Sub GetWebEdgeOverlapAndClearance(oACOrEC As Object, _
                                          bBottomEdge As Boolean, _
                                          Optional dInsideOverlap As Double, _
                                          Optional dOutsideOverlap As Double, _
                                          Optional dInsideClearance As Double, _
                                          Optional dOutsideClearance As Double, _
                                          Optional dEdgeLength As Double, _
                                          Optional bIsEdgeToEdge As Boolean)

    Const METHOD = "MbrMeasurementUtilities::GetWebEdgeOverlapAndClearance"

    On Error GoTo ErrorHandler
    Dim sError As String
        
    dInsideOverlap = 0#
    dOutsideOverlap = 0#
    dInsideClearance = 0#
    dOutsideClearance = 0#
    bIsEdgeToEdge = False
    
    ' -----------------------------
    'Get bounding and bounded ports
    ' -----------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    If TypeOf oACOrEC Is IJAppConnection Then
        Dim lStatus As Long
        Dim sMsg As String
        Dim oAppConn As IJAppConnection
        Set oAppConn = oACOrEC
        GetAssemblyConnectionInputs oAppConn, oBoundedPort, oBoundingPort
        
    ElseIf TypeOf oACOrEC Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Set oFeature = oACOrEC
        
        If oFeature.get_StructFeatureType = SF_FlangeCut Then
            Dim oSDOFlange As New StructDetailObjects.FlangeCut
            Set oSDOFlange.object = oACOrEC
            Set oBoundedPort = oSDOFlange.BoundedPort
            Set oBoundingPort = oSDOFlange.BoundingPort
        ElseIf oFeature.get_StructFeatureType = SF_WebCut Then
            Dim oSDOWeb As New StructDetailObjects.WebCut
            Set oSDOWeb.object = oACOrEC
            Set oBoundedPort = oSDOWeb.BoundedPort
            Set oBoundingPort = oSDOWeb.BoundingPort
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    ' ------------------------
    ' Not available for a tube
    ' ------------------------
    If IsTubularMember(oBoundingPort.Connectable) Then
        Exit Sub
    End If
    
    ' ----------------------------------------------------------------------
    ' Determine the relative orientation of the bounded and bounding objects
    ' ----------------------------------------------------------------------
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As New Collection

    Set oEdgeMapColl = GetEdgeMap(oACOrEC, oBoundingPort, oBoundedPort, , bPenetratesWeb)

    ' --------------------------
    ' Exit if not web-penetrated
    ' --------------------------
    If Not bPenetratesWeb Then
        Exit Sub
    End If
    
' This doesn't work because it is not using the simplfied alias.  For instance, if only reflected, it returns TopFlangeRight maps to TopFlangeLeft
'    ' -------------------------------------------------------------
'    ' Exit if the bounding alias does not have the specified edge
'    ' -------------------------------------------------------------
'    Dim realEdgeID As JXSEC_CODE
'    If bBottomEdge Then
'        realEdgeID = ReverseMap(JXSEC_BOTTOM_FLANGE_RIGHT, oEdgeMapColl)
'    Else
'        realEdgeID = ReverseMap(JXSEC_TOP_FLANGE_RIGHT, oEdgeMapColl)
'    End If
'
'    If realEdgeID = JXSEC_UNKNOWN Then
''        sError = "The specified edge does not exist"
''        GoTo ErrorHandler
'        Exit Sub
'    End If
    
    ' -------------------------------------------------------------------
    ' Determine the relative position of the bounded and bounding objects
    ' -------------------------------------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection
    Set oMeasurements = New Collection
    
    GetConnectedEdgeInfo oACOrEC, _
                         oBoundedPort, _
                         oBoundingPort, _
                         eTopOrWL, _
                         eBottomOrWR, _
                         eInsideTFOrTFL, _
                         eInsideBFOrTFR, _
                         oMeasurements
    
    ' -----------------------------------------
    ' Exit if the requested edge does not exist
    ' Calculate the edge length
    ' -----------------------------------------
    If bBottomEdge Then
        If Not KeyExists("DimPt21ToPt23", oMeasurements) Then
            Exit Sub
        Else
            dEdgeLength = oMeasurements.Item("DimPt21ToPt23")
        End If
    Else
        If Not KeyExists("DimPt15ToPt17", oMeasurements) Then
            Exit Sub
        Else
            dEdgeLength = oMeasurements.Item("DimPt15ToPt17")
        End If
    End If
    
    ' ----------------------------------------------
    ' If this is overlap at the bounding bottom edge
    ' ----------------------------------------------
    If bBottomEdge Then
        
        ' -------------------------------------------------------
        ' If the inside of the top flange is above the top corner
        ' -------------------------------------------------------
        Dim eTopInsidePos21 As eRelativePointPosition
        Dim eTopInsidePos23 As eRelativePointPosition
        Dim eBtmInsidePos21 As eRelativePointPosition
        Dim eBtmInsidePos23 As eRelativePointPosition
        Dim eTopOutsidePos21 As eRelativePointPosition
        Dim eTopOutsidePos23 As eRelativePointPosition
        Dim eBtmOutsidePos21 As eRelativePointPosition
        Dim eBtmOutsidePos23 As eRelativePointPosition
        
        eTopInsidePos21 = GetRelativePointPosition(21, InsideTopFlange, oMeasurements)
        eTopInsidePos23 = GetRelativePointPosition(23, InsideTopFlange, oMeasurements)
        eBtmInsidePos21 = GetRelativePointPosition(21, InsideBottomFlange, oMeasurements)
        eBtmInsidePos23 = GetRelativePointPosition(23, InsideBottomFlange, oMeasurements)
        eTopOutsidePos21 = GetRelativePointPosition(21, Top, oMeasurements)
        eTopOutsidePos23 = GetRelativePointPosition(23, Top, oMeasurements)
        eBtmOutsidePos21 = GetRelativePointPosition(21, Bottom, oMeasurements)
        eBtmOutsidePos23 = GetRelativePointPosition(23, Bottom, oMeasurements)
    
        If eTopInsidePos21 = Below Then
            ' -----------------------------------------------------------------------
            ' Inside clearance is distance from Pt 21 to the inside of the top flange
            ' -----------------------------------------------------------------------
            dInsideClearance = oMeasurements.Item("DimPt21ToTopInside")
            
            ' ----------------------------------------------------------
            ' If the inside of the bottom flange is below the top corner
            ' ----------------------------------------------------------
            If eBtmInsidePos21 = Above Then
                ' -------------------------------------------------------------
                ' If the inside of the bottom flange is below the bottom corner
                ' -------------------------------------------------------------
                If eBtmInsidePos23 = Above Or eBtmInsidePos23 = Coincident Then ' Conditions 11 and 12 in BoundingEdgeConditions.sha
                    ' -------------------------------------
                    ' Overlap is distance from Pt21 to Pt23
                    ' -------------------------------------
                    dInsideOverlap = dEdgeLength
                    
                    ' ------------------------------------------------------------------
                    ' Outside clearance is distance from Pt23 to flange right
                    ' ------------------------------------------------------------------
                    dOutsideClearance = oMeasurements.Item("DimPt23ToBottomInside")
                ' -------------------------------------------------------------------------------------------------------------------
                ' If the inside of the bottom flange intersects the edge, the overlap is from Pt21 to the inside of the bottom flange
                ' -------------------------------------------------------------------------------------------------------------------
                Else ' Conditions 9 and 10 in BoundingEdgeConditions.sha
                    bIsEdgeToEdge = HasBottomFlange(oBoundedPort.Connectable)
                    dInsideOverlap = oMeasurements.Item("DimPt21ToBottomInside")
                End If
            ElseIf eBtmOutsidePos21 = Above Then ' Conditions 3-8 in BoundingEdgeConditions.sha
                bIsEdgeToEdge = HasBottomFlange(oBoundedPort.Connectable)
                dInsideOverlap = -oMeasurements.Item("DimPt21ToBottomInside")
            End If
        ' ----------------------------------------------------------
        ' If the inside of the top flange is below the bottom corner (cases 4a, 5a, 7a, 8a, 9a, 10a, 11a)
        ' ----------------------------------------------------------
        ElseIf eTopInsidePos23 = Above Or eTopInsidePos23 = Coincident Then
            
            If eTopOutsidePos23 = Below Then '(excludes case 11a)
                bIsEdgeToEdge = HasTopFlange(oBoundedPort.Connectable)
            End If
            
            ' ------------------------------------------------------------------
            ' Outside clearance is distance from Pt23 to flange right
            ' ------------------------------------------------------------------
            dOutsideClearance = oMeasurements.Item("DimPt23ToBottomInside")
            
            'If eBtmInsidePos23 = Below Or eBtmInsidePos23 = Coincident Then
            If eTopInsidePos23 = Above Then ' Avoid very small values like 1.0e-16.  (cases 5a, 8a, 10a)
                dOutsideOverlap = -oMeasurements.Item("DimPt23ToTopInside")
            End If ' Else, cases 4a, 7a, 9a
        ' ---------------------------------------------------
        ' If the inside of the top flange intersects the edge
        ' ---------------------------------------------------
        Else ' cases 2a, 3a, 6a
        
            If eTopInsidePos21 = Above Then '(excludes case 2a)
                bIsEdgeToEdge = HasTopFlange(oBoundedPort.Connectable)
            End If
            
            ' -------------------------------------------------------------
            ' Overlap is distance from Pt23 to the inside of the top flange
            ' -------------------------------------------------------------
            dOutsideOverlap = oMeasurements.Item("DimPt23ToTopInside")
            
            ' ------------------------------------------------------------------
            ' Outside clearance is distance from Pt23 to flange right
            ' ------------------------------------------------------------------
            dOutsideClearance = oMeasurements.Item("DimPt23ToBottomInside")
        End If
    ' -------------------------------------------
    ' If this is overlap at the bounding top edge
    ' -------------------------------------------
    Else
        ' -------------------------------------------------------------
        ' If the inside of the bottom flange is below the bottom corner
        ' -------------------------------------------------------------
        Dim eBtmInsidePos15 As eRelativePointPosition
        Dim eBtmInsidePos17 As eRelativePointPosition
        Dim eTopInsidePos15 As eRelativePointPosition
        Dim eTopInsidePos17 As eRelativePointPosition
        Dim eBtmOutsidePos15 As eRelativePointPosition
        Dim eBtmOutsidePos17 As eRelativePointPosition
        Dim eTopOutsidePos15 As eRelativePointPosition
        Dim eTopOutsidePos17 As eRelativePointPosition
        
        eBtmInsidePos15 = GetRelativePointPosition(15, InsideBottomFlange, oMeasurements)
        eBtmInsidePos17 = GetRelativePointPosition(17, InsideBottomFlange, oMeasurements)
        eTopInsidePos15 = GetRelativePointPosition(15, InsideTopFlange, oMeasurements)
        eTopInsidePos17 = GetRelativePointPosition(17, InsideTopFlange, oMeasurements)
        eBtmOutsidePos15 = GetRelativePointPosition(15, Bottom, oMeasurements)
        eBtmOutsidePos17 = GetRelativePointPosition(17, Bottom, oMeasurements)
        eTopOutsidePos15 = GetRelativePointPosition(15, Top, oMeasurements)
        eTopOutsidePos17 = GetRelativePointPosition(17, Top, oMeasurements)
        
        If eBtmInsidePos17 = Above Then
            ' --------------------------------------------------------------------------
            ' Inside clearance is distance from Pt 17 to the inside of the bottom flange
            ' --------------------------------------------------------------------------
            dInsideClearance = oMeasurements.Item("DimPt17ToBottomInside")
            
            ' ----------------------------------------------------------
            ' If the inside of the top flange is above the bottom corner
            ' ----------------------------------------------------------
            If eTopInsidePos17 = Below Then
                ' -------------------------------------------------------
                ' If the inside of the top flange is above the top corner
                ' -------------------------------------------------------
                If eTopInsidePos15 = Below Or eTopInsidePos15 = Coincident Then ' Conditions 11 and 12 in BoundingEdgeConditions.sha
                    ' -------------------------------------
                    ' Overlap is distance from Pt15 to Pt17
                    ' -------------------------------------
                    dInsideOverlap = dEdgeLength
                    
                    ' ---------------------------------------------------------------
                    ' Outside clearance is distance from Pt15 to inside of top flange
                    ' ---------------------------------------------------------------
                    dOutsideClearance = oMeasurements.Item("DimPt15ToTopInside")
                ' -------------------------------------------------------------------------------------------------------------
                ' If the inside of the top flange intersects the edge, the overlap is from Pt17 to the inside of the top flange
                ' -------------------------------------------------------------------------------------------------------------
                Else ' Conditions 9 and 10 in BoundingEdgeConditions.sha
                    bIsEdgeToEdge = HasTopFlange(oBoundedPort.Connectable)
                    dInsideOverlap = oMeasurements.Item("DimPt17ToTopInside")
                End If
            ElseIf eTopOutsidePos17 = Below Then ' Conditions 3-8 in BoundingEdgeConditions.sha
                bIsEdgeToEdge = HasTopFlange(oBoundedPort.Connectable)
                dInsideOverlap = -oMeasurements.Item("DimPt17ToTopInside")
            End If
        ' --------------------------------------------------------
        ' If the inside of the bottom edge is above the top corner (cases 4a, 5a, 7a, 8a, 9a, 10a, 11a)
        ' --------------------------------------------------------
        ElseIf eBtmInsidePos15 = Below Or eBtmInsidePos15 = Coincident Then
            
            If eBtmOutsidePos15 = Above Then '(excludes case 11a)
                bIsEdgeToEdge = HasBottomFlange(oBoundedPort.Connectable)
            End If
            
            ' ---------------------------------------------------------------
            ' Outside clearance is distance from Pt15 to inside of top flange
            ' ---------------------------------------------------------------
            dOutsideClearance = oMeasurements.Item("DimPt15ToTopInside")
        
            'If HasTopFlange(oBoundedPort.Connectable) Then
            If eBtmInsidePos15 = Below Then ' Avoid very small values like 1.0e-16.  (cases 5a, 8a, 10a)
                dOutsideOverlap = -oMeasurements.Item("DimPt15ToBottomInside")
            End If
            
        ' ------------------------------------------------------
        ' If the inside of the bottom flange intersects the edge
        ' ------------------------------------------------------
        Else
            
            If eBtmInsidePos17 = Below Then '(excludes case 2a)
                bIsEdgeToEdge = HasBottomFlange(oBoundedPort.Connectable)
            End If
            
            ' ----------------------------------------------------------------
            ' Overlap is distance from Pt15 to the inside of the bottom flange
            ' ----------------------------------------------------------------
            dOutsideOverlap = oMeasurements.Item("DimPt15ToBottomInside")
            
            ' ---------------------------------------------------------------
            ' Outside clearance is distance from Pt15 to inside of top flange
            ' ---------------------------------------------------------------
            dOutsideClearance = oMeasurements.Item("DimPt15ToTopInside")
        End If
    
    End If
        
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
End Sub

' |                                                         ' |
' |                                                         ' |
' |---------------------------------------------            ' |
' |              ^                                          ' |--------------------------------------------
' |              |                                          ' |-------------------------------------------
' |              |                                          ' |              ^
' |        Inside clearance                                 ' |              |
' |              |                                          ' |        Inside clearance
' |              |                                          ' |              |
' |              v                                          ' |              v
' +------------------+                                      ' +------------------+
'                    |InsideOverlap                         '                    |
'                    |---------------------------           '                    |
'    Bounding        |---------------------------           '    Bounding        |Inside Overlap
'                    |                                      '                    |
'                    |Outside Overlap                       '                    |
' -------------------+                                      ' -------------------+
'          |        ^                  Bounded              '          |        ^                  Bounded
'          |        |                                       '          |        |
'          | Outside clearance                              '          | Outside clearance
'          |        |                                       '          |        |
'          |        v                                       '          |        v
'          +----------------------------------------        '          +----------------------------------------
          

Private Sub GetFlangeEdgeOverlapAndClearance(oACOrEC As Object, _
                                             bBottomEdge As Boolean, _
                                             bBottomFlange As Boolean, _
                                             Optional dInsideOverlap As Double, _
                                             Optional dOutsideOverlap As Double, _
                                             Optional dInsideClearance As Double, _
                                             Optional dOutsideClearance As Double, _
                                             Optional dEdgeLength As Double, _
                                             Optional bIsEdgeToEdge As Boolean)

    Const METHOD = "MbrMeasurementUtilities::GetFlangeEdgeOverlapAndClearance"

    On Error GoTo ErrorHandler
    Dim sError As String
        
    dInsideOverlap = 0#
    dOutsideOverlap = 0#
    dInsideClearance = 0#
    dOutsideClearance = 0#
    bIsEdgeToEdge = False
        
    ' -----------------------------
    'Get bounding and bounded ports
    ' -----------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    If TypeOf oACOrEC Is IJAppConnection Then
        Dim lStatus As Long
        Dim sMsg As String
        Dim oAppConn As IJAppConnection
        Set oAppConn = oACOrEC
        GetAssemblyConnectionInputs oAppConn, oBoundedPort, oBoundingPort
        
    ElseIf TypeOf oACOrEC Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Set oFeature = oACOrEC
        
        If oFeature.get_StructFeatureType = SF_FlangeCut Then
            Dim oSDOFlange As New StructDetailObjects.FlangeCut
            Set oSDOFlange.object = oACOrEC
            Set oBoundedPort = oSDOFlange.BoundedPort
            Set oBoundingPort = oSDOFlange.BoundingPort
        ElseIf oFeature.get_StructFeatureType = SF_WebCut Then
            Dim oSDOWeb As New StructDetailObjects.WebCut
            Set oSDOWeb.object = oACOrEC
            Set oBoundedPort = oSDOWeb.BoundedPort
            Set oBoundingPort = oSDOWeb.BoundingPort
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    ' ------------------------
    ' Not available for a tube
    ' ------------------------
    If IsTubularMember(oBoundingPort.Connectable) Then
        Exit Sub
    End If
    
    ' ----------------------------------------------------------------------
    ' Determine the relative orientation of the bounded and bounding objects
    ' ----------------------------------------------------------------------
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    Set oEdgeMapColl = GetEdgeMap(oACOrEC, oBoundingPort, oBoundedPort, , bPenetratesWeb)

    ' -----------------------------
    ' Exit if not flange-penetrated
    ' -----------------------------
    If bPenetratesWeb Then
        Exit Sub
    End If
    
    ' -----------------------------------------------------------
    ' Exit if the bounding alias does not have the specified edge
    ' -----------------------------------------------------------
    Dim RealEdgeID As JXSEC_CODE
    If bBottomEdge Then
        RealEdgeID = ReverseMap(JXSEC_BOTTOM_FLANGE_RIGHT, oEdgeMapColl)
    Else
        RealEdgeID = ReverseMap(JXSEC_TOP_FLANGE_RIGHT, oEdgeMapColl)
    End If
    
    If RealEdgeID = JXSEC_UNKNOWN Then
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------
    ' Determine the relative position of the bounded and bounding objects
    ' -------------------------------------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection
    
    GetConnectedEdgeInfo oACOrEC, _
                         oBoundedPort, _
                         oBoundingPort, _
                         eTopOrWL, _
                         eBottomOrWR, _
                         eInsideTFOrTFL, _
                         eInsideBFOrTFR, _
                         oMeasurements, , , , _
                         Not bBottomFlange
    
    ' -------------------------
    ' Calculate the edge length
    ' -------------------------
    If bBottomEdge Then
        If Not KeyExists("DimPt21ToPt23", oMeasurements) Then
            Exit Sub
        Else
            dEdgeLength = oMeasurements.Item("DimPt21ToPt23")
        End If
    Else
        If Not KeyExists("DimPt15ToPt17", oMeasurements) Then
            Exit Sub
        Else
            dEdgeLength = oMeasurements.Item("DimPt15ToPt17")
        End If
    End If
    
    ' --------------------------------------------------
    ' Determine what flanges exist on the bounded object
    ' --------------------------------------------------
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    Dim bHasRightFlange As Boolean
    Dim bHasLeftFlange As Boolean

    CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

    If bBottomFlange Then
        bHasRightFlange = bBFR
        bHasLeftFlange = bBFL
    Else
        bHasRightFlange = bTFR
        bHasLeftFlange = bTFL
    End If
    
    ' ----------------------------------------------
    ' If this is overlap at the bounding bottom edge
    ' ----------------------------------------------
    If bBottomEdge Then
        
        ' --------------------------------------------------------------------------------
        ' Get position of the four edges relative to the top and bottom corner of the edge
        ' --------------------------------------------------------------------------------
        Dim eWebRightPos21 As eRelativePointPosition
        Dim eWebRightPos23 As eRelativePointPosition
        Dim eWebLeftPos21 As eRelativePointPosition
        Dim eWebLeftPos23 As eRelativePointPosition
        Dim eFlangeRightPos21 As eRelativePointPosition
        Dim eFlangeRightPos23 As eRelativePointPosition
        Dim eFlangeLeftPos21 As eRelativePointPosition
        Dim eFlangeLeftPos23 As eRelativePointPosition
        
        eWebRightPos21 = GetRelativePointPosition(21, WebRight, oMeasurements)
        eWebRightPos23 = GetRelativePointPosition(21, WebRight, oMeasurements)
        eWebLeftPos21 = GetRelativePointPosition(21, WebLeft, oMeasurements)
        eWebLeftPos23 = GetRelativePointPosition(23, WebLeft, oMeasurements)
        eFlangeRightPos21 = GetRelativePointPosition(21, FlangeRight, oMeasurements)
        eFlangeRightPos23 = GetRelativePointPosition(23, FlangeRight, oMeasurements)
        eFlangeLeftPos21 = GetRelativePointPosition(21, FlangeLeft, oMeasurements)
        eFlangeLeftPos23 = GetRelativePointPosition(23, FlangeLeft, oMeasurements)
        
        ' ------------------------------------------------------
        ' If the web is entirely inside the edge (cases 1 and 2)
        ' ------------------------------------------------------
        If eWebRightPos21 = Below Or eWebRightPos21 = Coincident Then
        
            ' -----------------------------------------------------------
            ' There is no overlap or clearance if entire flange is inside
            ' -----------------------------------------------------------
            If eFlangeRightPos21 = Below Then
                Exit Sub
            End If
            
            ' ---------------------------------------------------------
            ' Inside clearance is distance from top corner to web right
            ' ---------------------------------------------------------
            dInsideClearance = oMeasurements.Item("DimPt21ToWR")
            
            ' ---------------------------------------
            ' If flange right is below the top corner
            ' ---------------------------------------
            If eFlangeRightPos21 = Above Then
                ' ------------------------------------------
                ' If flange right is below the bottom corner
                ' ------------------------------------------
                If eFlangeRightPos23 = Above Then
                    ' ------------------------------------------------------------
                    ' Outside overlap is distance from top corner to bottom corner
                    ' ------------------------------------------------------------
                    dInsideOverlap = oMeasurements.Item("DimPt21ToPt23")
                    
                    ' ----------------------------------------------------------------
                    ' Outside clearance is distance from bottom corner to flange right
                    ' ----------------------------------------------------------------
                    dOutsideClearance = oMeasurements.Item("DimPt23ToFR")
                ' -----------------------------------------------------------------------------------
                ' If flange right intersects the edge, the overlap is from top corner to flange right
                ' -----------------------------------------------------------------------------------
                Else
                    dInsideOverlap = oMeasurements.Item("DimPt21ToFR")
                End If
            End If
        ' -----------------------------------------------------
        ' If the web overlaps the inside corner (cases 3 and 6)
        ' -----------------------------------------------------
        ElseIf (eWebLeftPos21 = Below Or eWebLeftPos21 = Coincident) And (eWebRightPos21 = Above And eWebRightPos23 = Below) Then
            
            bIsEdgeToEdge = True
            
            ' -----------------------------------------------------------------------
            ' Inside overlap is distance from top corner to web left (negative value)
            ' -----------------------------------------------------------------------
            dInsideOverlap = -oMeasurements.Item("DimPt21ToWL")
            
            ' -------------------------------------------------------------------------------------
            ' If there is a left flange, the inside clearance is from the top corner to flange left
            ' -------------------------------------------------------------------------------------
            If bHasLeftFlange Then
                dInsideClearance = oMeasurements.Item("DimPt21ToFL")
            End If
            
            ' ------------------------
            ' If it has a right flange
            ' ------------------------
            If bHasRightFlange Then
                ' ------------------------------------------
                ' Overlap is from bottom corner to web right
                ' ------------------------------------------
                dOutsideOverlap = oMeasurements.Item("DimPt23ToWR")
        
                ' ----------------------------------------------------------------
                ' Outside clearance is distance from bottom corner to flange right
                ' ----------------------------------------------------------------
                dOutsideClearance = oMeasurements.Item("DimPt23ToFR")
            End If
        ' -------------------------------------------------------
        ' If the web completely overlaps the edge (cases 4,5,7,8)
        ' -------------------------------------------------------
        ElseIf (eWebLeftPos21 = Below Or eWebLeftPos21 = Coincident) And (eWebRightPos23 = Above Or eWebRightPos23 = Coincident) Then
            
            bIsEdgeToEdge = True
            
            ' -----------------------------------------------------------------------
            ' Inside overlap is distance from top corner to web left (negative value)
            ' -----------------------------------------------------------------------
            dInsideOverlap = -oMeasurements.Item("DimPt21ToWL")
            
            ' ----------------------------------------------------------------------------
            ' Outside overlap is distance from bottom corner to web right (negative value)
            ' ----------------------------------------------------------------------------
            dInsideOverlap = -oMeasurements.Item("DimPt23ToWR")
            
            ' -------------------------------------------------------------------------------------
            ' If there is a left flange, the inside clearance is from the top corner to flange left
            ' -------------------------------------------------------------------------------------
            If bHasLeftFlange Then
                dInsideClearance = oMeasurements.Item("DimPt21ToFL")
            End If
            
            ' -------------------------------------------------------------------------------------------
            ' If there is a right flange, the outside clearance is from the bottom corner to flange right
            ' -------------------------------------------------------------------------------------------
            If bHasRightFlange Then
                dOutsideClearance = oMeasurements.Item("DimPt23ToFR")
            End If
        ' ------------------------------------------------
        ' If the web is entirely within the edge (case 13)
        ' ------------------------------------------------
        ElseIf eWebLeftPos21 = Above And eWebRightPos23 = Below Then
            
            bIsEdgeToEdge = True
            
            ' -------------------------
            ' If there is a left flange
            ' -------------------------
            If bHasLeftFlange Then
                ' --------------------------------------------------
                ' Inside clearance is from top corner to flange left
                ' --------------------------------------------------
                dInsideClearance = oMeasurements.Item("DimPt21ToFL")
                
                ' ---------------------------------------------
                ' Inside overlap is from top corner to web left
                ' ---------------------------------------------
                dInsideOverlap = oMeasurements.Item("DimPt21ToWL")
            End If
            
            ' --------------------------
            ' If there is a right flange
            ' --------------------------
            If bHasRightFlange Then
                ' -------------------------------------------------
                ' Outside overlap is from bottom corner to web right
                ' -------------------------------------------------
                dOutsideOverlap = oMeasurements.Item("DimPt23ToWR")
                
                ' ----------------------------------------------------------------
                ' Outside clearance is distance from bottom corner to flange right
                ' ----------------------------------------------------------------
                dOutsideClearance = oMeasurements.Item("DimPt23ToFR")
            End If
        ' -------------------------------------------------------
        ' If the web overlaps the outside corner (cases 9 and 10)
        ' -------------------------------------------------------
        ElseIf (eWebLeftPos21 = Above And eWebLeftPos23 = Below) And (eWebRightPos23 = Above Or eWebRightPos23 = Coincident) Then
            
            bIsEdgeToEdge = True
            
            ' -------------------------
            ' If there is a left flange
            ' -------------------------
            If bHasLeftFlange Then
                ' ------------------------------------------------------
                ' Inside overlap is distance from top corner to web left
                ' ------------------------------------------------------
                dInsideOverlap = oMeasurements.Item("DimPt21ToWL")
            
                ' -----------------------------------------------------------
                ' Inside clearance is distance from top corner to flange left
                ' -----------------------------------------------------------
                dInsideClearance = oMeasurements.Item("DimPt21ToFL")
            End If
            
            ' ----------------------------------------------------------------------------
            ' Outside overlap is distance from bottom corner to web right (negative value)
            ' ----------------------------------------------------------------------------
            dOutsideOverlap = -oMeasurements.Item("DimPt23ToWR")
            
            ' ----------------------------------------------------------------------------------------------
            ' If it has a right flange, the outside clearance is distance from bottom corner to flange right
            ' ----------------------------------------------------------------------------------------------
            If bHasRightFlange Then
                dOutsideClearance = oMeasurements.Item("DimPt23ToFR")
            End If
        ' ---------------------------------------------------------
        ' If the web is entirely outside the edge (cases 11 and 12)
        ' ---------------------------------------------------------
        ElseIf eWebLeftPos23 = Above Then
            ' -------------------------
            ' If there is a left flange
            ' -------------------------
            If bHasLeftFlange Then
                ' -----------------------------------------------------
                ' Outside clearance is from outside corner to web left
                ' -----------------------------------------------------
                dOutsideClearance = oMeasurements.Item("DimPt23ToWL")
                
                ' --------------------------------------------
                ' If flange left is at or above the top corner
                ' --------------------------------------------
                If eFlangeLeftPos21 = Below Or eFlangeLeftPos21 = Coincident Then
                    ' --------------------------------------------------
                    ' Inside overlap is from top corner to bottom corner
                    ' --------------------------------------------------
                    dInsideOverlap = oMeasurements.Item("DimPt21ToPt23")
                    
                    ' --------------------------------------------------
                    ' Inside clearance is from top corner to flange left
                    ' --------------------------------------------------
                    dInsideClearance = oMeasurements.Item("DimPt21ToFL")
                ' ----------------------------------
                ' If flange left intersects the edge
                ' ----------------------------------
                ElseIf eFlangeLeftPos23 = Below Then
                    ' ---------------------------------------------------
                    ' Inside overlap is from bottom corner to flange left
                    ' ---------------------------------------------------
                    dInsideOverlap = oMeasurements.Item("DimPt23ToFL")
                End If
            End If
        End If

    ' -------------------------------------------
    ' If this is overlap at the bounding top edge
    ' -------------------------------------------
    Else
        ' --------------------------------------------------------------------------------
        ' Get position of the four edges relative to the top and bottom corner of the edge
        ' --------------------------------------------------------------------------------
        Dim eWebRightPos17 As eRelativePointPosition
        Dim eWebRightPos15 As eRelativePointPosition
        Dim eWebLeftPos17 As eRelativePointPosition
        Dim eWebLeftPos15 As eRelativePointPosition
        Dim eFlangeRightPos17 As eRelativePointPosition
        Dim eFlangeRightPos15 As eRelativePointPosition
        Dim eFlangeLeftPos17 As eRelativePointPosition
        Dim eFlangeLeftPos15 As eRelativePointPosition
        
        eWebRightPos17 = GetRelativePointPosition(17, WebRight, oMeasurements)
        eWebRightPos15 = GetRelativePointPosition(15, WebRight, oMeasurements)
        eWebLeftPos17 = GetRelativePointPosition(17, WebLeft, oMeasurements)
        eWebLeftPos15 = GetRelativePointPosition(15, WebLeft, oMeasurements)
        eFlangeRightPos17 = GetRelativePointPosition(17, FlangeRight, oMeasurements)
        eFlangeRightPos15 = GetRelativePointPosition(15, FlangeRight, oMeasurements)
        eFlangeLeftPos17 = GetRelativePointPosition(17, FlangeLeft, oMeasurements)
        eFlangeLeftPos15 = GetRelativePointPosition(15, FlangeLeft, oMeasurements)
        
        ' ------------------------------------------------------
        ' If the web is entirely inside the edge (cases 1 and 2)
        ' ------------------------------------------------------
        If eWebLeftPos17 = Above Or eWebLeftPos17 = Coincident Then
        
            ' -----------------------------------------------------------
            ' There is no overlap or clearance if entire flange is inside
            ' -----------------------------------------------------------
            If eFlangeLeftPos17 = Above Then
                Exit Sub
            End If
            
            ' -----------------------------------------------------------
            ' Inside clearance is distance from bottom corner to web left
            ' -----------------------------------------------------------
            dInsideClearance = oMeasurements.Item("DimPt17ToWL")
            
            ' -----------------------------------------
            ' If flange left is above the bottom corner
            ' -----------------------------------------
            If eFlangeLeftPos17 = Below Then
                ' ---------------------------------------
                ' If flange left is above the top corner
                ' ---------------------------------------
                If eFlangeLeftPos15 = Below Then
                    ' ------------------------------------------------------------
                    ' Outside overlap is distance from top corner to bottom corner
                    ' ------------------------------------------------------------
                    dInsideOverlap = oMeasurements.Item("DimPt15ToPt17")
                    
                    ' ------------------------------------------------------------
                    ' Outside clearance is distance from top corner to flange left
                    ' ------------------------------------------------------------
                    dOutsideClearance = oMeasurements.Item("DimPt15ToFL")
                ' ------------------------------------------------------------------------------------
                ' If flange left intersects the edge, the overlap is from bottom corner to flange left
                ' ------------------------------------------------------------------------------------
                Else
                    dInsideOverlap = oMeasurements.Item("DimPt17ToFL")
                End If
            End If
        ' -----------------------------------------------------
        ' If the web overlaps the bottom corner (cases 3 and 6)
        ' -----------------------------------------------------
        ElseIf (eWebRightPos17 = Above Or eWebRightPos17 = Coincident) And (eWebLeftPos15 = Above) Then

            bIsEdgeToEdge = True
            
            ' ---------------------------------------------------------------------------
            ' Inside overlap is distance from bottom corner to web right (negative value)
            ' ---------------------------------------------------------------------------
            dInsideOverlap = -oMeasurements.Item("DimPt17ToWR")
            
            ' ------------------------------------------------------------------------------------------
            ' If there is a right flange, the inside clearance is from the bottom corner to flange right
            ' ------------------------------------------------------------------------------------------
            If bHasRightFlange Then
                dInsideClearance = oMeasurements.Item("DimPt17ToFR")
            End If
            
            ' -----------------------
            ' If it has a left flange
            ' -----------------------
            If bHasLeftFlange Then
                ' --------------------------------------
                ' Overlap is from top corner to web left
                ' --------------------------------------
                dOutsideOverlap = oMeasurements.Item("DimPt15ToWL")
        
                ' ------------------------------------------------------------
                ' Outside clearance is distance from top corner to flange left
                ' ------------------------------------------------------------
                dOutsideClearance = oMeasurements.Item("DimPt15ToFL")
            End If
        ' -------------------------------------------------------
        ' If the web completely overlaps the edge (cases 4,5,7,8)
        ' -------------------------------------------------------
        ElseIf (eWebRightPos17 = Above Or eWebRightPos17 = Coincident) And (eWebLeftPos15 = Below Or eWebLeftPos15 = Coincident) Then
            
            bIsEdgeToEdge = True
            
            ' ---------------------------------------------------------------------------
            ' Inside overlap is distance from bottom corner to web right (negative value)
            ' ---------------------------------------------------------------------------
            dInsideOverlap = -oMeasurements.Item("DimPt17ToWR")
            
            ' ------------------------------------------------------------------------
            ' Outside overlap is distance from top corner to web left (negative value)
            ' ------------------------------------------------------------------------
            dInsideOverlap = -oMeasurements.Item("DimPt15ToWL")
            
            ' -------------------------------------------------------------------------------------------
            ' If there is a right flange, the inside clearance is from the bottom corner to flange right
            ' -------------------------------------------------------------------------------------------
            If bHasRightFlange Then
                dInsideClearance = oMeasurements.Item("DimPt17ToFR")
            End If
            
            ' --------------------------------------------------------------------------------------
            ' If there is a left flange, the outside clearance is from the top corner to flange left
            ' --------------------------------------------------------------------------------------
            If bHasLeftFlange Then
                dOutsideClearance = oMeasurements.Item("DimPt15ToFL")
            End If
            
        ' ------------------------------------------------
        ' If the web is entirely within the edge (case 13)
        ' ------------------------------------------------
        ElseIf eWebRightPos17 = Below And eWebLeftPos15 = Above Then
            
            bIsEdgeToEdge = True
            
            ' --------------------------
            ' If there is a right flange
            ' --------------------------
            If bHasRightFlange Then
                ' ---------------------------------------------------------------
                ' Inside clearance is distance from bottom corner to flange right
                ' ---------------------------------------------------------------
                dInsideClearance = oMeasurements.Item("DimPt17ToFR")
                
                ' -------------------------------------------------
                ' Inside overlap is from bottom corner to web right
                ' -------------------------------------------------
                dInsideOverlap = oMeasurements.Item("DimPt17ToWR")
            End If
            
            ' -------------------------
            ' If there is a left flange
            ' -------------------------
            If bHasLeftFlange Then
                ' ---------------------------------------------
                ' Inside overlap is from bottom corner to web left
                ' ---------------------------------------------
                dOutsideOverlap = oMeasurements.Item("DimPt15ToWL")
                                
                ' --------------------------------------------------
                ' Outside clearance is from top corner to flange left
                ' --------------------------------------------------
                dOutsideClearance = oMeasurements.Item("DimPt15ToFL")
            End If
            
        ' -------------------------------------------------------
        ' If the web overlaps the outside corner (cases 9 and 10)
        ' -------------------------------------------------------
        ElseIf (eWebRightPos15 = Below And eWebRightPos17 = Above) And (eWebLeftPos15 = Below Or eWebLeftPos15 = Coincident) Then
            
            bIsEdgeToEdge = True
            
            ' ------------------------
            ' If it has a right flange
            ' ------------------------
            If bHasRightFlange Then
                ' ----------------------------------------------------------
                ' Inside overlap is distance from bottom corner to web right
                ' ----------------------------------------------------------
                dInsideOverlap = oMeasurements.Item("DimPt17ToWR")
        
                ' ---------------------------------------------------------------
                ' Inside clearance is distance from bottom corner to flange right
                ' ---------------------------------------------------------------
                dInsideClearance = oMeasurements.Item("DimPt17ToFR")
            End If
            
            ' ------------------------------------------------------------------------
            ' Outside overlap is distance from top corner to web left (negative value)
            ' ------------------------------------------------------------------------
            dOutsideOverlap = -oMeasurements.Item("DimPt15ToWL")
            
            ' -----------------------------------------------------------------------------------------
            ' If it has a left flange, the outside clearance is distance from top corner to flange left
            ' -----------------------------------------------------------------------------------------
            If bHasLeftFlange Then
                dOutsideClearance = oMeasurements.Item("DimPt15ToFL")
            End If
            
        ' ---------------------------------------------------------
        ' If the web is entirely outside the edge (cases 11 and 12)
        ' ---------------------------------------------------------
        ElseIf eWebRightPos15 = Below Then
            ' --------------------------
            ' If there is a right flange
            ' --------------------------
            If bHasRightFlange Then
                ' -----------------------------------------------------
                ' Outside clearance is from outside corner to web right
                ' -----------------------------------------------------
                dOutsideClearance = oMeasurements.Item("DimPt15ToWR")
                
                ' ------------------------------------------------
                ' If flange right is at or below the bottom corner
                ' ------------------------------------------------
                If eFlangeRightPos17 = Above Or eFlangeRightPos17 = Coincident Then
                    ' --------------------------------------------------
                    ' Inside overlap is from top corner to bottom corner
                    ' --------------------------------------------------
                    dInsideOverlap = oMeasurements.Item("DimPt15ToPt17")
                    
                    ' ------------------------------------------------------
                    ' Inside clearance is from bottom corner to flange right
                    ' ------------------------------------------------------
                    dInsideClearance = oMeasurements.Item("DimPt17ToFR")
                ' -----------------------------------
                ' If flange right intersects the edge
                ' -----------------------------------
                ElseIf eFlangeRightPos15 = Above Then
                    ' -------------------------------------------------
                    ' Inside overlap is from top corner to flange right
                    ' -------------------------------------------------
                    dInsideOverlap = oMeasurements.Item("DimPt15ToFR")
                End If
            End If
        End If
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
End Sub

Public Sub GetEdgeOverlapAndClearance(oACOrEC As Object, _
                                      bIsBottomEdge As Boolean, _
                                      bIsBottomFlange As Boolean, _
                                      Optional dInsideOverlap As Double, _
                                      Optional dOutsideOverlap As Double, _
                                      Optional dInsideClearance As Double, _
                                      Optional dOutsideClearance As Double, _
                                      Optional dEdgeLength As Double, _
                                      Optional bIsEdgeToEdge As Boolean)

    ' -----------------------------
    'Get bounding and bounded ports
    ' -----------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    If TypeOf oACOrEC Is IJAppConnection Then
        Dim lStatus As Long
        Dim sMsg As String

        Dim oAppConn As IJAppConnection
        Set oAppConn = oACOrEC
        GetAssemblyConnectionInputs oAppConn, oBoundedPort, oBoundingPort
    ElseIf TypeOf oACOrEC Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Set oFeature = oACOrEC
        
        If oFeature.get_StructFeatureType = SF_FlangeCut Then
            Dim oSDOFlange As New StructDetailObjects.FlangeCut
            Set oSDOFlange.object = oACOrEC
            Set oBoundedPort = oSDOFlange.BoundedPort
            Set oBoundingPort = oSDOFlange.BoundingPort
        ElseIf oFeature.get_StructFeatureType = SF_WebCut Then
            Dim oSDOWeb As New StructDetailObjects.WebCut
            Set oSDOWeb.object = oACOrEC
            Set oBoundedPort = oSDOWeb.BoundedPort
            Set oBoundingPort = oSDOWeb.BoundingPort
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    If IsWebPenetrated(oBoundingPort, oBoundedPort) Then
        GetWebEdgeOverlapAndClearance oACOrEC, bIsBottomEdge, dInsideOverlap, dOutsideOverlap, dInsideClearance, dOutsideClearance, dEdgeLength, bIsEdgeToEdge
    Else
        GetFlangeEdgeOverlapAndClearance oACOrEC, bIsBottomEdge, bIsBottomFlange, dInsideOverlap, dOutsideOverlap, dInsideClearance, dOutsideClearance, dEdgeLength, bIsEdgeToEdge
    End If
    
End Sub

'*************************************************************************
' Function
'   AddMeasurementDimInfo
'   (like AddMeasurementSymbolDimInfo, but value is passed in rather than measurement symbol)
' Abstract
'   The parameter name and value are appended to the end of the string
'   passed in.  The value is truncated to 6 decimal places. The string is semi-colon delimited:
'   "DimPt11ToPt15=1.558790;DimPt15ToPt17=2.469276;DimPt17ToPt18=0.380347;"
'***************************************************************************
Private Sub AddMeasurementDimInfo(ByRef oMeasurements As Collection, dParameter As Double, sParameter As String, ByRef outDimInfo As String)

    oMeasurements.Add dParameter, sParameter
    outDimInfo = outDimInfo & sParameter & "=" & Round(dParameter, 6) & ";"
    
End Sub

'******************************************************************************************
' Function
'   GetMeasurements
' Abstract
'   Compute the measurements that were formally obtained using measurement symbols
' Arguments
'   oACorEC -        The assembly connection or end cut object
'   oBoundedPort -   Bounded port of assembly connection or end cut
'   oBoundingPort -  Bounding port of assembly connection or end cut
'   oBoundedData -   Collection of useful data about the bounded part
'                    Already populated by the one method expected to call this one, so passed in to avoid recalculating
'   sectionAlias -   Section alias as determined by mapping rule
'                    Mapping rule already called by the one method expected to call this one, so passed in to avoid recalculating
'   bPenetratesWeb - Flag indicating whether web or flange is more directly intersected by the bounding axis
'                    Determined by mapping rule.  See comment for sectionAlias
'   bUseTopFlange -  Flag indicating whether or not the measurements should be taken using the top or bottom flange, when flange-penetrated
'   oEdgeMap -       Collection from Edge Mapping Rule.  See comment for sectionAlias.
'   oMeasurements -  Output of key/value pairs.  The keys are the same as the outputs from the deprecated measurement symbols.
'   sCacheString -   Measurements as semi-colon delimted string that can be cached for increased performance.
'******************************************************************************************
Private Sub GetMeasurements(oACOrEC As Object, _
                            oBoundedPort As IJPort, _
                            oBoundingPort As IJPort, _
                            oBoundedData As MemberConnectionData, _
                            sectionAlias As Long, _
                            bPenetratesWeb As Boolean, _
                            bUseTopFlange As Boolean, _
                            oEdgeMap As Collection, _
                            ByRef oMeasurements As Collection, _
                            ByRef sCacheString As String)

    On Error GoTo ErrorHandler
    Dim oOrgMat As DT4x4
    
    Dim oBoundingMemberCommon As ISPSMemberPartCommon
    
    If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then
        'Get the bounding Member Part
        Set oBoundingMemberCommon = oBoundingPort.Connectable
    End If
    
    ' -------------------
    ' Get sketching plane
    ' -------------------
    Dim x As Double
    Dim y As Double
    Dim z As Double
        
    Dim oSketchPlane As IJPlane
    Set oSketchPlane = New Plane3d
    
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR
    If bTFL = False And bTFR = False Then bUseTopFlange = False
    
    GetSketchPlaneForMSSymbol oBoundedPort, oBoundingPort, bPenetratesWeb, bUseTopFlange, oSketchPlane, False
    
    Dim oSketchU As IJDVector
    Dim oSketchV As IJDVector
    Dim oSketchW As IJDVector
    Dim oSketchOrigin As IJDPosition
    
    Set oSketchU = New dVector
    Set oSketchV = New dVector
    Set oSketchOrigin = New DPosition
    
    'From from sketching plane interface U, V vectors and root point are obtained. W vector is computed as U x V
    oSketchPlane.GetRootPoint x, y, z
    oSketchOrigin.Set x, y, z
    oSketchPlane.GetUDirection x, y, z
    oSketchU.Set x, y, z
    oSketchPlane.GetVDirection x, y, z
    oSketchV.Set x, y, z
    Set oSketchW = oSketchU.Cross(oSketchV)
    
    ' ---------------------------------------------------
    ' Get matrix for transforming bounding points into 3D
    ' ---------------------------------------------------
    ' Get intersection of bounding axis and sketching plane
    Dim oBoundingCurve As Object
    Dim oWireUtil As IJSGOWireBodyUtilities
    Set oWireUtil = New SGOWireBodyUtilities
    Dim oGeomOps As New IMSModelGeomOps.DGeomOpsIntersect
    Dim oIntersection As Object
        
    Dim oTmpLandingCrv As IJWireBody
    If TypeOf oBoundingPort.Connectable Is IJStiffener Then
        Set oTmpLandingCrv = GetProfilePartLandingCurve(oBoundingPort.Connectable)
    Else
        'Use axis curve of bounding member as bounding curve
        'Set oBoundingCurve = oBoundingData.AxisCurve
        Dim osdoBdg As New StructDetailObjects.MemberPart
        Set osdoBdg.object = oBoundingPort.Connectable
        Dim oThkDir As IJDVector
        Dim bThkVal As Boolean
        osdoBdg.LandingCurve oTmpLandingCrv, oThkDir, bThkVal
    End If
    oWireUtil.ExtendWire oTmpLandingCrv, 2, 2, oBoundingCurve 'Extend bounding curve by two meeters
    oGeomOps.PlaceIntersectionObject Nothing, oBoundingCurve, oSketchPlane, Nothing, oIntersection
        
    If Not TypeOf oIntersection Is IJPointsGraphBody Then
        Exit Sub
    End If
    
    Dim oPointsGraph As IJPointsGraphBody
    Set oPointsGraph = oIntersection
    
    Dim oStructSymbolTools As SP3DStructGenericTools.IJStructSymbolTools
    Set oStructSymbolTools = New SP3DStructGenericTools.StructSymbolTools
    
    Dim oEnumVariant As IEnumVARIANT
    Dim oVertexList As IJElements
    Dim EnumVariantCount As Integer
    
    oPointsGraph.EnumPositions oEnumVariant, EnumVariantCount
    oStructSymbolTools.TransformIEnumVariantToIJElements oEnumVariant, oVertexList
        
    'Below is point of intersection of bounding axis and sketching plane
    Dim oPosOnAxis As IJDPosition
    Set oPosOnAxis = oVertexList.Item(1)
    
    Dim oBoundingMatrixAtSketchRoot As IJDT4x4
    
    'Get bounding matrix at point of intersection
    If Not oBoundingMemberCommon Is Nothing Then
        oBoundingMemberCommon.Rotation.GetTransformAtPosition oPosOnAxis.x, oPosOnAxis.y, oPosOnAxis.z, oBoundingMatrixAtSketchRoot, Nothing
    Else
        Dim oProfileAttributes As IJProfileAttributes
        Set oProfileAttributes = New ProfileUtils
        oProfileAttributes.GetProfileOrientationMatrix oBoundingPort.Connectable, oPosOnAxis, oBoundingMatrixAtSketchRoot
    End If
    
    Dim oBoundingX As IJDVector
    Dim oBoundingY As IJDVector
    Dim oBoundingZ As IJDVector
    
    Set oBoundingX = New dVector
    Set oBoundingY = New dVector
    Set oBoundingZ = New dVector
    
    oBoundingX.x = oBoundingMatrixAtSketchRoot.IndexValue(0)
    oBoundingX.y = oBoundingMatrixAtSketchRoot.IndexValue(1)
    oBoundingX.z = oBoundingMatrixAtSketchRoot.IndexValue(2)
    
    oBoundingY.x = oBoundingMatrixAtSketchRoot.IndexValue(4)
    oBoundingY.y = oBoundingMatrixAtSketchRoot.IndexValue(5)
    oBoundingY.z = oBoundingMatrixAtSketchRoot.IndexValue(6)
    
    oBoundingZ.x = oBoundingMatrixAtSketchRoot.IndexValue(8)
    oBoundingZ.y = oBoundingMatrixAtSketchRoot.IndexValue(9)
    oBoundingZ.z = oBoundingMatrixAtSketchRoot.IndexValue(10)
    
    ' --------------------------------------------------------------------------------
    ' Determine if we need to flip the matrix based on bounding matrix and mirror flag
    ' --------------------------------------------------------------------------------
    ' Somehow the mapping depends on maintaing a left-hand rule

    If Not oBoundingMemberCommon Is Nothing Then
        
        Dim oCalcZ As IJDVector
        Set oCalcZ = oBoundingX.Cross(oBoundingY)
        
        If oCalcZ.Dot(oBoundingZ) < 0# And oBoundingMemberCommon.Rotation.Mirror Or _
           oCalcZ.Dot(oBoundingZ) > 0# And Not oBoundingMemberCommon.Rotation.Mirror Then
           
        oBoundingMatrixAtSketchRoot.IndexValue(4) = -oBoundingMatrixAtSketchRoot.IndexValue(4)
        oBoundingMatrixAtSketchRoot.IndexValue(5) = -oBoundingMatrixAtSketchRoot.IndexValue(5)
        oBoundingMatrixAtSketchRoot.IndexValue(6) = -oBoundingMatrixAtSketchRoot.IndexValue(6)
        
        End If
    End If
    
    ' ------------------------------
    ' Determine which points we need
    ' ------------------------------
    Dim eBoundingAlias As eBounding_Alias
    eBoundingAlias = GetBoundingAliasSimplified(sectionAlias)
    
    Dim oCardinalPointCol As New Collection
    
    Select Case eBoundingAlias
    
        Case Web, FlangeLeftAndRightBottomWebs, FlangeLeftAndRightTopWebs, FlangeLeftAndRightWebs
            oCardinalPointCol.Add 3
            oCardinalPointCol.Add 11
            oCardinalPointCol.Add 15
            oCardinalPointCol.Add 23
        Case WebTopFlangeRight
            oCardinalPointCol.Add 3
            oCardinalPointCol.Add 11
            oCardinalPointCol.Add 15
            oCardinalPointCol.Add 17
            oCardinalPointCol.Add 18
            oCardinalPointCol.Add 23
        Case WebBuiltUpTopFlangeRight
            oCardinalPointCol.Add 3
            oCardinalPointCol.Add 11
            oCardinalPointCol.Add 14
            oCardinalPointCol.Add 15
            oCardinalPointCol.Add 17
            oCardinalPointCol.Add 18
            oCardinalPointCol.Add 23
            oCardinalPointCol.Add 50
        Case WebBottomFlangeRight
            oCardinalPointCol.Add 3
            oCardinalPointCol.Add 11
            oCardinalPointCol.Add 15
            oCardinalPointCol.Add 20
            oCardinalPointCol.Add 21
            oCardinalPointCol.Add 23
        Case WebBuiltUpBottomFlangeRight
            oCardinalPointCol.Add 3
            oCardinalPointCol.Add 11
            oCardinalPointCol.Add 15
            oCardinalPointCol.Add 20
            oCardinalPointCol.Add 21
            oCardinalPointCol.Add 23
            oCardinalPointCol.Add 24
            oCardinalPointCol.Add 51
        Case WebTopAndBottomRightFlanges
            oCardinalPointCol.Add 3
            oCardinalPointCol.Add 11
            oCardinalPointCol.Add 15
            oCardinalPointCol.Add 17
            oCardinalPointCol.Add 18
            oCardinalPointCol.Add 20
            oCardinalPointCol.Add 21
            oCardinalPointCol.Add 23
        Case Tube
            oCardinalPointCol.Add 3
            oCardinalPointCol.Add 11
            oCardinalPointCol.Add 15
            oCardinalPointCol.Add 23
    End Select
        
    Dim catalogPos As Collection
    Dim tempPos As IJDPosition
    Dim mappedCP As Long
    Dim trueCP As Long
    
    Set catalogPos = New Collection
    
    Dim i As Long
    Dim j As Long
            
    For i = 1 To oCardinalPointCol.Count
        'Initialze a temporary position and add it to collection
        Set tempPos = New DPosition
        tempPos.z = 0#
        catalogPos.Add tempPos, CStr(oCardinalPointCol.Item(i))
    Next i
                 
    ' --------------
    ' Map the points
    ' --------------
    Dim oPointMap As Collection
    Set oPointMap = GetPointMappingFromEdgeMapping(oEdgeMap)
    
    ' ------------------------------------------------
    ' If bounded and bounding are linear and untwisted
    ' ------------------------------------------------
    Dim isBoundedIsTwisted As Boolean
    Dim isBoundingIsTwisted As Boolean
    Dim oProfilePart As IJProfilePart
    
    If TypeOf oBoundedPort.Connectable Is IJProfilePart Then
        Set oProfilePart = oBoundedPort.Connectable
        isBoundedIsTwisted = oProfilePart.IsTwisted()
    End If
    
    If TypeOf oBoundingPort.Connectable Is IJProfilePart Then
        Set oProfilePart = oBoundingPort.Connectable
        isBoundingIsTwisted = oProfilePart.IsTwisted()
    End If
    
    Dim bDebug As Boolean
    bDebug = False
    
    Dim oModel As IJDModelBody
    Dim iFileNumber
    Dim sFileName As String
    Dim entityNumber As Long
    entityNumber = 2
    
    If bDebug Then
        Set oModel = oBoundedData.MemberPart
        oModel.DebugToSATFile "C:/000/bounded.sat"
        Set oModel = oBoundingPort.Connectable
        oModel.DebugToSATFile "C:/000/bounding.sat"

        iFileNumber = FreeFile
        sFileName = "M:\CommonShip\Tools\Bin\ship_Debug.txt"
        Open sFileName For Output As #iFileNumber
        Print #iFileNumber, "(view:gl)"
        Print #iFileNumber, "(part:load ""C:/000/bounded.sat"" )"
        Print #iFileNumber, "(part:load ""C:/000/bounding.sat"" )"
        Print #iFileNumber, "(iso)"
        Print #iFileNumber, "(zoom-all)"
        Print #iFileNumber, "(view:edges ON)"
        Print #iFileNumber, "(view:vertices ON)"
        Print #iFileNumber, "(view:shaded OFF)"
    End If
    
    ' Until CR XXXXXX is implemented, use intersection to measure bounding stiffeners

    If TypeOf oBoundedData.AxisCurve Is IJLine And TypeOf oBoundingCurve Is IJLine And _
       Not isBoundedIsTwisted And Not isBoundingIsTwisted And _
       Not TypeOf oBoundingPort.Connectable Is IJStiffener Then
        'Case: bounded and bounding are linear and untwisted, bounding is not stiffner.
        
        ' -----------------------------------------------------------
        ' Get the catalog position of the load point / cardinal point
        ' -----------------------------------------------------------
        Dim boundingCP As Long
        Dim boundingCPX As Double
        Dim boundingCPY As Double
        
        Dim oStiffenerSection As IJDProfileSection
        Dim oMemberSection As ISPSCrossSection
            
        If TypeOf oBoundingPort.Connectable Is IJStiffener Then
            Set oStiffenerSection = oBoundingPort.Connectable
            boundingCP = oStiffenerSection.LoadPoint
            oStiffenerSection.GetKeyPointCatalogCoordinates boundingCP, boundingCPX, boundingCPY
        Else
            Set oMemberSection = oBoundingMemberCommon.CrossSection
            boundingCP = oMemberSection.CardinalPoint
            oMemberSection.GetCardinalPointOffset boundingCP, boundingCPX, boundingCPY
        End If
        
        ' -----------------------
        ' Loop through the points
        ' -----------------------
        For i = 1 To oCardinalPointCol.Count
        
            ' ------------------------------------------------------------------------
            ' Get the position of the mapped cardinal point in the cross section space
            ' ------------------------------------------------------------------------
            mappedCP = oCardinalPointCol.Item(i)
            trueCP = oPointMap.Item(CStr(mappedCP))
                
            If TypeOf oBoundingPort.Connectable Is IJStiffener Then
                oStiffenerSection.GetKeyPointCatalogCoordinates trueCP, x, y
                catalogPos.Item(CStr(mappedCP)).x = x
                catalogPos.Item(CStr(mappedCP)).y = y
            Else
                GetStiffenerLoadPointPositionOnMemberSection oBoundingPort.Connectable, trueCP, x, y
            End If
            
            If TypeOf oBoundingPort.Connectable Is IJStiffener Then
                catalogPos.Item(CStr(mappedCP)).x = x - boundingCPX
                catalogPos.Item(CStr(mappedCP)).y = y - boundingCPY
                catalogPos.Item(CStr(mappedCP)).z = 0#
            Else
                catalogPos.Item(CStr(mappedCP)).y = x - boundingCPX
                catalogPos.Item(CStr(mappedCP)).z = y - boundingCPY
                catalogPos.Item(CStr(mappedCP)).x = 0#
            End If
            
            ' -----------------------
            ' Transform into 3D space
            ' -----------------------
            Set tempPos = oBoundingMatrixAtSketchRoot.TransformPosition(catalogPos.Item(CStr(mappedCP)))
            catalogPos.Item(CStr(mappedCP)).Set tempPos.x, tempPos.y, tempPos.z
            
            If bDebug Then
                Select Case mappedCP
                    Case 20, 21, 23, 3
                        entityNumber = entityNumber + 1
                        Print #iFileNumber, "(solid:sphere " & catalogPos.Item(CStr(mappedCP)).x & " " & catalogPos.Item(CStr(mappedCP)).y & " " & catalogPos.Item(CStr(mappedCP)).z & " 0.01)"
                        Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")YELLOW)"
                    Case 11, 15, 17, 18
                        entityNumber = entityNumber + 1
                        Print #iFileNumber, "(solid:sphere " & catalogPos.Item(CStr(mappedCP)).x & " " & catalogPos.Item(CStr(mappedCP)).y & " " & catalogPos.Item(CStr(mappedCP)).z & " 0.01)"
                        Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")MAGENTA)"
                End Select
            End If
            
            ' ----------------------------
            ' Project onto sketching plane
            ' ----------------------------
            Dim oBoundingPointToPlaneRoot As IJDVector
            Set oBoundingPointToPlaneRoot = oSketchOrigin.Subtract(catalogPos.Item(CStr(mappedCP)))
            
            Dim u As Double
            Dim offsetVector As IJDVector
            
            If TypeOf oBoundingPort.Connectable Is IJStiffener Then
                u = oSketchW.Dot(oBoundingPointToPlaneRoot) / oSketchW.Dot(oBoundingZ)
                Set offsetVector = oBoundingZ.Clone()
            Else
                u = oSketchW.Dot(oBoundingPointToPlaneRoot) / oSketchW.Dot(oBoundingX)
                Set offsetVector = oBoundingX.Clone()
            End If
            
            'Offset Vector
            offsetVector.Length = u
            
            Set tempPos = catalogPos.Item(CStr(mappedCP)).Offset(offsetVector)
            catalogPos.Item(CStr(mappedCP)).Set tempPos.x, tempPos.y, tempPos.z

            If bDebug Then
                Select Case mappedCP
                    Case 20, 21, 23, 3
                        entityNumber = entityNumber + 1
                        Print #iFileNumber, "(solid:sphere " & catalogPos.Item(CStr(mappedCP)).x & " " & catalogPos.Item(CStr(mappedCP)).y & " " & catalogPos.Item(CStr(mappedCP)).z & " 0.01)"
                        Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")RED)"
                    Case 11, 15, 17, 18
                        entityNumber = entityNumber + 1
                        Print #iFileNumber, "(solid:sphere " & catalogPos.Item(CStr(mappedCP)).x & " " & catalogPos.Item(CStr(mappedCP)).y & " " & catalogPos.Item(CStr(mappedCP)).z & " 0.01)"
                        Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")BLUE)"
                End Select
            End If
    
        Next i
    
    ' ---------------------------------------------------------
    ' If either bounded or bounding is not linear and untwisted
    ' ---------------------------------------------------------
    Else
        ' -----------------------
        ' Loop through the points
        ' -----------------------
        For i = 1 To oCardinalPointCol.Count
            ' -------------------------------------------------------------------
            ' Determine the point location on the sketching plane by intersection
            ' -------------------------------------------------------------------
            mappedCP = oCardinalPointCol.Item(i)
            Set tempPos = GetMappedPointLocationByIntersection(oBoundingPort.Connectable, oSketchPlane, mappedCP, sectionAlias, oEdgeMap)
            catalogPos.Item(CStr(mappedCP)).Set tempPos.x, tempPos.y, tempPos.z
        
            If bDebug Then
                Select Case mappedCP
                    Case 20, 21, 23, 3
                        entityNumber = entityNumber + 1
                        Print #iFileNumber, "(solid:sphere " & catalogPos.Item(CStr(mappedCP)).x & " " & catalogPos.Item(CStr(mappedCP)).y & " " & catalogPos.Item(CStr(mappedCP)).z & " 0.01)"
                        Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")RED)"
                    Case 11, 15, 17, 18
                        entityNumber = entityNumber + 1
                        Print #iFileNumber, "(solid:sphere " & catalogPos.Item(CStr(mappedCP)).x & " " & catalogPos.Item(CStr(mappedCP)).y & " " & catalogPos.Item(CStr(mappedCP)).z & " 0.01)"
                        Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")BLUE)"
                End Select
            End If

        Next i
    End If
    
    ' -----------------------------------------------------------------------------------------
    ' Get point on the bottom (web-penetrated) or right (flange-penetrated) of the bounded part
    ' -----------------------------------------------------------------------------------------
    Dim oPortOnBounded As IJPort
    Dim oPortOnBoundedGeom As IJModelBody
    Dim oPointOnBounded As IJDPosition
    Dim oModelBodyUtility As IJSGOModelBodyUtilities
    Dim dist As Double

    Dim bIsTubeBoundedCase As Boolean
    If IsTubularMember(oBoundedData.MemberPart) Then
        bIsTubeBoundedCase = True
    End If
    Dim eXSecCode As JXSEC_CODE
    If bPenetratesWeb Then
        eXSecCode = JXSEC_BOTTOM
    Else
        If bUseTopFlange Then
            If bTFR Then
                eXSecCode = JXSEC_TOP_FLANGE_RIGHT
                If bIsTubeBoundedCase Then eXSecCode = JXSEC_TOP
            Else
                eXSecCode = JXSEC_WEB_RIGHT
            End If
        Else
            If bBFR Then
                eXSecCode = JXSEC_BOTTOM_FLANGE_RIGHT
                If bIsTubeBoundedCase Then eXSecCode = JXSEC_BOTTOM
            Else
                eXSecCode = JXSEC_WEB_RIGHT
            End If
        End If
    End If
    
    Dim oTempAxis As IJCurve
    
    'Prepare sketch matrix using U, V and W vectors and sketch origin
    Dim oSketchMatrix As IJDT4x4
    Set oSketchMatrix = New DT4x4
    
    oSketchMatrix.IndexValue(0) = oSketchW.x
    oSketchMatrix.IndexValue(1) = oSketchW.y
    oSketchMatrix.IndexValue(2) = oSketchW.z
    oSketchMatrix.IndexValue(4) = oSketchU.x
    oSketchMatrix.IndexValue(5) = oSketchU.y
    oSketchMatrix.IndexValue(6) = oSketchU.z
    oSketchMatrix.IndexValue(8) = oSketchV.x
    oSketchMatrix.IndexValue(9) = oSketchV.y
    oSketchMatrix.IndexValue(10) = oSketchV.z
    oSketchMatrix.IndexValue(12) = oSketchOrigin.x
    oSketchMatrix.IndexValue(13) = oSketchOrigin.y
    oSketchMatrix.IndexValue(14) = oSketchOrigin.z

    'Keep a clone of above sketch matrix (which will be used to translate bounded axis to top/bottom)
    Set oOrgMat = oSketchMatrix.Clone
    
    oSketchMatrix.Invert
    
    Dim oBoundedPosition As IJDPosition
    
    'For Both Along case- Split None case
    
    If oBoundedData.ePortId = SPSMemberAxisAlong And TypeOf oBoundingPort Is ISPSSplitAxisAlongPort Then
            
        Dim oBoundedCurve As IJCurve
        Set oBoundedCurve = oBoundedData.AxisCurve
        
        Dim oBoundingMember As ISPSMemberPartCommon
        Set oBoundingMember = oBoundedPort.Connectable
        
        Dim dMinDist As Double
        Dim dx1 As Double, dy1 As Double, dz1 As Double
        Dim dX2 As Double, dY2 As Double, dZ2 As Double
        
        Dim oPosition As IJDPosition
        Dim dx As Double
        Dim dy As Double
        Dim dz As Double
                
        Dim oBoundingAxisCurve As IJCurve
        Set oBoundingAxisCurve = oBoundingMember.Axis
        
        'Get Bounded Location
        oBoundedCurve.DistanceBetween oBoundingAxisCurve, dMinDist, dx1, dy1, dz1, dX2, dY2, dZ2
        
        Set oBoundedPosition = New DPosition
        oBoundedPosition.Set dx1, dy1, dz1
            
    End If
    If Not bIsTubeBoundedCase Then
        Set oPortOnBounded = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, eXSecCode)
        
        Set oPortOnBoundedGeom = oPortOnBounded.Geometry
        Set oModelBodyUtility = New SGOModelBodyUtilities
        
        oModelBodyUtility.GetClosestPointOnBody oPortOnBoundedGeom, oSketchOrigin, oPointOnBounded, dist
    Else
        Dim oBddMat As DT4x4
        Set oBddMat = oOrgMat.Clone
        
        'Update matrix with bounded location
        oBddMat.IndexValue(12) = oBoundedData.Matrix.IndexValue(12)
        oBddMat.IndexValue(13) = oBoundedData.Matrix.IndexValue(13)
        oBddMat.IndexValue(14) = oBoundedData.Matrix.IndexValue(14)
        
        'Get bounded axis curve at top/bottom
        If eXSecCode = JXSEC_TOP Then
            TranslateAxisCurve oBoundedData.AxisPort, oBoundedData.MemberPart, _
                oBoundedData.AxisCurve, oBddMat, True, oTempAxis, , oBoundedPosition
            
        Else
            TranslateAxisCurve oBoundedData.AxisPort, oBoundedData.MemberPart, _
                oBoundedData.AxisCurve, oBddMat, False, oTempAxis, , oBoundedPosition
        End If
        Dim oCrv As IJCurve
        Dim oP3d As Point3d
        Set oP3d = New Point3d
        If TypeOf oTempAxis Is IJCurve Then
            Set oCrv = oTempAxis
            oP3d.SetPoint oSketchOrigin.x, oSketchOrigin.y, oSketchOrigin.z
            Dim dSrcX As Double
            Dim dSrcY As Double
            Dim dSrcZ As Double
            Dim dInX As Double
            Dim dInY As Double
            Dim dInZ As Double

            'Store point on bounded axis curve which is closest from sketch origin
            oCrv.DistanceBetween oP3d, dist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
            Set oPointOnBounded = New DPosition
            oPointOnBounded.Set dSrcX, dSrcY, dSrcZ
        Else
            'Not handled
        End If
    End If
    ' ----------------------------------------------------------
    ' Get matrix for transforming 3D points into sketching plane
    ' ----------------------------------------------------------
    
    ' ----------------------------
    ' Tranform the bounding points
    ' ----------------------------
    For i = 1 To oCardinalPointCol.Count
        
        mappedCP = oCardinalPointCol.Item(i)
        Set tempPos = oSketchMatrix.TransformPosition(catalogPos.Item(CStr(mappedCP)))
        catalogPos.Item(CStr(mappedCP)).Set tempPos.x, tempPos.y, tempPos.z
        If bDebug Then
            Select Case mappedCP
                Case 20, 21, 23, 3
                    entityNumber = entityNumber + 1
                    Print #iFileNumber, "(solid:sphere " & catalogPos.Item(CStr(mappedCP)).x & " " & catalogPos.Item(CStr(mappedCP)).y & " " & catalogPos.Item(CStr(mappedCP)).z & " 0.01)"
                    Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")RED)"
                Case 11, 15, 17, 18
                    entityNumber = entityNumber + 1
                    Print #iFileNumber, "(solid:sphere " & catalogPos.Item(CStr(mappedCP)).x & " " & catalogPos.Item(CStr(mappedCP)).y & " " & catalogPos.Item(CStr(mappedCP)).z & " 0.01)"
                    Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")BLUE)"
            End Select
        End If
            
    Next i

    ' --------------------------
    ' Tranform the bounded point
    ' --------------------------
    Set oPointOnBounded = oSketchMatrix.TransformPosition(oPointOnBounded)
    Set oSketchOrigin = oSketchMatrix.TransformPosition(oSketchOrigin) ' For debugging only.  Should be 0,0,0 (but getting 1,0,0, which works in this case)
        
    If bDebug Then
        entityNumber = entityNumber + 1
        Print #iFileNumber, "(solid:sphere " & oPointOnBounded.x & " " & oPointOnBounded.y & " " & oPointOnBounded.z & " 0.01)"
        Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")CYAN)"
        
        entityNumber = entityNumber + 1
        Print #iFileNumber, "(solid:sphere " & oSketchOrigin.x & " " & oSketchOrigin.y & " " & oSketchOrigin.z & " 0.01)"
        Print #iFileNumber, "(entity:set-color(entity " & entityNumber & ")WHITE)"
        
        entityNumber = entityNumber + 1
        Close #iFileNumber
    End If
    
    ' -----------------------
    ' Get some key dimensions
    ' -----------------------
    Dim dBoundedFlangeThickness As Double
    Dim dBoundedWebThickness As Double
    Dim dBoundedWidth As Double
    Dim dBoundedDepth As Double
    
    If TypeOf oBoundedData.MemberPart Is ISPSMemberPartCommon Then
        Dim oSDOBoundedMember As New StructDetailObjects.MemberPart
        Set oSDOBoundedMember.object = oBoundedData.MemberPart
    
        dBoundedFlangeThickness = oSDOBoundedMember.flangeThickness
        dBoundedWebThickness = oSDOBoundedMember.webThickness
        If Equal(dBoundedWebThickness, 0) Then
            dBoundedWebThickness = GetWidthFromStructDetailObjects(oSDOBoundedMember.object)
        End If
        
        
        dBoundedWidth = oSDOBoundedMember.FlangeLength
        dBoundedDepth = oSDOBoundedMember.Height
    Else
        Dim oSDOBoundedProfile As New StructDetailObjects.ProfilePart
        Set oSDOBoundedProfile.object = oBoundedData.MemberPart
    
        dBoundedFlangeThickness = oSDOBoundedProfile.flangeThickness
        dBoundedWebThickness = oSDOBoundedProfile.webThickness
        dBoundedWidth = oSDOBoundedProfile.FlangeLength
        If StrComp(oSDOBoundedProfile.sectionType, "BUT", vbTextCompare) = 0 Or _
            StrComp(oSDOBoundedProfile.sectionType, "BUTL2", vbTextCompare) = 0 Then
            dBoundedDepth = oSDOBoundedProfile.WebLength + dBoundedFlangeThickness
        Else
            dBoundedDepth = oSDOBoundedProfile.WebLength
        End If
    End If
    
    ' ------------------------------------------
    ' If the bounded is not linear or is twisted
    ' ------------------------------------------
    Set oMeasurements = New Collection
    
    Dim strMeasure As String
    Dim dToWR As Double
    Dim dToWL As Double
                
    If Not TypeOf oBoundedData.AxisCurve Is IJLine Or isBoundedIsTwisted Then
        
        Dim oTransform As IJTransform
        Dim oTangent As IJDVector
        Dim oDummyPoint As IJDPosition

        
        Dim dRatio As Double
            
        ' -----------------
        ' If web-penetrated
        ' -----------------
        If bPenetratesWeb Then
            
            ' ---------------------------------------
            ' Get a wire for the top and bottom ports
            ' ---------------------------------------
            Dim oTopWire As IJWireBody
            Dim oBtmWire As IJWireBody
            
            If Not bIsTubeBoundedCase Then
                Dim oTopPort As IJPort
                Dim oBtmPort As IJPort
                
                Set oTopPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)
                Set oBtmPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM)
            
                oGeomOps.PlaceIntersectionObject Nothing, oTopPort.Geometry, oSketchPlane, Nothing, oTopWire
                oGeomOps.PlaceIntersectionObject Nothing, oBtmPort.Geometry, oSketchPlane, Nothing, oBtmWire
            Else
                TranslateAxisCurve oBoundedData.AxisPort, oBoundedData.MemberPart, _
                    oBoundedData.AxisCurve, oBddMat, True, oTempAxis, , oBoundedPosition
                If TypeOf oTempAxis Is IJWireBody Then Set oTopWire = oTempAxis
        
                TranslateAxisCurve oBoundedData.AxisPort, oBoundedData.MemberPart, _
                    oBoundedData.AxisCurve, oBddMat, False, oTempAxis, , oBoundedPosition
                If TypeOf oTempAxis Is IJWireBody Then Set oBtmWire = oTempAxis
            End If
            
            ' ------------------------------------------
            ' Transform the wires to the sketching plane
            ' ------------------------------------------
            Set oTransform = oTopWire
            oTransform.Transform oSketchMatrix
        
            Set oTransform = oBtmWire
            oTransform.Transform oSketchMatrix
            
            ' --------------------------------
            ' Loop through the bounding points
            ' --------------------------------
            Dim oPointOnTop As IJDPosition
            Dim oPointOnBtm As IJDPosition
            Dim oPointOnTopInside As IJDPosition
            Dim oPointOnBtmInside As IJDPosition
            Dim oBtmToTopDir As IJDVector
            
            Dim dMeasuredDepth As Double
            
            For i = 1 To oCardinalPointCol.Count
                mappedCP = oCardinalPointCol.Item(i)
                
                ' -------------------------------
                ' Project to top and bottom wires
                ' -------------------------------
                oWireUtil.GetClosestPointOnWire oTopWire, catalogPos.Item(CStr(mappedCP)), oPointOnTop, oTangent
                oWireUtil.GetClosestPointOnWire oBtmWire, catalogPos.Item(CStr(mappedCP)), oPointOnBtm, oTangent
                                
                ' ----------------------
                ' Add measurement to top
                ' ----------------------
                dist = oPointOnTop.DistPt(catalogPos.Item(CStr(mappedCP)))
                
                strMeasure = "DimPt" & CStr(mappedCP) & "ToTop"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                                
                ' -------------------------
                ' Add measurement to bottom
                ' -------------------------
                dist = oPointOnBtm.DistPt(catalogPos.Item(CStr(mappedCP)))
                
                strMeasure = "DimPt" & CStr(mappedCP) & "ToBottom"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' -----------------------------------
                ' Add depth measurement at this point
                ' -----------------------------------
                Set oBtmToTopDir = oPointOnTop.Subtract(oPointOnBtm)
                
                dMeasuredDepth = oBtmToTopDir.Length
                
                strMeasure = "DepthAtPt" & CStr(mappedCP)
                AddMeasurementDimInfo oMeasurements, dBoundedWidth, strMeasure, sCacheString
                
                ' -----------------------------
                ' Add measurement to top inside
                ' -----------------------------
                dRatio = dBoundedDepth / dMeasuredDepth
                
                If HasTopFlange(oBoundedData.MemberPart) Then
                    oBtmToTopDir.Length = dMeasuredDepth - (dBoundedFlangeThickness * dRatio)
                    
                    Set oPointOnTopInside = oPointOnBtm.Offset(oBtmToTopDir)
                Else
                    Set oPointOnTopInside = oPointOnTop.Clone()
                End If
                
                dist = oPointOnTopInside.DistPt(catalogPos.Item(CStr(mappedCP)))
                    
                strMeasure = "DimPt" & CStr(mappedCP) & "ToTopInside"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' --------------------------------
                ' Add measurement to bottom inside
                ' --------------------------------
                If HasBottomFlange(oBoundedData.MemberPart) Then
                    oBtmToTopDir.Length = (dBoundedFlangeThickness * dRatio)
                    
                    Set oPointOnBtmInside = oPointOnBtm.Offset(oBtmToTopDir)
                Else
                    Set oPointOnBtmInside = oPointOnBtm.Clone()
                End If
                
                dist = oPointOnBtmInside.DistPt(catalogPos.Item(CStr(mappedCP)))
                    
                strMeasure = "DimPt" & CStr(mappedCP) & "ToBottomInside"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' ------------------------------------------
                ' Add inside depth measurement at this point
                ' ------------------------------------------
                dist = oPointOnBtmInside.DistPt(oPointOnTopInside)
                
                strMeasure = "InsideDepthAtPt" & CStr(mappedCP)
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
            
            Next i
        ' --------------------
        ' If flange-penetrated
        ' --------------------
        Else

            ' ---------------------------------------
            ' Get a wire for the left and right ports
            ' ---------------------------------------
            Dim oLeftPort As IJPort
            Dim oRightPort As IJPort
            
            Set oRightPort = oPortOnBounded
            
            If bUseTopFlange Then
                If bTFL Then
                    Set oLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_LEFT)
                Else
                    Set oLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                End If
            Else
                If bBFL Then
                    Set oLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_LEFT)
                Else
                    Set oLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                End If
            End If
            
            Dim oLeftWire As IJWireBody
            Dim oRightWire As IJWireBody
        
            oGeomOps.PlaceIntersectionObject Nothing, oLeftPort.Geometry, oSketchPlane, Nothing, oLeftWire
            oGeomOps.PlaceIntersectionObject Nothing, oRightPort.Geometry, oSketchPlane, Nothing, oRightWire
            
            ' ------------------------------------------
            ' Transform the wires to the sketching plane
            ' ------------------------------------------
            Set oTransform = oLeftWire
            oTransform.Transform oSketchMatrix
        
            Set oTransform = oRightWire
            oTransform.Transform oSketchMatrix
            
            ' --------------------------------
            ' Loop through the bounding points
            ' --------------------------------
            Dim oPointOnFlangeLeft As IJDPosition
            Dim oPointOnFlangeRight As IJDPosition
            Dim oPointOnWebLeft As IJDPosition
            Dim oPointOnWebRight As IJDPosition
            Dim oRightToLeftDir As IJDVector
            
            Dim dMeasuredWidth As Double
            
            For i = 1 To oCardinalPointCol.Count
                mappedCP = oCardinalPointCol.Item(i)
                
                ' -------------------------------
                ' Project to Left and Right wires
                ' -------------------------------
                oWireUtil.GetClosestPointOnWire oLeftWire, catalogPos.Item(CStr(mappedCP)), oPointOnFlangeLeft, oTangent
                oWireUtil.GetClosestPointOnWire oRightWire, catalogPos.Item(CStr(mappedCP)), oPointOnFlangeRight, oTangent
                                
                ' ----------------------
                ' Add measurement to Left
                ' ----------------------
                dist = oPointOnFlangeLeft.DistPt(catalogPos.Item(CStr(mappedCP)))
                
                strMeasure = "DimPt" & CStr(mappedCP) & "ToFL"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                                
                ' -------------------------
                ' Add measurement to Right
                ' -------------------------
                dist = oPointOnFlangeRight.DistPt(catalogPos.Item(CStr(mappedCP)))
                
                strMeasure = "DimPt" & CStr(mappedCP) & "ToFR"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' -----------------------------------
                ' Add width measurement at this point
                ' -----------------------------------
                Set oRightToLeftDir = oPointOnFlangeLeft.Subtract(oPointOnFlangeRight)
                
                dMeasuredWidth = oRightToLeftDir.Length
                
                strMeasure = "WidthAtPt" & CStr(mappedCP)
                AddMeasurementDimInfo oMeasurements, dBoundedWidth, strMeasure, sCacheString
                
                ' -----------------------------------------
                ' Add measurement to Web Left and Web Right
                ' -----------------------------------------
                dRatio = dBoundedWidth / dMeasuredWidth

                If bUseTopFlange Then
                    If bTFR Then
                        If bTFL Then
                            dToWL = dMeasuredWidth / 2# + (dBoundedWebThickness * dRatio) / 2#
                            dToWR = dMeasuredWidth / 2# - (dBoundedWebThickness * dRatio) / 2#
                        Else
                            dToWL = dMeasuredWidth
                            dToWR = dMeasuredWidth - (dBoundedWebThickness * dRatio)
                        End If
                    Else
                        dToWL = (dBoundedWebThickness * dRatio)
                        dToWR = 0#
                    End If
                Else
                    If bBFR Then
                        If bBFL Then
                            dToWL = dMeasuredWidth / 2# + (dBoundedWebThickness * dRatio) / 2#
                            dToWR = dMeasuredWidth / 2# - (dBoundedWebThickness * dRatio) / 2#
                        Else
                            dToWL = dMeasuredWidth
                            dToWR = dMeasuredWidth - (dBoundedWebThickness * dRatio)
                        End If
                    Else
                        dToWL = (dBoundedWebThickness * dRatio)
                        dToWR = 0#
                    End If
                End If
                
                oRightToLeftDir.Length = dToWL
                Set oPointOnWebLeft = oPointOnFlangeRight.Offset(oRightToLeftDir)
                
                dist = oPointOnWebLeft.DistPt(catalogPos.Item(CStr(mappedCP)))
                    
                strMeasure = "DimPt" & CStr(mappedCP) & "ToWL"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                oRightToLeftDir.Length = dToWR
                Set oPointOnWebRight = oPointOnFlangeRight.Offset(oRightToLeftDir)

                dist = oPointOnWebRight.DistPt(catalogPos.Item(CStr(mappedCP)))
                    
                strMeasure = "DimPt" & CStr(mappedCP) & "ToWR"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' -------------------------------------------
                ' Add web thickness measurement at this point
                ' -------------------------------------------
                dist = oPointOnWebRight.DistPt(oPointOnWebLeft)
                
                strMeasure = "WebThkAtPt" & CStr(mappedCP)
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
            
            Next i
        
        End If

    Else
        'Twisted case
        ' --------------------
        ' Make the bottom zero
        ' --------------------
        For i = 1 To oCardinalPointCol.Count
            mappedCP = oCardinalPointCol.Item(i)
            catalogPos.Item(CStr(mappedCP)).z = catalogPos.Item(CStr(mappedCP)).z - oPointOnBounded.z
        Next i
    
        oPointOnBounded.z = 0#
        
        ' -----------------
        ' If web-penetrated
        ' -----------------
        Dim dVector As IJDVector
    
        If bPenetratesWeb Then
            
            ' -----------------------
            ' Get some key dimensions
            ' -----------------------
            Dim dInsideDepth As Double
            Dim dTopInside As Double
            Dim dBtmInside As Double
            
            If HasTopFlange(oBoundedData.MemberPart) And HasBottomFlange(oBoundedData.MemberPart) Then
                dInsideDepth = dBoundedDepth - 2 * dBoundedFlangeThickness
                dTopInside = dBoundedDepth - dBoundedFlangeThickness
                dBtmInside = dBoundedFlangeThickness
            ElseIf HasTopFlange(oBoundedData.MemberPart) Then
                dInsideDepth = dBoundedDepth - dBoundedFlangeThickness
                dTopInside = dInsideDepth
                dBtmInside = 0#
            ElseIf HasBottomFlange(oBoundedData.MemberPart) Then
                dInsideDepth = dBoundedDepth - dBoundedFlangeThickness
                dTopInside = dBoundedDepth
                dBtmInside = dBoundedFlangeThickness
            Else
                dInsideDepth = dBoundedDepth
                dTopInside = dBoundedDepth
                dBtmInside = 0#
            End If
            
            ' -----------------------
            ' Loop through the points
            ' -----------------------
            For i = 1 To oCardinalPointCol.Count
                
                mappedCP = oCardinalPointCol.Item(i)
        
                ' ----------------------
                ' Add measurement to top
                ' ----------------------
                dist = Abs(catalogPos.Item(CStr(mappedCP)).z - dBoundedDepth)
                strMeasure = "DimPt" & mappedCP & "ToTop"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' -------------------------
                ' Add measurement to bottom
                ' -------------------------
                dist = Abs(catalogPos.Item(CStr(mappedCP)).z)
                strMeasure = "DimPt" & mappedCP & "ToBottom"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' -----------------------------
                ' Add measurement to top inside
                ' -----------------------------
                dist = Abs(catalogPos.Item(CStr(mappedCP)).z - dTopInside)
                strMeasure = "DimPt" & mappedCP & "ToTopInside"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' --------------------------------
                ' Add measurement to bottom inside
                ' --------------------------------
                dist = Abs(catalogPos.Item(CStr(mappedCP)).z - dBtmInside)
                strMeasure = "DimPt" & mappedCP & "ToBottomInside"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' -------------------------
                ' Add measurement for depth
                ' -------------------------
                strMeasure = "DepthAtPt" & mappedCP
                AddMeasurementDimInfo oMeasurements, dBoundedDepth, strMeasure, sCacheString
                
                ' --------------------------------
                ' Add measurement for inside depth
                ' --------------------------------
                strMeasure = "InsideDepthAtPt" & mappedCP
                AddMeasurementDimInfo oMeasurements, dInsideDepth, strMeasure, sCacheString
                
            Next i
        ' --------------------
        ' If flange-penetrated
        ' --------------------
        Else
            ' -----------------------
            ' Get some key dimensions
            ' -----------------------
            If bUseTopFlange Then
                If bTFR Then
                    If bTFL Then
                        dToWL = dBoundedWidth / 2# + dBoundedWebThickness / 2#
                        dToWR = dBoundedWidth / 2# - dBoundedWebThickness / 2#
                    Else
                        dToWL = dBoundedWidth
                        dToWR = dBoundedWidth - dBoundedWebThickness
                    End If
                Else
                'This code is implemented to make HSSR cross-sections work. For HSSR, distance from bounding top to bounded WebRight is returned correctly now.
                    If HasBottomFlange(oBoundedData.MemberPart) Then
                        Dim oModelBody As IJDModelBody
                        Set oModelBody = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT).Geometry
                        
                        Set oRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
                        oModelBody.GetMinimumDistance oRightPort.Geometry, Nothing, Nothing, dToWL
                    Else
                        dToWL = dBoundedWebThickness
                    End If
                    dToWR = 0#

                    dToWR = 0#
                End If
            Else
                If bBFR Then
                    If bBFL Then
                        dToWR = dBoundedWidth / 2# - dBoundedWebThickness / 2#
                        dToWL = dBoundedWidth / 2# + dBoundedWebThickness / 2#
                    Else
                        dToWR = dBoundedWidth - dBoundedWebThickness
                        dToWL = dBoundedWidth
                    End If
                Else
                    If HasBottomFlange(oBoundedData.MemberPart) Then

                        Set oModelBody = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT).Geometry
                        
                        Set oRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
                        oModelBody.GetMinimumDistance oRightPort.Geometry, Nothing, Nothing, dToWL
                    Else
                        dToWL = dBoundedWebThickness
                    End If
                    dToWR = 0#
                End If
            End If
                
            ' -----------------------
            ' Loop through the points
            ' -----------------------
            For i = 1 To oCardinalPointCol.Count
                
                mappedCP = oCardinalPointCol.Item(i)
            
                ' ------------------------------
                ' Add measurement to flange left
                ' ------------------------------
                dist = Abs(catalogPos.Item(CStr(mappedCP)).z - dBoundedWidth)
                strMeasure = "DimPt" & mappedCP & "ToFL"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' -------------------------------
                ' Add measurement to flange right
                ' -------------------------------
                dist = Abs(catalogPos.Item(CStr(mappedCP)).z)
                strMeasure = "DimPt" & mappedCP & "ToFR"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' ---------------------------
                ' Add measurement to web left
                ' ---------------------------
                dist = Abs(catalogPos.Item(CStr(mappedCP)).z - dToWL)
                strMeasure = "DimPt" & mappedCP & "ToWL"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' ----------------------------
                ' Add measurement to web right
                ' ----------------------------
                dist = Abs(catalogPos.Item(CStr(mappedCP)).z - dToWR)
                strMeasure = "DimPt" & mappedCP & "ToWR"
                AddMeasurementDimInfo oMeasurements, dist, strMeasure, sCacheString
                
                ' -------------------------
                ' Add measurement for width
                ' -------------------------
                strMeasure = "WidthAtPt" & mappedCP
                AddMeasurementDimInfo oMeasurements, dBoundedWidth, strMeasure, sCacheString
                
                ' --------------------------------
                ' Add measurement to web thickness
                ' --------------------------------
                strMeasure = "WebThkAtPt" & mappedCP
                AddMeasurementDimInfo oMeasurements, dBoundedWebThickness, strMeasure, sCacheString
                
            Next i
                
        End If
    End If
    
    ' -----------------------------
    ' Add point to point dimensions
    ' -----------------------------
    ' Point 3 is to either point 23 or 24
    If KeyExists(CStr(24), catalogPos) Then
        Set dVector = catalogPos.Item(CStr(3)).Subtract(catalogPos.Item(CStr(24)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt3ToPt24", sCacheString
    Else
        Set dVector = catalogPos.Item(CStr(3)).Subtract(catalogPos.Item(CStr(23)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt3ToPt23", sCacheString
    End If
    
    ' Point 11 is to either point 14 or 15
    If KeyExists(CStr(14), catalogPos) Then
        Set dVector = catalogPos.Item(CStr(11)).Subtract(catalogPos.Item(CStr(14)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt11ToPt14", sCacheString
    Else
        Set dVector = catalogPos.Item(CStr(11)).Subtract(catalogPos.Item(CStr(15)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt11ToPt15", sCacheString
    End If
    
    ' If 14 exists, 50 exists
    If KeyExists(CStr(14), catalogPos) Then
        Set dVector = catalogPos.Item(CStr(14)).Subtract(catalogPos.Item(CStr(50)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt14ToPt50", sCacheString
        
        Set dVector = catalogPos.Item(CStr(15)).Subtract(catalogPos.Item(CStr(50)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt15ToPt50", sCacheString
    End If
    
    ' Point 15 is to either point 17, 20, or 23
    ' If 17 exists, 18 exists
    If KeyExists(CStr(17), catalogPos) Then
        Set dVector = catalogPos.Item(CStr(15)).Subtract(catalogPos.Item(CStr(17)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt15ToPt17", sCacheString
        
        Set dVector = catalogPos.Item(CStr(17)).Subtract(catalogPos.Item(CStr(18)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt17ToPt18", sCacheString
    ElseIf KeyExists(CStr(20), catalogPos) Then
        Set dVector = catalogPos.Item(CStr(15)).Subtract(catalogPos.Item(CStr(20)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt15ToPt20", sCacheString
    Else
        Set dVector = catalogPos.Item(CStr(15)).Subtract(catalogPos.Item(CStr(23)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt15ToPt23", sCacheString
    End If
          
    ' If 20 exists, 21 exists
    If KeyExists(CStr(20), catalogPos) Then
        Set dVector = catalogPos.Item(CStr(20)).Subtract(catalogPos.Item(CStr(21)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt20ToPt21", sCacheString
    End If
    
    ' Point 18 is to either point 20 or 23
    If KeyExists(CStr(18), catalogPos) Then
        If KeyExists(CStr(20), catalogPos) Then
            Set dVector = catalogPos.Item(CStr(18)).Subtract(catalogPos.Item(CStr(20)))
            AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt18ToPt20", sCacheString
        Else
            Set dVector = catalogPos.Item(CStr(18)).Subtract(catalogPos.Item(CStr(23)))
            AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt18ToPt23", sCacheString
        End If
    End If
    
    ' If 21 exists, it is to 23
    If KeyExists(CStr(21), catalogPos) Then
        Set dVector = catalogPos.Item(CStr(21)).Subtract(catalogPos.Item(CStr(23)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt21ToPt23", sCacheString
    End If
    
    ' If 51 exists, we measure from 23 and 24
    If KeyExists(CStr(51), catalogPos) Then
        Set dVector = catalogPos.Item(CStr(51)).Subtract(catalogPos.Item(CStr(23)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt23ToPt51", sCacheString
    
        Set dVector = catalogPos.Item(CStr(51)).Subtract(catalogPos.Item(CStr(24)))
        AddMeasurementDimInfo oMeasurements, dVector.Length, "DimPt24ToPt51", sCacheString
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetMeasurements").Number
    
End Sub

' Compare measurements for debugging purposes
Private Sub TestMeasurements(oCodeMeasure As Collection, oSymMeasure As Collection, bIsWebPenetrated As Boolean)

    Dim oKeys As New Collection
        
    If bIsWebPenetrated Then
        oKeys.Add "DepthAtPt11"
        oKeys.Add "DepthAtPt14"
        oKeys.Add "DepthAtPt15"
        oKeys.Add "DepthAtPt17"
        oKeys.Add "DepthAtPt18"
        oKeys.Add "DepthAtPt20"
        oKeys.Add "DepthAtPt21"
        oKeys.Add "DepthAtPt23"
        oKeys.Add "DepthAtPt24"
        oKeys.Add "DepthAtPt3"
        oKeys.Add "DepthAtPt50"
        oKeys.Add "DepthAtPt51"
        oKeys.Add "DimPt11ToBottom"
        oKeys.Add "DimPt11ToBottomInside"
        oKeys.Add "DimPt11ToPt14"
        oKeys.Add "DimPt11ToPt15"
        oKeys.Add "DimPt11ToTop"
        oKeys.Add "DimPt11ToTopInside"
        oKeys.Add "DimPt14ToBottom"
        oKeys.Add "DimPt14ToBottomInside"
        oKeys.Add "DimPt14ToPt50"
        oKeys.Add "DimPt14ToTop"
        oKeys.Add "DimPt14ToTopInside"
        oKeys.Add "DimPt15ToBottom"
        oKeys.Add "DimPt15ToBottomInside"
        oKeys.Add "DimPt15ToPt17"
        oKeys.Add "DimPt15ToPt20"
        oKeys.Add "DimPt15ToPt50"
        oKeys.Add "DimPt15ToPt23"
        oKeys.Add "DimPt15ToTop"
        oKeys.Add "DimPt15ToTopInside"
        oKeys.Add "DimPt17ToBottom"
        oKeys.Add "DimPt17ToBottomInside"
        oKeys.Add "DimPt17ToPt18"
        oKeys.Add "DimPt17ToTop"
        oKeys.Add "DimPt17ToTopInside"
        oKeys.Add "DimPt18ToBottom"
        oKeys.Add "DimPt18ToBottomInside"
        oKeys.Add "DimPt18ToPt20"
        oKeys.Add "DimPt18ToPt23"
        oKeys.Add "DimPt18ToTop"
        oKeys.Add "DimPt18ToTopInside"
        oKeys.Add "DimPt20ToBottom"
        oKeys.Add "DimPt20ToBottomInside"
        oKeys.Add "DimPt20ToPt21"
        oKeys.Add "DimPt20ToTop"
        oKeys.Add "DimPt20ToTopInside"
        oKeys.Add "DimPt21ToBottom"
        oKeys.Add "DimPt21ToBottomInside"
        oKeys.Add "DimPt21ToPt23"
        oKeys.Add "DimPt21ToTop"
        oKeys.Add "DimPt21ToTopInside"
        oKeys.Add "DimPt23ToBottom"
        oKeys.Add "DimPt23ToBottomInside"
        oKeys.Add "DimPt23ToPt51"
        oKeys.Add "DimPt23ToTop"
        oKeys.Add "DimPt23ToTopInside"
        oKeys.Add "DimPt24ToBottom"
        oKeys.Add "DimPt24ToBottomInside"
        oKeys.Add "DimPt24ToPt51"
        oKeys.Add "DimPt24ToTop"
        oKeys.Add "DimPt24ToTopInside"
        oKeys.Add "DimPt3ToBottom"
        oKeys.Add "DimPt3ToBottomInside"
        oKeys.Add "DimPt3ToPt23"
        oKeys.Add "DimPt3ToPt24"
        oKeys.Add "DimPt3ToTop"
        oKeys.Add "DimPt3ToTopInside"
        oKeys.Add "DimPt50ToBottom"
        oKeys.Add "DimPt50ToBottomInside"
        oKeys.Add "DimPt50ToTop"
        oKeys.Add "DimPt50ToTopInside"
        oKeys.Add "DimPt51ToBottom"
        oKeys.Add "DimPt51ToBottomInside"
        oKeys.Add "DimPt51ToTop"
        oKeys.Add "DimPt51ToTopInside"
        oKeys.Add "InsideDepthAtPt11"
        oKeys.Add "InsideDepthAtPt14"
        oKeys.Add "InsideDepthAtPt15"
        oKeys.Add "InsideDepthAtPt17"
        oKeys.Add "InsideDepthAtPt18"
        oKeys.Add "InsideDepthAtPt20"
        oKeys.Add "InsideDepthAtPt21"
        oKeys.Add "InsideDepthAtPt23"
        oKeys.Add "InsideDepthAtPt24"
        oKeys.Add "InsideDepthAtPt3"
        oKeys.Add "InsideDepthAtPt50"
        oKeys.Add "InsideDepthAtPt51"
    Else
        oKeys.Add "WidthAtPt11"
        oKeys.Add "WidthAtPt14"
        oKeys.Add "WidthAtPt15"
        oKeys.Add "WidthAtPt17"
        oKeys.Add "WidthAtPt18"
        oKeys.Add "WidthAtPt20"
        oKeys.Add "WidthAtPt21"
        oKeys.Add "WidthAtPt23"
        oKeys.Add "WidthAtPt24"
        oKeys.Add "WidthAtPt3"
        oKeys.Add "WidthAtPt50"
        oKeys.Add "WidthAtPt51"
        oKeys.Add "DimPt11ToFR"
        oKeys.Add "DimPt11ToWR"
        oKeys.Add "DimPt11ToPt14"
        oKeys.Add "DimPt11ToPt15"
        oKeys.Add "DimPt11ToFL"
        oKeys.Add "DimPt11ToWL"
        oKeys.Add "DimPt14ToFR"
        oKeys.Add "DimPt14ToWR"
        oKeys.Add "DimPt14ToPt50"
        oKeys.Add "DimPt14ToFL"
        oKeys.Add "DimPt14ToWL"
        oKeys.Add "DimPt15ToFR"
        oKeys.Add "DimPt15ToWR"
        oKeys.Add "DimPt15ToPt17"
        oKeys.Add "DimPt15ToPt20"
        oKeys.Add "DimPt15ToPt50"
        oKeys.Add "DimPt15ToPt23"
        oKeys.Add "DimPt15ToFL"
        oKeys.Add "DimPt15ToWL"
        oKeys.Add "DimPt17ToFR"
        oKeys.Add "DimPt17ToWR"
        oKeys.Add "DimPt17ToPt18"
        oKeys.Add "DimPt17ToFL"
        oKeys.Add "DimPt17ToWL"
        oKeys.Add "DimPt18ToFR"
        oKeys.Add "DimPt18ToWR"
        oKeys.Add "DimPt18ToPt20"
        oKeys.Add "DimPt18ToPt23"
        oKeys.Add "DimPt18ToFL"
        oKeys.Add "DimPt18ToWL"
        oKeys.Add "DimPt20ToFR"
        oKeys.Add "DimPt20ToWR"
        oKeys.Add "DimPt20ToPt21"
        oKeys.Add "DimPt20ToFL"
        oKeys.Add "DimPt20ToWL"
        oKeys.Add "DimPt21ToFR"
        oKeys.Add "DimPt21ToWR"
        oKeys.Add "DimPt21ToPt23"
        oKeys.Add "DimPt21ToFL"
        oKeys.Add "DimPt21ToWL"
        oKeys.Add "DimPt23ToFR"
        oKeys.Add "DimPt23ToWR"
        oKeys.Add "DimPt23ToPt51"
        oKeys.Add "DimPt23ToFL"
        oKeys.Add "DimPt23ToWL"
        oKeys.Add "DimPt24ToFR"
        oKeys.Add "DimPt24ToWR"
        oKeys.Add "DimPt24ToPt51"
        oKeys.Add "DimPt24ToFL"
        oKeys.Add "DimPt24ToWL"
        oKeys.Add "DimPt3ToFR"
        oKeys.Add "DimPt3ToWR"
        oKeys.Add "DimPt3ToPt23"
        oKeys.Add "DimPt3ToPt24"
        oKeys.Add "DimPt3ToFL"
        oKeys.Add "DimPt3ToWL"
        oKeys.Add "DimPt50ToFR"
        oKeys.Add "DimPt50ToWR"
        oKeys.Add "DimPt50ToFL"
        oKeys.Add "DimPt50ToWL"
        oKeys.Add "DimPt51ToFR"
        oKeys.Add "DimPt51ToWR"
        oKeys.Add "DimPt51ToFL"
        oKeys.Add "DimPt51ToWL"
        oKeys.Add "WebThkAtPt11"
        oKeys.Add "WebThkAtPt14"
        oKeys.Add "WebThkAtPt15"
        oKeys.Add "WebThkAtPt17"
        oKeys.Add "WebThkAtPt18"
        oKeys.Add "WebThkAtPt20"
        oKeys.Add "WebThkAtPt21"
        oKeys.Add "WebThkAtPt23"
        oKeys.Add "WebThkAtPt24"
        oKeys.Add "WebThkAtPt3"
        oKeys.Add "WebThkAtPt50"
        oKeys.Add "WebThkAtPt51"
    End If
            
    Dim i As Long
    Dim nKeysFound As Long
    Dim diff As Double
    Dim keepChecking As VbMsgBoxResult
    keepChecking = vbYes
    
    Dim allOK As Boolean
    allOK = True
    
    For i = 1 To oKeys.Count
        If KeyExists(oKeys.Item(i), oCodeMeasure) And Not KeyExists(oKeys.Item(i), oSymMeasure) Then
            keepChecking = MsgBox(oKeys.Item(i) & " not found in symbol measurement", vbYesNo)
            nKeysFound = nKeysFound + 1
            allOK = False
        
        ElseIf KeyExists(oKeys.Item(i), oSymMeasure) And Not KeyExists(oKeys.Item(i), oCodeMeasure) Then
            keepChecking = MsgBox(oKeys.Item(i) & " not found in code measurement", vbYesNo)
            nKeysFound = nKeysFound + 1
            allOK = False
        
        ElseIf KeyExists(oKeys.Item(i), oSymMeasure) And KeyExists(oKeys.Item(i), oCodeMeasure) Then
            diff = Abs(oSymMeasure.Item(oKeys.Item(i)) - oCodeMeasure.Item(oKeys.Item(i)))
            If diff > 0.000011 Then
                keepChecking = MsgBox(oKeys.Item(i) & " does not match (sym|code|diff): " & oSymMeasure.Item(oKeys.Item(i)) & " | " & oCodeMeasure.Item(oKeys.Item(i)) & " | " & diff, vbYesNo)
                allOK = False
            End If
            nKeysFound = nKeysFound + 1
        End If
        
        If keepChecking = vbNo Then
            Exit For
        End If
    Next i
    
    If Not nKeysFound = oSymMeasure.Count And Not nKeysFound = oCodeMeasure.Count And keepChecking = vbYes Then
        MsgBox "Unexpected or missing keys. " & nKeysFound & " were found.  The code measurment defines " & oCodeMeasure.Count & " and the symbol version defines " & oSymMeasure.Count
        allOK = False
    End If
    
    If allOK Then
        MsgBox "All OK"
    End If
    
End Sub

' Get the location of a mapped point on a member or stiffener
' 1) determining the mapped faces of interest based on given key point and section alias
' 2) determining the true faces
' 3) getting the wire edge between the faces
' 4) intersecting the wire with the given plane
Private Function GetMappedPointLocationByIntersection(oBoundingPart As Object, _
                                                      oSketchPlane As IJPlane, _
                                                      mappedPoint As Long, _
                                                      sectionAlias As Long, _
                                                      edgeMap As Collection) As IJDPosition

    ' -------------------------------------------------
    ' Get the idealized face ports adjacent to the edge
    ' -------------------------------------------------
    ' By idealized, we mean as if this were in a member endcut symbol which uses the edge mapping rule
    ' The point passed in is also one from the perspective of an endcut symbol
    Dim eBoundingAlias As eBounding_Alias
    eBoundingAlias = GetBoundingAliasSimplified(sectionAlias)
    
    Dim faceID1 As JXSEC_CODE
    Dim faceID2 As JXSEC_CODE
    
    Select Case mappedPoint
    
        Case 3
        
            Select Case sectionAlias
            
                Case 0, 1, 2, 3, 4, 5, 6, 8, 9, 15, 17
                    faceID1 = JXSEC_BOTTOM
                    faceID2 = JXSEC_WEB_LEFT
                Case 7, 10, 12, 13, 19
                    faceID1 = JXSEC_BOTTOM
                    faceID2 = JXSEC_BOTTOM_FLANGE_LEFT
                Case 11
                    faceID1 = JXSEC_BOTTOM
                    faceID2 = JXSEC_WEB_LEFT_BOTTOM
                Case 14, 16, 18
                    faceID1 = JXSEC_INNER_WEB_RIGHT_BOTTOM
                    faceID2 = JXSEC_BOTTOM_FLANGE_RIGHT
                Case Else
                    GoTo ErrorHandler
            End Select
            
        Case 11
        
            Select Case sectionAlias
            
                Case 0, 1, 2, 3, 4, 5, 7, 10, 11, 14, 17
                    faceID1 = JXSEC_WEB_LEFT
                    faceID2 = JXSEC_TOP
                Case 6, 8, 12, 13, 19
                    faceID1 = JXSEC_TOP_FLANGE_LEFT
                    faceID2 = JXSEC_TOP
                Case 9
                    faceID1 = JXSEC_WEB_LEFT_TOP
                    faceID2 = JXSEC_TOP
                Case 15, 16, 18
                    faceID1 = JXSEC_RIGHT_WEB_TOP
                    faceID2 = JXSEC_INNER_WEB_RIGHT_TOP
                Case Else
                    GoTo ErrorHandler
            End Select
            
        Case 14
        
            Select Case sectionAlias
            
                Case 2
                    faceID1 = JXSEC_TOP
                    faceID2 = JXSEC_WEB_RIGHT_TOP
                Case Else
                    GoTo ErrorHandler
            End Select
        
        Case 15
        
            Select Case sectionAlias
            
                Case 0, 3, 4, 7, 8, 9, 10, 11, 12, 14, 17
                    faceID1 = JXSEC_TOP
                    faceID2 = JXSEC_WEB_RIGHT
                Case 1, 5, 6, 13, 19
                    faceID1 = JXSEC_TOP
                    faceID2 = JXSEC_TOP_FLANGE_RIGHT
                Case 2
                    faceID1 = JXSEC_TOP_FLANGE_RIGHT_TOP
                    faceID2 = JXSEC_TOP_FLANGE_RIGHT
                Case 15, 16, 18
                    faceID1 = JXSEC_RIGHT_WEB_TOP
                    faceID2 = JXSEC_WEB_RIGHT
                Case Else
                    GoTo ErrorHandler
            End Select
            
        Case 17
        
            Select Case sectionAlias
            
                Case 1, 2, 5, 6, 13, 19
                    faceID1 = JXSEC_TOP_FLANGE_RIGHT
                    faceID2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
                Case Else
                    GoTo ErrorHandler
            End Select
            
        Case 18
        
            Select Case sectionAlias
            
                Case 1, 2, 5, 6, 13, 19
                    faceID1 = JXSEC_WEB_RIGHT
                    faceID2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
                Case Else
                    GoTo ErrorHandler
            End Select
            
        Case 20
        
            Select Case sectionAlias
            
                Case 3, 4, 5, 7, 13, 19
                    faceID1 = JXSEC_WEB_RIGHT
                    faceID2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
                Case Else
                    GoTo ErrorHandler
            End Select
        
        Case 21
        
            Select Case sectionAlias
            
                Case 3, 4, 5, 7, 13, 19
                    faceID1 = JXSEC_BOTTOM_FLANGE_RIGHT
                    faceID2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
                Case Else
                    GoTo ErrorHandler
            End Select
        
        Case 23
        
            Select Case sectionAlias
            
                Case 0, 1, 2, 6, 8, 9, 10, 11, 12, 15, 17
                    faceID1 = JXSEC_BOTTOM
                    faceID2 = JXSEC_WEB_RIGHT
                Case 3, 5, 7, 13, 19
                    faceID1 = JXSEC_BOTTOM
                    faceID2 = JXSEC_BOTTOM_FLANGE_RIGHT
                Case 4
                    faceID1 = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM
                    faceID2 = JXSEC_BOTTOM_FLANGE_RIGHT
                Case 14, 16, 18
                    faceID1 = JXSEC_RIGHT_WEB_BOTTOM
                    faceID2 = JXSEC_WEB_RIGHT
                Case Else
                    GoTo ErrorHandler
            End Select
        
        Case 24
        
            Select Case sectionAlias
            
                Case 4
                    faceID1 = JXSEC_BOTTOM
                    faceID2 = JXSEC_WEB_RIGHT_BOTTOM
                Case Else
                    GoTo ErrorHandler
            End Select
    
        Case 50
        
            Select Case sectionAlias
            
                Case 2
                    faceID1 = JXSEC_TOP_FLANGE_RIGHT_TOP
                    faceID2 = JXSEC_WEB_RIGHT_TOP
                Case Else
                    GoTo ErrorHandler
            End Select
            
        Case 51
        
            Select Case sectionAlias
            
                Case 4
                    faceID1 = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM
                    faceID2 = JXSEC_WEB_RIGHT_BOTTOM
                Case Else
                    GoTo ErrorHandler
            End Select
            
    End Select
    
    Dim trueFaceID1 As Long
    Dim truefaceID2 As Long
    
    trueFaceID1 = edgeMap(CStr(faceID1))
    truefaceID2 = edgeMap(CStr(faceID2))
        
    Dim oFace1Port As IJPort
    Dim oFace2Port As IJPort
    
    Set oFace1Port = GetLateralSubPortBeforeTrim(oBoundingPart, trueFaceID1)
    Set oFace2Port = GetLateralSubPortBeforeTrim(oBoundingPart, truefaceID2)
    
    ' ---------------------------------------------------------
    ' Get a wire where each port intersects the sketching plane
    ' ---------------------------------------------------------
    ' Because of fillets, an actual edge may not exist.  We could use a VB-friendly version of GetEdgesByFaceAttributes
    ' as a quick check.  We can try intersecting the edges, but that will only succeed roughly half of the time, and we'll then
    ' have to do the fallback that requires two intersections (3 total).
    ' 50% 1 intersection and 50% 3 intersections = average of 2
    ' We may as well go straight for the fall-back solution.
    
    Dim oGeomOps As New IMSModelGeomOps.DGeomOpsIntersect
        
    Dim oFace1Wire As Object
    Dim oFace2Wire As Object
    
    oGeomOps.PlaceIntersectionObject Nothing, oFace1Port.Geometry, oSketchPlane, Nothing, oFace1Wire
    oGeomOps.PlaceIntersectionObject Nothing, oFace2Port.Geometry, oSketchPlane, Nothing, oFace2Wire
    
    ' --------------------------------------------------------
    ' Find the closest point of approach between the two wires
    ' --------------------------------------------------------
    Dim oModelUtil As IJSGOModelBodyUtilities
    Set oModelUtil = New SGOModelBodyUtilities
    
    Dim dist As Double
    Dim oPointOnWire1 As IJDPosition
    Dim oPointOnWire2 As IJDPosition
    oModelUtil.GetClosestPointsBetweenTwoBodies oFace1Wire, oFace2Wire, oPointOnWire1, oPointOnWire2, dist
    
    ' -----------------------------------------------------------------------------
    ' Create infinite lines through these points using the tangents at these points
    ' -----------------------------------------------------------------------------
    Dim oWireUtil As IJSGOWireBodyUtilities
    Set oWireUtil = New SGOWireBodyUtilities
    
    Dim oInfLine1 As IJLine
    Dim oInfLine2 As IJLine
    Set oInfLine1 = New Line3d
    Set oInfLine2 = New Line3d
    
    Dim oTangent As IJDVector
    Dim oDummyPoint As IJDPosition
    
    oWireUtil.GetClosestPointOnWire oFace1Wire, oPointOnWire1, oDummyPoint, oTangent
    
    oInfLine1.SetRootPoint oPointOnWire1.x, oPointOnWire1.y, oPointOnWire1.z
    oInfLine1.SetDirection oTangent.x, oTangent.y, oTangent.z
     
    Set oDummyPoint = Nothing
    Set oTangent = Nothing
    
    oWireUtil.GetClosestPointOnWire oFace2Wire, oPointOnWire2, oDummyPoint, oTangent
     
    oInfLine2.SetRootPoint oPointOnWire2.x, oPointOnWire2.y, oPointOnWire2.z
    oInfLine2.SetDirection oTangent.x, oTangent.y, oTangent.z
    
    oInfLine1.Infinite = True
    oInfLine2.Infinite = True
    
    ' ---------------------------------------------------------------------
    ' Set the result to the closest point of approach between the two lines
    ' ---------------------------------------------------------------------
    ' Since they are on the same plane, they should intersect exactly
    ' We can use the location on either line
    Dim oCurve1 As IJCurve
    Dim oCurve2 As IJCurve
    Set oCurve1 = oInfLine1
    Set oCurve2 = oInfLine2
    
    Dim x1 As Double
    Dim y1 As Double
    Dim z1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim z2 As Double
    
    oCurve1.DistanceBetween oCurve2, dist, x1, y1, z1, x2, y2, z2
    
    Set GetMappedPointLocationByIntersection = New DPosition
    
    GetMappedPointLocationByIntersection.Set x1, y1, z1
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetMappedPointLocationByIntersection").Number
    
End Function

' Determines the x-y location of a stiffener load point (see Symbol 2D help) on a member cross section.
' Stiffener load points are 1-27.  Member cardinal points are 1-15.

Private Sub GetStiffenerLoadPointPositionOnMemberSection(oMember As ISPSMemberPartCommon, cp As Long, x As Double, y As Double)

    On Error GoTo ErrorHandler
    
    ' --------------------------
    ' Get the section properties
    ' --------------------------
    Dim oSDOMember As New StructDetailObjects.MemberPart
    Set oSDOMember.object = oMember
    
    Dim flangeThickness As Double
    Dim webThickness As Double
    Dim depth As Double
    Dim sectionType As String
    
    flangeThickness = oSDOMember.flangeThickness
    webThickness = oSDOMember.webThickness
    depth = oSDOMember.WebLength
    sectionType = oSDOMember.sectionType

    ' --------------------
    ' Get the section type
    ' --------------------
    Dim oMemberSection As ISPSCrossSection
    Set oMemberSection = oMember.CrossSection

        
    If sectionType = "W" Or sectionType = "M" Or sectionType = "S" Or sectionType = "HP" Then
        Select Case cp
            Case 1
                oMemberSection.GetCardinalPointOffset 2, x, y
            Case 2
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x - webThickness / 2#
            Case 3
                oMemberSection.GetCardinalPointOffset 1, x, y
            Case 4
                oMemberSection.GetCardinalPointOffset 1, x, y
                If Not sectionType = "S" Then
                    y = y + flangeThickness / 2#
                End If
            Case 5
                oMemberSection.GetCardinalPointOffset 1, x, y
                If Not sectionType = "S" Then
                    y = y + flangeThickness
                End If
            Case 6
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x - webThickness / 2#
                y = y + flangeThickness
            Case 7
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x - webThickness / 2#
                y = y + depth / 2#
            Case 8
                oMemberSection.GetCardinalPointOffset 8, x, y
                x = x - webThickness / 2#
                y = y - flangeThickness
            Case 9
                oMemberSection.GetCardinalPointOffset 7, x, y
                If Not sectionType = "S" Then
                    y = y - flangeThickness
                End If
            Case 10
                oMemberSection.GetCardinalPointOffset 7, x, y
                If Not sectionType = "S" Then
                    y = y - flangeThickness / 2#
                End If
            Case 11
                oMemberSection.GetCardinalPointOffset 7, x, y
            Case 12
                oMemberSection.GetCardinalPointOffset 8, x, y
                x = x - webThickness / 2#
            Case 13
                oMemberSection.GetCardinalPointOffset 8, x, y
            Case 14
                oMemberSection.GetCardinalPointOffset 8, x, y
                x = x + webThickness / 2#
            Case 15
                oMemberSection.GetCardinalPointOffset 9, x, y
            Case 16
                oMemberSection.GetCardinalPointOffset 9, x, y
                If Not sectionType = "S" Then
                    y = y - flangeThickness / 2#
                End If
            Case 17
                oMemberSection.GetCardinalPointOffset 9, x, y
                If Not sectionType = "S" Then
                    y = y - flangeThickness
                End If
            Case 18
                oMemberSection.GetCardinalPointOffset 8, x, y
                x = x + webThickness / 2#
                y = y - flangeThickness
            Case 19
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x + webThickness / 2#
                y = y + depth / 2#
            Case 20
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x + webThickness / 2#
                y = y + flangeThickness
            Case 21
                oMemberSection.GetCardinalPointOffset 3, x, y
                If Not sectionType = "S" Then
                    y = y + flangeThickness
                End If
            Case 22
                oMemberSection.GetCardinalPointOffset 3, x, y
                If Not sectionType = "S" Then
                    y = y + flangeThickness / 2#
                End If
            Case 23
                oMemberSection.GetCardinalPointOffset 3, x, y
            Case 24
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x + webThickness / 2#
            Case 25
                oMemberSection.GetCardinalPointOffset 4, x, y
            Case 26
                oMemberSection.GetCardinalPointOffset 6, x, y
            Case 27
                oMemberSection.GetCardinalPointOffset 5, x, y
        End Select
        
    ElseIf sectionType = "WT" Or sectionType = "MT" Or sectionType = "ST" Then
        Select Case cp
            Case 1
                oMemberSection.GetCardinalPointOffset 2, x, y
            Case 2, 3, 4, 5, 6
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x - webThickness / 2#
            Case 7
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x - webThickness / 2#
                y = y + depth / 2#
            Case 8
                oMemberSection.GetCardinalPointOffset 8, x, y
                x = x - webThickness / 2#
                y = y - flangeThickness
            Case 9
                oMemberSection.GetCardinalPointOffset 7, x, y
                If Not sectionType = "ST" Then
                    y = y - flangeThickness
                End If
            Case 10
                oMemberSection.GetCardinalPointOffset 7, x, y
                If Not sectionType = "ST" Then
                    y = y - flangeThickness / 2#
                End If
            Case 11
                oMemberSection.GetCardinalPointOffset 7, x, y
            Case 12
                oMemberSection.GetCardinalPointOffset 8, x, y
                x = x - webThickness / 2#
            Case 13
                oMemberSection.GetCardinalPointOffset 8, x, y
            Case 14
                oMemberSection.GetCardinalPointOffset 8, x, y
                x = x + webThickness / 2#
            Case 15
                oMemberSection.GetCardinalPointOffset 9, x, y
            Case 16
                oMemberSection.GetCardinalPointOffset 9, x, y
                If Not sectionType = "S" Then
                    y = y - flangeThickness / 2#
                End If
            Case 17
                oMemberSection.GetCardinalPointOffset 9, x, y
                If Not sectionType = "S" Then
                    y = y - flangeThickness
                End If
            Case 18
                oMemberSection.GetCardinalPointOffset 8, x, y
                x = x + webThickness / 2#
                y = y - flangeThickness
            Case 19
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x + webThickness / 2#
                y = y + depth / 2#
            Case 20, 21, 22, 23, 24
                oMemberSection.GetCardinalPointOffset 2, x, y
                x = x + webThickness / 2#
            Case 25
                oMemberSection.GetCardinalPointOffset 4, x, y
            Case 26
                oMemberSection.GetCardinalPointOffset 6, x, y
            Case 27
                oMemberSection.GetCardinalPointOffset 5, x, y
        End Select
    ElseIf sectionType = "C" Or sectionType = "MC" Then
        Select Case cp
            Case 1
                oMemberSection.GetCardinalPointOffset 1, x, y
                x = x + webThickness / 2#
            Case 2, 3, 4
                oMemberSection.GetCardinalPointOffset 1, x, y
            Case 5, 6
                x = x + flangeThickness
            Case 7, 25
                oMemberSection.GetCardinalPointOffset 4, x, y
            Case 8, 9
                oMemberSection.GetCardinalPointOffset 7, x, y
                x = x - flangeThickness
            Case 10, 11, 12
                oMemberSection.GetCardinalPointOffset 7, x, y
            Case 13
                oMemberSection.GetCardinalPointOffset 7, x, y
                x = x + webThickness / 2#
            Case 14
                oMemberSection.GetCardinalPointOffset 7, x, y
                x = x + webThickness
            Case 15
                oMemberSection.GetCardinalPointOffset 9, x, y
            Case 16
                oMemberSection.GetCardinalPointOffset 9, x, y
                If Not sectionType = "No S-Channel Currently Defined in Catalog" Then
                    y = y - flangeThickness / 2#
                End If
            Case 17
                oMemberSection.GetCardinalPointOffset 9, x, y
                If Not sectionType = "No S-Channel Currently Defined in Catalog" Then
                    y = y - flangeThickness
                End If
            Case 18
                oMemberSection.GetCardinalPointOffset 7, x, y
                x = x + webThickness
                y = y - flangeThickness
            Case 19
                oMemberSection.GetCardinalPointOffset 7, x, y
                x = x + webThickness
            Case 20
                oMemberSection.GetCardinalPointOffset 1, x, y
                x = x + webThickness
                y = y + flangeThickness
            Case 21
                oMemberSection.GetCardinalPointOffset 3, x, y
                If Not sectionType = "No S-Channel Currently Defined in Catalog" Then
                    y = y + flangeThickness
                End If
            Case 22
                oMemberSection.GetCardinalPointOffset 3, x, y
                If Not sectionType = "No S-Channel Currently Defined in Catalog" Then
                    y = y + flangeThickness / 2#
                End If
            Case 23
                oMemberSection.GetCardinalPointOffset 3, x, y
            Case 24
                oMemberSection.GetCardinalPointOffset 1, x, y
                x = x + webThickness
            Case 26
                oMemberSection.GetCardinalPointOffset 6, x, y
            Case 27
                oMemberSection.GetCardinalPointOffset 4, x, y
                x = x + webThickness / 2#
        End Select
        
    ElseIf sectionType = "L" Then
        Select Case cp
            Case 1
                oMemberSection.GetCardinalPointOffset 1, x, y
                x = x + webThickness / 2#
            Case 2, 3, 4
                oMemberSection.GetCardinalPointOffset 1, x, y
            Case 5, 6
                x = x + flangeThickness
            Case 7, 25
                oMemberSection.GetCardinalPointOffset 4, x, y
            Case 8, 9, 10, 11, 12
                oMemberSection.GetCardinalPointOffset 7, x, y
            Case 13
                oMemberSection.GetCardinalPointOffset 7, x, y
                x = x + webThickness / 2#
            Case 14, 15, 16, 17, 18
                oMemberSection.GetCardinalPointOffset 7, x, y
                x = x + webThickness
            Case 19
                oMemberSection.GetCardinalPointOffset 7, x, y
                x = x + webThickness
            Case 20
                oMemberSection.GetCardinalPointOffset 1, x, y
                x = x + webThickness
                y = y + flangeThickness
            Case 21
                oMemberSection.GetCardinalPointOffset 3, x, y
                y = y + flangeThickness
            Case 22
                oMemberSection.GetCardinalPointOffset 3, x, y
                y = y + flangeThickness / 2#
            Case 23
                oMemberSection.GetCardinalPointOffset 3, x, y
            Case 24
                oMemberSection.GetCardinalPointOffset 1, x, y
                x = x + webThickness
            Case 26
                oMemberSection.GetCardinalPointOffset 6, x, y
            Case 27
                oMemberSection.GetCardinalPointOffset 4, x, y
                x = x + webThickness / 2#
        End Select
        
    ElseIf sectionType = "No FB Currently Defined in Catalog" Then

        Select Case cp
            Case 1
                oMemberSection.GetCardinalPointOffset 2, x, y
            Case 2, 3, 4, 5, 6
                oMemberSection.GetCardinalPointOffset 1, x, y
            Case 7, 25
                oMemberSection.GetCardinalPointOffset 4, x, y
            Case 8, 9, 10, 11, 12
                oMemberSection.GetCardinalPointOffset 7, x, y
            Case 13
                oMemberSection.GetCardinalPointOffset 8, x, y
            Case 14, 15, 16, 17, 18
                oMemberSection.GetCardinalPointOffset 9, x, y
            Case 19, 26
                oMemberSection.GetCardinalPointOffset 6, x, y
            Case 20, 21, 22, 23, 24
                oMemberSection.GetCardinalPointOffset 3, x, y
            Case 27
                oMemberSection.GetCardinalPointOffset 5, x, y
        End Select
            
    ElseIf sectionType = "HSSR" Or sectionType = "RS" Then

        Select Case cp
            Case 1, 2, 6, 20, 24
                oMemberSection.GetCardinalPointOffset 2, x, y
            Case 3, 4, 5
                oMemberSection.GetCardinalPointOffset 1, x, y
            Case 7, 19, 27
                oMemberSection.GetCardinalPointOffset 5, x, y
            Case 8, 12, 13, 14, 18
                oMemberSection.GetCardinalPointOffset 8, x, y
            Case 9, 10, 11
                oMemberSection.GetCardinalPointOffset 7, x, y
            Case 15, 16, 17
                oMemberSection.GetCardinalPointOffset 9, x, y
            Case 26
                oMemberSection.GetCardinalPointOffset 6, x, y
            Case 21, 22, 23
                oMemberSection.GetCardinalPointOffset 3, x, y
            Case 25
                oMemberSection.GetCardinalPointOffset 4, x, y
        End Select
            
    ElseIf sectionType = "HSSC" Or sectionType = "Pipe" Then

        Select Case cp
            Case 1, 2, 6, 20, 21, 22, 23, 24
                oMemberSection.GetCardinalPointOffset 2, x, y
            Case 3, 4, 5, 25
                oMemberSection.GetCardinalPointOffset 4, x, y
            Case 7, 19, 27
                oMemberSection.GetCardinalPointOffset 5, x, y
            Case 8, 9, 10, 11, 12, 13, 14, 18
                oMemberSection.GetCardinalPointOffset 8, x, y
            Case 15, 16, 17, 26
                oMemberSection.GetCardinalPointOffset 6, x, y
        End Select
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetStiffenerLoadPointPositionOnMemberSection").Number

End Sub

' Based on the way the edges are mapped, determine how to map the points at key intersections for measurement
' These points mirror those in the measurement symbols used up to and including R1.  They differ from the official key point description
' in the Symbol 2D help as noted below

Public Function GetPointMappingFromEdgeMapping(oEdgeMap As Collection) As Collection

    On Error GoTo ErrorHandler
    
    Dim oPointMap As New Collection
    
    Dim eTrueT As JXSEC_CODE        ' Top
    Dim eTrueWRT As JXSEC_CODE      ' Web Right Top
    Dim eTrueTFRT As JXSEC_CODE     ' Top Flange Right Top
    Dim eTrueTFR As JXSEC_CODE      ' Top Flange Right
    Dim eTrueTFRB As JXSEC_CODE     ' Top Flange Right Bottom
    Dim eTrueWR As JXSEC_CODE       ' Web Right
    Dim eTrueBFRT As JXSEC_CODE     ' Bottom Flange Right Top
    Dim eTrueBFR As JXSEC_CODE      ' Bottom Flange Right
    Dim eTrueBFRB As JXSEC_CODE     ' Bottom Flange Right Bottom
    Dim eTrueWRB As JXSEC_CODE      ' Web Right Bottom
    Dim eTrueB As JXSEC_CODE        ' Bottom
    Dim eTrueTFLB As JXSEC_CODE     ' Top Flange Left Bottom (for differentiating T and UA)
    Dim eTrueBFLT As JXSEC_CODE     ' Bottom Flange Left Top (for differentiating T and UA)
    Dim eTrueRWT As JXSEC_CODE      ' Right Web Top (compare Web Right Top)
    Dim eTrueRWB As JXSEC_CODE      ' Right Web Bottom (compare Web Right Bottom)
    Dim eTrueWLT As JXSEC_CODE      ' Web Left Top
    Dim etrueWLB As JXSEC_CODE      ' Web Right Top
    Dim eTrueOT As JXSEC_CODE
    
    If KeyExists(CStr(JXSEC_TOP), oEdgeMap) Then
        eTrueT = oEdgeMap.Item(CStr(JXSEC_TOP))
    End If
    
    
    If KeyExists(CStr(JXSEC_WEB_RIGHT_TOP), oEdgeMap) Then
        eTrueWRT = oEdgeMap.Item(CStr(JXSEC_WEB_RIGHT_TOP))
    Else
        eTrueWRT = JXSEC_UNKNOWN
    End If
    
    If KeyExists(CStr(JXSEC_TOP_FLANGE_RIGHT_TOP), oEdgeMap) Then
        eTrueTFRT = oEdgeMap.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_TOP))
    Else
        eTrueTFRT = JXSEC_UNKNOWN
    End If
    
    If KeyExists(CStr(JXSEC_TOP_FLANGE_RIGHT), oEdgeMap) Then
        eTrueTFR = oEdgeMap.Item(CStr(JXSEC_TOP_FLANGE_RIGHT))
        eTrueTFRB = oEdgeMap.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_BOTTOM))
    Else
        eTrueTFR = JXSEC_UNKNOWN
        eTrueTFRB = JXSEC_UNKNOWN
    End If
    
    If KeyExists(CStr(JXSEC_WEB_RIGHT), oEdgeMap) Then
        eTrueWR = oEdgeMap.Item(CStr(JXSEC_WEB_RIGHT))
    End If
    
    If KeyExists(CStr(JXSEC_BOTTOM_FLANGE_RIGHT), oEdgeMap) Then
        eTrueBFR = oEdgeMap.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT))
        eTrueBFRT = oEdgeMap.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_TOP))
    Else
        eTrueBFR = JXSEC_UNKNOWN
        eTrueBFRT = JXSEC_UNKNOWN
    End If

    If KeyExists(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM), oEdgeMap) Then
        eTrueBFRB = oEdgeMap.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM))
    Else
        eTrueBFRB = JXSEC_UNKNOWN
    End If

    If KeyExists(CStr(JXSEC_WEB_RIGHT_BOTTOM), oEdgeMap) Then
        eTrueWRB = oEdgeMap.Item(CStr(JXSEC_WEB_RIGHT_BOTTOM))
    Else
        eTrueWRB = JXSEC_UNKNOWN
    End If

    If KeyExists(CStr(JXSEC_TOP_FLANGE_LEFT_BOTTOM), oEdgeMap) Then
        eTrueTFLB = oEdgeMap.Item(CStr(JXSEC_TOP_FLANGE_LEFT_BOTTOM))
    Else
        eTrueTFLB = JXSEC_UNKNOWN
    End If
    
    If KeyExists(CStr(JXSEC_BOTTOM_FLANGE_LEFT_TOP), oEdgeMap) Then
        eTrueBFLT = oEdgeMap.Item(CStr(JXSEC_BOTTOM_FLANGE_LEFT_TOP))
    Else
        eTrueBFLT = JXSEC_UNKNOWN
    End If

    If KeyExists(CStr(JXSEC_BOTTOM), oEdgeMap) Then
        eTrueB = oEdgeMap.Item(CStr(JXSEC_BOTTOM))
    End If
    
    If KeyExists(CStr(JXSEC_RIGHT_WEB_TOP), oEdgeMap) Then
        eTrueRWT = oEdgeMap.Item(CStr(JXSEC_RIGHT_WEB_TOP))
    Else
        eTrueRWT = JXSEC_UNKNOWN
    End If
    
    If KeyExists(CStr(JXSEC_RIGHT_WEB_BOTTOM), oEdgeMap) Then
        eTrueRWB = oEdgeMap.Item(CStr(JXSEC_RIGHT_WEB_BOTTOM))
    Else
        eTrueRWB = JXSEC_UNKNOWN
    End If
    
    If KeyExists(CStr(JXSEC_WEB_LEFT_TOP), oEdgeMap) Then
        eTrueWLT = oEdgeMap.Item(CStr(JXSEC_WEB_LEFT_TOP))
    Else
        eTrueWLT = JXSEC_UNKNOWN
    End If
    
    If KeyExists(CStr(JXSEC_WEB_LEFT_BOTTOM), oEdgeMap) Then
        etrueWLB = oEdgeMap.Item(CStr(JXSEC_WEB_LEFT_BOTTOM))
    Else
        etrueWLB = JXSEC_UNKNOWN
    End If
    
    If KeyExists(CStr(JXSEC_OUTER_TUBE), oEdgeMap) Then
        eTrueOT = oEdgeMap.Item(CStr(JXSEC_OUTER_TUBE))
    End If
    
    If eTrueWR = JXSEC_WEB_RIGHT Then
        
        If eTrueTFRB = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
            ' 0 deg, no reflection, I, C, T, UA, BUTL3
            oPointMap.Add 3, CStr(3)
            oPointMap.Add 11, CStr(11)
            oPointMap.Add 17, CStr(17)
            oPointMap.Add 18, CStr(18)
                
            If eTrueBFR = JXSEC_BOTTOM_FLANGE_RIGHT Then
                ' I, C
                oPointMap.Add 15, CStr(15)
                oPointMap.Add 20, CStr(20)
                oPointMap.Add 21, CStr(21)
                oPointMap.Add 23, CStr(23)
            ElseIf eTrueWRT = JXSEC_WEB_RIGHT_TOP Then
                ' BUTL3
                oPointMap.Add 14, CStr(14)
                oPointMap.Add 15, CStr(15)
                oPointMap.Add 22, CStr(23) ' point 23 in the measurement symbol is not in the same location as the section symbol
                oPointMap.Add 50, CStr(50)
            ElseIf eTrueTFLB = JXSEC_TOP_FLANGE_LEFT_BOTTOM Then
                ' T
                oPointMap.Add 15, CStr(15)
                oPointMap.Add 23, CStr(23)
            Else
                ' UA
                oPointMap.Add 15, CStr(15)
                oPointMap.Add 22, CStr(23) ' point 23 in the measurement symbol is not in the same location as the section symbol
            End If
        
        ElseIf eTrueBFRT = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
            ' 0 deg, no reflection, L
            oPointMap.Add 3, CStr(3)
            oPointMap.Add 11, CStr(11)
            oPointMap.Add 14, CStr(15) ' Presumes that point 15 on an L would be defined like point 23 on a UA.
                                       ' Use point 14 instead of 15.  See note on point 23 for UA (above).
            oPointMap.Add 20, CStr(20)
            oPointMap.Add 21, CStr(21)
            oPointMap.Add 23, CStr(23)
        
        ElseIf eTrueTFRB = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
            ' 180 deg, reflected, I, C, L
            oPointMap.Add 11, CStr(3)
            oPointMap.Add 3, CStr(11)
            oPointMap.Add 21, CStr(17)
            oPointMap.Add 20, CStr(18)
            oPointMap.Add 23, CStr(15)
            oPointMap.Add 15, CStr(23)
            
            If Not eTrueBFRT = JXSEC_UNKNOWN Then
                oPointMap.Add 18, CStr(20)
                oPointMap.Add 17, CStr(21)
            End If
        
        ElseIf eTrueBFRT = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
            ' 180 deg, reflected, T, UA, BUTL3
            oPointMap.Add 11, CStr(3)
            oPointMap.Add 3, CStr(11)
            oPointMap.Add 22, CStr(15) ' point 23 in UA and BUTL3 measurement symbol is not in the same location as the section symbol
                                       ' Point 22 also works for T, since points 22 and 23 are at the same point
            oPointMap.Add 18, CStr(20)
            oPointMap.Add 17, CStr(21)
            oPointMap.Add 15, CStr(23)
        
            If eTrueWRB = JXSEC_WEB_RIGHT_TOP Then
                ' BUTL3
                oPointMap.Add 50, CStr(51)
                oPointMap.Add 14, CStr(24)
            End If
        
        ElseIf eTrueT = JXSEC_TOP Then
            ' 0 deg, no reflection, FB, RECT
            oPointMap.Add 3, CStr(3)
            oPointMap.Add 11, CStr(11)
            oPointMap.Add 15, CStr(15)
            oPointMap.Add 23, CStr(23)
        
        Else
            ' 180 deg, reflection, FB, RECT
            oPointMap.Add 11, CStr(3)
            oPointMap.Add 3, CStr(11)
            oPointMap.Add 23, CStr(15)
            oPointMap.Add 15, CStr(23)
        End If
    
    ElseIf eTrueWR = JXSEC_BOTTOM Then
    
        If Not eTrueRWT = JXSEC_UNKNOWN And Not eTrueRWB = JXSEC_UNKNOWN Then
            
            If eTrueRWT = JXSEC_BOTTOM_FLANGE_RIGHT Then
                ' 90 deg, no reflect, I
                oPointMap.Add 5, CStr(3)
                oPointMap.Add 21, CStr(11)
                oPointMap.Add 23, CStr(15)
                oPointMap.Add 3, CStr(23)
            Else
                ' 90 deg, reflect, I
                oPointMap.Add 21, CStr(3)
                oPointMap.Add 5, CStr(11)
                oPointMap.Add 3, CStr(15)
                oPointMap.Add 23, CStr(23)
            End If
            
        ElseIf Not eTrueRWT = JXSEC_UNKNOWN Then
            ' 90 deg, no reflect, C
            oPointMap.Add 11, CStr(3)
            oPointMap.Add 21, CStr(11)
            oPointMap.Add 23, CStr(15)
            oPointMap.Add 3, CStr(23)
        
        ElseIf Not eTrueRWB = JXSEC_UNKNOWN Then
            ' 90 deg, reflect, C
            oPointMap.Add 21, CStr(3)
            oPointMap.Add 11, CStr(11)
            oPointMap.Add 3, CStr(15)
            oPointMap.Add 23, CStr(23)
        
        ElseIf eTrueBFRT = JXSEC_UNKNOWN And Not eTrueBFLT = JXSEC_UNKNOWN Then
            ' 90 deg, no reflect, L
            oPointMap.Add 11, CStr(3)
            oPointMap.Add 21, CStr(11)
            oPointMap.Add 23, CStr(15)
            oPointMap.Add 3, CStr(23)
            
        ElseIf eTrueTFRB = JXSEC_UNKNOWN And Not eTrueTFLB = JXSEC_UNKNOWN Then
            ' 90 deg, refect, L
            oPointMap.Add 21, CStr(3)
            oPointMap.Add 11, CStr(11)
            oPointMap.Add 3, CStr(15)
            oPointMap.Add 23, CStr(23)
                                
        ElseIf eTrueT = JXSEC_WEB_RIGHT Then
            ' 90 deg, no reflect, FB, RECT
            oPointMap.Add 11, CStr(3)
            oPointMap.Add 15, CStr(11)
            oPointMap.Add 23, CStr(15)
            oPointMap.Add 3, CStr(23)
        
        Else
            ' 90 deg, reflect, FB, RECT
            oPointMap.Add 15, CStr(3)
            oPointMap.Add 11, CStr(11)
            oPointMap.Add 3, CStr(15)
            oPointMap.Add 23, CStr(23)
            
        End If
    
    ElseIf eTrueWR = JXSEC_WEB_LEFT Then
    
        If Not eTrueTFR = JXSEC_UNKNOWN And Not eTrueBFR = JXSEC_UNKNOWN Then
            
            If eTrueTFR = JXSEC_BOTTOM_FLANGE_LEFT Then
                '180 deg, no reflect, I
                oPointMap.Add 15, CStr(3)
                oPointMap.Add 23, CStr(11)
                oPointMap.Add 3, CStr(15)
                oPointMap.Add 5, CStr(17)
                oPointMap.Add 6, CStr(18)
                oPointMap.Add 8, CStr(20)
                oPointMap.Add 9, CStr(21)
                oPointMap.Add 11, CStr(23)
            
            Else
                '0 deg, reflect, I
                oPointMap.Add 23, CStr(3)
                oPointMap.Add 15, CStr(11)
                oPointMap.Add 11, CStr(15)
                oPointMap.Add 9, CStr(17)
                oPointMap.Add 8, CStr(18)
                oPointMap.Add 6, CStr(20)
                oPointMap.Add 5, CStr(21)
                oPointMap.Add 3, CStr(23)
                            
            End If
        
        ElseIf Not eTrueTFLB = JXSEC_UNKNOWN And Not eTrueBFLT = JXSEC_UNKNOWN Then
        
            If eTrueTFLB = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
                ' 180 deg, no reflect, C
                oPointMap.Add 15, CStr(3)
                oPointMap.Add 23, CStr(11)
                oPointMap.Add 3, CStr(15)
                oPointMap.Add 11, CStr(23)

            Else
                ' 0 deg, reflect, C
                oPointMap.Add 23, CStr(3)
                oPointMap.Add 15, CStr(11)
                oPointMap.Add 11, CStr(15)
                oPointMap.Add 3, CStr(23)
            End If
        
        ElseIf Not eTrueBFLT = JXSEC_UNKNOWN And Not eTrueBFRT = JXSEC_UNKNOWN Then
            ' 180 deg, no reflect, T
            oPointMap.Add 15, CStr(3)
            oPointMap.Add 23, CStr(11)
            oPointMap.Add 3, CStr(15)
            oPointMap.Add 8, CStr(20)
            oPointMap.Add 9, CStr(21)
            oPointMap.Add 11, CStr(23)
        
        ElseIf Not eTrueTFLB = JXSEC_UNKNOWN And Not eTrueTFRB = JXSEC_UNKNOWN Then
            ' 0 deg, reflect, T
            oPointMap.Add 23, CStr(3)
            oPointMap.Add 15, CStr(11)
            oPointMap.Add 11, CStr(15)
            oPointMap.Add 9, CStr(17)
            oPointMap.Add 8, CStr(18)
            oPointMap.Add 3, CStr(23)
                
        ElseIf eTrueTFLB = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
            ' 180 deg, no reflect, L
            oPointMap.Add 15, CStr(3)
            oPointMap.Add 23, CStr(11)
            oPointMap.Add 3, CStr(15)
            oPointMap.Add 11, CStr(23)
            
        ElseIf eTrueBFLT = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
            ' 0 deg, reflect, L
            oPointMap.Add 23, CStr(3)
            oPointMap.Add 14, CStr(11)
            oPointMap.Add 11, CStr(15)
            oPointMap.Add 3, CStr(23)
            
        ElseIf eTrueBFLT = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
            
            If etrueWLB = JXSEC_UNKNOWN Then
                ' 180 deg, no reflect, A
                oPointMap.Add 15, CStr(3)
                oPointMap.Add 22, CStr(11)
                oPointMap.Add 3, CStr(15)
                oPointMap.Add 11, CStr(23)
            Else
                ' 180 deg, no reflect, BUTL3
                oPointMap.Add 14, CStr(3)
                oPointMap.Add 22, CStr(11)
                oPointMap.Add 3, CStr(15)
                oPointMap.Add 11, CStr(23)
            End If
            
        ElseIf eTrueTFLB = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
        
            If eTrueWLT = JXSEC_UNKNOWN Then
                ' 0 deg, reflect, A
                oPointMap.Add 22, CStr(3)
                oPointMap.Add 15, CStr(11)
                oPointMap.Add 11, CStr(15)
                oPointMap.Add 3, CStr(23)
            Else
                ' 0 deg, reflect, BUTL3
                oPointMap.Add 22, CStr(3)
                oPointMap.Add 14, CStr(11)
                oPointMap.Add 11, CStr(15)
                oPointMap.Add 3, CStr(23)
            End If
        
        ElseIf eTrueT = JXSEC_BOTTOM Then
            ' 180 deg, no reflect, FB, RECT
            oPointMap.Add 15, CStr(3)
            oPointMap.Add 23, CStr(11)
            oPointMap.Add 3, CStr(15)
            oPointMap.Add 11, CStr(23)
            
        Else
            ' 0 deg, reflect, FB, RECT
            oPointMap.Add 23, CStr(3)
            oPointMap.Add 15, CStr(11)
            oPointMap.Add 11, CStr(15)
            oPointMap.Add 3, CStr(23)
            
        End If
    
    ElseIf eTrueWR = JXSEC_TOP Then
            
        If Not eTrueRWT = JXSEC_UNKNOWN And Not eTrueRWB = JXSEC_UNKNOWN Then
            
            If eTrueRWT = JXSEC_TOP_FLANGE_LEFT Then
                ' 270 deg, no reflect, I
                oPointMap.Add 17, CStr(3)
                oPointMap.Add 9, CStr(11)
                oPointMap.Add 11, CStr(15)
                oPointMap.Add 15, CStr(23)
            Else
                ' 270 deg, reflect, I
                oPointMap.Add 9, CStr(3)
                oPointMap.Add 17, CStr(11)
                oPointMap.Add 15, CStr(15)
                oPointMap.Add 11, CStr(23)
            End If
            
        ElseIf Not eTrueRWB = JXSEC_UNKNOWN Then
            ' 270 deg, no reflect, C
            oPointMap.Add 17, CStr(3)
            oPointMap.Add 3, CStr(11)
            oPointMap.Add 11, CStr(15)
            oPointMap.Add 15, CStr(23)
        
        ElseIf Not eTrueRWT = JXSEC_UNKNOWN Then
            ' 270 deg, reflect, C
            oPointMap.Add 3, CStr(3)
            oPointMap.Add 17, CStr(11)
            oPointMap.Add 15, CStr(15)
            oPointMap.Add 11, CStr(23)
            
        ElseIf eTrueWLT = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or etrueWLB = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
            ' 270 deg, no reflect, T
            oPointMap.Add 17, CStr(3)
            oPointMap.Add 9, CStr(11)
            oPointMap.Add 11, CStr(15)
            oPointMap.Add 15, CStr(23)
            
        ElseIf etrueWLB = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or eTrueWLT = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
            ' 270 deg, reflect, T
            oPointMap.Add 9, CStr(3)
            oPointMap.Add 17, CStr(11)
            oPointMap.Add 15, CStr(15)
            oPointMap.Add 11, CStr(23)
            
        ElseIf eTrueTFLB = JXSEC_WEB_RIGHT And eTrueTFRB = JXSEC_UNKNOWN Then
            ' 270 deg, no reflect, A
            oPointMap.Add 17, CStr(3)
            oPointMap.Add 3, CStr(11)
            oPointMap.Add 11, CStr(15)
            oPointMap.Add 15, CStr(23)

        ElseIf eTrueBFLT = JXSEC_WEB_RIGHT And eTrueBFRT = JXSEC_UNKNOWN Then
            ' 270 deg, reflect, A
            oPointMap.Add 3, CStr(3)
            oPointMap.Add 17, CStr(11)
            oPointMap.Add 15, CStr(15)
            oPointMap.Add 11, CStr(23)
                        
        ElseIf eTrueT = JXSEC_WEB_LEFT Then
            ' 270 deg, no reflect, FB, RECT
            oPointMap.Add 23, CStr(3)
            oPointMap.Add 3, CStr(11)
            oPointMap.Add 11, CStr(15)
            oPointMap.Add 15, CStr(23)
            
        Else
            ' 270 deg, no reflect, FB, RECT
            oPointMap.Add 3, CStr(3)
            oPointMap.Add 23, CStr(11)
            oPointMap.Add 15, CStr(15)
            oPointMap.Add 11, CStr(23)
        
        End If
    
    ElseIf eTrueWR = JXSEC_TOP_FLANGE_LEFT_BOTTOM Then
            
        ' 90 deg, no reflect, T
        oPointMap.Add 11, CStr(3)
        oPointMap.Add 15, CStr(11)
        oPointMap.Add 17, CStr(14)
        oPointMap.Add 18, CStr(50)
        oPointMap.Add 21, CStr(15)
        oPointMap.Add 3, CStr(17)
        oPointMap.Add 8, CStr(18)
        oPointMap.Add 9, CStr(23)
 
    ElseIf eTrueWR = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
    
        If eTrueWRT = JXSEC_TOP_FLANGE_LEFT_BOTTOM Then
    
            '90 deg, reflect, T
            oPointMap.Add 15, CStr(3)
            oPointMap.Add 11, CStr(11)
            oPointMap.Add 9, CStr(14)
            oPointMap.Add 8, CStr(50)
            oPointMap.Add 3, CStr(15)
            oPointMap.Add 23, CStr(17)
            oPointMap.Add 18, CStr(18)
            oPointMap.Add 17, CStr(23)
        
        ElseIf eTrueBFLT = JXSEC_WEB_RIGHT_TOP Then
        
            '90 deg, no reflect, BUTL3
            oPointMap.Add 11, CStr(3)
            oPointMap.Add 15, CStr(11)
            oPointMap.Add 17, CStr(15)
            oPointMap.Add 18, CStr(20)
            oPointMap.Add 21, CStr(21)
            oPointMap.Add 3, CStr(23)
            
        ElseIf eTrueTFLB = JXSEC_WEB_RIGHT_TOP Then
            '90 deg, reflect, BUTL3
            oPointMap.Add 15, CStr(3)
            oPointMap.Add 11, CStr(11)
            oPointMap.Add 3, CStr(15)
            oPointMap.Add 21, CStr(17)
            oPointMap.Add 8, CStr(18)
            oPointMap.Add 17, CStr(23)
            
        ElseIf eTrueBFR = JXSEC_BOTTOM Then
        
            '90 deg, no reflect, UA
            oPointMap.Add 11, CStr(3)
            oPointMap.Add 15, CStr(11)
            oPointMap.Add 17, CStr(15)
            oPointMap.Add 18, CStr(20)
            oPointMap.Add 21, CStr(21)
            oPointMap.Add 3, CStr(23)
                   
        ElseIf eTrueTFR = JXSEC_BOTTOM Then
        
           '90 deg, reflect, UA
            oPointMap.Add 15, CStr(3)
            oPointMap.Add 11, CStr(11)
            oPointMap.Add 3, CStr(15)
            oPointMap.Add 21, CStr(17)
            oPointMap.Add 18, CStr(18)
            oPointMap.Add 17, CStr(23)

        End If
        
    ElseIf eTrueWR = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
    
        If eTrueTFRB = JXSEC_WEB_RIGHT And eTrueTFLB = JXSEC_UNKNOWN Then
            ' 270 deg, no reflect, L
            oPointMap.Add 23, CStr(3)
            oPointMap.Add 3, CStr(11)
            oPointMap.Add 11, CStr(15)
            oPointMap.Add 14, CStr(17)
            oPointMap.Add 20, CStr(18)
            oPointMap.Add 21, CStr(23)

        ElseIf eTrueBFRT = JXSEC_WEB_RIGHT And eTrueBFLT = JXSEC_UNKNOWN Then
            ' 270 deg, reflect, L
            oPointMap.Add 3, CStr(3)
            oPointMap.Add 23, CStr(11)
            oPointMap.Add 21, CStr(15)
            oPointMap.Add 20, CStr(20)
            oPointMap.Add 14, CStr(21)
            oPointMap.Add 11, CStr(23)
        End If

    ElseIf eTrueWR = JXSEC_TOP_FLANGE_RIGHT_TOP Then
            
        If eTrueTFR = JXSEC_TOP Then
        
            ' 270 deg, no reflect, BUTL3
            oPointMap.Add 17, CStr(3)
            oPointMap.Add 3, CStr(11)
            oPointMap.Add 11, CStr(15)
            oPointMap.Add 14, CStr(17)
            oPointMap.Add 50, CStr(18)
            oPointMap.Add 15, CStr(23)

        Else
        
            ' 270 deg, no reflect, BUTL3
            oPointMap.Add 3, CStr(3)
            oPointMap.Add 17, CStr(11)
            oPointMap.Add 15, CStr(15)
            oPointMap.Add 50, CStr(20)
            oPointMap.Add 14, CStr(21)
            oPointMap.Add 11, CStr(23)
        
        End If
    ElseIf eTrueOT = JXSEC_OUTER_TUBE Then
                    
            oPointMap.Add 2, CStr(3)
            oPointMap.Add 8, CStr(11)
            oPointMap.Add 8, CStr(15)
            oPointMap.Add 2, CStr(23)
    
    End If
    
    Set GetPointMappingFromEdgeMapping = oPointMap
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetPointMappingFromEdgeMapping").Number
    
End Function

'*****************************************************************************************
'Function
'   MbrToTube_IsthisBorderACCase
'
'Description:
'   For Tube-bounding Cases, given the bounding part and bounded part
'    returns if this is a border case of 'ToCenter' where flange slightly goes out of the tube.
'
'Input
'   oSelectorLogic
'   bDonotCheckWPCase: If already checked in calling method use this flag to
'                      avoid duplicate check in current method
'Return
' One of the following Const values:
'BorderAC_OSOS = 1
'BorderAC_FCOS_TOP = 2
'BorderAC_FCOS_BTM = 3
'BorderAC_ToCenter = 4
'NOT_BorderAC = 5
'
'Assumptions:
'  1. Top/Bottom Flange is assumed to be existing for the bounded part:
'  Consider case bounded is similar to FlatBar without Top/Bottom flange: 'GetSelForMbrBoundedToTube'
'  already identifies 'ToCenter' case correctly. Border cases are applicable only if bounded has
'  Top/Bottom flange where portion of flange is hitting the bounding member and another portion of
'  the flange is outside the tube.
' 2. Current Case is Web Penetrated
' 3. Bounding should be a Tube/BUTube
'
'Exceptions
' 1. Bounded is a Member
'***************************************************************************
Private Function MbrToTube_IsthisBorderACCase(oBoundedPart As Object, oBoundingPart As Object) As Integer
    Const MT = "MbrToTube_IsthisBorderACCase"
    On Error GoTo ErrorHandler
    
    If Not TypeOf oBoundedPart Is ISPSMemberPartCommon Then Exit Function 'Check Input

    MbrToTube_IsthisBorderACCase = 0 'Initialize return value to Zero
    
    Dim sMsg As String
    
    'Set bounded to SDO wrapper
    Dim oSDO_Bounded As StructDetailObjects.MemberPart
    Set oSDO_Bounded = New StructDetailObjects.MemberPart
    Set oSDO_Bounded.object = oBoundedPart
    
    'Need to determine which of the Top/Bottom flanges exist
    Dim bTopFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bBottomFlangeRight As Boolean
    CrossSection_Flanges oSDO_Bounded.object, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
    If Not (bTopFlangeLeft Or bBottomFlangeLeft Or bTopFlangeRight Or bBottomFlangeRight) Then
        '*** No flanges exists, EXIT ***
        Exit Function
    End If
       
    'Get bounded flange thickness
    Dim dFlangeThickness As Double
    dFlangeThickness = oSDO_Bounded.flangeThickness
    
    Dim dTF_DistanceFromBdg  As Double
    Dim dBF_DistanceFromBdg  As Double
    
    'Get distance from top flange top
    dTF_DistanceFromBdg = GetDistanceFromBounding(oBoundingPart, oBoundedPart, _
                                JXSEC_TOP)
    If Not (bTopFlangeLeft Or bTopFlangeRight) Then
        'Top flange does not exist. So if dTF_DistanceFromBdg > 0.1mm exit
        If dTF_DistanceFromBdg > TOLERANCE_VALUE Then Exit Function
    End If
    
    'Get distance from bottom flange bottom
    dBF_DistanceFromBdg = GetDistanceFromBounding(oBoundingPart, oBoundedPart, _
                JXSEC_BOTTOM)
    If Not (bBottomFlangeLeft Or bBottomFlangeRight) Then
        'Bottom flange does not exist. So if bBottomFlangeLeft > 0.1mm exit
        If dBF_DistanceFromBdg > TOLERANCE_VALUE Then Exit Function
    End If
        
    'Identify AC Cases
    If dTF_DistanceFromBdg > dFlangeThickness Or _
       dBF_DistanceFromBdg > dFlangeThickness Then
        'This is a NOT a BorderAC case
        MbrToTube_IsthisBorderACCase = NOT_BorderAC
    ElseIf dTF_DistanceFromBdg > dFlangeThickness / 2 And _
            dBF_DistanceFromBdg > dFlangeThickness / 2 Then
        'Identified this is a BorderAC Outside-and-Outside case: flange center plane is above/below the
        'bounding tube
        MbrToTube_IsthisBorderACCase = BorderAC_OSOS
    ElseIf dTF_DistanceFromBdg > dFlangeThickness / 2 Then
        'Identified this is a BorderAC Face-and-Outside at Top case: flange center plane is above
        'bounding tube
        MbrToTube_IsthisBorderACCase = BorderAC_FCOS_TOP
    ElseIf dBF_DistanceFromBdg > dFlangeThickness / 2 Then
        'Identified this is a BorderAC Face-and-Outside at Bottom case: flange center plane is below
        'bounding tube
        MbrToTube_IsthisBorderACCase = BorderAC_FCOS_BTM
    ElseIf dTF_DistanceFromBdg < dFlangeThickness / 2 And dBF_DistanceFromBdg < dFlangeThickness / 2 Then
        MbrToTube_IsthisBorderACCase = BorderAC_ToCenter
    Else
        'Code execution is not expected to reach here; just a standard practice
        'Return Zero, already initialized it.
    End If
    
    'Cleanup
    Set oSDO_Bounded = Nothing

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT, sMsg).Number
End Function


'********************************************************************
' TranslateAxisCurve
'   Translates the Axis Curve of the Member Part
'    Firstly from the Current Postion to Center of the Tube
'    Then From Center of the Tube translates the Axis curve to the
'     if bTanslateToTop = true
'         +ve V direction of Sketching plane
'     else
'         to -ve V direction of Sketching plane
'    And then, if The Axis is Curve then after the translation
'    projects the translated Axis Curve on the Sketching plane in the
'    skecthing plane Normal Direction
'
'   In:
'   Out:
'
' Delete below notes later -------------->
'NOTES THAT compares below method with that in SDBoundedUSSRule **********************
' 1. If we are writing a new method here, 'oEndCutObject' can be deleted since it is NOT used.
' 2. 'Dim oConverter As CommonSymbolUtils.STGeomConverter' is unused in BoundedCustomUtility.bas
' 3. In BoundedCustomRule.cls, IJBoundedUSSRule_ResolveBoundedEdge passes oConverter.ViewTransform
'    to the method where as here clone of 'oSketchMatrix' (Alligators prepared this) is used.
' 4. Based on point 3, inputs to oSketchingPlane.SetNormal and oSketchingPlane.SetUDirection differ.
'
'********************************************************************
Public Sub TranslateAxisCurve(oBoundedObject As Object, _
                              oMemberPart As Object, _
                              oAxisCurve As Object, _
                              oTransform As DT4x4, _
                              bTanslateToTop As Boolean, _
                              oNewAxisCurve As Object, _
                              Optional oEndCutObject As Object, _
                              Optional oBoundedPosition As Object = Nothing)

    Const MT = "TranslateAxisCurve"
    On Error GoTo ErrorHandler
    Dim dBoundedSize As Double
    
    Dim oTemp_Vvec As IJDVector
    Dim oTranslateMatrix As AutoMath.DT4x4
    
    Dim oCmplx As ComplexString3d
    Dim curveElms As IJElements
    Dim oGeometryFactory As GeometryFactory
    
    Set curveElms = New JObjectCollection
    If TypeOf oAxisCurve Is ComplexString3d Then
        Set oCmplx = oAxisCurve
        oCmplx.GetCurves curveElms
    Else
        curveElms.Add oAxisCurve
    End If

    Set oGeometryFactory = New GeometryFactory
    Set oCmplx = oGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, _
                                                                  curveElms)
    Set curveElms = Nothing
    
    Dim oBoundedPort As IJPort
    Dim oBoundedPart As ISPSMemberPartCommon
    Dim oStructPorfile As IJStructProfilePart
    Dim oPos As IJDPosition
    Dim oPoint As IJPoint
    Dim oXSecMatirix As IJDT4x4
    Dim oXSecUvec As IJDVector
    Dim oXSecVvec As IJDVector
    Dim oTranslateMatrixUDir As IJDT4x4
    Dim oTranslateMatrixVDir As IJDT4x4
    Dim oSPSCrossSec As ISPSCrossSection
    Dim dU As Double
    Dim dV As Double
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    Dim dCardinalPoint As Double
    Dim oSPSSplitAxisPort As ISPSSplitAxisPort
    Dim eSPSPortIndex As SPSMemberAxisPortIndex

    '-------------------------------------------------------------------
    'Translates the Axis Curve of the Member Part
    'from the Current Postion to Center of the Tube i.e Cardinal Point 5
    '-------------------------------------------------------------------
    
    If TypeOf oBoundedObject Is IJPort Then
        Set oBoundedPort = oBoundedObject
        Set oBoundedPart = oBoundedPort.Connectable

        If TypeOf oBoundedObject Is ISPSSplitAxisPort Then
             Set oSPSSplitAxisPort = oBoundedObject
             eSPSPortIndex = oSPSSplitAxisPort.PortIndex
        Else
             '(unknown case)need to know the type of End Port
             'so that Strat/End point of Member can be known
             'probably need to throw error here
        End If
        
        Set oPos = New DPosition
        
        If oBoundedPosition Is Nothing Then
            Set oPoint = oBoundedPart.PointAtEnd(eSPSPortIndex)
            oPoint.GetPoint dx, dy, dz
            oPos.Set dx, dy, dz
        Else
            Set oPos = oBoundedPosition
        End If

        If TypeOf oBoundedPart Is IJStructProfilePart Then
            Set oStructPorfile = oBoundedPart
        Else
           'Unknown case
           'currently need to throw error
        End If

        Set oXSecMatirix = oStructPorfile.GetCrossSectionMatrixAtPoint(oPos)

        Set oXSecUvec = New dVector
        Set oXSecVvec = New dVector

        oXSecUvec.Set oXSecMatirix.IndexValue(0), oXSecMatirix.IndexValue(1), oXSecMatirix.IndexValue(2)
        oXSecVvec.Set oXSecMatirix.IndexValue(4), oXSecMatirix.IndexValue(5), oXSecMatirix.IndexValue(6)

        Set oSPSCrossSec = oBoundedPart.CrossSection

        dCardinalPoint = oSPSCrossSec.CardinalPoint
        oSPSCrossSec.GetCardinalPointDelta Nothing, dCardinalPoint, 5, dU, dV

        oXSecUvec.Length = dU
        oXSecVvec.Length = dV

        Set oTranslateMatrixUDir = New DT4x4
        Set oTranslateMatrixVDir = New DT4x4
        oTranslateMatrixUDir.LoadIdentity
        oTranslateMatrixVDir.LoadIdentity

        oTranslateMatrixUDir.Translate oXSecUvec
        oCmplx.Transform oTranslateMatrixUDir

        oTranslateMatrixVDir.Translate oXSecVvec
        oCmplx.Transform oTranslateMatrixVDir
    Else
      ' Need to know the position of the Start/End Port
      ' currently probably need to throw error here
    End If
    
    '-------------------------------------------------------------------
    'From Center of the Tube translates the Axis curve to the Top/Btm
    ' of the Tube
    '-------------------------------------------------------------------
    dBoundedSize = GetMemberTubeRadius(oMemberPart)
    
    Set oTemp_Vvec = New AutoMath.dVector
    oTemp_Vvec.Set oTransform.IndexValue(8), _
                   oTransform.IndexValue(9), _
                   oTransform.IndexValue(10)
                                       

    If bTanslateToTop Then
        oTemp_Vvec.Length = dBoundedSize
    Else
        oTemp_Vvec.Length = -dBoundedSize
    End If
        
    Set oTranslateMatrix = New AutoMath.DT4x4
    oTranslateMatrix.LoadIdentity
    oTranslateMatrix.Translate oTemp_Vvec
    oCmplx.Transform oTranslateMatrix
    Set oNewAxisCurve = oCmplx
    
    Dim oWireBodyToProject As IJWireBody
    Dim oProjectedWireBody As IJWireBody
    Dim oProjectUtil As IMSModelGeomOps.Project
    Dim oSketchingPlane As IJPlane
    Dim oSketchingPlaneNormal As IJDVector
    
    '-------------------------------------------------------------------
    'if The Axis is Curve(i.e Non Linear) then after the translation
    'project the translated Axis Curve on the Sketching plane in the
    'sketching plane Normal Direction
    '-------------------------------------------------------------------
    If TypeOf oMemberPart Is ISPSMemberPartCurve Then
      'Continue
    Else
      Exit Sub
    End If
    
    Set oSketchingPlaneNormal = New dVector

    oSketchingPlaneNormal.Set oTransform.IndexValue(0), oTransform.IndexValue(1), oTransform.IndexValue(2)
    oSketchingPlaneNormal.Length = 1

    Set oSketchingPlane = New Plane3d
    oSketchingPlane.SetNormal oTransform.IndexValue(0), oTransform.IndexValue(1), oTransform.IndexValue(2)
    oSketchingPlane.SetRootPoint oTransform.IndexValue(12), oTransform.IndexValue(13), oTransform.IndexValue(14)
    oSketchingPlane.SetUDirection oTransform.IndexValue(4), oTransform.IndexValue(5), oTransform.IndexValue(6)

    Dim GeomOpr As IMSModelGeomOps.DGeomWireFrameBody
    Dim oElemCurves As IJElements
    oCmplx.GetCurves oElemCurves

    Set GeomOpr = New DGeomWireFrameBody
    Dim oObj As Object

    Set oObj = GeomOpr.CreateSmartWireBodyFromGTypedCurves(Nothing, oElemCurves)

    Set oWireBodyToProject = oObj
    
    Set oProjectUtil = New Project
    oProjectUtil.CurveAlongVectorOnToSurface Nothing, _
                                             oSketchingPlane, _
                                             oWireBodyToProject, _
                                             oSketchingPlaneNormal, _
                                             Nothing, _
                                             oProjectedWireBody

    Set oNewAxisCurve = oProjectedWireBody

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
    
End Sub



'********************************************************************
' GetMemberTubeRadius
'
'   In:
'   Out:
'********************************************************************
Public Function GetMemberTubeRadius(oMemberObject As ISPSMemberPartCommon) As Double
Const MT = "GetMemberTubeRadius"
On Error GoTo ErrorHandler
    
    Dim dXpnt2 As Double
    Dim dXpnt5 As Double
    Dim dYpnt2 As Double
    Dim dYpnt5 As Double
    
    Dim oSPSCrossSection As ISPSCrossSection
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic

    GetMemberTubeRadius = 0#
    If oMemberObject Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oMemberObject Is ISPSMemberPartPrismatic Then
        Exit Function
    End If
    
    Set oMemberPartPrismatic = oMemberObject
    Set oSPSCrossSection = oMemberPartPrismatic.CrossSection
    oSPSCrossSection.GetCardinalPointOffset 2, dXpnt2, dYpnt2
    oSPSCrossSection.GetCardinalPointOffset 5, dXpnt5, dYpnt5
    GetMemberTubeRadius = Abs(dYpnt5 - dYpnt2)
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function

'*******************************
'AreTwoportNormalsParallel
'This method will Check whether two port normals are parallel or not
'Bounding - Plate
'Bounded- Member
'*******************************

Public Function AreTwoportNormalsParallel(oBoundingMbr As Object, oBoundedMbr As Object) As Boolean

    Dim oStructGeomUtils As GSCADStructGeomUtilities.PartInfo
    Set oStructGeomUtils = New PartInfo
        Dim oWLport As IJPort
        Dim oBasePort As IJPort
        Set oBasePort = GetBaseOffsetOrLateralPortForObject(oBoundingMbr, BPT_Base)
        Set oWLport = GetLateralSubPortBeforeTrim(oBoundedMbr, JXSEC_WEB_LEFT)
        Dim oNormal1 As IJDVector
        Dim oNormal2 As IJDVector
        Set oNormal1 = oStructGeomUtils.GetPortNormal(oBasePort, True)
        Set oNormal2 = oStructGeomUtils.GetPortNormal(oWLport, True)
        Dim dDot As Double
        dDot = oNormal1.Dot(oNormal2)
        If dDot = 1 Or dDot = -1 Then
            AreTwoportNormalsParallel = True
        End If
End Function
' ***************************************************************************
' Method : Get_CacheEdgeMapping
' In  : AC or EC
' Out :As Boolean value whether the cached data is available or not
' And also this method will return the edgeMapping collection, webpenetrated as boolean,section alias if cached
' ***************************************************************************
Private Function Get_CacheEdgeMapping(oACOrEC As Object, _
                                      ByRef oEdgeMapping As Collection, _
                                      bPenetratesWeb As Boolean, _
                                      sectionAlias As Long) As Boolean
    Const MT = "Get_CacheEdgeMapping"
    ' --------------------------------
    ' Retrieve the edge mapping string
    ' --------------------------------
    On Error GoTo ErrorHandler
    
    Get_CacheEdgeMapping = False

    Dim oAttributes As IJDAttributes
    Set oAttributes = oACOrEC

    Dim oAttributesCol As IJDAttributesCol
    On Error Resume Next
    Set oAttributesCol = oAttributes.CollectionOfAttributes("IJUAMbrACCacheStorage2")
    Err.Clear
    On Error GoTo ErrorHandler
    
    If oAttributesCol Is Nothing Then
        GoTo CleanUp
    End If
    
    Dim strValue As String
    strValue = oAttributesCol.Item("EdgeMapping").value
    If strValue = vbNullString Or strValue = "" Then
        GoTo CleanUp
    End If
    
    ' -------------------------------------------------------
    ' Split the attribute into an array of name/value strings
    ' -------------------------------------------------------
    Dim attrArray() As String
    attrArray = Split(strValue, "+", , vbTextCompare)
    
    ' -----------------------------------
    ' Loop through the name/value strings
    ' -----------------------------------
    Dim i As Long
    Dim equalPos As Integer
    
    Dim key As String
    Dim value As String
    
    For i = 0 To UBound(attrArray)
        equalPos = InStr(1, attrArray(i), "=", vbTextCompare)

        key = Left(attrArray(i), equalPos - 1)
        value = Right(attrArray(i), Len(attrArray(i)) - equalPos)

        oEdgeMapping.Add value, key
    Next i
    
    Dim dValue As Long
    
    dValue = oAttributesCol.Item("SectionAlias").value
    If dValue < 0 Then GoTo CleanUp
    sectionAlias = dValue
    
    On Error Resume Next
    Set oAttributesCol = oAttributes.CollectionOfAttributes("IJUAMbrACCacheStorage")
    Err.Clear
    On Error GoTo ErrorHandler
    
    bPenetratesWeb = oAttributesCol.Item("IsWebPenetrated").value
    Get_CacheEdgeMapping = True
    
CleanUp:

    Set oAttributesCol = Nothing
    Set oAttributes = Nothing
        
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function
' ***************************************************************************
' Method : Set_CacheEdgeMapping
' In  : AC or EC
' return : edgemap as collection
' ***************************************************************************
Private Sub Set_CacheEdgeMapping(oACOrEC As Object, edgeMap As Collection)
    
    Const MT = "Set_CacheEdgeMapping"
    On Error GoTo ErrorHandler

    Dim oAttributes As IJDAttributes
    Set oAttributes = oACOrEC
    
    Dim oAttributesCol As IJDAttributesCol
    
    On Error Resume Next
    Set oAttributesCol = oAttributes.CollectionOfAttributes("IJUAMbrACCacheStorage2")
    Err.Clear
    On Error GoTo ErrorHandler
    
    If oAttributesCol Is Nothing Then
        Exit Sub
    End If
    
    Dim strEdgeMap As String
    Dim key As JXSEC_CODE
    
    If edgeMap.Count > 0 Then
        key = ReverseMap(edgeMap.Item(1), edgeMap)
        strEdgeMap = key & "=" & edgeMap.Item(1)
    End If
    
    Dim i As Long
    
    For i = 2 To edgeMap.Count
        key = ReverseMap(edgeMap.Item(i), edgeMap)
        If key > 0 Then
            strEdgeMap = strEdgeMap & "+" & key & "=" & edgeMap.Item(i)
        Else
            Exit For
        End If
    Next i
    
    oAttributesCol.Item("EdgeMapping").value = strEdgeMap
        
CleanUp:

    Set oAttributesCol = Nothing
    Set oAttributes = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub
' **************************************************************************
' Method : GetEdgeMap
' In     : Only AC or EC, bounding port and bounded port
' Out    : Edgemapping collection
' if cached data is available, get that info. If not, call edgemapping rule to get the same and cache the data.
'***************************************************************************
Public Function GetEdgeMap(oACOrEC As Object, _
                           oBoundingPort As IJPort, _
                           oBoundedPort As IJPort, _
                           Optional ByRef sectionAlias As Long, _
                           Optional ByRef bPenetratesWeb As Boolean, _
                           Optional bForceRecompute As Boolean = False) As Collection

    Const MT = "GetEdgeMap"
    On Error GoTo ErrorHandler
    
        ' ------------------------------------------------------
        ' Try getting the cache data:
        ' When it is axis AC, get the cached data on AC (input can be AC or EC)
        ' When it is Generic AC and bounding collection is one, get the cached data on AC (input can be AC or EC)
        ' When it is Generic AC and bounding collection is more than one, and input is EC, get cached data on EC
        '                                                                     input is AC, making object as nothing to avoid getting cached data on AC
        ' ------------------------------------------------------
    Dim oObjectWithCachedData As Object
    Set oObjectWithCachedData = oACOrEC
        
    Dim eGetACType As eACType
    Dim sACItemName As String
    Dim oACObject As IJAppConnection

    If Not oACOrEC Is Nothing Then
    
            If TypeOf oACOrEC Is IJStructFeature Then
                AssemblyConnection_SmartItemName oACOrEC, sACItemName, oACObject
            Else
                Set oACObject = oACOrEC
            End If
            eGetACType = GetMbrAssemblyConnectionType(oACObject)
            If eGetACType = ACType_Mbr_Generic Or eGetACType = ACType_Stiff_Generic Then
                Dim oReferencesCollection As IJDReferencesCollection
                Dim oEditJDArgument As IJDEditJDArgument
                Dim oBoundingObjectColl As IJElements
                Dim iBoundingObjectsCount As Integer
                
                'Get Bounding Object Collection Count
                Set oReferencesCollection = GetRefCollFromSmartOccurrence(oACObject)
                Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
                Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
                iBoundingObjectsCount = oBoundingObjectColl.Count
                If iBoundingObjectsCount = 1 Then
                    Set oObjectWithCachedData = oACObject
                ElseIf iBoundingObjectsCount > 1 Then
                    If TypeOf oACOrEC Is IJAppConnection Then
                        Set oObjectWithCachedData = Nothing
                    End If
                End If
            ElseIf Not eGetACType = ACType_Mbr_Generic And Not eGetACType = ACType_Stiff_Generic Then
                
                Set oObjectWithCachedData = oACObject
                
            End If
        End If
    
    ' -----------------------------------------------------------------------------------
    ' If force recompute is off, and AC or web cut is passed in, retrieve data from cache
    ' -----------------------------------------------------------------------------------
    Dim oTempMap As New Collection

    If (bForceRecompute = False) And (Not oObjectWithCachedData Is Nothing) Then
        If Get_CacheEdgeMapping(oObjectWithCachedData, _
                                oTempMap, _
                                bPenetratesWeb, _
                                sectionAlias) = True Then
            ' Got the details from Cache.. exit here....
            Set GetEdgeMap = oTempMap
            Exit Function
        End If
    End If

    ' ---------------------------------------------------------------
    ' Calculate the mapping, section alias, and penetration condition
    ' ---------------------------------------------------------------
    Dim oEndCutMappingRule As IJEndCutEdgeMappingRule

    Set oTempMap = New Collection
    Set oEndCutMappingRule = CreateEdgeMappingRuleSymbolInstance
    oEndCutMappingRule.GetEdgeMapping oBoundingPort, oBoundedPort, sectionAlias, bPenetratesWeb, oTempMap

    Set GetEdgeMap = oTempMap
    
    Set oEndCutMappingRule = Nothing
        
    ' ----------------------
    ' Cache the mapping data
    ' ----------------------
    If oObjectWithCachedData Is Nothing Then
       'if the object on which to be cached is not available
       'not need to cache.
       Exit Function
    End If
    
    Set_CacheEdgeMapping oObjectWithCachedData, oTempMap
    
    ' -------------------------------
    ' Cache the penetration condition
    ' -------------------------------
    Dim oAttributes As IJDAttributes
    Set oAttributes = oObjectWithCachedData
    
    Dim oAttributesCol As IJDAttributesCol
    Dim IsAttExists As Boolean
    On Error Resume Next
    Set oAttributesCol = oAttributes.CollectionOfAttributes("IJUAMbrACCacheStorage")
        
    If Not oAttributesCol Is Nothing Then
        oAttributesCol.Item("IsWebPenetrated").value = bPenetratesWeb
    End If
    
    ' -----------------------
    ' Cache the section alias
    ' -----------------------
    On Error Resume Next
    Dim oCache2AttributeCol As IJDAttributesCol
    
    Set oCache2AttributeCol = oAttributes.CollectionOfAttributes("IJUAMbrACCacheStorage2")
    Err.Clear
    On Error GoTo ErrorHandler
    
    If oCache2AttributeCol Is Nothing Then
        Exit Function
    End If
    
    oCache2AttributeCol.Item("SectionAlias").value = sectionAlias
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function

Public Sub GetMultiBoundingEdgeMap(oACOrEC As Object, eEndCutType As eEndCutTypes, oMappedPortsCollection As JCmnShp_CollectionAlias, _
                    Optional AnglesColl As Collection = Nothing, Optional PostionColl As Collection = Nothing, _
                    Optional TopLeftInsidePort As IJPort = Nothing, Optional TopLeftInsidePos As IJDPosition = Nothing, _
                    Optional BtmRightInsidePort As IJPort = Nothing, Optional BtmRightInsidePos As IJDPosition = Nothing, _
                    Optional oBoundedPart As Object = Nothing)

    Const MT = "GetMultiBoundingEdgeMap"
    On Error GoTo ErrorHandler
    
    Dim oRefPortColl As Collection
    Dim oEndCutMappingRule As IJEndCutEdgeMappingRuleEx
    Dim oACObject As Object
    
    If TypeOf oACOrEC Is IJAppConnection Then
        Set oACObject = oACOrEC
    Else
        AssemblyConnection_SmartItemName oACOrEC, , oACObject
    End If
        
    If oACObject Is Nothing Then
        GoTo ErrorHandler
    End If
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    Dim sMsg As String
    
    InitMemberConnectionData oACObject, oBoundedData, oBoundingData, lStatus, sMsg
    If GetMbrAssemblyConnectionType(oACObject) = ACType_Mbr_Generic Then
        GetRefPortColl oACObject, oRefPortColl
    Else
        'Other than Genric AC, no other type of AC is currently supported
        Exit Sub
    End If
    
    Set oBoundedPart = oBoundedData.MemberPart
        
    Dim eEndCutPlane As StrDetEndCutMappingPlane
    
    If eEndCutType = WebCut Then
        eEndCutPlane = WebMidThickness
    ElseIf eEndCutType = FlangeCutTop Then
        eEndCutPlane = TopFlangeMidThickness
    ElseIf eEndCutType = FlangeCutBottom Then
        eEndCutPlane = BottomFlangeMidThickness
    Else
        GoTo ErrorHandler
    End If
    
    Set oMappedPortsCollection = New Collection
    Set AnglesColl = New Collection
    Set PostionColl = New Collection
    
    Set oEndCutMappingRule = CreateEdgeMappingRuleSymbolInstance
    oEndCutMappingRule.GetMultipleBoundaryEdgeMapping oRefPortColl, oBoundedData.AxisPort, eEndCutPlane, oMappedPortsCollection, AnglesColl, PostionColl, TopLeftInsidePort, _
                                                TopLeftInsidePos, BtmRightInsidePort, BtmRightInsidePos
    
    Dim bTestReults As Boolean
    bTestReults = False
    
    If bTestReults Then
        
        Dim oMB As IJDModelBody
        Dim oTXPort As IJPort
        Dim iIndex As Long
        Dim i As Long
        
        'Dump Mapped Ports
        For iIndex = 1 To oMappedPortsCollection.Count
            Set oTXPort = oMappedPortsCollection.Item(iIndex)
            Set oMB = oTXPort.Geometry
            oMB.DebugToSATFile ("C:\MappedPort" & iIndex & ".sat")
        Next
    
        'Dump Angles Collection
        For iIndex = 1 To AnglesColl.Count
            LogError Err, eEndCutType & " Angle no." & iIndex & " = " & AnglesColl.Item(iIndex)
        Next
        
        'Dump Position Collection
        Dim oPos As IJDPosition
        Set oPos = New DPosition
        For i = 1 To PostionColl.Count
            Set oPos = PostionColl.Item(i)
            LogError Err, eEndCutType & " Pos. No. " & i & " = " & oPos.x & "; " & oPos.y & "; " & oPos.z
        Next
        
        'Dump Top Inside Info.
        If Not (TopLeftInsidePort Is Nothing) Then
             Set oMB = TopLeftInsidePort.Geometry
             oMB.DebugToSATFile ("C:\TopInsideIntersectedPort.sat")
             LogError Err, eEndCutType & " Top Inside Pos." & " = " & TopLeftInsidePos.x & "; " & TopLeftInsidePos.y & "; " & TopLeftInsidePos.z
        End If
        
        'Dump Btm Inside Info.
        If Not (BtmRightInsidePort Is Nothing) Then
             Set oMB = BtmRightInsidePort.Geometry
             oMB.DebugToSATFile ("C:\BtmInsideIntersectedPort.sat")
             LogError Err, eEndCutType & " Btm Inside Pos." & " = " & BtmRightInsidePos.x & "; " & BtmRightInsidePos.y & "; " & BtmRightInsidePos.z
        End If
    End If
    
    Exit Sub
    '***************NEW MAPPING RULE USED******************************************************************
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub

Public Sub GetRefPortColl(oACOrEC As Object, oRefPortColl As Collection)

    Const MT = "GetRefPortColl"
    On Error GoTo ErrorHandler
    
    Dim oAppConnection As IJAppConnection
    
    AssemblyConnection_SmartItemName oACOrEC, , oAppConnection
    
    If oAppConnection Is Nothing Then
        Exit Sub
    End If
    
    Dim oReferencesCollection As IJDReferencesCollection
    Dim oEditJDArgument As IJDEditJDArgument
    
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    
    Dim nRefArguments As Long
    nRefArguments = oEditJDArgument.GetCount
    If nRefArguments < 1 Then
        Exit Sub
    End If
    
    Dim iIndex As Long
    Dim oArgObject As Object
    Set oRefPortColl = New Collection
    
    For iIndex = 1 To nRefArguments
        Set oArgObject = oEditJDArgument.GetEntityByIndex(iIndex)
        If TypeOf oArgObject Is IJPort Then
            oRefPortColl.Add oArgObject
        Else
        End If
    Next iIndex
       
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub

Public Function GetMappedPortFromGivenEdgeID(oMappedPortsCollection As Collection, lEdgeId As Long) As IJPort

    Dim lIndex As Integer
    lIndex = lEdgeId - 5000
    
    If (lIndex > 0) And lIndex <= oMappedPortsCollection.Count Then
        Set GetMappedPortFromGivenEdgeID = oMappedPortsCollection.Item(lIndex)
    Else
        Set GetMappedPortFromGivenEdgeID = Nothing
    End If
    
End Function

Public Function GetAngleBtwEdgeIDs(Edge1ID As Long, Edge2ID As Long, AngleCollection As Collection) As Double
    
 Const METHOD = "IsMultiBoundingCase"
 On Error GoTo ErrorHandler
    
    If (Edge1ID = e_JXSEC_MultipleBounding_5001 And Edge2ID = e_JXSEC_MultipleBounding_5002) Or _
         (Edge1ID = e_JXSEC_MultipleBounding_5002 And Edge2ID = e_JXSEC_MultipleBounding_5001) Then
        GetAngleBtwEdgeIDs = AngleCollection.Item(1)
    
    ElseIf (Edge1ID = e_JXSEC_MultipleBounding_5002 And Edge2ID = e_JXSEC_MultipleBounding_5003) Or _
         (Edge1ID = e_JXSEC_MultipleBounding_5003 And Edge2ID = e_JXSEC_MultipleBounding_5002) Then
        GetAngleBtwEdgeIDs = AngleCollection.Item(2)
        
    ElseIf (Edge1ID = e_JXSEC_MultipleBounding_5003 And Edge2ID = e_JXSEC_MultipleBounding_5004) Or _
         (Edge1ID = e_JXSEC_MultipleBounding_5004 And Edge2ID = e_JXSEC_MultipleBounding_5003) Then
        GetAngleBtwEdgeIDs = AngleCollection.Item(3)
    
    ElseIf (Edge1ID = e_JXSEC_MultipleBounding_5004 And Edge2ID = e_JXSEC_MultipleBounding_5005) Or _
         (Edge1ID = e_JXSEC_MultipleBounding_5005 And Edge2ID = e_JXSEC_MultipleBounding_5004) Then
        GetAngleBtwEdgeIDs = AngleCollection.Item(4)
    End If
 Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Function

Public Function IsMultiBoundingCase(ACorEndCut As Object, Optional pMappedPortCollection As JCmnShp_CollectionAlias = Nothing) As Boolean

    Const METHOD = "IsMultiBoundingCase"
    On Error GoTo ErrorHandler
                
    IsMultiBoundingCase = False
    
    Dim oBoundedPart As Object
    GetMultiBoundingEdgeMap ACorEndCut, WebCut, pMappedPortCollection, , , , , , , oBoundedPart
        
    Dim IsFlangeMBounding As Boolean
    If Not pMappedPortCollection Is Nothing Then
        If pMappedPortCollection.Count > 1 Then
            IsMultiBoundingCase = True
            Exit Function
        End If
    End If
    
    IsMultiBoundingCase = IsFlangeMultiBounding(ACorEndCut, oBoundedPart, pMappedPortCollection)
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Function
'*************************************************************************
'Function
'GetBoundingEndPort
'
'Abstract
'   Given the Bounded and Bounding members for an End to End Assembly Connection
'   Determine:
'       1. If the Bounding Member Port is an End port or the AlongAxis port
'       2. If Bounding Member Port is the AlongAxis port
'       3. Return the Bounding Member End Port that is closest to the Bounded Member End Port
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub GetBoundingEndPort(oBoundedPort As IJPort, _
                                oBoundingPort As IJPort, _
                                oBoundingAxisEndPort As ISPSSplitAxisPort)
Const METHOD = "::GetBoundingEndPort"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dDistEnd As Double
    Dim dDistStart As Double
    
    Dim oBoundedPoint As IJPoint
    Dim oBoundingPoint_End As IJPoint
    Dim oBoundingPoint_Start As IJPoint
    
    Dim oBoundingPart As ISPSMemberPartCommon
    Dim oBoundedPart As ISPSMemberPartCommon
    
    
    ' Check that the Supported and Supporting Member data is valid
    If oBoundedPort Is Nothing Then
        sMsg = "Supported Member data is not valid"
        GoTo ErrorHandler
    
    ElseIf oBoundingPort Is Nothing Then
        sMsg = "Supported Member Port data is not valid"
        GoTo ErrorHandler
    
    ElseIf Not TypeOf oBoundedPort Is ISPSSplitAxisPort Then
        sMsg = "Supporting Member data is not valid"
        GoTo ErrorHandler
    End If
    
    Set oBoundedPart = oBoundedPort.Connectable
    Set oBoundingPart = oBoundingPort.Connectable
    
    sMsg = ""
    
    Dim oBoundedAxisPort As ISPSSplitAxisPort
    Set oBoundedAxisPort = oBoundedPort
    '
    ' Have a valid Supported Member data: (Bounded Port is SPSMemberAxisStart or SPSMemberAxisEnd)
    ' Have a valid Supporting Member data: (Bounding Port is SPSMemberAxisAlong)
    ' Calculate the distance from the Supported End Point and the Supporting End points
    ' return the closest Bounding End Port
    Set oBoundedPoint = oBoundedPart.PointAtEnd(oBoundedAxisPort.PortIndex)
    Set oBoundingPoint_End = oBoundingPart.PointAtEnd(SPSMemberAxisEnd)
    Set oBoundingPoint_Start = oBoundingPart.PointAtEnd(SPSMemberAxisStart)
    dDistEnd = oBoundedPoint.DistFromPt(oBoundingPoint_End)
    dDistStart = oBoundedPoint.DistFromPt(oBoundingPoint_Start)
    If dDistStart < dDistEnd Then
        Set oBoundingAxisEndPort = oBoundingPart.AxisPort(SPSMemberAxisStart)
    Else
        Set oBoundingAxisEndPort = oBoundingPart.AxisPort(SPSMemberAxisEnd)
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'GetFlippedPorts
'
'input - App Connection
'Output - Flipped Bounding and Bounded ports as objects
'Abstract - Getting the Flip Primary and Secondary attribute value on the given app connection and outputs the flipped bounding and bounded ports
'           and flip value as boolean based on various conditions
    '           whether to flip the primary and secondary like primary criteria attribute and question "Flip Priamry and Secondary" question answer.
'*************************************************************************
Public Sub GetFlippedPorts(oAppConnection As IJAppConnection, FlippedBoundedPort As Object, FlippedBoundingPort As Object, Optional bConsiderFlipping As Boolean)

    Const MT = "GetFlippedPorts"
    Dim sMsg As String
    sMsg = "Getting Flipped Ports"
    
    Dim AttributeValue As Double
    
    Dim vValue As Variant
    
    Get_AttributeValue oAppConnection, "FlipPrimaryandSecondary", vValue
    
    If vValue Then
        bConsiderFlipping = True
    Else
        bConsiderFlipping = False
    End If
    
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort

    Dim ACPorts As IJElements
    oAppConnection.enumPorts ACPorts
    
    If TypeOf ACPorts.Item(1) Is ISPSSplitAxisAlongPort Then
        Set oBoundingPort = ACPorts.Item(1)
        Set oBoundedPort = ACPorts.Item(2)
    ElseIf TypeOf ACPorts.Item(2) Is ISPSSplitAxisAlongPort Then
        Set oBoundingPort = ACPorts.Item(2)
        Set oBoundedPort = ACPorts.Item(1)
    Else
        Set oBoundedPort = ACPorts.Item(1)
        Set oBoundingPort = ACPorts.Item(2)
    End If
        
    If bConsiderFlipping Then
        If TypeOf oBoundingPort Is ISPSSplitAxisEndPort And TypeOf oBoundedPort Is ISPSSplitAxisEndPort Then

            Set FlippedBoundedPort = oBoundingPort
            Set FlippedBoundingPort = oBoundedPort
        Else
            Dim oBoundedMbr As ISPSMemberPartCommon
            Set oBoundedMbr = oBoundedPort.Connectable
                    
            GetBoundingEndPort oBoundedPort, oBoundingPort, FlippedBoundedPort
            Set FlippedBoundingPort = oBoundedMbr.AxisPort(SPSMemberAxisAlong)
        End If
    Else
        Set FlippedBoundedPort = oBoundedPort
        Set FlippedBoundingPort = oBoundingPort
    End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT, sMsg
End Sub
    
Public Function IsFlangeMultiBounding(ACorEndCut As Object, oBoundedPart As Object, pWebCutMappedPortCollection As JCmnShp_CollectionAlias) As Boolean

    Const METHOD = "::IsFlangeMultiBounding"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Is Flange Multi Bounding ? "
    
    IsFlangeMultiBounding = False
    
    Dim pMappedPortCollection As JCmnShp_CollectionAlias

    If Not pWebCutMappedPortCollection Is Nothing Then
        If pWebCutMappedPortCollection.Count <= 1 Then 'Web is not multi-bounding case, so check for top flange
            If ((HasTopFlange(oBoundedPart)) And Not (IsRectangularMember(oBoundedPart))) Then
                GetMultiBoundingEdgeMap ACorEndCut, FlangeCutTop, pMappedPortCollection
            End If
        End If
    End If
    
    If Not pWebCutMappedPortCollection Is Nothing Then
        If pWebCutMappedPortCollection.Count <= 1 Then 'Web and top flange are not multi-bounding case, so check for bottom flange
            If ((HasBottomFlange(oBoundedPart)) And Not (IsRectangularMember(oBoundedPart))) Then
                GetMultiBoundingEdgeMap ACorEndCut, FlangeCutBottom, pMappedPortCollection
            End If
        End If
    End If
    
    If Not pMappedPortCollection Is Nothing Then
        If pMappedPortCollection.Count > 1 Then
            IsFlangeMultiBounding = True
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Function GetMemberDescriptionFromPropertyDescs(pPropertyDescriptions As IJDPropertyDescription, pObject As Object) As IJDMemberDescription

  Const METHOD = "::GetMemberDescriptionFromPropertyDescs"
  On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Get MemberDescription by looping through all"
    
    Dim oMemberObjects As IJDMemberObjects
    Dim oMD As IJDMemberDescription
    Dim oMemberDescriptions As IJDMemberDescriptions
    Dim iCount As Integer
    
    'Get all the Item Objects
    Set oMemberObjects = pPropertyDescriptions.CAO
    Set oMemberDescriptions = oMemberObjects.MemberDescriptions
    
    'Loop through all the Member descriptions and get the required one
    For iCount = 1 To oMemberDescriptions.Count
        Set oMD = oMemberDescriptions.ItemByDispid(iCount)
        
        'Get Struct feature Member description
        Dim oObject As Object
        Set oObject = pObject
        
        If oMD.object Is oObject Then
            Set GetMemberDescriptionFromPropertyDescs = oMD
            Exit Function
        End If
    Next
  Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function
'*************************************************************************
'Function
'IsConnectionWebPenetrated
'
'Abstract
'   Given the Struct feature or assembly connection returns the boolean value of web pentrated or not
'   Determine:
'       1. whether the Input is struct feature or assembly connection
'            a. if struct fetaure --> get the parent assembly connection object
'       2. Bounding and bounded from the parent assembly connection
'       3. web penetrated or not
'
'input - connection or its child struct feature object
'
'Return -  boolean value, returns true if it is web penetrated case otherwise false
'
'***************************************************************************
Public Function IsConnectionWebPenetrated(connectionOrChildFeature As Object) As Boolean

    IsConnectionWebPenetrated = False
    
    ' Get information about the connection
    Dim oAppConnection As IJAppConnection
    
    If TypeOf connectionOrChildFeature Is IJStructFeature Then
        Dim sACItemName As String
        Dim oACObj As Object
        AssemblyConnection_SmartItemName connectionOrChildFeature, sACItemName, oAppConnection
    ElseIf TypeOf connectionOrChildFeature Is IJAppConnection Then
        Set oAppConnection = connectionOrChildFeature
    Else
        ' Not yet handled
    End If
    
    ' IsWebPenetrated does not work for a generic connection
    If GetMbrAssemblyConnectionType(oAppConnection) = ACType_Mbr_Generic Then
        'yet to handle
    Else
        Dim oBoundedPort As IJPort
        Dim oBoundingPort As IJPort
        GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
        IsConnectionWebPenetrated = IsWebPenetrated(oBoundingPort, oBoundedPort)
    End If

End Function

