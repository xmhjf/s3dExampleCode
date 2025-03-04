Attribute VB_Name = "MemberUtilities"
'*******************************************************************
'
'Copyright (C) 2007-15 Intergraph Corporation. All rights reserved.
'
'File : MemberUtilities.bas
'
'Author : D.A. Trent
'
'Description :
'   Utilites for determining Type Of EndCuts to be Placed on Member to Member Assembly Connections
'   most of these Utilities are copied from (they are copied here for convenience)
'   S:\SmartPlantStructure\Symbols\AssemblyConnections\ConnectionUtilities.bas
'
'History:
'25/Aug/2011  - pnalugol - Addedd new methods to create InsetBrace
'22/Sep/2011  - GH/CM - TR-201985 "Inside" answer to "ShapeAtFace" question has no effect
'                       Modified CreateCornerFeatureBetweenTwoEndCutsByDispID() to support creation
'                       of CF between End Cuts via Axis AC's
'01/12/2011   - vbbheema -  TR-CP-206494  Split selector answers doesn't call proper web cut item.
'                           Modified "AreCrossSectionsIdentical" method to get the correct cross section
'                           for both Bounded and Bounding members
'5/Dec/2011 - svsmylav TR-205302: Replaced '.Subport' method call with 'GetLateralSubPortBeforeTrim'.
'3/Feb/2012 - CM - TR-202532 : Modified CMNeedToCompute() for updating Corner Feature when necessary.
'                              Added IsCornerFeatureInputsModificatonNeeded() and MbrCornerFeatureDataByObject()
'                              methods, for determining CF modified ports if necessary
'9/Feb/2012 - GH - TR-210715 : Modified MigrateAssemblyConnection() to handle Split Migration When
'                               Bounding is splitted. Considered Splits, Miters and Axis Along cases.
'    11/Apr/2012 - svsmylav
'       DM-213229: In 'CrossSection_Flanges' method updated few section name strings to upper case.
''16/May/2012 -GH - TR-212435 : Added a new mehtod GetNearestBoundingToPort()
'    11/Oct/2012 - Alligators CR-207934/DM-220895(for v2013)
'                  (i)Set axis curve to member data in 'InitEndCutConnectionData_Stiffener'
'                  (ii) Added 'GetSectionType' method.
' 7/23/2013 - NK - Alligators TR-CP-234739 (for v2013)
'                  As part of this TR, we added a new method "IsRectangularMember"
' 26/Aug/2013 - skcheeka - TR-237880 :(ISPSFACInputHelper_ValidateObjects) Added proper checks to handle cases where the
'                                      ports in a collection gets reversed.
' 19/Mar/2014 - CM - TR-250554 : Fix for Access Violation errors. Checking if any object is nothing.
'                                Updated the project using Middle dlls but not Client Dlls
'                                Updated InItEndCutConnectionData() to get proper context port from
'                                connectable object and use it accrodingly to determine proper bounded obj.
'    11/Aug/2014 - knukala
'         CR-CP-250020: Create AC for ladder support or cage support to ladder rail
'         Added optional arguments for HasTopFlange() method and HasBottomFlange() method.
'        20/Aug/2014 - mchandak
'         CR-CP-240787: Added CreateBearingPlate ,GetACInputs ,SetMatlGradeThickness and SetPlatePartProperties methods
' 04/Nov/14 NK  CR-CP-262343  Create AC ladder rung penetrating support
'    03/Nov/2014 - MDT/GH
'         CR-CP-250198  Lapped AC for traffic items
' 24/Feb/2015 - MDT/GH
'         CR-265236    Added optional argument whether to flip the primary or secondary to method GetAssemblyConnectionInputs() and flip will be considered when true.
'   19/Feb/15   NK   CR-CP-267170  Mitered Stadnard AC to be considered even for Mbr End to End Offset case
' 23/Mar/2015 - MDT/RPK Added the "GetNearestboundingBUPort()" method for TR-269306,
'               when the bounding member is builtup/design member to get the Nearest bounding port.
' 21/July/2015 - GH  SI-CP-275688    Investigate split migration failure
'                Updated MigrateAssemblyConnection() and MigrateEndCutObject()
' 16/Jul/2015 - svsmylav TR-268173: Added a check in 'SetPlatePartProperties' method such that for creation of
'               new bearing plate naming rule is instantiated otherwise it is not instantiated to avoid asserts.
' 09/Sept/15   knukala   CR-CP-226692  GetAssemblyConnectionInputs() methos is modified inorder to handle stiffener cases also.
' 03/Nov/15 -  pkakula TR-CP-278336 Generic Assebly connection fails when bounding is split
'                Updated MigrateAssemblyConnection() and GetPointOnObject()
' 11/Dec/15 - PYK  Added GetFrameConnectionType, AreMembersIdentical from AssemblyConnCommon.bas file and Moved the constant E_FAIL to CommonEnumsAndConstants.bas file.
' 15/Feb/16 - NK   Few errors and warnings were reported when AssemblyConnectionTest ATP
'
' 11/Feb/2016 - svsmylav TR-287668 InitEndcutConnectionData method to exit if any of the two input
'                        connectables are of plate type and if so exit the method.
' 21/Apr/2016 - mkonduri
'                        TR-CP-292596: Asserts pop-up while placing AC on a SplitNone Frame Connection.
'*****************************************************************************
Option Explicit
Private Const MODULE = "StructDetail\Data\Include\MemberUtilities"
'
Public Const gsMbrAxisToCenter = "MbrAxis_ToCenter"
Public Const gsMbrAxisToEdgeAndEdge = "MbrAxis_ToEdgeAndEdge"
Public Const gsMbrAxisToEdgeAndOutSide1Edge = "MbrAxis_ToEdgeAndOutSide1Edge"
Public Const gsMbrAxisToEdgeAndOutSide2Edge = "MbrAxis_ToEdgeAndOutSide2Edge"
Public Const gsMbrAxisToEdge = "MbrAxis_ToEdge"
Public Const gsMbrAxisToFaceAndEdge = "MbrAxis_ToFaceAndEdge"
Public Const gsMbrAxisToFaceAndOutSide1Edge = "MbrAxis_ToFaceAndOutSide1Edge"
Public Const gsMbrAxisToFaceAndOutSideNoEdge = "MbrAxis_ToFaceAndOutSideNoEdge"
Public Const gsMbrAxisToOnMember = "MbrAxis_ToOnMember"
Public Const gsMbrAxisToOutSideAndOutSide1Edge = "MbrAxis_ToOutSideAndOutSide1Edge"
Public Const gsMbrAxisToOutSideAndOutSide2Edge = "MbrAxis_ToOutSideAndOutSide2Edge"
Public Const gsMbrAxisToOutSideAndOutSideNoEdge = "MbrAxis_ToOutSideAndOutSideNoEdge"

Public Const gsStiffEndToMbrCenter = "StiffEndToMbrFace_Center"
Public Const gsStiffEndToMbrEdgeAndEdge = "StiffEndToMbrFace_EdgeAndEdge"
Public Const gsStiffEndToMbrEdgeAndOutSide1Edge = "StiffEndToMbrFace_EdgeAndOS1Edge"
Public Const gsStiffEndToMbrEdgeAndOutSide2Edge = "StiffEndToMbrFace_EdgeAndOS2Edge"
Public Const gsStiffEndToMbrEdge = "StiffEndToMbrFace_Edge"
Public Const gsStiffEndToMbrFaceAndEdge = "StiffEndToMbrFace_FCAndEdge"
Public Const gsStiffEndToMbrFaceAndOutSide1Edge = "StiffEndToMbrFace_FCAndOS1Edge"
Public Const gsStiffEndToMbrFaceAndOutSideNoEdge = "StiffEndToMbrFace_FCAndOSNoEdge"
Public Const gsStiffEndToOnMember = "StiffEndToMbrFace_OnMember"
Public Const gsStiffEndToMbrOutSideAndOutSide1Edge = "StiffEndToMbrFace_OSAndOS1Edge"
Public Const gsStiffEndToMbrOutSideAndOutSide2Edge = "StiffEndToMbrFace_OSAndOS2Edge"
Public Const gsStiffEndToMbrOutSideAndOutSideNoEdge = "StiffEndToMbrFace_OSAndOSNoEdge"
'
Public Const eIdealized_Unk = "Unknown"
Public Const eIdealized_Top = "Top"
Public Const eIdealized_Bottom = "Bottom"
Public Const eIdealized_WebLeft = "Web_Left"
Public Const eIdealized_WebRight = "Web_Right"
Public Const eIdealized_EndBaseFace = "End_Base"
Public Const eIdealized_EndOffsetFace = "End_Offset"
Public Const eIdealized_BoundingTube = "Bounding_TubeType"
Const CosTheta = 0.99619 'where Theta = 5PI/180

Private m_ReplacedParts As Collection
Private m_ReplacingParts As Collection

Private m_MigratedFeaturesSize As Long
Private m_MigratedFeaturesCount As Long
Private m_MigratedFeatures() As FeatureMigrateData

'Structure of FeatureMigrateData
Private Type FeatureMigrateData
    Feature As Object
    ReplacedOpt As Long
    ReplacedOpr As Long
    ReplacedMember As Object
    ReplacedOperation As Object
    ReplacingOpt As Long
    ReplacingOpr As Long
    ReplacingMember As Object
    ReplacingOperation As Object
End Type

'

'*************************************************************************
'Function
'InitMemberConnectionData
'
'Abstract
'   Given the Assembly Connection (IJAppConnection Interface)
'   Initialize the BoundedConnection and Bounding Connection Data structures
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub InitMemberConnectionData(oAppConnection As IJAppConnection, _
                                    oBoundedData As MemberConnectionData, _
                                    oBoundingData As MemberConnectionData, _
                                    lStatus As Long, sMsg As String, Optional bConsiderFlipping As Boolean = True)
Const METHOD = "::InitMemberConnectionData"
    On Error GoTo ErrorHandler
    
    Dim iIndex As Long
    Dim lCount As Long
    
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    
    Dim oPoint As IJPoint
    Dim oPosition As IJDPosition
    Dim oElements_Ports As IJElements
    
    Dim oPort As IJPort
    Dim oPortObj As Object
    Dim oSplitAxisPort As ISPSSplitAxisPort
    
    Dim ePortId As SPSMemberAxisPortIndex
    
    sMsg = ""
    lStatus = 0
    If Not oAppConnection Is Nothing Then
        ' Get the Assembly Connection Ports from the IJAppConnection
        oAppConnection.enumPorts oElements_Ports
        lCount = oElements_Ports.Count
        
        ' for Member EndCuts, require two(2) Ports
        '   one Port will be PortId type of SPSMemberAxisAlong (Bounding Member)
        '   one Port will be PortId type of SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
        If lCount <> 2 Then
            sMsg = "Member Assembly Connection requires two(2) Ports"
            GoTo StatusFalse
        End If
        
        Dim oConnectedObject1 As IJPort
        Dim oConnectedObject2 As IJPort
        'EX: For box connections, whenever the axis end frame connection is on ladder rail, then flipping need to be handled in order to
        ' get the same endcuts when axis end FC is on ladder rung
        If bConsiderFlipping Then
            GetFlippedPorts oAppConnection, oConnectedObject1, oConnectedObject2
        Else
            Set oConnectedObject1 = oElements_Ports.Item(1)
            Set oConnectedObject2 = oElements_Ports.Item(2)
        End If
        'Check if Both are Axis Along Along -could be split none case
        If TypeOf oConnectedObject1 Is ISPSSplitAxisAlongPort And TypeOf oConnectedObject2 Is ISPSSplitAxisAlongPort Then
        
            InitAlongConnectionData_ForBothAlongPorts oConnectedObject1, oConnectedObject2, oAppConnection, _
                                            oBoundedData, oBoundingData, lStatus, sMsg
        Else
            InitEndCutConnectionData oConnectedObject1, oConnectedObject2, _
                                            oBoundedData, oBoundingData, lStatus, sMsg
        End If
    End If
    Exit Sub
    
StatusFalse:
    lStatus = 1
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    lStatus = E_FAIL
End Sub


'*********************************************************************************************
' Method      : GetMbrAssemblyConnectionType
' Description : Returns the Assembly connection type (Generic/Axis)
'
'*********************************************************************************************
Public Function GetMbrAssemblyConnectionType(oAppConnection As IJAppConnection) As eACType
    Const METHOD = "::GetMbrAssemblyConnectionType"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim oElements_Ports As IJElements
    Dim oPort As IJPort
    Dim oMemberPart1 As Object
    Dim oMemberPart2 As Object
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
  
    GetMbrAssemblyConnectionType = eACType.ACType_None
    
    If oAppConnection Is Nothing Then
        sMsg = "Invalid Argument passed : Argument passed is Nothing. Error Out"
        GoTo ErrorHandler
    End If
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    oAppConnection.enumPorts oElements_Ports

    Set oPort = oElements_Ports.Item(1)
    Set oMemberPart1 = oPort.Connectable
    
    Set oPort = oElements_Ports.Item(2)
    Set oMemberPart2 = oPort.Connectable
        
    'Compare member part 1 and member part 2
        
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    
    Dim bColinear As Boolean
    Dim bEndToEnd As Boolean
    Dim bRightAngle As Boolean
    
    If TypeOf oAppConnection Is IJAssemblyConnection Then
        Set oConnAttrbs = New SDConnectionUtils
        If GreaterThanOrEqualTo(oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count, 1) Then
            GetMbrAssemblyConnectionType = eACType.ACType_Stiff_Generic
        Else
            GetMbrAssemblyConnectionType = eACType.ACType_Bounded
        End If
    ElseIf Not (oMemberPart1 Is oMemberPart2) Then
        InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
        If lStatus <> 0 Then
            Exit Function
        End If
        CheckEndToEndConnection oBoundedData.MemberPart, oBoundingData.MemberPart, bEndToEnd, bColinear, bRightAngle
    
        If bEndToEnd Then
            If bColinear Then
                GetMbrAssemblyConnectionType = eACType.ACType_Split
            Else
                GetMbrAssemblyConnectionType = eACType.ACType_Miter
            End If
        Else
            GetMbrAssemblyConnectionType = eACType.ACType_Axis
        End If
    Else
        GetMbrAssemblyConnectionType = eACType.ACType_Mbr_Generic
    End If
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg

End Function

'*************************************************************************
'Function
'InitEndCutConnectionData
'
'Abstract
'   Given the Assembly Connection (IJAppConnection Interface)
'   Initialize the BoundedConnection and Bounding Connection Data structures
'
'input
'    oConnectionObject1  = Bounded Port
'    oConnectionObject2  = Bounding Port
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub InitEndCutConnectionData(oConnectionObject1 As Object, _
                                    oConnectionObject2 As Object, _
                                    oBoundedData As MemberConnectionData, _
                                    oBoundingData As MemberConnectionData, _
                                    lStatus As Long, sMsg As String)
Const METHOD = "::InitEndCutConnectionData"
    On Error GoTo ErrorHandler
    
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    Dim oStructPort1 As IJStructPort
    Dim oStructPort2 As IJStructPort
    Dim eStructPort1Context As eUSER_CTX_FLAGS
    Dim eStructPort2Context As eUSER_CTX_FLAGS
    Dim oSplitAxisPort1 As ISPSSplitAxisPort
    Dim oSplitAxisPort2 As ISPSSplitAxisPort
    Dim ePortId1 As SPSMemberAxisPortIndex
    Dim ePortId2 As SPSMemberAxisPortIndex
    
    'inittialize
    eStructPort1Context = CTX_INVALID
    eStructPort2Context = CTX_INVALID
    
    'From two objects paased as arguments, we will try to get its context id from Ports
    'through which we can determine which port obj is Bounded End Port
    If TypeOf oConnectionObject1 Is IJPort Then
       Set oPort1 = oConnectionObject1
       If TypeOf oPort1.Connectable Is IJProfile Or TypeOf oPort1.Connectable Is IJPlate Then
            Set oStructPort1 = oPort1
            If Not oStructPort1 Is Nothing Then eStructPort1Context = oStructPort1.ContextID
       ElseIf TypeOf oPort1.Connectable Is ISPSMemberPartCommon Then
            If TypeOf oPort1 Is ISPSSplitAxisPort Then
                Set oSplitAxisPort1 = oPort1
                ePortId1 = oSplitAxisPort1.PortIndex
                If ePortId1 = SPSMemberAxisAlong Then
                    eStructPort1Context = CTX_LATERAL
                ElseIf ePortId1 = SPSMemberAxisEnd Then
                    eStructPort1Context = CTX_OFFSET
                ElseIf ePortId1 = SPSMemberAxisStart Then
                    eStructPort1Context = CTX_BASE
                End If
            ElseIf TypeOf oPort1 Is IJStructPort Then
                Set oStructPort1 = oPort1
                eStructPort1Context = oStructPort1.ContextID
            Else
                'Unhandled Case
                'Unable to determine Port Context
            End If
       End If
    Else
       'Proper arguments not passed
       sMsg = "Proper Arguments not being passed....."
       GoTo ErrorHandler
    End If
    
    If TypeOf oConnectionObject2 Is IJPort Then
       Set oPort2 = oConnectionObject2
       If TypeOf oPort2.Connectable Is IJProfile Or TypeOf oPort2.Connectable Is IJPlate Then
            Set oStructPort2 = oPort2
            If Not oStructPort2 Is Nothing Then eStructPort2Context = oStructPort2.ContextID
       ElseIf TypeOf oPort2.Connectable Is ISPSMemberPartCommon Then
            If TypeOf oPort2 Is ISPSSplitAxisPort Then
                Set oSplitAxisPort2 = oPort2
                ePortId2 = oSplitAxisPort2.PortIndex
                If ePortId2 = SPSMemberAxisAlong Then
                    eStructPort2Context = CTX_LATERAL
                ElseIf ePortId2 = SPSMemberAxisEnd Then
                    eStructPort2Context = CTX_OFFSET
                ElseIf ePortId2 = SPSMemberAxisStart Then
                    eStructPort2Context = CTX_BASE
                End If
            ElseIf TypeOf oPort2 Is IJStructPort Then
                Set oStructPort2 = oPort2
                eStructPort2Context = oStructPort2.ContextID
            Else
                'Unhandled Case
                'Unable to determine Port Context
            End If
       End If
    Else
       'Proper arguments not passed
       sMsg = "Proper Arguments not being passed....."
       GoTo ErrorHandler
    End If
        
    
    If TypeOf oPort1.Connectable Is ISPSMemberPartCommon And _
           (eStructPort1Context = CTX_BASE Or eStructPort1Context = CTX_OFFSET) Then
        InitEndCutConnectionData_Mbr oConnectionObject1, _
                                oConnectionObject2, _
                                oBoundedData, _
                                oBoundingData, _
                                lStatus, sMsg
    ElseIf TypeOf oPort1.Connectable Is IJProfile And _
           (eStructPort1Context = CTX_BASE Or eStructPort1Context = CTX_OFFSET) Then
        InitEndCutConnectionData_Stiffener oConnectionObject1, _
                                oConnectionObject2, _
                                oBoundedData, _
                                oBoundingData, _
                                lStatus, sMsg
                                
    ElseIf TypeOf oPort2.Connectable Is IJProfile And _
           (eStructPort2Context = CTX_BASE Or eStructPort2Context = CTX_OFFSET) Then
        InitEndCutConnectionData_Stiffener oConnectionObject2, _
                                oConnectionObject1, _
                                oBoundedData, _
                                oBoundingData, _
                                lStatus, sMsg
    ElseIf TypeOf oPort2.Connectable Is ISPSMemberPartCommon And _
           (eStructPort2Context = CTX_BASE Or eStructPort2Context = CTX_OFFSET) Then
        InitEndCutConnectionData_Mbr oConnectionObject2, _
                                oConnectionObject1, _
                                oBoundedData, _
                                oBoundingData, _
                                lStatus, sMsg
    
    ElseIf (TypeOf oPort1.Connectable Is IJPlate) Or (TypeOf oPort2.Connectable Is IJPlate) Then
        'Yet to handle such cases
        Exit Sub
    
    Else
        'Other cases not handled currently
        sMsg = "cases not handled currently....."
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    lStatus = E_FAIL
End Sub

'*************************************************************************
'Function
'InitEndCutConnectionData_Mbr
'
'Abstract
'   Given the Assembly Connection (IJAppConnection Interface)
'   Initialize the BoundedConnection and Bounding Connection Data structures
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub InitEndCutConnectionData_Mbr(oConnectionObject1 As Object, _
                                    oConnectionObject2 As Object, _
                                    oBoundedData As MemberConnectionData, _
                                    oBoundingData As MemberConnectionData, _
                                    lStatus As Long, sMsg As String)
Const METHOD = "::InitEndCutConnectionData_Mbr"
    On Error GoTo ErrorHandler
    
    Dim iIndex As Long
    Dim lCount As Long
    
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    
    Dim bSetBoundedData As Boolean
    Dim bSetBoundingData As Boolean
    Dim bGenericBounding As Boolean
    
    Dim oPoint As IJPoint
    Dim oPosition As IJDPosition
    
    Dim oPort As IJPort
    Dim oPortObj As Object
    Dim oSplitAxisPort As ISPSSplitAxisPort
    
    Dim ePortId As SPSMemberAxisPortIndex
    
    Dim sCStype As String
    Dim oCrossSection As IJCrossSection
    Dim oBounded_CrossSection As ISPSCrossSection
    Dim oBounding_CrossSection As ISPSCrossSection
    Dim oBounded_PartDesigned As ISPSDesignedMember
    Dim oBounding_PartDesigned As ISPSDesignedMember
    Dim oBounded_PartPrismatic As ISPSMemberPartPrismatic
    Dim oBounding_PartPrismatic As ISPSMemberPartPrismatic
    
    Dim oBoundedPart As ISPSMemberPartCommon
    Dim oBoundingPart As ISPSMemberPartCommon
    
    sMsg = ""
    lStatus = 0
    bSetBoundedData = True
    bSetBoundingData = True
    bGenericBounding = False

    ' for Member EndCuts, require two(2) Ports
    '   one Port will be Bounding Object:
    '           PortId type of SPSMemberAxisAlong (Bounding MemberPart)
    '           Plate Base/Offset/Lateral Face Port
    '           Profile Base/Offset/ Lateral SubPort
    '           MemberPart Base/Offset/ Lateral SubPort
    '           Reference(Grid) Plane (IJPlane)
    '           Point (IJPoint)
    '   one Port will be PortId type of SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
    sMsg = "Checking Member Assembly Connection Ports"
    lCount = 2
    For iIndex = 1 To lCount
        If iIndex = 1 Then
            Set oPortObj = oConnectionObject1
        Else
            Set oPortObj = oConnectionObject2
        End If
        
        If oPortObj Is Nothing Then
        ElseIf TypeOf oPortObj Is ISPSSplitAxisPort Then
            Set oPort = oPortObj
            Set oSplitAxisPort = oPortObj
            ePortId = oSplitAxisPort.PortIndex
            
            If ePortId = SPSMemberAxisAlong Then
                If TypeOf oPortObj Is ISPSSplitAxisAlongPort Then
                    ' 1st AxisAlong is Bounding object
                    If bSetBoundingData Then
                        bSetBoundingData = False
                        oBoundingData.ePortId = ePortId
                        Set oBoundingData.AxisPort = oSplitAxisPort
                        If TypeOf oPort.Connectable Is ISPSMemberPartCommon Then
                            Set oBoundingPart = oPort.Connectable
                            Set oBoundingData.MemberPart = oBoundingPart
                        Else
                            Set oBoundingData.MemberPart = oPort.Connectable
                        End If
                    Else
                        ' Bounding Object already set: 2nd AxisAlong is Bounded
                        oBoundedData.ePortId = ePortId
                        Set oBoundedData.AxisPort = oSplitAxisPort
                        If TypeOf oPort.Connectable Is ISPSMemberPartCommon Then
                            Set oBoundedPart = oPort.Connectable
                            Set oBoundedData.MemberPart = oBoundedPart
                         Else
                            Set oBoundedData.MemberPart = oPort.Connectable
                        End If
                    End If
                End If
            
            ElseIf ePortId = SPSMemberAxisEnd Then
                If TypeOf oPortObj Is ISPSSplitAxisEndPort Then
                    ' 1st AxisEnd is Bounded object
                    If bSetBoundedData Then
                        bSetBoundedData = False
                        oBoundedData.ePortId = ePortId
                        Set oBoundedData.AxisPort = oSplitAxisPort
                        If TypeOf oPort.Connectable Is ISPSMemberPartCommon Then
                            Set oBoundedPart = oPort.Connectable
                            Set oBoundedData.MemberPart = oBoundedPart
                         Else
                            Set oBoundedData.MemberPart = oPort.Connectable
                        End If
                    Else
                        ' Bounded Object already set: 2nd AxisEnd is Bounding
                        oBoundingData.ePortId = ePortId
                        Set oBoundingData.AxisPort = oSplitAxisPort
                        If TypeOf oPort.Connectable Is ISPSMemberPartCommon Then
                            Set oBoundingPart = oPort.Connectable
                            Set oBoundingData.MemberPart = oBoundingPart
                        Else
                            Set oBoundingData.MemberPart = oPort.Connectable
                        End If
                    End If
                End If
                
            ElseIf ePortId = SPSMemberAxisStart Then
                If TypeOf oPortObj Is ISPSSplitAxisEndPort Then
                    ' 1st AxisStart is Bounded object
                    If bSetBoundedData Then
                        bSetBoundedData = False
                        oBoundedData.ePortId = ePortId
                        Set oBoundedData.AxisPort = oSplitAxisPort
                        If TypeOf oPort.Connectable Is ISPSMemberPartCommon Then
                            Set oBoundedPart = oPort.Connectable
                            Set oBoundedData.MemberPart = oBoundedPart
                         Else
                            Set oBoundedData.MemberPart = oPort.Connectable
                        End If
                    Else
                        ' Bounded Object already set: 2nd AxisStart is Bounding
                        oBoundingData.ePortId = ePortId
                        Set oBoundingData.AxisPort = oSplitAxisPort
                        If TypeOf oPort.Connectable Is ISPSMemberPartCommon Then
                            Set oBoundingPart = oPort.Connectable
                            Set oBoundingData.MemberPart = oBoundingPart
                        Else
                            Set oBoundingData.MemberPart = oPort.Connectable
                        End If
                    End If
                End If
            
            End If
            
            Set oSplitAxisPort = Nothing
        
        ElseIf iIndex = 2 Then
            ' Bounding Object is NOT Member Part Port
            bGenericBounding = True
            If TypeOf oPortObj Is IJPort Then
            ElseIf TypeOf oPortObj Is IJPlane Then
            ElseIf TypeOf oPortObj Is IJPoint Then
            ElseIf TypeOf oPortObj Is IJSurfaceBody Then
            Else
                bGenericBounding = False
            End If
        
        Else
            ' Bounded Object is NOT Member Part Port
        End If
        
        Set oPort = Nothing
        Set oPortObj = Nothing
    Next iIndex
    
    ' Verify have valid Ports from the Assembly Connection
    If oBoundedData.AxisPort Is Nothing Then
        lStatus = 10
        sMsg = "Member Assembly Connection Bounded Port could not be determined"
        GoTo StatusFalse
    
    ElseIf bGenericBounding Then
        ' Member Part bounded by Generic Port/Object
        ' (Bounding Object is NOT a MemberPart)
    
    ElseIf oBoundingData.AxisPort Is Nothing Then
        lStatus = 11
        sMsg = "Member Assembly Connection Bounding Port could not be determined"
        GoTo StatusFalse
    End If
        
    ' Verify Bounded is valid MemberPartPrismatic object
    sCStype = ""
    If oBoundedData.MemberPart.IsPrismatic Then
        ' If Bounded object is MemberPartPrismatic:
        Set oBounded_PartPrismatic = oBoundedData.MemberPart
        Set oBounded_CrossSection = oBounded_PartPrismatic.CrossSection
    
    ElseIf TypeOf oBoundedData.MemberPart Is ISPSDesignedMember Then
        ' If Bounded object is DesignedMember:
        Set oBounded_PartDesigned = oBoundedData.MemberPart
        Set oBounded_CrossSection = oBounded_PartDesigned
        
        lStatus = 12
        sMsg = "Bounded object is ISPSDesignedMember (NOT MemberPartPrismatic)"
        GoTo StatusFalse
    
    Else
        lStatus = 13
        sMsg = "Bounded object type is NOT Known"
        GoTo StatusFalse
    End If
    
    If oConnectionObject1 Is oConnectionObject2 Then
        bGenericBounding = True
    End If
    
    ' Verify Bounding have valid Cross Section Type
    If Not bGenericBounding Then
        If oBoundingData.MemberPart.IsPrismatic Then
            ' If Bounding object is MemberPartPrismatic:
            Set oBounding_PartPrismatic = oBoundingData.MemberPart
            Set oBounding_CrossSection = oBounding_PartPrismatic.CrossSection
        
        ElseIf TypeOf oBoundingData.MemberPart Is ISPSDesignedMember Then
            ' If Bounding object is DesignedMember:
            Set oBounding_PartDesigned = oBoundingData.MemberPart
            Set oBounding_CrossSection = oBounding_PartDesigned
        
            ' Check if all of the children are valid Detailed Parts
            If Not IsDesignedMemberDetailed(oBounding_PartDesigned) Then
                lStatus = 14
                sMsg = "Bounding object is NOT Detailed ISPSDesignedMember"
                GoTo StatusFalse
            End If
            
        Else
            ' If Bounding object is unKnown
            lStatus = 15
            sMsg = "Bounding object type is NOT Known"
            GoTo StatusFalse
        End If
    End If
    
    ' Verify Bounded have valid Cross Section Type
    If TypeOf oBounded_CrossSection.definition Is IJCrossSection Then
        Set oCrossSection = oBounded_CrossSection.definition
        sCStype = oCrossSection.Type
       
        sMsg = "Bounded CrossSection Type is not valid: " & sCStype
'''        If Trim(LCase(sCStype)) = LCase("CS") Then
'''            GoTo StatusFalse
'''        ElseIf Trim(LCase(sCStype)) = LCase("HSSC") Then
'''            GoTo StatusFalse
'''        ElseIf Trim(LCase(sCStype)) = LCase("PIPE") Then
'''            GoTo StatusFalse
'''        End If
        
    Else
        lStatus = 20
        sMsg = "Bounded CrossSection Definition is not valid"
        GoTo StatusFalse
    End If
    
    ' Get the Bounded Port Axis curve and Point on Axis Curve (x,y,z Location)
    sMsg = "Calculating Bounded Axis/Location"
    If TypeOf oBoundedData.AxisPort Is IJPort Then
        Set oPort = oBoundedData.AxisPort
        If TypeOf oPort.Geometry Is IJPoint Then
            Set oPoint = oPort.Geometry
            oPoint.GetPoint dx, dy, dz
                
            ' Matrix: U is direction Along Axis
            ' Matrix: V is direction normal to Web (from Web Right to Web Left)
            ' Matrix: W is direction normal to Flange (from Flange Bottom to Flange Top)
            Set oBoundedData.AxisCurve = GetAxisCurveAtPosition(dx, dy, dz, oBoundedData.MemberPart)
            oBoundedData.MemberPart.Rotation.GetTransformAtPosition dx, dy, dz, oBoundedData.Matrix, Nothing
        Else
            lStatus = 30
            sMsg = "Member Assembly Connection Bounded Port geometry does not support IJPoint interface"
            GoTo StatusFalse
        End If
    Else
        lStatus = 31
        sMsg = "Member Assembly Connection Bounded Port does not support IJPort interface"
        GoTo StatusFalse
    End If
        
    ' Get the Bounding Port Axis curve and Point on Axis Curve (x,y,z Location)
    If Not bGenericBounding Then
        sMsg = "Calculating Bounding Axis/Location"
        Set oPosition = GetConnectionPositionOnSupping(oBoundingData.MemberPart, _
                                                       oBoundedData.MemberPart, _
                                                       oBoundedData.ePortId)
        oPosition.Get dx, dy, dz
                
        ' Matrix: U is direction Along Axis
        ' Matrix: V is direction normal to Web (from Web Right to Web Left)
        ' Matrix: W is direction normal to Flange (from Flange Bottom to Flange Top)
        Set oBoundingData.AxisCurve = GetAxisCurveAtPosition(dx, dy, dz, oBoundingData.MemberPart)
        oBoundingData.MemberPart.Rotation.GetTransformAtPosition dx, dy, dz, oBoundingData.Matrix, Nothing
    End If
    
    ' Verify that Bounded and Bounding Axis Curves were determined (retrieved) and are valid type
    If oBoundedData.AxisCurve Is Nothing Then
        sMsg = "Member Assembly Connection Bounded Axis Curve is Nothing"
        lStatus = 40
        GoTo StatusFalse
    ElseIf (TypeOf oBoundedData.AxisCurve Is IJCurve) Then
    
    ElseIf Not TypeOf oBoundedData.AxisCurve Is IJLine Then
        lStatus = 41
        sMsg = "Member Assembly Connection Bounded Axis Curve does not support IJLine interface"
        GoTo StatusFalse
    
    ElseIf bGenericBounding Then
        ' Member Part bounded by Generic Port/Object
        ' (Bounding Object is NOT a MemberPart)
    
    ElseIf oBoundingData.AxisCurve Is Nothing Then
        lStatus = 42
        sMsg = "Member Assembly Connection Bounding Axis Curve is Nothing"
        GoTo StatusFalse
    ElseIf (TypeOf oBoundingData.AxisCurve Is IJCurve) Then
    
    ElseIf (Not TypeOf oBoundingData.AxisCurve Is IJLine) Then
        lStatus = 43
        sMsg = "Member Assembly Connection Bounding Axis Curve does not support IJLine interface"
        GoTo StatusFalse
    End If
    
    Exit Sub
    
StatusFalse:
    If lStatus = 0 Then
        lStatus = 1
    End If
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    lStatus = E_FAIL
End Sub
'*************************************************************************
'Function
'InitEndCutConnectionData_Stiffener
'
'Abstract
'   Given the Assembly Connection (IJAppConnection Interface)
'   Initialize the BoundedConnection and Bounding Connection Data structures
'
'input
'    oConnectionObject1  = Bounded Port
'    oConnectionObject2  = Bounding Port
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub InitEndCutConnectionData_Stiffener(oBoundedObject As Object, _
                                    oBoundingObject As Object, _
                                    oBoundedData As MemberConnectionData, _
                                    oBoundingData As MemberConnectionData, _
                                    lStatus As Long, sMsg As String)
Const METHOD = "::InitEndCutConnectionData_Stiffener"
    On Error GoTo ErrorHandler
    
       
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    If TypeOf oBoundedObject Is IJPort Then
        Set oBoundedPort = oBoundedObject
    Else
        GoTo ErrorHandler
    End If
    
    If TypeOf oBoundingObject Is IJPort Then
        Set oBoundingPort = oBoundingObject
    Else
        GoTo ErrorHandler
    End If
    
    Dim oBoundedStructPort As IJStructPort
    Dim eContext_Id As eUSER_CTX_FLAGS
    
    'This method needs to get called only if
    'BoundedPort.COnnectable is of Type Profile/Stiffener
    'Hence Bounded Port needs to even support IJStructPort
    If TypeOf oBoundedPort Is IJStructPort Then
        Set oBoundedStructPort = oBoundedPort
    Else
        '!!! Error Case
        GoTo ErrorHandler
    End If
    
    Dim oLandCurve As IJWireBody
    Dim oThicknessDir As IJDVector
    Dim bThicknessCentered As Boolean
    
    eContext_Id = oBoundedStructPort.ContextID
    
    Set oBoundedData.AxisPort = oBoundedPort
    Set oBoundedData.MemberPart = oBoundedPort.Connectable
    oBoundedData.ePortId = eContext_Id
    
    Set oLandCurve = GetProfilePartLandingCurve(oBoundedPort.Connectable)
    If Not oLandCurve Is Nothing Then Set oBoundedData.AxisCurve = oLandCurve
    Set oLandCurve = Nothing
        
    If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then
        Dim oBoundingObj As ISPSMemberPartCommon
        Dim oSplitAxisPort As ISPSSplitAxisPort
        
        'Set oSplitAxisPort = oBoundingPort
        Set oBoundingObj = oBoundingPort.Connectable
        
        Set oBoundingData.AxisPort = oBoundingPort
        Set oBoundingData.MemberPart = oBoundingObj
        oBoundingData.ePortId = CTX_LATERAL
        
    ElseIf TypeOf oBoundingPort Is ISPSSplitAxisPort Then
        'Already this case is covered
       
    ElseIf TypeOf oBoundingPort.Connectable Is IJProfile Then
        Dim oBoundingStiffObj As IJProfile
        
        Set oBoundingStiffObj = oBoundingPort.Connectable
        
        Set oBoundingData.AxisPort = oBoundingPort
        Set oBoundingData.MemberPart = oBoundingStiffObj
        oBoundingData.ePortId = CTX_LATERAL

        Set oLandCurve = GetProfilePartLandingCurve(oBoundingStiffObj)
        If Not oLandCurve Is Nothing Then Set oBoundingData.AxisCurve = oLandCurve
        Set oLandCurve = Nothing
    Else
        ''????? Error Case!!!!
    End If
    
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    lStatus = E_FAIL
End Sub

'*************************************************************************
'Function
'CheckEndToEndConnection
'
'Abstract
'   Given the Bounded and Bounding Connection Data
'   Determine if Connection is: End To End
'                                   if End To End,
'                                       Check if Colinear
'                                       Check if at Right Angles
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub CheckEndToEndConnection(oBoundedPart As Object, _
                                   oBoundingPart As Object, _
                                   bEndToEnd As Boolean, _
                                   bColinear As Boolean, _
                                   bRightAngle As Boolean)
Const METHOD = "::CheckEndToEndConnection"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim oPos As IJDPosition, oTangent1 As IJDVector, oTangent2 As IJDVector
    
    If (Not oBoundedPart Is Nothing) And (Not oBoundingPart Is Nothing) Then
        If (TypeOf oBoundedPart Is ISPSMemberPartCommon) And (TypeOf oBoundingPart Is ISPSMemberPartCommon) Then
    
            ' Check if Assembly Connection is End To End Type
            AreMembersEndConnected oBoundedPart, oBoundingPart, bEndToEnd, oPos, oTangent1, oTangent2
            
            If bEndToEnd Then
                Dim oLine1 As IJLine, oLine2 As IJLine
                Set oLine1 = New Line3d: Set oLine2 = New Line3d
                oLine1.SetRootPoint oPos.x, oPos.y, oPos.z
                oLine1.SetDirection oTangent1.x, oTangent1.y, oTangent1.z
                oLine1.Length = 1
                
                oLine2.SetRootPoint oPos.x, oPos.y, oPos.z
                oLine2.SetDirection oTangent2.x, oTangent2.y, oTangent2.z
                oLine2.Length = 1
                ' Assembly Connection is End To End Type
                ' Check if Axis are Colinear
                    bColinear = IsMemberAxesColinear(oLine1, oLine2)
                    If bColinear Then
                        ' Assembly Connection is End To End Colinear Type
                        bRightAngle = False
                    Else
                        ' Check if Axis are Normal to each Other
                        bRightAngle = IsMemberAxesAtRightAngles(oLine1, oLine2)
                        If bRightAngle Then
                            ' Assembly Connection is End To End Normal Type
                        Else
                            ' Assembly Connection is End To End Non-Linear Type
                        End If
                    End If
            Else
                ' Assembly Connection is Along Axis Type
                bColinear = False
                bRightAngle = False
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub



'Description: To determine if two members are end connected.
'if yes, the common position and tangents at the common position of the two memberparts will be returned.
'this is common for both linear and curved members.
Public Sub AreMembersEndConnected(oBoundedData As ISPSMemberPartCommon, oBoundingData As ISPSMemberPartCommon, _
     bEndToEnd As Boolean, oEndToEndPos As IJDPosition, oTangent1 As IJDVector, oTangent2 As IJDVector)
     
    Const METHOD = "AreMembersEndConnected"
    On Error GoTo ErrorHandler
     
    bEndToEnd = False
    If oBoundedData Is Nothing Or oBoundingData Is Nothing Then Exit Sub
    
    Dim oLine1 As IJLine, oLine2 As IJLine
    Dim oCurve1 As IJCurve, oCurve2 As IJCurve
    
    Dim oWBodyUtil As New GSCADShipGeomOps.SGOWireBodyUtilities
    Dim oStartPos1 As IJDPosition, oStartTan1 As IJDVector, oEndPos1 As IJDPosition, oEndTan1 As IJDVector
    Dim oStartPos2 As IJDPosition, oStartTan2 As IJDVector, oEndPos2 As IJDPosition, oEndTan2 As IJDVector
    
    Set oStartPos1 = New DPosition: Set oStartPos2 = New DPosition
    Set oEndPos1 = New DPosition: Set oEndPos2 = New DPosition
    
    Set oStartTan1 = New dVector: Set oStartTan2 = New dVector
    Set oEndTan1 = New dVector: Set oEndTan2 = New dVector
    
    oWBodyUtil.GetEndPointsAndTangents oBoundedData.Axis, oStartPos1, oStartTan1, oEndPos1, oEndTan1
    oWBodyUtil.GetEndPointsAndTangents oBoundingData.Axis, oStartPos2, oStartTan2, oEndPos2, oEndTan2
    
    Dim dx As Double, dy As Double, dz As Double
    Dim dist1 As Double
    
    dx = oStartPos1.x - oStartPos2.x
    dy = oStartPos1.y - oStartPos2.y
    dz = oStartPos1.z - oStartPos2.z
    dist1 = Sqr(dx * dx + dy * dy + dz * dz)
    
    If Abs(dist1) < distTol Then
        bEndToEnd = True
        Set oEndToEndPos = oStartPos1
        Set oTangent1 = oStartTan1
        Set oTangent2 = oStartTan2
    Else
        dx = oStartPos1.x - oEndPos2.x
        dy = oStartPos1.y - oEndPos2.y
        dz = oStartPos1.z - oEndPos2.z
        If Abs(Sqr(dx * dx + dy * dy + dz * dz)) < distTol Then
            bEndToEnd = True
            Set oEndToEndPos = oStartPos1
            Set oTangent1 = oStartTan1
            Set oTangent2 = oEndTan2
        End If
    End If
    
    dx = oEndPos1.x - oStartPos2.x
    dy = oEndPos1.y - oStartPos2.y
    dz = oEndPos1.z - oStartPos2.z
    dist1 = Sqr(dx * dx + dy * dy + dz * dz)
    If Abs(dist1) < distTol Then
        bEndToEnd = True
        Set oEndToEndPos = oStartPos2
        Set oTangent1 = oEndTan1
        Set oTangent2 = oStartTan2
    Else
        dx = oEndPos1.x - oEndPos2.x
        dy = oEndPos1.y - oEndPos2.y
        dz = oEndPos1.z - oEndPos2.z
        If Abs(Sqr(dx * dx + dy * dy + dz * dz)) < distTol Then
            bEndToEnd = True
            Set oEndToEndPos = oEndPos1
            Set oTangent1 = oEndTan1
            Set oTangent2 = oEndTan2
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
    
End Sub

'*************************************************************************
'Function
'AreCrossSectionsIdentical
'
'Abstract
'   Given the Bounded and Bounding Connection Data
'   Determine if Connection is: End To End
'                                   if End To End,
'                                       Check if Colinear
'                                       Check if at Right Angles
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub AreCrossSectionsIdentical(oBoundedData As MemberConnectionData, _
                                   oBoundingData As MemberConnectionData, _
                                   bIdentical As Boolean)
Const METHOD = "::AreCrossSectionsIdentical"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim sData1 As String
    Dim sData2 As String
    
    Dim oBounded_CrossSection As ISPSCrossSection
    Dim oBounded_PartDesigned As ISPSDesignedMember
    Dim oBounded_PartPrismatic As ISPSMemberPartPrismatic
    
    Dim oBounding_CrossSection As ISPSCrossSection
    Dim oBounding_PartDesigned As ISPSDesignedMember
    Dim oBounding_PartPrismatic As ISPSMemberPartPrismatic
    
    bIdentical = False
    If oBoundedData.MemberPart.IsPrismatic Then
        ' If Bounded object is MemberPartPrismatic:
        Set oBounded_PartPrismatic = oBoundedData.MemberPart
        Set oBounded_CrossSection = oBounded_PartPrismatic.CrossSection
    
    ElseIf TypeOf oBoundedData.MemberPart Is ISPSDesignedMember Then
        ' If Bounded object is DesignedMember:
        Set oBounded_PartDesigned = oBoundedData.MemberPart
        Set oBounded_CrossSection = oBounded_PartDesigned
    End If
    
    If oBoundingData.MemberPart.IsPrismatic Then
        ' If Bounded object is MemberPartPrismatic:
        Set oBounding_PartPrismatic = oBoundingData.MemberPart
        Set oBounding_CrossSection = oBounding_PartPrismatic.CrossSection
    
    ElseIf TypeOf oBoundingData.MemberPart Is ISPSDesignedMember Then
        ' If Bounded object is DesignedMember:
        Set oBounding_PartDesigned = oBoundingData.MemberPart
        Set oBounding_CrossSection = oBounding_PartDesigned
    End If
    
    sData1 = oBounded_CrossSection.sectionType
    sData2 = oBounding_CrossSection.sectionType
    If Trim(LCase(sData1)) = Trim(LCase(sData2)) Then
        sData1 = oBounded_CrossSection.SectionName
        sData2 = oBounding_CrossSection.SectionName
        If Trim(LCase(sData1)) = Trim(LCase(sData2)) Then
            sData1 = oBounded_CrossSection.SectionStandard
            sData2 = oBounding_CrossSection.SectionStandard
            If Trim(LCase(sData1)) = Trim(LCase(sData2)) Then
                bIdentical = True
            End If
        End If
    End If
    
'################################################################################
'################################################################################
If False Then
    sMsg = METHOD
    sMsg = sMsg & "   ...bIdentical = " & bIdentical & _
           vbCrLf & _
           "   ...Bounded Section Type :" & oBounded_CrossSection.sectionType & _
           "   ...Name:" & oBounded_CrossSection.SectionName & _
           "   ...Standard:" & oBounded_CrossSection.SectionStandard & _
           vbCrLf & _
           "   ...Bounding Section Type:" & oBounding_CrossSection.sectionType & _
           "   ...Name:" & oBounding_CrossSection.SectionName & _
           "   ...Standard:" & oBounding_CrossSection.SectionStandard
    zMsgBox sMsg
End If
'################################################################################
'################################################################################
    

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub

'*************************************************************************
'Function
'AlongAxis_CheckCopeCut
'
'Abstract
'   Given the Bounded and Bounding members, Determine:
'       1. The Idealized Boundary, Web_Left, Web_Right, Top, or Bottom
'
'
'input
'
'   sConfig = "Top_Top"     : Bounded/Bounding Top Flanges are in same general direction
'   sConfig = "Bottom_Top"  : Bounded/Bounding Top Flanges are in the opposite general direction
'   sConfig = "Left_Top"    : Bounded Web_Left/Bounding Top Flange are in same general direction
'   sConfig = "Right_Top"   : Bounded Web_Right/Bounding Top Flange are in same general direction
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub AlongAxis_CheckCopeCut(oBoundedData As MemberConnectionData, _
                                  oBoundingData As MemberConnectionData, _
                                  sIdealizedBoundary As String, _
                                  sConfig As String, _
                                  bUseCopeCut As Boolean, _
                                  Optional dTolerance As Double = 0.025)
Const METHOD = "::AlongAxis_CheckCopeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim oBoundedTopPos As IJDPosition
    Dim oBoundedBottomPos As IJDPosition
    Dim oTopFlangeBottomPos As IJDPosition
    Dim oBottomFlangeTopPos As IJDPosition
    
    Dim oBounded_WebFlangePoints As Collection
    Dim oBounding_WebFlangePoints As Collection
    
    bUseCopeCut = False
    ' Calculate the local 2D Points that defined Cross Section Flange and Web Sections
    ' to determine if Bounded Member is:
    '       Inside the Bounding Member Top Flange Bottom and Bottom Flange Top
    '       Outside the Bounding Member Top and Bottom
    '       Above Bounding Member Top and Above Bounding Member Bottom Flange Top
    '       Below Bounding Member Top Flange Bottom and Below Bounding Member Bottom
    '       ---------------- ......Top
    '       |              |
    '       ------    ------ ......Top Flange Bottom
    '             |  |
    '             |  |
    '             |  |
    '             |  |
    '             |  |
    '             |  |
    '       ------    ------ ......Bottom Flange Top
    '       |              |
    '       ---------------- ......Bottom
    '
    ' Convert 2D Points to 3D that represent the Bounding Web/Flange sections
    Set oBounding_WebFlangePoints = GetMemberWebFlangePoints(oBoundingData.MemberPart, _
                                                             oBoundingData.Matrix, _
                                                             sIdealizedBoundary)
    If oBounding_WebFlangePoints.Count < 4 Then
        Exit Sub
    End If
    
    ' Convert 2D Points to 3D to represent the Bounded Web/Flange sections
    Set oBounded_WebFlangePoints = GetMemberWebFlangePoints(oBoundedData.MemberPart, _
                                                            oBoundedData.Matrix)
    If oBounding_WebFlangePoints.Count < 4 Then
        Exit Sub
    End If
    
    ' Convert 3D Points to 2D that represent the Bounding Web/Flange sections
    ' in the Bounded Web Plane (Top Flange Bottom point and Bottom Flange Top point)
    BoundingPointToWebPlane oBounding_WebFlangePoints.Item(2), oBoundedData, oBoundingData, _
                            True, oTopFlangeBottomPos, True
    
    BoundingPointToWebPlane oBounding_WebFlangePoints.Item(3), oBoundedData, oBoundingData, _
                            True, oBottomFlangeTopPos, True
    
    ' Convert 3D Points to 2D that represent the Bounded Web/Flange sections
    ' in the Bounded Web Plane (Top Flange Top point and Bottom Flange Bottom point)
    BoundingPointToWebPlane oBounded_WebFlangePoints.Item(1), oBoundedData, oBoundingData, _
                            True, oBoundedTopPos, False
    
    BoundingPointToWebPlane oBounded_WebFlangePoints.Item(4), oBoundedData, oBoundingData, _
                            True, oBoundedBottomPos, False
    
If False Then
    zMsgBox vbCrLf & "MemberUtilities:" & METHOD & " 2D ...sConfig:" & sConfig

    Debug_Position "oBoundedTopPos      ", oBoundedTopPos
    Debug_Position "oBoundedBottomPos   ", oBoundedBottomPos
    
    Debug_Position "oTopFlangeBottomPos ", oTopFlangeBottomPos
    Debug_Position "oBottomFlangeTopPos ", oBottomFlangeTopPos
End If
    
    If Trim(LCase(sConfig)) = LCase("Top_Top") Then
        ' Supported and Supporting Configuration is "Top_Top"
        If oTopFlangeBottomPos.z > (oBoundedTopPos.z + dTolerance) Then
            If oBottomFlangeTopPos.z < (oBoundedBottomPos.z - dTolerance) Then
                ' Supported Member is completely inside the Supporting Flanges
                ' Web Cope Cut is not required as default
                Exit Sub
            End If
        End If
    Else
        ' Supported and Supporting Configuration is "Bottom_Top"
        If oBottomFlangeTopPos.z > (oBoundedTopPos.z + dTolerance) Then
            If oBoundedTopPos.z < (oBoundedBottomPos.z - dTolerance) Then
                ' Supported Member is completely inside the Supporting Flanges
                ' Web Cope Cut is not required as default
                Exit Sub
            End If
        End If
    End If
    
    ' Due to Supporting Member Flange(s),
    ' Supported Member requires a Web Cope Cut around the the Supporting Flange(s)
    bUseCopeCut = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub

'*************************************************************************
'Function
'AlongAxis_IdealizedFlange
'
'Abstract
'   Given the Bounded and Bounding members when the Idealized Boundary is Top/Bottom Flange
'   Determine:
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub AlongAxis_IdealizedFlange(oBoundedData As MemberConnectionData, _
                                     oBoundingData As MemberConnectionData, _
                                     sConfig As String)
Const METHOD = "::AlongAxis_IdealizedFlange"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub

'*************************************************************************
'Function
'EndCut_WebFlangeConfig
'
'Abstract
'   Given the Bounded and Bounding members
'   Determine:
'       1. If the Bounded and Bounding Up direction vectors are in same general direction
'       2. If the Bounded and Bounding Up direction vectors are in opposite general direction
'       3. If the Bounded Up direction is same general Direction as Bounding Web_Left Direction
'       4. If the Bounded Up direction is same general Direction as Bounding Web_Right Direction
'
'input
'
'Return
'   sConfig = "Top_Top"     : Bounded/Bounding Top Flanges are in same general direction
'   sConfig = "Bottom_Top"  : Bounded/Bounding Top Flanges are in the opposite general direction
'   sConfig = "Left_Top"    : Bounded Web_Left/Bounding Top Flange are in same general direction
'   sConfig = "Right_Top"   : Bounded Web_Right/Bounding Top Flange are in same general direction
'
'Exceptions
'
'***************************************************************************
Public Sub EndCut_WebFlangeConfig(oBoundedData As MemberConnectionData, _
                                  oBoundingData As MemberConnectionData, _
                                  sConfig As String)
Const METHOD = "::EndCut_WebFlangeConfig"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dDot_UpUp As Double
    Dim dDot_WebUp As Double
    Dim dDistTolerance As Double
    
    Dim oBounded_UpDir As IJDVector
    Dim oBounded_WebDir As IJDVector
    Dim oBounding_UpDir As IJDVector
    
    Dim oGeometryServices As IGeometryServices
    
    Set oGeometryServices = New IngrGeom3D.GeometryFactory
    dDistTolerance = oGeometryServices.DistTolerance
    Set oGeometryServices = Nothing
    
    sConfig = ""
    ' Matrix.IndexValue(0,1,2)   : U is direction Along Axis
    ' Matrix.IndexValue(4,5,6)   : V is direction normal to Web (from Web Right to Web Left)
    ' Matrix.IndexValue(8,9,10)  : W is direction normal to Flange (from Flange Bottom to Flange Top)
    ' Matrix.IndexValue(12,13,14): Root/Origin Point
    '
    ' Calculate Dot Product between the Bounded and Bounding Up direction vectors
    Set oBounded_UpDir = New dVector
    oBounded_UpDir.Set oBoundedData.Matrix.IndexValue(8), oBoundedData.Matrix.IndexValue(9), _
                       oBoundedData.Matrix.IndexValue(10)
                       
    Set oBounding_UpDir = New dVector
    oBounding_UpDir.Set oBoundingData.Matrix.IndexValue(8), oBoundingData.Matrix.IndexValue(9), _
                        oBoundingData.Matrix.IndexValue(10)
    
    dDot_UpUp = oBounded_UpDir.Dot(oBounding_UpDir)

    ' Calculate Dot Product between the Bounded Up direction vector and Bounding AlongAxis vector
    Set oBounded_WebDir = New dVector
    oBounded_WebDir.Set oBoundedData.Matrix.IndexValue(4), oBoundedData.Matrix.IndexValue(5), _
                        oBoundedData.Matrix.IndexValue(6)
    
    dDot_WebUp = oBounded_WebDir.Dot(oBounding_UpDir)
        
    ' Check if the Up/Up directions are better choice then the Web/Up directions
    If Abs(dDot_UpUp) < dDistTolerance And Abs(dDot_WebUp) < dDistTolerance Then
    
    ElseIf Abs(dDot_UpUp) >= Abs(dDot_WebUp) Then
        ' The Bounded and Bounding Up direction vectors are in same general direction
        ' check if Top Flanges are in same general direction
        If dDot_UpUp < 0# Then
            ' Bounded/Bounding Top Flanges are in the opposite general direction
            sConfig = "Bottom_Top"
        Else
            ' Bounded/Bounding Top Flanges are in same general direction
            sConfig = "Top_Top"
        End If
        
    Else
        ' The Bounded Up direction and Bounding Web direction vectors are in same general direction
        ' check if Bounded Web and Bounding Top Flange are in same general direction
        If dDot_WebUp < 0# Then
            ' Bounded Web_Right/Bounding Top Flange are in the same general direction
            sConfig = "Right_Top"
        Else
            ' Bounded Web_Left/Bounding Top Flange are in the same general direction
            sConfig = "Left_Top"
        End If
    End If
    
If False Then
    zMsgBox METHOD & "   ...sConfig:" & sConfig
'''    Debug_Matrix "oBounded_Matrix ", oBoundedData.Matrix
'''    Debug_Matrix "oBounding_Matrix", oBoundingData.Matrix
'''    Debug_Vector "oBounded_UpDir  ", oBounded_UpDir
'''    Debug_Vector "oBounded_WebDir ", oBounded_WebDir
'''    Debug_Vector "oBounding_UpDir ", oBounding_UpDir
'''    zMsgBox "dDot_UpUp:" & Format(dDot_UpUp, "0.0000") & _
'''            "   ...dDot_WebUp:" & Format(dDot_WebUp, "0.0000")
End If

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub

'*************************************************************************
'Function
'GetMemberWebFlangePoints
'
'Abstract
'   Given the Bounded and Bounding members when the Idealized Boundary is Web_Left/Web_Right
'   Determine: Points that locate the Cross Section Flange sections
'   The Points are based on Top/Bottom/Web_Left/Web_Right Edges
'   (all Cross Sections should have these Edges if Flanges are defined)
'       ---------------- ......Top
'       |              |
'       ------    ------ ......Top Flange Bottom
'             |  |
'             |  |
'             |  |
'             |  |
'             |  |
'             |  |
'       ------    ------ ......Bottom Flange Top
'       |              |
'       ---------------- ......Bottom
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function GetMemberWebFlangePoints(oMemberPart As ISPSMemberPartCommon, _
                                         Optional oMatrix As IJDT4x4 = Nothing, _
                                         Optional sWebFace As String = eIdealized_Unk) _
                                         As Collection
Const METHOD = "::GetMemberWebFlangePoints"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dWeb_X1 As Double
    Dim dWeb_Y1 As Double
    Dim dWeb_X2 As Double
    Dim dWeb_Y2 As Double
    
    Dim oPos11 As IJDPosition
    Dim oPos12 As IJDPosition
    Dim oPos21 As IJDPosition
    Dim oPos22 As IJDPosition
    Dim oPos31 As IJDPosition
    Dim oPos32 As IJDPosition
    Dim oPos41 As IJDPosition
    Dim oPos42 As IJDPosition
    Dim oPosAlongAxis As IJDPosition
    Dim oWebFlangePoint As IJDPosition
    
    Dim oSymbol As IJDSymbol
    Dim oWebLeft As IJWireBody
    Dim oWebRight As IJWireBody
    Dim oFlangeTop As IJWireBody
    Dim oFlangeBottom As IJWireBody
    
    Dim oCrossSection As ISPSCrossSection
    Dim oPartDesigned As ISPSDesignedMember
    Dim oPartPrismatic As ISPSMemberPartPrismatic
    Dim oCollectionPoints As Collection
    Dim oBounded_PortFaces As IJElements
    
    Set oCollectionPoints = New Collection
    If oMemberPart.IsPrismatic Then
        ' Member Part is ISPSMemberPartPrismatic
        Set oPartPrismatic = oMemberPart
        Set oCrossSection = oPartPrismatic.CrossSection
        
    ElseIf TypeOf oMemberPart Is ISPSDesignedMember Then
        ' Member Part is ISPSDesignedMember
        Set oPartDesigned = oMemberPart
        Set oCrossSection = oPartDesigned
        
'$$$Klidge
'       ISPDesignedMember does not currently support the IJDSymbol
'       for now, return empty collection
'       future: need to determine how to determine shape points for ISPSDesignedmember objects
        Set GetMemberWebFlangePoints = oCollectionPoints
        Exit Function

    Else
        ' Member Part type is unKnown
        Set GetMemberWebFlangePoints = oCollectionPoints
        Exit Function
    End If
    
    If TypeOf oCrossSection.symbol Is IJDSymbol Then
        ' Get the Edges from the current symbol in the following order:
        ' JXSEC_WEB_LEFT : JXSEC_WEB_RIGHT : JXSEC_TOP : JXSEC_BOTTOM
        Set oSymbol = oCrossSection.symbol
        Set oBounded_PortFaces = GetSymbolWebFlangeEdges(oSymbol)
        If oBounded_PortFaces Is Nothing Then
            ' Did Not Find the requested Edges, assume there is No Web or Flange
        ElseIf oBounded_PortFaces.Count < 4 Then
            ' Did Not Find ALL of the requested Edges, assume there is No Web or Flange
        Else
            ' Set up data for converting 2D Points to 3D points
            If Not oMatrix Is Nothing Then
                Set oPosAlongAxis = New DPosition
                oPosAlongAxis.Set oMatrix.IndexValue(12), _
                                  oMatrix.IndexValue(13), oMatrix.IndexValue(14)
            End If

            ' Extract the End Points for Each Wire
            Set oWebLeft = oBounded_PortFaces.Item(1)
            oWebLeft.GetEndPoints oPos11, oPos12
            
            Set oWebRight = oBounded_PortFaces.Item(2)
            oWebRight.GetEndPoints oPos21, oPos22
        
            Set oFlangeTop = oBounded_PortFaces.Item(3)
            oFlangeTop.GetEndPoints oPos31, oPos32
            
            Set oFlangeBottom = oBounded_PortFaces.Item(4)
            oFlangeBottom.GetEndPoints oPos41, oPos42
            
If False Then
    zMsgBox vbCrLf & "MemberUtilities:" & METHOD & " 2D"

    Debug_Position "oWebLeft      oPos11      ", oPos11
    Debug_Position "oWebLeft      oPos12      ", oPos12
    Debug_Position "oWebRight     oPos21      ", oPos21
    Debug_Position "oWebRight     oPos22      ", oPos22
    Debug_Position "oFlangeTop    oPos31      ", oPos31
    Debug_Position "oFlangeTop    oPos32      ", oPos32
    Debug_Position "oFlangeBottom oPos41      ", oPos41
    Debug_Position "oFlangeBottom oPos42      ", oPos42
End If
        
            ' Determine Local u (X position) based on Idealized Boundary given
            If sWebFace = eIdealized_WebLeft Then
                dWeb_X1 = oPos11.x
                dWeb_X2 = oPos12.x
            ElseIf sWebFace = eIdealized_WebRight Then
                dWeb_X1 = oPos21.x
                dWeb_X2 = oPos22.x
            Else
                dWeb_X1 = (oPos11.x + oPos21.x) / 2#
                dWeb_X2 = (oPos12.x + oPos22.x) / 2#
            End If
        
            ' set the Top Flange Bottom position (local v) value:
            ' set the Bottom Flange Top position (local v) value:
            ' set the Top Flange Bottom position / Bottom Flange Top position :
            '   (local v) value using the end points of "Web_Left" or "Web_Right" Edge
            If sWebFace = eIdealized_WebLeft Then
                ' calculate distance from Top Edge to Web_Left end points
                If Abs(oPos11.y - oPos31.y) > Abs(oPos12.y - oPos31.y) Then
                    ' Web_Left end point is closest to Top
                    ' Web_Left start point is closest to Bottom
                    dWeb_Y1 = oPos12.y
                    dWeb_Y2 = oPos11.y
                Else
                    ' Web_Left start point is closest to Top
                    ' Web_Left start point is closest to Bottom
                    dWeb_Y1 = oPos11.y
                    dWeb_Y2 = oPos12.y
                End If
            ElseIf sWebFace = eIdealized_WebRight Then
                ' calculate distance from Top Edge to Web_Right end points
                If Abs(oPos21.y - oPos31.y) > Abs(oPos22.y - oPos31.y) Then
                    ' Web_Right end point is closest to Top
                    ' Web_Right start point is closest to Bottom
                    dWeb_Y1 = oPos22.y
                    dWeb_Y2 = oPos21.y
                Else
                    ' Web_Right start point is closest to Top
                    ' Web_Right end point is closest to Bottom
                    dWeb_Y1 = oPos21.y
                    dWeb_Y2 = oPos22.y
                End If
            Else
                ' calculate distance from Top Edge to Web_Left end points
                If Abs(oPos11.y - oPos31.y) > Abs(oPos12.y - oPos31.y) Then
                    ' Web_Left end point is closest to Top
                    ' Web_Left start point is closest to Bottom
                    dWeb_Y1 = oPos12.y
                    dWeb_Y2 = oPos11.y
                Else
                    ' Web_Left start point is closest to Top
                    ' Web_Left end point is closest to Bottom
                    dWeb_Y1 = oPos11.y
                    dWeb_Y2 = oPos12.y
                End If
                
                ' calculate distance from Top Edge to Web_Right end points
                If Abs(oPos21.y - oPos31.y) > Abs(oPos22.y - oPos31.y) Then
                    ' Web_Right end point is closest to Top
                    If Abs(dWeb_Y1 - oPos31.y) > Abs(oPos22.y - oPos31.y) Then
                        ' Web_Left Edge point is farthest from the Top
                    Else
                        ' Web_Right Edge point is farthest from the Top
                        dWeb_Y1 = oPos22.y
                    End If
                
                    If Abs(dWeb_Y2 - oPos41.y) > Abs(oPos21.y - oPos41.y) Then
                        ' Web_Left Edge point is farthest from the Bottom
                    Else
                        ' Web_Right Edge point is farthest from the Bottom
                        dWeb_Y2 = oPos21.y
                    End If
                
                Else
                    ' Web_Right start point is closest to Top
                    If Abs(dWeb_Y1 - oPos31.y) > Abs(oPos21.y - oPos31.y) Then
                        ' Web_Left Edge point is farthest from the Top
                    Else
                        ' Web_Right Edge point is farthest from the Top
                        dWeb_Y1 = oPos21.y
                    End If
                
                    If Abs(dWeb_Y2 - oPos41.y) > Abs(oPos22.y - oPos41.y) Then
                        ' Web_Left Edge point is farthest from the Bottom
                    Else
                        ' Web_Right Edge point is farthest from the Bottom
                        dWeb_Y2 = oPos22.y
                    End If
                End If
            End If
            
            ' set the Top of the Top Flange (local v) value using the "Top" Edge
            Set oWebFlangePoint = New DPosition
            If oPos31.y > oPos32.y Then
                oWebFlangePoint.Set dWeb_X1, oPos31.y, 0#
            Else
                oWebFlangePoint.Set dWeb_X1, oPos32.y, 0#
            End If
            If Not oMatrix Is Nothing Then
                Set oWebFlangePoint = z_TransformPosToGlobal(oMemberPart, _
                                                           oWebFlangePoint, oPosAlongAxis)
            End If
            oCollectionPoints.Add oWebFlangePoint
            
            ' set the Bottom of the Top Flange (local v) value using "Web_Left"/"Web_Right" Edge
            Set oWebFlangePoint = New DPosition
            oWebFlangePoint.Set dWeb_X1, dWeb_Y1, 0#
            If Not oMatrix Is Nothing Then
                Set oWebFlangePoint = z_TransformPosToGlobal(oMemberPart, _
                                                           oWebFlangePoint, oPosAlongAxis)
            End If
            oCollectionPoints.Add oWebFlangePoint
            
            ' set the Top of the Bottom Flange (local v) value using "Web_Left"/"Web_Right" Edge
            Set oWebFlangePoint = New DPosition
            oWebFlangePoint.Set dWeb_X1, dWeb_Y2, 0#
            If Not oMatrix Is Nothing Then
                Set oWebFlangePoint = z_TransformPosToGlobal(oMemberPart, _
                                                           oWebFlangePoint, oPosAlongAxis)
            End If
            oCollectionPoints.Add oWebFlangePoint
        
            ' set the Bottom of the Bottom Flange (local v) value using the "Bottom" Edge
            Set oWebFlangePoint = New DPosition
            If oPos41.y > oPos42.y Then
                oWebFlangePoint.Set dWeb_X1, oPos42.y, 0#
            Else
                oWebFlangePoint.Set dWeb_X1, oPos41.y, 0#
            End If
            If Not oMatrix Is Nothing Then
                Set oWebFlangePoint = z_TransformPosToGlobal(oMemberPart, _
                                                           oWebFlangePoint, oPosAlongAxis)
            End If
            oCollectionPoints.Add oWebFlangePoint
            
If False Then
    If Not oMatrix Is Nothing Then
        zMsgBox vbCrLf & "MemberUtilities:" & METHOD & " 3D"
    Else
        zMsgBox vbCrLf & "MemberUtilities:" & METHOD & " 2D"
    End If

    Debug_Position "Flange_Top_Top      ", oCollectionPoints.Item(1)
    Debug_Position "Flange_Top_Bottom   ", oCollectionPoints.Item(2)
    Debug_Position "Flange_Bottom_Top   ", oCollectionPoints.Item(3)
    Debug_Position "Flange_Bottom_Bottom", oCollectionPoints.Item(4)
End If
        
        End If
    End If
    
    Set GetMemberWebFlangePoints = oCollectionPoints
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'GetSymbolWebFlangeEdges
'
'Abstract
'   Given the CrossSection Symbol,
'   Retreive the Edges (WireBody) that represent Web_left, Web_Right, Bottom, Top
'   Determine:
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function GetSymbolWebFlangeEdges(oSymbol As IJDSymbol) As IJElements
Const METHOD = "::GetSymbolWebFlangeEdges"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim oCollectionEdgeIds As Collection
    
    ' Return the list of Edges
    Set oCollectionEdgeIds = New Collection
    oCollectionEdgeIds.Add JXSEC_WEB_LEFT   ' Edge Id:   = 257
    oCollectionEdgeIds.Add JXSEC_WEB_RIGHT  ' Edge Id:   = 258
    oCollectionEdgeIds.Add JXSEC_TOP        ' Edge Id:   = 514
    oCollectionEdgeIds.Add JXSEC_BOTTOM     ' Edge Id:   = 513
    
    Set GetSymbolWebFlangeEdges = GetSymbolEdgeIds(oSymbol, oCollectionEdgeIds)
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'EndCut_IdealizedWebFlanges
'
'Abstract
'   Given the CrossSection Symbol and Idealized Boundary
'   Retreive the Edges (WireBody) that represent Web_left, Web_Right, Bottom, Top
'   Determine:
'       If the Cross Section has Top_Flange (Left or Right) section
'       If the Cross Section has Bottom_Flange (Left or Right) section
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub EndCut_IdealizedWebFlanges(oMemberPart As ISPSMemberPartCommon, _
                                      sIdealizedBoundary As String, _
                                      bTopFlange As Boolean, bBottomFlange As Boolean)
Const METHOD = "::EndCut_IdealizedWebFlanges"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim bWebLeft As Boolean
    Dim bWebRight As Boolean
    
    Dim oSymbol As IJDSymbol
    Dim oPartDesigned As ISPSDesignedMember
    Dim oPartPrismatic As ISPSMemberPartPrismatic
    Dim oSPS_CrossSection As ISPSCrossSection

    Dim oCollectionFlange As IJElements
    Dim oCollectionEdgeIds As Collection
    
    bTopFlange = False
    bBottomFlange = False
    If oMemberPart.IsPrismatic Then
        Set oPartPrismatic = oMemberPart
        Set oSPS_CrossSection = oPartPrismatic.CrossSection
        
    ElseIf TypeOf oMemberPart Is ISPSDesignedMember Then
        Set oPartDesigned = oMemberPart
        Set oSPS_CrossSection = oMemberPart
'$$$Klidge
'       ISPDesignedMember does not currently support the IJDSymbol
'       for now, just exit sub
'       future:
'       need to determine how to determine Edge Ids for ISPSDesignedmember objects
        Exit Sub
    Else
        Exit Sub
    End If
    
    If Not TypeOf oSPS_CrossSection.symbol Is IJDSymbol Then
        Exit Sub
    End If
    
    bWebLeft = False
    bWebRight = False
    If Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_Unk) Then
        bWebLeft = True
        bWebRight = True
    ElseIf Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_WebLeft) Then
        bWebLeft = True
    ElseIf Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_WebRight) Then
        bWebRight = True
    End If
    
    ' The code below is a very general way to determine if a Cross Section contain Flange sections
    ' It is more costly w.r.t. Performance then basing the answer on Cross Section "Type" property
    ' but it is also more gneral, Need to determine if the Performance cost is acceptable
    Set oSymbol = oSPS_CrossSection.symbol
    If bWebLeft Then
        ' Check if any Bottom Flange "Left" Edges exist
        Set oCollectionEdgeIds = New Collection
        oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_LEFT_TOP             ' Edge Id:   = 769
        oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_LEFT                 ' Edge Id:   = 1025
        oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_LEFT_TOP_CORNER      ' Edge Id:   = 1541
        oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM_CORNER   ' Edge Id:   = 1796
        oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM          ' Edge Id:   = 2049
        Set oCollectionFlange = GetSymbolEdgeIds(oSymbol, oCollectionEdgeIds, True)
        If oCollectionFlange.Count > 0 Then
            bBottomFlange = True
        End If
        
        Set oCollectionFlange = Nothing
        Set oCollectionEdgeIds = Nothing
        
        ' Check if any Top Flange "Left" Edges exist
        Set oCollectionEdgeIds = New Collection
        oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_LEFT_BOTTOM         ' Edge Id:   = 770
        oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_LEFT                ' Edge Id:   = 1026
        oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER  ' Edge Id:   = 1540
        oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_LEFT_TOP_CORNER     ' Edge Id:   = 1795
        oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_LEFT_TOP            ' Edge Id:   = 2050
        Set oCollectionFlange = GetSymbolEdgeIds(oSymbol, oCollectionEdgeIds, True)
        If oCollectionFlange.Count > 0 Then
            bTopFlange = True
        End If
        
        Set oCollectionFlange = Nothing
        Set oCollectionEdgeIds = Nothing
        
    End If
    
    If bWebRight Then
        If Not bBottomFlange Then
            ' Check if any Bottom Flange "Right" Edges exist
            Set oCollectionEdgeIds = New Collection
            oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_RIGHT_TOP            ' Edge Id:   = 771
            oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_RIGHT                ' Edge Id:   = 1027
            oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER     ' Edge Id:   = 1538
            oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM_CORNER  ' Edge Id:   = 1793
            oCollectionEdgeIds.Add JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM         ' Edge Id:   = 2051
            Set oCollectionFlange = GetSymbolEdgeIds(oSymbol, oCollectionEdgeIds, True)
            If oCollectionFlange.Count > 0 Then
                bBottomFlange = True
            End If
            
            Set oCollectionFlange = Nothing
            Set oCollectionEdgeIds = Nothing
        End If
    
        If Not bTopFlange Then
            Set oCollectionEdgeIds = New Collection
            oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_RIGHT_BOTTOM        ' Edge Id:   = 772
            oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_RIGHT               ' Edge Id:   = 1028
            oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER ' Edge Id:   = 1537
            oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER    ' Edge Id:   = 1794
            oCollectionEdgeIds.Add JXSEC_TOP_FLANGE_RIGHT_TOP           ' Edge Id:   = 2052
            Set oCollectionFlange = GetSymbolEdgeIds(oSymbol, oCollectionEdgeIds, True)
            If oCollectionFlange.Count > 0 Then
                bTopFlange = True
            End If
            
            Set oCollectionFlange = Nothing
            Set oCollectionEdgeIds = Nothing
        End If
    End If
    
If False Then
zMsgBox METHOD & " ...bTopFlange:" & bTopFlange & " ...bBottomFlange:" & bBottomFlange
End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'GetSymbolEdgeIds
'
'Abstract
'   Given the CrossSection Symbol,
'   Retreive the Edges (WireBody) that represent given list of Edge Ids
'   Determine:
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function GetSymbolEdgeIds(oSymbol As IJDSymbol, oCollectionEdgeIds As Collection, _
                                 Optional bReturnFirst As Boolean = False) As IJElements
Const METHOD = "::GetSymbolEdgeIds"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim jIndex As Long
    Dim kIndex As Long
    Dim lEdgeId As Long
    
    Dim sText As String
    Dim sRepName As String
    Dim sOutputName As String
    
    Dim oOutput As Object
    Dim oSubOutput As IMSSymbolEntities.IJDOutput
    Dim oSubOutputs As IMSSymbolEntities.IJDOutputs
    
    Dim oSymbolOutput As IMSSymbolEntities.IJDOutput
    Dim oSymbolOutputs As IMSSymbolEntities.IJDOutputs
    Dim oRepresentation As IMSSymbolEntities.IJDRepresentation
    Dim oRepresentations As IMSSymbolEntities.IJDRepresentations
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim oCollOfOutputElements As IJElements
    
    Dim oEdgeWire As IJWireBody
    Set oCollOfOutputElements = New JObjectCollection
    
    ' Retreive the list of Representations from the Cross Section Symbol Definition
    Set oSymbolDefinition = oSymbol.IJDSymbolDefinition(0)
    Set oRepresentations = oSymbolDefinition.IJDRepresentations
sText = vbCrLf & "MemberUtilities:" & METHOD
sText = sText & vbCrLf & "Representations.Count =" & Trim(Str(oRepresentations.Count))

    For iIndex = 1 To oRepresentations.Count
        Set oRepresentation = oRepresentations.Item(iIndex)
        sRepName = oRepresentation.Name
        Set oRepresentation = Nothing
sText = sText & vbCrLf & "(" & Trim(Str(iIndex)) & ") sRepName: " & sRepName
    Next iIndex
    
    ' Loop the the list of Representations searching for "DetailedPhysicalWireBody"
    For iIndex = 1 To oRepresentations.Count
        Set oRepresentation = oRepresentations.Item(iIndex)
        sRepName = oRepresentation.Name
        If Trim(LCase(sRepName)) = LCase("DetailedPhysicalWireBody") Then
sText = sText & vbCrLf & "(" & Trim(Str(iIndex)) & ") sRepName: " & sRepName
            
            ' Retreive the list of Symbol Outputs from the Representation
            Set oRepresentation = oRepresentations.GetRepresentationByName(sRepName)
            Set oSymbolOutputs = oRepresentation
        
            ' Loop thru the list of Symbol Outputs for the Attributed WireBody
            ' (expect only one Output)
            On Error Resume Next
            jIndex = 0
            For Each oSymbolOutput In oSymbolOutputs
                sOutputName = oSymbolOutput.Name
jIndex = jIndex + 1
sText = sText & vbCrLf & "  ...(" & Trim(Str(jIndex)) & ") sOutputName: " & sOutputName
                
                Set oOutput = oSymbol.BindToOutput(sRepName, sOutputName)
                If Not oOutput Is Nothing Then
                    If TypeOf oOutput Is IJWireBody Then
                        ' Retrieve the Edges used to defined the following Edges:
                        For kIndex = 1 To oCollectionEdgeIds.Count
                            lEdgeId = oCollectionEdgeIds(kIndex)
                            Set oEdgeWire = GetEdgeFromSymbolOutput(oOutput, lEdgeId)
                            If Not oEdgeWire Is Nothing Then
                                oCollOfOutputElements.Add oEdgeWire
sText = sText & vbCrLf & "  ...   ... found Edge Id:" & Str(lEdgeId)
                                If bReturnFirst Then
                                    GoTo ExitMethod
                                End If
                            End If
                        Next kIndex
                    End If
                
                End If
            
                Set oOutput = Nothing
                Set oSymbolOutput = Nothing
            Next
        
            Err.Clear
            On Error GoTo ErrorHandler
        
        End If
    
    Next iIndex
    
    ' Rerurn the list of Edges
ExitMethod:
If False Then
    zMsgBox sText
End If
    Set GetSymbolEdgeIds = oCollOfOutputElements
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'GetEdgeFromSymbolOutput
'
'Abstract
'   Given the Symbol Output (Attributed WireBody)
'   Retreive the Edge (WireBody segment) that represents given Edge Id
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function GetEdgeFromSymbolOutput(oWireBody As IJWireBody, lEdgeId As Long) As IJWireBody
Const METHOD = "::GetEdgeFromSymbolOutput"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim oEdgeObject As Object
    Dim oEnumUnkEdges As IEnumUnknown
    Dim oCollectionEdges As Collection
    
    Dim oConvertUtils As CCollectionConversions
    Dim oModelBodyUtils As IJSGOModelBodyUtilities
    
    Set GetEdgeFromSymbolOutput = Nothing
    Set oModelBodyUtils = New SGOModelBodyUtilities
    '   The Symbol Output WireBody Attributes are:
    '       Context Id: -999999
    '       Operation Id: 2000
    '       Operator Id: (equals Edge Id)
    '       Xid: (equals Edge Id)
    '
    oModelBodyUtils.GetEdgesByAttributes oWireBody, -999999, 2000, lEdgeId, lEdgeId, oEnumUnkEdges
    If Not oEnumUnkEdges Is Nothing Then
        Set oConvertUtils = New CCollectionConversions
        oConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumUnkEdges, oCollectionEdges
        
        For Each oEdgeObject In oCollectionEdges
            If TypeOf oEdgeObject Is IJWireBody Then
                Set GetEdgeFromSymbolOutput = oEdgeObject
                Exit Function
            End If
        Next
    End If
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'CrossSection_Flanges
'
'Abstract
'   Given the CrossSection Symbol Type
'   Determine: If Cross Section contains the following:
'            Top Flange Left section
'            Top Flange Right section
'            Bottom Flange Left section
'            Bottom Flange Right section
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub CrossSection_Flanges(oPart As Object, _
                                bTopFlangeLeft As Boolean, bBottomFlangeLeft As Boolean, _
                                bTopFlangeRight As Boolean, bBottomFlangeRight As Boolean)
Const METHOD = "::CrossSection_Flanges"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim sCStype As String
    
    Dim oCrossSection As IJCrossSection
    Dim oPartDeisgned As ISPSDesignedMember
    Dim oPartPrismatic As ISPSMemberPartPrismatic
    Dim oSPS_CrossSection As ISPSCrossSection
    Dim oMemberPart As ISPSMemberPartCommon
    Dim oStiffener As New StructDetailObjects.ProfilePart
    
    bTopFlangeLeft = False
    bBottomFlangeLeft = False
    
    bTopFlangeRight = False
    bBottomFlangeRight = False
    
    If oPart Is Nothing Then
        sMsg = "Invalid Argument passed : Argument passed is Nothing. Error Out"
        GoTo ErrorHandler
    End If
    
    If TypeOf oPart Is ISPSMemberPartCommon Then
        Set oMemberPart = oPart
        
        If oMemberPart.IsPrismatic Then
            Set oPartPrismatic = oMemberPart
            Set oSPS_CrossSection = oPartPrismatic.CrossSection
        ElseIf TypeOf oMemberPart Is ISPSDesignedMember Then
            Set oPartDeisgned = oMemberPart
            Set oSPS_CrossSection = oPartDeisgned
        Else
            Exit Sub
        End If
        
    ElseIf TypeOf oPart Is IJProfile Then
        Set oStiffener.object = oPart
        sCStype = oStiffener.sectionType
    Else
        'Unknown Type
        Exit Sub
    End If
    
    If TypeOf oPart Is ISPSMemberPartCommon Then
        If TypeOf oSPS_CrossSection.definition Is IJCrossSection Then
            Set oCrossSection = oSPS_CrossSection.definition
            sCStype = oCrossSection.Type
        Else
            sCStype = ""
        End If
    End If
    
    Select Case UCase(sCStype)
        Case "2C"
        bTopFlangeLeft = True
        bTopFlangeRight = True
        bBottomFlangeLeft = True
        bBottomFlangeRight = True
        Case "2L"
        bBottomFlangeLeft = True
        bBottomFlangeRight = True
        Case "C", "MC", "C_S", "C_SS", "CSTYPE"
        bTopFlangeRight = True
        bBottomFlangeRight = True
        Case "HSSC", "PIPE"
        Case "HSSR", "RS", "RECT"
        Case "L"
        bBottomFlangeRight = True
        Case "T", "MT", "ST", "WT", "BUT", "T_XTYPE", "TSTYPE"
        bTopFlangeLeft = True
        bTopFlangeRight = True
        Case "W", "M", "HP", "S", "H", "I", "ISTYPE"
        bTopFlangeLeft = True
        bTopFlangeRight = True
        bBottomFlangeLeft = True
        bBottomFlangeRight = True
        Case "EA", "UA"
            bTopFlangeRight = True
        Case "BUTL3", "BUTL2"
            bTopFlangeRight = True
        Case Else
    End Select
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub
'*************************************************************************
'Function
'ISPSFACInputHelper_ValidateObjects
'***************************************************************************
Public Function InputHelper_ValidateObjects( _
                            ByVal oInputObjs As SP3DStructInterfaces.IJElements, _
                            oRelationObjs As SP3DStructInterfaces.IJElements) _
                            As SP3DStructInterfaces.SPSFACInputHelperStatus
Const METHOD = "::InputHelper_ValidateObjects"
    
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lCount As Long
    Dim oPortCol As IJElements
    
    Dim ePortIdx As SPSMemberAxisPortIndex
    
    Dim oInputObj1 As Object
    Dim oInputObj2 As Object
    Dim oFrmConn As ISPSFrameConnection
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppingPort As ISPSSplitAxisPort
    
    
    InputHelper_ValidateObjects = SPSFACInputHelper_Ok
    Set oPortCol = New JObjectCollection
    'filter out ports to portCol
    For lCount = 1 To oInputObjs.Count
        If oInputObjs.Item(lCount) Is Nothing Then
        ElseIf TypeOf oInputObjs.Item(lCount) Is IJPort Then
            oPortCol.Add oInputObjs.Item(lCount)
        End If
    Next lCount
    
    '  make sure there are only two ports
    If oPortCol.Count <> 2 Then
        InputHelper_ValidateObjects = SPSFACInputHelper_BadNumberOfObjects
        Exit Function
    End If
    ' Currently only supporting:
    '   Member End to Member Along Axis cases
    '   Member End to Member End cases (Split by Point)
    
    ' First Port MUST BE ISPSSplitAxisEndPort
    Set oInputObj1 = oPortCol.Item(1)
    Set oInputObj2 = oPortCol.Item(2)
    
    If Not TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
        Dim oTemp As Object
        Set oTemp = oInputObj1
        oInputObj1 = oInputObj2
        oInputObj2 = oTemp
    End If
    
    If Not TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
        InputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
        Exit Function
    End If
    
    ' Second Port MUST BE ISPSSplitAxisEndPort or ISPSSplitAxisAlongPort
    '   ISPSSplitAxisEndPort is expected for Member Split By Point cases
    '   ISPSSplitAxisAlongPort is expected for Bounded cases
    
    If Not TypeOf oInputObj2 Is ISPSSplitAxisEndPort Then
        If Not TypeOf oInputObj2 Is ISPSSplitAxisAlongPort Then
            InputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
            Exit Function
        End If
    End If
            
    ' Inputs are as expected: this is a valid Assembly Connection
    Set oSuppedPort = oInputObj1
    Set oSuppingPort = oInputObj2
                
    If oRelationObjs Is Nothing Then
        Set oRelationObjs = New JObjectCollection
    End If
    
    oRelationObjs.Clear
    oRelationObjs.Add oSuppedPort
    oRelationObjs.Add oSuppingPort
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'GetSupportingEndPort
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
Public Sub GetSupportingEndPort(oBoundedData As MemberConnectionData, _
                                oBoundingData As MemberConnectionData, _
                                oSupportingEndPort As ISPSSplitAxisPort)
Const METHOD = "::GetSupportingEndPort"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dDistEnd As Double
    Dim dDistStart As Double
    
    Dim oBoundedPoint As IJPoint
    Dim oBoundingPoint_End As IJPoint
    Dim oBoundingPoint_Start As IJPoint
    
    sMsg = ""
    
    ' Check that the Supported and Supporting Member data is valid
    If oBoundedData.MemberPart Is Nothing Then
        sMsg = "Supported Member data is not valid"
        GoTo ErrorHandler
    
    ElseIf oBoundedData.ePortId = SPSMemberAxisAlong Then
        sMsg = "Supported Member Port data is not valid"
        GoTo ErrorHandler
    
    ElseIf oBoundingData.MemberPart Is Nothing Then
        sMsg = "Supporting Member data is not valid"
        GoTo ErrorHandler
    
    ElseIf oBoundingData.ePortId <> SPSMemberAxisAlong Then
        Set oSupportingEndPort = oBoundingData.AxisPort
        Exit Sub
    End If
    
    ' Have a valid Supported Member data: (Bounded Port is SPSMemberAxisStart or SPSMemberAxisEnd)
    ' Have a valid Supporting Member data: (Bounding Port is SPSMemberAxisAlong)
    ' Calculate the distance from the Supported End Point and the Supporting End points
    ' return the closest Supporting End Port
    Set oBoundedPoint = oBoundedData.MemberPart.PointAtEnd(oBoundedData.ePortId)
    Set oBoundingPoint_End = oBoundingData.MemberPart.PointAtEnd(SPSMemberAxisEnd)
    Set oBoundingPoint_Start = oBoundingData.MemberPart.PointAtEnd(SPSMemberAxisStart)
    dDistEnd = oBoundedPoint.DistFromPt(oBoundingPoint_End)
    dDistStart = oBoundedPoint.DistFromPt(oBoundingPoint_Start)
    If dDistStart < dDistEnd Then
        Set oSupportingEndPort = oBoundingData.MemberPart.AxisPort(SPSMemberAxisStart)
    Else
        Set oSupportingEndPort = oBoundingData.MemberPart.AxisPort(SPSMemberAxisEnd)
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'BoundingPointToWebPlane
'
'Abstract
'   Given the Bounded and Bounding members for an End to End Assembly Connection
'   Determine:
'       1. If the Bounding Member Port is an end port or the AlongAxis port
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
Public Sub BoundingPointToWebPlane(oPos As IJDPosition, _
                                   oBoundedData As MemberConnectionData, _
                                   oBoundingData As MemberConnectionData, _
                                   bReturn2DPosition As Boolean, _
                                   oProjectedPos As IJDPosition, _
                                   Optional bBoundingAlongAxis As Boolean = False)
Const METHOD = "::BoundingPointToWebPlane"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dDist As Double
    
    Dim oWebNormal As IJDVector
    Dim oProjectVec As IJDVector
    Dim oBoundingAlongAxis As IJDVector
    
    Dim o3DspacePos As IJDPosition
    Dim oInvertMatrix As IJDT4x4
    
    sMsg = ""
    
    ' given oBoundedData.Matrix and oBoundingData.Matrix
    ' Matrix.IndexValue(0,1,2)   : U is direction Along Axis
    ' Matrix.IndexValue(4,5,6)   : V is direction normal to Web (from Web Right to Web Left)
    ' Matrix.IndexValue(8,9,10)  : W is direction normal to Flange (from Flange Bottom to Flange Top)
    ' Matrix.IndexValue(12,13,14): Root/Origin Point
    Set oWebNormal = New dVector
    oWebNormal.Set oBoundedData.Matrix.IndexValue(4), _
                     oBoundedData.Matrix.IndexValue(5), oBoundedData.Matrix.IndexValue(6)
    
    ' Calculate vector from  Point to be Projected to Web Plane Root Point
    Set oProjectVec = New dVector
    oProjectVec.Set oBoundedData.Matrix.IndexValue(12) - oPos.x, _
                    oBoundedData.Matrix.IndexValue(13) - oPos.y, _
                    oBoundedData.Matrix.IndexValue(14) - oPos.z
             
    ' Normalize the Web Plane Normal
    oWebNormal.Length = 1#
    
    'Calculate the projected length of oVec along plane normal
    dDist = oWebNormal.Dot(oProjectVec)
    If bBoundingAlongAxis Then
        ' User requested that the Poisiton be projected onto the Web Plane
        ' based on the Bounding Axis Curve
        Set oBoundingAlongAxis = New dVector
        oBoundingAlongAxis.Set oBoundingData.Matrix.IndexValue(0), _
                               oBoundingData.Matrix.IndexValue(1), _
                               oBoundingData.Matrix.IndexValue(2)
        oBoundingAlongAxis.Length = 1#
        dDist = dDist / Abs(oBoundingAlongAxis.Dot(oWebNormal))
    End If
    
    'Project the Position to the plane
    Set oProjectedPos = New DPosition
    oProjectedPos.x = dDist * oWebNormal.x + oPos.x
    oProjectedPos.y = dDist * oWebNormal.y + oPos.y
    oProjectedPos.z = dDist * oWebNormal.z + oPos.z

    If bReturn2DPosition Then
        Set oInvertMatrix = oBoundedData.Matrix.Clone
        oInvertMatrix.Invert
        Set o3DspacePos = oProjectedPos.Clone
        Set oProjectedPos = oInvertMatrix.TransformPosition(o3DspacePos)
    End If
    

    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'z_TransformPosToGlobal
'
'Abstract
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function z_TransformPosToGlobal(oMemberPart As ISPSMemberPartCommon, _
                                       oInputPos As IJDPosition, _
                                       Optional oPosAlongAxis As IJDPosition) As IJDPosition
Const MT = "z_TransformPosToGlobal"
  On Error GoTo ErrorHandler
    Dim oProfileBO As ISPSCrossSection
    Dim oPartDesigned As ISPSDesignedMember
    Dim oPartPrismatic As ISPSMemberPartPrismatic
    
    Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
    Dim oMat1 As New DT4x4, oMat2 As DT4x4
    Dim oVec As New dVector

    If oMemberPart.IsPrismatic Then
        Set oPartPrismatic = oMemberPart
        Set oProfileBO = oPartPrismatic.CrossSection
    
    ElseIf TypeOf oMemberPart Is ISPSDesignedMember Then
        Set oPartDesigned = oPartDesigned
        Set oProfileBO = oPartDesigned
    End If
    
    If oProfileBO Is Nothing Then
        xOffset = 0
        yOffset = 0
    
    Else
        oProfileBO.GetCardinalPointOffset oProfileBO.CardinalPoint, xOffsetCP, yOffsetCP 'Returns x and y of the current CP in RAD coordinates, which is member negative y and z.
        oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5
        xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
        yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
        xOffset = -xOffsetCP  ' x offset of current cp from cp5
        yOffset = -yOffsetCP  ' y offset of current cp from cp5
    End If
    
    oVec.Set xOffset, yOffset, 0
    oMat1.LoadIdentity
    oMat1.Translate oVec
    
    If oPosAlongAxis Is Nothing Then
        oMemberPart.Rotation.GetTransform oMat2
    Else
        oMemberPart.Rotation.GetTransformAtPosition oPosAlongAxis.x, oPosAlongAxis.y, oPosAlongAxis.z, oMat2, Nothing
    End If
    'from cross section coordinates to member coordinates
    Set oMat2 = CreateCSToMembTransform(oMat2, oMemberPart.Rotation.Mirror)
    oMat2.MultMatrix oMat1
    Set z_TransformPosToGlobal = oMat2.TransformPosition(oInputPos)
  Exit Function
ErrorHandler:    HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'CheckSupportedSupportingSize
'
'Abstract
'   Given the Bounded and Bounding members, Determine:
'       1. The Idealized Boundary, Web_Left, Web_Right, Top, or Bottom
'
'
'input
'
'Return
'       lBoundingIsLarger : -1 Supported Member Cross Section Size is Larger then Supporting
'       lBoundingIsLarger :  0 Size comparision could not be determined
'       lBoundingIsLarger : +1 Supported Member Cross Section Size is Smaller then Supporting
'
'Exceptions
'
'***************************************************************************
Public Sub CheckSupportedSupportingSize(oBoundedData As MemberConnectionData, _
                                        oBoundingData As MemberConnectionData, _
                                        sIdealizedBoundary As String, _
                                        lBoundingIsLarger As Long, _
                                        Optional dTolerance As Double = 0.025)
Const METHOD = "::CheckSupportedSupportingSize"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dBoundedSize As Double
    Dim dBoundingSize As Double
    
    Dim oBoundedTopPos As IJDPosition
    Dim oBoundedBottomPos As IJDPosition
    Dim oBoundingTopPos As IJDPosition
    Dim oBoundingBottomPos As IJDPosition
    
    Dim oBounded_WebFlangePoints As Collection
    Dim oBounding_WebFlangePoints As Collection
    
    lBoundingIsLarger = 0
    ' Retrieve the 3D points that represent the Bounding Member Flange sections
    ' Top, Top_Flange_Bottom, Bottom_Flange_top, and Bottom locations
    Set oBounding_WebFlangePoints = GetMemberWebFlangePoints(oBoundingData.MemberPart, _
                                                             oBoundingData.Matrix, _
                                                             sIdealizedBoundary)
    If oBounding_WebFlangePoints.Count < 4 Then
        Exit Sub
    End If
    
    ' Retrieve the 3D points that represent the Bounded Member Flange Sections
    ' Top, Top_Flange_Bottom, Bottom_Flange_top, and Bottom locations
    Set oBounded_WebFlangePoints = GetMemberWebFlangePoints(oBoundedData.MemberPart, _
                                                            oBoundedData.Matrix)
    If oBounding_WebFlangePoints.Count < 4 Then
        Exit Sub
    End If
    
    Set oBoundedTopPos = oBounded_WebFlangePoints.Item(1)
    Set oBoundedBottomPos = oBounded_WebFlangePoints.Item(4)
    dBoundedSize = oBoundedTopPos.DistPt(oBoundedBottomPos)
    
    Set oBoundingTopPos = oBounding_WebFlangePoints.Item(1)
    Set oBoundingBottomPos = oBounding_WebFlangePoints.Item(4)
    dBoundingSize = oBoundingTopPos.DistPt(oBoundingBottomPos)
    
    ' Compare the distance between the Top and Bottom Points
    ' to Determine if the Bounding Member is larger,same,smaller then Bounded Member
    If dBoundedSize > (dBoundingSize - dTolerance) Then
        lBoundingIsLarger = -1
    Else
        lBoundingIsLarger = 1
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'GetSmartItemSelection
'
'Abstract
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function GetSmartItemSelection(oSmartObject As Object) As String
Const METHOD = "::GetSmartItemSelection"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sSelection As String
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition
    
    Dim oHelper As GSCADSmartOccurrence.IJSmartOccurrenceHelper
    sSelection = ""
    If TypeOf oSmartObject Is IJSmartOccurrence Then
        Set oHelper = New GSCADSmartOccurrence.CSmartOccurrenceCES
        Set oSmartOccurrence = oSmartObject
        Set oSmartItem = oSmartOccurrence.SmartItemObject
        Set oSmartClass = oSmartItem.Parent
        Set oSymbolDefinition = oSmartClass.SelectionRuleDef
        sSelection = oHelper.CurrentSelection(oSymbolDefinition, oSmartObject)
        If Len(Trim(sSelection)) < 1 Then
            ' No Selection Rule Choice,
            ' User must have choice the Item directly, return the the Smart Item
            sSelection = oSmartItem.Name
        End If
    End If
    
    GetSmartItemSelection = sSelection
  Exit Function
  
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'CheckIdealizedBoundary
'
'Abstract
'   Given the Bounded and Bounding Connection Data
'   Determine IdealizedBoundary is: TopEnd To End
'                                   eIdealized_Unk = "Unknown"
'                                   eIdealized_Top = "Top"
'                                   eIdealized_Bottom = "Bottom"
'                                   eIdealized_WebLeft = "Web_Left"
'                                   eIdealized_WebRight = "Web_Right"
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub CheckIdealizedBoundary(oBoundedData As MemberConnectionData, _
                                  oBoundingData As MemberConnectionData, _
                                  sIdealizedBoundary As String)
Const METHOD = "::CheckIdealizedBoundary"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim bTubular As Boolean
    
    Dim dDot As Double
    Dim dBest As Double
    Dim oBoundedAxis_Vector As IJDVector
    Dim oBoundingWeb_Vector As IJDVector
    Dim oBoundingFlange_Vector As IJDVector
    
    Dim lIdealEdgeId As Long
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    Dim oBoundedConnectable  As IJConnectable
    
    ' Check if Bounding member Cross Section is Tubular
    sIdealizedBoundary = eIdealized_Unk
    bTubular = IsTubularMember(oBoundingData.MemberPart)
    If bTubular Then
        sIdealizedBoundary = eIdealized_BoundingTube
        Exit Sub
    End If
    
    ' "BuiltUP" DesignedMember Parts do NOT support IJStructProfilePart
    Set oBoundedConnectable = oBoundedData.MemberPart
    If Not TypeOf oBoundedConnectable Is IJStructProfilePart Then
        Exit Sub
    End If
    
    Set oStructProfilePart = oBoundedConnectable
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    
    If oBoundedData.AxisPort Is Nothing Or oBoundingData.AxisPort Is Nothing Then Exit Sub
    
    lIdealEdgeId = oStructEndCutUtil.GetIdealizedBoundaryId(oBoundedData.AxisPort, _
                                                            oBoundingData.AxisPort)
    If lIdealEdgeId = CTX_BASE Then
        sIdealizedBoundary = eIdealized_EndBaseFace
    ElseIf lIdealEdgeId = CTX_OFFSET Then
        sIdealizedBoundary = eIdealized_EndOffsetFace
    ElseIf lIdealEdgeId = JXSEC_TOP Then
        sIdealizedBoundary = eIdealized_Top
    ElseIf lIdealEdgeId = JXSEC_BOTTOM Then
        sIdealizedBoundary = eIdealized_Bottom
    ElseIf lIdealEdgeId = JXSEC_WEB_LEFT Then
        sIdealizedBoundary = eIdealized_WebLeft
    ElseIf lIdealEdgeId = JXSEC_WEB_RIGHT Then
        sIdealizedBoundary = eIdealized_WebRight
    End If
    
    Exit Sub
    
    ' Matrix.IndexValue(0,1,2)   : U is direction Along Axis
    ' Matrix.IndexValue(4,5,6)   : V is direction normal to Web (from Web Right to Web Left)
    ' Matrix.IndexValue(8,9,10)  : W is direction normal to Flange (from Flange Bottom to Flange Top)
    ' Matrix.IndexValue(12,13,14): Root/Origin Point
    '
    Set oBoundedAxis_Vector = New dVector
    If oBoundedData.ePortId = SPSMemberAxisStart Then
        oBoundedAxis_Vector.Set oBoundedData.Matrix.IndexValue(0), _
                                oBoundedData.Matrix.IndexValue(1), _
                                oBoundedData.Matrix.IndexValue(2)
    ElseIf oBoundedData.ePortId = SPSMemberAxisEnd Then
        oBoundedAxis_Vector.Set -oBoundedData.Matrix.IndexValue(0), _
                                -oBoundedData.Matrix.IndexValue(1), _
                                -oBoundedData.Matrix.IndexValue(2)
    Else
        Exit Sub
    End If
    
    Set oBoundingWeb_Vector = New dVector
    oBoundingWeb_Vector.Set oBoundingData.Matrix.IndexValue(4), _
                            oBoundingData.Matrix.IndexValue(5), oBoundingData.Matrix.IndexValue(6)
    
    Set oBoundingFlange_Vector = New dVector
    oBoundingFlange_Vector.Set oBoundingData.Matrix.IndexValue(8), _
                               oBoundingData.Matrix.IndexValue(9), oBoundingData.Matrix.IndexValue(10)
    
    dBest = oBoundedAxis_Vector.Dot(oBoundingWeb_Vector)
    If dBest > 0.00001 Then
        sIdealizedBoundary = eIdealized_WebLeft
    ElseIf dBest < -0.00001 Then
        dBest = Abs(dBest)
        sIdealizedBoundary = eIdealized_WebRight
    Else
        dBest = 0#
    End If
    
    dDot = oBoundedAxis_Vector.Dot(oBoundingFlange_Vector)
    If dDot > dBest Then
        sIdealizedBoundary = eIdealized_Top
    ElseIf Abs(dDot) > dBest Then
        sIdealizedBoundary = eIdealized_Bottom
    End If
    
'################################################################################
'################################################################################
If False Then
    sMsg = METHOD & "   ...sIdealizedBoundary = " & sIdealizedBoundary
    zMsgBox sMsg
End If
'################################################################################
'################################################################################
    

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub

'*************************************************************************
'Function
'GetPositionFromElementsList
'
'Abstract
'   Given the list of Elements (ports) for the Assembly Connection
'   Determine the Assembly Connection Location Point
'   Expecting at least One Point to be a Point object
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub GetPositionFromElementsList(oElements As IJElements, _
                                       ByRef oRefPoint As IJDPosition)
Const METHOD = "::GetPositionFromElementsList"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim nCount As Long
    
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    
    Dim oObj As Object
    Dim oPort As IJPort
    Dim oPoint As IJPoint
    Dim oGeometry As Object
    Dim oSurfaceBody As IJSurfaceBody
    Dim oTopologyLocate As New TopologyLocate
    Dim oNormalVec As IJDVector
    Dim oPosition As IJDPosition
    
    If oRefPoint Is Nothing Then
        Set oRefPoint = New AutoMath.DPosition
        oRefPoint.Set 0#, 0#, 0#
    End If

    nCount = oElements.Count
    For iIndex = 1 To nCount
        Set oObj = oElements(iIndex)
        If TypeOf oObj Is IJPort Then
            Set oPort = oObj
            Set oGeometry = oPort.Geometry
            If TypeOf oGeometry Is IJPoint Then
                Set oPoint = oGeometry
                oPoint.GetPoint dx, dy, dz
                oRefPoint.Set dx, dy, dz
                Exit Sub
            ElseIf TypeOf oGeometry Is IJSurfaceBody Then
                Set oSurfaceBody = oGeometry
                oTopologyLocate.FindApproxCenterAndNormal oSurfaceBody, oPosition, oNormalVec
                oPosition.Get dx, dy, dz
                oRefPoint.Set dx, dy, dz
                Exit Sub
            End If
        ElseIf TypeOf oObj Is IJPoint Then
            Set oPoint = oObj
            oPoint.GetPoint dx, dy, dz
            oRefPoint.Set dx, dy, dz
            Exit Sub
        End If
    Next iIndex
    
    'TODO: in the future, we could support get xyz from the intersection of two line ports.
    Err.Raise E_FAIL            'did NOT find one. x,y,z not set.

ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'SelectReplacingObject
'
'Abstract
'   Loop thru Collection of Replacing Objects
'   Return the Object closest to the given Reference Location
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub SelectReplacingObject(oObjectCollectionReplacing As IJDObjectCollection, _
                                 oLocation As IJDPosition, _
                                 oReplacingObject As Object)
Const METHOD = "::SelectReplacingObject"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    
    Dim dDist As Double
    Dim dMinDist As Double
    
    Dim oPosOnReplacing As IJDPosition
    Dim oObjectReplacing As Object
    
    iIndex = 0
    dMinDist = -1#
    sMsg = "Checking/Verifing Collection of Replacing Objects"
    If oObjectCollectionReplacing Is Nothing Then
        sMsg = "Collection of Replacing Objects is Nothing"
    
    ElseIf oObjectCollectionReplacing.Count < 1 Then
        sMsg = "Collection of Replacing Objects count < 1"

    ElseIf oObjectCollectionReplacing.Count = 1 Then
        ' Need to use For...Each... Loop to get object from Collection
        sMsg = "Returning single item in ObjectCollectionReplacing"
        For Each oObjectReplacing In oObjectCollectionReplacing
            Set oReplacingObject = oObjectReplacing
            Exit For
        Next
    
    Else
        sMsg = "Looping Thru the Collection of replacing Objects"
        For Each oObjectReplacing In oObjectCollectionReplacing
            iIndex = iIndex + 1
            sMsg = "Getting Point on Replacing Object:" & iIndex
            GetPointOnObject oObjectReplacing, oLocation, oPosOnReplacing
            
            If oPosOnReplacing Is Nothing Then
                sMsg = "Failed getting Point on Replacing Object:" & iIndex
    
            Else
                dDist = oLocation.DistPt(oPosOnReplacing)
            
                sMsg = "Checking if Replacing Object is closest:" & iIndex & _
                        "  ... Dist:" & dDist
                If dMinDist < -0.1 Then
                    dMinDist = dDist
                    Set oReplacingObject = oObjectReplacing
                ElseIf dDist < dMinDist Then
                    dMinDist = dDist
                    Set oReplacingObject = oObjectReplacing
                End If
            End If
    
        Next
    End If

    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub

' ******************************************************************************************
' ******************************************************************************************
' ******************************************************************************************

'*************************************************************************
'Function
'MigrateAssemblyConnection
'
'Abstract
'   Migrate the Assembly Connection Port Objects and Member Objects (End Cuts)
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub MigrateAssemblyConnection(pAggregatorDescription As IJDAggregatorDescription, _
                                     pMigrateHelper As IJMigrateHelper)
Const METHOD = "::MigrateAssemblyConnection"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    
    Dim oAppConn As Object
    Dim oMemberObjects As IJDMemberObjects
    
    Dim oMemberDatatypes As Collection
    Dim oMemberDataObjects As Collection
    
    sMsg = "Check/verify Smart Occurrence Item"
    Set oAppConn = pAggregatorDescription.CAO
    
    If oAppConn Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oAppConn Is IJAppConnection Then
        Exit Sub
    ElseIf Not TypeOf oAppConn Is IJStructAssemblyConnection Then
        Exit Sub
    End If
    
If bSM_trace Then
    zSM_trace vbCrLf
    Migrate_TraceMembers oAppConn, pMigrateHelper, "000"
    zSM_trace "*** MigrateAssemblyConnection ... oAppConn : " & Debug_ObjectName(oAppConn, True)
End If

    ' Migrate the IJStructAssemblyConnection input ports
    ' ... Migrate ALL dependent Member items
    Set m_ReplacedParts = Nothing
    Set m_ReplacingParts = Nothing
    m_MigratedFeaturesSize = 100
    m_MigratedFeaturesCount = 0
    ReDim m_MigratedFeatures(m_MigratedFeaturesSize)
    Migrate_MemberItems oAppConn, pMigrateHelper, True, True, True
    
If bSM_trace Then
    zSM_trace vbCrLf
    Migrate_TraceMembers oAppConn, pMigrateHelper, "999"
End If
        
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'MigrateEndCutObject
'
'Abstract
'   Migrate the Physical Connections created by the End Cut Object
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub MigrateEndCutObject(oAggregatorDescription As IJDAggregatorDescription, _
                               oMigrateHelper As IJMigrateHelper)
Const METHOD = "::MigrateEndCutObject"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim iIndex As Long
    
    Dim oEndCutObj As Object
    Dim oSmartParent As Object
    
    sMsg = "Check/verify Smart Occurrence Item"
    Set oEndCutObj = oAggregatorDescription.CAO
    If oEndCutObj Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oEndCutObj Is IJSmartOccurrence Then
        Exit Sub
    End If
    
If bSM_trace Then
    zSM_trace vbCrLf
    Dim oMemberDatatypes As Collection
    Dim oMemberDataOjects As Collection
    Migrate_GetMemberData oEndCutObj, oMemberDataOjects, oMemberDatatypes
    Migrate_TraceObject oEndCutObj, oMemberDataOjects, oMemberDatatypes, oMigrateHelper, "888"
End If
        
    ' Migration of Endcuts is done from the AssemblyConnection or FreeEndCut
    ' i.e.: there is no need to process each invdivual EndCut object
    Exit Sub
    
    ' Based on Content,
    ' if Content does NOT migrate Free EndCuts,
    ' ... then Free EndCuts can be migrated by checking the end cut Parent
    ' BUT
    ' Expect the Free End Cut objects to be Migrated on thier own
    ' ... Current uses C# for Free EndCuts
    ' ... EndCuts,Ingr.SP3D.Content.Structure.FreeEndCutDefinition
    ' ... ... Free End Cut Split/Migration is not as expected
    '
    ' check if EndCut Parent is a IJFreeEndCut
    ' ... EndCuts are migrated by the Assembly Connections
    ' ... Expect for EndCuts controlled by the FreeEndCut
    GetSmartOccurrenceParent oEndCutObj, oSmartParent
    If oSmartParent Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oSmartParent Is IJFreeEndCut Then
        Exit Sub
    End If
    
If bSM_trace Then
    zSM_trace vbCrLf
    Migrate_TraceMembers oSmartParent, oMigrateHelper, "000"
    zSM_trace "*** MigrateEndCutObject ... oSmartParent : " & Debug_ObjectName(oSmartParent, True)
End If

    ' Migrate the IJFreeEndCut input ports
    ' ... Migrate ALL dependent Member items
    Set m_ReplacedParts = Nothing
    Set m_ReplacingParts = Nothing
    Migrate_MemberItems oSmartParent, oMigrateHelper, True, True, True

If bSM_trace Then
    zSM_trace vbCrLf
    Migrate_TraceMembers oSmartParent, oMigrateHelper, "999"
End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'MigrateFreeEndCut
'
'Abstract
'   Migrate the Frr EndCut Input Objects and Member Objects (End Cuts)
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub MigrateFreeEndCut(pAggregatorDescription As IJDAggregatorDescription, _
                             pMigrateHelper As IJMigrateHelper)
Const METHOD = "::MigrateFreeEndCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    
    Dim oFreeEndcut As Object
    Dim oMemberObjects As IJDMemberObjects
    
    Dim oMemberDatatypes As Collection
    Dim oMemberDataObjects As Collection
    
    sMsg = "Check/verify Smart Occurrence Item"
    Set oFreeEndcut = pAggregatorDescription.CAO
    
    If oFreeEndcut Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oFreeEndcut Is IJFreeEndCut Then
        Exit Sub
    End If
    
If bSM_trace Then
    zSM_trace vbCrLf
    Migrate_TraceMembers oFreeEndcut, pMigrateHelper, "000"
    zSM_trace "*** MigrateFreeEndCut ... oFreeEndcut : " & Debug_ObjectName(oFreeEndcut, True)
End If

    ' Migrate the IJFreeendcut input ports
    ' ... Migrate ALL dependent Member items
    Set m_ReplacedParts = Nothing
    Set m_ReplacingParts = Nothing
    m_MigratedFeaturesSize = 100
    m_MigratedFeaturesCount = 0
    ReDim m_MigratedFeatures(m_MigratedFeaturesSize)
    Migrate_MemberItems oFreeEndcut, pMigrateHelper, True, True, True
    
If bSM_trace Then
    zSM_trace vbCrLf
    Migrate_TraceMembers oFreeEndcut, pMigrateHelper, "999"
End If
        
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

' ******************************************************************************************
' ******************************************************************************************
' ******************************************************************************************

'*************************************************************************
'Function
'AxisSel_WebCutType
'
'Abstract
'   Given the Bounded and Bounding members, Determine:
'       1. The Idealized Boundary, Web_Left, Web_Right, Top, or Bottom
'
'
'input
'   oBoundedData
'   oBoundingData
'   sIdealizedBoundary
'   sConfig
'
'   sConfig = "Top_Top"     : Bounded/Bounding Top Flanges are in same general direction
'   sConfig = "Bottom_Top"  : Bounded/Bounding Top Flanges are in the opposite general direction
'   sConfig = "Left_Top"    : Bounded Web_Left/Bounding Top Flange are in same general direction
'   sConfig = "Right_Top"   : Bounded Web_Right/Bounding Top Flange are in same general direction
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub AxisSel_WebCutType(oBoundedData As MemberConnectionData, _
                              oBoundingData As MemberConnectionData, _
                              sIdealizedBoundary As String, _
                              sConfig As String, _
                              sWebCutType As String, _
                              Optional dTolerance As Double = 0.005)
Const METHOD = "::AxisSel_WebCutType"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim oPosition As IJDPosition
    Dim oBounded_Position As IJDPosition
    Dim oBounding_Position As IJDPosition
    
    Dim oBounded_Points_2d As Collection
    Dim oBounded_Points_3d As Collection
    Dim oBounding_Points_2d As Collection
    Dim oBounding_Points_3d As Collection
    
    ' Calculate the local 2D Points that defined Cross Section Flange and Web Sections
    ' to determine if Bounded Member is:
    '       Inside the Bounding Member Top Flange Bottom and Bottom Flange Top
    '       Outside the Bounding Member Top and Bottom
    '       Above Bounding Member Top and Above Bounding Member Bottom Flange Top
    '       Below Bounding Member Top Flange Bottom and Below Bounding Member Bottom
    '       ---------------- ......Top
    '       |              |
    '       ------    ------ ......Top Flange Bottom
    '             |  |
    '             |  |
    '             |  |
    '             |  |
    '             |  |
    '             |  |
    '       ------    ------ ......Bottom Flange Top
    '       |              |
    '       ---------------- ......Bottom
    '
    ' Convert 2D Points to 3D that represent the Bounding Web/Flange sections
    Set oBounding_Points_3d = GetMemberWebFlangePoints(oBoundingData.MemberPart, _
                                                       oBoundingData.Matrix, _
                                                       sIdealizedBoundary)
    If oBounding_Points_3d.Count < 4 Then
        Exit Sub
    End If
    
    ' Convert 2D Points to 3D to represent the Bounded Web/Flange sections
    Set oBounded_Points_3d = GetMemberWebFlangePoints(oBoundedData.MemberPart, _
                                                            oBoundedData.Matrix)
    If oBounded_Points_3d.Count < 4 Then
        Exit Sub
    End If
    
    ' Convert 3D Points to 2D that represent the Bounding Web/Flange sections
    ' in the Bounded Web Plane (Top Flange Bottom point and Bottom Flange Top point)
    Set oBounded_Points_2d = New Collection
    Set oBounding_Points_2d = New Collection
    
    If Trim(LCase(sConfig)) = LCase("Bottom_Top") Then
        For iIndex = 1 To 4
            BoundingPointToWebPlane oBounding_Points_3d.Item(5 - iIndex), _
                                    oBoundedData, oBoundingData, _
                                    True, oPosition, True
            oBounding_Points_2d.Add oPosition
        Next iIndex
    Else
        For iIndex = 1 To 4
            BoundingPointToWebPlane oBounding_Points_3d.Item(iIndex), _
                                    oBoundedData, oBoundingData, _
                                    True, oPosition, True
            oBounding_Points_2d.Add oPosition
        Next iIndex
    End If
    ' Convert 3D Points to 2D that represent the Bounded Web/Flange sections
    ' in the Bounded Web Plane (Top Flange Top point and Bottom Flange Bottom point)
    For iIndex = 1 To 4
        BoundingPointToWebPlane oBounded_Points_3d.Item(iIndex), _
                                oBoundedData, oBoundingData, _
                                True, oPosition, True
        oBounded_Points_2d.Add oPosition
    Next iIndex
    
    ' Check if Bounding has a Top Flange Section
    sWebCutType = "W1"
    Set oBounded_Position = oBounded_Points_2d.Item(1)
    
    Set oPosition = oBounding_Points_2d.Item(1)
    Set oBounding_Position = oBounding_Points_2d.Item(2)
    If oPosition.z > (oBounding_Position.z + 0.0025) Then
        ' Bounding Member has (Top) Flange Section
        If oBounded_Position.z > (oBounding_Position.z - dTolerance) Then
            sWebCutType = "C1"
        Else
        End If
    Else
        ' Bounding Member does NOT have (Top) Flange Section
        If oBounded_Position.z > (oBounding_Position.z + dTolerance) Then
            sWebCutType = "S1"
        End If
    End If
    
    ' Check if Bounding has a Bottom Flange Section
    Set oBounded_Position = oBounded_Points_2d.Item(4)
    
    Set oPosition = oBounding_Points_2d.Item(4)
    Set oBounding_Position = oBounding_Points_2d.Item(3)
    If oPosition.z < (oBounding_Position.z - 0.0025) Then
        ' Bounding Member has (Bottom) Flange Section
        If oBounded_Position.z < (oBounding_Position.z + dTolerance) Then
            sWebCutType = sWebCutType & "C1"
        Else
            sWebCutType = sWebCutType & "W1"
        End If
    Else
        ' Bounding Member does NOT have (Bottom) Flange Section
        If oBounded_Position.z < (oBounding_Position.z - dTolerance) Then
            sWebCutType = sWebCutType & "S1"
        Else
            sWebCutType = sWebCutType & "W1"
        End If
    End If
    
If False Then
    zMsgBox vbCrLf & "MemberUtilities:" & METHOD
    zMsgBox "...sConfig:" & sConfig
    
    For iIndex = 1 To 4
        Set oPosition = oBounded_Points_2d.Item(iIndex)
        Debug_Position "oBounded(" & Trim(Str(iIndex)) & ") ...", oPosition
    Next iIndex

    For iIndex = 1 To 4
        Set oPosition = oBounding_Points_2d.Item(iIndex)
        Debug_Position "oBounding(" & Trim(Str(iIndex)) & ")...", oPosition
    Next iIndex

    zMsgBox "...sWebCutType:" & sWebCutType
End If
    
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub

'*************************************************************************
'Function
'IsTubularMember
'
'Abstract
'   Given the Member Part or member Axis Port (ISPSSplitAxisPort)
'   Check if the Cross Section type is Tubular
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function IsTubularMember(oMemberObject As Object) As Boolean
Const METHOD = "::IsTubularMember"
    
    Dim sMsg As String
    On Error GoTo ErrorHandler
    Dim sCStype As String
    sCStype = GetMemberCrossSectionName(oMemberObject)
    
    IsTubularMember = False
    
    If Trim(LCase(sCStype)) = LCase("CS") Then
        IsTubularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("HSSC") Then
        IsTubularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("PIPE") Then
        IsTubularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("P") Or Trim(LCase(sCStype)) = LCase("R") Then
        IsTubularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("BUTube") Then
        IsTubularMember = True
    Else
        IsTubularMember = False
    End If

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function
Public Function IsRectangularMember(oMemberObject As Object) As Boolean
Const METHOD = "::IsRectangularMember"
    
    Dim sMsg As String
    On Error GoTo ErrorHandler
    
    Dim sCStype As String
    sCStype = GetMemberCrossSectionName(oMemberObject)

    IsRectangularMember = False

    If Trim(LCase(sCStype)) = LCase("RS") Then
        IsRectangularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("HSSR") Then
        IsRectangularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("BUBoxFM") Then
        IsRectangularMember = True
    Else
        IsRectangularMember = False
    End If
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'GetPointOnObject
'
'Abstract
'   Given an Geometry object nd a Reference Location
'   Return the closest point on the geometry to the given Reference Location
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub GetPointOnObject(oOnObject As Object, _
                            oAtLocation As IJDPosition, _
                            oPosOnObject As IJDPosition)
Const METHOD = "::GetPointOnObject"
    On Error GoTo ErrorHandler
    Dim sMsg As String
        
    Dim nNumPts As Long
    
    Dim dDist As Double
    Dim dPntX As Double
    Dim dPntY As Double
    Dim dPntZ As Double
    Dim dSrcX As Double
    Dim dSrcY As Double
    Dim dSrcZ As Double
    Dim dPar1 As Double

    Dim dPnts1() As Double
    Dim dPnts2() As Double
    Dim dPars1() As Double
    Dim dPars2() As Double

    Dim oGeometry As Object
    
    Dim oPort As IJPort
    Dim oPoint As IJPoint
    Dim oCurve As IJCurve
    Dim oPlane As IJPlane
    Dim oNormal As IJDVector
    Dim oVector As IJDVector
    Dim oSurface As IJSurface
    Dim oAtPoint As IJPoint
    
    Dim oGeom3DFactory As GeometryFactory
    Dim oModelUtil As IJSGOModelBodyUtilities

    sMsg = "Check if OnObject is IJPort"
    If TypeOf oOnObject Is IJPort Then
        Set oPort = oOnObject
        Set oGeometry = oPort.Geometry
    Else
        Set oGeometry = oOnObject
    End If
    Dim dist As Double
    If TypeOf oGeometry Is IJModelBody Then
        sMsg = "IJModelBody Geometry Type of oOnObject"
        Set oModelUtil = New SGOModelBodyUtilities
        oModelUtil.GetClosestPointOnBody oGeometry, oAtLocation, oPosOnObject, dist
        
    ElseIf TypeOf oGeometry Is IJPoint Then
        sMsg = "IJPoint Geometry Type of oOnObject"
        Set oPoint = oGeometry
        oPoint.GetPoint dSrcX, dSrcY, dSrcZ
        
        Set oPosOnObject = oAtLocation.Clone
        oPosOnObject.Set dSrcX, dSrcY, dSrcZ
        
    ElseIf TypeOf oGeometry Is IJCurve Then
        sMsg = "IJCurve Geometry Type of oOnObject"
        Set oGeom3DFactory = New GeometryFactory
        Set oAtPoint = oGeom3DFactory.Points3d.CreateByPoint(Nothing, _
                                                             oAtLocation.x, _
                                                             oAtLocation.y, _
                                                             oAtLocation.z)
        Set oCurve = oGeometry
        oCurve.DistanceBetween oAtPoint, dPar1, _
                               dSrcX, dSrcY, dSrcZ, dPntX, dPntY, dPntZ

        Set oPosOnObject = oAtLocation.Clone
        oPosOnObject.Set dSrcX, dSrcY, dSrcZ

    ElseIf TypeOf oGeometry Is IJSurface Then
        sMsg = "IJSurface Geometry Type of oOnObject"
        Set oGeom3DFactory = New GeometryFactory
        Set oAtPoint = oGeom3DFactory.Points3d.CreateByPoint(Nothing, _
                                                             oAtLocation.x, _
                                                             oAtLocation.y, _
                                                             oAtLocation.z)
        
        Set oSurface = oGeometry
        oSurface.DistanceBetween oAtPoint, dDist, _
                                 dSrcX, dSrcY, dSrcZ, dPntX, dPntY, dPntZ, _
                                 nNumPts, dPnts1, dPnts2, dPars1, dPars2
    
        Set oPosOnObject = oAtLocation.Clone
        oPosOnObject.Set dSrcX, dSrcY, dSrcZ
    
    ElseIf TypeOf oGeometry Is IJPlane Then
        sMsg = "IJPlane Geometry Type of oOnObject"
        Set oPlane = oGeometry
        oPlane.GetNormal dSrcX, dSrcY, dSrcZ
        oPlane.GetRootPoint dPntX, dPntY, dPntY
        
        ' Project the Point onto the Plane
        Set oPosOnObject = oAtLocation.Clone
        oPosOnObject.Set dPntX, dPntY, dSrcZ
        
        dDist = oAtLocation.DistPt(oPosOnObject)
        Set oVector = oAtLocation.Subtract(oPosOnObject)
        oVector.Length = 1#
        
        Set oNormal = oVector.Clone
        oNormal.Set dSrcX, dSrcY, dSrcZ
        
        dDist = oNormal.Dot(oVector)
        oPosOnObject.x = oAtLocation.x - (oNormal.x * dDist)
        oPosOnObject.y = oAtLocation.y - (oNormal.y * dDist)
        oPosOnObject.z = oAtLocation.z - (oNormal.z * dDist)
        
    Else
        ' unknown geometry.  perhaps a compound surface from intelliship.
        Set oPosOnObject = Nothing
        sMsg = "Unknown Geometry Type of oOnObject: " & TypeName(oOnObject)
        Err.Raise E_FAIL
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'***********************************************************************
' This is a copy of a method  by the same name in StructDetailObjects.ProfilePart
' It is copied so we can use it for members.  If it works, a CR should be filed
' to have this added to StructDetailObjects.MemberPart
'***********************************************************************
Public Function CutoutSubPort(ByVal oPartObject As Object, _
                              ByVal oFeature As Object, _
                              ByVal lXId As Long, _
                              Optional bLast As Boolean = False) As IJPort

    Const METHOD = "CutoutSubPort"
    On Error GoTo ErrorHandler
     
    Dim GeomUtils As IJTopologyLocate
    Set GeomUtils = New TopologyLocate

    Dim lCtx As Long
    Dim lOptId As Long
    Dim lOprId As Long
    
    Dim lStructPortCtx As Long
    Dim lStructPortXid As Long
    Dim lStructPortOptId As Long
    Dim lStructPortOprId As Long
    
    Dim oStructPort As IJStructPort
    
    GetLateBindMbrFeatureData oFeature, oPartObject, lCtx, lOptId, lOprId
       
    Set CutoutSubPort = GetLateBindMbrFeaturePort(oFeature, oPartObject, lOptId, lCtx, lOptId, lOprId, lXId)

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, "Error in CutoutSubPort"
End Function

Public Sub MbrCornerFeatureData(ByVal pMD As IJDMemberDescription, _
                              portXid1 As JXSEC_CODE, _
                              portXid2 As JXSEC_CODE, _
                              oFacePort As IJPort, _
                              oEdgePort1 As IJPort, _
                              oEdgePort2 As IJPort)
    Dim sMETHOD As String
    Dim sError As String
    sMETHOD = "MbrCornerFeatureData"
    
    On Error GoTo ErrorHandler
    MbrCornerFeatureDataByObject pMD.CAO, portXid1, portXid2, oFacePort, oEdgePort1, oEdgePort2
    
    Exit Sub
ErrorHandler:
      Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'***********************************************************************
' METHOD:  GetMbrEndCutCornerPorts
'
' DESCRIPTION:
'
'***********************************************************************
Public Sub GetMbrEndCutCornerPorts(oEndCut As Object, _
                                 oBoundedPart As Object, _
                                 eEndCutXid1 As Long, _
                                 eEndCutXid2 As Long, _
                                 oFacePort As IJPort, _
                                 oEdgePort1 As IJPort, _
                                 oEdgePort2 As IJPort)
    Dim sMETHOD As String
    sMETHOD = "GetMbrEndCutCornerPorts"
    
    On Error GoTo ErrorHandler
    
    Dim sError As String
    
    Dim lFeatureCtx As Long
    Dim lFeatureOptId As Long
    Dim lFeatureOprId As Long
    
    Dim lStructPortCtx As Long
    Dim lStructPortXid As Long
    Dim lStructPortOptId As Long
    Dim lStructPortOprId As Long
    
    Dim oStructPort As IJStructPort
    
    ' Verify given Bounded Part is valid object
    sError = "Member EndCut Object not initialized"
    If oEndCut Is Nothing Then
        LogError Err, MODULE, sMETHOD, sError
        Exit Sub
    End If
    sError = "Member EndCut Bounded Part not initialized"
    If oBoundedPart Is Nothing Then
        LogError Err, MODULE, sMETHOD, sError
        Exit Sub
    End If
    
    ' save Local copy of Ports passed in
    Dim z_oFacePort As IJPort
    Set z_oFacePort = oFacePort
    
    sError = "GetLateBindMbrFeatureData failed"
    ' Get EndCut Feature Context, Operation, Operator data
    GetLateBindMbrFeatureData oEndCut, oBoundedPart, _
                           lFeatureCtx, lFeatureOptId, lFeatureOprId
    
    ' Replace Face Port with Port After EndCut is applied (Trim Operation)
    Set oStructPort = z_oFacePort
    lStructPortCtx = oStructPort.ContextID
    lStructPortXid = oStructPort.SectionID
    lStructPortOptId = oStructPort.OperationID
    lStructPortOprId = oStructPort.OperatorID
    Set oFacePort = GetLateBindMbrFeaturePort(oEndCut, oBoundedPart, lFeatureOptId, _
                                           lStructPortCtx, _
                                           lStructPortOptId, _
                                           lStructPortOprId, _
                                           lStructPortXid)
    Set oStructPort = oFacePort
    ' Get 1st port after endCut is applied
    Set oEdgePort1 = GetLateBindMbrFeaturePort(oEndCut, oBoundedPart, lFeatureOptId, _
                                               lFeatureCtx, _
                                               lFeatureOptId, _
                                               lFeatureOprId, _
                                               eEndCutXid1)
    Set oStructPort = oEdgePort1
    ' Get 2nd port after endCut is applied
    Set oEdgePort2 = GetLateBindMbrFeaturePort(oEndCut, oBoundedPart, lFeatureOptId, _
                                            lFeatureCtx, _
                                            lFeatureOptId, _
                                            lFeatureOprId, _
                                            eEndCutXid2)
    Set oStructPort = oEdgePort2

    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'***********************************************************************
' METHOD:  GetLateBindMbrFeatureData
'
' DESCRIPTION:
'
'***********************************************************************
Public Sub GetLateBindMbrFeatureData(oFeatureObject As Object, _
                                      oPartObject As Object, _
                                      lFeatureCtx As Long, _
                                      lFeatureOptId As Long, _
                                      lFeatureOprId As Long)
    Dim sMETHOD As String
    sMETHOD = "GetLateBindMbrFeatureData"
    
    On Error GoTo ErrorHandler
    Dim sError As String

    Dim eTmpType As Long    'JS_TOPOLOGY_PROXY_TYPE
    Dim lTmpXid As Long

    Dim oLatePort As Object
    Dim oMemberStructPort As IJStructPort
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart

    Dim oFeatureUtils As GSCADSDCreateModifyUtilities.IJSDFeatureAttributes
    
    lFeatureCtx = 0
    lFeatureOptId = 0
    lFeatureOprId = 0
    
    If TypeOf oPartObject Is ISPSMemberPartPrismatic Then
        ' Get the (Last) Global Port based on the given EndCut Feature
        ' This Port would be used for Physical Connections
        Set oStructProfilePart = oPartObject
        Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
        oStructEndCutUtil.GetLatePortForFeatureSegment oFeatureObject, -1, oLatePort
        
        ' From the Global Port, get Ctx, Opt, Opr values
        If oLatePort Is Nothing Then
        ElseIf TypeOf oLatePort Is IJStructPort Then
            Set oMemberStructPort = oLatePort
            oMemberStructPort.GetAttributes eTmpType, lFeatureCtx, _
                                            lFeatureOptId, lFeatureOprId, _
                                            lTmpXid
        End If
    Else
        Set oFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
        lFeatureCtx = oFeatureUtils.DetermineFeatureContext(oFeatureObject, oPartObject)
        lFeatureOptId = oFeatureUtils.GetFeatureOperationID(oFeatureObject)
        lFeatureOprId = oFeatureUtils.FeatureOperand(oFeatureObject, oPartObject)
    End If
    
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'***********************************************************************
' METHOD:  GetLateBindMbrFeaturePort
'
' DESCRIPTION:  Get Ports required for Corner Feature
'
'***********************************************************************
Public Function GetLateBindMbrFeaturePort(oFeatureObject As Object, _
                                          oPartObject As Object, _
                                          lFeatureOptId As Long, _
                                          lCtx As Long, _
                                          lOptId As Long, _
                                          lOprId As Long, _
                                          lXId As Long) As Object
    Dim sMETHOD As String
    sMETHOD = "GetLateBindMbrFeaturePort"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    Dim oMoniker As IMoniker
    Dim oLateBindFeaturePort As IJPort
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate

    If TypeOf oPartObject Is ISPSMemberPartPrismatic Then
        
        Set oLateBindFeaturePort = GetLateBindFeaturePort_Member(oFeatureObject, _
                                                                 oPartObject, _
                                                                 lFeatureOptId, _
                                                                 lCtx, _
                                                                 lOptId, _
                                                                 lOprId, _
                                                                 lXId)

    
    Else
    
        Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
        oTopologyLocate.GetLateBindingPortMonikerAfterCutout lOptId, _
                                                             lOprId, _
                                                             lCtx, _
                                                             lXId, _
                                                             oMoniker
        
        Set oLateBindFeaturePort = oTopologyLocate.BindMonikerToPort(oPartObject, _
                                                                     oMoniker, _
                                                                     lFeatureOptId)
        
    End If
    Set GetLateBindMbrFeaturePort = oLateBindFeaturePort
    
    Exit Function
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function GetLateBindFeaturePort_Member(oFeatureObject As Object, _
                                       oPartObject As Object, _
                                       lFeatureOptId As Long, _
                                       lCtx As Long, _
                                       lOptId As Long, _
                                       lOprId As Long, _
                                       lXId As Long) As Object
    Dim sMETHOD As String
    sMETHOD = "GetLateBindFeaturePort_Member"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    Dim iIndex As Long
    
    Dim vPort As Variant
    Dim eChkCtx As Long     'eUSER_CTX_FLAGS
    Dim eChkType As Long    'JS_TOPOLOGY_PROXY_TYPE
    Dim lChkXid As Long
    Dim lChkOptId As Long
    Dim lChkOprId As Long
    
    ' Get Stable Geometry (as opposed to last geometry)
    Dim oGraphConn As IJStructGraphConnectable
    Dim oGeomPorts As IJElements
    Dim oMemberStructPort As IJStructPort
    
    Set oGraphConn = oPartObject
    
    ' Get List of Ports AFTER EndCuts
    Dim oACTools As AssemblyConnectionTools
    Set oACTools = New AssemblyConnectionTools
    oACTools.GetBindingPort oPartObject, JS_TOPOLOGY_PROXY_LFACE, _
                            lOptId, lOprId, lCtx, lXId, _
                            "SPSMembers.SPSPartPrismaticGenerator", _
                            oMemberStructPort
    
    Set GetLateBindFeaturePort_Member = oMemberStructPort
    
    Set oACTools = Nothing
    GoTo CleanUp ' Exit Function from here
    
    Set oGraphConn = Nothing
    
    ' Loop thru the Ports searching for requested Ctx, Opt, Opr, Xid
    iIndex = 0
    For Each vPort In oGeomPorts
        iIndex = iIndex + 1
        Set oMemberStructPort = vPort
        oMemberStructPort.GetAttributes eChkType, eChkCtx, _
                                        lChkOptId, lChkOprId, lChkXid
        
        If lCtx = -1 Then
            eChkCtx = lCtx
        End If
        
        If lOptId = -1 Then
            lChkOptId = lOptId
        End If
        
        If lOprId = -1 Then
            lChkOprId = lOprId
        End If
        
        If lXId = -1 Then
            lChkXid = lXId
        End If
        
        If (eChkCtx And lCtx) <> lCtx Then
        ElseIf lChkOptId <> lOptId Then
        ElseIf lChkOprId <> lOprId Then
        ElseIf lChkXid <> lXId Then
        Else
            Set GetLateBindFeaturePort_Member = oMemberStructPort
            Exit For
        End If
    Next vPort
    
CleanUp:
    Set oGeomPorts = Nothing
    Set oMemberStructPort = Nothing

    Exit Function
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'***********************************************************************
' METHOD:
' CAConstruct_CornerFeatureBetweenTwoCuts
'
' DESCRIPTION:
' Create a corner feature between two endcuts on a member
' The cuts are indentified by the dispID, which is an index for IJDMemberObjects.ItemByIndex
' A value of -1 indicates that the object should be the object itself (oMD.CAO)
' Only the first dispID can be set to -1
' The feature goes on same face of the member as dispID1 (i.e. if dispID is a web cut, it is placed
' on the web; if dispID1 is a cut on the top flange, it goes on the top flange)
'
' Usage:
' 1) To create a corner feature between to webcuts or two flange cuts in a generic assembly connection,
'    - in the definition rule of the AC, pass the AC as oMD
'    - pass the dispID of one endcut as dispID1 (the feature will be applied on the same face)
'    - pass the dispID of a second endcut as dispID2
' 2) To create a corner feature between a web webcut and the flange cut it creates for the top flange
'    - in the definition rule of the web cut, pass the web cut as oMD
'    - pass -1 for dispID1 (the feature will be applied on the same face as the web cut)
'    - pass the dispID of the flange cut it creates at the top flange
'***********************************************************************
Public Function CreateCornerFeatureBetweenTwoEndCutsByDispID(ByVal oMD As IJDMemberDescription, _
                                                             ByVal oResMgr As IUnknown, _
                                                             dispID1 As Long, _
                                                             dispID2 As Long, _
                                                             Optional strSelection As String = vbNullString) As Object
    On Error GoTo ErrorHandler

    Dim sMETHOD As String
    Dim sError As String
    sMETHOD = "CreateCornerFeatureBetweenTwoEndCutsByDispID"
    
    ' ---------------
    ' Get the objects
    ' ---------------
    Dim oObj1 As Object
    Dim oObj2 As Object
    Dim oACObject As Object
    Dim oMemObjs As IJDMemberObjects
    
    AssemblyConnection_SmartItemName oMD.CAO, , oACObject
    '------------------------------------------------------
    'if the oACObject is MbrAxis-AC then set the oMemObjs
    'as oACObject since FlangeCut is created by Axis AC
    'Else if the oACObject is Generic-AC then set oMemObjs
    'as Webcut itself creates the FlangeCut
    '------------------------------------------------------
    If GetMbrAssemblyConnectionType(oACObject) = ACType_Axis Or GetMbrAssemblyConnectionType(oACObject) = ACType_Bounded Then
        ' If the AC is Axis AC
        Set oMemObjs = oACObject
    Else ' If the AC is Generic AC
        Set oMemObjs = oMD.CAO
    End If
    
    
    If dispID1 = dispID2 Then
        sError = "The dispIDs cannot be the same"
        GoTo ErrorHandler
    End If
    
    If dispID1 < 0 Then
        Set oObj1 = oMD.CAO
    Else
        If dispID1 > oMemObjs.Count Then
            sError = "DispID1 is out of range"
            GoTo ErrorHandler
        End If
        
        Set oObj1 = oMemObjs.ItemByDispid(dispID1)
    End If
    
    If dispID2 < 0 Then
        sError = "Only dispID1 can be set to -1"
        GoTo ErrorHandler
    Else
        If dispID2 > oMemObjs.Count Then
            sError = "DispID2 is out of range"
            GoTo ErrorHandler
        End If
        
        Set oObj2 = oMemObjs.ItemByDispid(dispID2)
    End If
        
    Set CreateCornerFeatureBetweenTwoEndCutsByDispID = _
                        CreateCornerFeatureBetweenTwoEndCuts(oObj1, oObj2, oResMgr, oMD.CAO, strSelection)
 
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CreateCornerFeatureBetweenTwoEndCuts(oFeature1 As IJStructFeature, _
                                                     oFeature2 As IJStructFeature, _
                                                     ByVal oResMgr As IUnknown, _
                                                     Optional oParent As Object, _
                                                     Optional strSelection As String = vbNullString) As Object
        
    On Error GoTo ErrorHandler

    Dim sMETHOD As String
    Dim sError As String
    sMETHOD = "CreateCornerFeatureBetweenTwoEndCuts"
    
    If strSelection = vbNullString Then
        strSelection = "MbrEndCutCorner"
    End If
    
    ' -------------------------
    ' Verify object are endcuts
    ' -------------------------
    If Not (oFeature1.get_StructFeatureType = SF_WebCut Or oFeature1.get_StructFeatureType = SF_FlangeCut) Then
        sError = "Item1 is not an endcut"
        GoTo ErrorHandler
    End If
        
    If Not (oFeature2.get_StructFeatureType = SF_WebCut Or oFeature2.get_StructFeatureType = SF_FlangeCut) Then
        sError = "Item2 is not an endcut"
        GoTo ErrorHandler
    End If
    
    ' ----------------------------------
    ' Verify object are on the same part
    ' ----------------------------------
    Dim oPart1 As Object
    Dim oPart2 As Object
    Dim oBounding1 As Object
    Dim oBounding2 As Object
    Dim oBoundingPort1 As IJPort
    Dim oBoundingPort2 As IJPort
    
    Dim oSDOWebCut As New StructDetailObjects.WebCut
    Dim oSDOFlangeCut As New StructDetailObjects.FlangeCut
    
    If oFeature1.get_StructFeatureType = SF_WebCut Then
        Set oSDOWebCut.object = oFeature1
        Set oPart1 = oSDOWebCut.Bounded
        Set oBounding1 = oSDOWebCut.Bounding
        Set oBoundingPort1 = oSDOWebCut.BoundingPort
    Else
        Set oSDOFlangeCut.object = oFeature1
        Set oPart1 = oSDOFlangeCut.Bounded
        Set oBounding1 = oSDOFlangeCut.Bounding
        Set oBoundingPort1 = oSDOFlangeCut.BoundingPort
    End If
    
    If oFeature2.get_StructFeatureType = SF_WebCut Then
        Set oSDOWebCut.object = oFeature2
        Set oPart2 = oSDOWebCut.Bounded
        Set oBounding2 = oSDOWebCut.Bounding
        Set oBoundingPort2 = oSDOWebCut.BoundingPort
    Else
        Set oSDOFlangeCut.object = oFeature2
        Set oPart2 = oSDOFlangeCut.Bounded
        Set oBounding2 = oSDOFlangeCut.Bounding
        Set oBoundingPort2 = oSDOFlangeCut.BoundingPort
    End If
    
    If Not oPart1 Is oPart2 Then
        sError = "The cuts identified by dispID1 and dispID2 must be on the same part"
        GoTo ErrorHandler
    End If
    
    ' ---------------------------------------------------------------
    ' Determine if feature is on the web, topflange, or bottom flange
    ' ---------------------------------------------------------------
    Dim faceXid As JXSEC_CODE
    
    If oFeature1.get_StructFeatureType = SF_WebCut Then
        faceXid = JXSEC_WEB_LEFT
    Else
        Dim isBottom As String
    
        GetSelectorAnswer oFeature1, "BottomFlange", isBottom
        
        If isBottom = "Yes" Then
            faceXid = JXSEC_BOTTOM
        Else
            faceXid = JXSEC_TOP
        End If
    End If
    
'    ' -------------------------------------------------------
'    ' If either part is a plate, check if the ports intersect
'    ' -------------------------------------------------------
'    If TypeOf oBounding1 Is IJPlate Or TypeOf oBounding2 Is IJPlate Then
'        Dim oBoundingPortGeom1 As IJSurfaceBody
'        Dim oBoundingPortGeom2 As IJSurfaceBody
'
'        Set oBoundingPortGeom1 = oBoundingPort1.Geometry
'        Set oBoundingPortGeom2 = oBoundingPort2.Geometry
'
'        Dim oModelUtil As IJSGOModelBodyUtilities
'        Set oModelUtil = New SGOModelBodyUtilities
'
'        Dim oPointOn1 As IJDPosition
'        Dim oPointOn2 As IJDPosition
'        Dim dist As Double
'
'        oModelUtil.GetClosestPointsBetweenTwoBodies oBoundingPortGeom1, oBoundingPortGeom2, oPointOn1, oPointOn2, dist
'
'        ' ---------------------------------------------------------------
'        ' If they don't intersect, use the lateral port, it it intersects
'        ' ---------------------------------------------------------------
'        If dist > 0.00001 Then
'            Dim oSDOPlate As New StructDetailObjects.PlatePart
'            dim oLateralPort as
'            If TypeOf oBounding1 Is IJPlate Then
'                Set oSDOPlate.object = oBounding1
'
    
    ' -------------
    ' Get the ports
    ' -------------
    Dim ctx1 As Long
    Dim opt1 As Long
    Dim opr1 As Long
    
    Dim ctx2 As Long
    Dim opt2 As Long
    Dim opr2 As Long
    
    ' Get the face port, beginning with Current Geometry
 
    Dim oFacePort As IJStructPort
    
    Dim oEdgePort1 As IJStructPort
    Dim oEdgePort2 As IJStructPort
    
    Set oFacePort = GetLateralSubPortBeforeTrim(oPart1, faceXid)
    
    ' Get the information for the two edge ports
    GetLateBindMbrFeatureData oFeature1, oPart1, ctx1, opt1, opr1
    GetLateBindMbrFeatureData oFeature2, oPart2, ctx2, opt2, opr2
    
    ' Replace face port with port after endCut is applied
    Set oFacePort = GetLateBindMbrFeaturePort(oFeature1, _
                                              oPart1, _
                                              opt1, _
                                              oFacePort.ContextID, _
                                              oFacePort.OperationID, _
                                              oFacePort.OperatorID, _
                                              oFacePort.SectionID)
    
    ' Get the edge ports
    Set oEdgePort1 = GetLateBindMbrFeaturePort(oFeature1, oPart1, opt1, ctx1, opt1, opr1, -1)
    Set oEdgePort2 = GetLateBindMbrFeaturePort(oFeature2, oPart2, opt2, ctx2, opt2, opr2, -1)
    
    ' -------------------------
    ' Create the corner feature
    ' -------------------------
    If Not oFacePort Is Nothing And Not oEdgePort1 Is Nothing And Not oEdgePort2 Is Nothing Then
        Dim oSystemParent As Object
        If oParent Is Nothing Then
            Set oSystemParent = oFeature1
        Else
            Set oSystemParent = oParent
        End If
        
        If Not oFacePort Is Nothing And Not oEdgePort1 Is Nothing And Not oEdgePort2 Is Nothing Then
            sError = "Creating Member Corner Feature"
            Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
            Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
            oSDO_CornerFeature.Create oResMgr, _
                                      oFacePort, _
                                      oEdgePort1, _
                                      oEdgePort2, _
                                      strSelection, _
                                      oSystemParent
        
            sError = "Returning Member CornerFeature just created"
            
            Set CreateCornerFeatureBetweenTwoEndCuts = oSDO_CornerFeature.object
        
        End If
    End If
                               
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'***********************************************************************
' METHOD:
' CAConstruct_CornerFeatureBetweenTwoCuts
'
' DESCRIPTION:
' Create a corner feature between an endcut on a member and a (skinny) lateral face
' The broad face port is implied by the input member description
'***********************************************************************
Public Function CreateCornerFeatureBetweenEndAndLateralPort(oFeature As IJStructFeature, _
                                                            ByVal oResMgr As IUnknown, _
                                                            lateralPortXid As JXSEC_CODE, _
                                                            IsBottomFlange As Boolean) As Object

    On Error GoTo ErrorHandler

    Dim sMETHOD As String
    Dim sError As String
    sMETHOD = "CreateCornerFeatureBetweenEndAndLateralPort"
    
    ' ------------------------------
    ' Verify the object is an endcut
    ' ------------------------------
    If Not oFeature.get_StructFeatureType = SF_WebCut Or oFeature.get_StructFeatureType = SF_FlangeCut Then
        sError = "Item is not an endcut"
        GoTo ErrorHandler
    End If
    
    Dim oSDOWebCut As New StructDetailObjects.WebCut
    Dim oSDOFlangeCut As New StructDetailObjects.FlangeCut
    Dim oPart As Object
    
    If oFeature.get_StructFeatureType = SF_WebCut Then
        Set oSDOWebCut.object = oFeature
        Set oPart = oSDOWebCut.Bounded
    Else
        Set oSDOFlangeCut.object = oFeature
        Set oPart = oSDOFlangeCut.Bounded
    End If
    
    ' ----------------------------------------------------------------
    ' Determine if feature is on the web, top flange, or bottom flange
    ' ----------------------------------------------------------------
    Dim faceXid As JXSEC_CODE
    
    If oFeature.get_StructFeatureType = SF_WebCut Then
        faceXid = JXSEC_WEB_LEFT
    Else
        If IsBottomFlange Then
            faceXid = JXSEC_BOTTOM
        Else
            faceXid = JXSEC_TOP
        End If
    End If
    
    ' -------------
    ' Get the ports
    ' -------------
    Dim ctx As Long
    Dim opt As Long
    Dim oPR As Long
    
    ' Get the face port, beginning with Current Geometry
    Dim oMemberPart As New StructDetailObjects.MemberPart
    Set oMemberPart.object = oPart
    
    Dim oFacePort As IJStructPort
    Dim oEdgePort1 As IJStructPort
    Dim oEdgePort2 As IJStructPort
    
    Set oFacePort = GetLateralSubPortBeforeTrim(oMemberPart.object, faceXid)
    
    ' Get the information for this end cut
    GetLateBindMbrFeatureData oFeature, oPart, ctx, opt, oPR
    
    ' Replace face port with port after endCut is applied
    Set oFacePort = GetLateBindMbrFeaturePort(oFeature, _
                                              oPart, _
                                              opt, _
                                              oFacePort.ContextID, _
                                              oFacePort.OperationID, _
                                              oFacePort.OperatorID, _
                                              oFacePort.SectionID)
    
    ' The lateral port is the same, but with different operator and xid (which are expected to be the same)
    Set oEdgePort2 = GetLateBindMbrFeaturePort(oFeature, _
                                              oPart, _
                                              opt, _
                                              oFacePort.ContextID, _
                                              oFacePort.OperationID, _
                                              lateralPortXid, _
                                              lateralPortXid)
    
    ' Get the end cut port
    Set oEdgePort1 = GetLateBindMbrFeaturePort(oFeature, oPart, opt, ctx, opt, oPR, -1)
    
    ' -------------------------
    ' Create the corner feature
    ' -------------------------
    If Not oFacePort Is Nothing And Not oEdgePort1 Is Nothing And Not oEdgePort2 Is Nothing Then
        sError = "Creating Member Corner Feature"
        Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
        Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
        oSDO_CornerFeature.Create oResMgr, _
                                  oFacePort, _
                                  oEdgePort1, _
                                  oEdgePort2, _
                                  "MbrEndCutCorner", _
                                  oFeature
    
        sError = "Returning Member CornerFeature just created"
        
        Set CreateCornerFeatureBetweenEndAndLateralPort = oSDO_CornerFeature.object
    
    End If
                               
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Sub NameMemberObject(oBO As Object, strName As String)
    Const METHOD = "NameMemberObject"
    On Error GoTo ErrorHandler
    
    If Not (TypeOf oBO Is IJNamedItem) Then
        Err.Raise LogError(Err, MODULE, METHOD, "Not a valid named object").Number
    End If
    
    Dim oIJNamedItem As IJNamedItem
    Set oIJNamedItem = oBO
    oIJNamedItem.Name = strName
    Set oIJNamedItem = Nothing
    
    Exit Sub
ErrorHandler:

End Sub


'***********************************************************************
' METHOD:
' GetBraceCrossSection
'
' DESCRIPTION:
' The inset brace is expected to be of the same cross section as bounded of the parent as the inset is usuallu cut from the Bounded member.
' However, if the brace cross section is not sufficient to fill the gap between Bounded&Bounding of parent AC, then look for a cross section from the catalog big enough
' to fill the gap.
' The logic in Else case gets the items from the catalog, filters them with the conditions required for the right brace for a given Bounded/Bounding
' Collection1 -- Gets all the "W" items defined in catalog
' Collection2 -- Items of which Depth > (dMinBraceHeight)
' Collection3 -- Items of which FlangeWidth >= FlangeWidth of Bounded of Parent AC
' Collection4 -- Items of which WebThickness >= WebThickness of Bounded of Parent AC
' Collection5 -- Items of which FlangeThickness <= f(FlangeThickness of Bounded of ParentAC)
'
'***********************************************************************
Public Sub GetBraceCrossSection(oBounded As StructDetailObjects.MemberPart, oBounding As StructDetailObjects.MemberPart, _
                        ByRef oCrossSection As IJCrossSection, Optional dMinBraceHeight As Double)
    On Error GoTo ErrorHandler
    Const METHOD = "GetBraceCrossSection"
    
    'Get the FW, WT, FT of bounded
    Dim dFW_Bounded As Double, dWT_Bounded As Double, dFT_Bounded As Double
    
    dFW_Bounded = oBounded.FlangeLength
    dWT_Bounded = oBounded.webThickness
    dFT_Bounded = oBounded.flangeThickness
    
    Dim oRefDataQuery As RefDataMiddleServices.RefdataSOMMiddleServices
    Set oRefDataQuery = New RefDataMiddleServices.RefdataSOMMiddleServices
    
    Dim sSectionName As String
    sSectionName = oBounded.SectionName
    
    Set oCrossSection = oRefDataQuery.GetCrossSection("AISC-LRFD-3.1", sSectionName)
    
    Dim oAttrs As IJDAttributes
    Set oAttrs = oCrossSection
    
    If oAttrs.CollectionOfAttributes("IStructCrossSectionDimensions").Item("Depth").value >= (dMinBraceHeight) Then    'make sure 0.01 is valid
        'retain the Bounded cross section
    Else
        Dim oSRDQuery As IJSRDStructQuery
        Dim oCollection1 As IJDCollection
        
        Set oSRDQuery = New SRDQuery
        Set oCollection1 = oSRDQuery.GetSections("AISC-LRFD-3.1", "W")
    
        If Not oCollection1.Size > 0 Then
            Exit Sub
        End If
    
        Dim oCollection2 As Collection
        Set oCollection2 = New Collection
    
        Dim lCollCount As Long
        lCollCount = oCollection1.Size
    
        Dim i As Integer
        For i = 1 To lCollCount
            Set oCrossSection = oCollection1.Item(i)
            Set oAttrs = oCrossSection
            If oAttrs.CollectionOfAttributes("IStructCrossSectionDimensions").Item("Depth").value > (dMinBraceHeight + 0.01) Then   'make sure 0.01 is valid
                oCollection2.Add oCrossSection
            End If
        Next i
    
        If Not oCollection2.Count > 0 Then
            Set oCrossSection = oCollection1.Item(1)
            Exit Sub
        Else
            Set oCrossSection = Nothing
        End If
    
        Set oCollection1 = Nothing
        
        Dim oCollection3 As Collection
        Set oCollection3 = New Collection
    
        For i = 1 To oCollection2.Count
            Set oCrossSection = oCollection2.Item(i)
            Set oAttrs = oCrossSection
            If oAttrs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("bf").value >= dFW_Bounded Then
                oCollection3.Add oCrossSection
            End If
        Next
    
        If Not oCollection3.Count > 0 Then
            'return the first from yColl
            Set oCrossSection = oCollection2.Item(1)
            Exit Sub
        Else
            Set oCrossSection = Nothing
        End If
    
        Set oCollection2 = Nothing

        Dim oCollection4 As Collection
        Set oCollection4 = New Collection
    
        For i = 1 To oCollection3.Count
            Set oCrossSection = oCollection3.Item(i)
            Set oAttrs = oCrossSection
            If oAttrs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").value >= dWT_Bounded Then
                oCollection4.Add oCrossSection
            End If
        Next
    
        If Not oCollection4.Count > 0 Then
            'return the first from zColl
            Set oCrossSection = oCollection3.Item(1)
            Exit Sub
        Else
            Set oCrossSection = Nothing
        End If
    
        Dim oCollection5 As Collection
        Set oCollection5 = New Collection
    
        For i = 1 To oCollection4.Count
            Set oCrossSection = oCollection4.Item(i)
            Set oAttrs = oCrossSection
            If oAttrs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").value >= dFT_Bounded Then
                oCollection5.Add oCrossSection
            End If
        Next
    
        If Not oCollection5.Count > 0 Then
            'return the first from pColl
            Set oCrossSection = oCollection4.Item(1)
            Exit Sub
        Else
            Set oCrossSection = Nothing
        End If
    
        'by the time it comes here, it is sure that oCollection5 is not empty
        'so take the first item from oCollection5
        Set oCrossSection = oCollection5.Item(1)
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "Could not retrieve CrossSection").Number

End Sub

'***********************************************************************
' METHOD:
' SetAssemblyConnectionNameRule
'
' DESCRIPTION:
' This method sets naming rule to name the AC created as a child of
' the AC; as part of B-52232(Creating InsetBrace)
'
'***********************************************************************
Public Sub SetAssemblyConnectionNameRule(oAsmConn As IJStructAssemblyConnection)
    Const METHOD = "SetAssemblyConnectionNameRule"
    On Error GoTo ErrorHandler

    Dim oNamingRules As IJElements
    Dim oNameRuleHlpr As IJDNamingRulesHelper
    Set oNameRuleHlpr = New NamingRulesHelper

    oNameRuleHlpr.GetEntityNamingRulesGivenProgID "StructConnections.StructAssemblyConnection", oNamingRules
    
    Dim oNameRuleHolder As IJDNameRuleHolder
    Dim oNameRuleAE As IJNameRuleAE

    If oNamingRules.Count > 0 Then
        Set oNameRuleHolder = oNamingRules.Item(1)
    End If

    Call oNameRuleHlpr.AddNamingRelations(oAsmConn, oNameRuleHolder, oNameRuleAE)

    Set oNameRuleHolder = Nothing
    Set oNamingRules = Nothing
    Set oNameRuleAE = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "Error with AC Naming rule").Number
    
End Sub
'*************************************************************************************
'IsCornerFeatureInputsModificatonNeeded()
'           This method dtermines whether the passed CF obj inputs needs any
'           modification or not. If yes modified ports are returned
'           Currently when the user changes question answer from one Codelist to
'           other codelist  at AC/end-cut level there is a chance that
'           input ports get modified.
'
'Inputs : pPropertyDescription ---> property of a MemberDescription of SymbolDefintion
'         oCornerFeature -----> Corner Feature object(for which inputs modification
'                                                      needs to be checked)
'
'Outputs :oModifiedFacePort,oModifiedEdgePort1,oModifiedEdgePort2
'
'if Corner Featire Modification is needed then modified(new) ports are to be returned
'otherwise Nothing(null) is returned
'
'
'**************************************************************************************
Public Function IsCornerFeatureInputsModificatonNeeded(pPropertyDescription As IJDPropertyDescription, _
                                                        oCornerFeature As IJStructFeature, _
                                                        oModifiedFacePort As IJPort, _
                                                        oModifiedEdgePort1 As IJPort, _
                                                        oModifiedEdgePort2 As IJPort) As Boolean
    Const sMETHOD = "IsCornerFeatureInputsModificatonNeeded"
    On Error GoTo ErrorHandler
    
    Dim bTopEdgeCorner As Boolean
    IsCornerFeatureInputsModificatonNeeded = False
    
    '---------------------------------------------------------
    'Need to check for which properties mofdification
    'is needed currently we need only for Top/Btm Edge Corners
    'but not Face Top/Btm Corner
    '---------------------------------------------------------
    If StrComp(pPropertyDescription.Name, "NeedToComputeCF3", vbTextCompare) = 0 Then
        'continue
        bTopEdgeCorner = True
    ElseIf StrComp(pPropertyDescription.Name, "NeedToComputeCF4", vbTextCompare) = 0 Then
        'contine
        bTopEdgeCorner = False
    Else
       IsCornerFeatureInputsModificatonNeeded = False
       Set oModifiedFacePort = Nothing
       Set oModifiedEdgePort1 = Nothing
       Set oModifiedEdgePort2 = Nothing
       Exit Function
    End If

    Dim sACItemName As String
    Dim sShapeAtEdge As String
    Dim oACObj As Object
    
    '---------------------------------------------------------
    'Get the question answer at AC level....
    '---------------------------------------------------------
    Parent_SmartItemName pPropertyDescription.CAO, sACItemName, oACObj
    Select Case LCase(sACItemName)
    Case LCase(gsMbrAxisToEdgeAndOutSide2Edge), (gsStiffEndToMbrEdgeAndOutSide2Edge)
        GetSelectorAnswer oACObj, "ShapeAtEdgeOverlap", sShapeAtEdge
    Case LCase(gsMbrAxisToFaceAndOutSide1Edge), LCase(gsMbrAxisToOutSideAndOutSide1Edge), _
         LCase(gsStiffEndToMbrFaceAndOutSide1Edge), LCase(gsStiffEndToMbrOutSideAndOutSide1Edge)
        GetSelectorAnswer oACObj, "ShapeAtEdge", sShapeAtEdge
    Case LCase(gsMbrAxisToOutSideAndOutSide2Edge), LCase(gsStiffEndToMbrOutSideAndOutSide2Edge)
        If bTopEdgeCorner Then
            GetSelectorAnswer oACObj, "ShapeAtTopEdge", sShapeAtEdge
        Else
            GetSelectorAnswer oACObj, "ShapeAtBottomEdge", sShapeAtEdge
        End If
    End Select

    Dim oFeatureUtils As IJSDFeatureAttributes
    Dim oFacePort As IJPort
    Dim oEdgePort1 As IJPort
    Dim oEdgePort2 As IJPort
    Dim oStructEdgePort1 As IJStructPort
    Dim oStructEdgePort2 As IJStructPort
    
    'get the CF inputs
    Set oFeatureUtils = New SDFeatureUtils
    oFeatureUtils.get_CornerCutInputs oCornerFeature, oFacePort, oEdgePort1, oEdgePort2
    
    Set oStructEdgePort1 = oEdgePort1
    Set oStructEdgePort2 = oEdgePort2
    
    Dim eXidEdgePort1 As JXSEC_CODE
    Dim eXidEdgePort2 As JXSEC_CODE
    Dim eNewEdgePort1 As JXSEC_CODE
    Dim eNewEdgePort2 As JXSEC_CODE
    
    eXidEdgePort1 = oStructEdgePort1.SectionID
    eXidEdgePort2 = oStructEdgePort2.SectionID
    
    '--------------------------------------------------------------------
    'if question's answer results CF of different ports to what currently
    'CF has then we need to Modfiy the current CF to new Ports
    '--------------------------------------------------------------------
    If StrComp(sShapeAtEdge, gsEdgeToOutside, vbTextCompare) = 0 Or _
        StrComp(sShapeAtEdge, gsInsideToOutsideCorner, vbTextCompare) = 0 Then
        
        If eXidEdgePort1 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
           eXidEdgePort2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
            IsCornerFeatureInputsModificatonNeeded = True
            eNewEdgePort1 = JXSEC_TOP
            eNewEdgePort2 = JXSEC_TOP_FLANGE_RIGHT
        ElseIf eXidEdgePort1 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
           eXidEdgePort2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
            IsCornerFeatureInputsModificatonNeeded = True
            eNewEdgePort1 = JXSEC_BOTTOM
            eNewEdgePort2 = JXSEC_BOTTOM_FLANGE_RIGHT
        Else
            IsCornerFeatureInputsModificatonNeeded = False
        End If
        
    Else
        If eXidEdgePort1 = JXSEC_TOP Or _
           eXidEdgePort2 = JXSEC_TOP Then
            IsCornerFeatureInputsModificatonNeeded = True
            eNewEdgePort1 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
            eNewEdgePort2 = JXSEC_TOP_FLANGE_RIGHT
        ElseIf eXidEdgePort1 = JXSEC_BOTTOM Or _
           eXidEdgePort2 = JXSEC_BOTTOM Then
            IsCornerFeatureInputsModificatonNeeded = True
            eNewEdgePort1 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
            eNewEdgePort2 = JXSEC_BOTTOM_FLANGE_RIGHT
        Else
            IsCornerFeatureInputsModificatonNeeded = False
        End If
    
    End If

    If IsCornerFeatureInputsModificatonNeeded Then
        MbrCornerFeatureDataByObject pPropertyDescription.CAO, eNewEdgePort1, _
                                eNewEdgePort2, oModifiedFacePort, _
                                oModifiedEdgePort1, oModifiedEdgePort2
    End If

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, "CornerFeature Inputs Modificaton Failed").Number
    
End Function
'***********************************************************************
' METHOD:
' MbrCornerFeatureDataByObject
'
' DESCRIPTION:
'      Similar implemenation as of MbrCornerFeatureData just argument has changed
'      to Corner Feature Object
'
'***********************************************************************
Public Sub MbrCornerFeatureDataByObject(ByVal oEndCutObject As Object, _
                              portXid1 As JXSEC_CODE, _
                              portXid2 As JXSEC_CODE, _
                              oFacePort As IJPort, _
                              oEdgePort1 As IJPort, _
                              oEdgePort2 As IJPort)
    Dim sMETHOD As String
    Dim sError As String
    sMETHOD = "MbrCornerFeatureDataByObject"
    
    On Error GoTo ErrorHandler
    
    'Determine feature type and fill bounded and bounding ports
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Dim oEndCutBoundedPort As IJPort
    Dim oBoundedPart As Object

    Set oFeature = oEndCutObject
    eFeatureType = oFeature.get_StructFeatureType
    Select Case eFeatureType
        Case SF_WebCut
            Dim oWebCut As New StructDetailObjects.WebCut
            Set oWebCut.object = oFeature
            Set oEndCutBoundedPort = oWebCut.BoundedPort
            Set oBoundedPart = oWebCut.Bounded
        Case SF_FlangeCut
            Dim oFlangeCut As New StructDetailObjects.FlangeCut
            Set oFlangeCut.object = oFeature
            Set oEndCutBoundedPort = oFlangeCut.BoundedPort
            Set oBoundedPart = oFlangeCut.Bounded
        Case Else
            GoTo ErrorHandler
    End Select
    
    ' Verify the following objects are valid
    If (oEndCutBoundedPort.Connectable Is Nothing) Then
        sError = "oEndCutBoundedPort.Connectable Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    If (oBoundedPart Is Nothing) Then
        sError = "oBoundedPart Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    If eFeatureType = SF_WebCut Then
        'Web penetrated case
        Set oFacePort = GetLateralSubPortBeforeTrim(oEndCutBoundedPort.Connectable, JXSEC_WEB_LEFT)
    Else
        'Flange penetrated case: check if this CF is on top and bottom flange
        Dim bBottomFlange As Boolean
        Dim strAnswer As String
        
        GetSelectorAnswer oEndCutObject, "BottomFlange", strAnswer
        If strAnswer = "Yes" Then
            Set oFacePort = GetLateralSubPortBeforeTrim(oEndCutBoundedPort.Connectable, JXSEC_BOTTOM)
        Else
            Set oFacePort = GetLateralSubPortBeforeTrim(oEndCutBoundedPort.Connectable, JXSEC_TOP)
        End If
    End If

    GetMbrEndCutCornerPorts oEndCutObject, oBoundedPart, portXid1, portXid2, oFacePort, oEdgePort1, oEdgePort2
    
    Exit Sub
ErrorHandler:
      Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'*************************************************************************************
'GetNearestBoundingToPort()
'           This method dtermines the nearest bounding to the given input port(Bounded port)
'
'Inputs : oBoundingObjColl----> Bounding object collection
'         oAppConnection --->Connection between the bounded and bounding
'         oPort ---> Bounded Port to get the nearest Bounding
'
'Outputs : GetNearestBoundingToPort ---> Bounding nearer to the input port
'
'
'**************************************************************************************
Public Function GetNearestBoundingToPort(ByVal oBoundingObjColl As IJElements, ByVal oEditJDArgument As IJDEditJDArgument, _
                                        ByVal oAppConnection As Object, ByVal oPort As IJPort) As Object

    Const METHOD = "GetNearestBoundingToPort"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim iCount As Integer
    Dim lStatus As Long
    Dim dDistance As Double
    Dim oEndNormal As IJDVector
    Dim oEndSurface As IJSurfaceBody
    Dim oEndPort As IJPort
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oPointOnEnd As IJDPosition
    Dim oSGOModelUtil As IJSGOModelBodyUtilities
    Dim oPortGeometry As IJSurfaceBody
    Dim oStructGeomUtils As GSCADStructGeomUtilities.PartInfo
    
    Set oSGOModelUtil = New SGOModelBodyUtilities
    Set oStructGeomUtils = New PartInfo
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    Set oEndPort = Member_GetSolidPort(oBoundedData.AxisPort, True)
       
    'Get EndPort Normal
    Set oEndNormal = oStructGeomUtils.GetPortNormal(oEndPort, True)
    Set oPortGeometry = GetExtendedPort(oPort)
    Set oEndSurface = oEndPort.Geometry
    ' ---------------------------------
    ' Loop through all Boundings
    ' ---------------------------------
    Dim oIntersect As IJDTopologyIntersect
    Dim oCommonGeom As Object
    Dim oPointOnCommon As IJDPosition
    Dim oDirVector As IJDVector
    Dim maxDist As Double
    
    Set oIntersect = New DGeomOpsIntersect
    maxDist = -1000000#

    For iCount = 1 To oBoundingObjColl.Count

        Dim oBoundingPort As IJPort
        Dim oBoundingPortGeom As Object
        
        Set oCommonGeom = Nothing
        
        ' -----------------------------------
        ' Check if it is a profile or member if it is use the global lateral port
        ' If it is a plate use the web cut port
        ' -----------------------------------
        If TypeOf oBoundingObjColl.Item(iCount) Is IJProfile Then
            'Profile
            Dim oProfilePart As New StructDetailObjects.ProfilePart
            Set oProfilePart.object = oBoundingObjColl.Item(iCount)
            Set oBoundingPort = oProfilePart.BasePort(BPT_Lateral)
            Set oBoundingPortGeom = oBoundingPort.Geometry
        ElseIf TypeOf oBoundingObjColl.Item(iCount) Is ISPSMemberPartPrismatic Then
            'Member
            Dim oMemberPart As New StructDetailObjects.MemberPart
            Set oMemberPart.object = oBoundingObjColl.Item(iCount)
            Set oBoundingPort = oMemberPart.BasePortBeforeTrim(BPT_Lateral)
            Set oBoundingPortGeom = oBoundingPort.Geometry
        Else
            'Plate
            Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjColl.Item(iCount), oEditJDArgument).Item(1)
            Set oBoundingPortGeom = oBoundingPort.Geometry
        End If

        If TypeOf oBoundingPortGeom Is IJSurfaceBody Then
            On Error Resume Next
            oIntersect.PlaceIntersectionObject Nothing, oPortGeometry, oBoundingPortGeom, Nothing, oCommonGeom
            On Error GoTo ErrorHandler
        End If

         ' ----------------------------------------------
        ' Get distance from intersection to the end port
        ' ----------------------------------------------
        dDistance = -1000000#
        If Not oCommonGeom Is Nothing Then
            If TypeOf oBoundingObjColl.Item(iCount) Is IJProfile Or TypeOf oBoundingObjColl.Item(iCount) Is ISPSMemberPartPrismatic Then
                ' If the bounding object is a stiffener or member then loop through all the vertices on the intersecting
                ' geometry to get the vertex that is the furthest from the bounded member End Surface
                Dim oVertexColl As Collection
                oSGOModelUtil.GetVertices oCommonGeom, oVertexColl

                Dim dVertexDist As Double
                Dim dVertexMaxDist As Double
                Dim oVertexPos As IJDPosition
                Dim oClosestPoint As IJDPosition

                dVertexMaxDist = -1000000# 'Initial value

                For Each oVertexPos In oVertexColl
                    oSGOModelUtil.GetClosestPointOnBody oEndSurface, oVertexPos, oClosestPoint, dVertexDist
                    If GreaterThan(dVertexDist, dVertexMaxDist) Then
                        dVertexMaxDist = dVertexDist
                        Set oPointOnCommon = oVertexPos
                        Set oPointOnEnd = oClosestPoint
                    End If
                Next oVertexPos
                dDistance = dVertexMaxDist

            Else
                'If it is a plate, get the distance from the EndSurface to the intersecting geometry
                oSGOModelUtil.GetClosestPointsBetweenTwoBodies oEndSurface, oCommonGeom, oPointOnEnd, oPointOnCommon, dDistance
            End If

            If GreaterThan(dDistance, 0.00001) Then   '0.001 mm
                Set oDirVector = oPointOnCommon.Subtract(oPointOnEnd)
                If GreaterThanZero(oDirVector.Dot(oEndNormal)) Then
                    dDistance = -dDistance
                End If
            End If
        End If

        ' -------------------------------
        ' Save the boundary furthest away
        ' -------------------------------
        If GreaterThan(dDistance, maxDist) Then
            maxDist = dDistance
            Set GetNearestBoundingToPort = oBoundingObjColl.Item(iCount)
        End If
    Next

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

Public Function HasTopFlange(oPart As Object, Optional bHasLeft As Boolean = False, Optional bHasRight As Boolean = False) As Boolean

    HasTopFlange = False

    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oPart, bTFL, bBFL, bTFR, bBFR
    
    If bTFL Then
        bHasLeft = True
    End If
    
    If bTFR Then
        bHasRight = True
    End If
    
    Dim oInnerPort As IJPort
    Set oInnerPort = GetLateralSubPortBeforeTrim(oPart, JXSEC_INNER_WEB_RIGHT)
    
    If bTFL Or bTFR Or (Not oInnerPort Is Nothing) Then
        HasTopFlange = True
    End If

End Function

Public Function HasBottomFlange(oPart As Object, Optional bHasLeft As Boolean = False, Optional bHasRight As Boolean = False) As Boolean

    HasBottomFlange = False

    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oPart, bTFL, bBFL, bTFR, bBFR
    
    If bBFL Then
        bHasLeft = True
    End If
    
    If bBFR Then
        bHasRight = True
    End If
    
    Dim oInnerPort As IJPort
    Set oInnerPort = GetLateralSubPortBeforeTrim(oPart, JXSEC_INNER_WEB_RIGHT)
    
    If bBFL Or bBFR Or (Not oInnerPort Is Nothing) Then
        HasBottomFlange = True
    End If

End Function
'*************************************************************************************
'GetSectionType()
'           This method gets the SectionType for a member or a stiffener
'
'Inputs : port as IJPort
'
'Outputs : SectionType
'
'**************************************************************************************
Public Function GetSectionType(oPort As IJPort) As String
    Dim strSectionType As String
    strSectionType = ""
    If Not oPort Is Nothing Then
        If TypeOf oPort.Connectable Is ISPSMemberPartCommon Then
            Dim oSDOMember As New StructDetailObjects.MemberPart
            Set oSDOMember.object = oPort.Connectable
            strSectionType = oSDOMember.sectionType
        ElseIf TypeOf oPort.Connectable Is IJProfile Then
            Dim oSDOStiffnerPart As New StructDetailObjects.ProfilePart
            Set oSDOStiffnerPart.object = oPort.Connectable
            strSectionType = oSDOStiffnerPart.sectionType
        End If
    End If
    GetSectionType = strSectionType
End Function

Public Function GetMemberCrossSectionName(oMemberObject As Object) As String
Const METHOD = "::GetMemberCrossSectionName"
    
    Dim sMsg As String
    On Error GoTo ErrorHandler
    
    'Validate
    If oMemberObject Is Nothing Then Exit Function
    
    Dim sCStype As String
    
    Dim oPort As IJPort
    Dim oMemberPart As ISPSMemberPartCommon
    Dim oSplitAxisPort As ISPSSplitAxisPort
    
    Dim oCrossSection As IJCrossSection
    Dim oPartDesigned As ISPSDesignedMember
    Dim oPartPrismatic As ISPSMemberPartPrismatic
    Dim oSPSCrossSection As ISPSCrossSection
    Dim oStiffenerPart As StructDetailObjects.ProfilePart
    
    Set oStiffenerPart = New StructDetailObjects.ProfilePart
    
    Dim bIsBuiltup As Boolean
    Dim oBUMember As ISPSDesignedMember
    IsFromBuiltUpMember oMemberObject, bIsBuiltup, oBUMember
    Dim otempObj As Object
    
    Set otempObj = oMemberObject
    
    If bIsBuiltup Then
        Set oMemberObject = oBUMember
    End If
    
    If TypeOf oMemberObject Is ISPSSplitAxisPort Then
        Set oPort = oMemberObject
        Set oSplitAxisPort = oMemberObject
        Set oMemberPart = oPort.Connectable
    
    ElseIf TypeOf oMemberObject Is ISPSMemberPartCommon Then
        Set oMemberPart = oMemberObject
    ElseIf TypeOf oMemberObject Is IJProfile Then
        Set oStiffenerPart.object = oMemberObject
        sCStype = oStiffenerPart.sectionType
    Else
        sCStype = ""
        GoTo CleanUp
    End If
    If Not oMemberPart Is Nothing Then
        If oMemberPart.IsPrismatic Then
            Set oPartPrismatic = oMemberPart
            Set oSPSCrossSection = oPartPrismatic.CrossSection
        
        ElseIf TypeOf oMemberPart Is ISPSDesignedMember Then
            Set oPartDesigned = oMemberPart
            Set oSPSCrossSection = oMemberPart
        ElseIf Not oMemberPart Is Nothing Then
            sCStype = ""
            GoTo CleanUp
        End If
    End If
        
    ' Verify Bounded have valid Cross Section Type
    If Not oSPSCrossSection Is Nothing Then
        If TypeOf oSPSCrossSection.definition Is IJCrossSection Then
            Set oCrossSection = oSPSCrossSection.definition
            sCStype = oCrossSection.Type
        Else
            sCStype = ""
        End If
    End If
    GetMemberCrossSectionName = sCStype

CleanUp:
    Set oMemberObject = otempObj
    Set otempObj = Nothing
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function
Public Function HasRightWeb(oPart As Object) As Boolean

    HasRightWeb = False
    
    Dim oInnerPort As IJPort
    Set oInnerPort = GetLateralSubPortBeforeTrim(oPart, JXSEC_INNER_WEB_RIGHT)
    
    If Not oInnerPort Is Nothing Then
        HasRightWeb = True
    End If
    
End Function
' The Standard Member AC definition accounts for separate left and right web cuts, even though they are not
' currently supported.  If there is only one web, the cut specified for the left web is applied.
' The rules use this method to determine whether or not to apply the specified right web cut.
' This method currently returns false in all cases.  If/when separate web cuts are allowed, it should be
' modified by un-commenting the code indicated below.
Public Function ConsiderRightWebAttributes(oPart As Object) As Boolean

    ConsiderRightWebAttributes = False
    
    ' Uncomment the code below if/when the system allows for separate cuts on left and right webs
'    If HasRightWeb(oPart) Then
'        ConsiderRightWebAttributes = True
'    End If
    
End Function

'*************************************************************************
'Function
'Create_BearingPlate
'
'
'Abstract
'   Create BearingPlate :
'        Given the MemberDescription and Root Selection Rule
'
'input
'   pMemberDescription
'   pResourceManager
'   sEndCutSelRule
'   bUseBoundingEndPort
'
'Return
'   pEndCutObject
'
'Exceptions
'
'***************************************************************************
Public Sub Create_BearingPlate(pMemberDescription As IJDMemberDescription, _
                               pResourceManager As IUnknown, _
                               sEndCutSelRule As String, _
                               bUseBoundingEndPort As Boolean, _
                               pBearingPlateObject As Object)
Const METHOD = "MbrAssemblyUtilities::Create_BearingPlate"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Dim oBearingPlate As IJSmartPlate
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oGraphicInputs As JCmnShp_CollectionAlias

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSPDefinition As GSCADSDCreateModifyUtilities.IJSDSmartPlateDefinition

    sMsg = "Creating WebCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                             lStatus, sMsg
    
    Set oBoundedPort = oBoundedData.AxisPort
    Set oBoundingPort = oBoundingData.AxisPort
    
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection
    Dim oBoundingObjectColl As IJElements
    
    If oBoundingPort Is Nothing Then
        'Possible Axis AC
        Dim oPortElements As IJElements
        oAppConnection.enumPorts oPortElements
        Set oBoundingPort = oPortElements.Item(2)
        
    ElseIf oBoundedPort.Connectable Is oBoundingPort.Connectable Then
        'Generic AC
        Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
        Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
        ' -------------------------------------------------------------
        ' Get all the bounding objects from the ports related to the AC
        ' -------------------------------------------------------------

        Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(1), oEditJDArgument).Item(1)
        
    End If
    
    If bUseBoundingEndPort Then
        GetSupportingEndPort oBoundedData, oBoundingData, oBoundingPort
    End If
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for WebCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    Dim oNearBoundingBUPort As New StructDetailObjects.PlatePart
    If TypeOf oBoundingPort.Connectable Is ISPSDesignedMember Then
        
        Dim oBoundingBU As ISPSDesignedMember
        Set oBoundingBU = oBoundingPort.Connectable
        Dim oBoundedStart As IJPoint
        Dim oBoundedEnd As IJPoint
        Dim oBoundedRefPos As New DPosition
        Dim dx As Double
        Dim dy As Double
        Dim dz As Double
   
        If TypeOf oBoundedPort Is ISPSSplitAxisPort Then
            Dim oBoundedMemberPort As ISPSSplitAxisPort
            Set oBoundedMemberPort = oBoundedPort
            Dim oBoundedMemberPart As ISPSMemberPartCommon
            Set oBoundedMemberPart = oBoundedPort.Connectable
            If oBoundedMemberPort.PortIndex = SPSMemberAxisStart Then
                Set oBoundedEnd = oBoundedMemberPart.PointAtEnd(SPSMemberAxisEnd)
                oBoundedEnd.GetPoint dx, dy, dz
                oBoundedRefPos.Set dx, dy, dz
            ElseIf oBoundedMemberPort.PortIndex = SPSMemberAxisEnd Then
                Set oBoundedStart = oBoundedMemberPart.PointAtEnd(SPSMemberAxisStart)
                oBoundedStart.GetPoint dx, dy, dz
                oBoundedRefPos.Set dx, dy, dz
            End If
            GetNearestboundingBUPort oBoundedRefPos, oBoundedPort, oBoundingBU, oBoundingPort
        End If
    
        
    End If

    
    ' Create the Bearing Plate
    sMsg = "...Creating Bearing Plate Object"
    Set oGraphicInputs = New Collection
    oGraphicInputs.Add oBoundingPort
    oGraphicInputs.Add oBoundedPort
    
    Set oSPDefinition = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    Set oBearingPlate = oSPDefinition.CreateBearingPlatePart(pResourceManager, _
                                                             sEndCutSelRule, _
                                                             oGraphicInputs, _
                                                             oSystemParent)
    
    sMsg = "...Setting Bearing Plate Properties"
    SetPlatePartProperties oBearingPlate, pResourceManager, _
                           "CPlatePart", "BearingPlateCategory", _
                           NonTight, Standalone, _
                           "Steel - Carbon", "A", 0.01
            
    sMsg = "Return the created Bearing Plate"
    Set pBearingPlateObject = oBearingPlate
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'SetPlatePartProperties
'
'Abstract
'   Initialize/Set required PlatePart Properties
'
'input
'   oBearingPlate
'   oResourceManager
'   strEntity
'   strNamingCategoryTable
'   eTightness
'   ePlateType
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub SetPlatePartProperties(oBearingPlate As Object, _
                    oResourceManager As IUnknown, _
                    strEntity As String, _
                    strNamingCategoryTable As String, _
                    eTightness As GSCADShipGeomOps.StructPlateTightness, _
                    ePlateType As GSCADShipGeomOps.StructPlateType, _
                    strMatl As String, _
                    strGrade As String, _
                    dThickness As Double)
Const METHOD = "MbrAssemblyUtilities::SetPlatePartProperties"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim strLongNames() As String
    Dim strShortNames() As String

    Dim iIndex As Long
    Dim lPriority() As Long
    
    Dim oRules As IJElements
    Dim oDummyAE As IJNameRuleAE
    Dim oQueryUtil As IJMetaDataCategoryQuery
    Dim oNamingObject As IJDNamingRulesHelper
    
    Dim oPlate As IJPlate
    Set oPlate = oBearingPlate
    
    'Need to assign naming rule for a new bearing plate
    If TypeOf oBearingPlate Is IJNamedItem Then
        Dim oNamedItem As IJNamedItem
        Set oNamedItem = oBearingPlate
        If Trim(oNamedItem.Name) = vbNullString Then
            'Retrieve first default naming rule
            Set oNamingObject = New NamingRulesHelper
            oNamingObject.GetEntityNamingRulesGivenName strEntity, oRules
    
            If oRules.Count >= 1 Then
                oNamingObject.AddNamingRelations oBearingPlate, oRules.Item(1), oDummyAE
            End If
            Set oDummyAE = Nothing
            Set oNamingObject = Nothing
            
            ' Default naming category to first non-negative value
            Set oQueryUtil = New CMetaDataCategoryQuery
            oQueryUtil.GetCategoryInfo oResourceManager, _
                                       strNamingCategoryTable, _
                                       strLongNames, _
                                       strShortNames, _
                                       lPriority
            Set oQueryUtil = Nothing
            oPlate.NamingCategory = -1
            For iIndex = LBound(lPriority) To UBound(lPriority)
                If lPriority(iIndex) >= 0 Then
                    oPlate.NamingCategory = lPriority(iIndex)
                    Exit For
                End If
            Next iIndex
        
            Erase strLongNames
            Erase strShortNames
            Erase lPriority
        End If
    End If
            
    'Set Plate Type
    'Set Plate Tightness
    oPlate.plateType = ePlateType
    oPlate.Tightness = eTightness
            
    'Set Plate Material, Grade and Thickness
    SetMatlGradeThickness oBearingPlate, strMatl, strGrade, dThickness
  
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'SetMatlGradeThickness
'
'Abstract
'   set Plate's Material Type, Grade, and Thickness
'
'input
'   oStructMaterial
'   strMatl
'   strGrade
'   dThickness
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub SetMatlGradeThickness(oStructMaterial As IJStructureMaterial, _
                                 strMatl As String, _
                                 strGrade As String, _
                                 dThickness As Double)
Const METHOD = "MbrAssemblyUtilities::InitializePlatePartProperties"
    Dim sMsg As String
  
    Dim iIndex As Long
    Dim nCount As Long
    
    Dim oPlate As IJPlate
    Dim oMatlObj As IJDMaterial
    Dim oPlateDims As IJDPlateDimensions
    Dim oRefDataQuery As RefDataMiddleServices.RefdataSOMMiddleServices
    Dim matlThickCol As IJDCollection
    Dim oMatProxy As Object
    Dim oPlateDimProxy As Object
    On Error GoTo ErrorHandler

    Dim oResMgr As IJDPOM

    Set oResMgr = GetResourceMgr
    
    
    Set oRefDataQuery = New RefDataMiddleServices.RefdataSOMMiddleServices
    Set oMatlObj = oRefDataQuery.GetMaterialByGrade(strMatl, strGrade)
    
    Set oMatProxy = oResMgr.GetProxy(oMatlObj)
    
    oStructMaterial.Material = oMatProxy
    
    If Not TypeOf oStructMaterial Is IJPlate Then
        Exit Sub
    End If
    
    Set oPlate = oStructMaterial
    Set matlThickCol = oRefDataQuery.GetPlateDimensions(oMatlObj.MaterialType, _
                                                        oMatlObj.MaterialGrade)
  
    nCount = matlThickCol.Size
    For iIndex = 1 To nCount
        Set oPlateDims = matlThickCol.Item(iIndex)
        If Abs(oPlateDims.thickness - dThickness) < 0.000005 Then
            Exit For
        End If
        Set oPlateDims = Nothing
    Next
    
    If oPlateDims Is Nothing Then
        Set oPlateDims = matlThickCol.Item(1)
    End If
    
    Set oPlateDimProxy = oResMgr.GetProxy(oPlateDims)
    oPlate.Dimensions = oPlateDimProxy
    Set oPlateDims = Nothing
    Set oPlateDimProxy = Nothing
    Set oRefDataQuery = Nothing
  
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Function GetResourceMgr() As IJDPOM

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

'*************************************************************************
' Function
' GetAssemblyConnectionInputs
'
' Returns Bounding and Bounded ports given the Assembly Connection.
'
' Abstract
'
'***************************************************************************
Public Sub GetAssemblyConnectionInputs(oAppConnection As IJAppConnection, oBoundedOrPenetratingPort As IJPort, oBoundingOrPenetratedPort As IJPort, Optional bFlip As Boolean = True)
                                    
    Const METHOD = "::GetAssemblyConnectionInputs"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lCount As Long
    Dim oACPorts As IJElements
    
    oAppConnection.enumPorts oACPorts
    lCount = oACPorts.Count
        
    If lCount <> 2 Then
        sMsg = "Member Assembly Connection requires two(2) Ports"
        GoTo ErrorHandler
    End If
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    
    Set oPort1 = oACPorts.Item(1)
    Set oPort2 = oACPorts.Item(2)
    
    If bFlip Then
        GetFlippedPorts oAppConnection, oPort1, oPort2
    End If
    
    If TypeOf oPort1 Is ISPSSplitAxisEndPort Then
        Set oBoundedOrPenetratingPort = oPort1
        Set oBoundingOrPenetratedPort = oPort2
    ElseIf TypeOf oPort2 Is ISPSSplitAxisEndPort Then
        Set oBoundedOrPenetratingPort = oPort2
        Set oBoundingOrPenetratedPort = oPort1
    ElseIf TypeOf oPort1 Is ISPSSplitAxisAlongPort And TypeOf oPort2 Is ISPSSplitAxisAlongPort Then
        GetBoundedAndBounding_ForBothAlongPorts oPort1, oPort2, oAppConnection, oBoundedOrPenetratingPort, oBoundingOrPenetratedPort
    ElseIf TypeOf oPort1 Is IJPort Or TypeOf oPort2 Is IJPort Then
        ' Here stiffeners and plates are handled
        Dim oStructPort1 As IJStructPort
        Dim oStructPort2 As IJStructPort
        Dim eStructPort1Context As eUSER_CTX_FLAGS
        Dim eStructPort2Context As eUSER_CTX_FLAGS
        Dim oSplitAxisPort1 As ISPSSplitAxisPort
        Dim oSplitAxisPort2 As ISPSSplitAxisPort
        Dim ePortId1 As SPSMemberAxisPortIndex
        Dim ePortId2 As SPSMemberAxisPortIndex
        
        'inittialize
        eStructPort1Context = CTX_INVALID
        eStructPort2Context = CTX_INVALID
        
        'From two objects paased as arguments, we will try to get its context id from Ports
        'through which we can determine which port obj is Bounded End Port
        If TypeOf oPort1.Connectable Is IJProfile Then
            Set oStructPort1 = oPort1
            If Not oStructPort1 Is Nothing Then eStructPort1Context = oStructPort1.ContextID
        End If
        If TypeOf oPort2.Connectable Is IJProfile Then
             Set oStructPort2 = oPort2
             If Not oStructPort2 Is Nothing Then eStructPort2Context = oStructPort2.ContextID
        End If
        If TypeOf oPort1.Connectable Is IJProfile And (eStructPort1Context = CTX_BASE Or eStructPort1Context = CTX_OFFSET) Then
            Set oBoundedOrPenetratingPort = oPort1
            Set oBoundingOrPenetratedPort = oPort2
        ElseIf TypeOf oPort2.Connectable Is IJProfile And (eStructPort2Context = CTX_BASE Or eStructPort2Context = CTX_OFFSET) Then
            Set oBoundedOrPenetratingPort = oPort2
            Set oBoundingOrPenetratedPort = oPort1
        End If
    Else
        Set oBoundedOrPenetratingPort = oPort1
        Set oBoundingOrPenetratedPort = oPort2
    End If

    Dim oRefPortColl As Collection
    Dim oReferencesCollection As IJDReferencesCollection
    Dim oBdngObjectColl As IJElements
    Dim oEditJDArgument As IJDEditJDArgument
    'If the assembly connection is generic, then the bounding and bounded ports are same.
    'If there is only one bounding object, then the bounding port is replaced with the referenceport collection object.
    If (GetMbrAssemblyConnectionType(oAppConnection) = ACType_Mbr_Generic) Or (GetMbrAssemblyConnectionType(oAppConnection) = ACType_Stiff_Generic) Then
        GetRefPortColl oAppConnection, oRefPortColl
        Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
        Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
        Set oBdngObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
        If oBdngObjectColl.Count = 1 Then
            Set oBoundingOrPenetratedPort = oRefPortColl.Item(1)
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'   GetLogicalBoundedAndBounding_ForBothAlongPorts
'
'Abstract
'   Bounded and Bounding are decided based on
'       In ladder, the hoop is the bounded
'       In handrail, the post is the bounded
'       Based on the plane normal defined by the two member axes
'input
'   oConnectionObject1, oConnectionObject2
'Return
'   oBoundedPort, oBoundingPort
'
'***************************************************************************
Public Sub GetLogicalBoundedAndBounding_ForBothAlongPorts(oConnectionObject1 As Object, oConnectionObject2 As Object, _
                                        oBoundedPort As Object, oBoundingPort As Object)

Const METHOD = "::GetLogicalBoundedAndBounding_ForBothAlongPorts"
    On Error GoTo ErrorHandler
    
    Dim lCount As Long
    Dim sMsg  As String
    ' require two Ports
    sMsg = "Checking Member Assembly Connection Ports"
    
    If oConnectionObject1 Is Nothing Or oConnectionObject2 Is Nothing Then
        GoTo ErrorHandler
    End If
    
    lCount = 2
        
    Dim oPort As IJPort
    Dim oMemberPart1 As ISPSMemberPartPrismatic
    Dim oMemberPart2 As ISPSMemberPartPrismatic
    
    Set oPort = oConnectionObject1
    Set oMemberPart1 = oPort.Connectable
    
    Set oPort = oConnectionObject2
    Set oMemberPart2 = oPort.Connectable
    
    Dim bFirstIsBounded As Boolean
    Dim bSecondIsBounded As Boolean
    
    bFirstIsBounded = False
    bSecondIsBounded = False
    
    'Check if Member part is Ladder Hoop or Handrail Post
    If (oMemberPart1.MemberType.TypeCategory = MemberCategoryAndType.LadderElement And oMemberPart1.MemberType.Type = MemberCategoryAndType.LHoop) Or _
        (oMemberPart1.MemberType.TypeCategory = MemberCategoryAndType.HandRailElement And oMemberPart1.MemberType.Type = MemberCategoryAndType.HRPost) Then
        
        bFirstIsBounded = True
        
    ElseIf (oMemberPart2.MemberType.TypeCategory = MemberCategoryAndType.LadderElement And oMemberPart2.MemberType.Type = MemberCategoryAndType.LHoop) Or _
        (oMemberPart2.MemberType.TypeCategory = MemberCategoryAndType.HandRailElement And oMemberPart2.MemberType.Type = MemberCategoryAndType.HRPost) Then
        bSecondIsBounded = True
        
    Else 'If not Hoop or Post
        Dim oCurve1 As IJCurve
        Dim oCurve2 As IJCurve
        Dim oIntersection As IJDPosition
        
        Set oCurve1 = oMemberPart1.Axis
        Set oCurve2 = oMemberPart2.Axis
    
        Dim oPositionOnCurve1 As IJDPosition
        Dim oPositionOnCurve2 As IJDPosition
        Dim dx As Double, dy As Double, dz As Double
        Dim dx1 As Double, dy1 As Double, dz1 As Double
        Dim dX2 As Double, dY2 As Double, dZ2 As Double
        Dim dMinDist As Double
        
        Set oPositionOnCurve1 = New DPosition
        Set oPositionOnCurve2 = New DPosition
        
        'Get the Distance between Bounding and Bounded Axis curves
        oCurve1.DistanceBetween oCurve2, dMinDist, dx1, dy1, dz1, dX2, dY2, dZ2
        
        oPositionOnCurve1.Set dx1, dy1, dz1
        oPositionOnCurve2.Set dX2, dY2, dZ2
    
        Dim dU As Double
        Dim dV As Double
        Dim oSPSCrossSec As ISPSCrossSection
        Dim lCardinalPoint As Long
        Dim oXSecUvec As IJDVector
        Dim oXSecVvec As IJDVector
        Dim oXSecZvec As IJDVector
        Dim oAxisVector1 As IJDVector
        Dim oAxisVector2 As IJDVector
        Dim oStructProfile As IJStructProfilePart
        Dim oResVec As IJDVector
        Dim oMatrix As DT4x4
        
        Set oStructProfile = oMemberPart1
        'Get Member Part Matrix
        Set oMatrix = oStructProfile.GetCrossSectionMatrixAtPoint(oPositionOnCurve1)
        
        Set oXSecUvec = New dVector
        Set oXSecVvec = New dVector
        Set oXSecZvec = New dVector
        
        'Get the Member Cross section vectors
        oXSecUvec.Set oMatrix.IndexValue(0), oMatrix.IndexValue(1), oMatrix.IndexValue(2)
        oXSecVvec.Set oMatrix.IndexValue(4), oMatrix.IndexValue(5), oMatrix.IndexValue(6)
        oXSecZvec.Set oMatrix.IndexValue(8), oMatrix.IndexValue(9), oMatrix.IndexValue(10)
        
        Set oAxisVector1 = New dVector
        Set oAxisVector1 = oXSecZvec
        
        Set oSPSCrossSec = oMemberPart1.CrossSection
        lCardinalPoint = oSPSCrossSec.CardinalPoint
        
        If lCardinalPoint <> 5 Then
        'Offet the Curve Postion to 5-Center
        
            oSPSCrossSec.GetCardinalPointDelta Nothing, lCardinalPoint, 5, dU, dV
            
            'Compute bounding U-vector
            oXSecUvec.Length = dU
            
            'Compute bounding V-vector
            oXSecVvec.Length = dV
                
            'Compute resultant vector
            Set oResVec = oXSecUvec.Add(oXSecVvec)
            
            'Get the Postion on 5-Center
            Set oPositionOnCurve1 = oPositionOnCurve1.Offset(oResVec)
        
        End If
        
        'For Second member
        Set oMatrix = New DT4x4
        Set oStructProfile = oMemberPart2
        
        'Get Member Part Matrix
        Set oMatrix = oStructProfile.GetCrossSectionMatrixAtPoint(oPositionOnCurve2)
        
        Set oXSecUvec = New dVector
        Set oXSecVvec = New dVector
        Set oXSecZvec = New dVector
        
        'Get the Member Cross section vectors
        oXSecUvec.Set oMatrix.IndexValue(0), oMatrix.IndexValue(1), oMatrix.IndexValue(2)
        oXSecVvec.Set oMatrix.IndexValue(4), oMatrix.IndexValue(5), oMatrix.IndexValue(6)
        oXSecZvec.Set oMatrix.IndexValue(8), oMatrix.IndexValue(9), oMatrix.IndexValue(10)
        
        Set oAxisVector2 = New dVector
        Set oAxisVector2 = oXSecZvec
        
        Set oSPSCrossSec = oMemberPart2.CrossSection
        lCardinalPoint = oSPSCrossSec.CardinalPoint
        
        If lCardinalPoint <> 5 Then
        'Offet the Curve Postion to 5-Center
        
            oSPSCrossSec.GetCardinalPointDelta Nothing, lCardinalPoint, 5, dU, dV
            
            'Compute bounding U-vector
            oXSecUvec.Length = dU
            
            'Compute bounding V-vector
            oXSecVvec.Length = dV
                
            'Compute resultant vector
            Set oResVec = oXSecUvec.Add(oXSecVvec)
            
            'Get the Postion on 5-Center
            Set oPositionOnCurve2 = oPositionOnCurve2.Offset(oResVec)
        End If
        
        Dim oFirstPos As IJDPosition
        Dim oSecondPos As IJDPosition
        
        Dim oZvector As IJDVector
        Set oZvector = New dVector
        
        Dim oYvector As IJDVector
        Set oYvector = New dVector
        
        Dim oXvector As IJDVector
        Set oXvector = New dVector
        
        'Create Unit Vectors in X,Y and Z directions
        oZvector.Set 0, 0, 1
        oYvector.Set 0, 1, 0
        oXvector.Set 1, 0, 0
        
        Dim oPlaneNormal As IJDVector
        Set oPlaneNormal = New dVector
        
        'Get the normal of the plane defined by member axes
        Set oPlaneNormal = oXSecZvec.Cross(oXSecZvec)
        
        Dim dZvalue As Double
        Dim dXvalue As Double
        Dim dYvalue As Double
        
        'Compute the dot product of Plane normal with X,Y and Z vectors
        dZvalue = Abs(oPlaneNormal.Dot(oZvector))
        dXvalue = Abs(oPlaneNormal.Dot(oXvector))
        dYvalue = Abs(oPlaneNormal.Dot(oYvector))
        
        Dim bZdirection As Boolean
        Dim bXdirection As Boolean
        Dim bYdirection As Boolean
        
        'Check in which direction the noraml is more oriented
        If GreaterThan(dZvalue, dXvalue) Then
            If GreaterThan(dZvalue, dYvalue) Then
                bZdirection = True
            Else
                bYdirection = True
            End If
        Else
            If GreaterThan(dXvalue, dYvalue) Then
                bXdirection = True
            ElseIf GreaterThan(dXvalue, dYvalue) Then
                bYdirection = True
            Else
                bZdirection = True
            End If
        End If
        
        'If the Normal is more oriented in the +z or -z direction consider z component
        If bZdirection Then
            'If both are positive then the one with the highest value gets the PC
            'If both are negative then the one with the lowest value gets the PC
            If (GreaterThanZero(oPositionOnCurve1.z) And GreaterThanZero(oPositionOnCurve2.z)) Or _
                (LessThanZero(oPositionOnCurve1.z) And LessThanZero(oPositionOnCurve2.z)) Then
                
                If GreaterThan(Abs(oPositionOnCurve1.z), Abs(oPositionOnCurve2.z)) Then
                    bFirstIsBounded = True
                Else
                    bSecondIsBounded = True
                End If
                
            Else
                If GreaterThan(oPositionOnCurve1.z, oPositionOnCurve2.z) Then
                    bFirstIsBounded = True
                Else
                    bSecondIsBounded = True
                End If
            End If
            
        'If the Normal is more oriented in the +x or -x direction consider x component
        ElseIf bXdirection Then
            'If both are positive then the one with the highest value gets the PC
            'If both are negative then the one with the lowest value gets the PC
            If (GreaterThanZero(oPositionOnCurve1.x) And GreaterThanZero(oPositionOnCurve2.x)) Or _
                (LessThanZero(oPositionOnCurve1.x) And LessThanZero(oPositionOnCurve2.x)) Then
                
                If GreaterThan(Abs(oPositionOnCurve1.x), Abs(oPositionOnCurve2.x)) Then
                    bFirstIsBounded = True
                Else
                    bSecondIsBounded = True
                End If
                
            Else
                If GreaterThan(oPositionOnCurve1.x, oPositionOnCurve2.x) Then
                    bFirstIsBounded = True
                Else
                    bSecondIsBounded = True
                End If
            End If
            
        'If the Normal is more oriented in the +y or -y direction consider y component
        ElseIf bYdirection Then
            'If both are positive then the one with the highest value gets the PC
            'If both are negative then the one with the lowest value gets the PC
            If (GreaterThanZero(oPositionOnCurve1.y) And GreaterThanZero(oPositionOnCurve2.y)) Or _
                (LessThanZero(oPositionOnCurve1.y) And LessThanZero(oPositionOnCurve2.y)) Then
                
                If GreaterThan(Abs(oPositionOnCurve1.y), Abs(oPositionOnCurve2.y)) Then
                    bFirstIsBounded = True
                Else
                    bSecondIsBounded = True
                End If
                
            Else
                If GreaterThan(oPositionOnCurve1.y, oPositionOnCurve2.y) Then
                    bFirstIsBounded = True
                Else
                    bSecondIsBounded = True
                End If
            End If
        Else
            'If both are positive then the one with the highest z value gets the PC
            'If both are negative then the one with the lowest z value gets the PC
            If (GreaterThanZero(oPositionOnCurve1.z) And GreaterThanZero(oPositionOnCurve2.z)) Or _
                (LessThanZero(oPositionOnCurve1.z) And LessThanZero(oPositionOnCurve2.z)) Then
                
                If GreaterThan(Abs(oPositionOnCurve1.z), Abs(oPositionOnCurve2.z)) Then
                    bFirstIsBounded = True
                Else
                    bSecondIsBounded = True
                End If
                
            Else
                If GreaterThan(oPositionOnCurve1.z, oPositionOnCurve2.z) Then
                    bFirstIsBounded = True
                Else
                    bSecondIsBounded = True
                End If
            End If
        
        End If
    End If
    
    'Set Bounded and Bounding Ports
    If bFirstIsBounded Then
        Set oBoundedPort = oConnectionObject1
        Set oBoundingPort = oConnectionObject2
    ElseIf bSecondIsBounded Then
        Set oBoundedPort = oConnectionObject2
        Set oBoundingPort = oConnectionObject1
    Else
        'Not handled
    End If
    
Exit Sub

ErrorHandler:
''MsgBox Err.Description
    HandleError MODULE, METHOD, sMsg
    
End Sub

'*************************************************************************
'Function
'InitAlongConnectionData_ForBothAlongPorts
'
'Abstract
'   Initialize the BoundedConnection and Bounding Connection Data structures
'
'input
'   oConnectionObject1, oConnectionObject2 and oAppconnection
'Return
'   Bounded Data ,Bounding Data, lStatus, sMsg
'Exceptions
'
'***************************************************************************
Public Sub InitAlongConnectionData_ForBothAlongPorts(oConnectionObject1 As Object, _
                                    oConnectionObject2 As Object, oAppConnection As IJAppConnection, _
                                    oBoundedData As MemberConnectionData, _
                                    oBoundingData As MemberConnectionData, _
                                    lStatus As Long, sMsg As String)
                                    
    Const METHOD = "::InitAlongConnectionData_ForBothAlongPorts"
    On Error GoTo ErrorHandler
    
    Dim iIndex As Long
    Dim lCount As Long
    
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    
    Dim oPort As IJPort
    Dim ePortId As SPSMemberAxisPortIndex
    Dim sCStype As String
    Dim oCrossSection As IJCrossSection
    Dim oBounded_CrossSection As ISPSCrossSection
    Dim oBounding_CrossSection As ISPSCrossSection
    Dim oBounded_PartDesigned As ISPSDesignedMember
    Dim oBounding_PartDesigned As ISPSDesignedMember
    Dim oBounded_PartPrismatic As ISPSMemberPartPrismatic
    Dim oBounding_PartPrismatic As ISPSMemberPartPrismatic
    
    Dim oBoundedPart As ISPSMemberPartCommon
    Dim oBoundingPart As ISPSMemberPartCommon
    
    sMsg = "validate input objects"
    lStatus = 0
    
    If oConnectionObject1 Is Nothing Or oConnectionObject2 Is Nothing Then
        GoTo ErrorHandler
    End If

    ' for Member EndCuts, require two(2) Ports
    sMsg = "Checking Member Assembly Connection Ports"
    lCount = 2
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim oTempPort As IJPort
    
    'Get Bounding and Bounded Ports
    GetBoundedAndBounding_ForBothAlongPorts oConnectionObject1, oConnectionObject2, oAppConnection, oBoundedPort, oBoundingPort
        
    Set oBoundedPart = oBoundedPort.Connectable
    Set oBoundingPart = oBoundingPort.Connectable
    
    Set oBoundedData.AxisPort = oBoundedPort
    Set oBoundedData.MemberPart = oBoundedPart
    oBoundedData.ePortId = SPSMemberAxisAlong
    
    Set oBoundingData.AxisPort = oBoundingPort
    Set oBoundingData.MemberPart = oBoundingPart
    oBoundingData.ePortId = SPSMemberAxisAlong
    
    Dim oCurve1 As IJCurve
    Dim oCurve2 As IJCurve
    
    Set oCurve1 = oBoundedPart.Axis
    Set oCurve2 = oBoundingPart.Axis
    
    Dim dX2 As Double, dY2 As Double, dZ2 As Double
    Dim dMinDist As Double
    
    oCurve1.DistanceBetween oCurve2, dMinDist, dx, dy, dz, dX2, dY2, dZ2
    
    Dim oPosition1 As IJDPosition
    Dim oPosition2 As IJDPosition
    
    Set oPosition1 = New DPosition
    Set oPosition2 = New DPosition
    
    oPosition1.Set dx, dy, dz
    oPosition2.Set dX2, dY2, dZ2
          
    ' Verify Bounded is valid MemberPartPrismatic object
    sCStype = ""
    If oBoundedData.MemberPart.IsPrismatic Then
        ' If Bounded object is MemberPartPrismatic:
        Set oBounded_PartPrismatic = oBoundedData.MemberPart
        Set oBounded_CrossSection = oBounded_PartPrismatic.CrossSection
    
    ElseIf TypeOf oBoundedData.MemberPart Is ISPSDesignedMember Then
        ' If Bounded object is DesignedMember:
        Set oBounded_PartDesigned = oBoundedData.MemberPart
        Set oBounded_CrossSection = oBounded_PartDesigned

        lStatus = 12
        sMsg = "Bounded object is ISPSDesignedMember (NOT MemberPartPrismatic)"
        GoTo StatusFalse
    
    Else
        lStatus = 13
        sMsg = "Bounded object type is NOT Known"
        GoTo StatusFalse
    End If
    
    ' Verify Bounding have valid Cross Section Type
    If oBoundingData.MemberPart.IsPrismatic Then
        ' If Bounding object is MemberPartPrismatic:
        Set oBounding_PartPrismatic = oBoundingData.MemberPart
        Set oBounding_CrossSection = oBounding_PartPrismatic.CrossSection
    
    ElseIf TypeOf oBoundingData.MemberPart Is ISPSDesignedMember Then
        ' If Bounding object is DesignedMember:
        Set oBounding_PartDesigned = oBoundingData.MemberPart
        Set oBounding_CrossSection = oBounding_PartDesigned
    
        ' Check if all of the children are valid Detailed Parts
        If Not IsDesignedMemberDetailed(oBounding_PartDesigned) Then
            lStatus = 14
            sMsg = "Bounding object is NOT Detailed ISPSDesignedMember"
            GoTo StatusFalse
        End If
        
    Else
        ' If Bounding object is unKnown
        lStatus = 15
        sMsg = "Bounding object type is NOT Known"
        GoTo StatusFalse
    End If
    
    ' Verify Bounded have valid Cross Section Type
    If TypeOf oBounded_CrossSection.definition Is IJCrossSection Then
        Set oCrossSection = oBounded_CrossSection.definition
        sCStype = oCrossSection.Type
       
        sMsg = "Bounded CrossSection Type is not valid: " & sCStype
        
    Else
        lStatus = 20
        sMsg = "Bounded CrossSection Definition is not valid"
        GoTo StatusFalse
    End If
    
    ' Get the Bounded Port Axis curve and Point on Axis Curve (x,y,z Location)
    sMsg = "Calculating Bounded Axis/Location"
    Set oPort = oBoundedData.AxisPort
    
    If TypeOf oPort.Geometry Is IJCurve Then
        
        oPosition1.Get dx, dy, dz
        
        Set oBoundedData.AxisCurve = GetAxisCurveAtPosition(dx, dy, dz, oBoundedData.MemberPart)
        oBoundedData.MemberPart.Rotation.GetTransformAtPosition dx, dy, dz, oBoundedData.Matrix, Nothing

    Else
        lStatus = 30
        sMsg = "Member Assembly Connection Bounded Port geometry does not support IJCurve interface"
        GoTo StatusFalse
    End If
        
    ' Get the Bounding Port Axis curve and Point on Axis Curve (x,y,z Location)
    sMsg = "Calculating Bounding Axis/Location"
    oPosition2.Get dx, dy, dz
    
    If TypeOf oBoundedData.AxisCurve Is IJCurve Then
        Set oBoundingData.AxisCurve = GetAxisCurveAtPosition(dx, dy, dz, oBoundingData.MemberPart)
        oBoundingData.MemberPart.Rotation.GetTransformAtPosition dx, dy, dz, oBoundingData.Matrix, Nothing
    Else
        lStatus = 41
        sMsg = "Member Assembly Connection Bounded Axis Curve does not support IJCurve interface"
        GoTo StatusFalse
    End If
    
    Exit Sub
    
StatusFalse:
    If lStatus = 0 Then
        lStatus = 1
    End If
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    lStatus = E_FAIL
End Sub

'*************************************************************************
'Function
'GetBoundedAndBounding_ForBothAlongPorts
'
'Abstract
'   Gets Bounded and Bounding Ports
'   Bounded and Bounding are swaped when Flip Flag is ON.
'input
'   oConnectionObject1, oConnectionObject2 and oAppConnection
'Return
'   oBoundedPort, oBoundingPort
'
'***************************************************************************

Public Sub GetBoundedAndBounding_ForBothAlongPorts(oConnectionObject1 As Object, oConnectionObject2 As Object, _
                                        oAppConnection As Object, oBoundedPort As Object, oBoundingPort As Object)

    On Error GoTo ErrorHandler
    Const METHOD = "::GetBoundedAndBounding_ForBothAlongPorts"
    
    Dim sMsg As String
    
    'Get Bounded and Bounding based on the
    GetLogicalBoundedAndBounding_ForBothAlongPorts oConnectionObject1, oConnectionObject2, oBoundedPort, oBoundingPort
    
    If oAppConnection Is Nothing Then Exit Sub
    
    Dim sFlipped As String
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    
    Set oSmartOccurrence = oAppConnection
    Set oSmartItem = oSmartOccurrence.ItemObject

    If Not oSmartItem Is Nothing Then
        GetSelectorAnswer oAppConnection, "Flip Primary and Secondary", sFlipped
    Else
        Dim oAttributes As IJDAttributes
        Set oAttributes = oSmartOccurrence
        
        Dim oAttribute As IJDAttribute
        Set oAttribute = oAttributes.CollectionOfAttributes("{DAC6AAA1-6CA1-41F0-A424-4AABF08AEA78}").Item("Flip Primary and Secondary")
        
        Dim pObject As IJDObject
        Dim pCodeListMD As IJDCodeListMetaData
        
        Set pObject = oSmartOccurrence
        Set pCodeListMD = pObject.ResourceManager
        sFlipped = pCodeListMD.ShortStringValue(oAttribute.AttributeInfo.CodeListTableName, oAttribute.value)
    End If
    
    Dim oTempPort As IJPort
    
    If sFlipped = gsYes Then
        Set oTempPort = oBoundedPort
        Set oBoundedPort = oBoundingPort
        Set oBoundingPort = oTempPort
    End If
    
  Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'IsMiterConnectionPossible
'
'Abstract
'Caluculate the flanges orientation and gives whether miters could be palced between the members or not.
'
'Arguments
'Bounded and bounding parts as objects
'
'Return
'Boolean value
'
'Exceptions
'
'***************************************************************************
Public Sub IsMiterConnectionPossible(oBoundedPort As Object, _
                                   oBoundingPort As Object, _
                                   IsMiterConnectionPossible As Boolean)
    Const METHOD = "::IsMiterConnectionPossible"
    On Error GoTo ErrorHandler
    IsMiterConnectionPossible = False
    Dim sMsg As String
    
    If oBoundedPort Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oBoundedPort Is ISPSSplitAxisPort Then
        Exit Sub
    End If
    
    If oBoundingPort Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oBoundingPort Is ISPSSplitAxisPort Then
        Exit Sub
    End If
    
    Dim oMat As IJDT4x4
    Dim oBddUDir As New dVector
    Dim oBdgUDir As New dVector
    Dim oBddVDir As New dVector
    Dim oBdgVDir As New dVector
    Dim oBddWDir As New dVector
    Dim oBdgWDir As New dVector
    Dim dBdgWdotBddU As Double
    Dim dBdgUdotBddW As Double
    Dim dBdgWdotBddV As Double
    Dim dBddWdotBdgV As Double
    Dim dVdotV As Double
    Dim dUdotU As Double
    
    Dim oBoundedPart As ISPSMemberPartPrismatic
    Dim oBoundingPart As ISPSMemberPartPrismatic
    Dim oBoundedAxisPort As ISPSSplitAxisPort
    Dim oBoundingAxisPort As ISPSSplitAxisPort
    Dim sX0#, sY0#, sZ0#, eX0#, eY0#, eZ0#
    Dim sX1#, sY1#, sZ1#, eX1#, eY1#, eZ1#

    Dim oBdgStartPoint As IJPoint
    Dim oBdgEndPoint As IJPoint
    Dim oBddPoint As IJPoint
    Dim dDistStart As Double
    Dim dDistEnd As Double
    
    Set oBoundingAxisPort = oBoundingPort
    Set oBoundedAxisPort = oBoundedPort
    Set oBoundedPart = oBoundedAxisPort.Part
    Set oBoundingPart = oBoundingAxisPort.Part

    Dim bHasTopLeft As Boolean
    Dim bHasBtmLeft As Boolean
    Dim bHasTopFlng As Boolean
    Dim bHasBtmFlng As Boolean
    bHasTopFlng = HasTopFlange(oBoundedPart, bHasTopLeft)
    bHasBtmFlng = HasBottomFlange(oBoundedPart, bHasBtmLeft)
    
    'HasTopFlange() and HasBottomFlange()methods return boolean as tru for W and C cross-sections as well.
    'To avoid C-cross section in the following If clause, optional arguments bHasTopLeft,bHasBtmLeft are passed. Calculation for 'C' will be done later.
    'These checks only involve W, rectangular and circular cross-scetions. For W, rectangular cross-sections, if both the bounding and bounded rotational angles are same, then boolean is true.
    'Also if Sine of both the angles is '0' that is multiples of PI, then boolean is returned as True.
    If bHasTopFlng And bHasBtmFlng And bHasTopLeft And bHasBtmLeft Then  'W
        If (Equal(Sin(oBoundedAxisPort.Part.Rotation.BetaAngle), 0) And _
            Equal(Sin(oBoundingAxisPort.Part.Rotation.BetaAngle), 0)) Or _
            (Equal(oBoundedAxisPort.Part.Rotation.BetaAngle, oBoundingAxisPort.Part.Rotation.BetaAngle)) Then
            IsMiterConnectionPossible = True
            Exit Sub
        End If
    ElseIf IsTubularMember(oBoundedPart) Then
        IsMiterConnectionPossible = True
        Exit Sub
    ElseIf IsRectangularMember(oBoundedPart) Then
        If (Equal(Sin(oBoundedAxisPort.Part.Rotation.BetaAngle), 0) And _
            Equal(Sin(oBoundingAxisPort.Part.Rotation.BetaAngle), 0)) Or _
        Equal(oBoundedAxisPort.Part.Rotation.BetaAngle, oBoundingAxisPort.Part.Rotation.BetaAngle) Then
            IsMiterConnectionPossible = True
            Exit Sub
        End If
    End If
    
    oBoundedPart.Axis.EndPoints sX0, sY0, sZ0, eX0, eY0, eZ0
    oBoundingPart.Axis.EndPoints sX1, sY1, sZ1, eX1, eY1, eZ1

    'Get U,V,W vectors.
    'W-vector is the Axis vector which always directs from start to End.
    'If Bounded Axis Port is Start, then set W-vector direction to negative value. If Axis Port is at End, then W-vector will be as it is.
    'W- vector is in the direction oh Baseport normal or offset port normal based on the direction of the W-vector set.
    'When the member is reflected, then U-vector direction will not change,for remaining cases(i.e;rotation), U-vector direction changes.
    'This is for finding bounded vectors. The same procedure is involved to calculate bounding vectors.
    If oBoundedAxisPort.PortIndex = SPSMemberAxisStart Then
        oBoundedPart.Rotation.GetTransformAtPosition sX0, sY0, sZ0, oMat, Nothing
        oBddWDir.Set -oMat.IndexValue(0), -oMat.IndexValue(1), -oMat.IndexValue(2) 'Parallel to Base/Offset port normal
        Set oBddPoint = oBoundedPart.PointAtEnd(SPSMemberAxisStart)
    ElseIf oBoundedAxisPort.PortIndex = SPSMemberAxisEnd Then
        oBoundedPart.Rotation.GetTransformAtPosition eX0, eY0, eZ0, oMat, Nothing
        oBddWDir.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2) 'Parallel to Base/Offset port normal
        Set oBddPoint = oBoundedPart.PointAtEnd(SPSMemberAxisEnd)
    End If
    If oBoundedPart.Rotation.Mirror = True Then
        oBddUDir.Set oMat.IndexValue(4), oMat.IndexValue(5), oMat.IndexValue(6) 'Parallel to web normal
    Else
        oBddUDir.Set -oMat.IndexValue(4), -oMat.IndexValue(5), -oMat.IndexValue(6) 'Parallel to web normal
    End If
    oBddVDir.Set oMat.IndexValue(8), oMat.IndexValue(9), oMat.IndexValue(10) 'Parallel to Topflng port normal
    
    'To know, whether the Axis port is Start or end:
    Set oBdgStartPoint = oBoundingPart.PointAtEnd(SPSMemberAxisStart)
    Set oBdgEndPoint = oBoundingPart.PointAtEnd(SPSMemberAxisEnd)
    dDistStart = oBddPoint.DistFromPt(oBdgStartPoint)
    dDistEnd = oBddPoint.DistFromPt(oBdgEndPoint)
    
    If dDistStart < dDistEnd Then 'Axis-port is at start
        oBoundingPart.Rotation.GetTransformAtPosition sX1, sY1, sZ1, oMat, Nothing
        oBdgWDir.Set -oMat.IndexValue(0), -oMat.IndexValue(1), -oMat.IndexValue(2) 'Parallel to Base/Offset port normal
    Else 'Axis port is at end
        oBoundingPart.Rotation.GetTransformAtPosition eX1, eY1, eZ1, oMat, Nothing
        oBdgWDir.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2) 'Parallel to Base/Offset port normal
    End If
    If oBoundingPart.Rotation.Mirror = True Then
        oBdgUDir.Set oMat.IndexValue(4), oMat.IndexValue(5), oMat.IndexValue(6) 'Parallel to web normal
    Else
        oBdgUDir.Set -oMat.IndexValue(4), -oMat.IndexValue(5), -oMat.IndexValue(6) 'Parallel to web normal
    End If
    oBdgVDir.Set oMat.IndexValue(8), oMat.IndexValue(9), oMat.IndexValue(10) 'Parallel to Topflng port normal
 
    dBdgWdotBddU = oBddUDir.Dot(oBdgWDir)
    dBdgUdotBddW = oBdgUDir.Dot(oBddWDir)
    dVdotV = oBddVDir.Dot(oBdgVDir)
    dUdotU = oBddUDir.Dot(oBdgUDir)
    dBdgWdotBddV = oBdgWDir.Dot(oBddVDir)
    dBddWdotBdgV = oBddWDir.Dot(oBdgVDir)
    
    If bHasTopFlng And bHasBtmFlng And Not bHasTopLeft And Not bHasBtmLeft Then 'C
        ' When members are in Flange plane, get
        ' i) the dot product between bounded W-vector and bounding v-vector
        ' ii) the dot product between bounding W-vector and bounded v-vector. When both are positive/Negative, then miters are possible for such cases.
        ' When both dt products are of opposite signs, then miters cannot be given n such cases.
        If GreaterThanZero(dBdgWdotBddU * dBdgUdotBddW) Then
            IsMiterConnectionPossible = True
        ' When members are in Web Plane, then above mentioned dot products are zero. In such cases, if both bounding and bounded U-vectors are in same direction, niter is possible.
        ElseIf (Equal((dBdgWdotBddU * dBdgUdotBddW), 0)) Then
            If GreaterThan(dUdotU, CosTheta) Then
                IsMiterConnectionPossible = True
            End If
        Else
        End If
    
    ElseIf bHasBtmFlng And Not bHasTopFlng Then  'L
        ' When members are in Flange plane, get
        ' i) the dot product between bounded W-vector and bounding v-vector
        ' ii) the dot product between bounding W-vector and bounded v-vector. When both are positive/Negative, then miters are possible for such cases.
        ' When both dt products are of opposite signs, then miters cannot be given n such cases.
        If GreaterThanZero(dBdgWdotBddU * dBdgUdotBddW) Then
            If GreaterThan(dVdotV, CosTheta) Then
                IsMiterConnectionPossible = True
            End If
        ' When members are in Web Plane, then above mentioned dot products are zero. In such cases, if both bounding and bounded U-vectors are in same direction, niter is possible.
        ElseIf (Equal((dBdgWdotBddU * dBdgUdotBddW), 0)) Then
            If GreaterThan(dUdotU, CosTheta) Then
                'For L cross-sections, if though it satisfies above cases and both flanges do not intersect each other, no miters should be given.
                'When V-vectors are perpendicular to each other,such cases we depend on :
                'i) dot product of bounding W-vector and bounded U-vetor
                'ii) dot product of bounded W-vector and bounding U-vetor.When both are positive/Negative, then miters are possible for such cases.
                ' When both dt products are of opposite signs, then miters cannot be given n such cases.
                If Equal(dVdotV, 0) Then
                    If GreaterThanZero(dBdgWdotBddV * dBddWdotBdgV) Then
                        IsMiterConnectionPossible = True
                    End If
                End If
            End If
        End If
    
    ElseIf bHasTopFlng And Not bHasBtmFlng Then 'T
        'When members are in flange plane:
        'For T cross-sections, if both flanges do not intersect each other, no miters should be given.
        'For such cases, if V-vectors are in same direction (that means Top flnage normals in same direction,miters are accepted.
        If GreaterThanZero(dVdotV) Then
            If ((Sin(oBoundedAxisPort.Part.Rotation.BetaAngle) = 0) And (Sin(oBoundingAxisPort.Part.Rotation.BetaAngle) = 0)) Or _
            (oBoundedAxisPort.Part.Rotation.BetaAngle = oBoundingAxisPort.Part.Rotation.BetaAngle) Then
                IsMiterConnectionPossible = True
            End If
        'When members are in web plane, dot product of V-vectors is zero. such cases we depend on :
        'i) dot product of bounding W-vector and bounded U-vetor
        'ii) dot product of bounded W-vector and bounding U-vetor.When both are positive/Negative, then miters are possible for such cases.
        ' When both dt products are of opposite signs, then miters cannot be given n such cases.
        ElseIf Equal(dVdotV, 0) Then
            If GreaterThanZero(dBdgWdotBddV * dBddWdotBdgV) Then
                IsMiterConnectionPossible = True
            End If
        End If
    End If
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'*************************************************************************
'Method
'GetNearestboundingBUPort
'
'Abstract
'Returns the nearest Builtup bounding port for the given bounded far end, bounded data and builtup member
'
'Arguments
'Bounded far end as IJDPosition, bounded data as MemberConnectionData and bounding member as ISPSDesignedMember
'
'Return
'Nearest bounding port as IJPort
'
'Exceptions
'
'***************************************************************************

Public Sub GetNearestboundingBUPort(oBoundedRefPos As IJDPosition, oBoundedPort As IJPort, _
                                        oBoundingBU As ISPSDesignedMember, oNearBoundingBUPort As IJPort)
    Const METHOD = "::GetNearestboundingBUPort"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim oDesignParent As IJDesignParent
    Set oDesignParent = oBoundingBU
    Dim oRootPlateSystems As IJDObjectCollection
    Dim oLeafPlateSystems As IJDObjectCollection
    Dim oLeafPlatesColl As IJDObjectCollection
    Dim oReqPlatesColl As IJDObjectCollection
    Set oRootPlateSystems = New JObjectCollection
    Dim oPlateParts As IJDObjectCollection
    Dim oPlatePartsColl As IJDObjectCollection
    Set oPlatePartsColl = New JObjectCollection
    Set oLeafPlatesColl = New JObjectCollection
    Set oReqPlatesColl = New JObjectCollection
    Dim oRootPlate As IJPlate
    Dim oLeafPlate As IJPlate
    
      '*******getting the children plate systems**********
    oDesignParent.GetChildren oRootPlateSystems
    Dim Obj1 As Object
     If oRootPlateSystems.Count <> 0 Then
        For Each oRootPlate In oRootPlateSystems
             Set oDesignParent = oRootPlate
             oDesignParent.GetChildren oLeafPlateSystems
             If oLeafPlateSystems.Count <> 0 Then
                For Each Obj1 In oLeafPlateSystems
                    If TypeOf Obj1 Is IJPlate Then
                        oLeafPlatesColl.Add Obj1
                    End If
                Next
             Else
                GoTo ErrorHandler
             End If
        Next
     Else
        GoTo ErrorHandler
     End If
     
    '********getting the bounded_W vector************
    Dim oNormal As IJDVector
    Set oNormal = New dVector
    Dim oBoundedAxis_W As IJDVector
    Set oBoundedAxis_W = New dVector
    Dim oBddAxisPort As ISPSSplitAxisPort
    Set oBddAxisPort = oBoundedPort
    Dim oMemberPart As ISPSMemberPartCommon
    Dim oCSMatrix As IJDT4x4
    Set oCSMatrix = New DT4x4
    Set oMemberPart = oBoundedPort.Connectable
    oMemberPart.Rotation.GetTransform oCSMatrix
    If oBddAxisPort.PortIndex = SPSMemberAxisStart Then
    oBoundedAxis_W.Set oCSMatrix.IndexValue(0), _
            oCSMatrix.IndexValue(1), _
            oCSMatrix.IndexValue(2)
    ElseIf oBddAxisPort.PortIndex = SPSMemberAxisEnd Then
    oBoundedAxis_W.Set -oCSMatrix.IndexValue(0), _
            -oCSMatrix.IndexValue(1), _
            -oCSMatrix.IndexValue(2)
    Else
    Exit Sub
    End If
    oBoundedAxis_W.Length = 1
    
    'getting the normal on the bounding plate part & checking the dot product with the bounded vector
    
    Dim oSBody As IJSurfaceBody
    Dim oGeomOpr As GSCADShipGeomOps.SGOModelBodyUtilities
    Dim oModelBody As IJDModelBody
    Set oGeomOpr = New SGOModelBodyUtilities
    Dim oCOG As IJDPosition
    Set oCOG = New DPosition
    Dim oPlatePart As New StructDetailObjects.PlatePart
    For Each oLeafPlate In oLeafPlatesColl
        If TypeOf oLeafPlate Is IJPlate Then
            Set oPlatePart.object = oLeafPlate
            Set oSBody = oLeafPlate
            oSBody.GetCenterOfGravity oCOG
            oSBody.GetNormalFromPosition oCOG, oNormal
            oNormal.Length = 1
            If Equal(Abs(oBoundedAxis_W.Dot(oNormal)), 1) Then
                oReqPlatesColl.Add oLeafPlate
            End If
        End If
    Next
    
    Dim oTopologyLoc As GSCADStructGeomUtilities.TopologyLocate
    Set oTopologyLoc = New TopologyLocate
    Dim oProjectedPoint As IJDPosition
    Set oProjectedPoint = New DPosition
    Dim oBoundedWLPort As IJPort
    Dim IsIntersecting As Boolean
    Dim oMbrPart As New StructDetailObjects.MemberPart
    Set oMbrPart.object = oMemberPart
    Set oBoundedWLPort = GetLateralSubPortBeforeTrim(oMbrPart.object, JXSEC_WEB_LEFT)
    Dim oExtWLPort As Object
    If Not oBoundedWLPort Is Nothing Then
        Set oExtWLPort = GetExtendedPort(oBoundedWLPort)
    End If
    Dim dMin As Double
    Dim dDistance As Double
    dMin = 1000#  ' distance
    Dim oBasePort As IJPort
    Dim oOffsetPort As IJPort
    Dim oBaseModelBody As IJDModelBody
    Dim oOffsetModelBody As IJDModelBody
    Dim dMinDist1 As Double
    Dim dMinDist2 As Double
    Dim oReqPlate As New StructDetailObjects.PlatePart
    
    '*********Inclined Cases # checking for the intersections
    
    If oReqPlatesColl.Count = 0 Then
        For Each oLeafPlate In oLeafPlatesColl
            IsIntersecting = oGeomOpr.HasIntersectingGeometry(oExtWLPort, oLeafPlate)
            If IsIntersecting Then
                oReqPlatesColl.Add oLeafPlate
            End If
        Next
        If oReqPlatesColl.Count <> 0 Then
            For Each oLeafPlate In oReqPlatesColl
                If TypeOf oLeafPlate Is IJPlate Then
                    Set oDesignParent = oLeafPlate
                    oDesignParent.GetChildren oPlateParts
                    If oPlateParts.Count <> 0 Then
                        For Each Obj1 In oPlateParts
                            oPlatePartsColl.Add Obj1
                        Next
                    Else
                        GoTo ErrorHandler
                    End If
                End If
            Next
        Else
            GoTo ErrorHandler
        End If
        
            '***getting the nearest plate part from the collection
            
        Set Obj1 = Nothing
        If oPlatePartsColl.Count <> 0 Then
            For Each Obj1 In oPlatePartsColl
                Set oPlatePart.object = Obj1
                Set oBasePort = oPlatePart.BasePort(BPT_Base)
                Set oBaseModelBody = oBasePort.Geometry
                Set oSBody = oBasePort.Geometry
                oSBody.GetCenterOfGravity oCOG
                dDistance = oCOG.DistPt(oBoundedRefPos)
                If dMin > dDistance Then
                    dMin = dDistance
                    Set oReqPlate.object = Obj1
                End If
            Next
        Else
            GoTo ErrorHandler
        End If
    Else
        'Ordinary Cases
        Dim Obj2 As Object
        Set Obj1 = Nothing
        For Each Obj2 In oReqPlatesColl
            If TypeOf Obj2 Is IJPlate Then
                Set oSBody = Obj2
                oTopologyLoc.GetProjectedPointOnModelBody oSBody, oBoundedRefPos, oProjectedPoint, Nothing
                dDistance = oProjectedPoint.DistPt(oBoundedRefPos)
                If dMin > dDistance Then
                    dMin = dDistance
                    Set oDesignParent = Obj2
                    oDesignParent.GetChildren oPlatePartsColl
                    For Each Obj1 In oPlatePartsColl
                        If TypeOf Obj1 Is IJPlatePart Then
                            Set oReqPlate.object = Obj1
                        End If
                    Next
                End If
            End If
        Next
    End If
    
    Set oBasePort = oReqPlate.BasePort(BPT_Base)
    Set oOffsetPort = oReqPlate.BasePort(BPT_Offset)
    Set oBaseModelBody = oBasePort.Geometry
    Set oOffsetModelBody = oOffsetPort.Geometry
    oBaseModelBody.GetMinimumDistanceFromPosition oBoundedRefPos, Nothing, dMinDist1
    oOffsetModelBody.GetMinimumDistanceFromPosition oBoundedRefPos, Nothing, dMinDist2
 
    If LessThan(dMinDist1, dMinDist2) Then
        Set oNearBoundingBUPort = oBasePort
    Else
        Set oNearBoundingBUPort = oOffsetPort
    End If
    
        Exit Sub
ErrorHandler:
        HandleError MODULE, METHOD, sMsg
End Sub


Public Sub AreMembersIdentical(oBoundedPort As Object, _
                                   oBoundingPort As Object, _
                                   bIdentical As Boolean)
Const METHOD = "::AreMembersIdentical"
    On Error GoTo ErrorHandler
    bIdentical = False
    
    Dim oBoundingBuiltupPort As IJPort
    Set oBoundingBuiltupPort = oBoundingPort
    If oBoundedPort Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oBoundedPort Is ISPSSplitAxisPort Then
        Exit Sub
    End If
    
    If oBoundingPort Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oBoundingPort Is ISPSSplitAxisPort Then
        Exit Sub
    ElseIf TypeOf oBoundingBuiltupPort.Connectable Is ISPSDesignedMember Then
        Exit Sub
    End If
    
    Dim oBoundedPart As ISPSCrossSection
    Dim oBoundingPart As ISPSCrossSection
    Dim oBoundedAxisPort As ISPSSplitAxisPort
    Dim oBoundingAxisPort As ISPSSplitAxisPort
    
    Set oBoundingAxisPort = oBoundingPort
    Set oBoundingPart = oBoundingAxisPort.Part.CrossSection
        
    Set oBoundedAxisPort = oBoundedPort
    Set oBoundedPart = oBoundedAxisPort.Part.CrossSection

    Dim sData1 As String
    Dim sData2 As String

    sData1 = oBoundedPart.sectionType
    sData2 = oBoundingPart.sectionType
    
    If Trim(LCase(sData1)) = Trim(LCase(sData2)) Then
        sData1 = oBoundedPart.SectionName
        sData2 = oBoundingPart.SectionName
        If Trim(LCase(sData1)) = Trim(LCase(sData2)) Then
            sData1 = oBoundedPart.SectionStandard
            sData2 = oBoundingPart.SectionStandard
            If Trim(LCase(sData1)) = Trim(LCase(sData2)) Then
                bIdentical = True
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
    
End Sub



'*************************************************************************
'Function
'GetFrameConnectionType
'
'Abstract
'   Given the Frameconnection type of member at end
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function GetFrameConnectionType(oPort As Object) As SPSFCBoundingType

    Const METHOD = "::GetFrameConnectionType"
    On Error GoTo ErrorHandler
    
    GetFrameConnectionType = 0  '0 is Default type of "Unknown"
         
    If oPort Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oPort Is ISPSSplitAxisPort Then
        Exit Function
    End If
    

    Dim oMemberPart As ISPSMemberPartCommon
    Dim oMemberAxisPort As ISPSSplitAxisPort

    Set oMemberAxisPort = oPort
    Set oMemberPart = oMemberAxisPort.Part
    
    Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
    Dim oMemberSystem As ISPSMemberSystem
    Dim oFC As ISPSFrameConnection

    Set oMemberSystem = oMemberPart.MemberSystem
    If Not oMemberAxisPort.PortIndex = SPSMemberAxisAlong Then
        Set oFC = oMemberSystem.FrameConnectionAtEnd(oMemberAxisPort.PortIndex)
    End If
    If Not oFC Is Nothing Then
        Set oMemberConnectionServices = oFC.Services
        oMemberConnectionServices.GetFCBoundingType oFC, GetFrameConnectionType
    End If
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

' ******************************************************************************************
' ******************************************************************************************
' ******************************************************************************************

Public Sub Migrate_StructAssemblyConnection(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJStructAssemblyConnection Then
        Exit Sub
    Else
        sMemberObjectType = "IJStructAssemblyConnection"
    End If
    
    Dim bNewPort As Boolean
    Dim oPort1 As Object
    Dim oPort2 As Object
    Dim oPart1 As Object
    Dim oPart2 As Object
    Dim oNewPort1 As Object
    Dim oNewPort2 As Object
    Dim oNewPart1 As Object
    Dim oNewPart2 As Object
    
    Dim oPort As IJPort
    Dim oAppConn As IJAppConnection
    Dim oElements As IJElements
    Dim oCollectionAlias As JCmnShp_CollectionAlias
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
    
    ' Retreive Assembly Connection Ports (expect two)
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJAppConnection Then
        Exit Sub
    End If
    
    Set oAppConn = oToBeMigrated
    oAppConn.enumPorts oElements
    If oElements Is Nothing Then
        Exit Sub
    ElseIf oElements.Count < 2 Then
        Exit Sub
    End If
    
    Set oPort1 = oElements.Item(1)
    Set oPort2 = oElements.Item(2)
    
    ' Check First Port's  Part
    If TypeOf oPort1 Is IJPort Then
        Set oPort = oPort1
        Set oPart1 = oPort.Connectable
    End If
    
    ' Check Second Port's Part
    If TypeOf oPort2 Is IJPort Then
        Set oPort = oPort2
        Set oPart2 = oPort.Connectable
    End If
    
    ' Check/Migrate Inputs
    Migrate_GetReplacingObject oPort1, oMigrateHelper, oNewPort1, oNewPart1
    
    Migrate_GetReplacingObject oPort2, oMigrateHelper, oNewPort2, oNewPart2
    
    ' Want to keep the order of Ports the same
    ' Remove both Ports from Assembly Conenction,
    ' Then add both Ports back to the Assembly Connection
    bNewPort = False
    If Not oPort1 Is oNewPort1 Then
        bNewPort = True
    ElseIf Not oPort2 Is oNewPort2 Then
        bNewPort = True
    End If
    
    If bNewPort Then
        oAppConn.removePort oPort1
        oAppConn.removePort oPort2
        
        oAppConn.addPort oNewPort1
        oAppConn.addPort oNewPort2
    End If
        
If bSM_trace Then
    If Not oPort1 Is oNewPort1 Then
        zSM_trace "*** Migrate_StructAssemblyConnection ... migrate port1 : " & Debug_ObjectName(oNewPort1, True)
    End If
    
    If Not oPort2 Is oNewPort2 Then
        zSM_trace "*** Migrate_StructAssemblyConnection ... migrate port2 : " & Debug_ObjectName(oNewPort2, True)
    End If

    If Not bNewPort Then
        zSM_trace "*** Migrate_StructAssemblyConnection *** no migration of input ports required : "
    End If
End If

    ' Check/Migrate Auxiliary Ports
    ' ... Note
    ' ... ... Need to expose StructGenericUtilities.IJMigrationHelper in middle tier
    Dim oPoint As IJDPosition
    Dim bIsDeleted As Boolean
    Dim bIsRefColMigrated As Boolean
    Dim oAssocHelper_MigrationHelper As StructGenericUtilities.IJMigrationHelper
    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection
    
    Set oConnAttrbs = New SDConnectionUtils
    Set oCollectionAlias = oConnAttrbs.get_AuxiliaryPorts(oAppConn)
    If oCollectionAlias Is Nothing Then
    ElseIf oCollectionAlias.Count > 0 Then
        Set oAssocHelper_MigrationHelper = New StructGenericUtilities.AssocHelper
        For iIndex = 1 To oCollectionAlias.Count
            oMigrateHelper.ObjectsReplacing oCollectionAlias.Item(iIndex), oObjectCollectionReplacing, bIsDeleted
            If oObjectCollectionReplacing Is Nothing Then
            ElseIf oObjectCollectionReplacing.Count < 1 Then
            Else
                oAssocHelper_MigrationHelper.MigrateAuxiliaryPortsRefCol oAppConn, oAppConn, oCollectionAlias.Item(iIndex), oObjectCollectionReplacing
            End If
        Next iIndex
        
    End If
        
    ' Check/Migrate ReferencesCollection Ports
    Migrate_ReferencesCollection oToBeMigrated, oMigrateHelper
    
    ' Check/Migrate Member Split By Plate
    ' ... Member Splits are based on the DesignParent Member System
    Dim oMbrByPlateSplits As Collection
    If Migrate_IsMbrSplitByPlate(oToBeMigrated, oPort1, oPort2, oMigrateHelper, oMbrByPlateSplits) Then
    End If
    
    ' Check/Migrate Design Parent
    Dim oParentNew As Object
    Migrate_SystemParent oToBeMigrated, oMigrateHelper, oParentNew
        
    If bNewPort Then
        Migrate_SetReplacedReplacing oToBeMigrated, oToBeMigrated, oMigrateHelper
    End If
        
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_StructAssemblyConnection"
    Err.Raise LogError(Err, MODULE, "Migrate_StructAssemblyConnection", strError).Number
End Sub

Public Sub Migrate_FreeEndCut(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJFreeEndCut Then
        Exit Sub
    Else
        sMemberObjectType = "IJFreeEndCut"
    End If
    
    Dim bNewPort As Boolean
    
    Dim oPort1 As Object
    Dim oPort2 As Object
    Dim oPart1 As Object
    Dim oPart2 As Object
    Dim oNewPort1 As Object
    Dim oNewPort2 As Object
    Dim oNewPart1 As Object
    Dim oNewPart2 As Object
    
    Dim oPort As IJPort
    Dim oFreeEndcut As IJFreeEndCut
    Dim oNewFreeEndCut As IJFreeEndCut
    Dim oResourceManager As IUnknown
    
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDFeatureDefinition As IJSDFeatureDefinition
    
    Set oFreeEndcut = oToBeMigrated
    oFreeEndcut.get_FreeEndCutInputs oPort1, oPort2
    
    ' Check First Port's  Part
    If oPort1 Is Nothing Then
        Exit Sub
    ElseIf TypeOf oPort1 Is IJPort Then
        Set oPort = oPort1
        Set oPart1 = oPort.Connectable
    End If
    
    ' Check Second Port's Part
    If oPort2 Is Nothing Then
        Exit Sub
    ElseIf TypeOf oPort2 Is IJPort Then
        Set oPort = oPort2
        Set oPart2 = oPort.Connectable
    End If
    
    ' Check/Migrate Inputs
    Migrate_GetReplacingObject oPort1, oMigrateHelper, oNewPort1, oNewPart1
    
    Migrate_GetReplacingObject oPort2, oMigrateHelper, oNewPort2, oNewPart2
    
    ' Check if the inputs Ports have been migrated
    bNewPort = False
    If Not oPort1 Is oNewPort1 Then
        bNewPort = True
    ElseIf Not oPort2 Is oNewPort2 Then
        bNewPort = True
    End If
    
    If bNewPort Then
        strError = "Replacing FreeEndCut ... "
        Set oSDO_Helper = New StructDetailObjects.Helper
        Set oResourceManager = oSDO_Helper.GetResourceManagerFromObject(oToBeMigrated)
        
        Set oSDFeatureDefinition = New GSCADSDCreateModifyUtilities.SDFeatureUtils
        Set oNewFreeEndCut = oSDFeatureDefinition.CreateModifyFreeEndCut(oResourceManager, oNewPort1, oNewPort2, vbNullString, Nothing, oFreeEndcut)
    End If
        
If bSM_trace Then
    If Not oPort1 Is oNewPort1 Then
        zSM_trace "*** Migrate_FreeEndCut ... migrate port1 : " & Debug_ObjectName(oNewPort1, True)
    End If
    
    If Not oPort2 Is oNewPort2 Then
        zSM_trace "*** Migrate_FreeEndCut ... migrate port2 : " & Debug_ObjectName(oNewPort2, True)
    End If

    If Not bNewPort Then
        zSM_trace "*** Migrate_FreeEndCut *** no migration of input ports required : "
    End If
End If

    ' Check/Migrate Design Parent
    Dim oParentNew As Object
    Migrate_SystemParent oToBeMigrated, oMigrateHelper, oParentNew
        
    If bNewPort Then
        Migrate_SetReplacedReplacing oToBeMigrated, oToBeMigrated, oMigrateHelper
    End If
        
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_FreeEndCut"
    Err.Raise LogError(Err, MODULE, "Migrate_FreeEndCut", strError).Number
End Sub

Public Sub Migrate_WebCut(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
        On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJStructFeature Then
        Exit Sub
    Else
        sMemberObjectType = "IJStructFeature"
    End If
    
    Dim oStructFeature As IJStructFeature
    Dim eStructFeatureType As StructFeatureTypes
    
    Dim oResourceManager As IUnknown
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
    Dim oSDFeatureAttributes As IJSDFeatureAttributes
    
    Set oStructFeature = oToBeMigrated
    eStructFeatureType = oStructFeature.get_StructFeatureType
    If Not eStructFeatureType = SF_WebCut Then
        Exit Sub
    End If
    
    Set oSDFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
    Set oSDFeatureAttributes = oSDFeatureUtils
    
    Dim oBounded As Object
    Dim oBounding As Object
    Dim oNewBounded As Object
    Dim oNewBounding As Object
    Dim oNewBoundedPart As Object
    Dim oNewBoundingPart As Object
    
    sMemberObjectType = "SF_WebCut"
    oSDFeatureAttributes.get_WebCutInputs oToBeMigrated, oBounding, oBounded
        
    ' Check/Migrate Inputs
    Migrate_GetReplacingObject oBounding, oMigrateHelper, oNewBounding, oNewBoundingPart
    
    Migrate_GetReplacingObject oBounded, oMigrateHelper, oNewBounded, oNewBoundedPart
        
    ' Check/Migrate Design Parent
    Dim oParentNew As Object
    Dim oSystemNew As IJSystem
    Migrate_SystemParent oToBeMigrated, oMigrateHelper, oParentNew
    
    Set oSystemNew = Nothing
    If oParentNew Is Nothing Then
    ElseIf TypeOf oParentNew Is IJSystem Then
        Set oSystemNew = oParentNew
    End If
        
    ' Check/Migrate Inputs
    Dim oModWebCut As Object
    If Not oBounding Is oNewBounding Then
    ElseIf Not oBounded Is oNewBounded Then
    Else
        Exit Sub
    End If
    
    strError = "Replacing EndCut Inputs ... WebCut"
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oResourceManager = oSDO_Helper.GetResourceManagerFromObject(oToBeMigrated)
        
    Set oModWebCut = oSDFeatureUtils.CreateWebCut(oResourceManager, oNewBounding, oNewBounded, _
                                                  vbNullString, oSystemNew, oToBeMigrated)
    If Not oBounded Is oNewBounded Then
        ' The Feature's Opt/Opr pair might change when placed during migration
        ' ... Need to save the Feature's Member Part, Opt/Opr pair Before migration
        ' ... Need to save the Feature's Member Part, Opt/Opr pair After migration
        ' ... This data is used by Migrate_GetOptOprFromFeature and Migrate_CreateReplacingPort
        ' ... for creating the correct Late Binding Ports based on the Feature
        Migrate_SetFeatureMigrateData oToBeMigrated, oBounded, True

        EndCut_FinalConstruct oModWebCut
        
        Migrate_SetFeatureMigrateData oToBeMigrated, oNewBounded, False
        
    End If
        
    ForceUpdateOnMemberObjects oModWebCut
        
If bSM_trace Then
    zSM_trace "Migrate_WebCut ...   : " & Debug_ObjectName(oModWebCut, True)
    If Not oBounded Is oNewBounded Then
        zSM_trace "*** ... ...    oBounding : " & Debug_PortData(oNewBounding, True)
    End If
    
    If Not oBounded Is oNewBounded Then
        zSM_trace "*** *** ... ... oBounded : " & Debug_PortData(oNewBounded, True)
    End If
End If
        
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_WebCut"
    Err.Raise LogError(Err, MODULE, "Migrate_WebCut", strError).Number
End Sub

Public Sub Migrate_FlangeCut(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJStructFeature Then
        Exit Sub
    Else
        sMemberObjectType = "IJStructFeature"
    End If
    
    Dim oStructFeature As IJStructFeature
    Dim eStructFeatureType As StructFeatureTypes
    
    Dim oResourceManager As IUnknown
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
    Dim oSDFeatureAttributes As IJSDFeatureAttributes
    
    Set oStructFeature = oToBeMigrated
    eStructFeatureType = oStructFeature.get_StructFeatureType
    If Not eStructFeatureType = SF_FlangeCut Then
        Exit Sub
    End If
    
    Set oSDFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
    Set oSDFeatureAttributes = oSDFeatureUtils
    
    Dim oWebCut As Object
    Dim oBounded As Object
    Dim oBounding As Object
    Dim oNewWebCut As Object
    Dim oNewBounded As Object
    Dim oNewBounding As Object
    Dim oNewBoundedPart As Object
    Dim oNewBoundingPart As Object
    
    sMemberObjectType = "SF_FlangeCut"
    oSDFeatureAttributes.get_FlangeCutInputs oToBeMigrated, oBounding, oBounded, oWebCut
        
    ' Check/Migrate Inputs
    Migrate_GetReplacingObject oBounding, oMigrateHelper, oNewBounding, oNewBoundingPart
        
    Migrate_GetReplacingObject oBounded, oMigrateHelper, oNewBounded, oNewBoundedPart
        
    ' do not expect webcuts to be Replaced
    ' ... Migrate_GetReplacingObject oWebCut, oMigrateHelper, oNewWebCut, oNewWebCutPart
    Set oNewWebCut = oWebCut
        
    ' Check/Migrate Design Parent
    Dim oParentNew As Object
    Dim oSystemNew As IJSystem
    Migrate_SystemParent oToBeMigrated, oMigrateHelper, oParentNew
    
    Set oSystemNew = Nothing
    If oParentNew Is Nothing Then
    ElseIf TypeOf oParentNew Is IJSystem Then
        Set oSystemNew = oParentNew
    End If
        
    ' Check/Migrate Inputs
    Dim oModFlangeCut As Object
    If Not oBounding Is oNewBounding Then
    ElseIf Not oBounded Is oNewBounded Then
    Else
        Exit Sub
    End If
    
    strError = "Replacing EndCut Inputs ... FlangeCut"
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oResourceManager = oSDO_Helper.GetResourceManagerFromObject(oToBeMigrated)
    
    Set oModFlangeCut = oSDFeatureUtils.CreateFlangeCut(oResourceManager, oNewBounding, oNewBounded, oNewWebCut, _
                                                     vbNullString, oSystemNew, oToBeMigrated)
    
    If Not oBounded Is oNewBounded Then
    
        ' The Feature's Opt/Opr pair might change when placed during migration
        ' ... Need to save the Feature's Member Part, Opt/Opr pair Before migration
        ' ... Need to save the Feature's Member Part, Opt/Opr pair After migration
        ' ... This data is used by Migrate_GetOptOprFromFeature and Migrate_CreateReplacingPort
        ' ... for creating the correct Late Binding Ports based on the Feature
        Migrate_SetFeatureMigrateData oToBeMigrated, oBounded, True

        EndCut_FinalConstruct oModFlangeCut
        
        Migrate_SetFeatureMigrateData oToBeMigrated, oNewBounded, False
        
    End If

    ForceUpdateOnMemberObjects oModFlangeCut

If bSM_trace Then
    zSM_trace "Migrate_FlangeCut ... : " & Debug_ObjectName(oModFlangeCut, True)
    If Not oBounded Is oNewBounded Then
        zSM_trace "*** ... ...     oBounding : " & Debug_PortData(oNewBounding, True)
    End If
    
    If Not oBounded Is oNewBounded Then
        zSM_trace "*** *** ... ...  oBounded : " & Debug_PortData(oNewBounded, True)
    End If
End If
    
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_FlangeCut"
    Err.Raise LogError(Err, MODULE, "Migrate_FlangeCut", strError).Number
End Sub

Public Sub Migrate_Bearing(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    Dim bSetInputs As Boolean
    Dim bPlaceBearing As Boolean
    Dim oCollectionAlias As JCmnShp_CollectionAlias

    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJSmartPlate Then
        Exit Sub
    Else
        sMemberObjectType = "IJSmartPlate"
    End If
    
    Dim oSmartPlate As IJSmartPlate
    Dim eSmartPlateType As SmartPlateTypes
    
    Dim oCustomGeometry As Collection
    Dim oStructCustomGeometry As IJDStructCustomGeometry
    
    Dim oResourceManager As IUnknown
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSPDefinition As GSCADSDCreateModifyUtilities.IJSDSmartPlateDefinition
    Dim oSDSmartPlateAtt As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
    Dim oSDSmartPlateOps As GSCADSDCreateModifyUtilities.IJSDSmartPlateOperations
    
    Set oSmartPlate = oToBeMigrated
    eSmartPlateType = oSmartPlate.SmartPlateType
    If Not eSmartPlateType = spType_BEARING Then
        Exit Sub
    End If
    
    Set oSDSmartPlateAtt = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    
    Dim oBounded As Object
    Dim oBounding As Object
    Dim oNewWebCut As Object
    Dim oNewBounded As Object
    Dim oNewBounding As Object
    Dim oNewBoundedPart As Object
    Dim oNewBoundingPart As Object
    
    sMemberObjectType = "spType_BEARING"
    oSDSmartPlateAtt.GetInputs_BearingPlate oToBeMigrated, oCollectionAlias
        
    ' Check/Migrate Inputs
    Set oBounding = oCollectionAlias.Item(1)
    Set oBounded = oCollectionAlias.Item(2)
    
    Migrate_GetReplacingObject oBounding, oMigrateHelper, oNewBounding, oNewBoundingPart
        
    Migrate_GetReplacingObject oBounded, oMigrateHelper, oNewBounded, oNewBoundedPart
        
    ' Check/Migrate Design Parent
    Dim oParentNew As Object
    Migrate_SystemParent oToBeMigrated, oMigrateHelper, oParentNew
        
    ' Check/Migrate Inputs
    Dim oModBearing As Object
    bSetInputs = False
    bPlaceBearing = False
    If Not oBounding Is oNewBounding Then
        bSetInputs = True
    ElseIf Not oBounded Is oNewBounded Then
        bSetInputs = True
    End If
    
    If bSetInputs Then
        strError = "Replacing Bearing Inputs ..."
        bPlaceBearing = True
        
        Set oSDO_Helper = New StructDetailObjects.Helper
        Set oResourceManager = oSDO_Helper.GetResourceManagerFromObject(oToBeMigrated)
            
        Set oCollectionAlias = New Collection
        oCollectionAlias.Add oNewBounding
        oCollectionAlias.Add oNewBounded
        
        Set oSPDefinition = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
        Set oModBearing = oSPDefinition.CreateBearingPlatePart(oResourceManager, vbNullString, _
                                                               oCollectionAlias, Nothing, oToBeMigrated)
    Else
        Set oModBearing = oToBeMigrated
    End If
    
    ' Placment of Bearing Plate is dependent on BOTH Bounded and Bounding
    If bSetInputs Then
    ElseIf TypeOf oModBearing Is IJDStructCustomGeometry Then
        ' Inputs were not changed
        ' But ...
        ' ... The Bounding could have already been replaced
        ' ... Check the CreatePlatePart_AE Inputs are as expected
        ' ... 3 inputs ... Bounding, Bounded, and BearingPlate
        Set oCustomGeometry = New Collection
        Set oStructCustomGeometry = oModBearing
        oStructCustomGeometry.GetCustomGeometry "CreatePlatePart.CreatePlatePart_AE.1", oCustomGeometry
        If oCustomGeometry Is Nothing Then
            bPlaceBearing = True
        ElseIf oCustomGeometry.Count < 3 Then
            bPlaceBearing = True
        Else
        End If
    End If
        
    If bPlaceBearing Then
        Set oSDSmartPlateOps = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
        oSDSmartPlateOps.PlaceBearingPlate oModBearing
        
        ForceUpdateOnMemberObjects oModBearing
    End If
        
If bSM_trace Then
    zSM_trace "Migrate_Bearing ... : " & Debug_ObjectName(oModBearing, True)
    If bSetInputs Then
        If Not oBounding Is oNewBounding Then
            zSM_trace "*** ... ...   oBounding : " & Debug_PortData(oNewBounding, True)
            bPlaceBearing = False
        End If
        
        If Not oBounded Is oNewBounded Then
            zSM_trace "*** *** ... .  oBounded : " & Debug_PortData(oNewBounded, True)
            bPlaceBearing = False
        End If
    
    ElseIf bPlaceBearing Then
        zSM_trace "*** *** ... .  PlaceBearingPlate (Inputs are set): "
    End If
    
End If
    
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_Bearing"
    Err.Raise LogError(Err, MODULE, "Migrate_Bearing", strError).Number
End Sub

Public Sub Migrate_CornerFeature(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJStructFeature Then
        Exit Sub
    Else
        sMemberObjectType = "IJStructFeature"
    End If
    
    Dim oStructFeature As IJStructFeature
    Dim eStructFeatureType As StructFeatureTypes
    
    Dim oResourceManager As IUnknown
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
    Dim oSDFeatureAttributes As IJSDFeatureAttributes
    Dim oSDFeatureDefinition As IJSDFeatureDefinition
    
    Set oStructFeature = oToBeMigrated
    eStructFeatureType = oStructFeature.get_StructFeatureType
    If Not eStructFeatureType = SF_CornerFeature Then
        Exit Sub
    End If
    
    Set oSDFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
    Set oSDFeatureAttributes = oSDFeatureUtils
    
    Dim oFacePort As Object
    Dim oEdgePort1 As Object
    Dim oEdgePort2 As Object
    Dim oNewFacePort As Object
    Dim oNewEdgePort1 As Object
    Dim oNewEdgePort2 As Object
    Dim oNewFacePart As Object
    Dim oNewEdgePart1 As Object
    Dim oNewEdgePart2 As Object
    
    sMemberObjectType = "SF_CornerFeature"
    oSDFeatureAttributes.get_CornerCutInputsEx2 oToBeMigrated, oFacePort, oEdgePort1, oEdgePort2
        
    ' Check/Migrate Inputs
    Migrate_GetReplacingObject oFacePort, oMigrateHelper, oNewFacePort, oNewFacePart
        
    Migrate_GetReplacingObject oEdgePort1, oMigrateHelper, oNewEdgePort1, oNewEdgePart1
        
    Migrate_GetReplacingObject oEdgePort2, oMigrateHelper, oNewEdgePort2, oNewEdgePart2
        
    ' Check/Migrate Design Parent
    Dim oParentNew As Object
    Migrate_SystemParent oToBeMigrated, oMigrateHelper, oParentNew
        
    ' Check/Migrate Inputs
    Dim oModCorner As Object
    If Not oFacePort Is oNewFacePort Then
    ElseIf Not oEdgePort1 Is oNewEdgePort1 Then
    ElseIf Not oEdgePort2 Is oNewEdgePort2 Then
    Else
        Exit Sub
    End If
    
    strError = "Replacing StructFeature Inputs ... Corner Feature"
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oResourceManager = oSDO_Helper.GetResourceManagerFromObject(oToBeMigrated)
        
    Set oModCorner = oSDFeatureUtils.CreateCornerCutEx(oResourceManager, oNewFacePort, oNewEdgePort1, oNewEdgePort2, _
                                                       vbNullString, oParentNew, oToBeMigrated)
    
    ' PlaceFeature
    ' The Feature's Opt/Opr pair might change when placed during migration
    ' ... Need to save the Feature's Member Part, Opt/Opr pair Before migration
    ' ... Need to save the Feature's Member Part, Opt/Opr pair After migration
    ' ... This data is used by Migrate_GetOptOprFromFeature and Migrate_CreateReplacingPort
    ' ... for creating the correct Late Binding Ports based on the Feature
    Migrate_SetFeatureMigrateData oToBeMigrated, oFacePort, True

    Set oSDFeatureDefinition = New SDFeatureUtils
    oSDFeatureDefinition.PlaceFeature oModCorner, oNewFacePart
        
    Migrate_SetFeatureMigrateData oToBeMigrated, oNewFacePort, False
        
If bSM_trace Then
    zSM_trace "Migrate_CornerFeature ... : " & Debug_ObjectName(oModCorner, True)
    If Not oFacePort Is oNewFacePort Then
        zSM_trace "*** *** ...     FacePort      : " & Debug_PortData(oNewFacePort, True)
    End If
    
    If Not oEdgePort1 Is oNewEdgePort1 Then
        zSM_trace "*** *** ... ... EdgePort1 : " & Debug_PortData(oNewEdgePort1, True)
    End If
    
    If Not oEdgePort2 Is oNewEdgePort2 Then
        zSM_trace "*** *** ... ... EdgePort2 : " & Debug_PortData(oNewEdgePort2, True)
    End If
End If
    
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_CornerFeature"
    Err.Raise LogError(Err, MODULE, "Migrate_CornerFeature", strError).Number
End Sub

Public Sub Migrate_PC(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJStructPhysicalConnection Then
        Exit Sub
    Else
        sMemberObjectType = "IJStructPhysicalConnection"
    End If
    
    Dim oSDConnectionAttributes As IJSDConnectionAttributes
    Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
    
    Dim oPort1 As Object
    Dim oPort2 As Object
    Dim oNewPort1 As Object
    Dim oNewPort2 As Object
    Dim oNewPart1 As Object
    Dim oNewPart2 As Object
    
    sMemberObjectType = "IJStructPhysicalConnection"
    Set oSDConnectionAttributes = New SDConnectionUtils
    oSDConnectionAttributes.get_ConnectionInputs oToBeMigrated, oPort1, oPort2
        
    ' Check/Migrate Inputs
    Migrate_GetReplacingObject oPort1, oMigrateHelper, oNewPort1, oNewPart1
        
    Migrate_GetReplacingObject oPort2, oMigrateHelper, oNewPort2, oNewPart2
        
    ' Check/Migrate Design Parent
    Dim oParentNew As Object
    Migrate_SystemParent oToBeMigrated, oMigrateHelper, oParentNew
        
    ' Check/Migrate Inputs
    If Not oPort1 Is oNewPort1 Then
    ElseIf Not oPort2 Is oNewPort2 Then
    Else
        Exit Sub
    End If
    
    strError = "Replacing PC Inputs "
    Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
        
    If Not oPort1 Is oNewPort1 Then
        oConnectionDefinition.ReplacePhysicalConnectionPort oToBeMigrated, oPort1, oNewPort1
    End If
        
    If Not oPort2 Is oNewPort2 Then
        oConnectionDefinition.ReplacePhysicalConnectionPort oToBeMigrated, oPort2, oNewPort2
    End If
    
    ' Flag Physical Connection as being in Split\Migration state
    ' such that the semnatic will NOT mark it to be deleted if the geometry is not valid (one time only)
    Migrate_SetMigrateControlFlag oToBeMigrated, True
    
If bSM_trace Then
    zSM_trace "Migrate_PC ... : " & Debug_ObjectName(oToBeMigrated, True)
    If Not oPort1 Is oNewPort1 Then
        zSM_trace "*** ... ...  Port1 : " & Debug_PortData(oNewPort1, True)
    End If
        
    If Not oPort2 Is oNewPort2 Then
        zSM_trace "*** *** ... .Port2 : " & Debug_PortData(oNewPort2, True)
    End If
End If
    
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_PC"
    Err.Raise LogError(Err, MODULE, "Migrate_PC", strError).Number
End Sub

Public Sub Migrate_SystemParent(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper, oParentNew As Object)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    
    Dim bIsDeleted As Boolean
    
    Dim oParentObj As Object
    Dim oSmartParent As Object
    Dim oDesignChild As IJDesignChild
    Dim oSystemChild As IJSystemChild
    Dim oDesignParent As IJDesignParent
    Dim oObjectReplacing As Object
    
    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection

    Set oParentNew = Nothing
    If oToBeMigrated Is Nothing Then
        Exit Sub
    End If
    
    ' get current Object's IJDesignParent (Parent System)
    ' Note: BearingPlate does NOT support IJDesignParent
    '       ... therefore can NOT be a valid parent
    '       ... Use the oToBeMigrated true IJDesignChild parent
    Set oParentObj = Nothing
    If TypeOf oToBeMigrated Is IJSmartOccurrence Then
        GetSmartOccurrenceParent oToBeMigrated, oSmartParent
        If oSmartParent Is Nothing Then
            If TypeOf oToBeMigrated Is IJDesignChild Then
                Set oDesignChild = oToBeMigrated
                Set oParentObj = oDesignChild.GetParent
            ElseIf TypeOf oToBeMigrated Is IJSystemChild Then
                Set oSystemChild = oToBeMigrated
                Set oParentObj = oSystemChild.GetParent
            End If
        ElseIf TypeOf oSmartParent Is IJDesignChild Then
            Set oParentObj = oSmartParent
        End If
    End If
    
    If Not oParentObj Is Nothing Then
    ElseIf TypeOf oToBeMigrated Is IJDesignChild Then
        Set oDesignChild = oToBeMigrated
        Set oParentObj = oDesignChild.GetParent
    ElseIf TypeOf oToBeMigrated Is IJSystemChild Then
        Set oSystemChild = oToBeMigrated
        Set oParentObj = oSystemChild.GetParent
    Else
        Exit Sub
    End If
    
    ' check if current Parent is in current list of ReplacedParts
    For iIndex = 1 To m_ReplacedParts.Count
        If oParentObj Is m_ReplacedParts.Item(iIndex) Then
            Set oParentNew = m_ReplacingParts.Item(iIndex)
            Exit For
        End If
    Next iIndex
                
    ' check if current Parent is in current list of Replaced/Replacing Migration data
    If oParentNew Is Nothing Then
        Set oParentNew = oParentObj
        oMigrateHelper.ObjectsReplacing oParentObj, oObjectCollectionReplacing, bIsDeleted
        If oObjectCollectionReplacing Is Nothing Then
        ElseIf oObjectCollectionReplacing.Count < 1 Then
        Else
            For iIndex = 1 To m_ReplacingParts.Count
                If oObjectReplacing Is m_ReplacingParts.Item(iIndex) Then
                    Set oParentNew = oObjectReplacing
                    Exit For
                End If
            Next iIndex
        End If
    End If
    
    ' Update current objects IJDesignParent (Parent System)
    If oParentNew Is Nothing Then
    ElseIf Not oParentObj Is oParentNew Then
        If TypeOf oParentObj Is IJDesignParent Then
            Set oDesignParent = oParentObj
            oDesignParent.RemoveChild oToBeMigrated
        End If
        
        If TypeOf oParentNew Is IJDesignParent Then
            Set oDesignParent = oParentNew
            oDesignParent.AddChild oToBeMigrated
        Else
            Set oParentNew = Nothing
        End If
    ElseIf TypeOf oParentNew Is IJDesignParent Then
    Else
        Set oParentNew = Nothing
    End If
    
    
If bSM_trace Then
    If Not oParentObj Is oParentNew Then
        zSM_trace "*** Migrate_SystemParent ... migrate parent : " & Debug_ObjectName(oParentNew, True)
    End If
End If
    
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_SystemParent"
    Err.Raise LogError(Err, MODULE, "Migrate_SystemParent", strError).Number
End Sub

Public Sub Migrate_InsertPlateSlot(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJStructFeature Then
        Exit Sub
    Else
        sMemberObjectType = "IJStructFeature"
    End If
    
    Dim oStructFeature As IJStructFeature
    Dim eStructFeatureType As StructFeatureTypes
    
    Dim oResourceManager As IUnknown
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
    Dim oSDFeatureAttributes As IJSDFeatureAttributes
    Dim oSDFeatureDefinition As IJSDFeatureDefinition
    
    Set oStructFeature = oToBeMigrated
    eStructFeatureType = oStructFeature.get_StructFeatureType
    If Not eStructFeatureType = SF_Slot Then
        Exit Sub
    End If
    
    Set oSDFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
    Set oSDFeatureAttributes = oSDFeatureUtils
    
    Dim oBasePlate As Object
    Dim oPenetrated As Object
    Dim oPenetrating As Object
    Dim oNewBasePlate As Object
    Dim oNewPenetrated As Object
    Dim oNewPenetrating As Object
    Dim oNewBasePlatePart As Object
    Dim oNewPenetratedPart As Object
    Dim oNewPenetratingPart As Object
    
    Dim oPort As IJPort
    Dim oPenetratedObj As Object
    Dim oPenetratedUnk As IUnknown
    Dim oPenetratedGeometry As IUnknown
    Dim oStructGeometryHelper As StructGeometryHelper
            
    sMemberObjectType = "SF_Slot"
    oSDFeatureAttributes.get_SlotInputs oToBeMigrated, oPenetrating, oPenetrated, oBasePlate
            
    Dim oSmartParent As Object
    Dim oSmartObject As Object
    Set oSmartObject = oToBeMigrated
    While Not oSmartObject Is Nothing
        Set oSmartParent = Nothing
        GetSmartOccurrenceParent oSmartObject, oSmartParent
        
        Set oSmartObject = Nothing
        If oSmartParent Is Nothing Then
        ElseIf TypeOf oSmartParent Is IJStructAssemblyConnection Then
        ElseIf TypeOf oSmartParent Is IJSmartOccurrence Then
            Set oSmartObject = oSmartParent
        End If
    Wend
            
    ' The penetrated object input is the geometry, change it to the part
    Set oPenetratedGeometry = oPenetrated
    Set oStructGeometryHelper = New StructGeometryHelper
    oStructGeometryHelper.RecursiveGetStructEntityIUnknown oPenetratedGeometry, oPenetratedUnk
    Set oPenetratedObj = oPenetratedUnk
        
    ' Check/Migrate Inputs (the slots inputs are Parts?)
    Migrate_GetReplacingObject oPenetrating, oMigrateHelper, oNewPenetrating, oNewPenetratingPart
        
    Migrate_GetReplacingObject oPenetratedObj, oMigrateHelper, oNewPenetrated, oNewPenetratedPart
    
    If oSmartParent Is Nothing Then
    ElseIf TypeOf oSmartParent Is IJStructAssemblyConnection Then
        Set oNewPenetrated = Nothing
        Set oNewPenetrating = Nothing
        GetPenetratedAndPenetratingPorts oSmartParent, oNewPenetrated, oNewPenetrating
        If TypeOf oNewPenetrated Is IJPort Then
            Set oPort = oNewPenetrated
            Set oNewPenetratedPart = oPort.Connectable
            If Not Migrate_IsPortGeometryValid(oPort) Then
                Migrate_SetReplacedReplacing oToBeMigrated, oToBeMigrated, oMigrateHelper
                Migrate_SetReplacedReplacing oSmartParent, oSmartParent, oMigrateHelper
                Exit Sub
            End If
        End If
        
        If TypeOf oNewPenetrating Is IJPort Then
            Set oPort = oNewPenetrating
            Set oNewPenetratingPart = oPort.Connectable
            If Not Migrate_IsPortGeometryValid(oPort) Then
                Migrate_SetReplacedReplacing oToBeMigrated, oToBeMigrated, oMigrateHelper
                Migrate_SetReplacedReplacing oSmartParent, oSmartParent, oMigrateHelper
                Exit Sub
            End If
        End If
    End If
    
    ' BasePlate is not used in creating/modifing the Slot
    ' ... Migrate_GetReplacingObject oBasePlate, oMigrateHelper, oNewBasePlate, oNewBasePlatePart
        
    ' Check/Migrate Design Parent
    Dim oParentNew As Object
    Dim oSystemNew As IJSystem
    Migrate_SystemParent oToBeMigrated, oMigrateHelper, oParentNew
    
    Set oSystemNew = Nothing
    If oParentNew Is Nothing Then
    ElseIf TypeOf oParentNew Is IJSystem Then
        Set oSystemNew = oParentNew
    End If
        
    ' Check/Migrate Inputs
    Dim oModSlot As Object
    If Not oPenetrating Is oNewPenetratingPart Then
    ElseIf Not oPenetratedObj Is oNewPenetratedPart Then
    Else
        Exit Sub
    End If
    
    strError = "Replacing StructFeature Inputs ... Slot"
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oResourceManager = oSDO_Helper.GetResourceManagerFromObject(oToBeMigrated)
        
    Set oModSlot = oSDFeatureUtils.CreateSlot(oResourceManager, oNewPenetratingPart, oNewPenetratedPart, _
                                              vbNullString, oSystemNew, oToBeMigrated)
        
    ' Add the Modified Slot to operations
    If Not oPenetratedObj Is oNewPenetratedPart Then
        
        ' The Feature's Opt/Opr pair might change when placed during migration
        ' ... Need to save the Feature's Member Part, Opt/Opr pair Before migration
        ' ... Need to save the Feature's Member Part, Opt/Opr pair After migration
        ' ... This data is used by Migrate_GetOptOprFromFeature and Migrate_CreateReplacingPort
        ' ... for creating the correct Late Binding Ports based on the Feature
        Migrate_SetFeatureMigrateData oToBeMigrated, oPenetratedObj, True
    
''        ' Remove the InsertPlateSlot from the Replaced Penetrated Member Part
''        Dim oOperationAE As Object
''        Dim OperationPattern As IJStructOperationPattern
''        Dim oCollectionOfOperators As IJElements
''        If TypeOf oPenetratedObj Is IJStructOperationPattern Then
''            Set OperationPattern = oPenetratedObj
''            OperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oOperationAE
''            If oCollectionOfOperators Is Nothing Then
''            ElseIf oCollectionOfOperators.Count < 1 Then
''            Else
''                For iIndex = 1 To oCollectionOfOperators.Count
''                    If oCollectionOfOperators.Item(iIndex) Is oModSlot Then
''                        oCollectionOfOperators.Remove (iIndex)
''                        OperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oOperationAE
''                        Exit For
''                    End If
''                Next iIndex
''            End If
''        End If
        
        Set oSDFeatureDefinition = New SDFeatureUtils
        oSDFeatureDefinition.PlaceFeature oModSlot, oNewPenetratedPart
            
        Migrate_SetFeatureMigrateData oToBeMigrated, oNewPenetratedPart, False
        
    End If
    
If bSM_trace Then
    zSM_trace "Migrate_InsertPlateSlot ... : " & Debug_ObjectName(oModSlot, True)
    zSM_trace "*** *** ...     Penetrating     : " & Debug_ObjectName(oNewPenetratingPart, True)
    If Not oPenetratedObj Is oNewPenetratedPart Then
        zSM_trace "*** *** ... ... oPenetrated : " & Debug_ObjectName(oNewPenetratedPart, True)
    End If
End If
    
    ' Flag the InsertPlateSlot as Migrated
    Migrate_SetReplacedReplacing oModSlot, oModSlot, oMigrateHelper
    Migrate_SetReplacedReplacing oSmartParent, oSmartParent, oMigrateHelper
    
If bSM_trace Then
    'Get Endcuts on pattern
    ' ... SPSMembers.SPSPartPrismaticGenerator.1
    'Get Cutouts on pattern
    ' ... SP3DStructGeneric.StructCutoutOperation.1 ... StructGeneric.StructCutoutOperationAE.1
    Migrate_TraceOperators oNewPenetratedPart
End If
    
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_InsertPlateSlot"
    Err.Raise LogError(Err, MODULE, "Migrate_InsertPlateSlot", strError).Number
End Sub

Public Sub Migrate_InsertPlateChamfer(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJChamfer Then
        Exit Sub
    Else
        sMemberObjectType = "IJChamfer"
    End If
    
    Dim oResourceManager As IUnknown
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oChamferUtils As GSCADSDCreateModifyUtilities.SDChamferUtils
    Dim oSDChamferAttributes As IJSDChamferAttributes
    Dim oSDChamferDefinition As IJSDChamferDefinition
    
    Set oChamferUtils = New GSCADSDCreateModifyUtilities.SDChamferUtils
    Set oSDChamferAttributes = oChamferUtils
    
    Dim oPort1 As Object
    Dim oPort2 As Object
    Dim oPart1 As Object
    Dim oPart2 As Object
    Dim oNewPort1 As Object
    Dim oNewPort2 As Object
    Dim oNewPart1 As Object
    Dim oNewPart2 As Object
    
    Dim oPort As IJPort
    Dim oChamfer As IJChamfer
        
    sMemberObjectType = "IJChamfer"
    Set oChamfer = oToBeMigrated
    oSDChamferAttributes.get_ChamferInputs oChamfer, oPort1, oPort2
    
    Dim oSmartParent As Object
    Dim oSmartObject As Object
    Set oSmartObject = oToBeMigrated
    While Not oSmartObject Is Nothing
        Set oSmartParent = Nothing
        GetSmartOccurrenceParent oSmartObject, oSmartParent
        
        Set oSmartObject = Nothing
        If oSmartParent Is Nothing Then
        ElseIf TypeOf oSmartParent Is IJStructAssemblyConnection Then
        ElseIf TypeOf oSmartParent Is IJSmartOccurrence Then
            Set oSmartObject = oSmartParent
        End If
    Wend
    
    If oPort1 Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oPort1 Is IJPort Then
        Exit Sub
    Else
        Set oPort = oPort1
        Set oPart1 = oPort.Connectable
    End If
    
    If oPort2 Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oPort2 Is IJPort Then
        Exit Sub
    Else
        Set oPort = oPort2
        Set oPart2 = oPort.Connectable
    End If
    
    ' Check/Migrate Inputs (the slots inputs are Parts?)
    Migrate_GetReplacingObject oPort1, oMigrateHelper, oNewPort1, oNewPart1
        
    Migrate_GetReplacingObject oPort2, oMigrateHelper, oNewPort2, oNewPart2
        
    ' Expecting only one(1) to be ISPSMemberPartCommon
    Dim oPlatePart As Object
    Dim oPlatePort As Object
    Dim oMemberPart As Object
    Dim oNewMemberPart As Object
    Dim oNewMemberPort As Object
    Dim oStructGeomBasicPort As StructGeomBasicPort
    If TypeOf oNewPart1 Is ISPSMemberPartCommon Then
        Set oPlatePart = oNewPart2
        Set oPlatePort = oNewPort2
        Set oMemberPart = oPart1
        Set oNewMemberPart = oNewPart1
        Set oStructGeomBasicPort = oNewPort1
    ElseIf TypeOf oNewPart2 Is ISPSMemberPartCommon Then
        Set oPlatePart = oNewPart1
        Set oPlatePort = oNewPort1
        Set oMemberPart = oPart2
        Set oNewMemberPart = oNewPart2
        Set oStructGeomBasicPort = oNewPort2
    Else
        Exit Sub
    End If
        
    ' Determine the Member Port (Xid) used in the Chamfer
    ' ... expecting Web_left (257) or Web_Right (258)
    Dim lXId As Long
    Dim lOprId As Long
    Dim lOptId As Long
    Dim lCtxId As Long
    Dim lPortType As Long
    Dim lReplacingOptId As Long
    Dim lReplacingOprId As Long
    
    Dim oOperation As Object
    Dim oCutFeature As Object
    Dim eStructOperation As cmnstrStructOperation
    Dim oStructOperationAE As IJStructOperationAE
    
    Dim oPortMoniker As IUnknown
    Dim oSP3D_StructPort As SP3DStructPorts.IJStructPort
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    
    Dim oACTools As AssemblyConnectionTools
    Dim oPortHelper As PORTHELPERLib.PortHelper
    
    Set oPort = oStructGeomBasicPort
    Set oSP3D_StructPort = oStructGeomBasicPort
    Set oPortMoniker = oSP3D_StructPort.PortMoniker
    Set oPortHelper = New PORTHELPERLib.PortHelper
    oPortHelper.DecodeTopologyProxyMoniker oPortMoniker, lPortType, lCtxId, lOptId, lOprId, lXId
    
    'get the Member Chamfered Port from the Slot Cut-out Port(Late Port)
    Migrate_GetFeatureFromOptOpr oMemberPart, lOptId, lOprId, oCutFeature, oOperation
    If oOperation Is Nothing Then
    Else
        Set oStructOperationAE = oOperation
        oStructOperationAE.GetOperationType eStructOperation
    End If
    
    ' No Feature was found on the Replaced Part with the given Opt,Opr
    ' check if the Feature has been moved to the Replacing Part
    If oCutFeature Is Nothing Then
        Migrate_GetFeatureFromOptOpr oNewMemberPart, lOptId, lOprId, oCutFeature, oOperation
    End If
    
    Set oACTools = New AssemblyConnectionTools
    If oCutFeature Is Nothing Then
        ' NO CutFeature (Slot) found
        ' do we want Port after EndCuts or after Cut operation(?)
        oACTools.GetBindingPort oNewMemberPart, JS_TOPOLOGY_PROXY_LFACE, lOptId, lOprId, lCtxId, lXId, _
                                "SPSMembers.SPSPartPrismaticGenerator", oNewMemberPort
    
    ElseIf eStructOperation = cmnstrCutoutOperation Then
        ' GetLatePortForFeatureSegment returns the Port based on "last geometry"
        ' ... valid for Webcut, Flangecut and Slot features only
        ' ... calls CStructSymbolTools::BindMonikerToStructLastPort
        ''        Set oStructProfilePart = oNewMemberPart
        ''        Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
        ''        oStructEndCutUtil.GetLatePortForFeatureSegment oCutFeature, lXId, oNewMemberPort
        Migrate_GetOptOprFromFeature oCutFeature, oNewMemberPart, lReplacingOptId, lReplacingOprId, oOperation
        oACTools.GetBindingPort oNewMemberPart, JS_TOPOLOGY_PROXY_LFACE, lReplacingOptId, lReplacingOprId, lCtxId, lXId, _
                                "StructGeneric.StructCutoutOperationAE.1", oNewMemberPort
        
    
    Else
        Migrate_GetOptOprFromFeature oCutFeature, oNewMemberPart, lReplacingOptId, lReplacingOprId, oOperation
        oACTools.GetBindingPort oNewMemberPart, JS_TOPOLOGY_PROXY_LFACE, lReplacingOptId, lReplacingOprId, lCtxId, lXId, _
                                "SPSMembers.SPSPartPrismaticGenerator", oNewMemberPort
    End If
    
    ' Check/Migrate Design Parent
    Dim oParentNew As Object
    Migrate_SystemParent oToBeMigrated, oMigrateHelper, oParentNew
        
    ' Check/Migrate Inputs
    Dim oModChamfer As Object
    If Not oPort1 Is oNewPort1 Then
    ElseIf Not oPort2 Is oNewPort2 Then
    Else
        Exit Sub
    End If
    
    strError = "Replacing Chamfer ... "
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oResourceManager = oSDO_Helper.GetResourceManagerFromObject(oToBeMigrated)
        
    Set oModChamfer = oChamferUtils.CreateModifyChamfer(oResourceManager, oPlatePort, oNewMemberPort, _
                                                        "", Nothing, oToBeMigrated)
    ' Add the Modified Chamfer to operations
    If Not oPort1 Is oPlatePort Then
        Set oSDChamferDefinition = New SDChamferUtils
        oSDChamferDefinition.PlaceChamfer oModChamfer, oPlatePart
    Else
        Set oSDChamferDefinition = New SDChamferUtils
        oSDChamferDefinition.PlaceChamfer oModChamfer, oPlatePart
    End If

If bSM_trace Then
    zSM_trace "Migrate_InsertPlateChamfer ... : " & Debug_ObjectName(oModChamfer, True)
    If Not oPort1 Is oNewPort1 Then
        zSM_trace "*** *** ...     Port1     : " & Debug_PortData(oNewPort1, True)
    End If
    
    If Not oPort2 Is oNewPort2 Then
        zSM_trace "*** *** ... ... Port2 : " & Debug_PortData(oNewPort2, True)
    End If
End If
    
    ' Flag the InsertPlateChamfer as Migrated
    Migrate_SetReplacedReplacing oModChamfer, oModChamfer, oMigrateHelper
    Migrate_SetReplacedReplacing oSmartParent, oSmartParent, oMigrateHelper
    
If bSM_trace Then
    'Get Endcuts on pattern
    ' ... SPSMembers.SPSPartPrismaticGenerator.1
    'Get Cutouts on pattern
    ' ... SP3DStructGeneric.StructCutoutOperation.1 ... StructGeneric.StructCutoutOperationAE.1
    Migrate_TraceOperators oNewPart1
End If
    
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_InsertPlateChamfer"
    Err.Raise LogError(Err, MODULE, "Migrate_InsertPlateChamfer", strError).Number
End Sub

Public Sub Migrate_ReferencesCollection(oToBeMigrated As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    sMemberObjectType = "?Type?"
    If oToBeMigrated Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oToBeMigrated Is IJSmartOccurrence Then
        Exit Sub
    Else
        sMemberObjectType = "ReferencesCollection"
    End If
    
    Dim bIsMigrated As Boolean
    Dim oReplacingPart As Object
    Dim oReplacedObject As Object
    Dim oReplacingObject As Object
    
    Dim RefCollection As Collection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection
    
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oToBeMigrated)
    If oReferencesCollection Is Nothing Then
        Exit Sub
    End If
    
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    If oEditJDArgument Is Nothing Then
        Exit Sub
    ElseIf oEditJDArgument.GetCount < 1 Then
        Exit Sub
    End If
    
    ' Loop thru each object in the oReferencesCollection
    bIsMigrated = False
    Set RefCollection = New Collection
    For iIndex = 1 To oEditJDArgument.GetCount
        Set oReplacedObject = oEditJDArgument.GetEntityByIndex(iIndex)
        If TypeOf oReplacedObject Is IJPort Then
            Migrate_GetReplacingObject oReplacedObject, oMigrateHelper, oReplacingObject, oReplacingPart
            If Not oReplacedObject Is oReplacingObject Then
If bSM_trace Then
    If Not bIsMigrated Then
        zSM_trace "*** Migrate_ReferencesCollection ... : " & Debug_ObjectName(oToBeMigrated, True)
    End If
    zSM_trace "*** *** ...  ReferencesCollection : " & Debug_PortData(oReplacingObject, True)
End If
                bIsMigrated = True
            End If
        Else
            Set oReplacingObject = oReplacedObject
        End If
        
        RefCollection.Add oReplacingObject
        Set oReplacedObject = Nothing
        
    Next iIndex
     
    If bIsMigrated Then
        oEditJDArgument.RemoveAll
        
        For iIndex = 1 To RefCollection.Count
            If Not RefCollection(iIndex) Is Nothing Then
                oEditJDArgument.SetEntity iIndex, RefCollection(iIndex)
            End If
        Next iIndex

    End If

    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_ReferencesCollection"
    Err.Raise LogError(Err, MODULE, "Migrate_ReferencesCollection", strError).Number
End Sub

' ******************************************************************************************
' ******************************************************************************************
' ******************************************************************************************

Public Sub Migrate_MemberItems(oParentMemberItem As Object, oMigrateHelper As IJMigrateHelper, _
                               bMigrateParent As Boolean, bMigrateMemberItems As Boolean, bMigrateMemberItemChildren As Boolean)
Const METHOD = "::Migrate_MemberItems"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim sMemberObjectName As String
    Dim sMemberObjectType As String
    
    Dim iIndex As Long
    Dim jIndex As Long
    Dim lDispId As Long
    Dim lMbrIdx As Long
    Dim nMemberItems As Long
    
    Dim oMemberItem As Object
    Dim oMemberObjects As IJDMemberObjects
    
    Dim oMemberDatatypes As Collection
    Dim oMemberDataObjects As Collection
    
    Dim oSaveReplacedParts As Collection
    Dim oSaveReplacingParts As Collection
    
    Dim sSmartItemDef As String
    Dim sMemberItemName As String
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oMemberDescription As IJDMemberDescription
    Dim oMemberDescriptions As IJDMemberDescriptions
    
    nMemberItems = 0
    Set oSaveReplacedParts = Nothing
    Set oSaveReplacingParts = Nothing
    
    ' check if processing just the given object
    If oParentMemberItem Is Nothing Then
        Exit Sub
    ElseIf TypeOf oParentMemberItem Is IJDMemberObjects Then
        If bMigrateMemberItems Then
            ' process the given object
            ' AND all of its MemberItems objects
            Set oMemberObjects = oParentMemberItem
            Set oMemberDescriptions = oMemberObjects.MemberDescriptions
            If oMemberObjects.Count > 0 Then
                nMemberItems = oMemberObjects.Count
            End If
        End If
    End If
        
    For iIndex = 0 To nMemberItems
        sSmartItemDef = ""
        sMemberItemName = ""
        Set oMemberItem = Nothing
        If iIndex = 0 Then
            ' First Item is the given SmartOccurrence that creates the dependent MemberItems
            Set oMemberItem = oParentMemberItem
            
        ElseIf Not oMemberObjects.Item(iIndex) Is Nothing Then
            ' next are the given SmartOccurrence dependent MemberItems
            Set oMemberItem = oMemberObjects.Item(iIndex)
            ' get the dependent MemberItem's Smart Occurrence Definition Name
            If TypeOf oMemberItem Is IJSmartOccurrence Then
                Set oSmartOccurrence = oMemberItem
                Set oSmartItem = oSmartOccurrence.ItemObject
                If Not oSmartItem Is Nothing Then
                    sSmartItemDef = oSmartItem.definition
                    If Len(Trim(sSmartItemDef)) < 1 Then
                        sSmartItemDef = oSmartItem.SymbolDefinition
                    End If
                    
                    oMemberObjects.GetItemDispid oSmartOccurrence, lDispId, lMbrIdx
                    Set oMemberDescription = oMemberDescriptions.ItemByDispid(lDispId)
                    sMemberItemName = oMemberDescription.Name
                    
                End If
            End If
        End If
        
        If Not oMemberItem Is Nothing Then
            ' get the cuurent Object's input objects
            sMemberObjectType = ""
            Set oMemberDatatypes = Nothing
            Set oMemberDataObjects = Nothing
            Migrate_GetMemberData oMemberItem, oMemberDataObjects, oMemberDatatypes
            
            If oMemberDataObjects Is Nothing Then
            ElseIf oMemberDataObjects.Count < 1 Then
            Else
                sMemberObjectType = oMemberDatatypes.Item(1)
            End If
            
            ' check if current Item is the given SmartOccurrence that creates the dependent MemberItems
            ' AND the m_ReplacedParts/m_ReplacingParts collections are nothing
            ' ... set m_ReplacedParts/m_ReplacingParts collections based on the Parent migration data
            ' ... expecting m_ReplacedParts/m_ReplacingParts collections to be empty on very first ParentMemberItem
            ' ... else m_ReplacedParts/m_ReplacingParts collections contain migration from Parent's Parent object
            If oMemberItem Is oParentMemberItem Then
                If m_ReplacedParts Is Nothing Then
                    Migrate_SetReplacingParts oMemberDataObjects, oMemberDatatypes, oMigrateHelper, m_ReplacedParts, m_ReplacingParts
                    ' if m_ReplacedParts/m_ReplacingParts collections are empty
                    ' ... there is nothing to Migrate
                    If m_ReplacedParts Is Nothing Then
                        Exit Sub
                    ElseIf m_ReplacedParts.Count < 1 Then
'' '''                        Exit Sub
                    End If
                End If
                
                If bMigrateParent Then
If bSM_trace Then zSM_trace "Migrate_MemberItems ... oParentMemberItem: " & Debug_ObjectName(oParentMemberItem, True)
    
                Else
                    sMemberObjectType = ""
                End If
            
            Else
                If m_ReplacedParts Is Nothing Then
                    Migrate_SetReplacingParts oMemberDataObjects, oMemberDatatypes, oMigrateHelper, m_ReplacedParts, m_ReplacingParts
                ElseIf m_ReplacedParts.Count < 1 Then
                    Migrate_SetReplacingParts oMemberDataObjects, oMemberDatatypes, oMigrateHelper, m_ReplacedParts, m_ReplacingParts
                End If
            End If
            
            ' For each of the expected Object Type under the IJStructAssemblyConnection
            ' ... Check/Create Replaced/Replacing data for the Object inputs as required
            ' ... Migrate the Object inputs and Place the Object on new Split Part as required
            ' ... this is dependent on the m_ReplacedParts and m_ReplacingParts collections
            If Len(sMemberObjectType) > 0 Then
                Migrate_CreateReplacingObjects oMemberItem, oMigrateHelper, False
            End If
            
            If Len(sMemberObjectType) < 1 Then
                ' Skip Migrating Parent
            ElseIf sMemberObjectType = "SF_WebCut" Then
                Migrate_WebCut oMemberItem, oMigrateHelper
            ElseIf sMemberObjectType = "SF_FlangeCut" Then
                Migrate_FlangeCut oMemberItem, oMigrateHelper
            ElseIf sMemberObjectType = "spType_BEARING" Then
                Migrate_Bearing oMemberObjects.Item(iIndex), oMigrateHelper

            ' Following objects are (usually) dependent on the above Ports
            ' need to check if the input Ports require Replaced/Replacing Split/Migration data
            ElseIf sMemberObjectType = "SF_CornerFeature" Then
                Migrate_CornerFeature oMemberItem, oMigrateHelper
            ElseIf sMemberObjectType = "SF_Slot" Then
                ' currently only expecting Slot for "InsertPlate" (cut)
                Migrate_InsertPlateSlot oMemberItem, oMigrateHelper
            ElseIf sMemberObjectType = "IJCollarPart" Then
            
            ElseIf sMemberObjectType = "IJChamfer" Then
                ' currently only expecting Chamfer for "InsertPlate" (Slot cut)
                Migrate_InsertPlateChamfer oMemberItem, oMigrateHelper
            
            ElseIf sMemberObjectType = "IJStructPhysicalConnection" Then
                Migrate_PC oMemberObjects.Item(iIndex), oMigrateHelper

            ' Following objects are (usually) not placed on Members
            ' ... currently not supporting them for Member Split/Migration
            ElseIf sMemberObjectType = "spType_BRACKET" Then
            ElseIf sMemberObjectType = "spType_COLLAR" Then
            ElseIf sMemberObjectType = "spType_INSERT" Then
            ElseIf sMemberObjectType = "spType_PAD" Then
            ElseIf sMemberObjectType = "spType_PARAMETRIC" Then

            ElseIf sMemberObjectType = "SF_EdgeFeature" Then
            ElseIf sMemberObjectType = "SF_FaceFeature" Then
            ElseIf sMemberObjectType = "SF_WaterStop" Then

            ' Following objects are (usually) Smart Occurrence Parent object
            ' ... used to create the dependent Features (Smart Occurrence MemberItems)
            ElseIf sMemberObjectType = "IJFreeEndCut" Then
                ' Not expecting FreeEndCuts to be "nested" as children
                Migrate_FreeEndCut oMemberItem, oMigrateHelper
            
            ElseIf sMemberObjectType = "IJAssemblyConnection" Then
            ElseIf sMemberObjectType = "IJStructAssemblyConnection" Then
                ' Check if this IJStructAssemblyConnection is the start
                ' or if it is a child (nested) of MemberItem Objects
                ' ... set m_ReplacedParts/m_ReplacingParts collections based on this (current) Object migration data
                If Not oMemberItem Is oParentMemberItem Then
                    Set oSaveReplacedParts = Nothing
                    Set oSaveReplacingParts = Nothing
                    If m_ReplacedParts Is Nothing Then
                    Else
                        ' save a copy of the m_ReplacedParts/m_ReplacingParts collections
                        ' ... to be restored After this object's MemberItem (children) have been processed
                        Set oSaveReplacedParts = New Collection
                        Set oSaveReplacingParts = New Collection
                        If m_ReplacedParts.Count > 0 Then
                            For jIndex = 1 To m_ReplacedParts.Count
                                oSaveReplacedParts.Add m_ReplacedParts.Item(jIndex)
                                oSaveReplacingParts.Add m_ReplacingParts.Item(jIndex)
                            Next jIndex
                        End If
                    End If
                    
                    Set m_ReplacedParts = Nothing
                    Set m_ReplacingParts = Nothing
                    Migrate_SetReplacingParts oMemberDataObjects, oMemberDatatypes, oMigrateHelper, m_ReplacedParts, m_ReplacingParts
                End If
                
                Migrate_StructAssemblyConnection oMemberItem, oMigrateHelper
            
            End If
        
            ' if current Item IS NOT the owning Parent
            ' check if to process this item list MemberItems
            If Not oMemberItem Is oParentMemberItem Then
                If bMigrateMemberItemChildren Then
                    ' if m_ReplacedParts/m_ReplacingParts collections are empty
                    ' ... there is nothing to Migrate
                    If m_ReplacedParts Is Nothing Then
                    ElseIf m_ReplacedParts.Count < 1 Then
                    Else
                        ' current Member Item has already been migrated above
                        ' ... do not need to MigrateParent when migrating Member Items children
If bSM_trace Then zSM_trace "Migrate_MemberItems ... oMemberItem(" & Trim(Str(iIndex)) & "): " & Debug_ObjectName(oMemberItem, True)
                        Migrate_MemberItems oMemberItem, oMigrateHelper, False, bMigrateMemberItems, bMigrateMemberItemChildren
                    End If
                End If
            End If
            
            ' restore m_ReplacedParts/m_ReplacingParts collections if saved data exists
            If oSaveReplacedParts Is Nothing Then
            Else
                Set m_ReplacedParts = New Collection
                Set m_ReplacingParts = New Collection
                If oSaveReplacedParts.Count > 0 Then
                    For jIndex = 1 To oSaveReplacedParts.Count
                        m_ReplacedParts.Add oSaveReplacedParts.Item(jIndex)
                        m_ReplacingParts.Add oSaveReplacingParts.Item(jIndex)
                    Next jIndex
                End If
            
                Set oSaveReplacedParts = Nothing
                Set oSaveReplacingParts = Nothing
            End If
            
        End If
    Next iIndex

   Exit Sub

ErrorHandler:
''MsgBox "*** ... ERROR ... " & METHOD
    HandleError MODULE, METHOD, sMsg
End Sub

' ******************************************************************************************
' ******************************************************************************************
' ******************************************************************************************

Public Sub Migrate_GetMemberData(oMigrateObject As Object, oMemberDataOjects As Collection, oMemberDatatypes As Collection)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    Dim jIndex As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    Dim oPort As IJPort
    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection
    
    If oMemberDataOjects Is Nothing Then
        Set oMemberDataOjects = New Collection
    End If
    
    If oMemberDatatypes Is Nothing Then
        Set oMemberDatatypes = New Collection
    End If
    
    ' For the given Object, get the object dtype and the Inputs used
    ' ... Connections, SmartPlate, StructFeatures, Collars, etc.
    If oMigrateObject Is Nothing Then
    
    ElseIf TypeOf oMigrateObject Is IJStructAssemblyConnection Then
        Dim oPort1 As Object
        Dim oPort2 As Object
        Dim oAppConn As IJAppConnection
        Dim oElements As IJElements
        Dim oCollectionAlias As JCmnShp_CollectionAlias
        Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
        
        sMemberObjectType = "IJStructAssemblyConnection"
        Set oAppConn = oMigrateObject
        oAppConn.enumPorts oElements
        Set oPort1 = oElements.Item(1)
        Set oPort2 = oElements.Item(2)
        
        Set oConnAttrbs = New SDConnectionUtils
        Set oCollectionAlias = oConnAttrbs.get_AuxiliaryPorts(oAppConn)
        
        oMemberDataOjects.Add oMigrateObject
        oMemberDataOjects.Add oPort1
        oMemberDataOjects.Add oPort2
        
        oMemberDatatypes.Add sMemberObjectType
        oMemberDatatypes.Add "Port1"
        oMemberDatatypes.Add "Port2"
        
        If oCollectionAlias Is Nothing Then
        ElseIf oCollectionAlias.Count > 0 Then
            For iIndex = 1 To oCollectionAlias.Count
                If oCollectionAlias.Item(iIndex) Is Nothing Then
                Else
                    oMemberDataOjects.Add oCollectionAlias.Item(iIndex)
                    oMemberDatatypes.Add "Auxiliary"
                End If
            Next iIndex
        End If
    
        Dim oEditJDArgument As IJDEditJDArgument
        Dim oReferencesCollection As IJDReferencesCollection
    
        If TypeOf oMigrateObject Is IJSmartOccurrence Then
            Set oReferencesCollection = GetRefCollFromSmartOccurrence(oMigrateObject)
            If oReferencesCollection Is Nothing Then
            Else
                Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
                If oEditJDArgument.GetCount > 0 Then
                    For iIndex = 1 To oEditJDArgument.GetCount
                        oMemberDataOjects.Add oEditJDArgument.GetEntityByIndex(iIndex)
                        oMemberDatatypes.Add "ReferencesCollection"
                    Next iIndex
                End If
            End If
        End If
    
    ElseIf TypeOf oMigrateObject Is IJAssemblyConnection Then
        
        sMemberObjectType = "IJAssemblyConnection"
        Set oAppConn = oMigrateObject
        oAppConn.enumPorts oElements
        Set oPort1 = oElements.Item(1)
        Set oPort2 = oElements.Item(2)
        
        Set oConnAttrbs = New SDConnectionUtils
        Set oCollectionAlias = oConnAttrbs.get_AuxiliaryPorts(oAppConn)
        
        oMemberDataOjects.Add oMigrateObject
        oMemberDataOjects.Add oPort1
        oMemberDataOjects.Add oPort2
        
        oMemberDatatypes.Add sMemberObjectType
        oMemberDatatypes.Add "Port1"
        oMemberDatatypes.Add "Port2"
        
        If oCollectionAlias Is Nothing Then
        ElseIf oCollectionAlias.Count > 0 Then
            For iIndex = 1 To oCollectionAlias.Count
                If oCollectionAlias.Item(iIndex) Is Nothing Then
                Else
                    oMemberDataOjects.Add oCollectionAlias.Item(iIndex)
                    oMemberDatatypes.Add "Auxiliary"
                End If
            Next iIndex
        End If

        If TypeOf oMigrateObject Is IJSmartOccurrence Then
            Set oReferencesCollection = GetRefCollFromSmartOccurrence(oMigrateObject)
            If oReferencesCollection Is Nothing Then
            Else
                Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
                If oEditJDArgument.GetCount > 0 Then
                    For iIndex = 1 To oEditJDArgument.GetCount
                        oMemberDataOjects.Add oEditJDArgument.GetEntityByIndex(iIndex)
                        oMemberDatatypes.Add "ReferencesCollection"
                    Next iIndex
                End If
            End If
        End If

    ElseIf TypeOf oMigrateObject Is IJSmartPlate Then
        Dim oRefPlane As Object
        Dim oSmartPlate As IJSmartPlate
        Dim eSmartPlateType As SmartPlateTypes
        Dim oSDSmartPlateAtt As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
        
        Set oSmartPlate = oMigrateObject
        eSmartPlateType = oSmartPlate.SmartPlateType
        Set oSDSmartPlateAtt = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
        
        If eSmartPlateType = spType_BEARING Then
            sMemberObjectType = "spType_BEARING"
            oSDSmartPlateAtt.GetInputs_BearingPlate oMigrateObject, oCollectionAlias
            
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
            
            If oCollectionAlias Is Nothing Then
            ElseIf oCollectionAlias.Count > 0 Then
                For iIndex = 1 To oCollectionAlias.Count
                    If oCollectionAlias.Item(iIndex) Is Nothing Then
                    Else
                        oMemberDataOjects.Add oCollectionAlias.Item(iIndex)
                        oMemberDatatypes.Add "Support"
                    End If
                Next iIndex
            End If
            
        ElseIf eSmartPlateType = spType_BRACKET Then
            sMemberObjectType = "spType_BRACKET"
            oSDSmartPlateAtt.GetInputs_Bracket oMigrateObject, oRefPlane, oElements
            
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
            
            If oElements Is Nothing Then
            ElseIf oElements.Count > 0 Then
                For iIndex = 1 To oElements.Count
                    If oElements.Item(iIndex) Is Nothing Then
                    Else
                        oMemberDataOjects.Add oElements.Item(iIndex)
                        oMemberDatatypes.Add "Support"
                    End If
                Next iIndex
            End If
            
        ElseIf eSmartPlateType = spType_COLLAR Then
            sMemberObjectType = "spType_COLLAR"
            oSDSmartPlateAtt.GetInputs_SmartCollar oMigrateObject, oCollectionAlias
            
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
            
            If oCollectionAlias Is Nothing Then
            ElseIf oCollectionAlias.Count > 0 Then
                For iIndex = 1 To oCollectionAlias.Count
                    If oCollectionAlias.Item(iIndex) Is Nothing Then
                    Else
                        oMemberDataOjects.Add oCollectionAlias.Item(iIndex)
                        oMemberDatatypes.Add "Input"
                    End If
                Next iIndex
            End If
            
        ElseIf eSmartPlateType = spType_INSERT Then
            sMemberObjectType = "spType_INSERT"
            oSDSmartPlateAtt.GetInputs_InsertPlate oMigrateObject, oCollectionAlias
            
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
            
            If oCollectionAlias Is Nothing Then
            ElseIf oCollectionAlias.Count > 0 Then
                For iIndex = 1 To oCollectionAlias.Count
                    If oCollectionAlias.Item(iIndex) Is Nothing Then
                    Else
                        oMemberDataOjects.Add oCollectionAlias.Item(iIndex)
                        oMemberDatatypes.Add "Input"
                    End If
                Next iIndex
            End If
            
        ElseIf eSmartPlateType = spType_PAD Then
            sMemberObjectType = "spType_PAD"
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
        
        ElseIf eSmartPlateType = spType_PARAMETRIC Then
            sMemberObjectType = "spType_PARAMETRIC"
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
        
        ElseIf eSmartPlateType = spType_UNTYPED Then
            sMemberObjectType = "spType_UNTYPED"
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
        
        Else
            sMemberObjectType = "sp_?" & Trim(Str(eSmartPlateType))
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
        End If

    ElseIf TypeOf oMigrateObject Is IJStructFeature Then
        Dim oWebCut As Object
        Dim oBounded As Object
        Dim oBounding As Object
        Dim oFacePort As Object
        Dim oLocation As Object
        Dim oBasePlate As Object
        Dim oPenetrated As Object
        Dim oPenetrating As Object
    
        Dim oStructFeature As IJStructFeature
        Dim eStructFeatureType As StructFeatureTypes
        
        Dim oSDFeatureAttributes As IJSDFeatureAttributes
        Set oSDFeatureAttributes = New SDFeatureUtils
        
        Set oStructFeature = oMigrateObject
        eStructFeatureType = oStructFeature.get_StructFeatureType
        
        If eStructFeatureType = SF_FlangeCut Then
            sMemberObjectType = "SF_FlangeCut"
            oSDFeatureAttributes.get_FlangeCutInputs oMigrateObject, oBounding, oBounded, oWebCut
        
            oMemberDataOjects.Add oMigrateObject
            oMemberDataOjects.Add oBounding
            oMemberDataOjects.Add oBounded
            oMemberDataOjects.Add oWebCut
        
            oMemberDatatypes.Add sMemberObjectType
            oMemberDatatypes.Add "Bounding"
            oMemberDatatypes.Add "Bounded"
            oMemberDatatypes.Add "WebCut"
        
        ElseIf eStructFeatureType = SF_WebCut Then
            sMemberObjectType = "SF_WebCut"
            oSDFeatureAttributes.get_WebCutInputs oMigrateObject, oBounding, oBounded
        
            oMemberDataOjects.Add oMigrateObject
            oMemberDataOjects.Add oBounding
            oMemberDataOjects.Add oBounded
        
            oMemberDatatypes.Add sMemberObjectType
            oMemberDatatypes.Add "Bounding"
            oMemberDatatypes.Add "Bounded"
        
        ElseIf eStructFeatureType = SF_CornerFeature Then
            sMemberObjectType = "SF_CornerFeature"
            ''oSDFeatureAttributes.get_CornerCutInputs oMigrateObject, oFacePort, oPort1, oPort2
            ''oSDFeatureAttributes.get_CornerCutInputsEx oMigrateObject, oFacePort, oPort1, oPort2
            oSDFeatureAttributes.get_CornerCutInputsEx2 oMigrateObject, oFacePort, oPort1, oPort2
        
            oMemberDataOjects.Add oMigrateObject
            oMemberDataOjects.Add oFacePort
            oMemberDataOjects.Add oPort1
            oMemberDataOjects.Add oPort2
            
            oMemberDatatypes.Add sMemberObjectType
            oMemberDatatypes.Add "FacePort"
            oMemberDatatypes.Add "EdgePort1"
            oMemberDatatypes.Add "EdgePort2"
        
        ElseIf eStructFeatureType = SF_EdgeFeature Then
            sMemberObjectType = "SF_EdgeFeature"
            oSDFeatureAttributes.Get_Inputs_EdgeCut oMigrateObject, oPort1, oLocation
            
            oMemberDataOjects.Add oMigrateObject
            oMemberDataOjects.Add oPort1
            oMemberDataOjects.Add oLocation
        
            oMemberDatatypes.Add sMemberObjectType
            oMemberDatatypes.Add "EgePort"
            oMemberDatatypes.Add "Location"
        
        ElseIf eStructFeatureType = SF_FaceFeature Then
            sMemberObjectType = "SF_FaceFeature"
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
        
        ElseIf eStructFeatureType = SF_Slot Then
            Dim oPenetratedObj As Object
            Dim oPenetratedUnk As IUnknown
            Dim oPenetratedGeometry As IUnknown
            Dim oStructGeometryHelper As StructGeometryHelper
            
            sMemberObjectType = "SF_Slot"
            oSDFeatureAttributes.get_SlotInputs oMigrateObject, oPenetrating, oPenetrated, oBasePlate
            
            oMemberDataOjects.Add oMigrateObject
            oMemberDataOjects.Add oPenetrating
            
            oMemberDatatypes.Add sMemberObjectType
            oMemberDatatypes.Add "Penetrating"
        
            ' The penetrated object input is the geometry, change it to the part
            Set oPenetratedGeometry = oPenetrated
            Set oStructGeometryHelper = New StructGeometryHelper
            oStructGeometryHelper.RecursiveGetStructEntityIUnknown oPenetratedGeometry, oPenetratedUnk
            Set oPenetratedObj = oPenetratedUnk
            
            oMemberDataOjects.Add oPenetrated
            oMemberDataOjects.Add oPenetratedObj
            
            oMemberDatatypes.Add " Penetrated"
            oMemberDatatypes.Add "*Penetrated"
            
            oMemberDataOjects.Add oBasePlate
            oMemberDatatypes.Add "BasePlate"
            
        ElseIf eStructFeatureType = SF_WaterStop Then
            sMemberObjectType = "SF_WaterStop"
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
        
        Else
            sMemberObjectType = "SF_?" & Trim(Str(eStructFeatureType))
            oMemberDataOjects.Add oMigrateObject
            oMemberDatatypes.Add sMemberObjectType
        End If
    
    ElseIf TypeOf oMigrateObject Is IJCollarPart Then
        Dim oSlot As Object
        Dim strRootClassName As String
        Dim oCollarAttributes As IJSDCollarAttributes
        Set oCollarAttributes = New SDCollarUtils
        
        sMemberObjectType = "IJCollarPart"
        oCollarAttributes.GetInput_Collar oMigrateObject, oBasePlate, oPenetrating, strRootClassName, oSlot
    
        oMemberDataOjects.Add oMigrateObject
        oMemberDataOjects.Add oBasePlate
        oMemberDataOjects.Add oPenetrating
        oMemberDataOjects.Add oSlot
    
        oMemberDatatypes.Add sMemberObjectType
        oMemberDatatypes.Add "BasePlate"
        oMemberDatatypes.Add "Penetrating"
        oMemberDatatypes.Add "Slot"
    
    ElseIf TypeOf oMigrateObject Is IJStructPhysicalConnection Then
        Dim oSDConnectionAttributes As IJSDConnectionAttributes
        Set oSDConnectionAttributes = New SDConnectionUtils
        
        sMemberObjectType = "IJStructPhysicalConnection"
        oSDConnectionAttributes.get_ConnectionInputs oMigrateObject, oPort1, oPort2
    
        oMemberDataOjects.Add oMigrateObject
        oMemberDataOjects.Add oPort1
        oMemberDataOjects.Add oPort2
        
        oMemberDatatypes.Add sMemberObjectType
        oMemberDatatypes.Add "Port1"
        oMemberDatatypes.Add "Port2"
    
    ElseIf TypeOf oMigrateObject Is IJFreeEndCut Then
        Dim oFreeEndcut As IJFreeEndCut
        
        sMemberObjectType = "IJFreeEndCut"
        Set oFreeEndcut = oMigrateObject
        
        oFreeEndcut.get_FreeEndCutInputs oPort1, oPort2
    
        oMemberDataOjects.Add oMigrateObject
        oMemberDataOjects.Add oPort1
        oMemberDataOjects.Add oPort2
        
        oMemberDatatypes.Add sMemberObjectType
        oMemberDatatypes.Add "Bounding"
        oMemberDatatypes.Add "Bounded"
    
    ElseIf TypeOf oMigrateObject Is IJChamfer Then
        Dim oChamfer As IJChamfer
        Dim oChamfered As Object
        Dim oDrivesChamfer As Object
        Dim oSDChamferAttributes As IJSDChamferAttributes
        
        sMemberObjectType = "IJChamfer"
        Set oChamfer = oMigrateObject
        Set oSDChamferAttributes = New SDChamferUtils
        oSDChamferAttributes.get_ChamferInputs oChamfer, oPort1, oPort2
    
        oMemberDataOjects.Add oMigrateObject
        oMemberDataOjects.Add oPort1
        oMemberDataOjects.Add oPort2
        
        oMemberDatatypes.Add sMemberObjectType
        oMemberDatatypes.Add "Input1"
        oMemberDatatypes.Add "Input2"
    
        If TypeOf oPort1 Is IJStructFeature Then
            Set oStructFeature = oPort1
            eStructFeatureType = oStructFeature.get_StructFeatureType
            Set oSDFeatureAttributes = New SDFeatureUtils

            If eStructFeatureType = SF_WebCut Then
                oSDFeatureAttributes.get_WebCutInputs oPort1, oBounding, oBounded
                Set oChamfered = oBounded
                
                If oBounding Is Nothing Then
                    Set oDrivesChamfer = oBounded
                ElseIf TypeOf oBounding Is IJPort Then
                    Set oDrivesChamfer = oBounding
                Else
                    Set oDrivesChamfer = oBounded
                End If
           
           ElseIf eStructFeatureType = SF_FlangeCut Then
                oSDFeatureAttributes.get_FlangeCutInputs oPort1, oBounding, oBounded, oWebCut
                If oBounding Is Nothing Then
                    Set oDrivesChamfer = oBounded
                ElseIf TypeOf oBounding Is IJPort Then
                    Set oDrivesChamfer = oBounding
                Else
                    Set oDrivesChamfer = oBounded
                End If
           End If
           
            oMemberDataOjects.Add oChamfered
            oMemberDatatypes.Add "Chamfered"
            If Not oChamfered Is oChamfered Then
                oMemberDataOjects.Add oDrivesChamfer
                oMemberDatatypes.Add "DrivesChamfer"
            End If
           
        End If
    
    Else
        sMemberObjectType = "_?" & TypeName(oMigrateObject)
        oMemberDataOjects.Add oMigrateObject
        oMemberDatatypes.Add "MigrateObject"
    End If
        
    Dim bFreeEndCut As Boolean
    Dim oParentObj As Object
    Dim oSmartParent As Object
    Dim oDesignChild As IJDesignChild
    Dim oSystemChild As IJSystemChild
    If TypeOf oMigrateObject Is IJSmartOccurrence Then
        GetSmartOccurrenceParent oMigrateObject, oSmartParent
        If oSmartParent Is Nothing Then
            If TypeOf oMigrateObject Is IJDesignChild Then
                Set oDesignChild = oMigrateObject
                Set oParentObj = oDesignChild.GetParent
                oMemberDataOjects.Add oParentObj
                oMemberDatatypes.Add "DesignParent"
            ElseIf TypeOf oMigrateObject Is IJSystemChild Then
                Set oSystemChild = oMigrateObject
                Set oParentObj = oSystemChild.GetParent
                oMemberDataOjects.Add oParentObj
                oMemberDatatypes.Add "SystemParent"
            End If
        Else
            oMemberDataOjects.Add oSmartParent
            oMemberDatatypes.Add "SmartParent"
        End If
    
    ElseIf TypeOf oMigrateObject Is IJDesignChild Then
        Set oDesignChild = oMigrateObject
        Set oParentObj = oDesignChild.GetParent
        oMemberDataOjects.Add oParentObj
        oMemberDatatypes.Add "DesignParent"
    ElseIf TypeOf oMigrateObject Is IJSystemChild Then
        Set oSystemChild = oMigrateObject
        Set oParentObj = oSystemChild.GetParent
        oMemberDataOjects.Add oParentObj
        oMemberDatatypes.Add "SystemParent"
    End If
    
    Exit Sub
    
ErrorHandler:
If bSM_trace Then zSM_trace "*** ... ERROR ... Migrate_GetMemberData"
''MsgBox "*** ... ERROR ... Migrate_GetMemberData"
    Err.Raise LogError(Err, MODULE, "Migrate_GetMemberData", strError).Number
End Sub

Public Sub Migrate_SetReplacingParts(oMemberDataOjects As Collection, oMemberDatatypes As Collection, oMigrateHelper As IJMigrateHelper, _
                                     oReplacedParts As Collection, oReplacingParts As Collection)
Const METHOD = "::Migrate_SetReplacingParts"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim jIndex As Long
    Dim dDist As Double
    Dim dMinDist As Double
    Dim bIsDeleted As Boolean

    Dim oChkRange As IJRangeAlias
    Dim oRefRange As IJRangeAlias
    Dim gChkRangeBox As IMSEntitySupport.GBox
    Dim gRefRangeBox As IMSEntitySupport.GBox
    Dim oChkPosition As IJDPosition
    Dim oRefPosition As IJDPosition
    
    Dim oPort As IJPort
    Dim oReplacedPart As Object
    Dim oReplacingPort As Object
    Dim oReplacingPart As Object
    Dim oObjectReplacing As Object
    Dim oChkReplacingPart As Object
    
    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection
    
    Set oReplacedParts = New Collection
    Set oReplacingParts = New Collection
    
    If oMemberDataOjects Is Nothing Then
        Exit Sub
    ElseIf oMemberDataOjects.Count < 1 Then
        Exit Sub
    ElseIf oMemberDatatypes Is Nothing Then
        Exit Sub
    ElseIf oMemberDatatypes.Count < 1 Then
        Exit Sub
    End If
    
    ' Currently using range box to determine replaced and replacing parts
    ' ... Need something that is more reliable ... *** To Do ***
    sMsg = "setup Replaced Part data "
    
    If oMemberDataOjects.Item(1) Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oMemberDataOjects.Item(1) Is IJRangeAlias Then
        Exit Sub
    End If
    
    Set oChkPosition = New AutoMath.DPosition
    Set oRefPosition = New AutoMath.DPosition
    
    ' Get RangeBox mid-point of object being Migratioed
    ' ... IJStructureFeature, IJSmartPlate, IJChamfer, IJCollarPart, Assembly/Physical Connection
    Set oRefRange = oMemberDataOjects.Item(1)
    gRefRangeBox = oRefRange.GetRange()
    oRefPosition.Set (gRefRangeBox.m_high.x + gRefRangeBox.m_low.x) / 2#, _
                     (gRefRangeBox.m_high.y + gRefRangeBox.m_low.y) / 2#, _
                     (gRefRangeBox.m_high.z + gRefRangeBox.m_low.z) / 2#

    For iIndex = 2 To oMemberDataOjects.Count
        ' Get IJConnectable of Migrated object's Inputs
        Set oReplacedPart = Nothing
        If oMemberDataOjects.Item(iIndex) Is Nothing Then
        ElseIf TypeOf oMemberDataOjects.Item(iIndex) Is IJPort Then
            Set oPort = oMemberDataOjects.Item(iIndex)
            Set oReplacedPart = oPort.Connectable
        ElseIf TypeOf oMemberDataOjects.Item(iIndex) Is IJConnectable Then
            Set oReplacedPart = oMemberDataOjects.Item(iIndex)
        End If
        
        ' for the Replaced Input Port/Part, Determine Replacing Part
        If Not oReplacedPart Is Nothing Then
            ' Get Migrated object's Input's Replacing objects
            Set oReplacingPart = oReplacedPart
            oMigrateHelper.ObjectsReplacing oMemberDataOjects.Item(iIndex), oObjectCollectionReplacing, bIsDeleted
            If oObjectCollectionReplacing Is Nothing Then
                ' Input is NOT replaced, get Inputs's IJConnectable Replacing objects
                oMigrateHelper.ObjectsReplacing oReplacedPart, oObjectCollectionReplacing, bIsDeleted
            ElseIf oObjectCollectionReplacing.Count < 1 Then
                ' Input is NOT replaced, get Inputs's IJConnectable Replacing objects
                oMigrateHelper.ObjectsReplacing oReplacedPart, oObjectCollectionReplacing, bIsDeleted
            End If
            
            If oObjectCollectionReplacing Is Nothing Then
            ElseIf oObjectCollectionReplacing.Count < 1 Then
            Else
                dMinDist = -1#
                For Each oObjectReplacing In oObjectCollectionReplacing
                    ' Get RangeBox mid-point of Replacing object (Part.geometry, Port.geometry, object.geometry)
                    Set oChkPosition = Nothing
                    Set oChkReplacingPart = Nothing
                    If TypeOf oObjectReplacing Is IJPort Then
                        Set oPort = oObjectReplacing
                        Set oChkReplacingPart = oPort.Connectable
                    ElseIf TypeOf oObjectReplacing Is IJConnectable Then
                        Set oChkReplacingPart = oObjectReplacing
                    End If
                    
                    If oChkReplacingPart Is Nothing Then
                    Else
                        GetPointOnObject oObjectReplacing, oRefPosition, oChkPosition

''                    ElseIf TypeOf oObjectReplacing Is ISPSSplitAxisAlongPort Then
''                    ElseIf TypeOf oObjectReplacing Is ISPSSplitAxisEndPort Then
''                        Dim dX As Double
''                        Dim dY As Double
''                        Dim dZ As Double
''                        Dim oPoint As IJPoint
''                        Dim oSplitAxisPort As ISPSSplitAxisPort
''                        Dim oMemberPartCommon As ISPSMemberPartCommon
''                        Set oSplitAxisPort = oObjectReplacing
''                        Set oMemberPartCommon = oSplitAxisPort.Part
''                        Set oPoint = oMemberPartCommon.PointAtEnd(oSplitAxisPort.PortIndex)
''                        oPoint.GetPoint dX, dY, dZ
''                        oChkPosition.Set dX, dY, dZ
''
''                    ElseIf TypeOf oObjectReplacing Is IJRangeAlias Then
''                        Set oChkRange = oObjectReplacing
''                        gChkRangeBox = oChkRange.GetRange()
''                        oChkPosition.Set (gChkRangeBox.m_high.x + gChkRangeBox.m_low.x) / 2#, _
''                                         (gChkRangeBox.m_high.y + gChkRangeBox.m_low.y) / 2#, _
''                                         (gChkRangeBox.m_high.z + gChkRangeBox.m_low.z) / 2#
                    End If
                        
                    If Not oChkPosition Is Nothing Then
                        dDist = oRefPosition.DistPt(oChkPosition)
                        If dMinDist < 0# Then
                            dMinDist = dDist
                            Set oReplacingPart = oChkReplacingPart
                        ElseIf dDist < dMinDist Then
                            dMinDist = dDist
                            Set oReplacingPart = oChkReplacingPart
                        End If
                    End If
                Next
                
                ' add the Replaced/Replacing Parts to collection
                If Not oReplacedPart Is oReplacingPart Then
                    bIsDeleted = False
                    For jIndex = 1 To oReplacedParts.Count
                        If oReplacedPart Is oReplacedParts.Item(jIndex) Then
                            bIsDeleted = True
                            Exit For
                        End If
                    Next jIndex
                    
                    If Not bIsDeleted Then
                        oReplacedParts.Add oReplacedPart
                        oReplacingParts.Add oReplacingPart
                    End If
                End If
            End If
        End If
    Next iIndex
    
If bSM_trace Then
    zSM_trace "... " & METHOD & " ... oReplacedParts.Count : " & oReplacedParts.Count & " ... " & oMemberDatatypes.Item(1)
    For iIndex = 1 To oReplacedParts.Count
        zSM_trace "... (" & Trim(Str(iIndex)) & ") ... Replaced  : " & Debug_ObjectName(oReplacedParts.Item(iIndex), True)
        zSM_trace "... (" & Trim(Str(iIndex)) & ") ... Replacing : " & Debug_ObjectName(oReplacingParts.Item(iIndex), True)
    Next iIndex
End If

   Exit Sub

ErrorHandler:
''MsgBox "*** ... ERROR ... " & METHOD
    HandleError MODULE, METHOD, sMsg
End Sub

Public Sub Migrate_GetReplacingObject(oToBeReplaced As Object, oMigrateHelper As IJMigrateHelper, oReplacingObj As Object, oReplacingPart As Object)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    
    Dim strError As String
    
    Dim bIsDeleted As Boolean
    
    Dim oPort As IJPort
    Dim oObjectReplacing As Object
    
    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection

    Set oReplacingObj = oToBeReplaced
    Set oReplacingPart = Nothing
    If oToBeReplaced Is Nothing Then
        Exit Sub
    End If
    
    oMigrateHelper.ObjectsReplacing oToBeReplaced, oObjectCollectionReplacing, bIsDeleted
    If oObjectCollectionReplacing Is Nothing Then
    ElseIf oObjectCollectionReplacing.Count < 1 Then
    Else
        For Each oObjectReplacing In oObjectCollectionReplacing
            If oObjectCollectionReplacing.Count = 1 Then
                Set oReplacingObj = oObjectReplacing
                Exit For
            ElseIf TypeOf oObjectReplacing Is IJPort Then
                Set oPort = oObjectReplacing
                For iIndex = 1 To m_ReplacingParts.Count
                    If oPort.Connectable Is m_ReplacingParts.Item(iIndex) Then
                        Set oReplacingObj = oObjectReplacing
                        Exit For
                    End If
                Next iIndex
            ElseIf TypeOf oObjectReplacing Is IJConnectable Then
                For iIndex = 1 To m_ReplacingParts.Count
                    If oObjectReplacing Is m_ReplacingParts.Item(iIndex) Then
                        Set oReplacingObj = oObjectReplacing
                        Exit For
                    End If
                Next iIndex
            End If
        Next
    End If
    
    If TypeOf oReplacingObj Is IJPort Then
        Set oPort = oReplacingObj
        Set oReplacingPart = oPort.Connectable
    ElseIf TypeOf oReplacingObj Is IJConnectable Then
        Set oReplacingPart = oReplacingObj
    End If
                
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_GetReplacingObject"
    Err.Raise LogError(Err, MODULE, "Migrate_GetReplacingObject", strError).Number
End Sub

Public Sub Migrate_GetFeatureFromOptOpr(oMemberPart As Object, lOptId As Long, lOprId As Long, oFeature As Object, oOperation As Object)
    On Error GoTo ErrorHandler
    Dim strError As String
    Dim iIndex As Long
    Dim lFeatureCtx As Long
    Dim lFeatureOptId As Long
    Dim lFeatureOprId As Long
    
    Dim eStructOperation As cmnstrStructOperation
    Dim oOperatorsList As IJElements
    Dim oOperationObject As Object
    Dim oStructOperationAE As IJStructOperationAE
    Dim oStructGraphNavigate As IJStructGraphNavigate
    
    Set oFeature = Nothing
    Set oOperation = Nothing
    
    If TypeOf oMemberPart Is IJStructGraphNavigate Then
    
        If m_MigratedFeaturesCount > 0 Then
            For iIndex = 1 To m_MigratedFeaturesCount
                If m_MigratedFeatures(iIndex).ReplacedMember Is oMemberPart Then
                    If m_MigratedFeatures(iIndex).ReplacedOpt = lOptId Then
                        If m_MigratedFeatures(iIndex).ReplacedOpr = lOprId Then
                            Set oFeature = m_MigratedFeatures(iIndex).Feature
                            Set oOperation = m_MigratedFeatures(iIndex).ReplacedOperation
                            Exit Sub
                        End If
                    End If
                ElseIf m_MigratedFeatures(iIndex).ReplacingMember Is oMemberPart Then
                    If m_MigratedFeatures(iIndex).ReplacingOpt = lOptId Then
                        If m_MigratedFeatures(iIndex).ReplacingOpr = lOprId Then
                            Set oFeature = m_MigratedFeatures(iIndex).Feature
                            Set oOperation = m_MigratedFeatures(iIndex).ReplacingOperation
                            Exit Sub
                        End If
                    End If
                End If
            Next iIndex
        End If
    
        Set oStructGraphNavigate = oMemberPart
        oStructGraphNavigate.FindOperationInGraph lOptId, oOperationObject
        If oOperationObject Is Nothing Then
        Else
            Set oOperation = oOperationObject
            Set oStructOperationAE = oOperationObject
            oStructOperationAE.GetOperationType eStructOperation
            oStructOperationAE.GetOperatorById lOprId, oOperatorsList
            If oOperatorsList Is Nothing Then
            ElseIf oOperatorsList.Count < 1 Then
            Else
                Set oFeature = oOperatorsList.Item(1)
                Exit Sub
            End If
        End If
    End If
   
     ''===================================
     ''an alternate way to get the Feature
     If oFeature Is Nothing Then
        'Get Endcuts on pattern
        ' ... SPSMembers.SPSPartPrismaticGenerator.1
        'Get Cutouts on pattern
        ' ... SP3DStructGeneric.StructCutoutOperation.1 ... StructGeneric.StructCutoutOperationAE.1
        Dim oOperator As Object
        Dim OperationPattern As IJStructOperationPattern
        Dim oCollectionOfOperators As IJElements
        If TypeOf oMemberPart Is IJStructOperationPattern Then
            Set OperationPattern = oMemberPart
            If eStructOperation = cmnstrCutoutOperation Then
                OperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oOperationObject
            Else
                OperationPattern.GetOperationPattern "SPSMembers.SPSPartPrismaticGenerator.1", oCollectionOfOperators, oOperationObject
            End If
    
            ' search collection for given Opt,Opr pair
            ' NOTE: currently only handles WebCut, FlangeCut, and Slot Features
            ' ... does not handle Corner Features
            If oCollectionOfOperators Is Nothing Then
            ElseIf oCollectionOfOperators.Count > 0 Then
                For Each oOperator In oCollectionOfOperators
                    If TypeOf oOperator Is IJStructFeature Then
                        If eStructOperation = cmnstrCutoutOperation Then
                            GetLateBindMbrFeatureData oOperator, oMemberPart, lFeatureCtx, lFeatureOptId, lFeatureOprId
                        Else
                            GetLateBindMbrFeatureData oOperator, oMemberPart, lFeatureCtx, lFeatureOptId, lFeatureOprId
                        End If
                        If lFeatureOptId = lOptId Then
                            If lFeatureOprId = lOprId Then
                                Set oFeature = oOperator
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
If bSM_trace Then zSM_trace "*** ... ERROR ... Migrate_GetFeatureFromOptOpr"
''MsgBox "*** ... ERROR ... Migrate_GetFeatureFromOptOpr"
    Err.Raise LogError(Err, MODULE, "Migrate_GetFeatureFromOptOpr", strError).Number
End Sub

Public Sub Migrate_GetOptOprFromFeature(oFeature As Object, oMemberPart As Object, lFeatureOptId As Long, lFeatureOprId As Long, oOperation As Object)
    On Error GoTo ErrorHandler
    Dim strError As String
    Dim iIndex As Long
    Dim lFeatureCtx As Long
    
    Dim eStructOperation As cmnstrStructOperation
    Dim oOperator As Object
    Dim oOperationObject As Object
    Dim oStructOperationAE As IJStructOperationAE
    Dim OperationPattern As IJStructOperationPattern
    Dim oCollectionOfOperators As IJElements
    
    lFeatureOptId = 0
    lFeatureOprId = 0
    Set oOperation = Nothing
    
    'Get Endcuts on pattern
    ' ... SPSMembers.SPSPartPrismaticGenerator.1
    'Get Cutouts on pattern
    ' ... SP3DStructGeneric.StructCutoutOperation.1 ... StructGeneric.StructCutoutOperationAE.1
    If oMemberPart Is Nothing Then
       Exit Sub
    ElseIf Not TypeOf oMemberPart Is IJStructOperationPattern Then
       Exit Sub
    ElseIf oFeature Is Nothing Then
       Exit Sub
    End If
    
    For iIndex = 1 To 2
        Set OperationPattern = oMemberPart
        If iIndex = 1 Then
            OperationPattern.GetOperationPattern "SPSMembers.SPSPartPrismaticGenerator.1", oCollectionOfOperators, oOperationObject
        Else
            OperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oOperationObject
        End If
    
        If oOperationObject Is Nothing Then
            eStructOperation = cmnstrNoOperation
        ElseIf TypeOf oOperationObject Is IJStructOperationAE Then
            Set oStructOperationAE = oOperationObject
            oStructOperationAE.GetOperationType eStructOperation
        Else
            eStructOperation = cmnstrNoOperation
        End If
    
        ' search collection for given Feature
        If oCollectionOfOperators Is Nothing Then
        ElseIf oCollectionOfOperators.Count > 0 Then
            For Each oOperator In oCollectionOfOperators
                If oOperator Is Nothing Then
                ElseIf oOperator Is oFeature Then
                    Set oOperation = oOperationObject
                    If eStructOperation = cmnstrCutoutOperation Then
                        GetLateBindMbrFeatureData oOperator, oMemberPart, lFeatureCtx, lFeatureOptId, lFeatureOprId
                    Else
                        GetLateBindMbrFeatureData oOperator, oMemberPart, lFeatureCtx, lFeatureOptId, lFeatureOprId
                    End If
                    Exit Sub
                End If
            Next
        End If
    Next iIndex
    
    Exit Sub
    
ErrorHandler:
If bSM_trace Then zSM_trace "*** ... ERROR ... Migrate_GetOptOprFromFeature"
''MsgBox "*** ... ERROR ... Migrate_GetFeatureFromOptOpr"
    Err.Raise LogError(Err, MODULE, "Migrate_GetOptOprFromFeature", strError).Number
End Sub

Public Function Migrate_IsPortGeometryValid(oPortToCheck As Object) As Boolean
    On Error Resume Next
    
    Dim oPort As IJPort
    Migrate_IsPortGeometryValid = False
    If TypeOf oPortToCheck Is IJPort Then
        Set oPort = oPortToCheck
        If oPort.Geometry Is Nothing Then
        Else
            Migrate_IsPortGeometryValid = True
        End If
    End If
    
    Exit Function
    
End Function

Public Sub Migrate_SetReplacedReplacing(oToBeReplaced As Object, oWithReplacing As Object, oMigrateHelper As IJMigrateHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection

    If oToBeReplaced Is Nothing Then
        Exit Sub
    ElseIf oWithReplacing Is Nothing Then
        Exit Sub
    End If
    
    ' Flag the InsertPlateChamfer as Migrated
    Set oObjectCollectionReplaced = New JObjectCollection
    Set oObjectCollectionReplacing = New JObjectCollection

    oObjectCollectionReplaced.Add oToBeReplaced
    oObjectCollectionReplacing.Add oWithReplacing
        
    oMigrateHelper.ObjectsReplaced oObjectCollectionReplaced, oObjectCollectionReplacing, False
    
    Set oObjectCollectionReplaced = Nothing
    Set oObjectCollectionReplacing = Nothing
    
    Exit Sub
    
ErrorHandler:
''MsgBox "*** ... ERROR ... Migrate_SetReplacedReplacing"
    Err.Raise LogError(Err, MODULE, "Migrate_SetReplacedReplacing", strError).Number
End Sub

Public Function Migrate_IsMbrSplitByPlate(oAppConn As Object, oInput1 As Object, oInput2 As Object, oMigrateHelper As IJMigrateHelper, _
                                          oMbrByPlateSplits As Collection, Optional bAllSplitConections = False) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    Dim iIndex As Long
    Dim jIndex As Long
    
    Dim bIsDeleted As Boolean
    Dim bPlateSystem As Boolean
    Dim bMembersystem As Boolean
    Dim bMbrSplitByPlate As Boolean
    
    Dim oPort As IJPort
'    Dim oParent As Object
'    Dim oDesignChild As IJDesignChild
    
    Dim oInputPart1 As Object
    Dim oInputPart2 As Object
    Dim oPlatePart As Object
    Dim oMemberPart As Object
    Dim oPlateSystem As Object
    Dim oObjectReplacing As Object
    
    Dim oMemberSystem As ISPSMemberSystem
    Dim oSplitAxisPort As ISPSSplitAxisPort
    Dim oMemberPartCommon As ISPSMemberPartCommon
    
    Dim oSplitInputs As IJElements
    Dim oSplitConnections As IJElements
    Dim oSPlitMbrConn As ISPSSplitMemberConnection

    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection

    Dim oStructDetailHelper As StructDetailHelper
    
    Migrate_IsMbrSplitByPlate = False
    Set oMbrByPlateSplits = New Collection
    
    If oAppConn Is Nothing Then
        Exit Function
    ElseIf oInput1 Is Nothing Then
        Exit Function
    ElseIf oInput2 Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oInput1 Is IJPort Then
        Exit Function
    ElseIf Not TypeOf oInput2 Is IJPort Then
        Exit Function
    End If
    
    ' Expecting one input to be a IJPlatePart
    ' And one input to be a ISPSMemberPartPrismatic
    Set oPort = oInput1
    Set oInputPart1 = oPort.Connectable
    
    Set oPort = oInput2
    Set oInputPart2 = oPort.Connectable
    
    If oInputPart1 Is Nothing Then
        Exit Function
    ElseIf oInputPart2 Is Nothing Then
        Exit Function
    End If
    
    If TypeOf oInputPart1 Is ISPSMemberPartCommon Then
        Set oMemberPart = oInputPart1
        If TypeOf oInput1 Is ISPSSplitAxisPort Then
            Set oSplitAxisPort = oInput1
        End If
    ElseIf TypeOf oInputPart1 Is IJPlatePart Then
        Set oPlatePart = oInputPart1
    Else
        Exit Function
    End If
    
    If TypeOf oInputPart2 Is ISPSMemberPartCommon Then
        Set oMemberPart = oInputPart2
        If TypeOf oInput2 Is ISPSSplitAxisPort Then
            Set oSplitAxisPort = oInput2
        End If
    ElseIf TypeOf oInputPart2 Is IJPlatePart Then
        Set oPlatePart = oInputPart2
    Else
        Exit Function
    End If
    
    If oMemberPart Is Nothing Then
        Exit Function
    ElseIf oPlatePart Is Nothing Then
        Exit Function
    End If
    
    'Get member system and root plate system
    Set oMemberPartCommon = oMemberPart
    If oMemberPartCommon.MemberSystem Is Nothing Then
        ' the MemberPart is NOT connected to an existing Member System
        ' ... assume the Member Part is being split
        ' ... use the Replacing Member Parts to determine the Mmember System
        oMigrateHelper.ObjectsReplacing oMemberPart, oObjectCollectionReplacing, bIsDeleted
        If oObjectCollectionReplacing Is Nothing Then
        ElseIf oObjectCollectionReplacing.Count < 1 Then
        Else
            For Each oObjectReplacing In oObjectCollectionReplacing
                If TypeOf oObjectReplacing Is ISPSMemberPartCommon Then
                    Set oMemberPartCommon = oObjectReplacing
                    If Not oMemberPartCommon.MemberSystem Is Nothing Then
                        Set oMemberSystem = oMemberPartCommon.MemberSystem
                        Exit For
                    End If
                End If
            Next
        End If
    Else
        Set oMemberSystem = oMemberPartCommon.MemberSystem
    End If
    
    Set oStructDetailHelper = New StructDetailHelper
    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oPlateSystem, True
    
    If oMemberSystem Is Nothing Then
        Exit Function
    ElseIf oPlateSystem Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oPlateSystem Is IJPlateSystem Then
        Exit Function
    End If

    'Get inputs of each split connection on the member system:
    ' if connectable of an input port happens to be the plate system
    ' ... then it is member-split-by-plate case
    Set oSplitConnections = oMemberSystem.SplitConnections
    If oSplitConnections Is Nothing Then
        Exit Function
    ElseIf oSplitConnections.Count < 1 Then
        Exit Function
    End If
        
    On Error Resume Next
    For iIndex = 1 To oSplitConnections.Count
        bPlateSystem = False
        bMembersystem = False
        bMbrSplitByPlate = False
        
        Set oSPlitMbrConn = oSplitConnections(iIndex)
        Set oSplitInputs = oSPlitMbrConn.InputObjects
        If oSplitInputs Is Nothing Then
        ElseIf oSplitInputs.Count < 1 Then
        Else
            
            For jIndex = 1 To oSplitInputs.Count
                If TypeOf oSplitInputs(jIndex) Is ISPSMemberSystem Then
                    bMembersystem = True
                ElseIf TypeOf oSplitInputs(jIndex) Is IJPort Then
                    Set oPort = oSplitInputs(jIndex)
                    If oPort.Connectable Is oPlateSystem Then
                        bPlateSystem = True
                    End If
                End If
                
                If bPlateSystem Then
                    If bMembersystem Then
                        oMbrByPlateSplits.Add oSPlitMbrConn
                        bMbrSplitByPlate = True
                        Migrate_IsMbrSplitByPlate = True
                        Exit For
                    End If
                
                End If
            Next jIndex
            
            If bAllSplitConections Then
                If Not bMbrSplitByPlate Then
                    oMbrByPlateSplits.Add oSPlitMbrConn
                End If
            End If
            
        End If
        
    Next iIndex
    On Error GoTo ErrorHandler
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "Migrate_IsMbrSplitByPlate", strError).Number
End Function

' ******************************************************************************************
' ******************************************************************************************
' ******************************************************************************************

Public Sub Migrate_CreateReplacingObjects(oMemberItem As Object, oMigrateHelper As IJMigrateHelper, bChildrenMemberItems As Boolean)
Const METHOD = "::Migrate_CreateReplacingObjects"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim bTrace As Boolean
    Dim sMemberObjectType As String
    
    Dim iIndex As Long
    Dim jIndex As Long
    Dim bIsDeleted As Boolean
    Dim bHasReplacing As Boolean

    Dim oPort As IJPort
    Dim oChkObject1 As Object
    Dim oChkObject2 As Object
    Dim oObjectReplacing As Object
    Dim oMemberObjects As IJDMemberObjects
    
    Dim oMemberDatatypes As Collection
    Dim oMemberDataObjects As Collection
    
    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection
    
    bTrace = True
    If oMemberItem Is Nothing Then
        Exit Sub
    ElseIf m_ReplacedParts Is Nothing Then
        Exit Sub
    ElseIf m_ReplacedParts.Count < 1 Then
        Exit Sub
    End If
    
    sMsg = "setup Replaced Port data "
    Set oMemberDatatypes = Nothing
    Set oMemberDataObjects = Nothing
    Migrate_GetMemberData oMemberItem, oMemberDataObjects, oMemberDatatypes
    
    For iIndex = 1 To oMemberDataObjects.Count
        bHasReplacing = False
        If oMemberDataObjects.Item(iIndex) Is Nothing Then
        Else
            If iIndex = 1 Then
                sMemberObjectType = oMemberDatatypes(iIndex)
            End If
            ' Check if the Input Port has Replaced/Replacing data
            ' if it Does NOT
            ' ... check if the Input Port Connectable has Replaced/Replacing data
            ' ... if it Does
            ' ... ... Create Replaced/Replacing data for the Input Port
            Set oChkObject1 = oMemberDataObjects.Item(iIndex)
            oMigrateHelper.ObjectsReplacing oChkObject1, oObjectCollectionReplacing, bIsDeleted
            If oObjectCollectionReplacing Is Nothing Then
            ElseIf oObjectCollectionReplacing.Count > 0 Then
                bHasReplacing = True
            End If
            
            If bHasReplacing Then
            ElseIf TypeOf oChkObject1 Is IJPort Then
                Set oPort = oChkObject1
                Set oChkObject2 = oPort.Connectable
                oMigrateHelper.ObjectsReplacing oChkObject2, oObjectCollectionReplacing, bIsDeleted
                If oObjectCollectionReplacing Is Nothing Then
                ElseIf oObjectCollectionReplacing.Count > 0 Then
                
If bSM_trace Then
    If bTrace Then
        zSM_trace "*** Migrate_CreateReplacingObjects ... oMemberItem: " & sMemberObjectType & " : " & Debug_ObjectName(oMemberItem, True)
    End If
    bTrace = False
End If

                    If sMemberObjectType = "IJStructPhysicalConnection" Then
                        ' Based on "Last" Port
                        Migrate_CreateReplacingPort oMemberItem, oChkObject1, oMigrateHelper, True
                    
                    Else
                        ' Based on Port "After EndCut" or "After Cutout"
                        ' ... If sMemberObjectType = "SF_CornerFeature" Then
                        Migrate_CreateReplacingPort oMemberItem, oChkObject1, oMigrateHelper, False
                    End If
                    
                End If
            ElseIf TypeOf oChkObject1 Is IJConnectable Then
            End If
        
            ' process ALL of the MemberItems created by the current MemberItem
            If bChildrenMemberItems Then
                If TypeOf oMemberItem Is IJDMemberObjects Then
                    Set oMemberObjects = oMemberItem
                    If oMemberObjects.Count > 0 Then
                        For jIndex = 1 To oMemberObjects.Count
                            If oMemberObjects.Item(jIndex) Is Nothing Then
                            Else
                                Migrate_CreateReplacingObjects oMemberObjects.Item(jIndex), oMigrateHelper, bChildrenMemberItems
                            End If
                        Next jIndex
                    End If
                End If
            End If
        End If
        
    Next iIndex
    
   Exit Sub

ErrorHandler:
''MsgBox "*** ... ERROR ... " & METHOD
    HandleError MODULE, METHOD, sMsg
End Sub

Public Sub Migrate_CreateReplacingPort(oMemberObj As Object, oPortToReplace As Object, oMigrateHelper As IJMigrateHelper, bFromLastGeometry As Boolean)
Const METHOD = "::Migrate_CreateReplacingPort"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim iIndex As Long

    Dim lXId As Long
    Dim lCtxId As Long
    Dim lOptId As Long
    Dim lOprId As Long
    Dim lPortType As Long
    Dim lReplacingOptId As Long
    Dim lReplacingOprId As Long
    
    Dim oPort As IJPort
    Dim oReplacingPort As Object
    Dim oPortMoniker As IUnknown
    Dim oSP3D_StructPort As SP3DStructPorts.IJStructPort
    Dim oStructGeomBasicPort As StructGeomBasicPort
    Dim oReplacedMemberPart As Object
    Dim oReplacingMemberPart As Object
    
    Dim oFeature As Object
    Dim oOperation As Object
    Dim eStructOperation As cmnstrStructOperation
    Dim oStructOperationAE As IJStructOperationAE
    
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    
    Dim oACTools As AssemblyConnectionTools
    Dim oSD_Helper As StructDetailObjects.Helper
    Dim oPortHelper As PORTHELPERLib.PortHelper

    If oPortToReplace Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oPortToReplace Is IJPort Then
        Exit Sub
    ElseIf Not TypeOf oPortToReplace Is StructGeomBasicPort Then
        Exit Sub
    End If
    
    sMsg = "setup Replaced Port data "
    Set oPort = oPortToReplace
    Set oPortHelper = New PORTHELPERLib.PortHelper
    Set oReplacedMemberPart = oPort.Connectable
    Set oStructProfilePart = oReplacedMemberPart
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    
    Set oReplacingMemberPart = Nothing
    For iIndex = 1 To m_ReplacedParts.Count
        If oStructProfilePart Is m_ReplacedParts.Item(iIndex) Then
            Set oReplacingMemberPart = m_ReplacingParts.Item(iIndex)
        End If
    Next iIndex
    
    If oReplacingMemberPart Is Nothing Then
        Exit Sub
    ElseIf oReplacingMemberPart Is oReplacedMemberPart Then
        Exit Sub
    End If
    
    ' get Ports Info Xid from the given Replaced Port
    ' ... expect these Ports to be from the Replaced Member Part
    sMsg = "Get Replaced Port data "
    Set oStructGeomBasicPort = oPortToReplace
    Set oSP3D_StructPort = oStructGeomBasicPort
    Set oPortMoniker = oSP3D_StructPort.PortMoniker
    oPortHelper.DecodeTopologyProxyMoniker oPortMoniker, lPortType, lCtxId, lOptId, lOprId, lXId
                
    eStructOperation = cmnstrNoOperation
    Set oACTools = New AssemblyConnectionTools
    
    If lOptId < 1 Then
        ' No Operation Id for the Replaced Port
        ' ... Use the Global Base, Offset, or Lateral Port
        ' ... Assuming OprId and Xid are not meaningful in this case
        oACTools.GetBindingPort oReplacingMemberPart, lPortType, lOptId, lOprId, lCtxId, lXId, _
                                "SPSMembers.SPSPartPrismaticGenerator", oReplacingPort
    
    ElseIf lOprId > 0 Then
        ' get the Operation and Feature that the Replaced Port is created from
        Migrate_GetFeatureFromOptOpr oStructProfilePart, lOptId, lOprId, oFeature, oOperation
        If oOperation Is Nothing Then
        Else
            Set oStructOperationAE = oOperation
            oStructOperationAE.GetOperationType eStructOperation
        End If
    
        ' No Feature was found on the Replaced Part with the given Opt,Opr
        ' check if the Feature has been moved to the Replacing Part
        If oFeature Is Nothing Then
            Migrate_GetFeatureFromOptOpr oReplacingMemberPart, lOptId, lOprId, oFeature, oOperation
            If oOperation Is Nothing Then
            Else
                Set oStructOperationAE = oOperation
                oStructOperationAE.GetOperationType eStructOperation
            End If
        End If
    End If
        
    If lOptId > 0 Then
        ' Get the late binding Port based on Replaced Port Opt,Opr
        ' ... expect the EndCut and/or Cut object to have already been migrated to the New Member Part
        sMsg = "Get Replacing Port data "
        If oFeature Is Nothing Then
            If bFromLastGeometry Then
                oACTools.GetBindingPort oReplacingMemberPart, JS_TOPOLOGY_PROXY_LFACE, lOptId, lOprId, lCtxId, lXId, _
                                        "", oReplacingPort
            Else
                oACTools.GetBindingPort oReplacingMemberPart, JS_TOPOLOGY_PROXY_LFACE, lOptId, lOprId, lCtxId, lXId, _
                                        "SPSMembers.SPSPartPrismaticGenerator", oReplacingPort
            End If
        
        ElseIf bFromLastGeometry Then
            ' GetLatePortForFeatureSegment returns the Port based on "last geometry"
            ' ... valid for Webcut, Flangecut and Slot features only
            ' ... calls CStructSymbolTools::BindMonikerToStructLastPort
            Set oStructProfilePart = oReplacingMemberPart
            Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
            oStructEndCutUtil.GetLatePortForFeatureSegment oFeature, lXId, oReplacingPort

        ElseIf eStructOperation = cmnstrCutoutOperation Then
            ' Get Late Port AFTER the StructCutoutOperationAE operation (after cutouts)
            Migrate_GetOptOprFromFeature oFeature, oReplacingMemberPart, lReplacingOptId, lReplacingOprId, oOperation
            oACTools.GetBindingPort oReplacingMemberPart, JS_TOPOLOGY_PROXY_LFACE, lReplacingOptId, lReplacingOprId, lCtxId, lXId, _
                                    "StructGeneric.StructCutoutOperationAE.1", oReplacingPort
        
        Else
            ' Get Late Port AFTER the SPSPartPrismaticGenerator operation (after endcuts)
            Migrate_GetOptOprFromFeature oFeature, oReplacingMemberPart, lReplacingOptId, lReplacingOprId, oOperation
            oACTools.GetBindingPort oReplacingMemberPart, JS_TOPOLOGY_PROXY_LFACE, lReplacingOptId, lReplacingOprId, lCtxId, lXId, _
                                    "SPSMembers.SPSPartPrismaticGenerator", oReplacingPort
        End If
    
    End If
    
    'Set Replaced, Replacing Port data
    If Not oReplacingPort Is Nothing Then
        sMsg = "Set Replaced/Replacing Port data "
        Migrate_SetReplacedReplacing oPortToReplace, oReplacingPort, oMigrateHelper
        
'' do not need to flag creating Feature as Replaced/Replacing (?)
''        Migrate_SetReplacedReplacing oMemberObj, oMemberObj, oMigrateHelper
    End If
    
If bSM_trace Then
    zSM_trace "*** ... " & METHOD & " : Opt " & Trim(Str(eStructOperation)) & " (FromLast: " & bFromLastGeometry & ") : " & Debug_ObjectName(oMemberObj, True)
    If oPortToReplace Is Nothing Then
        zSM_trace "*** ... Replaced      : (Is Nothing) ... ReplacedPart: " & Debug_ObjectName(oReplacedMemberPart, True)
    ElseIf TypeOf oPortToReplace Is IJPort Then
        zSM_trace "*** ... Replaced      : " & Debug_PortData(oPortToReplace, True)
    Else
        zSM_trace "*** ... Replaced      : " & Debug_ObjectName(oPortToReplace, True)
    End If
    
    If oReplacingPort Is Nothing Then
        zSM_trace "*** ... ... Replacing : (Is Nothing) ... ReplacingPart: " & Debug_ObjectName(oReplacingMemberPart, True)
    ElseIf TypeOf oReplacingPort Is IJPort Then
        zSM_trace "*** ... ... Replacing : " & Debug_PortData(oReplacingPort, True)
    Else
        zSM_trace "*** ... ... Replacing : " & Debug_ObjectName(oReplacingPort, True)
    End If
    
End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Sub

Public Function Migrate_GetMigrateControlFlag(oMigrateObject As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim lValue As Long
    
    ' Flag given object as being in Split\Migration state
    Migrate_GetMigrateControlFlag = False
    If TypeOf oMigrateObject Is IJStructGenericControlFlags Then
        Dim oControlFlags As IJStructGenericControlFlags    'IJControlFlags
        Set oControlFlags = oMigrateObject
        '   CTL_FLAG_USER_MASK      = 0xffff0000
        '   CTL_FLAG_IN_MIGRATE     = 0x10000000
        oControlFlags.GetControlFlags &H10000000, lValue
        If lValue <> 0 Then
            Migrate_GetMigrateControlFlag = True
        End If
    End If
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "Migrate_GetMigrateControlFlag", strError).Number
End Function

Public Function Migrate_SetMigrateControlFlag(oMigrateObject As Object, bFlagState As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim lValue As Long
    
    ' Flag given object as being in Split\Migration state
    Migrate_SetMigrateControlFlag = False
    If TypeOf oMigrateObject Is IJStructGenericControlFlags Then
        Dim oControlFlags As IJStructGenericControlFlags    'IJControlFlags
        Set oControlFlags = oMigrateObject
        '   CTL_FLAG_USER_MASK      = 0xffff0000
        '   CTL_FLAG_IN_MIGRATE     = 0x10000000
        oControlFlags.GetControlFlags &H10000000, lValue
        If lValue <> 0 Then
            Migrate_SetMigrateControlFlag = True
        End If
        
        If bFlagState Then
            oControlFlags.PutControlFlags &H10000000, &H10000000
        Else
            oControlFlags.PutControlFlags &H10000000, &H0
        End If
        Set oControlFlags = Nothing
''    Else
''If bSM_trace Then
''    zSM_trace vbCrLf
''    zSM_trace "*** Migrate_SetMigrateControlFlag ... IJStructGenericControlFlags Not Supported : " & Debug_ObjectName(oMigrateObject, True)
''End If
    End If
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "Migrate_SetMigrateControlFlag", strError).Number
End Function

Public Sub Migrate_SetFeatureMigrateData(oMigrateFeature As Object, oDataFrom As Object, bReplacedData As Boolean)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    Dim lIndex As Long
    Dim lOprId As Long
    Dim lOptId As Long
    Dim lReplacingIndex As Long
    Dim eStructFeatureType As StructFeatureTypes
    
    Dim oPort As IJPort
    Dim oOperation As Object
    Dim oMemberPart As Object
    Dim oStructFeature As IJStructFeature
    
    ' The Feature's Opt/Opr pair might change when placed during migration
    ' ... Need to save the Feature's Member Part, Opt/Opr pair Before migration
    ' ... Need to save the Feature's Member Part, Opt/Opr pair After migration
    ' ... This data is used by Migrate_GetOptOprFromFeature and Migrate_CreateReplacingPort
    ' ... for creating the correct Late Binding Ports based on the Feature
    If oMigrateFeature Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oMigrateFeature Is IJStructFeature Then
        Exit Sub
    ElseIf oDataFrom Is Nothing Then
        Exit Sub
    ElseIf TypeOf oDataFrom Is IJPort Then
        Set oPort = oDataFrom
        Set oMemberPart = oPort.Connectable
    ElseIf TypeOf oDataFrom Is ISPSMemberPartCommon Then
        Set oMemberPart = oDataFrom
    Else
        Exit Sub
    End If

    ' can only get Opt/Opr from specific Struct feature types
    Set oStructFeature = oMigrateFeature
    eStructFeatureType = oStructFeature.get_StructFeatureType
    If eStructFeatureType = SF_FlangeCut Then
    ElseIf eStructFeatureType = SF_WebCut Then
    ElseIf eStructFeatureType = SF_Slot Then
    Else
        Exit Sub
    End If
    
    lReplacingIndex = 0
    For lIndex = 1 To m_MigratedFeaturesCount
        If m_MigratedFeatures(lIndex).Feature Is oMigrateFeature Then
            If bReplacedData Then
                If m_MigratedFeatures(lIndex).ReplacedMember Is oMemberPart Then
                    ' Replaced Feature's Opt/Opr data already exists
                    Exit Sub
                End If
            Else
                ' Save index for the Feature's replacing Opt/Opr
                lReplacingIndex = lIndex
                Exit For
            End If
        End If
    Next lIndex

    ' Get the given Features's Opt/Opr pair
    Migrate_GetOptOprFromFeature oMigrateFeature, oMemberPart, lOptId, lOprId, oOperation
    
    If bReplacedData Then
        m_MigratedFeaturesCount = m_MigratedFeaturesCount + 1
        If m_MigratedFeaturesCount > m_MigratedFeaturesSize Then
            m_MigratedFeaturesSize = m_MigratedFeaturesSize + 100
            ReDim Preserve m_MigratedFeatures(m_MigratedFeaturesSize)
        End If
    
        Set m_MigratedFeatures(m_MigratedFeaturesCount).Feature = oMigrateFeature
    
        m_MigratedFeatures(m_MigratedFeaturesCount).ReplacedOpr = lOprId
        m_MigratedFeatures(m_MigratedFeaturesCount).ReplacedOpt = lOptId
        Set m_MigratedFeatures(m_MigratedFeaturesCount).ReplacedMember = oMemberPart
        Set m_MigratedFeatures(m_MigratedFeaturesCount).ReplacedOperation = oOperation
    
        m_MigratedFeatures(m_MigratedFeaturesCount).ReplacingOpr = 0
        m_MigratedFeatures(m_MigratedFeaturesCount).ReplacingOpt = 0
        Set m_MigratedFeatures(m_MigratedFeaturesCount).ReplacingMember = Nothing
        Set m_MigratedFeatures(m_MigratedFeaturesCount).ReplacingOperation = Nothing
        
    ElseIf lReplacingIndex > 0 Then
        m_MigratedFeatures(lReplacingIndex).ReplacingOpr = lOprId
        m_MigratedFeatures(lReplacingIndex).ReplacingOpt = lOptId
        Set m_MigratedFeatures(lReplacingIndex).ReplacingMember = oMemberPart
        Set m_MigratedFeatures(lReplacingIndex).ReplacingOperation = oOperation
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "Migrate_SetFeatureMigrateData", strError).Number
End Sub

' ******************************************************************************************
' ******************************************************************************************
' ******************************************************************************************

Public Sub Migrate_TraceMembers(oMigrateObject As Object, oMigrateHelper As IJMigrateHelper, sLevel As String)
    On Error GoTo ErrorHandler

    Dim iIndex As Long
    Dim lDispId As Long
    Dim lMbrIdx As Long
    Dim sLevel1 As String
    Dim strError As String
    Dim sSmartItemDef As String
    Dim sMemberItemName As String
    
    Dim oSmartItem As IJSmartItem
    Dim oMemberObjects As IJDMemberObjects
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oMemberDescription As IJDMemberDescription
    Dim oMemberDescriptions As IJDMemberDescriptions
    
    Dim oMemberDatatypes As Collection
    Dim oMemberDataOjects As Collection

    If oMigrateObject Is Nothing Then
        Exit Sub
    End If
    
    Migrate_GetMemberData oMigrateObject, oMemberDataOjects, oMemberDatatypes
    Migrate_TraceObject oMigrateObject, oMemberDataOjects, oMemberDatatypes, oMigrateHelper, sLevel
    
    If Not TypeOf oMigrateObject Is IJDMemberObjects Then
        Exit Sub
    End If
    
    Set oMemberObjects = oMigrateObject
    If oMemberObjects.Count < 1 Then
        Exit Sub
    End If
    
    Set oMemberDescriptions = oMemberObjects.MemberDescriptions
    
    ' For each Member Item
    For iIndex = 1 To oMemberObjects.Count
        If oMemberObjects.Item(iIndex) Is Nothing Then
        Else
            sLevel1 = sLevel & "." & Format(iIndex, "000")
            
            If TypeOf oMemberObjects.Item(iIndex) Is IJSmartOccurrence Then
                Set oSmartOccurrence = oMemberObjects.Item(iIndex)
                Set oSmartItem = oSmartOccurrence.ItemObject
                If Not oSmartItem Is Nothing Then
                    sSmartItemDef = oSmartItem.definition
                    If Len(Trim(sSmartItemDef)) < 1 Then
                        sSmartItemDef = oSmartItem.SymbolDefinition
                    End If
                    
                    oMemberObjects.GetItemDispid oSmartOccurrence, lDispId, lMbrIdx
                    Set oMemberDescription = oMemberDescriptions.ItemByDispid(lDispId)
                    sMemberItemName = oMemberDescription.Name
            
                    zSM_trace sLevel1 & "... SmartItemDef: " & sSmartItemDef & "... (" & Trim(Str(lDispId)) & ") MemberItemName: " & sMemberItemName
                    
                End If
            End If
                
            Migrate_TraceMembers oMemberObjects.Item(iIndex), oMigrateHelper, sLevel1
        End If
    Next iIndex
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "MigrateAssemblyConnection_EndCutsBounding", strError).Number
End Sub

Public Sub Migrate_TraceObject(oMigrateObject As Object, oMemberDataOjects As Collection, oMemberDatatypes As Collection, oMigrateHelper As IJMigrateHelper, sLevel As String)
    On Error GoTo ErrorHandler
    Dim iIndex As Long
    Dim jIndex As Long
    Dim lMbrIdx As Long
    Dim lDispId As Long
    
    Dim strError As String
    Dim sMemberObjectType As String
    
    Dim oPort As IJPort
    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection

''If bSM_trace Then zSM_trace sLevel & "... TypeOf oMigrateObject Is : " & sMemberObjectType & " : " & Debug_ObjectName(oMigrateObject, True)
        
    Dim bFreeEndCut As Boolean
    Dim bHasReplacing As Boolean
    
    Dim sData1 As String
    Dim sData2 As String
    Dim sData3 As String
    Dim sData4 As String
    Dim bIsDeleted As Boolean
    
    Dim oChkObject1 As Object
    Dim oChkObject2 As Object
    Dim oSmartParent As Object
    
    iIndex = 0
    If oMemberDataOjects Is Nothing Then
    ElseIf oMemberDataOjects.Count > 0 Then
        For iIndex = 1 To oMemberDataOjects.Count
            sData1 = ""
            sData2 = ""
            sData3 = ""
            sData4 = ""
            bFreeEndCut = False
            bHasReplacing = False
            
            If oMemberDataOjects.Item(iIndex) Is Nothing Then
            Else
                Set oChkObject1 = oMemberDataOjects.Item(iIndex)
                sMemberObjectType = oMemberDatatypes.Item(iIndex)
                
                sData1 = sLevel & "..."
                If iIndex > 1 Then
                    sData1 = sData1 & Format(iIndex - 1, "000")
                End If
                sData1 = sData1 & " " & oMemberDatatypes.Item(iIndex) & " "
                sData3 = sData1
                
                If oMemberDatatypes.Item(iIndex) = "SmartParent" Then
                    If TypeOf oChkObject1 Is IJFreeEndCut Then
                        bFreeEndCut = True
                        Set oSmartParent = oChkObject1
                    End If
                End If
                
                oMigrateHelper.ObjectsReplacing oChkObject1, oObjectCollectionReplacing, bIsDeleted
                If oObjectCollectionReplacing Is Nothing Then
                ElseIf oObjectCollectionReplacing.Count > 0 Then
                    bHasReplacing = True
                    For Each oChkObject2 In oObjectCollectionReplacing
                        If Len(Trim(sData2)) > 0 Then
                            sData2 = sData2 & vbCrLf & " " & sData1 & "... ...Replacing "
                        Else
                            sData2 = sData1 & "... ...Replacing "
                        End If
                        
                        If oChkObject2 Is Nothing Then
                        ElseIf TypeOf oChkObject2 Is IJPort Then
                            sData2 = sData2 & Debug_PortData(oChkObject2, True)
                        Else
                            sData2 = sData2 & Debug_ObjectName(oChkObject2, True)
                        End If
                    Next oChkObject2
                    
                    sData1 = sData1 & "... Replaced " & Format(oObjectCollectionReplacing.Count, "000") & " "
                    Set oObjectCollectionReplacing = Nothing
                End If
                
                If TypeOf oChkObject1 Is IJPort Then
                    sData1 = sData1 & Debug_PortData(oChkObject1, True)
                Else
                    sData1 = sData1 & Debug_ObjectName(oChkObject1, True)
                End If
            
                ' check if Port has valid Replacing Ports
                ' if not, check if the Port's Part has valid Replacing Parts
                ' ... this indicates the Port needs to have Replacing Port(s)
                If bHasReplacing Then
                ElseIf TypeOf oChkObject1 Is IJPort Then
                    Set oPort = oChkObject1
                    Set oChkObject2 = oPort.Connectable
                    oMigrateHelper.ObjectsReplacing oChkObject2, oObjectCollectionReplacing, bIsDeleted
                    If oObjectCollectionReplacing Is Nothing Then
                    ElseIf oObjectCollectionReplacing.Count > 0 Then
                        sData2 = sData3 & "... ***Replacing Ports required "
                    End If
                End If
                
                ' Find the Operation that current Port is from
                Dim lXId As Long
                Dim lOprId As Long
                Dim lOptId As Long
                Dim lCtxId As Long
                Dim lPortType As Long
                Dim eStructOperation As cmnstrStructOperation
                
                Dim oPortMoniker As IUnknown
                Dim oSP3D_StructPort As SP3DStructPorts.IJStructPort
                Dim oStructGeomBasicPort As StructGeomBasicPort
                Dim oPortHelper As PORTHELPERLib.PortHelper
    
                Dim oOperatorsList As IJElements
                Dim oOperationObject As Object
                Dim oStructOperationAE As IJStructOperationAE
                Dim oStructGraphNavigate As IJStructGraphNavigate
                
''                Dim oIPersist As IPersist
''                oIPersist.GetClassID
                
                If Not TypeOf oChkObject1 Is IJPort Then
                ElseIf Not TypeOf oChkObject1 Is StructGeomBasicPort Then
                ElseIf Not TypeOf oChkObject1 Is IJStructPort Then
                Else
                    Set oPort = oChkObject1
                    Set oChkObject2 = oPort.Connectable
                    If Not TypeOf oChkObject2 Is ISPSMemberPartCommon Then
                    ElseIf Not TypeOf oChkObject2 Is IJStructGraphNavigate Then
                    Else
                        Set oStructGeomBasicPort = oChkObject1
                        Set oSP3D_StructPort = oStructGeomBasicPort
                        Set oPortMoniker = oSP3D_StructPort.PortMoniker
                        Set oPortHelper = New PORTHELPERLib.PortHelper
                        oPortHelper.DecodeTopologyProxyMoniker oPortMoniker, lPortType, lCtxId, lOptId, lOprId, lXId
                        
                        Set oStructGraphNavigate = oChkObject2
                        oStructGraphNavigate.FindOperationInGraph lOptId, oOperationObject
                        If oOperationObject Is Nothing Then
                        Else
                            Set oStructOperationAE = oOperationObject
                            oStructOperationAE.GetOperationType eStructOperation
                            oStructOperationAE.GetOperatorById lOprId, oOperatorsList
                            If oOperatorsList Is Nothing Then
                            ElseIf oOperatorsList.Count < 1 Then
                            Else
                                If eStructOperation = cmnstrAddFeatureOperation Then
                                ElseIf eStructOperation = cmnstrBoundOperation Then
                                    sData4 = "BoundOperation"
                                ElseIf eStructOperation = cmnstrChamferCutOperation Then
                                    sData4 = "ChamferCutOperation"
                                ElseIf eStructOperation = cmnstrConnectionSplitOperation Then
                                    sData4 = "ConnectionSplitOperation"
                                ElseIf eStructOperation = cmnstrCutoutOperation Then
                                    sData4 = "CutoutOperation"
                                ElseIf eStructOperation = cmnstrDesignSplitOperation Then
                                    sData4 = "DesignSplitOperation"
                                ElseIf eStructOperation = cmnstrFlangedEdgeOperation Then
                                    sData4 = "FlangedEdgeOperation"
                                ElseIf eStructOperation = cmnstrGeneratePartOperation Then
                                    sData4 = "GeneratePartOperation"
                                ElseIf eStructOperation = cmnstrLimitOperation Then
                                    sData4 = "LimitOperation"
                                ElseIf eStructOperation = cmnstrNoOperation Then
                                    sData4 = "NoOperation"
                                ElseIf eStructOperation = cmnstrPartFinalTrimOperation Then
                                    sData4 = "PartFinalTrimOperation"
                                ElseIf eStructOperation = cmnstrPartSplitOperation Then
                                    sData4 = "PartSplitOperation"
                                ElseIf eStructOperation = cmnstrPlanningSplitOperation Then
                                    sData4 = "PlanningSplitOperation"
                                ElseIf eStructOperation = cmnstrSatFileOperation Then
                                    sData4 = "SatFileOperation"
                                ElseIf eStructOperation = cmnstrSplitOperation Then
                                    sData4 = "SplitOperation"
                                ElseIf eStructOperation = cmnstrStrakingSplitOperation Then
                                    sData4 = "StrakingSplitOperation"
                                ElseIf eStructOperation = cmnstrSweepOperation Then
                                    sData4 = "SweepOperation"
                                ElseIf eStructOperation = cmnstrThickenPlateOperation Then
                                    sData4 = "ThickenPlateOperation"
                                ElseIf eStructOperation = cmnstrTrimOperation Then
                                    sData4 = "TrimOperation"
                                ElseIf eStructOperation = cmnstrTrimPlateOperation Then
                                    sData4 = "TrimPlateOperation"
                                Else
                                    sData4 = Trim(Str(eStructOperation))
                                End If
                            
                                If Len(Trim(sData2)) > 0 Then
                                    sData2 = sData2 & vbCrLf & " "
                                End If
                                sData2 = sData2 & _
                                         sData3 & "... ***StructOperation = " & sData4 & " : " & TypeName(oStructOperationAE)
                            End If
                        End If
                    End If
                End If
            
                If Len(Trim(sData1)) > 0 Then
                    zSM_trace sData1
                    If Len(Trim(sData2)) > 0 Then
                        zSM_trace sData2
                    End If
                End If
            End If
        Next
    End If
        
    If TypeOf oMemberDataOjects.Item(1) Is IJStructAssemblyConnection Then
        Dim oMbrByPlateSplits As Collection
        Migrate_IsMbrSplitByPlate oMemberDataOjects.Item(1), oMemberDataOjects.Item(2), oMemberDataOjects.Item(3), oMigrateHelper, _
                                  oMbrByPlateSplits, True
        If oMbrByPlateSplits Is Nothing Then
        ElseIf oMbrByPlateSplits.Count < 1 Then
        Else
            For iIndex = 1 To oMbrByPlateSplits.Count
                If oMbrByPlateSplits.Item(iIndex) Is Nothing Then
                Else
                    sData1 = sLevel & "......" & Format(iIndex, "000") & " SplitConnections: "
                    sData1 = sData1 & Debug_ObjectName(oMbrByPlateSplits.Item(iIndex), True)
                    zSM_trace sData1
                End If
            Next iIndex
        End If
        
    End If
    
'' No need to trace owning Free EndCut
''    If bFreeEndCut Then
''        Dim oMemberDataTypes1 As Collection
''        Dim oMemberDataOjects1 As Collection
''        Migrate_GetMemberData oSmartParent, oMemberDataOjects1, oMemberDataTypes1
''        Migrate_TraceObject oSmartParent, oMemberDataOjects1, oMemberDataTypes1, oMigrateHelper, sLevel & ".FEC"
''    End If
    
    Exit Sub
    
ErrorHandler:
If bSM_trace Then zSM_trace "*** ... ERROR ... Migrate_TraceObject"
''MsgBox "*** ... ERROR ... Migrate_TraceObject"
    Err.Raise LogError(Err, MODULE, "Migrate_TraceObject", strError).Number
End Sub

Public Sub Migrate_TraceOperators(oMemberPart As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    'Get Endcuts on pattern
    ' ... SPSMembers.SPSPartPrismaticGenerator.1
    'Get Cutouts on pattern
    ' ... SP3DStructGeneric.StructCutoutOperation.1 ... StructGeneric.StructCutoutOperationAE.1
    Dim oOperator As Object
    Dim OperationPattern As IJStructOperationPattern
    Dim oStructOperationAE As Object
    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    If TypeOf oMemberPart Is IJStructOperationPattern Then
                zSM_trace "*** *** ... MemberPart Operators : " & Debug_ObjectName(oMemberPart, True)
        
        Set OperationPattern = oMemberPart
        OperationPattern.GetOperationPattern "SPSMembers.SPSPartPrismaticGenerator.1", oCollectionOfOperators, oStructOperationAE
        If oCollectionOfOperators Is Nothing Then
        ElseIf oCollectionOfOperators.Count > 0 Then
            For Each oOperator In oCollectionOfOperators
                zSM_trace "*** *** ... ... Endcut Operator  : " & Debug_ObjectName(oOperator, True)
            Next
        End If
        
        OperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructOperationAE
        If oCollectionOfOperators Is Nothing Then
        ElseIf oCollectionOfOperators.Count > 0 Then
            For Each oOperator In oCollectionOfOperators
                zSM_trace "*** *** ... ... Cutout Operator  : " & Debug_ObjectName(oOperator, True)
            Next
        End If
    
        OperationPattern.GetOperationPattern "SP3DStructGeneric.StructCutoutOperation.1", oCollectionOfOperators, oStructOperationAE
        If oCollectionOfOperators Is Nothing Then
        ElseIf oCollectionOfOperators.Count > 0 Then
            For Each oOperator In oCollectionOfOperators
                zSM_trace "*** *** ... ... Cutout Operator  : " & Debug_ObjectName(oOperator, True)
            Next
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
If bSM_trace Then zSM_trace "*** ... ERROR ... Migrate_TraceOperators"
    Err.Raise LogError(Err, MODULE, "Migrate_TraceOperators", strError).Number
End Sub

