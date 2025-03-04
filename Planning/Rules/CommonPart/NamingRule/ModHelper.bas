Attribute VB_Name = "ModHelper"
 
Option Explicit
 
Public Sub UpdateAlphabet(strBracketName As String, strLastAlphabet As String)

    'Get the last char of BracketName
    Dim strChar As String
    strChar = Right$(strBracketName, 1)
    
    Dim iChar As Integer
    iChar = Asc(strChar)
    If iChar >= 65 And iChar <= 90 Then
        If iChar >= Asc(strLastAlphabet) Then
            strLastAlphabet = Chr(iChar + 1)
        End If
    End If
    
End Sub


Public Function GetConnectedProfilePart(oBracket As Object, bSmartPlate As Boolean) As IJProfilePart
Const METHOD = "GetConnectedProfilePart"
On Error GoTo ErrorHandler
    
    Dim oProfilepartColl1 As IJElements
    Set oProfilepartColl1 = New JObjectCollection
            
    Dim oProfilePartColl2 As IJElements
    Set oProfilePartColl2 = New JObjectCollection
        
        
    ' If the Bracket is a StructDetailing object
    If bSmartPlate Then
                
        Dim oSDBracket As StructDetailObjects.Bracket
        Dim oRefPlane  As IJPlane
        Dim oSupport1  As IJPort
        Dim oSupport2  As IJPort
        Dim oSupport3  As Object
        Dim oSupport4  As Object
        Dim oSupport5  As Object
        Dim nSupports  As Long
    
        Set oSDBracket = New StructDetailObjects.Bracket
        Set oSDBracket.object = oBracket

        oSDBracket.GetInputs nSupports, oRefPlane, oSupport1, oSupport2, _
                                        oSupport3, oSupport4, oSupport5
        
        If TypeOf oSupport1.Connectable Is IJProfilePart Then
            oProfilepartColl1.Add oSupport1.Connectable
        ElseIf TypeOf oSupport2.Connectable Is IJProfilePart Then
            oProfilepartColl1.Add oSupport2.Connectable
        End If
        
    Else ' MoledForm Bracket
        
        Dim oBracketPlatePart              As IJDesignChild
        Dim oBracketPlateLeaf              As IJDesignChild
        Dim oBracketPlateSys               As IJPlate
                     
        Set oBracketPlatePart = oBracket
        Set oBracketPlateLeaf = oBracketPlatePart.GetParent
        
        Set oBracketPlateSys = oBracketPlateLeaf.GetParent
                
        Dim oPlateAttr As IJPlateAttributes
        Set oPlateAttr = New PlateUtils
                                
        Dim oprofileSys     As IJSystem
        
        If oPlateAttr.IsBracketByPlane(oBracketPlateSys) Then
            
            Dim oBracketPlane       As IJPlane
            Dim oSupportColl        As IJElements
            Dim strRootSelector     As String
            Dim oUVec               As IJPoint
            Dim oVVec               As IJPoint
            Dim i                   As Integer
            
            oPlateAttr.GetInput_BracketByPlane oBracketPlateSys, oBracketPlane, _
                                    oUVec, oVVec, strRootSelector, oSupportColl
            
            For i = 1 To oSupportColl.Count
            
                If TypeOf oSupportColl.Item(i) Is IJProfile Then
                    Set oprofileSys = oSupportColl.Item(i)
                    Exit For
                End If
            Next
            
        ElseIf oPlateAttr.IsTrippingBracket(oBracketPlateSys) Then
        
            Dim oBracketUVector As IJDVector
            Dim oBracketVVector As IJDVector
            Dim oInput1 As Object
            Dim oInput2 As Object
            Dim oInput3 As Object
            Dim oBracketAE As IJPlaneByElements_AE
            
            oPlateAttr.GetInput_TrippingBracket oBracketPlateSys, oBracketUVector, oBracketVVector, _
                                            oInput1, oInput2, oInput3, oBracketAE
                                            
                    
            If TypeOf oInput1 Is IJProfile Then
                Set oprofileSys = oInput1
            ElseIf TypeOf oInput2 Is IJProfile Then
                Set oprofileSys = oInput2
            ElseIf TypeOf oInput3 Is IJProfile Then
                Set oprofileSys = oInput3
            End If
            
        End If
        
        
        'Get all the ProfileParts under Profile System.
        Dim oCol            As IJDTargetObjectCol
        Dim iProfileIndex   As Integer
        
        Set oCol = oprofileSys.GetChildren
        
        'First get the leaf system
        For iProfileIndex = 1 To oCol.Count
            If TypeOf oCol.Item(iProfileIndex) Is IJProfile Then
                Set oprofileSys = oCol.Item(iProfileIndex)
                Exit For
            End If
        Next
    
        Set oCol = Nothing
        
        Set oCol = oprofileSys.GetChildren
    
        For iProfileIndex = 1 To oCol.Count
            If TypeOf oCol.Item(iProfileIndex) Is IJProfilePart Then
                oProfilepartColl1.Add oCol.Item(iProfileIndex)
            End If
        Next
    
    End If
                
            
    'Get all the ProfileParts connected to Bracket
    Dim iIndex As Long
        
    Dim oSDHelper As StructDetailObjects.Helper
    Dim nConnectionData As Long
    Dim aConnectionData As ConnectionData
    Dim zConnectionData() As ConnectionData
    Dim oConnection As IJStructPhysicalConnection
    
    Set oSDHelper = New StructDetailObjects.Helper
    
    ' retrieve all of the physical connections from the object
    nConnectionData = 0
    
    oSDHelper.Object_AppConnections oBracket, AppConnectionType_Physical, _
                                    nConnectionData, zConnectionData
    
    For iIndex = 1 To nConnectionData
        aConnectionData = zConnectionData(iIndex)
        Set oConnection = aConnectionData.AppConnection
        
        If Not oConnection Is Nothing Then
                   
            'use StructDetailObjects.PhysicalConn to get the connected objects
            Dim oWrapSDPhysConn As StructDetailObjects.PhysicalConn
            Set oWrapSDPhysConn = New StructDetailObjects.PhysicalConn
            Set oWrapSDPhysConn.object = oConnection
                                
            If TypeOf oWrapSDPhysConn.ConnectedObject1 Is IJProfilePart Then
                oProfilePartColl2.Add oWrapSDPhysConn.ConnectedObject1
            ElseIf TypeOf oWrapSDPhysConn.ConnectedObject2 Is IJProfilePart Then
                oProfilePartColl2.Add oWrapSDPhysConn.ConnectedObject2
            End If
        
        End If
        
        Set oConnection = Nothing
    Next
    
    Dim lIndex1       As Long, lIndex2 As Long
    Dim oProfilePart1 As IJProfilePart
    Dim oProfilePart2 As IJProfilePart
    
    'Compare the 2 collections and select the common ProfilPart
    For lIndex1 = 1 To oProfilepartColl1.Count
    
        Set oProfilePart1 = oProfilepartColl1.Item(lIndex1)
        
        For lIndex2 = 1 To oProfilePartColl2.Count
        
            Set oProfilePart2 = oProfilePartColl2.Item(lIndex2)
            
            If oProfilePart1 = oProfilePart2 Then
                Set GetConnectedProfilePart = oProfilePart1
                GoTo CleanUp
            End If
        Next
        
    Next
    
CleanUp:

    Set oProfilepartColl1 = Nothing
    Set oProfilePartColl2 = Nothing
    Set oSDHelper = Nothing
    
Exit Function
                 
ErrorHandler:
    GoTo CleanUp
End Function

Public Function IsBracket(oPlateObj As Object) As Boolean
Const METHOD = "IsBracket"
On Error GoTo ErrorHandler

    If TypeOf oPlateObj Is IJSmartPlate Then
        IsBracket = True
        GoTo CleanUp
    End If
    
    Dim oBracketPlateSys               As IJPlate
    
    If TypeOf oPlateObj Is IJPlatePart Then
        Dim oBracketPlatePart              As IJDesignChild
        Dim oBracketPlateLeaf              As IJDesignChild
                     
        Set oBracketPlatePart = oPlateObj
        Set oBracketPlateLeaf = oBracketPlatePart.GetParent
        Set oBracketPlateSys = oBracketPlateLeaf.GetParent
    Else
        Set oBracketPlateSys = oPlateObj
    End If
    
    Dim oBracketAttr As IJPlateAttributes
    Set oBracketAttr = New PlateUtils
    
    On Error Resume Next
    IsBracket = oBracketAttr.IsBracketByPlane(oBracketPlateSys)
    
    If Not IsBracket Then
        IsBracket = oBracketAttr.IsTrippingBracket(oBracketPlateSys)
    End If
    
CleanUp:
    Set oBracketAttr = Nothing
    
Exit Function
                 
ErrorHandler:
    GoTo CleanUp
End Function

Public Function GetBracketItemName(oBracketObj As IJPlatePart) As String
Const METHOD = "GetBracketItemName"
On Error GoTo ErrorHandler
    
    Dim oPlateSys       As IJPlate
    Dim oPlateLeaf      As IJDesignChild
    Dim oBracketSO      As IJSmartOccurrence
    Dim oAttribute      As IJDAttribute
    Dim bAttrFound      As Boolean
    Dim oBracketItem    As IJSmartItem
    
    Dim strBracketByPlaneType           As String
    Dim oCodelistMetadata               As IJDCodeListMetaData
    
    Set oPlateLeaf = oBracketObj
    
    Set oPlateLeaf = oPlateLeaf.GetParent
    Set oPlateSys = oPlateLeaf.GetParent

    Dim oPlateUtils As IJBracketAttributes
    Set oPlateUtils = New PlateUtils
    
    oPlateUtils.GetBracketByPlaneSO oPlateSys, oBracketSO
    
    'BracketByPlane
    If Not oBracketSO Is Nothing Then
        Set oBracketItem = oBracketSO.ItemObject
        GetBracketItemName = oBracketItem.Name
    End If
    
CleanUp:

    Set oBracketSO = Nothing
    Set oPlateSys = Nothing
    Set oPlateLeaf = Nothing
    Set oPlateUtils = Nothing
    Set oAttribute = Nothing
    Set oCodelistMetadata = Nothing
    
    Exit Function
                 
ErrorHandler:
    MsgBox Err.Description, vbInformation, METHOD
    GoTo CleanUp
End Function

Private Function GetAttribute(oObject As Object, strAttrName As String, bAttrFound As Boolean) As IJDAttribute
Const METHOD = "GetAttribute"
On Error GoTo ErrorHandler

    Dim varInterfaceID                  As Variant
    Dim oAttributes                     As IJDAttributes
    Dim oAttr                           As IJDAttribute
    Dim oAttrCol                        As IJDAttributesCol
    
    Set oAttributes = oObject
    bAttrFound = False

    If oAttributes Is Nothing Then Exit Function
    
    'iterate through all interfaces and search for the attribute
    For Each varInterfaceID In oAttributes
        Set oAttrCol = oAttributes.CollectionOfAttributes(varInterfaceID)
        
        If Not oAttrCol Is Nothing Then
            
            If oAttrCol.InterfaceInfo.IsHardCoded = False Then
                 For Each oAttr In oAttrCol
                    If Not oAttr Is Nothing Then
                        If (oAttr.AttributeInfo.Name = strAttrName) And (oAttr.AttributeInfo.OnPropertyPage = True) Then
                            If IsEmpty(oAttr.Value) = False Then
                                Set GetAttribute = oAttr
                                bAttrFound = True
                                GoTo wrapup
                            End If
                        End If
                    End If
                    Set oAttr = Nothing
                Next
                Set oAttrCol = Nothing
            End If
            
        End If
        
    Next
    
wrapup:
    Set oAttributes = Nothing
    Set oAttr = Nothing
    Set oAttrCol = Nothing
        
Exit Function
ErrorHandler:
'    ReportUnanticipatedError MODULE, METHOD
End Function

Private Function GetResMgrByDBType(strDBType As String) As Object
On Error GoTo ErrorHandler

    Dim jContext            As IJContext
    Dim oResourceMgr        As IUnknown
    Dim oDBTypeConfig       As IJDBTypeConfiguration
    Dim oConnectMiddle      As IJDAccessMiddle
    Dim strDBID             As String
    Dim oPOM                As IJDPOM
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strDBID = oDBTypeConfig.get_DataBaseFromDBType(strDBType)
    Set oResourceMgr = oConnectMiddle.GetResourceManager(strDBID)
    
    If Not oResourceMgr Is Nothing Then
        Set GetResMgrByDBType = oResourceMgr
    End If

Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, "GetResMgrByDBType"
End Function

Public Function IsWebBaseAngle90(oCollar As IJCollarPart) As Boolean
On Error GoTo ErrorHandler

   Dim oCollarWrapper   As Object
   Set oCollarWrapper = CreateObject("StructDetailObjects.Collar")

   Set oCollarWrapper.object = oCollar

   Dim oSlotWrapper As Object
   Set oSlotWrapper = CreateObject("StructDetailObjects.Slot")

   Dim oTopologyLocate As TopologyLocate
   Dim oPenetrationPoint As IJDPosition

   Dim oBasePlatePort As IJPort
   Dim oPointOnBasePlate As IJDPosition
   Dim oBasePortGeom As Object
   Dim oBasePlateNormal As IJDVector

   On Error Resume Next

   Set oSlotWrapper.object = oCollarWrapper.Slot
   Set oPenetrationPoint = oSlotWrapper.PenetrationLocation
   Set oTopologyLocate = New TopologyLocate

   ' Get base normal
   Set oBasePortGeom = oTopologyLocate.GetBasePlatePort(oSlotWrapper.Penetrating)
   oTopologyLocate.GetProjectedPointOnModelBody oBasePortGeom, _
                                                oPenetrationPoint, _
                                                oPointOnBasePlate, _
                                                oBasePlateNormal

   Dim oPenetratedPort As IJPort

   If TypeOf oSlotWrapper.Penetrated Is IJPlatePart Then
      Dim oPlatePartWrapper As Object
      Set oPlatePartWrapper = CreateObject("StructDetailObjects.PlatePart")

      ' Penetrated is plate,use base or offset
      Set oPlatePartWrapper.object = oSlotWrapper.Penetrated
      If oCollar.SideOfPlate = 0 Then
         Set oPenetratedPort = oPlatePartWrapper.BasePort(BPT_Base)
      Else
         Set oPenetratedPort = oPlatePartWrapper.BasePort(BPT_Offset)
      End If
   Else
      ' Penetrated is profile part,use web left port for now
      Dim oProfilePartWrapper As Object
      Set oProfilePartWrapper = CreateObject("StructDetailObjects.ProfilePart")

      Set oProfilePartWrapper.object = oSlotWrapper.Penetrated
      Set oPenetratedPort = oProfilePartWrapper.SubPort(JXSEC_WEB_LEFT)
   End If

   Dim oPointOnPenetrated As IJDPosition
   Dim oPenetratedNormal As IJDVector

   oTopologyLocate.GetProjectedPointOnModelBody oPenetratedPort.Geometry, _
                                                oPenetrationPoint, _
                                                oPointOnPenetrated, _
                                                oPenetratedNormal
   Dim dDot As Double
   Dim dAngle As Double
      Dim oProfileWrapper As Object
      Set oProfileWrapper = CreateObject("StructDetailObjects.ProfilePart")

   Set oProfileWrapper.object = oSlotWrapper.Penetrating

   ' Get angle between web port and base plate
   Dim oWebPort As IJPort
   Dim oPointOnWeb As IJDPosition
   Dim oWebNormal As IJDVector

   Set oWebPort = oTopologyLocate.GetSubPort(oSlotWrapper.Penetrating, JXSEC_WEB_RIGHT)
   oTopologyLocate.GetProjectedPointOnModelBody oWebPort.Geometry, _
                                                oPenetrationPoint, _
                                                oPointOnWeb, _
                                                oWebNormal
   Dim oPenetratedCrossBase As IJDVector
   Dim oPenetratedCrossWeb As IJDVector

   Set oPenetratedCrossBase = oPenetratedNormal.Cross(oBasePlateNormal)
   Set oPenetratedCrossWeb = oPenetratedNormal.Cross(oWebNormal)

   oPenetratedCrossBase.Length = 1
   oPenetratedCrossWeb.Length = 1
   dDot = oPenetratedCrossBase.Dot(oPenetratedCrossWeb)

   If dDot < -1 Then
      dDot = -1
   ElseIf dDot > 1 Then
      dDot = 1
   End If

   If Abs(dDot) < 0.000001 Then
      IsWebBaseAngle90 = True
   Else
      IsWebBaseAngle90 = False
   End If

Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, "IsWebBaseAngle90"
End Function

Public Function GetBracketClassName(oBracketObj As IJPlatePart) As String
Const METHOD = "GetBracketClassName"
On Error GoTo ErrorHandler
    
    Dim oPlateSys       As IJPlate
    Dim oPlateLeaf      As IJDesignChild
    Dim oBracketSO      As IJSmartOccurrence
    Dim oAttribute      As IJDAttribute
    Dim bAttrFound      As Boolean
    
    Dim strBracketByPlaneType           As String
    Dim oCodelistMetadata               As IJDCodeListMetaData
    
    Set oPlateLeaf = oBracketObj
    
    Set oPlateLeaf = oPlateLeaf.GetParent
    Set oPlateSys = oPlateLeaf.GetParent

    Dim oPlateUtils As IJBracketAttributes
    Set oPlateUtils = New PlateUtils
    
    oPlateUtils.GetBracketByPlaneSO oPlateSys, oBracketSO
    
    'BracketByPlane
    If Not oBracketSO Is Nothing Then
        Set oAttribute = GetAttribute(oBracketSO, "BracketByPlaneType", bAttrFound)
        
        If bAttrFound Then
          
            Set oCodelistMetadata = GetResMgrByDBType("Catalog")
            
            strBracketByPlaneType = oCodelistMetadata.LongStringValue(oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
        
            GetBracketClassName = strBracketByPlaneType
        Else
            GetBracketClassName = ""
        End If
    End If
    
CleanUp:

    Set oBracketSO = Nothing
    Set oPlateSys = Nothing
    Set oPlateLeaf = Nothing
    Set oPlateUtils = Nothing
    Set oAttribute = Nothing
    Set oCodelistMetadata = Nothing
    
    Exit Function
                 
ErrorHandler:
    MsgBox Err.Description, vbInformation, METHOD
    GoTo CleanUp
End Function

Public Function IsStandardCollar(oCollarPart As Object) As Boolean
    Const METHOD = "IsStandardCollar"
    On Error GoTo ErrorHandler
    
    IsStandardCollar = IsWebBaseAngle90(oCollarPart)
    
Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, METHOD
End Function

Public Function IsStandardBracket(oBracketPart As Object) As Boolean
    Const METHOD = "IsStandardBracket"
    On Error GoTo ErrorHandler
    
    Dim oSmartPlate                 As IJSmartPlate
    Dim oSDOBracket                 As Object
    Dim sObjectType                 As String
    Dim strBracketItemName          As String
    
    Set oSDOBracket = CreateObject("StructDetailObjects.Bracket")

    If TypeOf oBracketPart Is IJSmartPlate Then
        Set oSmartPlate = oBracketPart
    
        If Not oSmartPlate Is Nothing Then ' For StructDetailing Brackets

            Set oSDOBracket.object = oSmartPlate
            sObjectType = oSDOBracket.ClassName

            If (sObjectType = "2SBracketLinear") Then  ' "2SLT" And sObjectType <> "2SLT_wScallop") Then
                IsStandardBracket = True
            End If
        Else
            IsStandardBracket = False
        End If
    Else  ' For MoldedForms Brackets
        
        strBracketItemName = GetBracketClassName(oBracketPart)
        If strBracketItemName = "2SLT" Then
            IsStandardBracket = True
        End If
    End If
    
Exit Function
ErrorHandler:
    MsgBox Err.Description, vbInformation, METHOD
End Function
