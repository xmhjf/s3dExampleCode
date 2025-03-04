VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Selector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' For testing, 8 cases total:
' Bounding section W30x124, CP2
' Bounded sections, CP2
' - W24x55 (test t<13)
' - W24x55 (test t<=13<=20, C>=6)
' - W24x55 (test t<=13<=20, C<6)
' - W24x55 (test t>20)
' Repeat above with CP8 for both bounding and bounded

Public Function GetAllowedItems(oSmartOccurrence As IJSmartOccurrence) As Collection

    On Error GoTo ErrorHandler
    
    Set GetAllowedItems = New Collection

    ' ----------------
    ' Exit if not a PC
    ' ----------------
    If Not TypeOf oSmartOccurrence Is IJStructPhysicalConnection Then
        Exit Function
    End If
     
    Dim oSDO_PC As StructDetailObjects.PhysicalConn
    Set oSDO_PC = New StructDetailObjects.PhysicalConn
    Set oSDO_PC.object = oSmartOccurrence
     
    ' -------------------------------
    ' Exit if bounded is not a member
    ' -------------------------------
    If Not TypeOf oSDO_PC.ConnectedObject1 Is ISPSMemberPartPrismatic Then
        Exit Function
    End If
    
    Dim oSDO_Member As StructDetailObjects.MemberPart
    Set oSDO_Member = New StructDetailObjects.MemberPart
    Set oSDO_Member.object = oSDO_PC.ConnectedObject1
    
    ' ----------
    ' Get the AC
    ' ----------
    Dim sACName As String
    Dim oACObject As Object
    
    AssemblyConnection_SmartItemName oSmartOccurrence, sACName, oACObject
    
    ' ----------------
    ' Get the edge map
    ' ----------------
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    Dim sMsg As String
    
    InitMemberConnectionData oACObject, oBoundedData, oBoundingData, lStatus, sMsg

    Dim sectionAlias As Long
    Dim oEdgeMap As Collection
    Dim bPenetratesWeb As Boolean
    
    Set oEdgeMap = New Collection
    Set oEdgeMap = GetEdgeMap(oACObject, oBoundingData.AxisPort, oBoundedData.AxisPort, sectionAlias, bPenetratesWeb)

    ' -------------------------
    ' Get the bounded thickness
    ' -------------------------
    Dim dThickness As Double
    
    If bPenetratesWeb Then
        dThickness = oSDO_Member.flangeThickness
    Else
        dThickness = oSDO_Member.webThickness
    End If
    
    ' ------------------------------------------------------------------------------
    ' Determine if bounded by the top or bottom edge, relative to the bounded object
    ' ------------------------------------------------------------------------------
    ' Get the bounding port edge ID and reverse the mapping
    
    Dim oStructPort As IJStructPort
    Set oStructPort = oSDO_PC.Port2
    
    Dim realBoundingEdge As JXSEC_CODE
    Dim mappedBoundingEdge As JXSEC_CODE
    
    realBoundingEdge = oStructPort.SectionID
    mappedBoundingEdge = ReverseMap(realBoundingEdge, oEdgeMap)
    
    Dim isBottomEdge As Boolean
    isBottomEdge = True
    
    If mappedBoundingEdge = JXSEC_TOP_FLANGE_RIGHT Then
        isBottomEdge = False
    ElseIf Not mappedBoundingEdge = JXSEC_BOTTOM_FLANGE_RIGHT Then
        Exit Function ' unexpected bounding port for this table
    End If
    
    ' --------------------------------------------------
    ' Determine if the bounded bottom flange is involved
    ' --------------------------------------------------
    ' If web is involved, then set to false
    Dim IsBottomFlange As Boolean
    IsBottomFlange = False
    
    ' Get the parent of the PC
    Dim oParentName As String
    Dim oParentObj As Object
    
    Parent_SmartItemName oSmartOccurrence, oParentName, oParentObj
   
    ' Exit if parent is not a web or flange cut
    Dim featureType As StructFeatureTypes
    featureType = SF_CornerFeature
    
    If TypeOf oParentObj Is IJStructFeature Then
        Dim oStructFeature As IJStructFeature
        Set oStructFeature = oParentObj
        
        If oStructFeature.get_StructFeatureType = SF_FlangeCut Then
            GetSelectorAnswer oParentObj, "BottomFlange", IsBottomFlange
        ElseIf Not oStructFeature.get_StructFeatureType = SF_WebCut Then
            Exit Function
        End If
    End If
    
    ' ------------------------------------------------------------
    ' Get the overlap between the bounding edge and bounded flange
    ' ------------------------------------------------------------
    Dim dInsideOverlap As Double ' parameter 'c'
    GetEdgeOverlapAndClearance oACObject, isBottomEdge, IsBottomFlange, dInsideOverlap
    
    If dThickness < 0.013 Or dInsideOverlap < 0.006 Or dThickness > 0.02 Then
        GetAllowedItems.Add "ButtWeldV"
    ElseIf 0.013 <= dThickness And dThickness <= 0.02 And dInsideOverlap >= 0.006 Then
        GetAllowedItems.Add "ButtWeldY"
    End If

    Exit Function
    
ErrorHandler:
    
End Function
