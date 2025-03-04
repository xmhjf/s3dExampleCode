Attribute VB_Name = "MemberUtilities"
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
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
'*****************************************************************************
Option Explicit
Private Const MODULE = "StructDetail\Data\SmartOccurrence\MemberAssyConn\MemberUtilities"
Public Const E_FAIL = -2147467259
'
'Structure of Member Connection data
Public Type MemberConnectionData
    Matrix As IJDT4x4
    ePortId As SPSMemberAxisPortIndex
    AxisPort As ISPSSplitAxisPort
    AxisCurve As IJCurve
    MemberPart As ISPSMemberPartPrismatic
End Type

Public Const eIdealized_Unk = "Unknown"
Public Const eIdealized_Top = "Top"
Public Const eIdealized_Bottom = "Bottom"
Public Const eIdealized_WebLeft = "Web_Left"
Public Const eIdealized_WebRight = "Web_Right"
Public Const eIdealized_EndBaseFace = "End_Base"
Public Const eIdealized_EndOffsetFace = "End_Offset"
'


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
Public Sub HandleError(sModule As String, sMethod As String, Optional sExtraInfo As String = "")
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.AddFromErr Err, sExtraInfo, sMethod, sModule
    End If
    Set oEditErrors = Nothing
End Sub

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
                                    lStatus As Long, sMsg As String)
Const METHOD = "::InitMemberConnectionData"
    On Error GoTo ErrorHandler
    
    Dim iIndex As Long
    Dim lCount As Long
    
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    
    Dim oPoint As IJPoint
    Dim oPosition As IJDPosition
    Dim oElements_Ports As IJElements
    
    Dim oPort As IJPort
    Dim oPortObj As Object
    Dim oSplitAxisPort As ISPSSplitAxisPort
    
    Dim ePortId As SPSMemberAxisPortIndex
    
    sMsg = ""
    lStatus = 0
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    oAppConnection.enumPorts oElements_Ports
    lCount = oElements_Ports.Count
    
    ' for Member EndCuts, require two(2) Ports
    '   one Port will be PortId type of SPSMemberAxisAlong (Bounding Member)
    '   one Port will be PortId type of SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
    If lCount <> 2 Then
        sMsg = "Member Assembly Connection requires two(2) Ports"
'''        GoTo StatusFalse
    End If
    
    InitEndCutConnectionData oElements_Ports.Item(1), oElements_Ports.Item(2), _
                             oBoundedData, oBoundingData, lStatus, sMsg

    
    Exit Sub
    
StatusFalse:
    lStatus = 1
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    lStatus = E_FAIL
End Sub

'*************************************************************************
'Function
'InitEndCutConnectionData
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
Public Sub InitEndCutConnectionData(oBoundedObject As Object, _
                                    oBoundingObject As Object, _
                                    oBoundedData As MemberConnectionData, _
                                    oBoundingData As MemberConnectionData, _
                                    lStatus As Long, sMsg As String)
Const METHOD = "::InitEndCutConnectionData"
    On Error GoTo ErrorHandler
    
    Dim iIndex As Long
    Dim lCount As Long
    
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    
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
    
    
    sMsg = ""
    lStatus = 0
    
    ' for Member EndCuts, require two(2) Ports
    '   one Port will be PortId type of SPSMemberAxisAlong (Bounding Member)
    '   one Port will be PortId type of SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
    sMsg = "Checking Member Assembly Connection Ports"
    lCount = 2
    For iIndex = 1 To lCount
        If iIndex = 1 Then
            Set oPortObj = oBoundedObject
        Else
            Set oPortObj = oBoundingObject
        End If
        
        If oPortObj Is Nothing Then
        ElseIf TypeOf oPortObj Is ISPSSplitAxisPort Then
            Set oSplitAxisPort = oPortObj
            ePortId = oSplitAxisPort.portIndex
            
            If ePortId = SPSMemberAxisAlong Then
                If TypeOf oPortObj Is ISPSSplitAxisAlongPort Then
                    If oBoundingData.MemberPart Is Nothing Then
                        oBoundingData.ePortId = ePortId
                        Set oBoundingData.AxisPort = oSplitAxisPort
                        Set oBoundingData.MemberPart = oSplitAxisPort.Part
                    Else
                        oBoundedData.ePortId = ePortId
                        Set oBoundedData.AxisPort = oSplitAxisPort
                        Set oBoundedData.MemberPart = oSplitAxisPort.Part
                    End If
                End If
            
            ElseIf ePortId = SPSMemberAxisEnd Then
                If TypeOf oPortObj Is ISPSSplitAxisEndPort Then
                    If oBoundedData.MemberPart Is Nothing Then
                        oBoundedData.ePortId = ePortId
                        Set oBoundedData.AxisPort = oSplitAxisPort
                        Set oBoundedData.MemberPart = oSplitAxisPort.Part
                    Else
                        oBoundingData.ePortId = ePortId
                        Set oBoundingData.AxisPort = oSplitAxisPort
                        Set oBoundingData.MemberPart = oSplitAxisPort.Part
                    End If
                End If
                
            ElseIf ePortId = SPSMemberAxisStart Then
                If TypeOf oPortObj Is ISPSSplitAxisEndPort Then
                    If oBoundedData.MemberPart Is Nothing Then
                        oBoundedData.ePortId = ePortId
                        Set oBoundedData.AxisPort = oSplitAxisPort
                        Set oBoundedData.MemberPart = oSplitAxisPort.Part
                    Else
                        oBoundingData.ePortId = ePortId
                        Set oBoundingData.AxisPort = oSplitAxisPort
                        Set oBoundingData.MemberPart = oSplitAxisPort.Part
                    End If
                End If
            
            End If
            
            Set oSplitAxisPort = Nothing
        End If
        
        Set oPort = Nothing
    Next iIndex
    
    ' Verify have valid Ports from the Assembly Connection
    If oBoundedData.AxisPort Is Nothing Then
        sMsg = "Member Assembly Connection Bounded Port could not be determined"
        GoTo StatusFalse
    ElseIf oBoundingData.AxisPort Is Nothing Then
        sMsg = "Member Assembly Connection Bounding Port could not be determined"
        GoTo StatusFalse
    End If
        
    ' Verify Bounded have valid Cross Section Type
    Set oBounded_CrossSection = oBoundedData.MemberPart.CrossSection
    Set oBounding_CrossSection = oBoundingData.MemberPart.CrossSection
    
    If TypeOf oBounded_CrossSection.Definition Is IJCrossSection Then
        Set oCrossSection = oBounded_CrossSection.Definition
        sCStype = oCrossSection.Type
        
        sMsg = "Bounded CrossSection Type is not valid: " & sCStype
        If Trim(LCase(sCStype)) = LCase("CS") Then
            GoTo StatusFalse
        ElseIf Trim(LCase(sCStype)) = LCase("HSSC") Then
            GoTo StatusFalse
        ElseIf Trim(LCase(sCStype)) = LCase("PIPE") Then
            GoTo StatusFalse
        End If
        
    Else
        sMsg = "Bounded CrossSection Definition is not valid"
        GoTo StatusFalse
    End If
    
    ' Get the Bounded Port Axis curve and Point on Axis Curve (x,y,z Location)
    sMsg = "Calculating Bounded Axis/Location"
    If TypeOf oBoundedData.AxisPort Is IJPort Then
        Set oPort = oBoundedData.AxisPort
        If TypeOf oPort.Geometry Is IJPoint Then
            Set oPoint = oPort.Geometry
            oPoint.GetPoint dX, dY, dZ
                
            ' Matrix: U is direction Along Axis
            ' Matrix: V is direction normal to Web (from Web Right to Web Left)
            ' Matrix: W is direction normal to Flange (from Flange Bottom to Flange Top)
            Set oBoundedData.AxisCurve = GetAxisCurveAtPosition(dX, dY, dZ, oBoundedData.MemberPart)
            oBoundedData.MemberPart.Rotation.GetTransformAtPosition dX, dY, dZ, oBoundedData.Matrix, Nothing
        Else
            sMsg = "Member Assembly Connection Bounded Port geometry does not support IJPoint interface"
            GoTo StatusFalse
        End If
    Else
        sMsg = "Member Assembly Connection Bounded Port does not support IJPort interface"
        GoTo StatusFalse
    End If
        
    ' Get the Bounding Port Axis curve and Point on Axis Curve (x,y,z Location)
    sMsg = "Calculating Bounding Axis/Location"
    Set oPosition = GetConnectionPositionOnSupping(oBoundingData.MemberPart, oBoundedData.MemberPart, _
                                                   oBoundedData.ePortId)
    oPosition.Get dX, dY, dZ
                
    ' Matrix: U is direction Along Axis
    ' Matrix: V is direction normal to Web (from Web Right to Web Left)
    ' Matrix: W is direction normal to Flange (from Flange Bottom to Flange Top)
    Set oBoundingData.AxisCurve = GetAxisCurveAtPosition(dX, dY, dZ, oBoundingData.MemberPart)
    oBoundingData.MemberPart.Rotation.GetTransformAtPosition dX, dY, dZ, oBoundingData.Matrix, Nothing
    
    ' Verify that Bounded and Bounding Axis Curves were determined (retrieved) and are valid type
    If oBoundedData.AxisCurve Is Nothing Then
        sMsg = "Member Assembly Connection Bounded Axis Curve is Nothing"
        GoTo StatusFalse
    ElseIf oBoundingData.AxisCurve Is Nothing Then
        sMsg = "Member Assembly Connection Bounding Axis Curve is Nothing"
        GoTo StatusFalse
    ElseIf Not TypeOf oBoundedData.AxisCurve Is IJLine Then
        sMsg = "Member Assembly Connection Bounded Axis Curve does not support IJLine interface"
        GoTo StatusFalse
    ElseIf Not TypeOf oBoundingData.AxisCurve Is IJLine Then
        sMsg = "Member Assembly Connection Bounding Axis Curve does not support IJLine interface"
        GoTo StatusFalse
    End If
    
    Exit Sub
    
StatusFalse:
    lStatus = 1
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
Public Sub CheckEndToEndConnection(oBoundedData As MemberConnectionData, _
                                   oBoundingData As MemberConnectionData, _
                                   bEndToEnd As Boolean, _
                                   bColinear As Boolean, _
                                   bRightAngle As Boolean)
Const METHOD = "::CheckEndToEndConnection"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' Check if Assembly Connection is End To End Type
    bEndToEnd = IsMemberAxesEndToEnd(oBoundedData.AxisCurve, oBoundingData.AxisCurve)
    If bEndToEnd Then
        ' Assembly Connection is End To End Type
        ' Check if Axis are Colinear
            bColinear = IsMemberAxesColinear(oBoundedData.AxisCurve, oBoundingData.AxisCurve)
            If bColinear Then
                ' Assembly Connection is End To End Colinear Type
                bRightAngle = False
            Else
                ' Check if Axis are Normal to each Other
                bRightAngle = IsMemberAxesAtRightAngles(oBoundedData.AxisCurve, _
                                                        oBoundingData.AxisCurve)
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
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
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
    Dim oBounding_CrossSection As ISPSCrossSection
    
    bIdentical = False
    Set oBounded_CrossSection = oBoundedData.MemberPart.CrossSection
    Set oBounding_CrossSection = oBoundingData.MemberPart.CrossSection
    
    sData1 = oBounded_CrossSection.SectionType
    sData2 = oBounding_CrossSection.SectionType
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
           "   ...Bounded Section Type :" & oBounded_CrossSection.SectionType & _
           "   ...Name:" & oBounded_CrossSection.SectionName & _
           "   ...Standard:" & oBounded_CrossSection.SectionStandard & _
           vbCrLf & _
           "   ...Bounding Section Type:" & oBounding_CrossSection.SectionType & _
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
    Set oBounded_UpDir = New DVector
    oBounded_UpDir.Set oBoundedData.Matrix.IndexValue(8), oBoundedData.Matrix.IndexValue(9), _
                       oBoundedData.Matrix.IndexValue(10)
                       
    Set oBounding_UpDir = New DVector
    oBounding_UpDir.Set oBoundingData.Matrix.IndexValue(8), oBoundingData.Matrix.IndexValue(9), _
                        oBoundingData.Matrix.IndexValue(10)
    
    dDot_UpUp = oBounded_UpDir.Dot(oBounding_UpDir)

    ' Calculate Dot Product between the Bounded Up direction vector and Bounding AlongAxis vector
    Set oBounded_WebDir = New DVector
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
Public Function GetMemberWebFlangePoints(oMemberPart As ISPSMemberPartPrismatic, _
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
    Dim oCollectionPoints As Collection
    Dim oBounded_PortFaces As IJElements
    
    Set oCollectionPoints = New Collection
    Set oCrossSection = oMemberPart.CrossSection
    If TypeOf oCrossSection.Symbol Is IJDSymbol Then
        ' Get the Edges from the current symbol in the following order:
        ' JXSEC_WEB_LEFT : JXSEC_WEB_RIGHT : JXSEC_TOP : JXSEC_BOTTOM
        Set oSymbol = oCrossSection.Symbol
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
Public Sub EndCut_IdealizedWebFlanges(oMemberPart As ISPSMemberPartPrismatic, _
                                      sIdealizedBoundary As String, _
                                      bTopFlange As Boolean, bBottomFlange As Boolean)
Const METHOD = "::EndCut_IdealizedWebFlanges"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim bWebLeft As Boolean
    Dim bWebRight As Boolean
    
    Dim oSymbol As IJDSymbol
    Dim oSPS_CrossSection As ISPSCrossSection
    
    Dim oCollectionFlange As IJElements
    Dim oCollectionEdgeIds As Collection
    
    bTopFlange = False
    bBottomFlange = False
    Set oSPS_CrossSection = oMemberPart.CrossSection
    If Not TypeOf oSPS_CrossSection.Symbol Is IJDSymbol Then
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
    Set oSymbol = oSPS_CrossSection.Symbol
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
Public Sub CrossSection_Flanges(oMemberPart As ISPSMemberPartPrismatic, _
                                bTopFlangeLeft As Boolean, bBottomFlangeLeft As Boolean, _
                                bTopFlangeRight As Boolean, bBottomFlangeRight As Boolean)
Const METHOD = "::CrossSection_Flanges"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim sCStype As String
    
    Dim oCrossSection As IJCrossSection
    Dim oSPS_CrossSection As ISPSCrossSection
    
    bTopFlangeLeft = False
    bBottomFlangeLeft = False
    
    bTopFlangeRight = False
    bBottomFlangeRight = False
    
    Set oSPS_CrossSection = oMemberPart.CrossSection
    If TypeOf oSPS_CrossSection.Definition Is IJCrossSection Then
        oCrossSection = oSPS_CrossSection.Definition
        sCStype = oCrossSection.Type
    Else
        sCStype = ""
    End If
    
    If Trim(LCase(sCStype)) = LCase("2C") Then
        bTopFlangeLeft = True
        bTopFlangeRight = True
        bBottomFlangeLeft = True
        bBottomFlangeRight = True
    
    ElseIf Trim(LCase(sCStype)) = LCase("2L") Then
        bBottomFlangeLeft = True
        bBottomFlangeRight = True
    
    ElseIf Trim(LCase(sCStype)) = LCase("C") Then
        bTopFlangeRight = True
        bBottomFlangeRight = True
    
    ElseIf Trim(LCase(sCStype)) = LCase("C_S") Then
        bTopFlangeRight = True
        bBottomFlangeRight = True
    
    ElseIf Trim(LCase(sCStype)) = LCase("HSSC") Then
    
    ElseIf Trim(LCase(sCStype)) = LCase("HSSR") Then
    
    ElseIf Trim(LCase(sCStype)) = LCase("L") Then
    
    ElseIf Trim(LCase(sCStype)) = LCase("RECT") Then
    
    ElseIf Trim(LCase(sCStype)) = LCase("RS") Then
    
    ElseIf Trim(LCase(sCStype)) = LCase("S") Then
        bTopFlangeLeft = True
        bTopFlangeRight = True
        bBottomFlangeLeft = True
        bBottomFlangeRight = True
    
    ElseIf Trim(LCase(sCStype)) = LCase("ST") Then
        bTopFlangeLeft = True
        bTopFlangeRight = True
    
    ElseIf Trim(LCase(sCStype)) = LCase("T") Then
        bTopFlangeLeft = True
        bTopFlangeRight = True
    
    ElseIf Trim(LCase(sCStype)) = LCase("W") Then
        bTopFlangeLeft = True
        bTopFlangeRight = True
        bBottomFlangeLeft = True
        bBottomFlangeRight = True
    End If
    
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
        If TypeOf oInputObjs.Item(lCount) Is IJPort Then
            oPortCol.Add oInputObjs.Item(lCount)
        End If
    Next lCount
    
    '  make sure there are only two ports
    If oPortCol.Count = 2 Then
        Set oInputObj1 = oPortCol.Item(1)
        Set oInputObj2 = oPortCol.Item(2)
        If TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
            Set oSuppedPort = oInputObj1
            If TypeOf oInputObj2 Is ISPSSplitAxisAlongPort Then
                Set oSuppingPort = oInputObj2
                
                If oRelationObjs Is Nothing Then
                    Set oRelationObjs = New JObjectCollection
                End If
                
                oRelationObjs.Clear
                oRelationObjs.Add oSuppedPort
                oRelationObjs.Add oSuppingPort
            Else
                InputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
            End If
        Else
            InputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
        End If
    Else
        InputHelper_ValidateObjects = SPSFACInputHelper_BadNumberOfObjects
    End If
    
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
'       1. If the Bounding Member Port is and end port or the AlongAxis port
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
    Set oWebNormal = New DVector
    oWebNormal.Set oBoundedData.Matrix.IndexValue(4), _
                     oBoundedData.Matrix.IndexValue(5), oBoundedData.Matrix.IndexValue(6)
    
    ' Calculate vector from  Point to be Projected to Web Plane Root Point
    Set oProjectVec = New DVector
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
        Set oBoundingAlongAxis = New DVector
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
Public Function z_TransformPosToGlobal(oMemb As ISPSMemberPartPrismatic, _
                                       oInputPos As IJDPosition, _
                                       Optional oPosAlongAxis As IJDPosition) As IJDPosition
Const MT = "z_TransformPosToGlobal"
  On Error GoTo ErrorHandler
    Dim oProfileBO As ISPSCrossSection
    Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
    Dim oMat1 As New DT4x4, oMat2 As DT4x4
    Dim oVec As New DVector

    
    Set oProfileBO = oMemb.CrossSection
    oProfileBO.GetCardinalPointOffset oProfileBO.CardinalPoint, xOffsetCP, yOffsetCP 'Returns x and y of the current CP in RAD coordinates, which is member negative y and z.
    oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5
    xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
    yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
    xOffset = -xOffsetCP  ' x offset of current cp from cp5
    yOffset = -yOffsetCP  ' y offset of current cp from cp5
    
    oVec.Set xOffset, yOffset, 0
    oMat1.LoadIdentity
    oMat1.Translate oVec
    
    If oPosAlongAxis Is Nothing Then
        oMemb.Rotation.GetTransform oMat2
    Else
        oMemb.Rotation.GetTransformAtPosition oPosAlongAxis.x, oPosAlongAxis.y, oPosAlongAxis.z, oMat2, Nothing
    End If
    'from cross section coordinates to member coordinates
    Set oMat2 = CreateCSToMembTransform(oMat2, oMemb.Rotation.Mirror)
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
    If dBoundedSize > (dBoundedSize - dTolerance) Then
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
    
    Dim dDot As Double
    Dim dBest As Double
    Dim oBoundedAxis_Vector As IJDVector
    Dim oBoundingWeb_Vector As IJDVector
    Dim oBoundingFlange_Vector As IJDVector
    
    Dim lIdealEdgeId As Long
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    Dim oBoundedConnectable  As IJConnectable
    
    sIdealizedBoundary = eIdealized_Unk
    
    Set oBoundedConnectable = oBoundedData.MemberPart
    Set oStructProfilePart = oBoundedConnectable
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    lIdealEdgeId = oStructEndCutUtil.GetIdealizedBoundaryId(oBoundedData.AxisPort, _
                                                            oBoundingData.AxisPort)
    If lIdealEdgeId = JXSEC_TOP Then
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
    Set oBoundedAxis_Vector = New DVector
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
    
    Set oBoundingWeb_Vector = New DVector
    oBoundingWeb_Vector.Set oBoundingData.Matrix.IndexValue(4), _
                            oBoundingData.Matrix.IndexValue(5), oBoundingData.Matrix.IndexValue(6)
    
    Set oBoundingFlange_Vector = New DVector
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
    
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    
    Dim oObj As Object
    Dim oPort As IJPort
    Dim oPoint As IJPoint
    Dim oGeometry As Object
    
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
                oPoint.GetPoint dX, dY, dZ
                oRefPoint.Set dX, dY, dZ
                Exit Sub
            End If

        ElseIf TypeOf oObj Is IJPoint Then
            Set oPoint = oObj
            oPoint.GetPoint dX, dY, dZ
            oRefPoint.Set dX, dY, dZ
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
    Dim nNumPts As Long
    
    Dim dPntX As Double
    Dim dPntY As Double
    Dim dPntZ As Double
    Dim dSrcX As Double
    Dim dSrcY As Double
    Dim dSrcZ As Double
    Dim dPnts1() As Double
    Dim dPnts2() As Double
    Dim dPars1() As Double
    Dim dPars2() As Double

    Dim dDist As Double
    Dim dMinDist As Double
    
    Dim oGeometry As Object
    Dim oObjectReplacing As Object
    
    Dim oPort As IJPort
    Dim oPoint As IJPoint
    Dim oCurve As IJCurve
    Dim oPlane As IJPlane
    Dim oVector As IJDVector
    Dim oSurface As IJSurface
    Dim oRefPoint As IJPoint
    Dim oPosition As IJDPosition
    
    Dim oGeom3DFactory As GeometryFactory
    Dim oMemberFactory As SPSMemberFactory
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    Dim oFeatureServices As ISPSMemberFeatureServices

    sMsg = "Createing 3D Reference Point from given Position"
    Set oGeom3DFactory = New GeometryFactory
    Set oRefPoint = oGeom3DFactory.Points3d.CreateByPoint(Nothing, _
                                                          oLocation.x, _
                                                          oLocation.y, _
                                                          oLocation.z)

    iIndex = 0
    dMinDist = -1#
    sMsg = "Looping Thru the Collection of replacing Objects"
    For Each oObjectReplacing In oObjectCollectionReplacing
        iIndex = iIndex + 1
        sMsg = "Setting Geometry Type of Replacing Object:" & iIndex
        If TypeOf oObjectReplacing Is IJPort Then
            Set oPort = oObjectReplacing
            Set oGeometry = oPort.Geometry
        Else
            Set oGeometry = oObjectReplacing
        End If
        

        If TypeOf oGeometry Is IJModelBody Then
            sMsg = "IJModelBody Geometry Type of Replacing Object:" & iIndex
            Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
            oTopologyLocate.FindProjectedPointOnPart oGeometry, oLocation, _
                                                     oVector, oPosition
            dDist = oLocation.DistPt(oPosition)
            
        ElseIf TypeOf oGeometry Is IJPoint Then
            sMsg = "IJPoint Geometry Type of Replacing Object:" & iIndex
            Set oPoint = oGeometry
            dDist = oPoint.DistFromPt(oRefPoint)
        
        ElseIf TypeOf oGeometry Is IJCurve Then
            sMsg = "IJCurve Geometry Type of Replacing Object:" & iIndex
            Set oCurve = oGeometry
            oCurve.DistanceBetween oRefPoint, dDist, _
                                   dSrcX, dSrcY, dSrcZ, dPntX, dPntY, dPntZ

        ElseIf TypeOf oGeometry Is IJSurface Then
            sMsg = "IJSurface Geometry Type of Replacing Object:" & iIndex
            Set oSurface = oGeometry
            oSurface.DistanceBetween oRefPoint, dDist, _
                                     dSrcX, dSrcY, dSrcZ, dPntX, dPntY, dPntZ, _
                                     nNumPts, dPnts1, dPnts2, dPars1, dPars2
        
        ElseIf TypeOf oGeometry Is IJPlane Then
            sMsg = "IJPlane Geometry Type of Replacing Object:" & iIndex
            Set oMemberFactory = New SPSMemberFactory
            Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices
            oFeatureServices.GetMinDistFromPlane oGeometry, oRefPoint, dDist
            Set oMemberFactory = Nothing
            
        Else
            ' unknown geometry.  perhaps a compound surface from intelliship.
            sMsg = "Unknown Geometry Type of Replacing Object:" & iIndex
            Err.Raise E_FAIL
        End If
        
        sMsg = "UChecking if Replacing Object is closest:" & iIndex & _
               "  ... Dist:" & dDist
        If dMinDist < -0.1 Then
            dMinDist = dDist
            Set oReplacingObject = oObjectReplacing
        ElseIf dDist < dMinDist Then
            dMinDist = dDist
            Set oReplacingObject = oObjectReplacing
        End If

    Next

    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub

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
    Dim nCount As Long
    Dim lStatus As Long
    
    Dim bIsDeleted As Boolean
    Dim bIsMigrated As Boolean
    
    Dim eEndCutType As eEndCutTypes
    
    Dim oEndCutObject As Object
    Dim oAppConnPort1 As Object
    Dim oAppConnPort2 As Object
    Dim oBoundingEndPort As Object
    
    Dim oNewItemPort As Object
    Dim oNewConnPort1 As Object
    Dim oNewConnPort2 As Object
    Dim oNewBoundingEndPort As Object
    
    Dim oPoint As IJDPosition
    Dim oAppConn As IJAppConnection
    Dim oAppConnPorts As IJElements

    Dim oMemberItems As IJElements
    Dim oMemberObjects As IJDMemberObjects
    
    Dim oBoundedPart As Object
    Dim oBoundingPart As Object
    Dim oBounded_SplitAxisPort As ISPSSplitAxisPort
    Dim oBounding_SplitAxisPort As ISPSSplitAxisPort

    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oNewBoundedData As MemberConnectionData
    Dim oNewBoundingData As MemberConnectionData

    'Get the input ports.
    'If an input is on the RefColl, we assume it is not split.
    sMsg = "Retreiving Ports from Assembly Connection"
    bIsMigrated = False
    Set oAppConn = pAggregatorDescription.CAO
    InitMemberConnectionData oAppConn, oBoundedData, oBoundingData, lStatus, sMsg

    'Get the end position of the Bounded Part using the ports
    Call oAppConn.enumPorts(oAppConnPorts)
    GetPositionFromElementsList oAppConnPorts, oPoint
        
    nCount = oAppConnPorts.Count
    If nCount < 2 Then
        sMsg = "invalid Number of Ports from Assembly Connection:" & nCount
        GoTo ErrorHandler
    End If
    
    Set oNewConnPort1 = Nothing
    Set oAppConnPort1 = oAppConnPorts.Item(1)
    sMsg = "Checking if Port1 from Assembly Connection has been replaced"
    pMigrateHelper.ObjectsReplacing oAppConnPort1, oObjectCollectionReplacing, bIsDeleted
    If Not oObjectCollectionReplacing Is Nothing Then
        SelectReplacingObject oObjectCollectionReplacing, oPoint, oNewConnPort1
        If Not oNewConnPort1 Is Nothing Then
            bIsMigrated = True
        End If
        Set oObjectCollectionReplacing = Nothing
    End If
    
    Set oNewConnPort2 = Nothing
    Set oAppConnPort2 = oAppConnPorts.Item(2)
    sMsg = "Checking if Port2 from Assembly Connection has been replaced"
    pMigrateHelper.ObjectsReplacing oAppConnPort2, oObjectCollectionReplacing, bIsDeleted
    If Not oObjectCollectionReplacing Is Nothing Then
        SelectReplacingObject oObjectCollectionReplacing, oPoint, oNewConnPort2
        If Not oNewConnPort2 Is Nothing Then
            bIsMigrated = True
        End If
        Set oObjectCollectionReplacing = Nothing
    End If
        
    sMsg = "Checking if Assembly Connection Member Items are to be Migrated"
    If bIsMigrated Then
        If Not oNewConnPort1 Is Nothing Then
            sMsg = "Assembly Connection Port1 is Migrated"
            oAppConn.removePort oAppConnPort1
            oAppConn.addPort oNewConnPort1
        End If
        
        If Not oNewConnPort2 Is Nothing Then
            sMsg = "Assembly Connection Port2 is Migrated"
            oAppConn.removePort oAppConnPort2
            oAppConn.addPort oNewConnPort2
        End If
    
        ' Update the Bounded/Bounding data based on the Migrated Ports
        InitMemberConnectionData oAppConn, _
                                 oNewBoundedData, oNewBoundingData, lStatus, sMsg
    Else
        Exit Sub
    End If
    
    ' Verify Assembly Connections have valid Member Items that also require to be Migrated
    If Not TypeOf oAppConn Is IJDMemberObjects Then
        Exit Sub
    End If
    
    Set oMemberObjects = oAppConn
    If oMemberObjects.Item(1) Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oMemberObjects.Item(1) Is IJElements Then
        Exit Sub
    End If
        
    Set oMemberItems = oMemberObjects.Item(1)
    If oMemberItems.Count < 1 Then
        Exit Sub
    End If
    
    sMsg = "Assembly Connection Member Items are to be Migrated:" & oMemberItems.Count
    For iIndex = 1 To oMemberItems.Count
        bIsMigrated = False
        Set oEndCutObject = oMemberItems.Item(iIndex)
        If TypeOf oEndCutObject Is IJStructFeature Then
            EndCut_InputData oEndCutObject, oBounded_SplitAxisPort, oBoundedPart, _
                             oBounding_SplitAxisPort, oBoundingPart, eEndCutType
                             
            ' Check if the Bounding Port is SPSMemberAxisAlong type
            ' (Assembly Connections are always between an End and SPSMemberAxisAlong)
            If oBounding_SplitAxisPort.portIndex <> SPSMemberAxisAlong Then
                ' Member Item is End to End case
                ' Need to Retreive the Bounding End Port that was actually Used
                ' Need to Retreive the Bounding End Port that will be actually Used
                GetSupportingEndPort oBoundedData, oBoundingData, oBoundingEndPort
                GetSupportingEndPort oNewBoundedData, oNewBoundingData, _
                                     oNewBoundingEndPort
            Else
                Set oBoundingEndPort = Nothing
            End If
            
            sMsg = "Checking if EndCut Bounded Port has been replaced: " & Trim(Str(iIndex))
            If oBounded_SplitAxisPort Is oAppConnPort1 Then
                If Not oNewConnPort1 Is Nothing Then
                    bIsMigrated = True
                    Set oBounded_SplitAxisPort = oNewConnPort1
                End If
                
            ElseIf oBounded_SplitAxisPort Is oAppConnPort2 Then
                If Not oNewConnPort2 Is Nothing Then
                    bIsMigrated = True
                    Set oBounded_SplitAxisPort = oNewConnPort2
                End If
                
            ElseIf oBounded_SplitAxisPort Is oBoundingEndPort Then
                If Not oNewBoundingEndPort Is Nothing Then
                    bIsMigrated = True
                    Set oBounded_SplitAxisPort = oNewBoundingEndPort
                End If
            End If
            
            sMsg = "Checking if EndCut Bounding Port has been replaced: " & Trim(Str(iIndex))
            If oBounding_SplitAxisPort Is oAppConnPort1 Then
                If Not oNewConnPort1 Is Nothing Then
                    bIsMigrated = True
                    Set oBounding_SplitAxisPort = oNewConnPort1
                End If
            
            ElseIf oBounding_SplitAxisPort Is oAppConnPort2 Then
                If Not oNewConnPort2 Is Nothing Then
                    bIsMigrated = True
                    Set oBounding_SplitAxisPort = oNewConnPort2
                End If
            
            ElseIf oBounding_SplitAxisPort Is oBoundingEndPort Then
                If Not oNewBoundingEndPort Is Nothing Then
                    bIsMigrated = True
                    Set oBounding_SplitAxisPort = oNewBoundingEndPort
                End If
            End If
            
            If bIsMigrated Then
                sMsg = "Set EndCut Bounded/Bounding Ports with Migrated Ports: " & Trim(Str(iIndex))
                MigrateModifyEndCutData oEndCutObject, _
                                        oBounded_SplitAxisPort, oBounding_SplitAxisPort
                                        
                sMsg = "Setting EndCut as Migrated: " & Trim(Str(iIndex))
                Set oObjectCollectionReplaced = New JObjectCollection
                Set oObjectCollectionReplacing = New JObjectCollection
            
                oObjectCollectionReplaced.Add oEndCutObject
                oObjectCollectionReplacing.Add oEndCutObject
                    
                pMigrateHelper.ObjectsReplaced oObjectCollectionReplaced, _
                                               oObjectCollectionReplacing, False
                Set oObjectCollectionReplaced = Nothing
                Set oObjectCollectionReplacing = Nothing
            End If
            
            
        End If
    Next iIndex

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
    Dim jIndex As Long
    
    Dim lXId As Long
    Dim lCtxId As Long
    Dim lOptId As Long
    Dim lOprId As Long
    Dim lPortType As Long
        
    Dim bIsMigrated As Boolean

    Dim oPortOld As Object
    Dim oPortNew As Object
    
    Dim oPortObj1 As Object
    Dim oPortObj2 As Object
    
    Dim oEndCutObj As Object
    Dim oMemberPart As Object
    Dim oBoundedPart As Object
    Dim oBoundingPart As Object
    
    Dim oPort As IJPort
    Dim oNewPort As IJPort
    Dim oSP3D_StructPort As SP3DStructPorts.IJStructPort
    Dim oStructGeomBasicPort As StructGeomBasicPort
    
    Dim oMemberItems As IJElements
    Dim oMemberObjects As IJDMemberObjects
    
    Dim eEndCutType As eEndCutTypes
    Dim eBoundedSubPort As JXSEC_CODE
    
    Dim oPortMoniker As IUnknown
    Dim oSplitAxisPort As ISPSSplitAxisPort
    Dim oBounded_SplitAxisPort As ISPSSplitAxisPort
    Dim oBounding_SplitAxisPort As ISPSSplitAxisPort
    
    Dim oPhysConnection As IJStructPhysicalConnection
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    Dim oObjectCollectionReplaced As IJDObjectCollection
    Dim oObjectCollectionReplacing As IJDObjectCollection
    
    Dim oPortHelper As PORTHELPERLib.PortHelper
    Dim oSDO_PhysicalConn As StructDetailObjects.PhysicalConn
    Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition

    sMsg = "Verify EndCut have MemberItems to be Migrated"
    Set oEndCutObj = oAggregatorDescription.CAO
    If Not TypeOf oEndCutObj Is IJDMemberObjects Then
        Exit Sub
    End If
    
    Set oMemberObjects = oEndCutObj
    If oMemberObjects.Item(1) Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oMemberObjects.Item(1) Is IJElements Then
        Exit Sub
    End If
        
    Set oMemberItems = oMemberObjects.Item(1)
    If oMemberItems.Count < 1 Then
        Exit Sub
    End If
    
    sMsg = "Retreive the EndCut Bounded/Bounding Ports/Parts/EndCutType Data"
    If TypeOf oEndCutObj Is IJStructFeature Then
        EndCut_InputData oEndCutObj, oBounded_SplitAxisPort, oBoundedPart, _
                         oBounding_SplitAxisPort, oBoundingPart, eEndCutType
    Else
        Exit Sub
    End If
    
    ' Initialize Utility to Modify Physical Connections
    Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
    
    ' Since we are using Multiple (nCount) of the same Member Item
    ' The MemberObjects.item(1) is a List (Collection) of Objects being created
    ' For Each member Object, expect it to be a Physical Connection
    ' Need to Migrate the Ports for each Physical Connection
    sMsg = "EndCut Member Items are to be Migrated:" & oMemberItems.Count
    For iIndex = 1 To oMemberItems.Count
        bIsMigrated = False
        If TypeOf oMemberItems.Item(iIndex) Is IJStructPhysicalConnection Then
            Set oPhysConnection = oMemberItems.Item(iIndex)
            Set oSDO_PhysicalConn = New StructDetailObjects.PhysicalConn
            Set oSDO_PhysicalConn.object = oPhysConnection
            
            Set oPortObj1 = oSDO_PhysicalConn.Port1
            Set oPortObj2 = oSDO_PhysicalConn.Port2
            
            ' Loop thru both Physical Connection Ports
            For jIndex = 1 To 2
                If jIndex = 1 Then
                    Set oPortOld = oPortObj1
                Else
                    Set oPortOld = oPortObj2
                End If
                
                ' Check if Physical Connection Port is to be replaced
                Set oPort = oPortOld
                Set oMemberPart = oPort.Connectable
                If oMemberPart Is oBoundedPart Then
                    ' Physical Connection Port has not been replaced
                ElseIf oMemberPart Is oBoundingPart Then
                    ' Physical Connection Port has not been replaced
                Else
                    ' Physical Connection Port1 has BEEN replaced
                    sMsg = "Set Physical Connection Port with Migrated Port: " & Trim(Str(iIndex))
                    bIsMigrated = True
                    
                    ' Note:
                    ' Need to access IJStructPort Interface thru the StructGeomBasicPort Interface
                    Set oStructGeomBasicPort = oPortOld
                    Set oSP3D_StructPort = oStructGeomBasicPort
                    Set oPortMoniker = oSP3D_StructPort.PortMoniker
                    
                    Set oPortHelper = New PORTHELPERLib.PortHelper
                    oPortHelper.DecodeTopologyProxyMoniker oPortMoniker, _
                                                           lPortType, lCtxId, lOptId, lOprId, lXId
                    If lXId > 0 Then
                        ' Assuming the Port with Xid > 0 is associcated with the Bounded Part
                        sMsg = "GetLatePortForFeatureSegment..XId: " & Trim(Str(lXId))
                        Set oStructProfilePart = oBoundedPart
                        Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
                        oStructEndCutUtil.GetLatePortForFeatureSegment oEndCutObj, lXId, oPortNew
                    Else
                        ' Assuming the Port with Xid < 1 is associcated with the Bounding Part
                        sMsg = "Member_GetSolidPort..oBounding_SplitAxisPort: "
                        Set oPortNew = Member_GetSolidPort(oBounding_SplitAxisPort)
                    End If
                    
                    ' Update the Physical Connection Port
                    oConnectionDefinition.ReplacePhysicalConnectionPort oPhysConnection, _
                                                                        oPortOld, oPortNew
                    
                    ' Flag the Physical Connection Port as Migrated
                    Set oObjectCollectionReplaced = New JObjectCollection
                    Set oObjectCollectionReplacing = New JObjectCollection
                
                    oObjectCollectionReplaced.Add oPortOld
                    oObjectCollectionReplacing.Add oPortNew
                        
                    oMigrateHelper.ObjectsReplaced oObjectCollectionReplaced, _
                                                   oObjectCollectionReplacing, False
                    
                    Set oObjectCollectionReplaced = Nothing
                    Set oObjectCollectionReplacing = Nothing
                
                End If
            Next jIndex
            
            ' If Member Item (Physical Connection) has been Migrated (Updated)
            If bIsMigrated Then
                sMsg = "Set Physical Connection  as Migrated: " & Trim(Str(iIndex))
                Set oObjectCollectionReplaced = New JObjectCollection
                Set oObjectCollectionReplacing = New JObjectCollection
            
                oObjectCollectionReplaced.Add oPhysConnection
                oObjectCollectionReplacing.Add oPhysConnection
                    
                oMigrateHelper.ObjectsReplaced oObjectCollectionReplaced, _
                                               oObjectCollectionReplacing, False
                
                Set oObjectCollectionReplaced = Nothing
                Set oObjectCollectionReplacing = Nothing
            End If
        
        End If
    Next iIndex

    Exit Sub

ErrorHandler:
MsgBox "MigrateEndCutObject::ErrorHandler"
    HandleError MODULE, METHOD, sMsg
End Sub

'*************************************************************************
'Function
'MigrateModifyEndCutData
'
'Abstract
'   Given the EndCut Smart Occurrence object
'   Modify the Bounded and Bounding Ports used by the EndCut to the Replaced (Migrated) Ports
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub MigrateModifyEndCutData(oEndCutObject As Object, _
                                   oBoundedPort As Object, _
                                   oBoundingPort As Object)
Const METHOD = "::MigrateModifyEndCutData"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim eStructFeature As StructFeatureTypes
    
    Dim oModEndCut As Object
    Dim oParentObject As Object
    Dim oWebcutObject As Object
    
    Dim oStructfeature As IJStructFeature
    Dim oResourceManager As IUnknown
    
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDO_Webcut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
    
    If TypeOf oEndCutObject Is IJStructFeature Then
        Set oStructfeature = oEndCutObject
        eStructFeature = oStructfeature.get_StructFeatureType
        
        Set oSDO_Helper = New StructDetailObjects.Helper
        Set oResourceManager = oSDO_Helper.GetResourceManagerFromObject(oEndCutObject)
        
        Set oFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
        If eStructFeature = SF_WebCut Then
            sMsg = "Replacing EndCut Inputs ... WebCut"
            Set oModEndCut = oFeatureUtils.CreateWebCut(oResourceManager, _
                                                        oBoundingPort, _
                                                        oBoundedPort, _
                                                        vbNullString, _
                                                        oParentObject, _
                                                        oEndCutObject)
        ElseIf eStructFeature = SF_FlangeCut Then
            sMsg = "Replacing EndCut Inputs ... FlangeCut"
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oEndCutObject
            Set oWebcutObject = oSDO_FlangeCut.WebCut
            Set oModEndCut = oFeatureUtils.CreateFlangeCut(oResourceManager, _
                                                           oBoundingPort, _
                                                           oBoundedPort, _
                                                           oWebcutObject, _
                                                           vbNullString, _
                                                           oParentObject, _
                                                           oEndCutObject)
        
        Else
            sMsg = "Given EndCut Object StructFeature Type is not known"
            GoTo ErrorHandler
        End If
    
    Else
        sMsg = "Given EndCut Object is NOT IJStructFeature"
        GoTo ErrorHandler
    End If
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Sub
