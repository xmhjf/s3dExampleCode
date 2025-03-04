Attribute VB_Name = "ConvertHelper"
Option Explicit
Option Private Module
' ******************************************************************
'  Copyright (c) 2010, Intergraph Corporation. All rights reserved.
'
'  File               ConvertHelper.bas
'  ProgID             SPSHandrailMacros.ConvertHelper
'  Author             Black Pearl
'  Creation Date      [March 04, 2010]
'  Description        Creates frame connections between all of the
'                     components of a "dropped" handrail
' ******************************************************************

Private m_TopRails As IJElements
Private m_MidRails As IJElements
Private m_ToePlates As IJElements
Private m_Posts As IJElements
Private m_BeginTreatment As Object
Private m_EndTreatment As Object
Private m_MaterialProxy As IJDProxy
Private m_PlateDimensionsProxy As IJDProxy

Private m_oMemberFactory As SPSMembers.SPSMemberFactory
Private m_oErrors As IJEditErrors
Private Const ASCONNECPROGID = "StructConnections.StructAssemblyConnection"
Private Const strVisiblePorts = "VisiblePorts"
Private Const strStablePorts = "StablePorts"
Private Const strGraphics = "Graphics"

' find the first Element within 1 mm of Object.
Private Sub FindFirstMemberNearElement(oObject As Object, Elements As IJElements, ByRef MemberSys As ISPSMemberSystem)

    Const METHOD = "FindFirstMemberNearElement"
    On Error GoTo ErrHandler

    Dim iCurve As IJCurve
    Dim dDist As Double
    Dim dPostX  As Double, dPostY  As Double, dPostZ As Double
    Dim dRailX As Double, dRailY As Double, dRailZ As Double
    
    Set MemberSys = Nothing

    For Each iCurve In Elements
        iCurve.DistanceBetween oObject, dDist, dPostX, dPostY, dPostZ, dRailX, dRailY, dRailZ
        If dDist < 0.001 Then
            Set MemberSys = iCurve
            GoTo wrapup
        End If
    Next iCurve

wrapup:
    Exit Sub

ErrHandler:
    m_oErrors.Add Err.Number, "FindFirstMemberNearElement", Err.Description
    Exit Sub
End Sub

' Attach posts' to rails with AxisAlong or Split-None, or railFC's to posts.
' This function assumes that every rail's startFC is already attached if it is supposed to be.
' For linear rails, attach their endFC to the post even if the rail extends past the post, as for end-treatment or no-post-at-turns
' For curved rails, attach their endFC to the post if it is pointOn to the post, else attach the post to the rail.
' This behavior satisifes TR 256683 for linear rails because it leaves all posts that control the rails height with an unsupported top end.

Public Function ConnectPostsAndRails(PostElements As IJElements, RailElements As IJElements, _
                            bIsTopRail As Boolean, bCurvedRailExists As Boolean) As StructHandrailConvertHelperStatus
    Const METHOD = "ConnectPostsAndRails"
    On Error GoTo ErrHandler

    Dim eStatus As StructHandrailConvertHelperStatus
    eStatus = StructHandrailConvertHelperStatus_Unexpected

    Dim lPostCount As Long, lPostNumber As Long
    Dim NextRailMemberSys As ISPSMemberSystem
    Dim postMemberSys As ISPSMemberSystem
    Dim RailMemberSys As ISPSMemberSystem
    Dim RailFC As ISPSFrameConnection
    Dim TopEnd As SPSMemberAxisPortIndex
    
    Dim iLine As IJLine
    Dim iCurve As IJCurve
    Dim dist As Double
    Dim sX As Double, sY As Double, sZ As Double, eX As Double, eY As Double, eZ As Double
    Dim iPoint As IJPoint
    
    lPostCount = PostElements.Count
    If lPostCount < 1 Then
        eStatus = StructHandrailConvertHelperStatus_NoPosts
        GoTo ExitConnect
    End If
    
    Set iLine = New Line3d
    
    For lPostNumber = 1 To lPostCount
    
        Set postMemberSys = PostElements(lPostNumber)

        If lPostNumber = 1 Then     ' are posts pointed up or down ?
            postMemberSys.FrameConnectionAtEnd(SPSMemberAxisStart).Joint.Point.GetPoint sX, sY, sZ
            postMemberSys.FrameConnectionAtEnd(SPSMemberAxisEnd).Joint.Point.GetPoint eX, eY, eZ
            If eZ > sZ Then
                TopEnd = SPSMemberAxisEnd
            Else
                TopEnd = SPSMemberAxisStart
            End If
        End If
        
        FindFirstMemberNearElement postMemberSys, RailElements, RailMemberSys
        
        If RailMemberSys Is Nothing Then      ' should never happen
            GoTo NextPost
        End If
        
        ' if the startFC is not set, then this is a new rail.   attach it to the post.
        Set RailFC = RailMemberSys.FrameConnectionAtEnd(SPSMemberAxisStart)
        If Not RailFC.definition Is Nothing Then
            Set RailFC = RailMemberSys.FrameConnectionAtEnd(SPSMemberAxisEnd)
        
            If Not RailFC.definition Is Nothing Then    'end FC already set, probably to an endtreatment
                Set RailFC = Nothing
            ElseIf lPostNumber < lPostCount Then           ' not the last post.  is it last for this Rail ?
                FindFirstMemberNearElement PostElements(lPostNumber + 1), RailElements, NextRailMemberSys
                If NextRailMemberSys Is RailMemberSys Then    ' same rail at next post.  this is a mid-post
                    Set RailFC = Nothing
                End If
            End If
        End If
                    
        If RailFC Is Nothing Then           ' this is a mid-post to rail connection
        
            If bIsTopRail Then  ' attach the post to the rail
                eStatus = CreateConnection(postMemberSys.FrameConnectionAtEnd(TopEnd), "Axis-Along", RailMemberSys)
                ' this is where the post positioning can be updated
                ' SetPositionRule postMemberSys.FrameConnectionAtEnd(TopEnd), SPSPODimension_EndSelection_Start, SPSPODimension_Position_Distance
            Else
                eStatus = CreateConnection(postMemberSys, "Split-None", RailMemberSys)
            End If

        ' attach the end FC of the rail to the post
        Else

            If bCurvedRailExists Or TypeOf RailMemberSys Is ISPSMemberSystemCurve Then
                
                ' is the curved rail pointOn to  the post ?
                Set iPoint = RailFC
                iPoint.GetPoint sX, sY, sZ
                iLine.DefineBy2Points sX, sY, sZ, sX, sY, sZ
                Set iCurve = postMemberSys
                iCurve.DistanceBetween iLine, dist, sX, sY, sZ, eX, eY, eZ
                
                'If curved rail is very close to the post, then attach the rail to the post.   Else,
                'leave the end of the curved member unsupported but do connect this post to the rail.
                If bCurvedRailExists = False And dist < 0.001 Then
                     eStatus = CreateConnection(RailFC, "Axis-Along", postMemberSys)
                ElseIf bIsTopRail Then      ' else attach the post to the rail if it a TopRail
                     eStatus = CreateConnection(postMemberSys.FrameConnectionAtEnd(TopEnd), "Axis-Along", RailMemberSys)
                Else                        ' else conect this midrail to the post with SplitNone
                    eStatus = CreateConnection(postMemberSys, "Split-None", RailMemberSys)
                End If
                
            Else    ' attach the linear rail to the post, even if the rail extends past the post.
                 eStatus = CreateConnection(RailFC, "Axis-Along", postMemberSys)
                ' this is where the post positioning can be updated
                 SetPositionRule RailFC, SPSPODimension_EndSelection_Start, SPSPODimension_Position_Distance
             
                ' for the sake of FC offsets, represent the Rail FC as a curve
                Set iPoint = RailFC
                iPoint.GetPoint sX, sY, sZ
                iLine.DefineBy2Points sX, sY, sZ, sX, sY, sZ
                SetFCOffsetProperties RailFC, iLine, postMemberSys
            End If
        End If

NextPost:
    Next lPostNumber

ExitConnect:
    ConnectPostsAndRails = eStatus
    Exit Function

ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    ConnectPostsAndRails = StructHandrailConvertHelperStatus_Unexpected

End Function

' connect EndTreatment to top rail and mid rail using Axis-End FC
Public Function ConnectTreatmentToRails(Treatment As Object, TopRail As Object, MidRail As Object) As StructHandrailConvertHelperStatus
Const METHOD = "ConnectTreatmentToRails"
On Error GoTo ErrHandler
    Dim eStatus As StructHandrailConvertHelperStatus
    eStatus = StructHandrailConvertHelperStatus_Unexpected
    Dim TreatmentMemberSys As ISPSMemberSystem
    Dim TopEnd As SPSMemberAxisPortIndex
    Dim BottomEnd As SPSMemberAxisPortIndex
    Dim dStartFCx As Double, dStartFCy As Double, dStartFCz As Double
    Dim dEndFCx As Double, dEndFCy As Double, dEndFCz As Double
    
    Dim bPointOn As Boolean
    Dim iCurve As IJCurve
    
    If Treatment Is Nothing Then
        eStatus = StructHandrailConvertHelperStatus_Ok
        Exit Function
    End If

    Set TreatmentMemberSys = Treatment
    
    TreatmentMemberSys.FrameConnectionAtEnd(SPSMemberAxisStart).Joint.Point.GetPoint dStartFCx, dStartFCy, dStartFCz
    TreatmentMemberSys.FrameConnectionAtEnd(SPSMemberAxisEnd).Joint.Point.GetPoint dEndFCx, dEndFCy, dEndFCz
    If dEndFCz > dStartFCz Then
        TopEnd = SPSMemberAxisEnd
        BottomEnd = SPSMemberAxisStart
    Else
        TopEnd = SPSMemberAxisStart
        BottomEnd = SPSMemberAxisEnd
    End If
        
    If Not TopRail Is Nothing Then
        If TypeOf TreatmentMemberSys Is ISPSMemberSystemCurve Then      ' check whether ends meet
            Set iCurve = TopRail
            If TopEnd = SPSMemberAxisStart Then
                bPointOn = iCurve.IsPointOn(dStartFCx, dStartFCy, dStartFCz)
            Else
                bPointOn = iCurve.IsPointOn(dEndFCx, dEndFCy, dEndFCz)
            End If
        Else
            bPointOn = True     ' linear members can introduce an offset
        End If
        
        If bPointOn Then
            eStatus = CreateConnection(TreatmentMemberSys.FrameConnectionAtEnd(TopEnd), "Axis-End", TopRail)
        Else
            eStatus = StructHandrailConvertHelperStatus_Ok
        End If
    End If

    If Not MidRail Is Nothing Then
        If TypeOf TreatmentMemberSys Is ISPSMemberSystemCurve Then      ' check whether ends meet
            Set iCurve = MidRail
            If BottomEnd = SPSMemberAxisStart Then
                bPointOn = iCurve.IsPointOn(dStartFCx, dStartFCy, dStartFCz)
            Else
                bPointOn = iCurve.IsPointOn(dEndFCx, dEndFCy, dEndFCz)
            End If
        Else
            bPointOn = True
        End If

        If bPointOn Then
            eStatus = CreateConnection(TreatmentMemberSys.FrameConnectionAtEnd(BottomEnd), "Axis-End", MidRail)
        Else
            eStatus = StructHandrailConvertHelperStatus_Ok
        End If
    End If

    ConnectTreatmentToRails = eStatus
    Exit Function

ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    ConnectTreatmentToRails = StructHandrailConvertHelperStatus_Unexpected

End Function
' connect Rail start to a Post, or to previous Rail end if no post existed at the start.
Public Function ConnectRails(PostElements As IJElements, RailElements As IJElements, bCurvedRailExists As Boolean) As SPSHandrails.StructHandrailConvertHelperStatus

    Const METHOD = "ConnectRails"
    On Error GoTo ErrHandler

    Dim eStatus As StructHandrailConvertHelperStatus
    eStatus = StructHandrailConvertHelperStatus_Unexpected
    Dim lRailCount As Long, lRailNumber As Long
    
    Dim prevRail As ISPSMemberSystem, currRail As ISPSMemberSystem
    Dim postMemberSys As ISPSMemberSystem
    Dim RailFC As ISPSFrameConnection
    Dim prevRailConnectedOnEnd As Boolean

    lRailCount = RailElements.Count
    prevRailConnectedOnEnd = False

    If lRailCount > 1 Then

        For lRailNumber = 2 To lRailCount
        
            Set currRail = RailElements.Item(lRailNumber)
            Set RailFC = currRail.FrameConnectionAtEnd(SPSMemberAxisStart)
            
            Set postMemberSys = Nothing
            
            If Not bCurvedRailExists Then
                ' look for a Post near this FC
                FindFirstMemberNearElement RailFC, PostElements, postMemberSys
            End If
            
            If postMemberSys Is Nothing Then
            ' no post was found, so connect the Rail to the end of the previous one.
            
                Set prevRail = RailElements.Item(lRailNumber - 1)
                If prevRail.FrameConnectionAtEnd(SPSMemberAxisEnd).definition Is Nothing Then
                    eStatus = CreateConnection(RailFC, "Axis-End", prevRail.FrameConnectionAtEnd(SPSMemberAxisEnd))
                End If
            
                ' If this is a very short rail with no post along its length, attach end FC to the next rail.
                ' This assumes that the next rail will have a start FC that can attach to a post.
                If lRailNumber < lRailCount Then
                    FindFirstMemberNearElement currRail, PostElements, postMemberSys
                    If postMemberSys Is Nothing Then
                        Dim iAlignment As ISPSMemberSystemAlignment
                        Set RailFC = currRail.FrameConnectionAtEnd(SPSMemberAxisEnd)
                        eStatus = CreateConnection(RailFC, "Axis-End", RailElements.Item(lRailNumber + 1))
                        If prevRailConnectedOnEnd Then
                            Set iAlignment = currRail
                            iAlignment.Align = False
                        End If
                        prevRailConnectedOnEnd = True
                    Else
                        prevRailConnectedOnEnd = False
                    End If
                End If
                    
            Else
                'found a Post, so connect the Rail start to the post and set positionRule
                
                eStatus = CreateConnection(RailFC, "Axis-Along", postMemberSys)
                SetPositionRule RailFC, SPSPODimension_EndSelection_Start, SPSPODimension_Position_Distance
            
            End If
        Next lRailNumber
    End If

    eStatus = StructHandrailConvertHelperStatus_Ok
    ConnectRails = eStatus
    Exit Function

ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    ConnectRails = StructHandrailConvertHelperStatus_Unexpected
End Function

Private Sub SeparateRails(Rails As IJElements, RailSLO As IJStructListOwner, ByRef NRailRows As Long)

    Dim dPrevStartX As Double, dPrevStartY As Double, dPrevStartZ As Double
    Dim dPrevEndX As Double, dPrevEndY As Double, dPrevEndZ As Double
    Dim dCurrStartX As Double, dCurrStartY As Double, dCurrStartZ As Double
    Dim dCurrEndX As Double, dCurrEndY As Double, dCurrEndZ As Double
    Dim iCurve As IJCurve
    Dim Row As Long, Index As Long

    ' take the list of rails, assuming they are created in a sequential order along the path,
    ' create separate lists wherever the start of one does not match the end of the previous one

    NRailRows = 0
    If Rails Is Nothing Then
        Exit Sub
    End If
    If Rails.Count = 0 Then
        Exit Sub
    End If

    NRailRows = 1
    Set iCurve = Rails.Item(1)
    iCurve.EndPoints dPrevEndX, dPrevEndY, dPrevEndZ, dPrevStartX, dPrevStartY, dPrevStartZ
    
    For Each iCurve In Rails
    
        iCurve.EndPoints dCurrStartX, dCurrStartY, dCurrStartZ, dCurrEndX, dCurrEndY, dCurrEndZ
        
        ' if this curve is not end-matched to previous, then start a new row.
        If Abs(dCurrStartX - dPrevEndX) > 0.001 Or Abs(dCurrStartY - dPrevEndY) > 0.001 Or Abs(dCurrStartZ - dPrevEndZ) > 0.001 Then
            NRailRows = NRailRows + 1
        End If
    
        dPrevEndX = dCurrEndX
        dPrevEndY = dCurrEndY
        dPrevEndZ = dCurrEndZ
        
        RailSLO.List(Str(NRailRows)).Add iCurve
    
    Next iCurve

    Exit Sub

End Sub


' this is the Main function called from some client
Public Function ConnectComponents(oHandrail As ISPSHandrail) As SPSHandrails.StructHandrailConvertHelperStatus

    Const METHOD = "ConnectComponents"
    On Error GoTo ErrHandler
       
    Dim iMemberSystem As ISPSMemberSystem
    Dim lRailCount As Long, lRailNumber As Long
    Dim eStatus As SPSHandrails.StructHandrailConvertHelperStatus
    Dim Row As Long, NMidRailRows As Long
    Dim MidRailSLO As IJStructListOwner
    Dim bCurvedRailExists As Boolean

    eStatus = StructHandrailConvertHelperStatus_Unexpected
    
    If Not oHandrail.IsDesignedHandrail Then
        ' Cannot try to connect a handrail that has not been converted yet.
        eStatus = StructHandrailConvertHelperStatus_InvalidHandrail
        GoTo ExitConnect
    End If

    GetComponents oHandrail
    ' connect the top rails to each other using Axis-End Connections

    lRailCount = m_TopRails.Count
    If lRailCount = 0 Then
        ' if there are no top rails, something is wrong
        eStatus = StructHandrailConvertHelperStatus_NoTopRails
        GoTo ExitConnect
    End If
    
    ' detect whether any curved rails exist.  this boolean is also set when we have curved end treatment.
    ' the effect of this is that posts are attached to the rails instead of attaching rails to posts.
    ' that is done so that if an object supporting a post that controls the height of rail is moved up or down,
    ' the posts get shorter or longer and leaves the rails in place.   Then to keep the handrail height the same,
    ' user can do a simple vertical move of the rails.   If we connect curved rails to the posts at their ends, then the
    ' curved rail will distort which is more difficult to fix.
    ' See also TR 256683 for rationale and behaviors.
    
    bCurvedRailExists = False
    For lRailNumber = 1 To lRailCount
        If TypeOf m_TopRails(lRailNumber) Is ISPSMemberSystemCurve Then
            bCurvedRailExists = True
        End If
    Next lRailNumber
    If Not bCurvedRailExists Then
        If Not m_BeginTreatment Is Nothing Then
            If TypeOf m_BeginTreatment Is ISPSMemberSystemCurve Then
                bCurvedRailExists = True
            End If
        End If
    End If
    If Not bCurvedRailExists Then
        If Not m_EndTreatment Is Nothing Then
            If TypeOf m_EndTreatment Is ISPSMemberSystemCurve Then
                bCurvedRailExists = True
            End If
        End If
    End If

    ' Setting bCurvedRailExists to True makes posts connect to rails.
    ' Setting bCurvedRailExists to False connects rails to posts.
    ' Commenting this line makes it depend on the code above that detects curved rails or endTreatments
    
    ' bCurvedRailExists = True
    
    ' user may have created several midrails.   Separate into rows of endpoint-matched member systems.
    ' the first one is the upper midrail and the last is the lowest midrail
    Set MidRailSLO = New StructListOwner
        
    SeparateRails m_MidRails, MidRailSLO, NMidRailRows

    'connect each rail start either to a post or the previous rail element
    eStatus = ConnectRails(m_Posts, m_TopRails, bCurvedRailExists)
    For Row = 1 To NMidRailRows
        eStatus = ConnectRails(m_Posts, MidRailSLO.List(Str(Row)), bCurvedRailExists)
    Next Row
    eStatus = ConnectRails(m_Posts, m_ToePlates, bCurvedRailExists)
    
    ' if more than one midRail exists, the middle midRails should be connected to the endTreatment
    If NMidRailRows > 1 And Not m_BeginTreatment Is Nothing Then
        For Row = 1 To NMidRailRows - 1
            Set iMemberSystem = MidRailSLO.List(Str(Row))(1)
            eStatus = CreateConnection(iMemberSystem.FrameConnectionAtEnd(SPSMemberAxisStart), "Axis-Along", m_BeginTreatment)
            SetPositionRule iMemberSystem.FrameConnectionAtEnd(SPSMemberAxisStart), SPSPODimension_EndSelection_Start, SPSPODimension_Position_Distance
        Next Row
    End If
    If NMidRailRows > 1 And Not m_EndTreatment Is Nothing Then
        For Row = 1 To NMidRailRows - 1
            Set iMemberSystem = MidRailSLO.List(Str(Row))(MidRailSLO.List(Str(Row)).Count)
            eStatus = CreateConnection(iMemberSystem.FrameConnectionAtEnd(SPSMemberAxisEnd), "Axis-Along", m_EndTreatment)
            SetPositionRule iMemberSystem.FrameConnectionAtEnd(SPSMemberAxisEnd), SPSPODimension_EndSelection_End, SPSPODimension_Position_Distance
        Next Row
    End If

    'At this point, rails have their startFC set to a pointOn post or prev rail.
    'Now, connect unset startFC's to a post, mid-posts to rails, and rail endFC's to a post.
    For Row = 1 To NMidRailRows
        eStatus = ConnectPostsAndRails(m_Posts, MidRailSLO.List(Str(Row)), False, bCurvedRailExists)
    Next Row

    eStatus = ConnectPostsAndRails(m_Posts, m_ToePlates, False, bCurvedRailExists)

    ' connect the posts to the top rails
    eStatus = ConnectPostsAndRails(m_Posts, m_TopRails, True, bCurvedRailExists)

    If Not m_BeginTreatment Is Nothing Then
        eStatus = ConnectTreatmentToRails(m_BeginTreatment, m_TopRails(1), MidRailSLO.List(Str(NMidRailRows))(1))
    End If

    If Not m_EndTreatment Is Nothing Then
        eStatus = ConnectTreatmentToRails(m_EndTreatment, m_TopRails(m_TopRails.Count), MidRailSLO.List(Str(NMidRailRows))(MidRailSLO.List(Str(NMidRailRows)).Count))
    End If

    ' this function connects the bottom of posts to Structure objects, such as Members, Slabs or SmartMarine Plates.
    ' to disable that, merely comment this function so it does not get called.
    ConnectPostsToStructureObjects oHandrail
    
    ' cleanup local vars
    Set m_TopRails = Nothing
    Set m_MidRails = Nothing
    Set m_ToePlates = Nothing
    Set m_Posts = Nothing
    Set m_BeginTreatment = Nothing
    Set m_EndTreatment = Nothing
    Set m_oMemberFactory = Nothing
    Set MidRailSLO = Nothing

ExitConnect:
    ConnectComponents = eStatus
    Exit Function
    
ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    ConnectComponents = StructHandrailConvertHelperStatus_Unexpected
End Function

Private Sub GetComponents(oHandrail As ISPSHandrailConvertHelper)
Const METHOD = "GetComponents"
On Error GoTo ErrHandler
Set m_oErrors = New IMSErrorLog.JServerErrors
Set m_oMemberFactory = New SPSMembers.SPSMemberFactory

    With oHandrail
        Set m_TopRails = .TopRails
        Set m_MidRails = .MidRails
        Set m_ToePlates = .ToePlates
        Set m_Posts = .Posts
        Set m_BeginTreatment = .BeginTreatment
        Set m_EndTreatment = .EndTreatment
    End With

Exit Sub
ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    With m_oErrors.Item(m_oErrors.Count)
        Err.Raise .Number, .Source, .Description & .ExtraInfo, .HelpFile, .HelpContext
    End With
End Sub


Private Function CreateConnection(oFirstObj As Object, strFCTypeName As String, oSecondObj As Object) As StructHandrailConvertHelperStatus
Const METHOD = "CreateConnection"
On Error GoTo ErrHandler
    Dim eStatus As SPSHandrails.StructHandrailConvertHelperStatus
    eStatus = StructHandrailConvertHelperStatus_Unexpected
    Dim oFirstFC As ISPSFrameConnection
    If TypeOf oFirstObj Is ISPSFrameConnection Then
        Set oFirstFC = oFirstObj
        If Not oFirstFC.definition Is Nothing Then
            ' this frame connection is already connected
            eStatus = StructHandrailConvertHelperStatus_Unneeded
            GoTo ExitCreate
        End If
        ' set the type of frame connection
        oFirstFC.DefinitionName = strFCTypeName
        'get its input helper
        Dim oFirstInputHelper As ISPSFCInputHelper
        Set oFirstInputHelper = oFirstFC.InputHelper
        Dim eFCStatus As SPSFCInputHelperStatus
        On Error Resume Next
        eFCStatus = oFirstInputHelper.SetRelatedObjects(oFirstFC, oSecondObj, Nothing)
        On Error GoTo ErrHandler
        
        If eFCStatus = SPSFCInputHelper_Ok Then
            eStatus = StructHandrailConvertHelperStatus_Ok
        End If
    Else ' first object is not a frame connection.  we need to create a split conn
        ' create the member factory
        
        'get the resource manager
        Dim oDObject As iJDObject
        Set oDObject = oFirstObj
        
        'call the factory to create the Split connection
        Dim oSplitNoneConnection As ISPSSplitMemberConnection
        Set oSplitNoneConnection = m_oMemberFactory.CreateSplitMemberConnection(oDObject.ResourceManager)
        
        'get the input helper from the (newly created) split connection
        Dim oSplitHelper As IJStructILCHelper
        Set oSplitHelper = oSplitNoneConnection.Helper
        
        'add the rail member to an elements collection
        Dim oRailElems As IJElements
        Set oRailElems = New JObjectCollection
        Dim addedIdx As Long
        addedIdx = oRailElems.Add(oSecondObj)
        
        Dim eHelperStatus As StructSOCInputHelperStatus
        Dim pRelObj1 As Object
        Dim pRelParents As IJElements
        eHelperStatus = oSplitHelper.ValidateParents(oSplitNoneConnection, _
                                     oFirstObj, _
                                     oRailElems, _
                                     pRelObj1, _
                                     pRelParents)
                                     
        If eHelperStatus = StructSOCInputHelper_Ok Then
           eHelperStatus = oSplitHelper.SetParents(oSplitNoneConnection, _
                                    pRelObj1, _
                                    pRelParents)
           Dim oSplitConn As IJStructILCConnection
           Set oSplitConn = oSplitNoneConnection
           
           oSplitConn.SplitParentStatus = ssSplitNone
           eStatus = StructHandrailConvertHelperStatus_Ok
        End If
    End If

ExitCreate:
    CreateConnection = eStatus

Exit Function
ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    CreateConnection = StructHandrailConvertHelperStatus_Unexpected
End Function


'Only used for TopMountedToPad
'Create assembly connection at the end of each post to be the Pad
'The material and the thickness of the AC plate will be set in next run after the computing (ResetAllPadsSize)
Public Function AddPadsToPosts(oHandrail As ISPSHandrail, oSize As IJDVector) As StructHandrailConvertHelperStatus
Const METHOD = "AddPadsToPosts"
On Error GoTo ErrHandler
    Dim bCanAddPads As Boolean
    bCanAddPads = Not oHandrail.IsDesignedHandrail
    
    Dim eStatus As SPSHandrails.StructHandrailConvertHelperStatus
    eStatus = StructHandrailConvertHelperStatus_Unexpected
    
    If bCanAddPads Then
        Dim oHandrailHelp As ISPSHandrailConvertHelper
        Set oHandrailHelp = oHandrail
        Dim oPostMemberSys As ISPSMemberSystem
        Dim eIntersectionLocation As SPSMemberAxisPortIndex
        Dim oMemberPart As ISPSMemberPartCommon
        For Each oPostMemberSys In oHandrailHelp.Posts
            Dim x1 As Double, y1 As Double, z1 As Double
            Dim x2 As Double, y2 As Double, z2 As Double
            Dim x3 As Double, y3 As Double, z3 As Double
                        'Pad is always on the lower end of the post
            oPostMemberSys.LogicalAxis.GetLogicalEndPoint x2, y2, z2
            oPostMemberSys.LogicalAxis.GetLogicalStartPoint x1, y1, z1
            If z1 > z2 Then
                eIntersectionLocation = SPSMemberAxisEnd
            Else
                eIntersectionLocation = SPSMemberAxisStart
            End If
                        'Get Proxies for IJDMaterial and IJDPlateDimensions for later use
            If m_MaterialProxy Is Nothing Then
                If oHandrailHelp.MaterialProxy Is Nothing Then
                    If oMemberPart Is Nothing Then
                        Set oMemberPart = oPostMemberSys.MemberPartAtEnd(eIntersectionLocation)
                    End If
                    oHandrailHelp.MaterialProxy = GetMaterialProxy(oMemberPart)
                End If
                Set m_MaterialProxy = oHandrailHelp.MaterialProxy
                Dim oMaterial As IJDMaterial
                Set oMaterial = m_MaterialProxy
            End If
            If m_PlateDimensionsProxy Is Nothing Then
                Set oMemberPart = oPostMemberSys.MemberPartAtEnd(eIntersectionLocation)
                If oHandrailHelp.PlateDimensionsProxy Is Nothing Then
                    Set oHandrailHelp.PlateDimensionsProxy = GetPlateDimensionsProxy(oMemberPart, m_MaterialProxy, oSize.z)
                End If
                Set m_PlateDimensionsProxy = oHandrailHelp.PlateDimensionsProxy
            End If
                'Creat AC for each post
            eStatus = CreateFrameAssemblyConnectionAsPad(oPostMemberSys, eIntersectionLocation, oSize)
        Next oPostMemberSys ' post
        Set m_PlateDimensionsProxy = Nothing
        Set m_MaterialProxy = Nothing
    End If
    AddPadsToPosts = eStatus
Exit Function
ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    AddPadsToPosts = StructHandrailConvertHelperStatus_Unexpected
End Function


Private Sub SetPositionRule(oFC As ISPSFrameConnection, _
                            Optional eEndRule As StructMemberEndSelectionRule = SPSPODimension_EndSelection_Start, _
                            Optional ePositionRule As StructMemberPositionRule = SPSPODimension_Position_Distance)
Const METHOD = "SetPositionRule"
On Error GoTo ErrHandler
    ' For any rule other than "Intersection", we have to measure and set the rule
   If ePositionRule <> SPSPODimension_Position_Intersection Then
        Dim dDistOrRatio As Double
        With oFC.Joint.PointOnDimension
            ' measure to get the distance or ratio
            .Measure eEndRule, ePositionRule, dDistOrRatio
            
            ' set the value using the appropriate method
            If ePositionRule = SPSPODimension_Position_Ratio Then
                .PointOnRatio = dDistOrRatio
            ElseIf ePositionRule = SPSPODimension_Position_Distance Then
                .PointOnDistance = dDistOrRatio
            End If
            ' set the rule
            .PointOnEndRule = eEndRule
            .PointOnPositionRule = ePositionRule
        End With
    End If
Exit Sub
ErrHandler:

    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub

Private Function CreateFrameAssemblyConnectionAsPad(oMemSys As ISPSMemberSystem, whichEnd As SPSMemberAxisPortIndex, oSize As IJDVector) As StructHandrailConvertHelperStatus
Const METHOD = "CreateFrameAssemblyConnectionAsPad"
On Error GoTo ErrorHandler
    Dim eStatus As SPSHandrails.StructHandrailConvertHelperStatus
    eStatus = StructHandrailConvertHelperStatus_Ok

    Dim oAssConn As IJStructAssemblyConnection
    Dim oSmartOcc As IJSmartOccurrence
    Dim oAsmConnItm As IJSmartItem
    Dim oAssConnFac As StructConnectionsFactory
    Dim oRefColl As IJDReferencesCollection
    Dim oSystem As IJSystem
    Dim oDesChild As IJDesignChild
    Set oDesChild = oMemSys
    Dim oSymFac  As IJDSymbolEntitiesFactory
    Dim oRscMgr As Object
    Dim oStructConn As IJAppConnection
    Dim strSelectedItm As String, strSelectedClass As String
    Dim strAsmConn As String
    Dim iJDObject As iJDObject
    Dim oMemberPart As ISPSMemberPartCommon
    Dim oPort As IJPort
    Dim oAttrColl As Object

    
    Set oMemberPart = oMemSys.MemberPartAtEnd(whichEnd)
    
    Set iJDObject = oMemberPart
    Set oPort = oMemberPart.AxisPort(whichEnd)
    Set oSystem = oDesChild.GetParent
        'create ref entity
    Set oSymFac = New DSymbolEntitiesFactory
    Set oRefColl = oSymFac.CreateEntity(referencesCollection, iJDObject.ResourceManager)
    Set oSymFac = Nothing
    'create AC
    Set oAssConnFac = New StructConnectionsFactory
    Set oAssConn = oAssConnFac.CreateStructAssemblyConnection(iJDObject.ResourceManager)
    Set oAssConnFac = Nothing
    
    Set oSmartOcc = oAssConn
    'set ref for the AC
    SetRefColl oSmartOcc, oRefColl
    strSelectedItm = "BasePlateAsmConn_1"
    On Error Resume Next
    Set oAsmConnItm = GetConnectionDefinition(strSelectedItm)
    If Err.Number <> 0 Then
        Err.Raise Err.Number, METHOD, Err.Description
    End If
    On Error GoTo ErrorHandler
        'set parent system
    If Not oAsmConnItm Is Nothing Then
        strSelectedClass = oAsmConnItm.Parent.SCName
        oSmartOcc.RootSelection = strSelectedItm
        Set oAsmConnItm = Nothing
    End If
    
    Set oStructConn = oAssConn
        'set connection to the port
    oStructConn.addPort oPort
    oSystem.AddChildMember oStructConn
    SetAssemblyConnectionNameRule oAssConn
    'set plate length and width, thickness will be set in next run
    Set oAttrColl = GetAttributeCollection(oAssConn, "IJUASPSBasePlateAsmConn")
    SetAttributeValue oAttrColl, "DepthClearance", oSize.x
    SetAttributeValue oAttrColl, "WidthClearance", oSize.y
    
    Set oAssConn = Nothing
    Set oRefColl = Nothing
    Set oStructConn = Nothing
    CreateFrameAssemblyConnectionAsPad = eStatus
    Exit Function
ErrorHandler:
    CreateFrameAssemblyConnectionAsPad = StructHandrailConvertHelperStatus_Unexpected
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

'Copied from SPSAsmConnCommands for the AC creation
Public Function GetConnectionDefinition(ByVal name As String) As Object
    Const METHOD = "GetConnectionDefinition"
    On Error GoTo ErrHandler

    Dim oNamingCntx As IJDNamingContext
    Set oNamingCntx = GetCatalogDBConnection

    Set GetConnectionDefinition = oNamingCntx.Resolve(name)

        If GetConnectionDefinition Is Nothing Then
                MsgBox "GetConnectionDefinition: Could not find item " & name
        End If
    Set oNamingCntx = Nothing
    Exit Function
ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Private Sub SetRefColl(pFC As Object, pRefColl As IJDReferencesCollection)
Const METHOD = "SetRefColl"
 On Error GoTo ErrorHandler
  
   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pFC
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    pCollectionHelper.Add pRefColl, "RC", pRelationshipHelper
    Set pRevision = New JRevision
    pRevision.AddRelationship pRelationshipHelper
  
   Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub

Public Function GetCatalogDBConnection() As IJDConnection
    Const METHOD = "GetCatalogDBConnection"
    On Error GoTo ErrHandler
    Dim oTrader As IMSTrader.Trader
    Set oTrader = New Trader
    Dim oWorkingSet As IJDWorkingSet
    Set oWorkingSet = oTrader.Service("WorkingSet", "")
    Dim strCatlogDB As String
    Dim oAppCtx As IJApplicationContext
    Set oAppCtx = oTrader.Service("ApplicationContext", "")
    strCatlogDB = oAppCtx.DBTypeConfiguration.get_DataBaseFromDBType("Catalog")
    Set GetCatalogDBConnection = oWorkingSet.Item(strCatlogDB)
    Set oTrader = Nothing
    Set oWorkingSet = Nothing
    Set oAppCtx = Nothing
Exit Function
ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Private Sub SetAssemblyConnectionNameRule(oAsmConn As IJStructAssemblyConnection)  ', NewNameRule As String)
Const METHOD = "SetAssemblyConnectionNameRule"
On Error GoTo ErrorHandler
    Dim NameRule As String
    Dim found As Boolean
    found = False

    Dim NamingRules As IJElements
    Dim oNameRuleHolder As IJDNameRuleHolder
    Dim oActiveNRHolder As IJDNameRuleHolder
    Dim oNameRuleHlpr As IJDNamingRulesHelper
    Set oNameRuleHlpr = New NamingRulesHelper

    oNameRuleHlpr.GetEntityNamingRulesGivenProgID ASCONNECPROGID, NamingRules
    Dim ncount As Integer
    Dim oNameRuleAE As IJNameRuleAE

    If NamingRules.Count > 0 Then
        Set oNameRuleHolder = NamingRules.Item(1)
    End If

    Call oNameRuleHlpr.AddNamingRelations(oAsmConn, oNameRuleHolder, oNameRuleAE)

    Set oNameRuleHolder = Nothing

    Set oActiveNRHolder = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
Exit Sub
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub
'End of -Copied from SPSAsmConnCommands for the AC creation

'Material and thickness can only be set on base plate which is the child of AC
Public Function ResetAllPadsSize(oHandrail As ISPSHandrail) As StructHandrailConvertHelperStatus
Const METHOD = "ResetAllPadsSize"
On Error GoTo ErrorHandler
    Dim bCanConnect As Boolean
    bCanConnect = oHandrail.IsDesignedHandrail
    
    Dim eStatus As SPSHandrails.StructHandrailConvertHelperStatus
    eStatus = StructHandrailConvertHelperStatus_Unexpected
    
    If bCanConnect Then
        Dim oHandrailHelp As ISPSHandrailConvertHelper
        Dim oPostMemberSys As ISPSMemberSystem
        Dim eIntersectionLocation As SPSMemberAxisPortIndex
        Set oHandrailHelp = oHandrail
        For Each oPostMemberSys In oHandrailHelp.Posts
            Dim x1 As Double, y1 As Double, z1 As Double
            Dim x2 As Double, y2 As Double, z2 As Double
            Dim x3 As Double, y3 As Double, z3 As Double
            oPostMemberSys.LogicalAxis.GetLogicalEndPoint x2, y2, z2
            oPostMemberSys.LogicalAxis.GetLogicalStartPoint x1, y1, z1
            If z1 > z2 Then
                eIntersectionLocation = SPSMemberAxisEnd
            Else
                eIntersectionLocation = SPSMemberAxisStart
            End If
            eStatus = SetPadDimensions(oHandrailHelp, oPostMemberSys, eIntersectionLocation)
        Next oPostMemberSys ' post
    End If
ExitCreate:
    ResetAllPadsSize = eStatus
Exit Function
ErrorHandler:
    ResetAllPadsSize = StructHandrailConvertHelperStatus_Unexpected
    m_oErrors.Add Err.Number, METHOD, Err.Description

End Function

'Set material and thickness for the Pad (AC base plate)
Public Function SetPadDimensions(oHandrail As ISPSHandrail, oMemSys As ISPSMemberSystem, whichEnd As SPSMemberAxisPortIndex) As StructHandrailConvertHelperStatus
Const METHOD = "SetPadDimensions"
On Error GoTo ErrorHandler
    Dim eStatus As SPSHandrails.StructHandrailConvertHelperStatus
    eStatus = StructHandrailConvertHelperStatus_Ok
    Dim oAssConn As IJStructAssemblyConnection
    Dim oSystem As IJSystem
    Dim oStructConn As IJAppConnection
    Dim strSelectedItm As String, strSelectedClass As String
    Dim strAsmConn As String

    Dim iJDObject As iJDObject

    Dim oMemberPart As ISPSMemberPartCommon
    Dim oPort As IJPort
    Dim oConnection As Object
    Dim oConnections As IJElements
    Dim oPlate As Object
    Dim oDesignParent As IJDesignParent
    Dim oDesignChildren As IJDObjectCollection
    Dim oSmartOcc As IJSmartOccurrence

        'Get MemberPart from Member System
    Set oMemberPart = oMemSys.MemberPartAtEnd(whichEnd)
    'Get Port
        Set oPort = oMemberPart.AxisPort(whichEnd)
        'Get Connection
    oPort.enumConnections oConnections
    For Each oConnection In oConnections
        Set oAssConn = oConnection
        If Not oAssConn Is Nothing Then Exit For
    Next
        'Get base plate
    Set oDesignParent = oAssConn
    oDesignParent.GetChildren oDesignChildren, vbNullString
    For Each oPlate In oDesignChildren
        Set oSmartOcc = oPlate
        If oSmartOcc.RootSelection = "GenericRectPlatePart" Then Exit For
    Next
        'Set Material and thickness for the base plate
    If Not oPlate Is Nothing Then
        Dim oStructPlate As IJStructPlate
        Dim oPlateMaterial As IJStructMaterial
        Dim oMaterial As IJDMaterial
        Dim oHandrailHelp As ISPSHandrailConvertHelper
        Set oHandrailHelp = oHandrail
        Set oStructPlate = oPlate
        If Not oStructPlate Is Nothing Then
            oStructPlate.Dimensions = oHandrailHelp.PlateDimensionsProxy
        End If
        Set oPlateMaterial = oPlate
        Set oMaterial = oHandrailHelp.MaterialProxy
        If Not oPlateMaterial Is Nothing Then
            oPlateMaterial.StructMaterial = oHandrail.ConvertHelper.MaterialProxy   'GetMaterialProxyFromNames(oStructPlate, oMaterial.MaterialType, oMaterial.MaterialGrade)   'oHandrailHelp.MaterialProxy
        End If
    End If
    Set oConnections = Nothing
    Set oMemberPart = Nothing
    Set oStructConn = Nothing
    Set oPlate = Nothing
ExitCreate:
    SetPadDimensions = eStatus
Exit Function
ErrorHandler:
    SetPadDimensions = StructHandrailConvertHelperStatus_Unexpected
    m_oErrors.Add Err.Number, METHOD, Err.Description

End Function


Public Function GetMaterialProxy(oModelObj As Object) As IJDProxy
    Const METHOD = "GetMaterialProxy"
    On Error GoTo ErrorHandler

    Dim oMatlObj As IJDMaterial
    Dim oModelPOM As IJDPOM
    Dim iJDObject As iJDObject
    
    Set oMatlObj = oModelObj

    Set iJDObject = oModelObj
    Set oModelPOM = iJDObject.ResourceManager
    Set GetMaterialProxy = oModelPOM.GetProxy(oMatlObj, True)

    Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Public Function GetMaterialProxyFromNames(oModelObj As Object, ByVal strType As String, ByVal strGrade As String) As IJDProxy
    Const METHOD = "GetMaterialProxyFromNames"
    On Error GoTo ErrorHandler

    Dim oMatlObj As IJDMaterial
    Dim oRefDataQuery As IJDStructServices
    Dim oPlateMaterial As IJStructMaterial
    Dim oCatalogPOM As IJDPOM
    Dim oModelPOM As IJDPOM
    Dim iJDObject As iJDObject
    
    Set oRefDataQuery = New StructServices
    Set oCatalogPOM = GetCatalogResourceManager()
    Set oMatlObj = oRefDataQuery.GetMaterialFromGradeAndType(oCatalogPOM, strType, strGrade)

    Set iJDObject = oModelObj
    Set oModelPOM = iJDObject.ResourceManager
    Set GetMaterialProxyFromNames = oModelPOM.GetProxy(oMatlObj, True)

    Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Public Function ConnectPostsToStructureObjects(oHandrail As ISPSHandrail) As StructHandrailConvertHelperStatus
    Const METHOD = "ConnectPostsToStructureObjects"
    On Error GoTo ErrorHandler
    
    Dim dtol As Double
    Dim oPostMemberSystem As ISPSMemberSystem
    Dim oStructPersistTools As IJStructPersistTools

    Dim eWhichEnd As SPSMemberAxisPortIndex
    Dim iMemberPartCommon As ISPSMemberPartCommon
    Dim ICrossSection As ISPSCrossSection
    Dim dWidth As Double, dDepth As Double, dPostSize As Double

    Dim gboxRange As GBox
    Dim strQuery As String

    Dim oBO As Object
    Dim oPort As IJPort
    Dim oSLOStructBO As IJStructListOwner, oSLOPort As IJStructListOwner
    Dim elesStructureSLOs As IJElements, elesPortSLOs As IJElements
    
    Dim elesGraphics As IJElements
    Dim dDistance As Double, dDistanceMinSurface As Double, dDistanceMinMember As Double
    Dim dPostX As Double, dPostY As Double, dPostZ As Double
    Dim dPortX As Double, dPortY As Double, dPortZ As Double, dZMaxSurface As Double
    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    Dim dEndX As Double, dEndY As Double, dEndZ As Double
    
    Dim oSLOStructBOMinMember As IJStructListOwner, oSLOStructBOMinSurface As IJStructListOwner
    Dim oSLOPortMinMember As IJStructListOwner
    Dim oSLOPortMinSurface As IJStructListOwner
    
    Dim bClosed As Boolean
    Dim oPortGraphic As Object
    Dim oPortMinSurfaceGraphic As Object
    
    'use this line to find distance to the post start.
    Dim iLinePostStart As IJLine        ' needed to set ends for each post
    Dim iCurvePostStart As IJCurve      ' used to get mindist to the ports from the bottom of the post
    Dim iCurvePost As IJCurve           ' used to get the post's ends
    Dim oGeometryFactory As GeometryFactory
    
    ConnectPostsToStructureObjects = StructHandrailConvertHelperStatus_Ok

    If m_Posts Is Nothing Then
        Exit Function
    ElseIf m_Posts.Count = 0 Then
        Exit Function
    End If
    
    'get cross-section size of the posts
    Set oPostMemberSystem = m_Posts.Item(1)
    Set iMemberPartCommon = oPostMemberSystem.DesignPartAtEnd(SPSMemberAxisStart)
    Set ICrossSection = iMemberPartCommon.CrossSection
    ICrossSection.GetCardinalPointDelta Nothing, 1, 9, dWidth, dDepth
    dPostSize = Sqr(dWidth * dWidth + dDepth * dDepth)
    Set oPostMemberSystem = Nothing
    Set iMemberPartCommon = Nothing
    Set ICrossSection = Nothing
    
    If dPostSize < 0.05 Then    'dPostSize is at least 50 mm
        dPostSize = 0.05
    End If
    dtol = 0.000001

    BuildRangeBox m_Posts, gboxRange
    
    'enlarge the range box by post size in X and Y, and slightly in Z
    gboxRange.m_low.x = gboxRange.m_low.x - dPostSize
    gboxRange.m_low.y = gboxRange.m_low.y - dPostSize
    gboxRange.m_low.z = gboxRange.m_low.z - 0.05        'include objects below bottom of posts by 50 mm
    gboxRange.m_high.x = gboxRange.m_high.x + dPostSize
    gboxRange.m_high.y = gboxRange.m_high.y + dPostSize
    gboxRange.m_high.z = gboxRange.m_high.z + 0.05      'include objects above bottom of posts by 50 mm
    
    'classids..     240012 = MemberPartPrismatic
    '               240029 = MemberPartCurve
    '               240024 = Slab
    '              1090063 = SM Plate Part
    
    'range query..  exclude all objects, if for any axis, its min is > sketch max, or its max < sketch min."
    'which leads to: for all three axes, BO min < sketch max and BO max > sketch min.
    strQuery = "select CB.oid from COREBaseClass CB join CORESpatialIndex CSI on CB.oid = CSI.oid where " & _
            "(CB.ClassId in (240012,240024,240029,1090063)" & _
            " and (CSI.xmin < " & Str(gboxRange.m_high.x) & " and CSI.xmax > " & Str(gboxRange.m_low.x) & _
            " and  CSI.ymin < " & Str(gboxRange.m_high.y) & " and CSI.ymax > " & Str(gboxRange.m_low.y) & _
            " and  CSI.zmin < " & Str(gboxRange.m_high.z) & " and CSI.zmax > " & Str(gboxRange.m_low.z) & "))"

    Set oStructPersistTools = New StructPersistTools
    oStructPersistTools.QueryDBAndBind oHandrail, strQuery, True, elesStructureSLOs
    
    If elesStructureSLOs Is Nothing Then
        Exit Function
    End If
    
    'for each located structure business object, get its face ports, and for each port cache a list of graphics
    'determine if it belongs to a MemberSystem, and if so record the member system object in the SLO Data variant
    For Each oSLOStructBO In elesStructureSLOs
        CreateSurfacePortListWithGraphics oSLOStructBO, strVisiblePorts, False
        Set oBO = GetParentMemberSystem(oSLOStructBO.OwnerObject, 0)
        If Not oBO Is Nothing Then
            oSLOStructBO.Data = oBO
        End If
        Set oBO = Nothing
    Next oSLOStructBO

    'using the data created above, for each post find minimum-distance port

    Set oGeometryFactory = New GeometryFactory
    Set iLinePostStart = oGeometryFactory.Lines3d.CreateBy2Points(Nothing, 0#, 0#, 0#, 1#, 0#, 0#)
    Set iCurvePostStart = iLinePostStart
    Set oGeometryFactory = Nothing
    
    ' from the post bottom, find the closest member port and the closest non-member port.
    ' if there is a surface port, and it is closer than any member port, check if the post intersects the port completely, and if so, use the port.
    
    For Each oPostMemberSystem In m_Posts
        
        dDistanceMinMember = 50000000
        dDistanceMinSurface = 50000000
        dZMaxSurface = 50000000
        
        Set oSLOPortMinMember = Nothing
        Set oSLOPortMinSurface = Nothing
        Set oSLOStructBOMinMember = Nothing
        Set oSLOStructBOMinSurface = Nothing
        Set oPortMinSurfaceGraphic = Nothing

        'set iCurvePostStart to be a degenerate line at the bottom of the post
        Set iCurvePost = oPostMemberSystem
        iCurvePost.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
        If dStartZ < dEndZ Then
            eWhichEnd = SPSMemberAxisStart
            iLinePostStart.DefineBy2Points dStartX, dStartY, dStartZ, dStartX, dStartY, dStartZ
        Else
            eWhichEnd = SPSMemberAxisEnd
            iLinePostStart.DefineBy2Points dEndX, dEndY, dEndZ, dEndX, dEndY, dEndZ
        End If

        'loop through the located BO's and find the nearest member and nearest other surface
        For Each oSLOStructBO In elesStructureSLOs                      ' for each located BO
            
            If Not IsEmpty(oSLOStructBO.Data) Then
                Set oBO = oSLOStructBO.Data                             ' is this a child of a MemberSystem ?
            End If
            Set elesPortSLOs = oSLOStructBO.List(strVisiblePorts)       ' for each visible Port

            For Each oSLOPort In elesPortSLOs
                Set elesGraphics = oSLOPort.List(strGraphics)
                
                For Each oPortGraphic In elesGraphics                   ' for each IJGraphicEntity of that port
                    
                    iCurvePost.DistanceBetween oPortGraphic, dDistance, dPostX, dPostY, dPostZ, dPortX, dPortY, dPortZ
                    
                    If dDistance < dPostSize Then
                        If Not oBO Is Nothing Then
                            If dDistance < dDistanceMinMember Then
                                dDistanceMinMember = dDistance
                                Set oSLOPortMinMember = oSLOPort
                                Set oSLOStructBOMinMember = oSLOStructBO
                            End If
                        Else
                            'for surfaces, keep the closest one.  or the higher one if they are both on the post axis
                            If dDistance < dDistanceMinSurface - dtol Or dPortZ > dZMaxSurface Then
                                dZMaxSurface = dPortZ
                                dDistanceMinSurface = dDistance
                                Set oSLOPortMinSurface = oSLOPort
                                Set oPortMinSurfaceGraphic = oPortGraphic       ' save this particular minimum surface
                                Set oSLOStructBOMinSurface = oSLOStructBO
                            End If
                        End If
                    End If
                Next oPortGraphic
            Next oSLOPort
            Set oBO = Nothing
        Next oSLOStructBO
        
        bClosed = False
        ' found a mindist surface.   if it is much closer than the nearest member and closer than dPostSize...
        ' only connect to surface if the post intersects with the surface completely.
        If (Not oSLOPortMinSurface Is Nothing) And (dDistanceMinSurface < dDistanceMinMember - dtol) And dDistanceMinSurface < dPostSize Then
            
            'if no member is nearby, okay to make the surface FC even if partial or no contact.
            If dDistanceMinMember > dPostSize Then
                bClosed = True
            Else
                'since there is a member nearby, connect the post to the surface only if the post intersects the surface completely.
                IsPostPortSurfaceIntersectionClosed oPostMemberSystem, oPortMinSurfaceGraphic, bClosed
            End If
        
        End If
        
        ' found visible surface port that is close enough with closed post-surface intersection.   make surface connection, but using stable port
        If bClosed Then
            ' find a stable port
            dDistanceMinSurface = 50000000
            dZMaxSurface = 50000000
            Set oSLOPortMinSurface = Nothing
            CreateSurfacePortListWithGraphics oSLOStructBOMinSurface, strStablePorts, True
            Set elesPortSLOs = oSLOStructBOMinSurface.List(strStablePorts)
            For Each oSLOPort In elesPortSLOs
                Set elesGraphics = oSLOPort.List(strGraphics)
                For Each oPortGraphic In elesGraphics
                    iCurvePost.DistanceBetween oPortGraphic, dDistance, dPostX, dPostY, dPostZ, dPortX, dPortY, dPortZ
                    If dDistance < dPostSize Then
                        If dDistance < dDistanceMinSurface - dtol Or dPortZ > dZMaxSurface Then
                            dZMaxSurface = dPortZ
                            dDistanceMinSurface = dDistance
                            Set oSLOPortMinSurface = oSLOPort
                            Set oPortMinSurfaceGraphic = oPortGraphic
                         End If
                    End If
                Next oPortGraphic
            Next oSLOPort
            
            If (Not oSLOPortMinSurface Is Nothing) And dDistanceMinSurface < dPostSize Then
                ConnectPostToStructureObject oPostMemberSystem.FrameConnectionAtEnd(eWhichEnd), oSLOPortMinSurface.OwnerObject, "Surface-Default"
                SetFCOffsetProperties oPostMemberSystem.FrameConnectionAtEnd(eWhichEnd), iLinePostStart, oPortMinSurfaceGraphic
            End If
        
        ElseIf (Not oSLOPortMinMember Is Nothing) And dDistanceMinMember < dPostSize Then
        
            Set oBO = oSLOStructBOMinMember.Data
            ConnectPostToStructureObject oPostMemberSystem.FrameConnectionAtEnd(eWhichEnd), oBO, "Axis-Along"
            SetFCOffsetProperties oPostMemberSystem.FrameConnectionAtEnd(eWhichEnd), iLinePostStart, oBO
            Set oBO = Nothing
                
        End If
   
    Next oPostMemberSystem

    Set iLinePostStart = Nothing
    ConnectPostsToStructureObjects = StructHandrailConvertHelperStatus_Ok

    Exit Function

ErrorHandler:
   ConnectPostsToStructureObjects = StructHandrailConvertHelperStatus_Unexpected
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function  ' ConnectPostsToStructureObjects

Private Sub ConnectPostToStructureObject(oFC As ISPSFrameConnection, oRelatedObject1 As Object, strFCDef As String)
    Const METHOD = "ConnectPostToStructureObject"
    On Error GoTo ErrorHandler

    Dim x As Double, y As Double, z As Double
    Dim oReturn1 As Object, oReturn2 As Object
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus

    If oRelatedObject1 Is Nothing Then
        Exit Sub
    End If

    oFC.DefinitionName = strFCDef

    IHStatus = oFC.InputHelper.ValidateLocatedObjects(oFC, 0, 0, oRelatedObject1, Nothing, 0, 0, 0, oReturn1, oReturn2, x, y, z)
    If IHStatus = SPSFCInputHelper_Ok Then
        IHStatus = oFC.InputHelper.SetRelatedObjects(oFC, oReturn1, oReturn2)
    End If
    
    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub  ' ConnectPostToStructureObject

Private Sub CreateSurfacePortListWithGraphics(oSLOStructBO As IJStructListOwner, strSLOCollName As String, bStablePorts As Boolean)
    Const METHOD = "CreateSurfacePortListWithGraphics"
    On Error GoTo ErrorHandler
    
    Dim bIsPlate As Boolean
    Dim bIsMember As Boolean
    Dim bIsSlab As Boolean
    Dim oStructBO As Object
    Dim strNull As String
    Dim elesSLOPorts As IJElements
    Dim elesPorts As IJElements
    Dim oPort As Object
    Dim oSLOPort As IJStructListOwner
    
    Dim elesGraphics As IJElements
    Dim eleGraphic As Object
    Dim iPlane As IJPlane
    Dim dNormalZMax As Double, dtol As Double
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double

    Dim varOperation As Variant
    Dim iCommonStructEntity As IJCommonStructEntity
    Dim iStructGraphConnectable As IJStructGraphConnectable
    
    Dim iStructPort As IJStructPort
    Dim ctx As eUSER_CTX_FLAGS
    Dim oPt As Long, opr As Long, sectionID As Long
    Dim topologyProxyType As JS_TOPOLOGY_PROXY_TYPE
    
    bIsSlab = False
    bIsPlate = False
    bIsMember = False
    
    ' first check if this collection is already loaded.
    Set elesSLOPorts = oSLOStructBO.List(strSLOCollName)
    If elesSLOPorts.Count > 0 Then
        Exit Sub
    End If

    Set oStructBO = oSLOStructBO.OwnerObject

    If TypeOf oStructBO Is IJCommonStructEntity Then             ' a plate
        Set iCommonStructEntity = oStructBO
        If bStablePorts Then
            iCommonStructEntity.enumConnectableGlobalPortsByOperation elesPorts, "CreatePlatePart.GeneratePlatePart_AE.1", PortFace, False
        Else
            iCommonStructEntity.enumConnectableGlobalPortsByOperation elesPorts, strNull, PortFace, False
        End If
        bIsPlate = True
            
    ElseIf TypeOf oStructBO Is IJStructGraphConnectable Then    ' a prismatic member or slab
        Set iStructGraphConnectable = oStructBO
        If bStablePorts Then
            iStructGraphConnectable.enumPortsInGraphByTopologyFilter elesPorts, JS_TOPOLOGY_FILTER_ALL_LFACES, StableGeometry, varOperation
            If Not elesPorts Is Nothing Then
                If elesPorts.Count = 0 Then
                    Set elesPorts = Nothing                         ' StructTrimAE
                End If
            End If
            If elesPorts Is Nothing Then
                iStructGraphConnectable.enumPortsInGraphByTopologyFilter elesPorts, JS_TOPOLOGY_FILTER_ALL_LFACES, GeometryByIIDInGraph, "{632B16FB-4EB7-4DA7-9C04-550477711149}"
            End If
        Else
            iStructGraphConnectable.enumPortsInGraphByTopologyFilter elesPorts, JS_TOPOLOGY_FILTER_ALL_LFACES, CurrentGeometry, varOperation
        End If
        If TypeOf oStructBO Is ISPSMemberPartCommon Then
            bIsMember = True
        Else
            bIsSlab = True
        End If
    End If

    dtol = 0.0001

    'for each okay port, create a StructListOwner wrapper, and load it with the port's graphic objects
    If Not elesPorts Is Nothing Then
        For Each oPort In elesPorts
            'use only base/offset surfaces of slabs and plates.   lateral faces of members
            Set iStructPort = oPort
            iStructPort.GetAttributes topologyProxyType, ctx, oPt, opr, sectionID
            'MsgBox "topologyProxyType=" & topologyProxyType & ", ctx=" & CStr(ctx) & ", opt=" & CStr(oPt) & ", opr=" & CStr(opr) & ", sectionID=" & sectionID
            If (bIsMember And sectionID > 0) Or (bIsSlab) Or (bIsPlate And (ctx = CTX_BASE_NMINUS_LFACE Or ctx = CTX_OFFSET_NPLUS_LFACE)) Then
                Set oSLOPort = New StructListOwner
                Set oSLOPort.OwnerObject = oPort
                oSLOPort.CreateGraphicEntityList strGraphics
                
                'check if the port is planar, and if so keep it only if its normal is upward
                Set elesGraphics = oSLOPort.List(strGraphics)
                dNormalZMax = -2
                For Each eleGraphic In elesGraphics                   ' for each IJGraphicEntity of that port
                    If TypeOf eleGraphic Is IJPlane Then
                        Set iPlane = eleGraphic
                        iPlane.GetNormal dNormalX, dNormalY, dNormalZ
                    Else
                        dNormalZ = 1#           ' non-planar port.  keep it.
                    End If
                    If dNormalZ > dNormalZMax Then
                        dNormalZMax = dNormalZ
                    End If
                Next eleGraphic
                
                If dNormalZMax > dtol Then            'keep this port if it is non-planar or its normal is upward
                    elesSLOPorts.Add oSLOPort         'add this SLO wrapper to the BO's list of ports
                End If
                Set oSLOPort = Nothing
            End If
        Next oPort
    End If

    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub

Private Function GetParentMemberSystem(oBO As Object, recurseDepth As Long) As Object
    Const METHOD = "GetParentMemberSystem"
    On Error GoTo ErrorHandler

    Dim oPort As IJPort
    Dim oBONext As Object
    Dim iDesignChild As IJDesignChild
    
    Set GetParentMemberSystem = Nothing

    'if oBO is Nothing, then return.
    If oBO Is Nothing Then
        Exit Function
    End If
    
    'if it is a MemberSystem, we have found it, return it.
    If TypeOf oBO Is ISPSMemberSystem Then
        Set GetParentMemberSystem = oBO
        Exit Function
    'if it is a generic system, then MemberSystem is not in the parent heirarchy
    ElseIf TypeOf oBO Is IJAllowableSpecs Then
        Exit Function
    End If

    'if it is a port, get its connectable
    If TypeOf oBO Is IJPort Then
        Set oPort = oBO
        Set oBONext = oPort.Connectable

    'if it is a designChild, get its parent
    ElseIf TypeOf oBO Is IJDesignChild Then
        Set iDesignChild = oBO
        Set oBONext = iDesignChild.GetParent
    End If
    
    If Not oBONext Is oBO And recurseDepth < 6 Then      ' block against indefinite recursion
        Set GetParentMemberSystem = GetParentMemberSystem(oBONext, recurseDepth + 1)
    End If
    
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Exit Function
End Function

Public Sub SetFCOffsetProperties(oFC As ISPSFrameConnection, iLinePostStart As IJCurve, oSupportingGraphic As Object)
    Const METHOD = "SetFCOffsetProperties"
    On Error GoTo ErrorHandler

    Dim pIJAttrbs As IJDAttributes
    Dim oAttrCollection As CollectionProxy
    Dim oAttr As IJDAttribute

    Dim dDistance As Double, dSupportingX As Double, dSupportingY As Double, dSupportingZ As Double
    Dim dFCx As Double, dFCy As Double, dFCz As Double
    
    If oFC Is Nothing Then
        Exit Sub
    End If
    If oSupportingGraphic Is Nothing Then
        Exit Sub
    End If
    If iLinePostStart Is Nothing Then
        Exit Sub
    End If
    
    iLinePostStart.DistanceBetween oSupportingGraphic, dDistance, dFCx, dFCy, dFCz, dSupportingX, dSupportingY, dSupportingZ
    
    'technically, we should use the post's logicalAxis Start.   But, we know there is no offset yet, so we use the FC.
    'dFCx,y,z is the vector from the the supporting object to the post's physical axis atart.
    dFCx = dFCx - dSupportingX
    dFCy = dFCy - dSupportingY
    dFCz = dFCz - dSupportingZ
    
    Set pIJAttrbs = oFC
    On Error Resume Next
    Set oAttrCollection = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset")
    Err.Clear
    If oAttrCollection Is Nothing Then
        Exit Sub
    End If
    
    ' might be nice to transform this vector to local coordinates, so they are in terms of the supporting member c/sys.
    Set oAttr = oAttrCollection.Item("XOffset")
    oAttr.Value = dFCx
    Set oAttr = oAttrCollection.Item("YOffset")
    oAttr.Value = dFCy
    Set oAttr = oAttrCollection.Item("ZOffset")
    oAttr.Value = dFCz
    
    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Exit Sub
End Sub

Private Sub IsPostPortSurfaceIntersectionClosed(oPost As ISPSMemberSystem, oPortGraphic As Object, ByRef bClosed As Boolean)
    Const METHOD = "IsPostPortSurfaceIntersectionClosed"
    On Error GoTo ErrorHandler

    Dim bAnyOpen As Boolean
    Dim eIntersectCode As Geom3dIntersectConstants
    Dim eForm As Geom3dCurveFormConstants
    Dim iCurve As IJCurve
    Dim iSurface As IJSurface
    Dim elesCurves As IJElements, elesSurfaces As IJElements
    Dim elesIntersectCurves As IJElements

    Dim oMF As SPSMemberFactory
    Dim elesPostSurfaces As IJElements
    Dim oMemberPart As ISPSMemberPartCommon
    Dim iSPSMemberFeatureServices As iSPSMemberFeatureServices
    
    bClosed = False
    
    If oPost Is Nothing Then
        Exit Sub
    ElseIf oPortGraphic Is Nothing Then
        Exit Sub
    End If
    If Not TypeOf oPortGraphic Is IJSurface Then
        Exit Sub
    End If

    Set oMF = New SPSMemberFactory
    Set iSPSMemberFeatureServices = oMF.CreateMemberFeatureServices
    
    Set oMemberPart = oPost.MemberPartAtEnd(SPSMemberAxisStart)
    iSPSMemberFeatureServices.CreateMemberDesignSurface oMemberPart, 0, 0, elesPostSurfaces
    Set oMF = Nothing
    Set oMemberPart = Nothing
    Set iSPSMemberFeatureServices = Nothing
    
    If elesPostSurfaces Is Nothing Then
        bAnyOpen = True
    Else
        bAnyOpen = False
        For Each iSurface In elesPostSurfaces
            iSurface.Intersect oPortGraphic, elesCurves, eIntersectCode
            If elesCurves Is Nothing Then
                bAnyOpen = True
            Else
                For Each iCurve In elesCurves
                    eForm = iCurve.Form
                    If eForm >= CURVE_FORM_CLOSED Then
                        bClosed = True
                    Else
                        bAnyOpen = True
                    End If
                Next iCurve
                Set elesCurves = Nothing
            End If
        Next iSurface
    End If
    
    If bAnyOpen Then
        bClosed = False
    Else
        bClosed = True
    End If

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Exit Sub
End Sub

' build a range box using the posts lower start or end points
    
Public Sub BuildRangeBox(eles As IJElements, ByRef gboxRange As GBox)
    Const METHOD = "BuildRangeBox"
    On Error GoTo ErrorHandler
    
    Dim iCurvePost As IJCurve
    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    Dim dEndX As Double, dEndY As Double, dEndZ As Double
    
    ' initialize rangebox to invalid value
    gboxRange.m_low.x = 1
    gboxRange.m_low.y = 1
    gboxRange.m_low.z = 1
    gboxRange.m_high.x = -1
    gboxRange.m_high.y = -1
    gboxRange.m_high.z = -1
    
    If eles Is Nothing Then
        Exit Sub
    ElseIf eles.Count = 0 Then
        Exit Sub
    End If
    
    ' now initialize rangebox to the first point
    Set iCurvePost = eles.Item(1)
    iCurvePost.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

    If dStartZ < dEndZ Then
        gboxRange.m_low.x = dStartX
        gboxRange.m_low.y = dStartY
        gboxRange.m_low.z = dStartZ
        gboxRange.m_high.x = dStartX
        gboxRange.m_high.y = dStartY
        gboxRange.m_high.z = dStartZ
    Else
        gboxRange.m_low.x = dEndX
        gboxRange.m_low.y = dEndY
        gboxRange.m_low.z = dEndZ
        gboxRange.m_high.x = dEndX
        gboxRange.m_high.y = dEndY
        gboxRange.m_high.z = dEndZ
    End If

    For Each iCurvePost In eles
        
        iCurvePost.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
        
        If dStartZ < dEndZ Then
            If dStartX < gboxRange.m_low.x Then gboxRange.m_low.x = dStartX
            If dStartY < gboxRange.m_low.y Then gboxRange.m_low.y = dStartY
            If dStartZ < gboxRange.m_low.z Then gboxRange.m_low.z = dStartZ
            If dStartX > gboxRange.m_high.x Then gboxRange.m_high.x = dStartX
            If dStartY > gboxRange.m_high.y Then gboxRange.m_high.y = dStartY
            If dStartZ > gboxRange.m_high.z Then gboxRange.m_high.z = dStartZ
        Else
            If dEndX < gboxRange.m_low.x Then gboxRange.m_low.x = dEndX
            If dEndY < gboxRange.m_low.y Then gboxRange.m_low.y = dEndY
            If dEndZ < gboxRange.m_low.z Then gboxRange.m_low.z = dEndZ
            If dEndX > gboxRange.m_high.x Then gboxRange.m_high.x = dEndX
            If dEndY > gboxRange.m_high.y Then gboxRange.m_high.y = dEndY
            If dEndZ > gboxRange.m_high.z Then gboxRange.m_high.z = dEndZ
        End If
    
    Next iCurvePost
    
    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Exit Sub
End Sub
