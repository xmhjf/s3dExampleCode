VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberByMemberSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : MemberByMemberSel.cls
'
'Author : RP
'
'Description :
'    Macro for selecting a 2 member connection based on input properties
'
'History:
'
' 04 -14-04  RP   Added check for proper orientation before selecting GussetPlatAsmConn
' 06/13/06  RP   Changes due to impact from curved members(DI#84001)
' 12/11/06  RP   Do not select gussetplate  connection when supported member has I-shaped
'                section
' 18/06/14  knukala CR-CP-223345  Rules should not replace root selector inputs
'***************************************************************************************

Option Explicit

Const m_SelectorProgid As String = "SPSACMacros.MemberByMemberSel"
Const m_SelectorName As String = "SPSACMacros.MemberByMemberSel"

Implements IJDUserSymbolServices


Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add input description below
    
    Exit Sub
ErrorHandler:
    pIH.ReportError
End Sub


Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    
    Exit Sub
ErrorHandler:
    pQH.ReportError
End Sub


Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppingPort As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppingPart As ISPSMemberPartPrismatic
    Dim oConnPoint As IJPoint
    Dim x As Double, y As Double, z As Double
    Dim oSuppingCurve As IJCurve, oSuppedCurve As IJCurve
    Dim oSuppingVec As New DVector, oSuppedVec As New DVector
    Dim oConnPos As IJDPosition
    Dim oMat As IJDT4x4
    
    Set oStructConn = pSL.SmartOccurrence
    oStructConn.enumPorts colPorts
    Set oSuppedPort = GetEndPort(colPorts)
    If oSuppedPort Is colPorts.Item(1) Then
        Set oSuppingPort = colPorts.Item(2)
    Else
        Set oSuppingPort = colPorts.Item(1)
    End If
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part
    
    
    'get the connection point on supping member. It is possible that axes don't touch each other
    ' so we need to find a position which is closest to supped end
    Set oConnPos = GetConnectionPositionOnSupping(oSuppingPart, oSuppedPart, oSuppedPort.portIndex)
        
    If oConnPos Is Nothing Then
        Exit Sub
    End If
    
    'get the curve segment at this position. The axis may be made up of many curve segments
    Set oSuppingCurve = GetAxisCurveAtPosition(oConnPos.x, oConnPos.y, oConnPos.z, oSuppingPart)
    oSuppingPart.Rotation.GetTransformAtPosition oConnPos.x, oConnPos.y, oConnPos.z, oMat, Nothing
    oSuppingVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    
    Set oConnPoint = oSuppedPort.Port.Geometry 'this better be a point. The suppedPort always has a point geometry
    oConnPoint.GetPoint x, y, z
    
    'get the curve segment at this position. The axis may be made up of many curve segments
    Set oSuppedCurve = GetAxisCurveAtPosition(x, y, z, oSuppedPart)
    oSuppedPart.Rotation.GetTransformAtPosition x, y, z, oMat, Nothing
    oSuppedVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    
    If IsMemberAxesColinearAndEndToEnd(oSuppedCurve, oSuppingCurve) Then
        Dim oPos1 As New DPosition, oPos2 As New DPosition
        Dim x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double
        'it is possible that the suppincurve is an interior segment. We don't want a splice in that case
        
        oSuppingPart.Axis.EndPoints x1, y1, z1, x2, y2, z2
        oPos1.Set x1, y1, z1
        oPos2.Set x2, y2, z2
        'check start of curve segment is start of memb axis or end curve segment is end of memb axis
        If (oPos1.DistPt(oConnPos) < distTol) Or (oPos2.DistPt(oConnPos) < distTol) Then
            pSL.Add "SpliceAsmConn"
        End If
    ElseIf (Not AreVectorsParallel(oSuppingVec, oSuppedVec)) And (Not IsSuppedSeatedOnSupping(oSuppedPart, oSuppingPart)) Then
        
        'check if we can apply a miter connection
        
        'first chek if types are same
        If oSuppingPart.MemberType.TypeCategory = oSuppedPart.MemberType.TypeCategory Then
            'miter only if the sections are the same
            If oSuppingPart.CrossSection.SectionName = oSuppedPart.CrossSection.SectionName Then
                
                Dim oZSupped As IJDVector, oZSupping As IJDVector
                Dim oYSupped As IJDVector, oYSupping As IJDVector
                Dim cosZ As Double, cosY As Double
                Dim bZParallel As Boolean, bYParallel As Boolean
                
                Set oZSupped = GetZVector(oSuppedPart)
                Set oZSupping = GetZVector(oSuppingPart)
                
                Set oYSupped = GetYVector(oSuppedPart)
                Set oYSupping = GetYVector(oSuppingPart)

                cosZ = oZSupped.Dot(oZSupping)
                cosY = oYSupped.Dot(oYSupping)
                
                If Abs(cosZ - 1) < angleTol Then
                    bZParallel = True
                ElseIf Abs(cosY - 1) < angleTol Then
                    bYParallel = True
                End If
                
                'Either the z vecs of supped and supping or the y vecs of supped and supping should be parallel
                If bZParallel Or bYParallel Then
                    
                    Dim bMirrorSupped As Boolean, bMirrorSupping As Boolean, bMirror As Boolean
                    Dim portIdSupped As SPSMemberAxisPortIndex, portIdSupping As SPSMemberAxisPortIndex
                    Dim oSuppedPortPos As IJPoint, oSuppingPortPos0 As IJPoint, oSuppingPortPos1 As IJPoint
                    Dim oSuppedSys As ISPSMemberSystem, oSuppingSys As ISPSMemberSystem
                    Dim oJoint As ISPSAxisJoint, oSuppedJoint As ISPSAxisJoint, oSuppingJoint As ISPSAxisJoint
                    Dim oCol As IJElements
                    Dim obj1 As Object, obj2 As Object
                    Dim bSelectMiter As Boolean
                    Dim oLine1 As IJLine, oLine2 As IJLine
                    
                    bMirrorSupped = oSuppedPart.Rotation.Mirror
                    bMirrorSupping = oSuppingPart.Rotation.Mirror
                    bMirror = bMirrorSupped Xor bMirrorSupping ' true when only one is mirrored

                    portIdSupped = oSuppedPort.portIndex
                    portIdSupping = oSuppingPort.portIndex
                    
                    'get coordinates of supped port
                    Set oSuppedPortPos = oSuppedPort.Port.Geometry
                    Set oSuppingPortPos0 = oSuppingPart.PointAtEnd(SPSMemberAxisStart)
                    Set oSuppingPortPos1 = oSuppingPart.PointAtEnd(SPSMemberAxisEnd)
                    If portIdSupping = SPSMemberAxisAlong Then
                        If oSuppedPortPos.IsEqual(oSuppingPortPos0) Then
                            portIdSupping = SPSMemberAxisStart
                        ElseIf oSuppedPortPos.IsEqual(oSuppingPortPos1) Then
                            portIdSupping = SPSMemberAxisEnd
                        End If
                    End If
                    'check if coordinate is at the end of supping part
                    If portIdSupping <> SPSMemberAxisAlong Then
                        If (TypeOf oSuppedCurve Is IJLine) And (TypeOf oSuppingCurve Is IJLine) Then
                            Set oLine1 = oSuppedCurve
                            Set oLine2 = oSuppingCurve
                        
                             If IsMemberAxesEndToEnd(oLine1, oLine2) Then
                            'this means that supping and supped have same orientation
                                'get member systems
                                Set oSuppedSys = oSuppedPart.MemberSystem
                                Set oSuppingSys = oSuppingPart.MemberSystem
                                If (oSuppedSys Is Nothing) And (oSuppingSys Is Nothing) Then
                                    'two stand alone member parts connected together at the ends
                                    'may occur in equipment foundation
                                    'select miter for this case
                                    bSelectMiter = True
                                Else
                                    'atleast one is a member system
                                    'test if both are memb systems
                                    If (Not oSuppedSys Is Nothing) And (Not oSuppingSys Is Nothing) Then
                                        oSuppedPortPos.GetPoint x, y, z
                                        'check if the pos is at the ends of the member systems
                                        Set oSuppedJoint = oSuppedSys.JointAtEndPosition(x, y, z)
                                        Set oSuppingJoint = oSuppingSys.JointAtEndPosition(x, y, z)
                                        If (Not oSuppedJoint Is Nothing) And (Not oSuppingJoint Is Nothing) Then
                                            'make sure that another membsys is not point on to supped membsys at this location
                                            Set oCol = oSuppedSys.PointOnJointsAtPosition(x, y, z)
                                            If oCol.count = 0 Then
                                                'get point on joints at supping end
                                                Set oCol = oSuppingSys.PointOnJointsAtPosition(x, y, z)
                                                If oCol.count = 0 Then
                                                    If oSuppedJoint Is oSuppingJoint Then
                                                        'suppedsys and suppingsys share the joints at ends
                                                        'two member systems connected together at the ends
                                                        Set oCol = oSuppedJoint.EndMemberSystems
                                                        If oCol.count = 2 Then
                                                            oSuppedJoint.GetPointOn obj1, obj2
                                                            'check only for obj1 now as point on to two members, two joints
                                                            'or a member and joint is not supported
                                                            If Not obj1 Is Nothing Then
                                                                'check if it is point on to a third memb sys or to a joint. Either case there are more
                                                                'than two members here so we cannot select miter
                                                                If (Not TypeOf obj1 Is ISPSMemberSystemLinear) And (Not TypeOf obj1 Is ISPSAxisJoint) Then
                                                                    'select miter for this case
                                                                    bSelectMiter = True
                                                                End If
                                                            Else
                                                                'the joint is not point on to anything
                                                                'select miter for this case
                                                                bSelectMiter = True
                                                            End If
                                                        End If
                                                    End If
                                                ElseIf oCol.count = 1 Then
                                                    If oSuppedJoint Is oCol.Item(1) Then
                                                        'suppedsys is point on to suppingsys at the end
                                                        'make only one   memberend at this joint and that is of oSuppedsys
                                                        Set oCol = oSuppedJoint.EndMemberSystems
                                                        If oCol.count = 1 Then
                                                            'make sure no joint is point on to oSuppedJoint
                                                            'this is the case where a brace is connected to beam and column
                                                            Set oCol = oSuppedJoint.PointOnJoints
                                                            If oCol.count = 0 Then
                                                                bSelectMiter = True
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If bSelectMiter Then
            pSL.Add "MitreAsmConn"
        ElseIf (oSuppedPart.MemberType.TypeCategory = 3) And (oSuppingPart.MemberType.TypeCategory <> 3) Then ' Brace connected to beam, column
            Dim strSectionType As String
            strSectionType = oSuppedPart.CrossSection.SectionType
            If strSectionType <> "W" And strSectionType <> "S" And strSectionType <> "HP" And strSectionType <> "M" Then
                'do not select gusset plate connection for a I-shaped supported member
                If CheckForGussetPlateAsmConn(oSuppedPart, oSuppingPart) Then
                    pSL.Add "GussetPlateAsmConn"
                Else
                    pSL.Add "FittedAsmConn"
                End If
            Else
                pSL.Add "FittedAsmConn"
            End If
        Else
            pSL.Add "FittedAsmConn"
        End If
    End If
    Exit Sub
ErrorHandler:
    pSL.ReportError
End Sub
  
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

' ** End CM **



' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.Definition = pSelector
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub


Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function


Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub


Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function


Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************







