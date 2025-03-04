VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberEndSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : MemberEndSel.cls
'
'Author : RP
'
'Description :
'    Macro for selecting a 1 member connection based on input properties
'
'History:
'
'Feb 5 2004     RP  Modified selection rule to take
'                            in to account items in the RefColl
'1/7/05     JS Added code to the SelectorLogic to not choose
'                       a "MemberEnd" type assembly connection
'                       when the member is parallel to the plane - TR#67729
'09/27/06   RP IsMemberNearlyParallelToPlane() now handles curved members(TR#106105)
' 01-Feb-07 AS TR#109129 GenSurface AC should validate the geometric condition of supg and supd
'*************************************************************************************

Option Explicit

Const m_SelectorProgid As String = "SPSACMacros.MemberEndSel"
Const m_SelectorName As String = "SPSACMacros.MemberEndSel"
Const m_FamilyProgid As String = "SPSACMacros.Family"

Implements IJDUserSymbolServices


Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    
    Exit Sub
ErrorHandler:
    pIH.ReportError
End Sub


Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler

    
    Exit Sub
ErrorHandler:
    pQH.ReportError
End Sub


Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
  
    Dim oNavigate As ISPGNavigate
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oElvPlane As ISPGElevationPlane
    Dim oElm1 As Object
    Dim oRefColl As IJDReferencesCollection
    Dim oPort As IJPort
    Dim oConnectable As Object
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oPlane As IJPlane
    
    Set oStructConn = pSL.SmartOccurrence
    oStructConn.enumPorts colPorts
    If colPorts.count = 1 Then
        If TypeOf colPorts.Item(1) Is ISPSSplitAxisEndPort Then
            Set oRefColl = GetRefCollFromSmartOccurrence(pSL.SmartOccurrence)
            If Not oRefColl Is Nothing Then
                If oRefColl.IJDEditJDArgument.GetCount > 0 Then
                    'get the first item in the refcoll and check if an AsmConn
                    'can be selected based on it
                    Set oElm1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
                    If TypeOf oElm1 Is ISPGGridData Then
                        'oElm1 could be a gridline or a grid plane
                        If TypeOf oElm1 Is SPGGridLine Then
                            'then get the associated elevation plane and
                            Set oNavigate = oElm1
                            oNavigate.GetParent oElvPlane
                        ElseIf TypeOf oElm1 Is ISPGElevationPlane Then
                            Set oElvPlane = oElm1
                        End If
                        If Not oElvPlane Is Nothing Then
                            'check the type is Bottom Of Baseplate, if so
                            'create a BaseplateAsmConn
                            If oElvPlane.Type = 15 Then '15 is Bottom of Baseplate
                                Set oSuppedPort = GetEndPort(colPorts)
                                Set oPlane = oElvPlane
                                
                                If Not IsMemberNearlyParallelToPlane(oSuppedPort, oPlane) Then
                                    pSL.Add "BasePlateAsmConn"
                                End If
                            End If
                        End If
                    ElseIf TypeOf oElm1 Is IJPlane Then
                        Set oSuppedPort = GetEndPort(colPorts)
                        Set oPlane = oElm1
                        
                        If Not IsMemberNearlyParallelToPlane(oSuppedPort, oPlane) And IsSurfaceTypeAcceptableForSurfaceTrim(oSuppedPort, oPlane) Then
                            'for any plane that is not part of a grid system
                            'create a GeneralSurfaceAsmConn
                            pSL.Add "GeneralSurfaceAsmConn"
                        End If
                    
                    ElseIf IsSurfaceTypeAcceptableForSurfaceTrim(oSuppedPort, oElm1) Then
                        pSL.Add "GeneralSurfaceAsmConn"
                    End If
                    
                    
                End If
            End If
        End If
    ElseIf colPorts.count = 2 Then
        Set oSuppedPort = GetEndPort(colPorts)
        If Not oSuppedPort Is Nothing Then
            If oSuppedPort Is colPorts.Item(1) Then
                Set oPort = colPorts.Item(2)
            Else
                Set oPort = colPorts.Item(1)
            End If
        
            If TypeOf oPort Is IJPlane Then
                Set oPlane = oPort
                
                If Not IsMemberNearlyParallelToPlane(oSuppedPort, oPlane) Then
                    Set oConnectable = oPort.Connectable
                    If TypeOf oConnectable Is SPSSlabEntity Then
                        pSL.Add "BasePlateAsmConn"
                    Else
                        pSL.Add "GeneralSurfaceAsmConn"
                    End If
                End If
            ElseIf IsSurfaceTypeAcceptableForSurfaceTrim(oSuppedPort, oPort) Then
                pSL.Add "GeneralSurfaceAsmConn"
            End If
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


Private Function IsMemberNearlyParallelToPlane(oPort As ISPSSplitAxisPort, oPlane As IJPlane) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    Dim oMembDir As New DVector
    Dim nx As Double, ny As Double, nz As Double
    Dim oZDir As New DVector
    Dim cosA As Double
    Dim oPart As ISPSMemberPartPrismatic
    Dim oMat As IJDT4x4
    
    IsMemberNearlyParallelToPlane = False ' Assume they are not parallel
    
    Set oPart = oPort.Part
    
    oPart.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    
    'get the tangent to member axis at end point
    If oPort.portIndex = SPSMemberAxisStart Then
        oPart.Rotation.GetTransformAtPosition sX, sY, sZ, oMat, Nothing
    Else
        oPart.Rotation.GetTransformAtPosition eX, eY, eZ, oMat, Nothing
    End If
    
    'get x vector from matrix
    oMembDir.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    oMembDir.length = 1 ' Normalize the vector
    oPlane.GetNormal nx, ny, nz
    oZDir.Set nx, ny, nz
    cosA = oZDir.Dot(oMembDir) 'A is the angle between plane's normal and member axis
    If Abs(cosA) < 0.1 Then     ' member is greater than 6 degrees of parallel to surface
        IsMemberNearlyParallelToPlane = True
    End If
    
    Exit Function

ErrorHandler:
End Function
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************







