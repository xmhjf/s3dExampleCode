VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PhysConnParmCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "PhysConnParmCM"

Private sError As String
Private sMETHOD As String

Public Sub GetFilletWeldParameters(oPRL As IJDParameterLogic, _
                                   strStandardName As String, _
                                   bUseThinnerPart As Boolean, _
                                   bAdjustNose As Boolean, _
                                   dNoseOrientationAngle As Double, _
                                   lNoseMethod As Long, _
                                   dNose As Double, _
                                   strRefPartName As String, _
                                   dMoldedFillet As Double, _
                                   dAntiMoldedFillet As Double, _
                                   lFilletMeasureMethod As Long, _
                                   strReferenceSide As String, _
                                   lCategory As Long, _
                                   bRefIsMolded As Boolean)

    On Error GoTo ErrorHandler
  
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Set oPhysConn.object = oPRL.SmartOccurrence
    
    ' ------------------------------------------
    ' Get some required data via a helper method
    ' ------------------------------------------
    Dim sCategory As String
    Dim sBevelMethod As String
    Dim sClassSociety As String
    Dim dThickness1 As Double
    Dim dThickness2 As Double
     
    Get_ParameterRuleData oPRL, strStandardName, sClassSociety, sCategory, sBevelMethod, dThickness1, dThickness2
    
    ' -------------------------------------------------------------
    ' Calculate NoseOrientationAngle, depending on the bevel method
    ' -------------------------------------------------------------
    If bAdjustNose Then
        Select Case sBevelMethod
            Case gsConstant
                dNoseOrientationAngle = PI - oPhysConn.MountingAngle
                lNoseMethod = 65536
        
            Case gsVarying
                dNoseOrientationAngle = 0
                lNoseMethod = 65537
        End Select
    Else
        Select Case sBevelMethod
           Case gsConstant
               dNoseOrientationAngle = 1.570796
               lNoseMethod = 65536
           Case gsVarying
               Dim dAngle As Double
               dAngle = Abs(1.570796 - (oPhysConn.MountingAngle - TOL))
               dNoseOrientationAngle = dAngle
               lNoseMethod = 65537
        End Select
    End If
    
    ' ------------------
    ' Set the nose value
    ' ------------------
    dNose = dThickness1
     
    ' ------------------
    ' Get the part names
    ' ------------------
    Dim sPartName1 As String
    Dim oNamedItem1 As IJNamedItem
    Set oNamedItem1 = oPhysConn.ConnectedObject1
    sPartName1 = oNamedItem1.Name
    
    strRefPartName = sPartName1
    
    ' ----------------------------------------------------
    ' Calculate the fillet size and set measurement method
    ' ----------------------------------------------------
    Dim dFillet As Double
    
    If dThickness1 > dThickness2 And bUseThinnerPart Then
        dFillet = dThickness2 * 0.2
    Else
        dFillet = dThickness1 * 0.2
    End If
    
    dMoldedFillet = dFillet
    dAntiMoldedFillet = dFillet
    lFilletMeasureMethod = 65536
    
    ' ----------------------
    ' Set the reference side
    ' ----------------------
    Dim sRefSide As String
    sRefSide = GetRefSide(oPhysConn.object)

    If sRefSide = "Base" Then
        strReferenceSide = "molded"
        bRefIsMolded = True
    ElseIf sRefSide = "Offset" Then
        strReferenceSide = "antimolded"
        bRefIsMolded = False
    Else
        strReferenceSide = sRefSide
        bRefIsMolded = True
    End If
    
    ' ----------------
    ' Set the category
    ' ----------------
    If sCategory = gsNormal Then
        lCategory = 65537
    ElseIf sCategory = gsDeep Then
        lCategory = 65538
    ElseIf sCategory = gsFull Then
        lCategory = 65539
    ElseIf sCategory = gsChain Then
        lCategory = 65540
    ElseIf sCategory = "Staggered" Then
        lCategory = 65541
    ElseIf sCategory = "OneSidedBevel" Then
        lCategory = 65542
    ElseIf sCategory = "TwoSidedBevel" Then
        lCategory = 65543
    Else
        lCategory = 65536
    End If
     
    Exit Sub

ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub

Public Function GetWebFillet(oPart As Object) As Double

    On Error GoTo ErrorHandler
    
    Dim dWebThickness As Double
    
    If TypeOf oPart Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As New StructDetailObjects.MemberPart
        Set oBoundedMemberPart.object = oPart
        dWebThickness = oBoundedMemberPart.webthickness
    ElseIf TypeOf oPart Is IJProfile Then
        Dim oBoundedProfilePart As New StructDetailObjects.ProfilePart
        Set oBoundedProfilePart.object = oPart
        dWebThickness = oBoundedProfilePart.webthickness
    End If

    If dWebThickness <= 0.008 Then
        GetWebFillet = 0.005
    ElseIf dWebThickness <= 0.01 Then
        GetWebFillet = 0.006
    ElseIf dWebThickness <= 0.012 Then
        GetWebFillet = 0.008
    ElseIf dWebThickness <= 0.015 Then
        GetWebFillet = 0.01
    ElseIf dWebThickness <= 0.02 Then
        GetWebFillet = 0.012
    ElseIf dWebThickness <= 0.025 Then
        GetWebFillet = 0.015
    Else
        GetWebFillet = 0.02
    End If
    
    Exit Function
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Function

Public Function GetFlangeFillet(oPart As Object) As Double

    On Error GoTo ErrorHandler
    
    Dim dFlangeThickness As Double
    
    If TypeOf oPart Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As New StructDetailObjects.MemberPart
        Set oBoundedMemberPart.object = oPart
        dFlangeThickness = oBoundedMemberPart.flangeThickness
    ElseIf TypeOf oPart Is IJProfile Then
        Dim oBoundedProfilePart As New StructDetailObjects.ProfilePart
        Set oBoundedProfilePart.object = oPart
        dFlangeThickness = oBoundedProfilePart.flangeThickness
    End If
    
    If dFlangeThickness <= 0.008 Then
        GetFlangeFillet = 0.008
    ElseIf dFlangeThickness <= 0.01 Then
        GetFlangeFillet = 0.01
    ElseIf dFlangeThickness <= 0.012 Then
        GetFlangeFillet = 0.012
    ElseIf dFlangeThickness <= 0.015 Then
        GetFlangeFillet = 0.015
    ElseIf dFlangeThickness <= 0.02 Then
        GetFlangeFillet = 0.02
    ElseIf dFlangeThickness <= 0.025 Then
        GetFlangeFillet = 0.025
    Else
        GetFlangeFillet = 0.03
    End If
    
    Exit Function
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Function

Public Function UsePreSelection(oSL As IJDSelectorLogic) As Boolean

    UsePreSelection = False
    
    ' ----------------------------------------------
    ' Use filter in place of selection logic, if set
    ' ----------------------------------------------
    Dim strFilter As String
    strFilter = GetCustomAttribute(oSL.SmartOccurrence, "IJUASelectionFilter", "FilterProgID")
    
    If strFilter <> vbNullString Then
        
        UsePreSelection = True
        
        ' ---------------------------------------------------------------------------------------------
        ' If the filter does not have the form Text.Text, treat as comma-delimited set of allowed items
        ' ---------------------------------------------------------------------------------------------
        Dim lPosition As Long
        
        lPosition = InStr(1, strFilter, ".", vbTextCompare)
        
        Dim strItem As String
            
        If lPosition = 0 Then
           
            lPosition = 1
            
            Do While (lPosition > 0)
                
                lPosition = InStr(1, strFilter, ",", vbTextCompare)
                
                If lPosition > 0 Then
                    strItem = Left(strFilter, lPosition - 1)
                    strFilter = Right(strFilter, Len(strFilter) - lPosition)
                    oSL.Add strItem
                Else
                    oSL.Add strFilter
                End If
            Loop
            
        ' --------------------------------------------------------------
        ' If the filter has the form Text.Text, treat as a filter progID
        ' --------------------------------------------------------------
        Else
            ' --------------------------
            ' Create and call the filter
            ' --------------------------
            Dim oFilter As Object
            Set oFilter = SP3DCreateObject(strFilter)
    
            Dim oAllowedItems As Collection
            Set oAllowedItems = oFilter.GetAllowedItems(oSL.SmartOccurrence)
    
            ' -------------------------------------
            ' Add all the strings that are returned
            ' -------------------------------------
            Dim i As Long
        
            For i = 1 To oAllowedItems.Count
                strItem = oAllowedItems.Item(i)
                oSL.Add strItem
            Next i
        End If
    End If

End Function

