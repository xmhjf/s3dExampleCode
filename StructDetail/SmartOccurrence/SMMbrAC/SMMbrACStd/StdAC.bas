Attribute VB_Name = "StdAC"
'*******************************************************************
'
'Copyright (C) 2014-15 Intergraph Corporation. All rights reserved.
'
'File : StdAC.bas
'
'Author : Alligators
'History :
'   25/Feb/2015   - Addedd new method Set_FlipAttributeValue
'   12/May/2015 GH- CR-260982 - Added New method UpdatePCWithNewFilterProgID()
'   07/Aug/2015 NK- CR-266151 - Create StdAC for OutsideAndOutside-NoEdge case
'   30/11/2015    dsmamidi    DI-284513  Character limit restricts definition of Standard ACs
'*******************************************************************

Option Explicit

Public Const INPUT_PORT1FACE = "LogicalFace"
Public Const INPUT_PORT2EDGE = "Support1"
Public Const INPUT_PORT3EDGE = "Support2"

Private Const MODULE As String = "StructDetail\SmartOccurrence\SMMbrAC\SMMbrACStd\"
Public Const m_sStdACProjectName As String = CUSTOMERID + "MbrACStd"
Public Const m_sStdACProjectPath As String = "S:\StructDetail\SmartOccurrence\SMMbrAC\" + m_sStdACProjectName + "\"

Public Const IID_IJProxyMember = "{C1FD8CAF-9767-11D1-9425-0060973D4777}"


Public Sub Set_FlipAttributeValue(oAppConnection As IJAppConnection)

    Const MT = "Set_FlipAttributeValue"
    On Error GoTo ErrorHandler
    Dim sAttributeVal As String
    Dim oACDef As New ACDef
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim sMsg As String

    Dim ACPorts As IJElements
    oAppConnection.enumPorts ACPorts
    Set oBoundedPort = ACPorts.Item(1)
    Set oBoundingPort = ACPorts.Item(2)
    
    'Get the Primary PrimaryCriteria Attribute value
    sAttributeVal = oACDef.GetStandardACAttribute(oAppConnection, "PrimaryCriteria")
    
    Dim oSMOCC As IJSmartOccurrence
    Set oSMOCC = oAppConnection
    
    'Default set to False
    Set_AttributeValue oAppConnection, "IJUAFlipAttribute", "FlipPrimaryandSecondary", False
    
    Dim bFlip As Boolean
    bFlip = False
    Dim FlipPrimaryAndSec As Double
    
    If Len(sAttributeVal) Then
        
        Dim sSplitAttribute() As String
        Dim sValue As String
        Dim sCriteira As String
        Dim sValues() As String
        
        sSplitAttribute = Split(Trim$(sAttributeVal), ":")
        
        If UBound(sSplitAttribute) = 1 Then
        
            sCriteira = sSplitAttribute(0)
            sValues = Split(sSplitAttribute(1), ",")
                    
            Select Case LCase(Trim$(sCriteira))
                                   
                Case LCase("CS") ' Check Cross section Type
                    
                    Dim oBoundedPart As ISPSMemberPartCommon
                    Set oBoundedPart = oBoundedPort.Connectable
                    
                    Dim oBoundingPart As ISPSMemberPartCommon
                    Set oBoundingPart = oBoundingPort.Connectable
                    
                    Dim sBoundedCStype As String
                    Dim sBoundingCStype As String
                    Dim iCount As Integer
                    
                    'Get Cross sections
                    sBoundedCStype = oBoundedPart.CrossSection.sectionType
                    sBoundingCStype = oBoundingPart.CrossSection.sectionType
                                        
                    For iCount = 0 To (UBound(sValues))
                        'Check both Bounded and Bounding Cross sections
                        If StrComp(sBoundingCStype, Trim$(sValues(iCount)), vbTextCompare) = 0 And _
                            Not StrComp(sBoundedCStype, Trim$(sValues(iCount)), vbTextCompare) = 0 Then
                            bFlip = True
                            Exit For
                        'If bounded is satisfied then no need to swap
                        ElseIf StrComp(sBoundedCStype, Trim$(sValues(iCount)), vbTextCompare) = 0 Then
                            bFlip = False
                            Exit For
                        End If
                    Next
                                            
                Case LCase("FC") ' Check Frame connection
                                        
                    If InStr(1, Trim$(sValues(0)), "U", vbTextCompare) Then
                        bFlip = True
                    End If
                Case LCase("BC") ' Check Bounding Cases
                    Dim eBoundingCase As eMemberBoundingCase
                    eBoundingCase = GetMemberBoundingCase(oAppConnection)
                    
                    If sValues(0) = "F" Then 'To Center Case
                        If eBoundingCase = Center Then
                            'No Action is needed
                        Else
                            Dim oAttributes As IJDAttributes
                            Set oAttributes = oAppConnection
                            
                            On Error Resume Next
                            
                            Dim vValue As Variant
                            
                            'Set the attribute value to check the bounding case from other side
                            vValue = oAttributes.CollectionOfAttributes("IJUAFlipAttribute").Item("FlipPrimaryandSecondary").Value
                            oAttributes.CollectionOfAttributes("IJUAFlipAttribute").Item("FlipPrimaryandSecondary").Value = True
                            eBoundingCase = GetMemberBoundingCase(oAppConnection)
                            
                            If eBoundingCase = Center Then 'To Center Case
                                bFlip = True
                            End If
                            
                            'Set back the attribute value to initial value
                            oAttributes.CollectionOfAttributes("IJUAFlipAttribute").Item("FlipPrimaryandSecondary").Value = vValue
                            On Error GoTo ErrorHandler
                        End If
                    End If
            End Select
            
        End If
    End If
        
    Dim StrFlipAnswer As String
    'Get selector question answer
    GetSelectorAnswer oAppConnection, "Flip Primary and Secondary", StrFlipAnswer
    
    If StrFlipAnswer = gsYes Then
    
        If bFlip Then
            bFlip = False
        Else
            bFlip = True
        End If
    Else
        'No action required
    End If
    
    'Finally Set the attribute value
    Set_AttributeValue oAppConnection, "IJUAFlipAttribute", "FlipPrimaryandSecondary", bFlip
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT, sMsg
End Sub

Public Sub UpdatePCWithNewFilterProgID(oPhysicalConn As IJSmartOccurrence, strNewFilter As String)

  Const METHOD = "::UpdatePCWithNewFilterProgID"
  On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Update the PC if needed"
    
    'Get current PC FilterProgID value
    Dim strOldFilter As String
    strOldFilter = GetCustomAttribute(oPhysicalConn, "IJUASelectionFilter", "FilterProgID")

    Dim oCM As New CustomMethods
    'Update PCs
    If Not StrComp(strOldFilter, strNewFilter, vbTextCompare) = 0 Then
        oCM.AddPCFilter oPhysicalConn, strNewFilter
        oPhysicalConn.Update
    End If
    
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
'******************************************************************************************
' Method:
' GetSnipeOutsideFaceParameters
'
' Description: Get the parameters for achieving top and bottom snipe cuts for XX0003 AC (outside and outside case StdAc)
' *******************************************************************************************
Public Sub GetSnipeOutsideFaceParameters(oEndCut As IJStructFeature, _
                                         isOutsideBottom As Boolean, _
                                         dNose As Double, _
                                         dSlope As Double, _
                                         dSlopeOutside As Double, _
                                         dFlangeThickness As Double)
 Const METHOD = "::GetSnipeOutsideFaceParameters"
  On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "parameters are not filled properly"
    
    Dim oSDOWebCut As New StructDetailObjects.WebCut
    If oEndCut.get_StructFeatureType = SF_WebCut Then
        Set oSDOWebCut.object = oEndCut
    Else
        Exit Sub
    End If
    
    ' ---------------------------------------------------------------------
    ' If there is a flange outside the bounding geometry, get the thickness
    ' ---------------------------------------------------------------------
    dFlangeThickness = 0.0001
    
    If (isOutsideBottom And HasBottomFlange(oSDOWebCut.Bounded)) Or (Not isOutsideBottom And HasTopFlange(oSDOWebCut.Bounded)) Then
    
        Dim oSDOBoundedPart As New StructDetailObjects.MemberPart
        Set oSDOBoundedPart.object = oSDOWebCut.Bounded
        
        dFlangeThickness = oSDOBoundedPart.flangeThickness
    End If
    
    ' -------------------------------------------------
    ' Determine how much material lies outside the face
    ' -------------------------------------------------
    Dim cWLOrTop As ConnectedEdgeInfo
    Dim cWROrBtm As ConnectedEdgeInfo
    Dim cTFIOrFL As ConnectedEdgeInfo
    Dim cBFIOrFR As ConnectedEdgeInfo

    Dim oMeasurements As Collection
    Set oMeasurements = New Collection
    
    GetConnectedEdgeInfo oEndCut, _
                         oSDOWebCut.BoundedPort, _
                         oSDOWebCut.BoundingPort, _
                         cWLOrTop, _
                         cWROrBtm, _
                         cTFIOrFL, _
                         cBFIOrFR, _
                         oMeasurements
    
    Dim dlengthOutside As Double
    dlengthOutside = 0
    
    If isOutsideBottom Then
        If Not KeyExists("DimPt23ToBottom", oMeasurements) Then
            Exit Sub
        Else
            dlengthOutside = oMeasurements.Item("DimPt23ToBottom")
        End If
    Else
        If Not KeyExists("DimPt15ToTop", oMeasurements) Then
            Exit Sub
        Else
            dlengthOutside = oMeasurements.Item("DimPt15ToTop")
        End If
    End If
        
    ' ---------------------------------------------------------------
    ' If there is less than 10mm of clearance there shall be no snipe
    ' If the nose is less than the flange thickness or less than 5mm, there shall be no snipe
    ' ---------------------------------------------------------------
    Dim chamferIncrements As Integer
    chamferIncrements = Int((dlengthOutside - 0.01) * 1000 / 5)
    
    If (dlengthOutside - dFlangeThickness < 0.01) Or (chamferIncrements * 0.005 < dFlangeThickness) Then
        ' Set nose to be very small and the slopes to 90°
        dNose = 0.0001
        dSlope = Atn(1) * 2
        dSlopeOutside = dSlope
        Exit Sub
    End If
    ' -------------------------------------------------------------------------------------------
    ' Otherwise, set the angles to 45° and the nose to get an even increment of 5mm for the snipe
    ' -------------------------------------------------------------------------------------------
    If IsTubularMember(oSDOWebCut.Bounded) Or IsRectangularMember(oSDOWebCut.Bounded) Then
        dNose = 0.0001
    Else
        dNose = dlengthOutside - (chamferIncrements * 0.005)
    End If
    dSlope = Atn(1)
    dSlopeOutside = dSlope
Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'---------------------------------------------------------------------------------------------------------------
'Function: GetSymbolShare
'
'DESCRIPTION: Get the path of SymbolShare location
'
'Input:
'
'Output:
'GetSymbolShare,string
'---------------------------------------------------------------------------------------------------------------
Public Function GetSymbolShare() As String
    On Error GoTo ErrorHandler
    Const METHOD As String = "GetSymbolShare"

    Dim objJContext As IJContext
    Set objJContext = GetJContext
    If Not objJContext Is Nothing Then
        GetSymbolShare = objJContext.GetVariable("OLE_SERVER")
    End If

    Exit Function
    
ErrorHandler:
Err.Raise LogError(Err, MODULE, METHOD).Number
End Function
