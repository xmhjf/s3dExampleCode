VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SampleRule3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'   SampleRule3 NamingRule
'
'Description:
'History :
'   Siva   20th July  2012      Added new class
'************************************************************************************************************
Option Explicit
Implements IJMfgTemplateNamingRule

Private Const MODULE = "TemplateNamingRule.SampleRule2"

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize"
On Error GoTo ErrorHandler

Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJMfgTemplateNamingRule_CreateTemplateName(ByVal pPlatePart As Object, ByVal pTemplateSet As Object)
Const METHOD = "IJMfgTemplateNamingRule_CreateTemplateName"
On Error GoTo ErrorHandler

    Dim oMfgTemplateSet As IJDMfgTemplateSet
    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Dim oTemplates As IJElements
    Dim oTemplate As IJMfgTemplate
    Dim oIJNamedItem As IJNamedItem
    Dim strPosition As String, strDirection As String
    Dim nNumTemplates As Integer, nCount As Integer, lCode As Long
    
    Dim oSPGMiddleHelper As ISPGMiddleHelper
    Set oSPGMiddleHelper = New SPGMiddleHelper
    
    Dim oFrameSystem As IHFrameSystem
    Dim oMfgFrameSys As IJDMfgFrameSystem
    Set oMfgFrameSys = pTemplateSet
    If Not oMfgFrameSys Is Nothing Then
        Set oFrameSystem = oMfgFrameSys.FrameSysParent
    End If
    
    lCode = 65 'A - for even templates
    
    Set oMfgTemplateSet = pTemplateSet
    Set oProcessSettings = oMfgTemplateSet.GetProcessSettings
    
    strDirection = oProcessSettings.TemplateDirection
    strPosition = oProcessSettings.TemplatePosition
    
    Set oTemplates = oMfgTemplateSet.GetTemplatesSorted(GirthBetweenBaseCtlPointAndAftPoint)
    nNumTemplates = oTemplates.Count
    
    'first label all templates
    For nCount = 1 To nNumTemplates
        Set oTemplate = oTemplates.Item(nCount)
        Set oIJNamedItem = oTemplate
        
        If nCount = 1 Then
            ' Template generated on the aft/ fore seam or Up/Lower seam should named AFT/FORE, UP/LO
            If strDirection = "Transversal" Then 'AFT/LO name
                oIJNamedItem.Name = "AFT"
            ElseIf strDirection = "Longitudinal" Or strDirection = "Waterline" Then
                oIJNamedItem.Name = "LO"
            End If
            
        ElseIf nCount = nNumTemplates Then
            ' Template generated on the aft/ fore seam or Up/Lower seam should named AFT/FORE, UP/LO
            If strDirection = "Transversal" Then 'FORE/UP name
                oIJNamedItem.Name = "FORE"
            ElseIf strDirection = "Longitudinal" Or strDirection = "Waterline" Then
                oIJNamedItem.Name = "UP"
            End If
            
        Else
            ' Templates at Frames/even positions
            Dim strName As String
            If Not (strPosition = "EvenOnly" Or strPosition = "EvenAndEdges") Then
                
                Dim oTemplateReport As IJMfgTemplateReport
                Set oTemplateReport = oTemplate
            
                Dim oBCPos As IJDPosition
                Set oBCPos = oTemplateReport.GetPoint(BaseControlPoint)
                
                Dim dRootX As Double, dRootY As Double, dRootZ As Double, dLocation As Double
                oBCPos.Get dRootX, dRootY, dRootZ
                
                Dim eAxis As AxisType
                Dim strAxis As String
                
                Select Case strDirection
                  Case "Transversal"
                      eAxis = X
                      strAxis = "X"
                      dLocation = dRootX
                      strName = "F"
                  Case "Longitudinal"
                      eAxis = Y
                      strAxis = "Y"
                      dLocation = dRootY
                      strName = "L"
                  Case "Waterline"
                      eAxis = Z
                      strAxis = "Z"
                      dLocation = dRootZ
                      strName = "W"
                End Select
                
                If IsSketchTemplate(oTemplate) = False Then
                    
                    If Not oFrameSystem Is Nothing Then
                        Dim oGridPlane As IHFrame
                        oSPGMiddleHelper.IsPlaneAtLocation oFrameSystem, eAxis, Round(dLocation, 3), oGridPlane
                        
                        oIJNamedItem.Name = strName & Mid(oGridPlane.Name, 2)
                        Set oGridPlane = Nothing
                    End If
                Else 'sketch template
                    
                    ' If the sketch template is with offset of 200mm from F23. Template set name requested 'F23 + 200'
                    ' If the sketch template is with offset of 300mm from W23. Template set name requested 'W23 + 300'
                    strName = GetPlaneNameWithOffset(oBCPos, strAxis, oFrameSystem)
                    oIJNamedItem.Name = strName
                End If
                
                Set oTemplateReport = Nothing
                Set oBCPos = Nothing
            Else ' Even templates
                
                ' Even template need names like A, B, C, D ...
                strName = Chr(lCode)
                oIJNamedItem.Name = strName
                
                lCode = lCode + 1
                
            End If
        End If
               
        Set oIJNamedItem = Nothing
        Set oTemplate = Nothing
    Next nCount
    
CleanUp:
    Set oMfgTemplateSet = Nothing
    Set oProcessSettings = Nothing
    Set oTemplates = Nothing
    Set oSPGMiddleHelper = Nothing
    Set oMfgFrameSys = Nothing
    Set oFrameSystem = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3012, , "RULES")
    GoTo CleanUp
End Sub
 
Private Function IsSketchTemplate(ByVal oMfgTemplate As Object) As Boolean
    Const METHOD = "IsSketchTemplate"
    On Error GoTo ErrorHandler
    
    Dim oMfgGeomCol3d As IJMfgGeomCol3d
    Set oMfgGeomCol3d = oMfgTemplate

    Dim eMfgGeomType As StrMfgGeometryType
    eMfgGeomType = oMfgGeomCol3d.GetGeometryType
    
    If eMfgGeomType = STRMFG_SketchedTemplate Then
        IsSketchTemplate = True
    End If

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetPlaneNameWithOffset(oBCPos As IJDPosition, strAxis As String, oFrameSystem As IHFrameSystem) As String
Const METHOD = "GetPlaneNameWithOffset"
On Error GoTo ErrorHandler

    Dim oRefplanes  As IJDRefPlaneCollection
    Dim oBCPosVal   As Double
    Dim strName     As String
        
    If strAxis = "X" Then
        Set oRefplanes = oFrameSystem.PrincipalXAxis.RefPlaneCollection
        oBCPosVal = oBCPos.X
        strName = "F"
    ElseIf strAxis = "Y" Then
        Set oRefplanes = oFrameSystem.PrincipalYAxis.RefPlaneCollection
        oBCPosVal = oBCPos.Y
        strName = "L"
    Else
        Set oRefplanes = oFrameSystem.PrincipalZAxis.RefPlaneCollection
        oBCPosVal = oBCPos.Z
        strName = "W"
    End If
    
    ' Iterate through the planes and find the nearest plane with offset
    Dim i                       As Long
    For i = 1 To oRefplanes.Count - 1
    
        Dim oCurrGridData As ISPGGridData
        Dim oNextGridData As ISPGGridData
                
        Set oCurrGridData = oRefplanes.Item(i)
        Set oNextGridData = oRefplanes.Item(i + 1)
        
        If (Abs(oBCPosVal) >= oCurrGridData.Position) And (Abs(oBCPosVal) <= oNextGridData.Position) Then
            Dim dOffset     As Double
            dOffset = Abs(Abs(oBCPosVal) - oCurrGridData.Position)
            GetPlaneNameWithOffset = strName & Mid(oRefplanes.Item(i).Name, 2) & " + " & Round(dOffset * 1000, 0)
            Exit For
        End If
        
        Set oCurrGridData = Nothing
        Set oNextGridData = Nothing
        
    Next i
    
    Set oRefplanes = Nothing
    
Exit Function
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function

