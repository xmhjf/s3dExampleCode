VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Implements IJPMSortGroups

Dim m_objBins As StrMfgDrawingAVP.CBins
Dim m_objRects As StrMfgDrawingAVP.CRectangles

Dim m_objRects1 As RAD2D.Groups
Dim m_objRect1 As RAD2D.Group

Dim m_objApp As Object
Dim m_objRadApp As RAD2D.Application
Dim m_objDoc As RAD2D.Document
Dim m_objSheets As RAD2D.Sheets
Dim m_objOrigSheet As RAD2D.Sheet

Dim m_dBinMinX As Double, m_dBinMinY As Double, m_dBinMaxX As Double, m_dBinMaxY As Double
Dim m_dBinCenterX As Double, m_dBinCenterY As Double

Dim m_lRotations() As Long

Dim m_oProfileCollection As Collection
Dim m_oBinKeys As Collection
Dim m_oCurrProfileCollection As Collection

Const TOLERANCE As Double = 0.000001

'********************************************************************
' Routine:  Class_Initialize
' Abstract: Initialize all needed objects
'********************************************************************
Private Sub Class_Initialize()

End Sub

'********************************************************************
' Routine:  IJPMSortGroups_SortGroups
' Abstract: Decide Which SortGroups rule to trigger
'********************************************************************
Private Sub IJPMSortGroups_SortGroups(oSheet As Object)

    Dim lCount As Long
    Dim lIdx As Long
    Dim strPartType As String
    
    Set m_objOrigSheet = oSheet
    
    'Determine the part type we will be sorting
    strPartType = "SMS_PLATES"
    lCount = m_objOrigSheet.AttributeSets.Count
    For lIdx = 1 To lCount
        If m_objOrigSheet.AttributeSets.Item(lIdx).SetName = "PartType" Then
            strPartType = m_objOrigSheet.AttributeSets.Item(lIdx).Item(1).Value
        End If
    Next
    
    If strPartType = "SMS_PROFILES" Then
        SortProfiles
    Else
        SortPlates
    End If

End Sub

'********************************************************************
' Routine:  SortPlates
' Abstract: Trigger the sort groups rule for plate parts.
'********************************************************************
Private Sub SortPlates()
    
    Dim objRect As StrMfgDrawingAVP.CRectangle
    Dim lCount As Long, i As Long
    Dim oColl As New Collection
    Dim objRect1 As RAD2D.Group

    Set m_objBins = Nothing
    Set m_objRects = Nothing
    
    Set m_objBins = New StrMfgDrawingAVP.CBins
    Set m_objRects = New StrMfgDrawingAVP.CRectangles
    
    Set m_objRects1 = m_objOrigSheet.Groups
    ReDim m_lRotations(1 To 1)
   
    GetBinCoordinates
    
    m_objBins.AddByDefiningCoordinates m_dBinMinX, m_dBinMinY, m_dBinMaxX, m_dBinMaxY
    
    Set oColl = PopulateRectanglesCollection
    Set oColl = MergeSort(oColl, 1, m_objRects1.Count)
    
    For Each objRect1 In oColl
        HandleNewRectangle objRect1
    Next objRect1
    
    Set oColl = Nothing
    Set objRect1 = Nothing
    Set objRect = Nothing
    
End Sub

'********************************************************************
' Routine:  SortProfiles
' Abstract: Trigger the sort groups rule for profile parts.
'********************************************************************
Private Sub SortProfiles()
    
    Dim lCount As Long
    Dim lIdx As Long
    Dim oColl As New Collection
    Dim objRect1 As RAD2D.Group
    
    Set m_oProfileCollection = New Collection
    Set m_oCurrProfileCollection = New Collection
    Set m_oBinKeys = New Collection
    Set m_objBins = Nothing
    Set m_objRects = Nothing
    
    Set m_objBins = New StrMfgDrawingAVP.CBins
    Set m_objRects = New StrMfgDrawingAVP.CRectangles
    
    Set m_objRects1 = m_objOrigSheet.Groups
    ReDim m_lRotations(1 To 1)
    
    GetBinCoordinates
    FillProfileCollection
    InitBinsForProfiles
    
    lCount = m_oProfileCollection.Count
    For lIdx = 1 To lCount
        Set m_oCurrProfileCollection = m_oProfileCollection(lIdx)
        Set oColl = PopulateProfileRectanglesCollection()
        Set oColl = MergeSort(oColl, 1, m_oCurrProfileCollection.Count)
        For Each objRect1 In oColl
            HandleNewProfileRectangle objRect1, lIdx
            lRect = lRect + 1
        Next objRect1
        Set oColl = Nothing
        Set objRect1 = Nothing
        Set objRect = Nothing
    Next
      
End Sub

'********************************************************************
' Routine:  FillProfileCollection
' Abstract: Fills the profile collection the groups sorted by
'           section type
'********************************************************************
Private Sub FillProfileCollection()
    On Error Resume Next
    Dim lCount As Long
    Dim lIdx As Long
    Dim oGroup As RAD2D.Group
    Dim oAttrSet As RAD2D.AttributeSet
    Dim oAttr As RAD2D.Attribute
    Dim strCurrXSection As String
    Dim oGrpArray As Collection
    Dim bFound As Boolean
    
    For Each oGroup In m_objOrigSheet.Groups
        For Each oAttrSet In oGroup.AttributeSets
            If oAttrSet.SetName = "General" Then
                lCount = oAttrSet.Count
                bFound = False
                For lIdx = 1 To lCount
                    Set oAttr = oAttrSet.Item(lIdx)
                    If oAttr.Name = "SMS_PART_INFO||XSECTION_TYPE" Then
                        bFound = True
                        strCurrXSection = oAttr.Value
                        Set oGrpArray = m_oProfileCollection.Item(strCurrXSection)
                        If oGrpArray Is Nothing Then
                            Set oGrpArray = New Collection
                            oGrpArray.Add oGroup
                            m_oProfileCollection.Add oGrpArray, strCurrXSection
                            m_oBinKeys.Add strCurrXSection
                            Set oGrpArray = Nothing
                        Else
                            oGrpArray.Add oGroup
                            Set oGrpArray = Nothing
                        End If
                        Exit For
                    End If
                Next
                If Not bFound Then
                    strCurrXSection = "PLATE"
                    Set oGrpArray = m_oProfileCollection.Item(strCurrXSection)
                    If oGrpArray Is Nothing Then
                        Set oGrpArray = New Collection
                        oGrpArray.Add oGroup
                        m_oProfileCollection.Add oGrpArray, strCurrXSection
                        m_oBinKeys.Add strCurrXSection
                        Set oGrpArray = Nothing
                    Else
                        oGrpArray.Add oGroup
                        Set oGrpArray = Nothing
                    End If
                End If
                Exit For
            End If
        Next oAttrSet
    Next oGroup
    
    Set oGroup = Nothing
    Set oGrpArray = Nothing
    Set oAttrSet = Nothing
    Set oAttr = Nothing
End Sub

'********************************************************************
' Routine:  InitBinsForProfiles
' Abstract: Initialize the bin objects for the profile parts
'********************************************************************
Private Sub InitBinsForProfiles()
    
    Dim lFactor As Long
    Dim lCount As Long
    Dim lIdx As Long
    Dim oColl As Collection
    Dim oGroup As RAD2D.Group
    Dim oLayer As RAD2D.Layer
    Dim dTempX As Double
    Dim dLargeX As Double
    Dim dHeight As Double
    Dim dWidth As Double
    Dim strCurrLayer As String
    
    Dim strSectionType As String
    Dim oTxt As RAD2D.TextBox
    
    lCount = m_oProfileCollection.Count
    
    For lIdx = 1 To lCount
        
        dTempX = 0
        
        Set oColl = m_oProfileCollection.Item(lIdx)
        For Each oGroup In oColl
            GetRange oGroup, dHeight, dWidth
            If dTempX < dWidth Then
                dTempX = dTempX + dWidth
                If (dTempX + 0.4) > 15 Then
                    dTempX = dWidth
                End If
            End If
        Next oGroup
            
        m_objBins.AddByDefiningCoordinates m_dBinMinX, m_dBinMinY, m_dBinMaxX, m_dBinMaxY
        dLargeX = dLargeX + dTempX + 0.2
        Set oColl = Nothing
        m_dBinMinX = dLargeX
        m_dBinMaxX = m_dBinMinX + GLOBAL_X_WIDTH

        strCurrLayer = m_objOrigSheet.Layers.ActiveLayer
        
        For Each oLayer In m_objOrigSheet.Layers
            If Not oLayer Is Nothing Then
                If oLayer.Name = "AnnotationsLayer" Then
                    oLayer.Activate
                    Exit For
                End If
            End If
        Next oLayer
        
        Set oTxt = m_objOrigSheet.TextBoxes.Add((m_dBinMinX - (dTempX / 2) - 0.2), 0.1)
        oTxt.Units = igPaperStyleUnits
        oTxt.Text = m_oBinKeys.Item(lIdx)
        oTxt.Visible = True
        oTxt.Edit.TextSize = 0.2
        oTxt.Edit.Font = "Arial"
        oTxt.Edit.Color = "&H1"
        oTxt.VerticalJustification = igVerticalShapeCenter
        oTxt.HorizontalJustification = igHorizontalShapeCenter
        
        For Each oLayer In m_objOrigSheet.Layers
            If Not oLayer Is Nothing Then
                If oLayer.Name = strCurrLayer Then
                    oLayer.Activate
                    Exit For
                End If
            End If
        Next oLayer
        
    Next
    
    Set oGroup = Nothing
    Set oLayer = Nothing
    Set oColl = Nothing
    
End Sub

'********************************************************************
' Routine:  PopulateRectanglesCollection
' Abstract: Populate the rectangle collection containing bin
'           information for each part
'********************************************************************
Private Function PopulateRectanglesCollection() As Collection
    On Error Resume Next
    Dim objRADRects As RAD2D.Rectangles2d
    Dim objRADRect As RAD2D.Rectangle2d
    Dim objAVPRect As StrMfgDrawingAVP.CRectangle
    Dim objAttSets As RAD2D.AttributeSets
    Dim objAttSet As RAD2D.AttributeSet
    Dim i As Long, lCount As Long
    Dim objRADRects1 As RAD2D.Groups
    Dim objRADRect1 As RAD2D.Group
    Dim objAVPRect1 As RAD2D.Group
    Dim objAVPRects1 As RAD2D.Groups
    Dim oGrpCol As New Collection
    
    Randomize
    
    Set objRADRects1 = m_objOrigSheet.Groups
    
    If Check2.Value = vbUnchecked Then
        For Each objRADRect1 In objRADRects1
            With objRADRect1
                Set objAttSets = .AttributeSets
                Set objAttSet = objAttSets.Item("Bin")
                If objAttSet Is Nothing Then
                    oGrpCol.Add objRADRect1
                Else
                    Set objAttSet = Nothing
                End If
            End With
        Next objRADRect1
    Else
        lCount = CLng(Text1.Text)
        For i = 1 To lCount
            Set objAVPRect = m_objRects.AddByDimensions(Rnd * 0.08795 + 0.02, Rnd * 0.1197 + 0.02)
        Next i
    End If
    
    Set PopulateRectanglesCollection = oGrpCol
    
End Function

'********************************************************************
' Routine:  PopulateRectanglesCollection
' Abstract: Populate the rectangle collection containing bin
'           information for each part
'********************************************************************
Private Function PopulateProfileRectanglesCollection() As Collection
    On Error Resume Next
    Dim objRADRects As RAD2D.Rectangles2d
    Dim objRADRect As RAD2D.Rectangle2d
    Dim objAVPRect As StrMfgDrawingAVP.CRectangle
    Dim objAttSets As RAD2D.AttributeSets
    Dim objAttSet As RAD2D.AttributeSet
    Dim i As Long, lCount As Long
    Dim objRADRect1 As RAD2D.Group
    Dim objAVPRect1 As RAD2D.Group
    Dim objAVPRects1 As RAD2D.Groups
    Dim oGrpCol As New Collection
    Dim oObj As Object
    
    Randomize
  
    If Check2.Value = vbUnchecked Then
        For Each objRADRect1 In m_oCurrProfileCollection
            With objRADRect1
                Set objAttSets = .AttributeSets
                Set objAttSet = objAttSets.Item("Bin")
                If objAttSet Is Nothing Then
                    oGrpCol.Add objRADRect1
                Else
                    Set objAttSet = Nothing
                End If
            End With
        Next objRADRect1
    Else
        lCount = CLng(Text1.Text)
        For i = 1 To lCount
            Set objAVPRect = m_objRects.AddByDimensions(Rnd * 0.08795 + 0.02, Rnd * 0.1197 + 0.02)
        Next i
    End If
    
    Set PopulateProfileRectanglesCollection = oGrpCol
    
End Function

'********************************************************************
' Routine:  GetBinCoordinates
' Abstract: Set the cooddinates for the bin object.
'********************************************************************
Private Sub GetBinCoordinates()
    On Error Resume Next
       
    m_dBinMinX = X_ORIGIN
    m_dBinMinY = Y_ORIGIN
    m_dBinMaxX = GLOBAL_X_WIDTH
    m_dBinMaxY = GLOBAL_Y_WIDTH
    
    m_dBinCenterX = m_dBinMinX
    m_dBinCenterY = m_dBinMaxY / 2
        
End Sub

'********************************************************************
' Routine:  HandleNewRectangle
' Abstract: Try to find a greenspace into which the given object
'           will fit. If a greenspace cannot be found create a new
'           greenspace for the object.
'********************************************************************
Private Sub HandleNewRectangle(ByRef objRect As RAD2D.Group)
    
    Dim dRectWidth As Double, dRectHeight As Double
    Dim bFoundFit As Boolean
    Dim objBin As StrMfgDrawingAVP.CBin
    Dim objGSS As StrMfgDrawingAVP.CGreenSpaces
    Dim objGS As StrMfgDrawingAVP.CGreenSpace
    Dim lBin As Long
    Dim dHeight As Double, dWidth As Double
    
    GetRange objRect, dRectHeight, dRectWidth
    lBin = lIdx
    ReDim Preserve m_lRotations(1 To m_objBins.Count)
    
    For Each objBin In m_objBins
        
        lBin = lBin + 1
        Set objGSS = objBin.GreenSpaces
        For Each objGS In objGSS
            bFoundFit = True
            With objGS
                Select Case True
                    Case .Width < dRectWidth, .Height < dRectHeight
                        bFoundFit = False
                End Select
            End With
            If bFoundFit Then
                AddRectangleToGreenSpace objBin, objGS, objRect, dRectWidth, dRectHeight
                GoTo DONE
            End If
        Next objGS
        
        If Not bFoundFit Then
            For Each objGS In objGSS
                bFoundFit = True
                With objGS
                    Select Case True
                        Case .Height < dRectWidth, .Width < dRectHeight
                            bFoundFit = False
                    End Select
                End With
                If bFoundFit Then
                    m_lRotations(lBin) = m_lRotations(lBin) + 1
                    AddRectangleToGreenSpace objBin, objGS, objRect, dRectHeight, dRectWidth
                    GoTo DONE
                End If
            Next objGS
        End If
        
    Next objBin
    
DONE:
    
    If Not bFoundFit Then
        Set objBin = m_objBins.AddByDefiningCoordinates(m_dBinMinX, m_dBinMinY, m_dBinMaxX, m_dBinMaxY)
        Set objGSS = objBin.GreenSpaces
        Set objGS = objGSS.Item(1)
        AddRectangleToGreenSpace objBin, objGS, objRect, dRectWidth, dRectHeight
    End If
    
    ShrinkGreenSpaces objRect, objGSS
    
    objGSS.Sort

End Sub

'********************************************************************
' Routine:  HandleNewProfileRectangle
' Abstract: Try to find a greenspace into which the given object
'           will fit. If a greenspace cannot be found create a new
'           greenspace for the object with respect to section type
'           columns
'********************************************************************
Private Sub HandleNewProfileRectangle(ByRef objRect As RAD2D.Group, ByVal lCount As Long)
    
    Dim dRectWidth As Double, dRectHeight As Double
    Dim bFoundFit As Boolean
    Dim objBin As StrMfgDrawingAVP.CBin
    Dim objGSS As StrMfgDrawingAVP.CGreenSpaces
    Dim objGS As StrMfgDrawingAVP.CGreenSpace
    Dim lBin As Long
    Dim dHeight As Double, dWidth As Double
    Dim dMinX As Double
    Dim dMinY As Double
    Dim dMaxX As Double
    Dim dMaxY As Double
    
    lBin = 0
    
    For Each objBin In m_objBins
        
        lBin = lBin + 1
        If lBin < lCount Then GoTo NextBin
        
        Set objGSS = objBin.GreenSpaces
        For Each objGS In objGSS
            bFoundFit = True
            With objGS
                Select Case True
                    Case .Width < dRectWidth, .Height < dRectHeight
                        bFoundFit = False
                End Select
            End With
            If bFoundFit Then
                AddRectangleToGreenSpace objBin, objGS, objRect, dRectWidth, dRectHeight
                GoTo DONE
            End If
        Next objGS
        
        If Not bFoundFit Then
            For Each objGS In objGSS
                bFoundFit = True
                With objGS
                    Select Case True
                        Case .Height < dRectWidth, .Width < dRectHeight
                            bFoundFit = False
                    End Select
                End With
                If bFoundFit Then
                    m_lRotations(lBin) = m_lRotations(lBin) + 1
                    AddRectangleToGreenSpace objBin, objGS, objRect, dRectHeight, dRectWidth
                    GoTo DONE
                End If
            Next objGS
        End If
NextBin:
    Next objBin
    
DONE:
    
    If Not bFoundFit Then
        Set objBin = m_objBins.Item(lCount)
        objBin.GetDefiningCoordinates dMinX, dMinY, dMaxX, dMaxY
        Set objBin = Nothing
        Set objBin = m_objBins.AddByDefiningCoordinates(dMinX, dMinY, dMaxX, dMaxY)
        Set objGSS = objBin.GreenSpaces
        Set objGS = objGSS.Item(1)
        AddRectangleToGreenSpace objBin, objGS, objRect, dRectWidth, dRectHeight
    End If
    
    ShrinkGreenSpaces objRect, objGSS
    
    objGSS.Sort

End Sub

'********************************************************************
' Routine:  FindLargeX
' Abstract: Find the x in the collection of objects and set the new
'           bin to have that as its minimum x value
'********************************************************************
Sub FindLargeX()
    
    Dim oProfileCollection As Collection
    Dim oGroup As RAD2D.Group
    Dim lCount As Long
    Dim lIdx As Long
    Dim maxX As Double
    Dim minX As Double
    Dim largeX As Double
    Dim TempX As Double
    
    largeX = 0
    TempX = 0
    
    lCount = m_objBins.Count
    If lCount > 0 Then
        For lIdx = 1 To lCount
            Set oProfileCollection = m_oProfileCollection(lIdx)
            TempX = 0
            For Each oGroup In oProfileCollection
                GetRange oGroup, minX, maxX
                If TempX < maxX Then
                    TempX = TempX + maxX
                End If
            Next oGroup
            largeX = largeX + TempX
        Next
        If largeX > 0 Then
            m_dBinMinX = largeX
            m_dBinMaxX = m_dBinMinX + m_dBinMaxX + 0.05
        End If
    End If
    
    Set oGroup = Nothing
    Set oProfileCollection = Nothing
    
End Sub

'********************************************************************
' Routine:  AddRectangleToGreenSpace
' Abstract: Add the object to greenspace using bottom left location
'           preferences.
' ToDo:     Check object for location preferences
'********************************************************************
Private Sub AddRectangleToGreenSpace(ByVal objBin As StrMfgDrawingAVP.CBin, ByVal objGS As StrMfgDrawingAVP.CGreenSpace, ByRef objRect As RAD2D.Group, ByVal dRectWidth As Double, dRectHeight As Double)
          
    Dim dRectMinX As Double, dRectMinY As Double, dRectMaxX As Double, dRectMaxY As Double
    Dim dGSMinX As Double, dGSMinY As Double, dGSMaxX As Double, dGSMaxY As Double
    Dim dGSCenterX As Double, dGSCenterY As Double
    Dim dHalfRectWidth As Double, dHalfRectHeight As Double
    
    objGS.GetDefiningCoordinates dGSMinX, dGSMinY, dGSMaxX, dGSMaxY
    dRectMinX = m_dBinCenterX - dRectWidth / 2#
    dRectMinY = m_dBinCenterY - dRectHeight / 2#
    dRectMaxX = m_dBinCenterX + dRectWidth / 2#
    dRectMaxY = m_dBinCenterY + dRectHeight / 2#
    
    'Bottom side of greenspace is closer to center of bin than top side
    If FPGreaterThan(Abs(dGSMaxY - m_dBinCenterY), Abs(dGSMinY - m_dBinCenterY)) Then
        
        'Right side of greenspace is closer to center of bin than left side
        If FPGreaterThan(Abs(dGSMinX - m_dBinCenterX), Abs(dGSMaxX - m_dBinCenterX)) Then
            'Place rectangle at bottom right of greenspace.
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMaxX - dRectWidth, dGSMinY, dGSMaxX, dGSMinY + dRectHeight, objRect)
        
        'Left side of greenspace is closer to center of bin than right side
        ElseIf FPGreaterThan(Abs(dGSMaxX - m_dBinCenterX), Abs(dGSMinX - m_dBinCenterX)) Then
            'Place rectangle at bottom left of greenspace.
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMinX, dGSMinY, dGSMinX + dRectWidth, dGSMinY + dRectHeight, objRect)
         
        'Left and right sides of greenspace are equidistant from center
        Else
            'Place rectangle at bottom center of greenspace.
            dGSCenterX = (dGSMinX + dGSMaxX) / 2#
            dHalfRectWidth = dRectWidth / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSCenterX - dHalfRectWidth, dGSMinY, dGSCenterX + dHalfRectWidth, dGSMinY + dRectHeight, objRect)
        
        End If
        
    'Top side of greenspace is closer to center of bin than bottom side
    ElseIf FPGreaterThan(Abs(dGSMinY - m_dBinCenterY), Abs(dGSMaxY - m_dBinCenterY)) Then
      
        'Right side of greenspace is closer to center of bin than left side
        If FPGreaterThan(Abs(dGSMinX - m_dBinCenterX), Abs(dGSMaxX - m_dBinCenterX)) Then
            'Place rectangle at top right of greenspace.
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMaxX - dRectWidth, dGSMaxY - dRectHeight, dGSMaxX, dGSMaxY, objRect)
            
        'Left side of greenspace is closer to center of bin than right side
        ElseIf FPGreaterThan(Abs(dGSMaxX - m_dBinCenterX), Abs(dGSMinX - m_dBinCenterX)) Then
            'Place rectangle at top left of greenspace.
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMinX, dGSMaxY - dRectHeight, dGSMinX + dRectWidth, dGSMaxY, objRect)
            
        'Left and right sides of greenspace are equidistant from center
        Else
            'Place rectangle at top center of greenspace.
            dGSCenterX = (dGSMinX + dGSMaxX) / 2#
            dHalfRectWidth = dRectWidth / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSCenterX - dHalfRectWidth, dGSMaxY - dRectHeight, dGSCenterX + dHalfRectWidth, dGSMaxY, objRect)
        
        End If
    
    'Top and bottom sides of greenspace are equidistant from center
    Else
         
        'Right side of greenspace is closer to center of bin than left side
        If FPGreaterThan(Abs(dGSMinX - m_dBinCenterX), Abs(dGSMaxX - m_dBinCenterX)) Then
            'Place rectangle at center right of greenspace.
            dGSCenterY = (dGSMinY + dGSMaxY) / 2#
            dHalfRectHeight = dRectHeight / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMaxX - dRectWidth, dGSCenterY - dHalfRectHeight, dGSMaxX, dGSCenterY + dHalfRectHeight, objRect)
        
        'Left side of greenspace is closer to center of bin than right side
        ElseIf FPGreaterThan(Abs(dGSMaxX - m_dBinCenterX), Abs(dGSMinX - m_dBinCenterX)) Then
            'Place rectangle at center left of greenspace.
            dGSCenterY = (dGSMinY + dGSMaxY) / 2#
            dHalfRectHeight = dRectHeight / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMinX, dGSCenterY - dHalfRectHeight, dGSMinX + dRectWidth, dGSCenterY + dHalfRectHeight, objRect)
            
        'Left and right sides of greenspace are equidistant from center of bin.
        Else
            'Place rectangle at center center of greenspace.
            dGSCenterX = (dGSMinX + dGSMaxX) / 2#
            dGSCenterY = (dGSMinY + dGSMaxY) / 2#
            dHalfRectWidth = dRectWidth / 2#
            dHalfRectHeight = dRectHeight / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSCenterX - dHalfRectWidth, dGSCenterY - dHalfRectHeight, dGSCenterX + dHalfRectWidth, dGSCenterY + dHalfRectHeight, objRect)
        
        End If
    
    End If

End Sub

'********************************************************************
' Routine:  ShrinkGreenSpaces
' Abstract: Shrink all greenspaces the rectangle belongs to.
'********************************************************************
Private Sub ShrinkGreenSpaces(ByVal objRect As RAD2D.Group, ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces)

    Dim i As Long, lGSCount As Long, lContainCount As Long, lSideMask As Long
    Dim lGSArray() As Long

    lContainCount = 0
    
    lGSCount = objGSS.Count
    For i = 1 To lGSCount
        If RectangleInGreenSpace(objRect, objGSS.Item(i)) Then
            lContainCount = lContainCount + 1
            ReDim Preserve lGSArray(1 To lContainCount)
            lGSArray(lContainCount) = i
        End If
    Next i
    
    lSideMask = FindUnboundedRectangleSides(objRect)
    DoShrink objGSS, lGSArray, lContainCount, objRect, lSideMask

End Sub

'********************************************************************
' Routine:  RectangleInGreenSpace
' Abstract: Determine whether a rectangle is located in the given
'           greenspace.
'********************************************************************
Private Function RectangleInGreenSpace(ByVal objRect As RAD2D.Group, ByVal objGS As StrMfgDrawingAVP.CGreenSpace) As Boolean

    Dim dRectMinX As Double, dRectMinY As Double, dRectMaxX As Double, dRectMaxY As Double
    Dim dGSMinX As Double, dGSMinY As Double, dGSMaxX As Double, dGSMaxY As Double
    
    GetRangeWithMargin objRect, dRectMinX, dRectMinY, dRectMaxX, dRectMaxY
    
    objGS.GetDefiningCoordinates dGSMinX, dGSMinY, dGSMaxX, dGSMaxY

    RectangleInGreenSpace = BoxAIntersectsBoxB(dRectMinX, dRectMinY, dRectMaxX, dRectMaxY, dGSMinX, dGSMinY, dGSMaxX, dGSMaxY)

End Function

'********************************************************************
' Routine:  BoxAIntersectsBoxB
' Abstract: Determine if two boxes intersect
'********************************************************************
Private Function BoxAIntersectsBoxB(ByVal dAMinX As Double, ByVal dAMinY As Double, ByVal dAMaxX As Double, ByVal dAMaxY As Double, ByVal dBMinX As Double, ByVal dBMinY As Double, ByVal dBMaxX As Double, ByVal dBMaxY As Double) As Boolean
    
    BoxAIntersectsBoxB = False
    Select Case True
        Case PointInBox(dAMinX, dAMinY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             PointInBox(dAMinX, dAMaxY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             PointInBox(dAMaxX, dAMinY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             PointInBox(dAMaxX, dAMaxY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             PointInBox(dBMinX, dBMinY, dAMinX, dAMinY, dAMaxX, dAMaxY), _
             PointInBox(dBMinX, dBMaxY, dAMinX, dAMinY, dAMaxX, dAMaxY), _
             PointInBox(dBMaxX, dBMinY, dAMinX, dAMinY, dAMaxX, dAMaxY), _
             PointInBox(dBMaxX, dBMaxY, dAMinX, dAMinY, dAMaxX, dAMaxY), _
             BoxAOverlapsBoxB(dAMinX, dAMinY, dAMaxX, dAMaxY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             BoxAOverlapsBoxB(dBMinX, dBMinY, dBMaxX, dBMaxY, dAMinX, dAMinY, dAMaxX, dAMaxY)
             
             BoxAIntersectsBoxB = True
    End Select
    
End Function

'********************************************************************
' Routine:  PointInBox
' Abstract: Determine if a point lies within a box
'********************************************************************
Private Function PointInBox(ByVal dPointX As Double, ByVal dPointY As Double, ByVal dBoxMinX As Double, ByVal dBoxMinY As Double, ByVal dBoxMaxX As Double, ByVal dBoxMaxY As Double) As Boolean
    
    PointInBox = False
    If FPGreaterThanOrEqualTo(dPointX, dBoxMinX) Then
        If FPLessThanOrEqualTo(dPointX, dBoxMaxX) Then
            If FPGreaterThanOrEqualTo(dPointY, dBoxMinY) Then
                If FPLessThanOrEqualTo(dPointY, dBoxMaxY) Then
                    PointInBox = True
                End If
            End If
        End If
    End If

End Function

'********************************************************************
' Routine:  BoxAOverlapsBoxB
' Abstract: Determine if two boxes overlap
'********************************************************************
Private Function BoxAOverlapsBoxB(ByVal dAMinX As Double, ByVal dAMinY As Double, ByVal dAMaxX As Double, ByVal dAMaxY As Double, ByVal dBMinX As Double, ByVal dBMinY As Double, ByVal dBMaxX As Double, ByVal dBMaxY As Double) As Boolean

    BoxAOverlapsBoxB = False
    If FPGreaterThan(dBMinX, dAMinX) Then
        If FPGreaterThan(dAMaxX, dBMaxX) Then
            If FPGreaterThan(dAMinY, dBMinY) Then
                If FPGreaterThan(dBMaxY, dAMaxY) Then
                    BoxAOverlapsBoxB = True
                End If
            End If
        End If
    End If

End Function

'********************************************************************
' Routine:  FPGreaterThanOrEqualTo
' Abstract: Determine if the first value is greater than or equal to
'           the second
'********************************************************************
Private Function FPGreaterThanOrEqualTo(ByVal dA As Double, ByVal dB As Double) As Boolean

    If dA >= dB - TOLERANCE Then
        FPGreaterThanOrEqualTo = True
    Else
        FPGreaterThanOrEqualTo = False
    End If

End Function

'********************************************************************
' Routine:  FPLessThanOrEqualTo
' Abstract: Determine if the first object is less than or equal to
'           the second
'********************************************************************
Private Function FPLessThanOrEqualTo(ByVal dA As Double, ByVal dB As Double) As Boolean

    If dA <= dB + TOLERANCE Then
        FPLessThanOrEqualTo = True
    Else
        FPLessThanOrEqualTo = False
    End If

End Function

'********************************************************************
' Routine:  FPEqualTo
' Abstract: Determine if two values are equal
'********************************************************************
Private Function FPEqualTo(ByVal dA As Double, ByVal dB As Double) As Boolean

    If Abs(dA - dB) <= TOLERANCE Then
        FPEqualTo = True
    Else
        FPEqualTo = False
    End If

End Function

'********************************************************************
' Routine:  FPBetween
' Abstract: Determine if the value lies between the other two
'********************************************************************
Private Function FPBetween(ByVal dA As Double, dB1 As Double, dB2 As Double) As Boolean

    FPBetween = False
    If FPGreaterThanOrEqualTo(dA, dB1) Then
        If FPLessThanOrEqualTo(dA, dB2) Then
            FPBetween = True
        End If
    End If
        
End Function

'********************************************************************
' Routine:  FPGreaterThan
' Abstract: Determine if the first value is greater thean the second
'********************************************************************
Private Function FPGreaterThan(ByVal dA As Double, ByVal dB As Double) As Boolean
    
    If dA > dB + TOLERANCE Then
        FPGreaterThan = True
    Else
        FPGreaterThan = False
    End If

End Function

'********************************************************************
' Routine:  FindUnboundedRectangleSides
' Abstract: Find sides of the rectangle that are unbounded
'********************************************************************
Private Function FindUnboundedRectangleSides(ByVal objRect As RAD2D.Group) As Long

    Dim dTopClearDist As Double, dBottomClearDist As Double
    Dim dLeftClearDist As Double, dRightClearDist As Double
    Dim dRectWidth As Double, dRectHeight As Double
         
    GetRange objRect, dRectHeight, dRectWidth
    dTopClearDist = dRectWidth
    dBottomClearDist = dRectWidth
    dLeftClearDist = dRectHeight
    dRightClearDist = dRectHeight
    
    FindRectangleBinAdjacency objRect, dTopClearDist, dBottomClearDist, dLeftClearDist, dRightClearDist
    FindRectangleRectangleAdjacency objRect, dTopClearDist, dBottomClearDist, dLeftClearDist, dRightClearDist
    
    FindUnboundedRectangleSides = 0
    
    If FPGreaterThan(dTopClearDist, 0#) Then
        FindUnboundedRectangleSides = FindUnboundedRectangleSides Or 1
    End If
    
    If FPGreaterThan(dBottomClearDist, 0#) Then
        FindUnboundedRectangleSides = FindUnboundedRectangleSides Or 2
    End If
    
    If FPGreaterThan(dLeftClearDist, 0#) Then
        FindUnboundedRectangleSides = FindUnboundedRectangleSides Or 4
    End If
    
    If FPGreaterThan(dRightClearDist, 0#) Then
        FindUnboundedRectangleSides = FindUnboundedRectangleSides Or 8
    End If

End Function

'********************************************************************
' Routine:  FindRectangleBinAdjacency
' Abstract: Find which side of the rectangle borders the bin if any
'********************************************************************
Private Sub FindRectangleBinAdjacency(ByVal objRect As RAD2D.Group, ByRef dTopClearDist As Double, ByRef dBottomClearDist As Double, ByRef dLeftClearDist As Double, ByRef dRightClearDist As Double)

    Dim lWhichSides As Long

    If RectangleBordersBin(objRect, lWhichSides) Then
        If (lWhichSides And 1) = 1 Then
            dTopClearDist = 0#
        End If
        
        If (lWhichSides And 2) = 2 Then
            dBottomClearDist = 0#
        End If
        
        If (lWhichSides And 4) = 4 Then
            dLeftClearDist = 0#
        End If
        
        If (lWhichSides And 8) = 8 Then
            dRightClearDist = 0#
        End If
    End If

End Sub

'********************************************************************
' Routine:  RectangleBordersBin
' Abstract: Determine whether the rectangle borders the bin
'********************************************************************
Private Function RectangleBordersBin(ByVal objRect As RAD2D.Group, ByRef lWhichSides As Long) As Boolean

    Dim dRectMinX As Double, dRectMinY As Double, dRectMaxX As Double, dRectMaxY As Double
      
    GetRangeWithMargin objRect, dRectMinX, dRectMinY, dRectMaxX, dRectMaxY
    
    lWhichSides = BoxABordersWhichInternalSidesOfBoxB(dRectMinX, dRectMinY, dRectMaxX, dRectMaxY, m_dBinMinX, m_dBinMinY, m_dBinMaxX, m_dBinMaxY)
    
    If lWhichSides > 0 Then
        RectangleBordersBin = True
    Else
       RectangleBordersBin = False
    End If
    
End Function

'********************************************************************
' Routine:  BoxABordersWhichInternalSidesOfBoxB
' Abstract: Determine which sides of Box A border BoxB
'********************************************************************
Private Function BoxABordersWhichInternalSidesOfBoxB(ByVal dAMinX As Double, ByVal dAMinY As Double, ByVal dAMaxX As Double, ByVal dAMaxY As Double, ByVal dBMinX As Double, ByVal dBMinY As Double, ByVal dBMaxX As Double, ByVal dBMaxY As Double) As Long

    BoxABordersWhichInternalSidesOfBoxB = 0
    
    'Check to see if the top side of Box A borders the top side of Box B.
    If FPEqualTo(dAMaxY, dBMaxY) Then
        BoxABordersWhichInternalSidesOfBoxB = BoxABordersWhichInternalSidesOfBoxB Or 1
    End If
        
    'Check to see if the bottom side of Box A borders the bottom side of Box B.
    If FPEqualTo(dAMinY, dBMinY) Then
        BoxABordersWhichInternalSidesOfBoxB = BoxABordersWhichInternalSidesOfBoxB Or 2
    End If
    
    'Check to see if the left side of Box A borders the left side of Box B.
    If FPEqualTo(dAMinX, dBMinX) Then
        BoxABordersWhichInternalSidesOfBoxB = BoxABordersWhichInternalSidesOfBoxB Or 4
    End If
    
    'Check to see if the right side of Box A borders the right side of Box B.
    If FPEqualTo(dAMaxX, dBMaxX) Then
        BoxABordersWhichInternalSidesOfBoxB = BoxABordersWhichInternalSidesOfBoxB Or 8
    End If

End Function

'********************************************************************
' Routine:  FindRectangleRectangleAdjacency
' Abstract: Find which sides of the rectangle border other rectangles
'********************************************************************
Private Sub FindRectangleRectangleAdjacency(ByVal objRect As RAD2D.Group, ByRef dTopClearDist As Double, ByRef dBottomClearDist As Double, ByRef dLeftClearDist As Double, ByRef dRightClearDist As Double)
    
    Dim objRects As RAD2D.Groups, objCompareRect As RAD2D.Group
    Dim lWhichSide As Long
    Dim dHorizDistInCommon As Double, dVertDistInCommon As Double
    Dim objRects1 As Object
    
    Set objRects1 = objRect.Parent
       
    For Each objCompareRect In m_objRects1
        If Not objRect Is objCompareRect Then
            If RectangleABordersRectangleB(objCompareRect, objRect, lWhichSide, dHorizDistInCommon, dVertDistInCommon) Then
                'If lWhichSide = 0, then no side of Box A borders any side of Box B.
                'If lWhichSide = 1, then bottom side of Box A borders top side of Box B.
                'If lWhichSide = 2, then top side of Box A borders bottom side of Box B.
                'If lWhichSide = 3, then right side of Box A borders left side of Box B.
                'If lWhichSide = 4, then left side of Box A borders right side of Box B.
                Select Case lWhichSide
                
                    Case 1
                        dTopClearDist = dTopClearDist - dHorizDistInCommon
                    Case 2
                        dBottomClearDist = dBottomClearDist - dHorizDistInCommon
                    Case 3
                        dLeftClearDist = dLeftClearDist - dVertDistInCommon
                    Case 4
                        dRightClearDist = dRightClearDist - dVertDistInCommon
                    
                End Select
            End If
        End If
    Next objCompareRect
    
End Sub

'********************************************************************
' Routine:  PopulateRectanglesCollection
' Abstract: Determine if rectangle A borders rectangle B
'********************************************************************
Private Function RectangleABordersRectangleB(ByVal objRectA As RAD2D.Group, ByVal objRectB As RAD2D.Group, ByRef lWhichSide As Long, ByRef dHorizDistInCommon As Double, ByRef dVertDistInCommon As Double) As Boolean

    Dim dAMinX As Double, dAMinY As Double, dAMaxX As Double, dAMaxY As Double
    Dim dBMinX As Double, dBMinY As Double, dBMaxX As Double, dBMaxY As Double
        
    GetRangeWithMargin objRectA, dAMinX, dAMinY, dAMaxX, dAMaxY
    GetRangeWithMargin objRectB, dBMinX, dBMinY, dBMaxX, dBMaxY
    
    'If BoxABordersWhichSideOfBoxB = 0, then no side of Box A borders any side of Box B.
    'If BoxABordersWhichSideOfBoxB = 1, then bottom side of Box A borders top side of Box B.
    'If BoxABordersWhichSideOfBoxB = 2, then top side of Box A borders bottom side of Box B.
    'If BoxABordersWhichSideOfBoxB = 3, then right side of Box A borders left side of Box B.
    'If BoxABordersWhichSideOfBoxB = 4, then left side of Box A borders right side of Box B.
    lWhichSide = BoxABordersWhichSideOfBoxB(dAMinX, dAMinY, dAMaxX, dAMaxY, dBMinX, dBMinY, dBMaxX, dBMaxY, dHorizDistInCommon, dVertDistInCommon)
    
    If lWhichSide = 0 Then
        RectangleABordersRectangleB = False
    Else
        RectangleABordersRectangleB = True
    End If
    
End Function

'********************************************************************
' Routine:  BoxABordersWhichSideOfBoxB
' Abstract: Determine which sides of Box A boder Box b if any
'********************************************************************
Private Function BoxABordersWhichSideOfBoxB(ByVal dAMinX As Double, ByVal dAMinY As Double, ByVal dAMaxX As Double, ByVal dAMaxY As Double, ByVal dBMinX As Double, ByVal dBMinY As Double, ByVal dBMaxX As Double, ByVal dBMaxY As Double, ByRef dHorizDistInCommon As Double, ByRef dVertDistInCommon As Double) As Long

    BoxABordersWhichSideOfBoxB = 0
    
    If BoxASideBordersBoxBSide(dAMinX, dAMaxX, dBMinX, dBMaxX, dHorizDistInCommon) Then
    
        'Check to see if the bottom side of Box A borders the top side of Box B.
        If FPEqualTo(dAMinY, dBMaxY) Then
            BoxABordersWhichSideOfBoxB = 1
            
        'Check to see if the top side of Box A borders the bottom side of Box B.
        ElseIf FPEqualTo(dAMaxY, dBMinY) Then
            BoxABordersWhichSideOfBoxB = 2
            
        End If
        
    End If
        
    'Check to see if Box A's height range coordinates are within Box B's height range coordinates.
    If BoxASideBordersBoxBSide(dAMinY, dAMaxY, dBMinY, dBMaxY, dVertDistInCommon) Then
    
        'Check to see if the right side of Box A borders the left side of Box B.
        If FPEqualTo(dAMaxX, dBMinX) Then
            BoxABordersWhichSideOfBoxB = 3
        
        'Check to see if the left side of Box A borders the right side of Box B.
        ElseIf FPEqualTo(dAMinX, dBMaxX) Then
            BoxABordersWhichSideOfBoxB = 4
            
        End If
        
    End If

End Function

'********************************************************************
' Routine:  BoxASideBordersBoxBSide
' Abstract: Determine if two sides border each other
'********************************************************************
Private Function BoxASideBordersBoxBSide(ByVal dAMin As Double, ByVal dAMax As Double, ByVal dBMin As Double, ByVal dBMax As Double, ByRef dDistInCommon As Double) As Boolean

    dDistInCommon = 0#
    BoxASideBordersBoxBSide = True

    If FPBetween(dAMin, dBMin, dBMax) Then
        dDistInCommon = DistanceInCommon(dAMin, dAMax, dBMin, dBMax, True)
    ElseIf FPBetween(dAMax, dBMin, dBMax) Then
        dDistInCommon = DistanceInCommon(dAMin, dAMax, dBMin, dBMax, False)
    ElseIf FPBetween(dBMin, dAMin, dAMax) Then
        dDistInCommon = dBMax - dBMin
    Else
        BoxASideBordersBoxBSide = False
    End If
            
End Function

'********************************************************************
' Routine:  DistanceInCommon
' Abstract: Determine the common distance between two sides
'********************************************************************
Private Function DistanceInCommon(ByVal dAMin As Double, ByVal dAMax As Double, ByVal dBMin As Double, ByVal dBMax As Double, ByVal bAIsMin As Boolean) As Double
    
    If bAIsMin Then
        If FPLessThanOrEqualTo(dBMax, dAMax) Then
            DistanceInCommon = dBMax - dAMin
        Else
            DistanceInCommon = dAMax - dAMin
        End If
    Else
       DistanceInCommon = dAMax - dBMin
    End If
    
End Function

'********************************************************************
' Routine:  DoShrink
' Abstract: Shrink the greenspace
'********************************************************************
Private Sub DoShrink(ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces, ByRef lGSArray() As Long, ByVal lUBound As Long, ByVal objRect As RAD2D.Group, ByVal lSideMask As Long)

    Dim i As Long
    Dim dGSMinXArray() As Double, dGSMinYArray() As Double
    Dim dGSMaxXArray() As Double, dGSMaxYArray() As Double
    Dim dRectMinX As Double, dRectMinY As Double, dRectMaxX As Double, dRectMaxY As Double
    
    ReDim dGSMinXArray(1 To lUBound)
    ReDim dGSMinYArray(1 To lUBound)
    ReDim dGSMaxXArray(1 To lUBound)
    ReDim dGSMaxYArray(1 To lUBound)

    If lSideMask = 0 Then
        'The rectangle fits perfectly into one greenspace.  We just need to delete that greenspace.
        objGSS.Item(1).Delete
        Exit Sub
    End If
        
    'Populate the intersected greenspace coordinate arrays.
    For i = 1 To lUBound
        With objGSS.Item(lGSArray(i))
            .GetDefiningCoordinates dGSMinXArray(i), dGSMinYArray(i), dGSMaxXArray(i), dGSMaxYArray(i)
        End With
    Next i
    
    GetRangeWithMargin objRect, dRectMinX, dRectMinY, dRectMaxX, dRectMaxY
    
    If (lSideMask And 1) = 1 Then
        'Top side of objRect is at least partially clear.  Shrink all intersected
        'greenspaces toward the top of objRect and determine which has the largest area.
        AddNewGreenSpace objGSS, dGSMinXArray, dGSMinYArray, dGSMaxXArray, dGSMaxYArray, lUBound, dRectMaxY, 1
    End If
    
    If (lSideMask And 2) = 2 Then
        'Bottom side of objRect is at least partially clear.  Shrink all intersected
        'greenspaces toward the bottom of objRect and determine which has the largest area.
        AddNewGreenSpace objGSS, dGSMinXArray, dGSMinYArray, dGSMaxXArray, dGSMaxYArray, lUBound, dRectMinY, 2
    End If
    
    If (lSideMask And 4) = 4 Then
        'Left side of objRect is at least partially clear.  Shrink all intersected
        'greenspaces toward the left of objRect and determine which has the largest area.
        AddNewGreenSpace objGSS, dGSMinXArray, dGSMinYArray, dGSMaxXArray, dGSMaxYArray, lUBound, dRectMinX, 3
    End If
    
    If (lSideMask And 8) = 8 Then
        'Right side of objRect is at least partially clear.  Shrink all intersected
        'greenspaces toward the right of objRect and determine which has the largest area.
        AddNewGreenSpace objGSS, dGSMinXArray, dGSMinYArray, dGSMaxXArray, dGSMaxYArray, lUBound, dRectMaxX, 4
    End If
    
    DeleteGreenSpaces objGSS, lGSArray, lUBound
    
End Sub

'********************************************************************
' Routine:  AddNewGreenSpace
' Abstract: Add the new greenspace to the collection
'********************************************************************
Private Sub AddNewGreenSpace(ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces, ByRef dGSMinXArray() As Double, ByRef dGSMinYArray() As Double, ByRef dGSMaxXArray() As Double, ByRef dGSMaxYArray() As Double, ByVal lUBound As Long, ByVal dRectSideCoord As Double, ByVal lWhichSide As Long)

    Dim i As Long
    Dim dNewGSMinXArray() As Double, dNewGSMinYArray() As Double
    Dim dNewGSMaxXArray() As Double, dNewGSMaxYArray() As Double
    Dim lDiscardArray() As Long
    Dim bFound As Boolean
    
    ReDim dNewGSMinXArray(1 To lUBound)
    ReDim dNewGSMinYArray(1 To lUBound)
    ReDim dNewGSMaxXArray(1 To lUBound)
    ReDim dNewGSMaxYArray(1 To lUBound)
    ReDim lDiscardArray(1 To 1)
    
    lDiscardArray(1) = 0
    
    Select Case lWhichSide
        Case 1  'Shrink toward top of rectangle.
            For i = 1 To lUBound
                dNewGSMinXArray(i) = dGSMinXArray(i)
                dNewGSMinYArray(i) = dRectSideCoord
                dNewGSMaxXArray(i) = dGSMaxXArray(i)
                dNewGSMaxYArray(i) = dGSMaxYArray(i)
            Next i
        Case 2  'Shrink toward bottom of rectangle.
            For i = 1 To lUBound
                dNewGSMinXArray(i) = dGSMinXArray(i)
                dNewGSMinYArray(i) = dGSMinYArray(i)
                dNewGSMaxXArray(i) = dGSMaxXArray(i)
                dNewGSMaxYArray(i) = dRectSideCoord
            Next i
        Case 3  'Shrink toward left of rectangle.
            For i = 1 To lUBound
                dNewGSMinXArray(i) = dGSMinXArray(i)
                dNewGSMinYArray(i) = dGSMinYArray(i)
                dNewGSMaxXArray(i) = dRectSideCoord
                dNewGSMaxYArray(i) = dGSMaxYArray(i)
            Next i
        Case 4  'Shrink toward right of rectangle.
            For i = 1 To lUBound
                dNewGSMinXArray(i) = dRectSideCoord
                dNewGSMinYArray(i) = dGSMinYArray(i)
                dNewGSMaxXArray(i) = dGSMaxXArray(i)
                dNewGSMaxYArray(i) = dGSMaxYArray(i)
            Next i
    End Select
    
    GetDiscardArray lUBound, dNewGSMinXArray, dNewGSMinYArray, dNewGSMaxXArray, dNewGSMaxYArray, lDiscardArray, lWhichSide
    AddNonEnvelopedGreenSpaces objGSS, lUBound, dNewGSMinXArray, dNewGSMinYArray, dNewGSMaxXArray, dNewGSMaxYArray, lDiscardArray
    
End Sub

'********************************************************************
' Routine:  GetDiscardArray
' Abstract: Get the array of greenspaces to be discarded
'********************************************************************
Private Sub GetDiscardArray(ByVal lUBound As Long, ByRef dNewGSMinXArray() As Double, ByRef dNewGSMinYArray() As Double, ByRef dNewGSMaxXArray() As Double, ByRef dNewGSMaxYArray() As Double, ByRef lDiscardArray() As Long, ByVal lWhichSide As Long)

    Dim i As Long, j As Long

    For i = 1 To lUBound
        If ((lWhichSide = 1 Or lWhichSide = 2) And FPGreaterThan(dNewGSMaxYArray(i), dNewGSMinYArray(i))) Or _
           ((lWhichSide = 3 Or lWhichSide = 4) And FPGreaterThan(dNewGSMaxXArray(i), dNewGSMinXArray(i))) Then
            For j = 1 To lUBound
                If i <> j Then
                    If FPLessThanOrEqualTo(dNewGSMinXArray(i), dNewGSMinXArray(j)) Then
                        If FPLessThanOrEqualTo(dNewGSMinYArray(i), dNewGSMinYArray(j)) Then
                            If FPLessThanOrEqualTo(dNewGSMaxXArray(j), dNewGSMaxXArray(i)) Then
                                If FPLessThanOrEqualTo(dNewGSMaxYArray(j), dNewGSMaxYArray(i)) Then
                                    AddToDiscardArray lDiscardArray, j
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
        Else
            AddToDiscardArray lDiscardArray, i
        End If
    Next i

End Sub

'********************************************************************
' Routine:  AddToDiscardArray
' Abstract: Redim and object to Array
'********************************************************************
Private Sub AddToDiscardArray(ByRef lDiscardArray() As Long, ByVal lItemToDiscard As Long)

    Dim lDiscardArrayUBound As Long

    lDiscardArrayUBound = UBound(lDiscardArray)
    lDiscardArray(lDiscardArrayUBound) = lItemToDiscard
    lDiscardArrayUBound = lDiscardArrayUBound + 1
    ReDim Preserve lDiscardArray(1 To lDiscardArrayUBound)
    lDiscardArray(lDiscardArrayUBound) = 0

End Sub

'********************************************************************
' Routine:  AddNonEnvelopedGreenSpaces
' Abstract: ADd greenspaces that are not present in the discard array
'********************************************************************
Private Sub AddNonEnvelopedGreenSpaces(ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces, ByVal lUBound As Long, ByRef dNewGSMinXArray() As Double, ByRef dNewGSMinYArray() As Double, ByRef dNewGSMaxXArray() As Double, ByRef dNewGSMaxYArray() As Double, ByRef lDiscardArray() As Long)

    Dim i As Long, j As Long, lDiscardArrayUBound As Long
    Dim bFound As Boolean
    For i = 1 To lUBound
            
        bFound = False
        lDiscardArrayUBound = UBound(lDiscardArray)
        For j = 1 To lDiscardArrayUBound
            If lDiscardArray(j) = i Then
                bFound = True
                Exit For
            End If
        Next j
        
        If Not bFound Then
            objGSS.AddByDefiningCoordinates dNewGSMinXArray(i), dNewGSMinYArray(i), dNewGSMaxXArray(i), dNewGSMaxYArray(i)
        End If
    Next i

End Sub

'********************************************************************
' Routine:  DeleteGreenSpaces
' Abstract: Remove unnecessary green spaces
'********************************************************************
Private Sub DeleteGreenSpaces(ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces, ByRef lGSArray() As Long, ByVal lUBound As Long)

    Dim i As Long
    
    For i = lUBound To 1 Step -1
        With objGSS.Item(lGSArray(i))
            .Delete
        End With
    Next i

End Sub

'********************************************************************
' Routine:  MergeSort
' Abstract: Merge sort the rad objects based on rangebox size
'********************************************************************
Private Function MergeSort(ByRef objCollection As Collection, _
                      ByVal lStart As Long, _
                      ByVal lFinish As Long) As Collection
                      
                      
    Dim lMiddle     As Long
    Dim objStart    As Group
    Dim objFinish   As Group
    Dim dStartRange As Double
    Dim dFinishRange As Double
    Dim objTempCol As Collection
    Dim dHeight As Double, dWidth As Double
     
    If lFinish - lStart <= 1 Then
        If (lFinish - lStart) > 0 Then
            With objCollection
                Set objStart = .Item(lStart)
                Set objFinish = .Item(lFinish)
                
                GetRange objStart, dHeight, dStartRange
                GetRange objFinish, dHeight, dFinishRange
                
                If FPGreaterThanOrEqualTo(dFinishRange, dStartRange) Then
                    .Remove lFinish
                    .Add objFinish, , lStart
                    '2nd element is bigger than 1st one. Since for Check drawing, we need descending sort,
                    Set objTempCol = New Collection
                    objTempCol.Add objFinish
                    objTempCol.Add objStart
                Else
                    '1st element is bigger than 2nd one. Since for Check drawing, we need descending sort,
                    Set objTempCol = New Collection
                    objTempCol.Add objStart
                    objTempCol.Add objFinish
                End If ' modUtils.FPGreaterThanOrEqualTo(objFinish.Area, objStart.Area)
            End With ' objCollection
        Else  ' There is only one element to sort
            Set objStart = objCollection.Item(lStart)
            Set objTempCol = New Collection
            objTempCol.Add objStart
        End If ' lFinish - lStart <> 0
    Else
        lMiddle = (lStart + lFinish) \ 2
        MergeSort objCollection, lStart, lMiddle
        MergeSort objCollection, lMiddle + 1, lFinish
        Set objTempCol = Merge(objCollection, lStart, lMiddle, lFinish)
    End If

    Set MergeSort = objTempCol

End Function

'********************************************************************
' Routine:  GetRange
' Abstract: Get the range box of the object group
'********************************************************************
Private Sub GetRange(objGroup As Group, dHeight As Double, dWidth As Double)

    Dim dXMin As Double, dXMax As Double, dYMin As Double, dYMax As Double
    Dim dRange As Double, dMax As Double
    
    objGroup.Range dXMin, dYMin, dXMax, dYMax
    
    dWidth = dXMax - dXMin
    dHeight = dYMax - dYMin

End Sub

'********************************************************************
' Routine:  GetRangeWithMargin
' Abstract: Get the range box of the object group with added margin
'********************************************************************
Private Sub GetRangeWithMargin(objGroup As Group, dXMin As Double, dYMin As Double, dXMax As Double, dYMax As Double)
    
    objGroup.Range dXMin, dYMin, dXMax, dYMax
    
    dXMin = dXMin + X_MARGIN
    dYMin = dYMin + X_MARGIN
    dXMax = dXMax + Y_MARGIN
    dYMax = dYMax + Y_MARGIN
    
End Sub

'********************************************************************
' Routine:  Merge
' Abstract: Perform the merge part of the merge sort
'********************************************************************
Private Function Merge(ByRef objCollection As Collection, ByVal lStart As Long, _
                  ByVal lMiddle As Long, ByVal lFinish As Long) As Collection

    Dim i           As Long
    Dim lTemp       As Long
    Dim lStart1     As Long
    Dim lStart2     As Long
    Dim lFinish1    As Long
    Dim lFinish2    As Long
    Dim objStart1   As Group
    Dim objStart2   As Group
    Dim objTempCol  As Collection
    Dim dStartRange As Double
    Dim dFinishRange As Double
    Dim dHeight As Double, dWidth As Double
    
    Set objTempCol = New Collection
    
    lStart1 = lStart
    lFinish1 = lMiddle
    
    lStart2 = lMiddle + 1
    lFinish2 = lFinish
    
    With objCollection
        Do While lStart1 <= lFinish1 And lStart2 <= lFinish2
            Set objStart1 = .Item(lStart1)
            Set objStart2 = .Item(lStart2)
            
            GetRange objStart1, dHeight, dStartRange
            GetRange objStart2, dHeight, dFinishRange
            
            If FPGreaterThanOrEqualTo(dStartRange, dFinishRange) Then
                objTempCol.Add objStart1
                lStart1 = lStart1 + 1
            Else
                objTempCol.Add objStart2
                lStart2 = lStart2 + 1
            End If
        Loop
    
        If lStart1 <= lFinish1 Then
            For i = lStart1 To lFinish1
                objTempCol.Add .Item(i)
            Next i
        ElseIf lStart <= lFinish2 Then
            For i = lStart2 To lFinish2
                objTempCol.Add .Item(i)
            Next i
        End If
        
        For i = lStart To lFinish
            .Remove i
            If i <= .Count Then
                .Add objTempCol.Item(objTempCol.Count - (lFinish - i)), , i
            Else
                .Add objTempCol.Item(objTempCol.Count - (lFinish - i))
            End If
        Next i
    End With
      
    Set Merge = objTempCol
    
End Function





