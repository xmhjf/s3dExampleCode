VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Similar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  FILE:  Similar.cls
'
'  DESCRIPTION:
'  Rule for comparing two detailed parts for "topological" similarity.
'
'  AUTHOR:  Mike and Jim
'
'  HISTORY
'  2002    (MP) created
'    020423 (JB) additional stiffener and seam checks working. Added Beam Part.

Option Explicit
Implements IJBoardManagementRule
'Struct of Corner data
Private Type CornerData
    Angle As Double
    EdgePort1 As IJPort
    EdgePort2 As IJPort
    FacePort As IJPort
    Location As IJDPosition
    Normal As IJDVector
End Type

Private Const IID_IJHoleTraceAE As String = "{A87716F0-0198-4D5D-879E-234DE527C2A0}"

Private Const sMODULE As String = "Similar.cls: "

   
Private Function DiffPlateBoundaries(ByRef oCandidatePlate As StructDetailObjects.PlatePart, _
                                     ByRef oTargetPlate As StructDetailObjects.PlatePart) As Boolean
'
'only returns one boundary????
'wrapper method currently only supports detailed parts of plate system
'
    DiffPlateBoundaries = True

    Dim oTargetCollection As Collection
    Dim oCandidateCollection As Collection
    
    Dim cTargetBoundaryData As BoundaryData
    Dim cCandidateBoundaryData As BoundaryData
    Dim n As Long
    Dim nCount As Long
    Dim TempDiff As Boolean
    
    TempDiff = True
    
    Set oTargetCollection = oTargetPlate.PlateBoundaries
    Set oCandidateCollection = oCandidatePlate.PlateBoundaries

    Open "c:\BoardMgtLog.txt" For Append Access Write As #1
    Print #1, Now(); Tab; "Candidate Count"; Tab; oCandidateCollection.Count; _
                    Tab; "Target Count"; Tab; oTargetCollection.Count
    
    'same number of boundaries?
    If oCandidateCollection.Count <> oTargetCollection.Count Then
        TempDiff = True
'
'    'if number of boundaries are the same and not zero
'        'perform check of boundary types
    ElseIf oTargetCollection.Count > 0 Then
        
        'check first boundary
        cTargetBoundaryData = oTargetCollection.Item(1)
        cCandidateBoundaryData = oCandidateCollection.Item(1)
        
        Print #1, Now(); Tab; "candidate object type, type of"; Tab; _
                    cCandidateBoundaryData.ObjectType; Tab; cCandidateBoundaryData.TypeObject
        Print #1, Now(); Tab; "target    object type, type of"; Tab; _
                    cTargetBoundaryData.ObjectType; Tab; cTargetBoundaryData.TypeObject
        
        If cTargetBoundaryData.ObjectType <> cCandidateBoundaryData.ObjectType Then
            TempDiff = True
        ElseIf cTargetBoundaryData.TypeObject <> cCandidateBoundaryData.TypeObject Then
            TempDiff = True
        Else
            TempDiff = False
        End If
        
        'if plate has more than one boundary and first boundary was ok,
        'check mirrored boundries, check one plate from
        'top of list, and second plate from bottom of list.
        If (oTargetCollection.Count > 1 And TempDiff = False) Then
            nCount = oTargetCollection.Count
            For n = 2 To nCount
                
                cTargetBoundaryData = oTargetCollection.Item(n)
                cCandidateBoundaryData = oCandidateCollection.Item(nCount + 2 - n)
                
                Print #1, Now(); Tab; "candidate object type, type of"; Tab; _
                    cCandidateBoundaryData.ObjectType; Tab; cCandidateBoundaryData.TypeObject
                Print #1, Now(); Tab; "target    object type, type of"; Tab; _
                    cTargetBoundaryData.ObjectType; Tab; cTargetBoundaryData.TypeObject
                
                If cTargetBoundaryData.ObjectType <> cCandidateBoundaryData.ObjectType Then
                    TempDiff = True
                ElseIf cTargetBoundaryData.TypeObject <> cCandidateBoundaryData.TypeObject Then
                    TempDiff = True
                Else
                    TempDiff = False
                End If
                
                'check state after each boundary, cause exit if different boundary found
                If TempDiff = True Then n = nCount + 10
            Next
            
        End If
        
'   else both plates have no boundaries, this check is thus indeterminate,
'   plates may be similar so let them pass this test????
    Else
        TempDiff = False
    End If
    
    If TempDiff = False Then
        DiffPlateBoundaries = False
    Else
        DiffPlateBoundaries = True
    End If
    
    Close #1

'    'release collections
    Set oCandidateCollection = Nothing
    Set oTargetCollection = Nothing

End Function
Private Function DiffPlateAnglesAndNormals(ByRef oCandidatePlate As StructDetailObjects.PlatePart, _
                                           ByRef oTargetPlate As StructDetailObjects.PlatePart) As Boolean
    DiffPlateAnglesAndNormals = True

'wrapper method not ready yet - call to this routine should be commented out

    Dim AngleTol As Double
    Dim NormalTol As Double
    
    Dim CandidateCornerData As CornerData
    Dim TargetCornerData As CornerData
    
    Dim oCandidateCornerCollection As IJElements
    Dim oTargetCornerCollection As IJElements
    
    Dim n As Long
    Dim nCount As Long
    Dim bTempDiff As Boolean 'interim varible for DiffPlateAnglesAndNormals

    bTempDiff = True
    AngleTol = 5 ' assumes angle is in degrees
    NormalTol = 0.02 'assumes vector is normalized

    Set oCandidateCornerCollection = oCandidatePlate.GetPlateCornerData(True)
    Set oTargetCornerCollection = oTargetPlate.GetPlateCornerData(True)

    'same number of corners?
    If oCandidateCornerCollection.Count <> oTargetCornerCollection.Count Then
        bTempDiff = True
    ElseIf oCandidateCornerCollection.Count < 3 Then
        bTempDiff = True 'error condition, this should never happen?
    
    'plates have same number of corners
    Else
        'check first corner
        With CandidateCornerData
            .Angle = oCandidateCornerCollection.Item(1).Angle
            Set .EdgePort1 = oCandidateCornerCollection.Item(1).EdgePort1
            Set .EdgePort2 = oCandidateCornerCollection.Item(1).EdgePort2
            Set .FacePort = oCandidateCornerCollection.Item(1).FacePort
            Set .Location = oCandidateCornerCollection.Item(1).Location
            Set .Normal = oCandidateCornerCollection.Item(1).Normal
        End With
        
        With TargetCornerData
            .Angle = oTargetCornerCollection.Item(1).Angle
            Set .EdgePort1 = oTargetCornerCollection.Item(1).EdgePort1
            Set .EdgePort2 = oTargetCornerCollection.Item(1).EdgePort2
            Set .FacePort = oTargetCornerCollection.Item(1).FacePort
            Set .Location = oTargetCornerCollection.Item(1).Location
            Set .Normal = oTargetCornerCollection.Item(1).Normal
        End With
        
        'Normalize normal vector just to be safe
        CandidateCornerData.Normal.Length = 1
        TargetCornerData.Normal.Length = 1
        
        If Abs(CandidateCornerData.Angle - TargetCornerData.Angle) > AngleTol Then
            bTempDiff = True
        ElseIf Abs(CandidateCornerData.Normal.x - TargetCornerData.Normal.x) > NormalTol Then
            bTempDiff = True
        ElseIf Abs(CandidateCornerData.Normal.y - (-1 * TargetCornerData.Normal.y)) > NormalTol Then
            bTempDiff = True
        ElseIf Abs(CandidateCornerData.Normal.z - TargetCornerData.Normal.z) > NormalTol Then
            bTempDiff = True
        Else
            bTempDiff = False
        End If
        
        'if first corner ok
        'check remaining corner angles and normals, since plates should be mirrored
        'check one plate from top of list, and second plate from bottom of list.
        If bTempDiff = False Then
            nCount = oCandidateCornerCollection.Count
            For n = 2 To nCount
                With CandidateCornerData
                    .Angle = oCandidateCornerCollection.Item(n).Angle
                    Set .EdgePort1 = oCandidateCornerCollection.Item(n).EdgePort1
                    Set .EdgePort2 = oCandidateCornerCollection.Item(n).EdgePort2
                    Set .FacePort = oCandidateCornerCollection.Item(n).FacePort
                    Set .Location = oCandidateCornerCollection.Item(n).Location
                    Set .Normal = oCandidateCornerCollection.Item(n).Normal
                End With
                
                With TargetCornerData
                    .Angle = oTargetCornerCollection.Item(nCount + 2 - n).Angle
                    Set .EdgePort1 = oTargetCornerCollection.Item(nCount + 2 - n).EdgePort1
                    Set .EdgePort2 = oTargetCornerCollection.Item(nCount + 2 - n).EdgePort2
                    Set .FacePort = oTargetCornerCollection.Item(nCount + 2 - n).FacePort
                    Set .Location = oTargetCornerCollection.Item(nCount + 2 - n).Location
                    Set .Normal = oTargetCornerCollection.Item(nCount + 2 - n).Normal
                End With
                
                'Normalize normal vector just to be safe
                CandidateCornerData.Normal.Length = 1
                TargetCornerData.Normal.Length = 1
                
                If Abs(CandidateCornerData.Angle - TargetCornerData.Angle) > AngleTol Then
                    bTempDiff = True
                ElseIf Abs(CandidateCornerData.Normal.x - TargetCornerData.Normal.x) > NormalTol Then
                    bTempDiff = True
                ElseIf Abs(CandidateCornerData.Normal.y - (-1 * TargetCornerData.Normal.y)) > NormalTol Then
                    bTempDiff = True
                ElseIf Abs(CandidateCornerData.Normal.z - TargetCornerData.Normal.z) > NormalTol Then
                    bTempDiff = True
                Else
                    bTempDiff = False
                End If
                
                'break out of loop if non similar case is found
                If bTempDiff = True Then n = nCount + 10
            Next
        End If
     End If
     
    If bTempDiff = False Then
        DiffPlateAnglesAndNormals = False
    Else
        DiffPlateAnglesAndNormals = True
    End If
'
    'release collections
    Set oCandidateCornerCollection = Nothing
    Set oTargetCornerCollection = Nothing
'
End Function

Private Function DiffStiffenerParentTypes(ByRef oCandidateStiffener As ProfilePart, _
                                       ByRef oTargetStiffener As ProfilePart) As Boolean
    DiffStiffenerParentTypes = False
    
    ' Everything passes this check for now.
    
    ' Orginal intent was based on assumption that parent of profile system would always
    ' be a plate system.  THis is not true.
    
    ' Therefore submitting a CR requesting new wrapper method that will specifically
    ' return the plate / plate system that the stiffener / stiffener system is
    ' stiffening.
    
    
    

End Function
Private Function DiffStiffenerBoundaries(ByRef oCandidateStiffener As ProfilePart, _
                                       ByRef oTargetStiffener As ProfilePart) As Boolean
    DiffStiffenerBoundaries = True
    
    Dim TempDiff As Boolean
    TempDiff = True
    
    Dim oTargetCollection As Collection
    Dim oCandidateCollection As Collection
    
    Dim oTargetBoundaryData As BoundaryData
    Dim oCandidateBoundaryData As BoundaryData
    Dim n As Long
    
    Set oTargetCollection = oTargetStiffener.ProfileBoundaries
    Set oCandidateCollection = oCandidateStiffener.ProfileBoundaries
    
'   This routine can be re-written after wrapper method is modified to
'   to always return two items in the collection!!!!
'   The comparisons may not be valid if each collection only
'   has one item, do not know which end the boundary data was found.

    
        
    
    
'   if number of boundaries are different then not symmetical
    If oTargetCollection.Count <> oCandidateCollection.Count Then
        DiffStiffenerBoundaries = True

'   if both profiles have no boundaries then they may be similar
'   otherwise need to compare object types and type of object for each boundary
    ElseIf oTargetCollection.Count > 0 Then
        For n = 1 To oTargetCollection.Count
            oTargetBoundaryData = oTargetCollection.Item(n)
            oCandidateBoundaryData = oCandidateCollection.Item(n)
            
            If oTargetBoundaryData.ObjectType <> oCandidateBoundaryData.ObjectType Then
                TempDiff = True
            ElseIf oTargetBoundaryData.TypeObject <> oCandidateBoundaryData.TypeObject Then
                TempDiff = True
            Else
                TempDiff = False
            End If

'           check state after each boundary, cause exit if different boundary found
            If TempDiff = True Then
                DiffStiffenerBoundaries = True
                n = oTargetCollection.Count + 10 'different boundaries found, end loop
            End If
            
'
            If TempDiff = False Then DiffStiffenerBoundaries = False
            
        Next

'   boundaries do not exist for either profile
    Else
        DiffStiffenerBoundaries = False
    End If
    
'
'   release collections
    Set oCandidateCollection = Nothing
    Set oTargetCollection = Nothing
'    Set oTargetBoundaryData = Nothing
 '   Set oCandidateBoundaryData = Nothing
    
End Function
Private Function DiffStiffenerDirections(ByRef oCandidateStiffener As ProfilePart, _
                                       ByRef oTargetStiffener As ProfilePart) As Boolean
    DiffStiffenerDirections = True
    

    Dim TempDiff As Boolean
    TempDiff = True

    Dim TangentTol As Double
    TangentTol = 0.02 'assumes vector is normalized, this allows a 2% difference in each direction

    Dim oCandidateCollection As Collection
    Dim oTargetCollection As Collection
    
    Dim oCandidateTangent As IJDVector
    Dim oTargetTangent As IJDVector

    Set oCandidateCollection = New Collection
    Set oCandidateCollection = oCandidateStiffener.LandingCurveDirection

    Set oTargetCollection = New Collection
    Set oTargetCollection = oTargetStiffener.LandingCurveDirection

'   compare x, y and z tangent components for start and finish points
    Dim n As Long
    For n = 1 To 2
        Set oCandidateTangent = oCandidateCollection.Item(n)
        Set oTargetTangent = oTargetCollection.Item(n)
    
'       assumes tangent is normalized coming from wrapper class
'       should probably normalize it just in case...
        oCandidateTangent.Length = 1
        oTargetTangent.Length = 1
        
'        Open "c:\BoardMgtLog.txt" For Append Access Write As #1
'        Print #1, Now(); Tab; "Candidate xyz"; Tab; oCandidateTangent.x; oCandidateTangent.y; oCandidateTangent.z
'        Print #1, Now(); Tab; "Target xyz"; Tab; oTargetTangent.x; oTargetTangent.y; oTargetTangent.z
'        Close #1
        
        If (Abs(oCandidateTangent.x - oTargetTangent.x) > TangentTol And _
           (Abs(oCandidateTangent.x + oTargetTangent.x) > TangentTol)) Then
            TempDiff = True
'       check mirrored y value
        ElseIf (Abs(oCandidateTangent.y - oTargetTangent.y) > TangentTol And _
               (Abs(oCandidateTangent.y + oTargetTangent.y) > TangentTol)) Then
            TempDiff = True
        ElseIf (Abs(oCandidateTangent.z - oTargetTangent.z) > TangentTol And _
               (Abs(oCandidateTangent.z + oTargetTangent.z) > TangentTol)) Then
            TempDiff = True
        Else
            TempDiff = False
        End If

'       check state at end each loop, cause exit if difference found
        If TempDiff = True Then
            DiffStiffenerDirections = True
            n = 3 ' cause exit from loop
        End If
    Next

'   if both tangent vectors are approximately the same, function should return as False
    If TempDiff = False Then DiffStiffenerDirections = False

'
'   release collections
    Set oCandidateCollection = Nothing
    Set oTargetCollection = Nothing
    Set oCandidateTangent = Nothing
    Set oTargetTangent = Nothing
    
End Function
Private Function DiffSeamBoundaries(ByRef oCandidateSeam As StructDetailObjects.Seam, _
                                    ByRef oTargetSeam As StructDetailObjects.Seam) As Boolean
'psuedo code only - waiting for wrapper methods
'
'    DiffSeamBoundaries = True
'
'    Dim NB As Long
'    Dim oCandidateBoundary As String
'    Dim oTargetBoundary As String
'
'    Dim oCandidateBoundaries As Collection
'    Dim oTargetBoundaries As Collection
'
'    Set oCandidateBoundaries = New Collection
'    Set oCandidateBoundaries = oCandidateSeam.Boundaries 'not sure how to get colletion
'
'    Set oTargetBoundaries = New Collection
'    Set oTargetBoundaries = oTargetSeam.Boundaries
'
'    'same number of boundaries?
'    If oCandidateBoundaries.Count = oTargetBoundaries.Count Then
'        DiffSeamBoundaries = False
'    End If
'
'    'if number of boundaries are the same
'    If DiffSeamBoundaries = False Then
'        'perform check of boundary types
'        NB = oCandidateBoundaries.Count
'        For i = 1 To NB
'            Set oCandidateBoundary = oCandidateBoundaries.Item(i)
'            Set oTargetBoundary = oTargetBoundaries.Item(i)
'
'            'not sure how to get Type from collection
'            If StrComp(oCandidateBoundary.Type, oTargetBoundary.Type) <> 0 Then
'                DiffSeamBoundaries = True
'            End If
'        Next
'    End If
'
'    'release collections
'    Set oCandidateBoundaries = Nothing
'    Set oTargetBoundaries = Nothing
End Function
Private Function DiffSeamDirections(ByRef oCandidateSeam As StructDetailObjects.Seam, _
                                    ByRef oTargetSeam As StructDetailObjects.Seam) As Boolean
    DiffSeamDirections = True
    
    Dim TempDiff As Boolean
    TempDiff = True

    Dim TangentTol As Double
    TangentTol = 0.02 'assumes vector is normalized, this allows a 2% difference in each direction

    Dim oCandidateCollection As Collection
    Dim oTargetCollection As Collection
    
    Dim oCandidateTangent As IJDVector
    Dim oTargetTangent As IJDVector

    Set oCandidateCollection = New Collection
    Set oCandidateCollection = oCandidateSeam.LandingCurveDirection

    Set oTargetCollection = New Collection
    Set oTargetCollection = oTargetSeam.LandingCurveDirection

'   compare x, y and z tangent components for start and finish points
    Dim n As Long
    For n = 1 To 2
        Set oCandidateTangent = oCandidateCollection.Item(n)
        Set oTargetTangent = oTargetCollection.Item(n)
    
'       assumes tangent is normalized coming from wrapper class
'       should probably normalize it just in case...
        oCandidateTangent.Length = 1
        oTargetTangent.Length = 1
        
'        Open "c:\BoardMgtLog.txt" For Append Access Write As #1
'        Print #1, Now(); Tab; "Candidate xyz"; Tab; oCandidateTangent.x; oCandidateTangent.y; oCandidateTangent.z
'        Print #1, Now(); Tab; "Target xyz"; Tab; oTargetTangent.x; oTargetTangent.y; oTargetTangent.z
'        Close #1
        
        'Checking the vector component by component to see whether they are similar or not
        If (Abs(oCandidateTangent.x - oTargetTangent.x) > TangentTol And _
           (Abs(oCandidateTangent.x + oTargetTangent.x) > TangentTol)) Then
            TempDiff = True
'       check mirrored y value
        ElseIf (Abs(oCandidateTangent.y - oTargetTangent.y) > TangentTol And _
               (Abs(oCandidateTangent.y + oTargetTangent.y) > TangentTol)) Then
            TempDiff = True
        ElseIf (Abs(oCandidateTangent.z - oTargetTangent.z) > TangentTol And _
               (Abs(oCandidateTangent.z + oTargetTangent.z) > TangentTol)) Then
            TempDiff = True
        Else
            TempDiff = False
        End If

'       check state at end each loop, cause exit if difference found
        If TempDiff = True Then
            DiffSeamDirections = True
            n = 3 ' cause exit from loop
        End If
    Next

'   if both tangent vectors are approximately the same, function should return as False
    If TempDiff = False Then DiffSeamDirections = False

'
'   release collections
    Set oCandidateCollection = Nothing
    Set oTargetCollection = Nothing
    Set oCandidateTangent = Nothing
    Set oTargetTangent = Nothing
End Function
Private Function DiffSeamIntersections(ByRef oCandidateSeam As StructDetailObjects.Seam, _
                                       ByRef oTargetSeam As StructDetailObjects.Seam) As Boolean
'psuedo code only - waiting for wrapper methods
'
'    DiffSeamIntersections = True
'    Dim CoordTol As Double
'
'    'determine seam orientation to select best reference plane to intersect seam
'        '1 - MostlyVertical - Z
'        '2 - MostlyTransversal - Y
'        '3 - MostlyLongitudinal - X
'        '4 - Vertical - Z
'        '5 - Transversal - Y
'        '6 - Longitudinal - X
'        '-1 - unknown - ????
'
'
'    Dim oCandidateIntersections As Collection
'    Dim oTargetIntersections As Collection
'
'    Set oCandidateIntersections = New Collection
'    Set oTargetIntersections = New Collection
'
'    If oCandidateSeam.Orientation = oTargetSeam.Orientation Then
'        DiffSeamIntersections = False
'        If oCandidateSeam.Orientation = Longitudinal Or MostlyLongitudinal Then
'            Set oCandidateIntersections = oCandidateSeam.SeamRefPlaneIntsections("CS_0", "X")
'            Set oTargetIntersections = oTargetSeam.SeamRefPlaneIntsections("CS_0", "X")
'        ElseIf oCandidateSeam.Orientation = Vertical Or MostlyVertical Then
'            Set oCandidateIntersections = oCandidateSeam.SeamRefPlaneIntsections("CS_0", "Z")
'            Set oTargetIntersections = oTargetSeam.SeamRefPlaneIntsections("CS_0", "Z")
'        ElseIf oCandidateSeam.Orientation = Transveral Or MostlyTransveral Then
'            Set oCandidateIntersections = oCandidateSeam.SeamRefPlaneIntsections("CS_0", "Y")
'            Set oTargetIntersections = oTargetSeam.SeamRefPlaneIntsections("CS_0", "Y")
'        Else
'            DiffSeamIntersections = True
'        End If
'    End If
'
'    'same number of Intersections?
'    If DiffSeamIntersections = False Then
'        If oCandidateIntersections.Count = oTargetIntersections.Count Then
'            DiffSeamIntersections = False
'        End If
'    End If
'
'    'if number of intersections are the same
'    If DiffSeamIntersections = False Then
'        'perform check of coordinates, include allowance for specified tolerance
'        CoordTol = 0.5 'mm ?????
'        NC = oCandidateIntersections.Count
'        For i = 1 To NC
'            Set oCandidateInfo = oCandidateIntersections.Item(i)
'            Set oTargetInfo = oTargetIntersections.Item(i)
'
'            'not sure how to get normals from collection
'            If Abs(oCandidateInfo.Xcoord - oTargetInfo.Xcoord) > CoordTol Then
'                DiffSeamIntersections = True
'            ElseIf Abs(oCandidateInfo.Ycoord - (-1 * oTargetInfo.Ycoord)) > CoordTol Then
'                DiffSeamIntersections = True
'            ElseIf Abs(oCandidateInfo.Zcoord - oTargetInfo.Zcoord) > CoordTol Then
'                DiffSeamIntersections = True
'            End If
'        Next
'    End If
'
'    'release collections
'    Set oCandidateIntersections = Nothing
'    Set oTargetIntersections = Nothing
'
End Function
Private Function CheckOpeningSymmetry(oCandidateOpening As IJOpeningEntity, otargetOpening As IJOpeningEntity) As Boolean
    Dim oCurve1 As IJCurve
    Dim oCurve2 As IJCurve
    
    Dim CircumTol As Double
    CircumTol = 0.000001

    Set oCurve1 = oCandidateOpening
    Set oCurve2 = otargetOpening
    ' It is always advisable to have a tolerance when comparing a double precision
    If Abs(oCurve1.Length - oCurve2.Length) < CircumTol Then
        CheckOpeningSymmetry = True
    Else
        CheckOpeningSymmetry = False
    End If
End Function
Private Function CheckSketchFeatureSymmetry(oCandidateFeature As Object, oTargetFeature As Object) As Boolean
    
    Dim oWire1 As IJStructGenericContour
    Dim oWire2 As IJStructGenericContour
    
     Dim CircumTol As Double
     CircumTol = 0.000001

    Set oWire1 = oCandidateFeature
    Set oWire2 = oTargetFeature
    
    Dim pAttrWire1 As IJWireBody
    Dim pAttrWire2 As IJWireBody
    
    oWire1.GetAttributedWire pAttrWire1
    oWire2.GetAttributedWire pAttrWire2
    
    Dim oCurve1 As IJCurve
    Dim oCurve2 As IJCurve
    
    If (Not pAttrWire1 Is Nothing) And (Not pAttrWire2 Is Nothing) Then
      
        Dim oSgoWireBodyUtils As SGOWireBodyUtilities
        Set oSgoWireBodyUtils = New SGOWireBodyUtilities
        
        Set oCurve1 = oSgoWireBodyUtils.GetCurveFromWireBody(pAttrWire1)
        Set oCurve2 = oSgoWireBodyUtils.GetCurveFromWireBody(pAttrWire2)
       ' It is always advisable to have a tolerance when comparing a double precision
   
        If Abs(oCurve1.Length - oCurve2.Length) < 0.000001 Then
          
            CheckSketchFeatureSymmetry = True
        Else
            CheckSketchFeatureSymmetry = False
        End If
    End If
       
        
End Function




Private Function IJBoardManagementRule_IsSymmetrical(ByVal oCandidatePart As Object, ByVal oTargetPart As Object) As Boolean

 On Error GoTo ErrorHandler
    
    'turn logging on/off
    Dim LoggingOn As Boolean
    LoggingOn = False
    Dim sLogFileName As String
    sLogFileName = "c:\BoardMgtLog.txt"
    Dim IsSymmetrical As Boolean
    
    'set defaults
    IsSymmetrical = False
    
    Dim oCandNamedItem As IJNamedItem
    Set oCandNamedItem = oCandidatePart
    Dim oTargNamedItem As IJNamedItem
    Set oTargNamedItem = oTargetPart
   
    
    'determine object types
    Dim sPartKind As String
    If (TypeOf oCandidatePart Is IJPlatePart And TypeOf oTargetPart Is IJPlatePart) Then
        sPartKind = "PlatePart"
    ElseIf (TypeOf oCandidatePart Is IJStiffenerPart And _
            TypeOf oTargetPart Is IJStiffenerPart) Then
        sPartKind = "StiffenerPart"
    ElseIf (TypeOf oCandidatePart Is IJBeamPart And TypeOf oTargetPart Is IJBeamPart) Then
        sPartKind = "BeamPart"
    ElseIf (TypeOf oCandidatePart Is IJSeam And TypeOf oTargetPart Is IJSeam) Then
        sPartKind = "Seam"
    ElseIf (TypeOf oCandidatePart Is IJOpeningEntity) And (TypeOf oTargetPart Is IJOpeningEntity) Then
         sPartKind = "Opening"
    ElseIf ((TypeOf oCandidatePart Is IJSketchFeature) And (TypeOf oTargetPart Is IJSketchFeature)) Then
         sPartKind = "SketchFeature"
    ElseIf ((oCandNamedItem.TypeString = "Hole Trace") And (oTargNamedItem.TypeString = "Hole Trace")) Then
    
        sPartKind = "HoleTrace"
    ElseIf ((oCandNamedItem.TypeString = "StructFeature") And (oTargNamedItem.TypeString = "StructFeature")) Then
    
        sPartKind = "StructFeature"
    Else
    
        sPartKind = "NotSupported"
    End If
    
    
    Select Case sPartKind
    
        Case "PlatePart"
        
            Dim oCandidatePlate As StructDetailObjects.PlatePart
            Dim oTargetPlate As StructDetailObjects.PlatePart
            
            Set oCandidatePlate = New StructDetailObjects.PlatePart
            Set oCandidatePlate.object = oCandidatePart
            
            Set oTargetPlate = New StructDetailObjects.PlatePart
            Set oTargetPlate.object = oTargetPart
            
            If StrComp(oCandidatePlate.Material.MaterialType, _
                        oTargetPlate.Material.MaterialType) <> 0 Then
                IsSymmetrical = False
            ElseIf StrComp(oCandidatePlate.Grade, oTargetPlate.Grade) <> 0 Then
                IsSymmetrical = False
            ElseIf oCandidatePlate.PlateThickness <> oTargetPlate.PlateThickness Then
                IsSymmetrical = False
'            ElseIf DiffPlateBoundaries(oCandidatePlate, oTargetPlate) Then
'                IsSymmetrical = False
'            ElseIf DiffPlateAnglesAndNormals(oCandidatePlate, oTargetPlate) Then
'                IsSymmetrical = False
            Else
                IsSymmetrical = True
            End If
            
            'log results
            If LoggingOn Then
                Open sLogFileName For Append Access Write As #1
                Print #1, Now(); Tab; "Plate"; Tab; oCandidatePlate.Name; Tab; oTargetPlate.Name; Tab; IsSymmetrical
                Close #1
            End If
            
            'release objects
            Set oCandidatePlate = Nothing
            Set oTargetPlate = Nothing
            
        Case "StiffenerPart"
        
            Dim oCandidateStiffener As StructDetailObjects.ProfilePart
            Dim oTargetStiffener As StructDetailObjects.ProfilePart
            
            Set oCandidateStiffener = New StructDetailObjects.ProfilePart
            Set oCandidateStiffener.object = oCandidatePart
            
            Set oTargetStiffener = New StructDetailObjects.ProfilePart
            Set oTargetStiffener.object = oTargetPart
            
            If StrComp(oCandidateStiffener.MaterialType, oTargetStiffener.MaterialType) <> 0 Then
                IsSymmetrical = False
            ElseIf StrComp(oCandidateStiffener.Grade, oTargetStiffener.Grade) <> 0 Then
                IsSymmetrical = False
            ElseIf StrComp(oCandidateStiffener.SectionType, oTargetStiffener.SectionType) <> 0 Then
                IsSymmetrical = False
            ElseIf StrComp(oCandidateStiffener.SectionName, oTargetStiffener.SectionName) <> 0 Then
                IsSymmetrical = False
            ElseIf DiffStiffenerParentTypes(oCandidateStiffener, oTargetStiffener) Then
                IsSymmetrical = False
            ElseIf DiffStiffenerBoundaries(oCandidateStiffener, oTargetStiffener) Then
                IsSymmetrical = False
            ElseIf DiffStiffenerDirections(oCandidateStiffener, oTargetStiffener) Then
                IsSymmetrical = False
            Else
                IsSymmetrical = True
            End If
            
            'log results
            If LoggingOn Then
                Open sLogFileName For Append Access Write As #1
                Print #1, Now(); Tab; "Stiffener"; Tab; oCandidateStiffener.Name; Tab; oTargetStiffener.Name; Tab; IsSymmetrical
                Close #1
            End If
            
            
            'release objects
            Set oCandidateStiffener = Nothing
            Set oTargetStiffener = Nothing
        
        Case "BeamPart"
        
            Dim oCandidateBeam As StructDetailObjects.BeamPart
            Dim oTargetBeam As StructDetailObjects.BeamPart
            
            Set oCandidateBeam = New StructDetailObjects.BeamPart
            Set oCandidateBeam.object = oCandidatePart
            
            Set oTargetBeam = New StructDetailObjects.BeamPart
            Set oTargetBeam.object = oTargetPart
            
    '        If StrComp(oCandidateBeam.MaterialType, oTargetBeam.MaterialType) <> 0 Then
    '            IsSymmetrical = False
    '        ElseIf StrComp(oCandidateBeam.Grade, oTargetBeam.Grade) <> 0 Then
    '            IsSymmetrical = False
    '        ElseIf StrComp(oCandidateBeam.SectionType, oTargetBeam.SectionType) <> 0 Then
    '            IsSymmetrical = False
    '        ElseIf StrComp(oCandidateBeam.SectionName, oTargetBeam.SectionName) <> 0 Then
    '            IsSymmetrical = False
           If oCandidateBeam.BeamType <> oTargetBeam.BeamType Then
                IsSymmetrical = False
    '        ElseIf DiffBeamBoundaries(oCandidateBeam, oTargetBeam) Then
    '            IsSymmetrical = False
    '        ElseIf DiffBeamDirections(oCandidateBeam, oTargetBeam) Then
    '            IsSymmetrical = False
            Else
                IsSymmetrical = True
            End If
            
            'log results
            If LoggingOn Then
                Open sLogFileName For Append Access Write As #1
                Print #1, Now(); Tab; "Beam"; Tab; oCandidateBeam.Name; Tab; oTargetBeam.Name; Tab; IsSymmetrical
                Close #1
            End If
            
            
            'release objects
            Set oCandidateBeam = Nothing
            Set oTargetBeam = Nothing
        
        Case "Seam"
        
            Dim oCandidateSeam As StructDetailObjects.Seam
            Dim oTargetSeam As StructDetailObjects.Seam
            
            Set oCandidateSeam = New StructDetailObjects.Seam
            Set oCandidateSeam.object = oCandidatePart
            
            Set oTargetSeam = New StructDetailObjects.Seam
            Set oTargetSeam.object = oTargetPart
            
            Dim oCandidateParent As IJPlate
            Set oCandidateParent = oCandidateSeam.ParentSystem
            
            Dim oTargetParent As IJPlate
            Set oTargetParent = oTargetSeam.ParentSystem
            
            
            If oCandidateParent.PlateType <> oTargetParent.PlateType Then
                IsSymmetrical = False
                
            ElseIf oCandidateSeam.Orientation <> oTargetSeam.Orientation Then
                IsSymmetrical = False
                'Print #1, Now(); Tab; "Seam-ori"; Tab; oCandidateSeam.Orientation; Tab; oTargetSeam.Orientation;
                'MsgBox ("Orientation" & oCandidateSeam.Orientation & oTargetSeam.Orientation)
    '        ElseIf DiffSeamBoundaries(oCandidateSeam, oTargetSeam) Then
    '            IsSymmetrical = False
            ElseIf DiffSeamDirections(oCandidateSeam, oTargetSeam) Then
                IsSymmetrical = False
                'Print #1, Now(); Tab; "Seam-dir";
    '        ElseIf DiffSeamIntersections(oCandidateSeam, oTargetSeam) Then
    '            IsSymmetrical = False
            Else
                IsSymmetrical = True
            End If
            
            
            'log results
            If LoggingOn Then
                Open sLogFileName For Append Access Write As #1
                Print #1, Now(); Tab; "Seam"; Tab; oCandidateSeam.Name; Tab; oTargetSeam.Name; Tab; IsSymmetrical
                Close #1
            End If
            
            'release objects
            Set oCandidateSeam = Nothing
            Set oTargetSeam = Nothing
            
            Set oCandidateParent = Nothing
            Set oTargetParent = Nothing
        
        Case "Opening"
            IsSymmetrical = CheckOpeningSymmetry(oCandidatePart, oTargetPart)
        
        Case "SketchFeature", _
             "StructFeature"
            IsSymmetrical = CheckSketchFeatureSymmetry(oCandidatePart, oTargetPart)
        Case "HoleTrace"
            IsSymmetrical = CheckSketchHoleTraceSymmetry(oCandidatePart, oTargetPart)
        Case Else
            IsSymmetrical = False
            
            'log results
            If LoggingOn Then
                Open sLogFileName For Append Access Write As #1
                Print #1, Now(); Tab; "Objects not detected to be a Plate, Stiffener, Beam or Seam"; Tab; IsSymmetrical
                Close #1
            End If
            
            
    End Select
    IJBoardManagementRule_IsSymmetrical = IsSymmetrical

Exit Function

ErrorHandler:
    If LoggingOn Then
        Open sLogFileName For Append Access Write As #1
        Print #1, Now(); Tab; "Exception caught by ErrorHandler"; Tab; IsSymmetrical
        Close #1
    End If
Debug.Assert False


End Function

Private Function CheckSketchHoleTraceSymmetry(oCandidateTrace As Object, oTargetTrace As Object) As Boolean
  
    On Error Resume Next
    Dim oWire1 As IJStructGenericContour
    Dim oWire2 As IJStructGenericContour
    
    'Passed in objects are hole traces,get corresponding features based on "HoleTraceFeature" relationship
    Dim oAssocRelation As IJDAssocRelation
    Dim oTargetObjCol As IJDTargetObjectCol
    
    Dim oCandidateHoleFeature As Object
    Dim oTargetHoleFeature As Object
    
    Set oAssocRelation = oCandidateTrace
    Set oTargetObjCol = oAssocRelation.CollectionRelations(IID_IJHoleTraceAE, "HoleTraceFeature_Feature")
    Set oAssocRelation = Nothing
    If Not oTargetObjCol Is Nothing Then
        Set oCandidateHoleFeature = oTargetObjCol.Item(1)
        If Not oCandidateHoleFeature Is Nothing Then
            Set oWire1 = oCandidateHoleFeature
            Set oCandidateHoleFeature = Nothing
        End If
        Set oTargetObjCol = Nothing
    End If
    
    Set oAssocRelation = oTargetTrace
    Set oTargetObjCol = oAssocRelation.CollectionRelations(IID_IJHoleTraceAE, "HoleTraceFeature_Feature")
    Set oAssocRelation = Nothing
    If Not oTargetObjCol Is Nothing Then
        Set oTargetHoleFeature = oTargetObjCol.Item(1)
        If Not oTargetHoleFeature Is Nothing Then
            Set oWire2 = oTargetHoleFeature
            Set oTargetHoleFeature = Nothing
        End If
        Set oTargetObjCol = Nothing
    End If
        
    If oWire1 Is Nothing Or oWire2 Is Nothing Then
        CheckSketchHoleTraceSymmetry = False
        Exit Function
    End If
    
    Dim CircumTol As Double
    CircumTol = 0.000001

    Dim pAttrWire1 As IJWireBody
    Dim pAttrWire2 As IJWireBody
    
    oWire1.GetAttributedWire pAttrWire1
    oWire2.GetAttributedWire pAttrWire2
    
    Dim oCurve1 As IJCurve
    Dim oCurve2 As IJCurve
    
    If (Not pAttrWire1 Is Nothing) And (Not pAttrWire2 Is Nothing) Then
    
        Dim oSgoWireBodyUtils As SGOWireBodyUtilities
        Set oSgoWireBodyUtils = New SGOWireBodyUtilities
        
        Set oCurve1 = oSgoWireBodyUtils.GetCurveFromWireBody(pAttrWire1)
        Set oCurve2 = oSgoWireBodyUtils.GetCurveFromWireBody(pAttrWire2)
        If Abs(oCurve1.Length - oCurve2.Length) < CircumTol Then
          
            CheckSketchHoleTraceSymmetry = True
        Else
            CheckSketchHoleTraceSymmetry = False
        End If
    End If
       
        
End Function

