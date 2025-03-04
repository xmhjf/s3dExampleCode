VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarkingMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''
''   Copyright (C) 2002, Intergraph Corporation. All rights reserved.
''
''   Abstract:
''              MarkingMigration.cls
''   ProgID:
''              StrMfgMigrationRules.MarkingMigration
''   Author:
''              Suresh Kadambari.
''   Creation Date:
''              30 Oct 2006
''   Description:
''      Struct Mfg should check split notification and keep the marking lines.
''
''      In a simplest test, a part "P1" is split into two smaller parts "P2" and "P3" and "P1" will be deleted.
''      Old part P1, along with new parts P2 and P3 will be passed to a user defined VB rule.  This rule will be
''      bulkloaded into the catalog data in StructMfgSetting.xls spread sheet.
''
''   Change History:
''   dd.mmm.yyyy     who     change description
''   -----------     ---     ------------------
''
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'

Option Explicit
Const MODULE = "MarkingMigration"
Implements IJMfgSplitMigration


Private Function IJMfgSplitMigration_MigrateObject(ByVal pOldObj As Object, ByVal pReplacingObjColl As GSCADStructMfgGlobals.IJElements, ByVal pMfgObject As Object, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
'Arguments:
' pOldObj           ==> part "P1" which is being deleted.
' pReplacingObjColl ==> Collection of replacing new parts ( P2 and P3 ).
' pbCreateMfg       ==> Boolean flag indicating if software should automatically create Markinglines
'                       for the new parts.

' Output of this routine will be one of the part in pReplacingObjColl collection.
' NOTE:  The rule is not limited to two parts.  "n" number of parts could be passed to the rule
' after a split.  The rule would still only return one part meeting the criteria. But, it can create
' new markinglines on other parts based on the boolean flag pbCreateMfg
'
' pbCreateMfg   Output of MigrateObject()   Action
' ===========   =========================   ======
' False         Nothing           Marking ( ML1 ) will be deleted. NO new marking lines will be created.
' False         Object P2         Marking ( ML1 ) will be re-connected from P1 to P2. New marking lines will NOT be created for the remaining new parts ( for P3 ).
' True          Nothing           Marking ( ML1 ) will be deleted. New marking lines will be created for all the new parts ( for P2 and P3 ).
' True          Object P2         Marking ( ML1 ) will be re-connected from P1 to P2.  New marking lines will be created for the remaining new parts ( for P3 ).
'
Const METHOD = "IJMfgSplitMigration_MigrateObject"
On Error GoTo ErrorHandler

    Dim lNewPartCount As Long
    Dim indPart As Long
    Dim dMaxSurfaceArea As Double
    Dim lMaxSurfaceIndex As Long

    'Don't create markinglines automatically for the new parts.
    pbCreateMfg = False ' test with false

    'Get the new part count
    lNewPartCount = pReplacingObjColl.Count

    ' If the new part count is zero, there is nothing to do.
    If lNewPartCount = 0 Then
        Exit Function
    End If

    Dim oRuleHlpr As IJMfgRuleHelpersSupport
    Set oRuleHlpr = New MfgRuleHelpersSupport
        
    Dim oPartsCollection As IJElements
    Set oPartsCollection = New JObjectCollection
    
    ' Loop for each element in the new part collection.
    Dim oDetailedPart  As Object
    
    For indPart = 1 To lNewPartCount
        Dim bOverlap As Boolean
        bOverlap = False
        
        Set oDetailedPart = pReplacingObjColl.Item(indPart)
        bOverlap = oRuleHlpr.RangesOverlap(oDetailedPart, pMfgObject)
        
        If bOverlap = True Then
            ' add the parts which are overlapping to this range in a separate collection
            oPartsCollection.Add oDetailedPart ', Str$(iKey)
        End If
        
        Set oDetailedPart = Nothing
    Next indPart

    'Get the new part count
    lNewPartCount = oPartsCollection.Count
    
    'Something went wrong. Quit here
    If lNewPartCount = 0 Then
        Exit Function
    End If
        
    ' validate the range range check to determine if there is really topology overlap
    oPartsCollection.WaitForUpdate = True
    
    If lNewPartCount > 1 Then
        ' Initialization.
        dMaxSurfaceArea = -1# 'maximum surface area
        lMaxSurfaceIndex = 1 'index of the biggest  part in pReplacingObjColl
        
        Dim oTopologyIntersect As IJDTopologyIntersect
        Set oTopologyIntersect = New DGeomOpsIntersect
        
        For indPart = 1 To lNewPartCount
            bOverlap = True
            
            Set oDetailedPart = oPartsCollection.Item(indPart)
            
            Dim oMfgMarkingLine As IJMfgMarkingLines_AE
            Set oMfgMarkingLine = pMfgObject
            
            Dim oMarkingGeometry As IJComplexString
            Set oMarkingGeometry = oMfgMarkingLine.GeometryAsComplexString
            
            oTopologyIntersect.HasOverlappingGeometry oDetailedPart, oMarkingGeometry, bOverlap
            If bOverlap = False Then
                oPartsCollection.Remove (indPart)
            End If
        Next indPart
        
    End If

    oPartsCollection.WaitForUpdate = False

    'Get the new part count
    lNewPartCount = oPartsCollection.Count
    
    If lNewPartCount > 1 Then
        
        ' from the moniker check if it supports ijplatepart or ijprofilepart interfaces
        If TypeOf pOldObj Is IJPlatePart Then
        
            ' Loop for each element in the new part collection.
            For indPart = 1 To lNewPartCount
                Dim oPlatePartSupport As IJPlatePartSupport
                Dim oBaseSurface As IJSurfaceBody
        
                'Create PlatePart support to get the base surface of the plate without any features/camfers
                Set oPlatePartSupport = New PlatePartSupport
        
                Dim oPartSupport As IJPartSupport
                Set oPartSupport = oPlatePartSupport
        
                'Get the base surface of this plate part
                Set oPartSupport.Part = oPartsCollection.Item(indPart)
        
                oPlatePartSupport.GetSurface PlateBaseSide, oBaseSurface
        
                Dim oBaseSurfaceModelBody As IJDModelBody
                Set oBaseSurfaceModelBody = oBaseSurface
        
                'Calculate the surface area
                Dim dDummyLength As Double
                Dim dEstRelAccyAchieved As Double
                Dim dSurfaceArea As Double
                Dim dDummyVolume As Double
        
                oBaseSurfaceModelBody.GetDimMetrics 0.001, dEstRelAccyAchieved, dDummyLength, dSurfaceArea, dDummyVolume
        
                ' See if this area is larger than the previous one
                If dSurfaceArea > dMaxSurfaceArea Then
                     ' If so, store this area as maximum and the element index
                     dMaxSurfaceArea = dSurfaceArea
                     lMaxSurfaceIndex = indPart
                End If
                Set oBaseSurfaceModelBody = Nothing
        
                Set oPlatePartSupport = Nothing
                Set oPartSupport = Nothing
                Set oBaseSurface = Nothing
            Next indPart
            
            'Something went wrong. Quit here
            If lMaxSurfaceIndex = 0 Then
                Exit Function
            End If
            
            'Return the plate part with largest surface area.
            Set IJMfgSplitMigration_MigrateObject = oPartsCollection.Item(lMaxSurfaceIndex)
            
        ElseIf TypeOf pOldObj Is IJStructProfilePart Then
            
            ' Initialization.
            Dim dMaxLength As Double
            Dim lMaxLengthIndex As Long
            dMaxLength = -1# 'maximum landing curve length
            lMaxLengthIndex = 1
    
            
            ' Loop for each element in the new Profile collection.
            Dim indProfile As Long
            
            For indProfile = 1 To oPartsCollection.Count
                
                'Create ProfilePart support
                Dim oProfilePartSupport As IJProfilePartSupport
                Set oProfilePartSupport = New ProfilePartSupport
                Set oPartSupport = oProfilePartSupport
                
                Dim dLength As Double
                'Get the base surface of this plate part
                Set oPartSupport.Part = oPartsCollection.Item(indProfile)
    
                dLength = oProfilePartSupport.ApproximateLength

                ' See if this length is larger than the previous one
                If dLength > dMaxLength Then
                     ' If so, store this length as maximum and the element index
                    dMaxLength = dLength
                    lMaxLengthIndex = indProfile
                End If
    
                dLength = -1#
            Next indProfile
    
            Set oProfilePartSupport = Nothing
            Set oPartSupport = Nothing
            ' take the first element. ideally it must be the profile with the biggest length
            Set IJMfgSplitMigration_MigrateObject = oPartsCollection.Item(lMaxLengthIndex)
            
        End If
        
        ' create markings on other parts as they are having overlapping geometry
        pbCreateMfg = True ' true indicates create new marking lines on ther parts
        
    Else
        Set IJMfgSplitMigration_MigrateObject = oPartsCollection.Item(1)
    End If
    
    ' no setting to clone
    pOptionalArg = False
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4003, , "RULES")
End Function
'If plate, return the plate with maximum surface area. If Profile/member, return the one having maximum length
Private Function IJMfgSplitMigration_ReverseMigrate(ByVal pReplacedPartColl As GSCADStructMfgGlobals.IJElements, ByVal pReplacingPart As Object, ByVal pMfgObjsColl As GSCADStructMfgGlobals.IJElements, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
    Const METHOD = "IJMfgSplitMigration_ReverseMigrate"
    On Error GoTo ErrorHandler
        
    pbCreateMfg = False ' set this to true only if migration is not needed but want to create a new marking line
    pOptionalArg = False ' this makes send to clone the settings only when the above the flag is true

    ' Initialization.
    Dim indMarking          As Long
    Dim lMaxIndex           As Long
    Dim oPartSupport        As IJPartSupport
    Dim oMfgMarking         As IJMfgChild
    Dim oMfgMarkingFolder   As IJMfgChild

    lMaxIndex = 1 'index of the biggest plate part in pReplacingObjColl

    Dim oMfgUpdateInfo As IJMfgUpdateInfo
    
    If TypeOf pReplacingPart Is IJStructProfilePart Then
    
        ' Initialization.
        Dim dMaxLength As Double
        dMaxLength = -1# 'maximum landing curve length

        
        ' Loop for each element in the new Profile collection.
        For indMarking = 1 To pMfgObjsColl.Count
            Set oMfgMarking = pMfgObjsColl.Item(indMarking)
    
            Set oMfgMarkingFolder = oMfgMarking.GetParent
            
            Set oMfgUpdateInfo = oMfgMarking
            oMfgUpdateInfo.UpToDate = FLAG_PART_DELETED
        
            'Create ProfilePart support
            Dim oProfilePartSupport As IJProfilePartSupport
            Set oProfilePartSupport = New ProfilePartSupport
            Set oPartSupport = oProfilePartSupport
            
            Dim dLength As Double
                
            'Get the base surface of this plate part
            Set oPartSupport.Part = oMfgMarkingFolder.GetParent

            dLength = oProfilePartSupport.ApproximateLength

            ' See if this length is larger than the previous one
            If dLength > dMaxLength Then
                 ' If so, store this length as maximum and the element index
                dMaxLength = dLength
                lMaxIndex = indMarking
            End If

            dLength = -1#
        Next indMarking

        Set oProfilePartSupport = Nothing
        Set oPartSupport = Nothing
    
    Else
        
        Dim dMaxSurfaceArea     As Double
        dMaxSurfaceArea = -1# 'maximum surface area

        'Create PlatePart support to get the base surface of the plate without any features/camfers
        Dim oPlatePartSupport As IJPlatePartSupport
        Set oPlatePartSupport = New PlatePartSupport
    
        Set oPartSupport = oPlatePartSupport
        
        ' Loop for each element in the manufactruing objects collection.
        For indMarking = 1 To pMfgObjsColl.Count
        
            Set oMfgMarking = pMfgObjsColl.Item(indMarking)
            Set oMfgMarkingFolder = oMfgMarking.GetParent
            
            Set oMfgUpdateInfo = oMfgMarking
            oMfgUpdateInfo.UpToDate = FLAG_PART_DELETED
            
            'Get the base surface of this plate part
            Set oPartSupport.Part = oMfgMarkingFolder.GetParent
    
            Dim oBaseSurface As IJSurfaceBody
            oPlatePartSupport.GetSurface PlateBaseSide, oBaseSurface
    
            Dim oBaseSurfaceModelBody As IJDModelBody
            Set oBaseSurfaceModelBody = oBaseSurface
    
            'Calculate the surface area
            Dim dDummyLength            As Double
            Dim dEstRelAccyAchieved     As Double
            Dim dSurfaceArea            As Double
            Dim dDummyVolume            As Double
    
            oBaseSurfaceModelBody.GetDimMetrics 0.001, dEstRelAccyAchieved, dDummyLength, dSurfaceArea, dDummyVolume
    
            ' See if this area is larger than the previous one
            If dSurfaceArea > dMaxSurfaceArea Then
                 ' If so, store this area as maximum and the element index
                 dMaxSurfaceArea = dSurfaceArea
                 lMaxIndex = indMarking
            End If
            
            Set oBaseSurfaceModelBody = Nothing
            Set oBaseSurface = Nothing
            
        Next indMarking
        
        Set oPlatePartSupport = Nothing
        Set oPartSupport = Nothing
    End If

    Set oMfgUpdateInfo = pMfgObjsColl.Item(lMaxIndex)
    oMfgUpdateInfo.UpToDate = FLAG_MIGRATED
    
    'Return the manufacturing part with largest surface area.
    Set IJMfgSplitMigration_ReverseMigrate = pMfgObjsColl.Item(lMaxIndex)
   
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1035, , "RULES")
End Function
