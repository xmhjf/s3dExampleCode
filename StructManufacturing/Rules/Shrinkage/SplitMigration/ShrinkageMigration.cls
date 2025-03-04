VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShrinkageMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJMfgSplitMigration
Private Const MODULE = "ShrinkageMigration"

Private Function IJMfgSplitMigration_MigrateObject(ByVal pOldObj As Object, ByVal pReplacingObjColl As GSCADStructMfgGlobals.IJElements, ByVal pMfgObject As Object, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
'Arguments:
' pOldObj ==>  Plate "P1" which is being deleted.
' pReplacingObjColl ==> Collection of replacing new plates ( P2 and P3 ).
' pbCreateMfg ==>   Boolean flag indicating if software should automatically create Shrinkage for the new plates.
' Output of this routine will be one of the plate in pReplacingObjColl collection.
' NOTE:  The rule is not limited to two plates.  "n" number of plates could be passed to the rule after a split.  The rule would still only return one plate meeting the criteria.
'
' pbCreateMfg   Output of MigrateObject()   Action
' ===========   =========================   ======
' False         Nothing           Shrinkage ( S1 ) will be deleted. NO new Shrinkage will be created.
' True          Nothing           Shrinkage ( S1 ) will be deleted. New Shrinkage  will be created for all the new plates ( for P2 and P3 ).
' False         Object P2         Shrinkage ( S1 ) will be re-connected from P1 to P2. New Shrinkage  will NOT be created for the remaining new plates ( for P3 ).
' True          Object P2         Shrinkage ( S1 ) will be re-connected from P1 to P2.  New Shrinkage will be created for the remaining new plates ( for P3 ).

    Const METHOD = "IJMfgSplitMigration_MigrateObject"
    On Error GoTo ErrorHandler

    Dim oShrinkage As IJScalingShr
    Set oShrinkage = pMfgObject
    'Create Shrinkage automatically for the new Parts.
    pbCreateMfg = True
    Dim oPartSupport As IJPartSupport

    If TypeOf pOldObj Is IJStructProfilePart Then
        
        Dim lNewProfileCount As Long
        Dim indProfile As Long
        Dim dMaxLength As Double
        Dim dMaxLengthIndex As Long
        
        'Get the new Profile count
        lNewProfileCount = pReplacingObjColl.Count

        ' If the new Profile count is zero, there is nothing to do.
        If lNewProfileCount = 0 Then
            Exit Function
        End If

        ' Initialization.
        dMaxLength = -1# 'maximum landing curve length
        dMaxLengthIndex = 1 'index of the biggest Profile part in pReplacingObjColl

        ' Loop for each element in the new Profile collection.
        For indProfile = 1 To lNewProfileCount
            Dim dLength As Double
            Dim oProfilePartSupport As IJProfilePartSupport
            
            'Create ProfilePart support
            Set oProfilePartSupport = New ProfilePartSupport
            Set oPartSupport = oProfilePartSupport

            Set oPartSupport.Part = pReplacingObjColl.Item(indProfile)

            dLength = oProfilePartSupport.ApproximateLength
    
            ' See if this length is larger than the previous one
            If dLength > dMaxLength Then
                 ' If so, store this length as maximum and the element index
                dMaxLength = dLength
                dMaxLengthIndex = indProfile
            End If

            dLength = -1#

            Set oProfilePartSupport = Nothing
            Set oPartSupport = Nothing

        Next indProfile

        'Something went wrong. Quit here
        If dMaxLengthIndex = 0 Then
            Exit Function
        End If
        
        'Return the Profile part with largest landing curve length.
        Set IJMfgSplitMigration_MigrateObject = pReplacingObjColl.Item(dMaxLengthIndex)
        
        Exit Function
    
    Else 'This is for Plate Shrinkage.
        Dim lNewPlateCount As Long
        Dim indPlate As Long
        Dim dMaxSurfaceArea As Double
        Dim lMaxSurfaceIndex As Long
    
        'Get the new Plate count
        lNewPlateCount = pReplacingObjColl.Count
    
        ' If the new plate count is zero, there is nothing to do.
        If lNewPlateCount = 0 Then
            Exit Function
        End If
    
        ' Initialization.
        dMaxSurfaceArea = -1# 'maximum surface area
        lMaxSurfaceIndex = 0 'index of the biggest plate part in pReplacingObjColl
    
        ' Loop for each element in the new plate collection.
        For indPlate = 1 To lNewPlateCount
            Dim oPlatePartSupport As IJPlatePartSupport
            Dim oBaseSurface As IJSurfaceBody
    
            'Create PlatePart support to get the base surface of the plate without any features/camfers
            Set oPlatePartSupport = New PlatePartSupport
    
            Set oPartSupport = oPlatePartSupport
    
            'Get the base surface of this plate part
            Set oPartSupport.Part = pReplacingObjColl.Item(indPlate)
    
             If (TypeOf oPartSupport.Part Is IJPlatePart) Then
            'If oPartSupport.IsDetailedPart Then
    
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
                     lMaxSurfaceIndex = indPlate
                End If
                Set oBaseSurfaceModelBody = Nothing
            End If
    
            Set oPlatePartSupport = Nothing
            Set oPartSupport = Nothing
            Set oBaseSurface = Nothing
    
        Next indPlate
    
        'Something went wrong. Quit here
        If lMaxSurfaceIndex = 0 Then
            Exit Function
        End If
        
        'Return the plate part with largest surface area.
         Set IJMfgSplitMigration_MigrateObject = pReplacingObjColl.Item(lMaxSurfaceIndex)
         
    End If
    
    ' no settings to clone
    pOptionalArg = False
    
   Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7003, , "RULES")
End Function

Private Function IJMfgSplitMigration_ReverseMigrate(ByVal pReplacedPartColl As GSCADStructMfgGlobals.IJElements, ByVal pReplacingPart As Object, ByVal pMfgObjsColl As GSCADStructMfgGlobals.IJElements, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
    Const METHOD = "IJMfgSplitMigration_ReverseMigrate"
    On Error GoTo ErrorHandler

    'Return the plate part with largest surface area.
        
    pbCreateMfg = False ' set this to true only if migration is not needed, but want to create a new shrinkage
    pOptionalArg = False ' this makes send to clone the settings only when the above the flag is true
    
    ' Initialization.
    Dim indShr              As Long
    Dim lMaxIndex           As Long
    Dim oPartSupport        As IJPartSupport
    Dim oShrinkage          As IJMfgChild

    lMaxIndex = 1 'index of the biggest plate part in pReplacingObjColl

    If TypeOf pReplacingPart Is IJStructProfilePart Then
    
        ' Initialization.
        Dim dMaxLength As Double
        dMaxLength = -1# 'maximum landing curve length

        'Create ProfilePart support
        Dim oProfilePartSupport As IJProfilePartSupport
        
        ' Loop for each element in the new Profile collection.
        For indShr = 1 To pMfgObjsColl.Count
            Set oShrinkage = pMfgObjsColl.Item(indShr)
            
            Dim dLength As Double
            
            Set oProfilePartSupport = New ProfilePartSupport
            Set oPartSupport = oProfilePartSupport
            'Get the base surface of this plate part
            Set oPartSupport.Part = oShrinkage.GetParent

            dLength = oProfilePartSupport.ApproximateLength

            ' See if this length is larger than the previous one
            If dLength > dMaxLength Then
                 ' If so, store this length as maximum and the element index
                dMaxLength = dLength
                lMaxIndex = indShr
            End If

            dLength = -1#
        Next indShr

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
        For indShr = 1 To pMfgObjsColl.Count
        
            Set oShrinkage = pMfgObjsColl.Item(indShr)
            
            'Get the base surface of this plate part
            Set oPartSupport.Part = oShrinkage.GetParent
    
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
                 lMaxIndex = indShr
            End If
            
            Set oBaseSurfaceModelBody = Nothing
            Set oBaseSurface = Nothing
            
        Next indShr
        
        Set oPlatePartSupport = Nothing
        Set oPartSupport = Nothing
    End If

    'Return the manufacturing part with largest surface area.
    Set IJMfgSplitMigration_ReverseMigrate = pMfgObjsColl.Item(lMaxIndex)
   
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7003, , "RULES")
End Function
