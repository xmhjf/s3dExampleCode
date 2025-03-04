VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfgPlateMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''
''   Copyright (C) 2002, Intergraph Corporation. All rights reserved.
''
''   Abstract:
''              MfgPlateMigration.cls
''   ProgID:
''              StrMfgMigrationRules.MfgPlateMigration
''   Author:
''              Ingr StructMfg team.
''   Creation Date:
''              19 April 2006
''   Description:
''      CR#84861: Struct Mfg should check split notification and keep the GNEST lot number.
''      When a part has already been nested, it has a Lot number to link the part between
''      IntelliShip and GNEST.
''
''      In a simplest test, a plate "P1" is split into two smaller plates "P2" and "P3" and "P1" will be deleted.
''      Old plate P1, along with new Plates P2 and P3 will be passed to a user defined VB rule.  This rule will be
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
Private Const MODULE = "MfgPlateMigration"
Implements IJMfgSplitMigration

Private Function IJMfgSplitMigration_MigrateObject(ByVal pOldObj As Object, ByVal pReplacingObjColl As GSCADStructMfgGlobals.IJElements, ByVal pMfgObject As Object, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
'Arguments:
' pOldObj ==>  Plate "P1" which is being deleted.
'  pReplacingObjColl ==> Collection of replacing new plates ( P2 and P3 ).
' pbCreateMfg ==>   Boolean flag indicating if software should automatically create MfgPlates for the new plates.
' Output of this routine will be one of the plate in pReplacingObjColl collection. "M1" in the above example will then be reattached to that plate and thus maintain the lot# for this plate part.
' NOTE:  The rule is not limited to two plates.  "n" number of plates could be passed to the rule after a split.  The rule would still only return one plate meeting the criteria.
'
' pbCreateMfg   Output of MigrateObject()   Action
' ===========   =========================   ======
' False         Nothing           MfgPlate ( M1 ) will be deleted. NO new MfgPlates will be created.
' True          Nothing           MfgPlate ( M1 ) will be deleted. New MfgPlates will be created for all the new plates ( for P2 and P3 ).
' False         Object P2         MfgPlate ( M1 ) will be re-connected from P1 to P2. New MfgPlates will NOT be created for the remaining new plates ( for P3 ).
' True          Object P2         MfgPlate ( M1 ) will be re-connected from P1 to P2.  New MfgPlates will be created for the remaining new plates ( for P3 ).
'

    Const METHOD = "IJMfgSplitMigration_MigrateObject"
    On Error GoTo ErrorHandler

    Dim lNewPlateCount      As Long
    Dim indPlate            As Long
    Dim dMaxSurfaceArea     As Double
    Dim lMaxSurfaceIndex    As Long

    'Create MfgPlates automatically for the new plates.
    pbCreateMfg = True

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

        Dim oPartSupport As IJPartSupport
        Set oPartSupport = oPlatePartSupport

        'Get the base surface of this plate part
        Set oPartSupport.Part = pReplacingObjColl.Item(indPlate)

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
    
    ' clone the process and marking settings
    pOptionalArg = True

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1035, , "RULES")
End Function

Private Function IJMfgSplitMigration_ReverseMigrate(ByVal pReplacedPartColl As GSCADStructMfgGlobals.IJElements, ByVal pReplacingPart As Object, ByVal pMfgObjsColl As GSCADStructMfgGlobals.IJElements, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
    Const METHOD = "IJMfgSplitMigration_ReverseMigrate"
    On Error GoTo ErrorHandler

    'Return the plate part with largest surface area.
        
    pbCreateMfg = False ' set this to true only if migration is not needed but want to create a new mfg plate
    pOptionalArg = False ' this makes send to clone the settings only when the above the flag is true
    
    ' Initialization.
    Dim lNewPlateCount      As Long
    Dim indPlate            As Long
    Dim dMaxSurfaceArea     As Double
    Dim lMaxSurfaceIndex    As Long
    
    dMaxSurfaceArea = -1# 'maximum surface area
    lMaxSurfaceIndex = 0 'index of the biggest plate part in pReplacingObjColl

    'Create PlatePart support to get the base surface of the plate without any features/camfers
    Dim oPlatePartSupport As IJPlatePartSupport
    Set oPlatePartSupport = New PlatePartSupport

    Dim oPartSupport As IJPartSupport
    Set oPartSupport = oPlatePartSupport
    
    Dim oMfgUpdateInfo As IJMfgUpdateInfo
    
    ' Loop for each element in the new plate collection.
    For indPlate = 1 To pMfgObjsColl.Count
        
        Dim oMfgPlate As IJMfgChild
        Set oMfgPlate = pMfgObjsColl.Item(indPlate)
        
        Set oMfgUpdateInfo = oMfgPlate
        oMfgUpdateInfo.UpToDate = FLAG_PART_DELETED
                
        'Get the base surface of this plate part
        Set oPartSupport.Part = oMfgPlate.GetParent

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
             lMaxSurfaceIndex = indPlate
        End If
        
        Set oBaseSurfaceModelBody = Nothing

        Set oBaseSurface = Nothing
        
    Next indPlate
    
    Set oPlatePartSupport = Nothing
    Set oPartSupport = Nothing
    
    Set oMfgUpdateInfo = pMfgObjsColl.Item(lMaxSurfaceIndex)
    oMfgUpdateInfo.UpToDate = FLAG_MIGRATED
    
    'Return the manufacturing part with largest surface area.
    Set IJMfgSplitMigration_ReverseMigrate = pMfgObjsColl.Item(lMaxSurfaceIndex)
   
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1035, , "RULES")
End Function

