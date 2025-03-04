VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PinJigMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJMfgSplitMigration
Private Const MODULE = "PinJigMigration"

Private Function IJMfgSplitMigration_MigrateObject(ByVal pOldObj As Object, ByVal pReplacingObjColl As GSCADStructMfgGlobals.IJElements, ByVal pMfgObject As Object, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
    Const METHOD = "IJMfgSplitMigration_MigrateObject"
    On Error GoTo ErrorHandler

    Dim lNewPlateCount As Long
    Dim indPlate As Long
    Dim dMaxSurfaceArea As Double
    Dim lMaxSurfaceIndex As Long

    'Don't Create PinJigs automatically for the new plates.
    pbCreateMfg = False

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

    ' no settings to clone
    pOptionalArg = False
        
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5002, , "RULES")
End Function
Private Function IJMfgSplitMigration_ReverseMigrate(ByVal pReplacedPartColl As GSCADStructMfgGlobals.IJElements, ByVal pReplacingPart As Object, ByVal pMfgObjsColl As GSCADStructMfgGlobals.IJElements, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
    Const METHOD = "IJMfgSplitMigration_ReverseMigrate"
    On Error GoTo ErrorHandler

    'Return the pinjig on the plate part with largest surface area.
        
    pbCreateMfg = False ' set this to true only if migration is not needed but want to create a new pin jig
    pOptionalArg = False ' this makes send to clone the settings only when the above the flag is true
    
    ' Initialization.
    Dim indPinJig           As Long
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
    
    ' Loop for each element in the manufactruing objects collection.
    For indPinJig = 1 To pMfgObjsColl.Count
        
        Dim oPinJig As IJMfgChild
        Set oPinJig = pMfgObjsColl.Item(indPinJig)
        
        Set oMfgUpdateInfo = oPinJig
        oMfgUpdateInfo.UpToDate = FLAG_PART_DELETED
        
        'Get the base surface of this plate part
        Set oPartSupport.Part = oPinJig.GetParent

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
             lMaxSurfaceIndex = indPinJig
        End If
        
        Set oBaseSurfaceModelBody = Nothing
        Set oBaseSurface = Nothing
        
    Next indPinJig
    
    Set oPlatePartSupport = Nothing
    Set oPartSupport = Nothing
    
    Set oMfgUpdateInfo = pMfgObjsColl.Item(lMaxSurfaceIndex)
    oMfgUpdateInfo.UpToDate = FLAG_MIGRATED
    
    'Return the manufacturing part with largest surface area.
    Set IJMfgSplitMigration_ReverseMigrate = pMfgObjsColl.Item(lMaxSurfaceIndex)
   
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5002, , "RULES")

End Function
