VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Filters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function IsSidePlate(oObject As Object) As Integer
    Let IsSidePlate = 0
    Dim sType As String
    Let sType = GetMacroType(oObject)
    If Mid(sType, 1, 9) = "SidePlate" _
    Or Mid(sType, 1, 9) = "RingPlate" Then
        Let IsSidePlate = 1
    End If
End Function
Public Function IsTransitionOrContinuityPlate(oObject As Object) As Integer
    Let IsTransitionOrContinuityPlate = 0
    Dim sType As String
    Let sType = GetMacroType(oObject)
    If sType = "TransitionPlate" _
    Or sType = "ContinuityPlate" Then
        Let IsTransitionOrContinuityPlate = 1
    End If
End Function
Public Function IsContinuityPlate(oObject As Object) As Integer
    Let IsContinuityPlate = 0
    Dim sType As String
    Let sType = GetMacroType(oObject)
    If sType = "ContinuityPlate" Then
        Let IsContinuityPlate = 1
        
    End If
End Function
Public Function IsTransitionPlate(oObject As Object) As Integer
    Let IsTransitionPlate = 0
    Dim sType As String
    Let sType = GetMacroType(oObject)
    If sType = "TransitionPlate" Then
        Let IsTransitionPlate = 1
    End If
End Function
Public Function IsRootPlateSystem(oObject As Object) As Integer
    Let IsRootPlateSystem = 0
    If TypeOf oObject Is IJPlateSystem Then
        If TypeOf oObject Is IJDesignChild Then
            Dim pDesignChild As IJDesignChild: Set pDesignChild = oObject
            If Not TypeOf pDesignChild.GetParent Is IJPlateSystem Then
                Let IsRootPlateSystem = 1
            End If
        End If
    End If
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Filter out any Plate System that is not an APS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsAdvancedPlateSystem(oObject As Object) As Integer
    Let IsAdvancedPlateSystem = 0
    If TypeOf oObject Is IJPlateSystem Then
        If TypeOf oObject Is IJDesignChild Then
            Dim pDesignChild As IJDesignChild: Set pDesignChild = oObject
            If Not TypeOf pDesignChild.GetParent Is IJPlateSystem Then
                If GetMacroType(oObject) <> "" Then
                    Let IsAdvancedPlateSystem = 1
                End If
            End If
        End If
    End If
End Function
Public Function IsRootStiffenerSystem(oObject As Object) As Integer
    Let IsRootStiffenerSystem = 0
    If TypeOf oObject Is IJStiffenerSystem Then
        If TypeOf oObject Is IJDesignChild Then
            Dim pDesignChild As IJDesignChild: Set pDesignChild = oObject
            If Not TypeOf pDesignChild.GetParent Is IJStiffenerSystem Then
                Let IsRootStiffenerSystem = 1
            End If
        End If
    End If
End Function
Public Function IsLinearStiffenerSystem(oObject As Object) As Integer
    Let IsLinearStiffenerSystem = 0
    If TypeOf oObject Is IJStiffenerSystem Then
'        Dim pDesignChild As IJDesignChild: Set pDesignChild = oObject
'        Dim oLeafProfileSystem As Object: Set oLeafProfileSystem = pDesignChild.GetParent()
        Dim pProfileAttributes As IJProfileAttributes: Set pProfileAttributes = New ProfileUtils
        Dim oLandingCurve As Object: Call pProfileAttributes.GetLandingCurveFromProfile(oObject, oLandingCurve)
        If TypeOf oLandingCurve Is IJLine Then
            Let IsLinearStiffenerSystem = 1
        End If
    End If
End Function
Private Function GetMacroType(oObject As Object) As String
    Let GetMacroType = ""
    If TypeOf oObject Is IJPlateSystem Then
        Dim pGeometricConstruction As IJGeometricConstruction
        Set pGeometricConstruction = PlateSystem_GetMacro(oObject)
        
        If Not pGeometricConstruction Is Nothing Then
            Let GetMacroType = pGeometricConstruction.Type
        End If
    End If
End Function

