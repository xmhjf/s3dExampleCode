'***************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project  : M:\SharedContent\Src\CommonRoute\Rules\RouteStockPartNameRules\
'
'  Class    : CondPartNameRule
'
'  Abstract : The file contains implementation of the naming rules for ConduitParts.
'
'***************************************************************************************
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports SystemsAndSpecsMiddle = Ingr.SP3D.Systems.Middle
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle


Public Class CondPartNameRule
    Inherits NameRuleBase
    Private Const strCountFormat = "0000"       ' define fixed-width number field for name
    Public Overrides Sub ComputeName(ByVal oEntity As Common.Middle.BusinessObject, ByVal oParents As System.Collections.ObjectModel.ReadOnlyCollection(Of Common.Middle.BusinessObject), ByVal oActiveEntity As Common.Middle.BusinessObject)
        Dim oNamedItem As NamedItemHelper, oParent As NamedItemHelper
        Dim strName As String, strRunName As String, strSequenceNo As String, strLocation As String, strConduit As String
        Dim lcount As Long
        strConduit = "COND"
        Dim oLogError As LogError
        oLogError = MiddleServiceProvider.ErrorLogger()
        Dim sMethod = "CondPartNameRule : ComputeName() "
        If oEntity Is Nothing Or oParents Is Nothing Then
            Exit Sub
        End If
        Try
            'To get the SystemParentName
            oParent = New NamedItemHelper(oParents.Item(0))
            strRunName = oParent.Name

            If strRunName <> GetNamingParentsString(oActiveEntity) Then
                'Adding the RunName as the Naming ParentString
                SetNamingParentsString(oActiveEntity, strRunName)
                strLocation = vbNullString
                GetCountAndLocationID(strRunName, lcount, strLocation)
                strSequenceNo = Format(lcount, strCountFormat)
                'Name is Formed in the Required Sequence
                strName = strRunName + "-" + strConduit + "-" + strSequenceNo
                'Setting Name on the Object
                oNamedItem = New NamedItemHelper(oEntity)
                SetName(oEntity, strName)
            End If

        Catch ex As Exception
            Dim sException As String = sMethod + ex.Message
            If (Not oLogError Is Nothing) Then
                oLogError.Log(sException)
            End If
            Throw ex
        End Try
    End Sub
    Public Overrides Function GetNamingParents(ByVal oEntity As Common.Middle.BusinessObject) As System.Collections.ObjectModel.Collection(Of Common.Middle.BusinessObject)
        Dim oChild As SystemChildHelper
        Dim oParent As BusinessObject
        Dim oParentColl As New Collection(Of BusinessObject)
        GetNamingParents = Nothing
        If oEntity Is Nothing Then
            Exit Function
        End If
        Try
            oChild = New SystemChildHelper(oEntity)
            oParent = oChild.SystemParent
            If Not oParent Is Nothing Then
                oParentColl.Add(oParent)
            End If
            GetNamingParents = oParentColl
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
