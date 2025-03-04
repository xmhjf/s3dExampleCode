'=====================================================================================================
'
'Copyright 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'Equipment foundation custom assemblies localizer
'
'File
'  EquipmentFoundationLocalizer.vb
'
'=====================================================================================================
Imports Ingr.SP3D.Common.Middle.Services

Namespace Ingr.SP3D.Content.Structure

    NotInheritable Class EquipmentFoundationLocalizer
        Public Shared Function GetString(ByVal iID As Integer, ByVal defMsgStr As String) As String
            Return CmnLocalizer.GetString(iID, defMsgStr, EquipmentFoundationResourceIDs.DEFAULT_RESOURCE, EquipmentFoundationResourceIDs.DEFAULT_ASSEMBLY)
        End Function
    End Class

End Namespace
