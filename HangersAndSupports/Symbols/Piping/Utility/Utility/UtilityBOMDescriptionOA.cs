//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   UtilityBOMDescriptionOA.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.UtilityBOMDescriptionOA.cs
//   Author       :  Hema
//   Creation Date:  31/10/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   31/10/2012      Hema   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		 Hema 	DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;


#region "ICustomHgrBOMDescription Members"
namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [SymbolVersion("1.0.0.0")]
    public class UtilityBOMDescriptionOA : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_END_PLATE_HOLED"))
	            {
                    bomString = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE_HOLED", "BOM_DESC")).PropValue;
	            }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_CURVED_PLATE"))
                {
                    bomString = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_CURVED_PLATE", "BOM_DESC")).PropValue;
                }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_GENERIC_L"))
                {
                    bomString = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "BOM_DESC")).PropValue;
                }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_GENERIC_T"))
                {
                    bomString =(string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_T", "BOM_DESC")).PropValue;
                }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_GENERIC_W"))
                {
                    bomString =(string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "BOM_DESC")).PropValue;
                }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_USER_FIXED_BOX"))
                {
                    bomString =(string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC")).PropValue;
                }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_USER_FIXED_CYL"))
                {
                    bomString =(string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_USER_FIXED_CYL", "BOM_DESC")).PropValue;
                }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_VARIABLE_BOX"))
                {
                    bomString =(string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_VARIABLE_BOX", "BOM_DESC")).PropValue;
                }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_VARIABLE_CYL"))
                {
                    bomString =(string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_VARIABLE_CYL", "BOM_DESC")).PropValue;
                }
                return bomString;
            }
            catch
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of UtilityBOMDescriptionOA"));
                return "";
            }
        }
    }
}
#endregion
