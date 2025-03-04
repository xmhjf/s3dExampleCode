//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMPipeStachion.cs
//     Author       :  Manikanth
//   Creation Date:  16-05-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-05-2013   Manikanth  CR-CP-222292 & CR-CP-233113 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

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
    public class AnvilBOMPipeStachion : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomString = "";
            double stanchDiameter;
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double length = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                PropertyValueCodelist stanchionNomDiaCodeList = null;

                if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG63A"))
                    stanchionNomDiaCodeList = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG63A", "STANCHION_NOM_DIA"));
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG63B"))
                    stanchionNomDiaCodeList = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG63B", "STANCHION_NOM_DIA"));
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG63C"))
                    stanchionNomDiaCodeList = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG63C", "STANCHION_NOM_DIA"));

                if (stanchionNomDiaCodeList != null)
                {
                    if (stanchionNomDiaCodeList.PropValue <= 0 || stanchionNomDiaCodeList.PropValue > 15)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidStanchDiavalue, "Stanchion diameter value should be between 1 and 15"));
                        return "";
                    }
                    stanchDiameter = Convert.ToDouble(stanchionNomDiaCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(stanchionNomDiaCodeList.PropValue).DisplayName);
                    bomString = part.PartDescription + ", " + "Stanchion Pipe Size:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, stanchDiameter, UnitName.DISTANCE_INCH) + "D=" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_INCH);
                }
               
                return bomString;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMPipeStachion"));
                return "";
            }
        }
    }
}
#endregion
