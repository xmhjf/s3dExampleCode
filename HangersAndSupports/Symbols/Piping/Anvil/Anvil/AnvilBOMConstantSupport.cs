//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMConstantSupport.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.AnvilBOMConstantSupport
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
    public class AnvilBOMConstantSupport : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomString = "";
            double load = 0, totalTravel = 0, actTravel = 0;
            PropertyValueCodelist travelDir = null;
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG80A"))
                {
                    load = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80A", "LOAD")).PropValue;
                    totalTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80A", "TOTAL_TRAV")).PropValue;
                    actTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80A", "ACT_TRAVEL")).PropValue;
                    travelDir = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80A", "TRAVEL_DIR"));
                }
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG80B"))
                {
                    load = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80B", "LOAD")).PropValue;
                    totalTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80B", "TOTAL_TRAV")).PropValue;
                    actTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80B", "ACT_TRAVEL")).PropValue;
                    travelDir = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80B", "TRAVEL_DIR"));
                }
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG81D"))
                {
                    load = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81D", "LOAD")).PropValue;
                    totalTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81D", "TOTAL_TRAV")).PropValue;
                    actTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81D", "ACT_TRAVEL")).PropValue;
                    travelDir = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81D", "TRAVEL_DIR"));
                }
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG80C"))
                {
                    load = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80C", "LOAD")).PropValue;
                    totalTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80C", "TOTAL_TRAV")).PropValue;
                    actTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80C", "ACT_TRAVEL")).PropValue;
                    travelDir = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG80C", "TRAVEL_DIR"));
                }
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG81F"))
                {
                    load = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81F", "OPER_LOAD")).PropValue;
                    totalTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81F", "TOTAL_TRAV")).PropValue;
                    actTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81F", "ACT_TRAVEL")).PropValue;
                    travelDir = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81F", "DIR"));
                }

                if (travelDir != null)
                {
                    if (travelDir.PropValue != 1 && travelDir.PropValue != 2)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTravelDirValue, "Travel direction value should be 1 or 2"));
                        return "";
                    }
                    string direction = travelDir.PropertyInfo.CodeListInfo.GetCodelistItem(travelDir.PropValue).DisplayName;
                    double travel = 0.5 * ((int)((totalTravel / 0.0254 + 0.499) * 2.0)) * 25.4 / 1000;

                    bomString = part.PartDescription + "LOAD " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, load, UnitName.FORCE_POUND_FORCE) + "," + "TotalTravel " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, travel, UnitName.DISTANCE_INCH_SYMBOL) + "," + "Actual Travel " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, actTravel, UnitName.DISTANCE_INCH_SYMBOL) + ", " + "Direction of Travel " + direction;
                }
                
                return bomString;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMConstantSupport"));
                return "";
            }
        }
    }
}
#endregion