//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMConstantSupport2.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.AnvilBOMConstantSupport2
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
    public class AnvilBOMConstantSupport2 : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "", suspension, travelDirection;
            double load = 0, totalTravel = 0, actTravel = 0;
            PropertyValueCodelist travelDirCodelist = null, suspensionCodelist = null;
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                if (oSupportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG81A"))
                {
                    load = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81A", "LOAD")).PropValue;
                    totalTravel = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81A", "TOTAL_TRAV")).PropValue;
                    actTravel = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81A", "ACT_TRAVEL")).PropValue;
                    travelDirCodelist = ((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81A", "TRAVEL_DIR"));
                    suspensionCodelist = ((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81A", "SUSPENSION"));
                }
                if (oSupportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG81B"))
                {
                    load = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81B", "LOAD")).PropValue;
                    totalTravel = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81B", "TOTAL_TRAV")).PropValue;
                    actTravel = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81B", "ACT_TRAVEL")).PropValue;
                    travelDirCodelist = ((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81B", "TRAVEL_DIR"));
                    suspensionCodelist = ((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81B", "SUSPENSION"));
                }
                if (oSupportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG81C"))
                {
                    load = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81C", "LOAD")).PropValue;
                    totalTravel = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81C", "TOTAL_TRAV")).PropValue;
                    actTravel = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81C", "ACT_TRAVEL")).PropValue;
                    travelDirCodelist = ((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81C", "TRAVEL_DIR"));
                    suspensionCodelist = ((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG81C", "SUSPENSION"));
                }

                if (travelDirCodelist != null && suspensionCodelist != null)
                {
                    if (travelDirCodelist.PropValue != 1 && travelDirCodelist.PropValue != 2)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTravelDirValue, "Travel direction value should be 1 or 2"));
                        return "";
                    }
                    if (suspensionCodelist.PropValue != 1 && suspensionCodelist.PropValue != 2)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidSuspensionValue, "Suspension value should be 1 or 2"));
                        return "";
                    }

                    travelDirection = travelDirCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(travelDirCodelist.PropValue).DisplayName;
                    suspension = suspensionCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(suspensionCodelist.PropValue).DisplayName;
                    double travel = 0.5 * ((int)((totalTravel / 0.0254 + 0.499) * 2.0)) * 25.4 / 1000;

                    bomString = part.PartDescription + "LOAD " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, load, UnitName.FORCE_POUND_FORCE) + "," + "TotalTravel " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, travel, UnitName.DISTANCE_INCH_SYMBOL) + "," + "Actual Travel " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, actTravel, UnitName.DISTANCE_INCH_SYMBOL) + ", " + "Direction of Travel " + travelDirection + ", " + "Suspension" + suspension;
                    bomString = part.PartDescription;
                }

                return bomString;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMConstantSupport2"));
                return "";
            }
        }
    }
}
#endregion
