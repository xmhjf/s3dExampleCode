//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMSpringsVariable3.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.AnvilBOMSpringsVariable3
//   Author       :  ManiKanth
//   Creation Date:  13-05-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-05-2013   Manikanth  CR-CP-222292 & CR-CP-233113 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
    public class AnvilBOMSpringsVariable3 : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            PropertyValueCodelist dir = null;
            double p = 0, span = 0, workingTravel = 0, hotLoad = 0;
            string bomString = "";
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                if ((part.SupportsInterface("IJOAHgrAnvil_FIGB268G") && (part.SupportsInterface("IJUAHgrAnvil_FIGB268G"))))
                {
                    dir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268G", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB268G", "WORKING_TRAV")).PropValue;
                    p = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIGB268G", "P")).PropValue;
                    span = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB268G", "SPAN")).PropValue;
                }
                else if ((part.SupportsInterface("IJOAHgrAnvil_FIGB98G") && (part.SupportsInterface("IJUAHgrAnvil_FIGB98G"))))
                {
                    dir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB98G", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB98G", "WORKING_TRAV")).PropValue;
                    p = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIGB98G", "P")).PropValue;
                    span = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB98G", "SPAN")).PropValue;
                }
                else if ((part.SupportsInterface("IJOAHgrAnvil_FIGB82G") && (part.SupportsInterface("IJUAHgrAnvil_FIGB82G"))))
                {
                    dir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB82G", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB82G", "WORKING_TRAV")).PropValue;
                    p = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIGB82G", "P")).PropValue;
                    span = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB82G", "SPAN")).PropValue;
                }
                if (dir !=null)
                {
                    if (dir.PropValue <= 0 || dir.PropValue > 15)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDirValue, "Direction value should be between 1 and 15"));
                        return "";
                    }
                    string direction = dir.PropertyInfo.CodeListInfo.GetCodelistItem(dir.PropValue).DisplayName;
                    hotLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrHotLoad", "HOT_LOAD")).PropValue;
                    bomString = part.PartDescription + ", C-C" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, span, UnitName.DISTANCE_INCH) + " ,P =" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, p, UnitName.DISTANCE_INCH) + " Hot Load Per Spring: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, hotLoad, UnitName.FORCE_POUND_FORCE) + " Total Load: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, hotLoad * 2, UnitName.FORCE_POUND_FORCE) + " Travel: " + direction;
                }

                return bomString;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMSpringsVariable3"));
                return "";
            }
        }
    }
}
#endregion
