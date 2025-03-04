//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMSpringsVariable.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.AnvilBOMSpringsVariable
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
    public class AnvilBOMSpringsVariable : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomString = "";
            double hotLoad, workingTravel = 0;
            PropertyValueCodelist travelDir = null;
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                hotLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrHotLoad", "HOT_LOAD")).PropValue;

                if (part.SupportsInterface("IJOAHgrAnvil_FIGB268E"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268E", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB268E", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIGB268D"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268D", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB268D", "WORKING_TRAV")).PropValue;
                }
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIGB268C"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268C", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB268C", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIGB268B"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268B", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB268B", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIGB268A"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268A", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB268A", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG98E"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98E", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG98E", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG98D"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98D", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG98D", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG98C"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98C", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG98C", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG98B"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98B", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG98B", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG98A"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98A", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG98A", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG82E"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG82E", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG82E", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG82D"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG82D", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG82D", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG82C"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG82C", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG82C", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG82B"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG82B", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG82B", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG82A"))
                {
                    travelDir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG82A", "DIR"));
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG82A", "WORKING_TRAV")).PropValue;
                }
                if (travelDir != null)
                {
                    if (travelDir.PropValue != 1 && travelDir.PropValue != 2)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTravelDirValue, "Travel direction value should be 1 or 2"));
                        return "";
                    }
                    string travelDirection = travelDir.PropertyInfo.CodeListInfo.GetCodelistItem(travelDir.PropValue).DisplayName;
                    if (travelDirection.Length > 0)
                        travelDirection = " " + travelDirection;

                    bomString = part.PartDescription + ",Hot Load:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, hotLoad, UnitName.FORCE_POUND_FORCE) + ", Travel: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, workingTravel, UnitName.DISTANCE_INCH) + travelDirection;
                }
               
                return bomString;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMSpringsVariable"));
                return "";
            }
        }
    }
}
#endregion
