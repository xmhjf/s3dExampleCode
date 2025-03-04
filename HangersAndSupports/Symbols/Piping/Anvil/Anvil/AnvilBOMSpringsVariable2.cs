//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMSpringsVariable2.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.AnvilBOMSpringsVariable2
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
    public class AnvilBOMSpringsVariable2 : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomString = "", ndUnitType;
            double hotLoad, pipeDia, workingTravel = 0;
            PropertyValueCodelist dir = null, coltype = null, topCodelist = null, roleMaterialCodeList = null;
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                if (part.SupportsInterface("IJOAHgrAnvil_FIGB268F"))
                {
                    dir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268F", "DIR"));
                    coltype = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268F", "COL_TYP"));
                    roleMaterialCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268F", "ROLL_MATERIAL");
                    topCodelist = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIGB268F", "TOP");
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIGB268F", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG98F"))
                {
                    dir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98F", "DIR"));
                    coltype = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98F", "COL_TYP"));
                    roleMaterialCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98F", "ROLL_MATERIAL");
                    topCodelist = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98F", "TOP");
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG98F", "WORKING_TRAV")).PropValue;
                }
                else if (part.SupportsInterface("IJOAHgrAnvil_FIG82F"))
                {
                    dir = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG82F", "DIR"));
                    coltype = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG82F", "COL_TYP"));
                    roleMaterialCodeList = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG82F", "ROLL_MATERIAL"));
                    topCodelist = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG82F", "TOP");
                    workingTravel = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrAnvil_FIG82F", "WORKING_TRAV")).PropValue;
                }

                hotLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrHotLoad", "HOT_LOAD")).PropValue;
                ndUnitType = (string)((PropertyValueString)part.GetPropertyValue("IJHgrDiameterSelection", "NDUnitType")).PropValue;
                pipeDia = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;

                if (dir != null && coltype != null && topCodelist != null && roleMaterialCodeList != null)
                {
                    if (dir.PropValue <= 0 || dir.PropValue > 15)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDirValue, "Direction value between 1 and 15"));
                        return "";
                    }

                    if (coltype.PropValue != 1 && coltype.PropValue != 2)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidcolumntypeValue, "COL_TYPE value should be 1 or 2"));
                        return "";
                    }

                    if (topCodelist.PropValue != 1 && topCodelist.PropValue != 2 && topCodelist.PropValue != 3)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidcolumntypeValue, "TOP value should be 1 or 2 or 3"));
                        return "";
                    }

                    if (roleMaterialCodeList.PropValue != 1 && roleMaterialCodeList.PropValue != 2)
                    {
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRollMaterialValue, "Roll material value should be 1 or 2 or 3 "));
                        return "";
                    }

                    string top = topCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(topCodelist.PropValue).DisplayName;
                    string coltype1 = coltype.PropertyInfo.CodeListInfo.GetCodelistItem(coltype.PropValue).DisplayName;
                    string dir1 = dir.PropertyInfo.CodeListInfo.GetCodelistItem(dir.PropValue).DisplayName;
                    string rollmaterial = roleMaterialCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(roleMaterialCodeList.PropValue).DisplayName;
                    string bomEnd;
                    if (coltype1 == "Guided")
                        bomEnd = ", Guided Load Column.  Hot Load: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, hotLoad, UnitName.FORCE_POUND_FORCE) + "  Travel: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, workingTravel, UnitName.FORCE_POUND_FORCE) + " " + dir1;
                    else
                        bomEnd = ", Hot Load: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, hotLoad, UnitName.FORCE_POUND_FORCE) + " Travel: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, workingTravel, UnitName.FORCE_POUND_FORCE) + " " + dir1;

                    if (top == "LOAD FlANGE")
                        bomString = part.PartDescription + ", LoadFlange" + bomEnd;
                    if (top == "Pipe Roll")
                        bomString = part.PartDescription + pipeDia + " " + ndUnitType + "  " + rollmaterial + " piperoll" + bomEnd;
                    if (top == "None")
                        bomString = part.PartDescription + bomEnd;
                }
                
                return bomString;
            }
            catch  //General Unhandled exception    
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMSpringsVariable2"));
                return "";
            }
        }
    }
}
#endregion
