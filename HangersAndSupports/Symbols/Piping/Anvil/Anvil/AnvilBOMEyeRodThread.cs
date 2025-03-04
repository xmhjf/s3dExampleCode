//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMEyeRodThread.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.AnvilBOMEyeRodThread
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
    public class AnvilBOMEyeRodThread : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomString = "", finish;
            double length, d = 0, d_User = 0, rodDiameter, holeClearance;
            PropertyValueCodelist finishCodelist;
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                length = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue;
                finishCodelist = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrRodFinish", "FINISH"));

                if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG278L"))
                {
                    d_User = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG278L", "D_USER")).PropValue;
                    d = (double)((PropertyValueDouble)part.GetPropertyValue("IJAHgrAnvil_FIG278L", "D")).PropValue;
                }
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG278"))
                {
                    d_User = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG278", "D_USER")).PropValue;
                    d = (double)((PropertyValueDouble)part.GetPropertyValue("IJAHgrAnvil_FIG278", "D")).PropValue;
                }
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG248L"))
                {
                    d_User = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG248L", "D_USER")).PropValue;
                    d = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIG248L", "D")).PropValue;
                }
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG248"))
                {
                    d_User = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG248", "D_USER")).PropValue;
                    d = (double)((PropertyValueDouble)part.GetPropertyValue("IJAHgrAnvil_FIG248", "D")).PropValue;
                }
                if (finishCodelist.PropValue != 1 && finishCodelist.PropValue != 2)
                {
                    ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFinishvalue, "Finish value should be 1 or 2"));
                    return "";
                }
                finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                holeClearance = 0.003175;
                if (rodDiameter > 0.036)
                    holeClearance = 0.00635;
                if (d_User != d)
                    d = d_User;
                bomString = part.PartDescription + "  " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length - rodDiameter / 2.0 - holeClearance, UnitName.DISTANCE_INCH) + " " + "long, center to end, D=" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, d, UnitName.DISTANCE_INCH) + ", " + "Finish " + finish;
                return bomString;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMEyeRodThread"));
                return "";
            }
        }
    }
}
#endregion
