//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMEndThreadedRod.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.AnvilBOMEndThreadedRod
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
    public class AnvilBOMEndThreadedRod : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomString = "", finish;
            double length, d = 0, d_User = 0, rodDiameter;
            PropertyValueCodelist finishCodelist;
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG253"))
                    d_User = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG253", "D_USER")).PropValue;
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_fig140"))
                    d_User = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_fig140", "D_USER")).PropValue;

                length = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                finishCodelist = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrRodFinish", "FINISH"));
                rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue;

                if (d_User > 0.0 && length < d_User * 2.0)
                    d = length / 2.0;
                else
                {
                    if (d_User > 0)
                        d = d_User;
                    else
                    {
                        if ((rodDiameter <= 0.015875) && (length <= 0.6096))
                            d = length / 2.0;
                    }
                }
                if (finishCodelist.PropValue != 1 && finishCodelist.PropValue != 2)
                {
                    ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFinishvalue, "Finish value should be 1 or 2"));
                    return "";
                }
                finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                bomString = part.PartDescription + "Length  " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_INCH) + "," + "D " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, d, UnitName.DISTANCE_INCH) + "," + "Finish " + finish;
                return bomString;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMEndThreadedRod"));
                return "";
            }
        }
    }
}
#endregion
