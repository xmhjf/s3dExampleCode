//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMwithFinish.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.AnvilBOMwithFinish
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
    public class AnvilBOMwithFinish : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomString = "", finish;
            PropertyValueCodelist finishCodelist=null;
            try
            {
                if (supportOrComponent.SupportsInterface("IJOAHgrRodFinish"))
                    finishCodelist = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrRodFinish", "FINISH"));
                else if (supportOrComponent.SupportsInterface("IJOAHgrFinish"))
                    finishCodelist = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrFinish", "FINISH"));
                else if (supportOrComponent.SupportsInterface("IJOAHgrAnvil_FIG86"))
                    finishCodelist = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG86", "FINISH"));

                if(finishCodelist!=null)
                {
                   if (finishCodelist.PropValue != 1 && finishCodelist.PropValue != 2)
                   {
                     ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFinishvalue, "Finish value should be 1 or 2"));
                     return "";
                   }
                   finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                   Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                   bomString = part.PartDescription + ", Finish: " + finish;
                }
                
                return bomString;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMwithFinish"));
                return "";
            }
        }
    }
}
#endregion
