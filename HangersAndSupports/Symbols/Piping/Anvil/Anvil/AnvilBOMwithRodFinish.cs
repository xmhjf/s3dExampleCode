//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AnvilBOMwithRodFinish.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.AnvilBOMwithRodFinish
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
    public class AnvilBOMwithRodFinish : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            double length;
            string bomString = "", finish;
            try
            {
                PropertyValueCodelist finishCodelist = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAHgrRodFinish", "FINISH"));
                length = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                bomString = part.PartDescription + ", " + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_INCH) + ", Finish: " + finish;
                return bomString;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of AnvilBOMwithRodFinish"));
                return "";
            }
        }
    }
}
#endregion
