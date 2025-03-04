//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ConstantSpring.cs
//    JIMCBOM,Ingr.SP3D.Content.Support.Symbols.ConstantSpring
//   Author       :  Hema
//   Creation Date:  07-10-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07-10-2013   Hema       CR-CP-240907  Convert HS_JIMCBOM to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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
    public class ConstantSpring : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                
                //Get the part description
                string partDescription = part.PartDescription;
                string[] splitPartDescription = partDescription.Split('-');
                string shortDescription = splitPartDescription[0];
                string travel = splitPartDescription[1];
                string sizeNumber = splitPartDescription[2];

                //Get the values from the Part / Part Occurence
                double totalTravel = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsTotalTravel", "TotalTravel")).PropValue;
                double maxTravel = totalTravel * 1000;
                bomDescription = shortDescription + "-" + maxTravel + "-" + sizeNumber;

                return bomDescription;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, JIMCBOMLocalizer.GetString(JIMCBOMResourceIDs.ErrBOMDescription, "Error in BOMDescription of AnvilBOMwithRodFinish"));
                return "";
            }
        }
    }
}
#endregion
