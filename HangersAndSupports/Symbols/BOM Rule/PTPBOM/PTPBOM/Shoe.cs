//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Shoe.cs
//   PTPBOM,Ingr.SP3D.Content.Support.Symbols.Shoe
//   Author       :  Hema
//   Creation Date:  07-10-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07-10-2013   Hema       CR-CP-240907  Convert HS_PTPBOM to C# .Net    
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
    public class Shoe : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                //Get the values from the Part / Part Occurence
                string partDescription = part.PartDescription;

                int finishValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPFinish", "Finish")).PropValue;
                string finish = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPFinish", "Finish")).PropertyInfo.CodeListInfo.GetCodelistItem(finishValue).DisplayName;

                int slideMaterialValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPSlideMaterial", "SlidePlateMaterial")).PropValue;
                string slideMaterial = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPSlideMaterial", "SlidePlateMaterial")).PropertyInfo.CodeListInfo.GetCodelistItem(slideMaterialValue).DisplayName;

                if (slideMaterialValue == -1 && finishValue == -1) //For both Finish and Slide Plate Material is nothing
                    bomDescription = partDescription;
                else if (slideMaterialValue == -1)     //For Slide Plate Material is nothing
                    bomDescription = partDescription + ", " + finish;
                else if (finishValue == -1)
                    bomDescription = partDescription + ", " + slideMaterial; //For Finish is nothing
                else
                    bomDescription = partDescription + ", " + slideMaterial + ", " + finish;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PTPBOMLocalizer.GetString(PTPBOMResourceIDs.ErrBOMDescription, "Error in BOMDescription of BOM.cs"));
                return "";
            }
            return bomDescription;
        }
    }
}
#endregion
