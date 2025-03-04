//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ConstantSpring.cs
//   PSLBOM,Ingr.SP3D.Content.Support.Symbols.RiserClamp
//   Author       :  PVK
//   Creation Date:  17-06-2014
//   Description:                         

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-06-2014      PVK     DM-CP-250540 Deliver new Smart Parts for PSL
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
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
    public class RiserClamp : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                
                
                //Get the part description
                string partDescription = part.PartDescription;
                string partnumber = part.PartNumber;
                double fmin = 0, fmax = 0, f = 0, pindiameter = 0;


                //Get the values from the Part / Part Occurence
                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass auxTable = (PartClass)cataloghelper.GetPartClass("HS_PSL_RISERCLAMP_AUX");
                ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in classItems)
                {
                    //string spartnumber = ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrPSL_Clamp_AUX", "PART_NUMBER")).PropValue);
                    bool isEqual = String.Equals(partnumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrPSL_Clamp_AUX", "PART_NUMBER")).PropValue), StringComparison.Ordinal);
                    if (isEqual == true)
                    {
                        fmax = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrPSL_Clamp_AUX", "FMAX")).PropValue);
                        fmin = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrPSL_Clamp_AUX", "FMIN")).PropValue);
                        break;
                    }
                }
                    f = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAhsWingWidth", "F")).PropValue;
                    pindiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsPin1", "Pin1Diameter")).PropValue;

                    if (f <=0)
                    {
                        f = fmin;
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLBOMLocalizer.GetString(PSLBOMResourceIDs.ErrInvalidMinMaxLength, "F can not be negative. " + "Resetting F to F Minimum"));
                    }

                    if (f < fmin)
                    {
                        f = fmin;
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLBOMLocalizer.GetString(PSLBOMResourceIDs.ErrInvalidMinMaxLength, "F is Less than F Minimum of " + fmin * 1000 + "mm. " + "Resetting F to F Minimum"));
                    }
                    if (f > fmax)
                    {
                        f=fmax;
                        ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLBOMLocalizer.GetString(PSLBOMResourceIDs.ErrInvalidMinMaxLength, "F is Greater than F Maximum of " + fmax * 1000 + "mm. " + "Resetting F to F Maximum"));
                    }

                    supportOrComponent.SetPropertyValue(f, "IJOAhsWingWidth", "F");
                    supportOrComponent.SetPropertyValue(f / 2 + (1.5 * pindiameter), "IJOAhsWingHeight1", "Height1");
                    supportOrComponent.SetPropertyValue(f / 2 + (1.5 * pindiameter), "IJOAhsWingHeight2", "Height2");
                    supportOrComponent.SetPropertyValue(f / 2, "IJOAhsWingOffset1", "Offset1");
                    supportOrComponent.SetPropertyValue(f / 2, "IJOAhsWingOffset2", "Offset2");

                    bomDescription = partDescription + "," + "F= " + f*1000;


                
                return bomDescription;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLBOMLocalizer.GetString(PSLBOMResourceIDs.ErrBOMDescription, "Error in BOMDescription of RiserClamp"));
                return "";
            }
        }
    }
}
#endregion
