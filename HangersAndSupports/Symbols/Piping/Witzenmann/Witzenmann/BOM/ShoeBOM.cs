//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ZVxBOM.cs
//   Witzenmann,Ingr.SP3D.Content.Support.Symbols.ShoeBOM
//   Author       : Vinod  
//   Creation Date:  01.12.2015
//   DI-CP-282684  Integrate the newly developed Witzenmann Parts into Product  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.ObjectModel;

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
    public class ShoeBOM : SmartPartComponentDefinition, ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                string figureNumber = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUAhsFigureNumber", "FigureNumber")).PropValue;

                double nominalDiameter;
                if (catalogPart.SupportsInterface("IJUAhsNomPipeDia"))
                {
                    nominalDiameter = ((double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsNomPipeDia", "NomPipeDia")).PropValue);
                }
                else
                {
                    nominalDiameter = ((double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJHgrDiameterSelection", "NDFrom")).PropValue);
                    string NDUnitType = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJHgrDiameterSelection", "NDUnitType")).PropValue;
                    if (NDUnitType.ToUpper() == "MM")
                    {
                        nominalDiameter = nominalDiameter / 1000;
                    }
                }
          
                double nominalShoeheight = ((double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsNomShoeHeight", "NomShoeHeight")).PropValue);

                double materialType=0;
                if (catalogPart.SupportsInterface("IJUAhsWZNMaterial"))
                    materialType = (double)((PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsWZNMaterial", "MaterialType")).PropValue;
                else
                    materialType = (double)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAhsWZNMaterial", "MaterialType")).PropValue;


                int surfaceFinish = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAhsWZNFinish", "SurfaceFinish")).PropValue;

               

                bomDescription = figureNumber + "." + (nominalDiameter*1000 ).ToString("0000") + "." + (nominalShoeheight*1000).ToString("000") + "-" + materialType + "." + surfaceFinish;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrStrutABOMDescription, "Error in BOMDescription of ShoeBOM.cs."));
            }
            return bomDescription;
        }
    }
}
#endregion