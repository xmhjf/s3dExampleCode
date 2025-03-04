//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ZVxBOM.cs
//   Witzenmann,Ingr.SP3D.Content.Support.Symbols.ZKNBOM
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
    public class ZKNBOM : SmartPartComponentDefinition, ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string figureNumber = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUAhsFigureNumber", "FigureNumber")).PropValue;

               // if (catalogPart.SupportsInterface("IJUAhsWDesign"))

                string design = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUAhsWDesign", "Design")).PropValue;
                int LGV = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAhsWLGV", "ULGV")).PropValue;
                double stuctureWidth = ((double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAhsStructWidth", "StructWidth")).PropValue)*1000;
                double clampThk = ((double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAhsClampThick", "ClampThickness")).PropValue)*1000;
                int surfaceFinish = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAhsWZNFinish", "SurfaceFinish")).PropValue;
                if (figureNumber == "ZKN")
                {
                    if (design == "1")
                        bomDescription = figureNumber + " " + design + "." + LGV + "." + stuctureWidth.ToString("000") + "-" + surfaceFinish;
                    else
                        bomDescription = figureNumber + " " + design + "." + LGV + "." + stuctureWidth.ToString("000") +"."+ clampThk.ToString("00") + "-" + surfaceFinish;
                }
                else
                {
                    bomDescription = figureNumber + " " + LGV + "." + stuctureWidth.ToString("000") + "."+clampThk.ToString("00") + "-" + surfaceFinish;
                }

            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrStrutABOMDescription, "Error in BOMDescription of Struct_A.cs."));
            }
            return bomDescription;
        }
    }
}
#endregion