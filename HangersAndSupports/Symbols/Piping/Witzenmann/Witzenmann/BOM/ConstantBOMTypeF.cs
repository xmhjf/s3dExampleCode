//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ZVxBOM.cs
//   Witzenmann,Ingr.SP3D.Content.Support.Symbols.ConstantBOMTypeF
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
    public class ConstantBOMTypeF : SmartPartComponentDefinition, ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                //Double loadValue = 0, rodSpacingValue = 0;
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
              
                string figureNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsFigureNumber", "FigureNumber")).PropValue;
                int LGV = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsWLGV", "OLGV")).PropValue;
                int size = Convert.ToInt32((string)((PropertyValueString)part.GetPropertyValue("IJUAhsSize", "Size")).PropValue);
                Double totalTravelValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsTotalTravel", "TotalTravel")).PropValue;

                bomDescription = figureNumber + " " + size.ToString("00") + "." + (totalTravelValue * 1000).ToString("000") + "." + (LGV).ToString("00");
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstantBOMDescription, "Error in BOMDescription of ConstantBOM.cs"));
                return "";
            }
            return bomDescription;
        }
    }
}
#endregion
