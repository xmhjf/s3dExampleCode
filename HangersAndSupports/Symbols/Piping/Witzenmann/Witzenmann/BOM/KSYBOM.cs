//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Class          : KSYBOM.cs
//   ProID          : HSSmartPart,Ingr.SP3D.Content.Support.Symbols.KSYBOM
//   Author         : VinodPeddi  
//   Creation Date  : 01.12.2015
//   Description    :
//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------  
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
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
    public class KSYBOM : SmartPartComponentDefinition, ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                string figureNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsFigureNumber", "FigureNumber")).PropValue;

                string sStructType;
                PropertyValueCodelist lStringType = (PropertyValueCodelist)part.GetPropertyValue("IJUAhsStructType", "StructType");
                sStructType = lStringType.PropertyInfo.CodeListInfo.GetCodelistItem((int )lStringType.PropValue).DisplayName;
                double sStructDepth = ((double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAhsStructDepth", "StructDepth")).PropValue);
                double sStructWidth = ((double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAhsStructWidth", "StructWidth")).PropValue);


                if (sStructType == "T") 
                {
                    bomDescription = figureNumber + "-" + sStructType  + (sStructWidth * 1000).ToString("000");
                }
                else if (sStructType == "L") 
                {
                    if (sStructWidth <= sStructDepth)
                    {
                        bomDescription = figureNumber + "-" + sStructType + (sStructWidth * 1000).ToString("000") + "x" + (sStructDepth * 1000).ToString("000");
                    }
                    else 
                    {
                        bomDescription = figureNumber + "-" + sStructType + (sStructDepth * 1000).ToString("000") + "x" + (sStructWidth * 1000).ToString("000");
                    }
                }
                else
                {
                    bomDescription = figureNumber + "-" + sStructType + (sStructWidth * 1000).ToString("000") + "x" + (sStructDepth * 1000).ToString("000");
                } 

            }

            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstantBOMDescription, "Error in BOMDescription of KSYBOM.cs"));
                return "";
            }
            return bomDescription;
        }
    }
}
#endregion