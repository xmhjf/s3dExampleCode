//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ZVxBOM.cs
//   Witzenmann,Ingr.SP3D.Content.Support.Symbols.ZVxBOM
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
    public class ZVxBOM : SmartPartComponentDefinition, ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string figureNumber = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUAhsFigureNumber", "FigureNumber")).PropValue;

                double Width1_B = ((double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue);
                double Length1_E = ((double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsLength1", "Length1")).PropValue);
                double HP2PosX = ((double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsHolePort2", "HP2PosX")).PropValue);
                double HP1PosX = ((double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsHolePort1", "HP1PosX")).PropValue);
                double EValue = Length1_E-(HP1PosX)-(Length1_E - HP2PosX);


                int materialType = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAhsWZNMaterial", "MaterialType")).PropValue;
                int surfaceFinish = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAhsWZNFinish", "SurfaceFinish")).PropValue;
                int LGV = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAhsWLGV", "ULGV")).PropValue;

                bomDescription = figureNumber + " " + (Width1_B * 1000).ToString("000") + "." + (EValue * 1000).ToString("000") + "." + LGV.ToString("00") + "-" + materialType + "." + surfaceFinish;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrStrutABOMDescription, "Error in BOMDescription of ZVxBOM.cs."));
            }
            return bomDescription;
        }
    }
}
#endregion