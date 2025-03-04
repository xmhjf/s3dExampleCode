//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Rod.cs
//    PSLBOM,Ingr.SP3D.Content.Support.Symbols.StrutSNCX
//   Author       : PVK
//   Creation Date:  17-06-2014
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-06-2014      PVK     DM-CP-250540 Deliver new Smart Parts for PSL
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System.Collections.Generic;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System;

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
    public class StrutSNCX : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double lengthValue,  minlen, maxlen ;

                lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAhsLength", "Length")).PropValue;
                minlen = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsMinLen", "MinLen")).PropValue;
                maxlen = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsMaxLen", "MaxLen")).PropValue;
                string length = string.Empty, bomValue = string.Empty;

                PropertyValueCodelist lBomLengthUnits = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits");
                bomValue = lBomLengthUnits.PropertyInfo.CodeListInfo.GetCodelistItem(lBomLengthUnits.PropValue).DisplayName;

                length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);


                if (lengthValue < minlen || lengthValue > maxlen)
                {
                    string maxLength = string.Empty, minLength = string.Empty;
                    PropertyValueCodelist bomList = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits");
                    bomValue = bomList.PropertyInfo.CodeListInfo.GetCodelistItem(bomList.PropValue).DisplayName;
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minlen, UnitName.DISTANCE_MILLIMETER);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxlen, UnitName.DISTANCE_MILLIMETER);

                    ToDoListMessage ToDolistMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLBOMLocalizer.GetString(PSLBOMResourceIDs.ErrInvalidMinMaxLength, "Length of the strut must be between" + minLength + "and" + maxLength));
                    if (lengthValue < minlen)
                        lengthValue = minlen;
                    if (lengthValue > maxlen)
                        lengthValue = maxlen;

                    try
                    {
                        oSupportOrComponent.SetPropertyValue(lengthValue, "IJUAhsLength", "Length");
                    }
                    catch { }
                }

                bomDescription = catalogPart.PartDescription;

                if (lengthValue > 0)
                    bomDescription = bomDescription + ",Length=" + length;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage ToDolistMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLBOMLocalizer.GetString(PSLBOMResourceIDs.ErrBOMDescription, "Error in BOMDescription of StructSNCX"));
            }
            return bomDescription;
        }

    }
}
#endregion
