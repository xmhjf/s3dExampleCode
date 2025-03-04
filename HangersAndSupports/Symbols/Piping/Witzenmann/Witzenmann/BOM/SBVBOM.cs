//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ZVxBOM.cs
//   Witzenmann,Ingr.SP3D.Content.Support.Symbols.SBVBOM
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
    public class SBVBOM : SmartPartComponentDefinition, ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double lengthValue, repOverLength1, minlen, maxlen;

                if (catalogPart.SupportsInterface("IJUAhsLength"))
                    lengthValue = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsLength", "Length")).PropValue;
                else
                    lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;

                try
                {
                    repOverLength1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsRepOverLen1", "RepOverLen1")).PropValue;
                }
                catch
                {
                    repOverLength1 = 0;
                }
                minlen = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsMinLen", "MinLen")).PropValue;
                maxlen = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsMaxLen", "MaxLen")).PropValue;
                string length = string.Empty, bomValue = string.Empty;
                try
                {
                    PropertyValueCodelist lBomLengthUnits = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits");
                    bomValue = lBomLengthUnits.PropertyInfo.CodeListInfo.GetCodelistItem(lBomLengthUnits.PropValue).DisplayName;
                }
                catch
                {
                    bomValue = "in";
                }

                if (bomValue.ToUpper() == "IN")
                    length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_INCH);
                else if (bomValue.ToUpper() == "FT")
                    length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_FOOT);
                else if (bomValue.ToUpper() == "MM")
                    length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);
                else if (bomValue.ToUpper() == "M")
                    length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_METER);

                if (lengthValue < minlen || lengthValue > maxlen)
                {
                    string maxLength = string.Empty, minLength = string.Empty;

                    try
                    {
                        PropertyValueCodelist bomList = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits");
                        bomValue = bomList.PropertyInfo.CodeListInfo.GetCodelistItem(bomList.PropValue).DisplayName;
                    }
                    catch
                    {
                        bomValue = "in";
                    }
                    if (bomValue.ToUpper() == "IN")
                    {
                        minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minlen, UnitName.DISTANCE_INCH);
                        maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxlen, UnitName.DISTANCE_INCH);
                    }
                    else if (bomValue.ToUpper() == "FT")
                    {
                        minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minlen, UnitName.DISTANCE_FOOT);
                        maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxlen, UnitName.DISTANCE_FOOT);
                    }
                    else if (bomValue.ToUpper() == "MM")
                    {
                        minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minlen, UnitName.DISTANCE_MILLIMETER);
                        maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxlen, UnitName.DISTANCE_MILLIMETER);
                    }
                    else if (bomValue.ToUpper() == "M")
                    {
                        minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minlen, UnitName.DISTANCE_METER);
                        maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxlen, UnitName.DISTANCE_METER);
                    }

                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidMinMaxLength, "Length of the strut must be between" + minLength + "and" + maxLength));

                    if (lengthValue < minlen)
                        lengthValue = minlen;
                    if (lengthValue > maxlen)
                        lengthValue = maxlen;

                    try
                    {
                        oSupportOrComponent.SetPropertyValue(lengthValue, "IJOAHgrOccLength", "Length");
                    }
                    catch { }
                }


                lengthValue = lengthValue + repOverLength1;
                //bomDescription = catalogPart.PartDescription;

                if (lengthValue > 0)
                {
                    
                    string design = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJUAhsWDesign", "Design")).PropValue;
                    double nominalLoad = ((double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAhsWNominalLoad", "NominalLoad")).PropValue);
                    bomDescription = "SBV" + " " + (nominalLoad/1000).ToString("0000") + "."   + (lengthValue * 1000).ToString("0000") +"."+ design;
                }
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrStrutABOMDescription, "Error in BOMDescription of SBVBOM.cs."));
            }
            return bomDescription;
        }
    }
}
#endregion