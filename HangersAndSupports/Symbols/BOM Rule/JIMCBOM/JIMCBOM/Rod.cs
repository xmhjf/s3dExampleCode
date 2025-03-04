//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Rod.cs
//    JIMCBOM,Ingr.SP3D.Content.Support.Symbols.Rod
//   Author       :  Hema
//   Creation Date:  07-10-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07-10-2013   Hema       CR-CP-240907  Convert HS_JIMCBOM to C# .Net  
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
    public class Rod : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            IEnumerable<BusinessObject> rodParts = null;
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                //Get the Length Units
                int bomLengthUnitsValue = (int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits")).PropValue;
                string bomLengthUnits = ((PropertyValueCodelist)part.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits")).PropertyInfo.CodeListInfo.GetCodelistItem(bomLengthUnitsValue).DisplayName;

                string partDescription = part.PartDescription; double lengthValue = 0;
                try
                {
                    lengthValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;
                }
                catch { lengthValue = 0; }

                //Convert the values from DB Units into the desired units
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass rodPartClass = (PartClass)catalogBaseHelper.GetPartClass("HgrSrv_BOMUnits");
                PrecisionType precisionType = new PrecisionType();
                UnitName primaryUnits = new UnitName(), secondaryUnits = new UnitName();
                if (rodPartClass.PartClassType.Equals("HgrServiceClass"))
                    rodParts = rodPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    rodParts = rodPartClass.Parts;
                rodParts = rodParts.Where(part1 => ((string)((PropertyValueString)part1.GetPropertyValue("IJUAhsUnits", "BOMUnitRule")).PropValue).Equals(bomLengthUnits));
                if (rodParts.Count() > 0)
                {
                    primaryUnits = ((UnitName)((PropertyValueInt)rodParts.ElementAt(0).GetPropertyValue("IJUAhsUnits", "PrimaryUnits")).PropValue);
                    try
                    {
                        secondaryUnits = ((UnitName)((PropertyValueInt)rodParts.ElementAt(0).GetPropertyValue("IJUAhsUnits", "SecondaryUnits")).PropValue);
                    }
                    catch
                    {
                        secondaryUnits = (UnitName)345;
                    }
                    precisionType = ((PrecisionType)((PropertyValueInt)rodParts.ElementAt(0).GetPropertyValue("IJUAhsUnits", "PrecisionType")).PropValue);
                }
                UOMFormat uomFormat = MiddleServiceProvider.UOMMgr.GetDefaultUnitFormat(UnitType.Distance);
                uomFormat.PrecisionType = precisionType;
                if (uomFormat.PrecisionType == PrecisionType.PRECISIONTYPE_FRACTIONAL)
                    uomFormat.FractionalPrecision = uomFormat.FractionalPrecision;
                if (uomFormat.PrecisionType == PrecisionType.PRECISIONTYPE_DECIMAL)
                    uomFormat.DecimalPrecision = uomFormat.DecimalPrecision;

                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, uomFormat, primaryUnits, secondaryUnits, UnitName.UNIT_NOT_SET);
                if (length.Contains('.'))
                    bomDescription = partDescription + " " + double.Parse(length.Split('.').GetValue(0).ToString()) + " " + length.Split(' ').GetValue(1); //Final BOM Description
                else
                    bomDescription = partDescription + " " + length; //Final BOM Description

                return bomDescription;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, JIMCBOMLocalizer.GetString(JIMCBOMResourceIDs.ErrBOMDescription, "Error in BOMDescription of AnvilBOMwithRodFinish"));
                return "";
            }
            finally
            {
                if (rodParts is IDisposable)
                {
                    ((IDisposable)rodParts).Dispose(); // This line will be executed
                }
            }
        }
    }
}
#endregion
