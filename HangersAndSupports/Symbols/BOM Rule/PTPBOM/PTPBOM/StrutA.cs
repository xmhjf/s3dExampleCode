//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   StrutA.cs.cs
//   PTPBOM,Ingr.SP3D.Content.Support.Symbols.StrutA.cs
//   Author       :  Hema
//   Creation Date:  09-10-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-10-2013   Hema       CR-CP-240907  Convert HS_PTPBOM to C# .Net    
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Linq;
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
    public class StrutA : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            IEnumerable<BusinessObject> rodParts = null;
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                //Get the values from the Part / Part Occurence
                string partDescription = part.PartDescription;
                string structAssemblyInfo = (string)((PropertyValueString)part.GetPropertyValue("IJOAhsPTPStrutAssyInfo", "StrutAssyInfo")).PropValue;

                string[] splitPartDescripition = partDescription.Split(',');
                string figureNumber = splitPartDescripition[0];
                string size = splitPartDescripition[1];
                string shortDescription = splitPartDescripition[2];
                string option = splitPartDescripition[3];

                double ccLengthValue = 0; 
                int unitTypeValue = 0;
                string unitType = string.Empty, ccLength = string.Empty, finish = string.Empty;
                int finishValue = 0;
                if (!string.IsNullOrEmpty(structAssemblyInfo))
                {
                    //get the unit type and CC Length from Custom BOM
                    string[] splitStructInfo = structAssemblyInfo.Split(':');
                    string splitCCLength = splitStructInfo[1];
                    string[] splitCCAndFinish = splitCCLength.Split(',');
                    ccLengthValue = Convert.ToDouble(splitCCAndFinish[0]);
                    string lengthUnitType = splitCCAndFinish[1];
                    string[] splitCCAndUnits = lengthUnitType.Split(',');
                    unitType = splitCCAndUnits[0];

                    //get the finish value from custom BOM
                    string[] finishArray = structAssemblyInfo.Split(',');
                    string finishAttribute = finishArray[2];
                    string[] splitFinish = finishAttribute.Split(':');
                    finishValue = int.Parse(splitFinish[1]);

                    unitTypeValue = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsBOMLenUnits", "BOMLenUnits")).PropertyInfo.CodeListInfo.GetCodelistItem(unitType).DisplayName.Length;
                    PrecisionType precisionType = new PrecisionType();

                    UnitName primaryUnits = new UnitName(), secondaryUnits = new UnitName();
                    //Declare Unit Types
                    if (unitTypeValue == 7)
                    {
                        CatalogBaseHelper cataloghelper = new CatalogBaseHelper();

                        CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                        PartClass rodPartClass = (PartClass)catalogBaseHelper.GetPartClass("HgrSrv_BOMUnits");

                        if (rodPartClass.PartClassType.Equals("HgrServiceClass"))
                            rodParts = rodPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        else
                            rodParts = rodPartClass.Parts;

                        rodParts = rodParts.Where(part1 => ((PropertyValueString)part1.GetPropertyValue("IJUAhsUnits", "BOMUnitRule")).PropValue.Equals(unitType));
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
                    }
                    ccLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, ccLengthValue, PTPBOMServices.GetUOMFormat(UnitType.Distance, precisionType), primaryUnits, secondaryUnits, UnitName.UNIT_NOT_SET);
                    finish = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPFinish", "Finish")).PropertyInfo.CodeListInfo.GetCodelistItem(finishValue).DisplayName;
                }
                if (!string.IsNullOrEmpty(structAssemblyInfo))
                {
                    if (finishValue == -1 && unitTypeValue == 7)  //For Finish is nothing and No Length Unit Type
                        bomDescription = figureNumber + ", (" + size + " )," + shortDescription + "," + option;
                    else if (finishValue == -1)  //For Finish is nothing
                        bomDescription = figureNumber + ", (" + size + " , C-C= " + ccLength + ")," + shortDescription + "," + option;
                    else if (unitTypeValue == 7) //For No Length Unit Type
                        bomDescription = figureNumber + ", (" + size + " )," + shortDescription + "," + option + ", " + finish;
                    else
                        bomDescription = figureNumber + ", (" + size + " , C-C= " + ccLength + ")," + shortDescription + "," + option + ", " + finish;
                }
                else
                    bomDescription = partDescription;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PTPBOMLocalizer.GetString(PTPBOMResourceIDs.ErrBOMDescription, "Error in BOMDescription of BOM.cs"));
                return "";
            }
            finally
            {
                if (rodParts is IDisposable)
                {
                    ((IDisposable)rodParts).Dispose(); // This line will be executed
                }
            }
            return bomDescription;
        }
    }
}
#endregion
