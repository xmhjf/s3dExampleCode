//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Rod.cs
//   PTPBOM,Ingr.SP3D.Content.Support.Symbols.Rod
//   Author       :  Hema
//   Creation Date:  09-10-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-10-2013   Hema       CR-CP-240907  Convert HS_PTPBOM to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using System.Linq;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Collections.Generic;
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

                int bomLengthUnitsValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsBOMLenUnits", "BOMLenUnits")).PropValue;
                if (bomLengthUnitsValue == -1)
                    bomLengthUnitsValue = 6;
                string bomLengthUnits = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsBOMLenUnits", "BOMLenUnits")).PropertyInfo.CodeListInfo.GetCodelistItem(bomLengthUnitsValue).DisplayName;

                //Get the values from the Part / Part Occurence
                string partDescription = part.PartDescription;

                string[] splitPartDescripition = partDescription.Split(',');
                string figureNumber = splitPartDescripition[0];
                string rodDiameter = splitPartDescripition[1];
                string shortDescription = splitPartDescripition[2];

                string threadLength = string.Empty;
                try
                {
                    threadLength = splitPartDescripition[3];
                }
                catch
                {
                    threadLength = string.Empty;
                }

                //Get the values from the Part / Part Occurence
                double lengthValue = 0;
                try
                {
                    lengthValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;
                }
                catch
                {
                    lengthValue = 0;
                }
                int finishValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPFinish", "Finish")).PropValue;

                PrecisionType precisionType = new PrecisionType();
                UnitName primaryUnits = new UnitName(), secondaryUnits = new UnitName();
                if (bomLengthUnitsValue != 7)
                {
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    PartClass rodPartClass = (PartClass)catalogBaseHelper.GetPartClass("HgrSrv_BOMUnits");

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
                }
                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, PTPBOMServices.GetUOMFormat(UnitType.Distance, precisionType), primaryUnits, secondaryUnits, UnitName.UNIT_NOT_SET);
                string finish = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPFinish", "Finish")).PropertyInfo.CodeListInfo.GetCodelistItem(finishValue).DisplayName;

                if (threadLength.Equals(""))
                {
                    if (finishValue == -1 && bomLengthUnitsValue == 7)  //For Finish is nothing and No Length Unit Type
                        bomDescription = figureNumber + ", (" + rodDiameter + ")," + shortDescription;
                    else if (finishValue == -1)  //For Finish is nothing
                        bomDescription = figureNumber + ", (" + rodDiameter + ", X " + length + ")," + shortDescription;
                    else if (bomLengthUnitsValue == 7) //For No Length Unit Type
                        bomDescription = figureNumber + ", (" + rodDiameter + ")," + shortDescription + ", " + finish;
                    else
                        bomDescription = figureNumber + ", (" + rodDiameter + " X " + length + ")," + shortDescription + ", " + finish;
                }
                else
                {
                    if (finishValue == -1 && bomLengthUnitsValue == 7)  //For Finish is nothing and No Length Unit Type
                        bomDescription = figureNumber + ", (" + rodDiameter + ")," + shortDescription + "," + threadLength;
                    else if (finishValue == -1)  //For Finish is nothing
                        bomDescription = figureNumber + ", (" + rodDiameter + " X " + length + ")," + shortDescription + "," + threadLength;
                    else if (bomLengthUnitsValue == 7) //For No Length Unit Type
                        bomDescription = figureNumber + ", (" + rodDiameter + " )," + shortDescription + "," + threadLength + ", " + finish;
                    else
                        bomDescription = figureNumber + ", (" + rodDiameter + " X " + length + ")," + shortDescription + "," + threadLength + ", " + finish;
                }
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PTPBOMLocalizer.GetString(PTPBOMResourceIDs.ErrBOMDescription, "Error in BOMDescription of Rod"));
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
