//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   VariableSpring.cs
//   PTPBOM,Ingr.SP3D.Content.Support.Symbols.VariableSpring
//   Author       :  Hema
//   Creation Date:  07-10-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07-10-2013     Hema       CR-CP-240907  Convert HS_PTPBOM to C# .Net  
//   06-06-2016     PVK        TR-CP-296065	Fix new coverity issues found in H&S Content 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Collections.Generic;

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
    public class VariableSpring : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            IEnumerable<BusinessObject> rodParts = null;
            try
            {
                Double installedLoadValue = 0, operatingLoadValue = 0, rodSpacingValue = 0;
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                // Get the values from the Part / Part Occurence
                string partDescription = part.PartDescription;
                string[] splitPartDescripition = partDescription.Split(',');
                string figureNumber = splitPartDescripition[0];
                string shortDescription = splitPartDescripition[1];

                Double movementValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsMovement", "Movement")).PropValue;
                if (supportOrComponent.SupportsInterface("IJUAhsOperatingLoad"))
                    operatingLoadValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsOperatingLoad", "OperatingLoad")).PropValue;
                else if (supportOrComponent.SupportsInterface("IJUAhsOperatingLoadG"))
                    operatingLoadValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsOperatingLoadG", "OperatingLoad")).PropValue;
                if (supportOrComponent.SupportsInterface("IJUAhsInstalledLoadG"))
                {
                    try
                    {
                        installedLoadValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsInstalledLoadG", "InstalledLoad")).PropValue;
                    }
                    catch { installedLoadValue = 0; }
                }
                else if (supportOrComponent.SupportsInterface("IJUAhsInstalledLoad"))
                {
                    try
                    {
                        installedLoadValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsInstalledLoad", "InstalledLoad")).PropValue;
                    }
                    catch { installedLoadValue = 0; }
                }
                int finishValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPFinish", "Finish")).PropValue;
                string springType = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsSpringType", "SpringType")).PropValue;
                string finish = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPFinish", "Finish")).PropertyInfo.CodeListInfo.GetCodelistItem(finishValue).DisplayName;

                UnitName primaryDistanceUnits = UnitName.DISTANCE_METER, secondaryDistanceUnits = UnitName.DISTANCE_METER, forceUnits = UnitName.FORCE_KILONEWTON;
                string movement, operatingLoad = "", installedLoad = "", rodSpacing;
                 PrecisionType distancePrecisionType = PrecisionType.PRECISIONTYPE_DECIMAL;
                BusinessObject support = supportOrComponent.GetRelationship("SupportHasComponents", "Support").TargetObjects[0];

                //Convert the values from DB Units into the desired units
                GenericHelper genericHelper = new GenericHelper((Ingr.SP3D.Support.Middle.Support)support);
                Collection<object> collection = new Collection<object>();
                genericHelper.GetDataByRule("PartBOMPrimarydistanceUnits", support, out collection);
                if (collection != null)
                    primaryDistanceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMSecondaryDistanceUnits", support, out collection);
                if (collection != null)
                    secondaryDistanceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMDistancePrecisionType", support, out collection);
                if (collection != null)
                    distancePrecisionType = (PrecisionType)collection[0];

                movement = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, movementValue, PTPBOMServices.GetUOMFormat(UnitType.Distance, distancePrecisionType), primaryDistanceUnits, secondaryDistanceUnits, UnitName.UNIT_NOT_SET);

                genericHelper.GetDataByRule("PartBOMForceUnits", support, out collection);
                if (collection != null)
                    forceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMForcePrecisionType", support, out collection);
                if (collection != null)
                {
                    operatingLoad = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, operatingLoadValue, PTPBOMServices.GetUOMFormat(UnitType.Force, (PrecisionType)collection[0]), forceUnits);
                    installedLoad = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, installedLoadValue, PTPBOMServices.GetUOMFormat(UnitType.Force, (PrecisionType)collection[0]), forceUnits);
                }
                if (finishValue == -1)  //For Finish is nothing
                    bomDescription = figureNumber + " , " + shortDescription + " , ( Hot Load = " + double.Parse(operatingLoad.Split(' ').GetValue(0).ToString()) + " ) , ( Cold Load = " + double.Parse(installedLoad.Split(' ').GetValue(0).ToString()) + " ) , Movement = " + movement;
                else
                    bomDescription = figureNumber + " , " + shortDescription + " , ( Hot Load = " + double.Parse(operatingLoad.Split(' ').GetValue(0).ToString()) + " ) , ( Cold Load = " + double.Parse(installedLoad.Split(' ').GetValue(0).ToString()) + " ) , Movement = " + movement + " , " + finish;
                if (springType.Equals("F"))
                {
                    double installedHeight = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsInstalledHeight", "InstalledHeight")).PropValue;
                    double offset1 = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsOffset1", "Offset1")).PropValue;
                    double heightValue = installedHeight + offset1;
                    string height = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, heightValue, PTPBOMServices.GetUOMFormat(UnitType.Distance, distancePrecisionType), primaryDistanceUnits, secondaryDistanceUnits, UnitName.UNIT_NOT_SET);
                    if (finishValue == -1)  //For Finish is nothing
                        bomDescription = figureNumber + " , " + shortDescription + " , ( Hot Load = " + double.Parse(operatingLoad.Split(' ').GetValue(0).ToString()) + " ) , ( Cold Load = " + double.Parse(installedLoad.Split(' ').GetValue(0).ToString()) + " ) , Movement = " + movement + " , ( Height = " + height + " ) ";
                    else
                        bomDescription = figureNumber + " , " + shortDescription + " , ( Hot Load=" + double.Parse(operatingLoad.Split(' ').GetValue(0).ToString()) + " ) , ( Cold Load=" + double.Parse(installedLoad.Split(' ').GetValue(0).ToString()) + " ) , Movement = " + movement + " , ( Height = " + height + " ) , " + finish;
                }
                try
                {
                    rodSpacingValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsCC", "CC")).PropValue;
                }
                catch { rodSpacingValue = 0; }

                if (rodSpacingValue != 0)
                {
                    //Declare Unit Types
                    int bomLengthUnitsValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsBOMLenUnits", "BOMLenUnits")).PropValue;
                    if (bomLengthUnitsValue == -1)
                        bomLengthUnitsValue = 6;
                    string bomLengthUnits = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsBOMLenUnits", "BOMLenUnits")).PropertyInfo.CodeListInfo.GetCodelistItem(bomLengthUnitsValue).DisplayName;
                    PrecisionType precisionType = new PrecisionType();
                    UnitName primaryUnits, secondaryUnits;
                    if (bomLengthUnitsValue != 7)
                    {
                        CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
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
                    rodSpacing = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, rodSpacingValue, PTPBOMServices.GetUOMFormat(UnitType.Distance, precisionType), primaryDistanceUnits, secondaryDistanceUnits, UnitName.UNIT_NOT_SET);
                    if (finishValue == -1 && bomLengthUnitsValue == 7) //For Finish is nothing and No Length Unit Type
                        bomDescription = figureNumber + "," + shortDescription + ", ( Hot Load = " + double.Parse(operatingLoad.Split(' ').GetValue(0).ToString()) + " ) , ( Cold Load = " + double.Parse(installedLoad.Split(' ').GetValue(0).ToString()) + " ) , Movement = " + movement;
                    else if (finishValue == -1) //For Finish is nothing
                        bomDescription = figureNumber + "," + shortDescription + ", ( Hot Load = " + double.Parse(operatingLoad.Split(' ').GetValue(0).ToString()) + " ) , ( Cold Load = " + double.Parse(installedLoad.Split(' ').GetValue(0).ToString()) + " ) , Movement = " + movement + ", ( C to C = " + rodSpacing + " ) ";
                    else if (bomLengthUnitsValue == 7)  //For No Length Unit Type
                        bomDescription = figureNumber + "," + shortDescription + ", ( Hot Load = " + double.Parse(operatingLoad.Split(' ').GetValue(0).ToString()) + " ) , ( Cold Load = " + double.Parse(installedLoad.Split(' ').GetValue(0).ToString()) + " ) , Movement = " + movement + ", " + finish;
                    else
                        bomDescription = figureNumber + "," + shortDescription + ", ( Hot Load = " + double.Parse(operatingLoad.Split(' ').GetValue(0).ToString()) + " ) , ( Cold Load = " + double.Parse(installedLoad.Split(' ').GetValue(0).ToString()) + " ) , Movement = " + movement + ", ( C to C = " + rodSpacing + ") , " + finish;
                }
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
