//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   BOM.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.VariableBOM
//   Author       :  Rajeswari
//   Creation Date:  10/06/2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   10/06/2013   Rajeswari  CR-CP-222469-Initial Creation
//   06/06/2016     Vinay    TR-CP-296065	Fix new coverity issues found in H&S Content 
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
    public class VariableBOM : SmartPartComponentDefinition, ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Double installedLoadValue = 0, operatingLoadValue = 0, rodSpacingValue = 0;
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                // Get the values from the Part / Part Occurence
                string partDescription = part.PartDescription;
                Double movementValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsMovement", "Movement")).PropValue;
                int movementDirValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJUAhsMovementDir", "MovementDirection")).PropValue;
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
                    catch
                    {
                        installedLoadValue = 0;
                    }
                }
                else if (supportOrComponent.SupportsInterface("IJUAhsInstalledLoad"))
                {
                    try
                    {
                        installedLoadValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsInstalledLoad", "InstalledLoad")).PropValue;
                    }
                    catch
                    {
                        installedLoadValue = 0;
                    }
                }
                int coatingValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJCoatingInfo", "CoatingType")).PropValue;
                int coilCoatingValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJUAhsCoilCoating", "CoilCoatingType")).PropValue;

                UnitName primaryDistanceUnits = UnitName.DISTANCE_METER, secondaryDistanceUnits = UnitName.DISTANCE_METER, forceUnits = UnitName.FORCE_KILONEWTON;
                string movement = "", movementDir, operatingLoad="", installedLoad="", coating, coilCoating, rodSpacing="";

                movementDir = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJUAhsMovementDir", "MovementDirection")).PropertyInfo.CodeListInfo.GetCodelistItem(movementDirValue).DisplayName;
                coating = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJCoatingInfo", "CoatingType")).PropertyInfo.CodeListInfo.GetCodelistItem(coatingValue).DisplayName;
                coilCoating = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJCoatingInfo", "CoatingType")).PropertyInfo.CodeListInfo.GetCodelistItem(coilCoatingValue).DisplayName;

                BusinessObject support = supportOrComponent.GetRelationship("SupportHasComponents", "Support").TargetObjects[0];

                GenericHelper genericHelper = new GenericHelper((Ingr.SP3D.Support.Middle.Support)support);
                Collection<object> collection = new Collection<object>();
                genericHelper.GetDataByRule("PartBOMPrimarydistanceUnits", support, out collection);
                if (collection!=null)
                    primaryDistanceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMSecondaryDistanceUnits", support, out collection);
                if (collection != null)
                    secondaryDistanceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMDistancePrecisionType", support, out collection);

                if (collection != null)
                    movement = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, movementValue, GetUOMFormat(UnitType.Distance, (PrecisionType)collection[0]), primaryDistanceUnits, secondaryDistanceUnits, UnitName.UNIT_NOT_SET);

                genericHelper.GetDataByRule("PartBOMForceUnits", support, out collection);
                if (collection != null)
                    forceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMForcePrecisionType", support, out collection);
                if (collection != null)
                {
                    operatingLoad = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, operatingLoadValue, GetUOMFormat(UnitType.Force, (PrecisionType)collection[0]), forceUnits);
                    installedLoad = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, installedLoadValue, GetUOMFormat(UnitType.Force, (PrecisionType)collection[0]), forceUnits);
                }
                
                bomDescription = partDescription + ", Operating Load = " + double.Parse(operatingLoad.Split(' ').GetValue(0).ToString()) + "." + " " + operatingLoad.Split(' ').GetValue(1) + ", Movement = " + movement + " " + movementDir + ", Installed Load = " + double.Parse(installedLoad.Split(' ').GetValue(0).ToString()) + "." + " " + installedLoad.Split(' ').GetValue(1);
                try
                {
                    rodSpacingValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsCC", "CC")).PropValue;
                }
                catch
                {
                    rodSpacingValue = 0;
                }

                if (rodSpacingValue != 0)
                {
                    genericHelper.GetDataByRule("PartBOMDistancePrecisionType", support, out collection);
                    if (collection != null)
                        rodSpacing = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, rodSpacingValue, GetUOMFormat(UnitType.Distance,(PrecisionType)collection[0]), primaryDistanceUnits, secondaryDistanceUnits, UnitName.UNIT_NOT_SET);
                    bomDescription = bomDescription + ", Rod Spacing = " + rodSpacing;
                }

                if (coating != "Undefined")
                    bomDescription = bomDescription + ", Finish = " + coating;

                if (coilCoating != "Undefined" && coilCoating != "" && coilCoating != coating)
                    bomDescription = bomDescription + ", Spring Coil Finish = " + coilCoating;
           
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrVariableSpringBOMDescription, "Error in BOMDescription of VariableBOM.cs"));
                return "";
            }
     return bomDescription;
        }
    }
}
#endregion
