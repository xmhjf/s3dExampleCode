//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ConstantBOM.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ConstantBOM
//   Author       : Vijay  
//   Creation Date:  06.June.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   06.June.2013    Vijay    CR-CP-222480  Convert HS_S3DConstant Smartpart VB Project to C# .Net  
//   06-June.2016    Vinay    TR-CP-296065	Fix new coverity issues found in H&S Content 
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
    public class ConstantBOM : SmartPartComponentDefinition, ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Double loadValue = 0, rodSpacingValue = 0;
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                //Get the values from the Part / Part Occurence
                string partDescription = part.PartDescription;
                Double movementValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsMovement", "Movement")).PropValue;
                int movementDirValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJUAhsMovementDir", "MovementDirection")).PropValue;
                if (supportOrComponent.SupportsInterface("IJUAhsLoad"))
                    loadValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsLoad", "Load")).PropValue;
                else if (supportOrComponent.SupportsInterface("IJUAhsLoadG"))
                    loadValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsLoadG", "Load")).PropValue;
                Double totalTravelValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsTotalTravel", "TotalTravel")).PropValue;
                int coatingValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJCoatingInfo", "CoatingType")).PropValue;

                UnitName primaryDistanceUnits = UnitName.DISTANCE_METER, secondaryDistanceUnits = UnitName.DISTANCE_METER, forceUnits = UnitName.FORCE_KILONEWTON;
                string movement="", movementDir, load ="", coating, totalTravel="", rodSpacing = "";

                movementDir = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJUAhsMovementDir", "MovementDirection")).PropertyInfo.CodeListInfo.GetCodelistItem(movementDirValue).DisplayName;
                coating = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJCoatingInfo", "CoatingType")).PropertyInfo.CodeListInfo.GetCodelistItem(coatingValue).DisplayName;

                RelationCollection hgrRelation = supportOrComponent.GetRelationship("SupportHasComponents", "Support");
                BusinessObject businessObject = hgrRelation.TargetObjects[0];
                GenericHelper genericHelper = new GenericHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                Collection<object> collection = new Collection<object>();

                genericHelper.GetDataByRule("PartBOMPrimarydistanceUnits", businessObject, out collection);
                if (collection != null)
                    primaryDistanceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMSecondaryDistanceUnits", businessObject, out collection);
                if (collection != null)
                    secondaryDistanceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMDistancePrecisionType", businessObject, out collection);
                if (collection != null)
                    movement = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, movementValue, GetUOMFormat(UnitType.Distance, (PrecisionType)collection[0]), primaryDistanceUnits, secondaryDistanceUnits, UnitName.UNIT_NOT_SET);

                genericHelper.GetDataByRule("PartBOMForceUnits", businessObject, out collection);
                if (collection != null)
                    forceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMForcePrecisionType", businessObject, out collection);
                if (collection != null)
                {
                    load = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, loadValue, GetUOMFormat(UnitType.Force, (PrecisionType)collection[0]), forceUnits);
                    totalTravel = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, totalTravelValue, GetUOMFormat(UnitType.Distance, (PrecisionType)collection[0]), primaryDistanceUnits, secondaryDistanceUnits, UnitName.UNIT_NOT_SET);
                }

                bomDescription = partDescription + ", Total Travel = " + totalTravel + ", Movement = " + movement + " " + movementDir + ", Load = " + double.Parse(load.Split(' ').GetValue(0).ToString()) + "." + " " + load.Split(' ').GetValue(1);
                if (part.SupportsInterface("IJUAhsCC"))
                    rodSpacingValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsCC", "CC")).PropValue;
                else
                    rodSpacingValue = 0;

                if (rodSpacingValue != 0)
                {
                    genericHelper.GetDataByRule("PartBOMDistancePrecisionType", businessObject, out collection);
                    if (collection != null)
                        rodSpacing = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, rodSpacingValue, GetUOMFormat(UnitType.Distance, (PrecisionType)collection[0]), primaryDistanceUnits, secondaryDistanceUnits, UnitName.UNIT_NOT_SET);
                    bomDescription = bomDescription + ", Rod Spacing = " + rodSpacing;
                }

                if (coating != "Undefined")
                    bomDescription = bomDescription + ", Finish = " + coating;
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