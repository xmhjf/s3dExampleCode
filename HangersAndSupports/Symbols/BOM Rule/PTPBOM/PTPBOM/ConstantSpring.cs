//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ConstantSpring.cs
//   PTPBOM,Ingr.SP3D.Content.Support.Symbols.ConstantSpring
//   Author       :  Hema
//   Creation Date:  09-10-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-10-2013     Hema       CR-CP-240907  Convert HS_PTPBOM to C# .Net
//   06-06-2016     PVK        TR-CP-296065	Fix new coverity issues found in H&S Content   
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
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
    public class ConstantSpring : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)supportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                // Get the values from the Part / Part Occurence
                string partDescription = part.PartDescription;
                string[] splitPartDescripition = partDescription.Split(',');
                string figureNumber = splitPartDescripition[0];
                string shortDescription = splitPartDescripition[1];

                //Get the values from the Part / Part Occurence
                double loadValue = 0;

                if (supportOrComponent.SupportsInterface("IJUAhsLoad"))
                    loadValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsLoad", "Load")).PropValue;
                else if (supportOrComponent.SupportsInterface("IJUAhsLoadG"))
                    loadValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsLoadG", "Load")).PropValue;

                double totalTravelValue = (double)((PropertyValueDouble)supportOrComponent.GetPropertyValue("IJUAhsTotalTravel", "TotalTravel")).PropValue;
                int finishValue = (int)((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPFinish", "Finish")).PropValue;

                UnitName primaryDistanceUnits = UnitName.DISTANCE_METER, secondaryDistanceUnits = UnitName.DISTANCE_METER, forceUnits = UnitName.FORCE_KILONEWTON;
                string totalTravel, load;
                int distancePrecision = 0;
                PrecisionType distancePrecisionType = PrecisionType.PRECISIONTYPE_DECIMAL;
                PrecisionType forcePrecitionType = PrecisionType.PRECISIONTYPE_DECIMAL;

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
                genericHelper.GetDataByRule("PartBOMDistancePrecision", support, out collection);
                if (collection != null)
                    distancePrecision = (int)collection[0];

                genericHelper.GetDataByRule("PartBOMForceUnits", support, out collection);
                if (collection != null)
                    forceUnits = (UnitName)collection[0];
                genericHelper.GetDataByRule("PartBOMForcePrecisionType", support, out collection);
                if (collection != null)
                    forcePrecitionType = (PrecisionType)collection[0];

                totalTravel = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, totalTravelValue, PTPBOMServices.GetUOMFormat(UnitType.Distance, distancePrecisionType), primaryDistanceUnits, secondaryDistanceUnits, UnitName.UNIT_NOT_SET);
                load = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Force, loadValue, PTPBOMServices.GetUOMFormat(UnitType.Force, forcePrecitionType), forceUnits);
                string finish = ((PropertyValueCodelist)supportOrComponent.GetPropertyValue("IJOAhsPTPFinish", "Finish")).PropertyInfo.CodeListInfo.GetCodelistItem(finishValue).DisplayName;
                if (finishValue == -1)  //For Finish is nothing
                    bomDescription = figureNumber + "," + shortDescription + ", ( Load = " + double.Parse(load.Split(' ').GetValue(0).ToString()) + "." + " " + load.Split(' ').GetValue(1) + " ), ( Total Load = " + totalTravel + " )";
                else
                    bomDescription = figureNumber + "," + shortDescription + ", ( Load = " + double.Parse(load.Split(' ').GetValue(0).ToString()) + "." + " " + load.Split(' ').GetValue(1) + " ), ( Total Load = " + totalTravel + " ), " + finish;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PTPBOMLocalizer.GetString(PTPBOMResourceIDs.ErrBOMDescription, "Error in BOMDescription of ConstantSpring"));
                return "";
            }
            return bomDescription;
        }
    }
}
#endregion
