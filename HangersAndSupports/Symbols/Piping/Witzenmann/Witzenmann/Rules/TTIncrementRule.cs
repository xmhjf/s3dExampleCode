//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ZVxBOM.cs
//   Witzenmann,Ingr.SP3D.Content.Support.Symbols.TTIncrementRule
//   Author       : Vinod  
//   Creation Date:  01.12.2015
//   DI-CP-282684  Integrate the newly developed Witzenmann Parts into Product  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class TTIncrementRule : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.SupportComponent supportComponent = (Ingr.SP3D.Support.Middle.SupportComponent)SupportOrComponent;
            Part oPart = (Part)supportComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

            double[] totalTravelIncrement = new double[1];

            double totalTravel;
            int percentOT;
            double movement;
            double minOT;

            if (supportComponent.SupportsInterface("IJUAhsMovement"))
                movement = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAhsMovement", "Movement")).PropValue;
            else
                movement = (double)((PropertyValueDouble)oPart.GetPropertyValue("IJUAhsMovement", "Movement")).PropValue;

            if (supportComponent.SupportsInterface("IJUAhsOverTravelPercent"))
                percentOT = (int)((PropertyValueInt)supportComponent.GetPropertyValue("IJUAhsOverTravelPercent", "OverTravelPercent")).PropValue;
            else
                percentOT = (int)((PropertyValueInt)oPart.GetPropertyValue("IJUAhsOverTravelPercent", "OverTravelPercent")).PropValue;


            if (supportComponent.SupportsInterface("IJUAhsOverTravelMin"))
                minOT = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAhsOverTravelMin", "OverTravelMin")).PropValue;
            else
                minOT = (double)((PropertyValueDouble)oPart.GetPropertyValue("IJUAhsOverTravelMin", "OverTravelMin")).PropValue;

            percentOT = percentOT / 100;
            if (movement * percentOT > minOT)
                totalTravel = movement * (1 + percentOT);
            else
                totalTravel = movement + minOT;

            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            PartClass TTIncrementByRule = (PartClass)catalogBaseHelper.GetPartClass("WZN_TTIncrementByRule");

            ReadOnlyCollection<BusinessObject> loadTableClassItems = TTIncrementByRule.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            foreach (BusinessObject classItem in loadTableClassItems)
            {
                double minTravel = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsMinTravel", "MinTravel")).PropValue) / 1000;
                double maxTravel = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsMaxTravel", "MaxTravel")).PropValue) / 1000;

                if (totalTravel > minTravel && totalTravel <= maxTravel)
                {
                    totalTravelIncrement[0] = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsTravelIncrement", "TravelIncrement")).PropValue;
                    break;
                }
            }
            return totalTravelIncrement;
        }
    }
}