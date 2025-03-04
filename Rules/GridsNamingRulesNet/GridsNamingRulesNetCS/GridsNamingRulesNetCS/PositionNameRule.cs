//***************************************************************************************
//  Copyright (C) 2009, Intergraph Corporation.  All rights reserved.
//
//  Project  : \GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetCS
//  Class    : PositionNameRule
//  Abstract : The file contains  implementation of Position NamingRule
//
//  Chaitanya   April '09       Creation
//  Chaitanya   09/07/2009      TR- 170922  Error is displayed when selecting .Net NameRule for a cylindrical plane. 
//***************************************************************************************
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Grids.Middle;

// Define user namespace
namespace GridsNamingRulesNetCS
{
    public class PositionNameRule : NameRuleBase
    {
        private const string strDecimalFormat = "0.000";   //Precision value of the Position 

        /// <summary>
        /// Creates a name for the object passed in. The name is based on the parents
        /// name and object name.The Naming Parents are added in AddNamingParents().
        /// Both these methods are called from naming rule semantic.
        /// </summary>
        /// <param name="oEntity">Child object that needs to have the naming rule naming.</param>
        /// <param name="oParents">Naming parents collection.</param>
        /// <param name="oActiveEntity">Naming rules active entity on which the NamingParentsString is stored.</param>
        /// <exception cref="ArgumentNullException">The Grids entity is null.</exception>
        public override void ComputeName(BusinessObject pEntity, ReadOnlyCollection<BusinessObject> pParents, BusinessObject pActiveEntity)
        {
            try
            {
               if (pEntity == null)
                {
                    throw new ArgumentNullException(); 
                }

                GridPlaneBase       pGridEntity = null;
                GridCylinder        pGridCylinderEntity = null;

                CoordinateSystem    oCS = null;

                // Initializing enums with default values.
                AxisType Axis = AxisType.X;
                CoordinateSystem.CoordinateSystemType CSType = CoordinateSystem.CoordinateSystemType.Grids;

                //Radial Cylinder supports GridCylinder
                if (pEntity is GridPlaneBase)
                {
                    pGridEntity = (GridPlaneBase)pEntity;
                    oCS = pGridEntity.Axis.CoordinateSystem;

                    Axis = pGridEntity.Axis.AxisType;
                    CSType = oCS.Type;                    
                }
                else if (pEntity is GridCylinder)
                {
                    pGridCylinderEntity = (GridCylinder)pEntity;
                    oCS = pGridCylinderEntity.Axis.CoordinateSystem;

                    Axis = pGridCylinderEntity.Axis.AxisType;
                }
     
               // Basing on the AxisType, set the name of the Plane.
               switch (Axis)
               {
                   case AxisType.X:
                       pEntity.SetPropertyValue((CSType == CoordinateSystem.CoordinateSystemType.Ship ? "F" : "E") + " " + (pGridEntity.DistanceFromOrigin).ToString(strDecimalFormat) + " m", "IJNamedItem", "Name");
                       break;
                   case AxisType.Y:
                       pEntity.SetPropertyValue((CSType == CoordinateSystem.CoordinateSystemType.Ship ? "L" : "N") + " " + (pGridEntity.DistanceFromOrigin).ToString(strDecimalFormat) + " m", "IJNamedItem", "Name");
                       break;
                   case AxisType.Z:
                       pEntity.SetPropertyValue((CSType == CoordinateSystem.CoordinateSystemType.Ship ? "D" : "El") + " " + (pGridEntity.DistanceFromOrigin).ToString(strDecimalFormat) + " m", "IJNamedItem", "Name");
                       break;
                   case AxisType.Radial:
                       pEntity.SetPropertyValue("R" + " " + (pGridEntity.DistanceFromOrigin * 180 / 3.14159265358979).ToString(strDecimalFormat) + " deg", "IJNamedItem", "Name");
                       break;
                   case AxisType.Cylindrical:
                       pEntity.SetPropertyValue("C" + " " + (pGridCylinderEntity.DistanceFromOrigin).ToString(strDecimalFormat) + " m", "IJNamedItem", "Name");
                       break;
               }          
            }
            catch (CmnException ex)
            {
                throw new Exception("GridsNamingRulesNetCS.PositionNameRule.ComputeName " + ex.Message);
            }
        }

        /// <summary>
        /// All the Naming Parents that need to participate in an objects naming are added here to the
        /// Collection(Of BusinessObject). The parents added here are used in computing the name of the object in
        /// ComputeName(). Both these methods are called from naming rule semantic.
        /// </summary>
        /// <param name="oEntity">Child object that needs to have the naming rule naming.</param>
        /// <returns> Collection of parents that participate in an objects naming.</returns>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
           Collection<BusinessObject> oParentsColl = new Collection<BusinessObject>();

           return oParentsColl;
        }
    }
}
