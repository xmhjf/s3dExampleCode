//----------------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  PlateKnuckleRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\ShipStructure\Rules\Release\PlateKnuckleRule.dll
//  Original Class Name: ‘BendOrSplitRule’ in VB content
//
//Abstract:
//  PlateKnuckleRule is a .NET PlateKnuckle rule which decides the PlateKnuckle’s treatment type and inner radius.
//  This class subclasses from PlateKnuckleRuleBase.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//----------------------------------------------------------------------------------------------------------------
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// PlateKnuckle rule to decide upon the PlateKnuckle’s treatment type and inner radius.
    /// </summary>
    public class PlateKnuckleRule : PlateKnuckleRuleBase
    {
        //==================================================================================================
        //DefinitionName/ProgID of this rule is "KnuckleRules,Ingr.SP3D.Content.Structure.PlateKnuckleRule"
        //==================================================================================================
        #region Public override methods

        /// <summary>
        /// Gets the PlateKnuckle inner radius (meters) based on the thickness of the PlateSystem.
        /// </summary>
        /// <param name="plateSystem">The PlateKnuckle’s parent PlateSystem.</param>
        /// <returns>The PlateKnuckle inner radius.</returns>
        public override double GetInnerRadius(PlateSystem plateSystem)
        {
            double innerRadius;

            //get the PlateKnuckle’s parent PlateSystem thickness  
            double plateSystemThickness = plateSystem.Thickness;

            //if the thickness is less than 10mm then the inner radius is 50mm
            //else if the thickness is less than or equal to 20mm ( 10mm < thickness <= 20mm) then the inner radius is 100mm
            //else if the thickness is less than 30mm ( 20mm < thickness < 30mm) then the inner radius is 150mm
            //else (thickness > 30mm) then the inner radius is 200mm
            if (plateSystemThickness < 0.01)
            {
                innerRadius = 0.05;
            }
            else if (plateSystemThickness < 0.02 || StructHelper.AreEqual(plateSystemThickness, 0.02))
            {
                innerRadius = 0.1;
            }
            else if (plateSystemThickness < 0.03)
            {
                innerRadius = 0.15;
            }
            else
            {
                innerRadius = 0.2;
            }

            return innerRadius;
        }

        #endregion Public override methods
    }
}