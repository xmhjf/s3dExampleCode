//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_NOTCH_PLATE.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_NOTCH_PLATE
//   Author       :  Hema
//   Creation Date:  04.11.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   04.11.2012      Hema     CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		 Hema 	  DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Utility_NOTCH_PLATE : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_NOTCH_PLATE"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "THICKNESS", "THICKNESS", 0.999999)]
        public InputDouble m_dTHICKNESS;
        [InputDouble(3, "WIDTH", "WIDTH", 0.999999)]
        public InputDouble m_dWIDTH;
        [InputDouble(4, "DEPTH", "DEPTH", 0.999999)]
        public InputDouble m_dDEPTH;
        [InputDouble(5, "X", "X", 0.999999)]
        public InputDouble m_dX;
        [InputDouble(6, "Z", "Z", 0.999999)]
        public InputDouble m_dZ;
        [InputString(7, "BOM_DESC1", "BOM_DESC1", "No Value")]
        public InputString m_oBOM_DESC1;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("BODY", "BODY")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part =(Part)m_PartInput.Value ;

                Double thickness = m_dTHICKNESS.Value;
                Double width = m_dWIDTH.Value;
                Double depth = m_dDEPTH.Value;
                Double x = m_dX.Value;
                Double z = m_dZ.Value;

                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }
                if (depth == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidDepth, "Depth cannot be zero"));
                    return;
                }
                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-width / 2 + x, -depth / 2 + z, 0));
                pointCollection.Add(new Position(-width / 2, -depth / 2 + z, 0));
                pointCollection.Add(new Position(-width / 2, depth / 2, 0));
                pointCollection.Add(new Position(width / 2, depth / 2, 0));
                pointCollection.Add(new Position(width / 2, -depth / 2, 0));
                pointCollection.Add(new Position(-width / 2 + x, -depth / 2, 0));
                pointCollection.Add(new Position(-width / 2 + x, -depth / 2 + z, 0));

                Vector projectionVector = new Vector(0, 0, thickness);
                Projection3d body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_NOTCH_PLATE"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double z = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_NOTCH_PLATE", "Z")).PropValue;
                double x = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_NOTCH_PLATE", "X")).PropValue;
                double width = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_NOTCH_PLATE", "WIDTH")).PropValue;
                double depth = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_NOTCH_PLATE", "DEPTH")).PropValue;
                string bomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_NOTCH_PLATE", "BOM_DESC1")).PropValue;
                double thickness = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_NOTCH_PLATE", "THICKNESS")).PropValue;

                if (bomDescription == null)
                    bomString = Microsoft.VisualBasic.Conversion.Str(thickness) + " Plate Steel, " + Microsoft.VisualBasic.Conversion.Str(width) + " X " + Microsoft.VisualBasic.Conversion.Str(depth) + ", Notch " + Microsoft.VisualBasic.Conversion.Str(x) + " X " + Microsoft.VisualBasic.Conversion.Str(z);
                else
                    bomString = bomDescription;
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_NOTCH_PLATE"));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {

                Double weight, cogX, cogY, cogZ;
                const int getSteelDensityKGPerM = 7900;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double z = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_NOTCH_PLATE", "Z")).PropValue;
                double x = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_NOTCH_PLATE", "X")).PropValue;
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_NOTCH_PLATE", "WIDTH")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_NOTCH_PLATE", "DEPTH")).PropValue;
                double thickness = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_NOTCH_PLATE", "THICKNESS")).PropValue;

                weight = ((depth * width * thickness) - (x * z * thickness)) * getSteelDensityKGPerM;
                cogX = 0;
                cogY = 0;
                cogZ = 0;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_NOTCH_PLATE"));
                }
            }
        }
        #endregion
    }
}
