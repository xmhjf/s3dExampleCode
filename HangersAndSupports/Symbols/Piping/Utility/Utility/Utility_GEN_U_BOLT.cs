﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GEN_U_BOLT.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_U_BOLT
//   Author       :  Hema
//   Creation Date:  01.11.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01.11.2012      Hema    CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		 Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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
    public class Utility_GEN_U_BOLT : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_U_BOLT"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(3, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(4, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputString(5, "BOM_DESC1", "BOM_DESC1", "No Value")]
        public InputString m_oBOM_DESC1;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("BEND", "BEND")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("LEFT", "LEFT")]
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

                Double D = m_dD.Value;
                Double H = m_dH.Value;
                Double R = m_dR.Value;

                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrRodDiameterGTZero, "Rod Diameter should be greater than zero"));
                    return;
                }
                if (H == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidHeight, "Height cannot be zero"));
                    return;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Revolution3d elbow = new Revolution3d(OccurrenceConnection, (new Circle3d(new Position(0, R, 0), new Vector(0, 0, 1), D / 2)), new Vector(1, 0, 0), new Position(0, 0, 0), Math.PI, true);
                m_Symbolic.Outputs["BEND"] = elbow;

                symbolGeometryHelper.ActivePosition = new Position(0, R, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(0, 1, 0));
                Projection3d right = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2, H);
                m_Symbolic.Outputs["RIGHT"] = right;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -R, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(0, 1, 0));
                Projection3d left = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2, H);
                m_Symbolic.Outputs["LEFT"] = left;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GEN_U_BOLT"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string BOMString = "";
            try 
            {
                double D = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_BOLT", "D")).PropValue;
                double H = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_BOLT", "H")).PropValue;
                double R = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_BOLT", "R")).PropValue;
                double nuts = (double)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_BOLT", "NUTS")).PropValue;
                double thread = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_BOLT", "THREAD")).PropValue;
                string bomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_BOLT", "BOM_DESC1")).PropValue;

                if (bomDescription == null)
                    BOMString = "Custom U-Bolt, H=" + Microsoft.VisualBasic.Conversion.Str(H) + " D=" + Microsoft.VisualBasic.Conversion.Str(D) + " R=" + Microsoft.VisualBasic.Conversion.Str(R) + " Nuts=" + Microsoft.VisualBasic.Conversion.Str(nuts) + " Thread=" + Microsoft.VisualBasic.Conversion.Str(thread);
                else
                    BOMString = bomDescription;

                return BOMString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_GEN_U_BOLT"));
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
                double D = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_U_BOLT", "D")).PropValue;
                double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_U_BOLT", "H")).PropValue;
                double R = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_U_BOLT", "R")).PropValue;
                weight = ((Math.PI * D / 2 * D / 2 * H) * 2 + (Math.PI * D / 2 * D / 2 * (Math.PI * R))) * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GEN_U_BOLT"));
                }
            }
        }
        #endregion
    }
}
