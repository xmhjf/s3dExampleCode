﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_USER_FIXED_CYL.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_USER_FIXED_CYL
//   Author       :  Vijay  
//   Creation Date:  30.10.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30.10.2012     Vijay    CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
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
    public class Utility_USER_FIXED_CYL : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_USER_FIXED_CYL"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(3, "RADIUS", "RADIUS", 0.999999)]
        public InputDouble m_dRADIUS;
        [InputString(4, "BOM_DESC", "BOM_DESC","No Value")]
        public InputString m_oBOM_DESC;
        [InputDouble(5, "DryWt", "DryWt", 0.999999)]
        public InputDouble m_dDryWt;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
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

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Double L = m_dL.Value;
                Double radius = m_dRADIUS.Value;
                Double drywt = m_dDryWt.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "StartOther", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "EndOther", new Position(0, 0, L), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidLength, "Length cannot be zero"));
                    return;
                }
                if (radius <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrRadiusGTZero, "Radius should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d top2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, radius, L);
                m_Symbolic.Outputs["BODY"] = top2;
               
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_USER_FIXED_CYL"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {

                Double weight, cogX, cogY, cogZ, dryWeight;
                if (supportComponentBO.SupportsInterface("IJOAHgrUtility_USER_FIXED_CYL"))
                {
                    try
                    {
                        dryWeight = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_USER_FIXED_CYL", "DryWt")).PropValue;
                    }
                    catch
                    {
                        dryWeight = 0;
                    }
                    weight = dryWeight;
                }
                else
                {
                    const int getSteelDensityKGPerM = 7900;
                    double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_USER_FIXED_CYL", "L")).PropValue;
                    double Radius = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_USER_FIXED_CYL", "RADIUS")).PropValue;

                    weight = (Math.PI * Radius * Radius * L) * getSteelDensityKGPerM;
                }
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_USER_FIXED_CYL"));
                }
            }
        }
        #endregion
    }
}
