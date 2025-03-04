//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GEN_REC_STRAP.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_REC_STRAP
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
    public class Utility_GEN_REC_STRAP : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_REC_STRAP"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(3, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(5, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(6, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
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

                Double A = m_dA.Value;
                Double D = m_dD.Value;
                Double B = m_dB.Value;
                Double T = m_dT.Value;
                Double pipe_Dia = m_dPIPE_DIA.Value;

                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrThicknessGTZero, "Thickness T should be greater than zero"));
                    return;
                }
                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWidthAGTZero, "Width A should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrHeightBGTZero, "Height B should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrDGTZero, "Depth D should be greater than zero"));
                    return;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                symbolGeometryHelper.ActivePosition = new Position(0, 0, (B - pipe_Dia / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(0, 1, 0));
                Projection3d top = (Projection3d)symbolGeometryHelper.CreateBox(null, T, A + 2 * T, D);
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, A / 2 + T / 2, (B - pipe_Dia / 2) - T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(0, 1, 0));
                Projection3d left = (Projection3d)symbolGeometryHelper.CreateBox(null, B, T, D);
                m_Symbolic.Outputs["LEFT"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -A / 2 - T / 2, (B - pipe_Dia / 2) - T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(0, 1, 0));
                Projection3d right = (Projection3d)symbolGeometryHelper.CreateBox(null, B, T, D);
                m_Symbolic.Outputs["RIGHT"] = right;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GEN_REC_STRAP"));
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

                Double weight, cogX, cogY, cogZ;
                const int getSteelDensityKGPerM = 7900;
                double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_REC_STRAP", "T")).PropValue;
                double A = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_REC_STRAP", "A")).PropValue;
                double B = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_REC_STRAP", "B")).PropValue;
                double D = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_REC_STRAP", "D")).PropValue;

                weight = (A * D * T) + (2 * (T * B * D)) * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GEN_REC_STRAP"));
                }
            }
        }
        #endregion
    }
}
