//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GEN_4_BOLT_CLAMP.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_4_BOLT_CLAMP
//   Author       :  Hema
//   Creation Date:  31.10.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   31.10.2012      Hema    CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		 Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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
    public class Utility_GEN_4_BOLT_CLAMP : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_4_BOLT_CLAMP"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(4, "K", "K", 0.999999)]
        public InputDouble m_dK;
        [InputDouble(5, "J", "J", 0.999999)]
        public InputDouble m_dJ;
        [InputDouble(6, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(7, "BOLT_DIA", "BOLT_DIA", 0.999999)]
        public InputDouble m_dBOLT_DIA;
        [InputDouble(8, "BOLT_L", "BOLT_L", 0.999999)]
        public InputDouble m_dBOLT_L;
        [InputDouble(9, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(10, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(11, "OPT_BOM", "OPT_BOM", 1)]
        public InputDouble m_oOPT_BOM;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("BOT1", "BOT1")]
        [SymbolOutput("BOT2", "BOT2")]
        [SymbolOutput("TOP_BOLT1", "TOP_BOLT1")]
        [SymbolOutput("TOP_BOLT2", "TOP_BOLT2")]
        [SymbolOutput("BOT_BOLT1", "BOT_BOLT1")]
        [SymbolOutput("BOT_BOLT2", "BOT_BOLT2")]
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

                Double R = m_dR.Value;
                Double A = m_dA.Value;
                Double K = m_dK.Value;
                Double J = m_dJ.Value;
                Double H = m_dH.Value;
                Double bolt_Dia = m_dBOLT_DIA.Value;
                Double bolt_L = m_dBOLT_L.Value;
                Double T = m_dT.Value;
                Double W = m_dW.Value;

                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrStockTGTZero, "Stock T should be greater than zero"));
                    return;
                }
                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidOverallWidth, "Overall Width should be greater than zero"));
                    return;
                }
                if (R <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrRGTZero, "Radius should be greater than zero"));
                    return;
                }
                if (bolt_Dia <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidBoltDia, "Bolt Dia should be greater than zero"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidStockWidth, "Stock Width should be greater than zero"));
                    return;
                }
                if (bolt_L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidBOlTLENGTH, "Bolt Length cannot be zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(-J / 2, 0, K / 2 - bolt_Dia / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Pin", new Position(J / 2, 0, K / 2 - bolt_Dia / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                symbolGeometryHelper.ActivePosition = new Position(0, (H / 2 + T), -(R + (A / 2 - R) / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(0, 0, 1));
                Projection3d top1 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, A / 2 - R, W);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(H / 2 + T), -(R + (A / 2 - R) / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d top2 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, A / 2 - R, W);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, (H / 2 + T), (R + (A / 2 - R) / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(0, 0, 1));
                Projection3d bottom1 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, A / 2 - R, W);
                m_Symbolic.Outputs["BOT1"] = bottom1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(H / 2 + T), (R + (A / 2 - R) / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d bottom2 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, A / 2 - R, W);
                m_Symbolic.Outputs["BOT2"] = bottom2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(-J / 2, -bolt_L / 2, K / 2).Subtract(new Position(-J / 2, bolt_L / 2, K / 2));
                symbolGeometryHelper.ActivePosition = new Position(-J / 2, -bolt_L / 2, K / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d top_bolt1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, bolt_Dia / 2, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT1"] = top_bolt1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position(J / 2, -bolt_L / 2, K / 2).Subtract(new Position(J / 2, bolt_L / 2, K / 2));
                symbolGeometryHelper.ActivePosition = new Position(J / 2, -bolt_L / 2, K / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d top_bolt2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, bolt_Dia / 2, normal1.Length);
                m_Symbolic.Outputs["TOP_BOLT2"] = top_bolt2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal2 = new Position(-J / 2, -bolt_L / 2, -K / 2).Subtract(new Position(-J / 2, bolt_L / 2, -K / 2));
                symbolGeometryHelper.ActivePosition = new Position(-J / 2, -bolt_L / 2, -K / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d bottom_bolt1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, bolt_Dia / 2, normal2.Length);
                m_Symbolic.Outputs["BOT_BOLT1"] = bottom_bolt1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal3 = new Position(J / 2, -bolt_L / 2, -K / 2).Subtract(new Position(J / 2, bolt_L / 2, -K / 2));
                symbolGeometryHelper.ActivePosition = new Position(J / 2, -bolt_L / 2, -K / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d bottom_bolt2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, bolt_Dia / 2, normal3.Length);
                m_Symbolic.Outputs["BOT_BOLT2"] = bottom_bolt2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-W / 2, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Projection3d body = (Projection3d)symbolGeometryHelper.CreateCylinder(null, R + T, W);
                m_Symbolic.Outputs["BODY"] = body;

            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GEN_4_BOLT_CLAMP"));
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
                double R = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "R")).PropValue;
                double A = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "A")).PropValue;
                double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "H")).PropValue;
                double boltDiameter = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "BOLT_DIA")).PropValue;
                double boltL = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "BOLT_L")).PropValue;
                double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "T")).PropValue;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "W")).PropValue;

                weight = (((R + T) * (R + T) * Math.PI * W) - (R * R * Math.PI * W) - ((T * W * H) * 2) + (((A / 2 - R - T) * W * T) * 4) + ((Math.PI * boltL * (boltDiameter / 2) * (boltDiameter / 2)) * 4)) * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GEN_4_BOLT_CLAMP"));
                }
            }
        }
    }
        #endregion
}
