//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Guide.cs
//    FLSample,Ingr.SP3D.Content.Support.Symbols.Guide
//   Author       :  Vijay
//   Creation Date:  13-Dec-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   13/12/2012     Vijay    CR222290 .Net FLSample Projected Creation
//	 20/03/2013		Vijay 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
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

    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class Guide : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FLSampleParts,Ingr.SP3D.Content.Support.Symbols.Guide"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(4, "Depth", "Depth", 0.999999)]
        public InputDouble m_dDepth;
        [InputDouble(5, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputString(6, "InputBomDesc", "InputBomDesc", "")]
        public InputString m_oInputBomDesc;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
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
                Part part = (Part)m_PartInput.Value;
                
                Double L = m_dL.Value;
                Double W = m_dW.Value;
                Double depth = m_dDepth.Value;
                Double T = m_dT.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, 0, 0), new Vector(1, 0, 0), new Vector(0, Math.Sqrt(2) / 2, Math.Sqrt(2) / 2));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(1, 0, 0), new Vector(0, Math.Sqrt(2) / 2, Math.Sqrt(2) / 2));
                m_Symbolic.Outputs["Port3"] = port3;

                if (L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrLArguments, "L value cannot be zero"));
                    return;
                }
                Collection<Position> pointCollection = new Collection<Position>();

                pointCollection.Add(new Position(0, 0, 0));
                pointCollection.Add(new Position(0, 0, depth));
                pointCollection.Add(new Position(0, -T, depth));
                pointCollection.Add(new Position(0, -T, T));
                pointCollection.Add(new Position(0, -W, T));
                pointCollection.Add(new Position(0, -W, 0));
                pointCollection.Add(new Position(0, 0, 0));

                Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(L, 0, 0), L, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrConstructOutputs, "Error in constructoutputs of Guide.cs"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                string inputBomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMMBomDesc", "InputBomDesc")).PropValue;
                bomDescription = inputBomDescription;
                return bomDescription;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Guide.cs"));
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
                const int getSteelDensityKGPerM = 7900;
                Double weight, cogX, cogY, cogZ;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGuide", "W")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGuide", "Depth")).PropValue;
                double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGuide", "T")).PropValue;
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGuide", "L")).PropValue;

                weight = ((L * (W - T) * T) + (L * depth * T)) * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrWeightCG, "Error in weightCG of Guide.cs"));
                }
            }
        }
        #endregion
    }
}

