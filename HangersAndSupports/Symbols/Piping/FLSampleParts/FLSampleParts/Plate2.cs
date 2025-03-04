//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Plate2.cs
//    FLSample,Ingr.SP3D.Content.Support.Symbols.Plate2
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
    public class Plate2 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FLSampleParts,Ingr.SP3D.Content.Support.Symbols.Plate2"
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
        [InputString(5, "InputBomDesc", "InputBomDesc", "")]
        public InputString m_oInputBomDesc;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("LINE", "LINE")]
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
                Part part = (Part)m_PartInput.Value;
                Double thickness = m_dTHICKNESS.Value;
                Double width = m_dWIDTH.Value;
                Double depth = m_dDEPTH.Value;

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "TopStructure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "BotStructure", new Position(0, 0, thickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrWidthArguments, "Width should be greater than zero"));
                    return;
                }
                if (depth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrDepthArguments, "Depth should be greater than zero"));
                    return;
                }
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrThicknessArguments, "THICKNESS should be greater than zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                Projection3d body = (Projection3d)symbolGeometryHelper.CreateBox(null, thickness, width, depth);
                m_Symbolic.Outputs["BODY"] = body;

                Line3d line = new Line3d(new Position(-depth / 2, 0, thickness / 2), new Position(depth / 2, 0, thickness / 2));
                m_Symbolic.Outputs["LINE"] = line;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrConstructOutputs, "Error in constructoutputs of Plate2.cs"));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Plate2.cs"));
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
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_PLATE", "WIDTH")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_PLATE", "DEPTH")).PropValue;
                double thickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_PLATE", "THICKNESS")).PropValue;

                weight = depth * width * thickness * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrWeightCG, "Error in weightCG of Plate2.cs"));
                }
            }
        }
        #endregion
    }
}


