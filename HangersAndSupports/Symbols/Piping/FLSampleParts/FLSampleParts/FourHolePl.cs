//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   FourHolePl.cs
//    FLSample,Ingr.SP3D.Content.Support.Symbols.FourHolePl
//   Author       :  Vijay
//   Creation Date:  12-Dec-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12/12/2012     Vijay    CR222290 .Net FLSample Projected Creation
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
    public class FourHolePl : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FLSampleParts,Ingr.SP3D.Content.Support.Symbols.FourHolePl"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(3, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(4, "Depth", "Depth", 0.999999)]
        public InputDouble m_dDepth;
        [InputString(5, "InputBomDesc", "InputBomDesc", "")]
        public InputString m_oInputBomDesc;
        [InputDouble(6, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(7, "HoleSize", "HoleSize", 0.999999)]
        public InputDouble m_dHoleSize;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("LINE", "LINE")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("R_F_HOLE", "R_F_HOLE")]
        [SymbolOutput("R_B_HOLE", "R_B_HOLE")]
        [SymbolOutput("L_F_HOLE", "L_F_HOLE")]
        [SymbolOutput("L_B_HOLE", "L_B_HOLE")]
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
                Double T = m_dT.Value;
                Double width = m_dWidth.Value;
                Double depth = m_dDepth.Value;
                Double C = m_dC.Value;
                Double holeSize = m_dHoleSize.Value;


                //ports
                Port port1 = new Port(OccurrenceConnection, part, "TopStructure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "BotStructure", new Position(0, 0, T), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
             
                if (holeSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrHoleSizeArguments, "HoleSize should be greater than zero"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrTArguments, "T value should be greater than zero"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrWidthArguments, "Width  should be greater than zero"));
                    return;
                }
                if (depth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrDepthArguments, "Depth should be greater than zero"));
                    return;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                Projection3d body = (Projection3d)symbolGeometryHelper.CreateBox(null, T, width, depth);
                m_Symbolic.Outputs["BODY"] = body;

                Line3d line = new Line3d(new Position(-depth / 2.0, 0, T / 2.0), new Position(depth / 2.0, 0, T / 2.0));
                m_Symbolic.Outputs["LINE"] = line;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2 - C), width / 2 - C, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rightfrontHole = symbolGeometryHelper.CreateCylinder(null, holeSize / 2.0, T + 0.00002);
                m_Symbolic.Outputs["R_F_HOLE"] = rightfrontHole;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position((depth / 2 - C), width / 2 - C, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rightbackHole = symbolGeometryHelper.CreateCylinder(null, holeSize / 2.0, T + 0.00002);
                m_Symbolic.Outputs["R_B_HOLE"] = rightbackHole;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2 - C), -(width / 2 - C), 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d leftfrontHole = symbolGeometryHelper.CreateCylinder(null, holeSize / 2.0, T + 0.00002);
                m_Symbolic.Outputs["L_F_HOLE"] = leftfrontHole;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position((depth / 2 - C), -(width / 2 - C), 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d leftbackHole = symbolGeometryHelper.CreateCylinder(null, holeSize / 2.0, T + 0.00002);
                m_Symbolic.Outputs["L_B_HOLE"] = leftbackHole;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrConstructOutputs, "Error in constructoutputs of FourHolePl.cs"));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of FourHolePl.cs"));
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
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_FourHolePl", "Width")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_FourHolePl", "Depth")).PropValue;
                double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_FourHolePl", "T")).PropValue;

                weight = depth * width * T * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrWeightCG, "Error in weightCG of FourHolePl.cs"));
                }
            }
        }
        #endregion
    }
}

