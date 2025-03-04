//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Cylinder.cs
//    FLSample,Ingr.SP3D.Content.Support.Symbols.Cylinder
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
    public class Cylinder : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FLSampleParts,Ingr.SP3D.Content.Support.Symbols.Cylinder"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(3, "RADIUS", "RADIUS", 0.999999)]
        public InputDouble m_dRADIUS;
        [InputString(4, "InputBomDesc", "InputBomDesc", "")]
        public InputString m_oInputBomDesc;
        [InputDouble(5, "InputWeight", "InputWeight", 0.999999)]
        public InputDouble m_dInputWeight;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("TOP_POINTPT0", "TOP_POINTPT0")]
        [SymbolOutput("BOTTOM_POINTPT0", "BOTTOM_POINTPT0")]
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
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Double L = m_dL.Value;
                Double radius = m_dRADIUS.Value;

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "StartOther", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "EndOther", new Position(0, 0, L), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                if (radius <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrRadiusArguments, "RADIUS should be greater than zero"));
                    return;
                }
                if (L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrLArguments, "L value cannot be zero"));
                    return;
                }
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d top1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, radius, L);
                m_Symbolic.Outputs["BODY"] = top1;


                Point3d point = new Point3d(new Position(0, 0, L));
                m_Symbolic.Outputs["TOP_POINTPT0"] = point;


                Point3d point1 = new Point3d(new Position(0, 0, 0));
                m_Symbolic.Outputs["BOTTOM_POINTPT0"] = point1;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrConstructOutputs, "Error in constructoutputs of cylinder.cs"));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of cylinder.cs"));
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
                Double inputWeight = (Double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrFlSampleWeight", "InputWeight")).PropValue;
                weight = inputWeight;

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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FLSampleLocalizer.GetString(FLSampleSymbolResourceIDs.ErrWeightCG, "Error in weightCG of cylinder.cs"));
                }
            }
        }
        #endregion
    }
}


