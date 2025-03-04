//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ProtectSaddle.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ProtectSaddle
//   Author       :  
//   Creation Date:  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-02-2013     Pradeep   CR-222486 Converted VB Project to .Net
//   25/Mar/2013    Rajeswari DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   30/10/2013     Hema      CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//   12-12-2014     PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
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
    [VariableOutputs]
    public class ProtectSaddle : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ProtectSaddle"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        public AspectDefinition m_oSymbolic;
        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddProtectSaddleInputs(2, out endIndex, additionalInputs);

                return additionalInputs;
            }
        }
        #endregion
        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            AddProtectSaddleOutputs(aspectName, additionalOutputs);
            AddPipeClampOutputs(aspectName, additionalOutputs, "clamp1");
            AddPipeClampOutputs(aspectName, additionalOutputs, "clamp2");
            return additionalOutputs;
        }
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
                ProtectSaddleInputs protectSaddle = LoadProtectSaddle(2);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                PipeClampInputs pipeClamp = new PipeClampInputs(); ;
                if (protectSaddle.ClampName != "No Value" || protectSaddle.ClampName == "")
                {
                    pipeClamp = LoadPipeClampDataByQuery(protectSaddle.ClampName);
                    protectSaddle.PipeOD = protectSaddle.PipeOD + (pipeClamp.Thickness1 + pipeClamp.Thickness1);
                }

                //protectsaddleshape
                Matrix4X4 matrix = new Matrix4X4();
                AddProtectSaddle(protectSaddle, matrix, m_oSymbolic.Outputs, "ProtectSaddle");
                if (protectSaddle.ClampName != "No Value" || protectSaddle.ClampName == "")
                {
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(matrix.Transform(new Vector(protectSaddle.Length1 / 2 - pipeClamp.Width1 / 2, 0, 0)));
                    AddPipeClamp(pipeClamp, matrix, m_oSymbolic.Outputs, "clamp1");
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(matrix.Transform(new Vector(-protectSaddle.Length1 / 2 + pipeClamp.Width1 / 2, 0, 0)));
                    AddPipeClamp(pipeClamp, matrix, m_oSymbolic.Outputs, "clamp2");
                }
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_oSymbolic.Outputs["Port1"] = port1;

                if (!(protectSaddle.Height1 > 0) || (protectSaddle.Height3 > 0))
                {
                    double RightFootZ;
                    double RightFootY;
                    double Beta;

                    if (protectSaddle.ToEdge == 1)
                    {
                        Beta = ((protectSaddle.Thickness1 / 2) / (protectSaddle.PipeOD / 2));
                        RightFootY = Math.Sin(protectSaddle.Angle2 - Beta - Beta) * (protectSaddle.PipeOD / 2);
                        RightFootZ = Math.Cos(protectSaddle.Angle2 - Beta - Beta) * (protectSaddle.PipeOD / 2);
                    }
                    else
                    {
                        Beta = ((protectSaddle.Thickness1 / 2) / (protectSaddle.PipeOD / 2));
                        RightFootY = Math.Sin(protectSaddle.Angle2 - Beta) * (protectSaddle.PipeOD / 2);
                        RightFootZ = Math.Cos(protectSaddle.Angle2 - Beta) * (protectSaddle.PipeOD / 2);
                    }
                    if (protectSaddle.Height3 > 0)
                    {
                        RightFootZ = Math.Cos(Math.Asin(protectSaddle.Width2 / (protectSaddle.PipeOD / 2))) * (protectSaddle.PipeOD / 2);
                        RightFootY = protectSaddle.Width2;
                    }

                    Port port2 = new Port(OccurrenceConnection, part, "Weld1", new Position(-protectSaddle.Length1 / 2, -RightFootY, -RightFootZ), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_oSymbolic.Outputs["Port2"] = port2;
                }
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of ProtectSaddle"));
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
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = 0;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of ProtectSaddle"));
                }
            }

        }
        #endregion
    }
}
