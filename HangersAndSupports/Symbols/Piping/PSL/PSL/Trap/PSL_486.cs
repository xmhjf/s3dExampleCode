//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_486.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_486
//   Author       :  Hema
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Hema    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.Generic;

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
    public class PSL_486 : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_486"

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "SHOE_H", "SHOE_H", 0.999999)]
        public InputDouble SHOE_H;
        [InputDouble(3, "S", "S", 0.999999)]
        public InputDouble S;
        [InputDouble(4, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(5, "BEAM_REF", "BEAM_REF", 1)]
        public InputDouble BEAM_REF;
        [InputString(6, "LOAD_GROUP", "LOAD_GROUP", "No Value")]
        public InputString LOAD_GROUP;
        [InputDouble(7, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputDouble(8, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(9, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(10, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(11, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(12, "T", "T", 0.999999)]
        public InputDouble T;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BEAM", "BEAM")]
        [SymbolOutput("LEFT1", "LEFT1")]
        [SymbolOutput("LEFT2", "LEFT2")]
        [SymbolOutput("RIGHT1", "RIGHT1")]
        [SymbolOutput("RIGHT2", "RIGHT2")]
        [SymbolOutput("LEFT_BOLT", "LEFT_BOLT")]
        [SymbolOutput("RIGHT_BOLT", "RIGHT_BOLT")]
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
                //-----------------------------------------------------------------------
                // Construction of Physical Aspect 
                //----------------------------------------------------------------------
                Part part = (Part)PartInput.Value;

                Double shoeHeight = SHOE_H.Value;
                Double s = S.Value;
                Double pipeDiameter = PIPE_DIA.Value;
                Double beamReference = BEAM_REF.Value;
                String loadGroup = LOAD_GROUP.Value;
                Double b = B.Value;
                Double c = C.Value;
                Double d = D.Value;
                Double e = E.Value;
                Double t = T.Value;
                Double maxSpan, minSpan;
                string[] beamarray = new String[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L" };

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                if (beamReference < 1 || beamReference > 11)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBeamReferenceCodelist, "BEAM_REF codelist values should be between 1 and 11"));
                    return;
                }
                PropertyValueCodelist beamCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrPSL_486", "BEAM_REF");
                String beamRef = beamCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)beamReference).ShortDisplayName.Trim();

                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_486_AUX", "BEAM_REF"), beamRef);

                double a = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_486_AUX", "IJUAHgrPSL_486_AUX", "A", parameter);

                parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_486_AUX", "LOAD_GROUP"), loadGroup);
                //Span error checking
                if (beamRef.Equals("A")) //check if range is between 0 and LA
                {
                    maxSpan = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_486_AUX", "IJUAHgrPSL_486_AUX", "L" + beamRef, parameter);
                    //error message
                    if (s < 0 || s > maxSpan)
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "Beam reference " + beamRef + " is not available for the selected rod center dimension(S). S must be less than " + Convert.ToString(maxSpan * 1000));
                }
                else
                {
                    //get MAX_SPAN
                    maxSpan = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_486_AUX", "IJUAHgrPSL_486_AUX", "L" + beamRef, parameter);
                    //error message
                    if (maxSpan == 0)
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "Beam reference " + beamRef + " is not available for the selected part");
                    else
                    {
                        //get the lower bound
                        int i = 1, minIndex = 0;
                        do
                        {
                            if (beamarray[i] == beamRef)
                                minIndex = i - 1;
                            i = i + 1;
                        } while (i < 12);
                        //error message
                        minSpan = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_486_AUX", "IJUAHgrPSL_486_AUX", "L" + beamarray[minIndex], parameter);
                        //error message
                        if (s <= minSpan || s > maxSpan)
                            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "Beam reference " + beamRef + " is not available for the selected rod center dimension(S). S must be between " + Convert.ToString(minSpan * 1000) + " and " + Convert.ToString(maxSpan * 1000));
                    }
                }
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, a), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                double down = shoeHeight;
                if (shoeHeight == 0)
                    down = pipeDiameter / 2.0;

                Port port2 = new Port(OccurrenceConnection, part, "Pin1", new Position(0, s / 2.0, a + c - down), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Pin2", new Position(0, -s / 2.0, a + c - down), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                //Validating Inputs
                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }
                if (b <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBGTZ, "B value should be greater than zero"));
                    return;
                }
                if (t <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidTGTZ, "T value should be greater than zero"));
                    return;
                }
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }
                if (d <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDGTZ, "D value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(-((2.0 * e) + s) / 2.0, -a / 2.0, -down);
                Projection3d beam = symbolGeometryHelper.CreateBox(null, (2.0 * e) + s, a, a, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                beam.Transform(matrix);
                m_Symbolic.Outputs["BEAM"] = beam;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(s / 2.0 + b / 2.0, -(1.25 * b) / 2.0, a - down);
                Projection3d left1 = symbolGeometryHelper.CreateBox(null, t, b * 1.25, 2 * c, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                left1.Transform(matrix);
                m_Symbolic.Outputs["LEFT1"] = left1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(s / 2.0 - b / 2.0 - t, -(1.25 * b) / 2.0, a - down);
                Projection3d left2 = symbolGeometryHelper.CreateBox(null, t, b * 1.25, 2 * c, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                left2.Transform(matrix);
                m_Symbolic.Outputs["LEFT2"] = left2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-s / 2.0 - b / 2.0 - t, -(1.25 * b) / 2.0, a - down);
                Projection3d right1 = symbolGeometryHelper.CreateBox(null, t, b * 1.25, 2 * c, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                right1.Transform(matrix);
                m_Symbolic.Outputs["RIGHT1"] = right1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-s / 2.0 + b / 2.0, -(1.25 * b) / 2.0, a - down);
                Projection3d right2 = symbolGeometryHelper.CreateBox(null, t, b * 1.25, 2 * c, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                right2.Transform(matrix);
                m_Symbolic.Outputs["RIGHT2"] = right2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, s / 2.0 - b, a - down + c).Subtract(new Position(0, s / 2.0 + b, a - down + c));
                symbolGeometryHelper.ActivePosition = new Position(0, s / 2.0 + b, a - down + c);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d leftBolt = symbolGeometryHelper.CreateCylinder(null, d / 2.0, normal.Length);
                m_Symbolic.Outputs["LEFT_BOLT"] = leftBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, -s / 2.0 + b, a - down + c).Subtract(new Position(0, -s / 2.0 - b, a - down + c));
                symbolGeometryHelper.ActivePosition = new Position(0, -s / 2.0 - b, a - down + c);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rightBolt = symbolGeometryHelper.CreateCylinder(null, d / 2.0, normal.Length);
                m_Symbolic.Outputs["RIGHT_BOLT"] = rightBolt;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_486."));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                int beamRefvalue = (int)((PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrPSL_486", "BEAM_REF")).PropValue;
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                string beamRef = metadataManager.GetCodelistInfo("PSL_486_BEAM_REF", "UDP").GetCodelistItem(beamRefvalue).ShortDisplayName;

                double unitWeight = PSLSymbolServices.GetDataByCondition("PSL_486_AUX", "IJUAHgrPSL_486_AUX", "UNIT_WEIGHT", "IJUAHgrPSL_486_AUX", "BEAM_REF", beamRef);
                double weightExclBeam = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_486", "WEIGHT_EXCL_BEAM")).PropValue;
                double s = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrPSL_486", "S")).PropValue;
                double e = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_486", "E")).PropValue;

                double weight = (s + 2 * e) * unitWeight + weightExclBeam;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, 0, 0, 0);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_486.cs."));
            }
        }
        #endregion
    }
}
