//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_487.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_487
//   Author       :  Hema
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Hema    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net 
//   30-Dec-2014    PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report   
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
    public class PSL_487 : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_487"

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "SHOE_H", "SHOE_H", 0.999999)]
        public InputDouble SHOE_H;
        [InputDouble(3, "S", "S", 0.999999)]
        public InputDouble S;
        [InputDouble(4, "BEAM_REF", "BEAM_REF", 1)]
        public InputDouble BEAM_REF;
        [InputDouble(5, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(6, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputDouble(7, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(8, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(9, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(10, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(11, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(12, "G", "G", 0.999999)]
        public InputDouble G;
        [InputDouble(13, "H", "H", 0.999999)]
        public InputDouble H;
        [InputDouble(14, "T", "T", 0.999999)]
        public InputDouble T;
        [InputString(15, "LOAD_GROUP", "LOAD_GROUP", "No Value")]
        public InputString LOAD_GROUP;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BEAM1", "BEAM1")]
        [SymbolOutput("BEAM2", "BEAM2")]
        [SymbolOutput("LEFT1", "LEFT1")]
        [SymbolOutput("LEFT2", "LEFT2")]
        [SymbolOutput("RIGHT1", "RIGHT1")]
        [SymbolOutput("RIGHT2", "RIGHT2")]
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

                Double beamReference = BEAM_REF.Value;
                String loadGroup = LOAD_GROUP.Value;
                Double shoeHeight = SHOE_H.Value;
                Double s = S.Value;
                Double pipeDiameter = PIPE_DIA.Value;
                Double c = C.Value;
                Double a = A.Value;
                Double f = F.Value;
                Double b = B.Value;
                Double e = E.Value;
                Double g = G.Value;
                Double h = H.Value;
                Double t = T.Value;

                Double maxSpan, minSpan;
                string[] stringarray = new String[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "J" };

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                if (beamReference < 1 || beamReference > 11)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBeamReferenceCodelist, "BEAM_REF codelist values should be between 1 and 11"));
                    return;
                }

                PropertyValueCodelist beamCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrPSL_487", "BEAM_REF");
                String beamRef = beamCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)beamReference).ShortDisplayName.Trim();

                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_487_AUX", "BEAM_REF"), beamRef);

                double k = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_487_AUX", "IJUAHgrPSL_487_AUX", "K", parameter);

                parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_487_AUX", "LOAD_GROUP"), loadGroup);

                //Span error checking
                if (beamRef.Equals("A")) //check if range is between 0 and LA
                {
                    maxSpan = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_487_AUX", "IJUAHgrPSL_487_AUX", "L" + beamRef, parameter);

                    //error message
                    if (s < 0 || s > maxSpan)
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "Beam reference " + beamRef + " is not available for the selected Rod Center dimension(S). S must be less than " + Convert.ToString(maxSpan * 1000));
                }
                else
                {
                    //get MAX_SPAN
                    maxSpan = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_487_AUX", "IJUAHgrPSL_487_AUX", "L" + beamRef, parameter);

                    //error message
                    if (HgrCompareDoubleService.cmpdbl(maxSpan , 0)==true)
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "Beam reference " + beamRef + " is not available for the selected part");
                    else
                    {
                        //get the lower bound
                        int i = 1, minIndex = 0;
                        do
                        {
                            if (stringarray[i] == beamRef)
                                minIndex = i - 1;
                            i = i + 1;
                        } while (i < 10);

                        //error message
                        minSpan = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_487_AUX", "IJUAHgrPSL_487_AUX", "L" + stringarray[minIndex], parameter);

                        //error message
                        if (s <= minSpan || s > maxSpan)
                            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning,"Beam reference " + beamRef + " is not available for the selected Rod Center dimension(S). S must be between " + Convert.ToString(minSpan * 1000) + " and " + Convert.ToString(maxSpan * 1000));
                    }
                }
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, k + e), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                double down = shoeHeight;
                if (HgrCompareDoubleService.cmpdbl(shoeHeight , 0)==true)
                    down = pipeDiameter / 2.0;

                Port port2 = new Port(OccurrenceConnection, part, "Hole1", new Position(0, s / 2.0, down + k + e), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Hole2", new Position(0, -s / 2.0, down + k + e), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                //Validating Inputs
                if (b <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBGTZ, "B value should be greater than zero"));
                    return;
                }
                if (k <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidKGTZ, "K value should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(2 * f + s ,  0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFAndS, "(2F + S) value cannot be zero"));
                    return;
                }
                if (t <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidTGTZ, "T value should be greater than zero"));
                    return;
                }
                if (g <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidGGTZ, "G value should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(2.0 * h + a,0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAAndH, "(2H + A) value cannot be zero"));
                    return;
                }
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }
                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(-((2.0 * f) + s) / 2.0, -a / 2.0 - b / 2.0, down);
                Projection3d beam1 = symbolGeometryHelper.CreateBox(null, (2.0 * f) + s, b / 2.0, k, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                beam1.Transform(matrix);
                m_Symbolic.Outputs["BEAM1"] = beam1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-((2.0 * f) + s) / 2.0, a / 2.0, down);
                Projection3d beam2 = symbolGeometryHelper.CreateBox(null, (2.0 * f) + s, b / 2.0, k, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                beam2.Transform(matrix);
                m_Symbolic.Outputs["BEAM2"] = beam2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(s / 2.0 - ((2.0 * h) + a) / 2.0, -g / 2.0, down - t);
                Projection3d left1 = symbolGeometryHelper.CreateBox(null, 2.0 * h + a, g, t, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                left1.Transform(matrix);
                m_Symbolic.Outputs["LEFT1"] = left1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(s / 2.0 - c / 2.0, -b / 2.0, k + down);
                Projection3d left2 = symbolGeometryHelper.CreateBox(null, c, b, e, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                left2.Transform(matrix);
                m_Symbolic.Outputs["LEFT2"] = left2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-s / 2.0 - ((2.0 * h) + a) / 2.0, -g / 2.0, down - t);
                Projection3d right1 = symbolGeometryHelper.CreateBox(null, 2.0 * h + a, g, t, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                right1.Transform(matrix);
                m_Symbolic.Outputs["RIGHT1"] = right1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-s / 2.0 - c / 2.0, -b / 2.0, k + down);
                Projection3d right2 = symbolGeometryHelper.CreateBox(null, c, b, e, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                right2.Transform(matrix);
                m_Symbolic.Outputs["RIGHT2"] = right2;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_487."));
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

                int beamReferencevalue = (int)((PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrPSL_487", "BEAM_REF")).PropValue;
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                String beamRef = metadataManager.GetCodelistInfo("PSL_486_BEAM_REF", "UDP").GetCodelistItem(beamReferencevalue).ShortDisplayName;

                double unitWeight = PSLSymbolServices.GetDataByCondition("PSL_486_AUX", "IJUAHgrPSL_486_AUX", "UNIT_WEIGHT", "IJUAHgrPSL_486_AUX", "BEAM_REF", beamRef);
                double weightExclBeam = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_487", "WEIGHT_EXCL_BEAM")).PropValue;
                double s = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrPSL_487", "S")).PropValue;
                double f = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_487", "F")).PropValue;

                double weight = 2 * (s + 2 * f) * unitWeight + weightExclBeam;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, 0, 0, 0);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_487.cs."));
            }
        }
        #endregion
    }
}
