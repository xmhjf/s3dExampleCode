﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_PB4_CM.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_PB4_CM
//   Author       :  Vijaya
//   Creation Date:  23-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-08-2013     Vijaya    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    public class PSL_PB4_CM : HangerComponentSymbolDefinition
    {

        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_PB4_CM"

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "COMLIN_T", "COMLIN_T", 0.999999)]
        public InputDouble COMLIN_T;
        [InputString(3, "CLAMP_SIZE", "CLAMP_SIZE", "No Value")]
        public InputString CLAMP_SIZE;
        [InputDouble(4, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(5, "END_T", "END_T", 0.999999)]
        public InputDouble END_T;
        [InputDouble(6, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(7, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(8, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(10, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(11, "WEB_T", "WEB_T", 0.999999)]
        public InputDouble WEB_T;
        [InputDouble(12, "FLANGE_T", "FLANGE_T", 0.999999)]
        public InputDouble FLANGE_T;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("BASEPLATE", "BASEPLATE")]
        [SymbolOutput("WEB", "WEB")]
        [SymbolOutput("EP1", "EP1")]
        [SymbolOutput("EP2", "EP2")]
        [SymbolOutput("BASEPLATE2", "BASEPLATE2")]
        [SymbolOutput("WEB2", "WEB2")]
        [SymbolOutput("EP3", "EP3")]
        [SymbolOutput("EP4", "EP4")]
        [SymbolOutput("BASEPLATE3", "BASEPLATE3")]
        [SymbolOutput("WEB3", "WEB3")]
        [SymbolOutput("EP5", "EP5")]
        [SymbolOutput("EP6", "EP6")]
        [SymbolOutput("BASEPLATE4", "BASEPLATE4")]
        [SymbolOutput("WEB4", "WEB4")]
        [SymbolOutput("EP7", "EP7")]
        [SymbolOutput("EP8", "EP8")]
        [SymbolOutput("CYL", "CYL")]
        [SymbolOutput("CYL2", "CYL2")]
        [SymbolOutput("COMLIN1", "COMLIN1")]
        [SymbolOutput("COMLIN2", "COMLIN2")]
        [SymbolOutput("BOX1", "BOX1")]
        [SymbolOutput("BOX2", "BOX2")]
        [SymbolOutput("BOX3", "BOX3")]
        [SymbolOutput("BOX4", "BOX4")]
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
                Part part = (Part)PartInput.Value;

                Double pipeDiameter = PIPE_DIA.Value, a = A.Value, b = B.Value, c = C.Value, d = D.Value, e = E.Value, webThickness = WEB_T.Value, comlinThickness = COMLIN_T.Value, flangeThickness = FLANGE_T.Value, endThickness = END_T.Value;
                string clampSize = CLAMP_SIZE.Value;
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                double clampF, clampE, clampB;//, alpha, endplateA;
                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_PB1_CM_AUX", "SIZE"), clampSize);
                clampF = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PB1_CM_AUX", "IJUAHgrPSL_PB1_CM_AUX", "F", parameter);
                clampE = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PB1_CM_AUX", "IJUAHgrPSL_PB1_CM_AUX", "E", parameter);
                clampB = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PB1_CM_AUX", "IJUAHgrPSL_PB1_CM_AUX", "B", parameter);

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, c), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "Structure", new Position(0, -c, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                Port port4 = new Port(OccurrenceConnection, part, "Structure", new Position(0, c, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port4"] = port4;

                Port port5 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -c), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port5"] = port5;

                //Validating Inputs  
                if (flangeThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFlangeThicknessGTZ, "FLANGE_T value should be greater than zero"));
                    return;
                }
                if (webThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidWebThicknessGTZ, "WEB_T value should be greater than zero"));
                    return;
                }
                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }
                if (endThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEndTGTZ, "End_T value should be greater than zero"));
                    return;
                }
                if (clampE <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(-a / 2, -(e - clampE - endThickness) / 2, c - flangeThickness);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d basePlate = symbolGeometryHelper.CreateBox(null, a, e - clampE - endThickness, flangeThickness, 9);
                basePlate.Transform(matrix);
                m_Symbolic.Outputs["BASEPLATE"] = basePlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-webThickness / 2, -(e - clampE - endThickness) / 2, c - flangeThickness - ((c - pipeDiameter / 2 - flangeThickness) * 0.75));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d web = symbolGeometryHelper.CreateBox(null, webThickness, e - clampE - endThickness, (c - pipeDiameter / 2 - flangeThickness) * 0.75, 9);
                web.Transform(matrix);
                m_Symbolic.Outputs["WEB"] = web;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-a / 2, -e / 2 + clampE / 2 - endThickness / 2, pipeDiameter / 2 + clampF);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d ep1 = symbolGeometryHelper.CreateBox(null, a, endThickness, c - pipeDiameter / 2 - clampF, 9);
                ep1.Transform(matrix);
                m_Symbolic.Outputs["EP1"] = ep1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-a / 2, e / 2 - clampE / 2 - endThickness / 2, pipeDiameter / 2 + clampF);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d ep2 = symbolGeometryHelper.CreateBox(null, a, endThickness, c - pipeDiameter / 2 - clampF, 9);
                ep2.Transform(matrix);
                m_Symbolic.Outputs["EP2"] = ep2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-c, -(e - clampE - endThickness) / 2, -a / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d basePlate2 = symbolGeometryHelper.CreateBox(null, flangeThickness, e - clampE - endThickness, a, 9);
                basePlate2.Transform(matrix);
                m_Symbolic.Outputs["BASEPLATE2"] = basePlate2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-c + flangeThickness, -(e - clampE - endThickness) / 2, -webThickness / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d web2 = symbolGeometryHelper.CreateBox(null, (c - pipeDiameter / 2 - flangeThickness) * 0.75, e - clampE - endThickness, webThickness, 9);
                web2.Transform(matrix);
                m_Symbolic.Outputs["WEB2"] = web2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-c, -e / 2 + clampE / 2 - endThickness / 2, -a / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d ep3 = symbolGeometryHelper.CreateBox(null, c - pipeDiameter / 2 - clampF, endThickness, a, 9);
                ep3.Transform(matrix);
                m_Symbolic.Outputs["EP3"] = ep3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-c, e / 2 - clampE / 2 - endThickness / 2, -a / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d ep4 = symbolGeometryHelper.CreateBox(null, c - pipeDiameter / 2 - clampF, endThickness, a, 9);
                ep4.Transform(matrix);
                m_Symbolic.Outputs["EP4"] = ep4;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(c - flangeThickness, -(e - clampE - endThickness) / 2, -a / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d basePlate3 = symbolGeometryHelper.CreateBox(null, flangeThickness, e - clampE - endThickness, a, 9);
                basePlate3.Transform(matrix);
                m_Symbolic.Outputs["BASEPLATE3"] = basePlate3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(c - flangeThickness - ((c - pipeDiameter / 2 - flangeThickness) * 0.75), -(e - clampE - endThickness) / 2, -webThickness / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d web3 = symbolGeometryHelper.CreateBox(null, (c - pipeDiameter / 2 - flangeThickness) * 0.75, e - clampE - endThickness, webThickness, 9);
                web3.Transform(matrix);
                m_Symbolic.Outputs["WEB3"] = web3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2 + clampF, -e / 2 + clampE / 2 - endThickness / 2, -a / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d ep5 = symbolGeometryHelper.CreateBox(null, c - pipeDiameter / 2 - clampF, endThickness, a, 9);
                ep5.Transform(matrix);
                m_Symbolic.Outputs["EP5"] = ep5;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2 + clampF, e / 2 - clampE / 2 - endThickness / 2, -a / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d ep6 = symbolGeometryHelper.CreateBox(null, c - pipeDiameter / 2 - clampF, endThickness, a, 9);
                ep6.Transform(matrix);
                m_Symbolic.Outputs["EP6"] = ep6;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-a / 2, -(e - clampE - endThickness) / 2, -c);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d basePlate4 = symbolGeometryHelper.CreateBox(null, a, e - clampE - endThickness, flangeThickness, 9);
                basePlate4.Transform(matrix);
                m_Symbolic.Outputs["BASEPLATE4"] = basePlate4;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-webThickness / 2, -(e - clampE - endThickness) / 2, -c + flangeThickness);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d web4 = symbolGeometryHelper.CreateBox(null, webThickness, e - clampE - endThickness, (c - pipeDiameter / 2 - flangeThickness) * 0.75, 9);
                web4.Transform(matrix);
                m_Symbolic.Outputs["WEB4"] = web4;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-a / 2, -e / 2 + clampE / 2 - endThickness / 2, -c);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d ep7 = symbolGeometryHelper.CreateBox(null, a, endThickness, c - pipeDiameter / 2 - clampF, 9);
                ep7.Transform(matrix);
                m_Symbolic.Outputs["EP7"] = ep7;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-a / 2, e / 2 - clampE / 2 - endThickness / 2, -c);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d ep8 = symbolGeometryHelper.CreateBox(null, a, endThickness, c - pipeDiameter / 2 - clampF, 9);
                ep8.Transform(matrix);
                m_Symbolic.Outputs["EP8"] = ep8;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(e / 2, 0, 0).Subtract(new Position(e / 2 - clampE, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(e / 2 - clampE, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d cylinder = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + clampF, normal.Length);
                m_Symbolic.Outputs["CYL"] = cylinder;


                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(-e / 2 + clampE, 0, 0).Subtract(new Position(-e / 2, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(-e / 2, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d cylinder2 = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + clampF, normal.Length);
                m_Symbolic.Outputs["CYL2"] = cylinder2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(e / 2 - clampE / 2 + clampE * 0.75, 0, 0).Subtract(new Position(e / 2 - clampE / 2 - clampE * 0.75, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(e / 2 - clampE / 2 - clampE * 0.75, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d comlin1 = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + comlinThickness, normal.Length);
                m_Symbolic.Outputs["COMLIN1"] = comlin1;


                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(-e / 2 + clampE / 2 + clampE * 0.75, 0, 0).Subtract(new Position(-e / 2 + clampE / 2 - clampE * 0.75, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(-e / 2 + clampE / 2 - clampE * 0.75, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d comlin2 = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + comlinThickness, normal.Length);
                m_Symbolic.Outputs["COMLIN2"] = comlin2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2, 0, -(clampB + clampF * 2) / 2);
                matrix.Rotate(Math.PI / 4, new Vector(0, 1, 0));
                matrix.Translate(new Vector(0, -e / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d box1 = symbolGeometryHelper.CreateBox(null, d / 2 - pipeDiameter / 2, clampE, clampB + clampF * 2, 9);
                box1.Transform(matrix);
                m_Symbolic.Outputs["BOX1"] = box1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-d / 2, 0, -(clampB + clampF * 2) / 2);
                matrix.Rotate(Math.PI / 4, new Vector(0, 1, 0));
                matrix.Translate(new Vector(0, -e / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d box2 = symbolGeometryHelper.CreateBox(null, d / 2 - pipeDiameter / 2, clampE, clampB + clampF * 2, 9);
                box2.Transform(matrix);
                m_Symbolic.Outputs["BOX2"] = box2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2, 0, -(clampB + clampF * 2) / 2);
                matrix.Rotate(Math.PI / 4, new Vector(0, 1, 0));
                matrix.Translate(new Vector(0, e / 2 - clampE, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d box3 = symbolGeometryHelper.CreateBox(null, d / 2 - pipeDiameter / 2, clampE, clampB + clampF * 2, 9);
                box3.Transform(matrix);
                m_Symbolic.Outputs["BOX3"] = box3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-d / 2, 0, -(clampB + clampF * 2) / 2);
                matrix.Rotate(Math.PI / 4, new Vector(0, 1, 0));
                matrix.Translate(new Vector(0, e / 2 - clampE, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d box4 = symbolGeometryHelper.CreateBox(null, d / 2 - pipeDiameter / 2, clampE, clampB + clampF * 2, 9);
                box4.Transform(matrix);
                m_Symbolic.Outputs["BOX4"] = box4;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_PB4_CM.cs"));
                    return;
                }

            }
        }

        #endregion

    }
}
