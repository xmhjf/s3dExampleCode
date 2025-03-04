//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_351.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_351
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
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
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
    public class PSL_351 : HangerComponentSymbolDefinition
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_351"

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "ANGLE", "ANGLE", 1)]
        public InputDouble ANGLE;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(4, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputString(5, "CLAMP", "CLAMP", "No Value")]
        public InputString CLAMP;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BASE", "BASE")]
        [SymbolOutput("WEB", "WEB")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("BOT1", "BOT1")]
        [SymbolOutput("BOT2", "BOT2")]
        [SymbolOutput("TOP_BOLT", "TOP_BOLT")]
        [SymbolOutput("BOT_BOLT", "BOT_BOLT")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("TOP1_B", "TOP1_B")]
        [SymbolOutput("TOP2_B", "TOP2_B")]
        [SymbolOutput("BOT1_B", "BOT1_B")]
        [SymbolOutput("BOT2_B", "BOT2_B")]
        [SymbolOutput("TOP_BOLT_B", "TOP_BOLT_B")]
        [SymbolOutput("BOT_BOLT_B", "BOT_BOLT_B")]
        [SymbolOutput("BODY_B", "BODY_B")]
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
                int angleValue = 0;
                int angle = (int)((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrPSL_351", "ANGLE")).PropValue;

                if (angle == 2)
                    angleValue = 90;
                else if (angle == 1)
                    angleValue = 45;

                Double pipeDiameter = PIPE_DIA.Value;
                Double a = A.Value;
                String clamp = CLAMP.Value;
                Double a351 = A.Value;

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                double k, l, webThickness, flangeThickness, b, c, d, e, f;

                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPipe_Dia_mm", "PIPE_DIA"), pipeDiameter);

                k = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_349", "IJUAHgrPSL_349", "K", parameter, 0.003, PSLSymbolServices.ComparisionOperator.BETWEEN_WITHOUT_LIMITS);
                l = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_349", "IJUAHgrPSL_349", "L", parameter, 0.003, PSLSymbolServices.ComparisionOperator.BETWEEN_WITHOUT_LIMITS);
                webThickness = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_349", "IJUAHgrPSL_FLANGE_DIM", "WEB_T", parameter, 0.003, PSLSymbolServices.ComparisionOperator.BETWEEN_WITHOUT_LIMITS);
                flangeThickness = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_349", "IJUAHgrPSL_FLANGE_DIM", "FLANGE_T", parameter, 0.003, PSLSymbolServices.ComparisionOperator.BETWEEN_WITHOUT_LIMITS);

                parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_SIZE", "SIZE"), clamp);

                a = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PC2", "IJUAHgrPSL_PC2", "A", parameter);
                b = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PC2", "IJUAHgrPSL_PC2", "B", parameter);
                c = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PC2", "IJUAHgrPSL_PC2", "C", parameter);
                d = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PC2", "IJUAHgrPSL_PC2", "D", parameter);
                e = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PC2", "IJUAHgrPSL_PC2", "E", parameter);
                f = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PC2", "IJUAHgrPSL_PC2", "F", parameter);

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -a351), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (k <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidKGTZ, "K value should be greater than zero"));
                    return;
                }
                if (l <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidLGTZ, "L value should be greater than zero"));
                    return;
                }
                if (flangeThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFlangeThicknessGTZ, "FLANGE_T value should be greater than zero"));
                    return;
                }
                if (webThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidWebThicknessGTZ, "WEB_T should be greater than zero"));
                    return;
                }                
                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }
                if (f <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFGTZ, "F value should be greater than zero"));
                    return;
                }
                if ((c + d - pipeDiameter / 2) <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCDPipeDiaGTZ, "( C + D - PIPE_DIA/2 ) value should be greater than zero"));
                    return;
                }
                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(b + 4 * f,0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBAndFNEZ, "(B + 4 * F) value cannot be zero"));
                    return;
                }

                //creating pipe shoe
                symbolGeometryHelper.ActivePosition = new Position(-k / 2, -l / 2, -a351);
                Projection3d baseBox = symbolGeometryHelper.CreateBox(null, k, l, flangeThickness, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                baseBox.Transform(matrix);
                m_Symbolic.Outputs["BASE"] = baseBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-webThickness / 2, -l / 2, flangeThickness - a351);
                Projection3d web = symbolGeometryHelper.CreateBox(null, webThickness, l, a351 - pipeDiameter / 2 - f - flangeThickness, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                web.Transform(matrix);
                m_Symbolic.Outputs["WEB"] = web;

                //creating first pipe clamps
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(b / 2.0 + f), -l / 2, pipeDiameter / 2.0);
                Projection3d top1 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2.0, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(angleValue * Math.PI / 180, new Vector(1, 0, 0));
                top1.Transform(matrix);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(b / 2.0, -l / 2, pipeDiameter / 2.0);
                Projection3d top2 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2.0, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(angleValue * Math.PI / 180, new Vector(1, 0, 0));
                top2.Transform(matrix);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(b / 2.0 + f), -l / 2, -(c + d));
                Projection3d bottom1 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2.0, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(angleValue * Math.PI / 180, new Vector(1, 0, 0));
                bottom1.Transform(matrix);
                m_Symbolic.Outputs["BOT1"] = bottom1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(b / 2.0, -l / 2, -(c + d));
                Projection3d bottom2 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2.0, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(angleValue * Math.PI / 180, new Vector(1, 0, 0));
                bottom2.Transform(matrix);
                m_Symbolic.Outputs["BOT2"] = bottom2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, a / 2, b + 4 * f);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-l / 2 + e / 2, c, -(b + 4 * f) / 2));
                matrix.Rotate(Math.PI / 2 + angleValue * Math.PI / 180, new Vector(1, 0, 0), new Position(0, 0, 0));
                topBolt.Transform(matrix);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d topBolt1 = symbolGeometryHelper.CreateCylinder(null, a / 2, b + 4 * f);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-l / 2 + e / 2, -c, -(b + 4 * f) / 2));
                matrix.Rotate(Math.PI / 2 + angleValue * Math.PI / 180, new Vector(1, 0, 0), new Position(0, 0, 0));
                topBolt1.Transform(matrix);
                m_Symbolic.Outputs["BOT_BOLT"] = topBolt1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2.0 + f, e);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -l / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;

                //creating second pipe clamps
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(b / 2.0 + f), l / 2 - e, pipeDiameter / 2.0);
                Projection3d top1B = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2.0, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(angleValue * Math.PI / 180, new Vector(1, 0, 0));
                top1B.Transform(matrix);
                m_Symbolic.Outputs["TOP1_B"] = top1B;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(b / 2.0, l / 2 - e, pipeDiameter / 2.0);
                Projection3d top2B = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2.0, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(angleValue * Math.PI / 180, new Vector(1, 0, 0));
                top2B.Transform(matrix);
                m_Symbolic.Outputs["TOP2_B"] = top2B;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(b / 2.0 + f), l / 2 - e, -(c + d));
                Projection3d bottom1B = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2.0, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(angleValue * Math.PI / 180, new Vector(1, 0, 0));
                bottom1B.Transform(matrix);
                m_Symbolic.Outputs["BOT1_B"] = bottom1B;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(b / 2.0, l / 2 - e, -(c + d));
                Projection3d bottom2B = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2.0, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(angleValue * Math.PI / 180, new Vector(1, 0, 0));
                bottom2B.Transform(matrix);
                m_Symbolic.Outputs["BOT2_B"] = bottom2B;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d topBoltB = symbolGeometryHelper.CreateCylinder(null, a / 2, b + 4 * f);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(l / 2 - e / 2, c, -(b + 4 * f) / 2));
                matrix.Rotate(Math.PI / 2 + angleValue * Math.PI / 180, new Vector(1, 0, 0), new Position(0, 0, 0));
                topBoltB.Transform(matrix);
                m_Symbolic.Outputs["TOP_BOLT_B"] = topBoltB;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d topBoltB1 = symbolGeometryHelper.CreateCylinder(null, a / 2, b + 4 * f);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(l / 2 - e / 2, -c, -(b + 4 * f) / 2));
                matrix.Rotate(Math.PI / 2 + angleValue * Math.PI / 180, new Vector(1, 0, 0), new Position(0, 0, 0));
                topBoltB1.Transform(matrix);
                m_Symbolic.Outputs["BOT_BOLT_B"] = topBoltB1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d bodyB = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2.0 + f, e);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, l / 2 - e, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                bodyB.Transform(matrix);
                m_Symbolic.Outputs["BODY_B"] = bodyB;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_351."));
                    return;
                }
            }
        }
        #endregion
    }
}
