//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_PC3.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_PC3
//   Author       :  Vijaya
//   Creation Date:  22-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22-08-2013     Vijaya    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//   30-12-2014     PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
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

    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_PC3 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_PC3"

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(3, "INSULAT", "INSULAT", 0.999999)]
        public InputDouble INSULAT;
        [InputDouble(4, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(5, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(6, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(7, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(8, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(9, "C", "C", 0.999999)]
        public InputDouble C;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("BOT1", "BOT1")]
        [SymbolOutput("BOT2", "BOT2")]
        [SymbolOutput("TOP_BOLT", "TOP_BOLT")]
        [SymbolOutput("TOP_BOLT1", "TOP_BOLT1")]
        [SymbolOutput("BOT_BOLT", "BOT_BOLT")]
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
                Part part = (Part)PartInput.Value;

                Double insulation = INSULAT.Value, pipeDiameter = PIPE_DIA.Value, boltSize = A.Value, a = A.Value, b = B.Value, c=0, cMinimum, d = D.Value, e = E.Value, f = F.Value;
                //this returns initial default value
                if (part != null)
                    c = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrPSL_PC3", "C")).PropValue;
                cMinimum = pipeDiameter / 2 + insulation + boltSize * 1.5;

                if (c < cMinimum)
                    c = 5 * (Math.Floor((pipeDiameter * 1000 / 2 + insulation * 1000 + boltSize * 1000 * 1.5) / 5) / 1000 + 0.001);
                //setting values back
                if (part != null)
                {
                    SupportComponent supportComponent = (SupportComponent)Occurrence;
                    supportComponent.SetPropertyValue(c, "IJOAHgrPSL_PC3", "C");
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, c), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs               
                if (f <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFGTZ, "F value should be greater than zero"));
                    return;
                }
                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
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

                symbolGeometryHelper.ActivePosition = new Position(-(b / 2 + f), -e / 2, pipeDiameter / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d top1 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2, 9);
                top1.Transform(matrix);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(b / 2, -e / 2, pipeDiameter / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d top2 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2, 9);
                top2.Transform(matrix);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-(b / 2 + f), -e / 2, -(pipeDiameter / 2 + f + d * 2));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d bottom1 = symbolGeometryHelper.CreateBox(null, f, e, d * 2 + f, 9);
                bottom1.Transform(matrix);
                m_Symbolic.Outputs["BOT1"] = bottom1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(b / 2, -e / 2, -(pipeDiameter / 2 + f + d * 2));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d bottom2 = symbolGeometryHelper.CreateBox(null, f, e, d * 2 + f, 9);
                bottom2.Transform(matrix);
                m_Symbolic.Outputs["BOT2"] = bottom2;


                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, b / 2 + 2 * f, c).Subtract(new Position(0, -(b / 2 + 2 * f), c));
                symbolGeometryHelper.ActivePosition = new Position(0, -(b / 2 + 2 * f), c);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, b / 2 + 2 * f, (c - ((c + d) - pipeDiameter / 2) / 3)).Subtract(new Position(0, -(b / 2 + 2 * f), (c - ((c + d) - pipeDiameter / 2) / 3)));
                symbolGeometryHelper.ActivePosition = new Position(0, -(b / 2 + 2 * f), (c - ((c + d) - pipeDiameter / 2) / 3));
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt1 = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT1"] = topBolt1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, b / 2 + 2 * f, -pipeDiameter / 2 - f - d).Subtract(new Position(0, -(b / 2 + 2 * f), -pipeDiameter / 2 - f - d));
                symbolGeometryHelper.ActivePosition = new Position(0, -(b / 2 + 2 * f), -pipeDiameter / 2 - f - d);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d botomBolt = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["BOT_BOLT"] = botomBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + f, e);
                matrix.SetIdentity();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -e / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_PC3.cs"));
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
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                string partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_PC3", "SIZE")).PropValue;
                double c = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrPSL_PC3", "C")).PropValue;

                bomDescription = "PSL " + partNumber + " Pipe Clamp Three Bolt Type, C= " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, c, UnitName.DISTANCE_MILLIMETER);

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_PC3.cs."));
                return "";
            }
        }
        #endregion
    }
}
