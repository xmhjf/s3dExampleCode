//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   SFS5368.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5368
//   Author       :   Vijaya
//   Creation Date:  18/3/2013
//   Description: CR-CP-222272 Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//    18/3/2013      Vijaya   CR-CP-222272 Initial Creation
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
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class SFS5368 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5368"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PipeND", "PipeND", 0.999999)]
        public InputDouble m_dPipeND;
        [InputDouble(3, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(4, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputString(5, "TypeandSize", "TypeandSize", "No Value")]
        public InputString m_TypeandSize;
        [InputDouble(6, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(7, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(8, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(9, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(10, "StopperType", "StopperType", 1)]
        public InputDouble m_StopperType;
        [InputString(11, "Material", "Material", "No Value")]
        public InputString m_Material;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Body", "Body")]
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

                Double pipeNominalDiameter = m_dPipeND.Value;
                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double length = m_dL.Value;
                String typeandSize = m_TypeandSize.Value;
                Double width = m_dWidth.Value;
                Double thickness = m_dThickness.Value;
                Double S = m_dS.Value;
                Double B = m_dB.Value;
                int type = (int)m_StopperType.Value;
                String material = m_Material.Value;
                Matrix4X4 rotateMatrix = new Matrix4X4();
                Line3d line;
                Collection<ICurve> curveCollection = new Collection<ICurve>();


                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                if (length == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrTnvalidLengthNZero, "Length value cannot be zero"));
                    return;
                }
                if (type == 1)
                {

                    Double y = width * Math.Sin(Math.PI / 4);
                    Double z = Math.Sqrt(Math.Pow(pipeDiameter / 2, 2) - Math.Pow((width * Math.Sin(Math.PI / 4)), 2));
                    Double offset = thickness * Math.Sin(Math.PI / 4);
                    Collection<Position> pointCollection = new Collection<Position>();
                    pointCollection.Add(new Position(-length / 2, -y, -z));
                    pointCollection.Add(new Position(-length / 2, 0, -z - y));
                    pointCollection.Add(new Position(-length / 2, y, -z));
                    pointCollection.Add(new Position(-length / 2, y + offset, -z - offset));
                    pointCollection.Add(new Position(-length / 2, 0, -z - y - offset * 2));
                    pointCollection.Add(new Position(-length / 2, -y - offset, -z - offset));
                    pointCollection.Add(new Position(-length / 2, -y, -z));
                    Vector projectionVector = new Vector(length, 0, 0);
                    Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), length, true);
                    m_Symbolic.Outputs["Body"] = body;
                }
                else
                {
                    Double angle = B / (pipeDiameter / 2);
                    Double calculation1 = Math.Sin(angle / 2) * (pipeDiameter / 2);
                    Double calculation2 = Math.Sqrt((pipeDiameter / 2 * pipeDiameter / 2) - (calculation1 * calculation1));
                    Double calculation3 = Math.Sin(angle / 2) * (pipeDiameter / 2 + S);
                    Double calculation4 = Math.Sqrt(((pipeDiameter / 2 + S) * (pipeDiameter / 2 + S)) - (calculation3 * calculation3));

                    if (angle < Math.PI)
                    {
                        calculation2 = -calculation2;
                        calculation4 = -calculation4;
                    }


                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d outerArc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2 + S, angle);
                    rotateMatrix.Rotate((Math.PI / 2 - angle / 2), new Vector(0, 0, 1));
                    rotateMatrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                    rotateMatrix.Translate(new Vector(0, -length / 2, 0));
                    rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    outerArc.Transform(rotateMatrix);
                    curveCollection.Add(outerArc);


                    line = new Line3d(new Position(-length / 2, calculation3, calculation4), new Position(-length / 2, calculation1, calculation2));
                    curveCollection.Add(line);


                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d innerArc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, angle);
                    rotateMatrix = new Matrix4X4();
                    rotateMatrix.Rotate((Math.PI / 2 - angle / 2), new Vector(0, 0, 1));
                    rotateMatrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                    rotateMatrix.Translate(new Vector(0, -length / 2, 0));
                    rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    innerArc.Transform(rotateMatrix);
                    curveCollection.Add(innerArc);


                    line = new Line3d(new Position(-length / 2, -calculation1, calculation2), new Position(-length / 2, -calculation3, calculation4));
                    curveCollection.Add(line);
                    
                    Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), length, true);
                    m_Symbolic.Outputs["Body"] = body;

                }

            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5368."));
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
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                PropertyValueCodelist typeCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAFINL_StopperType", "StopperType");
                long stopperType = typeCodeList.PropValue;
                if (stopperType <= 0 || stopperType > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidStopperType, "StopperType Code list value should be between 1 and 2"));
                }

                string typeandSize = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUAFINL_TypeandSize", "TypeandSize")).PropValue;               

                if (stopperType == 1) // L shape
                    bomDescription = "Stopper SFS 5368 A" + typeandSize.Trim();
                else // cuved plate
                    bomDescription = "Stopper SFS 5368 B" + typeandSize.Trim();

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SFS5368."));
                return "";
            }
        }
        #endregion

    }

}
