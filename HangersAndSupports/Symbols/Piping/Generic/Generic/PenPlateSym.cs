//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   PenPlateSym.cs
//    Generic,Ingr.SP3D.Content.Support.Symbols.PenPlateSym
//   Author       :  Hema
//   Creation Date:  19-11-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//    19-11-2012     Hema    CR-CP-222274  Converted HS_Generic VB Project to C# .Net  
//	  27/03/2013	 Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   6/11/2013     Vijaya   CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

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
    public class PenPlateSym : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Generic,Ingr.SP3D.Content.Support.Symbols.PenPlateSym"
        //----------------------------------------------------------------------------------
        double angle;

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "LeftPipeRadius", "LeftPipeRadius", 0.999999)]
        public InputDouble m_dLeftPipeRadius;
        [InputDouble(3, "RightPipeRadius", "RightPipeRadius", 0.999999)]
        public InputDouble m_dRightPipeRadius;
        [InputDouble(4, "DistBetPipes", "DistBetPipes", 0.999999)]
        public InputDouble m_dDistBetPipes;
        [InputDouble(5, "LeftOffset", "LeftOffset", 0.999999)]
        public InputDouble m_dLeftOffset;
        [InputDouble(6, "RightOffset", "RightOffset", 0.999999)]
        public InputDouble m_dRightOffset;
        [InputDouble(7, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(8, "NoOfRoutes", "NoOfRoutes",0.999999)]
        public InputDouble m_oNoOfRoutes;
        [InputString(9, "PlateShape", "PlateShape","No Value")]
        public InputString m_oPlateShape;
        [InputDouble(10, "FilletRadius", "FilletRadius", 0.999999)]
        public InputDouble m_dFilletRadius;
        [InputDouble(11, "PlateWidth", "PlateWidth", 0.999999)]
        public InputDouble m_dPlateWidth;
        [InputDouble(12, "PlateDepth", "PlateDepth", 0.999999)]
        public InputDouble m_dPlateDepth;
        [InputDouble(13, "WidthOffset", "WidthOffset", 0.999999)]
        public InputDouble m_dWidthOffset;
        [InputDouble(14, "DepthOffset", "DepthOffset", 0.999999)]
        public InputDouble m_dDepthOffset;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("PipePlate1", "PipePlate1")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Hole1", "Hole1")]
        [SymbolOutput("Hole2", "Hole2")]
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
                
                double leftPipeRadius = m_dLeftPipeRadius.Value;
                double rightPipeRadius = m_dRightPipeRadius.Value;
                double distanceBetweenPipes = m_dDistBetPipes.Value;
                double leftOffset = m_dLeftOffset.Value;                
                double rightOffset = m_dRightOffset.Value;
                double thickness = m_dThickness.Value;                
                double noOfRoutes = m_oNoOfRoutes.Value;
                string shape = m_oPlateShape.Value;
                double filletRadius = m_dFilletRadius.Value;
                double plateWidth = m_dPlateWidth.Value;
                double plateDepth = m_dPlateDepth.Value;
                double widthOffset = m_dWidthOffset.Value;
                double depthOffset = m_dDepthOffset.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                Collection<ICurve> curveCollection; 
                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                if (shape.ToUpper() != "ROUND")
                {
                    double angleValue = ((rightPipeRadius + rightOffset) - (leftPipeRadius + leftOffset)) / (distanceBetweenPipes);
                    angle = Math.Asin(angleValue);
                }
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (leftPipeRadius <= 0 && leftOffset <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidleftOffsetAndleftPipeRadius, "Left Pipe Radius and Left Pipe Offset should be greater than zero"));
                    return;
                }
                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }

                if (shape.ToUpper() == "ROUND")
                {
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d pipePlate1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, (leftPipeRadius + leftOffset), thickness);
                    m_Symbolic.Outputs["PipePlate1"] = pipePlate1;
                }
                else if (shape.ToUpper() == "RECTANGLE")
                {
                    curveCollection = new Collection<ICurve>();

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d arc = symbolGeometryHelper.CreateArc(null, filletRadius, Math.PI / 2);
                    matrix.SetIdentity();
                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                    arc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-((plateWidth + widthOffset) / 2 - filletRadius), (plateDepth + depthOffset) / 2 - filletRadius, 0));
                    arc.Transform(matrix);
                    m_Symbolic.Outputs["Arc1"] = arc;
                    curveCollection.Add(arc);

                    curveCollection.Add(new Line3d( new Position(-(plateWidth + widthOffset) / 2 + filletRadius, (plateDepth + depthOffset) / 2, 0), new Position((plateWidth + widthOffset) / 2 - filletRadius, (plateDepth + depthOffset) / 2, 0)));

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d arc1 = symbolGeometryHelper.CreateArc(null, filletRadius, Math.PI / 2);
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(((plateWidth + widthOffset) / 2 - filletRadius), (plateDepth + depthOffset) / 2 - filletRadius, 0));
                    arc1.Transform(matrix);
                    curveCollection.Add(arc1);

                    curveCollection.Add(new Line3d( new Position((plateWidth + widthOffset) / 2, (plateDepth + depthOffset) / 2 - filletRadius, 0), new Position((plateWidth + widthOffset) / 2, -(plateDepth + depthOffset) / 2 + filletRadius, 0)));

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d arc2 = symbolGeometryHelper.CreateArc(null, filletRadius, -(Math.PI / 2 + Math.PI));
                    matrix.SetIdentity();
                    matrix.Rotate(Math.PI + Math.PI / 2, new Vector(0, 0, 1));
                    arc2.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(((plateWidth + widthOffset) / 2 - filletRadius), -(plateDepth + depthOffset) / 2 + filletRadius, 0));
                    arc2.Transform(matrix);
                    curveCollection.Add(arc2);

                    curveCollection.Add(new Line3d( new Position((plateWidth + widthOffset) / 2 - filletRadius, -(plateDepth + depthOffset) / 2, 0), new Position(-(plateWidth + widthOffset) / 2 + filletRadius, -(plateDepth + depthOffset) / 2, 0)));

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d arc3 = symbolGeometryHelper.CreateArc(null, filletRadius, Math.PI / 2);
                    matrix.SetIdentity();
                    matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                    arc3.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-((plateWidth + widthOffset) / 2 - filletRadius), -(plateDepth + depthOffset) / 2 + filletRadius, 0));
                    arc3.Transform(matrix);
                    curveCollection.Add(arc3);

                    curveCollection.Add(new Line3d( new Position(-(plateWidth + widthOffset) / 2, -(plateDepth + depthOffset) / 2 + filletRadius, 0), new Position(-(plateWidth + widthOffset) / 2, (plateDepth + depthOffset) / 2 - filletRadius, 0)));

                    Vector lineVector = new Vector(0, 0, thickness);
                    Projection3d pipePlate1 = new Projection3d( new ComplexString3d(curveCollection), lineVector, lineVector.Length, true);
                    m_Symbolic.Outputs["PipePlate1"] = pipePlate1;
                }
                else
                {
                    if (noOfRoutes > 1)
                    {
                        Port port2 = new Port(OccurrenceConnection, part, "Route_2", new Position(0, -distanceBetweenPipes, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_Symbolic.Outputs["Port2"] = port2;
                    }
                    Collection<ICurve> curveCollection1 = new Collection<ICurve>();

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix.SetIdentity();
                    matrix.Rotate(Math.PI / 2 + angle, new Vector(0, 0, 1));

                    Arc3d arc = symbolGeometryHelper.CreateArc(null, leftPipeRadius + leftOffset, Math.PI - 2 * angle);
                    arc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                    arc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, -thickness / 2, 0));
                    arc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2 + Math.PI), new Vector(0, 0, 1));
                    arc.Transform(matrix);
                    curveCollection1.Add(arc);

                    curveCollection1.Add(new Line3d( new Position(-thickness / 2, (leftPipeRadius + leftOffset) * Math.Sin(angle), (leftPipeRadius + leftOffset) * Math.Cos(angle)), new Position(-thickness / 2, -distanceBetweenPipes + ((rightPipeRadius + rightOffset) * Math.Sin(angle)), (rightPipeRadius + rightOffset) * Math.Cos(angle))));

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix.SetIdentity();
                    matrix.Rotate(-Math.PI / 2 - angle, new Vector(0, 0, 1));

                    Arc3d arc1 = symbolGeometryHelper.CreateArc(null, rightPipeRadius + rightOffset, Math.PI + 2 * angle);
                    arc1.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                    arc1.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(distanceBetweenPipes, -thickness / 2, 0));
                    arc1.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                    arc1.Transform(matrix);
                    curveCollection1.Add(arc1);

                    curveCollection1.Add(new Line3d( new Position(-thickness / 2, -distanceBetweenPipes + ((rightPipeRadius + rightOffset) * Math.Sin(angle)), -(rightPipeRadius + rightOffset) * Math.Cos(angle)), new Position(-thickness / 2, (leftPipeRadius + leftOffset) * Math.Sin(angle), -(leftPipeRadius + leftOffset) * Math.Cos(angle))));

                    Vector lineVector = new Vector(1, 0, 0);
                    Projection3d pipePlate1 = new Projection3d( new ComplexString3d(curveCollection1), lineVector, thickness, true);
                    m_Symbolic.Outputs["PipePlate1"] = pipePlate1;
                }
            }
            catch    //General Unhandled exception
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrConstructOutputs, "Error in ConstructOutputs of PenPlateSym"));
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
                Double weight, cogX=0, cogY=0, cogZ=0;
                CatalogStructHelper query=new CatalogStructHelper(); 
                double Volume=0.0;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double Thickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrPntrPlate", "Thickness")).PropValue;
                string strMaterialType = (string)((PropertyValueString)part.GetPropertyValue("IJHgrPartMaterial", "MaterialType")).PropValue;
                string strMaterialGrade = (string)((PropertyValueString)part.GetPropertyValue("IJHgrPartMaterial", "MaterialGrade")).PropValue;

                Ingr.SP3D.ReferenceData.Middle.Material material;
                material= query.GetMaterial(strMaterialType,strMaterialGrade);
                weight = Volume * material.Density;
               
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrWeightCG, "Error in WeightCG of PenPlateSym"));
                }
            }
        }
        #endregion
    }
}
