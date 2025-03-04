//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ZVxBOM.cs
//   Witzenmann,Ingr.SP3D.Content.Support.Symbols.KSYBeamClamp
//   Author       : Vinod  
//   Creation Date:  01.12.2015
//   DI-CP-282684  Integrate the newly developed Witzenmann Parts into Product  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Structure.Middle;
using System.Collections.ObjectModel;

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
    public class KSYBeamClamp : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.KSYBeamClamp"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #endregion
       
        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Steel", "Steel")]
        public AspectDefinition m_oSymbolic;
        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                List<Input> additionalInputs = new List<Input>();

                int startIndex = 2;

                additionalInputs.Add(new InputDouble(startIndex, "Height1", "Height1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Height2", "Height2", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Height3", "Height3", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Height4", "Height4", 0, false));

                additionalInputs.Add(new InputDouble(++startIndex, "Length1", "Length1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Length2", "Length2", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Length3", "Length3", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Length4", "Length4", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Length5", "Length5", 0, false));

                additionalInputs.Add(new InputDouble(++startIndex, "Angle1", "Angle1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Angle2", "Angle2", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Thickness1", "Thickness1", 0, false));

                additionalInputs.Add(new InputDouble(++startIndex, "StructType", "StructType", 1, false));
                additionalInputs.Add(new InputDouble(++startIndex, "StructWidth", "StructWidth", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "StructDepth", "StructDepth", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "FlangeThk", "FlangeThk", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "WebThk", "WebThk", 0, false));

                additionalInputs.Add(new InputDouble(++startIndex, "RodDiameter", "RodDiameter", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Width1", "Width1", 0, false));
                

                return additionalInputs;
            }
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
                #region Declaration of variables and their initialisation

                int startIndex = 2;
                Double Height1 = GetDoubleInputValue(startIndex);
                Double Height2 = GetDoubleInputValue(++startIndex);
                Double Height3 = GetDoubleInputValue(++startIndex);
                Double Height4 = GetDoubleInputValue(++startIndex);
                Double Length1 = GetDoubleInputValue(++startIndex);
                Double Length2 = GetDoubleInputValue(++startIndex);
                Double Length3 = GetDoubleInputValue(++startIndex);
                Double Length4 = GetDoubleInputValue(++startIndex);
                Double Length5 = GetDoubleInputValue(++startIndex);
                Double Angle1 = GetDoubleInputValue(++startIndex);
                Double Angle2 = GetDoubleInputValue(++startIndex);
                Double Thickness1 = GetDoubleInputValue(++startIndex);
                int StructType = (int)GetDoubleInputValue(++startIndex);
                Double StructWidth = GetDoubleInputValue(++startIndex);
                Double StructDepth = GetDoubleInputValue(++startIndex);
                Double FlangeThk = GetDoubleInputValue(++startIndex);
                Double WebThk = GetDoubleInputValue(++startIndex);

                Double RodDiameter = GetDoubleInputValue(++startIndex);
                Double Width = GetDoubleInputValue(++startIndex);

                Part part = (Part)m_PartInput.Value;
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;
                #endregion
                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                SymbolGeometryHelper symbolGeomHlpr = new SymbolGeometryHelper();
                symbolGeomHlpr.ActivePosition = new Position(0, 0, 0);
                symbolGeomHlpr.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                
                Matrix4X4 matrix = new Matrix4X4();

                Port Steel = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["Steel"] = Steel;

                double Ang1Ycomponent = (Length3 + Length5 - Length4) * Math.Tan(Angle1);
                double Ang2Ycomponent = Length3 * Math.Tan(Angle2);

                List<Position> points = new List<Position>();

                points.Add(new Position(0, 0, 0));
                points.Add(new Position(0, Height1, 0));
                points.Add(new Position(Length1, Height1, 0));
                points.Add(new Position(Length1, Height2, 0));
                points.Add(new Position(Length2, Height2, 0));
                points.Add(new Position(Length2, -(Height3 - Height2 - Ang2Ycomponent), 0));
                points.Add(new Position(Length2 - Length3, -(Height3 - Height2), 0));
                points.Add(new Position(Length2 - Length3 - Length5, -(Height3 - Height2), 0));
                points.Add(new Position(Length2 - Length3 - Length5, -(Height4 + Ang1Ycomponent), 0));
                points.Add(new Position(Length2 - Length4, -Height4 , 0));
                points.Add(new Position(Length2 - Length4 , 0, 0));
                points.Add(new Position(0, 0, 0));
                
                LineString3d linestring1 = symbolGeomHlpr.CreateLineString(null, points, true);
                
                Projection3d extrusion1 = new Projection3d(linestring1, new Vector(0, 0, 1), Thickness1, true);
                Projection3d extrusion2 = new Projection3d(linestring1, new Vector(0, 0, 1), Thickness1, true);
                Projection3d extrusionL1 = new Projection3d(linestring1, new Vector(0, 0, 1), Thickness1, true);
                Projection3d extrusionL2 = new Projection3d(linestring1, new Vector(0, 0, 1), Thickness1, true);

                string sStructType ;
                 PropertyValueCodelist lStringType = (PropertyValueCodelist)part.GetPropertyValue("IJUAhsStructType", "StructType");
                 sStructType = lStringType.PropertyInfo.CodeListInfo.GetCodelistItem(StructType).DisplayName;



                 if (sStructType == "T")
                {
                    double Bl = Width  / 2;

                    matrix.SetIdentity();
                    matrix.Translate(new Vector((StructWidth - Length2) / 2, Thickness1 / 2 + Bl, 0));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                    extrusion1.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusion1", extrusion1);

                    matrix.SetIdentity();
                    matrix.Translate(new Vector((-StructWidth + Length2) / 2, -Thickness1 / 2 + Bl, 0));
                    matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                    extrusion2.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusion2", extrusion2);
                    
                    matrix.SetIdentity();
                    matrix.Translate(new Vector((StructWidth- Length2) / 2, Thickness1 / 2 - Bl, 0));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                    extrusionL1.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusionL1", extrusionL1);

                    matrix.SetIdentity();
                    matrix.Translate(new Vector((-StructWidth + Length2) / 2, -Thickness1 / 2 - Bl, 0));
                    matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                    extrusionL2.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusionL2", extrusionL2);

                    Line3d rodline = new Line3d(new Position(StructWidth / 2 +  Length2, 0, 0), new Position(-StructWidth / 2 - Length2, 0, 0));
                    Circle3d M12Circle = new Circle3d(new Position(StructWidth / 2 +  Length2, 0, 0), new Vector(-1, 0, 0), 0.006);
                    
                    matrix.SetIdentity(); int index = 0;
                    matrix.Translate(new Vector(0, Bl , Height2 /2));
                    Collection<Surface3d> Rod1SurfaceCol = new Collection<Surface3d>();
                    Rod1SurfaceCol = Surface3d.GetSweepSurfacesFromCurve( rodline, M12Circle, (SurfaceSweepOptions)1);
                    foreach (Surface3d item in Rod1SurfaceCol)
                    {
                        Rod1SurfaceCol[index].Transform(matrix);
                        m_oSymbolic.Outputs.Add("Rod1Surface" + index, item);
                        ++index;
                    }

                    matrix.SetIdentity();  index = 0;
                    matrix.Translate(new Vector(0, -Bl, Height2 / 2));
                    Collection<Surface3d> Rod2SurfaceCol = new Collection<Surface3d>();
                    Rod2SurfaceCol = Surface3d.GetSweepSurfacesFromCurve( rodline, M12Circle, (SurfaceSweepOptions)1);
                    foreach (Surface3d item in Rod2SurfaceCol)
                    {
                        Rod2SurfaceCol[index].Transform(matrix);
                        m_oSymbolic.Outputs.Add("Rod2Surface" + index, item);
                        ++index;
                    }
                }

                 if (sStructType == "C" || sStructType == "U")
                {
                    //StructDepth = 0.381; StructWidth = 0.094; 
                    double Bl = Width  / 2;

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-StructWidth+(Length2/2), Bl - Thickness1 / 2, StructDepth / 2));
                    matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                    extrusion2.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusion2", extrusion2);

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-StructWidth + (Length2 / 2), Bl + Thickness1 / 2, -StructDepth / 2));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                    extrusion1.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusion1", extrusion1);

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-StructWidth + (Length2 / 2), -Bl - Thickness1 / 2, StructDepth / 2));
                    matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                    extrusionL1.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusionL1", extrusionL1);

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-StructWidth + (Length2 / 2), -Bl + Thickness1 / 2, -StructDepth / 2));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Rotate(Math.PI, new Vector(0, 1, 0));                    
                    extrusionL2.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusionL2", extrusionL2);

                    double roddia = 0.012; double arcrad = StructWidth / 8;

                    Position P1 = new Position(-StructWidth -  Length2, 0, StructDepth / 2+Height2/2 );
                    Position P2 = new Position(-arcrad + roddia / 2, 0, StructDepth / 2 + Height2 / 2);
                    Position CP1 = new Position(-arcrad + roddia / 2, 0, StructDepth / 2 + Height2 / 2 - arcrad);
                    Position P3 = new Position(roddia / 2, 0, StructDepth / 2 + Height2 / 2 - arcrad);
                    Position P4 = new Position(roddia / 2, 0,- StructDepth / 2 - Height2 / 2 + arcrad);
                    Position CP2 = new Position(-arcrad + roddia / 2, 0, -StructDepth / 2 - Height2 / 2 + arcrad);
                    Position P5 = new Position(-arcrad + roddia / 2, 0, -StructDepth / 2 - Height2 / 2);
                    Position P6 = new Position(-StructWidth -  Length2, 0, -StructDepth / 2 - Height2 / 2);
                    
                    Collection<ICurve> curveCollection = new Collection<ICurve>();

                    Line3d rodline = new Line3d(P1, P2); 
                    Arc3d RodC1 = new Arc3d(CP1, new Vector(0, 1, 0), P2, P3);
                    Line3d rodline2 = new Line3d(P3, P4);
                    Arc3d RodC2 = new Arc3d(CP2, new Vector(0, 1, 0), P4, P5);
                    Line3d rodline3 = new Line3d(P5, P6);
                    
                    curveCollection.Add(rodline);
                    curveCollection.Add(RodC1);
                    curveCollection.Add(rodline2);
                    curveCollection.Add(RodC2);
                    curveCollection.Add(rodline3);

                    ComplexString3d lineString1 = new ComplexString3d(curveCollection);
                    Circle3d M12Circle = new Circle3d(P1, new Vector(-1, 0, 0), roddia / 2);

                    matrix.SetIdentity(); int index = 0;
                    matrix.Translate(new Vector(0, Bl, 0));
                    Collection<Surface3d> Rod1SurfaceCol = new Collection<Surface3d>();
                    Rod1SurfaceCol = Surface3d.GetSweepSurfacesFromCurve( lineString1, M12Circle, (SurfaceSweepOptions)1);
                    foreach (Surface3d item in Rod1SurfaceCol)
                    {
                        Rod1SurfaceCol[index].Transform(matrix);
                        m_oSymbolic.Outputs.Add("Rod1Surface" + index, item);
                        ++index;
                    }

                    matrix.SetIdentity(); index = 0;
                    matrix.Translate(new Vector(0, -Bl, 0));
                    Collection<Surface3d> Rod2SurfaceCol = new Collection<Surface3d>();
                    Rod2SurfaceCol = Surface3d.GetSweepSurfacesFromCurve( lineString1, M12Circle, (SurfaceSweepOptions)1);
                    foreach (Surface3d item in Rod2SurfaceCol)
                    {
                        Rod2SurfaceCol[index].Transform(matrix);
                        m_oSymbolic.Outputs.Add("Rod2Surface" + index, item);
                        ++index;
                    }
                }

                 if (sStructType == "L")
                {
                   double Bl = Width  / 2;

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-StructWidth + (Length2/2), Bl + Thickness1 / 2, 0));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                    extrusion1.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusion1", extrusion1);

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, Bl - Thickness1 / 2, StructDepth - (Length2 / 2)));
                    matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Rotate(-Math.PI/2, new Vector(0, 1, 0));
                    extrusion2.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusion2", extrusion2);                   

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-StructWidth + (Length2 / 2), -Bl + Thickness1 / 2, 0));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                    extrusionL1.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusionL1", extrusionL1);

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, -Bl - Thickness1 / 2, StructDepth - (Length2 / 2)));                    
                    matrix.Rotate(-Math.PI/2, new Vector(0, 0, 1));
                    matrix.Rotate(-Math.PI/2 , new Vector(1, 0, 0));
                    extrusionL2.Transform(matrix);
                    m_oSymbolic.Outputs.Add("extrusionL2", extrusionL2);

                    double roddia = 0.012;

                    Position P1 = new Position(Height2 /2, 0, StructDepth +  Length2);
                    Position P2 = new Position(Height2 / 2, 0, 0 + Height2 / 2);
                    Position CP1 = new Position(-Height2 / 2, 0,  Height2 / 2);
                    Position P3 = new Position(- Height2 / 2, 0, 0 - Height2 / 2);
                    Position P4 = new Position(-StructWidth - Length2, 0, 0 - Height2 / 2);
                    
                    Collection<ICurve> curveCollection = new Collection<ICurve>();

                    Line3d rodline = new Line3d(P1, P2);                    
                    Arc3d RodC2 = new Arc3d(CP1, new Vector(0, 1, 0), P2, P3);
                    Line3d rodline3 = new Line3d(P3, P4);
                    
                    curveCollection.Add(rodline);
                    curveCollection.Add(RodC2);
                    curveCollection.Add(rodline3);

                    ComplexString3d lineString1 = new ComplexString3d(curveCollection);
                    Circle3d M12Circle = new Circle3d(P1, new Vector(0, 0, -1), roddia / 2);
                    
                    matrix.SetIdentity(); int index = 0;
                    matrix.Translate(new Vector(0, Bl, 0));
                    Collection<Surface3d> Rod1SurfaceCol = new Collection<Surface3d>();
                    Rod1SurfaceCol = Surface3d.GetSweepSurfacesFromCurve( lineString1, M12Circle, (SurfaceSweepOptions)1);
                    foreach (Surface3d item in Rod1SurfaceCol)
                    {
                        Rod1SurfaceCol[index].Transform(matrix);
                        m_oSymbolic.Outputs.Add("Rod1Surface" + index, item);
                        ++index;
                    }

                    matrix.SetIdentity(); index = 0;
                    matrix.Translate(new Vector(0, -Bl, 0));
                    Collection<Surface3d> Rod2SurfaceCol = new Collection<Surface3d>();
                    Rod2SurfaceCol = Surface3d.GetSweepSurfacesFromCurve( lineString1, M12Circle, (SurfaceSweepOptions)1);
                    foreach (Surface3d item in Rod2SurfaceCol)
                    {
                        Rod2SurfaceCol[index].Transform(matrix);
                        m_oSymbolic.Outputs.Add("Rod2Surface" + index, item);
                        ++index;
                    }
                }
            } 
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of KSYBeamClamp.cs"));
                }
            }
        }
        #endregion

        #region "Caluculating WeightCG"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                      
                double  weight, cogx, cogy, cogz;
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
                    cogx = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogx = 0;
                }

                try
                {
                    cogy = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogy = 0;
                }

                try
                {
                    cogz = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogz = 0;
                }
                supportComponent.SetWeightAndCOG(weight, cogx, cogy, cogz);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of KSYBeamClamp.cs"));
                }
            }
        }
        #endregion

    }

}