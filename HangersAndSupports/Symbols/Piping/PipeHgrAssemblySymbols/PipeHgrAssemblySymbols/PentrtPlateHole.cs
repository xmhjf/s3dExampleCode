//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2016, Intergraph Corporation. All rights reserved.
//
//   PentrtPlate_1.cs
//   PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.PentrtPlateAspect
//   Author       :  Siva
//   Creation Date:  02-Nov-2016
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//	 13/04/2017		Siva 	 DM-CP-302096 Post - HangerSupport should relate with Structure Before Cut Port  
//   28/04/2017     Siva     TR-CP-313893  Not able to place hole trace on pen plate with hole aspect in specific case  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]

    public class PentrtPlateHole : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.PentrtPlateAspect"
        //----------------------------------------------------------------------------------
        //bool firstTime = false;
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "Thickness", "Thickness", 1)]
        public InputDouble Thickness;

        [InputDouble(3, "PipeRadius", "Pipe Radius", 0.25)]
        public InputDouble PipeRadius;

        [InputDouble(4, "Width", "Width", 2)]
        public InputDouble Width;

        [InputDouble(5, "Height", "Height", 2)]
        public InputDouble Height;

        [InputDouble(6, "PipePort_X", "Pipe Port X", 2)]
        public InputDouble PipePort_X;

        [InputDouble(7, "PipePort_Y", "Pipe Port Y", 2)]
        public InputDouble PipePort_Y;

        [InputDouble(8, "HoleAspThick", "Hole Aspect Thickness", 1)]
        public InputDouble HoleAspectThick;

        [InputDouble(9, "HoleAspGap", "Hole Aspect Gap", 1)]
        public InputDouble HoleAspectGap;

        [InputDouble(10, "AspectDirect", "Hole Aspect Direction", 1)]
        public InputDouble HoleAspectDir;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("StructPort", "Structure Port")]
        [SymbolOutput("PipePort", "Pipe Port")]
        [SymbolOutput("Plate1", "Plate1")]
        [SymbolOutput("Plate2", "Plate2")]
        [SymbolOutput("Plate3", "Plate3")]
        [SymbolOutput("Plate4", "Plate4")]
        [SymbolOutput("Plate5", "Plate5")]
        public AspectDefinition simplePhysicalAspect;

        [Aspect("Hole", "Hole Aspect", 524288)]
        [SymbolOutput("Cylinder", "Cylinder")]
        public AspectDefinition m_HoleAspect;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        SupportHelper supportHelper = null; BusinessObject businessObject = null;
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                Double thickness = Thickness.Value;
                Double pipeRadius = PipeRadius.Value;
                Double aspectThick = HoleAspectThick.Value;
                Double holeAspGap = HoleAspectGap.Value;
                int Direction = (int)HoleAspectDir.Value;


                if (aspectThick < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrInvalidHoleAspectThickness, "Invalid Thickness. Value should be greater than or equal to zero."));
                    return;
                }

                if (holeAspGap < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrInvalidHoleAspectGap, "Invalid Gap. Value should be greater than or equal to zero."));
                    return;
                }

                Boolean isPlaceByStruct = false;

                double width = 0, height = 0, pipePortX = 0, pipePortY = 0;

                RelationCollection hgrRelation = Occurrence.GetRelationship("SupportHasComponents", "Support");
                businessObject = hgrRelation.TargetObjects[0];
                supportHelper = new SupportHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                if (supportHelper.PlacementType == PlacementType.PlaceByStruct)
                    isPlaceByStruct = true;

                Position position = new Position(0, 0, 0);
                string boundingBox = string.Empty, boundingBoxName = string.Empty;
                BoundingBoxHelper boundingBoxHelper = new BoundingBoxHelper((Ingr.SP3D.Support.Middle.Support)businessObject);

                RefPortHelper refporthlpr = new RefPortHelper((Ingr.SP3D.Support.Middle.Support)businessObject);

                double toGlobalZ = 0; double structZtoGlobalX = 0; double structZtoGlobalY = 0;

                BusinessObject StructRefPort = refporthlpr.ReferencePort("Structure");



                if (StructRefPort != null)
                {
                    try
                    {
                        toGlobalZ = Math.Round(refporthlpr.AngleBetweenPorts("Structure", PortAxisType.Z, OrientationAlong.Global_Z) * (180.0 / Math.PI), 0);
                        structZtoGlobalX = Math.Round(refporthlpr.AngleBetweenPorts("Structure", PortAxisType.Z, OrientationAlong.Global_X) * (180.0 / Math.PI), 0);

                        Matrix4X4 Port = new Matrix4X4();
                        Vector Struct_Z = new Vector(0, 0, 0), GlobalY = new Vector(0, 0, 0);

                        Port = refporthlpr.PortLCS("Structure");
                        Struct_Z.Set(Port.ZAxis.X, Port.ZAxis.Y, Port.ZAxis.Z);

                        Port = new Matrix4X4();
                        GlobalY.Set(0, 1, 0);

                        structZtoGlobalY = Math.Round(GetAngleBetweenVectors(Struct_Z, GlobalY) * (180.0 / Math.PI), 0);

                    }
                    catch
                    {
                        toGlobalZ = 0;
                    }
                }
                else
                    toGlobalZ = 0;


                try
                {
                    boundingBoxHelper.CreateStandardBoundingBoxes(false);

                    // Get required information about the Bounding Box Surrounding the Pipe.
                    // The Bounding Box used depends on the command.

                    if (isPlaceByStruct)
                    {
                        // Use the Structure Bounding Box
                        boundingBox = "BBS";
                        boundingBoxName = "BBS_Low";
                    }
                    else
                    {
                        // Use the Route Bouding Box
                        boundingBox = "BBR";
                        boundingBoxName = "BBR_Low";
                    }

                    if (boundingBoxHelper.GetBoundingBox(boundingBox) != null)
                    {
                        width = boundingBoxHelper.GetBoundingBox(boundingBox).Width;
                        height = boundingBoxHelper.GetBoundingBox(boundingBox).Height;
                    }
                    else
                    {
                        width = 0.0;
                        height = 0.0;
                    }


                    position = boundingBoxHelper.GetBoundingBox(boundingBox).GetRelativeRouteCenterPosition(1);
                    if (position != null)
                    {
                        pipePortX = position.X;
                        pipePortY = position.Y;
                    }
                }
                catch
                {
                    width = 0.1;
                    height = 0.1;
                    pipePortX = 0.0;
                    pipePortY = 0.0;
                    pipeRadius = 0.05;
                    thickness = 0.025;
                }

                // Correct Height if necessary
                if (height <= 0)
                    height = 1.0;
                // Correct Width if necessary
                if (width <= 0)
                    width = 1.0;
                // Correct Thickness if necessary
                if (thickness <= 0.0)
                    thickness = 0.0;
                // Correct PipeRadius if necessary
                if (pipeRadius <= 0.0)
                    pipeRadius = 1.0;

                if (supportHelper.SupportedObjects == null)
                {
                    pipeRadius = 0.05;
                    thickness = 0.025;
                    width = 0.1;
                    height = 0.1;
                }

                // Set the Height value on the Plate.
                Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)businessObject;
                Occurrence.SetPropertyValue(width, "IJUAHgrOccGeometry", "Width");
                // Set the Height value on the Plate.
                Occurrence.SetPropertyValue(height, "IJUAHgrOccGeometry", "Height");
                // Set the PipePort_X value on the Plate.
                Occurrence.SetPropertyValue(pipePortX, "IJUAHgrOccPenPlate", "PipePort_X");
                // Set the PipePort_Y value on the Plate.
                Occurrence.SetPropertyValue(pipePortY, "IJUAHgrOccPenPlate", "PipePort_Y");
                // Set the Aspect thickness value on the Plate.
                Occurrence.SetPropertyValue(aspectThick, "IJUAHgrHoleAspectThk", "HoleAspThick");
                // Set the Aspect Gap value on the Plate.
                Occurrence.SetPropertyValue(holeAspGap, "IJUAHgrHoleAspectThk", "HoleAspGap");
                // Set the Aspect direction value on the Plate.
                Occurrence.SetPropertyValue(Direction, "IJUAHgrHoleAspectDir", "AspectDirect");

                // Define the offset from the provided bounding box.
                double offset = pipeRadius;
                double chamfer = 0.5 * offset;
                // ========================================
                // Bracket contour and projection
                // ========================================

                // Create Line String Representing the outline of the plate
                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0.0, -chamfer, -offset));
                pointCollection.Add(new Position(0.0, width + chamfer, -offset));
                pointCollection.Add(new Position(0.0, width + offset, -chamfer));
                pointCollection.Add(new Position(0.0, width + offset, height + chamfer));
                pointCollection.Add(new Position(0.0, width + chamfer, height + offset));
                pointCollection.Add(new Position(0.0, -chamfer, height + offset));
                pointCollection.Add(new Position(0.0, -offset, height + chamfer));
                pointCollection.Add(new Position(0.0, -offset, -chamfer));
                pointCollection.Add(new Position(0.0, -chamfer, -offset));


                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Plane3d plane = new Plane3d(pointCollection);
                SupportedHelper supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)businessObject);

                double[] minPipeRadiuscol = new double[supportHelper.SupportedObjects.Count];
                double minPipeRadius = 0;
                double temp = 0;

                if (supportHelper.SupportedObjects != null)
                {
                    for (int i = 1; i <= supportHelper.SupportedObjects.Count; i++)
                    {
                        SupportedObjectInfo supportedObjectInfo = supportedHelper.SupportedObjectInfo(i);
                        if (supportedObjectInfo.GetType() == typeof(PipeObjectInfo))
                            pipeRadius = ((PipeObjectInfo)supportedObjectInfo).OutsideDiameter / 2.0;
                        else if (supportedObjectInfo.GetType() == typeof(ConduitObjectInfo))
                            pipeRadius = ((ConduitObjectInfo)supportedObjectInfo).OutsideDiameter / 2.0;
                        try
                        {
                            position = boundingBoxHelper.GetBoundingBox(boundingBox).GetRelativeRouteCenterPosition(i);
                        }
                        catch
                        {
                            position = new Position(0, 0, 0);
                        }
                        Circle3d hole = new Circle3d(new Position(0.0, position.X, position.Y), new Vector(1.0, 0.0, 0.0), pipeRadius);
                        curveCollection.Add(hole);
                        Collection<ComplexString3d> collection = new Collection<ComplexString3d>();
                        ComplexString3d holeComplexString = new ComplexString3d(curveCollection);
                        plane.AddHole(holeComplexString);
                        curveCollection.Clear();
                        minPipeRadiuscol[i - 1] = pipeRadius;
                    }

                    //Get the max piperadius
                    for (int i = 0; i < supportHelper.SupportedObjects.Count; i++)
                    {
                        if (supportHelper.SupportedObjects.Count > 1)
                        {
                            for (int j = i + 1; j < supportHelper.SupportedObjects.Count; j++)
                            {
                                if (minPipeRadiuscol[j] < minPipeRadiuscol[i])
                                {
                                    temp = minPipeRadiuscol[i];
                                    minPipeRadiuscol[i] = minPipeRadiuscol[j];
                                    minPipeRadiuscol[j] = temp;
                                }
                            }
                        }
                    }
                    minPipeRadius = minPipeRadiuscol[0];
                }
                Line3d zVector = new Line3d(new Position(0, 0, 0), new Position(thickness, 0, 0));
                Collection<Surface3d> surface = Surface3d.GetSweepSurfacesFromPlane(plane, new Line3d(zVector), (SurfaceSweepOptions)1);

                int i1 = 1;
                foreach (Surface3d item in surface)
                {
                    simplePhysicalAspect.Outputs.Add("Plate" + i1, item);
                    i1++;
                }

                // Hanger Ports
                Port structPort = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0.0, 0.0), new Vector(1.0, 0.0, 0.0), new Vector(0.0, 0.0, 1.0));
                simplePhysicalAspect.Outputs["StructPort"] = structPort;
                Port pipePort = new Port(OccurrenceConnection, part, "Route", new Position(0.0, pipePortX, pipePortY), new Vector(1.0, 0.0, 0.0), new Vector(0.0, 0.0, 1.0));
                simplePhysicalAspect.Outputs["PipePort"] = pipePort;

                if (aspectThick > 0)
                {
                    double xVal = 0;

                    //For Vertical planes
                    if (structZtoGlobalY == 0 || structZtoGlobalX == 0)
                    {
                        if (Direction == 1)
                            xVal = 0.0;
                        else
                            xVal = (thickness - aspectThick);
                    }
                    else if (structZtoGlobalY == 180 || structZtoGlobalX == 180)
                    {
                        if (Direction == 1)
                            xVal = (thickness - aspectThick);
                        else
                            xVal = 0.0;
                    }
                    //for horizontal planes
                    else if (toGlobalZ == 0)
                    {
                        if (Direction == 1)
                            xVal = 0.0;
                        else
                            xVal = (thickness - aspectThick);
                    }
                    else if (toGlobalZ == 180)
                    {
                        if (Direction == 1)
                            xVal = (thickness - aspectThick);
                        else
                            xVal = 0.0;
                    }
                    //for inclined planes
                    else
                    {
                        if ((toGlobalZ > 0 && toGlobalZ <= 45) || (toGlobalZ > 90 && toGlobalZ <= 135))
                        {
                            if (Direction == 1)
                                xVal = 0.0;
                            else
                                xVal = (thickness - aspectThick);
                        }
                        else if ((toGlobalZ > 135 && toGlobalZ < 180) || (toGlobalZ > 45 && toGlobalZ <= 90))
                        {
                            if (Direction == 1)
                                xVal = (thickness - aspectThick);
                            else
                                xVal = 0.0;
                        }
                    }

                    if ((supportHelper.SupportedObjects.Count) == 1)
                    {
                        // create Hole aspect
                        Matrix4X4 matrix = new Matrix4X4();
                        Circle3d Circle = new Circle3d(new Position(xVal, pipePortX, pipePortY), new Vector(1.0, 0.0, 0.0), pipeRadius + holeAspGap);

                        Projection3d Cylinder = new Projection3d(Circle, new Vector(1.0, 0.0, 0.0), aspectThick, false);
                        Cylinder.Transform(matrix);
                        m_HoleAspect.Outputs["Cylinder"] = Cylinder;
                    }

                    else
                    {
                        //Create curve collection for 
                        Collection<ICurve> HoleAspCurveColl = new Collection<ICurve>();
                        if ((cmpdbl((height - minPipeRadius), minPipeRadius) == false) && (cmpdbl((width - minPipeRadius), minPipeRadius) == false))
                        {
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, minPipeRadius, minPipeRadius), null, new Position(xVal, minPipeRadius, -holeAspGap), new Position(xVal, -holeAspGap, minPipeRadius)));
                            HoleAspCurveColl.Add(new Line3d(new Position(xVal, -holeAspGap, minPipeRadius), new Position(xVal, -holeAspGap, height - minPipeRadius)));
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, minPipeRadius, height - minPipeRadius), null, new Position(xVal, -holeAspGap, height - minPipeRadius), new Position(xVal, minPipeRadius, height + holeAspGap)));
                            HoleAspCurveColl.Add(new Line3d(new Position(xVal, minPipeRadius, height + holeAspGap), new Position(xVal, width - minPipeRadius, height + holeAspGap)));
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, width - minPipeRadius, height - minPipeRadius), null, new Position(xVal, width - minPipeRadius, height + holeAspGap), new Position(xVal, width + holeAspGap, height - minPipeRadius)));
                            HoleAspCurveColl.Add(new Line3d(new Position(xVal, width + holeAspGap, height - minPipeRadius), new Position(xVal, width + holeAspGap, minPipeRadius)));
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, width - minPipeRadius, minPipeRadius), null, new Position(xVal, width + holeAspGap, minPipeRadius), new Position(xVal, width - minPipeRadius, -holeAspGap)));
                            HoleAspCurveColl.Add(new Line3d(new Position(xVal, width - minPipeRadius, -holeAspGap), new Position(xVal, minPipeRadius, -holeAspGap)));
                        }
                        else if ((cmpdbl((height - minPipeRadius), minPipeRadius) == true) && (cmpdbl((width - minPipeRadius), minPipeRadius) == false))
                        {
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, minPipeRadius, minPipeRadius), null, new Position(xVal, minPipeRadius, -holeAspGap), new Position(xVal, -holeAspGap, minPipeRadius)));
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, minPipeRadius, minPipeRadius), null, new Position(xVal, -holeAspGap, minPipeRadius), new Position(xVal, minPipeRadius, height + holeAspGap)));
                            HoleAspCurveColl.Add(new Line3d(new Position(xVal, minPipeRadius, height + holeAspGap), new Position(xVal, width - minPipeRadius, height + holeAspGap)));
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, width - minPipeRadius, minPipeRadius), null, new Position(xVal, width - minPipeRadius, height + holeAspGap), new Position(xVal, width + holeAspGap, minPipeRadius)));
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, width - minPipeRadius, minPipeRadius), null, new Position(xVal, width + holeAspGap, minPipeRadius), new Position(xVal, width - minPipeRadius, -holeAspGap)));
                            HoleAspCurveColl.Add(new Line3d(new Position(xVal, width - minPipeRadius, -holeAspGap), new Position(xVal, minPipeRadius, -holeAspGap)));
                        }
                        else if ((cmpdbl((height - minPipeRadius), minPipeRadius) == false) && (cmpdbl((width - minPipeRadius), minPipeRadius) == true))
                        {
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, minPipeRadius, minPipeRadius), null, new Position(xVal, minPipeRadius, -holeAspGap), new Position(xVal, -holeAspGap, minPipeRadius)));
                            HoleAspCurveColl.Add(new Line3d(new Position(xVal, -holeAspGap, minPipeRadius), new Position(xVal, -holeAspGap, height - minPipeRadius)));
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, minPipeRadius, height - minPipeRadius), null, new Position(xVal, -holeAspGap, height - minPipeRadius), new Position(xVal, minPipeRadius, height + holeAspGap)));
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, minPipeRadius, height - minPipeRadius), null, new Position(xVal, minPipeRadius, height + holeAspGap), new Position(xVal, width + holeAspGap, height - minPipeRadius)));
                            HoleAspCurveColl.Add(new Line3d(new Position(xVal, width + holeAspGap, height - minPipeRadius), new Position(xVal, width + holeAspGap, minPipeRadius)));
                            HoleAspCurveColl.Add(new Arc3d(new Position(xVal, minPipeRadius, minPipeRadius), null, new Position(xVal, width + holeAspGap, minPipeRadius), new Position(xVal, minPipeRadius, -holeAspGap)));
                        }


                        ComplexString3d AspComplexString = new ComplexString3d(HoleAspCurveColl);

                        Projection3d HoleAsp = new Projection3d(AspComplexString, new Vector(1.0, 0.0, 0.0), aspectThick, false);

                        Matrix4X4 matrix = new Matrix4X4();
                        HoleAsp.Transform(matrix);
                        m_HoleAspect.Outputs["Cylinder"] = HoleAsp;
                    }
                }

            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PentrtPlate_1.cs"));
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
                double width = 0, height = 0;
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double thickness = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrThickness", "Thickness")).PropValue;
                try
                {
                    width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccGeometry", "Width")).PropValue;
                }
                catch { width = 0; }
                try
                {
                    height = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccGeometry", "Height")).PropValue;
                }
                catch { height = 0; }
                double pipeRadius = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrRadius", "PipeRadius")).PropValue;
                double offset = pipeRadius;
                double chamfer = 0.5 * offset;
                RelationCollection hgrRelation = supportComponentBO.GetRelationship("SupportHasComponents", "Support");
                businessObject = hgrRelation.TargetObjects[0];
                supportHelper = new SupportHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                SupportedHelper supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                if (supportHelper.SupportedObjects != null)
                {
                    for (int i = 1; i <= supportHelper.SupportedObjects.Count; i++)
                    {
                        SupportedObjectInfo supportedObjectInfo = supportedHelper.SupportedObjectInfo(i);
                        if (supportedObjectInfo.GetType() == typeof(PipeObjectInfo))
                            pipeRadius = ((PipeObjectInfo)supportedObjectInfo).OutsideDiameter / 2.0;
                        else if (supportedObjectInfo.GetType() == typeof(ConduitObjectInfo))
                            pipeRadius = ((ConduitObjectInfo)supportedObjectInfo).OutsideDiameter / 2.0;
                    }
                }

                double volume = (width * height * thickness) - (4 * 1 / 2 * offset * chamfer * thickness) - (Math.PI * (pipeRadius * pipeRadius) * thickness);
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                string materialType = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJHgrPartMaterial", "MaterialType")).PropValue;
                string materialGrade = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJHgrPartMaterial", "MaterialGrade")).PropValue;
                Material material = catalogStructHelper.GetMaterial(materialType, materialGrade);

                double weight = volume * material.Density;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, 0, 0, 0);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PentrtPlate_1.cs."));
            }
        }
        private double GetAngleBetweenVectors(Vector vector1, Vector vector2)
        {
            try
            {
                double dblDotProd = (vector1.Dot(vector2) / (vector1.Length * vector2.Length));
                double dblArcCos = 0.0;
                if (cmpdbl(Math.Abs(dblDotProd), 1) == false)
                {
                    dblArcCos = Math.PI / 2 - Math.Atan(dblDotProd / Math.Sqrt(1 - dblDotProd * dblDotProd));
                }
                else if (cmpdbl(dblDotProd, -1) == true)
                {
                    dblArcCos = Math.PI;
                }
                else if (cmpdbl(dblDotProd, 1) == true)
                {
                    dblArcCos = 0;
                }
                return dblArcCos;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static bool cmpdbl(double leftvaule, double rightvalue)
        {
            bool result = true;
            if (leftvaule > rightvalue - 0.00001 && leftvaule < rightvalue + 0.00001)
                result = true;
            else
                result = false;
            return result;
        }

        #endregion

    }
}