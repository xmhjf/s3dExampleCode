//----------------------------------------------------------------------------------
//      Copyright (C) 2013 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   SplitMigration rule. This rule covers Plate,Profile,TemplateSet,PinJig,MarkingLine,Margin,Shrinkage, and Tab 
//
//      Author:  
//
//      History:
//      Oct 31th, 2014   Created by Natilus-HSV
//
//-----------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Content.Manufacturing;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Split Migration Rule for Plate,Profile,TemplateSet,PinJig,MarkingLine,Margin,Shrinkage, and Tab 
    /// </summary>
    public class SplitMigration : SplitMigrationRule
    {

        /// <summary>
        /// Evaluates Split Migration or Reverse Split Migration
        /// </summary>
        /// <param name="splitMigrationInfo">The split migration information.</param>
        /// <exception cref="System.ArgumentNullException">Input SplitMigrationInfo is empty</exception>
        public override void Evaluate(SplitMigrationInformation splitMigrationInfo)
        {
            try
            {


                if (splitMigrationInfo == null)
                    throw new ArgumentNullException("Input SplitMigrationInfo is empty");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs
                ReadOnlyCollection<BusinessObject> replacedObjects = null;
                ReadOnlyCollection<BusinessObject> replacingObjects = null;
                ReadOnlyCollection<BusinessObject> mfgObjects = null;

                if (splitMigrationInfo.ReplacedObjects != null)
                {
                    replacedObjects = splitMigrationInfo.ReplacedObjects;
                }

                if (splitMigrationInfo.ReplacingObjects != null)
                {
                    replacingObjects = splitMigrationInfo.ReplacingObjects;
                }

                if (splitMigrationInfo.MfgObjects != null)
                {
                    mfgObjects = splitMigrationInfo.MfgObjects;
                }

                MigrationStatus migrationStatus = splitMigrationInfo.MigrationStatus;
                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                BusinessObject outputObject = null;

                //Default values are defined in XML. User can override these value 
                //Example) User overides following values 
                //splitMigrationInfo.CreateMfgObject = true;
                //splitMigrationInfo.CopyMfgSetting = true;

                string controlAttribute;
                double distanceTolerance = Convert.ToDouble(splitMigrationInfo.GetArguments("Distance").FirstOrDefault().Value);
                bool checkOverlapping = false;
                bool setMfgUpdateInfo = false;
                ReadOnlyCollection<BusinessObject> finalReplacingObjects = null;

                if (mfgObjects != null)
                {
                    if (mfgObjects.Count == 0)
                        return;
                }
                else return;

                ManufacturingBase mfgObject = (ManufacturingBase)mfgObjects[0];

                #region Split Migration
                if (MigrationStatus.Migration == migrationStatus)
                {
                    controlAttribute = Convert.ToString(splitMigrationInfo.GetArguments("ProductionControl", "Migration").FirstOrDefault().Value);
                    checkOverlapping = Convert.ToBoolean(splitMigrationInfo.GetArguments("CheckOverlapping", "Migration").FirstOrDefault().Value);

                    // Checking Overlap 
                    finalReplacingObjects = GetReplacingObjects(mfgObject, replacingObjects, checkOverlapping);

                    if (finalReplacingObjects != null)
                    {
                        switch (controlAttribute.ToUpper())
                        {
                            case "BIGGESTAREA":
                                outputObject = base.GetBiggestAreaObject(mfgObject, finalReplacingObjects, setMfgUpdateInfo);
                                break;
                            case "BIGGESTLENGTH":
                                outputObject = base.GetBiggestLengthObject(mfgObject, finalReplacingObjects, setMfgUpdateInfo);
                                break;
                            case "BIGGESTAREAORLENGTH":
                                #region BIGGESTAREAORLENGTH
                                if (replacedObjects != null && replacedObjects.Count > 0)
                                {
                                    //Add MarkingLine logic to check Overlap 
                                    if (replacedObjects[0] is PlatePartBase)
                                    {
                                        outputObject = base.GetBiggestAreaObject(mfgObject, finalReplacingObjects, setMfgUpdateInfo);
                                    }
                                    else if (replacedObjects[0] is ProfilePart)
                                    {
                                        outputObject = base.GetBiggestLengthObject(mfgObject, finalReplacingObjects, setMfgUpdateInfo);
                                    }
                                    else
                                    {
                                        //ToDo 
                                    }
                                }
                                #endregion
                                break;
                            case "OBJECTWITHINDISTANCE":
                                //Add Logic for checking distance mfgObject and replacingObjects
                                Position position = null;
                                if (mfgObject is Tab)
                                {
                                    Tab tab = (Tab)mfgObject;
                                    if (tab.Position != null)
                                        position = tab.Position;

                                }
                                else if (mfgObject is IRange)
                                {
                                    IRange range = (IRange)mfgObject;
                                    RangeBox rangeBox = range.Range;

                                    if (rangeBox != null)
                                    {
                                        Position highPos = rangeBox.High;
                                        Position lowPos = rangeBox.Low;

                                        if (highPos != null && lowPos != null)
                                        {
                                            position = new Position((highPos.X + lowPos.X) / 2.0, (highPos.Y + lowPos.Y) / 2.0, (highPos.Z + lowPos.Z) / 2.0);
                                        }
                                    }
                                }
                                else
                                {
                                    //To Do
                                }
                                if (position != null)
                                    outputObject = base.GetObjectwithinDistance(position, finalReplacingObjects, setMfgUpdateInfo, distanceTolerance);
                                break;
                            case "DELETEALL":
                                splitMigrationInfo.CreateMfgObject = false;
                                break;
                            //Pick up Last one       
                            case "NOTDEFINDED":
                            default:
                                if (finalReplacingObjects.Count > 0)
                                    outputObject = finalReplacingObjects[finalReplacingObjects.Count - 1];
                                break;
                        }
                    }

                }
                #endregion

                #region Reverse Split Migration
                else if (MigrationStatus.ReverseMigration == migrationStatus)
                {
                    controlAttribute = Convert.ToString(splitMigrationInfo.GetArguments("ProductionControl", "ReverseMigration").FirstOrDefault().Value);
                    setMfgUpdateInfo = Convert.ToBoolean(splitMigrationInfo.GetArguments("SetMfgUpdateInfo", "ReverseMigration").FirstOrDefault().Value);

                    switch (controlAttribute.ToUpper())
                    {
                        case "BIGGESTAREA":
                            outputObject = base.GetBiggestAreaObject(mfgObject, mfgObjects, setMfgUpdateInfo);
                            break;
                        case "BIGGESTLENGTH":
                            outputObject = base.GetBiggestLengthObject(mfgObject, mfgObjects, setMfgUpdateInfo);
                            break;
                        case "BIGGESTAREAORLENGTH":
                            #region BIGGESTAREAORLENGTH
                            if (replacedObjects != null && replacedObjects.Count > 0)
                            {
                                //Add MarkingLine logic to check Overlap 
                                if (replacedObjects[0] is PlatePartBase)
                                {
                                    outputObject = base.GetBiggestAreaObject(mfgObject, mfgObjects, setMfgUpdateInfo);
                                }
                                else if (replacedObjects[0] is ProfilePart)
                                {
                                    outputObject = base.GetBiggestLengthObject(mfgObject, mfgObjects, setMfgUpdateInfo);
                                }
                                else
                                {
                                    //ToDo 
                                }
                            }
                            #endregion
                            break;

                        case "OBJECTWITHINDISTANCE":
                            Position position = null;

                            if (replacingObjects != null && replacingObjects.Count == 1)
                            {
                                if (replacingObjects[0] is IRange)
                                {
                                    IRange range = (IRange)replacingObjects[0];
                                    RangeBox rangeBox = range.Range;

                                    if (rangeBox != null)
                                    {
                                        Position highPos = rangeBox.High;
                                        Position lowPos = rangeBox.Low;

                                        if (highPos != null && lowPos != null)
                                        {
                                            position = new Position((highPos.X + lowPos.X) / 2.0, (highPos.Y + lowPos.Y) / 2.0, (highPos.Z + lowPos.Z) / 2.0);
                                        }
                                    }
                                }
                            }
                            if (position != null)
                                outputObject = base.GetObjectwithinDistance(position, mfgObjects, setMfgUpdateInfo, distanceTolerance);
                            break;

                        case "DELETEALL":
                            splitMigrationInfo.CreateMfgObject = false;
                            break;
                        //Pick up Last one  
                        case "NOTDEFEINED":
                        default:
                            if (mfgObjects.Count > 0)
                                outputObject = mfgObjects[mfgObjects.Count - 1];
                            break;
                    }
                }


                #endregion


                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs
                splitMigrationInfo.OutputObject = outputObject;
                #endregion
            }
            catch(Exception e)
            {
                LogForToDoList(1035, "Problem occurred in SplitMigration Rule " + e.Message);
            }


        }

        /// <summary>
        /// Gets the final replacing objects after checking if there is overlapping between input Manufacturing object and input replacing objects.
        /// </summary>
        /// <param name="mfgObject">The MFG object.</param>
        /// <param name="replacingObjects">The replacing objects.</param>
        /// <param name="checkOverlapping">if set to <c>true</c> [check overlapping].</param>
        /// <returns></returns>
        private ReadOnlyCollection<BusinessObject> GetReplacingObjects(BusinessObject mfgObject, ReadOnlyCollection<BusinessObject> replacingObjects, bool checkOverlapping)
        {
            ReadOnlyCollection<BusinessObject> finalReplacingObjects = null;
            BusinessObject objectTobeChecked = null;
            List<BusinessObject> listofReplacingObjects = new List<BusinessObject>();

            if (checkOverlapping == true)
            {
                Margin margin = null;
                Tab tab = null;
                TopologyPort marginPort = null;
                if (mfgObject is Margin)
                {
                    margin = (Margin)mfgObject;
                    marginPort = ((Margin)mfgObject).Port;
                    if (margin.IsRegional == false)
                    {
                        objectTobeChecked = (BusinessObject)marginPort;
                    }
                    else
                    {
                        TopologyPort edgePort = PortService.GetEdgeFromLateralFace(marginPort, ContextTypes.Base);

                        //Get related port at initial stage of geometry. It calls GetRelatedPort of EntityHelper. 
                        TopologyPort relatedPort = PortService.GetRelatedPortAtStage(edgePort, GeometryStage.Initial);
                        if (relatedPort.Geometry is ICurve)
                        {
                            ICurve inputCurve = (ICurve)relatedPort.Geometry;
                            if (inputCurve != null)
                            {
                                Position startPosition, endPosition;
                                inputCurve.EndPoints(out startPosition, out endPosition);

                                objectTobeChecked = (BusinessObject)base.GetTrimmedCurve(inputCurve, startPosition, endPosition, endPosition);
                            }
                        }
                    }
                }
                else if(mfgObject is Tab)
                {
                    //Get Overlap. but not checking for TabType 4
                    Part tabPart = null;
                    int tabType = -1;

                    tab = (Tab)mfgObject;
                    tabPart = tab.Part;                    

                    if (tabPart != null)
                    {
                        PropertyValue tabTypePropValue = tabPart.GetPropertyValue("IJUASMPlateTabType", "TabType");
                        if (tabTypePropValue != null && tabTypePropValue is PropertyValueInt)
                        {
                            tabType = (int)((PropertyValueInt)tabTypePropValue).PropValue;
                        }
                        else if (tabTypePropValue != null && tabTypePropValue is PropertyValueCodelist)
                        {
                            tabType = (int)((PropertyValueCodelist)tabTypePropValue).PropValue;
                        }
                        else
                        {
                            //To Do
                        }
                    }

                    if (tabType == 4) // Tab Along Edge, Don't Check Overlap
                    {
                        if(replacingObjects != null)
                            finalReplacingObjects = replacingObjects;
                        else
                            finalReplacingObjects = new ReadOnlyCollection<BusinessObject>(listofReplacingObjects);
                        return finalReplacingObjects;

                    }
                    else
                    {
                        TopologyPort drivenPort = null;
                        BusinessObject drivingEntity = null;

                        drivenPort = tab.DrivenPort;
                        drivingEntity = tab.DrivingEntity;
                        if (drivenPort != null && drivingEntity != null)
                        {
                            foreach (BusinessObject replacingObject in replacingObjects)
                            {
                                if (base.IsOverlapping(drivenPort, replacingObject) == true && base.IsOverlapping(drivingEntity, replacingObject) == true)
                                    listofReplacingObjects.Add(replacingObject);
                            }
                        }

                        finalReplacingObjects = new ReadOnlyCollection<BusinessObject>(listofReplacingObjects);

                    }
                }
                else
                {
                    objectTobeChecked = mfgObject;
                }

                //Check Range Box 
                if (objectTobeChecked != null)
                {
                    foreach (BusinessObject replacingObject in replacingObjects)
                    {
                        if (base.IsOverlapping(objectTobeChecked, replacingObject) == true)
                            listofReplacingObjects.Add(replacingObject);
                    }
                    finalReplacingObjects = new ReadOnlyCollection<BusinessObject>(listofReplacingObjects);
                }

            }
            else
            {
                finalReplacingObjects = replacingObjects;
            }
 
            return finalReplacingObjects;
        }
    }
}

