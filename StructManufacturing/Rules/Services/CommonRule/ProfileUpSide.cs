//----------------------------------------------------------------------------------
//      Copyright (C) 2013 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   BevelNo rule returns calculated deviation angle.   
//
//      Author:  
//
//      History:
//      October 14, 2013   Created by Natilus-HSV
//
//-----------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Content.Manufacturing;
using Ingr.SP3D.Structure.Middle;
using ProcessInfo = Ingr.SP3D.Content.Manufacturing.ProcessInfo;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// BevelNo Rule
    /// </summary>
    public class ProfileUpSide : ProfileUpSideRule
    {

        /// <summary>
        /// Gets the bevel deviation angle.
        /// </summary>
        /// <param name="processInfo">The process info.</param>
  
        public override void Evaluate(ProcessInformation processInfo)
        {
            try
            {
                if (processInfo == null)
                    throw new ArgumentNullException("Input ProcessInfo is empty");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs
                ProfilePart profilePart = null;
                ManufacturingBase mfgPart = null;

                if (processInfo.ManufacturingPart != null)
                {
                    mfgPart = (ManufacturingBase)processInfo.ManufacturingPart;
                }

                if (processInfo.ManufacturingParent != null)
                {
                    profilePart = (ProfilePart)processInfo.ManufacturingParent;
                }
                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing
                //Get Arguments
                Dictionary<int, object> args = processInfo.GetArguments("ProductionControl");
                if (args.Count == 0)
                {
                    return;
                }

                int profileSide = (int)SectionFaceType.Unknown;

                ReadOnlyCollection<BusinessObject> sectionFaces = null;
                ReadOnlyCollection<int> sectionFaceTypes = null;
                TopologySurface webLeftSurface = null;
                TopologySurface webRightSurface = null;
                Position centerPositionSurface = null;
                Vector normalVector = null;
                int webLeftIndex = -1;
                int webRightIndex = -1;
                //            ISurface surfaceGeometry = null;

                string controlAttribute = Convert.ToString(processInfo.GetArguments("ProductionControl").FirstOrDefault().Value);
                switch (controlAttribute.ToUpper())
                {
                    case "CHECKVALIDFACE":
                        #region Check Valid Face
                        int sectionFaceType = Convert.ToInt32(processInfo.GetArguments("SectionFaceType").FirstOrDefault().Value);

                        //Check ValidFace Port for Top & Bottom 
                        if (profilePart != null)
                        {
                            profilePart.GetSectionFaces(true, out sectionFaces, out sectionFaceTypes);
                            if (sectionFaces != null && sectionFaceTypes != null)
                            {
                                int sectionFaceIndex = sectionFaceTypes.IndexOf(sectionFaceType);
                                TopologySurface sectionSurface = null;

                                try
                                {
                                    if (sectionFaceIndex >= 0)
                                        sectionSurface = (TopologySurface)sectionFaces[sectionFaceIndex];
                                }
                                catch
                                {
                                }

                                if (sectionSurface != null)
                                {

                                    if (sectionFaceType == (int)SectionFaceType.Top)
                                    {
                                        //Check if there is Top_Flange_Right_Bottom or Top_Flange_Left_Bottom 
                                        if (sectionFaceTypes.Contains((int)SectionFaceType.Top_Flange_Left_Bottom) == true ||
                                            sectionFaceTypes.Contains((int)SectionFaceType.Top_Flange_Right_Bottom) == true)
                                        {
                                            profileSide = sectionFaceType;
                                            break;
                                        }
                                    }
                                    else if (sectionFaceType == (int)SectionFaceType.Bottom)
                                    {
                                        //Check if there is Top_Flange_Right_Bottom or Top_Flange_Left_Bottom 
                                        if (sectionFaceTypes.Contains((int)SectionFaceType.Bottom_Flange_Left_Top) == true ||
                                            sectionFaceTypes.Contains((int)SectionFaceType.Bottom_Flange_Right_Top) == true)
                                        {
                                            profileSide = sectionFaceType;
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        profileSide = sectionFaceType;
                                        break;
                                    }

                                }


                            }
                        }

                        #endregion
                        break;
                    case "LOADFACEDATAFROMXML":
                        #region LoadFaceDataFromXML
                        string xmlFilePath = Convert.ToString(processInfo.GetArguments("XMLFile").FirstOrDefault().Value);
                        string faceName = Convert.ToString(processInfo.GetArguments("FaceName").FirstOrDefault().Value);
                        profileSide = Convert.ToInt32(processInfo.GetArguments("SectionFaceType").FirstOrDefault().Value);
                        if (mfgPart is ManufacturingProfile && String.IsNullOrEmpty(xmlFilePath) == false && String.IsNullOrWhiteSpace(xmlFilePath) == false)
                        {
                            base.LoadFaceDataFromXML((ManufacturingProfile)mfgPart, faceName, xmlFilePath);
                        }
                        #endregion
                        break;
                    case "INNERWEB":
                        #region InnerWeb
                        //Logic
                        //Determines the marking side for the mfg profile if it is inner.
                        //Any web left or web right surface whose nomal is pointing to the centre of the hull is Inner side.
                        //Note: If a stiffener is not in standard directions(X,Y,Z), this might be helpful to identify the desired face 
                        if (profilePart != null)
                        {

                            if (profilePart is MemberPart)
                            {
                                profileSide = (int)SectionFaceType.Unknown;
                            }
                            else
                            {
                                profilePart.GetSectionFaces(true, out sectionFaces, out sectionFaceTypes);
                                if (sectionFaces != null && sectionFaceTypes != null)
                                {
                                    webLeftIndex = sectionFaceTypes.IndexOf((int)SectionFaceType.Web_Left);
                                    webRightIndex = sectionFaceTypes.IndexOf((int)SectionFaceType.Web_Right);

                                    try
                                    {
                                        if (webLeftIndex >= 0)
                                            webLeftSurface = (TopologySurface)sectionFaces[webLeftIndex];
                                        if (webRightIndex >= 0)
                                            webRightSurface = (TopologySurface)sectionFaces[webRightIndex];
                                    }
                                    catch
                                    {
                                    }

                                    if (webLeftSurface != null)
                                    {
                                        //There is matching FacePort 
                                        // Case 1 centerPositionSurface.y <= 0 and normalVector.Y > 0
                                        // Case 2 centerPositionSurface.y => 0 and normalVector.Y < 0

                                        base.GetSurfaceApproxNormalAndCenter(webLeftSurface, out centerPositionSurface, out normalVector);

                                        if (centerPositionSurface.Y * normalVector.Y < 0)
                                        {
                                            profileSide = (int)SectionFaceType.Web_Left;
                                            break;
                                        }
                                        else
                                        {

                                        }
                                    }

                                    if (webRightSurface != null)
                                    {
                                        // Case 1 centerPositionSurface.y < 0 and normalVector.Y > 0
                                        // Case 2 centerPositionSurface.y > 0 and normalVector.Y < 0

                                        base.GetSurfaceApproxNormalAndCenter(webRightSurface, out centerPositionSurface, out normalVector);

                                        if (centerPositionSurface.Y * normalVector.Y < 0)
                                        {
                                            profileSide = (int)SectionFaceType.Web_Right;
                                            break;
                                        }
                                        else
                                        {

                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                        break;
                    case "OUTERWEB":
                        #region OuterWeb
                        //Description:   Determines the marking side for the mfg profile if it is Outer side.
                        //Any web left or web right surface whose nomal is pointing to outside of the hull is Outer side.
                        //Note: If a stiffener is not in standard directions(X,Y,Z), this might be helpful to identify the desired face
                        if (profilePart != null)
                        {

                            if (profilePart is MemberPart)
                            {
                                profileSide = (int)SectionFaceType.Unknown;
                            }
                            else
                            {
                                profilePart.GetSectionFaces(true, out sectionFaces, out sectionFaceTypes);
                                if (sectionFaces != null && sectionFaceTypes != null)
                                {
                                    webLeftIndex = sectionFaceTypes.IndexOf((int)SectionFaceType.Web_Left);
                                    webRightIndex = sectionFaceTypes.IndexOf((int)SectionFaceType.Web_Right);

                                    try
                                    {
                                        if (webLeftIndex >= 0)
                                            webLeftSurface = (TopologySurface)sectionFaces[webLeftIndex];
                                        if (webRightIndex >= 0)
                                            webRightSurface = (TopologySurface)sectionFaces[webRightIndex];
                                    }
                                    catch
                                    {
                                    }

                                    if (webLeftSurface != null)
                                    {
                                        base.GetSurfaceApproxNormalAndCenter(webLeftSurface, out centerPositionSurface, out normalVector);

                                        if (centerPositionSurface.Y * normalVector.Y > 0)
                                        {
                                            profileSide = (int)SectionFaceType.Web_Left;
                                            break;
                                        }
                                        else
                                        {

                                        }

                                    }

                                    if (webRightSurface != null)
                                    {
                                        // Case 1 centerPositionSurface.y <= 0 and normalVector.Y > 0
                                        // Case 2 centerPositionSurface.y => 0 and normalVector.Y < 0
                                        base.GetSurfaceApproxNormalAndCenter(webRightSurface, out centerPositionSurface, out normalVector);

                                        if (centerPositionSurface.Y * normalVector.Y > 0)
                                        {
                                            profileSide = (int)SectionFaceType.Web_Right;
                                            break;
                                        }
                                        else
                                        {

                                        }

                                    }

                                }
                            }
                        }

                        #endregion
                        break;
                    case "ALONGDIRECTION":
                        #region AlongDirection
                        if (profilePart != null)
                        {
                            if (profilePart is MemberPart)
                            {
                                //Skip for MemberPart
                                profileSide = (int)SectionFaceType.Unknown;
                            }
                            else
                            {
                                //Logic
                                //1. Get the Center and Normal Vector approximately. 

                                double dirX = Convert.ToDouble(processInfo.GetArguments("DirX").FirstOrDefault().Value);
                                double dirY = Convert.ToDouble(processInfo.GetArguments("DirY").FirstOrDefault().Value);
                                double dirZ = Convert.ToDouble(processInfo.GetArguments("DirZ").FirstOrDefault().Value);
                                //Radian, it is converted into DB Unit- Radian 
                                double angleTol = Convert.ToDouble(processInfo.GetArguments("Angle").FirstOrDefault().Value);

                                Vector directionVec = new Vector(dirX, dirY, dirZ);
                                double angle = 0.0;

                                profilePart.GetSectionFaces(true, out sectionFaces, out sectionFaceTypes);
                                if (sectionFaces != null && sectionFaceTypes != null)
                                {
                                    webLeftIndex = sectionFaceTypes.IndexOf((int)SectionFaceType.Web_Left);
                                    webRightIndex = sectionFaceTypes.IndexOf((int)SectionFaceType.Web_Right);

                                    try
                                    {
                                        if (webLeftIndex >= 0)
                                            webLeftSurface = (TopologySurface)sectionFaces[webLeftIndex];
                                        if (webRightIndex >= 0)
                                            webRightSurface = (TopologySurface)sectionFaces[webRightIndex];
                                    }
                                    catch
                                    {
                                    }


                                    if (webLeftSurface != null)
                                    {
                                        base.GetSurfaceApproxNormalAndCenter(webLeftSurface, out centerPositionSurface, out normalVector);
                                        normalVector.Length = 1.0;

                                        //Radian 0<= angle <= PI
                                        angle = Math.Acos(directionVec.Dot(normalVector));

                                        if (angle <= angleTol)
                                        {
                                            profileSide = (int)SectionFaceType.Web_Left;
                                            break;
                                        }
                                    }

                                    if (webRightSurface != null)
                                    {
                                        base.GetSurfaceApproxNormalAndCenter(webRightSurface, out centerPositionSurface, out normalVector);
                                        normalVector.Length = 1.0;

                                        //Radian 0<= angle <= PI
                                        angle = Math.Acos(directionVec.Dot(normalVector));
                                        if (angle <= angleTol)
                                        {
                                            profileSide = (int)SectionFaceType.Web_Right;
                                            break;
                                        }
                                    }
                                }

                            }
                        }
                        #endregion
                        break;
                    default:
                        profileSide = (int)SectionFaceType.Unknown;
                        break;

                }
                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs
                processInfo.SetAttribute((int)ProcessInfo.ProcessValues.ProfileUpSide, profileSide);
                #endregion
            }
            catch(Exception e)
            {
                LogForToDoList(2032, "Problem occurred in ProfileUpSide Rule" + e.Message);
            }
        }
    }
}
