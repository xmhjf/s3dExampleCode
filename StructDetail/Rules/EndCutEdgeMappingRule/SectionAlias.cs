//-----------------------------------------------------------------------------
//      Copyright (C) 2010 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionAlias is a helper class that will use 
//                  the mapped edges to determine which type of SectionAliasType
//                  to return.  SectionAliasType is an enum that will tell the 
//                  user which type of cross section they are dealing with.  For 
//                  example, SectionAliasType.WebTopFlange lets the user know that
//                  the cross section has a web and a top flange.  The user then 
//                  knows this is a T type cross section.    
//
//      Author:  3XCalibur
//
//      History:
//      November 05, 2010       WR                  Created
//
//-----------------------------------------------------------------------------

using Ingr.SP3D.Structure.Middle;
using System.Collections.Generic;
namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Specifies the type of cross section.  For example, WebTopFlange, is a web with a top flange which
    /// denotes a T type cross section.
    /// </summary>
    public enum SectionAliasType
    {
        /// <summary>
        /// Cannot determine the cross section alias.
        /// </summary>
        UnknownAlias = -1,
        /// <summary>
        /// A cross section with only a web.
        /// </summary>
        Web = 0,
        /// <summary>
        /// A mirrored L type cross section with the top flange pointing to the right.
        /// </summary>
        WebTopFlangeRight = 1,
        /// <summary>
        /// A built up cross section with the top flange pointing to the right and the web coming off the top flange.      
        /// </summary>
        WebBuiltUpTopFlangeRight = 2,
        /// <summary>
        /// An L type cross section where the bottom flange points to the right.
        /// </summary>
        WebBottomFlangeRight = 3,
        /// <summary>
        /// A built up cross section with the bottom flange pointing to the right and the web coming off the bottom flange.
        /// </summary>
        WebBuiltUpBottomFlangeRight = 4,
        /// <summary>
        /// A C type cross section where the top and bottom flanges point to the right.
        /// </summary>
        WebTopAndBottomRightFlanges = 5,
        /// <summary>
        /// A T type cross section where the top flange points to the left and right.
        /// </summary>
        WebTopFlange = 6,
        /// <summary>
        /// A 2L type cross section where the bottom flange points to the left and right.
        /// </summary>
        WebBottomFlange = 7,
        /// <summary>
        /// A rotated L type cross section with the top flange pointing to the left.
        /// </summary>
        WebTopFlangeLeft = 8,
        /// <summary>
        /// A built up cross section mirrored with the top flange pointing to the left and the web coming off the top flange.
        /// </summary>
        WebBuitUpTopFlangeLeft = 9,
        /// <summary>
        /// A mirrored L type cross section with the bottom flange pointing to the left.
        /// </summary>
        WebBottomFlangeLeft = 10,
        /// <summary>
        /// A built up cross section mirrored with the bottom flange pointing to the left and the web coming off the bottom flange. 
        /// </summary>
        WebBuiltUpBottomFlangeLeft = 11,
        /// <summary>
        /// A mirrored C type cross section where the top and bottom flanges point to the left.
        /// </summary>
        WebTopAndBottomLeftFlanges = 12,
        /// <summary>
        /// An I type cross section where the top and bottom flanges point to the left and right.
        /// </summary>
        WebTopAndBottomFlanges = 13,
        /// <summary>
        /// A rotated C type cross section with the top and bottom flanges pointing to the bottom.
        /// </summary>
        FlangeLeftAndRightBottomWebs = 14,
        /// <summary>
        /// A mirrored and rotated C type cross section with the top and bottom flanges pointing to the top.
        /// </summary>
        FlangeLeftAndRightTopWebs = 15,
        /// <summary>
        /// A rotated I type cross section with the top and bottom flanges pointing to the top and bottom.
        /// </summary>
        FlangeLeftAndRightWebs = 16,
        /// <summary>
        /// A rectangular type cross section.
        /// </summary>
        TwoWebsTwoFlanges = 17,
        /// <summary>
        /// A rectangular type cross section with the left and right webs coming off the top and bottom flanges.
        /// </summary>
        TwoFlangesBetweenWebs = 18,
        /// <summary>
        /// A rectangular type cross section with the top and bottom flanges extending past the left and right webs.
        /// </summary>
        TwoWebsBetweenFlanges = 19,
        /// <summary>
        /// A circular type cross section.
        /// </summary>
        Tube = 20
    }
   
    /// <summary>
    /// The SectionAlias class is a helper class which will determine the type of section based 
    /// on the section edges.
    /// </summary>
    internal static class SectionAlias
    {
        /// <summary>
        /// Gets the section alias.
        /// </summary>
        /// <param name="mappedEdges">The mapped edges.</param>
        /// <returns>Returns the appropriate section alias based on the mapped edges.</returns>
        internal static SectionAliasType GetSectionAlias(Dictionary<SectionFaceType, SectionFaceType> mappedEdges)
        {
            if (mappedEdges.ContainsKey(SectionFaceType.Top_Flange_Right) || mappedEdges.ContainsKey(SectionFaceType.Top_Flange_Left))
            {
                return GetAliasThatContainsTopFlange(mappedEdges);
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Bottom_Flange_Right) || mappedEdges.ContainsKey(SectionFaceType.Bottom_Flange_Left))
            {
                return GetAliasThatContainsBottomFlange(mappedEdges);
            }
            else
            {
                return GetAliasThatContainsInner(mappedEdges);
            }
        }

        /// <summary>
        /// Gets the alias that contains top flanges.
        /// </summary>
        /// <param name="mappedEdges">The mapped edges.</param>
        /// <returns>Returns the appropriate section alias based on sections that contain top flanges.</returns>
        private static SectionAliasType GetAliasThatContainsTopFlange(Dictionary<SectionFaceType, SectionFaceType> mappedEdges)
        {
            if (mappedEdges.ContainsKey(SectionFaceType.Top_Flange_Right) && mappedEdges.ContainsKey(SectionFaceType.Bottom_Flange_Right))
            {
                if (mappedEdges.ContainsKey(SectionFaceType.Top_Flange_Left) && mappedEdges.ContainsKey(SectionFaceType.Bottom_Flange_Left))
                {
                    if (mappedEdges.ContainsKey(SectionFaceType.Inner_Web_Right))
                    {
                        //since it contains an inner web (here it could be left or right does not matter)
                        //then it must be a box type cross section
                        return SectionAliasType.TwoWebsBetweenFlanges;
                    }
                    else
                    {
                        //is an I type cross section
                        return SectionAliasType.WebTopAndBottomFlanges;
                    }
                }
                else
                {
                    //is a C type cross section pointing to the right
                    return SectionAliasType.WebTopAndBottomRightFlanges;
                }
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Top_Flange_Left) && mappedEdges.ContainsKey(SectionFaceType.Bottom_Flange_Left))
            {
                //is a C type cross section pointing to the left
                return SectionAliasType.WebTopAndBottomLeftFlanges;
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Top_Flange_Right) && mappedEdges.ContainsKey(SectionFaceType.Top_Flange_Left))
            {
                //is a T type cross section
                return SectionAliasType.WebTopFlange;
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Top_Flange_Right))
            {
                if (mappedEdges.ContainsKey(SectionFaceType.Web_Right_Top))
                {
                    //is a built up type cross section with the web sticking out of the top flange and flange pointing to the right
                    return SectionAliasType.WebBuiltUpTopFlangeRight;
                }
                else
                {
                    //is an L type cross section mirrored and up-side-down
                    return SectionAliasType.WebTopFlangeRight;
                }
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Top_Flange_Left))
            {
                if (mappedEdges.ContainsKey(SectionFaceType.Web_Left_Top))
                {
                    //is a built up type cross section with the web sticking out of the top flange and flange pointing to the left
                    return SectionAliasType.WebBuitUpTopFlangeLeft;
                }
                else
                {
                    //is an L type cross section up-side-down
                    return SectionAliasType.WebTopFlangeLeft;
                }
            }

            //return unknown if it cannot determine what type of TopFlange cross section this is.
            return SectionAliasType.UnknownAlias;
        }

        /// <summary>
        /// Gets the alias that contains bottom flange.
        /// </summary>
        /// <param name="mappedEdges">The mapped edges.</param>
        /// <returns>Returns the appropriate section alias based on sections that contain bottom flanges.</returns>
        private static SectionAliasType GetAliasThatContainsBottomFlange(Dictionary<SectionFaceType, SectionFaceType> mappedEdges)
        {
            //I and C type cross sections should have been detected by GetAliasThatContainsTopFlange() method.  Here we need to look
            //at 2L and L type cross sections.

            if (mappedEdges.ContainsKey(SectionFaceType.Bottom_Flange_Right) && mappedEdges.ContainsKey(SectionFaceType.Bottom_Flange_Left))
            {
                //is a 2L type cross section
                return SectionAliasType.WebBottomFlange;
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Bottom_Flange_Right))
            {
                //if it has a web right bottom it could be a built up type cross section, or a T, or a 2L type cross section
                if (mappedEdges.ContainsKey(SectionFaceType.Web_Right_Bottom))
                {
                    //for T web_left will be mapped to top for 2L web_left will be mapped to bottom
                    foreach (var pair in mappedEdges)
                    {
                        if (pair.Key.Equals(SectionFaceType.Web_Left) && pair.Value.Equals(SectionFaceType.Top))
                        {
                            //is a T which is pointing to the right
                            return SectionAliasType.WebBuiltUpTopFlangeRight;
                        }
                        else if (pair.Key.Equals(SectionFaceType.Web_Left) && pair.Value.Equals(SectionFaceType.Bottom))
                        {
                            //is a 2L which is pointing to the right
                            return SectionAliasType.WebBuiltUpBottomFlangeRight;
                        }
                    }

                    //it is not a T or an 2L type so,
                    //it's a built up type cross section with the web sticking out of the bottom flange and flange pointing to the right
                    return SectionAliasType.WebBuiltUpBottomFlangeRight;
                }
                else
                {
                    //is an L type cross section
                    return SectionAliasType.WebBottomFlangeRight;
                }
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Bottom_Flange_Left))
            {
                if (mappedEdges.ContainsKey(SectionFaceType.Web_Left_Bottom))
                {
                    //is a built up type cross section with the web sticking out of the bottom flange and flange pointing to the left
                    return SectionAliasType.WebBuiltUpBottomFlangeLeft;
                }
                else
                {
                    //is a mirrored L type cross section
                    return SectionAliasType.WebBottomFlangeLeft;
                }
            }

            //return unknown if it cannot determine what type of TopFlange cross section this is.
            return SectionAliasType.UnknownAlias;
        }

        /// <summary>
        /// Gets the alias that contains inner.
        /// </summary>
        /// <param name="mappedEdges">The mapped edges.</param>
        /// <returns>Returns the appropriate section alias based on sections that contain inner webs.</returns>
        private static SectionAliasType GetAliasThatContainsInner(Dictionary<SectionFaceType, SectionFaceType> mappedEdges)
        {
            if (mappedEdges.ContainsKey(SectionFaceType.Inner_Top) && mappedEdges.ContainsKey(SectionFaceType.Right_Web_Top))
            {
                //is a box type cross section standing on its webs 
                return SectionAliasType.TwoFlangesBetweenWebs;
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Inner_Top))
            {
                //is a box type cross section.  As of now this is not being mapped.  A box will return as a Web.
                //here just so if in the future they decide to mapp inners for rectangular cross sections.
                return SectionAliasType.TwoWebsTwoFlanges;
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Right_Web_Top) && !mappedEdges.ContainsKey(SectionFaceType.Right_Web_Bottom))
            {
                //is a C type cross section laying on the left web ( |__| )
                return SectionAliasType.FlangeLeftAndRightTopWebs;
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Left_Web_Bottom) && !mappedEdges.ContainsKey(SectionFaceType.Left_Web_Top))
            {
                //is a C type cross section standing on its webs
                return SectionAliasType.FlangeLeftAndRightBottomWebs;
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Right_Web_Top) && mappedEdges.ContainsKey(SectionFaceType.Left_Web_Bottom))
            {
                //is an I type cross section stanging on its webs
                return SectionAliasType.FlangeLeftAndRightWebs;
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Right_Web_Top) && mappedEdges.ContainsKey(SectionFaceType.Right_Web_Bottom))
            {
                //is a T (or 2L) type cross section on its side with bottom pointing to right. [if 2L top pointing to the right]                

                //if it has an inner web right top then the the flange is pointing to the left
                if (mappedEdges.ContainsKey(SectionFaceType.Inner_Web_Right_Top))
                {
                    foreach (var pair in mappedEdges)
                    {
                        if (pair.Key.Equals(SectionFaceType.Web_Left) && pair.Value.Equals(SectionFaceType.Top))
                        {
                            //is a 2L which is pointing to the left
                            return SectionAliasType.WebBuiltUpBottomFlangeLeft;
                        }
                        else if (pair.Key.Equals(SectionFaceType.Web_Left) && pair.Value.Equals(SectionFaceType.Bottom))
                        {
                            //is a T which is pointing to the Left
                            return SectionAliasType.WebBuitUpTopFlangeLeft;
                        }
                    }
                }

                //if it does not have an inner web right top then the flange must be pointing to the right                
                //Should not get down to this return.
                return SectionAliasType.WebBuiltUpTopFlangeRight;
            }
            else if (mappedEdges.ContainsKey(SectionFaceType.Outer_Tube))
            {
                return SectionAliasType.Tube;
            }
            else
            {
                return SectionAliasType.Web;
            }
        }
    }
}
