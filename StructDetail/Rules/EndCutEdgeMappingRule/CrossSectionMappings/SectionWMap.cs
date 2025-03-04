//-----------------------------------------------------------------------------
//      Copyright (C) 2010-12 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionWMap class is designed to return a cross section
//                  map based on section type W.  Other sections with similar
//                  shape as W will also use this class to return the appropriate map.
//
//      Author:  3XCalibur
//
//      History:
//      November 03, 2010       WR                  Created
//
//   October 11,2012     Alliagtors      CR-207934/DM-CP-220895(for v2013) 'pointMap' collection is filled appropriately.
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure.CrossSectionMappings
{
    /// <summary>
    /// The SectionWMap class is designed to return a cross section edge map based on I type shape cross sections.
    /// </summary>
    internal class SectionWMap : ICrossSectionMap
    {
        #region ICrossSectionMap Members

        /// <summary>
        /// Gets the cross section map.
        /// </summary>
        /// <param name="flipLeftAndRight">if set to <c>true</c> [is mirrored].</param>
        /// <param name="quadrant">The quadrant.</param>
        /// <returns>The cross section mapping of edges based on the quadrant and the flipLeftAndRight flag.</returns>
        /// <exception cref="S3DEndCutMappingException">Unable to map cross section.</exception>
        public Dictionary<SectionFaceType, SectionFaceType> GetCrossSectionMap(bool flipLeftAndRight, int quadrant, Dictionary<int, int> pointMap)
        {   
            Dictionary<SectionFaceType, SectionFaceType> mappedEdges = new Dictionary<SectionFaceType, SectionFaceType>();

            //      WR = Web_Right
            //      TFRB = Top_Flange_Right_Bottom
            //      TFR = Top_Flange_Right 
            //      Top               
            //      TFL = Top_Flange_Left
            //      TFLB = Top_Flange_Left_Bottom
            //      WL = Web_Left
            //      BFLT = Bottom_Flange_Left_Top
            //      BFL = Bottom_Flange_Left
            //      Bottom
            //      BFR = Bottom_Flange_Right
            //      BFRT = Bottom_Flange_Right_Top
            //
            //                       Top
            //           **************************
            //       TFL *                        * TFR
            //           ***********    ***********
            //               TFLB  *    *  TFRB
            //                     *    *
            //                     *    *  
            //                 WL  *    *  WR
            //                     *    *  
            //                     *    *
            //               BFLT  *    *  BFRT 
            //           ***********    ***********
            //       BFL *                        * BFR
            //           **************************
            //                     Bottom

            if (!flipLeftAndRight)
            {
                // Not mirrored case
                switch (quadrant)
                {
                    case 1:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Bottom_Flange_Right_Top);

                        for (int i = 1; i < 28; i++)
                        {
                            pointMap.Add(i, i);
                        }

                        break;
                    
                    case 2:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Right_Web_Top, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Top, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Left_Top, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Left_Web_Top, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Left_Web_Bottom, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Left_Bottom, SectionFaceType.Top_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Bottom, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Right_Web_Bottom, SectionFaceType.Bottom_Flange_Left);

                        pointMap.Add(1, 25);
                        pointMap.Add(2, 25);
                        pointMap.Add(3, 11);
                        pointMap.Add(4, 11);
                        pointMap.Add(5, 11);
                        pointMap.Add(6, 25);
                        pointMap.Add(7, 27);
                        pointMap.Add(8, 26);
                        pointMap.Add(9, 15);
                        pointMap.Add(10, 15);
                        pointMap.Add(11, 15);
                        pointMap.Add(12, 26);
                        pointMap.Add(13, 26);
                        pointMap.Add(14, 26);
                        pointMap.Add(15, 23);
                        pointMap.Add(16, 23);
                        pointMap.Add(17, 23);
                        pointMap.Add(18, 26);
                        pointMap.Add(19, 27);
                        pointMap.Add(20, 25);
                        pointMap.Add(21, 3);
                        pointMap.Add(22, 3);
                        pointMap.Add(23, 3);
                        pointMap.Add(24, 25);
                        pointMap.Add(25, 13);
                        pointMap.Add(26, 1);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Top_Flange_Left_Bottom);

                        pointMap.Add(1, 13);
                        pointMap.Add(2, 14);
                        pointMap.Add(3, 15);
                        pointMap.Add(4, 16);
                        pointMap.Add(5, 17);
                        pointMap.Add(6, 18);
                        pointMap.Add(7, 19);
                        pointMap.Add(8, 20);
                        pointMap.Add(9, 21);
                        pointMap.Add(10, 22);
                        pointMap.Add(11, 23);
                        pointMap.Add(12, 24);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 2);
                        pointMap.Add(15, 3);
                        pointMap.Add(16, 4);
                        pointMap.Add(17, 5);
                        pointMap.Add(18, 6);
                        pointMap.Add(19, 7);
                        pointMap.Add(20, 8);
                        pointMap.Add(21, 9);
                        pointMap.Add(22, 10);
                        pointMap.Add(23, 11);
                        pointMap.Add(24, 12);
                        pointMap.Add(25, 26);
                        pointMap.Add(26, 25);
                        pointMap.Add(27, 27);
                        
                        break;
                    
                    case 4:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Right_Web_Top, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Top, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Left_Top, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Left_Web_Top, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Left_Web_Bottom, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Left_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Bottom, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Right_Web_Bottom, SectionFaceType.Top_Flange_Right);

                        pointMap.Add(1, 26);
                        pointMap.Add(2, 26);
                        pointMap.Add(3, 23);
                        pointMap.Add(4, 23);
                        pointMap.Add(5, 23);
                        pointMap.Add(6, 26);
                        pointMap.Add(7, 27);
                        pointMap.Add(8, 25);
                        pointMap.Add(9, 3);
                        pointMap.Add(10, 3);
                        pointMap.Add(11, 3);
                        pointMap.Add(12, 25);
                        pointMap.Add(13, 25);
                        pointMap.Add(14, 25);
                        pointMap.Add(15, 11);
                        pointMap.Add(16, 11);
                        pointMap.Add(17, 11);
                        pointMap.Add(18, 25);
                        pointMap.Add(19, 27);
                        pointMap.Add(20, 26);
                        pointMap.Add(21, 15);
                        pointMap.Add(22, 15);
                        pointMap.Add(23, 15);
                        pointMap.Add(24, 26);
                        pointMap.Add(25, 1);
                        pointMap.Add(26, 13);
                        pointMap.Add(27, 27);                        
                        
                        break;
                    
                    default:
                        throw new EdgeMappingException(quadrant);
                }
            }
            else
            {
                switch (quadrant)
                {
                    case 1:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Bottom_Flange_Left_Top);

                        pointMap.Add(1, 1);
                        pointMap.Add(2, 24);
                        pointMap.Add(3, 23);
                        pointMap.Add(4, 22);
                        pointMap.Add(5, 21);
                        pointMap.Add(6, 20);
                        pointMap.Add(7, 19);
                        pointMap.Add(8, 18);
                        pointMap.Add(9, 17);
                        pointMap.Add(10, 16);
                        pointMap.Add(11, 15);
                        pointMap.Add(12, 14);
                        pointMap.Add(13, 13);
                        pointMap.Add(14, 12);
                        pointMap.Add(15, 11);
                        pointMap.Add(16, 10);
                        pointMap.Add(17, 9);
                        pointMap.Add(18, 8);
                        pointMap.Add(19, 7);
                        pointMap.Add(20, 6);
                        pointMap.Add(21, 5);
                        pointMap.Add(22, 4);
                        pointMap.Add(23, 3);
                        pointMap.Add(24, 2);
                        pointMap.Add(25, 26);
                        pointMap.Add(26, 25);
                        pointMap.Add(27, 27);
                        
                        break;
                    
                    case 2:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Right_Web_Top, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Top, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Left_Top, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Left_Web_Top, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Left_Web_Bottom, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Left_Bottom, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Right_Web_Bottom, SectionFaceType.Bottom_Flange_Right);

                        pointMap.Add(1, 26);
                        pointMap.Add(2, 26);
                        pointMap.Add(3, 15);
                        pointMap.Add(4, 15);
                        pointMap.Add(5, 15);
                        pointMap.Add(6, 26);
                        pointMap.Add(7, 27);
                        pointMap.Add(8, 25);
                        pointMap.Add(9, 11);
                        pointMap.Add(10, 11);
                        pointMap.Add(11, 11);
                        pointMap.Add(12, 25);
                        pointMap.Add(13, 25);
                        pointMap.Add(14, 25);
                        pointMap.Add(15, 3);
                        pointMap.Add(16, 3);
                        pointMap.Add(17, 3);
                        pointMap.Add(18, 25);
                        pointMap.Add(19, 27);
                        pointMap.Add(20, 26);
                        pointMap.Add(21, 23);
                        pointMap.Add(22, 23);
                        pointMap.Add(23, 23);
                        pointMap.Add(24, 26);
                        pointMap.Add(25, 13);
                        pointMap.Add(26, 1);
                        pointMap.Add(27, 27);
                        
                        break;
                    
                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Top_Flange_Right_Bottom);

                        pointMap.Add(1, 13);
                        pointMap.Add(2, 12);
                        pointMap.Add(3, 11);
                        pointMap.Add(4, 10);
                        pointMap.Add(5, 9);
                        pointMap.Add(6, 8);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 6);
                        pointMap.Add(9, 5);
                        pointMap.Add(10, 4);
                        pointMap.Add(11, 3);
                        pointMap.Add(12, 2);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 24);
                        pointMap.Add(15, 23);
                        pointMap.Add(16, 22);
                        pointMap.Add(17, 21);
                        pointMap.Add(18, 20);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 18);
                        pointMap.Add(21, 17);
                        pointMap.Add(22, 16);
                        pointMap.Add(23, 15);
                        pointMap.Add(24, 14);
                        pointMap.Add(25, 25);
                        pointMap.Add(26, 26);
                        pointMap.Add(27, 27);
                        
                        break;
                    
                    case 4:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Right_Web_Top, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Top, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Left_Top, SectionFaceType.Bottom_Flange_Right_Top);
                        mappedEdges.Add(SectionFaceType.Left_Web_Top, SectionFaceType.Bottom_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Left_Web_Bottom, SectionFaceType.Bottom_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Left_Bottom, SectionFaceType.Bottom_Flange_Left_Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Inner_Web_Right_Bottom, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Right_Web_Bottom, SectionFaceType.Top_Flange_Left);

                        pointMap.Add(1, 25);
                        pointMap.Add(2, 25);
                        pointMap.Add(3, 3);
                        pointMap.Add(4, 3);
                        pointMap.Add(5, 3);
                        pointMap.Add(6, 25);
                        pointMap.Add(7, 27);
                        pointMap.Add(8, 26);
                        pointMap.Add(9, 23);
                        pointMap.Add(10, 23);
                        pointMap.Add(11, 23);
                        pointMap.Add(12, 26);
                        pointMap.Add(13, 26);
                        pointMap.Add(14, 26);
                        pointMap.Add(15, 15);
                        pointMap.Add(16, 15);
                        pointMap.Add(17, 15);
                        pointMap.Add(18, 26);
                        pointMap.Add(19, 27);
                        pointMap.Add(20, 25);
                        pointMap.Add(21, 11);
                        pointMap.Add(22, 11);
                        pointMap.Add(23, 11);
                        pointMap.Add(24, 25);
                        pointMap.Add(25, 1);
                        pointMap.Add(26, 13);
                        pointMap.Add(27, 27);                        
                        
                        break;
                    
                    default:
                        throw new EdgeMappingException(quadrant);
                }
            }

            return mappedEdges;            
        }

        #endregion
    }
}
