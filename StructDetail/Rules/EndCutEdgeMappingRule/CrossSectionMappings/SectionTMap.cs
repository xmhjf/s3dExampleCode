//-----------------------------------------------------------------------------
//      Copyright (C) 2010-12 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionTMap class is designed to return a cross section
//                  map based on section type T.  Other sections with similar
//                  shape as T will also use this class to return the appropriate map.
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
    /// The SectionTMap class is designed to return a cross section edge map based on T type shape cross sections.
    /// </summary>
    internal class SectionTMap : ICrossSectionMap
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
            //      Bottom
            //
            //                       Top
            //           **************************
            //       TFL *                        * TFR
            //           ***********    ***********
            //               TFLB  *    *  TFRB
            //                     *    *
            //                 WL  *    *  WR
            //                     *    *
            //                     *    *  
            //                     *    *
            //                     *    *  
            //                     ******
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
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);

                        for (int i = 1; i < 28; i++)
                        {
                            pointMap.Add(i, i);
                        }                        
                        
                        break;
                    
                    case 2:
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Web_Right_Top, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Bottom);

                        pointMap.Add(1, 10);
                        pointMap.Add(2, 11);
                        pointMap.Add(3, 11);
                        pointMap.Add(4, 11);
                        pointMap.Add(5, 11);
                        pointMap.Add(6, 11);
                        pointMap.Add(7, 13);
                        pointMap.Add(8, 15);
                        pointMap.Add(9, 15);
                        pointMap.Add(10, 15);
                        pointMap.Add(11, 15);
                        pointMap.Add(12, 15);
                        pointMap.Add(13, 16);
                        pointMap.Add(14, 17);
                        pointMap.Add(15, 20);
                        pointMap.Add(16, 1);
                        pointMap.Add(17, 2);
                        pointMap.Add(18, 8);
                        pointMap.Add(19, 1);
                        pointMap.Add(20, 9);
                        pointMap.Add(21, 9);
                        pointMap.Add(22, 9);
                        pointMap.Add(23, -1);
                        pointMap.Add(24, 9);
                        pointMap.Add(25, 13);
                        pointMap.Add(26, 1);
                        pointMap.Add(27, 27);
                        pointMap.Add(50, 18);

                        break;
                    
                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left_Top, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Left, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Bottom_Flange_Right_Top, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        
                        pointMap.Add(1, 13);
                        pointMap.Add(2, 14);
                        pointMap.Add(3, 15);
                        pointMap.Add(4, 16);
                        pointMap.Add(5, 17);
                        pointMap.Add(6, 18);
                        pointMap.Add(7, 19);
                        pointMap.Add(8, 20);
                        pointMap.Add(9, 20);
                        pointMap.Add(10, 20);
                        pointMap.Add(11, 20);
                        pointMap.Add(12, 20);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 2);
                        pointMap.Add(15, 2);
                        pointMap.Add(16, 2);
                        pointMap.Add(17, 2);
                        pointMap.Add(18, 2);
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
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Web_Left_Top, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Top, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Bottom);
                        
                        pointMap.Add(1, 16);
                        pointMap.Add(2, 17);
                        pointMap.Add(3, -1);
                        pointMap.Add(4, 17);
                        pointMap.Add(5, 17);
                        pointMap.Add(6, 17);
                        pointMap.Add(7, 1);
                        pointMap.Add(8, 18);
                        pointMap.Add(9, 20);
                        pointMap.Add(10, 1);
                        pointMap.Add(11, 2);
                        pointMap.Add(12, 9);
                        pointMap.Add(13, 10);
                        pointMap.Add(14, 11);
                        pointMap.Add(15, 11);
                        pointMap.Add(16, 11);
                        pointMap.Add(17, 11);
                        pointMap.Add(18, 11);
                        pointMap.Add(19, 13);
                        pointMap.Add(20, 15);
                        pointMap.Add(21, 15);
                        pointMap.Add(22, 15);
                        pointMap.Add(23, 15);
                        pointMap.Add(24, 15);
                        pointMap.Add(25, 1);
                        pointMap.Add(26, 13);
                        pointMap.Add(27, 27);
                        pointMap.Add(52, 8);

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
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);
                        
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
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Bottom, SectionFaceType.Web_Right);                        
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right, SectionFaceType.Bottom);                        
                        mappedEdges.Add(SectionFaceType.Top_Flange_Right_Top, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Web_Right_Top, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top_Flange_Left);
                        
                        pointMap.Add(1, 16);
                        pointMap.Add(2, 15);
                        pointMap.Add(3, 15);
                        pointMap.Add(4, 15);
                        pointMap.Add(5, 15);
                        pointMap.Add(6, 15);
                        pointMap.Add(7, 13);
                        pointMap.Add(8, 11);
                        pointMap.Add(9, 11);
                        pointMap.Add(10, 11);
                        pointMap.Add(11, 11);
                        pointMap.Add(12, 11);
                        pointMap.Add(13, 10);
                        pointMap.Add(14, 9);
                        pointMap.Add(15, 2);
                        pointMap.Add(16, 1);
                        pointMap.Add(17, 20);
                        pointMap.Add(18, 18);
                        pointMap.Add(19, 1);
                        pointMap.Add(20, 17);
                        pointMap.Add(21, 17);
                        pointMap.Add(22, 17);
                        pointMap.Add(23, -1);
                        pointMap.Add(24, 17);
                        pointMap.Add(25, 13);
                        pointMap.Add(26, 1);
                        pointMap.Add(27, 27);
                        pointMap.Add(50, 8);

                        break;
                    
                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);                        
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);                        
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
                        pointMap.Add(8, 2);
                        pointMap.Add(9, 2);
                        pointMap.Add(10, 2);
                        pointMap.Add(11, 2);
                        pointMap.Add(12, 2);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 20);
                        pointMap.Add(15, 20);
                        pointMap.Add(16, 20);
                        pointMap.Add(17, 20);
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
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top_Flange_Left_Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top_Flange_Left);
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top_Flange_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left_Top, SectionFaceType.Top_Flange_Right_Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top_Flange_Left_Bottom, SectionFaceType.Web_Left);
                        
                        pointMap.Add(1, 10);
                        pointMap.Add(2, 9);
                        pointMap.Add(3, -1);
                        pointMap.Add(4, 9);
                        pointMap.Add(5, 9);
                        pointMap.Add(6, 9);
                        pointMap.Add(7, 1);
                        pointMap.Add(8, 8);
                        pointMap.Add(9, 2);
                        pointMap.Add(10, 1);
                        pointMap.Add(11, 20);
                        pointMap.Add(12, 17);
                        pointMap.Add(13, 16);
                        pointMap.Add(14, 15);
                        pointMap.Add(15, 15);
                        pointMap.Add(16, 15);
                        pointMap.Add(17, 15);
                        pointMap.Add(18, 15);
                        pointMap.Add(19, 13);
                        pointMap.Add(20, 11);
                        pointMap.Add(21, 11);
                        pointMap.Add(22, 11);
                        pointMap.Add(23, 11);
                        pointMap.Add(24, 11);
                        pointMap.Add(25, 1);
                        pointMap.Add(26, 13);
                        pointMap.Add(27, 27);
                        pointMap.Add(52, 18);
                        
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
