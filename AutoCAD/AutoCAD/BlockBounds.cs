using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;

namespace AutoCAD
{
    public class BlockBounds
    {
        [CommandMethod("BlockBounds")]
        public void BlockExt()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;
            PromptEntityOptions enOpt =
              new PromptEntityOptions("\nВыберите примитив: ");
            PromptEntityResult enRes = ed.GetEntity(enOpt);

            if (enRes.Status != PromptStatus.OK) return;

            Extents3d blockExt =
              new Extents3d(Point3d.Origin, Point3d.Origin);
            Matrix3d mat = Matrix3d.Identity;

            using (Entity en = enRes.ObjectId.Open(OpenMode.ForRead) as Entity)
            {
                GetBlockExtents(en, ref blockExt, ref mat);

            }
            
            string s =
              "MinPoint: " + blockExt.MinPoint.ToString() + " " +
              "MaxPoint: " + blockExt.MaxPoint.ToString();
            ed.WriteMessage(s);
            //------------------------------------------------------------
            // Только для тестирования полученного габаритного контейнера
            //------------------------------------------------------------
            #region TestinExts
            using (BlockTableRecord curSpace =
              doc.Database.CurrentSpaceId.Open(OpenMode.ForWrite) as BlockTableRecord)
            {
                Point3dCollection pts = new Point3dCollection();
                pts.Add(blockExt.MinPoint);
                pts.Add(new Point3d(blockExt.MinPoint.X,
                  blockExt.MaxPoint.Y, blockExt.MinPoint.Z));
                pts.Add(blockExt.MaxPoint);
                pts.Add(new Point3d(blockExt.MaxPoint.X,
                  blockExt.MinPoint.Y, blockExt.MinPoint.Z));
                using (Polyline3d poly =
                  new Polyline3d(Poly3dType.SimplePoly, pts, true))
                {
                    curSpace.AppendEntity(poly);
                }
            }
            #endregion

        }
        /// <summary>
        /// Рекурсивное получение габаритного контейнера для вставки блока.
        /// </summary>
        /// <param name="en">Имя примитива</param>
        /// <param name="ext">Габаритный контейнер</param>
        /// <param name="mat">Матрица преобразования из системы координат блока в МСК.</param>
        void GetBlockExtents(Entity en, ref Extents3d ext, ref Matrix3d mat)
        {
            if (!IsLayerOn(en.LayerId))
                return;
            if (en is BlockReference)
            {
                BlockReference bref = en as BlockReference;
                Matrix3d matIns = mat * bref.BlockTransform;
                using (BlockTableRecord btr =
                  bref.BlockTableRecord.Open(OpenMode.ForRead) as BlockTableRecord)
                {
                    foreach (ObjectId id in btr)
                    {
                        using (DBObject obj = id.Open(OpenMode.ForRead) as DBObject)
                        {
                            Entity enCur = obj as Entity;
                            if (enCur == null || enCur.Visible != true)
                                continue;
                            // Пропускаем неконстантные и невидимые определения атрибутов
                            AttributeDefinition attDef = enCur as AttributeDefinition;
                            if (attDef != null && (!attDef.Constant || attDef.Invisible))
                                continue;
                            GetBlockExtents(enCur, ref ext, ref matIns);
                        }
                    }
                }
                // Отдельно обрабатываем атрибуты блока
                //if (bref.AttributeCollection.Count > 0)
                //{
                //    foreach (ObjectId idAtt in bref.AttributeCollection)
                //    {
                //        using (AttributeReference attRef =
                //          idAtt.Open(OpenMode.ForRead) as AttributeReference)
                //        {
                //            if (!attRef.Invisible && attRef.Visible)
                //                GetBlockExtents(attRef, ref ext, ref mat);
                //        }
                //    }
                //}
            }
            else
            {
                if (mat.IsUniscaledOrtho())
                {
                    using (Entity enTr = en.GetTransformedCopy(mat))
                    {
                        if (IsEmptyExt(ref ext))
                        {
                            try { ext = enTr.GeometricExtents; } catch { };
                        }
                        else
                        {
                            try { ext.AddExtents(enTr.GeometricExtents); } catch { };
                        }
                        return;
                    }
                }
                else
                {
                    try
                    {
                        Extents3d curExt = en.GeometricExtents;
                        curExt.TransformBy(mat);
                        if (IsEmptyExt(ref ext))
                            ext = curExt;
                        else
                            ext.AddExtents(curExt);
                    }
                    catch { }
                    return;
                }
            }

            return;

        }
        /// <summary>
        /// Определяет видны ли объекты на слое, указанном его ObjectId.
        /// </summary>
        /// <param name="layerId">ObjectId слоя.</param>
        /// <returns></returns>
        bool IsLayerOn(ObjectId layerId)
        {
            List<string> listLayer = new List<string>() {   "Осевая",
                                                            "Осевая линия(ГОСТ)",
                                                            "i_Осевые линии",
                                                            "i_Отверстия",
                                                            "i_Резьбы",
                                                            "i_Шильд"};
            using (LayerTableRecord ltr = layerId.Open(OpenMode.ForRead) as LayerTableRecord)
            {
                if (ltr.IsFrozen) return false;  
                if (ltr.IsOff) return false;
                if (ltr.Name == "Defpoints") return false;
                if (listLayer.BinarySearch(ltr.Name) >= 0) return false;
            }
            return true;
        }
        /// <summary>
        /// Определят не пустой ли габаритный контейнер.
        /// </summary>
        /// <param name="ext">Габаритный контейнер.</param>
        /// <returns></returns>
        bool IsEmptyExt(ref Extents3d ext)
        {
            if (ext.MinPoint.DistanceTo(ext.MaxPoint) < Tolerance.Global.EqualPoint)
                return true;
            else
                return false;
        }
    }
}
