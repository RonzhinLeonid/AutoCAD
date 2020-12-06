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
        public void BlockFrame()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor edit = doc.Editor;
            PromptEntityOptions entityOpt =
              new PromptEntityOptions("\nВыберите примитив: ");
            PromptEntityResult entityRes = edit.GetEntity(entityOpt);

            if (entityRes.Status != PromptStatus.OK) return;

            Extents3d blockFrame =
              new Extents3d(Point3d.Origin, Point3d.Origin);

            using (Entity element = entityRes.ObjectId.Open(OpenMode.ForRead) as Entity)
            {
                GetBlockExtents(element, ref blockFrame);

            }
            
            string s =
              "MinPoint: " + blockFrame.MinPoint.ToString() + " " +
              "MaxPoint: " + blockFrame.MaxPoint.ToString();
            edit.WriteMessage(s);
            //------------------------------------------------------------
            // Только для тестирования полученного габаритного контейнера
            //------------------------------------------------------------
            #region TestinExts
            using (BlockTableRecord curSpace =
              doc.Database.CurrentSpaceId.Open(OpenMode.ForWrite) as BlockTableRecord)
            {
                Point3dCollection pts = new Point3dCollection();
                pts.Add(blockFrame.MinPoint);
                pts.Add(new Point3d(blockFrame.MinPoint.X,
                  blockFrame.MaxPoint.Y, blockFrame.MinPoint.Z));
                pts.Add(blockFrame.MaxPoint);
                pts.Add(new Point3d(blockFrame.MaxPoint.X,
                  blockFrame.MinPoint.Y, blockFrame.MinPoint.Z));
                using (Polyline3d poly =
                  new Polyline3d(Poly3dType.SimplePoly, pts, true))
                {
                    curSpace.AppendEntity(poly);
                }
            }
            #endregion

        }
        /// <summary>
        /// Рекурсивное получение габаритного контейнера блока.
        /// </summary>
        /// <param name="element">Имя примитива</param>
        /// <param name="boundaries">Габаритный контейнер</param>
        /// <param name="mat">Матрица преобразования из системы координат блока в МСК.</param>
        void GetBlockExtents(Entity element, ref Extents3d boundaries)
        {
            if (!IsLayerOn(element.LayerId))
                return;
            if (element is BlockReference)
            {
                BlockReference blockRef = element as BlockReference;
                using (BlockTableRecord blockTR =
                  blockRef.BlockTableRecord.Open(OpenMode.ForRead) as BlockTableRecord)
                {
                    foreach (ObjectId id in blockTR)
                    {
                        using (DBObject obj = id.Open(OpenMode.ForRead) as DBObject)
                        {
                            Entity entityCur = obj as Entity;
                            if (entityCur == null || entityCur.Visible != true)
                                continue;
                            // Пропускаем неконстантные и невидимые определения атрибутов
                            AttributeDefinition attDef = entityCur as AttributeDefinition;
                            if (attDef != null)
                                continue;
                            GetBlockExtents(entityCur, ref boundaries);
                        }
                    }
                }
            }
            else
            {
                if (IsEmptyBound(ref boundaries))
                {
                    try { boundaries = (Extents3d)element.Bounds; } catch { };
                }
                else
                {
                    try { boundaries.AddExtents((Extents3d)element.Bounds); } catch { };
                }
                return;
            }
            return;
        }
        /// <summary>
        /// Определяет видны ли объекты на слое, указанном его ObjectId и принадлежность "не нужным" слоям.
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
                                                            "i_Шильд",
                                                            "Зона сверления"};
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
        /// <param name="bound">Габаритный контейнер.</param>
        /// <returns></returns>
        bool IsEmptyBound(ref Extents3d bound)
        {
            if (bound.MinPoint.DistanceTo(bound.MaxPoint) < Tolerance.Global.EqualPoint)
                return true;
            else
                return false;
        }
    }
}
