using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Collections;

namespace AcadTextPlacement
{
    public class TextMark
    {

        public class DataSaver
        {
            public const string key = "scaleDef";

            public int scale_saved;
            public DataSaver()
            {
                scale_saved = 1000;
            }
        }
        
        [CommandMethod("TextMark", CommandFlags.UsePickSet)]
        public void GatherText()
        {
            Document doc =
                Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Hashtable ud = doc.UserData;
            DataSaver ds;

            // Initiate previously used scale saver
            ds = ud[DataSaver.key] as DataSaver;
            if (ds == null)
            {
                object obj = ud[DataSaver.key];
                if (obj == null)
                {
                    ds = new DataSaver();
                    ud.Add(DataSaver.key, ds);
                }
            }

            // Ask user to select entities

            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nВыберите полининии: ";
            pso.AllowDuplicates = false;
            pso.AllowSubSelections = true;
            pso.RejectObjectsFromNonCurrentSpace = true;
            pso.RejectObjectsOnLockedLayers = false;
        
            PromptSelectionResult psr = ed.GetSelection(pso);
            if (psr.Status != PromptStatus.OK)
                return;

            // Ask user for marking text

            PromptStringOptions txt = new PromptStringOptions("Введите текст");
            txt.Message = "\nВведите текст для маркировки линий: ";
            txt.AllowSpaces = false;
            PromptResult txt_result = ed.GetString(txt);
            String txt_value = txt_result.StringResult;

            if (txt_result.Status != PromptStatus.OK)
                return;


            // Ask user for map scale

            PromptIntegerOptions scale = new PromptIntegerOptions("Введите масштаб");
            scale.Message = "\nВведите масштаб плана: ";
            scale.AllowNone = false;
            scale.Keywords.Add("500");
            scale.Keywords.Add("1000");
            scale.Keywords.Add("2000");
            scale.Keywords.Add("5000");
            scale.AppendKeywordsToMessage = true;
            scale.UseDefaultValue = true;
            scale.DefaultValue = ds.scale_saved;
            PromptIntegerResult scale_result = ed.GetInteger(scale);
            int scale_value = scale_result.Value;
            ds.scale_saved = scale_value;


            if (scale_result.Status != PromptStatus.OK)
                return;

            // Collect points on the component entities
        
            Point3dCollection pts = new Point3dCollection();

            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTableRecord btr =
                (BlockTableRecord)tr.GetObject(
                    db.CurrentSpaceId,
                    OpenMode.ForWrite
                );

                foreach (SelectedObject so in psr.Value)
                {
                    Entity ent =
                        (Entity)tr.GetObject(
                        so.ObjectId,
                        OpenMode.ForRead
                        );

                    // Collect the points for each selected entity

                    Point3dCollection entPts = new Point3dCollection();
                    List<double> rot = new List<double>();
                    CollectPoints(tr, ent, entPts, rot, scale_value);

                    // Add a physical DBPoint at each Point3d

                    for (int i=0; i < entPts.Count; i++)
                    {
                        DBText dbt = new DBText();
                        dbt.SetDatabaseDefaults();
                        dbt.Position = entPts[i];
                        dbt.Height = scale_value / 100;
                        dbt.Rotation = rot[i];
                        dbt.Justify = AttachmentPoint.MiddleCenter;
                        dbt.AlignmentPoint = entPts[i];
                        dbt.TextString = txt_value;
                        btr.AppendEntity(dbt);
                        tr.AddNewlyCreatedDBObject(dbt, true);
                    }
                }
                tr.Commit();
            }
        }
        private void CollectPoints(Transaction tr, Entity ent, Point3dCollection pts, List<double> rot, int scale_value)
        {
            
            // Operations with polyline
            // Iterate vertices and generate points for polyline segments with preset step (currently 100)

            Polyline pline = ent as Polyline;
            if (pline != null)
            {
                int nv = pline.NumberOfVertices;
                for (int i = 0; i < nv - 1; i++)
                {
                    try
                    {
                        Point3d pt1 = pline.GetPoint3dAt(i);
                        Point3d pt2 = pline.GetPoint3dAt(i+1);
                        double dx = pt2.X - pt1.X;
                        double dy = pt2.Y - pt1.Y;
                        double d = Math.Sqrt(dx*dx + dy*dy);

                        double rotation = 0.0;

                        if ((dx > 0) && (dy > 0))
                        {
                            rotation = Math.Acos(dx/d);
                        }
                        else if ((dx > 0) && (dy < 0))
                        {
                            rotation = 2*Math.PI - Math.Acos(dx/d);
                        }
                        else if ((dx < 0) && (dy > 0))
                        {
                            rotation = Math.Acos(dx/d) + Math.PI;
                        }
                        else if ((dx < 0) && (dy < 0))
                        {
                            rotation = Math.PI - Math.Acos(dx/d);
                        }
                        else if ((dx == 0) && (dy != 0))
                        {
                            rotation = Math.PI / 2;
                        }
                        else
                        {
                            rotation = Math.Acos(dx/d);
                        }

                        if (d > scale_value / 40)
                        
                        {
                            for (double j = scale_value / 33; j < d - scale_value / 33; j += scale_value / 10)
                            {
                                Point3d pt = new Point3d(pt1.X + dx * j / d, pt1.Y + dy * j / d, pt1.Z);
                                pts.Add(pt);
                                rot.Add(rotation);
                            }
                        }
                    }

                    catch { }
                }
            }
        }
    }
}