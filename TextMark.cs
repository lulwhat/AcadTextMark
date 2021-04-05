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
        // class that saves previosly used options for TextMark
        public class DataSaver
        {
            public const string scaleKey = "scaleDef";

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
            DataSaver dss;

            // Initiate data saver
            dss = ud[DataSaver.scaleKey] as DataSaver;
            if (dss == null)
            {
                object obj = ud[DataSaver.scaleKey];
                if (obj == null)
                {
                    dss = new DataSaver();
                    ud.Add(DataSaver.scaleKey, dss);
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
            txt.AllowSpaces = true;
            PromptResult txt_result = ed.GetString(txt);
            String txt_value = txt_result.StringResult;

            if (txt_result.Status != PromptStatus.OK)
                return;
            
            // Ask user for text position

            PromptKeywordOptions pos = new PromptKeywordOptions(
                "\nВыберите положение текста относительно линии[Центр/Верх/Низ]:",
                "Центр Верх Низ"
                );
            pos.Keywords.Default = "Центр";
            PromptResult pos_result = ed.GetKeywords(pos);
            String pos_value = pos_result.StringResult;

            if ((pos_result.Status != PromptStatus.OK) & (pos_result.Status != PromptStatus.Keyword))
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
            scale.DefaultValue = dss.scale_saved;
            PromptIntegerResult scale_result = ed.GetInteger(scale);
            int scale_value = scale_result.Value;
            dss.scale_saved = scale_value;


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
                    CollectPoints(tr, ent, entPts, rot, scale_value, pos_value);

                    // Add a physical DBPoint at each Point3d

                    for (int i=0; i < entPts.Count; i++)
                    {
                        DBText dbt = new DBText();
                        var textStyles = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                        dbt.SetDatabaseDefaults();
                        dbt.Position = entPts[i];
                        dbt.Height = 2 / 1000.0 * scale_value;
                        dbt.Rotation = rot[i];
                        try
                        {
                            dbt.TextStyleId = textStyles["SPDS"];
                        }
                        catch
                        {  }
                        if ((pos_value == "Центр") | (pos_value == "Ц"))
                        {
                            dbt.Justify = AttachmentPoint.MiddleCenter;
                        }    
                        else if ((pos_value == "Низ") | (pos_value == "Н"))
                        {
                            dbt.Justify = AttachmentPoint.TopLeft;
                        }
                        else
                        {
                            dbt.Justify = AttachmentPoint.BottomLeft;
                        }
                        dbt.AlignmentPoint = entPts[i];
                        dbt.TextString = txt_value;
                        btr.AppendEntity(dbt);
                        tr.AddNewlyCreatedDBObject(dbt, true);
                    }
                }
                tr.Commit();
            }
        }
        private void CollectPoints(Transaction tr, Entity ent, Point3dCollection pts, List<double> rot, int scale_value, String pos_value)
        {
            
            // Operations with polyline
            // Iterate vertices and generate points for polyline segments based on scale

            Polyline pline = ent as Polyline;
            if (pline != null)
            {
                int nv = pline.NumberOfVertices;
                double shift = scale_value / 40.0;
                for (int i = 0; i < nv - 1; i++)
                {
                    try
                    {
                        Point3d pt1 = pline.GetPoint3dAt(i);
                        Point3d pt2 = pline.GetPoint3dAt(i+1);
                        double dx = pt2.X - pt1.X;
                        double dy = pt2.Y - pt1.Y;
                        double d = Math.Sqrt(dx*dx + dy*dy);
                        double offset_d = 0.0;
                        double offset_x = 0.0;
                        double offset_y = 0.0;
                        // offset is added when text signs are either above or under the line
                        // cases of line aligment that affect x and y offset
                        // x+ y-
                        if ((dy < 0) & (dx > 0))
                        {
                            if ((pos_value == "Верх") | (pos_value == "В"))
                            {
                                offset_d = 0.6 + (0.002 * scale_value);
                                offset_x = Math.Cos(Math.PI / 2 - Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value;
                                offset_y = Math.Sin(Math.PI / 2 - Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value;
                            }
                            else if ((pos_value == "Низ") | (pos_value == "Н"))
                            {
                                offset_d = 0.6 + (0.002 * scale_value);
                                offset_x = -(Math.Cos(Math.PI / 2 - Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value);
                                offset_y = -(Math.Sin(Math.PI / 2 - Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value);
                            }
                        }
                        // x- y-
                        else if ((dy < 0) & (dx < 0))
                        {
                            if ((pos_value == "Верх") | (pos_value == "В"))
                            {
                                offset_d = -(0.6 + (0.002 * scale_value));
                                offset_x = Math.Cos(1.5 * Math.PI + Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value;
                                offset_y = Math.Sin(1.5 * Math.PI + Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value;
                            }
                            else if ((pos_value == "Низ") | (pos_value == "Н"))
                            {
                                offset_d = -(0.6 + (0.002 * scale_value));
                                offset_x = Math.Cos(1.5 * Math.PI + Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value;
                                offset_y = -(Math.Sin(1.5 * Math.PI + Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value);
                            }
                        }
                        // x- y+
                        else if ((dx < 0) & (dy > 0))
                        {
                            if ((pos_value == "Верх") | (pos_value == "В"))
                            {
                                offset_d = -(0.6 + (0.002 * scale_value));
                                offset_x = Math.Cos(Math.PI / 2 - Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value;
                                offset_y = -(Math.Sin(Math.PI / 2 - Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value);
                            }
                            else if ((pos_value == "Низ") | (pos_value == "Н"))
                            {
                                offset_d = -(0.6 + (0.002 * scale_value));
                                offset_x = -(Math.Cos(Math.PI / 2 - Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value);
                                offset_y = Math.Sin(Math.PI / 2 - Math.Acos(dx / d)) * 0.5 / 1000.0 * scale_value;
                            }
                        }
                        // x- y0
                        else if ((dx < 0) & (dy == 0))
                        {
                            if ((pos_value == "Верх") | (pos_value == "В"))
                            {
                                offset_d = -(0.6 + (0.002 * scale_value));
                                offset_y = 0.5 / 1000.0 * scale_value;
                            }
                            else if ((pos_value == "Низ") | (pos_value == "Н"))
                            {
                                offset_d = -(0.6 + (0.002 * scale_value));
                                offset_y = - 0.5 / 1000.0 * scale_value;
                            }

                        }
                        // x0 y-
                        else if ((dx == 0) & (dy < 0))
                        {
                            if ((pos_value == "Верх") | (pos_value == "В"))
                            {
                                offset_d = -(0.6 + (0.002 * scale_value));
                                offset_x = - 0.5 / 1000.0 * scale_value;
                            }
                            else if ((pos_value == "Низ") | (pos_value == "Н"))
                            {
                                offset_d = -(0.6 + (0.002 * scale_value));
                                offset_x = 0.5 / 1000.0 * scale_value;
                            }

                        }
                        // x+ y+
                        else
                        {
                            if ((pos_value == "Верх") | (pos_value == "В"))
                            {
                                offset_d = 0.6 + (0.002 * scale_value);
                                offset_x = Math.Cos(Math.Acos(dx / d) + Math.PI / 2) * 0.5 / 1000.0 * scale_value;
                                offset_y = Math.Sin(Math.Acos(dx / d) + Math.PI / 2) * 0.5 / 1000.0 * scale_value;
                            }
                            else if ((pos_value == "Низ") | (pos_value == "Н"))
                            {
                                offset_d = 0.6 + (0.002 * scale_value);
                                offset_x = Math.Cos(Math.Acos(dx / d) - Math.PI / 2) * 0.5 / 1000.0 * scale_value;
                                offset_y = Math.Sin(Math.Acos(dx / d) - Math.PI / 2) * 0.5 / 1000.0 * scale_value;
                            }
                        }

                        // compute text rotation based on line coordinate sector position
                        double rotation = 0.0;
                        if ((dx > 0) & (dy > 0))
                        {
                            rotation = Math.Acos(dx/d);
                        }
                        else if ((dx > 0) & (dy < 0))
                        {
                            rotation = 2*Math.PI - Math.Acos(dx/d);
                        }
                        else if ((dx < 0) & (dy > 0))
                        {
                            rotation = Math.Acos(dx/d) + Math.PI;
                        }
                        else if ((dx < 0) & (dy < 0))
                        {
                            rotation = Math.PI - Math.Acos(dx/d);
                        }
                        else if ((dx == 0) & (dy != 0))
                        {
                            rotation = Math.PI / 2;
                        }
                        else if ((dx < 0) & (dy == 0))
                        {
                            rotation = 0;
                        }
                        else
                        {
                            rotation = Math.Acos(dx/d);
                        }

                        if (d > scale_value / 100)
                        {
                            // st is a step on a line segment
                            // shift stores distance between previous segment last sign and endpoint
                            // also case when shift is greater than step
                            double st;
                            if (shift >= (scale_value / 14.2))
                            {
                                st = scale_value / 100;
                            }
                            else
                            {
                                st = scale_value / 14.2 - shift;
                            }
                            while (st < (d - scale_value / 100.0))
                            {
                                Point3d pt = new Point3d(
                                    pt1.X + (st + offset_d) * dx / d + offset_x,
                                    pt1.Y + (st + offset_d) * dy / d + offset_y,
                                    pt1.Z
                                    );
                                pts.Add(pt);
                                rot.Add(rotation);
                                shift = d - st;
                                st += scale_value / 14.2;
                            }
                        }
                    }

                    catch { }
                }
            }
        }
    }
}