using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TunnalCal.Helper
{
    class CADops
    {
        /// <summary>
        /// Get layer name by select a object
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static string getLayerName(string msg, Document doc)
        {
            //Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            string sLayer = null;

            PromptEntityResult per = ed.GetEntity(msg);
            if (per.Status == PromptStatus.OK)
            {
                ObjectId objectId = per.ObjectId;
                try
                {
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        //Get the entity
                        Autodesk.AutoCAD.DatabaseServices.Entity ent = trans.GetObject(objectId, OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        sLayer = ent.Layer;
                    }
                }
                catch { }
            }

            return sLayer;
        }


        /// <summary>
        /// project 3d polyline in XY plane, return a flattened polyline on XY plane and orignal 3d polyline
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="id">ObjectId of the 3d Polyline</param>
        /// <param name="originalPoly3d"></param>
        /// <returns></returns>
        public static Polyline3d CreatePolylineOnXYPlane(Document doc, ObjectId id, ref Polyline3d originalPoly3d)
        {
            Polyline3d pl = new Polyline3d();

            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);

                DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                Entity ent = obj as Entity;
                string layerName = ent.Layer.ToString();

                pl.Layer = layerName;
                btr.AppendEntity(pl);
                tr.AddNewlyCreatedDBObject(pl, true);

                Polyline3d p3d = obj as Polyline3d;
                if (p3d != null)
                {
                    originalPoly3d = p3d;
                    foreach (ObjectId vId in p3d)
                    {
                        PolylineVertex3d v3d = (PolylineVertex3d)tr.GetObject(vId, OpenMode.ForRead);
                        PolylineVertex3d v3d_new = new PolylineVertex3d(new Point3d(v3d.Position.X, v3d.Position.Y, 0));
                        pl.AppendVertex(v3d_new);//apdd point into 3d polyline
                        tr.AddNewlyCreatedDBObject(v3d_new, true);
                    }
                }

                tr.Commit();
            }

            return pl;
        }

        /// <summary>
        /// project 3d polyline in XY plane, return a flattened polyline on XY plane and orignal 3d polyline
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="id">ObjectId of the 3d Polyline</param>
        /// <param name="originalPoly3d"></param>
        /// <returns></returns>
        public static Polyline3d CreatePolylineFromPoint(Document doc, List<Point3d> pts)
        {
            Polyline3d pl = new Polyline3d();

            Database db = doc.Database;
            #region create polyline with points
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);

                    pl.Layer = CADops.getLayerName("Select Any Point in Layer.", doc);

                    btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);

                    foreach (Point3d pt in pts)
                    {       
                        PolylineVertex3d vex3d = new PolylineVertex3d(pt);
                        pl.AppendVertex(vex3d);//apdd point into 3d polyline
                        tr.AddNewlyCreatedDBObject(vex3d, true);
                    }
                }
                catch { }
                tr.Commit();
            }
            #endregion
            return pl;
        }

        /// <summary>
        /// create the 3D polyline in specifed layer
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pts"></param>
        /// <param name="layerName"></param>
        /// <returns></returns>
        public static Polyline3d CreatePolylineFromPoint(Document doc, List<Point3d> pts, string layerName)
        {
            Polyline3d pl = new Polyline3d();

            Database db = doc.Database;
            #region create polyline with points
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);

                    pl.Layer = layerName;

                    btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);


                    foreach (Point3d pt in pts)
                    {
                        PolylineVertex3d vex3d = new PolylineVertex3d(pt);
                        pl.AppendVertex(vex3d);//apdd point into 3d polyline
                        tr.AddNewlyCreatedDBObject(vex3d, true);
                    }
                }
                catch { }
                tr.Commit();
            }

            #endregion
            return pl;
        }

        /// <summary>
        /// Create 3D polyline use 3D points and also ref return a 3D polyline that is in XY Plane, in the sepcified layerName
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pts"></param>
        /// <param name="polyInXYplane"></param>
        /// <param name="layerName"></param>
        /// <returns></returns>
        public static Polyline3d CreatePolylineFromPoint(Document doc, List<Point3d> pts, ref Polyline3d polyInXYplane, string layerName)
        {
            Polyline3d pl = new Polyline3d();

            Database db = doc.Database;
            #region create polyline with points
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);

                    pl.Layer = layerName;
                    polyInXYplane.Layer = layerName;

                    btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);

                    btr.AppendEntity(polyInXYplane);
                    tr.AddNewlyCreatedDBObject(polyInXYplane, true);

                    foreach (Point3d pt in pts)
                    {
                        //Point3d pt = new Point3d(Convert.ToDouble(myPointsXYZ[n, 0]), Convert.ToDouble(myPointsXYZ[n, 1]), Convert.ToDouble(myPointsXYZ[n, 2]));
                        PolylineVertex3d vex3d = new PolylineVertex3d(pt);
                        PolylineVertex3d vex3d0 = new PolylineVertex3d(new Point3d(pt.X, pt.Y, 0));
                        pl.AppendVertex(vex3d);//apdd point into 3d polyline
                        polyInXYplane.AppendVertex(vex3d0);
                        tr.AddNewlyCreatedDBObject(vex3d, true);
                        tr.AddNewlyCreatedDBObject(vex3d0, true);
                    }
                }
                catch { }
                tr.Commit();
            }
            #endregion
            return pl;
        }

        /// <summary>
        /// Get a list of Vectors that is vertical to path of all points
        /// </summary>
        /// <param name="pts"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static List<Vector3d> getVectors(List<Point3d> pts, Document doc, ref List<Vector3d> vectorsAlongPath)
        {
            Editor ed = doc.Editor;

            //===========  rotate, move the W beam profile to the starting point of the line
            Matrix3d curUCSMatrix = doc.Editor.CurrentUserCoordinateSystem;
            CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;

            List<Vector3d> vectors = new List<Vector3d>();
            int noOfpts = pts.Count();
            #region get the first points vector
            Point3d pt1 = new Point3d(pts[0].X, pts[0].Y, 0);
            Point3d pt2 = new Point3d(pts[1].X, pts[1].Y, 0);
            Vector3d vect1 = pt1.GetVectorTo(pt2);
            
            ed.WriteMessage($"vect1 => {vect1.X}, {vect1.Y}, {vect1.Z}\n");

            Vector3d vect2 = vect1.TransformBy(Matrix3d.Rotation(Helper.Angle.angToRad(90), curUCS.Zaxis, new Point3d(0, 0, 0)));//rotation 90 degree along Z axis
            ed.WriteMessage($"vect2 => {vect2.X}, {vect2.Y}, {vect2.Z}\n");

            vectorsAlongPath.Add(vect1);
            vectors.Add(vect2);
            #endregion

            #region get middle points vectors
            for (int i = 1; i <= noOfpts - 2; i++)
            {
                pt1 = new Point3d(pts[i - 1].X, pts[i - 1].Y, 0);
                pt2 = new Point3d(pts[i + 1].X, pts[i + 1].Y, 0);
                vect1 = pt1.GetVectorTo(pt2);
                ed.WriteMessage($"vect1 => {vect1.X}, {vect1.Y}, {vect1.Z}\n");

                vect2 = vect1.TransformBy(Matrix3d.Rotation(Helper.Angle.angToRad(90), curUCS.Zaxis, new Point3d(0, 0, 0)));//rotation 90 degree along Z axis
                ed.WriteMessage($"vect2 => {vect2.X}, {vect2.Y}, {vect2.Z}\n");

                vectorsAlongPath.Add(vect1);
                vectors.Add(vect2);
            }
            #endregion

            #region get last points vector
            pt1 = new Point3d(pts[noOfpts - 2].X, pts[noOfpts - 2].Y, 0);
            pt2 = new Point3d(pts[noOfpts - 1].X, pts[noOfpts - 1].Y, 0);

            vect1 = pt1.GetVectorTo(pt2);
            ed.WriteMessage($"vect1 => {vect1.X}, {vect1.Y}, {vect1.Z}\n");

            vect2 = vect1.TransformBy(Matrix3d.Rotation(Helper.Angle.angToRad(90), curUCS.Zaxis, new Point3d(0, 0, 0)));//rotation 90 degree along Z axis
            ed.WriteMessage($"vect2 => {vect2.X}, {vect2.Y}, {vect2.Z}\n");

            vectorsAlongPath.Add(vect1);
            vectors.Add(vect2);
            #endregion

            return vectors;
        }


        public static List<Vector3d> GetVectors(List<Point3d>pts, Document doc)
        {
            List<Vector3d> vectors = new List<Vector3d>();
            
            //create polyline

            return vectors;
        }
        /// <summary>
        /// Get all polylines in 'sLayer' layer
        /// </summary>
        /// <param name="sLayer"></param>
        /// <returns></returns>
        public static ObjectIdCollection SelectAllPolyline(string sLayer)
        {
            ObjectIdCollection retVal = null;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Get a selection set of all possible polyline entities on the requested layer
                PromptSelectionResult oPSR = null;

                TypedValue[] tvs = new TypedValue[] {
            new TypedValue(Convert.ToInt32(DxfCode.Operator), "<and"),
            new TypedValue(Convert.ToInt32(DxfCode.LayerName), sLayer),//this is not required
            new TypedValue(Convert.ToInt32(DxfCode.Operator), "<or"),
            new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE"),
            new TypedValue(Convert.ToInt32(DxfCode.Start), "LWPOLYLINE"),
            new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE2D"),
            new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE3d"),
            new TypedValue(Convert.ToInt32(DxfCode.Operator), "or>"),
            new TypedValue(Convert.ToInt32(DxfCode.Operator), "and>")
                };

                SelectionFilter oSf = new SelectionFilter(tvs);

                oPSR = ed.SelectAll(oSf);

                if (oPSR.Status == PromptStatus.OK)
                {
                    retVal = new ObjectIdCollection(oPSR.Value.GetObjectIds());
                }
                else
                {
                    retVal = new ObjectIdCollection();
                }
            }
            catch { }

            return retVal;
        }

        public static List<Point3d> GetPointsFrom3dPolyline(Polyline3d poly3d, Document doc)
        {
            Database db = doc.Database;
            Editor ed = doc.Editor;

            List<Point3d> pts = new List<Point3d>();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                if (poly3d != null)
                {
                    // Use foreach to get each contained vertex
                    foreach (ObjectId vId in poly3d)
                    {
                        PolylineVertex3d v3d = (PolylineVertex3d)tr.GetObject(vId, OpenMode.ForRead);
                        Point3d pt = new Point3d(v3d.Position.X, v3d.Position.Y, v3d.Position.Z);

                        pts.Add(pt);

                        ed.WriteMessage($"GetPointsFrom3dPolyline => {pt.X}, {pt.Y}, {pt.Z}\n");
                    }
                }

                tr.Commit();
            }
            return pts;
        }
        
        public static double scaleNmove(double input, long adjust, long scaler)
        {
            double output = input * scaler - adjust;
            return output;
        }

        public static string CreateAndAssignALayer(string sLayerName)
        {
            string name = string.Empty;
            // Get the current document and database
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = acDoc.Database;

            // Start a transaction
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable; // Open the Layer table for read

                if (acLyrTbl.Has(sLayerName) == false)
                {
                    using (LayerTableRecord acLyrTblRec = new LayerTableRecord())
                    {
                        // Assign the layer the ACI color 3 and a name
                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                        acLyrTblRec.Name = sLayerName;

                        acLyrTbl.UpgradeOpen();// Upgrade the Layer table for write

                        acLyrTbl.Add(acLyrTblRec);// Append the new layer to the Layer table and the transaction
                        tr.AddNewlyCreatedDBObject(acLyrTblRec, true);
                        //db.Clayer = acLyrTbl[sLayerName];//set current layer
                        name = sLayerName;
                    }
                }
                tr.Commit();
            }
            return name;
        }

        private static ObjectIdCollection GetEntitiesOnLayer(string layerName)
        {
            Document doc =Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
           
            // Build a filter list so that only entities
            // on the specified layer are selected
            TypedValue[] tvs = new TypedValue[1] {new TypedValue((int)DxfCode.LayerName,layerName)};

            SelectionFilter sf =new SelectionFilter(tvs);

            PromptSelectionResult psr =ed.SelectAll(sf);

            if (psr.Status == PromptStatus.OK)
                return new ObjectIdCollection(psr.Value.GetObjectIds());
            else
                return new ObjectIdCollection();
        }

        /// <summary>
        /// Insert block into layout
        /// </summary>
        /// <param name="db"></param>
        /// <param name="layoutName"></param>
        /// <param name="blockName"></param>
        /// <param name="insPnt"></param>
        /// <returns></returns>
        public static bool InsertBlock(Database db, string layoutName, string blockName, Point3d insPnt)
        {
            LayoutManager lm = LayoutManager.Current;
            using(Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (!bt.Has(blockName))
                    return false; ////// need to load the block
                
                //get the block id
                ObjectId blkId = bt[blockName];
                BlockTableRecord btr = null;
                if (string.IsNullOrEmpty(layoutName))
                    btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                else
                {
                    ObjectId layoutId = lm.GetLayoutId(layoutName);
                    Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                    btr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                }

                //insert the block
                BlockReference blkRef = new BlockReference(insPnt, blkId);
                btr.UpgradeOpen();
                btr.AppendEntity(blkRef);
                tr.AddNewlyCreatedDBObject(blkRef, true);
                tr.Commit();
            }
            return true;
        }
   
        public static List<string> ImportBlocks(Database db, Database sourceDb, List<string> blockNames)
        {
            List<string> output = new List<string>();
            try
            {
                //create a variable to store the list of block identifiers
                ObjectIdCollection blockIds = new ObjectIdCollection();
                Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = sourceDb.TransactionManager;
                using(Transaction tr = tm.StartTransaction())
                {
                    //open the block table
                    BlockTable bt = tm.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false) as BlockTable;
                    //check each block in the block table
                    foreach(ObjectId btrId in bt)
                    {
                        BlockTableRecord btr = tm.GetObject(btrId, OpenMode.ForRead, false) as BlockTableRecord;
                        //only add named & non-layout blocks to the copy list
                        if (blockNames.Contains(btr.Name))
                        {
                            blockIds.Add(btrId);
                            output.Add(btr.Name);
                        }

                        btr.Dispose();
                    }
                }

                //copy blocks from source to destination database
                IdMapping mapping = new IdMapping();
                sourceDb.WblockCloneObjects(blockIds, db.BlockTableId, mapping, DuplicateRecordCloning.Ignore, false);
            }
            catch { }
            return output;
        }
    }
}
