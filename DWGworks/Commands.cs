using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;

using TunnalCal.Helper;
using TunnalCal.C3D;
using System.IO;
using Xbim.Common.Exceptions;
using Autodesk.AutoCAD.GraphicsInterface;

namespace TunnalCal
{
    public class Commands
    {
        [CommandMethod("IFC_CREATE_ALIGNMENT")]
        static public void CreateAlignmentFromIFC()
        {
            string ifcpath = @"D:\_DE_Tech_Projects\ObjectCreator\Alignment.ifc";
            Alignments.CreateAlignmentFromIFC(ifcpath);
        }
        
        
        /// <summary>
        /// List out Vertices in a selected polyline
        /// </summary>
        [CommandMethod("TUN_ListVertices")]
        static public void ListVertices()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            HostApplicationServices hs =HostApplicationServices.Current;
            string path =hs.FindFile(doc.Name, doc.Database,FindFileHint.Default);

            //ed.WriteMessage("\nTo List all the points in the 3D string, type command LISTVERTICES.");

            long adjustX = 0;
            long adjustY = 0;
            long scaler = 1;

            PromptEntityResult per = ed.GetEntity("Select polylines");
            ObjectId oid = per.ObjectId;

            if (per.Status == PromptStatus.OK)
            {
                Transaction tr = db.TransactionManager.StartTransaction();
                DBObject objPick = tr.GetObject(oid, OpenMode.ForRead);
                Entity objEnt = objPick as Entity;
                string sLayer = objEnt.Layer.ToString();

                path += "_" + sLayer;

                ObjectIdCollection oDBO = CADops.SelectAllPolyline(sLayer);
                //Handle handseed = db.Handseed;

                try
                {
                    using (tr)
                    {
                        long shtNo = 1;
                        List<string> data = new List<string>();
                        foreach (ObjectId id in oDBO)
                        {
                            //DBObject obj = tr.GetObject(per.ObjectId, OpenMode.ForRead);
                            DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                            Entity ent = obj as Entity;
                            string layerName = ent.Layer.ToString();

                            //create list for storing x,y,z value of point
                            string StrId = "";


                            // If a "lightweight" (or optimized) polyline
                            Autodesk.AutoCAD.DatabaseServices.Polyline lwp = obj as Autodesk.AutoCAD.DatabaseServices.Polyline;

                            if (lwp != null)
                            {
                                //StrId = lwp.ObjectId.ToString();
                                StrId = lwp.Handle.ToString();
                                ed.WriteMessage("\n" + StrId);
                                // Use a for loop to get each vertex, one by one
                                int vn = lwp.NumberOfVertices;
                                for (int i = 0; i < vn; i++)
                                {
                                    // Could also get the 3D point here
                                    Point2d pt = lwp.GetPoint2dAt(i);
                                    string temp = CADops.scaleNmove(pt.X, adjustX, scaler) + "," + CADops.scaleNmove(pt.Y, adjustY, scaler) + ", ," + StrId + "," + layerName;
                                    data.Add(temp);
                                    ed.WriteMessage("\n" + pt.ToString());
                                }
                            }
                            else
                            {
                                // If an old-style, 2D polyline
                                Polyline2d p2d = obj as Polyline2d;
                                if (p2d != null)
                                {
                                    StrId = p2d.Handle.ToString();
                                    ed.WriteMessage("\n" + StrId);

                                    // Use foreach to get each contained vertex
                                    foreach (ObjectId vId in p2d)
                                    {
                                        Vertex2d v2d = (Vertex2d)tr.GetObject(vId, OpenMode.ForRead);
                                        string temp = CADops.scaleNmove(v2d.Position.X, adjustX, scaler) + "," + CADops.scaleNmove(v2d.Position.Y, adjustY, scaler) + "," + v2d.Position.Z * scaler + "," + StrId + "," + layerName;
                                        data.Add(temp);
                                        ed.WriteMessage("\n" + v2d.Position.ToString());
                                    }
                                }
                                else
                                {
                                    // If an old-style, 3D polyline
                                    Polyline3d p3d = obj as Polyline3d;
                                    if (p3d != null)
                                    {
                                        StrId = p3d.Handle.ToString();
                                        ed.WriteMessage("\n" + StrId);

                                        // Use foreach to get each contained vertex
                                        foreach (ObjectId vId in p3d)
                                        {
                                            PolylineVertex3d v3d = (PolylineVertex3d)tr.GetObject(vId, OpenMode.ForRead);
                                            string temp = CADops.scaleNmove(v3d.Position.X, adjustX, scaler) + "," + CADops.scaleNmove(v3d.Position.Y, adjustY, scaler) + "," + v3d.Position.Z * scaler + "," + StrId + "," + layerName;
                                            data.Add(temp);
                                            ed.WriteMessage("\n" + v3d.Position.ToString());
                                        }
                                    }
                                }
                            }
                            //create dataarray to populate in excel
                            int no = data.Count;
                            ed.WriteMessage("\n Number of Point:" + no);

                            //createCSV(data, shtNo);
                            shtNo += 1;

                        }
                        Excel.createCSV(data, path);
                        // Committing is cheaper than aborting
                        ed.WriteMessage("\ncsv file has been created under path " + path);
                        ed.WriteMessage("\nData format: Easting, Northing, Elevation, String ID, Layer Name.");
                        tr.Commit();
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// Create Polyline from a excel datasheet, require the data sheet path by selection and the name of sheet
        /// </summary>
        [CommandMethod("TUN_CreatePolyLine")]
        static public void CreatePolyLine()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            HostApplicationServices hs = HostApplicationServices.Current;
            string outputPath = hs.FindFile(doc.Name, doc.Database, FindFileHint.Default);

            //get user input
            //select point input file
            ed.WriteMessage("Select the Excel file contains the door step points");
            string excelPath = FileOps.SelectFile();

            #region get Excel Sheet Name
            PromptStringOptions pStrOpts = new PromptStringOptions("\nEnter Sheet Name: ");
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = doc.Editor.GetString(pStrOpts);
            string shtName = pStrRes.StringResult;
            #endregion

            List<Point3d> doorStepPts = Excel.getAllpoint(doc, excelPath, shtName);

            CADops.CreatePolylineFromPoint(doc, doorStepPts);

        }

        /// <summary>
        /// Select a list of point from excel as primary input, select excel file, and type in Sheet Name;
        /// OUTPUT is also a list of point as excel;
        /// This Function check this primary point list against two polyline (rail track) to:
        /// offset a tunnel radius from list of point, create a line between points and offseted point;
        /// find the intersect point between line and both rails, find the low rail point;
        /// compare the point and low rail point to check if there are at least 550mm clearance;
        /// if it is more than 550mm in clearance, output point is the same as input points;
        /// if it is less than 550mm in clearance, output point is use input point X,Y and add 550mm over low rail point's Z
        /// ------------- future works ------------------
        /// allow user to input clearance number for checking
        /// allow user to input tunnel diameter
        /// </summary>
        [CommandMethod("TUN_CheckElevationDifference")]
        //project to level 0 to find the intersect point
        static public void CheckElevationDifference()
        {
            string msg = string.Empty;

            double checkDistance = 0.55;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            HostApplicationServices hs = HostApplicationServices.Current;
            string outputPath = hs.FindFile(doc.Name, doc.Database, FindFileHint.Default);
            string logPath = outputPath + "_log.log";
            System.IO.File.WriteAllText(logPath, msg);

            //get user input
            //select point input file
            ed.WriteMessage("Select the Excel file contains the door step points");
            string excelPath = FileOps.SelectFile();

            #region get Excel Sheet Name
            PromptStringOptions pStrOpts = new PromptStringOptions("\nEnter Sheet Name: ");
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = doc.Editor.GetString(pStrOpts);
            string shtName = pStrRes.StringResult;
            #endregion

            List<Point3d> doorStepPts = Excel.getAllpoint(doc, excelPath, shtName);

            PromptEntityResult per = ed.GetEntity("Select polylines");
            ObjectId oid = per.ObjectId;

            ObjectIdCollection oDBO = new ObjectIdCollection();
            List<Polyline3d> poly3ds = new List<Polyline3d>();
            List<Polyline3d> poly3ds0 = new List<Polyline3d>();

            if (per.Status == PromptStatus.OK)
            {
                Transaction tr = db.TransactionManager.StartTransaction();

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);

                DBObject objPick = tr.GetObject(oid, OpenMode.ForRead);
                Entity objEnt = objPick as Entity;
                string sLayer = objEnt.Layer.ToString();
                
                oDBO = CADops.SelectAllPolyline(sLayer);
                foreach (ObjectId id in oDBO)
                {
                    Polyline3d pl = new Polyline3d();
                    System.IO.File.AppendAllText(logPath, $"{id}\n");
                    poly3ds0.Add(CADops.CreatePolylineOnXYPlane(doc, id, ref pl));
                    poly3ds.Add(pl);
                }
            }

            List<Point3d> output = new List<Point3d>();
            List<string> data = new List<string>();
            if (poly3ds.Count() > 0)
            {
                Transaction tr = db.TransactionManager.StartTransaction();

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);

                using (tr)
                {
                    List<Vector3d> vectorsAlongPath = new List<Vector3d>();
                    List<Vector3d> vectors = CADops.getVectors(doorStepPts, doc, ref vectorsAlongPath);

                    //foreach (Point3d pt in doorStepPts)
                    for (int i = 0; i < doorStepPts.Count(); i++)
                    {
                        Point3d pt = doorStepPts[i];

                        Point3d pt0 = new Point3d(pt.X, pt.Y, 0);

                        Vector3d v = vectors[i].GetNormal() * 5;
                        Matrix3d mat = Matrix3d.Displacement(v);
                        Point3d npt = pt0.TransformBy(mat);

                        v = vectors[i].GetNormal() * -5;
                        mat = Matrix3d.Displacement(v);
                        Point3d npt2 = pt0.TransformBy(mat);

                        //create a 2d line in XY plane
                        Line ln = new Line(npt, npt2);

                        btr.AppendEntity(ln);
                        tr.AddNewlyCreatedDBObject(ln, true);

                        msg = $"pt => {pt.X}, {pt.Y}, {pt.Z}\n";

                        #region get intersect point from point to polyline
                        Point3d ptNearest = Point3d.Origin;
                        for(int j = 0; j < poly3ds0.Count(); j ++)
                        {
                            Polyline3d p3d0 = poly3ds0[j];
                            Polyline3d p3d = poly3ds[j];

                            Point3d ptTemp = new Point3d();
                            #region get the alignment object and find the nearest point to the nominated point

                            Point3dCollection pts3D = new Point3dCollection();
                            p3d0.IntersectWith(ln, Intersect.OnBothOperands, pts3D, IntPtr.Zero, IntPtr.Zero);

                            try
                            {
                                if (pts3D.Count > 0)
                                {
                                    double para = p3d0.GetParameterAtPoint(pts3D[0]);
                                    //ed.WriteMessage($"{pts3D[0]}, {para}\n");
                                    ptTemp = p3d.GetPointAtParameter(para);
                                }
                            }
                            catch { }

                            #region get the point with lower Z
                            if (ptNearest == Point3d.Origin)
                                ptNearest = ptTemp;
                            else
                            {
                                if (ptNearest.Z > ptTemp.Z)
                                    ptNearest = ptTemp;
                            }
                            #endregion
                        }
                        #endregion
                        #endregion

                        msg += $"ptNearest: {ptNearest.X}, {ptNearest.Y}, {ptNearest.Z}\n";

                        try
                        {
                            double diff = ptNearest.Z - pt.Z;
                            if (Math.Abs(diff) <= checkDistance)
                            {
                                Point3d newPt = new Point3d(pt.X, pt.Y, ptNearest.Z + checkDistance);
                                output.Add(newPt);
                                data.Add($"{newPt.X},{newPt.Y},{newPt.Z}, less, {diff}, {ptNearest.Z}");
                                msg += $", Z diff: {diff} => less than {checkDistance} => {newPt.X}, {newPt.Y}, {newPt.Z}\n\n";
                            }
                            else
                            {
                                Point3d newPt = pt;
                                output.Add(newPt);
                                data.Add($"{newPt.X},{newPt.Y},{newPt.Z}, more, {diff}, {ptNearest.Z}");
                                msg += $", Z diff: {diff} => more than {checkDistance} => {newPt.X}, {newPt.Y}, {newPt.Z}\n\n";
                            }
                        }
                        catch { }

                        System.IO.File.AppendAllText(logPath, msg);
                    }
                    tr.Commit();
                }
            }
            Excel.createCSV(data, outputPath);
            ed.WriteMessage("\ncsv file has been created under path " + outputPath);

        }

        /// <summary>
        /// Create a Tunnel TBM solid
        /// </summary>
        [CommandMethod("TUN_CreateTBM")]
        static public void CreateTunnel()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            //===========  
            Matrix3d curUCSMatrix = doc.Editor.CurrentUserCoordinateSystem;
            CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;

            double TunnelDia = 6.04;
            PromptEntityResult per = ed.GetEntity("Select polylines");
            ObjectId oid = per.ObjectId;

            if (per.Status == PromptStatus.OK)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);

                    DBObject objPick = tr.GetObject(oid, OpenMode.ForRead);
                    Entity objEnt = objPick as Entity;
                    string sLayer = objEnt.Layer.ToString();

                    ObjectIdCollection oDBO = CADops.SelectAllPolyline(sLayer);

                    List<string> data = new List<string>();
                    foreach (ObjectId id in oDBO)
                    {
                        //get 3d polyline
                        Polyline3d poly3d = tr.GetObject(id, OpenMode.ForRead) as Polyline3d;
                    
                        //get points on 3d polyline
                        List<Point3d> pts0 = CADops.GetPointsFrom3dPolyline(poly3d, doc);
                        List<Point3d> pts = pts0.Distinct(new PointComparer(3)).ToList();

                        //get vectors on points
                        List<Vector3d> vectorAlongPath = new List<Vector3d>();
                        List<Vector3d> vectors = CADops.getVectors(pts, doc, ref vectorAlongPath);

                        List<LoftProfile> loftProfiles = new List<LoftProfile>();
                        for (int i = 0; i < pts.Count(); i=i+4)
                        {
                            Point3d pt = pts[i];
                            //ed.WriteMessage($"TUN_CreateTBM => {pt.X}, {pt.Y}, {pt.Z}; ");

                            Vector3d v = vectors[i].GetNormal() * TunnelDia/2;
                            Matrix3d mat = Matrix3d.Displacement(v);
                            Point3d npt = pt.TransformBy(mat);

                            //create a 2d line in XY plane
                            Line ln = new Line(npt, pt);

                            btr.AppendEntity(ln);
                            tr.AddNewlyCreatedDBObject(ln, true);

                            double ang = ln.Angle;//vectors[i].GetAngleTo(curUCS.Xaxis);
                                                  //ed.WriteMessage($"angle {ang}\n");

                            Region acRegion = new Region();
                            try
                            {
                                using (Circle acCirc = new Circle())
                                {
                                    acCirc.Center = new Point3d(pt.X, pt.Y, pt.Z);
                                    acCirc.Radius = TunnelDia /2;

                                    acCirc.TransformBy(Matrix3d.Rotation(Angle.angToRad(90), curUCS.Xaxis, acCirc.Center));
                                    acCirc.TransformBy(Matrix3d.Rotation(ang, curUCS.Zaxis, acCirc.Center));

                                    // Add the new object to the block table record and the transaction
                                    btr.AppendEntity(acCirc);
                                    tr.AddNewlyCreatedDBObject(acCirc, true);

                                    DBObjectCollection acDBObjColl = new DBObjectCollection();
                                    acDBObjColl.Add(acCirc);
                                    DBObjectCollection myRegionColl = new DBObjectCollection();
                                    myRegionColl = Region.CreateFromCurves(acDBObjColl);
                                    Region acRegion1 = myRegionColl[0] as Region;
                                    // Add the new object to the block table record and the transaction
                                    btr.AppendEntity(acRegion1);
                                    tr.AddNewlyCreatedDBObject(acRegion1, true);
                                    LoftProfile lp1 = new LoftProfile(acRegion1);
                                    loftProfiles.Add(lp1);
                                }
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage(ex.ToString()); }
                        }

                        LoftProfile[] lps = new LoftProfile[loftProfiles.Count()];
                        for (int i = 0; i < loftProfiles.Count(); i ++)
                        {
                            lps[i] = loftProfiles[i];
                        }

                        try
                        {
                            // =========== create loft solid
                            Solid3d sol = new Solid3d();
                            sol.SetDatabaseDefaults();
                            LoftOptions lpOptions = new LoftOptions();
                            //LoftProfile lpGuide = new LoftProfile(acLine);//create loft profile
                            //LoftProfile[] guideToLoft = new LoftProfile[1] { lpGuide };
                            //sol.CreateLoftedSolid(lps, guideToLoft, null, lpOptions);//guide to loft can not be null, this parameter is not optional
                            sol.CreateLoftedSolid(lps, null, null, lpOptions);
                            sol.Layer = sLayer;//assign layer

                            //=============save it into database
                            Autodesk.AutoCAD.DatabaseServices.Entity ent = sol;//this is created for using hyperlink
                            btr.AppendEntity(ent);
                            tr.AddNewlyCreatedDBObject(ent, true);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage(ex.ToString()); }
                    }

                    tr.Commit();
                }
            }

        }


        [CommandMethod("ACM_LAYERS_N_TITLEBLOCK_EXPORT")]
        static public void ExportLayersFromDrawings()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database docDB = doc.Database;
            Editor ed = doc.Editor;

            string[] fileList = FileOps.selectFiles("Select all the files to export layers.", "dwg", "Select all the files to export layers.");
            if (fileList == null)
            {
                ed.WriteMessage("\nNo file selected, operation Termiated.\n");
                return;
            }

            Dictionary<string, List<string>> dwgLayers = new Dictionary<string, List<string>>();//to store the layers in each drawings
            Dictionary<string, List<string>> titleBlocks = new Dictionary<string, List<string>>();//to store the number of titleblock in a drawing

            foreach (string file in fileList)
            {  
                if (File.Exists(file))
                {
                    ed.WriteMessage($"Read layer from {file}\n");
                    try
                    {
                        Database db = new Database(false, true);
                        using (db)
                        {
                            db.ReadDwgFile(file, FileOpenMode.OpenForReadAndAllShare, false, null);
                            db.CloseInput(true);

                            Transaction tr = db.TransactionManager.StartTransaction();

                            using (tr)
                            {
                                ed.WriteMessage($"Start transition....\n");
                                LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                                List<string> layers = new List<string>();
                                foreach (ObjectId id in lt)
                                {
                                    LayerTableRecord l = tr.GetObject(id, OpenMode.ForWrite) as LayerTableRecord;
                                    string name = l.Name;
                                    string revised = name.Trim().ToUpper(); //clean up the layer name

                                    if (!l.IsDependent)
                                    {
                                        if (!layers.Contains(revised))
                                            layers.Add(revised);
                                    }
                                }

                                List<string> blocks = new List<string>();
                                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                //get the record from paper space
                                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;
                                foreach(ObjectId oId in btr)
                                {
                                    BlockReference br = tr.GetObject(oId, OpenMode.ForRead) as BlockReference;//get block reference
                                    if (oId.ObjectClass.DxfName == "INSERT")//if this block is inserted
                                        blocks.Add(br.Name);
                                }
                                tr.Commit();

                                ed.WriteMessage($"End transition....\n");
                                dwgLayers.Add(file, layers);
                                titleBlocks.Add(file, blocks);
                            }
                            ed.WriteMessage("\n Finish reading " + System.IO.Path.GetFileName(file) + ".\n");
                        }
                    }
                    catch(Autodesk.AutoCAD.Runtime.Exception ex) 
                    { ed.WriteMessage(ex.ToString() + Environment.NewLine); }
                }
                else
                {
                    ed.WriteMessage("File " + file + " does not exist.\n");
                }
            }

            Excel.createCSV(dwgLayers, "D:\\dwgLayers.csv");
            Excel.createCSV(titleBlocks, "D:\\dwgBlocks.csv");
            ed.WriteMessage("Done.");
        }

        [CommandMethod("ACM_RENAME_LAYERS_BY_EXCEL")]
        static public void RenameDrawingsLayersByExcel()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            //string[] filePath = FileOps.getFilePaths("Select layer Excel file", "Excel files | *.xlsx", false);

            //if(filePath != null)
            //{
                string excelFile = "D:\\cad_layer_update_test.xlsx";

                Dictionary<string, Dictionary<string, string>> data = Excel.getDwgLayerSetup(excelFile);

                foreach(KeyValuePair<string, Dictionary<string, string>> d in data)
                {
                    string file = d.Key;
                    Dictionary<string, string> layers = d.Value;

                    if (File.Exists(file))
                    {
                        ed.WriteMessage($"Rename layers for {file}\n");
                        try
                        {
                        Database db = new Database(false, true);
                        ObjectIdCollection layIdsToBeRemoved = new ObjectIdCollection();
                        using (db)
                        {
                            db.ReadDwgFile(file, System.IO.FileShare.ReadWrite, true, null);

                            Transaction tr = db.TransactionManager.StartTransaction();
                            using (tr)
                            {
                                ed.WriteMessage($"Start transition....\n");
                                #region ///look into BlockTable and saved object id
                                Dictionary<string, ObjectIdCollection> objLayers = new Dictionary<string, ObjectIdCollection>();//key = layer name; value = ObjectIdCollection
                                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);//get all ObjectIdCollection in database
                                foreach (ObjectId btrId in blockTable)
                                {
                                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                    foreach(ObjectId id in btr)
                                    {
                                        Entity obj = tr.GetObject(id, OpenMode.ForRead) as Entity;
                                        string layerName = obj.Layer.Trim().ToUpper();
                                        if (objLayers.ContainsKey(layerName))
                                            objLayers[layerName].Add(id);//add the objectId to ObjectIdCollection if it is already exist
                                        else
                                        {
                                            ObjectIdCollection col = new ObjectIdCollection();
                                            col.Add(id);
                                            objLayers.Add(layerName, col);//add first item
                                        }
                                    }
                                }
                                #endregion

                                #region get layer table
                                LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                                
                                //set layer 0 as current layer
                                db.Clayer = lt["0"];

                                foreach (ObjectId lyId in lt)
                                {
                                    LayerTableRecord l = tr.GetObject(lyId, OpenMode.ForWrite) as LayerTableRecord;//get LayerTableRecord

                                    //only do something when the layer is not in xref
                                    if (!l.IsDependent)
                                    {
                                        string name = l.Name;
                                        string oldLayerName = name.Trim().ToUpper(); //clean up the layer name

                                        if (layers.ContainsKey(oldLayerName))
                                        {
                                            string newLayerName = layers[oldLayerName];
                                            if (string.IsNullOrEmpty(newLayerName))
                                                continue;

                                            if (lt.Has(newLayerName))
                                            {
                                                #region move the entities that is in the oldLayerName to the newLayerName if this new Layer already exist 
                                                //if this new layer name is already exists
                                                //move geometries in this old layer into that right layer
                                                if (objLayers.ContainsKey(oldLayerName))
                                                {
                                                    ObjectIdCollection col = objLayers[oldLayerName];
                                                    foreach (ObjectId objId in col)
                                                    {
                                                        Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                                                        ent.Layer = newLayerName;
                                                    }
                                                    #endregion
                                                    //add the old layer to the layIdsToBeRemoved
                                                    layIdsToBeRemoved.Add(lyId);
                                                }
                                            }
                                            else
                                            {
                                                //if this new layer name is not already exists
                                                //change current layer name to new name
                                                l.Name = newLayerName;
                                            }
                                        }
                                    }
                                }
                                #endregion


                                #region remove layers
                                foreach (ObjectId id in layIdsToBeRemoved)
                                {
                                    DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
                                    obj.Erase();
                                }
                                #endregion
                                tr.Commit();

                                ed.WriteMessage($"End transition....\n");
                            }
                            ed.WriteMessage("\n Finish update " + System.IO.Path.GetFileName(file) + ".\n");

                            db.SaveAs(file, DwgVersion.Current);//save changes
                        }
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    { ed.WriteMessage(ex.ToString() + Environment.NewLine); }
                }
                else
                {
                    ed.WriteMessage("File " + file + " does not exist.\n");
                }
            }
            //}
        }

        /// <summary>
        /// not done anything yet
        /// </summary>
        [CommandMethod("ACM_INSERT_BLOCKS_BYEXCEL")]
        static public void InsertBlocksByExcel()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            string excelFile = "D:\\dwgBlock_list.xlsx";

            Dictionary<string, List<string>> data = Excel.getDwgBlockList(excelFile);

            string blockTemplate = @"D:\_PROJECTs\60614740-MascotTunnel\_blocks\CADtemplate.dwg";
            List<string> blockRequired = new List<string>()
            {
                "RMS_TfNSW_DESIGN-BLK-00-0000-GE-ADM_A1",
                "RMS_TfNSW_DESIGN-BLK-00-0000-GE-ATT_A1"
            };
            //load block template file database
            Database sourceDb = new Database(false, true);
            if (File.Exists(blockTemplate))
            {
                //open block template file to read
                sourceDb.ReadDwgFile(blockTemplate, System.IO.FileShare.Read, true, null);

                foreach (KeyValuePair<string, List<string>> d in data)
                {
                    string file = d.Key;
                    List<string> blocks = d.Value;

                    if (File.Exists(file))
                    {
                        ed.WriteMessage($"Open {file} to insert blocks.....\n");
                        try
                        {
                            Database db = new Database(false, true);
                            using (db)
                            {
                                //open drawing database
                                db.ReadDwgFile(file, System.IO.FileShare.ReadWrite, true, null);

                                #region load blocks into drawings, make sure the block is available to be inserted
                                List<string> loadedBlock = new List<string>();
                                try
                                {
                                    //create a variable to store the list of block identifiers
                                    ObjectIdCollection blockIdsToCopied = new ObjectIdCollection();
                                    Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = sourceDb.TransactionManager;
                                    using (Transaction tr = tm.StartTransaction())
                                    {
                                        ed.WriteMessage($"start transaction to copy blocks.....\n");
                                        BlockTable bt = tm.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false) as BlockTable;//open the block table
                                        //check each block in the block table
                                        foreach (ObjectId btrId in bt)
                                        {
                                            BlockTableRecord btr = tm.GetObject(btrId, OpenMode.ForRead, false) as BlockTableRecord;

                                            if(!btr.IsAnonymous && !btr.IsLayout)
                                            {
                                                blockIdsToCopied.Add(btrId);
                                                loadedBlock.Add(btr.Name);
                                                ed.WriteMessage($"Block: {btr.Name}, will be loaded in {file}.\n");
                                            }
                                            btr.Dispose();
                                        }
                                    }

                                    //copy blocks from source to destination database
                                    if (blockIdsToCopied.Count > 0)
                                    {
                                        IdMapping mapping = new IdMapping();
                                        sourceDb.WblockCloneObjects(blockIdsToCopied, db.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);
                                        ed.WriteMessage($"copied blocks into {file}.....\n");
                                    }
                                }
                                catch { }
                                #endregion

                                #region determine what blocks needs to be added
                                //create the list of blocks that will be added
                                List<string> blocksTobeAdded = new List<string>();
                                foreach (string b in blockRequired)
                                {
                                    if (!blocks.Contains(b))
                                        blocksTobeAdded.Add(b);
                                }
                                #endregion

                                #region insert blocks into paper space
                                Point3d insPoint = new Point3d(0, 0, 0);
                                //string layoutName = "Layout1";//get all layout name
                                LayoutManager lm = LayoutManager.Current;
                                using (Transaction tr = db.TransactionManager.StartTransaction())
                                {
                                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                                    DBDictionary layoutDic = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, false) as DBDictionary;//get layout dictionary from current drawings
                                    foreach(DBDictionaryEntry entry in layoutDic)
                                    {
                                        ObjectId layoutId = entry.Value;//layout id
                                        Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;//get layout
                                        if (layout.LayoutName == "Model")
                                            continue;
                   
                                        BlockTableRecord btr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                        foreach (string blockName in blocksTobeAdded)
                                        {
                                            ObjectId blkId = bt[blockName];
                                            ed.WriteMessage($"ObjectId {blkId} in {layout.LayoutName}\n");

                                            //insert the block
                                            BlockReference blkRef = new BlockReference(insPoint, blkId);
                                            btr.UpgradeOpen();
                                            btr.AppendEntity(blkRef);
                                            tr.AddNewlyCreatedDBObject(blkRef, true);

                                            ed.WriteMessage($"Insert {blockName} in {layout.LayoutName}\n");
                                        }
                                    }
                                    
                                    tr.Commit();
                                }
                                //foreach (string blockName in blocksTobeAdded)
                                //{
                                //    try
                                //    {
                                //        CADops.InsertBlock(db, layoutName, blockName, insPoint);
                                //        ed.WriteMessage($"Insert {blockName} in {file}\n");
                                //    }
                                //    catch { }
                                //}
                                #endregion
                                db.SaveAs(file, DwgVersion.Current);//save changes
                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        { ed.WriteMessage(ex.ToString() + Environment.NewLine); }
                    }
                    else
                    {
                        ed.WriteMessage("File " + file + " does not exist.\n");
                    }
                }

                sourceDb.Dispose();
                
            }
            ed.WriteMessage($"Done.");
        }

    }
}
