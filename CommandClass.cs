using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Geometry;
using System.Runtime.InteropServices;
using System;
using System.Data;

namespace PlottingApplication

{

    public class SimplePlottingCommands

    {

        [DllImport("acad.exe",

                  CallingConvention = CallingConvention.Cdecl,

                  EntryPoint = "acedTrans")

        ]

        static extern int acedTrans(

          double[] point,

          IntPtr fromRb,

          IntPtr toRb,

          int disp,

          double[] result

        );


        [CommandMethod("winplot")]

        public void WindowPlot()

        {

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (var tr = db.TransactionManager.StartTransaction())

            {

                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                foreach (var ltrId in lt)

                {

                    bool lockZero = false;
                    if (ltrId != db.Clayer && (lockZero || ltrId != db.LayerZero))

                    {                                               
                        var ltr = (LayerTableRecord)tr.GetObject(ltrId, OpenMode.ForWrite);

                        if (ltr.Name == "TextDel")
                            ltr.IsPlottable = false;
                    }

                }

                tr.Commit();

            }


            ed.Regen();
            Object SysVarBackPlot = Application.GetSystemVariable("BACKGROUNDPLOT");
            Application.SetSystemVariable("BACKGROUNDPLOT", 0);
            string path = db.Filename;
            char ch = '\\';
            int lastIndexOfChar = path.LastIndexOf(ch);
            bool printOrient = true;

            path = path.Substring(0, lastIndexOfChar + 1);
            int iterationCount;
            double scale = 1;
            iterationCount = 0;
            DataSet printBase = new DataSet("AutocadPrintObject");
            System.Data.DataTable drawingTable = new System.Data.DataTable("Drawings");
            printBase.Tables.Add(drawingTable);
            System.Data.DataColumn idColumn = new System.Data.DataColumn("Id", Type.GetType("System.Int32"));
            idColumn.Unique = true; // столбец будет иметь уникальное значение
            idColumn.AllowDBNull = false; // не может принимать null
            idColumn.AutoIncrement = true; // будет автоинкрементироваться
            idColumn.AutoIncrementSeed = 1; // начальное значение
            idColumn.AutoIncrementStep = 1; // приращении при добавлении новой строки

            System.Data.DataColumn nameColumn = new System.Data.DataColumn("Name", Type.GetType("System.String"));
            System.Data.DataColumn leftPointXColumn = new System.Data.DataColumn("LeftPointX", Type.GetType("System.Double"));
            leftPointXColumn.DefaultValue = 0; // значение по умолчанию
            System.Data.DataColumn leftPointYColumn = new System.Data.DataColumn("LeftPointY", Type.GetType("System.Double"));
            leftPointXColumn.DefaultValue = 0; // значение по умолчанию
            System.Data.DataColumn rightPointXColumn = new System.Data.DataColumn("RightPointX", Type.GetType("System.Double"));
            System.Data.DataColumn rightPointYColumn = new System.Data.DataColumn("RightPointY", Type.GetType("System.Double"));
            //System.Data.DataColumn drawOrientationColumn = new System.Data.DataColumn("drawOrientation", Type.GetType("System.Bool"));

            drawingTable.Columns.Add(idColumn);
            drawingTable.Columns.Add(nameColumn);
            drawingTable.Columns.Add(leftPointXColumn);
            drawingTable.Columns.Add(leftPointYColumn);
            drawingTable.Columns.Add(rightPointXColumn);
            drawingTable.Columns.Add(rightPointYColumn);
            //drawingTable.Columns.Add(drawOrientationColumn);
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //Выделяем рамкой с координатами x1, y1, x2, y2
                //PromptSelectionResult prRes = ed.GetSelection();
                PromptSelectionResult selRes = ed.SelectAll();
                if (selRes.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nError! \n");
                    return;
                }

                ObjectId[] idsn = selRes.Value.GetObjectIds();
                foreach (ObjectId idn in idsn)
                {

                    Entity entity = (Entity)tr.GetObject(idn, OpenMode.ForRead);
                    if (entity.Layer == "Print")
                    {

                        DataRow row = drawingTable.NewRow();
                        row.ItemArray = new object[] { iterationCount, " ", 0, 0, 0, 0 };
                        drawingTable.Rows.Add(row);
                        Polyline pl = (Polyline)entity;
                        if (pl != null)
                        {
                            int nVertex = pl.NumberOfVertices; // Количество вершин полилинии

                            for (int i = 0; i < nVertex; i++)
                            {
                                Point3d p = pl.GetPoint3dAt(i); // Получаем очередную вершину
                                //ed.WriteMessage("\n\tp[{0}]={1}", i, p);

                                if (i == 0)
                                {
                                    //framePointsX[iterationCount] = p.X;
                                    drawingTable.Rows[iterationCount][2] = p.X;
                                    //framePointsY[iterationCount] = p.Y;
                                    drawingTable.Rows[iterationCount][3] = p.Y;

                                    //ed.WriteMessage("\n\tp[{0}]={1}{2}", i, x2, y2);
                                }
                                if (i == 2)
                                {
                                    //framePointsX1[iterationCount] = p.X;
                                    drawingTable.Rows[iterationCount][4] = p.X;
                                    //framePointsY1[iterationCount] = p.Y;
                                    drawingTable.Rows[iterationCount][5] = p.Y;
                                    //ed.WriteMessage("\n\tp[{0}]={1}{2}", i, x3, y3);
                                }

                            }

                            //Point3d x = pl.GetPoint3dAt(0);
                            //Point3d x1 = pl.GetPoint3dAt(2);
                            //ed.WriteMessage("\n\tp[0]={0}, {1}", framePointsX[iterationCount], framePointsY[iterationCount]);
                            //ed.WriteMessage("\n\tp[2]={0}, {1}", framePointsX1[iterationCount], framePointsY1[iterationCount]);
                            //ed.WriteMessage(x2.ToString(), y2.ToString(), "\n", x3.ToString(), y3.ToString(), "\n");
                        }
                        iterationCount = iterationCount + 1;
                    }

                }
                ed.WriteMessage("\n\titerationCount={0}", iterationCount);
                for (int i = 0; i < iterationCount; i++)
                {
                    //var selectedBooks = drawingTable.Select("Id < 120");
                    //foreach (var b in selectedBooks)
                    // ed.WriteMessage("\n\t{0} - {1}", b["id"], b["LeftPointX"]);
                    PromptSelectionResult prRes;
                    double x1 = Convert.ToDouble(drawingTable.Rows[i][2]);
                    double y1 = Convert.ToDouble(drawingTable.Rows[i][3]);
                    double x2 = Convert.ToDouble(drawingTable.Rows[i][4]);
                    double y2 = Convert.ToDouble(drawingTable.Rows[i][5]);
                    printOrient = (Math.Abs(y2 - y1) > Math.Abs(x2 - x1)) ? true : false;

                    //ed.WriteMessage("\n\tPoint1= {0}, {1}", x1, y1);
                    //ed.WriteMessage("\n\tPoint2= {0}, {1}", x2, y2);
                    prRes = ed.SelectCrossingWindow(new Point3d(x1, y1, 0),
                                                    new Point3d(x2, y2, 0));
                    if (prRes.Status != PromptStatus.OK)
                        return;
                    //Создаем коллекцию выделенных объектов
                    ObjectIdCollection objIds = new ObjectIdCollection();

                    ObjectId[] objIdArray = prRes.Value.GetObjectIds();

                    string name = "123";
                    string PDFname = "123";

                    //Перебираем все объекты в коллекции 

                    foreach (ObjectId id in objIdArray)
                    {

                        //Приводим все объекты к типу object
                        Entity entity = (Entity)tr.GetObject(id, OpenMode.ForRead);
                        //Фильтруем объекты, которые будут переноситься на новый лист
                        if (entity.Layer != "TextDel")
                        {

                            //Добавление нужных объектов
                            objIds.Add(id);

                        }

                        //Ищем в выбранной рамке текст на слое Text
                        if ((entity.GetType() == typeof(DBText)) & (entity.Layer == "Text"))
                        {

                            //Получаем значение объекта DBText на слое Text
                            DBText dt = (DBText)entity;
                            if (dt != null)
                            {
                                //namePrint[iterationCount] = dt.TextString;
                                drawingTable.Rows[i][1] = dt.TextString;
                                //name = path + namePrint[iterationCount] + ".dwg";
                                name = path + drawingTable.Rows[i][1].ToString() + ".dwg";
                                PDFname = path + drawingTable.Rows[i][1].ToString() + ".pdf";
                            }

                            //acad.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\nLayer:{0}; Type:{1}; Color: {2},{3},{4}\n",
                            //entity.Layer, entity.GetType().ToString(), entity.Color.Red.ToString(), entity.Color.Green.ToString(), entity.Color.Blue.ToString()));
                            //acad.DocumentManager.MdiActiveDocument.Editor.WriteMessage(name.ToString());
                            //ed.WriteMessage("\n\tЧертеж {0} распечатан! Полный путь: {1}", namePrint[iterationCount], name);
                            //acad.DocumentManager.MdiActiveDocument.Editor.WriteMessage(name.ToString());
                            //acad.DocumentManager.MdiActiveDocument.Editor.WriteMessage(path.ToString());
                        }

                    }

                    using (Database newDb = new Database(true, false))

                    {
                        db.Wblock(newDb, objIds, Point3d.Origin, DuplicateRecordCloning.Ignore);

                        string FileName = name;

                        newDb.SaveAs(FileName, DwgVersion.Newest);
                    }
                    ed.WriteMessage("\n\tЧертеж {0} сохранен отдельным файлом! Полный путь: {1}", drawingTable.Rows[i][1].ToString(), name);
                    Point3d first = new Point3d(x1, y1, 0);
                    Point3d second = new Point3d(x2, y2, 0);

                    // Перевод координат СК UCS в DCS

                    ResultBuffer rbFrom =

                      new ResultBuffer(new TypedValue(5003, 1)),

                                rbTo =

                      new ResultBuffer(new TypedValue(5003, 2));

                    double[] firres = new double[] { 0, 0, 0 };

                    double[] secres = new double[] { 0, 0, 0 };
                    acedTrans(first.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
                    acedTrans(second.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
                    Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);

                    BlockTableRecord btr =

                      (BlockTableRecord)tr.GetObject(

                        db.CurrentSpaceId,

                        OpenMode.ForRead

                      );

                    Layout lo =

                      (Layout)tr.GetObject(

                        btr.LayoutId,

                        OpenMode.ForRead

                      );
                    PlotInfo pi = new PlotInfo();

                    pi.Layout = btr.LayoutId;
                    PlotSettings ps = new PlotSettings(lo.ModelType);

                    ps.CopyFrom(lo);

                    PlotSettingsValidator psv = PlotSettingsValidator.Current;

                    psv.SetPlotWindowArea(ps, window);
                    psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                    psv.SetUseStandardScale(ps, true);
                    psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                    psv.SetPlotPaperUnits(ps, PlotPaperUnit.Millimeters);
                    psv.SetPlotCentered(ps, false);
                    psv.SetPlotRotation(ps, PlotRotation.Degrees000);
                    psv.SetUseStandardScale(ps, false);
                    psv.SetPlotConfigurationName(ps, "DWG To PDF.pc3", "ISO_A0_(841.00_x_1189.00_MM)");

                    if (printOrient == true)
                    {
                        psv.SetPlotRotation(ps, PlotRotation.Degrees090);
                        psv.SetPlotConfigurationName(ps, "DWG To PDF.pc3", "ISO_expand_A2_(594.00_x_420.00_MM)");
                        if ((Math.Abs(secres[0] - firres[0]) / 420) > (Math.Abs(secres[1] - firres[1]) / 594))
                            scale = (Math.Abs(secres[0] - firres[0]) + 30) / 420;
                        else
                            scale = (Math.Abs(secres[1] - firres[1]) + 30) / 594;
                        psv.SetCustomPrintScale(ps, new CustomScale(1, 1.004 * scale));
                    }

                    else
                    {
                        psv.SetPlotRotation(ps, PlotRotation.Degrees090);
                        psv.SetPlotConfigurationName(ps, "DWG To PDF.pc3", "ISO_expand_A1_(594.00_x_841.00_MM)");
                        //psv.SetPlotConfigurationName(ps, "DWG To PDF.pc3", "ISO_expand_A2_(594.00_x_420.00_MM)");
                        if ((Math.Abs(secres[0] - firres[0]) / 841) > (Math.Abs(secres[1] - firres[1]) / 594))
                            scale = (Math.Abs(secres[0] - firres[0]) + 30) / 841;
                        else
                            scale = (Math.Abs(secres[1] - firres[1]) + 30) / 594;
                        psv.SetCustomPrintScale(ps, new CustomScale(1, 1.004 * scale));
                    }
                    //0.7063020
                    //1.41428571429
                    //
                    //psv.SetPlotRotation(ps, printOrient == true ? PlotRotation.Degrees000 : PlotRotation.Degrees090);
                    //psv.SetPlotConfigurationName(ps, "DWG To PDF.pc3", "ISO_A0_(841.00_x_1189.00_MM)");
                    psv.SetPlotCentered(ps, true);
                    pi.OverrideSettings = ps;

                    PlotInfoValidator piv =

                      new PlotInfoValidator();

                    piv.MediaMatchingPolicy =

                      MatchingPolicy.MatchEnabled;

                    piv.Validate(pi);

                    if (PlotFactory.ProcessPlotState ==

                        ProcessPlotState.NotPlotting)

                    {

                        PlotEngine pe =

                          PlotFactory.CreatePublishEngine();

                        using (pe)

                        {

                            // Создаем Progress Dialog (с возможностью отмены пользователем)


                            PlotProgressDialog ppd =

                              new PlotProgressDialog(false, 1, true);

                            using (ppd)

                            {

                                ppd.set_PlotMsgString(

                                  PlotMessageIndex.DialogTitle,

                                  "Custom Plot Progress"

                                );

                                ppd.set_PlotMsgString(

                                  PlotMessageIndex.CancelJobButtonMessage,

                                  "Cancel Job"

                                );

                                ppd.set_PlotMsgString(

                                  PlotMessageIndex.CancelSheetButtonMessage,

                                  "Cancel Sheet"

                                );

                                ppd.set_PlotMsgString(

                                  PlotMessageIndex.SheetSetProgressCaption,

                                  "Sheet Set Progress"

                                );

                                ppd.set_PlotMsgString(

                                  PlotMessageIndex.SheetProgressCaption,

                                  "Sheet Progress"

                                );

                                ppd.LowerPlotProgressRange = 0;
                                ppd.UpperPlotProgressRange = 100;
                                ppd.PlotProgressPos = 0;

                                // Начинаем печать

                                ppd.OnBeginPlot();
                                ppd.IsVisible = true;
                                pe.BeginPlot(ppd, null);

                                pe.BeginDocument(

                                  pi,

                                  doc.Name,

                                  null,

                                  1,

                                  true, 

                                  PDFname

                                );

                                ppd.OnBeginSheet();
                                ppd.LowerSheetProgressRange = 0;
                                ppd.UpperSheetProgressRange = 100;
                                ppd.SheetProgressPos = 0;

                                PlotPageInfo ppi = new PlotPageInfo();

                                pe.BeginPage(

                                  ppi,

                                  pi,

                                  true,

                                  null

                                );

                                pe.BeginGenerateGraphics(null);
                                pe.EndGenerateGraphics(null);
                                pe.EndPage(null);
                                ppd.SheetProgressPos = 100;
                                ppd.OnEndSheet();

                                // Завершаем работу с документом

                                pe.EndDocument(null);

                                // Завершаем печать

                                ppd.PlotProgressPos = 100;
                                ppd.OnEndPlot();
                                pe.EndPlot(null);

                            }

                        }

                    }

                    else

                    {

                        ed.WriteMessage(

                          "\nAnother plot is in progress."

                        );

                    }
                }

            }

            using (var tr = db.TransactionManager.StartTransaction())

            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                foreach (var ltrId in lt)

                {
                    bool lockZero = false;

                    if (ltrId != db.Clayer && (lockZero || ltrId != db.LayerZero))

                    {
                        // Открываем слой для записи

                        var ltr = (LayerTableRecord)tr.GetObject(ltrId, OpenMode.ForWrite);
                        
                    }

                }

                tr.Commit();

            }

        }
        
        public void Terminate()
        {

        }
    }

}
