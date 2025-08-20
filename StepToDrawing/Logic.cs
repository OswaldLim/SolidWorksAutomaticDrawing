using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SolidWorksBatchDXF
{
    public static class Logic
    {
        private static HashSet<string> exportedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public static void RunBatchExport(string inputFolder, string outputFolder)
        {
            Directory.CreateDirectory(outputFolder);

            SldWorks swApp = new SldWorks();
            swApp.Visible = true;

            string asmFilePath = $@"{inputFolder}";
            int errs = 0, warns = 0;

            ModelDoc2 asmDoc = swApp.OpenDoc6(
                asmFilePath,
                (int)swDocumentTypes_e.swDocASSEMBLY,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errs,
                ref warns);

            if (asmDoc == null)
            {
                Console.WriteLine("Failed to open assembly.");
                return;
            }

            AssemblyDoc asm = (AssemblyDoc)asmDoc;
            object[] topComps = (object[])asm.GetComponents(true);

            if (topComps != null)
            {
                foreach (Component2 top in topComps)
                    ProcessComponentRecursive(swApp, top, outputFolder);
            }
        }

        private static void ProcessComponentRecursive(SldWorks swApp, Component2 comp, string outputFolder)
        {
            if (comp == null) return;

            ModelDoc2 model = comp.GetModelDoc2();
            if (model == null) return;

            int docType = model.GetType();

            if (docType == (int)swDocumentTypes_e.swDocPART)
            {
                TryExportSheetMetalOrPartDrawing(swApp, model, comp, outputFolder);
            }
            else if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                string asmName = comp.Name2 ?? "SubAssembly";
                string asmBase = Path.GetFileNameWithoutExtension(asmName);
                string safeName = string.Concat(
                    asmBase.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries)
                );
                string subFolder = Path.Combine(outputFolder, safeName);
                Directory.CreateDirectory(subFolder);

                object[] kids = (object[])comp.GetChildren();
                if (kids != null)
                {
                    foreach (Component2 child in kids)
                        ProcessComponentRecursive(swApp, child, subFolder);
                }
                return;
            }

            object[] children = (object[])comp.GetChildren();
            if (children != null)
            {
                foreach (Component2 child in children)
                    ProcessComponentRecursive(swApp, child, outputFolder);
            }
        }

        private static void TryExportSheetMetalOrPartDrawing(SldWorks swApp, ModelDoc2 partModel, Component2 comp, string outputFolder)
        {
            if (partModel == null) return;

            string modelPath = partModel.GetPathName();
            if (string.IsNullOrEmpty(modelPath)) return;

            if (exportedParts.Contains(modelPath))
            {
                Console.WriteLine($"⚠️ Skipping duplicate part: {modelPath}");
                return;
            }
            exportedParts.Add(modelPath);

            string name = Path.GetFileNameWithoutExtension(modelPath);
            if (name.EndsWith(".step", StringComparison.OrdinalIgnoreCase))
                name = name.Substring(0, name.Length - 5);

            string safeName = string.Concat(name.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries));
            string outPath = Path.Combine(outputFolder, safeName + ".slddrw");


            string template = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "A4 - Landscape.DRWDOT");

            ModelDoc2 drwModel = swApp.NewDocument(template, (int)swDwgPaperSizes_e.swDwgPaperA4size, 0, 0);
            if (drwModel == null)
            {
                Console.WriteLine($"❌ Failed to create drawing for {name}");
                return;
            }

            DrawingDoc drwDoc = (DrawingDoc)drwModel;
            drwDoc.Create3rdAngleViews2(partModel.GetPathName());

            //Get the front view(skip sheet format)
            View frontView = drwDoc.GetFirstView().GetNextView();
            while (frontView != null)
            {
                double[] vPos = frontView.Position;
                vPos[0] += 0.01;
                vPos[1] -= 0.01;
                frontView.Position = vPos;

                if (frontView != null)
                {
                    drwModel.Extension.SelectByID2(frontView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                    int autoDimSuccess = drwDoc.AutoDimension(
                        (int)swAutodimEntities_e.swAutodimEntitiesBasedOnPreselect,
                        (int)swAutodimScheme_e.swAutodimSchemeBaseline,
                        (int)swAutodimHorizontalPlacement_e.swAutodimHorizontalPlacementAbove,
                        (int)swAutodimScheme_e.swAutodimSchemeBaseline,
                        (int)swAutodimVerticalPlacement_e.swAutodimVerticalPlacementRight
                    );
                    Console.WriteLine($"Auto-dimensioned view: {autoDimSuccess}");
                }

                object[] anns = frontView.GetAnnotations();
                Console.WriteLine(frontView.GetAnnotationCount());

                if (anns != null)
                {
                    Console.WriteLine("Inside");
                    List<(Annotation ann, double val)> horizDims = new List<(Annotation, double)>();
                    List<(Annotation ann, double val)> vertDims = new List<(Annotation, double)>();

                    foreach (object a in anns)
                    {
                        Annotation swAnn = (Annotation)a;
                        DisplayDimension specAnn = swAnn.GetSpecificAnnotation();

                        if (specAnn != null && specAnn is DisplayDimension swDispDim)
                        {
                            Dimension swDim = swDispDim.GetDimension();
                            int type = swDispDim.GetType();

                            double val = swDim.GetSystemValue2(""); // always meters

                            // Try to determine orientation
                            // Use DisplayData line to figure out angle of the dimension line
                            DisplayData swDispData = swDispDim.GetDisplayData();

                            // Get first line segment of this annotation (dimension line)
                            object raw = swDispData.GetLineAtIndex3(0);
                            if (raw is Array arr && arr.Length >= 10)
                            {
                                double startX = (double)arr.GetValue(4);
                                double startY = (double)arr.GetValue(5);
                                double endX = (double)arr.GetValue(7);
                                double endY = (double)arr.GetValue(8);

                                double dx = endX - startX;
                                double dy = endY - startY;

                                // angle in radians
                                double angle = Math.Atan2(dy, dx);

                                // near 0° or 180° → horizontal
                                // near 90° or 270° → vertical
                                if (Math.Abs(Math.Sin(angle)) < 0.5) // closer to horizontal
                                    horizDims.Add((swAnn, val));
                                else
                                    vertDims.Add((swAnn, val));
                            }
                        }
                    }

                    // Sort both lists descending by dimension value
                    horizDims = horizDims.OrderByDescending(d => d.val).ToList();
                    vertDims = vertDims.OrderByDescending(d => d.val).ToList();

                    Console.WriteLine($"Horiz: {horizDims.Count}, Vert: {vertDims.Count}");


                    int deleted = 0;

                    // Keep only the 3 largest, delete the rest
                    for (int i = 1; i < horizDims.Count - 1; i++)
                    {
                        Annotation swAnn = horizDims[i].ann;
                        swAnn.Select2(false, 0);
                        drwModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                        deleted++;
                    }
                    Console.WriteLine($"Deleted {deleted} horizontal annotations");
                    deleted = 0;

                    // Trim vertical dimensions
                    for (int i = 1; i < vertDims.Count - 1; i++)
                    {
                        Annotation swAnn = vertDims[i].ann;
                        swAnn.Select2(false, 0);
                        drwModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                        deleted++;
                    }
                    Console.WriteLine($"Deleted {deleted} vertical annotations");
                }

                bool statuss = drwModel.Extension.AlignDimensions((int)swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.001);

                Console.WriteLine(statuss);

                frontView = frontView.GetNextView();
                drwDoc.EditRebuild();
            }

            // Auto-dimension based on preselected edges
            View firstView = drwDoc.GetFirstView().GetNextView(); // Skip sheet format


            // Add isometric view
            if (firstView != null)
            {
                object outlineObj = firstView.GetOutline();
                double[] outline = (double[])outlineObj;

                double xIso = outline[2] + 0.12;
                double yIso = outline[3] + 0.05;
                double scale = firstView.ScaleDecimal;

                View isoView = drwDoc.CreateDrawViewFromModelView3(partModel.GetPathName(), "*Isometric", xIso, yIso, scale);
                Console.WriteLine(isoView != null ? "Isometric view added." : "Failed to add isometric view.");
            }

            // Save drawing
            int saveErr = 0, saveWarn = 0;
            bool status = drwModel.Extension.SaveAs(outPath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref saveErr, ref saveWarn);

            Console.WriteLine(status ? $"✅ Drawing saved: {outPath}" : $"❌ Failed to save drawing (Err={saveErr}, Warn={saveWarn})");

            swApp.CloseDoc(drwModel.GetTitle());
        }
    }
}
