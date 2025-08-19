using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Reflection;

namespace SolidWorksBatchDXF
{
    class Program
    {
        private const string OutputFolder = @"C:\Users\Lenovo\Desktop\AssemblyDrawings";

        static void Main(string[] args)
        {
            Console.WriteLine("Starting DXF batch export...");

            Directory.CreateDirectory(OutputFolder);

            SldWorks swApp = new SldWorks();
            swApp.Visible = true;

            // Pick an assembly file
            string asmFilePath = @"C:\Users\Lenovo\Desktop\SCHOOL\Personal\NL METAL\Solidworks Automation\AssemblyFiles\BL1000_20230528.SLDASM";
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
                    ProcessComponentRecursive(swApp, top, OutputFolder);
            }

            Console.WriteLine("DXF batch export complete!");
        }

        private static void ProcessComponentRecursive(SldWorks swApp, Component2 comp, string outputFolder)
        {
            if (comp == null) return;

            ModelDoc2 model = comp.GetModelDoc2();
            if (model == null) return;

            int docType = model.GetType();

            if (docType == (int)swDocumentTypes_e.swDocPART)
            {
                // If it's a part, generate drawing normally
                TryExportSheetMetalOrPartDrawing(swApp, model, comp, outputFolder);
            }
            else if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                // For subassembly, create a separate drawing
                string asmName = comp.Name2 ?? "SubAssembly";
                string safeName = string.Join("_", asmName.Split(Path.GetInvalidFileNameChars()));
                string subFolder = Path.Combine(outputFolder, safeName);
                Directory.CreateDirectory(subFolder);


                // Process children inside this subassembly
                object[] kids = (object[])comp.GetChildren();
                if (kids != null)
                {
                    foreach (Component2 child in kids)
                        ProcessComponentRecursive(swApp, child, subFolder);
                }
                return; // already processed children recursively
            }

            // For parts under top assembly (already processed in subassembly block)
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

            string name = comp.Name2 ?? "Part";
            string safeName = string.Join("_", name.Split(Path.GetInvalidFileNameChars()));
            string outPath = Path.Combine(outputFolder, safeName + ".slddrw");

            // 🔹 Find the longest linear edges along X and Y
            double maxX = 0, maxY = 0;
            Edge longestX = null, longestY = null;

            Feature feat = (Feature)partModel.FirstFeature();
            while (feat != null)
            {
                object[] bodies = feat.GetBody();
                if (bodies != null)
                {
                    foreach (Body2 body in bodies)
                    {
                        object[] edges = (object[])body.GetEdges();
                        if (edges != null)
                        {
                            foreach (Edge edge in edges)
                            {
                                Curve curve = edge.GetCurve();
                                if (curve != null && curve.GetType().Name.Contains("Line"))
                                {
                                    Vertex startVertex = edge.GetStartVertex();
                                    Vertex endVertex = edge.GetEndVertex();

                                    double[] startPt = (double[])startVertex.GetPoint();
                                    double[] endPt = (double[])endVertex.GetPoint();

                                    double dx = endPt[0] - startPt[0];
                                    double dy = endPt[1] - startPt[1];
                                    double dz = endPt[2] - startPt[2];
                                    double length = Math.Sqrt(dx * dx + dy * dy + dz * dz);

                                    if (Math.Abs(dx) > 0.99 && length > maxX) { maxX = length; longestX = edge; }
                                    if (Math.Abs(dy) > 0.99 && length > maxY) { maxY = length; longestY = edge; }
                                }
                            }
                        }
                    }
                }
                feat = feat.GetNextFeature() as Feature;
            }

            // Select the longest edges
            if (longestX != null)
                partModel.Extension.SelectByID2("", "EDGE", 0, 0, 0, true, 0, null, 0);
            if (longestY != null)
                partModel.Extension.SelectByID2("", "EDGE", 0, 0, 0, true, 0, null, 0);

            // Load drawing template
            string template = @"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2021\lang\english\sheetformat\A4 - Landscape.DRWDOT";
            ModelDoc2 drwModel = swApp.NewDocument(template, (int)swDwgPaperSizes_e.swDwgPaperA4size, 0, 0);
            if (drwModel == null) { Console.WriteLine($"❌ Failed to create drawing for {name}"); return; }

            DrawingDoc drwDoc = (DrawingDoc)drwModel;


            // Insert 3rd-angle standard views
            drwDoc.Create3rdAngleViews2(partModel.GetPathName());

            // Get the front view (skip sheet format)
            View frontView = drwDoc.GetFirstView().GetNextView();
            if (frontView != null)
            {
                // Access the current transform
                MathTransform trans = frontView.ModelToViewTransform;

                // Create a translation vector: move down by 0.05 m each iteration
                double[] translationArray = new double[] { 0, -0.05, 0 };
                MathUtility swMath = (MathUtility)swApp.GetMathUtility();
                MathTransform translation = swMath.CreateTransform(translationArray);

                // Loop 3 times
                for (int i = 0; i < 3; i++)
                {
                    MathTransform newTransform = trans.IMultiply(translation);
                    frontView.ModelToViewTransform = newTransform;

                    // Update trans so next iteration applies on top of the previous
                    trans = frontView.ModelToViewTransform;

                    // Repeat for top/right views if needed
                    View topView = frontView.GetNextView();
                    if (topView != null)
                    {
                        MathTransform topTrans = topView.ModelToViewTransform;
                        topView.ModelToViewTransform = topTrans.IMultiply(translation);
                    }
                }
            }


            // Auto-dimension based on preselected edges
            View firstView = drwDoc.GetFirstView().GetNextView(); // Skip sheet format
            if (firstView != null)
            {
                drwModel.Extension.SelectByID2(firstView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                int autoDimSuccess = drwDoc.AutoDimension(
                    (int)swAutodimEntities_e.swAutodimEntitiesBasedOnPreselect,
                    (int)swAutodimScheme_e.swAutodimSchemeBaseline,
                    (int)swAutodimHorizontalPlacement_e.swAutodimHorizontalPlacementAbove,
                    (int)swAutodimScheme_e.swAutodimSchemeBaseline,
                    (int)swAutodimVerticalPlacement_e.swAutodimVerticalPlacementRight
                );
                Console.WriteLine($"Auto-dimensioned view: {autoDimSuccess}");
            }

            // Add isometric view
            if (firstView != null)
            {
                object outlineObj = firstView.GetOutline();
                double[] outline = (double[])outlineObj;

                double xIso = outline[2] + 0.1;
                double yIso = outline[3] + 0.06;
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