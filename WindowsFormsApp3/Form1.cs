using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Dialog.UIControls;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                string dirName = new DirectoryInfo(textBox1.Text).Name;
                DirectoryInfo parentDir = Directory.GetParent(textBox1.Text);
                Directory.Delete(textBox1.Text,true);
                string newDirectory = Path.Combine(parentDir.FullName, "CNC", "NC1");
                Directory.CreateDirectory(newDirectory);

                //Copy all *.nc1 files created in the CNC folder.
                string[] files = Directory.GetFiles(parentDir.FullName, "*.nc1");
                foreach (var item in files)
                {
                    string fileName = Path.GetFileName(item);
                    fileName = fileName.Replace(dirName, string.Empty);
                    File.Move(item, Path.Combine(newDirectory, fileName));
                }
                //Directory.EnumerateFiles(parentDir.FullName, "*.nc1").ToList().ForEach(x => File.Move(x, Path.Combine(newDirectory, Path.GetFileName(x))));
            }
        }

        private void cmdPrint_Click(object sender, EventArgs e)
        {
            DrawingHandler drawingHandler = new DrawingHandler();
            PrintAttributes printAttributes = new PrintAttributes();
            //LineTypeAttributes lineTypeAttributes = new LineTypeAttributes();

            //lineTypeAttributes.Color
            printAttributes.Scale = 1;
            //printAttributes.PrintToMultipleSheet = false;
            printAttributes.NumberOfCopies = 1;
            printAttributes.Orientation = DotPrintOrientationType.Auto;
            printAttributes.PrintArea = DotPrintAreaType.EntireDrawing;
            printAttributes.ScalingType = DotPrintScalingType.Scale;
            //printAttributes.PrinterInstance = "3. 8 1/2x11 ------- PDF";
            


            DrawingEnumerator selectedDrawings = drawingHandler.GetDrawingSelector().GetSelected();
            
            while (selectedDrawings.MoveNext())
            {

                Drawing currentDrawing = selectedDrawings.Current;
                

                string printerInstanceName = PrinterInstance(currentDrawing.GetSheet().Width, currentDrawing.GetSheet().Height);
                printAttributes.PrinterInstance = printerInstanceName;
                //drawingHandler.SetActiveDrawing(currentDrawing, false);


                // Make sure it is a valid drawing
                if (currentDrawing != null)
                {
                    // Check that the drawing is up-to-date before printing
                    if (currentDrawing.UpToDateStatus == DrawingUpToDateStatus.DrawingIsUpToDate)
                    {
                        string path = "C:\\temp\\opex\\test";
                        string fileName = Path.Combine(path, PrintToFile(currentDrawing) + ".pdf");
                        drawingHandler.PrintDrawing(currentDrawing, printAttributes);
                    }
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Text = EnvironmentVariables.GetEnvironmentVariable(textBox1.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                var strValue = string.Empty;
                bool result = Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption(textBox1.Text, ref strValue);
                MessageBox.Show(strValue);
            }
        }


        private string PrintToFile(Drawing currentDrawing)
        {
            string DrawingType = GetDrawingTypeCharacter(currentDrawing);
            string output_file = GetOutputFile(DrawingType);
            string converted_file = ConvertOutputFile(output_file, currentDrawing);
            


            return converted_file;
        }

        private string ConvertOutputFile(string output_file, Drawing currentDrawing)
        {
            string convertedFile = string.Empty;
            string drawingName = RemoveBrackets(currentDrawing.Mark);
            string drawingTitle = currentDrawing.Name;
            string revision = string.Empty;
            int RevNO = 0;

            Beam DWG = new Beam
            {
                Identifier =
                        (Identifier)currentDrawing.GetType().GetProperty("Identifier", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(currentDrawing, null)
            };

            DWG.GetReportProperty("REVISION.NUMBER", ref RevNO);

            if (RevNO > 0)
            {
                String MArkfortest = String.Empty;
                DWG.GetReportProperty("mark_" + RevNO, ref MArkfortest);
                revision = " - Rev " + MArkfortest;
            }

            convertedFile = drawingName + " - " + drawingTitle + revision;

            return convertedFile;
        }

        private static string GetOutputFile(string DrawingType)
        {
            string output_file = string.Empty;

            switch (DrawingType)
            {
                case "A": //Assembly Drawing
                    output_file = EnvironmentVariables.GetEnvironmentVariable("XS_DRAWING_PLOT_FILE_NAME_A");
                    break;
                case "W": //Single Part Drawing
                    output_file = EnvironmentVariables.GetEnvironmentVariable("XS_DRAWING_PLOT_FILE_NAME_W");
                    break;
                case "C": //CastUnit Drawing
                    output_file = EnvironmentVariables.GetEnvironmentVariable("XS_DRAWING_PLOT_FILE_NAME_C");
                    break;
                case "M": //MultiDrawing
                    output_file = EnvironmentVariables.GetEnvironmentVariable("XS_DRAWING_PLOT_FILE_NAME_M");
                    break;
                case "G": //GADrawing
                    output_file = EnvironmentVariables.GetEnvironmentVariable("XS_DRAWING_PLOT_FILE_NAME_G");
                    break;
                default:
                    output_file = "Unknown File";
                    break;
            }

            return output_file;
        }

        private string GetDrawingTypeCharacter(Drawing DrawingInstance)
        {
            string Result = "U"; // Unknown drawing

            if (DrawingInstance is GADrawing)
            {
                Result = "G";
            }
            else if (DrawingInstance is AssemblyDrawing)
            {
                Result = "A";
            }
            else if (DrawingInstance is CastUnitDrawing)
            {
                Result = "C";
            }
            else if (DrawingInstance is MultiDrawing)
            {
                Result = "M";
            }
            else if (DrawingInstance is SinglePartDrawing)
            {
                Result = "W";
            }

            return Result;
        }

        private void cmdPrintToFile_Click(object sender, EventArgs e)
        {
            DrawingHandler MyDrawingHandler = new DrawingHandler();
            if (MyDrawingHandler.GetConnectionStatus())
            {
                DrawingEnumerator SelectedDrawings = MyDrawingHandler.GetDrawingSelector().GetSelected();
                while (SelectedDrawings.MoveNext())
                {
                    Drawing currentDrawing = SelectedDrawings.Current;
                    //Identifier =
                    //            (Identifier)drawing.GetType().GetProperty("Identifier", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(drawing, null)

                    PrintToFile(currentDrawing);
                }
            }
        }

        private string RemoveBrackets(string inputString)
        {
            StringBuilder builder = new StringBuilder(inputString);
            builder.Replace("[", "");
            builder.Replace("]", "");
            builder.Replace(".", "");

            return builder.ToString();
        }

        private void SelectDrawingsWithRevision()
        {
            DrawingHandler drawingHandler = new DrawingHandler();
            Model teklaModel = new Model();
            PrintAttributes printAttributes = new PrintAttributes();
            printAttributes.Scale = 1;
            printAttributes.PrintToMultipleSheet = false;
            printAttributes.NumberOfCopies = 1;
            printAttributes.Orientation = DotPrintOrientationType.Auto;
            printAttributes.PrintArea = DotPrintAreaType.EntireDrawing;
            printAttributes.ScalingType = DotPrintScalingType.Auto;
            printAttributes.PrinterInstance = "PDFCreator";

            string path = "C:\\temp\\opex\\OFA\\New folder";
            Directory.CreateDirectory(path);

            if (teklaModel.GetConnectionStatus())
            {
                

                List<int> drawingWithRevision = new List<int>();
                DrawingEnumerator drawingsEnum = drawingHandler.GetDrawings();

                while (drawingsEnum.MoveNext())
                {
                    int revision = 0;
                    var drawing = drawingsEnum.Current as Drawing;
                    Beam DWG = new Beam
                    {
                        Identifier =
                                (Identifier)drawing.GetType().GetProperty("Identifier", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(drawing, null)
                    };

                    DWG.GetReportProperty("REVISION.NUMBER", ref revision);
                    string DrawingType = GetDrawingTypeCharacter(drawing);

                    //Should have revision and DrawingType is not SinglePart "W"
                    if ((revision > 0) && (DrawingType != "W"))
                    {
                        DWG.Select();
                        string filename = Path.Combine(path, PrintToFile(drawing) + ".pdf");
                        Console.WriteLine("Printing: " + filename);
                        drawingHandler.PrintDrawing(drawing, printAttributes, filename);
                    }


                }


                string defaultReport = "FabTrol_Drawing_List-v30.xsr";
                
                string template = "450   FabTrol Drawing List-v30";
                string fullPath = Path.Combine(path, defaultReport);
                Operation.CreateReportFromSelected(template, fullPath, "Drawing List", "", "");
            }
        }

        private void cmdSelect_Click(object sender, EventArgs e)
        {
            SelectDrawingsWithRevision();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string path = "C:\\temp\\opex\\OFA\\test";
            string fullPath = Path.Combine(path, textBox1.Text);
            Directory.CreateDirectory(fullPath);

            MessageBox.Show(fullPath.ToString());


        }

        private void button5_Click(object sender, EventArgs e)
        {
            Model CurrentModel = new Model();
            ModelInfo Info = CurrentModel.GetInfo();

            string ModelDatabase = Path.Combine(Info.ModelPath, Info.ModelName);

            MessageBox.Show(ModelDatabase);
            //Console.WriteLine(directoryInfo.Attributes);
        }

        private string PrinterInstance(double width, double height)
        {
            //Todo: Remove hardcoded Printer Instance and add this to fabricator settings.
            string PrinterInstanceName = "3. 8 1/2x11 ------- PDF"; //Default Printer instance
            if ((width > 400 && width < 450) && (height < 280 && height > 260))
            {
                PrinterInstanceName = "4. 11x17 ---------- PDF";
            }
            else if ((width > 850  && width < 1000) && (height < 600 && height > 550))
            {
                PrinterInstanceName = "5. 24x36 ---------- PDF";
            }
            else if ((width > 180 && width < 250) && (height < 300 && height > 240))
            {
                PrinterInstanceName = "3. 8 1/2x11 ------- PDF";
            }

            return PrinterInstanceName;
        }

        private void cmdConvert2DXF_Click(object sender, EventArgs e)
        {
            //string defFile = "tekla_dstv2dxf_imperial.def";
            //string exeFile = "tekla_dstv2dxf.exe";
            //string dxfOutputFolder = "NC_dxf";

            //try
            //{
            //    string modelDir;
            //    string dstvDir;
            //    string XS_Variable = string.Empty;
            //    TeklaStructuresSettings.GetAdvancedOption("XS_DIR", ref XS_Variable);

            //    Model model = new Model();
            //    ModelInfo Info = model.GetInfo();

            //    /** Get model directory **/
            //    modelDir = Info.ModelPath;

            //    /** Check for existence of dstv2dxf.exe **/
            //    string defFileFullPath = Path.Combine(XS_Variable, "nt\\dstv2dxf", defFile);
            //    Console.WriteLine(defFileFullPath);
            //    if (File.Exists(defFileFullPath))
            //    {
            //        dstvDir = Path.Combine(XS_Variable, "nt\\dstv2dxf");
            //    }
            //    else
            //    {
            //        MessageBox.Show("The conversion definition file " + defFile + " and/or the dstv2dxf directory could not be found.\n\npleases modify the macro script to point to the correct directory.");
            //        return;
            //    }

            //    /** Copy dstv2dxf.exe to the model folder **/
            //    if (!File.Exists(Path.Combine(modelDir, exeFile)))
            //    {
            //        new FileInfo(Path.Combine(dstvDir, "tekla_dst+v2dxf.exe")).CopyTo(Path.Combine(modelDir, "tekla_dstv2dxf.exe"), true);
            //        MessageBox.Show("File copied!");
            //    }

            //    /** Generate the "model local" version of the def file **/
            //    StreamReader sr = new StreamReader(Path.Combine(dstvDir, defFile), Encoding.Default);
            //    StreamWriter sw = new StreamWriter(Path.Combine(modelDir, defFile), false, Encoding.Default);
            //    string line = " ";
            //    string inputDir = Path.Combine(modelDir, "NC_files");
            //    while ((line = sr.ReadLine()) != null)
            //    {
            //        if (line.Trim().IndexOf("INPUT_FILE_DIR") >= 0)
            //            sw.WriteLine("INPUT_FILE_DIR=" + inputDir);
            //        else if (line.Trim().IndexOf("OUTPUT_FILE_DIR") >= 0)
            //            sw.WriteLine("OUTPUT_FILE_DIR=" + Path.Combine(modelDir, dxfOutputFolder));
            //        else
            //            sw.WriteLine(line);
            //    }
            //    sw.Flush();
            //    sw.Close();


            //    /** Check for dxfOutputFolder. Creat it if it doesn't exist **/
            //    if (!Directory.Exists(Path.Combine(modelDir, dxfOutputFolder)))
            //        Directory.CreateDirectory(Path.Combine(modelDir, dxfOutputFolder));

            //    /** Launch the local dstv2dxf **/
            //    defFile = Path.Combine(modelDir, defFile);
            //    //MessageBox.Show(Path.Combine(modelDir, "tekla_dstv2dxf.exe"), defFile);
            //    Console.WriteLine(Path.Combine(modelDir, defFile));
            //    Console.WriteLine(Path.Combine(modelDir, "tekla_dstv2dxf.exe"));
            //    Console.WriteLine(defFile);
            //    Process NCDXFConv = new Process
            //    {
            //        EnableRaisingEvents = false
            //    };
            //    NCDXFConv.StartInfo.CreateNoWindow = true;
            //    NCDXFConv.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            //    NCDXFConv.StartInfo.FileName = Path.Combine(modelDir, "tekla_dstv2dxf.exe");
            //    NCDXFConv.StartInfo.Arguments = " -cfg " + defFile + " -m batch -f *.nc1";
            //    NCDXFConv.Start();



            //    NCDXFConv.WaitForExit();
            //    NCDXFConv.Close();

            //    if (MessageBox.Show("Do you want to open the folder, which contains the DXF files?", "Tekla Structures", 
            //        MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            //    {
            //        /** Open a file explorer window in the output folder **/
            //        Process.Start(Path.Combine(modelDir, dxfOutputFolder));
            //        //System.Diagnostics.Process Explorer = new System.Diagnostics.Process();
            //        //Explorer.EnableRaisingEvents = false;
            //        //Explorer.StartInfo.FileName = "explorer";
            //        //Explorer.StartInfo.Arguments = "\"" + @modelDir + dxfOutputFolder + "\"";
            //        //Explorer.Start();
            //    }

            //}
            //catch (Exception)
            //{
            //    throw;
            //}

            Convert2DXF(textBox1.Text, textBox2.Text);
            MessageBox.Show("Done");
            
        }


        public bool Convert2DXF(string sourcePath, string destinationPath)
        {
            string defFile = "tekla_dstv2dxf_imperial.def";
            string exeFile = "tekla_dstv2dxf.exe";

            try
            {
                string modelDir;
                string dstvDir;
                string XS_Variable = string.Empty;
                TeklaStructuresSettings.GetAdvancedOption("XS_DIR", ref XS_Variable);

                Model model = new Model();
                ModelInfo modelInfo = model.GetInfo();
                /** Get model directory **/
                modelDir = modelInfo.ModelPath;

                /** Check for existence of dstv2dxf.exe **/
                string defFileFullPath = Path.Combine(XS_Variable, "nt\\dstv2dxf", defFile);
                Console.WriteLine(defFileFullPath);
                if (File.Exists(defFileFullPath))
                {
                    dstvDir = Path.Combine(XS_Variable, "nt\\dstv2dxf");
                }
                else
                {
                    //Generate error log
                    //MessageBox.Show("The conversion definition file " + defFile + " and/or the dstv2dxf directory could not be found.\n\npleases modify the macro script to point to the correct directory.");
                    return false;
                }

                /** Copy dstv2dxf.exe to the model folder **/
                if (!File.Exists(Path.Combine(modelDir, exeFile)))
                {
                    new FileInfo(Path.Combine(dstvDir, exeFile)).CopyTo(Path.Combine(modelDir, exeFile), true);
                }

                /** Generate the "model local" version of the def file **/
                StreamReader sr = new StreamReader(Path.Combine(dstvDir, defFile), Encoding.Default);
                StreamWriter sw = new StreamWriter(Path.Combine(modelDir, defFile), false, Encoding.Default);
                string line = " ";

                /***/
                //sourcePath = Path.Combine(modelDir, "NC_files");
                //destinationPath = Path.Combine(modelDir, "NC_dxf");
                /***/

                while ((line = sr.ReadLine()) != null)
                {
                    if (line.Trim().IndexOf("INPUT_FILE_DIR") >= 0)
                        sw.WriteLine("INPUT_FILE_DIR=" + '"' + sourcePath + '"');
                    else if (line.Trim().IndexOf("OUTPUT_FILE_DIR") >= 0)
                        sw.WriteLine("OUTPUT_FILE_DIR=" + '"' +  destinationPath + '"');
                    else
                        sw.WriteLine(line);
                }
                sw.Flush();
                sw.Close();

                /** Create output directory if not exist **/
                Directory.CreateDirectory(destinationPath);

                /** Launch the local dstv2dxf **/
                defFile = Path.Combine(modelDir, defFile);
                Process NCDXFConv = new Process
                {
                    EnableRaisingEvents = false
                };
                NCDXFConv.StartInfo.CreateNoWindow = true;
                NCDXFConv.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                NCDXFConv.StartInfo.FileName = Path.Combine(modelDir, exeFile);
                NCDXFConv.StartInfo.Arguments = " -cfg " + defFile + " -m batch -f *.nc1";
                NCDXFConv.Start();
                NCDXFConv.WaitForExit();
                NCDXFConv.Close();

            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }
    }
}
