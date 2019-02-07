using ClosedXML.Excel;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Dialog.UIControls;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using WindowsFormsApp3.Properties;
using pdfParser = iTextSharp.text;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        //[DllImport("user32.dll", EntryPoint = "FindWindowA")]
        //public static extern IntPtr FindWindowA(string lp1, string lp2);

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        public static extern IntPtr FindWindow(string lp1, string lp2);


        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);


        //public static extern IntPtr FindWindow(string lp1, string lp2);


        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(HandleRef hWnd, int nCmdShow);
        public const int SW_RESTORE = 9;
        Model teklaModel;

        public Form1()
        {
            InitializeComponent();
            teklaModel = new Model();
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

        private void cmdProjectProperties_Click(object sender, EventArgs e)
        {
            Model currentModel = new Model();
            ProjectInfo projectInfo = currentModel.GetProjectInfo();

            if (currentModel.GetConnectionStatus())
            {
                if (projectInfo.Name.Length > 0)
                {
                    string contractor = string.Empty;
                    projectInfo.GetUserProperty("GC", ref contractor);
                    projectInfo.GetUserProperty("GC", ref contractor);
                    MessageBox.Show(projectInfo.ProjectNumber + Environment.NewLine +  projectInfo.Name 
                        + Environment.NewLine  + " " + contractor);
                }

                

            }


        }

        private void cmdMiscList_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Export Drawing List";
                saveFileDialog.ShowDialog();

                if (saveFileDialog.FileName != "")
                {
                    var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("Misc_List");

                    FormatSheet(worksheet);
                    RowSheetData(worksheet);

                    //Save to file
                    workbook.SaveAs(saveFileDialog.FileName);

                    DialogResult dialogResult = MessageBox.Show("Done exporting to excel file. You want to open the folder?", "Show folder", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                        Process.Start(@Path.GetDirectoryName(saveFileDialog.FileName));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FormatSheet(IXLWorksheet worksheet)
        {
            Model currentModel = new Model();
            ProjectInfo projectInfo = currentModel.GetProjectInfo();
            string contractor = string.Empty;
            projectInfo.GetUserProperty("GC", ref contractor);

            //Header title
            worksheet.Cell("A1").Value = "JOB NO.";
            worksheet.Cell("A3").Value = "SHEET NO.";
            worksheet.Cell("E1").Value = "IN SHOP";
            worksheet.Cell("E3").Value = "BY";
            worksheet.Range("A1:E5").Style.Font.FontName = "Arial";
            worksheet.Range("A1:E1").Style.Font.FontSize = 8;
            worksheet.Range("A3:E3").Style.Font.FontSize = 8;
            worksheet.Range("A1:E4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range("A1:E4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Range("A1:A2").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Range("E1:E2").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            worksheet.Cell("A5").Value = "PROJECT:";
            worksheet.Cell("C5").Value = "CONTRACTOR:";
            worksheet.Range("A5:E5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            worksheet.Range("A5:E5").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Range("A5:E5").Style.Border.InsideBorder= XLBorderStyleValues.Thin;
            worksheet.Range("A5:E5").Style.Font.FontSize = 8;

            //Header Values
            worksheet.Cell("A2").Value = projectInfo.ProjectNumber;
            worksheet.Cell("A4").Value = "1 of 1";
            worksheet.Cell("E2").Value = "";
            worksheet.Cell("E4").Value = "";
            worksheet.Range("A2:E2").Style.Font.FontSize = 10;
            worksheet.Range("A4:E4").Style.Font.FontSize = 10;
            worksheet.Range("A2:E2").Style.Font.Bold = true;
            worksheet.Range("A4:E4").Style.Font.Bold = true;
            worksheet.Range("A3:A4").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Range("E3:E4").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            worksheet.Cell("B5").Value = projectInfo.Name;
            worksheet.Cell("B5").Style.Font.Bold = true;
            worksheet.Cell("D5").Value = contractor;
            worksheet.Cell("D5").Style.Font.Bold = true;
            worksheet.Range("D5:E5").Merge();

            //Logo
            //var logo = new Bitmap(WindowsFormsApp3.Properties.Resources.);
            worksheet.Range("B1:D4").Merge();
            worksheet.Range("B1:D4").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            var logo = Resources.trueNorthLogo;
            worksheet.AddPicture(logo).MoveTo(worksheet.Cell("B1").Address, 7, 7).Scale(0.95);

            //Report Title
            worksheet.Range("A6:E6").Merge();
            worksheet.Cell("A6").Style.Font.FontName = "Arial";
            worksheet.Cell("A6").Style.Font.FontSize = 11;
            worksheet.Range("A6:E6").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Cell("A6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range("A6:E6").Style.Font.Bold = true;
            worksheet.Cell("A6").Value = "BOLT & MISC: ITEM LIST";

            //Column Header
            worksheet.Range("B7:C7").Merge();
            worksheet.Range("D7:E7").Merge();

            worksheet.Cell("A7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range("B7:C7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Cell("D7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range("A7:E7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Range("A7:E7").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Range("A7:E7").Style.Font.FontName = "Arial";
            worksheet.Range("A7:E7").Style.Font.FontSize = 8;

            worksheet.Cell("A7").Value = "QUANTITY";
            worksheet.Cell("B7").Value = "DESCRIPTION";
            worksheet.Cell("D7").Value = "LOCATION";
            
            //Columns width
            worksheet.Column(1).Width = 10.86;
            worksheet.Column(2).Width = 35.71;
            worksheet.Column(3).Width = 11.43;
            worksheet.Column(4).Width = 16.43;
            worksheet.Column(5).Width = 10.86;

            //Row data border
            for (int i = 0; i < 31; i++)
            {
                int rowData = 8+i; //Start the data in row 8
                worksheet.Range("B" + rowData + ":C" + rowData).Merge();
                worksheet.Range("D" + rowData + ":E" + rowData).Merge();
                worksheet.Range("A" + rowData + ":E" + rowData).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("A" + rowData + ":E" + rowData).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }
        }

        private void RowSheetData(IXLWorksheet worksheet)
        {
            int currentRow = 8;
         
            var objModel = teklaModel.GetModelObjectSelector().GetAllObjects();
            objModel.SelectInstances = false;

            List<Tekla.Structures.Model.ModelObject> modelObjects = ToList(objModel);
            List<DrawingStatusModel> listModelObject = GetBoughtOutItem(modelObjects);
           
            var query = from p in listModelObject.GroupBy(p => p.CatalogNumber)
                        select new
                        {
                            count = p.Count(),
                            p.First().BoughtOutComment,
                            p.First().BoughtOutItem,
                            p.First().CatalogNumber,
                            p.First().Id
                        };


            foreach (var obj in query)
            {
                var rangeRow = worksheet.Range(currentRow, 1, currentRow, 5);
                rangeRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangeRow.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                rangeRow.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rangeRow.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rangeRow.Style.Font.FontSize = 8;
                worksheet.Cell(currentRow, 1).Value = obj.count;
                worksheet.Cell(currentRow, 2).Value = obj.CatalogNumber + " " + obj.BoughtOutComment;
                worksheet.Cell(currentRow, 2).Style.Alignment.WrapText = true;
                currentRow++;
            }
        }

        private static List<Tekla.Structures.Model.ModelObject> ToList(ModelObjectEnumerator selectedModels)
        {
            var modelObjects = new List<Tekla.Structures.Model.ModelObject>();
            while (selectedModels.MoveNext())
            {
                var modelObject = selectedModels.Current;
                if (modelObject == null) continue;
                modelObjects.Add(modelObject);
            }

            return modelObjects;
        }

        private List<DrawingStatusModel> GetBoughtOutItem(List<Tekla.Structures.Model.ModelObject> modelObjects)
        {
            List<DrawingStatusModel> listModelObject = null;

            if (teklaModel.GetConnectionStatus())
            {
                var data = modelObjects
                .Select(b =>
                {
                    ArrayList stringUDAList = new ArrayList { "BOUGHT_ITEM", "comment", "BOUGHT_ITEM_NAME" };
                    Hashtable stringUDA = new Hashtable();
                    b.GetStringReportProperties(stringUDAList, ref stringUDA);

                    
                    return new DrawingStatusModel
                    {
                        Id = b.Identifier.ID,
                        Type = b.GetType().Name.ToString(),
                        BoughtOutItem = stringUDA.ContainsKey("BOUGHT_ITEM") ? (string)stringUDA["BOUGHT_ITEM"] : string.Empty,
                        BoughtOutComment = stringUDA.ContainsKey("comment") ? (string)stringUDA["comment"] : string.Empty,
                        CatalogNumber = stringUDA.ContainsKey("BOUGHT_ITEM_NAME") ? (string)stringUDA["BOUGHT_ITEM_NAME"] : string.Empty,
                    };
                })
                .Where(b => b.BoughtOutItem.ToUpper() == "YES") 
                .Where(b => b.Type != "Assembly")
                .OrderBy(b => b.SeqNo);

                listModelObject = data.ToList<DrawingStatusModel>();
            }

            return listModelObject;
        }

        private void cmdPDFSize_Click(object sender, EventArgs e)
        {
            string[] filePaths = Directory.GetFiles(textBox1.Text, "*.pdf");
            foreach (string file in filePaths)
            {
                
                PdfReader pdfReader = new PdfReader(file);

                pdfParser.Rectangle mediabox = pdfReader.GetPageSize(1);
                MessageBox.Show(file);
            }
        }

        private void cmdMoveFile_Click(object sender, EventArgs e)
        {
            File.Move(textBox1.Text, textBox2.Text);
        }

        private void cmdRemovePlates_Click(object sender, EventArgs e)
        {
            if (ExportReport(textBox1.Text))
            {
                MessageBox.Show("Done");
            }
            else
            {
                MessageBox.Show("Error");
            }
        }

        private bool ExportReport(string path, string reportTemplate = "")
        {
            DrawingHandler DH = new DrawingHandler();
            DrawingEnumerator DE = DH.GetDrawingSelector().GetSelected();
            
            if (DE.GetSize() != 0)
            {
                while (DE.MoveNext())
                {
                    Drawing D = DE.Current;

                    if (!(D is SinglePartDrawing))
                    {
                        DH.SetActiveDrawing(D, false);
                    }
                }
                return Operation.CreateReportFromSelected(reportTemplate, path, "", "", "");
            }
            return false;
        }

        private void cmdTestStack_Click(object sender, EventArgs e)
        {
            DrawingHandler DH = new DrawingHandler();
            DrawingEnumerator DE = DH.GetDrawingSelector().GetSelected();
            int RevNO = 0;
            string revisionMark = string.Empty;
            string revisionLastMark = string.Empty;


            if (DE.GetSize() != 0)
            {
                while (DE.MoveNext())
                {
                    Drawing D = DE.Current;

                    Beam fakeBeam = new Beam
                    {
                        Identifier =
                            (Identifier)D.GetType().GetProperty("Identifier", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(D, null)
                    };

                    fakeBeam.GetReportProperty("REVISION.NUMBER", ref RevNO);
                    fakeBeam.GetReportProperty("REVISION.MARK", ref revisionMark);
                    fakeBeam.GetReportProperty("REVISION.LAST_MARK", ref revisionLastMark);
                    //if (!(D is SinglePartDrawing))
                    //{
                    //    DH.SetActiveDrawing(D, false);
                    //}

                    if (RevNO > 0)
                    {
                        DH.SetActiveDrawing(D, true);
                    }

                }
                Operation.CreateReportFromSelected("TN_FabTrol_Drawing_List-v30", "c:\\temp\\test\\TN_FabTrol_Drawing_List-v30.xsr", "", "", "");

                MessageBox.Show("Done");
            }
            else
            {
                MessageBox.Show("Nothing Selected");
            }
        }

        private void cmdSendMessage_Click(object sender, EventArgs e)
        {
            IntPtr handle = FindWindow(String.Empty, "Drawing List");

            if (SetForegroundWindow(handle))
            {
                SendKeys.Send("{H}{e}{l 2}{o} {W}{o}{r}{l}{d}{!}");
                SendKeys.Send("{TAB}");
                SendKeys.Send("{TAB}");
                SendKeys.Send("{TAB}");
                SendKeys.Send("{TAB}");
                SendKeys.Send("{TAB}");

                //    //SendKeys.Send("{ENTER}");
            }
        }

        private void cmdShowWindowAsync_Click(object sender, EventArgs e)
        {
            Process[] objProcesses = System.Diagnostics.Process.GetProcessesByName("Tekla Structures");
            if (objProcesses.Length > 0)
            {
                IntPtr hWnd = IntPtr.Zero;
                hWnd = objProcesses[0].MainWindowHandle;
                ShowWindowAsync(new HandleRef(null, hWnd), SW_RESTORE);
                SetForegroundWindow(objProcesses[0].MainWindowHandle);
            }
        }

        private void cmdFilter_Click(object sender, EventArgs e)
        {
            var selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            ArrayList modelObjects = new ArrayList();
            //GetModelsInDrawingsByUserIssue(textBox1.Text);
            SelectModelBasedonUserIssue(textBox1.Text, modelObjects);

            
            selector.Select(modelObjects);
            Operation.CreateMISFileFromSelected(Operation.MISExportTypeEnum.KISS, "c:\\temp\\test.kss");
        }

        private List<int> GetModelsInDrawingsByUserIssue(string UserIssue)
        {
            DrawingHandler DH = new DrawingHandler();
            DrawingEnumerator drawingsEnum = DH.GetDrawings();
            string userIssueProperty = string.Empty;
            List<int> selectedDrawings = new List<int>();
            var modelObjects = new ArrayList();


            if (teklaModel.GetConnectionStatus())
            {
                while (drawingsEnum.MoveNext())
                {
                    var drawing = drawingsEnum.Current as Drawing;
                    drawing.GetUserProperty("DR_USER_ISSUE", ref userIssueProperty);
                    if (userIssueProperty == UserIssue)
                    {
                        Type[] typeFilter = new Type[] { typeof(Tekla.Structures.Drawing.Part) };
                        foreach (DrawingObject drawingObject in drawing.GetSheet().GetObjects())
                        {

                            //----------------------------------------------------//
                            Tekla.Structures.Drawing.ModelObject oDrawingPart = drawingObject as Tekla.Structures.Drawing.ModelObject;
                            Beam fakeBeam = new Beam();
                            fakeBeam.Identifier = (drawingObject as Tekla.Structures.Drawing.Part).ModelIdentifier;
                            Console.WriteLine(fakeBeam.Identifier.ID.ToString());
                            //if (!selectedDrawings.Contains(oDrawingPart.ModelIdentifier.ID))
                            //{
                            //    selectedDrawings.Add(oDrawingPart.ModelIdentifier.ID);

                            //    Tekla.Structures.Model.ModelObject oModelObject = teklaModel.SelectModelObject(new Identifier(oDrawingPart.ModelIdentifier.ID));
                            //    Beam fakeBeam = oModelObject as Beam;

                            //    if (oModelObject != null && fakeBeam != null)
                            //    {
                            //        Console.WriteLine(fakeBeam.GetAssembly().GetAssemblyType().ToString());
                            //        if (fakeBeam.GetAssembly().GetAssemblyType() == Assembly.AssemblyTypeEnum.STEEL_ASSEMBLY &&
                            //            fakeBeam.AssemblyNumber.Prefix != "REF")
                            //        {
                            //            modelObjects.Add(oModelObject);
                            //            Console.WriteLine(oModelObject.Identifier.ID);
                            //        }
                            //    }
                            //}

                                //----------------------------------------------------//

                            //    Type[] typeFilter = new Type[] { typeof(Tekla.Structures.Drawing.Part) };
                            //DrawingObjectEnumerator drawingEnum = drawingObject.GetRelatedObjects(typeFilter);

                            //while (drawingEnum.MoveNext())
                            //{
                            //    //Beam fakeBeam = new Beam();
                            //    //fakeBeam.Identifier = (drawingEnum.Current as Tekla.Structures.Drawing.Part).ModelIdentifier;
                            //    //fakeBeam.Select()


                            //    //BoltGroup modelBoltGroup = new Model().SelectModelObject(drawingObject.Modify)
                            //    Tekla.Structures.Drawing.ModelObject myDrawingPart = drawingEnum.Current as Tekla.Structures.Drawing.ModelObject;

                            //    if (!selectedDrawings.Contains(myDrawingPart.ModelIdentifier.ID))
                            //    {
                            //        selectedDrawings.Add(myDrawingPart.ModelIdentifier.ID);
                            //        Tekla.Structures.Model.ModelObject modelObject = teklaModel.SelectModelObject(new Identifier(myDrawingPart.ModelIdentifier.ID));


                            //        Beam fakeBeam = modelObject as Beam;
                                    
                            //        if (modelObject != null && fakeBeam != null)
                            //        {
                            //            Console.WriteLine(fakeBeam.GetAssembly().GetAssemblyType().ToString());
                            //            if (fakeBeam.GetAssembly().GetAssemblyType() == Assembly.AssemblyTypeEnum.STEEL_ASSEMBLY &&
                            //                fakeBeam.AssemblyNumber.Prefix != "REF")
                            //            {
                            //                modelObjects.Add(modelObject);
                            //                Console.WriteLine(modelObject.Identifier.ID);
                            //            }
                                            
                            //        }
                                        

                            //    }
                            //}
                        }
                    }
                }

                var selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                selector.Select(modelObjects);
            }

            Operation.CreateMISFileFromSelected(Operation.MISExportTypeEnum.KISS, "c:\\temp\\test.kss");
            return selectedDrawings;
        }

        private void SelectModelBasedonUserIssue(string UserIssue, ArrayList modelObjects)
        {
            DrawingHandler DH = new DrawingHandler();
            DrawingEnumerator drawingListEnumerator = DH.GetDrawings(); // Get drawing list
            string userIssueProperty = string.Empty;

            while (drawingListEnumerator.MoveNext())
            {
                Drawing myDrawing = drawingListEnumerator.Current;

                myDrawing.GetUserProperty("DR_USER_ISSUE", ref userIssueProperty);
                if (userIssueProperty == UserIssue && myDrawing.GetType().Name.ToUpper() != "GADRAWING")
                {
                    Console.WriteLine("myDrawing.Mark: " + myDrawing.Mark);
                    AddChildDrawingObjectsToTreeNode(myDrawing.GetSheet(), modelObjects); // Add all objects placed to the sheet to the UI tree
                }
                    

            }
        }
        private void AddChildDrawingObjectsToTreeNode(Tekla.Structures.Drawing.IHasChildren parentObject, ArrayList modelObjects)
        {
            Type[] typeFilter = new Type[] { typeof(Tekla.Structures.Drawing.Part), typeof(Tekla.Structures.Drawing.View) };
            DrawingObjectEnumerator objectEnumerator = parentObject.GetObjects(typeFilter); // Gets the objects that are placed directly to the current container object
            //objectEnumerator.SelectInstances = false; // We don't want to automatically select the instances, we just need to know if the object exists (we're not using objects properties at all)
            
            

            while (objectEnumerator.MoveNext())
            {
                //Tekla.Structures.Drawing.ModelObject myDrawingPart = objectEnumerator.Current as Tekla.Structures.Drawing.ModelObject;
                //Tekla.Structures.Model.ModelObject modelObject = teklaModel.SelectModelObject(new Identifier(myDrawingPart.ModelIdentifier.ID));


                //--------------------------------------------//
                string type = objectEnumerator.Current.GetType().Name;

                Console.WriteLine("Object Type: " + type);
                //--------------------------------------------//

                if (type == "Part")
                {
                    Tekla.Structures.Model.ModelObject modelObject = teklaModel.SelectModelObject((objectEnumerator.Current as Tekla.Structures.Drawing.Part).ModelIdentifier);
                    Beam fakeBeam = modelObject as Beam;


                    if (fakeBeam != null && fakeBeam.GetAssembly().GetAssemblyType() == Assembly.AssemblyTypeEnum.STEEL_ASSEMBLY &&
                        fakeBeam.AssemblyNumber.Prefix != "REF")
                    {
                        modelObjects.Add(modelObject);
                        Console.WriteLine(modelObject.Identifier.ID);
                    }
                }

                //Beam fakeBeam = modelObject as Beam;
                //if (modelObject != null && fakeBeam != null)
                //{
                //    Console.WriteLine(fakeBeam.GetAssembly().GetAssemblyType().ToString());
                //    if (fakeBeam.GetAssembly().GetAssemblyType() == Assembly.AssemblyTypeEnum.STEEL_ASSEMBLY &&
                //        fakeBeam.AssemblyNumber.Prefix != "REF")
                //    {
                //        modelObjects.Add(modelObject);
                //        Console.WriteLine(modelObject.Identifier.ID);
                //    }

                //}

                if (objectEnumerator.Current is Tekla.Structures.Drawing.IHasChildren)
                        AddChildDrawingObjectsToTreeNode(objectEnumerator.Current as IHasChildren, modelObjects);
                
            }

            
        }

        private void cmdSegregateNC1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("Textbox 1 is empty!", "NC1", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox1.Focus();
                return;
            }

            Dictionary<string, string> nc1Path = GetNC1Settings();
            string path = textBox1.Text;
            string[] filePaths = Directory.GetFiles(path, "*.nc1");
            foreach (string file in filePaths)
            {
                string profile = GetNC1Profile(file);
                Console.WriteLine("File: " + file + " profile: " + profile);
                MoveNC1ToDesignatedFolder(file, profile, nc1Path);
            }

            string profileDirectory = string.Empty;


            //Compress the Folder and delete source when done.
            foreach (var item in nc1Path)
            {
                if (Directory.Exists(Path.Combine(path, item.Value)))
                {
                    CompressNC1Folder(path, item.Value, "F288");
                }
            }

            
        }

        private void CompressNC1Folder(string FilePath, string Profile, string ProjectNumber)
        {
            string profilePath = Path.Combine(FilePath, Profile);
            string projectNamePath = Path.Combine(FilePath, ProjectNumber);
            string compressedFileName = Path.Combine(FilePath, projectNamePath + ".zip");

            //Rename to projectName first
            Directory.Move(profilePath, projectNamePath);
            //Compress the folder
            
            ZipFile.CreateFromDirectory(projectNamePath, compressedFileName, CompressionLevel.Fastest, true);
            //Rename the zip file
            string newFileName = Path.GetFileNameWithoutExtension(compressedFileName) + " " + Profile + ".zip";
            newFileName = Path.Combine(FilePath, newFileName);
            File.Move(compressedFileName, newFileName);

            //Delete the source folder
            Console.WriteLine("Deleting: " + projectNamePath);
            Directory.Delete(projectNamePath, true);
        }

        private Dictionary<string, string> GetNC1Settings()
        {
            Dictionary<string, string> nc1Path = new Dictionary<string, string>();
            nc1Path.Add(ProfileTypes.Angle, GetOpexSettings("Lejeune_AnglesFolder"));
            nc1Path.Add(ProfileTypes.Channels, GetOpexSettings("Lejeune_ChannelsFolder"));
            nc1Path.Add(ProfileTypes.HSS, GetOpexSettings("Lejeune_HSSFolder"));
            nc1Path.Add(ProfileTypes.Plate, GetOpexSettings("Lejeune_PlateFolder"));
            nc1Path.Add(ProfileTypes.WideFlange, GetOpexSettings("Lejeune_WideFlangeFolder"));
            nc1Path.Add(ProfileTypes.MISC, GetOpexSettings("Lejeune_MMarkFolder"));

            return nc1Path;
        }

        private void MoveNC1ToDesignatedFolder(string NC1File, string Profile, Dictionary<string, string> Nc1Path)
        {
            
            string fileName = Path.GetFileName(NC1File);
            string currentPath = Path.GetDirectoryName(NC1File);
            string profileDirectory = string.Empty;
            Nc1Path.TryGetValue(Profile, out profileDirectory); //Get the profile directory set in opex.config file
            
            if (profileDirectory != null)
            {
                string newPath = Path.Combine(currentPath, profileDirectory); //Append the profile directory from the root NC1 directory
                Directory.CreateDirectory(newPath); //Create the profile directory if not exist
                Console.WriteLine("Moving " + NC1File + " to " + newPath);

                if (Profile == ProfileTypes.WideFlange)
                {

                    string mMarkFolder = string.Empty;
                    Nc1Path.TryGetValue(ProfileTypes.MISC, out mMarkFolder);
                    GroupBy100AndMove(NC1File, newPath, mMarkFolder);
                }
                else
                {
                    File.Move(NC1File, Path.Combine(newPath, fileName));
                }

                if (Profile == ProfileTypes.Plate)
                {
                    ExecutePlateSorter(Path.Combine(newPath, fileName));
                }
            }
            //Do nothing if unknown profile ???
        }

        private void ExecutePlateSorter(string FilePath)
        {
            string plateSorter = GetOpexSettings("PlateSorterPath");
            string perlPath = GetOpexSettings("PerlPath");

            //Move to TEKLA folder first
            string path = Path.GetDirectoryName(FilePath);
            string fileName = Path.GetFileName(FilePath);
            string newPath = Path.Combine(path, "TEKLA");
            Directory.CreateDirectory(newPath); //Create if not exist
            newPath = Path.Combine(newPath, fileName); //Append the filename to get full path
            File.Move(FilePath, newPath);

            ProcessStartInfo perlStartInfo = new ProcessStartInfo(perlPath);
            perlStartInfo.Arguments = plateSorter + " "+ newPath;
            perlStartInfo.UseShellExecute = false;
            perlStartInfo.RedirectStandardOutput = true;
            perlStartInfo.RedirectStandardError = true;
            perlStartInfo.RedirectStandardInput = true;
            perlStartInfo.CreateNoWindow = false;

            Process perl = new Process();
            perl.StartInfo = perlStartInfo;
            perl.Start();
            
            //string output = perl.StandardOutput.ReadToEnd();
            //Console.WriteLine("Perl Output: " + output);
            //string err = perl.StandardError.ReadToEnd();
            //Console.WriteLine("Perl Output: " + err);

            StreamWriter myStreamWriter = perl.StandardInput;
            myStreamWriter.WriteLine(" ");
            perl.WaitForExit();
            
        }

        private void GroupBy100AndMove(string FileName, string FilePath, string MMarkFolder)
        {
            string output = new string(Path.GetFileName(FileName).TakeWhile(char.IsDigit).ToArray());
            int numericValue = 0;
            string newPath = string.Empty;
            string newFullPath = string.Empty;

            if (Int32.TryParse(output, out numericValue))
            {
                //Group by 100
                double lowerBoundary = Math.Floor(numericValue / 100.0) * 100.0;
                double higherBoundary = lowerBoundary + 99;
                string groupFolder = lowerBoundary.ToString() + "-" + higherBoundary;
                newPath = Path.Combine(FilePath, groupFolder);
                newFullPath = Path.Combine(newPath, Path.GetFileName(FileName));
                Console.WriteLine("Moving " + FileName + " to " + newFullPath);
            }
            else
            {
                //Move to M Mark
                newPath = Path.Combine(FilePath, MMarkFolder);
                newFullPath = Path.Combine(newPath, Path.GetFileName(FileName));
                Console.WriteLine(FileName + " will be moved to MMARK");
            }

            Directory.CreateDirectory(newPath); //Create if not exist
            File.Move(FileName, newFullPath);
        }

        public string GetNC1Profile(string FileName)
        {
            const int count = 8; // Profile is in line 9, so need to skip the first 8 lines
            string profile = File.ReadLines(FileName).Skip(count).Take(1).First();
            string returnValue = string.Empty;

            if (profile.Trim().StartsWith("HSS"))
            {
                returnValue = ProfileTypes.HSS;
            }
            else if(profile.Trim().StartsWith("C"))
            {
                returnValue = ProfileTypes.Channels;
            }
            else if (profile.Trim().StartsWith("PL") || profile.Trim().StartsWith("BPL"))
            {
                returnValue = ProfileTypes.Plate;
            }
            else if (profile.Trim().StartsWith("W"))
            {
                returnValue = ProfileTypes.WideFlange;
            }
            else if (profile.Trim().StartsWith("L"))
            {
                returnValue = ProfileTypes.Angle;
            }
            //else
            //{
            //    returnValue = ProfileTypes.MISC;
            //}
         
            return returnValue;
        }

        public string GetOpexSettings(string Variable)
        {
            var data = File
                .ReadAllLines("opex.config")
                .Select(x => x.Split('='))
                .Where(x => x.Length > 1)
                .ToDictionary(x => x[0].Trim(), x => x[1]);

            return data[Variable];
        }

    }
}
