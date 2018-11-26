using System;
using System.Collections.Generic;
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

            printAttributes.Scale = 1;
            printAttributes.PrintToMultipleSheet = false;
            printAttributes.NumberOfCopies = 1;
            printAttributes.Orientation = DotPrintOrientationType.Auto;
            printAttributes.PrintArea = DotPrintAreaType.EntireDrawing;
            printAttributes.ScalingType = DotPrintScalingType.Auto;
            printAttributes.PrinterInstance = "PDFCreator";
            

            DrawingEnumerator selectedDrawings = drawingHandler.GetDrawingSelector().GetSelected();
            while (selectedDrawings.MoveNext())
            {

                Drawing currentDrawing = selectedDrawings.Current;

                // Make sure it is a valid drawing
                if (currentDrawing != null)
                {
                    // Check that the drawing is up-to-date before printing
                    if (currentDrawing.UpToDateStatus == DrawingUpToDateStatus.DrawingIsUpToDate)
                    {
                        
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

            convertedFile = drawingName + " - " + drawingTitle;

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
            MessageBox.Show(path, textBox1.Text);
            string fullPath = Path.Combine(path, textBox1.Text);
            MessageBox.Show(fullPath.ToString());


        }
    }
}
