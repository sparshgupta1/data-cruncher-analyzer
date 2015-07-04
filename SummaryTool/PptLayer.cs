using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Graph = Microsoft.Office.Interop.Graph;
using System.Runtime.InteropServices;
using System.Data;

namespace SummaryTool
{
    public class PptLayer
    {
        public static PowerPoint.Application objApp;
        public static PowerPoint.Presentations objPresSet;
        public static PowerPoint._Presentation objPres;
        public static PowerPoint.Slides objSlides;
        public static PowerPoint._Slide objSlide;
        public static PowerPoint.TextRange objTextRng;
        public static PowerPoint.Shapes objShapes;
        public static PowerPoint.Shape objShape;
        public static PowerPoint.SlideShowWindows objSSWs;
        public static PowerPoint.SlideShowTransition objSST;
        public static PowerPoint.SlideShowSettings objSSS;
        public static PowerPoint.SlideRange objSldRng;
        public static Graph.Chart objChart;

        public static void ShowPresentation(string path)
        {
            String strTemplate ; //, strPic;
            string Base = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

            string LocalDir = System.IO.Path.GetDirectoryName(Base).Replace("file:\\", "");

            strTemplate = System.IO.Path.Combine(LocalDir, "Resources", "Template.pptx"); 
            //"\\\\flrblr001\\windows_data\\From-SLI-Fileserver\\E\\Product Engineering\\Sundar\\svn\\trunk\\CommonBase\\DeviceSpecific\\Si53300\\sample_data\\Template.pptx";
            //strPic = "C:\\Windows\\Blue Lace 16.bmp";
            //bool bAssistantOn;
            if (!System.IO.File.Exists(strTemplate)) {
                System.Windows.MessageBox.Show("Template file is missing, create the file: " + strTemplate);
            }
            //Create a new presentation based on a template.
            objApp = new PowerPoint.Application();
            objApp.Visible = MsoTriState.msoTrue;
            objPresSet = objApp.Presentations;
            objPres = objPresSet.Open(strTemplate,
              MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
            objSlides = objPres.Slides;
            objApp.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized;
            
        }
        public static void ExportSummaryTables(DataTable dtSum=null)
        {
            //Add a 15 x 15 table
            int row = 25 ; //Number of Rows per sheet to export to ppt
            
            //Find how many level of Sub headers
            int HeaderRow = 1;
            for (int i = 0; i < dtSum.Columns.Count; i++) if (HeaderRow < dtSum.Columns[i].ToString().Split('#').Length) HeaderRow = dtSum.Columns[i].ToString().Split('#').Length;
            int Datarow = row - HeaderRow;
            //Create pages
            int pagecount = dtSum.Rows.Count / Datarow;
            if ((dtSum.Rows.Count % Datarow) > 0) pagecount = (dtSum.Rows.Count / Datarow) + 1;
            for (int k = 0; k < pagecount; k++)
            {
                //Add a blank slide
                PowerPoint.Slide sld = objPres.Slides.Add(objSlides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                //objApp.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized;
                int currentPageRowCount=row;
                if (dtSum.Rows.Count < row) { row = dtSum.Rows.Count + HeaderRow; Datarow = row - HeaderRow; currentPageRowCount = row; }
                else if (dtSum.Rows.Count - Datarow * k < row) { currentPageRowCount = dtSum.Rows.Count - Datarow * k + HeaderRow; Datarow = row - HeaderRow; }

                //                                                                 horiz,verti                      cell Width, Height
                PowerPoint.Shape shp = sld.Shapes.AddTable(currentPageRowCount, dtSum.Columns.Count, 10, 60, objPres.PageSetup.SlideWidth - 20, 200);
                //shp.TextFrame.TextRange.Font.Size = 20;
                PowerPoint.Table tbl = shp.Table;
                
                //Print the Header
                for (int i = 0; i < dtSum.Columns.Count; i++)
                {
                    if (dtSum.Columns[i].ToString().IndexOf('#') == -1) //Check for multiheader
                    {
                        tbl.Rows[1].Cells[i + 1].Shape.TextFrame.TextRange.Text = dtSum.Columns[i].ToString();
                        tbl.Rows[1].Cells[i + 1].Shape.TextFrame.TextRange.Font.Size = 8;
                        //tf.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    }
                    else // split the header and move to next row
                    {
                        string[] words = dtSum.Columns[i].ToString().Split('#');
                        for (int j = 0; j < words.Length; j++)
                        {
                            if (i > 0) if (tbl.Rows[j + 1].Cells[i].Shape.TextFrame.TextRange.Text == words[j]) tbl.Rows[j + 1].Cells[i + 1].Merge(tbl.Rows[j + 1].Cells[i]);
                                else tbl.Rows[j + 1].Cells[i + 1].Shape.TextFrame.TextRange.Text = words[j];
                            else tbl.Rows[j + 1].Cells[i + 1].Shape.TextFrame.TextRange.Text = words[j];
                            tbl.Rows[j + 1].Cells[i + 1].Shape.TextFrame.TextRange.Font.Size = 8;
                            //tbl.Rows[j + 1].Cells[i + 1].Shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        }
                        if (HeaderRow < words.Length) HeaderRow = words.Length + 1;
                    }
                }
                
                //Populate Rows
                for (int i = 0; i < Datarow; i++)
                {
                    if (Datarow * k + i >= dtSum.Rows.Count) continue;
                    for (int j = 0; j < dtSum.Rows[i].ItemArray.Length; j++)
                    {
                        string Datastr;
                        var curCell = tbl.Rows[i + 1 + HeaderRow].Cells[j + 1];
                        var preCell = tbl.Rows[i + HeaderRow].Cells[j + 1];
                        if (i > 0) if (preCell.Shape.TextFrame.TextRange.Text == dtSum.Rows[i + Datarow * k].ItemArray[j].ToString()) { preCell.Merge(curCell); continue; }
                            else Datastr = dtSum.Rows[i + Datarow * k].ItemArray[j].ToString();
                        else Datastr = dtSum.Rows[i + Datarow * k].ItemArray[j].ToString();
                        curCell.Shape.TextFrame.TextRange.Font.Size = 8;
                        
                        //for more format refer http://www.pptfaq.com/FAQ00790_Working_with_PowerPoint_tables.htm
                        double Datadbl;
                        if (Double.TryParse(Datastr, out Datadbl))
                            if (Datastr.Length > 8) curCell.Shape.TextFrame.TextRange.Text = Format("S4", Datadbl);
                            else curCell.Shape.TextFrame.TextRange.Text = Datadbl.ToString();
                        else
                            curCell.Shape.TextFrame.TextRange.Text = Datastr;
                    }
                }
            }
            
        }
        public static void IncludeOthers(string path)
        {//Build Slide #2:
            //Add text to the slide title, format the text. Also add a chart to the
            //slide and change the chart type to a 3D pie chart.
            objSlide = objSlides.Add(objSlides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Text = "My Chart";
            objTextRng.Font.Name = "Comic Sans MS";
            objTextRng.Font.Size = 48;
            objChart = (Graph.Chart)objSlide.Shapes.AddOLEObject(150, 150, 480, 320,
              "MSGraph.Chart.8", "", MsoTriState.msoFalse, "", 0, "",
              MsoTriState.msoFalse).OLEFormat.Object;
            objChart.ChartType = Graph.XlChartType.xl3DPie;
            objChart.Legend.Position = Graph.XlLegendPosition.xlLegendPositionBottom;
            objChart.HasTitle = true;
            objChart.ChartTitle.Text = "Here it is...";

            //Build Slide #3:
            //Change the background color of this slide only. Add a text effect to the slide
            //and apply various color schemes and shadows to the text effect.
            objSlide = objSlides.Add(objSlides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            objSlide.FollowMasterBackground = MsoTriState.msoFalse;
            objShapes = objSlide.Shapes;
            objShape = objShapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect27,
              "The End", "Impact", 96, MsoTriState.msoFalse, MsoTriState.msoFalse, 230, 200);

            //Save the presentation to disk
            objPres.SaveAs(path,
                  PowerPoint.PpSaveAsFileType.ppSaveAsPresentation,
                  Microsoft.Office.Core.MsoTriState.msoFalse);
            //Modify the slide show transition settings for all 3 slides in
            //the presentation.
            int[] SlideIdx = new int[3];
            for (int k = 0; k < 3; k++) SlideIdx[k] = k + 1;
            objSldRng = objSlides.Range(SlideIdx);
            objSST = objSldRng.SlideShowTransition;
            objSST.AdvanceOnTime = MsoTriState.msoTrue;
            objSST.AdvanceTime = 3;
            objSST.EntryEffect = PowerPoint.PpEntryEffect.ppEffectBoxOut;

            //Prevent Office Assistant from displaying alert messages:
            // bAssistantOn = objApp.Assistant.On;
            // objApp.Assistant.On = false;

            //Run the Slide show from slides 1 thru 3.
            objSSS = objPres.SlideShowSettings;
            objSSS.StartingSlide = 1;
            objSSS.EndingSlide = 3;
            objSSS.Run();

            //Wait for the slide show to end.
            objSSWs = objApp.SlideShowWindows;
            while (objSSWs.Count >= 1) System.Threading.Thread.Sleep(100);

            ////Reenable Office Assisant, if it was on:
            //if (bAssistantOn)
            //{
            //    objApp.Assistant.On = true;
            //    objApp.Assistant.Visible = false;
            //}

            //Close the presentation without saving changes and quit PowerPoint.
            //  objPres.Close();
            //  objApp.Quit();
        }
        public static void SavePresentation(string path)
        {
            //Save the presentation to disk
            objPres.SaveAs(path,
                  PowerPoint.PpSaveAsFileType.ppSaveAsPresentation,
                  Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        public static void IncludeImage(string path)
        {
            //Build Slide #1:
            //Add text to the slide, change the font and insert/position a 
            //picture on the first slide.
            objSlide = objSlides.Add(objSlides.Count+1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Text = "Summary Report";
            objTextRng.Font.Name = "Comic Sans MS";
            objTextRng.Font.Size = 48;
            objSlide.Shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue,
              0, 50, 700, 450);
        }
        /// <summary>
        /// Format a value using engineering notation.
        /// </summary>
        /// <example>
        ///     Format("S4",-12345678.9) = "-12.34e-6"
        ///     with 4 significant digits
        /// </example>
        /// <arg name="format">The format specifier</arg>
        /// <arg name="value">The value</arg>
        /// <returns>A string representing the value formatted according to the format specifier</returns>
        public static string Format(string format, double value)
        {
            if (format.StartsWith("S"))
            {
                string dg = format.Substring(1);
                int significant_digits;
                int.TryParse(dg, out significant_digits);
                if (significant_digits == 0) significant_digits = 4;
                int sign;
                double amt;
                int exponent;
                SplitEngineeringParts(value, out sign, out amt, out exponent);
                return ComposeEngrFormat(significant_digits, sign, amt, exponent);
            }
            else
            {
                return value.ToString(format);
            }
        }
        static void SplitEngineeringParts(double value,
                    out int sign,
                    out double new_value,
                    out int exponent)
        {
            sign = Math.Sign(value);
            value = Math.Abs(value);
            if (value > 0.0)
            {
                if (value > 1.0)
                {
                    exponent = (int)(Math.Floor(Math.Log10(value) / 3.0) * 3.0);
                }
                else
                {
                    exponent = (int)(Math.Ceiling(Math.Log10(value) / 3.0) * 3.0);
                }
            }
            else
            {
                exponent = 0;
            }
            new_value = value * Math.Pow(10.0, -exponent);
            if (new_value >= 1e3)
            {
                new_value /= 1e3;
                exponent += 3;
            }
            if (new_value <= 1e-3 && new_value > 0)
            {
                new_value *= 1e3;
                exponent -= 3;
            }
        }
        static string ComposeEngrFormat(int significant_digits, int sign, double v, int exponent)
        {
            int expsign = Math.Sign(exponent);
            try { exponent = Math.Abs(exponent); }
            catch { }
            int digits = v > 0 ? (int)Math.Log10(v) + 1 : 0;
            int decimals = Math.Max(significant_digits - digits, 0);
            double round = Math.Pow(10, -decimals);
            digits = v > 0 ? (int)Math.Log10(v + 0.5 * round) + 1 : 0;
            decimals = Math.Max(significant_digits - digits, 0);
            string t;
            string f = "0:F";
            if (exponent == 0)
            {
                t = string.Format("{" + f + decimals + "}", sign * v);
            }
            else
            {
                t = string.Format("{" + f + decimals + "}e{1}", sign * v, expsign * exponent);
            }
            return t;
        }

        public static void ValidateFormat()
        {
            Console.WriteLine("\t1.0      = {0}", Format("S3", 1.0));
            Console.WriteLine("\t0.1      = {0}", Format("S3", 0.1));
            Console.WriteLine("\t0.01     = {0}", Format("S3", 0.01));
            Console.WriteLine("\t0.001    = {0}", Format("S3", 0.001));
            Console.WriteLine("\t0.0001   = {0}", Format("S3", 0.0001));
            Console.WriteLine("\t0.00001  = {0}", Format("S3", 0.00001));
            Console.WriteLine("\t0.000001 = {0}", Format("S3", 0.000001));
        }

    }

}
