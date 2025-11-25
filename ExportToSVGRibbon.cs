using Microsoft.Office.Tools.Ribbon;
using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

namespace ChartToSVG
{
    public partial class ExportToSVGRibbon
    {
        private void ExportToSVGRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_svg_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Chart chart = null;

                if (app.ActiveChart != null)
                {
                    chart = app.ActiveChart;
                }
                else
                {
                    Forms.MessageBox.Show(text: "No active chart found. Please select a chart and try again.",
                        caption: "Error",
                        buttons: Forms.MessageBoxButtons.OK,
                        icon: Forms.MessageBoxIcon.Error);
                    return;
                }

                Forms.OpenFileDialog saveFileDialog = new Forms.OpenFileDialog
                {
                    Filter = "SVG Files (*.svg)|*.svg",
                    Title = "Save Chart as SVG",
                    CheckFileExists = false,
                    CheckPathExists = true,
                    DefaultExt = "svg"
                };

                if (saveFileDialog.ShowDialog() != Forms.DialogResult.OK)
                {
                    return;
                }

                string filePath = saveFileDialog.FileName;

                ProcessChart(chart, filePath);

                Forms.MessageBox.Show(text: "Chart exported successfully!",
                                      caption: "Success",
                                      buttons: Forms.MessageBoxButtons.OK,
                                      icon: Forms.MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                Forms.MessageBox.Show(text: ex.Message,
                                      caption: "Error",
                                      buttons: Forms.MessageBoxButtons.OK,
                                      icon: Forms.MessageBoxIcon.Error);



            }
        }

        private void ProcessChart(Excel.Chart chart, string filePath)
        {
            //// For any processing, export to a temp SVG file first and delete it after use
            //string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".svg");
            //chart.Export(tempPath, "SVG");
            //// DO things with the SVG file here
            //File.Delete(tempPath);
            // For now, just export directly to the specified file path
            chart.Export(filePath, "SVG");
        }

    }
}
