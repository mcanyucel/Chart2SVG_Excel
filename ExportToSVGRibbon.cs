using Microsoft.Office.Tools.Ribbon;
using SkiaSharp;
using Svg.Skia;
using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

namespace ChartToSVG
{
    public partial class ExportToSVGRibbon
    {
        private void ExportToSVGRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_export_click(object sender, RibbonControlEventArgs e)
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

                Forms.SaveFileDialog saveFileDialog = new Forms.SaveFileDialog
                {
                    Filter = "PDF Files (*.pdf)|*.pdf|SVG Files (*.svg)|*.svg",
                    Title = "Save Chart",
                    DefaultExt = "pdf",
                    FilterIndex = 1
                };

                if (saveFileDialog.ShowDialog() != Forms.DialogResult.OK)
                {
                    return;
                }

                string filePath = saveFileDialog.FileName;
                string extension = Path.GetExtension(filePath).ToLower();

                if (extension == ".pdf")
                {
                    ProcessChartToPDF(chart, filePath);
                }
                else if (extension == ".svg")
                {
                    ProcessChartToSVG(chart, filePath);
                }
                else
                {
                    Forms.MessageBox.Show(text: "Unsupported file format selected.",
                                          caption: "Error",
                                          buttons: Forms.MessageBoxButtons.OK,
                                          icon: Forms.MessageBoxIcon.Error);
                    return;
                }

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

        private void ProcessChartToSVG(Excel.Chart chart, string filePath)
        {
            chart.Export(filePath, "SVG");
        }

        private void ProcessChartToPDF(Excel.Chart chart, string filePath)
        {
            string tempSvgPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".svg");
            try
            {
                chart.Export(tempSvgPath, "SVG");

                var svg = new SKSvg();
                svg.Load(tempSvgPath);

                if (svg.Picture == null)
                {
                    throw new Exception("Failed to load SVG content.");
                }

                var bounds = svg.Picture.CullRect;
                float width = bounds.Width;
                float height = bounds.Height;

                using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    using (var pdfDocument = SKDocument.CreatePdf(stream))
                    {
                        using (var pdfCanvas = pdfDocument.BeginPage(width, height))
                        {
                            pdfCanvas.DrawPicture(svg.Picture);
                        }
                        pdfDocument.EndPage();
                        pdfDocument.Close();
                    }
                }
            }
            finally
            {
                if (File.Exists(tempSvgPath))
                {
                    File.Delete(tempSvgPath);
                }
            }
        }
    }
}
