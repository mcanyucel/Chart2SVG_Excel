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
        private static bool _isSkiaRegistered = false;

        private void ExportToSVGRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Start log at initialization
            ExportLogger.StartNewLog("Add-in Initialization");
            ExportLogger.Log("ChartToSVG loading...");
            RegisterSkiaSharp();
        }

        private void RegisterSkiaSharp()
        {
            if (_isSkiaRegistered)
            {
                ExportLogger.Log("SkiaSharp already registered");
                return;
            }

            ExportLogger.Log("Registering SkiaSharp...");

            try
            {
                var assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var assemblyDirectory = Path.GetDirectoryName(assemblyPath);
                var architecture = Environment.Is64BitProcess ? "x64" : "x86";
                var nativePath = Path.Combine(assemblyDirectory, architecture);

                ExportLogger.Log($"Architecture: {architecture}");
                ExportLogger.Log($"Native DLL path: {nativePath}");

                if (Directory.Exists(nativePath))
                {
                    var path = Environment.GetEnvironmentVariable("PATH") ?? String.Empty;
                    if (!path.Contains(nativePath))
                    {
                        Environment.SetEnvironmentVariable("PATH", $"{nativePath};{path}");
                        ExportLogger.Log("Added SkiaSharp path to PATH");
                    }
                    else
                    {
                        ExportLogger.Log("PATH already includes SkiaSharp");
                    }
                }
                else
                {
                    ExportLogger.Log("WARNING: Native DLL path not found");
                }

                _isSkiaRegistered = true;
                ExportLogger.Log("✓ SkiaSharp registration complete");
            }
            catch (Exception ex)
            {
                ExportLogger.Log($"✗ Registration failed: {ex.Message}");
                // Don't disrupt Excel startup even if registration fails
            }
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
                    Forms.MessageBox.Show(
                        text: "No active chart found. Please select a chart and try again.",
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

                // Start fresh log for this export operation (overwrites previous)
                ExportLogger.StartNewLog($"Export to {extension.ToUpper()}");
                ExportLogger.Log($"Chart: {chart.Name}");
                ExportLogger.Log($"Output: {filePath}");

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
                    Forms.MessageBox.Show(
                        text: "Unsupported file format selected.",
                        caption: "Error",
                        buttons: Forms.MessageBoxButtons.OK,
                        icon: Forms.MessageBoxIcon.Error);
                    return;
                }

                ExportLogger.LogSuccess(filePath);

                Forms.MessageBox.Show(
                    text: "Chart exported successfully!",
                    caption: "Success",
                    buttons: Forms.MessageBoxButtons.OK,
                    icon: Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ExportLogger.LogError(ex);

                Forms.MessageBox.Show(
                    text: ex.Message,
                    caption: "Error",
                    buttons: Forms.MessageBoxButtons.OK,
                    icon: Forms.MessageBoxIcon.Error);
            }
        }

        private void ProcessChartToSVG(Excel.Chart chart, string filePath)
        {
            ExportLogger.Log("Starting SVG export...");
            chart.Export(filePath, "SVG");
            ExportLogger.Log("SVG export completed");
        }

        private void ProcessChartToPDF(Excel.Chart chart, string filePath)
        {
            string tempSvgPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".svg");

            try
            {
                ExportLogger.Log("Step 1: Exporting chart to temporary SVG...");
                chart.Export(tempSvgPath, "SVG");
                ExportLogger.Log($"  Temp SVG: {tempSvgPath} ({new FileInfo(tempSvgPath).Length:N0} bytes)");

                ExportLogger.Log("Step 2: Loading SVG with SkiaSharp...");
                var svg = new SKSvg();
                svg.Load(tempSvgPath);

                if (svg.Picture == null)
                {
                    throw new Exception("Failed to load SVG content.");
                }

                var bounds = svg.Picture.CullRect;
                float width = bounds.Width;
                float height = bounds.Height;
                ExportLogger.Log($"  SVG dimensions: {width:F1} x {height:F1}");

                ExportLogger.Log("Step 3: Creating PDF document...");
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
                ExportLogger.Log("PDF creation completed");
            }
            catch (Exception ex)
            {
                ExportLogger.LogError(ex);
                throw;
            }
            finally
            {
                if (File.Exists(tempSvgPath))
                {
                    File.Delete(tempSvgPath);
                    ExportLogger.Log("Cleaned up temporary SVG file");
                }
            }
        }
    }
}