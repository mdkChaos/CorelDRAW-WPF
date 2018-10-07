using CorelDRAW_WPF.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorelDRAW_WPF
{
    class Controller
    {
        public Controller(MainWindow mainWindow)
        {
            this.mainWindow = mainWindow;
        }
        public string FileName { get; set; }
        List<DataModel> datas;
        MainWindow mainWindow;

        void ReadFromRow(Excel.Worksheet workSheet)
        {
            int row = 2;
            CultureInfo temp_culture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
            while (workSheet.Cells[row, 1].Text.Trim() != string.Empty)
            {
                DataModel dataModel = new DataModel
                {
                    Customer = workSheet.Cells[row, 1].Text,
                    ChildName = workSheet.Cells[row, 2].Text,
                    DoorWidth = Convert.ToDouble(workSheet.Cells[row, 4].Text),
                    Pocket = workSheet.Cells[row, 5].Text
                };

                if (int.TryParse(workSheet.Cells[row, 3].Text, out int imgNum))
                {
                    dataModel.ImageNumber = imgNum.ToString();
                }
                else
                {
                    dataModel.ImageNumber = "0";
                }
                if (Convert.ToInt32(workSheet.Cells[row, 6].Text) == 1)
                {
                    dataModel.IsUp = true;
                }
                else
                {
                    dataModel.IsUp = false;
                }

                datas.Add(dataModel);
                row++;
            }
            Thread.CurrentThread.CurrentCulture = temp_culture;
        }

        void OpenFile(string filter)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Filter = filter
            };
            if (openFile.ShowDialog() == true)
            {
                FileName = openFile.FileName;
            }
        }

        async void ExtractDataFromExcel(CancellationTokenSource cts)
        {
            Excel.Application excelApp = null;
            Excel.Workbooks workBooks = null;
            Excel.Workbook workBook = null;
            Excel.Sheets workSheets = null;
            Excel.Worksheet workSheet = null;
            datas = new List<DataModel>();
            Stopwatch sw = new Stopwatch();
            try
            {
                OpenFile("Excel files(*.xls*)|*.xls*");
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Подождите, идёт обработка файла Excel.\n";
                }));
                excelApp = new Excel.Application
                {
                    Visible = false,
                    ScreenUpdating = false,
                    EnableEvents = false
                };
                
                workBooks = excelApp.Workbooks;
                workBook = workBooks.Open(FileName);
                sw.Start();
                workSheets = workBook.Worksheets;
                workSheet = (Excel.Worksheet)workSheets.get_Item(1);

                ReadFromRow(workSheet);

                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Файл Excel, обработан. Можете продолжить работу.\n";

                    sw.Stop();

                    mainWindow.OutputText.Text += "Время обработки файла Excel: " + (sw.ElapsedMilliseconds / 1000.0).ToString() + " сек.\n";
                    mainWindow.OutputText.Text += "Обработанно: " + datas.Count + " строк.\n";
                    mainWindow.OutputText.ScrollToEnd();
                }));
                await mainWindow.ProcessCorelDRAWFile.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.ProcessCorelDRAWFile.IsEnabled = true;
                }));
            }
            catch (OperationCanceledException)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Операция была отменена пользователем!\n";
                    mainWindow.OutputText.ScrollToEnd();
                }));
            }
            catch (Exception ex)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += $"Work is failed.\n{ex.Message}\n";
                    mainWindow.OutputText.ScrollToEnd();
                }));
            }
            finally
            {
                workBook.Close();
                excelApp.Quit();
                await mainWindow.ProcessExcelFile.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.ProcessExcelFile.IsEnabled = true;
                }));
                Marshal.ReleaseComObject(workSheet);
                Marshal.ReleaseComObject(workSheets);
                Marshal.ReleaseComObject(workBook);
                Marshal.ReleaseComObject(workBooks);
                Marshal.ReleaseComObject(excelApp);
            }
        }

        public async Task StartExcelTaskAsync(CancellationTokenSource cts)
        {
            await Task.Run(() => ExtractDataFromExcel(cts));
        }

        public async Task StartCorelTaskAsync(CancellationTokenSource cts)
        {
            await Task.Run(() => StartCorelDRAWAsync(datas, cts));
        }

        async void StartCorelDRAWAsync(List<DataModel> datas, CancellationTokenSource cts)
        {
            CorelDRAW.Application corelApp = null;
            VGCore.Document document = null;
            VGCore.Page page = null;
            VGCore.Layer layer = null;
            VGCore.Shape shape = null;
            VGCore.Rect rect = null;
            VGCore.ImportFilter importFilter = null;
            VGCore.DataItem image = null;
            RectanglePosition rectanglePosition;
            RGBAssign rgbAssign;
            ArtisticText artisticText;
            CMYKAssign cmykAssign;
            ParagraphText paragraphText;
            Stopwatch sw = new Stopwatch();
            int count = 0;
            string fullPath;
            List<DataModel> data;

            try
            {
                data = datas;
                OpenFile("CorelDRAW files(*.cdr)|*.cdr");
                await mainWindow.ProgressBar.Dispatcher.BeginInvoke(DispatcherPriority.Send, new Action(delegate ()
                {
                    mainWindow.ProgressBar.Value = 0;
                }));
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Подождите, идёт обработка файла CorelDRAW.\n";
                }));
                corelApp = new CorelDRAW.Application
                {
                    Visible = false,
                    Optimization = true,
                    EventsEnabled = false
                };

                document = corelApp.OpenDocument(FileName, 1);
                document.BeginCommandGroup("Fast");
                document.SaveSettings();
                document.PreserveSelection = false;

                sw.Start();

                //document = corelApp.CreateDocument();
                document.Unit = VGCore.cdrUnit.cdrMillimeter;

                foreach (DataModel item in data)
                {
                    if (item.IsUp)
                    {
                        page = document.InsertPagesEx(1, false, document.ActivePage.Index, 297, 210);
                        layer = page.Layers[2];

                        rect = new VGCore.Rect
                        {
                            Width = 297,
                            Height = 100
                        };
                        rgbAssign = new RGBAssign(255, 255, 0);
                        rectanglePosition = new RectanglePosition(0, 210);
                        CreateRectangleRectAsync(rect, layer, rectanglePosition, rgbAssign, cts);

                        rgbAssign = new RGBAssign(255, 72, 41);
                        artisticText = new ArtisticText(31.369, 138.2776, item.ChildName, "Kabarett Simple", 205, "Name1");
                        CreateArtisticTextAsync(layer, artisticText, rgbAssign, cts);

                        cmykAssign = new CMYKAssign(0, 0, 0, 100);
                        paragraphText = new ParagraphText(13.555, 100.591, 29.112, 109.311, item.ImageNumber, "Arial", 18, "ImageNumber1");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(100, 0, 0, 0);
                        paragraphText = new ParagraphText(29.669, 100.591, 45.227, 109.311, item.DoorWidth.ToString(), "Arial", 18, "DoorWidth1");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(100, 0, 100, 0);
                        paragraphText = new ParagraphText(45.783, 100.591, 72.548, 109.311, item.Pocket, "Arial", 18, "Pocket1");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(0, 88, 97, 0);
                        paragraphText = new ParagraphText(73.104, 100.591, 149.288, 109.311, item.Customer, "Arial", 9, "Customer1");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        fullPath = Path.GetDirectoryName(FileName) + @"\img\" + item.ImageNumber + ".png";

                        if (item.ImageNumber != "0")
                        {
                            importFilter = layer.ImportEx(fullPath, VGCore.cdrFilter.cdrPNG);
                            importFilter.Finish();

                            shape = page.Shapes[item.ImageNumber + ".png"];
                            image = shape.ObjectData["Name"];
                            image.Value = item.ImageNumber;
                            if (item.DoorWidth <= 25)
                            {
                                shape.PositionX = 202.438 - shape.SizeWidth;
                                shape.PositionY = 159.9946 + shape.SizeHeight / 2;
                            }
                            else if (item.DoorWidth > 25 && item.DoorWidth < 29)
                            {
                                shape.PositionX = 258.8768 - shape.SizeWidth;
                                shape.PositionY = 159.9946 + shape.SizeHeight / 2;
                            }
                            else if (item.DoorWidth >= 29)
                            {
                                shape.PositionX = 272.669 - shape.SizeWidth;
                                shape.PositionY = 159.9946 + shape.SizeHeight / 2;
                            }
                        }

                        count++;
                        await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                        {
                            mainWindow.OutputText.Text += "Обработанно: " + count + " строк.\n";
                            mainWindow.OutputText.ScrollToEnd();
                        }));
                        await mainWindow.ProgressBar.Dispatcher.BeginInvoke(DispatcherPriority.Send, new Action(delegate ()
                        {
                            mainWindow.ProgressBar.Value = (double)count * 100 / data.Count;
                        }));
                    }
                    else
                    {
                        rect = new VGCore.Rect
                        {
                            Width = 297,
                            Height = 100
                        };
                        rgbAssign = new RGBAssign(255, 255, 0);
                        rectanglePosition = new RectanglePosition(0, 100);
                        CreateRectangleRectAsync(rect, layer, rectanglePosition, rgbAssign, cts);

                        rgbAssign = new RGBAssign(255, 72, 41);
                        artisticText = new ArtisticText(31.369, 28.7782, item.ChildName, "Kabarett Simple", 205, "Name2");
                        CreateArtisticTextAsync(layer, artisticText, rgbAssign, cts);

                        cmykAssign = new CMYKAssign(0, 0, 0, 100);
                        paragraphText = new ParagraphText(226.538, 100.591, 242.096, 109.311, item.ImageNumber, "Arial", 18, "ImageNumber2");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(100, 0, 0, 0);
                        paragraphText = new ParagraphText(242.652, 100.591, 258.21, 109.311, item.DoorWidth.ToString(), "Arial", 18, "DoorWidth2");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(100, 0, 100, 0);
                        paragraphText = new ParagraphText(258.767, 100.591, 285.868, 109.311, item.Pocket, "Arial", 18, "Pocket2");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(0, 88, 97, 0);
                        paragraphText = new ParagraphText(149.845, 100.591, 225.982, 109.311, item.Customer, "Arial", 9, "Customer2");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        fullPath = Path.GetDirectoryName(FileName) + @"\img\" + item.ImageNumber + ".png";

                        if (item.ImageNumber != "0")
                        {
                            importFilter = layer.ImportEx(fullPath, VGCore.cdrFilter.cdrPNG);
                            importFilter.Finish();

                            shape = page.Shapes[item.ImageNumber + ".png"];
                            image = shape.ObjectData["Name"];
                            image.Value = item.ImageNumber;
                            if (item.DoorWidth <= 25)
                            {
                                shape.PositionX = 202.438 - shape.SizeWidth;
                                shape.PositionY = 50.0126 + shape.SizeHeight / 2;
                            }
                            else if (item.DoorWidth > 25 && item.DoorWidth < 29)
                            {
                                shape.PositionX = 258.8768 - shape.SizeWidth;
                                shape.PositionY = 50.0126 + shape.SizeHeight / 2;
                            }
                            else if (item.DoorWidth >= 29)
                            {
                                shape.PositionX = 272.669 - shape.SizeWidth;
                                shape.PositionY = 50.0126 + shape.SizeHeight / 2;
                            }
                        }

                        count++;
                        await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                        {
                            mainWindow.OutputText.Text += "Обработанно: " + count + " строк.\n";
                            mainWindow.OutputText.ScrollToEnd();
                        }));
                        await mainWindow.ProgressBar.Dispatcher.BeginInvoke(DispatcherPriority.Send, new Action(delegate ()
                        {
                            mainWindow.ProgressBar.Value = (double)count * 100 / data.Count;
                        }));
                    }
                }

                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Файл CorelDRAW, обработан. Можете продолжить работу.\n";

                    sw.Stop();

                    mainWindow.OutputText.Text += "Время обработки файла CorelDRAW: " + (sw.ElapsedMilliseconds / 1000.0).ToString() + " сек.\n";
                    mainWindow.OutputText.Text += "Обработанно: " + data.Count + " строк.\n";
                    mainWindow.OutputText.ScrollToEnd();
                }));

                await mainWindow.ProgressBar.Dispatcher.BeginInvoke(DispatcherPriority.Send, new Action(delegate ()
                {
                    mainWindow.ProgressBar.Value = 0;
                }));

                document.PreserveSelection = true;
                document.ResetSettings();
                corelApp.EventsEnabled = true;
                corelApp.Optimization = false;
                document.EndCommandGroup();
                corelApp.Refresh();
                corelApp.ActiveWindow.Refresh();
                corelApp.Visible = true;
            }
            catch (OperationCanceledException)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Операция была отменена пользователем!\n";
                    mainWindow.OutputText.ScrollToEnd();
                }));
            }
            catch (Exception ex)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += $"Work is failed.\n{ex.Message}\n";
                    mainWindow.OutputText.ScrollToEnd();
                }));
            }
            finally
            {
                Marshal.ReleaseComObject(rect);
                Marshal.ReleaseComObject(image);
                Marshal.ReleaseComObject(importFilter);
                Marshal.ReleaseComObject(shape);
                Marshal.ReleaseComObject(layer);
                Marshal.ReleaseComObject(page);
                Marshal.ReleaseComObject(document);
                Marshal.ReleaseComObject(corelApp);
            }
        }

        async void CreateRectangleRectAsync(VGCore.Rect rect, VGCore.Layer layer, RectanglePosition rectanglePosition, RGBAssign rgbAssign, CancellationTokenSource cts)
        {
            VGCore.Shape shape = null;
            VGCore.Outline outline = null;
            VGCore.Fill fill = null;
            VGCore.Color color = null;

            try
            {
                shape = layer.CreateRectangleRect(rect);
                shape.PositionX = rectanglePosition.PositionX;
                shape.PositionY = rectanglePosition.PositionY;
                outline = shape.Outline;
                outline.Type = VGCore.cdrOutlineType.cdrNoOutline;
                fill = shape.Fill;
                color = fill.UniformColor;
                color.RGBAssign(
                    rgbAssign.R,
                    rgbAssign.G,
                    rgbAssign.B
                );
            }
            catch (OperationCanceledException)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Операция была отменена пользователем!\n";
                }));
            }
            catch (Exception ex)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += $"Work is failed.\n{ex.Message}\n";
                }));
            }
            finally
            {
                Marshal.ReleaseComObject(color);
                Marshal.ReleaseComObject(fill);
                Marshal.ReleaseComObject(outline);
                Marshal.ReleaseComObject(shape);
            }
        }

        async void CreateArtisticTextAsync(VGCore.Layer layer, ArtisticText artisticText, RGBAssign rgbAssign, CancellationTokenSource cts)
        {
            VGCore.Shape shape = null;
            VGCore.Fill fill = null;
            VGCore.Color color = null;

            try
            {
                shape = layer.CreateArtisticText(
                    artisticText.Left,
                    artisticText.Bottom,
                    artisticText.Text,
                    VGCore.cdrTextLanguage.cdrLanguageNone,
                    VGCore.cdrTextCharSet.cdrCharSetMixed,
                    artisticText.Font,
                    artisticText.Size,
                    VGCore.cdrTriState.cdrTrue,
                    VGCore.cdrTriState.cdrTrue,
                    VGCore.cdrFontLine.cdrMixedFontLine,
                    VGCore.cdrAlignment.cdrLeftAlignment
                );
                fill = shape.Fill;
                color = fill.UniformColor;
                color.RGBAssign(
                    rgbAssign.R,
                    rgbAssign.G,
                    rgbAssign.B
                );
                shape.Name = artisticText.Name;
                shape.SizeWidth = artisticText.SizeWidth;
            }
            catch (OperationCanceledException)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Операция была отменена пользователем!\n";
                }));
            }
            catch (Exception ex)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += $"Work is failed.\n{ex.Message}\n";
                }));
            }
            finally
            {
                Marshal.ReleaseComObject(color);
                Marshal.ReleaseComObject(fill);
                Marshal.ReleaseComObject(shape);
            }
        }

        async void CreateParagraphTextAsync(VGCore.Layer layer, ParagraphText paragraphText, CMYKAssign cmykAssign, CancellationTokenSource cts)
        {
            VGCore.Shape shape = null;
            VGCore.Outline outline = null;
            VGCore.Fill fill = null;
            VGCore.Color color = null;
            VGCore.Text text = null;
            VGCore.TextRange story = null;

            try
            {
                shape = layer.CreateParagraphText(
                    paragraphText.Left,
                    paragraphText.Top,
                    paragraphText.Right,
                    paragraphText.Bottom,
                    paragraphText.Text
                );
                fill = shape.Fill;
                color = fill.UniformColor;
                color.CMYKAssign(
                    cmykAssign.C,
                    cmykAssign.M,
                    cmykAssign.Y,
                    cmykAssign.K
                );
                text = shape.Text;
                story = text.Story;
                story.Style = VGCore.cdrFontStyle.cdrBoldFontStyle;
                story.Size = paragraphText.Size;
                story.Alignment = VGCore.cdrAlignment.cdrLeftAlignment;
                outline = shape.Outline;
                outline.SetNoOutline();
                shape.Name = paragraphText.Name;
            }
            catch (OperationCanceledException)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Операция была отменена пользователем!\n";
                }));
            }
            catch (Exception ex)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += $"Work is failed.\n{ex.Message}\n";
                }));
            }
            finally
            {
                Marshal.ReleaseComObject(story);
                Marshal.ReleaseComObject(text);
                Marshal.ReleaseComObject(color);
                Marshal.ReleaseComObject(fill);
                Marshal.ReleaseComObject(outline);
                Marshal.ReleaseComObject(shape);
            }
        }
    }
}
