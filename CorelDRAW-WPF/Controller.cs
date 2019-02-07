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
                    Customer = workSheet.Cells[row, 1].Text.Trim(),
                    ChildName = workSheet.Cells[row, 2].Text.Trim(),
                    DoorWidth = Convert.ToDouble(workSheet.Cells[row, 4].Text.Trim()),
                    Pocket = workSheet.Cells[row, 6].Text.Trim()
                };

                if (int.TryParse(workSheet.Cells[row, 3].Text.Trim(), out int imgNum))
                {
                    dataModel.ImageNumber = imgNum.ToString();
                }
                else
                {
                    dataModel.ImageNumber = "0";
                }

                if (int.TryParse(workSheet.Cells[row, 5].Text.Trim(), out int backgroundNum))
                {
                    dataModel.BackgroundNumber = backgroundNum.ToString();
                }
                else
                {
                    dataModel.BackgroundNumber = "1";
                }

                if (Convert.ToInt32(workSheet.Cells[row, 7].Text.Trim()) == 1)
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
            VGCore.Shape siteLogo = null;
            VGCore.Rect rect = null;
            VGCore.ImportFilter importFilter = null;
            VGCore.DataItem image = null;
            const string LOGO = "www.vash-sadik.com";
            RectanglePosition rectanglePosition;
            RGBAssign rgbAssign;
            ArtisticText artisticText;
            CMYKAssign cmykAssign;
            ParagraphText paragraphText;
            Stopwatch sw = new Stopwatch();
            int count = 0;
            string fullPath;
            string[] name;
            string fullName;
            float fontSize;
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

                document.Unit = VGCore.cdrUnit.cdrMillimeter;

                foreach (DataModel item in data)
                {
                    if (item.IsUp)
                    {
                        page = document.InsertPagesEx(1, false, document.ActivePage.Index, 297, 210);
                        layer = page.Layers[2];

                        //rect = new VGCore.Rect
                        //{
                        //    Width = 297,
                        //    Height = 100
                        //};
                        //rgbAssign = new RGBAssign(255, 255, 255);
                        //rectanglePosition = new RectanglePosition(0, 210);
                        //CreateRectangleRectAsync(rect, layer, rectanglePosition, rgbAssign, cts);

                        //Add background image
                        fullPath = Path.GetDirectoryName(FileName) + @"\fon\" + item.BackgroundNumber + ".jpeg";
                        if (item.BackgroundNumber != "0")
                        {
                            importFilter = layer.ImportEx(fullPath, VGCore.cdrFilter.cdrJPEG);
                            importFilter.Finish();

                            shape = page.Shapes[item.BackgroundNumber + ".jpeg"];
                            image = shape.ObjectData["Name"];
                            image.Value = item.BackgroundNumber;

                            shape.SizeWidth = 297;
                            shape.SizeHeight = 100;
                            shape.PositionX = 0;
                            shape.PositionY = 210;
                        }

                        name = item.ChildName.Split(' ');
                        if (name.Length > 1)
                        {
                            fullName = name[0] + "\r\n" + name[1];
                            fontSize = 102.5f;
                        }
                        else
                        {
                            fullName = name[0];
                            fontSize = 205f;
                        }

                        rgbAssign = new RGBAssign(255, 72, 41);
                        artisticText = new ArtisticText(31.369, 138.2776, fullName, "Kabarett Simple", fontSize, "Name1");
                        CreateArtisticTextAsync(layer, artisticText, rgbAssign, cts);
                        shape = page.Shapes["Name1"];
                        //shape.SizeHeight = 61.761;
                        shape.CenterY = 160;

                        cmykAssign = new CMYKAssign(0, 0, 0, 100);
                        paragraphText = new ParagraphText(13.555, 105, 29.112, 110, item.ImageNumber, "Arial", 12, "ImageNumber1");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(100, 0, 0, 0);
                        paragraphText = new ParagraphText(29.669, 105, 45.227, 110, item.DoorWidth.ToString(), "Arial", 12, "DoorWidth1");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(100, 0, 100, 0);
                        paragraphText = new ParagraphText(45.783, 105, 72.548, 110, item.Pocket, "Arial", 12, "Pocket1");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(0, 88, 97, 0);
                        paragraphText = new ParagraphText(73.104, 105, 149.288, 110, item.Customer, "Arial", 12, "Customer1");
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        fullPath = Path.GetDirectoryName(FileName) + @"\img\" + item.ImageNumber + ".png";

                        rgbAssign = new RGBAssign(255, 41, 41);
                        artisticText = new ArtisticText(193.745, 128.413, LOGO, "Arial", 16.591f, "Logo1", 63.174);
                        CreateArtisticTextAsync(layer, artisticText, rgbAssign, cts);
                        siteLogo = page.Shapes["Logo1"];
                        siteLogo.Rotate(90);
                        siteLogo.SizeHeight = 63.174;
                        siteLogo.SizeWidth = 4.255;

                        if (item.DoorWidth < 23)
                        {
                            siteLogo.CenterX = 195.872;
                        }
                        else if (item.DoorWidth >= 23 && item.DoorWidth < 25)
                        {
                            siteLogo.CenterX = 215.872;
                        }
                        else if (item.DoorWidth >= 25 && item.DoorWidth < 27)
                        {
                            siteLogo.CenterX = 235.872;
                        }
                        else if (item.DoorWidth >= 27 && item.DoorWidth < 29)
                        {
                            siteLogo.CenterX = 255.872;
                        }
                        else if (item.DoorWidth >= 29)
                        {
                            siteLogo.CenterX = 275.872;
                        }
                        siteLogo.CenterY = 160;

                        if (item.ImageNumber != "0")
                        {
                            importFilter = layer.ImportEx(fullPath, VGCore.cdrFilter.cdrPNG);
                            importFilter.Finish();

                            shape = page.Shapes[item.ImageNumber + ".png"];
                            image = shape.ObjectData["Name"];
                            image.Value = item.ImageNumber;

                            shape.CenterX = siteLogo.CenterX - siteLogo.SizeWidth / 2 - 10 - shape.SizeWidth / 2;
                            shape.CenterY = siteLogo.CenterY;
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
                        //rect = new VGCore.Rect
                        //{
                        //    Width = 297,
                        //    Height = 100
                        //};
                        //rgbAssign = new RGBAssign(255, 255, 255);
                        //rectanglePosition = new RectanglePosition(0, 100);
                        //CreateRectangleRectAsync(rect, layer, rectanglePosition, rgbAssign, cts);

                        fullPath = Path.GetDirectoryName(FileName) + @"\fon\" + item.BackgroundNumber + ".jpeg";
                        if (item.BackgroundNumber != "0")
                        {
                            importFilter = layer.ImportEx(fullPath, VGCore.cdrFilter.cdrJPEG);
                            importFilter.Finish();

                            shape = page.Shapes[item.BackgroundNumber + ".jpeg"];
                            image = shape.ObjectData["Name"];
                            image.Value = item.BackgroundNumber;

                            shape.SizeWidth = 297;
                            shape.SizeHeight = 100;
                            shape.PositionX = 0;
                            shape.PositionY = 100;
                        }

                        name = item.ChildName.Split(' ');
                        if (name.Length > 1)
                        {
                            fullName = name[0] + "\r\n" + name[1];
                            fontSize = 102.5f;
                        }
                        else
                        {
                            fullName = name[0];
                            fontSize = 205f;
                        }
                        rgbAssign = new RGBAssign(255, 72, 41);
                        artisticText = new ArtisticText(31.369, 28.7782, fullName, "Kabarett Simple", fontSize, "Name2");
                        CreateArtisticTextAsync(layer, artisticText, rgbAssign, cts);
                        shape = page.Shapes["Name2"];
                        shape.CenterY = 50;

                        cmykAssign = new CMYKAssign(0, 0, 0, 100);
                        paragraphText = new ParagraphText(226.538, 100, 242.096, 105, item.ImageNumber, "Arial", 12, "ImageNumber2", 2);
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(100, 0, 0, 0);
                        paragraphText = new ParagraphText(242.652, 100, 258.21, 105, item.DoorWidth.ToString(), "Arial", 12, "DoorWidth2", 2);
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(100, 0, 100, 0);
                        paragraphText = new ParagraphText(258.767, 100, 285.868, 105, item.Pocket, "Arial", 12, "Pocket2", 2);
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        cmykAssign = new CMYKAssign(0, 88, 97, 0);
                        paragraphText = new ParagraphText(149.845, 100, 225.982, 105, item.Customer, "Arial", 12, "Customer2", 2);
                        CreateParagraphTextAsync(layer, paragraphText, cmykAssign, cts);

                        fullPath = Path.GetDirectoryName(FileName) + @"\img\" + item.ImageNumber + ".png";

                        rgbAssign = new RGBAssign(255, 41, 41);
                        artisticText = new ArtisticText(193.745, 128.413, LOGO, "Arial", 16.591f, "Logo2", 63.174);
                        CreateArtisticTextAsync(layer, artisticText, rgbAssign, cts);
                        siteLogo = page.Shapes["Logo2"];
                        siteLogo.Rotate(90);
                        siteLogo.SizeHeight = 63.174;
                        siteLogo.SizeWidth = 4.255;

                        if (item.DoorWidth < 23)
                        {
                            siteLogo.CenterX = 195.872;
                        }
                        else if (item.DoorWidth >= 23 && item.DoorWidth < 25)
                        {
                            siteLogo.CenterX = 215.872;
                        }
                        else if (item.DoorWidth >= 25 && item.DoorWidth < 27)
                        {
                            siteLogo.CenterX = 235.872;
                        }
                        else if (item.DoorWidth >= 27 && item.DoorWidth < 29)
                        {
                            siteLogo.CenterX = 255.872;
                        }
                        else if (item.DoorWidth >= 29)
                        {
                            siteLogo.CenterX = 275.872;
                        }
                        siteLogo.CenterY = 50;

                        if (item.ImageNumber != "0")
                        {
                            importFilter = layer.ImportEx(fullPath, VGCore.cdrFilter.cdrPNG);
                            importFilter.Finish();

                            shape = page.Shapes[item.ImageNumber + ".png"];
                            image = shape.ObjectData["Name"];
                            image.Value = item.ImageNumber;

                            shape.CenterX = siteLogo.CenterX - siteLogo.SizeWidth / 2 - 10 - shape.SizeWidth / 2;
                            shape.CenterY = siteLogo.CenterY;
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
                    VGCore.cdrTriState.cdrFalse,
                    VGCore.cdrTriState.cdrFalse,
                    VGCore.cdrFontLine.cdrMixedFontLine,
                    (VGCore.cdrAlignment)artisticText.Alignment
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
                story.Alignment = (VGCore.cdrAlignment)paragraphText.Alignment;//VGCore.cdrAlignment.cdrLeftAlignment;
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
