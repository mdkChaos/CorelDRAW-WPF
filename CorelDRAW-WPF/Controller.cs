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

        public void ExtractDataFromExcel()
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
                excelApp = new Excel.Application
                {
                    Visible = false,
                    ScreenUpdating = false,
                    EnableEvents = false
                };
                OpenFile("Excel files(*.xls*)|*.xls*");

                workBooks = excelApp.Workbooks;
                workBook = workBooks.Open(FileName);
                sw.Start();
                workSheets = workBook.Worksheets;
                workSheet = (Excel.Worksheet)workSheets.get_Item(1);
                mainWindow.OutputText.Text += "Подождите, идёт обработка файла Excel.\n";
                ReadFromRow(workSheet);
                mainWindow.OutputText.Text += "Файл Excel, обработан. Можете продолжить работу.\n";
                sw.Stop();
                mainWindow.OutputText.Text += "Время обработки файла Excel: " + (sw.ElapsedMilliseconds / 1000.0).ToString() + " сек.\n";
                mainWindow.OutputText.Text += "Обработанно: " + datas.Count + " строк.\n";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                workBook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(workSheet);
                Marshal.ReleaseComObject(workSheets);
                Marshal.ReleaseComObject(workBook);
                Marshal.ReleaseComObject(workBooks);
                Marshal.ReleaseComObject(excelApp);
            }
        }

        public void InsertDataToCorel()
        {
            CorelDRAW.Application corelApp = null;
            VGCore.Document document = null;
            VGCore.Pages pages = null;
            VGCore.Page page = null;
            VGCore.Layer layer = null;
            VGCore.Shapes shapes = null;
            VGCore.ShapeRange shapeRange = null;
            Stopwatch sw = new Stopwatch();

            try
            {
                corelApp = new CorelDRAW.Application
                {
                    Visible = false,
                    Optimization = true,
                    EventsEnabled = false
                };

                OpenFile("CorelDRAW files(*.cdr)|*.cdr");
                document = corelApp.OpenDocument(FileName, 1);
                document.BeginCommandGroup("Fast");
                document.SaveSettings();
                document.PreserveSelection = false;

                sw.Start();

                pages = document.Pages;
                page = pages.First;
                layer = page.Layers["Layer1"];
                shapes = layer.Shapes;
                shapeRange = shapes.All();

                mainWindow.OutputText.Text += "Подождите, идёт обработка файла CorelDRAW.\n";
                ProcessToCorel(corelApp, document, shapeRange);
                mainWindow.OutputText.Text += "Файл CorelDRAW, обработан. Можете продолжить работу.\n";

                sw.Stop();

                mainWindow.OutputText.Text += "Время обработки файла CorelDRAW: " + (sw.ElapsedMilliseconds / 1000.0).ToString() + " сек.\n";
                mainWindow.OutputText.Text += "Обработанно: " + datas.Count + " строк.\n";

                document.PreserveSelection = true;
                document.ResetSettings();
                corelApp.EventsEnabled = true;
                corelApp.Optimization = false;
                document.EndCommandGroup();
                corelApp.Refresh();
                corelApp.ActiveWindow.Refresh();
                corelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(shapeRange);
                Marshal.ReleaseComObject(shapes);
                Marshal.ReleaseComObject(layer);
                Marshal.ReleaseComObject(page);
                Marshal.ReleaseComObject(pages);
                Marshal.ReleaseComObject(document);
                Marshal.ReleaseComObject(corelApp);
            }
        }

        private void ProcessToCorel(CorelDRAW.Application corelApp, VGCore.Document document, VGCore.ShapeRange shapeRange)
        {
            string fullPath;
            VGCore.Page newPage = null;
            VGCore.Layer newLayer = null;
            VGCore.Shape shape = null;
            VGCore.Text text = null;
            VGCore.TextRange textRange = null;
            VGCore.ImportFilter importFilter = null;
            VGCore.DataItem image = null;

            foreach (DataModel item in datas)
            {
                if (item.IsUp)
                {
                    shapeRange.Copy();
                    newPage = document.InsertPagesEx(1, false, document.ActivePage.Index, 11.692913, 8.267717);
                    newLayer = newPage.Layers["Layer1"];
                    newLayer.PasteEx(corelApp.CreateStructPasteOptions());

                    shape = newPage.Shapes["Name1"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.ChildName;
                    shape.SizeWidth = 4.5;

                    shape = newPage.Shapes["Customer1"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.Customer;

                    shape = newPage.Shapes["Pocket1"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.Pocket;

                    shape = newPage.Shapes["DoorWidth1"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.DoorWidth.ToString();

                    shape = newPage.Shapes["ImageNumber1"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.ImageNumber;

                    fullPath = Path.GetDirectoryName(FileName) + @"\img\" + item.ImageNumber + ".png";

                    if (item.ImageNumber != "0")
                    {
                        importFilter = newLayer.ImportEx(fullPath, VGCore.cdrFilter.cdrPNG);
                        importFilter.Finish();

                        shape = newPage.Shapes[item.ImageNumber + ".png"];
                        image = shape.ObjectData["Name"];
                        image.Value = item.ImageNumber;
                        if (Convert.ToDouble(item.DoorWidth) <= 25)
                        {
                            shape.PositionX = 7.97 - shape.SizeWidth;
                            shape.PositionY = 6.299 + shape.SizeHeight / 2;
                        }
                        else if (Convert.ToDouble(item.DoorWidth) > 25 && Convert.ToDouble(item.DoorWidth) < 29)
                        {
                            shape.PositionX = 10.192 - shape.SizeWidth;
                            shape.PositionY = 6.299 + shape.SizeHeight / 2;
                        }
                        else if (Convert.ToDouble(item.DoorWidth) >= 29)
                        {
                            shape.PositionX = 10.735 - shape.SizeWidth;
                            shape.PositionY = 6.299 + shape.SizeHeight / 2;
                        }
                    }
                }
                else
                {
                    newPage = document.Pages[2];
                    newLayer = newPage.Layers["Layer1"];

                    shape = newPage.Shapes["Name2"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.ChildName;
                    shape.SizeWidth = 4.5;

                    shape = newPage.Shapes["Customer2"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.Customer;

                    shape = newPage.Shapes["Pocket2"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.Pocket;

                    shape = newPage.Shapes["DoorWidth2"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.DoorWidth.ToString();

                    shape = newPage.Shapes["ImageNumber2"];
                    text = shape.Text;
                    textRange = text.Story;
                    textRange.Text = item.ImageNumber;

                    fullPath = Path.GetDirectoryName(FileName) + @"\img\" + item.ImageNumber + ".png";

                    if (item.ImageNumber != "0")
                    {
                        importFilter = newLayer.ImportEx(fullPath, VGCore.cdrFilter.cdrPNG);
                        importFilter.Finish();

                        shape = newPage.Shapes[item.ImageNumber + ".png"];
                        image = shape.ObjectData["Name"];
                        image.Value = item.ImageNumber;
                        if (item.DoorWidth <= 25)
                        {
                            shape.PositionX = 7.97 - shape.SizeWidth;
                            shape.PositionY = 1.969 + shape.SizeHeight / 2;
                        }
                        else if (item.DoorWidth > 25 && item.DoorWidth < 29)
                        {
                            shape.PositionX = 10.192 - shape.SizeWidth;
                            shape.PositionY = 1.969 + shape.SizeHeight / 2;
                        }
                        else if (item.DoorWidth >= 29)
                        {
                            shape.PositionX = 10.735 - shape.SizeWidth;
                            shape.PositionY = 1.969 + shape.SizeHeight / 2;
                        }
                    }
                }
            }

            Marshal.ReleaseComObject(image);
            Marshal.ReleaseComObject(importFilter);
            Marshal.ReleaseComObject(textRange);
            Marshal.ReleaseComObject(newLayer);
            Marshal.ReleaseComObject(newPage);
            Marshal.ReleaseComObject(text);
            Marshal.ReleaseComObject(shape);
        }

        public async Task GetTaskAsync(CancellationTokenSource cts)
        {
            await Task.Run(() => StartCorelDRAWAsync(cts));
        }

        async void StartCorelDRAWAsync(CancellationTokenSource cts)
        {
            CorelDRAW.Application corelApp = null;
            VGCore.Document document = null;
            VGCore.Page page = null;
            VGCore.Layer layer = null;
            VGCore.Rect rect = null;
            RectanglePosition rectanglePosition;
            RGBAssign rgbAssign;
            ArtisticText artisticText;
            CMYKAssign cmykAssign;
            ParagraphText paragraphText;

            try
            {
                corelApp = new CorelDRAW.Application
                {
                    Visible = false,
                    Optimization = true,
                    EventsEnabled = false
                };
                document = corelApp.CreateDocument();
                document.Unit = VGCore.cdrUnit.cdrMillimeter;
                page = document.InsertPagesEx(1, false, document.ActivePage.Index, 297, 210);

                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Файл создан.\n";
                    mainWindow.OutputText.Text += "Создадим прямоугольник.\n";
                }));

                //await MainWindow.progressBar.Dispatcher.BeginInvoke(DispatcherPriority.Send, new Action(delegate ()
                //{
                //    MainWindow.progressBar.Value = 0;
                //}));

                layer = page.Layers[2];

                rect = new VGCore.Rect
                {
                    Width = 297,
                    Height = 100
                };
                rgbAssign = new RGBAssign(255, 255, 0);
                rectanglePosition = new RectanglePosition(0, 210);
                CreateRectangleRect(rect, layer, rectanglePosition, rgbAssign);
                rectanglePosition = new RectanglePosition(0, 100);
                CreateRectangleRect(rect, layer, rectanglePosition, rgbAssign);

                rgbAssign = new RGBAssign(255, 72, 41);
                artisticText = new ArtisticText(31.369, 138.2776, "Name 1", "Kabarett Simple", 205, "Name1");
                CreateArtisticText(layer, artisticText, rgbAssign);
                artisticText = new ArtisticText(31.369, 28.7782, "Name 2", "Kabarett Simple", 205, "Name2");
                CreateArtisticText(layer, artisticText, rgbAssign);

                cmykAssign = new CMYKAssign(0, 0, 0, 100);
                paragraphText = new ParagraphText(13.555, 100.591, 29.112, 109.311, "номер рисунка", "Arial", 18, "ImageNumber1");
                CreateParagraphText(layer, paragraphText, cmykAssign);
                paragraphText = new ParagraphText(226.538, 100.591, 242.096, 109.311, "номер рисунка", "Arial", 18, "ImageNumber2");
                CreateParagraphText(layer, paragraphText, cmykAssign);

                cmykAssign = new CMYKAssign(100, 0, 0, 0);
                paragraphText = new ParagraphText(29.669, 100.591, 45.227, 109.311, "ширина дверцы", "Arial", 18, "DoorWidth1");
                CreateParagraphText(layer, paragraphText, cmykAssign);
                paragraphText = new ParagraphText(242.652, 100.591, 258.21, 109.311, "ширина дверцы", "Arial", 18, "DoorWidth2");
                CreateParagraphText(layer, paragraphText, cmykAssign);

                cmykAssign = new CMYKAssign(100, 0, 100, 0);
                paragraphText = new ParagraphText(45.783, 100.591, 72.548, 109.311, "кармашек", "Arial", 18, "Pocket1");
                CreateParagraphText(layer, paragraphText, cmykAssign);
                paragraphText = new ParagraphText(258.767, 100.591, 285.868, 109.311, "кармашек", "Arial", 18, "Pocket2");
                CreateParagraphText(layer, paragraphText, cmykAssign);

                cmykAssign = new CMYKAssign(0, 88, 97, 0);
                paragraphText = new ParagraphText(73.104, 100.591, 149.288, 109.311, "заказчик", "Arial", 9, "Customer1");
                CreateParagraphText(layer, paragraphText, cmykAssign);
                paragraphText = new ParagraphText(149.845, 100.591, 225.982, 109.311, "заказчик", "Arial", 9, "Customer2");
                CreateParagraphText(layer, paragraphText, cmykAssign);

                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Прямоугольник создан.\n";
                }));

                corelApp.EventsEnabled = true;
                corelApp.Optimization = false;
                corelApp.Visible = true;
            }
            catch (OperationCanceledException)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += "Операция была отменена пользователем!\n";
                }));
            }
            catch (Exception)
            {
                await mainWindow.OutputText.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    mainWindow.OutputText.Text += $"Work is failed.\n";
                }));
            }
            finally
            {
                Marshal.ReleaseComObject(rect);
                Marshal.ReleaseComObject(layer);
                Marshal.ReleaseComObject(page);
                Marshal.ReleaseComObject(document);
                Marshal.ReleaseComObject(corelApp);
            }
        }

        void CreateRectangleRect(VGCore.Rect rect, VGCore.Layer layer, RectanglePosition rectanglePosition, RGBAssign rgbAssign)
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
            catch (Exception)
            {
                throw;
            }
            finally
            {
                Marshal.ReleaseComObject(color);
                Marshal.ReleaseComObject(fill);
                Marshal.ReleaseComObject(outline);
                Marshal.ReleaseComObject(shape);
            }
        }

        void CreateArtisticText(VGCore.Layer layer, ArtisticText artisticText, RGBAssign rgbAssign)
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
            catch (Exception)
            {
                throw;
            }
            finally
            {
                Marshal.ReleaseComObject(color);
                Marshal.ReleaseComObject(fill);
                Marshal.ReleaseComObject(shape);
            }
        }

        void CreateParagraphText(VGCore.Layer layer, ParagraphText paragraphText, CMYKAssign cmykAssign)
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
            catch (Exception)
            {
                throw;
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
