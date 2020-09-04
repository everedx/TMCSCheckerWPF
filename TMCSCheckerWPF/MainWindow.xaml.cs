using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties;

namespace TMCSCheckerWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public delegate void UpdatePgBar(double value);
        public delegate void FinishExport();
        string[] arrayColumsExcel = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };



        List<String> typesList;
        List<String> devicesList;
        SqlConnection connection;
        string imagesDir;
        public MainWindow()
        {
            InitializeComponent();
       
            this.WindowStyle = WindowStyle.None;
        }


        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("Loaded");
            typesList = new List<string>();
            devicesList = new List<string>();
            imagesDir = Environment.GetEnvironmentVariable("ITS_CLIENT_PATH")+ @"\Resources\Images\ImagesMap\";
        }

        private void btnConnect_Click(object sender, RoutedEventArgs e)
        {
            bool enable = false;
            try
            {
                //Try to connect to DB
                string m_ConnectionString = "Server = " + tbIP.Text + "; Database = " + tbDB.Text + "; User Id = " + tbDB.Text + "; Password = " + tbPass.Text + ";";


                connection = new SqlConnection(m_ConnectionString);
                connection.Open();
                if (connection.State.Equals(ConnectionState.Open))
                {
                    enable = true;
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
              
                lblDvice.IsEnabled = enable;
                lblType.IsEnabled = enable;
                cbDevices.IsEnabled = enable;
                cbType.IsEnabled = enable;
                btnExport.IsEnabled = enable;
                btnExportImages.IsEnabled = enable;

                if (enable)
                {
                    LoadTypes();
                    btnConnect.IsEnabled = false;
                    
                }
            }

            


        }


        private void LoadTypes()
        {
            string queryString = "SELECT NOMBRE_TIPOEQUIPO FROM CONF_TIPOS_EQUIPOS WHERE ID_TEQUIP_PRINCIPAL = 'CPMV'";
            SqlCommand command = new SqlCommand(queryString, connection);
            SqlDataReader reader = command.ExecuteReader();
            
            while (reader.Read())
            {
                typesList.Add(reader.GetString(0));
                //MessageBox.Show(reader.GetString(0));
            }
            reader.Close();
            cbType.ItemsSource = typesList;
            cbType.SelectionChanged += LoadDevices;
            cbType.SelectedIndex = 0;
            
        }

        private void LoadDevices(Object source, EventArgs e)
        {
            cbDevices.SelectionChanged -= LoadIconsAssociated;
            cbDevices.ItemsSource = null;
            string queryString = "SELECT IDENT FROM CONF_EQUIPOS WHERE NOMBRE_TIPOEQUIPO = '"+ cbType.SelectedItem.ToString() + "'";
            SqlCommand command = new SqlCommand(queryString, connection);
            SqlDataReader reader = command.ExecuteReader();
            devicesList.Clear();
            while (reader.Read())
            {
                devicesList.Add(reader.GetString(0));
                //MessageBox.Show(reader.GetString(0));
            }
            reader.Close();
            
            cbDevices.ItemsSource = devicesList;
            // cbDevices.SelectionChanged += LoadDevices;
            cbDevices.SelectionChanged += LoadIconsAssociated;
            cbDevices.SelectedIndex = 0;
            
        }


        private void LoadIconsAssociated(Object source, EventArgs e)
        {
            int gridWidth = gridIcons.ColumnDefinitions.Count;
            string queryString = @"SELECT NOMBRE_ICONO FROM VIS_LISTA_ICONOS WHERE ID_LISTA IN
                                  (SELECT ID_LISTA FROM VIS_LISTA WHERE NOMBRE_LISTA IN(
                                   SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS 
                                   WHERE IDENT = '"+ cbDevices.SelectedItem.ToString() +"'"
                                  +@" and NOMBRE_CARACT = 'LISTAICONOS'))";
            int count = 0;
            SqlCommand command = new SqlCommand(queryString, connection);
            SqlDataReader reader = command.ExecuteReader();
            //gridIcons.Children.Add();
            System.Windows.Controls.Border b = (System.Windows.Controls.Border)gridIcons.Children[0];

            gridIcons.Children.Clear();
            gridIcons.Children.Add(b);
           
            bool addImage = true;
            while (reader.Read())
            {

                //var bitmap = new BitmapImage(uri);
                System.Windows.Controls.Image img = new System.Windows.Controls.Image();
                Label lab = new Label();
                try {
                    using (var fs = new System.IO.FileStream(imagesDir + "icono" + reader.GetString(0) + ".png", System.IO.FileMode.Open))
                    {
                        var bmp = new Bitmap(fs);
                        img.Source = BitmapToImageSource((Bitmap)bmp.Clone());
                    }
                    //BitmapImage bitmap = new BitmapImage(new Uri(imagesDir + "icono" + reader.GetString(0) + ".png"));
                    //img.Source = bitmap;
                    addImage = true;
                    ToolTip t = new ToolTip();
                    t.Content = imagesDir + "icono" + reader.GetString(0) + ".png";
                    img.ToolTip = t;
                    
               
                    
                 
                    img.Margin = new System.Windows.Thickness(8);
                  //  t.con
                    
                }
                catch (Exception ex)
                {
                    lab.Content = reader.GetString(0) + "\n MISSING!";
                    addImage = false;
                    ToolTip t = new ToolTip();
                    t.Content = imagesDir + "icono" + reader.GetString(0) + ".png";
                    lab.ToolTip = t;
                }
                finally
                {
                    if (addImage)
                    {
                        gridIcons.Children.Add(img);
                        Grid.SetRow(img, count / gridWidth);
                        Grid.SetColumn(img, count % gridWidth);
                    }
                    else
                    {
                        gridIcons.Children.Add(lab);
                        Grid.SetRow(lab, count / gridWidth);
                        Grid.SetColumn(lab, count % gridWidth);
                    }
                    count++;
                }
                
                //MessageBox.Show(reader.GetString(0));
            }
            reader.Close();


        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            pgBar.Value = 0;
            lblDvice.IsEnabled = false;
            lblType.IsEnabled = false;
            cbDevices.IsEnabled = false;
            cbType.IsEnabled = false;
            btnExport.IsEnabled = false;
            btnExportImages.IsEnabled = false;

            Thread exportThread = new Thread(new ParameterizedThreadStart(GenerateExportConfig));
            exportThread.Start(cbType.SelectedItem.ToString());
 


          

 
        }

        private void btnExportImages_Click(object sender, RoutedEventArgs e)
        {
            lblDvice.IsEnabled = false;
            lblType.IsEnabled = false;
            cbDevices.IsEnabled = false;
            cbType.IsEnabled = false;
            btnExport.IsEnabled = false;
            btnExportImages.IsEnabled = false;

            Thread exportThread = new Thread(new ParameterizedThreadStart(GenerateExportConfigImages));
            exportThread.Start(cbType.SelectedItem.ToString());




        }



        void InsertImage(Worksheet ws, long x, long y, long? width, long? height, string sImagePath)
        {
            try
            {
                WorksheetPart wsp = ws.WorksheetPart;
                DrawingsPart dp;
                ImagePart imgp;
                WorksheetDrawing wsd;

                ImagePartType ipt;
                switch (sImagePath.Substring(sImagePath.LastIndexOf('.') + 1).ToLower())
                {
                    case "png":
                        ipt = ImagePartType.Png;
                        break;
                    case "jpg":
                    case "jpeg":
                        ipt = ImagePartType.Jpeg;
                        break;
                    case "gif":
                        ipt = ImagePartType.Gif;
                        break;
                    default:
                        return;
                }

                if (wsp.DrawingsPart == null)
                {
                    //----- no drawing part exists, add a new one

                    dp = wsp.AddNewPart<DrawingsPart>();
                    imgp = dp.AddImagePart(ipt, wsp.GetIdOfPart(dp));
                    wsd = new WorksheetDrawing();
                }
                else
                {
                    //----- use existing drawing part

                    dp = wsp.DrawingsPart;
                    imgp = dp.AddImagePart(ipt);
                    dp.CreateRelationshipToPart(imgp);
                    wsd = dp.WorksheetDrawing;
                }

                using (FileStream fs = new FileStream(sImagePath, FileMode.Open))
                {
                    imgp.FeedData(fs);
                }

                int imageNumber = dp.ImageParts.Count<ImagePart>();
                if (imageNumber == 1)
                {
                    Drawing drawing = new Drawing();
                    drawing.Id = dp.GetIdOfPart(imgp);
                    ws.Append(drawing);
                }

                NonVisualDrawingProperties nvdp = new NonVisualDrawingProperties();
                nvdp.Id = new UInt32Value((uint)(1024 + imageNumber));
                nvdp.Name = "Picture " + imageNumber.ToString();
                nvdp.Description = "";
                DocumentFormat.OpenXml.Drawing.PictureLocks picLocks = new DocumentFormat.OpenXml.Drawing.PictureLocks();
                picLocks.NoChangeAspect = true;
                picLocks.NoChangeArrowheads = true;
                NonVisualPictureDrawingProperties nvpdp = new NonVisualPictureDrawingProperties();
                nvpdp.PictureLocks = picLocks;
                NonVisualPictureProperties nvpp = new NonVisualPictureProperties();
                nvpp.NonVisualDrawingProperties = nvdp;
                nvpp.NonVisualPictureDrawingProperties = nvpdp;

                DocumentFormat.OpenXml.Drawing.Stretch stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
                stretch.FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle();

                BlipFill blipFill = new BlipFill();
                DocumentFormat.OpenXml.Drawing.Blip blip = new DocumentFormat.OpenXml.Drawing.Blip();
                blip.Embed = dp.GetIdOfPart(imgp);
                blip.CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print;
                blipFill.Blip = blip;
                blipFill.SourceRectangle = new DocumentFormat.OpenXml.Drawing.SourceRectangle();
                blipFill.Append(stretch);

                DocumentFormat.OpenXml.Drawing.Transform2D t2d = new DocumentFormat.OpenXml.Drawing.Transform2D();
                DocumentFormat.OpenXml.Drawing.Offset offset = new DocumentFormat.OpenXml.Drawing.Offset();
                offset.X = 0;
                offset.Y = 0;
                t2d.Offset = offset;
                Bitmap bm = new Bitmap(sImagePath);

                DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();

                if (width == null)
                    extents.Cx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
                else
                    extents.Cx = width;

                if (height == null)
                    extents.Cy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
                else
                    extents.Cy = height;

                bm.Dispose();
                t2d.Extents = extents;
                ShapeProperties sp = new ShapeProperties();
                sp.BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto;
                sp.Transform2D = t2d;
                DocumentFormat.OpenXml.Drawing.PresetGeometry prstGeom = new DocumentFormat.OpenXml.Drawing.PresetGeometry();
                prstGeom.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle;
                prstGeom.AdjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
                sp.Append(prstGeom);
                sp.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

                DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();
                picture.NonVisualPictureProperties = nvpp;
                picture.BlipFill = blipFill;
                picture.ShapeProperties = sp;

                Position pos = new Position();
                pos.X = x;
                pos.Y = y;
                Extent ext = new Extent();
                ext.Cx = extents.Cx;
                ext.Cy = extents.Cy;
                AbsoluteAnchor anchor = new AbsoluteAnchor();
                anchor.Position = pos;
                anchor.Extent = ext;
                anchor.Append(picture);
                anchor.Append(new ClientData());
                wsd.Append(anchor);
                
                wsd.Save(dp);
            }
            catch (Exception ex)
            {
               // throw ex; // or do something more interesting if you want
            }
        }

        void InsertImage(Worksheet ws, long x, long y, string sImagePath)
        {
            InsertImage(ws, x, y, null, null, sImagePath);
        }
        BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
            }
        }




        private void GenerateExportConfig(object obj)
        {
            int indexRow = 1;
            int indexCols = 0;
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(AppDomain.CurrentDomain.BaseDirectory + "\\RESULTS_" + obj.ToString() + ".xlsx", SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = obj.ToString()
                };
                //var path = AppDomain.CurrentDomain.BaseDirectory + @"\RESULTS" + cbType.SelectedItem.ToString() + ".csv";
                // StreamWriter sw = new StreamWriter(path);
                foreach (string s in devicesList)
                {
                    string queryString = @"SELECT NOMBRE_ICONO, (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS WHERE IDENT = '" + s + @"' and NOMBRE_CARACT = 'LISTAICONOS') as LIST FROM VIS_LISTA_ICONOS WHERE ID_LISTA IN
                                      (SELECT ID_LISTA FROM VIS_LISTA WHERE NOMBRE_LISTA IN(
                                       SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS 
                                       WHERE IDENT = '" + s + "'"
                                     + @" and NOMBRE_CARACT = 'LISTAICONOS'))";
                    SqlCommand command = new SqlCommand(queryString, connection);
                    SqlDataReader reader = command.ExecuteReader();


                    bool printedList = false;
                    indexCols = 1;


                    Row row = new Row() { RowIndex = (uint)indexRow };
                    Cell cDev = new Cell() { CellReference = arrayColumsExcel[0] + indexRow, CellValue = new CellValue(s.ToString()), DataType = new EnumValue<CellValues>(CellValues.String) };
                    row.Append(cDev);

                    while (reader.Read())
                    {

                        if (!printedList)
                        {
                            printedList = true;
                            // writingString += "," + reader.GetString(1);
                            Cell c1 = new Cell() { CellReference = arrayColumsExcel[indexCols] + indexRow, CellValue = new CellValue(reader.GetString(1).ToString()), DataType = new EnumValue<CellValues>(CellValues.String) };
                            row.Append(c1);
                            indexCols++;
                        }
                        //writingString += "," + reader.GetString(0);

                        Cell c = new Cell() { CellReference = arrayColumsExcel[indexCols] + indexRow, CellValue = new CellValue(reader.GetString(0).ToString()), DataType = new EnumValue<CellValues>(CellValues.String) };
                        row.Append(c);




                        indexCols++;

                    }
                    sheetData.Append(row);
                    indexRow++;

                    this.Dispatcher.Invoke(
                    new UpdatePgBar(this.updatePgBar),
                    new object[] { (double.Parse(indexRow.ToString()) / double.Parse(devicesList.Count.ToString())) * 100 }
                     );

                    //sw.WriteLine(writingString) ;
                    reader.Close();
                }

                sheets.Append(sheet);

                workbookpart.Workbook.Save();

                // Close the document.
                spreadsheetDocument.Close();

                this.Dispatcher.Invoke(
                    new FinishExport(FinishThreadExport),
                    new object[] {}
                     );
            }
        }


        private void GenerateExportConfigImages(object obj)
        {
            int indexRow = 1;
            int indexCols = 0;
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(AppDomain.CurrentDomain.BaseDirectory + "\\RESULTS_" + obj.ToString() + "_ICONS.xlsx", SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = obj.ToString()
                };
                //var path = AppDomain.CurrentDomain.BaseDirectory + @"\RESULTS" + cbType.SelectedItem.ToString() + ".csv";
                // StreamWriter sw = new StreamWriter(path);
                foreach (string s in devicesList)
                {
                    string queryString = @"SELECT NOMBRE_ICONO, (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS WHERE IDENT = '" + s + @"' and NOMBRE_CARACT = 'LISTAICONOS') as LIST FROM VIS_LISTA_ICONOS WHERE ID_LISTA IN
                                      (SELECT ID_LISTA FROM VIS_LISTA WHERE NOMBRE_LISTA IN(
                                       SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS 
                                       WHERE IDENT = '" + s + "'"
                                     + @" and NOMBRE_CARACT = 'LISTAICONOS'))";
                    SqlCommand command = new SqlCommand(queryString, connection);
                    SqlDataReader reader = command.ExecuteReader();


                    bool printedList = false;
                    indexCols = 1;


                    Row row = new Row() { RowIndex = (uint)indexRow };
                    Cell cDev = new Cell() { CellReference = arrayColumsExcel[0] + indexRow, CellValue = new CellValue(s.ToString()), DataType = new EnumValue<CellValues>(CellValues.String) };
                    row.Append(cDev);

                    while (reader.Read())
                    {

                        if (!printedList)
                        {
                            printedList = true;
                            // writingString += "," + reader.GetString(1);
                            Cell c1 = new Cell() { CellReference = arrayColumsExcel[indexCols] + indexRow, CellValue = new CellValue(reader.GetString(1).ToString()), DataType = new EnumValue<CellValues>(CellValues.String) };
                            row.Append(c1);
                            indexCols++;
                        }
                        //writingString += "," + reader.GetString(0);
                        InsertImage(worksheetPart.Worksheet, (long)(indexCols * 20 * 914400 / 72) + (long)(200 * 914400) / 72, (long)(indexRow - 1) * (long)15 * (long)914400 / (long)72, imagesDir + "icono" + reader.GetString(0) + ".png");

                        // Cell c = new Cell() { CellReference = arrayColumsExcel[indexCols] + indexRow, CellValue = new CellValue(reader.GetString(0).ToString()), DataType = new EnumValue<CellValues>(CellValues.String) };
                        // row.Append(c);




                        indexCols++;
                    }
                    sheetData.Append(row);
                    indexRow++;

                    this.Dispatcher.Invoke(
                      new UpdatePgBar(this.updatePgBar),
                      new object[] { (double.Parse(indexRow.ToString()) / double.Parse(devicesList.Count.ToString())) * 100 }
                       );

                    //sw.WriteLine(writingString) ;
                    reader.Close();
                }

                sheets.Append(sheet);

                workbookpart.Workbook.Save();

                // Close the document.
                spreadsheetDocument.Close();
                this.Dispatcher.Invoke(
                    new FinishExport(FinishThreadExport),
                    new object[] { }
                     );

            }
        }





        void updatePgBar(double value)
        {
            pgBar.Value = value;
        }

        void FinishThreadExport()
        {
            //sw.Close();
            MessageBox.Show("Exported!");
            lblDvice.IsEnabled = true;
            lblType.IsEnabled = true;
            cbDevices.IsEnabled = true;
            cbType.IsEnabled = true;
            btnExport.IsEnabled = true;
            btnExportImages.IsEnabled = true;
            pgBar.Value = 0;
        }


        private void Window_MouseDown_1(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }


        private void OnMouseMoveHandler(object sender, MouseEventArgs args)
        {
            if ((sender as FrameworkElement).ToolTip == null)
                (sender as FrameworkElement).ToolTip = new ToolTip() { Placement = PlacementMode.Relative };
            double x = args.GetPosition((sender as FrameworkElement)).X;
            double y = args.GetPosition((sender as FrameworkElement)).Y;
            var tip = ((sender as FrameworkElement).ToolTip as ToolTip);
            //tip.Content = tooltip_text;
            tip.HorizontalOffset = x + 10;
            tip.VerticalOffset = y + 10;
        }
    }




}
