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
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
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
        public delegate void UpdateGUI();
        public delegate void FinishExport();
        string[] arrayColumsExcel = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };


        System.Timers.Timer timerClipboard;
        List<String> typesList;
        List<String> typesConnectionList;
        List<String> devicesList;
        List<String> protocolList;
        SqlConnection connection;
        string imagesDir;
        public MainWindow()
        {

            InitializeComponent();
            timerClipboard = new System.Timers.Timer();
            timerClipboard.Interval = 2000;
            timerClipboard.Elapsed += new ElapsedEventHandler(TimerElapsed);
            timerClipboard.Stop();
            this.WindowStyle = WindowStyle.None;
        }


        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("Loaded");
            typesList = new List<string>();
            devicesList = new List<string>();
            protocolList = new List<string>();
            typesConnectionList = new List<string>();
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
                btnCompare.IsEnabled = enable;
                cbTypeConnections.IsEnabled = enable;
                cbProtocol.IsEnabled = enable;
                lblProtocol.IsEnabled = enable;
                lblTypeConnections.IsEnabled = enable;

                if (enable)
                {
                    LoadTypes();
                    LoadProtocols();
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

        private void LoadProtocols()
        {
            string queryString = "SELECT DISTINCT PROTOCOLO FROM CONF_EQUIPOS WHERE PROTOCOLO NOT IN('ES','C2C', 'VIRTUAL','DGT') AND NOMBRE_TIPOEQUIPO NOT IN ('CANAL','WGT_SC','CML_SPC_CNTLR') AND NOMBRE_TIPOEQUIPO NOT LIKE '%GROUP%' ORDER BY PROTOCOLO";
            SqlCommand command = new SqlCommand(queryString, connection);
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                protocolList.Add(reader.GetString(0));
                //MessageBox.Show(reader.GetString(0));
            }
            reader.Close();
            cbProtocol.ItemsSource = protocolList;
            cbProtocol.SelectionChanged += LoadTypesConnectionsTab;
            cbProtocol.SelectedIndex = 0;
        }

        private void LoadTypesConnectionsTab(Object source, EventArgs e)
        {
            cbTypeConnections.SelectionChanged -= LoadConnections;
            cbTypeConnections.ItemsSource = null;
            string queryString = "SELECT DISTINCT NOMBRE_TIPOEQUIPO FROM CONF_EQUIPOS WHERE PROTOCOLO ='"+ cbProtocol.SelectedItem.ToString() +"' AND NOMBRE_TIPOEQUIPO NOT IN ('CANAL','WGT_SC','CML_SPC_CNTLR') AND NOMBRE_TIPOEQUIPO NOT LIKE '%GROUP%'";
            SqlCommand command = new SqlCommand(queryString, connection);
            SqlDataReader reader = command.ExecuteReader();
            typesConnectionList.Clear();
            while (reader.Read())
            {
                typesConnectionList.Add(reader.GetString(0));
                //MessageBox.Show(reader.GetString(0));
            }
            reader.Close();
            cbTypeConnections.ItemsSource = typesConnectionList;
            cbTypeConnections.SelectionChanged += LoadConnections;
            cbTypeConnections.SelectedIndex = 0;
        }

        private void LoadConnections(Object source, EventArgs e)
        {
            ClearGrid();
            switch (cbProtocol.SelectedItem.ToString())
            {
                case "GENETEC":
                    LoadGenetecConnections();
                    break;

                case "TSI_SP_003":
                    LoadRTAConnections();
                    break;

                default:
                    LoadIPChannelConnections();
                    break;
            }
        }

        private void LoadIPChannelConnections()
        {
            List<IPChannelConnection> ipConnections = new List<IPChannelConnection>();
            string queryString = @" SELECT DISTINCT VCE.IDENT AS DEVICE, 
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS WHERE IDENT = (SELECT top(1) CS.IDENT AS CHANNEL FROM CONF_SUBEQUIPOS CS, CONF_EQUIPOS CE WHERE CS.IDENT = CE.IDENT AND CE.NOMBRE_TIPOEQUIPO = 'CANAL' AND CS.IDENT1 = VCE.IDENT) AND NOMBRE_CARACT = 'IP1') AS IP_CHANNEL,
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS WHERE IDENT = (SELECT top(1) CS.IDENT AS CHANNEL FROM CONF_SUBEQUIPOS CS, CONF_EQUIPOS CE WHERE CS.IDENT = CE.IDENT AND CE.NOMBRE_TIPOEQUIPO = 'CANAL' AND CS.IDENT1 = VCE.IDENT) AND NOMBRE_CARACT = 'PUERTO1') AS PORT_CHANNEL
            FROM VIEW_CARACTERISTICAS_EQUIPOS VCE, CONF_TIPOS_EQUIPOS CTE, CONF_EQUIPOS CE
            WHERE VCE.NOMBRE_TIPOEQUIPO = CTE.NOMBRE_TIPOEQUIPO
            AND CE.IDENT = VCE.IDENT
            AND VCE.IDENT IS NOT NULL
            AND VCE.NOMBRE_TIPOEQUIPO = '" + cbTypeConnections.SelectedItem.ToString() + "' AND CE.PROTOCOLO = '"+ cbProtocol.SelectedItem.ToString() + "'";

            SqlCommand command = new SqlCommand(queryString, connection);
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                IPChannelConnection ipCon = new IPChannelConnection();
                if (!reader.IsDBNull(0))
                    ipCon.Device = reader.GetString(0);
                if (!reader.IsDBNull(1))
                    ipCon.IP = reader.GetString(1);
                if (!reader.IsDBNull(2))
                    ipCon.Port = reader.GetString(2);

                ipConnections.Add(ipCon);
            }
            reader.Close();

            dgConnections.Columns[0].Width = new DataGridLength(3, DataGridLengthUnitType.Star);
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "IP",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding("IP")
            });
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "Port",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding("Port")
            });

            dgConnections.AutoGenerateColumns = false;
            dgConnections.ItemsSource = ipConnections;
        }

        private void LoadRTAConnections()
        {

            List<RTAConnection> RTAConnections = new List<RTAConnection>();
            //NON VMS_C devices

            string queryString = @"SELECT DISTINCT VCE.IDENT AS DEVICE, 
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS CAR WHERE CAR.IDENT = VCE.IDENT AND NOMBRE_CARACT = 'NUMLINEAS') AS NUM_LINES,
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS CAR WHERE CAR.IDENT = VCE.IDENT AND NOMBRE_CARACT = 'NUMCARACTERES') AS NUM_CHARS,
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS CAR WHERE CAR.IDENT = VCE.IDENT AND NOMBRE_CARACT = 'BEACON_TYPE') AS BEACON_TYPE,
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS CAR WHERE CAR.IDENT = VCE.IDENT AND NOMBRE_CARACT = 'GROUP_ID') AS GROUP_ID,
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS CAR WHERE CAR.IDENT = VCE.IDENT AND NOMBRE_CARACT = 'MULTIPUNTO') AS ADRESS_ID,
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS CAR WHERE CAR.IDENT = VCE.IDENT AND NOMBRE_CARACT = 'SIGN_ID') AS SIGN_ID,
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS WHERE IDENT = (SELECT top(1) CS.IDENT AS CHANNEL FROM CONF_SUBEQUIPOS CS, CONF_EQUIPOS CE WHERE CS.IDENT = CE.IDENT AND CE.NOMBRE_TIPOEQUIPO = 'CANAL' AND CS.IDENT1 = VCE.IDENT) AND NOMBRE_CARACT = 'IP1') AS IP_CHANNEL,
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS WHERE IDENT = (SELECT top(1) CS.IDENT AS CHANNEL FROM CONF_SUBEQUIPOS CS, CONF_EQUIPOS CE WHERE CS.IDENT = CE.IDENT AND CE.NOMBRE_TIPOEQUIPO = 'CANAL' AND CS.IDENT1 = VCE.IDENT) AND NOMBRE_CARACT = 'PUERTO1') AS PORT_CHANNEL,
            (SELECT VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS WHERE IDENT = (SELECT top(1) CS.IDENT AS CHANNEL FROM CONF_SUBEQUIPOS CS, CONF_EQUIPOS CE WHERE CS.IDENT = CE.IDENT AND CE.NOMBRE_TIPOEQUIPO = 'CANAL' AND CS.IDENT1 = VCE.IDENT) AND NOMBRE_CARACT = 'PASSWORD') AS PASSWORD_RTA
            FROM VIEW_CARACTERISTICAS_EQUIPOS VCE, CONF_TIPOS_EQUIPOS CTE
            WHERE VCE.NOMBRE_TIPOEQUIPO = CTE.NOMBRE_TIPOEQUIPO
            AND CTE.ID_TEQUIP_PRINCIPAL = 'CPMV'
            AND VCE.IDENT IS NOT NULL
            AND VCE.NOMBRE_TIPOEQUIPO = '" + cbTypeConnections.SelectedItem.ToString()+"'";

            SqlCommand command = new SqlCommand(queryString, connection);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                RTAConnection rtaCon = new RTAConnection();
                if(!reader.IsDBNull(0))
                    rtaCon.Device = reader.GetString(0);
                if (!reader.IsDBNull(1))
                    rtaCon.NumLines = reader.GetString(1);
                if (!reader.IsDBNull(2))
                    rtaCon.NumCharacters = reader.GetString(2);
                if (!reader.IsDBNull(3))
                    rtaCon.Beacon_Type = reader.GetString(3);
                if (!reader.IsDBNull(4))
                    rtaCon.GroupId = reader.GetString(4);
                if (!reader.IsDBNull(5))
                    rtaCon.AdressId = reader.GetString(5);
                if (!reader.IsDBNull(6))
                    rtaCon.SignID = reader.GetString(6);
                if (!reader.IsDBNull(7))
                    rtaCon.IP = reader.GetString(7);
                if (!reader.IsDBNull(8))
                    rtaCon.Port = reader.GetString(8);
                if (!reader.IsDBNull(9))
                    rtaCon.RTAPass = reader.GetString(9);

                if (string.IsNullOrEmpty(rtaCon.Beacon_Type) || rtaCon.Beacon_Type == "none")
                    rtaCon.Beacon_Type = "N";
                else
                    rtaCon.Beacon_Type = "Y";
                RTAConnections.Add(rtaCon);
            }
            reader.Close();
            //ClearGrid();

            dgConnections.Columns[0].Width = new DataGridLength(2, DataGridLengthUnitType.Star);
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "Number of lines",
                Width = new DataGridLength(0.3, DataGridLengthUnitType.Star),
                Binding = new Binding("NumLines")
            });
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "Number of chars",
                Width = new DataGridLength(0.3, DataGridLengthUnitType.Star),
                Binding = new Binding("NumCharacters")
            });
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "Beacons",
                Width = new DataGridLength(0.3, DataGridLengthUnitType.Star),
                Binding = new Binding("Beacon_Type")
            });
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "Group ID",
                Width = new DataGridLength(0.3, DataGridLengthUnitType.Star),
                Binding = new Binding("GroupId")
            });
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "Adress ID",
                Width = new DataGridLength(0.3, DataGridLengthUnitType.Star),
                Binding = new Binding("AdressId")
            });
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "Sign ID",
                Width = new DataGridLength(0.3, DataGridLengthUnitType.Star),
                Binding = new Binding("SignID")
            });
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "IP",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                Binding = new Binding("IP")
            });
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "Port",
                Width = new DataGridLength(0.6, DataGridLengthUnitType.Star),
                Binding = new Binding("Port")
            });
            dgConnections.Columns.Add(new DataGridTextColumn
            {
                Header = "Password",
                Width = new DataGridLength(0.6, DataGridLengthUnitType.Star),
                Binding = new Binding("RTAPass")
            });

            dgConnections.AutoGenerateColumns = false;
            dgConnections.ItemsSource = RTAConnections;
        }

        private void LoadGenetecConnections()
        {

            List<GenetecConnection> genetecConnections = new List<GenetecConnection>();
            String queryString = "SELECT IDENT,NOMBRE_CARACT,VALOR FROM VIEW_CARACTERISTICAS_EQUIPOS WHERE NOMBRE_TIPOEQUIPO = '" + cbTypeConnections.SelectedItem.ToString() + "' AND NOMBRE_CARACT IN('GENETEC_ID','CITILOG_ID')";
            SqlCommand command = new SqlCommand(queryString, connection);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                GenetecConnection gen = new GenetecConnection();
                gen.Device = reader.GetString(0);
                if (reader.GetString(1).Equals("GENETEC_ID"))
                    gen.GenID = reader.GetString(2);
                else if(reader.GetString(1).Equals("CITILOG_ID"))
                    gen.CitiID = reader.GetString(2);
                genetecConnections.Add(gen);
            }
            reader.Close();

            if (genetecConnections.Find(x => !string.IsNullOrEmpty(x.GenID)) != null)
            {

                dgConnections.Columns.Add(new DataGridTextColumn
                {
                    Header = "Genetec ID", 
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                    Binding = new Binding("GenID")
                }) ;
            }
            if (genetecConnections.Find(x => !string.IsNullOrEmpty(x.CitiID)) != null)
            {
                dgConnections.Columns.Add(new DataGridTextColumn
                {
                    Header = "Citilog ID",
                    Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                    Binding = new Binding("CitiID")
                });
            }

            dgConnections.ItemsSource = genetecConnections;

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
                    t.Placement = PlacementMode.Relative;
                    img.ToolTip = t;
                    img.MouseRightButtonDown += new MouseButtonEventHandler(SetClipBoardString);
                    img.MouseMove += new MouseEventHandler(OnMouseMoveHandler);
               
                    
                 
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
                    lab.MouseRightButtonDown += new MouseButtonEventHandler(SetClipBoardString);
                    lab.MouseMove += new MouseEventHandler(OnMouseMoveHandler);
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
            btnCompare.IsEnabled = false;

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
            btnCompare.IsEnabled = false;

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
                        InsertImage(worksheetPart.Worksheet, (long)(indexCols * 20 * 914400 / 72) + (long)(200 * 914400) / 72, (long)(indexRow - 1) * (long)15 * (long)9144*100 / (long)72, imagesDir + "icono" + reader.GetString(0) + ".png");

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

            CustomModalWindow cmw = new CustomModalWindow("SICE", "Excel exported.");
            cmw.ShowDialog();
            lblDvice.IsEnabled = true;
            lblType.IsEnabled = true;
            cbDevices.IsEnabled = true;
            cbType.IsEnabled = true;
            btnExport.IsEnabled = true;
            btnExportImages.IsEnabled = true;
            btnCompare.IsEnabled = true;
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
            
            tip.HorizontalOffset = x+15;
            tip.VerticalOffset = y;


        }


        private void SetClipBoardString(object sender, MouseButtonEventArgs args)
        {

            Clipboard.SetText((sender as FrameworkElement).ToolTip.ToString().Remove(0, ("System.Windows.Controls.ToolTip: ").ToString().Length));
            txtLog.Content = "Copied to Clipboard";
            timerClipboard.Start();
        }

        private void ClearLog()
        {
            txtLog.Content = "";
        }
        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            
            this.Dispatcher.Invoke(
                      new UpdateGUI(this.ClearLog)
                       );
            timerClipboard.Stop();
        }

        private void ClearGrid()
        {
            dgConnections.Columns[0].Width = new DataGridLength(1, DataGridLengthUnitType.Star);
            for (int i = dgConnections.Columns.Count-1; i > 0; i--)
            {

                dgConnections.Columns.Remove(dgConnections.Columns[i]);
            }
        }


        private void btnCompare_Click(object sender, RoutedEventArgs e)
        {
            ComparisonIcons comp = new ComparisonIcons();
            comp.ShowDialog();
        }

    }


    public class ConnectionRow
    {
        public string Device { get; set; }
        public string Param1 { get; set; }
        public string Param2 { get; set; }
        public string Param3 { get; set; }
        public string Param4 { get; set; }
        public string Param5 { get; set; }
        public string Param6 { get; set; }

        public ConnectionRow(string dev, string p1, string p2, string p3, string p4, string p5, string p6)
        {
            Device = dev;
            Param1 = p1;
            Param2 = p2;
            Param3 = p3;
            Param4 = p4;
            Param5 = p5;
            Param6 = p6;
        }
    }

    public class GenetecConnection
    {
        public string Device { get; set; }
        public string GenID { get; set; }
        public string CitiID { get; set; }
    }


    public class RTAConnection
    {
        public string Device { get; set; }
        public string NumLines { get; set; }
        public string NumCharacters { get; set; }
        public string Beacon_Type { get; set; }
        public string GroupId { get; set; }
        public string AdressId { get; set; }
        public string SignID { get; set; }
        public string IP { get; set; }
        public string Port { get; set; }
        public string RTAPass { get; set; }
    }

    public class IPChannelConnection
    {
        public string Device { get; set; }
        public string IP { get; set; }
        public string Port { get; set; }
    }

}
