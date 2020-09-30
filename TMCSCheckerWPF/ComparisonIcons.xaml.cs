using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace TMCSCheckerWPF
{
    /// <summary>
    /// Interaction logic for ComparisonIcons.xaml
    /// </summary>
    public partial class ComparisonIcons : Window
    {
        public delegate void UpdatePgBar(double value);
        public delegate void FinishReading();
        private List<DeviceIconGroup> documentationImportList;
        private List<DeviceIconGroup> databaseImportList;

        public ComparisonIcons()
        {
            InitializeComponent();
            documentationImportList = new List<DeviceIconGroup>();
            databaseImportList = new List<DeviceIconGroup>();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_MouseDown_1(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void ImportDatabaseItems(object sender, RoutedEventArgs e)
        {
            DisableWindowOptions();
            //choose file and load them into the db list
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 0;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Title = "Select the database export: ";
            if (openFileDialog.ShowDialog() == true)
            {
                string fullNameOfFile = openFileDialog.FileName;
                textDBTitle.Text = fullNameOfFile.Split('\\').Last();
                Thread exportThread = new Thread(new ParameterizedThreadStart(ReadExcelFileDatabase));
                exportThread.Start(fullNameOfFile);
            }
            else
                EnableWindowOptions();
        }

        private void ImportDocumentationItems(object sender, RoutedEventArgs e)
        {
            DisableWindowOptions();
            //choose file and load them into the doc list
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 0;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Title = "Select the documentation export: ";
            if (openFileDialog.ShowDialog() == true)
            {
                String tabTitle = "";
                string fullNameOfFile = openFileDialog.FileName;
                textDocTitle.Text = fullNameOfFile.Split('\\').Last();
                InputModalWindow imw = new InputModalWindow("Select tab");
                imw.ShowDialog();
                tabTitle = imw.ReturnValue;
                ExcelObject eo = new ExcelObject() {path = fullNameOfFile, tab = tabTitle };
                Thread exportThread = new Thread(new ParameterizedThreadStart(ReadExcelFileDocument));
                exportThread.Start(eo);
                
            }
            else
                EnableWindowOptions();


        }


        private void ReadExcelFileDatabase(object fileName)
        {
            databaseImportList.Clear();
            int x, y;
            x = 0;
            y = 0;


            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName.ToString(), false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                SharedStringTablePart sstpart = null;
                SharedStringTable sst = null;
                if(workbookPart.GetPartsOfType<SharedStringTablePart>().Count() >0)
                {
                    sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    sst = sstpart.SharedStringTable;
                }
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();


                Worksheet sheet = worksheetPart.Worksheet;

                var cells = sheet.Descendants<Cell>();
                var rows = sheet.Descendants<Row>();

                Console.WriteLine("Row count = {0}", rows.LongCount());
                Console.WriteLine("Cell count = {0}", cells.LongCount());


                foreach (Row row in rows)
                {
                    DeviceIconGroup dig = new DeviceIconGroup();
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        if (y == 1)
                        {
                            y++;
                            continue;
                        }

                        if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(c.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            if (y == 0)
                                dig.deviceName = str;
                            else
                                dig.iconNames.Add(str);

                        }
                        else if (c.CellValue != null && c.DataType == CellValues.String)
                        {
                           if (y == 0)
                                dig.deviceName = c.CellValue.Text;
                           else
                                dig.iconNames.Add(c.CellValue.Text);
                        }
                        y++;
                    }
                    this.Dispatcher.Invoke(
                       new UpdatePgBar(this.updatePgBar),
                       new object[] { (double.Parse(x.ToString()) / double.Parse(rows.Count().ToString())) * 100 }
                    );
                    databaseImportList.Add(dig);

                    x++;
                    y = 0;
                }

                // Close the document.
                doc.Close();
            }

            this.Dispatcher.Invoke(
                       new FinishReading(EnableWindowOptions),
                       new object[] { }
                    );


        }


        private void ReadExcelFileDocument(object excelObj)
        {
            documentationImportList.Clear();
            int x, y;
            x = 0;
            y = 0;

            string filename = ((ExcelObject)excelObj).path;
            string tab = ((ExcelObject)excelObj).tab;

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filename, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                SharedStringTablePart sstpart = null;
                SharedStringTable sst = null;
                if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    sst = sstpart.SharedStringTable;
                }
                List<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>().ToList();
                List<WorksheetPart> sheetParts = workbookPart.WorksheetParts.ToList();
                int indexOfTab = workbookPart.Workbook.Descendants<Sheet>().ToList().IndexOf(workbookPart.Workbook.Descendants<Sheet>().ToList().Find(obj => obj.Name.Equals(tab)));
                WorksheetPart worksheetPart = GetWorksheetPart(workbookPart,tab);
              //  WorksheetPart worksheetPart = workbookPart.WorksheetParts.ToList()[indexOfTab];

                //workbookPart.WorksheetParts.ToList().Find(x=>x.Worksheet.XName)
                //string sheetName = workbookPart.Workbook.Descendants<Sheet>().ToList().IndexOf();


                //workbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex).Name

                Worksheet sheet = worksheetPart.Worksheet;

                var cells = sheet.Descendants<Cell>();
                var rows = sheet.Descendants<Row>();

                Console.WriteLine("Row count = {0}", rows.LongCount());
                Console.WriteLine("Cell count = {0}", cells.LongCount());


                foreach (Row row in rows)
                {
                    if (x < 2)
                    {
                        x++;
                        continue;
                    }
                    DeviceIconGroup dig = new DeviceIconGroup();
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        if (y == 1)
                        {
                            y++;
                            continue;
                        }

                        if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(c.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            if (y == 0)
                                dig.deviceName = str;
                            else
                                dig.iconNames.Add(str);

                        }
                        else if (c.CellValue != null && c.DataType == CellValues.String)
                        {
                            if (y == 0)
                                dig.deviceName = c.CellValue.Text;
                            else
                                dig.iconNames.Add(c.CellValue.Text);
                        }
                        y++;
                    }
                    this.Dispatcher.Invoke(
                       new UpdatePgBar(this.updatePgBar),
                       new object[] { (double.Parse(x.ToString()) / double.Parse(rows.Count().ToString())) * 100 }
                    );
                    documentationImportList.Add(dig);

                    x++;
                    y = 0;
                }

                // Close the document.
                doc.Close();
            }

            this.Dispatcher.Invoke(
                       new FinishReading(EnableWindowOptions),
                       new object[] { }
                    );


        }


        private void DisableWindowOptions()
        {
            btnLoadDB.IsEnabled = false;
            btnLoadDoc.IsEnabled = false;
            btnCompare.IsEnabled = false;
            btnExport.IsEnabled = false;
        }

        private void EnableWindowOptions()
        {
            btnLoadDB.IsEnabled = true;
            btnLoadDoc.IsEnabled = true;
            pgBar.Value = 0;
            if (documentationImportList.Count > 0 && databaseImportList.Count > 0)
            {
                btnCompare.IsEnabled = true;
                btnExport.IsEnabled = true;
            }
            else
            {
                btnCompare.IsEnabled = false;
                btnExport.IsEnabled = false;
            }
        }
        void updatePgBar(double value)
        {
            pgBar.Value = value;
        }

        private WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
        }

        private void btnCompare_Click(object sender, RoutedEventArgs e)
        {
            List<ComparedObject> objectComparisonList = new List<ComparedObject>();
            for (int i = 0; i < documentationImportList.Count; i++)
            {
                DeviceIconGroup databaseItem = databaseImportList.Find(x=>x.deviceName.Equals(documentationImportList[i].deviceName));
                ComparedObject co = new ComparedObject() { DeviceDoc = documentationImportList[i].deviceName, DeviceDB = databaseItem.deviceName,Result=documentationImportList[i].IsEqualTo(databaseItem) };
                objectComparisonList.Add(co);
            }
            dgConnections.ItemsSource = objectComparisonList;
        }
    }



    public class ExcelObject
    {
        public string path;
        public string tab;
    }

    public class ComparedObject
    {
      
        public string DeviceDoc { get; set; }
        public string DeviceDB { get; set; }
        public bool Result { get; set; }
    }

    public class DeviceIconGroup 
    {
        public string deviceName;
        public List<string> iconNames;

        public DeviceIconGroup()
        {
            iconNames = new List<string>();
        }

        public bool IsEqualTo(DeviceIconGroup other)
        {
            if (deviceName.Equals(other.deviceName) && (iconNames.Count == other.iconNames.Count))
            {
                for (int i = 0; i < iconNames.Count; i++)
                {
                    if (!iconNames[i].Equals(other.iconNames[i]))
                        return false;
                }
                return true;
            }
            else
                return false;
        }
    }
}
