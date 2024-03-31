using Microsoft.Win32;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System;
using Coordinates;
using OfficeOpenXml;
using System.IO;
using System.ComponentModel;
using System.Collections.ObjectModel;

namespace MishConvertor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private List<double> m_latList;
        private List<double> m_lonList;
        private Dictionary<int, SingleDotPosition> m_dots;
        private ObservableCollection<SingleDotPosition> m_dotsList;
        public event PropertyChangedEventHandler? PropertyChanged;

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            DataContext = this;

            MainTabControl.SelectionChanged += MainTabControl_SelectionChanged;
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MainTabControl.SelectedItem != null && MainTabControl.SelectedItem == ViewFile) 
            {
                if (m_dots != null)
                {
                    m_dotsList = new ObservableCollection<SingleDotPosition>(m_dots.Values);
                    Dots = m_dotsList;
                }
            }
        }

        private void UploadExcelButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                MessageBox.Show($"Opening file: {filePath}");

                if (TakeDataFromFile(filePath))
                {
                    SaveDataToDictionary();
                }
            }
        }

        private void SaveDataToDictionary()
        {
            m_dots = new Dictionary<int, SingleDotPosition>();

            for (int i = 0; i < m_latList.Count; i++)
            {
                m_dots.Add(i + 1, new SingleDotPosition(m_latList[i], m_lonList[i]));
            }
        }

        private void ExportNewFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;

                using (ExcelPackage excelPackage = new ExcelPackage()) 
                {
                    ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    ws.Cells[1, 1].Value = "Latitude";
                    ws.Cells[1, 2].Value = "Longitude";
                    ws.Cells[1, 3].Value = "MitEast";
                    ws.Cells[1, 4].Value = "MitNorth";

                    int row = 2;
                    foreach (var kvp in m_dots)
                    {
                        SingleDotPosition dot = kvp.Value;
                        ws.Cells[row, 1].Value = dot.Lat;
                        ws.Cells[row, 2].Value = dot.Longitude;
                        ws.Cells[row, 3].Value = dot.MitEast;
                        ws.Cells[row, 4].Value = dot.MitNorth;

                        ++row;
                    }

                    excelPackage.SaveAs(new FileInfo(filePath));
                }
            }
        }

        private bool TakeDataFromFile(string filePath)
        {
            m_latList = new List<double>();
            m_lonList = new List<double>();

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets[0];

                    for (int row = 1; row <= ws.Dimension.Rows; ++row)
                    {
                        for (int col = 1; col <= ws.Dimension.Columns; ++col)
                        {
                            object cellVal = ws.Cells[row, col].Value;
                            if (col == 1)
                            {
                                if (cellVal != null && double.TryParse(cellVal.ToString(), out double dblVal))
                                {
                                    m_latList.Add(dblVal);
                                }
                            }
                            else if (col == 2)
                            {
                                if (cellVal != null && double.TryParse(cellVal.ToString(), out double dblVal))
                                {
                                    m_lonList.Add(dblVal);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }

            return true;
        }

        private void ConvertToITM_Click(object sender, EventArgs e)
        {
            var lat = LatitudeReport;
            var lon = LongitudeReport;
            var dot = new SingleDotPosition(lat, lon);
            ITMEastReport = dot.MitEast;
            ITMNorthReport = dot.MitNorth;
        }

        private void ConvertToLanLon_Click(object sender, EventArgs e)
        {
            int itmEast = ITMEastReport2;
            int itmNorth = ITMNorthReport2;
            var dot = new SingleDotPosition(itmNorth, itmEast);
            LatitudeReport2 = dot.Lat;
            LongitudeReport2 = dot.Longitude;
        }

        #region Properties

        public string AppTitle => $"MishConvertor v{System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}";

        public ObservableCollection<SingleDotPosition> Dots
        {
            get { return m_dotsList; }
            set 
            { 
                m_dotsList = value;
                OnPropertyChanged("Dots");
            }
        }

        private double m_latitudeReport;

        public double LatitudeReport
        {
            get { return m_latitudeReport; }
            set
            {
                m_latitudeReport = value;
                OnPropertyChanged("LatitudeReport");
            }
        }

        private double m_longitudeReport;
        public double LongitudeReport
        {
            get { return m_longitudeReport; }
            set
            {
                m_longitudeReport = value;
                OnPropertyChanged("LongitudeReport");
            }
        }

        private int m_iTMEastReport;
        public int ITMEastReport
        {
            get { return m_iTMEastReport; }
            set
            {
                m_iTMEastReport = value;
                OnPropertyChanged("ITMEastReport");
            }
        }

        private int m_iTMNorthReport;
        public int ITMNorthReport
        {
            get { return m_iTMNorthReport; }
            set
            {
                m_iTMNorthReport = value;
                OnPropertyChanged("ITMNorthReport");
            }
        }

        private double m_latitudeReport2;

        public double LatitudeReport2
        {
            get { return m_latitudeReport2; }
            set
            {
                m_latitudeReport2 = value;
                OnPropertyChanged("LatitudeReport2");
            }
        }

        private double m_longitudeReport2;
        public double LongitudeReport2
        {
            get { return m_longitudeReport2; }
            set
            {
                m_longitudeReport2 = value;
                OnPropertyChanged("LongitudeReport2");
            }
        }

        private int m_iTMEastReport2;
        public int ITMEastReport2
        {
            get { return m_iTMEastReport2; }
            set
            {
                m_iTMEastReport2 = value;
                OnPropertyChanged("ITMEastReport2");
            }
        }

        private int m_iTMNorthReport2;
        public int ITMNorthReport2
        {
            get { return m_iTMNorthReport2; }
            set
            {
                m_iTMNorthReport2 = value;
                OnPropertyChanged("ITMNorthReport2");
            }
        }

        #endregion

        #region Others

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}