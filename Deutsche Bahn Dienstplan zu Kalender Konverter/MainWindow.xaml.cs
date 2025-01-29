using System.ComponentModel;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;

namespace Deutsche_Bahn_Dienstplan_zu_Kalender_Konverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private string _resultFileName = "MeinCalendar";
        public string ResultFileName
        {
            get { return _resultFileName; }
            set
            {
                _resultFileName = value;
                OnPropertyChanged(nameof(ResultFileName));
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this; // Set DataContext to the current instance
        }
        private void PickFile_Click(object sender, RoutedEventArgs e)
        {
            // OpenFileDialog for picking an Excel file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;

                try
                {
                    // Load the Excel file using EPPlus
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        // Assume the data is in the first worksheet
                        var worksheet = package.Workbook.Worksheets[0];

                        // Generate the calendar content based on the Excel sheet
                        string calendarContent = GenerateCalendarFromExcel(worksheet);

                        // Open SaveFileDialog to choose location for saving .ics file
                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        saveFileDialog.Filter = "iCalendar Files (*.ics)|*.ics|All Files (*.*)|*.*";
                        saveFileDialog.DefaultExt = ".ics";
                        saveFileDialog.FileName = ResultFileName + ".ics";

                        // If the user selects a location and clicks "Save"
                        if (saveFileDialog.ShowDialog() == true)
                        {
                            string calendarFilePath = saveFileDialog.FileName;

                            // Write the calendar content to the selected file
                            File.WriteAllText(calendarFilePath, calendarContent);
                            MessageBox.Show($"Kalender-Datei erfolgreich gespeichert unter: {calendarFilePath}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Fehler: {ex.Message}");
                }
            }
        }
        private string GenerateCalendarFromExcel(ExcelWorksheet worksheet)
        {
            var calendarContent = "BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\n";

            int totalRows = worksheet.Dimension.Rows;
            int totalColumns = worksheet.Dimension.Columns;

            // Process each "super-row" (4 rows per month)
            for (int row = 3; row <= totalRows; row += 4)
            {
                // 1st sub-row: Month and days
                string monthInfo = worksheet.Cells[row, 1].Text; // e.g., "01,2025"
                if (DateTime.TryParseExact(monthInfo, "MM,yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime monthStart))
                {
                    // Loop through the days in this month (columns 2 to totalColumns)
                    for (int col = 2; col <= totalColumns; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        string dayText = cell.Text ?? string.Empty;
                        // If it's a valid day, process it
                        if (int.TryParse(dayText, out int day))
                        {
                            DateTime eventStartDate = new DateTime(monthStart.Year, monthStart.Month, day);

                            // 2nd sub-row: Shift Type
                            string shiftType = worksheet.Cells[row + 1, col].Text;
                            if (!string.IsNullOrEmpty(shiftType) && !isShiftTypeFree(shiftType)) // Handle normal shift types
                            {
                                DateTime eventEndDate = eventStartDate.AddDays(1);
                                calendarContent += CreateFullDayEvent(eventStartDate, eventEndDate, shiftType);
                            }
                        }
                    }
                }
            }

            calendarContent += "END:VCALENDAR";
            return calendarContent;
        }

        private bool isShiftTypeFree(string shiftType)
        {
            return int.TryParse(shiftType, out int _);
        }


        private string CreateFullDayEvent(DateTime startDate, DateTime endDate, string summary)
        {
            return $"BEGIN:VEVENT\n" +
                   $"SUMMARY:{summary}\n" +
                   $"DTSTART;VALUE=DATE:{startDate:yyyyMMdd}\n" +
                   $"DTEND;VALUE=DATE:{endDate:yyyyMMdd}\n" +
                   "END:VEVENT\n";
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

