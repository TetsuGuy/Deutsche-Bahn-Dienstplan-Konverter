using System.ComponentModel;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;
using Spire.Pdf;

namespace Deutsche_Bahn_Dienstplan_zu_Kalender_Konverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this; // Set DataContext to the current instance
        }

        private void PickFile_Click_PDF(object sender, RoutedEventArgs e)
        {
            // OpenFileDialog for picking a PDF file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    MessageBox.Show("Die ausgewählte Datei existiert nicht.");
                    return;
                }
                LoadTable(GenerateXlsxFile(filePath));
            }
        }

        private string GenerateXlsxFile(string filePath)
        {
            string xlsxFilePath = Path.ChangeExtension(filePath, ".xlsx");
            PdfDocument pdf = new PdfDocument();
            pdf.LoadFromFile(filePath);
            pdf.SaveToFile(xlsxFilePath, FileFormat.XLSX);
            return xlsxFilePath;
        }


        private void LoadTable(string filePathXLSX)
        {
            // Load the Excel file using EPPlus
            using (var package = new ExcelPackage(new FileInfo(filePathXLSX)))
            {
                // Assume the data is in the first worksheet
                var worksheet = package.Workbook.Worksheets[0];

                // Generate the calendar content based on the Excel sheet
                string calendarContent = GenerateCalendarFromExcel(worksheet);

                // Open SaveFileDialog to choose location for saving .ics file
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "iCalendar Files (*.ics)|*.ics|All Files (*.*)|*.*";
                saveFileDialog.DefaultExt = ".ics";
                saveFileDialog.FileName = "Schichtplan.ics";

                // If the user selects a location and clicks "Save"
                if (saveFileDialog.ShowDialog() == true)
                {
                    string calendarFilePath = saveFileDialog.FileName;

                    // Write the calendar content to the selected file
                    File.WriteAllText(calendarFilePath, calendarContent);
                    File.Delete(filePathXLSX); // remove the temporary Excel file
                    MessageBox.Show($"Kalender-Datei erfolgreich gespeichert unter: {calendarFilePath}.\nBitte kontrolliere die erzeugten Termine nun auf Richtigkeit.");
                }
            }
        }

        private (int startRow, int startColumn, int totalRows, int totalColumns) FindStartIndices(ExcelWorksheet worksheet)
        {
            int totalRows = worksheet.Dimension.Rows;
            int totalColumns = worksheet.Dimension.Columns;
            // find start row
            int startRow = 1;
            int startColumn = 1;
            bool foundStartCell = false;
            // regex: start cell should be format "MM,yyyy" or "MM.yyyy"
            var regex = new System.Text.RegularExpressions.Regex(@"^\d{2}[,.]\d{4}$");
            for (int row = 1; row <= totalRows; row++)
            {
                if (foundStartCell) break;
                for (int column = 1; column <= totalColumns; column++)
                {
                    if (foundStartCell) break;
                    string cellText = worksheet.Cells[row, column].Text;
                    if (regex.IsMatch(cellText.Trim()))
                    {
                        startRow = row;
                        startColumn = column;
                        foundStartCell = true;
                    }
                }
            }
            if (!foundStartCell)
            {
                throw new System.Exception("Startzelle nicht gefunden. Bitte überprüfe das Format der Excel-Datei.");
            }
            return (startRow, startColumn, totalRows, totalColumns);
        }
        
        private string GenerateCalendarFromExcel(ExcelWorksheet worksheet)
        {
            var calendarContent = "BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\n";
            var (startRow, startColumn, totalRows, totalColumns) = FindStartIndices(worksheet);

            // Process each "super-row" (4 rows per month)
            for (int row = startRow; row <= totalRows; row += 4)
            {
                // 1st sub-row: Month and days
                string monthInfo = worksheet.Cells[row, startColumn].Text.Trim(); // e.g., "01,2025"
                bool isMonthInfoValid = DateTime.TryParseExact(monthInfo, new[] { "MM,yyyy", "MM.yyyy" }, null, System.Globalization.DateTimeStyles.None, out DateTime monthStart);
                if (!isMonthInfoValid)
                {
                    MessageBox.Show($"Ungültiges Monatsformat in der Excel-Datei entdeckt: {monthInfo}");
                    continue;  // Skip this month if the format is incorrect
                }
                if (DateTime.TryParseExact(monthInfo, new[] { "MM,yyyy", "MM.yyyy" }, null, System.Globalization.DateTimeStyles.None, out monthStart))
                {
                    // Loop through the days in this month (columns 2 to totalColumns)
                    for (int col = startColumn+1; col <= totalColumns; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        string dayText = cell.Text.Trim() ?? string.Empty;
                        // If it's a valid day, process it
                        if (int.TryParse(dayText, out int day))
                        {
                            DateTime eventStartDate = new DateTime(monthStart.Year, monthStart.Month, day);

                            // 2nd sub-row: Shift Type
                            string shiftType = worksheet.Cells[row + 1, col].Text.Trim();
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

