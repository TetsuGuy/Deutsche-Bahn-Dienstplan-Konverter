using System.IO;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;

namespace Deutsche_Bahn_Dienstplan_zu_Kalender_Konverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ResultFileName = "Calendar.ics";
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
                        saveFileDialog.FileName = ResultFileName.Text;

                        // If the user selects a location and clicks "Save"
                        if (saveFileDialog.ShowDialog() == true)
                        {
                            string calendarFilePath = saveFileDialog.FileName;

                            // Write the calendar content to the selected file
                            File.WriteAllText(calendarFilePath, calendarContent);
                            MessageBox.Show($"Calendar file saved successfully at {calendarFilePath}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}");
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
                        string dayText = string.Empty;

                        var cell = worksheet.Cells[row, col];

                        // Check if the current cell is merged
                        if (cell.Merge)
                        {
                            // Get the merged range's first (top-left) cell
                            var mergedRange = worksheet.Cells[cell.Start.Row, cell.Start.Column];
                            dayText = mergedRange.Text;

                            // Skip all columns in this merged block
                            col = cell.End.Column; // Move to the last column of the merged block
                        }
                        else
                        {
                            dayText = cell.Text;
                        }

                        // If it's a valid day, process it
                        if (int.TryParse(dayText, out int day))
                        {
                            DateTime eventStartDate = new DateTime(monthStart.Year, monthStart.Month, day);

                            // 2nd sub-row: Shift Type
                            string shiftType = worksheet.Cells[row + 1, col].Text;

                            // Handle whole-number shift types as continuous events
                            if (int.TryParse(shiftType, out int _)) // Whole number detected
                            {
                                int endCol = col;

                                // Continue until the shift type changes or the column ends
                                while (endCol <= totalColumns)
                                {
                                    string nextShiftType = worksheet.Cells[row + 1, endCol].Text;
                                    string nextCellText = worksheet.Cells[row, endCol].Merge ?
                                        worksheet.Cells[worksheet.Cells[row, endCol].Start.Row, worksheet.Cells[row, endCol].Start.Column].Text :
                                        worksheet.Cells[row, endCol].Text;

                                    bool isNextShiftType = !int.TryParse(nextShiftType, out int _);
                                    bool isEmptyCell = string.IsNullOrWhiteSpace(nextCellText);

                                    if (isNextShiftType || isEmptyCell)
                                        break;

                                    endCol++;
                                }

                                // The last day in this block
                                DateTime eventEndDate = new DateTime(monthStart.Year, monthStart.Month,
                                    int.Parse(worksheet.Cells[row, endCol - 1].Text)).AddDays(1);

                                // Add the full-day event to the calendar
                                calendarContent += CreateFullDayEvent(eventStartDate, eventEndDate, $"Free ({shiftType})");

                                // Skip processed columns
                                col = endCol - 1;
                            }
                            else if (!string.IsNullOrEmpty(shiftType)) // Handle normal shift types
                            {
                                DateTime eventEndDate = eventStartDate.AddDays(1); // Single-day event
                                calendarContent += CreateFullDayEvent(eventStartDate, eventEndDate, shiftType);
                            }
                        }
                    }
                }
            }

            calendarContent += "END:VCALENDAR";
            return calendarContent;
        }


        private string CreateFullDayEvent(DateTime startDate, DateTime endDate, string summary)
        {
            return $"BEGIN:VEVENT\n" +
                   $"SUMMARY:{summary}\n" +
                   $"DTSTART;VALUE=DATE:{startDate:yyyyMMdd}\n" +
                   $"DTEND;VALUE=DATE:{endDate:yyyyMMdd}\n" +
                   "END:VEVENT\n";
        }

        private void FilePathTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
    }
}

