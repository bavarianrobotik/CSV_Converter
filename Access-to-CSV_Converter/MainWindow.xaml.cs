using System;
using System.Data;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Xml;
using ExcelDataReader; 

namespace AccessToCsvConverter
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // Event für das Durchsuchen der Datei
        private void BrowseFile(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            if (AccessRadioButton.IsChecked == true)
            {
                openFileDialog.Filter = "Access-Datenbank (*.mdb; *.accdb)|*.mdb;*.accdb";
            }
            else if (ExcelRadioButton.IsChecked == true)
            {
                openFileDialog.Filter = "Excel-Datei (*.xls; *.xlsx)|*.xls;*.xlsx";
            }
            else if (XmlRadioButton.IsChecked == true)
            {
                openFileDialog.Filter = "XML-Datei (*.xml)|*.xml";
            }

            if (openFileDialog.ShowDialog() == true)
            {
                FilePath.Text = openFileDialog.FileName;
            }
        }

        // Event für das Durchsuchen des Ausgabeordners
        private void BrowseOutputFolder(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                CsvOutputPath.Text = dialog.SelectedPath;
            }
        }

        // Event für das Konvertieren der Datei (Access, Excel, XML) in CSV
        private void ConvertToCsv(object sender, RoutedEventArgs e)
        {
            string filePath = FilePath.Text;
            string csvOutputPath = CsvOutputPath.Text;
            bool includeHeader = (bool)IncludeHeaderCheckBox.IsChecked; // Kopfzeile überprüfen
            string delimiter = CommaRadioButton.IsChecked == true ? ";" : "."; // Auswahl des Trennzeichens

            // Zahlenformat basierend auf der Benutzerauswahl
            CultureInfo numberCulture = NumberCommaRadioButton.IsChecked == true
                ? new CultureInfo("de-DE") // Format: 1.000,00
                : new CultureInfo("en-US"); // Format: 1,000.00

            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(csvOutputPath))
            {
                StatusText.Text = "Bitte wählen Sie eine Datei und einen Zielordner.";
                return;
            }

            try
            {
                if (AccessRadioButton.IsChecked == true)
                {
                    ConvertAccessToCsv(filePath, csvOutputPath, includeHeader, delimiter, numberCulture);
                }
                else if (ExcelRadioButton.IsChecked == true)
                {
                    ConvertExcelToCsv(filePath, csvOutputPath, includeHeader);
                }
                else if (XmlRadioButton.IsChecked == true)
                {
                    ConvertXmlToCsv(filePath, csvOutputPath);
                }
                StatusText.Text = "Konvertierung erfolgreich abgeschlossen.";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Fehler: {ex.Message}";
            }
        }

        // Methode für Access zu CSV
        private void ConvertAccessToCsv(string accessFilePath, string csvOutputPath, bool includeHeader, string delimiter, CultureInfo numberCulture)
        {
            // ODBC Verbindungszeichenfolge für Access-Datenbanken
            string connectionString = $"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={accessFilePath};";

            using (OdbcConnection connection = new OdbcConnection(connectionString))
            {
                connection.Open();

                // Alle Tabellennamen der Datenbank abfragen
                DataTable schemaTable = connection.GetSchema("Tables");
                foreach (DataRow row in schemaTable.Rows)
                {
                    string tableName = row["TABLE_NAME"].ToString();

                    // Nur Benutzertabellen exportieren, keine Systemtabellen
                    if (!tableName.StartsWith("MSys"))
                    {
                        ExportTableToCsv(connection, tableName, csvOutputPath, includeHeader, delimiter, numberCulture);
                    }
                }

                connection.Close();
            }
        }

        // Methode zum Exportieren einer Access-Tabelle als CSV
        private void ExportTableToCsv(OdbcConnection connection, string tableName, string csvOutputPath, bool includeHeader, string delimiter, CultureInfo numberCulture)
        {
            string query = $"SELECT * FROM [{tableName}] ORDER BY ID";
            OdbcCommand command = new OdbcCommand(query, connection);
            OdbcDataReader reader = command.ExecuteReader();

            string csvFilePath = Path.Combine(csvOutputPath, $"{tableName}.csv");

            using (StreamWriter writer = new StreamWriter(csvFilePath))
            {
                // Kopfzeile nur hinzufügen, wenn sie erwünscht ist
                if (includeHeader)
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        writer.Write(reader.GetName(i));
                        writer.Write(delimiter); // Trennzeichen
                    }
                    writer.WriteLine();
                }

                // Zeilen in die CSV-Datei schreiben
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        string columnName = reader.GetName(i);

                        // Prüfen, ob die Spalte nicht "ID" ist und ob sie eine Zahl ist
                        if (columnName != "ID" && decimal.TryParse(reader[i].ToString(), out decimal number))
                        {
                            // Prüfen, ob es sich um eine 4-stellige Zahl handelt
                           // if (number >= 1000 && number < 10000)
                           // {
                                // Für vierstellige Zahlen: Kein Tausendertrennzeichen, Nachkommastellen beibehalten
                                writer.Write(number.ToString("0.#####", numberCulture)); // Ausgabe ohne Tausendertrennzeichen
                            //}
                            //else
                            //{
                                // Für andere Zahlen: Tausendertrennzeichen behalten, keine Rundung, Nachkommastellen beibehalten
                              //  writer.Write(number.ToString("N", numberCulture)); // Ausgabe mit Tausendertrennzeichen
                            //}
                        }
                        else
                        {
                            // ID oder Nicht-Zahlenformat unverändert schreiben
                            writer.Write(reader[i].ToString());
                        }

                        writer.Write(delimiter); // Trennzeichen
                    }
                    writer.WriteLine(); // Neue Zeile nach jeder Datensatzzeile
                }
            }

            reader.Close();
        }

        // Methode für Excel zu CSV
        private void ConvertExcelToCsv(string excelFilePath, string csvOutputPath, bool includeHeader)
        {
            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int sheetIndex = 0;
                    do
                    {
                        string csvFilePath = Path.Combine(csvOutputPath, $"Sheet{sheetIndex}.csv");
                        using (var writer = new StreamWriter(csvFilePath))
                        {
                            if (includeHeader && reader.FieldCount > 0)
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    writer.Write(reader.GetName(i) + ",");
                                }
                                writer.WriteLine();
                            }

                            while (reader.Read())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    writer.Write(reader.GetValue(i)?.ToString() + ",");
                                }
                                writer.WriteLine();
                            }
                        }
                        sheetIndex++;
                    } while (reader.NextResult());
                }
            }
        }

        // Methode für XML zu CSV
        private void ConvertXmlToCsv(string xmlFilePath, string csvOutputPath)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFilePath);
            XmlNodeList rows = doc.SelectNodes("//row");

            string csvFilePath = Path.Combine(csvOutputPath, "XmlData.csv");
            using (StreamWriter writer = new StreamWriter(csvFilePath))
            {
                foreach (XmlNode row in rows)
                {
                    foreach (XmlNode cell in row.ChildNodes)
                    {
                        writer.Write(cell.InnerText + ",");
                    }
                    writer.WriteLine();
                }
            }
        }
    }
}
