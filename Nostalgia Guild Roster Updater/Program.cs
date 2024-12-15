using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace NostalgiaGuildRosterPointsUpdater
{
    public class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public partial class MainForm : Form
    {
        private TextBox txtCsvInput;
        private Button btnUpload;
        private Label lblStatus;

        public MainForm()
        {
            txtCsvInput = new TextBox();
            btnUpload = new Button();
            lblStatus = new Label();
            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            // Setting up the form
            this.Text = "Nostalgia Guild Roster Points Updater";
            this.Size = new System.Drawing.Size(600, 400);
            this.BackColor = System.Drawing.Color.FromArgb(30, 30, 30);
            this.ForeColor = System.Drawing.Color.White;

            // TextBox for CSV input
            txtCsvInput = new TextBox
            {
                Multiline = true,
                Text = string.Empty,
                Dock = DockStyle.Top,
                Height = 200,
                BackColor = System.Drawing.Color.FromArgb(50, 50, 50),
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(txtCsvInput);

            // Upload Button
            btnUpload = new Button
            {
                Text = "Upload to Google Sheets",
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = System.Drawing.Color.FromArgb(70, 70, 70),
                ForeColor = System.Drawing.Color.White
            };
            btnUpload.Click += btnUpload_Click;
            this.Controls.Add(btnUpload);

            // Status Label
            lblStatus = new Label
            {
                Text = "Status: Ready",
                Dock = DockStyle.Bottom,
                Height = 30,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                ForeColor = System.Drawing.Color.White
            };
            this.Controls.Add(lblStatus);
        }

        private void btnUpload_Click(object? sender, EventArgs? e)
        {
            try
            {
                string csvData = txtCsvInput.Text;
                if (string.IsNullOrWhiteSpace(csvData))
                {
                    lblStatus.Text = "Status: No data entered!";
                    return;
                }

                var parsedData = ParseCsv(csvData);
                UpdateGoogleSheet(parsedData);
                lblStatus.Text = "Status: Successfully uploaded!";
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"Error: {ex.Message}";
            }
        }

        private List<(string Name, int Points)> ParseCsv(string csvData)
        {
            var result = new List<(string, int)>();
            var rows = csvData.Split('\n');
            foreach (var row in rows)
            {
                var columns = row.Split(',');
                if (columns.Length < 2) continue;

                string name = columns[0].Trim();
                if (int.TryParse(columns[1].Trim(), out int points))
                {
                    result.Add((name, points));
                }
            }
            return result;
        }

        private void UpdateGoogleSheet(List<(string Name, int Points)> data)
        {
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            string ApplicationName = "NostalgiaGuildRosterPointsUpdater";

            GoogleCredential credential;
            using (var stream = new FileStream("C:\\Users\\fuck\\Documents\\GitHub\\NostalgiaGuildRosterAndPointsUpdater\\NostalgiaGuildRosterUpdater\\key", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            string spreadsheetId = "1k-XR-RuO8QEnlJqNVSIfxG6-tgg2Qjp7qV88WNEDMBQ"; // Replace with your spreadsheet ID
            string range = "Q2:Q";

            var values = new List<IList<object>>();
            foreach (var item in data)
            {
                values.Add(new List<object> { item.Points });
            }

            var valueRange = new ValueRange
            {
                Values = values
            };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();
        }
    }
}
