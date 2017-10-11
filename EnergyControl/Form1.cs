using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EnergyControl
{
    public partial class Form1 : Form

    {
        static string ApplicationName = "Update Google Sheet Data with Google Sheets API v4";
        static String spreadsheetId = "1xFHgjLLInDenc3g1UpydfGET6nViTcsgu0v7Rr8fl7k";
        static string sheetName = "Лист1";
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };

        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DateTime dateOnly = DateTime.Now;
            textBox1.Text = dateOnly.ToString("d");

            var service = OpenSheet();
            textBox2.Text = GetPreDate(service, "B1:B");
            textBox4.Text = GetPreDate(service, "A1:A");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var service = OpenSheet();
            UpdateRow(service, textBox1.Text);

            textBox2.Text = GetPreDate(service, "B1:B");
            textBox4.Text = GetPreDate(service, "A1:A");
        }

        static SheetsService OpenSheet()
        {
            UserCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Path.Combine
                    (System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal),
                     ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        string UpdateRow(SheetsService service, string textBox)
        {
            ValueRange rVR;
            String sRange;
            int rowNumber = 1;
            String rowNumberString;

            sRange = String.Format("{0}!A:A", sheetName);
            SpreadsheetsResource.ValuesResource.GetRequest getRequest
                = service.Spreadsheets.Values.Get(spreadsheetId, sRange);
            rVR = getRequest.Execute();
            IList<IList<Object>> values = rVR.Values;

            if (values != null && values.Count > 0) rowNumber = values.Count + 1;
            sRange = String.Format("{0}!A{1}:B{1}", sheetName, rowNumber);

            ValueRange valueRange = new ValueRange();
            valueRange.Range = sRange;
            valueRange.MajorDimension = "ROWS";

            DateTime dt = new DateTime();
            dt = DateTime.Now;
            List<object> oblist = new List<object>() { textBox, String.Format("{0}", rowNumber)};
            valueRange.Values = new List<IList<object>> { oblist };
            //Console.WriteLine("{0}, {1}", oblist[0], oblist[1]);

            rowNumber = rowNumber - 1;
            rowNumberString = "A" + rowNumber.ToString();
            if (oblist[0].ToString() != GetPreDate(service, rowNumberString))
            {
                SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest
                = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, sRange);
                updateRequest.ValueInputOption
                = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                UpdateValuesResponse uUVR = updateRequest.Execute();
            }
            else
            {
                MessageBox.Show("Показания уже вносились сегодня", "Ахтунг!", MessageBoxButtons.OK);
            }
            return sRange;
        }

        String GetPreDate(SheetsService service, string rowNumber)
        {
            String range = "Лист1!" + rowNumber;
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;

                foreach (var row in values)
                {
                    foreach (var col in row)
                    {
                        range = col.ToString();
                    }
                }

            return range;
        }
    }
}
