using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using UnityEngine;
using Google.Apis.Services;
using System;
using Google.Apis.Sheets.v4.Data;
using System.IO;
using System.Threading.Tasks;

public class GoogleSheetsForUnity : MonoBehaviour
{
    [Header("GoogleSheets Information")]
    [SerializeField] private string spreadSheetID;
    [SerializeField] private string sheetID;

    [Header("Data from GoogleSheets")]
    [SerializeField] private string getDataInRange;

    private string serviceAccountEmail = "";
    private string certificateName = "";
    private string certificatePath;

    private static SheetsService googleSheetsService;
    [Serializable]
    public class Row
    {
        public List<string> cellData = new List<string>();
    }
    [Serializable]
    public class RowList
    {
        public List<Row> rows = new List<Row>();
    }

    public RowList DataFromGoogleSheets = new RowList();

    [Header("Write Data From Unity")]
    [SerializeField] private string writeDataInRange;

    public RowList WriteDataFromUnity = new RowList();

    [Header("Delete Data In GoogleSheets")]
    [SerializeField] private string deleteDataInRange;


    void Start()
    {
        //Remove comment to use on Android

/*        string tempPath = Path.Combine(Application.streamingAssetsPath, certificateName);
        WWW reader = new WWW(tempPath);
        while (!reader.isDone) { }
        certificatePath = Application.persistentDataPath + "/db";
        File.WriteAllBytes(certificatePath, reader.bytes);*/


        certificatePath = Application.dataPath + "/StreamingAssets/" + certificateName;  //Comment to use on Android

        var certificate = new X509Certificate2(certificatePath, "notasecret", X509KeyStorageFlags.Exportable);

        ServiceAccountCredential credential = new ServiceAccountCredential(
            new ServiceAccountCredential.Initializer(serviceAccountEmail)
            {
                Scopes = new[] { SheetsService.Scope.Spreadsheets }
            }.FromCertificate(certificate));

        googleSheetsService = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "GoogleSheets API for Unity"
        });

        //Use async methods to increase Android performance.
    }

    public void ReadData()
    {
        string range = sheetID + "!" + getDataInRange;

        var request = googleSheetsService.Spreadsheets.Values.Get(spreadSheetID, range);
        var reponse = request.Execute();
        var values = reponse.Values;
        if (values != null && values.Count > 0)
        {
            foreach (var row in values)
            {
                Row newRow = new Row();
                DataFromGoogleSheets.rows.Add(newRow);
                foreach (var value in row)
                {
                    newRow.cellData.Add(value.ToString());
                    Debug.Log(value.ToString());
                }

            }
        }
    }

    public async void ReadDataAsyn()
    {
        var task = await Task.Run(() =>
        {
            string range = sheetID + "!" + getDataInRange;

            var request = googleSheetsService.Spreadsheets.Values.Get(spreadSheetID, range);
            var reponse = request.Execute();
            var values = reponse.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    Row newRow = new Row();
                    DataFromGoogleSheets.rows.Add(newRow);
                    foreach (var value in row)
                    {
                        newRow.cellData.Add(value.ToString());
                    }

                }
            }
            return 0;
        });
    }

    public void WriteData()
    {
        string range = sheetID + "!" + writeDataInRange;
        var valueRange = new ValueRange();
        var cellData = new List<object>();
        var arrows = new List<IList<object>>();
        foreach (var row in WriteDataFromUnity.rows)
        {
            cellData = new List<object>();
            foreach (var data in row.cellData)
            {
                cellData.Add(data);
            }

            arrows.Add(cellData);
        }

        valueRange.Values = arrows;

        var request = googleSheetsService.Spreadsheets.Values.Append(valueRange, spreadSheetID, range);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        var reponse = request.Execute();
    }

    public async void WriteDataAsyn()
    {
        var task = await Task.Run(() =>
        {
            string range = sheetID + "!" + writeDataInRange;
            var valueRange = new ValueRange();
            var cellData = new List<object>();
            var arrows = new List<IList<object>>();
            foreach (var row in WriteDataFromUnity.rows)
            {
                cellData = new List<object>();
                foreach (var data in row.cellData)
                {
                    cellData.Add(data);
                }

                arrows.Add(cellData);
            }

            valueRange.Values = arrows;

            var request = googleSheetsService.Spreadsheets.Values.Append(valueRange, spreadSheetID, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var reponse = request.Execute();
            return 0;
        });
    }

    public void DeleteData()
    {
        var range = sheetID + "!" + deleteDataInRange;

        var deleteData = googleSheetsService.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadSheetID, range);
        deleteData.Execute();
    }

    public async void DeleteDataAsyn()
    {
        var task = await Task.Run(() =>
        {
            var range = sheetID + "!" + deleteDataInRange;

            var deleteData = googleSheetsService.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadSheetID, range);
            deleteData.Execute();
            return 0;
        });
    }
    
    
    public void UpdateData()
    {
        string range = sheetID + "!" + writeDataInRange;
        var valueRange = new ValueRange();
        var cellData = new List<object>();
        var arrows = new List<IList<object>>();
        foreach (var row in WriteDataFromUnity.rows)
        {
            cellData = new List<object>();
            foreach (var data in row.cellData)
            {
                cellData.Add(data);
            }

            arrows.Add(cellData);
        }

        valueRange.Values = arrows;

        var updateRequest = googleSheetsService.Spreadsheets.Values.Update(valueRange, spreadSheetID, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        var appendReponse = updateRequest.Execute();
    }

    public async void UpdateDataAsyn()
    {
        var task = await Task.Run(() =>
        {
            string range = sheetID + "!" + writeDataInRange;
            var valueRange = new ValueRange();
            var cellData = new List<object>();
            var arrows = new List<IList<object>>();
            foreach (var row in WriteDataFromUnity.rows)
            {
                cellData = new List<object>();
                foreach (var data in row.cellData)
                {
                    cellData.Add(data);
                }

                arrows.Add(cellData);
            }

            valueRange.Values = arrows;

            var updateRequest = googleSheetsService.Spreadsheets.Values.Update(valueRange, spreadSheetID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = updateRequest.Execute();
            return 0;
        });
    }


}
