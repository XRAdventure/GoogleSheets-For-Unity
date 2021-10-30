using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System.Collections;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using UnityEngine;
using Google.Apis.Services;
using System;
using Google.Apis.Sheets.v4.Data;
using System.IO;

public class GoogleSheetsAPI : MonoBehaviour
{
    [Header("GoogleSheets Information")]
    [SerializeField] private string spreadSheetID;
    [SerializeField] private string sheetID;

    [Header("Data from GoogleSheets")]
    [SerializeField] private string getDataInRange;

    private string serviceAccountEmail = "googlesheetsunity@unityapi-329521.iam.gserviceaccount.com";
    private string certificateName = "unityapi-329521-8ddbe692d4e0.p12";
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
        /* Uncomment if you will use Android
        string androidPath = "jar:file://" + Application.dataPath + "!/assets/";
        string tempPath =  Path.Combine(androidPath, certificateName);
        WWW reader = new WWW(tempPath);
        while (!reader.isDone) { }
        var androidJarPath = Application.persistentDataPath + "/db";
        File.WriteAllBytes(androidJarPath, reader.bytes);
        */

        certificatePath = "/StreamingAssets/" + certificateName;

        var certificate = new X509Certificate2(Application.dataPath + certificatePath, "notasecret", X509KeyStorageFlags.Exportable);

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

        ReadData();
    }

    public void ReadData()
    {
        string range = sheetID + "!" + getDataInRange;

        var request = googleSheetsService.Spreadsheets.Values.Get(spreadSheetID, range);
        var reponse = request.Execute();
        var values = reponse.Values;
        if(values != null && values.Count > 0)
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

    public void DeleteData()
    {
        var range = sheetID + "!" + deleteDataInRange;

        var deleteData = googleSheetsService.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadSheetID, range);
        deleteData.Execute();
    }
}
