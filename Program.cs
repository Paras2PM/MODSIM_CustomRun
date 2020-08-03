using Csu.Modsim.ModsimIO;
using Csu.Modsim.ModsimModel;
using OfficeOpenXml;
using System;
using System.Data.OleDb;
using System.IO;

public static class CustomMODSIM
{
    public static void Main(string[] CmdArgs)
    {
        string fileName = CmdArgs[0];
        string inputFileName = CmdArgs[1];
        string resultFileName = Path.GetDirectoryName(inputFileName) + "\\out.xlsx";

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        if (File.Exists(resultFileName))
        {
            File.Delete(resultFileName);
        }

        using (var package = new ExcelPackage(new FileInfo(inputFileName)))
        using (var resultPackage = new ExcelPackage(new FileInfo(resultFileName)))
        {
            var sheets = package.Workbook.Worksheets;

            foreach (var sheet in sheets)
            {
                var resultSheet = resultPackage.Workbook.Worksheets.Add(sheet.Name);

                Log($"Run modsim for sheet '{sheet.Name}'");

                RunSolver(fileName, sheet.Cells, resultSheet);

                Log($"Finish running modsim for sheet '{sheet.Name}'");
            }

            Log($"Save result file to '{resultFileName}'");
            resultPackage.Save();

            Log($"Finish");
            Log($"Press any key to exit");
            Console.ReadKey();
        }
    }

    private static void Log(string text)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine(text);
        Console.ResetColor();
    }
    private static void RunSolver(string fileName, ExcelRange rows, ExcelWorksheet resultSheet)
    {
        Model myModel = new Model();

        myModel.Init += () =>
        {
            var nodes = new Node[rows.Columns - 1];
            for (var i = 0; i < rows.Columns - 1; i++)
            {
                nodes[i] = myModel.FindNode(rows[1, i + 2].GetValue<string>(), false);
            }


            for (int i = 0; i < myModel.TimeStepManager.noDataTimeSteps; i++)
            {
                for (var j = 0; j < rows.Columns; j++)
                {
                    nodes[j].mnInfo.nodedemand[i, 0] = rows[i + 2, j + 2].GetValue<int>();
                }
            }

        };

        //myModel.OnMessage += (message) => { Console.WriteLine(message); };

        XYFileReader.Read(myModel, fileName);

        Modsim.RunSolver(myModel);

        Log($"Parse modsim output");

        ParseModsimResult(fileName, resultSheet);
    }

    private static void ParseModsimResult(string fileName, ExcelWorksheet resultSheet)
    {
        var source = fileName.Replace(".xy", "OUTPUT.mdb");
        using (var conection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + source + ";Persist Security Info=False;"))
        {
            conection.Open();
            var query = @"SELECT [DEMOutput].[Surf_In], [TimeSteps].[TSDate], [NodesInfo].[NName] from ([DEMOutput]
                         INNER JOIN [TimeSteps] ON [DEMOutput].[TSIndex] = [TimeSteps].[TSIndex])
                         INNER JOIN [NodesInfo] ON [DEMOutput].[NNo] = [NodesInfo].[NNumber]
                         ORDER BY [DEMOutput].[TSIndex],[DEMOutput].[NNo]";

            var command = new OleDbCommand(query, conection);
            var reader = command.ExecuteReader();

            resultSheet.Cells[1, 1].Value = "Date";

            var date = DateTime.MinValue;
            var row = 2;
            var col = 2;
            while (reader.Read())
            {
                if (date == DateTime.MinValue)
                {
                    date = reader.GetDateTime(1);
                }
                else if (date != reader.GetDateTime(1))
                {
                    col = 2;
                    row++;

                    date = reader.GetDateTime(1);
                }

                resultSheet.Cells[row, col].Value = reader.GetDouble(0);
                resultSheet.Cells[row, 1].Value = reader.GetDateTime(1).ToString("dd/MM/yyyy");
                resultSheet.Cells[1, col].Value = reader.GetString(2);
                col++;
            }

        }
    }

}
