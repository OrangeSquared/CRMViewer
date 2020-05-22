using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CRMViewer
{
    class Query
    {
        private Connection connection;
        private Stack<Result> data;

        public Query(Connection connection)
        {
            this.connection = connection;
            data = new Stack<Result>();
        }

        public void Begin()
        {
            data.Push(new Result() { Data = LoadData(), LogicalName = "querylist", Header = "Query List" });

            ResultResponse returnCode = new ResultResponse();
            while (true)
            {
                bool breakout = false;

                if (returnCode.ConsoleKey != ConsoleKey.F1)
                    returnCode = data.Peek().Draw();
                else
                    returnCode = data.Peek().Draw("F2: Paste New   F8: Save As   Selection: ");

                switch (returnCode.ConsoleKey)
                {
                    case ConsoleKey.Escape: breakout = true; break;

                    case ConsoleKey.F2:
                        Console.WriteLine("Paste Query:");
                        StringBuilder newquery = new StringBuilder();
                        while (true)
                        {
                            string line = Console.ReadLine();
                            newquery.AppendLine(line);
                            if (line.ToLower().Contains("</fetch>")) break;
                        }
                        DataTable newQueryData = Util.GetQueryDataTable(newquery.ToString(), connection);
                        Result newQueryResult = new Result() { Data = newQueryData };
                        returnCode = newQueryResult.Draw();

                        switch (returnCode.ConsoleKey)
                        {
                            case ConsoleKey.F8:
                                Console.Write("\r\nFilename to save: ");
                                string filename = Console.ReadLine();
                                StreamWriter sw = File.CreateText(filename + ".xml");
                                sw.WriteLine(newquery);
                                sw.Flush();
                                sw.Close();
                                breakout = true;
                                break;

                            default:
                                break;
                        }
                        break;


                    default:
                        int returnrow = 0;
                        if (int.TryParse(returnCode.Response, out returnrow))
                            if (returnrow >= 1 && returnrow <= data.Peek().Data.Rows.Count)
                            {
                                string query = string.Empty;
                                StreamReader sr = File.OpenText((string)data.Peek().Data.Rows[returnrow - 1]["Queries"] + ".xml");
                                query = sr.ReadToEnd();
                                sr.Close();

                                Regex regex = new Regex(@"\{[\w ]+\}");
                                MatchCollection matchCollection = regex.Matches(query);

                                if (matchCollection.Count > 0)
                                {
                                    Console.WriteLine("\r\nQuery Variables:");
                                    foreach (Match match in matchCollection)
                                    {
                                        Console.Write(match.Value.Substring(1, match.Value.Length - 2) + ": ");
                                        string value = Console.ReadLine();
                                        query = query.Replace(match.Value, value);
                                    }
                                }

                                DataTable queryData = Util.GetQueryDataTable(query.ToString(), connection);
                                Result queryResult = new Result() { Data = queryData, LogicalName = queryData.TableName };
                                returnCode = queryResult.Draw();

                                if (returnCode.ConsoleKey == ConsoleKey.Delete )
                                {
                                    File.Delete((string)data.Peek().Data.Rows[returnrow - 1]["Queries"] + ".xml");
                                    breakout = true;
                                }

                            }
                        break;
                }

                if (breakout) break;

            }
        }

        private DataTable LoadData()
        {
            DataTable data = new DataTable();
            data.Columns.AddRange(
                 new DataColumn[]
                 {
                     new DataColumn("Queries",typeof(string)),
                     new DataColumn("Modified On", typeof(DateTime))
                 }
                );

            string[] files = Directory.GetFiles("./", "*.xml");
            foreach (string file in files)
                if (!(file.Contains("Microsoft.Crm.Sdk.Proxy.xml") || file.Contains("Microsoft.Xrm.Sdk.xml")))
                {
                    DataRow dr = data.NewRow();
                    dr["Queries"] = Path.GetFileNameWithoutExtension(file);
                    dr["Modified On"] = File.GetLastWriteTime(file);
                    data.Rows.Add(dr);
                }

            return data;
        }
    }
}