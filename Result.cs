using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMViewer
{
    class Result
    {
        private DataTable _data;
        public DataTable Data
        {
            get { return _data; }
            set
            {
                _data = value;
                columnWidths = null;
                maxColumnWidths = null;
                page = 0;
            }
        }
        string header;
        public string Header
        {
            get
            {
                if (!string.IsNullOrEmpty(header))
                    return header;
                else
                    return Data.TableName;
            }
            set { header = value; }
        }

        public string LogicalName { get; internal set; }

        private int[] columnWidths;
        private int[] maxColumnWidths;
        private int lastConsoleWidth = 0;
        private int lastConsoleHeight = 0;
        private int indexWidth = 3;
        private int page = 0;

        public Result() { }
        public Result(DataTable data)
        {
            this.Data = data;
        }
        public Result(DataTable data, string header)
        {
            this.Data = data;
            this.Header = header;
        }

        public ResultResponse Draw()
        {
            return Draw("Selection: ");
        }
        public ResultResponse Draw(string PromptOverride)
        { 
            ResultResponse retVal = new ResultResponse();

            //while (retVal.ConsoleKey != ConsoleKey.Enter)
            //{

                DrawTable(Header);
                retVal = ProcessInput(PromptOverride);
            //}

            switch (retVal.ConsoleKey)
            {
                case ConsoleKey.F11:
                    Console.Write("\r\nFilename to save: ");
                    string filename = Console.ReadLine();
                    StreamWriter sw = File.CreateText(filename);
                    sw.WriteLine("Sep=^");
                    List<string> line = new List<string>();
                    foreach (DataColumn dc in Data.Columns)
                        line.Add(CSVFormat(dc.ColumnName));
                    sw.WriteLine(string.Join(@"^", line));

                    foreach (DataRow dr in Data.Rows)
                    {
                        line = new List<string>();
                        foreach (DataColumn dc in Data.Columns)
                            line.Add(CSVFormat(Util.GetString(dr[dc.ColumnName])));
                        sw.WriteLine(string.Join(@"^", line));
                    }

                    sw.Flush();
                    sw.Close();
                    retVal = new ResultResponse();
                    break;

                default:
                    break;
            }

            return retVal;
        }

        private string CSVFormat(object input)
        {
            string retVal = string.Empty;
            if (input == null)
                return string.Empty;
            else if (input.ToString().ToLower() == "<null>")
                return string.Empty;
            else
                retVal = input.ToString();

            bool addquotes = false;

            if (retVal.Contains(@""""))
            {
                retVal = retVal.Replace(@"""", @"""""");
                addquotes = true;
            }

            if (retVal.Contains(",") && !(retVal.StartsWith(@"""") && retVal.EndsWith(@"""")))
                addquotes = true;

            while (retVal.Contains("\r") || retVal.Contains("\n"))
            {
                retVal = retVal.Replace("\r", "<br/>");
                retVal = retVal.Replace("\n", "<br/>");
                addquotes = true;
            }

            if (addquotes)
                retVal = @"""" + retVal + @"""";

            return retVal.Trim();
        }

        private void CalculateWidths()
        {
            if (maxColumnWidths == null)
            {
                maxColumnWidths = new int[Data.Columns.Count];
                for (int i = 0; i <= maxColumnWidths.GetUpperBound(0); i++)
                    maxColumnWidths[i] = 0;

                foreach (DataRow r in Data.Rows)
                {
                    object[] vals = r.ItemArray;
                    for (int i = 0; i <= maxColumnWidths.GetUpperBound(0); i++)
                        maxColumnWidths[i] = Math.Max(maxColumnWidths[i], Util.GetString(vals[i]).Length);
                }
            }

            if (Console.WindowWidth != lastConsoleWidth)
            {
                //recalculate columns
                columnWidths = new int[Data.Columns.Count];
                for (int i = 0; i < Data.Columns.Count; i++)
                    columnWidths[i] = 5;

                indexWidth = Math.Max(3, (int)Math.Ceiling(Math.Log10(Data.Rows.Count)) + 2);
                int colGapsWidth = 3 * (Data.Columns.Count - 1);
                int remainingWidth = Console.WindowWidth - (indexWidth + colGapsWidth + columnWidths.Sum());
                bool breakout = false;
                while (remainingWidth > Data.Columns.Count)
                {
                    breakout = true;
                    for (int i = 0; i <= columnWidths.GetUpperBound(0); i++)
                        if (columnWidths[i] < maxColumnWidths[i])
                        {
                            columnWidths[i]++;
                            breakout = false;
                        }
                    if (breakout) break;
                    remainingWidth = Console.WindowWidth - (indexWidth + colGapsWidth + columnWidths.Sum());
                }
            }
        }

        ResultResponse ProcessInput(string prompt)
        {
            ConsoleKeyInfo retVal = new ConsoleKeyInfo();
            string input = "";
            bool inputDone = false;


            while (!inputDone)
            {
                Console.CursorLeft = 0;
                Console.Write(new string(' ', Console.WindowWidth - 1));
                Console.CursorLeft = 0;
                Console.Write(string.Format("{0}{1}", prompt, input));

                while (!Console.KeyAvailable)
                    System.Threading.Thread.Sleep(1);

                while (Console.KeyAvailable)
                {
                    retVal = Console.ReadKey(true);
                    switch (retVal.Key)
                    {
                        case ConsoleKey.Escape:
                            inputDone = true;
                            break;

                        case ConsoleKey.UpArrow:
                        case ConsoleKey.PageUp:
                            page = Math.Max(0, page - 1);
                            DrawTable(Header);
                            break;
                        case ConsoleKey.PageDown:
                        case ConsoleKey.DownArrow:
                            page = Math.Max(0, page + 1);
                            while (page * (Console.WindowHeight - 4) > Data.Rows.Count)
                                page--;
                            DrawTable(Header);
                            break;

                        case ConsoleKey.Backspace:
                            if (input.Length > 0)
                                input = input.Substring(0, input.Length - 1);
                            break;

                        case ConsoleKey.Enter:
                            inputDone = true;
                            break;

                        case ConsoleKey.F7:
                            Console.Write(" Sort column #: ");
                            string sortColumnString = Console.ReadLine();
                            int sortColumn = -1;

                            if (int.TryParse(sortColumnString, out sortColumn) &&
                                Math.Abs(sortColumn) <= Data.Columns.Count &&
                                Math.Abs(sortColumn) > 0)
                            {
                                //time to re-sort                            
                                if (sortColumn > 0)
                                    Data.DefaultView.Sort = string.Format("[{0}] ASC", Data.Columns[sortColumn - 1].ColumnName);
                                else
                                    Data.DefaultView.Sort = string.Format("[{0}] DESC", Data.Columns[Math.Abs(sortColumn) - 1].ColumnName);
                            }
                            DrawTable(Header);
                            break;

                        case ConsoleKey.D0:
                        case ConsoleKey.D1:
                        case ConsoleKey.D2:
                        case ConsoleKey.D3:
                        case ConsoleKey.D4:
                        case ConsoleKey.D5:
                        case ConsoleKey.D6:
                        case ConsoleKey.D7:
                        case ConsoleKey.D8:
                        case ConsoleKey.D9:
                        case ConsoleKey.NumPad0:
                        case ConsoleKey.NumPad1:
                        case ConsoleKey.NumPad2:
                        case ConsoleKey.NumPad3:
                        case ConsoleKey.NumPad4:
                        case ConsoleKey.NumPad5:
                        case ConsoleKey.NumPad6:
                        case ConsoleKey.NumPad7:
                        case ConsoleKey.NumPad8:
                        case ConsoleKey.NumPad9:
                            input += retVal.KeyChar;
                            break;

                        default:
                            inputDone = true;
                            break;
                    }

                }
            }

            return new ResultResponse() { ConsoleKey = retVal.Key, Modifiers = retVal.Modifiers, Response = input };
        }

        private void DrawTable(string header)
        {
            CalculateWidths();

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(header);

            object[] vals;
            string[] valsStrings = new string[Data.Columns.Count];
            int rowsAdded = 0;

            sb.AppendFormat(new string(' ', indexWidth));
            string[] colNames = new string[Data.Columns.Count];
            for (int i = 0; i < Data.Columns.Count; i++)
                colNames[i] = Data.Columns[i].ColumnName.PadRight(columnWidths[i]).Substring(0, columnWidths[i]); ;
            sb.AppendLine(string.Join(" | ", colNames));

            for (int i = page * (Console.WindowHeight - 4); i < Data.Rows.Count; i++)
            {
                sb.AppendFormat("{0," + (indexWidth - 2).ToString() + "}. ", i + 1);
                vals = Data.DefaultView[i].Row.ItemArray;
                for (int j = 0; j <= vals.GetUpperBound(0); j++)
                    valsStrings[j] = Util.GetString(vals[j]).PadRight(columnWidths[j]).Substring(0, columnWidths[j]);
                sb.AppendLine(string.Join(" | ", valsStrings));

                if (++rowsAdded >= (Console.WindowHeight - 4)) break;
            }
            sb.AppendLine();
            Console.Clear();
            Console.Write(sb.ToString());
        }
    }
}
