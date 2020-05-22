using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMViewer
{
    class Program
    {
        static void Main(string[] args)
        {
            int serverindex = 0;
            string[] servers = Directory.GetFiles(".\\", "server*.cfg");
            string user = string.Empty;
            string pass = string.Empty;
            string url = string.Empty;
            GetServerLogin(servers[serverindex], out user, out pass, out url);

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Menu", typeof(string)));

            dt.Rows.Add(new object[] { "Browse" });
            dt.Rows.Add(new object[] { "Query" });
            dt.Rows.Add(new object[] { "Search" });
            dt.Rows.Add(new object[] { "Utility" });
            dt.Rows.Add(new object[] { "Solutions" });
            dt.Rows.Add(new object[] { "Switch Server" });

            Result r = new Result() { Data = dt, Header = "Pointed to server " + url };

            while (true)
            {
                ResultResponse returnCode = r.Draw();

                if (returnCode.ConsoleKey == ConsoleKey.Escape)
                    break;

                int itemNumber = -1;

                if (int.TryParse(returnCode.Response, out itemNumber))
                {
                    itemNumber--;
                    string command = (string)dt.DefaultView[itemNumber][0];

                    switch (command)
                    {
                        case "Browse":
                            Console.Write("\r\nLoading Entities");
                            Console.CursorLeft = 0;
                            Browser b = new Browser(new Connection(user, pass, url));
                            b.Begin();
                            break;

                        case "Query":
                            Query q = new Query(new Connection(user, pass, url));
                            q.Begin();
                            break;

                        case "Search":
                            break;

                        case "Utility":
                            Utility u = new Utility(new Connection(user, pass, url));
                            u.Begin();
                            break;

                        case "Solutions":
                            Solutions s = new Solutions(new Connection(user, pass, url));
                            s.Begin();
                            break;

                        case "Switch Server":
                            serverindex++;
                            if (serverindex > servers.GetUpperBound(0)) serverindex = 0;
                            GetServerLogin(servers[serverindex], out user, out pass, out url);
                            r.Header = "Pointed to server " + url;
                            break;

                        default:
                            break;
                    }
                }
            }

            Console.WriteLine("done.");
        }

        private static void GetServerLogin(string path, out string user, out string pass, out string url)
        {
            string file = File.ReadAllText(path);
            string[] lines = file.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            user = lines[0];
            pass = lines[1];
            url = lines[2];

        }
    }
}
