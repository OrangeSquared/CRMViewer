using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using WikiClientLibrary.Client;
using WikiClientLibrary.Pages;
using WikiClientLibrary.Sites;
using Excel = Microsoft.Office.Interop.Excel;

namespace CRMViewer
{
    static class Util
    {
        public static async Task AddToWiki(string URL, string username, string password, string pagename, string content)
        {
            WikiClient wikiClient = new WikiClient() { ClientUserAgent = "CRMViewer/1.0" };
            WikiSite wikiSite = new WikiSite(wikiClient, URL);
            await wikiSite.Initialization;
            CancellationToken cancellationToken = new CancellationToken();
            await wikiSite.LoginAsync(username, password);


            WikiPage wikiPage = new WikiPage(wikiSite, pagename);
            await wikiPage.RefreshAsync();
            wikiPage.Content = content;
            await wikiPage.UpdateContentAsync("auto update", true, true);
            //await wikiSite.LogoutAsync();
        }

        internal static string GetString(object value)
        {
            string retVal = string.Empty;

            switch (value.GetType().Name)
            {
                case "String": retVal = (String)value; break;

                case "DateTime": retVal = ((DateTime)value).ToString("yyyy-MM-dd hh:mm:ss"); break;

                case "Int32": retVal = ((int)value).ToString(); break;

                case "Guid": retVal = ((Guid)value).ToString(); break;

                case "Money": retVal = ((Money)value).Value.ToString("C2"); break;

                case "EntityReference": retVal = "(ER)" + ((EntityReference)value).Name; break;
                //case "EntityReference": retVal = string.Format("{0} ({1})", ((EntityReference)value).LogicalName, ((EntityReference)value).Id); break;

                case "AliasedValue": retVal = GetString(((AliasedValue)value).Value); break;

                case "DBNull": retVal = "<null>"; break;

                default:
                    System.Diagnostics.Debug.WriteLine("no special format for " + value.GetType().Name);
                    retVal = value.ToString();
                    break;
            }


            return retVal;
        }

        internal static DataTable GetQueryDataTable(string newquery, Connection connection)
        {
            DataTable retVal = new DataTable();
            EntityCollection ec = connection.OrganizationService.RetrieveMultiple(new FetchExpression(newquery));
            List<DataColumn> columns = new List<DataColumn>();

            if (ec.Entities.Count > 0)
                retVal.TableName = ec.Entities[0].LogicalName;


            foreach (Entity e in ec.Entities)
                for (int i = 0; i < e.Attributes.Keys.Count; i++)
                //foreach (string k in e.Attributes.Keys)
                {
                    string k = e.Attributes.Keys.ToArray()[i];
                    if (!columns.Any(x => x.ColumnName == k))
                        columns.Insert(i, new DataColumn(k, typeof(string)));
                    //if (e[k].GetType() != typeof(AliasedValue))
                    //    columns.Insert(i, new DataColumn(k, e[k].GetType()));
                    //else
                    //    columns.Insert(i, new DataColumn(k, ((AliasedValue)e[k]).Value.GetType()));
                }

            foreach (DataColumn dataColumn in columns)
                retVal.Columns.Add(dataColumn);


            foreach (Entity e in ec.Entities)
            {
                DataRow newRow = retVal.NewRow();
                foreach (string k in e.Attributes.Keys)
                    newRow[k] = Util.EntityValueFormat(e, k, connection);
                //if (e[k].GetType() != typeof(AliasedValue))
                //    newRow[k] = e[k];
                //else
                //    newRow[k] = ((AliasedValue)e[k]).Value;
                retVal.Rows.Add(newRow);
            }

            return retVal;
        }

        internal static void ShowProgress(int value, int max)
        {
            int pos = Console.CursorLeft;
            Console.CursorLeft = 0;
            Console.Write(new string(' ', pos));
            int i = value - 2;
            List<string> output = new List<string>();
            while (output.Count < 5)
            {
                if (i > max) break;
                if (i > 0)
                {
                    if (i != value)
                        output.Add(i.ToString());
                    else
                        output.Add(string.Format("[{0}]", i));
                }
                i++;
            }
            Console.CursorLeft = 0;
            Console.Write(string.Join("..", output));
        }

        internal static List<Entity> GetAllUserPermissions(IOrganizationService orgsvc)
        {
            var moreRecords = false;
            int page = 1;
            var cookie = "";// string.Empty;
            List<Entity> Entities = new List<Entity>();
            do
            {
                var fetchXml = string.Format(@"
                <fetch {0} >
                  <entity name='systemuser' >
                    <attribute name='domainname' />
                    <attribute name='fullname' />
                    <filter>
                      <condition attribute='isdisabled' operator='eq' value='0' />
                    </filter>
                    <link-entity name='systemuserroles' from='systemuserid' to='systemuserid' intersect='true' >
                      <link-entity name='role' from='roleid' to='roleid' >
                        <attribute name='name' alias='RoleName' />
                        <link-entity name='roleprivileges' from='roleid' to='roleid' intersect='true' >
                          <link-entity name='privilege' from='privilegeid' to='privilegeid' >
                            <attribute name='accessright' alias='PrivilegeAccessRight' />
                            <attribute name='name' alias='PrivilegeName' />
                          </link-entity>
                        </link-entity>
                      </link-entity>
                    </link-entity>
                  </entity>
                </fetch>", cookie);
                EntityCollection collection = orgsvc.RetrieveMultiple(new FetchExpression(fetchXml));

                if (collection.Entities.Count >= 0) Entities.AddRange(collection.Entities);
                Console.WriteLine(string.Format("  Loaded {0}, ({1} total)", collection.Entities.Count, Entities.Count));
                System.Threading.Thread.Sleep(1);

                moreRecords = collection.MoreRecords;
                if (moreRecords)
                {
                    page++;
                    cookie = string.Format("paging-cookie='{0}' page='{1}'", System.Security.SecurityElement.Escape(collection.PagingCookie), page);
                }
            } while (moreRecords);

            return Entities;
        }

        internal static List<Entity> GetAllEntityRecords(string entityLogicalName, IOrganizationService orgsvc)
        {
            Console.WriteLine(string.Format("Loading {0}", entityLogicalName));
            var moreRecords = false;
            int page = 1;
            var cookie = "";// string.Empty;
            List<Entity> Entities = new List<Entity>();
            do
            {
                var fetchXml = string.Format(@"
                <fetch {1}>
                  <entity name='{0}' />
                </fetch>", entityLogicalName, cookie);
                EntityCollection collection = orgsvc.RetrieveMultiple(new FetchExpression(fetchXml));

                if (collection.Entities.Count >= 0) Entities.AddRange(collection.Entities);
                Console.WriteLine(string.Format("  Loaded {0}, ({1} total)", collection.Entities.Count, Entities.Count));
                System.Threading.Thread.Sleep(1);

                moreRecords = collection.MoreRecords;
                if (moreRecords)
                {
                    page++;
                    cookie = string.Format("paging-cookie='{0}' page='{1}'", System.Security.SecurityElement.Escape(collection.PagingCookie), page);
                }
            } while (moreRecords);

            return Entities;
        }

        internal static void SendToExcel(DataTable newDT)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook wb = app.Workbooks.Add();
            Excel.Worksheet ws = wb.Worksheets[1];

            int row = 1;
            int col = 1;
            foreach (DataColumn dc in newDT.Columns)
                ws.Cells[row, col++] = dc.ColumnName;
            col = 1;
            row++;
            foreach (DataRow dr in newDT.Rows)
            {
                foreach (DataColumn dc in newDT.Columns)
                    ws.Cells[row, col++] = dr[dc.ColumnName];
                col = 1;
                row++;
            }
        }

        internal static void SendToFile(List<Entity> results, string[] attributesToExport)
        {
            StreamWriter sw = File.CreateText("output.csv");

            Console.WriteLine("Loading into Excel");
            int max = results.Count;
            int pos = 1;
            foreach (Entity e in results)
            {
                foreach (string heading in attributesToExport)
                {
                    if (e.Contains(heading))
                    {
                        sw.Write("\"");
                        sw.Write(e[heading].GetType().Name != "AliasedValue" ? e[heading].ToString() : ((AliasedValue)e[heading]).Value.ToString());
                        sw.Write("\"");
                    }
                    sw.Write(",");
                }

                sw.WriteLine();
                if (pos++ % 100 == 99)
                    Util.ShowProgress(pos, max);
            }

            sw.Flush();
            sw.Close();
        }

        internal static void SendToExcel(List<Entity> results, string[] attributesToExport)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = false;
            app.ScreenUpdating = false;
            Excel.Workbook wb = app.Workbooks.Add();
            Excel.Worksheet ws = wb.Worksheets[1];

            Console.WriteLine("Loading into Excel");
            int max = results.Count;
            int pos = 1;
            int row = 1;
            int col = 1;
            foreach (string heading in attributesToExport)
                ws.Cells[row, col++] = heading;

            col = 1;
            row++;
            foreach (Entity e in results)
            {
                foreach (string heading in attributesToExport)
                    if (e.Contains(heading))
                        ws.Cells[row, col++] = e[heading].GetType().Name != "AliasedValue" ? e[heading].ToString() : ((AliasedValue)e[heading]).Value.ToString();
                    else
                        col++;
                col = 1;
                row++;
                if (pos++ % 100 == 99)
                    Util.ShowProgress(pos, max);
            }
            app.Visible = true;
            app.ScreenUpdating = true;
        }

        internal static string EntityValueFormat(Entity entity, string AttributeLogicalName, Connection connection)
        {
            switch (entity.Attributes[AttributeLogicalName].GetType().Name)
            {
                case "Int32": return entity.GetAttributeValue<Int32>(AttributeLogicalName).ToString(); break;
                case "String": return entity.GetAttributeValue<string>(AttributeLogicalName).Replace("\r\n", "<p>"); break;
                case "Boolean": return entity.GetAttributeValue<bool>(AttributeLogicalName) ? "True" : "False"; break;
                case "DateTime": return entity.GetAttributeValue<DateTime>(AttributeLogicalName).ToString("yyyy-MM-ddTHH:mm:ss"); break;
                case "Guid": return entity.GetAttributeValue<Guid>(AttributeLogicalName).ToString(); break;
                case "OptionSetValue":
                    //oh this is slow...
                    //RetrieveAttributeRequest rar = new RetrieveAttributeRequest()
                    //{
                    //    EntityLogicalName = entity.LogicalName,
                    //    LogicalName = AttributeLogicalName,
                    //    RetrieveAsIfPublished = true
                    //};
                    //RetrieveAttributeResponse rarr = (RetrieveAttributeResponse)connection.OrganizationService.Execute(rar);

                    //OptionMetadata op = null;
                    //if (rarr.AttributeMetadata.GetType().Name == "PicklistAttributeMetadata")
                    //    if (((PicklistAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options.Count > 0)
                    //        op = ((PicklistAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options.First(x => x.Value == entity.GetAttributeValue<OptionSetValue>(AttributeLogicalName).Value);

                    //if (rarr.AttributeMetadata.GetType().Name == "StateAttributeMetadata")
                    //    op = ((StateAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options.First(x => x.Value == entity.GetAttributeValue<OptionSetValue>(AttributeLogicalName).Value);

                    //if (rarr.AttributeMetadata.GetType().Name == "StatusAttributeMetadata")
                    //    op = ((StatusAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options.First(x => x.Value == entity.GetAttributeValue<OptionSetValue>(AttributeLogicalName).Value);

                    //if (op != null)
                    //    return string.Format("({0}) {1}", op.Value, op.Label.LocalizedLabels[0].Label);
                    //else
                    //    return string.Format("{0}", "UNKNOWN");

                    return connection.GetOptionSetValue(entity.LogicalName, AttributeLogicalName, entity.GetAttributeValue<OptionSetValue>(AttributeLogicalName).Value);
                    break;

                case "EntityReference":
                    EntityReference er = entity.GetAttributeValue<EntityReference>(AttributeLogicalName);
                    return string.Format("{0} ({1}:{2})", er.Name, er.LogicalName, er.Id);
                    break;

                case "Money":
                    return entity.GetAttributeValue<Money>(AttributeLogicalName).Value.ToString("c");

                default:
                    return entity[AttributeLogicalName].GetType().Name;
                    break;
            }
        }

    }
}
