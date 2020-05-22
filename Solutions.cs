using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMViewer
{
    class Solutions
    {
        private Connection connection;
        private Stack<Result> data;

        public Solutions(Connection connection)
        {
            this.connection = connection;
        }

        public void Begin()
        {
            data = new Stack<Result>();
            data.Push(LoadData());

            while (true)
            {
                bool breakout = false;

                ResultResponse returnCode = data.Peek().Draw();

                switch (returnCode.ConsoleKey)
                {
                    case ConsoleKey.Escape : breakout = true; break;

                    default:
                        int selection = -1;

                        if (int.TryParse(returnCode.Response, out selection))
                        {
                            switch (data.Count)
                            {
                                case 1:
                                    Console.Write("\r\nExporting solution? ([U]nmanaged/[M]anaged/[C]ancel): ");
                                    string answer = Console.ReadLine();
                                    if (string.IsNullOrEmpty(answer) || answer.ToUpper().Trim().StartsWith("U") || answer.ToUpper().Trim().StartsWith("M"))
                                    {
                                        ExportSolutionRequest exportSolutionRequest = new ExportSolutionRequest();
                                        exportSolutionRequest.Managed = answer.ToUpper().Trim().StartsWith("M");
                                        exportSolutionRequest.SolutionName = (string)data.Peek().Data.Rows[selection - 1]["Logical Name"];// solution.UniqueName;

                                        ExportSolutionResponse exportSolutionResponse = null;
                                        try
                                        {
                                            exportSolutionResponse = (ExportSolutionResponse)connection.OrganizationService.Execute(exportSolutionRequest);
                                            byte[] exportXml = exportSolutionResponse.ExportSolutionFile;
                                            string filename = (string)data.Peek().Data.Rows[selection - 1]["Logical Name"] + (exportSolutionRequest.Managed ? "-managed" : "-unmanaged") + ".zip";
                                            File.WriteAllBytes(filename, exportXml);

                                            Console.WriteLine("Solution exported to {0}.", filename);
                                            Console.WriteLine("\r\nEnter to continue");
                                            Console.ReadLine();
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine(ex.Message);

                                            Console.WriteLine("\r\nEnter to continue");
                                            Console.ReadLine();
                                        }
                                    }
                                    break;

                                default:
                                    break;
                            }
                        }
                        break;
                }

                if (breakout) break;

            }
        }

        private Result LoadData()
        {
            Result retVal = new Result();
            retVal.Data = new DataTable();
            retVal.Data.Columns.AddRange(
                 new DataColumn[]
                 {
                     new DataColumn("Solution",typeof(string)),
                     new DataColumn("Logical Name",typeof(string)),
                     new DataColumn("Published By", typeof(string)),
                     new DataColumn("Updated On", typeof(DateTime)),
                 }
                );

            string fetchXml = @"<fetch>
                                  <entity name='solution' >
                                    <attribute name='friendlyname' />
                                    <attribute name='publisherid' />
                                    <attribute name='uniquename' />
                                    <attribute name='modifiedon' />
                                    <attribute name='ismanaged' />
                                    <attribute name='isvisible' />
                                    <filter>
                                      <condition attribute='isvisible' operator='eq' value='1' />
                                    </filter>
                                    <order attribute='friendlyname' />
                                  </entity>
                                </fetch>";
            EntityCollection ec = connection.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));

            foreach (Entity entity in ec.Entities)
            {

                DataRow dr = retVal.Data.NewRow();
                dr["Solution"] = entity.GetAttributeValue<string>("friendlyname");
                dr["Logical Name"] = entity.GetAttributeValue<string>("uniquename");
                dr["Published By"] = entity.GetAttributeValue<EntityReference>("publisherid").Name;
                dr["Updated On"] = entity.GetAttributeValue<DateTime>("modifiedon");
                retVal.Data.Rows.Add(dr);
            }

            return retVal;
        }

        public DataTable GetPicklist(string EntityLogicalName, string AttributeLogicalName)
        {
            DataTable retVal = new DataTable() { TableName = "Picklist " + AttributeLogicalName };

            retVal.Columns.AddRange(
                new DataColumn[] {
                    new DataColumn("Value", typeof(int)),
                    new DataColumn("Label", typeof(string))
                });

            //need to detmine global or local
            RetrieveAttributeRequest rar = new RetrieveAttributeRequest()
            {
                EntityLogicalName = EntityLogicalName,
                LogicalName = AttributeLogicalName,
                RetrieveAsIfPublished = true
            };
            RetrieveAttributeResponse rarr = (RetrieveAttributeResponse)connection.OrganizationService.Execute(rar);

            if (rarr.AttributeMetadata.GetType().Name == "PicklistAttributeMetadata")
                foreach (OptionMetadata om in ((PicklistAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options)
                {
                    DataRow dr = retVal.NewRow();
                    dr["Value"] = om.Value ?? 0;
                    dr["Label"] = om.Label.LocalizedLabels[0].Label;
                    retVal.Rows.Add(dr);
                }

            else if (rarr.AttributeMetadata.GetType().Name == "StateAttributeMetadata")
                foreach (OptionMetadata om in ((StateAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options)
                {
                    DataRow dr = retVal.NewRow();
                    dr["Value"] = om.Value ?? 0;
                    dr["Label"] = om.Label.LocalizedLabels[0].Label;
                    retVal.Rows.Add(dr);
                }

            else if (rarr.AttributeMetadata.GetType().Name == "StatusAttributeMetadata")
                foreach (OptionMetadata om in ((StatusAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options)
                {
                    DataRow dr = retVal.NewRow();
                    dr["Value"] = om.Value ?? 0;
                    dr["Label"] = om.Label.LocalizedLabels[0].Label;
                    retVal.Rows.Add(dr);
                }

            else
                throw new Exception("unknown picklist");

            return retVal;
        }
    }
}