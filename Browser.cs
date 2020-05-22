using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace CRMViewer
{
    class Browser
    {
        private Connection connection;
        private Stack<Result> data;

        private Dictionary<string, DataTable> cache;

        public Browser(Connection connection)
        {
            this.connection = connection;
        }

        public void Begin()
        {
            data = new Stack<Result>();
            cache = new Dictionary<string, DataTable>();

            data.Push(new Result(LoadEntityData()));

            ResultResponse returnCode = new ResultResponse();
            while (data.Count > 0)
            {
                if (returnCode.ConsoleKey != ConsoleKey.F1)
                    returnCode = data.Peek().Draw();
                else
                    returnCode = data.Peek().Draw("F2: Read   F3: Search   F4: Open   F5: Refresh   Selection: ");

                string tableType = data.Peek().Data.TableName;
                if (tableType.Contains(" ")) tableType = tableType.Substring(0, tableType.IndexOf(" "));

                switch (returnCode.ConsoleKey)
                {
                    case ConsoleKey.Escape: data.Pop(); break;

                    case ConsoleKey.F1:
                        break;

                    case ConsoleKey.F2: //read
                        if (tableType == "Entity")
                        {
                            string entityToRead = data.Peek().Data.TableName.Substring(7).Replace("(cached)", "");
                            Console.Write("\r\nGUID, index or Enter: ");
                            string sGuid = Console.ReadLine();
                            Guid recordID = Guid.Empty;
                            int recordNumber = 0;
                            if (Guid.TryParse(sGuid, out recordID))
                                data.Push(new Result(GetRecord(entityToRead, recordID)) { LogicalName = entityToRead });
                            else if (int.TryParse(sGuid, out recordNumber))
                                data.Push(new Result(GetRecord(entityToRead, Math.Max(0, recordNumber))) { LogicalName = entityToRead });
                            else
                                data.Push(new Result(GetRecord(entityToRead, 1)) { LogicalName = entityToRead });
                        }
                        break;

                    case ConsoleKey.F3: //search
                        if (returnCode.Modifiers == 0 && tableType == "Entity")
                        {
                            Console.Write("\r\nAttribute index to search: ");
                            string searchAttributeString = Console.ReadLine();
                            int searchAttributeIndex = -1;
                            if (int.TryParse(searchAttributeString, out searchAttributeIndex))
                            {
                                string searchEntity = data.Peek().Data.TableName.Substring(7).Replace("(cached)", "");
                                string searchAttribute = (string)data.Peek().Data.DefaultView[searchAttributeIndex - 1]["LogicalName"];
                                string searchAttributeDataType = (string)data.Peek().Data.DefaultView[searchAttributeIndex - 1]["DataType"];
                                Console.WriteLine("  searching attribute {0}({1})", searchAttribute, searchAttributeDataType);
                                Console.Write("Value to search for: ");
                                string searchValue = Console.ReadLine();
                                RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest()
                                {
                                    EntityFilters = EntityFilters.Attributes,
                                    LogicalName = searchEntity
                                };

                                RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)connection.OrganizationService.Execute(retrieveEntityRequest);
                                EntityMetadata em = (EntityMetadata)retrieveEntityResponse.Results["EntityMetadata"];
                                string primaryIDAttribute = em.PrimaryIdAttribute;
                                string primaryIDName = em.PrimaryNameAttribute;
                                string filterOperator = "like";
                                string fetchXml = string.Format(@"
                                    <fetch>
                                      <entity name='{0}' >
                                        <attribute name='{1}' />
                                        <attribute name='{2}' />
                                        <attribute name='{3}' />
                                        <filter>
                                          <condition attribute='{3}' operator='{4}' value='{5}' />
                                        </filter>
                                      </entity>
                                    </fetch>", searchEntity, primaryIDAttribute, primaryIDName, searchAttribute, filterOperator, searchValue);

                                DataTable retVal = new DataTable() { TableName = "Search" };
                                retVal.Columns.AddRange(
                                    new DataColumn[] {
                                        new DataColumn(primaryIDAttribute, typeof(Guid)),
                                        new DataColumn(primaryIDName, typeof(string)),
                                        new DataColumn(searchAttribute, typeof(object)),
                                        new DataColumn("Entity", typeof(string))
                                    });

                                EntityCollection entityCollection = connection.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));
                                if (entityCollection.Entities.Count > 0)
                                {
                                    var attDataType = entityCollection.Entities[0][searchAttribute].GetType();
                                    foreach (Entity entity in entityCollection.Entities)
                                    {
                                        DataRow dataRow = retVal.NewRow();
                                        if (entity.Contains(primaryIDAttribute)) dataRow[primaryIDAttribute] = entity[primaryIDAttribute];
                                        if (entity.Contains(primaryIDName)) dataRow[primaryIDName] = entity[primaryIDName];
                                        if (entity.Contains(searchAttribute)) dataRow[searchAttribute] = entity[searchAttribute];
                                        dataRow["Entity"] = searchEntity;
                                        retVal.Rows.Add(dataRow);
                                    }
                                }
                                else
                                {
                                    DataRow dataRow = retVal.NewRow();
                                    dataRow[primaryIDAttribute] = "No";
                                    dataRow[primaryIDName] = "record";
                                    dataRow[searchAttribute] = "found";
                                    retVal.Rows.Add(dataRow);
                                }

                                data.Push(new Result(retVal, string.Format("Search on {0} for {1} in {2}", searchEntity, searchValue, searchAttribute)));
                            }
                        }
                        break;

                    case ConsoleKey.F4: //open
                        string url = ((OrganizationServiceProxy)connection.OrganizationService).ServiceConfiguration.CurrentServiceEndpoint.Address.Uri.Host;
                        int port = ((OrganizationServiceProxy)connection.OrganizationService).ServiceConfiguration.CurrentServiceEndpoint.Address.Uri.Port;
                        string target = string.Empty;

                        switch (tableType)
                        {
                            case "Entity":
                                string ename = data.Peek().LogicalName; // Data.TableName.Substring(7).Replace("(cached)", "");
                                target = string.Format("https://{0}:{1}/main.aspx?etn={2}&pagetype=entitylist"
                                    , url, port, ename);

                                break;

                            case "Record":
                                //https://test-ifma.crm.dynamics.com/main.aspx?etn=bpt_chapter&pagetype=entityrecord&id=f0adaa03-4d26-e511-80db-c4346bacb92c
                                string[] vals = data.Peek().Data.TableName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                string etn = data.Peek().LogicalName; // vals[1];
                                Guid id = Guid.Parse(vals[2]);

                                target = string.Format("https://{0}:{1}/main.aspx?etn={2}&pagetype=entityrecord&id={3}"
                                    , url, port, etn, id);
                                break;

                            default:
                                break;
                        }
                        //open browser
                        System.Diagnostics.Process.Start(target);
                        break;

                    case ConsoleKey.F5: //refresh
                        string tn = data.Peek().Data.TableName.Substring(7).Replace("(cached)", "");
                        data.Pop();
                        data.Push(new Result(GetEntity(tn, true)));
                        break;




                    default:
                        int selection = -1;

                        if (int.TryParse(returnCode.Response, out selection))
                            switch (tableType)
                            {
                                case "Search":
                                    string targetEntityx = (string)data.Peek().Data.DefaultView[selection - 1]["Entity"];
                                    Guid recordIdx = (Guid)data.Peek().Data.DefaultView[selection - 1][data.Peek().Data.Columns[0].ColumnName];
                                    data.Push(new Result(GetRecord(targetEntityx, recordIdx)));
                                    break;

                                case "Entities":
                                    if (selection <= data.First().Data.Rows.Count)
                                    {
                                        string entityName = (string)data.First().Data.DefaultView[selection - 1]["LogicalName"];

                                        DataTable newEntity = GetEntity(entityName);
                                        data.Push(new Result(newEntity) { LogicalName = (string)data.First().Data.DefaultView[selection - 1]["LogicalName"] });
                                    }
                                    break;

                                case "Entity":
                                    DataRowView dr = data.Peek().Data.DefaultView[selection - 1];
                                    switch ((string)dr["DataType"])
                                    {
                                        case "Picklist":
                                        case "Status":
                                        case "State":
                                            string targetPicklist = (string)dr["LogicalName"];
                                            string logName = data.Peek().Data.TableName.Substring(7).Replace("(cached)", "");
                                            data.Push(new Result(GetPicklist(logName, targetPicklist)));
                                            break;

                                        case "Lookup":
                                            string targetEntity = (string)dr["MetaData"];
                                            data.Push(new Result(GetEntity(targetEntity)){ LogicalName = targetEntity });
                                            break;

                                        default:
                                            break;
                                    }
                                    break;

                                default: break;
                            }

                        break;
                }

            }
        }

        DataTable GetRecord(string entityToRead, int recordNumber)
        {
            var fetchXml = string.Format(@"
            <fetch count='1' page='{1}'>
              <entity name='{0}'>
              </entity>
            </fetch>", entityToRead, recordNumber);

            EntityCollection ec = null;
            try
            {
                ec = connection.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));

                if (ec != null && ec.Entities.Count > 0)
                    return GetRecord(entityToRead, ec[0].Id);
                else
                    return GetRecord(entityToRead, Guid.NewGuid()); //will fail and return record not found
            }
            catch (Exception)
            {
                DataTable retVal = new DataTable();
                retVal.Columns.AddRange(
                    new DataColumn[]
                    {
                        new DataColumn("DisplayName", typeof(string)),
                        new DataColumn("LogicalName", typeof(string)),
                        new DataColumn("DataType", typeof(string)),
                        new DataColumn("Value", typeof(string))
                    });

                DataRow dr = retVal.NewRow();
                dr[0] = "Error";
                dr[1] = "in";
                dr[2] = "query.";
                retVal.Rows.Add(dr);
                return retVal;
            }
        }


        DataTable GetRecord(string entityToRead, Guid recordID)
        {
            DataTable retVal = new DataTable() { TableName = string.Format("Record {0} {{{1}}}", entityToRead, recordID) };

            retVal.Columns.AddRange(
                new DataColumn[]
                {
                                new DataColumn("DisplayName", typeof(string)),
                                new DataColumn("LogicalName", typeof(string)),
                                new DataColumn("DataType", typeof(string)),
                                new DataColumn("Value", typeof(string))
                });

            var fetchXml = string.Format($@"
            <fetch>
              <entity name='{{0}}'>
                <filter>
                  <condition attribute='{{0}}id' operator='eq' value='{{1}}'/>
                </filter>
              </entity>
            </fetch>", entityToRead, recordID);

            EntityCollection ec = connection.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));
            if (ec == null || ec.Entities.Count <= 0)
            {
                DataRow dr = retVal.NewRow();
                dr[0] = "Record";
                dr[1] = "not";
                dr[2] = "found.";
                retVal.Rows.Add(dr);
            }
            else
            {
                Entity entity = ec[0];
                DataTable attributes = GetEntity(entityToRead);
                //////// do sorting

                foreach (DataRow edr in attributes.DefaultView.Table.Rows)
                {
                    DataRow dr = retVal.NewRow();
                    dr[0] = edr[0];
                    dr[1] = edr[1];
                    dr[2] = edr[3];
                    if (entity.Contains((string)edr[1]))
                        dr[3] = Util.EntityValueFormat(entity, (string)edr[1], connection);
                    //else
                    //  dr[2] = "<NULL>";
                    retVal.Rows.Add(dr);
                }
            }

            return retVal;
        }

        private DataTable LoadEntityData()
        {
            DataTable newDT = new DataTable("Entities");
            newDT.Columns.AddRange(
                new DataColumn[] {
                    new DataColumn("LogicalName", typeof(string)),
                    new DataColumn("Label", typeof(string)),
                });

            RetrieveAllEntitiesRequest retrieveAllEntitiesRequest = new RetrieveAllEntitiesRequest()
            {
                EntityFilters = EntityFilters.Entity,
                RetrieveAsIfPublished = true
            };

            RetrieveAllEntitiesResponse retrieveAllEntitiesResponse = (RetrieveAllEntitiesResponse)connection.OrganizationService.Execute(retrieveAllEntitiesRequest);

            IEnumerable<EntityMetadata> subset = retrieveAllEntitiesResponse.EntityMetadata;

            int pos = 1;
            int max = subset.Count();
            Console.WriteLine(string.Format("\r\nLoading {0} entities", max));
            foreach (EntityMetadata e in subset)
            {
                if (e.IsCustomizable.Value && e.DisplayName.LocalizedLabels.Count > 0)
                {
                    DataRow newDR = newDT.NewRow();
                    newDR[0] = e.LogicalName;
                    newDR[1] = e.DisplayName.LocalizedLabels.Count > 0 ? e.DisplayName.LocalizedLabels.First(x => x.LanguageCode == 1033).Label : e.LogicalName;
                    newDT.Rows.Add(newDR);
                    Util.ShowProgress(pos++, max);
                }
            }

            newDT.DefaultView.Sort = "LogicalName ASC";

            cache["Entities"] = newDT;

            return newDT;
        }

        public DataTable GetEntity(string EntityLogicalName) { return GetEntity(EntityLogicalName, false); }
        public DataTable GetEntity(string EntityLogicalName, bool RefreshCache)
        {
            string cachePath = string.Format("{0}\\{1}.cache", ((OrganizationServiceProxy)connection.OrganizationService).ServiceConfiguration.CurrentServiceEndpoint.Address.Uri.Host, EntityLogicalName);

            if (RefreshCache)
            {
                if (cache.ContainsKey(EntityLogicalName))
                    cache.Remove(EntityLogicalName);

                if (File.Exists(cachePath))
                    File.Delete(cachePath);
            }


            if (!Directory.Exists(Path.GetDirectoryName(cachePath)))
                Directory.CreateDirectory(Path.GetDirectoryName(cachePath));
            Console.Write(" loading " + EntityLogicalName);
            DataTable retVal = null;
            if (cache.ContainsKey(EntityLogicalName))
            {
                retVal = cache[EntityLogicalName];
                if (!retVal.TableName.Contains("(cached)"))
                    retVal.TableName += "(cached)";
            }
            else if (File.Exists(cachePath))
            {
                DataTable loader = new DataTable();
                loader.ReadXml(cachePath);
                cache[EntityLogicalName] = loader;
                return GetEntity(EntityLogicalName);
            }
            else
            {
                retVal = new DataTable() { TableName = "Entity " + EntityLogicalName };
                retVal.Columns.AddRange(
                    new DataColumn[] {
                    new DataColumn("DisplayName", typeof(string)),
                    new DataColumn("LogicalName", typeof(string)),
                    new DataColumn("Required", typeof(string)),
                    new DataColumn("DataType", typeof(string)),
                    new DataColumn("MetaData", typeof(string))
                    });

                RetrieveEntityRequest request = new RetrieveEntityRequest()
                {
                    EntityFilters = EntityFilters.Attributes,
                    RetrieveAsIfPublished = true,
                    LogicalName = EntityLogicalName
                };
                RetrieveEntityResponse response = (RetrieveEntityResponse)connection.OrganizationService.Execute(request);

                int pos = 1;
                int max = response.EntityMetadata.Attributes.Count();
                Console.WriteLine(string.Format(" with {0} attributes", max));

                string primaryKey = EntityLogicalName + "id";
                foreach (AttributeMetadata am in response.EntityMetadata.Attributes)
                {
                    if (am.DisplayName.LocalizedLabels.Count > 0)
                    {
                        DataRow newDR = retVal.NewRow();
                        newDR["DisplayName"] = am.DisplayName.LocalizedLabels[0].Label;
                        newDR["LogicalName"] = am.LogicalName;
                        newDR["DataType"] = am.AttributeType.ToString();
                        switch (am.RequiredLevel.Value)
                        {
                            case AttributeRequiredLevel.None: newDR["Required"] = string.Empty; break;
                            case AttributeRequiredLevel.SystemRequired: newDR["Required"] = "Req'd"; break;
                            case AttributeRequiredLevel.ApplicationRequired: newDR["Required"] = "Req'd"; break;
                            case AttributeRequiredLevel.Recommended: newDR["Required"] = string.Empty; break;
                            //case AttributeRequiredLevel.Recommended: newDR["Required"] = "Rec'd"; break;
                            default:
                                break;
                        }
                        switch (am.AttributeType.ToString())
                        {
                            case "Lookup":
                                RetrieveAttributeRequest rar = new RetrieveAttributeRequest()
                                {
                                    EntityLogicalName = EntityLogicalName,
                                    LogicalName = am.LogicalName,
                                    RetrieveAsIfPublished = true
                                };
                                RetrieveAttributeResponse rarr = (RetrieveAttributeResponse)connection.OrganizationService.Execute(rar);
                                newDR["MetaData"] = string.Join(", ", ((LookupAttributeMetadata)rarr.AttributeMetadata).Targets);
                                break;

                            case "Picklist":
                                RetrieveAttributeRequest rarpl = new RetrieveAttributeRequest()
                                {
                                    EntityLogicalName = EntityLogicalName,
                                    LogicalName = am.LogicalName,
                                    RetrieveAsIfPublished = true
                                };
                                RetrieveAttributeResponse rarplr = (RetrieveAttributeResponse)connection.OrganizationService.Execute(rarpl);
                                newDR["MetaData"] = ((PicklistAttributeMetadata)rarplr.AttributeMetadata).OptionSet.Name;
                                break;

                            default:
                                newDR["MetaData"] = string.Empty;
                                break;
                        }
                        retVal.Rows.Add(newDR);
                    }

                    Util.ShowProgress(pos++, max);
                }

                cache[EntityLogicalName] = retVal;
                retVal.WriteXml(cachePath, XmlWriteMode.WriteSchema);
            }

            retVal.DefaultView.Sort = "DisplayName ASC";
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
