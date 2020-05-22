using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace CRMViewer
{
    class Utility
    {
        private Connection connection;
        private DataTable data;

        public Utility(Connection connection)
        {
            this.connection = connection;
        }

        internal void Begin()
        {
            LoadData();
            Result r = new Result() { Data = data };

            ResultResponse returnCode = r.Draw();

            int index = -1;
            if (int.TryParse(returnCode.Response, out index))
            {
                switch (index)
                {
                    case 1: RecalculateOrderTotal(); break;
                    case 2: BulkPurgeResponsibleParty(); break;
                    case 3: ClearPortalsSessions(); break;
                    case 4: ExportUserRoleToExcel(); break;
                    case 5: ExportEntitiesToExcel(); break;
                    case 6: UnlockQuoteProduct(); break;
                    case 7: UnlockOrderProduct(); break;
                    case 8: ReactivateQuote(); break;
                    default:
                        break;
                }
            }
        }

        private void RecalculateOrderTotal()
        {
            Guid orderid = Guid.Empty;
            string value = string.Empty;

            do
            {
                Console.Write("\r\nPaste Order Guid: ");
                value = Console.ReadLine();
                if (string.IsNullOrEmpty(value)) return;
                if (!Guid.TryParse(value, out orderid))
                    Console.WriteLine("Invalid Guid");
            } while (orderid == Guid.Empty);

            try
            {
                Entity pre = connection.OrganizationService.Retrieve("salesorder", orderid, new ColumnSet(new string[] { "totalamount" }));
                Entity post = new Entity("salesorder", orderid);
                post["totalamount"] = pre["totalamount"];
                connection.OrganizationService.Update(post);
                Console.WriteLine(" Order recalculated");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private void BulkPurgeResponsibleParty()
        {
            Console.Write("\r\nPaste GUID of account to purge as responsible party: ");
            string input = Console.ReadLine();
            Guid accountid = Guid.Empty;
            if (Guid.TryParse(input, out accountid))
            {
                Entity org = connection.OrganizationService.Retrieve("account", accountid, new ColumnSet(true));
                Console.WriteLine(org.GetAttributeValue<string>("name") + " found.  Type 'continue' to purge from every contact.responsibleparty.");
                input = Console.ReadLine();
                if (input == "continue")
                {
                    string fetchXml = string.Format(@"<fetch>
                      <entity name='contact' >
                        <attribute name='contactid' />
                        <attribute name='description' />
                        <filter>
                          <condition attribute='ifma_responsibleparty' operator='eq' value='{0}' />
                        </filter>
                      </entity>
                    </fetch>", accountid);

                    EntityCollection ec = connection.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));
                    Console.WriteLine(ec.Entities.Count().ToString() + " contacts found to remote responsible party");
                    int pos = 1;
                    foreach (Entity e in ec.Entities)
                    {
                        Entity update = new Entity("contact", e.Id);
                        update["ifma_responsibleparty"] = null;
                        update["description"] = e.GetAttributeValue<string>("description") + string.Format("\r\nRemoved {0} as responsible party on {1}", org.GetAttributeValue<string>("name"), DateTime.Now);
                        connection.OrganizationService.Update(update);
                        Util.ShowProgress(pos++, ec.Entities.Count);
                        System.Threading.Thread.Sleep(1);
                    }
                }
            }
        }

        private void ClearPortalsSessions()
        {
            var fetchXml = $@"
                <fetch>
                  <entity name='adx_webformsession' />
                </fetch>";

            EntityCollection ec = connection.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));
            Console.WriteLine(string.Format("\r\nClearing {0} portals sessions", ec.Entities.Count()));
            for (int i = 0; i < ec.Entities.Count; i++)
            {
                connection.OrganizationService.Delete("adx_webformsession", ec.Entities[i].Id);
                Util.ShowProgress(i + 1, ec.Entities.Count);
                System.Threading.Thread.Sleep(10);
            }
        }

        private void ExportUserRoleToExcel()
        {
            List<Entity> results = Util.GetAllUserPermissions(connection.OrganizationService);
            Util.SendToFile(results, results.First().Attributes.Keys.ToArray());
        }

        private void ExportEntitiesToExcel()
        {
            DataTable newDT = new DataTable("Entities");
            newDT.Columns.AddRange(
                new DataColumn[] {
                    new DataColumn("Label", typeof(string)),
                    new DataColumn("LogicalName", typeof(string)),
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
                DataRow newDR = newDT.NewRow();
                newDR[0] = e.DisplayName.LocalizedLabels.Count > 0 ? e.DisplayName.LocalizedLabels.First(x => x.LanguageCode == 1033).Label : e.LogicalName;
                newDR[1] = e.LogicalName;
                newDT.Rows.Add(newDR);
                Util.ShowProgress(pos++, max);
            }

            newDT.DefaultView.Sort = "Label ASC";

            Util.SendToExcel(newDT);
        }

        private void ReactivateQuote()
        {
            Console.Write("\r\n Quote number: ");
            string quotenumber = Console.ReadLine();

            QueryByAttribute qbaQuote = new QueryByAttribute("quote");
            qbaQuote.AddAttributeValue("quotenumber", quotenumber);
            qbaQuote.ColumnSet = new ColumnSet(new string[] { "quoteid", "name" });
            EntityCollection ec = connection.OrganizationService.RetrieveMultiple(qbaQuote);
            if (ec.Entities.Count == 1)
            {
                Entity entity = new Entity("quote", ec.Entities[0].Id);
                entity["statecode"] = new OptionSetValue(0);
                entity["statuscode"] = new OptionSetValue(1);
                connection.OrganizationService.Update(entity);
            }
        }
        private void UnlockQuoteProduct()
        {
            Console.Write("\r\n Quote number: ");
            string quotenumber = Console.ReadLine();

            QueryByAttribute qbaQuote = new QueryByAttribute("quote");
            qbaQuote.AddAttributeValue("quotenumber", quotenumber);
            qbaQuote.ColumnSet = new ColumnSet(new string[] { "quoteid", "name" });
            EntityCollection ec = connection.OrganizationService.RetrieveMultiple(qbaQuote);
            if (ec.Entities.Count == 1)
            {
                Console.WriteLine("looking at quote: " + ec.Entities[0].GetAttributeValue<string>("name"));
                QueryByAttribute qbaQuoteDetail = new QueryByAttribute("quotedetail");
                qbaQuoteDetail.AddAttributeValue("quoteid", ec.Entities[0].Id);
                qbaQuoteDetail.ColumnSet = new ColumnSet(new string[] { "quotedetailid", "productid", "quantity" });
                EntityCollection ec1 = connection.OrganizationService.RetrieveMultiple(qbaQuoteDetail);

                int i = 1;
                foreach (Entity e in ec1.Entities)
                    Console.WriteLine(string.Format("{0}. {1} x {2}", i++, e.GetAttributeValue<decimal>("quantity"), e.GetAttributeValue<EntityReference>("productid").Name));


                string sel = Console.ReadLine();
                if (int.TryParse(sel, out i))
                {
                    Entity qdChange = new Entity("quotedetail", ec1.Entities[i - 1].Id);
                    qdChange["ispriceoverridden"] = true;
                    connection.OrganizationService.Update(qdChange);
                }
            }
        }

        private void UnlockOrderProduct()
        {
            Console.Write("\r\n Order number: ");
            string quotenumber = Console.ReadLine();

            QueryByAttribute qbaSalesOrder = new QueryByAttribute("salesorder");
            qbaSalesOrder.AddAttributeValue("ordernumber", quotenumber);
            qbaSalesOrder.ColumnSet = new ColumnSet(new string[] { "salesorderid", "name" });
            EntityCollection ec = connection.OrganizationService.RetrieveMultiple(qbaSalesOrder);
            if (ec.Entities.Count == 1)
            {
                Console.WriteLine("looking at salesorder: " + ec.Entities[0].GetAttributeValue<string>("name"));
                QueryByAttribute qbaSalesOrderDetail = new QueryByAttribute("salesorderdetail");
                qbaSalesOrderDetail.AddAttributeValue("salesorderid", ec.Entities[0].Id);
                qbaSalesOrderDetail.ColumnSet = new ColumnSet(new string[] { "salesorderdetailid", "productid", "quantity" });
                EntityCollection ec1 = connection.OrganizationService.RetrieveMultiple(qbaSalesOrderDetail);

                int i = 1;
                foreach (Entity e in ec1.Entities)
                    Console.WriteLine(string.Format("{0}. {1} x {2}", i++, e.GetAttributeValue<decimal>("quantity"), e.GetAttributeValue<EntityReference>("productid").Name));


                string sel = Console.ReadLine();
                if (int.TryParse(sel, out i))
                {
                    Entity qdChange = new Entity("salesorder", ec.Entities[0].Id);
                    qdChange["ispricelocked"] = false;
                    connection.OrganizationService.Update(qdChange);

                    qdChange = new Entity("salesorderdetail", ec1.Entities[i - 1].Id);
                    qdChange["ispriceoverridden"] = true;
                    connection.OrganizationService.Update(qdChange);
                }
            }
        }

        private void LoadData()
        {
            data = new DataTable();
            data.Columns.Add(new DataColumn("Utility", typeof(string)));

            data.Rows.Add(new string[] { "Recalculate Order Total" });
            data.Rows.Add(new string[] { "Bluk Purge Responsible Party" });
            data.Rows.Add(new string[] { "Clear Portals WebForm Sessions" });
            data.Rows.Add(new string[] { "Export UserRolePermission to csv" });
            data.Rows.Add(new string[] { "Export Entities To Excel" });
            data.Rows.Add(new string[] { "Unlock Quote Item Record" });
            data.Rows.Add(new string[] { "Unlock SalesOrder Item Record" });
            data.Rows.Add(new string[] { "Reactivate Quote" });
        }
    }
}
