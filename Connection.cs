using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMViewer
{
    public class Connection
    {
        private List<Tuple<string, string, int, string>> _optionSetValues;

        public IOrganizationService OrganizationService { get; set; }

        public Connection(string user,string pass,string url)
        {
            this.OrganizationService = Connector.ConnectToMSCRM(user, pass, url);
            _optionSetValues = new List<Tuple<string, string, int, string>>();
        }

        public string GetOptionSetValue(string Entity, string Attribute, OptionSetValue Value) { return GetOptionSetValue(Entity, Attribute, Value.Value); }

        public string GetOptionSetValue(string Entity, string Attribute, int Value)
        {
            if (!_optionSetValues.Any(x => (x.Item1 == Entity && x.Item2 == Attribute && x.Item3 == Value)))
            {
                RetrieveAttributeRequest rar = new RetrieveAttributeRequest()
                {
                    EntityLogicalName = Entity,
                    LogicalName = Attribute,
                    RetrieveAsIfPublished = true
                };
                RetrieveAttributeResponse rarr = (RetrieveAttributeResponse)OrganizationService.Execute(rar);

                OptionMetadata op = null;
                if (rarr.AttributeMetadata.GetType().Name == "PicklistAttributeMetadata")
                    if (((PicklistAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options.Count > 0)
                        op = ((PicklistAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options.First(x => x.Value == Value);

                if (rarr.AttributeMetadata.GetType().Name == "StateAttributeMetadata")
                    op = ((StateAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options.First(x => x.Value == Value);

                if (rarr.AttributeMetadata.GetType().Name == "StatusAttributeMetadata")
                    op = ((StatusAttributeMetadata)rarr.AttributeMetadata).OptionSet.Options.First(x => x.Value == Value);

                if (op != null)
                    _optionSetValues.Add(new Tuple<string, string, int, string>(Entity, Attribute, Value, op.Label.LocalizedLabels[0].Label));
                else
                    _optionSetValues.Add(new Tuple<string, string, int, string>(Entity, Attribute, Value, "UNKNOWN"));
            }

            Tuple<string, string, int, string> record = _optionSetValues.Find(x => (x.Item1 == Entity && x.Item2 == Attribute && x.Item3 == Value));
            return string.Format("({0}) {1}", record.Item3, record.Item4);
        }

    }
}
