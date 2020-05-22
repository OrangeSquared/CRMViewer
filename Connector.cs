using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace CRMViewer
{
    static class Connector
    {
        public static IOrganizationService ConnectToMSCRM(string UserName, string Password, string SoapOrgServiceUri)
        {
            IOrganizationService _service = null;
            ClientCredentials credentials = new ClientCredentials();
            credentials.UserName.UserName = UserName;
            credentials.UserName.Password = Password;
            Uri serviceUri = new Uri(SoapOrgServiceUri);
            OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
            proxy.Timeout = new TimeSpan(1, 0, 0);
            proxy.EnableProxyTypes();
            _service = (IOrganizationService)proxy;

            return _service;
        }
    }
}
