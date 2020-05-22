using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMViewer
{
    class Browse
    {
        private IOrganizationService OrganizationService;
        private DataTable entities;
        private List<KeyValuePair<string, DataTable>> entityDetails;

        public Browse(IOrganizationService organizationService)
        {
            OrganizationService = organizationService;

        }

        private DataTable LoadEntities()
        {
            DataTable retVal = new DataTable();






            return retVal;
        }
    }
}
