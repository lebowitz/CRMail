using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.TeamFoundation.Client;

namespace crmail
{
    internal class TfServerUtil
    {
        // Methods
        public TfServerUtil()
        {
        }

        public static TeamFoundationServer GetServer(string serverName)
        {
            TeamFoundationServer server1 = null;
            if (string.IsNullOrEmpty(serverName))
            {
                throw new ArgumentException("serverName cannot be null or empty");
            }
            if (serverName != null)
            {
                server1 = TeamFoundationServerFactory.GetServer(serverName);
            }
            if (server1 == null)
            {
                throw new Exception("Cannot connect to Team Foundation Server");
            }
            return server1;
        }

    }
 

}
