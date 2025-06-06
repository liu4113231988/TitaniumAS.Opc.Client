﻿using System;
using System.Collections.Generic;
using NLog;
using TitaniumAS.Opc.Client.Common.Wrappers;

namespace TitaniumAS.Opc.Client.Common.Internal
{
    internal class OpcServerList2Enumerator : IOpcServerListEnumerator
    {
        private static readonly ILogger Log = LogManager.GetCurrentClassLogger();
        private readonly OpcServerList2 _opcServerList2;

        public OpcServerList2Enumerator(object enumerator)
        {
            _opcServerList2 = new OpcServerList2(enumerator);
        }

        public List<OpcServerDescription> Enumerate(string host, Guid[] categoriesGuids)
        {
            if (_opcServerList2 == null)
                return null;
            try
            {
                using (var enumerator = _opcServerList2.EnumClassesOfCategories(categoriesGuids, null))
                {
                    return EnumerateServers(host, enumerator);
                }
            }
            catch (Exception ex)
            {
                Log.Warn(ex, "Failed to enumerate classes of categories.");
                return null;
            }
        }

        public Guid CLSIDFromProgID(string progId)
        {
            return _opcServerList2.CLSIDFromProgID(progId);
        }

        private List<OpcServerDescription> EnumerateServers(string host, OpcEnumGuid enumGuid)
        {
            var servers = new List<OpcServerDescription>();
            try
            {
                var clsids = new Guid[OpcConfiguration.BatchSize];
                int fetched;

                do
                {
                    fetched = enumGuid.Next(clsids);
                    for (var i = 0; i < fetched; i++)
                    {
                        var description = TryGetServerDescription(host, clsids[i]);
                        if (description != null)
                        {
                            servers.Add(description);
                        }
                    }
                } while (fetched != 0);
            }
            catch (Exception ex)
            {
                Log.Warn(ex, "Failed to enumerate classes of categories.");
            }
            return servers;
        }

        public OpcServerDescription TryGetServerDescription(string host, Guid clsid)
        {
            try
            {
                return GetServerDescription(host, clsid);
            }
            catch (Exception ex)
            {
                Log.Warn(ex, "Failed to get class details.");
                return new OpcServerDescription(host, clsid);
            }
        }

        public OpcServerDescription GetServerDescription(string host, Guid clsid)
        {
            string userType;
            string vendorIndependentProgId;
            string progId = _opcServerList2.GetClassDetails(clsid, out userType, out vendorIndependentProgId);
            return new OpcServerDescription(host, clsid, progId, userType, vendorIndependentProgId);
        }
    }
}