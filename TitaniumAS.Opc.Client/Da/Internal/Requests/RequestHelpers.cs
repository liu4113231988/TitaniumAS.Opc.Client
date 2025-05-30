using System;
using System.Threading;
using NLog;

namespace TitaniumAS.Opc.Client.Da.Internal.Requests
{
    internal static class RequestHelpers
    {
        public static void SetCancellationHandler(CancellationToken token, Action callback)
        {
            token.Register(
                () =>
                {
                    try
                    {
                        callback();
                    }
                    catch (Exception ex)
                    {
                        LogManager.GetCurrentClassLogger().Error("Cancel failed.", ex);
                    }
                }
                );
        }
    }
}