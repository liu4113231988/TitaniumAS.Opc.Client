﻿using NLog;
using System;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using TitaniumAS.Opc.Client.Common.Wrappers;

namespace TitaniumAS.Opc.Client.Common
{
    public abstract class ComWrapper : IComWrapper
    {
        public static event EventHandler<RpcFailedEventArgs> RpcFailed;
        private const int RPC_S_SERVER_UNAVAILABLE = unchecked((int)0x800706BA);

        public object UserData { get; set; }
        private static readonly ILogger Log = LogManager.GetCurrentClassLogger();

        protected ComWrapper()
        {
        }

        protected ComWrapper(object userData)
        {
            UserData = userData;
        }

        protected virtual void OnRpcFailed(RpcFailedEventArgs args)
        {
            var handler = RpcFailed;
            if (handler != null) handler(this, args);
        }

        public TResult DoComCall<TResult>(object comObject, string methodName, Func<TResult> comAction, params object[] arguments)
        {
            try
            {
                Log.Trace($"Calling '{methodName}' Object: {comObject.GetHashCode()} {ArgumentsToString(arguments)}");
                var result = comAction();
                Log.Trace($"Calling '{methodName}' Object: {comObject.GetHashCode()} {ResultToString(result)}");
                return result;
            }
            catch (InvalidCastException ex) //workaround for correct Com.QueryInterface calls
            {
                if (ex.Message.Contains("0x800706BA"))  //find HRESULT: 0x800706BA - RPC_SERVER_S_UNAVAILABLE
                {
                    var errorString = string.Format("Invalid Cast Error: {0} Method name: {1} Object: {2}", ex, methodName, comObject.GetHashCode());
                    Log.Error(errorString);
                    object userData = UserData;
                    OnRpcFailed(new RpcFailedEventArgs(userData, RPC_S_SERVER_UNAVAILABLE));
                }
                throw;
            }
            catch (COMException ex)
            {
                var errorString = string.Format("COM Error: {0} Method name: {1} Object: {2}", ex, methodName, comObject.GetHashCode());
                Log.Error(errorString);
                if (IsRpcError(ex.ErrorCode))
                {
                    object userData = UserData;
                    if (ex.Message.Contains("0x800706BA"))
                    {
                        OnRpcFailed(new RpcFailedEventArgs(userData, RPC_S_SERVER_UNAVAILABLE));
                    }
                    else
                    {
                        OnRpcFailed(new RpcFailedEventArgs(userData, ex.ErrorCode));
                    }
                }
                throw;
            }
            catch (Exception ex)
            {
                var errorString = string.Format("Error3: {0} Method name: {1} Object: {2}", ex, methodName, comObject.GetHashCode());
                Log.Error(errorString);
                throw;
            }
        }

        private static string ResultToString<TResult>(TResult result)
        {
            var collection = result as IEnumerable;
            if (collection != null)
            {
                return string.Join(",", collection);
            }
            else
            {
                return result.ToString();
            }
        }

        private string ArgumentsToString(object[] arguments)
        {
            var builder = new StringBuilder();
            foreach (var argument in arguments)
            {
                if (argument is IEnumerable)
                {
                    builder.AppendFormat("[{0}]", string.Join(", ", argument));
                }
                else
                {
                    builder.AppendFormat(", {0}", argument);
                }

            }
            if (arguments.Any())
            {
                return string.Format("Arguments: {0}", builder);
            }
            else
            {
                return string.Empty;
            }
        }

        public void DoComCall(object comObject, string methodName, Action comAction, params object[] arguments)
        {
            try
            {
                Log.Trace($"Calling '{methodName}' Object: {comObject.GetHashCode()} {ArgumentsToString(arguments)}");
                comAction();
                Log.Trace($"Success: '{methodName}' Object: {comObject.GetHashCode()}", methodName);
            }
            catch (InvalidCastException ex) //workaround for correct Com.QueryInterface calls
            {
                if (ex.Message.Contains("0x800706BA"))  //find HRESULT: 0x800706BA - RPC_SERVER_S_UNAVAILABLE
                {
                    var errorString = string.Format("Invalid Cast Error: {0} Method name: {1} Object: {2}", ex, methodName, comObject.GetHashCode());
                    Log.Error(errorString);
                    object userData = UserData;
                    OnRpcFailed(new RpcFailedEventArgs(userData, RPC_S_SERVER_UNAVAILABLE));
                }
                throw;
            }
            catch (COMException ex)
            {
                var errorString = string.Format("COM Error: {0} Method name: {1} Object: {2}", ex, methodName, comObject.GetHashCode());
                Log.Error(errorString);
                if (IsRpcError(ex.ErrorCode))
                {
                    object userData = UserData;
                    if (ex.Message.Contains("0x800706BA"))
                    {
                        OnRpcFailed(new RpcFailedEventArgs(userData, RPC_S_SERVER_UNAVAILABLE));
                    }
                    else
                    {
                        OnRpcFailed(new RpcFailedEventArgs(userData, ex.ErrorCode));
                    }
                }
                throw;
            }
            catch (Exception ex)
            {
                var errorString = string.Format("Error: {0} Method name: {1} Object: {2}", ex, methodName, comObject.GetHashCode());
                Log.Error(errorString);
                throw;
            }
        }

        private bool IsRpcError(int errorCode)
        {
            return errorCode == HRESULT.RPC_E_CLIENT_DIED ||
                   errorCode == HRESULT.RPC_E_SERVER_DIED ||
                   errorCode == HRESULT.RPC_E_SERVER_DIED_DNE ||
                   errorCode == RPC_S_SERVER_UNAVAILABLE;
        }
    }
}
