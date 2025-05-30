﻿using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NLog;
using IConnectionPoint = System.Runtime.InteropServices.ComTypes.IConnectionPoint;

namespace TitaniumAS.Opc.Client.Common.Internal
{
    internal class ConnectionPoint<T> : IDisposable where T : class
    {
        private static readonly ILogger Log = LogManager.GetCurrentClassLogger();
        private readonly T _sink;
        private IConnectionPoint _connectionPoint;
        private int? _cookie;

        public ConnectionPoint(T sink)
        {
            _sink = sink;
        }

        public bool IsConnected
        {
            get { return _connectionPoint != null && _cookie != null; }
        }

        public T Sink
        {
            get { return _sink; }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            Disconnect();
        }

        ~ConnectionPoint()
        {
            Dispose(false);
        }

        public void TryConnect(object comServer)
        {
            try
            {
                Connect(comServer);
            }
            catch (Exception ex)
            {
                Log.Error($"Failed to connect to '{ex}' connection point.");
            }
        }

        /// <exception cref="InvalidOperationException">Already attached to IOPCShutdown connection point.</exception>
        public void Connect(object comServer)
        {
            if (IsConnected)
                throw new InvalidOperationException("Already attached to the connection point.");

            var connectionPointContainer = (IConnectionPointContainer)comServer;

            var riid = typeof(T).GUID;
            connectionPointContainer.FindConnectionPoint(ref riid, out _connectionPoint);
            int cookie;
            _connectionPoint.Advise(_sink, out cookie);
            _cookie = cookie;
        }

        public void Disconnect()
        {
            try
            {
                if (_connectionPoint == null)
                    return;

                if (_cookie.HasValue)
                {
                    // TODO: Fix: hangs when disposing.
                    _connectionPoint.Unadvise(_cookie.Value);
                    _cookie = null;
                }

                Marshal.ReleaseComObject(_connectionPoint);
                _connectionPoint = null;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to unsubscribe callback.");
            }
        }
    }
}