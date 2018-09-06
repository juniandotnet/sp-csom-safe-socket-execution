using System;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using static Microsoft.SharePoint.Client.ClientContextExtensions;

namespace SPSafeSocketDemo
{
    public static class SafeSocketProcess
    {
        public static void Execute(
            ClientContext clientContext,
            Action<ClientContext> process,
            int retryCount = 10,
            int delay = 500,
            string userAgent = null)
        {
            int retryAttempts = 0;
            int backoffInterval = delay;

            if (retryCount <= 0)
                throw new ArgumentException("Provide a retry count greater than zero.");

            if (delay <= 0)
                throw new ArgumentException("Provide a delay greater than zero.");

            // Do while retry attempt is less than retry count
            while (retryAttempts < retryCount)
            {
                try
                {
                    // Clone client context and execute desired process
                    using (var ctx = clientContext.Clone(clientContext.Url))
                    {
                        process?.Invoke(ctx);
                        ctx.ExecuteQueryRetry(retryCount, delay, userAgent);
                    }
                    return;
                }
                catch (Exception ex)
                {
                    // Check if it's a WSAEADDRINUSE socket error
                    if (ex is WebException webex
                        && webex.InnerException is SocketException sockex
                        && sockex?.SocketErrorCode == SocketError.AddressAlreadyInUse)
                    {
                        // Add delay for retry
                        Task.Delay(backoffInterval).Wait();

                        // Add to retry count and increase delay.
                        retryAttempts++;
                        backoffInterval = backoffInterval * 2;
                    }
                    else
                    {
                        // Forward exception if it's not a WSAEADDRINUSE socket error
                        throw;
                    }
                }
            }

            // Throw an exception when retry count has been attempted.
            throw new MaximumRetryAttemptedException(
                $"Maximum retry attempts {retryCount}, has been attempted.");
        }
    }
}
