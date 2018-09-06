using System;
using Microsoft.SharePoint.Client;

namespace SPSafeSocketDemo
{
    class Program
    {
        // Change it to your Sharepoint list URL
        static readonly string ListURL = "http://example.com/list";

        static void Main(string[] args)
        {
            // Change to larger loop number if you don't see SocketException
            var loop = 256;

            for (var i = 0; i < loop; i++)
            {
                // Uncomment one of these.
                // * * *
                // 1. Normal execution, without catching SocketException.
                // ExecuteNormally(i);
                // * * *
                // 2. Catch SocketException and see the error code,
                //    if it's a WSAEADDRINUSE error, try again.
                   ExecuteSafely(i);
            }
        }

        static void ExecuteNormally(int batch)
        {
            using (var context = new ClientContext(ListURL))
            {
                AddRandomItems(context, batch);
                context.ExecuteQueryRetry();
            }
        }

        static void ExecuteSafely(int batch)
        {
            using (var context = new ClientContext(ListURL))
            {
                SafeSocketProcess.Execute(
                    context,
                    (ctx) => AddRandomItems(ctx, batch));
            }
        }

        static void AddRandomItems(ClientContext ctx, int batch)
        {
            for (var i = 0; i < 100; i++)
            {
                // Assume that the web has a list named "Announcements". 
                var announcementsList = ctx.Web.Lists.GetByTitle("Announcements");

                // We are just creating a regular list item.
                var itemCreateInfo = new ListItemCreationInformation();
                var newItem = announcementsList.AddItem(itemCreateInfo);
                newItem["Title"] =
                    $"Batch #{batch}, Item #{DateTime.UtcNow.ToString("R")}";
                newItem.Update();
            }
        }
    }
}
