using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.Delta;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace DriveTrackChanges
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Track changes");

            var driveId = "Me";
            var driveItemId = "root";

            var cts = new CancellationTokenSource();
            cts.CancelAfter(TimeSpan.FromMinutes(10));

            var credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
            {
                ClientId = "<client_id>",
                TenantId = "<tenant_id>"
            });

            var graphClient = new GraphServiceClient(credential);
            var latestTokenResponse = await graphClient.Drives[driveId].Items[driveItemId].DeltaWithToken("latest").GetAsDeltaWithTokenGetResponseAsync(cancellationToken: cts.Token);
            var deltaLink = latestTokenResponse.OdataDeltaLink;
            var timestamp = DateTime.UtcNow;
            var trackingTask = Task.Run(async () =>
            {
                while (true)
                {
                    // change delay if you want to track changes more frequently
                    await Task.Delay(TimeSpan.FromMinutes(1), cts.Token);
                    var response = await graphClient.Drives[driveId].Items[driveItemId].Delta.WithUrl(deltaLink).GetAsDeltaGetResponseAsync(rc =>
                    {
                        // exclude parent folders from the response
                        rc.Headers.Add("Prefer", "deltaExcludeParent");
                    }, cts.Token);

                    var pageIterator = PageIterator<DriveItem, DeltaGetResponse>.CreatePageIterator(graphClient, response, (item) =>
                    {
                        // ignore root item
                        if (item.Name == "root")
                        {
                            return true;
                        }

                        if (item.Deleted != null)
                        {
                            Console.WriteLine($"Deleted item: {item.Id}, state: {item.Deleted.State}");
                            return true;
                        }
                        if (item.CreatedDateTime > timestamp)
                        {
                            Console.WriteLine($"New item: {item.Id}, created: {item.CreatedDateTime}, name: {item.Name}");
                            return true;
                        }

                        Console.WriteLine($"Modified item: {item.Id}, modified: {item.LastModifiedDateTime}, name: {item.Name}");
                        return true;
                    },
                    rc =>
                    {
                        // exclude parent folders from the response
                        rc.Headers.Add("Prefer", "deltaExcludeParent");
                        return rc;
                    });

                    await pageIterator.IterateAsync(cts.Token);
                    deltaLink = pageIterator.Deltalink;
                    timestamp = DateTime.UtcNow;
                }
            });

            try
            {
                await trackingTask;
            }
            catch (TaskCanceledException)
            {
                Console.WriteLine("Task was canceled.");
            }
            catch(ODataError error)
            {
                Console.WriteLine($"Error: {error.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            try
            {
                trackingTask.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
