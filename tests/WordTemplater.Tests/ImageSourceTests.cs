using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    public class ImageSourceTests
    {
        [Fact]
        public async Task Export_ImageFromHttpUrl_EmbedsDownloadedBytes()
        {
            var png = File.ReadAllBytes(TestPaths.Template("Rectangle.png"));
            int port = FreeLoopbackPort();
            int requests = 0;
            string? lastPath = null;

            var listener = new HttpListener();
            listener.Prefixes.Add($"http://localhost:{port}/");
            listener.Start();
            var server = Task.Run(async () =>
            {
                while (true)
                {
                    HttpListenerContext ctx;
                    try
                    {
                        ctx = await listener.GetContextAsync();
                    }
                    catch (Exception) when (!listener.IsListening)
                    {
                        return;
                    }
                    lastPath = ctx.Request.Url?.AbsolutePath;
                    Interlocked.Increment(ref requests);
                    ctx.Response.ContentType = "image/png";
                    ctx.Response.ContentLength64 = png.Length;
                    ctx.Response.OutputStream.Write(png, 0, png.Length);
                    ctx.Response.Close();
                }
            });

            try
            {
                using var templateStream = new MemoryStream(TestTemplates.Build(TestTemplates.P(TestTemplates.F("Url:image()"))));
                using var template = new WordTemplate(templateStream);
                const int exports = 3;
                for (int i = 0; i < exports; i++)
                {
                    using var output = template.Export(new { Url = $"http://localhost:{port}/x.png" });
                    Assert.NotNull(output);
                    using var doc = WordprocessingDocument.Open(output!, false);
                    var images = doc.MainDocumentPart!.ImageParts.Select(ReadAll).ToList();
                    Assert.True(images.Any(bytes => bytes.SequenceEqual(png)),
                        $"Export #{i + 1}: no ImagePart matches the served PNG ({images.Count} image part(s)).");
                }

                Assert.Equal(exports, Volatile.Read(ref requests));
                Assert.Equal("/x.png", lastPath);
            }
            finally
            {
                listener.Stop();
                listener.Close();
                await server.WaitAsync(TimeSpan.FromSeconds(5));
            }
        }

        private static byte[] ReadAll(ImagePart part)
        {
            using var stream = part.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }

        private static int FreeLoopbackPort()
        {
            var probe = new TcpListener(IPAddress.Loopback, 0);
            probe.Start();
            try
            {
                return ((IPEndPoint)probe.LocalEndpoint).Port;
            }
            finally
            {
                probe.Stop();
            }
        }
    }
}
