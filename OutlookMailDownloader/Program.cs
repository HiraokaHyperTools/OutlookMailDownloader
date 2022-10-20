using Azure.Core;
using Azure.Identity;
using CommandLine;
using LiteDB;
using Microsoft.Graph;
using OutlookMailDownloader.Models;
using OutlookMailDownloader.Utils;
using System.Globalization;
using System.Xml.Serialization;
using Directory = System.IO.Directory;
using File = System.IO.File;

namespace OutlookMailDownloader
{
    internal class Program
    {
        static int Main(string[] args) =>
            Parser.Default.ParseArguments<InitOpt, RunOpt>(args)
                .MapResult<InitOpt, RunOpt, int>(
                    DoInit,
                    DoRun,
                    ex => 1
                );

        private static readonly XmlSerializer _configDeser = new XmlSerializer(typeof(Config));

        [Verb("init")]
        class InitOpt
        {
            [Value(0, MetaName = "Path")]
            public string? Path { get; set; }
        }

        private static int DoInit(InitOpt arg)
        {
            var projDir = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, arg?.Path ?? "."));
            Directory.CreateDirectory(projDir);
            using var fs = File.Create(Path.Combine(projDir, "Config.xml"));
            _configDeser.Serialize(fs, new Config { AttachmentSaveTo = Path.Combine(projDir, "save", "{DATE}_{SUBJECT}_{FILENAME}"), });
            return 0;
        }

        [Verb("run")]
        class RunOpt
        {
            [Value(0, MetaName = "Path")]
            public string? Path { get; set; }
        }

        private static int DoRun(RunOpt arg)
        {
            var projDir = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, arg?.Path ?? "."));
            var authFile = Path.Combine(projDir, "Microsoft.bin");
            using var fs = File.OpenRead(Path.Combine(projDir, "Config.xml"));
            var config = (Config)(_configDeser.Deserialize(fs) ?? throw new Exception());

            using var db = new LiteDatabase(new ConnectionString { Filename = Path.Combine(projDir, "Sys.db") });
            var visitedMessages = db.GetCollection<VisitedMessage>();

            new MailReader().ReceiveAsync(
                authFile: authFile,
                saveAttachment: async (subject, fileName, timestamp, bytes) =>
                {
                    var saveTo = (config.AttachmentSaveTo ?? throw new Exception())
                        .Replace("{DATE}", (timestamp != null)
                            ? timestamp.Value.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                            : "000000"
                        )
                        .Replace("{SUBJECT}", FileNameNormalizer(subject))
                        .Replace("{FILENAME}", FileNameNormalizer(fileName));

                    var parentDir = Path.GetDirectoryName(saveTo);
                    if (parentDir != null)
                    {
                        Directory.CreateDirectory(parentDir);
                    }

                    await File.WriteAllBytesAsync(saveTo, bytes);
                },
                needDownload: async id =>
                {
                    if (visitedMessages.FindById(id) != null)
                    {
                        return false;
                    }
                    else
                    {
                        await Task.Run(() => visitedMessages.Insert(new VisitedMessage { Id = id, }));
                        return true;
                    }
                },
                cancellationToken: CancellationToken.None
            )
                .Wait();

            return 0;
        }

        private static string FileNameNormalizer(string name) =>
            name
                .Replace(":", "：")
                .Replace("/", "／")
                .Replace("\"", "”")
                .Replace("*", "＊")
                .Replace("?", "？")
                .Replace("<", "＜")
                .Replace("|", "｜")
                .Replace(">", "＞")
                ;
    }
}