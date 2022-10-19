using CommandLine;
using OutlookMailDownloader.Models;
using OutlookMailDownloader.Utils;
using System.Xml.Serialization;

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
            _configDeser.Serialize(fs, new Config { AttachmentSaveTo = Path.Combine(projDir, "{DATE}_{SUBJECT}_{FILENAME}"), });
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

            new MailReader().ReceiveAsync(authFile, CancellationToken.None).Wait();

            return 0;
        }
    }
}