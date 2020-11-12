using System;

using OpenXmlProto.Data;

namespace OpenXmlProto
{
    class Program
    {
        static void Main(string[] args)
        {
            var res = CreateWithData(args);

            if (res) Console.WriteLine("Document successfully created!");
        }

        static bool CreateWithData(string[] args)
        {
            var title = args.Length > 0 && !(string.IsNullOrEmpty(args[0]))
                ? args[0]
                : CreateManifestTitle();

            var manifest = Manifest.GenerateManifest(title);
            return manifest.GenerateDocument(Environment.CurrentDirectory);
        }

        static string CreateManifestTitle()
        {
            Console.WriteLine("What would you like to name this manifest?");
            var title = Console.ReadLine();

            while(string.IsNullOrEmpty(title))
            {
                title = Console.ReadLine();
            }

            return title;
        }
    }
}
