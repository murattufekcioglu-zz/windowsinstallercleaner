using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WindowsInstaller;

namespace WindowsInstallerCleaner
{
    class Program
    {
        static void Main(string[] args)
        {
            var hashset = new HashSet<string>();

            var type = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            var installer = (Installer)Activator.CreateInstance(type);

            var folder = new DirectoryInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Installer"));

            var products = installer.Products;
            var productEnumerator = products.GetEnumerator();

            while (productEnumerator.MoveNext())
            {
                var product = productEnumerator.Current as string;
                var productLocation = installer.ProductInfo[product, "LocalPackage"];

                hashset.Add(productLocation);

                var patches = installer.Patches[product];
                var patchEnumerator = patches.GetEnumerator();
                
                while (patchEnumerator.MoveNext())
                {
                    var patch = patchEnumerator.Current as string;
                    var patchLocation = installer.PatchInfo[patch, "LocalPackage"];

                    hashset.Add(patchLocation);
                }
            }

            var files = folder.GetFiles("*.ms*");
            var unusedFiles = from file in files
                              where !hashset.Contains(file.FullName)
                              orderby file.Length descending
                              select file;

            var totalBytes = (from file in unusedFiles
                              select file.Length).Sum();

            foreach (var file in unusedFiles)
            {
                if ((file.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    file.Attributes = file.Attributes ^ FileAttributes.ReadOnly;
                file.Delete();
            }
        }
    }
}
