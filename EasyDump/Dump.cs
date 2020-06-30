using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Compression;

namespace EasyDump
{
    internal class Dump
    {
        private static string GetWmiClass(string className)
        {
            var output = "...\n";
            var query = new SelectQuery(@"Select * from " + className);

            using (var searcher = new ManagementObjectSearcher(query))
            {
                foreach (ManagementObject process in searcher.Get())
                {
                    var searcherProperties = process.Properties;
                    foreach (var sp in searcherProperties)
                    {
                        output += sp.Name + ": " + sp.Value + "\n";
                    }
                }
            }

            return output;
        }
        
        public static void Run(string filename)
        {
            var directory = new DirectoryInfo(@"C:\Windows\Minidump");
            var miniDumpFile = directory.GetFiles()
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

            if (miniDumpFile == null)
            {
                MessageBox.Show("No minidump found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var date = DateTime.Now.ToString("yy-MM-dd_HH-mm-ss");
            var tempFolder = Directory.CreateDirectory("temp-" + date);

            File.Copy(miniDumpFile.FullName, tempFolder.FullName + "\\" + "minidump.dmp");

            var systemInformation = GetWmiClass("Win32_OperatingSystem")
                                    + GetWmiClass("Win32_ComputerSystem")
                                    + GetWmiClass("Win32_SystemDevices");

            File.WriteAllText(tempFolder.FullName + "\\" + "system_info.txt", systemInformation);
            ZipFile.CreateFromDirectory(tempFolder.FullName, filename);
            tempFolder.Delete(true);
            MessageBox.Show("Saved to " + filename + ".", "Ok", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
