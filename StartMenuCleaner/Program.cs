using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StartMenuCleaner {
    class Program {
        static void Main(string[] args) {
            try {
                string[] dirs = new string[] {
                    Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) + "\\Programs",
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu) + "\\Programs"
                };
                foreach (string dir in dirs) {
                    List<string> dirDeletionQueue = new List<string>();
                    foreach (string directory in Directory.GetDirectories(dir)) {
                        List<string> deletionQueue = new List<string>();
                        foreach (string filename in Directory.GetFiles(directory)) if (Path.GetExtension(filename).ToLower() == ".url" && args.Contains("/u")) deletionQueue.Add(filename);
                        foreach (string filename in deletionQueue) {
                            File.Delete(filename);
                            Console.WriteLine("Deleted " + filename);
                        }
                        if (Directory.GetFiles(directory).Length == 1) {
                            string filename = Directory.GetFiles(directory)[0];
                            if (Path.GetExtension(filename).ToLower() != ".lnk") continue;
                            File.Copy(filename, dir + "\\" + Path.GetFileName(filename), true);
                            File.Delete(filename);
                            Console.WriteLine("Moved " + filename + " out of folder.");
                            dirDeletionQueue.Add(directory);
                        } else if (Directory.GetFiles(directory).Length == 0) dirDeletionQueue.Add(directory);
                    }
                    foreach (string directory in dirDeletionQueue) {
                        if (Path.GetDirectoryName(directory).Contains("StartUp")) continue;
                        try {
                            Directory.Delete(directory, true);
                            Console.WriteLine("Deleted " + directory);
                        } catch (IOException e) {
                            Console.WriteLine("Could not delete " + directory + ": " + e.Message);
                        }
                    }
                }
            } catch(Exception e) {
                Console.WriteLine("An unhandled exception occurred:\n" + e.ToString());
            } finally {
                Console.WriteLine("Done!");
                Console.Read();
            }
        }
    }
}
