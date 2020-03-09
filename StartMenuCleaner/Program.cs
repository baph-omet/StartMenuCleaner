using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StartMenuCleaner {
    class Program {
        static void Main(string[] args) {
            bool deleteURLs = args.Contains("-u");
            try {
                //Folders to check
                string[] dirs = new string[] {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs")
                };

                foreach (string dir in dirs) {
                    List<string> dirDeletionQueue = new List<string>();
                    foreach (string directory in Directory.GetDirectories(dir)) {
                        List<string> deletionQueue = new List<string>();
                        
                        if (deleteURLs) foreach (string filename in Directory.GetFiles(directory,"*.url")) deletionQueue.Add(filename);
                        
                        foreach (string filename in deletionQueue) {
                            try {
                                File.Delete(filename);
                                Console.WriteLine("Deleted " + filename);
                            } catch (IOException e) {
                                Console.WriteLine(string.Format("Could not delete {0}: {1}", filename, e.ToString()));
                            } catch (UnauthorizedAccessException) {
                                Console.WriteLine(string.Format("User '{0}' does not have permission to delete {1}. Skipping.",Environment.UserName, filename));
                            }
                        }

                        if (Directory.GetFiles(directory).Length == 1) {
                            string filename = Directory.GetFiles(directory)[0];
                            if (Path.GetExtension(filename).ToLower() != ".lnk") continue;
                            File.Copy(filename, Path.Combine(dir,Path.GetFileName(filename)), true);
                            File.Delete(filename);
                            Console.WriteLine(string.Format("Moved {0} out of folder.", filename));
                            dirDeletionQueue.Add(directory);
                        } else if (Directory.GetFiles(directory).Length == 0) dirDeletionQueue.Add(directory);
                    }

                    foreach (string directory in dirDeletionQueue) {
                        if (Path.GetDirectoryName(directory).Contains("StartUp")) continue;
                        try {
                            Directory.Delete(directory, true);
                            Console.WriteLine("Deleted " + directory);
                        } catch (IOException e) {
                            Console.WriteLine(string.Format("Could not delete {0}: {1}", directory, e.Message));
                        } catch (UnauthorizedAccessException) {
                            Console.WriteLine(string.Format("User '{0}' does not have permission to delete {1}. Skipping.", Environment.UserName, directory));
                        }
                    }

                    List<string> rootDeletionQueue = new List<string>();
                    
                    if (deleteURLs) foreach (string filename in Directory.GetFiles(dir,"*.url")) rootDeletionQueue.Add(filename);
                    
                    foreach (string filename in rootDeletionQueue) {
                        File.Delete(filename);
                        Console.WriteLine("Deleted " + filename);
                    }
                }
            } catch(Exception e) {
                Console.WriteLine("An unhandled exception occurred:");
                Console.WriteLine(e.ToString());
            } finally {
                Console.WriteLine("Done!");
                Console.Read();
            }
        }
    }
}
