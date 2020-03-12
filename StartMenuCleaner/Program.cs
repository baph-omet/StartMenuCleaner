using Shell32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace StartMenuCleaner {
    class Program {
        [STAThread]
        static void Main(string[] args) {
            bool deleteURLs = args.Contains("-u");
            try {
                //Folders to check
                string[] dirs = new string[] {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs")
                };

                foreach (string dir in dirs) {
                    List<DirectoryInfo> dirDeletionQueue = new List<DirectoryInfo>();
                    foreach (string d in Directory.GetDirectories(dir)) {
                        DirectoryInfo directory = new DirectoryInfo(d);

                        // Check to see if folder is hidden
                        if (directory.Attributes.HasFlag(FileAttributes.Hidden)) {
                            Console.WriteLine(string.Format("Skipping hidden directory {0}", d));
                            continue;
                        }

                        // Check for write access
                        if (!HasWriteAccessDirectory(d)) {
                            Console.WriteLine(string.Format("User does not have write access to directory {0}. Skipping.", d));
                            continue;
                        }

                        List<FileInfo> deletionQueue = new List<FileInfo>();

                        // Queue URL files for deletion
                        if (deleteURLs) foreach (string filename in Directory.GetFiles(directory.FullName, "*.url")) deletionQueue.Add(new FileInfo(filename));

                        // Queue broken shortcuts for deletion
                        foreach (FileInfo f in directory.GetFiles()) {
                            Shell shell = new Shell();
                            Folder folder = shell.NameSpace(directory.FullName);
                            FolderItem fi = folder.ParseName(f.Name);
                            if (!fi.IsLink) continue;
                            try {
                                ShellLinkObject link = (ShellLinkObject)fi.GetLink;
                                if (!File.Exists(link.Path)) {
                                    deletionQueue.Add(f);
                                    Console.WriteLine(string.Format("Detected broken shortcut {0} pointing to {1}", f.FullName, link.Path));
                                }
                            } catch (UnauthorizedAccessException) {
                                Console.WriteLine(string.Format("User does not have access to check validity of shortcut {0}. Skipping.", f.FullName));
                            }
                        }

                        // Delete marked files
                        foreach (FileInfo f in deletionQueue) {
                            try {
                                f.Delete();
                                Console.WriteLine("Deleted " + f.FullName);
                            } catch (IOException e) {
                                Console.WriteLine(string.Format("Could not delete {0}: {1}", f.FullName, e.ToString()));
                            } catch (UnauthorizedAccessException) {
                                Console.WriteLine(string.Format("User '{0}' does not have permission to delete {1}. Skipping.", Environment.UserName, f.FullName));
                            }
                        }

                        // If folder only has one file in it, move that file out into the root
                        if (directory.GetFiles().Length == 1) {
                            FileInfo f = directory.GetFiles()[0];
                            f.CopyTo(Path.Combine(dir, f.Name), true);
                            f.Delete();
                            Console.WriteLine(string.Format("Moved {0} out of folder.", f.FullName));
                        }

                        // If directory is empty, queue it for deletion
                        if (directory.GetFiles().Length == 0) dirDeletionQueue.Add(directory);
                    }

                    // Delete each queued directory
                    foreach (DirectoryInfo directory in dirDeletionQueue) {
                        if (directory.Name.Contains("StartUp")) continue;
                        try {
                            directory.Delete(true);
                            Console.WriteLine("Deleted " + directory.FullName);
                        } catch (IOException e) {
                            Console.WriteLine(string.Format("Could not delete {0}: {1}", directory.FullName, e.Message));
                        } catch (UnauthorizedAccessException) {
                            Console.WriteLine(string.Format("User '{0}' does not have permission to delete {1}. Skipping.", Environment.UserName, directory.FullName));
                        }
                    }

                    List<string> rootDeletionQueue = new List<string>();

                    // Queue URL files in root for deletion
                    if (deleteURLs) foreach (string filename in Directory.GetFiles(dir, "*.url")) rootDeletionQueue.Add(filename);

                    // Queue broken shortcuts in root for deletion
                    DirectoryInfo rootDirectory = new DirectoryInfo(dir);
                    foreach (FileInfo f in rootDirectory.GetFiles()) {
                        if (!HasWriteAccessFile(f.FullName)) continue;
                        Shell shell = new Shell();
                        Folder folder = shell.NameSpace(rootDirectory.FullName);
                        FolderItem fi = folder.ParseName(f.Name);
                        if (!fi.IsLink) continue;
                        try {
                            ShellLinkObject link = (ShellLinkObject)fi.GetLink;
                            if (!File.Exists(link.Path)) {
                                rootDeletionQueue.Add(f.FullName);
                                Console.WriteLine(string.Format("Detected broken shortcut {0} pointing to {1}", f.FullName, link.Path));
                            }
                        } catch (UnauthorizedAccessException) {
                            Console.WriteLine(string.Format("User does not have access to check validity of shortcut {0}. Skipping.", f.FullName));
                        }
                    }

                    // Delete queued files in root
                    foreach (string filename in rootDeletionQueue) {
                        File.Delete(filename);
                        Console.WriteLine("Deleted " + filename);
                    }
                }
            } catch (Exception e) {
                Console.WriteLine("An unhandled exception occurred:");
                Console.WriteLine(e.ToString());
            } finally {
                Console.WriteLine("Done!");
                Console.Read();
            }
        }

        private static bool HasWriteAccessDirectory(string path) {
            try {
                Directory.GetAccessControl(path);
                return true;
            } catch (UnauthorizedAccessException) {
                return false;
            }
        }

        private static bool HasWriteAccessFile(string path) {
            try {
                File.GetAccessControl(path);
                return true;
            } catch (UnauthorizedAccessException) {
                return false;
            }
        }
    }
}
