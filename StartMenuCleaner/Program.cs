using Shell32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Principal;

namespace StartMenuCleaner {
    class Program {
        [STAThread]
        static void Main(string[] args) {
            if (!HasAdminRights()) {
                Console.WriteLine("This program requires administrator permissions since it needs to access Windows directories. Please re-launch as an administrator.");
                return;
            }

            bool deleteURLs = args.Contains("-u");
#if DEBUG
            deleteURLs = true;
#endif
            try {
                //Folders to check
                string[] dirs = new string[] {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs")
                };

                foreach (string dir in dirs) {
                    List<DirectoryInfo> dirDeletionQueue = new();
                    foreach (string d in Directory.GetDirectories(dir)) {
                        DirectoryInfo directory = new(d);

                        // Check to see if folder is hidden
                        if (directory.Attributes.HasFlag(FileAttributes.Hidden)) {
                            Console.WriteLine($"Skipping hidden directory {d}");
                            continue;
                        }

                        // Check for write access
                        if (!HasWriteAccessDirectory(d)) {
                            Console.WriteLine($"User does not have write access to directory {d}. Skipping.");
                            continue;
                        }

                        List<FileInfo> deletionQueue = new();

                        // Queue URL files for deletion
                        if (deleteURLs) foreach (string filename in Directory.GetFiles(directory.FullName, "*.url")) deletionQueue.Add(new FileInfo(filename));

                        // Queue broken shortcuts for deletion
                        foreach (FileInfo f in directory.GetFiles("*.lnk")) {
                            Shell shell = new();
                            Folder folder = shell.NameSpace(directory.FullName);
                            FolderItem fi = folder.ParseName(f.Name);
                            try {
                                ShellLinkObject link = (ShellLinkObject)fi.GetLink;
                                if (link.Path.ToLower().Contains(@"windows\system32")) continue;
                                try {
                                    using FileStream _ = File.OpenRead(link.Path);
                                } catch(Exception) {
                                    deletionQueue.Add(f);
                                    Console.WriteLine($"Detected broken shortcut {f.FullName} pointing to {link.Path}");
                                }
                            } catch (UnauthorizedAccessException) {
                                Console.WriteLine($"User does not have access to check validity of shortcut {f.FullName}. Skipping.");
                            }
                        }

                        // Delete marked files
                        foreach (FileInfo f in deletionQueue) {
                            try {
                                f.Delete();
                                Console.WriteLine($"Deleted {f.FullName}");
                            } catch (IOException e) {
                                Console.WriteLine($"Could not delete {f.FullName}: {e}");
                            } catch (UnauthorizedAccessException) {
                                Console.WriteLine($"User '{Environment.UserName}' does not have permission to delete {f.FullName}. Skipping.");
                            }
                        }

                        // If folder only has one file in it, move that file out into the root
                        if (directory.GetFiles().Length == 1) {
                            FileInfo f = directory.GetFiles()[0];
                            f.CopyTo(Path.Combine(dir, f.Name), true);
                            f.Delete();
                            Console.WriteLine($"Moved {f.FullName} out of folder.");
                        }

                        // If directory is empty, queue it for deletion
                        if (directory.GetFiles().Length == 0) dirDeletionQueue.Add(directory);
                    }

                    // Delete each queued directory
                    foreach (DirectoryInfo directory in dirDeletionQueue) {
                        if (directory.Name.Contains("StartUp")) continue;
                        try {
                            directory.Delete(true);
                            Console.WriteLine($"Deleted {directory.FullName}");
                        } catch (IOException e) {
                            Console.WriteLine($"Could not delete {directory.FullName}: {e.Message}");
                        } catch (UnauthorizedAccessException) {
                            Console.WriteLine($"User '{Environment.UserName}' does not have permission to delete {directory.FullName}. Skipping.");
                        }
                    }

                    List<string> rootDeletionQueue = new();

                    // Queue URL files in root for deletion
                    if (deleteURLs) foreach (string filename in Directory.GetFiles(dir, "*.url")) rootDeletionQueue.Add(filename);

                    // Queue broken shortcuts in root for deletion
                    DirectoryInfo rootDirectory = new(dir);
                    foreach (FileInfo f in rootDirectory.GetFiles("*.lnk")) {
                        if (!HasWriteAccessFile(f.FullName)) continue;
                        Shell shell = new();
                        Folder folder = shell.NameSpace(rootDirectory.FullName);
                        FolderItem fi = folder.ParseName(f.Name);
                        if (!fi.IsLink) continue;
                        try {
                            ShellLinkObject link = (ShellLinkObject)fi.GetLink;
                            if (link.Path.ToLower().Contains(@"windows\system32")) continue;
                            try {
                                using FileStream _ = File.OpenRead(link.Path);
                            } catch (Exception) {
                                rootDeletionQueue.Add(f.FullName);
                                Console.WriteLine($"Detected broken shortcut {f.FullName} pointing to {link.Path}");
                            }
                        } catch (UnauthorizedAccessException) {
                            Console.WriteLine($"User does not have access to check validity of shortcut {f.FullName}. Skipping.");
                        }
                    }

                    // Delete queued files in root
                    foreach (string filename in rootDeletionQueue) {
                        File.Delete(filename);
                        Console.WriteLine($"Deleted {filename}");
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

        private static bool HasAdminRights() {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
                return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
