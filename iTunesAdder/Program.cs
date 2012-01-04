using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;

namespace iTunesAdder
{
    public static class Program
    {
        private static string iTunesAddFolder = "E:\\iTunes\\Automatically Add to iTunes\\";
        
        static void Main(string[] args)
        {

            string input;
            string m4vpath;

            if (args.Length > 0)
            {
                input = args[1];
            }
            else
            {
                Console.Write("Enter a full path to file or a directory:");
                input = Console.ReadLine();
            }

            // get the file attributes for file or directory
            FileAttributes attr = File.GetAttributes(input);

            //detect whether its a directory or file
            if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
            {
                string[] filePaths = Directory.GetFiles(input, "*.avi"/*, SearchOption.AllDirectories*/);

                Console.Write("Processing following {0} files:", filePaths.Length);

                for (int i = 0; i < filePaths.Length; i++)
                {
                    Console.Write(filePaths[i]);
                }

                for (int i = 0; i < filePaths.Length; i++)
                {
                    filePaths[i] = filePaths[i].Trim('"');
                    m4vpath = Encode(filePaths[i]);
                    AddMetadata(m4vpath);
                    AddToItunes(m4vpath);
                    DeleteSource(filePaths[i]);
                }
            }
            else
            {
                input = input.Trim('"');
                m4vpath = Encode(input);
                AddMetadata(m4vpath);
                AddToItunes(m4vpath);
                DeleteSource(input);
            }

            Console.Write("All done.");
        }

        private static void DeleteSource(string fullpath)
        {
            FileSystem.DeleteFile(fullpath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
        }

        private static void AddToItunes(string m4vpath)
        {
            string addpath = iTunesAddFolder+Path.GetFileName(m4vpath);

            if (!File.Exists(addpath))
            {
                try
                {
                    File.Move(m4vpath, addpath);
                    Console.Write("File moved to iTunes directory for adding.");
                }
                catch (Exception e)
                {
                    Console.Write("Error in moving file: \n {0}", e);
                }
            }
            else
            {
                Console.Write("File exists start up iTunes.");
            }
        }
        public static void AddMetadata(string path)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = false;
            startInfo.FileName = "AtomicParsley.exe";
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            string atoms = getAtoms(Path.GetFileName(path));
            startInfo.Arguments = path + " " + atoms;

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                Console.WriteLine("Meta-Add Started");
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                }
                Console.WriteLine("Successfull meta-add to file {0}! \n With arguments {1}", Path.GetFileName(path), atoms);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error with meta-add: \n {0}", e);
            }
        }

        private static string getAtoms(string name)
        {
            string[] patterns = { @"s\d\de\d\d", @"\dx\d\d", @"\d\d\d" };
            string[] dividers = { @"\.", @" ", @"\_" };
            const string tv = "--overWrite --genre \"TV Shows\" --stik \"TV Show\"";
            //const string movie = "--overWrite --genre \"Movies\" --stik \"Movie\"";
            string title = "--title \"\"";
            string show = "--TVShowName ";
            string season = "--TVSeason ";
            string episode = "--TVEpisodeNum ";
            string description = "--description ";
            string[] splitName = null;
            string se;

            for (int d = 0; d < dividers.Length; d++)
            {
                for (int i = 0; i < patterns.Length; i++)
                {
                    if (Regex.IsMatch(name, ".*" + dividers[d] + patterns[i] + dividers[d] + ".*", RegexOptions.IgnoreCase))
                    {

                        splitName = Regex.Split(name, patterns[i], RegexOptions.IgnoreCase);
                        if (splitName != null && splitName.Length > 0)
                        {
                            //get show name
                            show += "\"" + Regex.Replace(splitName[0], dividers[d], " ").Trim() + "\"";

                            //throw additional data into description
                            description += "\"";
                            if (splitName.Length > 1)
                            {
                                for (int j = 1; j < splitName.Length; j++)
                                {
                                    description += splitName[j];
                                }
                            }
                            description += "\"";

                            //get season and episode string and clean
                            se = Regex.Match(name, patterns[i], RegexOptions.IgnoreCase).Value;
                            se = se.Trim(' ', '.', '_');
                            string div = "";

                            if (i == 0)
                            {
                                div = "e";
                            }
                            else if (i == 1)
                            {
                                div = "x";
                            }

                            else if (i == 2)
                            {
                                se.TrimStart('0');
                                splitName[0] = se[0].ToString();
                                splitName[1] = se.Substring(1);
                            }
                            //split up string
                            if (i != 2)
                            {
                                splitName = Regex.Split(se, div, RegexOptions.IgnoreCase);
                            }
                            //get season # without any preceding 0's
                            while (splitName[0].StartsWith("0") || splitName[0].StartsWith("s") || splitName[0].StartsWith("S") && splitName[0].Length != 1)
                            {
                                splitName[0] = splitName[0].Substring(1);
                            }
                            season += splitName[0];


                            //get episode #
                            if (splitName.Length >= 1)
                            {
                                while (splitName[1].StartsWith("0") && splitName[1].Length != 1)
                                {
                                    splitName[1] = splitName[1].Substring(1);
                                }
                                episode += splitName[1];
                            }

                        }
                        return tv + " " + title + " " + show + " " + season + " " + episode + " " + description;
                    }
                }
            }
            return "Error";
        }
        public static string Encode(string path)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = false;
            startInfo.FileName = "HandbrakeCLI.exe";
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            startInfo.Arguments = "-i " + path + " -o " 
                + Path.GetDirectoryName(path) +"\\"
                + Path.GetFileNameWithoutExtension(path)
                + ".m4v --preset \"High Profile\"";

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                Console.WriteLine("Encoding has started.");
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                    Console.WriteLine("Encoding has completed.");
                }

                return Path.GetDirectoryName(path) +"\\"+ Path.GetFileNameWithoutExtension(path)+ ".m4v";

            }
            catch (Exception e)
            {
                // Log error.
                Console.WriteLine("Encoding has failed. \n {0}", e);
            }
            return "Error";
        }
    }
}
