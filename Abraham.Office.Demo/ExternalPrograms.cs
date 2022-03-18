using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;

namespace Abraham.ShellFunctions
{
    /// <summary>
    /// Helper class to open files with registered program
    /// Abraham Beratung 10/2013
    /// Oliver Abraham
    /// www.oliver-abraham.de
    /// mail@oliver-abraham.de
    /// </summary>
	public class ExternalPrograms
    {
        #region ------------- Methods -------------------------------------------------------------
        public static void StartProcess(string program, string arguments)
        {
            Process.Start(program, arguments);
        }

        public static void StartBatchfile(string path)
		{
			string AktuellesArbeitsverzeichnis = Directory.GetCurrentDirectory();
			StartBatchfile_internal(path);
			Directory.SetCurrentDirectory(AktuellesArbeitsverzeichnis);
		}

		private static void StartBatchfile_internal(string path)
		{
			var workingDirectory = Path.GetPathRoot(path);
			Directory.SetCurrentDirectory(workingDirectory);
			var programm = "cmd.exe";
			Process.Start(programm, " /k \"" + path + "\"");
		}

		public static void OpenDirectoryInExplorer(string path)
        {
            if (Directory.Exists(path) == false)
                throw new Exception("path does not exist");

            var programm = "explorer.exe";
            Process.Start(programm, " \"" + path + "\"");
        }

        public static void OpenFileInStandardBrowser(string filename)
        {
            var extension = Path.GetExtension(filename);
            var program = FindAssociatedProgramFor(extension);
            if (program == null)
                return;

            Process.Start(program, " \"" + filename + "\"");
        }

        public static void OpenFileInEditor(string filename)
        {
            var extension = Path.GetExtension(filename);
            string program = FindAssociatedProgramFor(extension);
            if (program == null)
                return;

            bool BracketsAreAlreadyThere = (program.Contains("\"%1\"") || program.Contains("\" %1\""));
            string filenameInBrackets;
            if (BracketsAreAlreadyThere)
                filenameInBrackets = filename;
            else
                filenameInBrackets = $"\"{filename}\"";

            var fullcommand = program.Replace("%1", filenameInBrackets);
            Process.Start(fullcommand, "");
        }

		public static string FindAssociatedProgramFor(string extension)
		{
            var docName = Registry.GetValue($@"HKEY_CLASSES_ROOT\{extension}", string.Empty, string.Empty) as string;
            if (!string.IsNullOrEmpty(docName))
            {
                var associatedProgram = Registry.GetValue($@"HKEY_CLASSES_ROOT\{docName}\shell\open\command", string.Empty, string.Empty) as string;
                return associatedProgram;
            }
            return null;
        }
        #endregion
    }
}
