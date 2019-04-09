

using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace TLTC_VaultService.Lib
{
    public class OpenXMLWord : IDisposable
    {
        // Specify whether the instance is disposed.
        private bool disposed = false;

        // The word package
        private static WordprocessingDocument package = null;
        private static XLWorkbook workbook = null;
        private static IXLWorksheet worksheet = null;
       
        private static string FileName = string.Empty;

        public static string ParseDirectoryFileContentReturnString(string dir, string condition)
        {
            List<string> result = ParseDirectoryFileContent(dir, condition);
            return string.Join(",", result);
        }
        public static List<string> ParseDirectoryFileContent(string dir,string condition)
        {
            List<string> matchFiles = new List<string>();
            string[] files = Directory.GetFiles(dir);
            foreach (string file in files)
            {
                FileInfo f = new FileInfo(file);
                bool flag = false;
                switch (f.Extension.ToLower())
                {
                    case ".docx":
                        flag = OpenXMLWord.SearchWordContains(file, condition);
                        if (flag)
                        {
                            matchFiles.Add(file);
                        }
                        break;

                }
                //DirectoryInfo sdir = new DirectoryInfo(dir);
                //string[] directories = Directory.GetDirectories(dir);
                //List<string> revMatchFiles = RecursiveDir(directories, condition);
                //matchFiles.AddRange(revMatchFiles);
            }
            return matchFiles;
        }
        private static List<string> RecursiveDir(string[] sourceDir, string cond)
        {
            List<string> matchFiles = new List<string>();
            foreach (var dir in sourceDir)
            {
                string[] files = Directory.GetFiles(dir);
                foreach (string file in files)
                {
                    FileInfo f = new FileInfo(file);
                    bool flag = false;
                    switch (f.Extension.ToLower())
                    {
                        case ".docx":
                            flag = OpenXMLWord.SearchWordContains(file, cond);
                            if (flag)
                            {
                                matchFiles.Add(file);
                            }
                            break;
                    }
                    DirectoryInfo sdir = new DirectoryInfo(dir);
                    string[] directories = Directory.GetDirectories(dir);
                    List<string> revMatchFiles = RecursiveDir(directories, cond);
                    matchFiles.AddRange(revMatchFiles);
                }
            }
            return matchFiles;
        }
        public static bool SearchWordContains(string filepath, string condition)
        {
            List<string> words = new List<string>();

            if (string.IsNullOrEmpty(filepath) || !File.Exists(filepath))
            {
                return false;
            }
            WordprocessingDocument package = WordprocessingDocument.Open(filepath, true);
            OpenXmlElement element = package.MainDocumentPart.Document.Body;
            if (element == null)
            {
                package.Close();
                return false;
            }
            string innText = element.InnerText;

            string[] conditions = ParseCondition(condition);
            foreach (var cond in conditions)
            {
                if (innText.IndexOf(cond, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    package.Close();
                    return true;
                }
            }
            package.Close();
            return false;
        }
        public static string[] ParseCondition(string condition)
        {
            string[] conditions = null;

            if (condition.Contains(",") == true)
            {
                conditions = condition.Split(',');
            }
            else
            {
                conditions = new string[1];
                conditions[0] = condition;
            }

            return conditions;
        }


        
        #region IDisposable interface

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            // Protect from being called multiple times.
            if (disposed)
            {
                return;
            }

            if (disposing)
            {
                // Clean up all managed resources.
                if (package != null)
                {
                    package.Dispose();
                }
            }

            disposed = true;
        }
        #endregion
    }
}
