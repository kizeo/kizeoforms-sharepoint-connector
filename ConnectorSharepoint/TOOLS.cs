using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Collections;

namespace TestClientObjectModel
{
    public static class TOOLS
    {

        public static log4net.ILog Log = log4net.LogManager.GetLogger(typeof(TOOLS));


        /// <summary>
        /// Removes All unothorized characters in a Sharepoint file name : \ " / > % * : ? | ! #
        /// </summary>
        /// <param name="s"> the string to clean</param>
        /// <returns>a clean string</returns>
        public static string CleanString(string s)
        {
            return s.Replace(@"""", "").Replace(@"\", "").Replace(@"/", "").Replace(">", "").Replace("<", "").Replace("!", "")
                .Replace("#", "").Replace("%", "").Replace("*", "").Replace(":", "").Replace("|", "").Replace("?", "");
        }

        /// <summary>
        /// take sharepoint path and create folders hiearchy one by one one by one 
        /// </summary>
        /// <param name="Context">sharepoint context</param>
        /// <param name="path">the path </param>
        /// <param name="formsLibrary"> form to spLibrary config </param>
        internal static void CreateSpPath(ClientContext Context, string path, List formsLibrary)
        {
            string folderPath;
            bool pathCreated = false;
            int slashIndex = 0;

            while (slashIndex != -619)
            {
                slashIndex = path.IndexOf("/", slashIndex + 1);

                if (slashIndex > 0)
                    folderPath = path.Substring(0, slashIndex);
                else
                {
                    folderPath = path;
                    slashIndex = -619;
                }

                pathCreated = CreateFolderIfNotExist(Context, folderPath, formsLibrary);

            }
            if(pathCreated)
            {
                Log.Debug("Path created");
            }
            else
            {
                Log.Debug("Path Already exist");
            }


        }

        /// <summary>
        /// everything is in the name
        /// </summary>
        /// <param name="Context"></param>
        /// <param name="folderPath"></param>
        /// <param name="formsLibrary"></param>
        /// <returns></returns>
        internal static bool CreateFolderIfNotExist(ClientContext Context, string folderPath, List formsLibrary)
        {
            var folderCreateInfo = new ListItemCreationInformation();
            folderCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
            folderCreateInfo.LeafName = folderPath;

            var folderItem = formsLibrary.AddItem(folderCreateInfo);
            folderItem.Update();

            try
            {
                Context.ExecuteQuery();
                Log.Warn($"Folder Created : {folderPath}");
                return (true);
            }
            catch (Exception )
            {
                Log.Warn($"Folder already exist : {folderPath}");
                return false;
            }
            
        }


        /// <summary>
        /// this methode converts to the correct type
        /// </summary>
        /// <param name="item"></param>
        /// <param name="dataMapping"></param>
        /// <param name="columnValue"></param>
        /// <returns></returns>
        public static ListItem ConvertToCorrectTypeAndSet(ListItem item, DataMapping dataMapping, string columnValue)
        {
                if (dataMapping.SpecialType == "Date")
                {
                    if (DateTime.TryParse(columnValue, out DateTime date))
                    {

                        item[dataMapping.SpColumnId] = DateTime.ParseExact(columnValue, "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture); ;
                    }
                    else
                    {
                        TOOLS.LogErrorwithoutExitProgram($"Error Parsing date value : {columnValue}");
                    item[dataMapping.SpColumnId] = "";
                    }
                }
                else
                {
                    if (columnValue != "") item[dataMapping.SpColumnId] = columnValue;
                } 

            return item;
        }


        public static void LogErrorAndExitProgram(string message)
        {
            Log.Error(message);
            Console.ReadKey();
            Environment.Exit(0);
        }

        public static void LogErrorwithoutExitProgram(string message)
        {
            Log.Error(message);

        }

        /// <summary>
        /// unzip zip file and return dictionay of (fileName : memory stream)
        /// </summary>
        /// <param name="ms"></param>
        /// <returns></returns>
      
    }

    /// <summary>
    /// thread safe HashSet
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class LockedHashset<T> : IEnumerable<T>
    {
        private object UnParUnSvp;
        private HashSet<T> TheHashSet;

        public LockedHashset()
        {
            UnParUnSvp = new object();
            TheHashSet = new HashSet<T>();
        }

        public void Add(T item)
        {
            lock (UnParUnSvp)
            {
                TheHashSet.Add(item);
            }
        }

      

        public IEnumerator<T> GetEnumerator()
        {
            return TheHashSet.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return TheHashSet.GetEnumerator();
        }
    }

   

}
