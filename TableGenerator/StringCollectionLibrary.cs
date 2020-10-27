using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableGenerator
{
    /// <summary>
    /// Library object for working with array of strings.
    /// </summary>
    public static class StringCollectionLibrary
    {

        /// <summary>
        /// Method for getting collection of files names from array of paths.
        /// </summary>
        /// <param name="filesPaths">String array of paths.</param>
        /// <returns>
        /// List of files names strings.
        /// </returns>
        public static List<string> getNamesOfFilesFromPaths(string[] filesPaths)
        {
            List<string> filesNames = new List<string>();
            for(int i = 0; i < filesPaths.Length; i++)
            {
                int pos = filesPaths[i].LastIndexOf("\\") + 1;
                filesNames.Add(filesPaths[i].Substring(pos, filesPaths[i].Length - pos));
            }

            return filesNames;
        }

        /// <summary>
        /// Method for getting formatted names in collection of components.
        /// </summary>
        /// <param name="filesNames">Collection of names.</param>
        /// <returns>
        /// List of formatted components.
        /// </returns>
        public static List<string[]> getCollectionOfFormattedNames(List<string> filesNames)
        {
            List<string[]> dataCollection = new List<string[]>();
            filesNames.ForEach(delegate (string name)
            {
                int fileExtPos = name.LastIndexOf(".");
                if (fileExtPos >= 0)
                    name = name.Substring(0, fileExtPos);

                string[] nameComponents = name.Split('_');
                string[] dividedComponents = new string[5] { "", "", "", "", "" };
                for(int i = 0; i < dividedComponents.Length; i++)
                {
                    var nameComponentExists = nameComponents.ElementAtOrDefault(i) != null;
                    if (nameComponentExists == true)
                        dividedComponents[i] = nameComponents[i];
                }

                dataCollection.Add(dividedComponents);
            });

            return dataCollection;
        }
    }
}
