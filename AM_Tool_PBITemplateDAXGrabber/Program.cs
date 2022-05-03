using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json;
using System.IO;

namespace AM_Tool_PBITemplateDAXGrabber {
    class Program {
        static void Main(string[] args) {

            //PBI Template files come in with the UTF-16 LE encoding, which is supposed to be Unicode
            //Theres a lot of tricks in this code to get Unicode to play nice with basic strings

            //How to get a Power BI template file
            // 1. Open the PBIX file and save it as a Template
            // 2. Change the Template file type to .zip
            // 3. Extract the zip somewhere
            // 4. DataModelSchema is the JSON file that contains any DAX code from the PBIX file

            StringBuilder DAXHarvest = new StringBuilder();

            string utfnewline = "\n";
            byte[] utfBytes = Encoding.UTF8.GetBytes(utfnewline);
            byte[] unicodeBytes = Encoding.Convert(Encoding.UTF8, Encoding.Unicode, utfBytes);
            char[] unicodeChars = new char[Encoding.Unicode.GetCharCount(unicodeBytes, 0, unicodeBytes.Length)];
            Encoding.Unicode.GetChars(unicodeBytes, 0, unicodeBytes.Length, unicodeChars, 0);
            string unicodeNewline = new string(unicodeChars);


            Console.Write("Path to DataModelSchema: ");
            string jsonPath = Console.ReadLine();
            if (jsonPath.Length == 0) { 
                //local to exe
                jsonPath = Path.GetFullPath("DataModelSchema");
            } else { 
                //user path given
                jsonPath = Path.GetFullPath(jsonPath);
            }

            StringBuilder dataFile = new StringBuilder();
            dataFile.Append(File.ReadAllText(jsonPath, Encoding.Unicode));

            JObject dataModel = JObject.Parse(dataFile.ToString());

            JArray tables = (JArray)dataModel["model"]["tables"];

            for (int i = 0; i < tables.Count; i++) {
                
                JObject table = (JObject)tables[i];
                if (table.ContainsKey("isHidden")) continue;
                else {
                    DAXHarvest.AppendLine("=====================================");
                    DAXHarvest.AppendLine($"\t{(string)table["name"]}");
                    DAXHarvest.AppendLine("=====================================");
                    DAXHarvest.AppendLine();
                    

                    Dictionary<string, Dictionary<string, List<string>>> folders = new Dictionary<string, Dictionary<string, List<string>>>();
                    folders.Add("No Folder", new Dictionary<string, List<string>>());
                    folders["No Folder"].Add("Columns", new List<string>());
                    folders["No Folder"].Add("Measures", new List<string>());

                    foreach (JObject column in table["columns"]) {
                        if (column.ContainsKey("isHidden")) continue;
                        if (column.ContainsKey("type") && (string)column["type"] == "calculated") {
                            if (column.ContainsKey("displayFolder")) {
                                if (!folders.ContainsKey((string)column["displayFolder"])) {
                                    folders.Add((string)column["displayFolder"], new Dictionary<string, List<string>>());
                                    folders[(string)column["displayFolder"]].Add("Columns", new List<string>());
                                    folders[(string)column["displayFolder"]].Add("Measures", new List<string>());
                                }
                                folders[(string)column["displayFolder"]]["Columns"].Add($"{(string)column["name"]}={((string)column["expression"]).Replace(unicodeNewline, "")}");
                            } else {
                                folders["No Folder"]["Columns"].Add($"{(string)column["name"]}={((string)column["expression"]).Replace(unicodeNewline, "")}");
                            }
                        }
                    }

                    foreach (JObject measure in table["measures"]) {
                        if (measure.ContainsKey("displayFolder")) {
                            if (!folders.ContainsKey((string)measure["displayFolder"])) {
                                folders.Add((string)measure["displayFolder"], new Dictionary<string, List<string>>());
                                folders[(string)measure["displayFolder"]].Add("Columns", new List<string>());
                                folders[(string)measure["displayFolder"]].Add("Measures", new List<string>());
                            }
                            folders[(string)measure["displayFolder"]]["Measures"].Add($"{(string)measure["name"]}={((string)measure["expression"]).Replace(unicodeNewline, "")}");
                        } else {
                            folders["No Folder"]["Measures"].Add($"{(string)measure["name"]}={((string)measure["expression"]).Replace(unicodeNewline, "")}");
                        }   
                    }

                    foreach (string key in folders.Keys) {
                        DAXHarvest.AppendLine($"[[{key}]]");
                        if (folders[key]["Columns"].Count > 0) {
                            DAXHarvest.AppendLine("\t[Columns]");
                            foreach (string DAX in folders[key]["Columns"]) {
                                DAXHarvest.AppendLine($"\t\t{DAX}");
                            }
                            DAXHarvest.AppendLine();
                        }
                        if (folders[key]["Measures"].Count > 0) {
                            DAXHarvest.AppendLine("\t[Measures]");
                            foreach (string DAX in folders[key]["Measures"]) {
                                DAXHarvest.AppendLine($"\t\t{DAX}");
                            }
                            DAXHarvest.AppendLine();
                        }
                    }
                }
                DAXHarvest.AppendLine();
            }

            File.WriteAllText("DAX Harvest.txt", DAXHarvest.ToString());
        } 
    }
}
