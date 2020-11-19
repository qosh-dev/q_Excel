using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static Newtonsoft.Json.JsonConvert;


namespace Excel
{
    class Methods
    {
        public static List<Dictionary<string, string>> toJson<T>(List<T> listOfObj)
        {
            string[] converted = listOfObj.Select(e => SerializeObject(e)
                                          .Replace("{", "")
                                          .Replace("}", "")
                                          .Replace("\"", ""))
                                          .ToArray();

            // return converted.Select(e => 
            //                         e.Split(",")
            //                 .Select(e => new {
            //                     key = e.Split(":",2)[0],
            //                     value = e.Split(":",2)[1] 
            //         }
            //     ).ToDictionary(obj => obj.key, obj => obj.value))
            //      .ToList();


            List<Dictionary<string, string>> dict = new List<Dictionary<string, string>>();
            for (int i = 0; i < converted.Count(); i++)
            {
                Dictionary<string, string> tempDict = new Dictionary<string, string>();
                string[] tempArr = converted[i].Split(",");
                for (int o = 0; o < tempArr.Count(); o++)
                {
                    string[] finalEl = tempArr[o].Split(":", 2);
                    tempDict[finalEl[0]] = finalEl[1];
                }
                dict.Add(tempDict);
            }
            return dict;
        }
        public static Dictionary<string, string> toJson<T>(T obj)
        {
            string converted = SerializeObject(obj).Replace("{", "").Replace("}", "").Replace("\"", "");
            Dictionary<string, string> Dict = new Dictionary<string, string>();
            string[] tempArr = converted.Split(",");
            var first = tempArr[0].Split(":");
            if (first[1] == "0")
            {
                Dict[first[0]] = "null";
            }
            else
            {
                Dict[first[0]] = first[1];
            }
            for (int o = 1; o < tempArr.Count(); o++)
            {
                string[] finalEl = tempArr[o].Split(":", 2);
                Dict[finalEl[0]] = finalEl[1];
            }
            return Dict;
        }
    }
}
