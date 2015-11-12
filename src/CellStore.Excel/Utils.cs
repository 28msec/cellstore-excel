using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using ExcelDna.Integration;


namespace CellStore.Excel.Tools
{

    /// <summary>
    /// Represents a collection of functions to help with several tasks
    /// </summary>
    public class Utils
    {

        private static String logPathDir =
            System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData) 
                + "\\28msec";
        public static String logPath = logPathDir + "\\CellStore.Excel.log";

        public static void log(String message)
        {
            if (!Directory.Exists(logPathDir))
            {
                Directory.CreateDirectory(logPathDir);
            }
            System.IO.StreamWriter file = new System.IO.StreamWriter(logPath, true);
            file.WriteLine("LOG {0}", message);
            file.Close();
        }

        private static string factToString(JObject fact)
        {
            StringBuilder sb = new StringBuilder();
            JObject aspects = (JObject)fact["Aspects"];
            String archive = ((JValue)aspects["xbrl28:Archive"]).Value<String>();
            String entity = ((JValue)aspects["xbrl:Entity"]).Value<String>();
            String period = ((JValue)aspects["xbrl:Period"]).Value<String>();
            String concept = ((JValue)aspects["xbrl:Concept"]).Value<String>();
            String unit = ((JValue)aspects["xbrl:Unit"]).Value<String>();
            String value = ((JValue)fact["Value"]).Value<String>();

            sb.AppendLine(concept + "= " + value);
            sb.AppendLine(entity + " (" + archive + ", " + period + ", " + unit + ")");
            List<string> ignore = new List<string> { "xbrl28:Archive", "xbrl:Entity", "xbrl:Period", "xbrl:Concept", "xbrl:Unit" };
            foreach (JValue keyAspect in fact["KeyAspects"].Children())
            {
                String dim = keyAspect.Value<String>();
                if (ignore.Contains(dim))
                    continue;
                String mem = ((JValue)aspects[dim]).Value<String>();
                sb.AppendLine(dim + "= " + mem);
            }
            sb.AppendLine("------------------------------------");
            return sb.ToString();
        }

        public static Object[,] defaultErrorHandler(Exception ex)
        {
            String caption = "ERROR";
            String msg = ex.Message;
            log(caption + ": " + msg);
            Object[,] error = new Object[,] { { "# ERROR " + msg } };
            return error;
        }

        public static Object[,] getFactTableResult(dynamic response, bool debugInfo = false)
        {
            JArray facts = (JArray)response["FactTable"];
            Object[,] results;
            if (facts.Count > 0)
            {
                results = new Object[facts.Count, 1];
            }
            else
            {
                results = new Object[1, 1] { { ExcelEmpty.Value } };
            }

            StringBuilder sb = new StringBuilder();
            int row = 0;
            foreach (JObject fact in facts.Children())
            {
                JValue value = (JValue)fact["Value"];
                results[row, 0] = value.Value<double>();
                if (debugInfo)
                    sb.Append(factToString(fact));
                row++;
            }
            if (debugInfo)
            {
                MessageBox.Show(sb.ToString(), "Facts Debug Info");
            }
            return results;
        }

        public static bool hasEntityFilter(String eid, String ticker, String tag)
        {
            return eid != null || ticker != null || tag != null;
        }

        public static bool hasConceptFilter(String concept)
        {
            return concept != null;
        }

        public static bool hasAdditionalFilter(String fiscalYear, String fiscalPeriod, Dictionary<string, string> dimensions)
        {
            return fiscalYear != null || fiscalPeriod != null || dimensions != null;
        }

        public static String castParamString(
            Object param, String paramName, bool isMandatory, String defaultVal = null)
        {
            String param_casted;
            if (param is string)
            {
                param_casted = Convert.ToString(param);
            }
            else if (param is double)
            {
                param_casted = Convert.ToString(param);
            }
            else if (!isMandatory && (param == null || param is ExcelEmpty || param is ExcelMissing))
            {
                param_casted = defaultVal;
            }
            else if (isMandatory && (param == null || param is ExcelEmpty || param is ExcelMissing))
            {
                throw new ArgumentException("Mandatory Parameter missing: '" + paramName + "'.", paramName);
            }
            else
            {
                throw new ArgumentException("Invalid Parameter value '" + param.ToString() + "'.", paramName);
            }
            return param_casted;
        }

        public static bool castParamBool(
            Object param, String paramName, bool defaultVal)
        {
            bool param_casted;
            if (param is bool)
            {
                param_casted = Convert.ToBoolean(param);
            }
            else if (param is double)
            {
                param_casted = Convert.ToBoolean(param);
            }
            else if (param is ExcelEmpty || param is ExcelMissing)
            {
                param_casted = defaultVal;
            }
            else
            {
                throw new ArgumentException("Invalid Boolean Parameter value '" + param.ToString() + "'.", paramName);
            }
            return param_casted;
        }

        public static int? castParamInt(
            Object param, String paramName, int? defaultVal = null)
        {
            int? param_casted;
            if (param is string)
            {
                param_casted = Convert.ToInt32(param);
            }
            else if (param is double)
            {
                param_casted = Convert.ToInt32(param);
            }
            else if (param is ExcelEmpty || param is ExcelMissing)
            {
                param_casted = defaultVal;
            }
            else
            {
                throw new ArgumentException("Invalid Parameter value '" + param.ToString() + "'.", paramName);
            }
            return param_casted;
        }

        public static Dictionary<string, string> castStringDictionary(
            Object param, String paramName)
        {
            if (param == null || param is ExcelEmpty || param is ExcelMissing)
            {
                return null;
            }

            Dictionary<string, string> dict = null;
            if (param is Object[])
            {
                dict = new Dictionary<string, string>();
                Object[] param_casted = (Object[])param;
                int param_d1 = param_casted.Length;
                for (int i = 0; i < param_d1; i++)
                {
                    if (param_casted[i] == null || param_casted[i] is ExcelEmpty || param_casted[i] is ExcelMissing)
                    {
                        continue;
                    }
                    string param_Val = Convert.ToString(param_casted[i]);
                    string[] tokenz = param_Val.Split('=');
                    string errormsg = "Invalid Parameter value '" + param_Val + "' for parameter '" + paramName + "'. Accepted format: 'prefix:Dimension=value'.";
                    if (tokenz.Length != 2)
                    {
                        throw new ArgumentException(errormsg, paramName);
                    }
                    Regex regex = new Regex("^[^:]+:[^:]+$");
                    string param_Key = tokenz[0];
                    if (!regex.IsMatch(param_Key))
                    {
                        param_Key = tokenz[0] + "";
                    }
                    if (!regex.IsMatch(param_Key))
                    {
                        throw new ArgumentException(errormsg, paramName);
                    }
                    string param_Value = Convert.ToString(tokenz[1]);
                    dict.Add(param_Key, param_Value);
                }
                return dict;
            }
            else
            {
                throw new ArgumentException("Invalid Parameter value '" + paramName + "'.", paramName);
            }
        }

        public static Dictionary<string, bool> castBoolDictionary(
            Object param, String paramName)
        {
            Dictionary<string, string> dictString = castStringDictionary(param, paramName);
            if (dictString == null)
            {
                return null;
            }

            Dictionary<string, bool> dict = new Dictionary<string, bool>(dictString.Count);
            foreach (KeyValuePair<string, string> entry in dictString)
            {
                dict.Add(entry.Key, Convert.ToBoolean(entry.Value));
            }
            return dict;
        }

        public static Dictionary<string, int> castIntDictionary(
            Object param, String paramName)
        {
            Dictionary<string, string> dictString = castStringDictionary(param, paramName);
            if (dictString == null)
            {
                return null;
            }

            Dictionary<string, int> dict = new Dictionary<string, int>(dictString.Count);
            foreach (KeyValuePair<string, string> entry in dictString)
            {
                dict.Add(entry.Key, Convert.ToInt32(entry.Value));
            }
            return dict;
        }
    }
}
