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
        private static System.IO.StreamWriter logWriter;

        public static void initLogWriter()
        {
            if (!Directory.Exists(logPathDir))
            {
                Directory.CreateDirectory(logPathDir);
            }
            logWriter = new System.IO.StreamWriter(logPath, true);
        }

        public static void closeLogWriter()
        {
            logWriter.Close();
        }

        public static void log(String message)
        {
            logWriter.WriteLine("{1:HH:mm:ss.fff} LOG {0}", message, DateTime.Now);
        }

        private static double? getValueDouble(JObject fact)
        {
            JValue value = ((JValue)fact["Value"]);
            if (value != null && value.Type != JTokenType.Null && value.ToString() != String.Empty)
            {
                return value.Value<double>();
            }
            return null;
        }

        private static string getAspectValue(JObject fact, String aspectName)
        {
            JObject aspects = (JObject)fact["Aspects"];
            JValue aspect = ((JValue)aspects[aspectName]);
            if (aspect != null && aspect.Type != JTokenType.Null)
            {
                return aspect.Value<String>();
            }
            else if (aspect != null && aspect.Type == JTokenType.Null)
            {
                return "null";
            }
            return String.Empty;
        }

        private static string factToString(JObject fact)
        {
            StringBuilder sb = new StringBuilder();
            JObject aspects = (JObject)fact["Aspects"];
            String archive = getAspectValue(fact, "xbrl28:Archive");
            String entity = getAspectValue(fact, "xbrl:Entity");
            String period = getAspectValue(fact, "xbrl:Period");
            String concept = getAspectValue(fact, "xbrl:Concept");
            String unit = getAspectValue(fact, "xbrl:Unit");
            double? value = getValueDouble(fact);

            sb.AppendLine("Value=" + value);
            sb.AppendLine("Concept=" + concept + " Entity=" + entity + " Archive=" + archive + " Period=" + period + " Unit=" + unit);
            List<string> ignore = new List<string> { "xbrl28:Archive", "xbrl:Entity", "xbrl:Period", "xbrl:Concept", "xbrl:Unit" };
            foreach (JValue keyAspect in fact["KeyAspects"].Children())
            {
                String dim = keyAspect.Value<String>();
                if (ignore.Contains(dim))
                    continue;
                String mem = ((JValue)aspects[dim]).Value<String>();
                sb.Append(dim + "=" + mem + " ");
            }
            sb.AppendLine("");
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

        public static Object getFactTableResult(dynamic response, bool debugInfo = false)
        {
            JArray facts = (JArray)response["FactTable"];
            Object[] results;
            if (facts.Count > 0)
            {
                results = new Object[facts.Count];
            }
            else
            {
                results = new Object[] { ExcelEmpty.Value };
            }

            StringBuilder sb = new StringBuilder();
            int row = 0;
            foreach (JObject fact in facts.Children())
            {
                double? value = getValueDouble(fact);
                if (value != null)
                {
                    results[row] = value;
                    if (debugInfo)
                        sb.Append(factToString(fact));
                    row++;
                }
            }
            if (debugInfo)
            {
                Utils.log(facts.ToString());
                MessageBox.Show(sb.ToString(), "Facts Debug Info");
            }
            return results;
        }

        public static bool hasEntityFilter(String eid, String ticker, String tag, Dictionary<string, string> dimensions)
        {
            return eid != null || ticker != null || tag != null ||
                (dimensions != null && dimensions.ContainsKey("xbrl:Entity"));
        }

        public static bool hasConceptFilter(String concept, Dictionary<string, string> dimensions)
        {
            return concept != null ||
                (dimensions != null && dimensions.ContainsKey("xbrl:Concept"));
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
            if (param is Boolean)
            {
                param_casted = Convert.ToBoolean(param);
            }
            else if (param is double)
            {
                param_casted = Convert.ToBoolean(param);
            }
            else if (param == null || param is ExcelEmpty || param is ExcelMissing)
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
            else if (param == null || param is ExcelEmpty || param is ExcelMissing)
            {
                param_casted = defaultVal;
            }
            else
            {
                throw new ArgumentException("Invalid Parameter value '" + param.ToString() + "'.", paramName);
            }
            return param_casted;
        }

        private static void readFromObject(Object param, String paramName, String suffix, Dictionary<string, string> dict)
        {
            if (param is Object[])
            {
                //Utils.log("Object[]");
                Object[] param_casted = (Object[])param;
                int param_d1 = param_casted.Length;
                for (int i = 0; i < param_d1; i++)
                {
                    if (param_casted[i] == null || param_casted[i] is ExcelEmpty || param_casted[i] is ExcelMissing)
                    {
                        continue;
                    }
                    readFromObject(param_casted[i], paramName, suffix, dict);
                }
            }
            else if (param is Object[,])
            {
                //Utils.log("Object[,]");
                Object[,] param_casted = (Object[,])param;
                int param_d1 = param_casted.GetLength(0);
                int param_d2 = param_casted.GetLength(1);
                for (int i = 0; i < param_d1; i++)
                {
                    for (int j = 0; j < param_d2; j++)
                    {
                        if (param_casted[i, j] == null || param_casted[i, j] is ExcelEmpty || param_casted[i, j] is ExcelMissing)
                        {
                            continue;
                        }
                        readFromObject(param_casted[i, j], paramName, suffix, dict);
                    }
                }
            }
            else if (param is ExcelReference)
            {
                //Utils.log("ExcelReference");
                ExcelReference reference = (ExcelReference) param;
                List<ExcelReference> list = reference.InnerReferences;
                if (reference.GetValue() is ExcelError && list != null && list.ToArray().Length > 0)
                {
                    foreach (ExcelReference refer in list)
                    {
                        Object val = refer.GetValue();
                        readFromObject(val, paramName, suffix, dict);
                    }
                }
                else
                {
                    readFromObject(reference.GetValue(), paramName, suffix, dict);
                }
            }
            else if (param is string)
            {
                string param_Val = Convert.ToString(param);
                //Utils.log("val: " + param_Val);
                string[] tokenz = param_Val.Split('=');
                string errormsg = "Invalid Parameter value '" + param_Val + "' for parameter '" + paramName + "'. Accepted format: 'prefix:Dimension=value'.";
                if (tokenz.Length != 2)
                {
                    throw new ArgumentException(errormsg, paramName);
                }
                Regex regex = new Regex("^[^:]+:[^:]+" + suffix + "$");
                string param_Key = tokenz[0] + suffix;
                if (!regex.IsMatch(param_Key))
                {
                    param_Key = tokenz[0] + "";
                }
                if (!regex.IsMatch(param_Key))
                {
                    throw new ArgumentException(errormsg, paramName);
                }
                string param_Value = Convert.ToString(tokenz[1]);
                //Utils.log(param_Key + "=" + param_Value);
                dict.Add(param_Key, param_Value);
            }
            else if (param == null || param is ExcelEmpty || param is ExcelMissing)
            {
                ; // skip
            }
            else
            {
                throw new ArgumentException("Invalid Parameter value '" + Convert.ToString(param)  + "' for '" + paramName + "'.", paramName);
            }
        }

        public static Dictionary<string, string> castStringDictionary(
            Object param, String paramName, String suffix)
        {
            if (param == null || param is ExcelEmpty || param is ExcelMissing)
            {
                return null;
            }

            Dictionary<string, string> dict = new Dictionary<string, string>();
            readFromObject(param, paramName, suffix, dict);
            return dict;
        }

        public static Dictionary<string, bool> castBoolDictionary(
            Object param, String paramName, String suffix)
        {
            Dictionary<string, string> dictString = castStringDictionary(param, paramName, suffix);
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
            Object param, String paramName, String suffix)
        {
            Dictionary<string, string> dictString = castStringDictionary(param, paramName, suffix);
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
