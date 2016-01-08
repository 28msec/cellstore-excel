using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using CellStore.Excel.Client;
using CellStore.Excel.Tools;
using CellStore.Excel.Tasks;
using System.Text;

namespace CellStore.Excel.Api
{

    /// <summary>
    /// Represents a collection of helper functions for all APIs
    /// </summary>
    public class UtilFunctions
    {

        private static string parameter(
            Object name, 
            string firstParamName, 
            Object value,
            string secondParamName,
            string suffix = "")
        {
            string name_casted = Utils.castParamString(name, firstParamName, true) + suffix;
            string value_casted = Utils.castParamString(value, secondParamName, true);
            return name_casted + "=" + value_casted;
        }

        #region MemberFilter
        /// <summary>
        /// Create a Member Filter for a specific dimension.
        /// </summary>
        /// <param name="dimension">The name of dimension, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.</param> 
        /// <param name="member">The member to filter or ALL for a wildcard filter.</param> 
        /// <returns>Object</returns>
        [ExcelFunction(
            "Create a Member Filter for a specific dimension.",
            Name = "MemberFilter",
            Category = "CellStore.Utils")]
        public static Object MemberFilter(
          [ExcelArgument("The name of dimension, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.", Name="dimension")]
            Object dimension = null,
          [ExcelArgument("The member to filter or ALL for a wildcard filter.", Name="member")]
            Object member = null
          )
        {
            try
            {
                return parameter(dimension, "dimension", member, "member");
            }
            catch (Exception e)
            {
                return Utils.defaultErrorHandler(e);
            }
        }
        #endregion

        #region EntityFilter
        /// <summary>
        /// Create a Member Filter for Entities.
        /// </summary>
        /// <param name="scheme">The scheme of the entity identifyer, e.g. "http://www.sec.gov/CIK".</param> 
        /// <param name="id">The entity identifyer, e.g. "0001070750" or "ALL".</param> 
        /// <returns>Object</returns>
        [ExcelFunction(
            "Create a Member Filter for Entities.",
            Name = "EntityFilter",
            Category = "CellStore.Utils")]
        public static Object EntityFilter(
          [ExcelArgument("The scheme of the entity identifyer, e.g. http://www.sec.gov/CIK.", Name="scheme")]
            Object scheme = null,
          [ExcelArgument("The entity identifyer, e.g. '0001070750' or 'ALL'.", Name="id")]
            Object id = null
          )
        {
            try
            {
                string scheme_casted = Utils.castParamString(scheme, "scheme", false, "http://www.sec.gov/CIK");
                string id_casted = Utils.castParamString(id, "id", true, "ALL");
                string member = id_casted;
                Regex regex = new Regex("^" + scheme + " [^ ]+$");
                if (member != "ALL" && !regex.IsMatch(member))
                {
                    member = scheme + " " + id_casted;
                }
                return parameter("xbrl:Entity", "dimension", member, "scheme + ' ' + id");
            }
            catch (Exception e)
            {
                return Utils.defaultErrorHandler(e);
            }
        }
        #endregion

        #region Parameter
        /// <summary>
        /// Create a Parameter for the Rest API.
        /// </summary>
        /// <param name="parameter">The name of parameter, for example open, ticker, function, groupby.</param> 
        /// <param name="value">The member to filter or ALL for a wildcard filter.</param> 
        /// <returns>Object</returns>
        [ExcelFunction(
            "Create a Parameter for the Rest API.",
            Name = "Parameter",
            Category = "CellStore.Utils")]
        public static Object Parameter(
          [ExcelArgument("The name of parameter, for example open, ticker, function, groupby.", Name="name")]
            Object name = null,
          [ExcelArgument("The value of the API parameter.", Name="value")]
            Object value = null
          )
        {
            try
            {
                return parameter(name, "name", value, "value");
            }
            catch (Exception e)
            {
                return Utils.defaultErrorHandler(e);
            }
        }
        #endregion


        #region DimensionDefault
        /// <summary>
        /// Define a default value for a dimension.
        /// </summary>
        /// <param name="dimension">The name of dimension, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.</param> 
        /// <param name="default">The default value for a dimension.</param> 
        /// <returns>Object</returns>
        [ExcelFunction(
            "Define a default value for a dimension.",
            Name = "DimensionDefault",
            Category = "CellStore.Utils")]
        public static Object DimensionDefault(
          [ExcelArgument("The name of dimension, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.", Name="dimension")]
            Object dimension = null,
          [ExcelArgument("The default value for a dimension.", Name="default")]
            Object defaultV = null
          )
        {
            try
            {
                return parameter(dimension, "dimension", defaultV, "default", "::default");
            }
            catch (Exception e)
            {
                return Utils.defaultErrorHandler(e);
            }
        }
        #endregion

        #region DimensionType
        /// <summary>
        /// Define the type of a dimension.
        /// </summary>
        /// <param name="dimension">The name of dimension, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.</param> 
        /// <param name="type">The type of the dimension, e.g. string, integer.</param> 
        /// <returns>Object</returns>
        [ExcelFunction(
            "Define the type of a dimension.",
            Name = "DimensionType",
            Category = "CellStore.Utils")]
        public static Object DimensionType(
          [ExcelArgument("The name of dimension, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.", Name="dimension")]
            Object dimension = null,
          [ExcelArgument("The type of the dimension, e.g. string, integer.", Name="type")]
            Object type = null
          )
        {
            try
            {
                return parameter(dimension, "dimension", type, "type", "::type");
            }
            catch (Exception e)
            {
                return Utils.defaultErrorHandler(e);
            }
        }
        #endregion

        private static void parseParameters(Object param, string suffix, string value, ref List<string> parameters)
        {
            if (param is Object[])
            {
                Object[] param_casted = (Object[])param;
                int param_d1 = param_casted.Length;
                for (int i = 0; i < param_d1; i++)
                {
                    if (param_casted[i] == null || param_casted[i] is ExcelEmpty || param_casted[i] is ExcelMissing)
                    {
                        continue;
                    }
                    parseParameters(param_casted[i], suffix, value, ref parameters);
                }
            }
            else if (param is Object[,])
            {
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
                        parseParameters(param_casted[i,j], suffix, value, ref parameters);
                    }
                }
            }
            else if (param is ExcelReference)
            {
                ExcelReference reference = (ExcelReference)param;
                List<ExcelReference> list = reference.InnerReferences;
                if (reference.GetValue() is ExcelError && list != null && list.ToArray().Length > 0)
                {
                    foreach (ExcelReference refer in list)
                    {
                        Object val = refer.GetValue();
                        parseParameters(val, suffix, value, ref parameters);
                    }
                }
                else
                {
                    parseParameters(reference.GetValue(), suffix, value, ref parameters);
                }
            }
            else if (param is string)
            {
                string param_Key = Convert.ToString(param);
                string errormsg = "Invalid Parameter value '" + param_Key + "'. Accepted format: 'prefix:Dimension'.";
                Regex regex = new Regex("^[^:]+:[^:]+" + suffix + "$");
                param_Key = param_Key + suffix;
                if (!regex.IsMatch(param_Key))
                {
                    throw new ArgumentException(errormsg, param_Key);
                }
                parameters.Add(param_Key + "=" + value);
            }
            else if (param == null || param is ExcelEmpty || param is ExcelMissing)
            {
                ; // skip
            }
            else
            {
                throw new ArgumentException("Invalid Parameter value '" + Convert.ToString(param) + ".", Convert.ToString(param));
            }
        }

        #region Aggregation
        /// <summary>
        /// Define an aggregation for one or several Dimensions.
        /// </summary>
        /// <param name="AggregatedDimensions">List of dimensions, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.</param>
        /// <param name="GroupedDimensions">List of dimensions, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.</param>
        /// <param name="AggregationFunction">A function used for aggregation, e.g. avg, sum.</param>
        /// <returns>Object</returns>
        [ExcelFunction(
            "Define an aggregation for one or several Dimensions.",
            Name = "Aggregation",
            Category = "CellStore.Utils")]
        public static Object Aggregation(
          [ExcelArgument("List of dimensions, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.", 
            Name="AggregatedDimensions", AllowReference = true)]
            Object AggregatedDimensions = null,
          [ExcelArgument("List of dimensions, for example xbrl:Entity, xbrl:Concept, xbrl28:FiscalYear, dei:LegalEntityAxis.", 
            Name="GroupedDimensions", AllowReference = true)]
            Object GroupedDimensions = null,
          [ExcelArgument("A function used for aggregation, e.g. avg, sum. (Default: sum)", 
            Name="AggregationFunction")]
            Object AggregationFunction = null
          )
        {
            try
            {
                List<string> parameters = new List<string>();
                parseParameters(AggregatedDimensions, "::aggregation", "aggregate", ref parameters);
                parseParameters(GroupedDimensions, "::aggregation", "group", ref parameters);
                string function = Utils.castParamString(AggregationFunction, "AggregationFunction", false, "sum");
                parameters.Add("aggregation-function=" + function);
                StringBuilder sb = new StringBuilder();
                Boolean isFirst = true;
                foreach (string param in parameters)
                {
                    if (!isFirst)
                        sb.Append("&");
                    sb.Append(param);
                    isFirst = false;
                }
                return sb.ToString();
            }
            catch (Exception e)
            {
                return Utils.defaultErrorHandler(e);
            }
        }
        #endregion

    }

}
