using System;
using System.Collections.Generic;
using CellStore.Excel.Client;
using CellStore.Excel.Tools;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace CellStore.Excel.Tasks
{
    
    public class Parameters
    {
        Dictionary<string, Parameter> parameters = new Dictionary<string, Parameter>();
        Dictionary<string, Parameter> dimensions = new Dictionary<string, Parameter>();
        Dictionary<string, Parameter> dimensionDefaults = new Dictionary<string, Parameter>();
        Dictionary<string, Parameter> dimensionTypes = new Dictionary<string, Parameter>();
        Dictionary<string, Parameter> dimensionAggregations = new Dictionary<string, Parameter>();

        public Parameters(
            Object parameters = null
          )
        {
            parse(parameters);
        }

        public Parameter getParameter(string name)
        {
            Parameter param;
            if (parameters.TryGetValue(name, out param)
                || dimensions.TryGetValue(name, out param)
                || dimensionDefaults.TryGetValue(name, out param)
                || dimensionTypes.TryGetValue(name, out param)
                || dimensionAggregations.TryGetValue(name, out param))
            {
                return param;
            };
            return null;
        }

        public void addParameter(string name, Parameter param)
        {
            if (param.isDimension())
            {
                dimensions.Add(name, param);
            } else if (param.isDimensionDefault())
            {
                dimensionDefaults.Add(name, param);
            }
            else if(param.isDimensionType())
            {
                dimensionTypes.Add(name, param);
            } else if (param.isDimensionAggregation())
            {
                dimensionAggregations.Add(name, param);
            }
            else
            {
                parameters.Add(name, param);
            }
        }

        public string getParamValue(string name)
        {
            Parameter param = getParameter(name);
            return getParamValue(param);
        }

        private string getParamValue(Parameter param)
        {
            if(param == null || param.size() == 0)
            {
                return null;
            }
            
            if(param.size() > 1)
            {
                throw new ArgumentException("Currently, only one value per parameter allowed: " + param.ToString(), param.getName());
            }
            return param.getValue(0);
        }

        private Dictionary<string, string> getValues(ref Dictionary<string, Parameter> paramtrs)
        {
            Dictionary<string, string> values = new Dictionary<string, string>();
            foreach (Parameter param in paramtrs.Values)
            {
                string value = getParamValue(param);
                if (value != null)
                {
                    values.Add(param.getName(), value);
                }
            }
            return values;
        }

        public Dictionary<string, string> getDimensionsValues()
        {
            return getValues(ref dimensions);
        }

        public Dictionary<string, string> getDimensionDefaultsValues()
        {
            return getValues(ref dimensionDefaults);
        }

        public Dictionary<string, string> getDimensionTypesValues()
        {
            return getValues(ref dimensionTypes);
        }

        public Dictionary<string, string> getDimensionAggregationsValues()
        {
            return getValues(ref dimensionAggregations);
        }

        public void parse(Object parameters)
        {
            if (parameters is Object[])
            {
                //Utils.log("Object[]");
                Object[] param_casted = (Object[]) parameters;
                int param_d1 = param_casted.Length;
                for (int i = 0; i < param_d1; i++)
                {
                    if (param_casted[i] == null || param_casted[i] is ExcelEmpty || param_casted[i] is ExcelMissing)
                    {
                        continue;
                    }
                    parse(param_casted[i]);
                }
            }
            else if (parameters is Object[,])
            {
                //Utils.log("Object[,]");
                Object[,] param_casted = (Object[,])parameters;
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
                        parse(param_casted[i, j]);
                    }
                }
            }
            else if (parameters is ExcelReference)
            {
                //Utils.log("ExcelReference");
                ExcelReference reference = (ExcelReference)parameters;
                List<ExcelReference> list = reference.InnerReferences;
                if (reference.GetValue() is ExcelError && list != null && list.ToArray().Length > 0)
                {
                    foreach (ExcelReference refer in list)
                    {
                        Object val = refer.GetValue();
                        parse(val);
                    }
                }
                else
                {
                    parse(reference.GetValue());
                }
            }
            else if (parameters is string)
            {
                string paramStr = Convert.ToString(parameters);

                string[] paramTokenz = paramStr.Split('&');

                if (paramTokenz.Length > 1)
                {
                    foreach (string param in paramTokenz)
                    {
                        parse(param);
                    }
                }
                else
                {
                    string[] tokenz = paramStr.Split('=');
                    string errormsg = "Invalid Parameter '" + paramStr + "'. Accepted format: 'parameter=value'.";
                    if (tokenz.Length != 2)
                    {
                        throw new ArgumentException(errormsg, "parameters");
                    }
                    string param_Key = Convert.ToString(tokenz[0]);
                    string param_Value = Convert.ToString(tokenz[1]);
                    Parameter param = getParameter(param_Key);
                    if (param != null)
                    {
                        param.addValue(param_Value);
                    }
                    else
                    {
                        param = new Parameter(param_Key, param_Value);
                        addParameter(param_Key, param);
                    }
                }
            }
            else if (parameters == null || parameters is ExcelEmpty || parameters is ExcelMissing)
            {
                ; // skip
            }
            else
            {
                throw new ArgumentException("Invalid Parameter value '" + Convert.ToString(parameters) + "' in 'parameters'.", "parameters");
            }
        }
        
        private string ToString(Dictionary<string, Parameter> parameters, string prefix = "")
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(prefix);
            Boolean isFirst = true;
            foreach (Parameter param in parameters.Values)
            {
                if (!isFirst)
                    sb.Append(" | ");
                sb.Append(param.ToString());
                isFirst = false;
            }
            isFirst = true;
            return sb.ToString();
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(ToString(parameters));
            if(dimensions.Count > 0)
                sb.AppendLine(ToString(dimensions, "    Dimensions           : "));
            if (dimensionDefaults.Count > 0)
                sb.AppendLine(ToString(dimensionDefaults, "    DimensionDefaults    : "));
            if (dimensionTypes.Count > 0)
                sb.AppendLine(ToString(dimensionTypes, "    DimensionTypes       : "));
            if (dimensionAggregations.Count > 0)
                sb.AppendLine(ToString(dimensionAggregations, "    DimensionAggregations: "));
            return sb.ToString();
        }

    }
    
}
