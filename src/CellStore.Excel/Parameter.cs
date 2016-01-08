using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace CellStore.Excel.Tasks
{
    
    public class Parameter
    {
        string name;
        List<string> values = new List<string>();

        public Parameter(
            string name,
            string value
          )
        {
            this.name = name;
            addValue(value);
        }
        
        public void addValue(string value)
        {
            this.values.Add(value);
        }

        public int size()
        {
            return values.Count;
        }

        public string getName()
        {
            return name;
        }

        public string getValue(int pos)
        {
            if(pos >= size())
            {
                return null;
            }
            return values[pos];
        }

        public bool isDimension()
        {
            Regex regex = new Regex("^[^:]+:[^:]+$");
            return regex.IsMatch(name);
        }

        public bool isDimensionDefault()
        {
            Regex regex = new Regex("^[^:]+:[^:]+::default$");
            return regex.IsMatch(name);
        }

        public bool isDimensionType()
        {
            Regex regex = new Regex("^[^:]+:[^:]+::type");
            return regex.IsMatch(name);
        }

        public bool isDimensionAggregation()
        {
            Regex regex = new Regex("^[^:]+:[^:]+::aggregation");
            return regex.IsMatch(name);
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(name);
            sb.Append("=");
            Boolean isFirst = true;
            foreach (string value in values)
            {
                if(!isFirst)
                    sb.Append(",");
                sb.Append(value);
                isFirst = false;
            }
            return sb.ToString();
        }
        
    }
    
}
