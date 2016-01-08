using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using ExcelDna.Integration;
using System.Runtime.Caching;

namespace CellStore.Excel.Tools
{

    /// <summary>
    /// Represents a Cache
    /// </summary>
    public class Cache
    {
        private static string LOADING = "LOADING:";

        public static bool get(string key, out Object result)
        {
            ObjectCache cache = MemoryCache.Default;
            result = cache[key] as Object;
            if (result != null)
                return true;
            return false;
        }

        public static Object getLoading(string key)
        {
            ObjectCache cache = MemoryCache.Default;
            Object result = cache[LOADING + key] as Object;
            if (result != null)
            {
                return result;
            }                
            return "# Loading";
        }

        public static void set(string key, Object result)
        {
            ObjectCache cache = MemoryCache.Default;
            cache.Add(key, result, DateTime.Now.AddMinutes(1), null);
            cache.Add(LOADING + key, result, DateTime.Now.AddMinutes(2), null);
        }
    }
            
}
