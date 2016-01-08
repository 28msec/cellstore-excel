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
    /// Represents a collection of functions to interact with the API endpoints
    /// </summary>
    public class DataApiFunctions
    {

        #region ApiFacts
        /// <summary>
        /// Retrieve one or more facts for a combination of filings. 
        /// </summary>
        /// <param name="basePath">A URL to a specific endpoint, e.g. http://domain/v1/_queries/public</param> 
        /// <param name="token">The token of the current session</param> 
        /// <param name="parameters">All parameters to pass to the /api/facts endpoint</param> 
        /// <returns>Object</returns>            
        [ExcelFunction(
            "Retrieve one or more facts for a combination of filings. Please note, that in case multiple facts are returned Excel will only display the first one in a Cell. Nevertheless, you can apply aggregation function such as sum to all of them.",
            Name = "ApiFacts",
            Category = "CellStore.DataApi",
            IsVolatile = true)]
        public static Object ApiFacts(
          [ExcelArgument("The base Path of the Cell Store endpoint (default: 'http://secxbrl.28.io/v1/_queries/public')", Name="basePath")]
            Object basePath = null,
          [ExcelArgument("The token of the current session", Name="token")]
            Object token = null,
          [ExcelArgument("All parameters to pass to the /api/facts endpoint", Name="parameters", AllowReference = true)]
            Object parametersOrig = null,
          [ExcelArgument("Show debug info in a Message Box.", Name="debugInfo")]
            Object debugInfo = null
          )
        {
            try
            {
                Parameters parameters = new Parameters(parametersOrig);
                ListFactsTask task = new ListFactsTask(basePath, token, parameters, debugInfo);
                String taskId = task.id();
                Object result;
                if (Cache.get(taskId, out result))
                {
                    return result;
                }

                Utils.log("ListFacts Parameters: " + parameters.ToString());
                Object loading = Cache.getLoading(taskId);
                result = ExcelAsyncUtil.Run("ListFacts", taskId, delegate {
                    try
                    {
                        return task.run();
                    }
                    catch (Exception e)
                    {
                        return Utils.defaultErrorHandler(e);
                    }
                });
                if (result.Equals(ExcelError.ExcelErrorNA))
                {
                    return loading;
                }
                else
                {
                    Cache.set(taskId, result);
                    return result;
                }
            }
            catch (Exception e)
            {
                return Utils.defaultErrorHandler(e);
            }
        }
        #endregion

        #region ApiFactsRequest
        /// <summary>
        /// Retrieve the request to get one or more facts for a combination of filings.
        /// </summary>
        /// <param name="basePath">A URL to a specific endpoint, e.g. http://domain/v1/_queries/public</param> 
        /// <param name="token">The token of the current session</param> 
        /// <param name="parameters">All parameters to pass to the /api/facts endpoint</param> 
        /// <returns>Object</returns>
        [ExcelFunction(
            "Retrieve one or more facts for a combination of filings. Please note, that in case multiple facts are returned Excel will only display the first one in a Cell. Nevertheless, you can apply aggregation function such as sum to all of them.",
            Name = "ApiFactsRequest",
            Category = "CellStore.DataApi")]
        public static Object ApiFactsRequest(
          [ExcelArgument("The base Path of the Cell Store endpoint (default: 'http://secxbrl.28.io/v1/_queries/public')", Name="basePath")]
            Object basePath = null,
          [ExcelArgument("The token of the current session", Name="token")]
            Object token = null,
          [ExcelArgument("All parameters to pass to the /api/facts endpoint", Name="parameters", AllowReference = true)]
            Object parametersOrig = null
          )
        {
            try
            {
                Parameters parameters = new Parameters(parametersOrig);
                ListFactsTask task = new ListFactsTask(basePath, token, parameters, null);
                return task.request();
            }
            catch (Exception e)
            {
                return Utils.defaultErrorHandler(e);
            }
        }
        #endregion
    }

}
