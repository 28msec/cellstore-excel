using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.IO;
using System.Web;
using System.Linq;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using RestSharp;
using RestSharp.Extensions;
using CellStore.Api;

namespace CellStore.Excel.Client
{
    /// <summary>
    /// API client holds all the apis.
    /// </summary>
    public class ApiClients
    {
        private static Dictionary<String, CellStore.Api.DataApi> dataApis = new Dictionary<String, CellStore.Api.DataApi>();
        private static Dictionary<String, CellStore.Api.UsersApi> usersApis = new Dictionary<String, CellStore.Api.UsersApi>();
        private static Dictionary<String, CellStore.Api.SessionsApi> sessionsApis = new Dictionary<String, CellStore.Api.SessionsApi>();
        private static Dictionary<String, CellStore.Api.ReportsApi> reportsApis = new Dictionary<String, CellStore.Api.ReportsApi>();
    
        /// <summary>
        /// Get a client connecting to the DataApi of a specific endpoint.
        /// </summary>
        /// <param name="basePath">A URL to a specific endpoint, e.g. http://domain/v1/_queries/public</param> 
        /// <returns>CellStore.Api.DataApi</returns>            
        public static CellStore.Api.DataApi getDataApiClient(String basePath = "http://secxbrl.28.io/v1/_queries/public")
        {
            CellStore.Api.DataApi dataApi = null;
            if(!dataApis.TryGetValue(basePath, out dataApi))
            {
              dataApi = new CellStore.Api.DataApi(new CellStore.Client.ApiClient(basePath));
              dataApis.Add(basePath, dataApi);
            }
            return dataApi;
        }
        
        /// <summary>
        /// Get a client connecting to the UsersApi of a specific endpoint.
        /// </summary>
        /// <param name="basePath">A URL to a specific endpoint, e.g. http://domain/v1/_queries/public</param> 
        /// <returns>CellStore.Api.UsersApi</returns>            
        public static CellStore.Api.UsersApi getUsersApiClient(String basePath = "http://secxbrl.28.io/v1/_queries/public")
        {
            CellStore.Api.UsersApi userApi = null;
            if(!usersApis.TryGetValue(basePath, out userApi))
            {
              userApi = new CellStore.Api.UsersApi(new CellStore.Client.ApiClient(basePath));
              usersApis.Add(basePath, userApi);
            }
            return userApi;
        }
        
        /// <summary>
        /// Get a client connecting to the SessionsApi of a specific endpoint.
        /// </summary>
        /// <param name="basePath">A URL to a specific endpoint, e.g. http://domain/v1/_queries/public</param> 
        /// <returns>CellStore.Api.SessionsApi</returns>            
        public static CellStore.Api.SessionsApi getSessionsApiClient(String basePath = "http://secxbrl.28.io/v1/_queries/public")
        {
            CellStore.Api.SessionsApi sessionApi = null;
            if(!sessionsApis.TryGetValue(basePath, out sessionApi))
            {
              sessionApi = new CellStore.Api.SessionsApi(new CellStore.Client.ApiClient(basePath));
              sessionsApis.Add(basePath, sessionApi);
            }
            return sessionApi;
        }
        
        /// <summary>
        /// Get a client connecting to the ReportsApi of a specific endpoint.
        /// </summary>
        /// <param name="basePath">A URL to a specific endpoint, e.g. http://domain/v1/_queries/public</param> 
        /// <returns>CellStore.Api.ReportsApi</returns>            
        public static CellStore.Api.ReportsApi getReportsApiClient(String basePath = "http://secxbrl.28.io/v1/_queries/public")
        {
            CellStore.Api.ReportsApi reportsApi = null;
            if(!reportsApis.TryGetValue(basePath, out reportsApi))
            {
              reportsApi = new CellStore.Api.ReportsApi(new CellStore.Client.ApiClient(basePath));
              reportsApis.Add(basePath, reportsApi);
            }
            return reportsApi;
        }
  
    }
}
