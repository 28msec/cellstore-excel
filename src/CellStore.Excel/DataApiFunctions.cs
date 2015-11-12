using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using CellStore.Excel.Client;
using CellStore.Excel.Tools;

namespace CellStore.Excel.Api
{
    
    /// <summary>
    /// Represents a collection of functions to interact with the API endpoints
    /// </summary>
    public class DataApiFunctions
    {
                
        #region ListFacts
        /// <summary>
        /// Retrieve one or more facts for a combination of filings. 
        /// </summary>
        /// <param name="basePath">A URL to a specific endpoint, e.g. http://domain/v1/_queries/public</param> 
        /// <param name="token">The token of the current session</param> 
        /// <param name="eid">An Entity ID (a value of the xbrl:Entity aspect)</param> 
        /// <param name="cik">A CIK number</param> 
        /// <param name="edinetcode">An EDINET Code</param> 
        /// <param name="ticker">The ticker of the entity</param> 
        /// <param name="tag">The tag to filter entities</param> 
        /// <param name="sic">The industry group</param> 
        /// <param name="aid">The id of the filing</param> 
        /// <param name="fiscalYear">The fiscal year of the fact to retrieve (default: no filtering)</param> 
        /// <param name="concept">The name of the concept to retrieve the fact for (alternatively, a parameter with name xbrl:Concept can be used)</param> 
        /// <param name="fiscalPeriod">The fiscal period of the fact to retrieve (default: no filtering)</param> 
        /// <param name="fiscalPeriodType">The fiscal period type of the fact to retrieve (default: no filtering). Can select multiple</param> 
        /// <param name="report">The report to use as a context to retrieve the facts. In particular, concept maps and rules found in this report will be used. (default: none)</param> 
        /// <param name="additionalRules">The name of a report from which to use rules in addition to a report&#39;s rules (e.g. FundamentalAccountingConcepts)</param> 
        /// <param name="labels">Whether human readable labels should be included for concepts in each fact. (default: false)</param> 
        /// <param name="open">Whether the query has open hypercube semantics. (default: false)</param> 
        /// <param name="mostRecentVersionAspect">A transaction-time aspect of which only the latest value per cell is retained (default: sec:Accepted for SEC, fsa:Submitted for japan)</param> 
        /// <param name="profileName">Specifies which profile to use. The default depends on the underlying repository</param> 
        /// <param name="dimensions">A set of dimension names and values used for filtering. As a value, the value of the dimension or ALL can be provided if all facts with this dimension should be retrieved. Each key is in the form prefix:dimension, each value is a string</param> 
        /// <param name="dimensionTypes">Sets the given dimensions to be typed dimensions with the specified type. (Default: xbrl:Entity/xbrl:Period/xbrl:Unit/xbrl28:Archive are typed string, others are explicit dimensions. Some further dimensions may have default types depending on the profile.). Each key is in the form prefix:dimension::type, each value is a string</param> 
        /// <param name="defaultDimensionValues">Specifies the default value of the given dimensions that should be returned if the dimension was not provided explicitly for a fact. Each key is in the form  prefix:dimension::default, each value is a string</param> 
        /// <param name="dimensionSlicers">Specifies whether the given dimensions are slicers (true) or a dicers (false). Slicer dimensions do not appear in the output fact table (default: false). Each key is in the form prefix:dimension::slicer, each value is a boolean</param> 
        /// <param name="dimensionColumns">Specifies the position at which dicer dimensions appear in the output fact table (default: arbitrary order). Each key is in the form prefix:dimension::column, each value is an integer</param> 
        /// <param name="count">If true, only outputs statistics (default is false)</param> 
        /// <param name="top">Output only the first [top] results (default: no limit)</param> 
        /// <param name="skip">Skip the first [skip] results (default: 0)</param> 
        /// <returns>Object[,]</returns>            
        [ExcelFunction(
            "Retrieve one or more facts for a combination of filings. Please note, that in case multiple facts are returned Excel will only display the first one in a Cell. Nevertheless, you can apply aggregation function such as sum to all of them.",
            Name = "ListFacts", 
            Category = "CellStore.Excel.DataApi")]
        public static Object[,] ListFacts (
          [ExcelArgument("The base Path of the Cell Store endpoint (default: 'http://secxbrl.28.io/v1/_queries/public')", Name="basePath")]
            Object basePath = null,
          [ExcelArgument("The token of the current session", Name="token")]
            Object token = null, 
          [ExcelArgument("An Entity ID (a value of the xbrl:Entity aspect)", Name="eid")]
            Object eid = null, 
          [ExcelArgument("The ticker of the entity", Name="ticker")]
            Object ticker = null, 
          [ExcelArgument("The tag to filter entities", Name="tag")]
            Object tag = null, 
          [ExcelArgument("The id of the filing", Name="aid")]
            Object aid = null, 
          [ExcelArgument("The fiscal year of the fact to retrieve (default: no filtering)", Name="fiscalYear")]
            Object fiscalYear = null, 
          [ExcelArgument("The name of the concept to retrieve the fact for (alternatively, a parameter with name xbrl:Concept can be used)", Name="concept")]
            Object concept = null, 
          [ExcelArgument("The fiscal period of the fact to retrieve (default: no filtering)", Name="fiscalPeriod")]
            Object fiscalPeriod = null, 
          [ExcelArgument("The fiscal period type of the fact to retrieve (default: no filtering). Can select multiple", Name="fiscalPeriodType")]
            Object fiscalPeriodType = null, 
          [ExcelArgument("The report to use as a context to retrieve the facts. In particular, concept maps and rules found in this report will be used. (default: none)", Name="report")]
            Object report = null, 
          [ExcelArgument("The name of a report from which to use rules in addition to a report's rules (e.g. FundamentalAccountingConcepts)", Name="additionalRules")]
            Object additionalRules = null, 
          [ExcelArgument("Whether the query has open hypercube semantics. (default: false)", Name="open")]
            Object open = null, 
          [ExcelArgument("Specify an aggregation function to aggregate facts. Will aggregate facts, grouped by key aspects, with this function.", Name="aggregationFunction")]
            Object aggregationFunction = null, 
          [ExcelArgument("Specifies which profile to use. The default depends on the underlying repository", Name="profileName")]
            Object profileName = null, 
          [ExcelArgument("A set of dimension names and values used for filtering. As a value, the value of the dimension or ALL can be provided if all facts with this dimension should be retrieved. Each key is in the form prefix:dimension, each value is a string", Name="dimensions")]
            Object[] dimensions = null, 
          [ExcelArgument("Sets the given dimensions to be typed dimensions with the specified type. (Default: xbrl:Entity/xbrl:Period/xbrl:Unit/xbrl28:Archive are typed string, others are explicit dimensions. Some further dimensions may have default types depending on the profile.). Each key is in the form prefix:dimension::type, each value is a string", Name="dimensionTypes")]
            Object[] dimensionTypes = null,
          [ExcelArgument("Excludes (\"aggregate\") or includes (\"group\") the dimension in those used to group facts with the supplied aggregation function. By default, all key aspects are used as grouping keys and facts are aggregated along non-key aspects. Has no effect if no aggregation function is supplied.", Name="dimensionAggregation")]
            Object[] dimensionAggregation = null,
          [ExcelArgument("If true, only returns count of the result (default is false)", Name="count")]
            Object count = null, 
          [ExcelArgument("Output only the first [top] results (default: 100).", Name="top")]
            Object top = null/*,
          [ExcelArgument("Skip the first [skip] results (default: 0)", Name="skip")]
            Object skip = null*/,
          [ExcelArgument("Show debug info in a Message Box.", Name="debugInfo")]
            Object debugInfo = null
          )
        {
            try
            {
                string basePath_casted = Utils.castParamString(basePath, "basePath", false, "http://secxbrl.28.io/v1/_queries/public");
                CellStore.Api.DataApi api = ApiClients.getDataApiClient(basePath_casted);
                
                string token_casted = Utils.castParamString(token, "token", true);
                string eid_casted = Utils.castParamString(eid, "eid", false);
                string ticker_casted = Utils.castParamString(ticker, "ticker", false);
                string tag_casted = Utils.castParamString(tag, "tag", false);
                string aid_casted = Utils.castParamString(aid, "aid", false);
                string fiscalYear_casted = Utils.castParamString(fiscalYear, "fiscalYear", false);
                string concept_casted = Utils.castParamString(concept, "concept", false);
                string fiscalPeriod_casted = Utils.castParamString(fiscalPeriod, "fiscalPeriod", false);
                string fiscalPeriodType_casted = Utils.castParamString(fiscalPeriodType, "fiscalPeriodType", false);
                string report_casted = Utils.castParamString(report, "report", false);
                string additionalRules_casted = Utils.castParamString(additionalRules, "additionalRules", false);
                bool open_casted = Utils.castParamBool(open, "open", false);
                string aggregationFunction_casted = Utils.castParamString(aggregationFunction, "aggregationFunction", false);
                string profileName_casted = Utils.castParamString(profileName, "profileName", false);

                Dictionary<string, string> dimensions_casted = 
                    Utils.castStringDictionary(dimensions, "dimensions");
                Dictionary<string, string> dimensionTypes_casted = 
                    Utils.castStringDictionary(dimensionTypes, "dimensionTypes");
                /*Dictionary<string, bool?> dimensionSlicers_casted =
                    Utils.castBoolDictionary(dimensionSlicers, "dimensionSlicers");*/
                Dictionary<string, string> dimensionAggregation_casted =
                    Utils.castStringDictionary(dimensionAggregation, "dimensionAggregation");

                bool count_casted = Utils.castParamBool(count, "count", false);
                int? top_casted = Utils.castParamInt(top, "top", 100);
                //int? skip_casted = Utils.castParamInt(skip, "skip", 0);

                bool debugInfo_casted = Utils.castParamBool(debugInfo, "debugInfo", false); 

                if(!(Utils.hasEntityFilter(eid_casted, ticker_casted, tag_casted)
                    && Utils.hasConceptFilter(concept_casted)
                    && Utils.hasAdditionalFilter(fiscalYear_casted, fiscalPeriod_casted, dimensions_casted)))
                {
                    throw new Exception("Too generic filter.");
                }

                dynamic response = api.ListFacts(
                  token: token_casted,
                  eid: eid_casted,
                  ticker: ticker_casted,
                  tag: tag_casted,
                  aid: aid_casted,
                  fiscalYear: fiscalYear_casted,
                  concept: concept_casted,
                  fiscalPeriod: fiscalPeriod_casted,
                  fiscalPeriodType: fiscalPeriodType_casted,
                  report: report_casted,
                  additionalRules: additionalRules_casted,
                  open: open_casted,
                  aggregationFunction: aggregationFunction_casted,
                  profileName: profileName_casted,
                  dimensions: dimensions_casted,
                  dimensionTypes: dimensionTypes_casted,
                  dimensionAggregation: dimensionAggregation_casted,
                  count: count_casted,
                  top: top_casted);
                  //skip: skip_casted);
                return Utils.getFactTableResult(response, debugInfo_casted);
            }
            catch (Exception e)
            {
                return Utils.defaultErrorHandler(e);
            }
        }
        #endregion
        
    }
    
}
