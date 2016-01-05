using System;
using System.Collections.Generic;
using CellStore.Excel.Client;
using CellStore.Excel.Tools;
using System.Text;

namespace CellStore.Excel.Tasks
{
    
    public class ListFactsTask
    {
        CellStore.Api.DataApi api;

        string basePath_casted;
        string token_casted;
        string eid_casted;
        string ticker_casted;
        string tag_casted;
        string aid_casted;
        string fiscalYear_casted;
        string concept_casted;
        string fiscalPeriod_casted;
        string fiscalPeriodType_casted;
        string report_casted;
        string additionalRules_casted;
        bool open_casted;
        string aggregationFunction_casted;
        string profileName_casted;

        Dictionary<string, string> dimensions_casted;
        Dictionary<string, string> dimensionDefaults_casted;
        Dictionary<string, string> dimensionTypes_casted;
        /*Dictionary<string, bool?> dimensionSlicers_casted;*/
        Dictionary<string, string> dimensionAggregation_casted;

        bool count_casted;
        int? top_casted;
        //int? skip_casted;

        bool debugInfo_casted;

        public ListFactsTask(
            Object basePath = null,
            Object token = null,
            Object eid = null,
            Object ticker = null,
            Object tag = null,
            Object aid = null,
            Object fiscalYear = null,
            Object concept = null,
            Object fiscalPeriod = null,
            Object fiscalPeriodType = null,
            Object report = null,
            Object additionalRules = null,
            Object open = null,
            Object aggregationFunction = null,
            Object profileName = null,
            Object dimensions = null,
            Object dimensionDefaults = null,
            Object dimensionTypes = null,
            Object dimensionAggregation = null,
            Object count = null,
            Object top = null/*,
            Object skip = null*/,
            Object debugInfo = null
          )
        {
            basePath_casted = Utils.castParamString(basePath, "basePath", false, "http://secxbrl.28.io/v1/_queries/public");
            api = ApiClients.getDataApiClient(basePath_casted);

            token_casted = Utils.castParamString(token, "token", true);
            eid_casted = Utils.castParamString(eid, "eid", false);
            ticker_casted = Utils.castParamString(ticker, "ticker", false);
            tag_casted = Utils.castParamString(tag, "tag", false);
            aid_casted = Utils.castParamString(aid, "aid", false);
            fiscalYear_casted = Utils.castParamString(fiscalYear, "fiscalYear", false);
            concept_casted = Utils.castParamString(concept, "concept", false);
            fiscalPeriod_casted = Utils.castParamString(fiscalPeriod, "fiscalPeriod", false);
            fiscalPeriodType_casted = Utils.castParamString(fiscalPeriodType, "fiscalPeriodType", false);
            report_casted = Utils.castParamString(report, "report", false);
            additionalRules_casted = Utils.castParamString(additionalRules, "additionalRules", false);
            open_casted = Utils.castParamBool(open, "open", false);
            aggregationFunction_casted = Utils.castParamString(aggregationFunction, "aggregationFunction", false);
            profileName_casted = Utils.castParamString(profileName, "profileName", false);

            dimensions_casted =
                Utils.castStringDictionary(dimensions, "dimensions", "");
            dimensionDefaults_casted =
                Utils.castStringDictionary(dimensionDefaults, "dimensionDefaults", "::default");
            dimensionTypes_casted =
                Utils.castStringDictionary(dimensionTypes, "dimensionTypes", "::type");
            /*dimensionSlicers_casted =
                Utils.castBoolDictionary(dimensionSlicers, "dimensionSlicers");*/
            dimensionAggregation_casted =
                Utils.castStringDictionary(dimensionAggregation, "dimensionAggregation", "::aggregation");

            count_casted = Utils.castParamBool(count, "count", false);
            top_casted = Utils.castParamInt(top, "top", 100);
            //skip_casted = Utils.castParamInt(skip, "skip", 0);

            debugInfo_casted = Utils.castParamBool(debugInfo, "debugInfo", false);

            if (!(Utils.hasEntityFilter(eid_casted, ticker_casted, tag_casted, dimensions_casted)
                && Utils.hasConceptFilter(concept_casted, dimensions_casted)
                && Utils.hasAdditionalFilter(fiscalYear_casted, fiscalPeriod_casted, dimensions_casted)))
            {
                throw new Exception("Too generic filter.");
            }
            Utils.log("Created Task " + ToString());
        }

        private void append(StringBuilder sb, String key, String value, bool last = false)
        {
            if (value != null)
            {
                sb.Append(key);
                sb.Append("=");
                sb.Append(value);
                if (!last) sb.Append(" | ");
            }
        }

        private void append(StringBuilder sb, Dictionary<string, string> dict)
        {
            foreach (KeyValuePair<string, string> entry in dimensions_casted)
            {
                append(sb, entry.Key, entry.Value);
            }
        }

        public string id()
        {
            StringBuilder sb = new StringBuilder();
            append(sb, "basePath", basePath_casted);
            //append(sb, "token", token_casted);
            append(sb, "eid", eid_casted);
            append(sb, "ticker", ticker_casted);
            append(sb, "tag", tag_casted);
            append(sb, "aid", aid_casted);
            append(sb, "fiscalYear", fiscalYear_casted);
            append(sb, "concept", concept_casted);
            append(sb, "fiscalPeriod", fiscalPeriod_casted);
            append(sb, "report", report_casted);
            append(sb, "additionalRules", additionalRules_casted);
            append(sb, "open", Convert.ToString(open_casted));
            append(sb, "aggregationFunction", aggregationFunction_casted);
            append(sb, "profileName", profileName_casted);
            append(sb, "fiscalPeriodType", fiscalPeriodType_casted);
            append(sb, dimensions_casted);
            append(sb, dimensionDefaults_casted);
            append(sb, dimensionTypes_casted);
            append(sb, dimensionAggregation_casted);
            append(sb, "count", Convert.ToString(count_casted));
            append(sb, "top", Convert.ToString(top_casted));
            //append(sb, "skip", Convert.ToString(skip_casted));
            return sb.ToString();
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("ListFactsTask: ");
            sb.Append(id());
            return sb.ToString();
        }      

        public Object run()
        {
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
                defaultDimensionValues: dimensionDefaults_casted,
                dimensionTypes: dimensionTypes_casted,
                dimensionAggregation: dimensionAggregation_casted,
                count: count_casted,
                top: top_casted);
                //skip: skip_casted);

            Utils.log("Received response for " + ToString());
            return Utils.getFactTableResult(response, debugInfo_casted);
        }
        
    }
    
}
