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
        bool labels_casted = false;

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
            Object basePath,
            Object token,
            Parameters parameters,
            Object debugInfo
          )
        {
            basePath_casted = Utils.castParamString(basePath, "basePath", false, "http://secxbrl.28.io/v1/_queries/public");
            api = ApiClients.getDataApiClient(basePath_casted);

            token_casted = Utils.castParamString(token, "token", true);
            eid_casted = Utils.castParamString(parameters.getParamValue("eid"), "eid", false);
            ticker_casted = Utils.castParamString(parameters.getParamValue("ticker"), "ticker", false);
            tag_casted = Utils.castParamString(parameters.getParamValue("tag"), "tag", false);
            aid_casted = Utils.castParamString(parameters.getParamValue("aid"), "aid", false);
            fiscalYear_casted = Utils.castParamString(parameters.getParamValue("fiscalYear"), "fiscalYear", false);
            concept_casted = Utils.castParamString(parameters.getParamValue("concept"), "concept", false);
            fiscalPeriod_casted = Utils.castParamString(parameters.getParamValue("fiscalPeriod"), "fiscalPeriod", false);
            fiscalPeriodType_casted = Utils.castParamString(parameters.getParamValue("fiscalPeriodType"), "fiscalPeriodType", false);
            report_casted = Utils.castParamString(parameters.getParamValue("report"), "report", false);
            additionalRules_casted = Utils.castParamString(parameters.getParamValue("additionalRules"), "additionalRules", false);
            open_casted = Utils.castParamBool(parameters.getParamValue("open"), "open", false);
            aggregationFunction_casted = Utils.castParamString(parameters.getParamValue("aggregation-function"), "aggregationFunction", false);
            profileName_casted = Utils.castParamString(parameters.getParamValue("profile-name"), "profileName", false);

            dimensions_casted = parameters.getDimensionsValues();
            dimensionDefaults_casted = parameters.getDimensionDefaultsValues();
            dimensionTypes_casted = parameters.getDimensionTypesValues();
            dimensionAggregation_casted = parameters.getDimensionAggregationsValues();

            count_casted = Utils.castParamBool(parameters.getParamValue("count"), "count", false);
            top_casted = Utils.castParamInt(parameters.getParamValue("top"), "top", 100);
            //skip_casted = Utils.castParamInt(skip, "skip", 0);

            debugInfo_casted = Utils.castParamBool(debugInfo, "debugInfo", false);

            if (!(Utils.hasEntityFilter(eid_casted, ticker_casted, tag_casted, dimensions_casted)
                && Utils.hasConceptFilter(concept_casted, dimensions_casted)
                && Utils.hasAdditionalFilter(fiscalYear_casted, fiscalPeriod_casted, dimensions_casted)))
            {
                throw new Exception("Too generic filter.");
            }
            //Utils.log("Created Task " + ToString());
        }

        private void appendRequest(ref StringBuilder sb, String key, String value, bool last = false)
        {
            if (value != null)
            {
                sb.Append(key);
                sb.Append("=");
                sb.Append(value);
                if (!last) sb.Append("&");
            }
        }

        private void appendRequest(ref StringBuilder sb, Dictionary<string, string> dict)
        {
            foreach (KeyValuePair<string, string> entry in dict)
            {
                appendRequest(ref sb, entry.Key, entry.Value);
            }
        }

        public string request()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(basePath_casted + "/api/facts?");
            appendRequest(ref sb, "eid", eid_casted);
            appendRequest(ref sb, "ticker", ticker_casted);
            appendRequest(ref sb, "tag", tag_casted);
            appendRequest(ref sb, "aid", aid_casted);
            appendRequest(ref sb, "fiscalYear", fiscalYear_casted);
            appendRequest(ref sb, "concept", concept_casted);
            appendRequest(ref sb, "fiscalPeriod", fiscalPeriod_casted);
            appendRequest(ref sb, "report", report_casted);
            appendRequest(ref sb, "additional-rules", additionalRules_casted);
            appendRequest(ref sb, "open", open_casted ? "true" : "false" );
            appendRequest(ref sb, "aggregation-function", aggregationFunction_casted);
            appendRequest(ref sb, "profile-name", profileName_casted);
            appendRequest(ref sb, "fiscalPeriodType", fiscalPeriodType_casted);
            appendRequest(ref sb, "count", count_casted ? "true" : "false" );
            appendRequest(ref sb, "top", Convert.ToString(top_casted));
            appendRequest(ref sb, "labels", labels_casted ? "true" : "false");
            appendRequest(ref sb, dimensions_casted);
            appendRequest(ref sb, dimensionDefaults_casted);
            appendRequest(ref sb, dimensionTypes_casted);
            appendRequest(ref sb, dimensionAggregation_casted);
            //append(ref sb, "skip", Convert.ToString(skip_casted));
            appendRequest(ref sb, "token", token_casted, true);
            return sb.ToString();
        }

        private void append(ref StringBuilder sb, String key, String value, bool last = false)
        {
            if (value != null)
            {
                sb.Append(key);
                sb.Append("=");
                sb.Append(value);
                if (!last) sb.Append(" | ");
            }
        }

        private void append(ref StringBuilder sb, Dictionary<string, string> dict, string prefix)
        {
            sb.Append(prefix);
            foreach (KeyValuePair<string, string> entry in dict)
            {
                append(ref sb, entry.Key, entry.Value);
            }
        }

        public string id()
        {
            StringBuilder sb = new StringBuilder();
            append(ref sb, "basePath", basePath_casted);
            //append(ref sb, "token", token_casted);
            append(ref sb, "eid", eid_casted);
            append(ref sb, "ticker", ticker_casted);
            append(ref sb, "tag", tag_casted);
            append(ref sb, "aid", aid_casted);
            append(ref sb, "fiscalYear", fiscalYear_casted);
            append(ref sb, "concept", concept_casted);
            append(ref sb, "fiscalPeriod", fiscalPeriod_casted);
            append(ref sb, "report", report_casted);
            append(ref sb, "additionalRules", additionalRules_casted);
            append(ref sb, "open", Convert.ToString(open_casted));
            append(ref sb, "aggregationFunction", aggregationFunction_casted);
            append(ref sb, "profileName", profileName_casted);
            append(ref sb, "fiscalPeriodType", fiscalPeriodType_casted);
            append(ref sb, "count", Convert.ToString(count_casted));
            append(ref sb, "top", Convert.ToString(top_casted));
            append(ref sb, "labels", Convert.ToString(labels_casted));
            append(ref sb, dimensions_casted, "[Dimensions:] ");
            append(ref sb, dimensionDefaults_casted, "[Defaults:] ");
            append(ref sb, dimensionTypes_casted, "[Types:] ");
            append(ref sb, dimensionAggregation_casted,"[Aggregations:] ");
            //append(ref sb, "skip", Convert.ToString(skip_casted));
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
                top: top_casted,
                labels: labels_casted);
                //skip: skip_casted);

            Object result = Utils.getFactTableResult(response, debugInfo_casted);
            Utils.log("Received '" + string.Join(", ", (Object[])result) + "' as response for " + ToString());
            return result;
        }
        
    }
    
}
