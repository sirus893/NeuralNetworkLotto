// This file was auto-generated by ML.NET Model Builder. 

using Microsoft.ML.Data;

namespace MyAIML.Model
{
    public class ModelInput
    {
        [ColumnName("Sentiment"), LoadColumn(0)]
        public bool Sentiment { get; set; }


        [ColumnName("SentimentText"), LoadColumn(1)]
        public string SentimentText { get; set; }


    }
}
