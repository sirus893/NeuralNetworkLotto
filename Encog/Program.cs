using Encog.Engine.Network.Activation;
using Encog.ML.Data.Basic;
using Encog.Neural.Networks;
using Encog.Neural.Networks.Layers;
using Encog.Neural.Networks.Training.Propagation.Resilient;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Encog
{
    class Program
    {
        //public int V1 { get; private set; }
        //public int V2 { get; private set; }
        //public int V3 { get; private set; }
        //public int V4 { get; private set; }
        //public int V5 { get; private set; }

        static void Main(string[] args)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            //ListResults results = CreateDataOLE();
            ListResults results = CreateData();

            var deep = 20;
            var network = new BasicNetwork();

            network.AddLayer(new BasicLayer(null, true, 5 * deep));

            network.AddLayer(new BasicLayer(new ActivationSigmoid(), true, 5 * 5 * deep));

            network.AddLayer(new BasicLayer(new ActivationSigmoid(), true, 5 * 5 * deep));

            network.AddLayer(new BasicLayer(new ActivationLinear(), true, 5));

            network.Structure.FinalizeStructure();

            var learningInput = new double[deep][];

            for (int i = 0; i < deep; ++i)
            {
                learningInput[i] = new double[deep * 6];

                for (int j = 0, k = 0; j < deep; ++j)
                {
                    var idx = 2 * deep - i - j;
                    var data = results[idx];

                    learningInput[i][k++] = data.V1;
                    learningInput[i][k++] = data.V2;
                    learningInput[i][k++] = data.V3;
                    learningInput[i][k++] = data.V4;
                    learningInput[i][k++] = data.V5;
                }
            }

            var learningOutput = new double[deep][];
            for (int i = 0; i < deep; ++i)
            {
                var idx = deep - 1 - i;
                var data = results[idx];
                learningOutput[i] = new double[5]
                {
                    data.V1,
                    data.V2,
                    data.V3,
                    data.V4,
                    data.V5,
                };
            }

            var trainingSet = new BasicMLDataSet(learningInput, learningOutput);

            var train = new ResilientPropagation(network, trainingSet);

            train.NumThreads = Environment.ProcessorCount;

        START: network.Reset();

        RETRY:
            var step = 0;
            do
            {
                train.Iteration();
                Console.WriteLine("Train Error: {0}", train.Error);
                ++step;
            }
            while (train.Error > 0.001 && step < 20);

            var passedCount = 0;
            for (var i = 0; i < deep; ++i)
            {
               var should = new Results(learningOutput[i]);

               var inputn = new BasicMLData(6 * deep);

                Array.Copy(learningInput[i], inputn.Data, inputn.Data.Length);

                var comput = new Results(((BasicMLData)network.Compute(inputn)).Data);

                var passed = should.ToString() == comput.ToString();

                if (passed)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    ++passedCount;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                }
                Console.WriteLine("{0} {1} {2} {3}",should.ToString().PadLeft(17, ' '),passed ? "==" : "!=",
                        comput.ToString().PadRight(17, ' '),
                        passed ? "PASS" : "FAIL");

                Console.ResetColor();
            }

            var input = new BasicMLData(5 * deep);

            for (int i = 0, k = 0; i < deep; ++i)
            {
                var idx = deep - 1 - i;
                var data = results[idx];
                input.Data[k++] = data.V1;
                input.Data[k++] = data.V2;
                input.Data[k++] = data.V3;
                input.Data[k++] = data.V4;
                input.Data[k++] = data.V5;
            }

            var perfect = results[0];
            var predict = new Results(((BasicMLData)network.Compute(input)).Data);

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Predict: {0}", predict);
            Console.ResetColor();
            if (predict.IsOut())
                goto START;
            if ((double)passedCount < (deep * (double)9 / (double)10) || !predict.IsValid())
                goto RETRY;
            Console.WriteLine("Press any key for close...");
            Console.ReadKey(true);

        }

        //private static ListResults CreateDataOLE()
        //{
        //    string dir = Directory.GetCurrentDirectory();

        //}

        private static ListResults CreateData()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            string dir = Directory.GetCurrentDirectory();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(Path.Combine(dir, @"..\..\Excel\EuroMillions.xlsx")));
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            ListResults results = new ListResults();

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            for (int i = 2; i <= rowCount; i++)
            {
                List<int> numbers = new List<int>();

                for (int j = 2; j <= colCount - 2; j++)
                {
                    var cell = xlRange.Cells[i, j];

                    var number = Convert.ToInt32(cell.Value2);

                    if (number >= 1)
                    {
                        numbers.Add(number);
                    }
                }

                if (numbers.Count == 5)
                {
                    results.Add(new Results
                    (
                        numbers[0],
                        numbers[1],
                        numbers[2],
                        numbers[3],
                        numbers[4]
                    ));
                }
                else
                {
                    Console.Write("Something didnt get counted");
                }
            }
            stopwatch.Stop();

            return results;
        }
    }

    public class Results
    {
        public int V1 { get; private set; }
        public int V2 { get; private set; }
        public int V3 { get; private set; }
        public int V4 { get; private set; }
        public int V5 { get; private set; }


        public Results(int v1, int v2, int v3, int v4, int v5)

        {
            V1 = v1;
            V2 = v2;
            V3 = v3;
            V4 = v4;
            V5 = v5;

        }

        public Results(double[] values)

        {
            V1 = (int)values[0];
            V2 = (int)values[1];
            V3 = (int)values[2];
            V4 = (int)values[3];
            V5 = (int)values[4];
        }

        public bool IsValid()
        {
            return
            V1 >= 1 && V1 <= 49 &&
            V2 >= 1 && V2 <= 49 &&
            V3 >= 1 && V3 <= 49 &&
            V4 >= 1 && V4 <= 49 &&
            V5 >= 1 && V5 <= 49 &&
            V1 != V2 &&
            V1 != V3 &&
            V1 != V4 &&
            V1 != V5 &&
            V2 != V3 &&
            V2 != V4 &&
            V2 != V5 &&
            V3 != V4 &&
            V3 != V5 &&
            V4 != V5;

        }

        public bool IsOut()
        {
            return
            !(
            V1 >= 1 && V1 <= 49 &&
            V2 >= 1 && V2 <= 49 &&
            V3 >= 1 && V3 <= 49 &&
            V4 >= 1 && V4 <= 49 &&
            V5 >= 1 && V5 <= 49);
        }

        public override string ToString()
        {
            return string.Format(
            "{0},{1},{2},{3},{4}",
            V1, V2, V3, V4, V5);
        }



    }

    class ListResults : List<Results> { }



}
