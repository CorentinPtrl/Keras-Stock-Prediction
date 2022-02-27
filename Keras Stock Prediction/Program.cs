using CsvHelper;
using Keras.Layers;
using Keras.Models;
using Keras_Stock_Prediction.Utils;
using Microsoft.Office.Interop.Excel;
using Numpy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Keras_Stock_Prediction
{
    internal class Program
    {
        public static double[] getHigh(String file)
        {
            using (var reader = new StreamReader(file))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = new List<double>();
                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                    records.Add(csv.GetField<double>("High"));
                }
                return records.ToArray();
            }
        }


        public static void CreateExcelFile(String path,double[] actualPrice, float[] predictedPrice)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                return;
            }


            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Actual Price";
            xlWorkSheet.Cells[1, 2] = "Predict";
            for (int i = 0; i < actualPrice.Length; i++)
            {
                string doub = actualPrice[i].ToString().Replace(",", ".");
                xlWorkSheet.Cells[i + 2, 1] = "=" + doub;
            }
            for (int i = 0; i < predictedPrice.Length; i++)
            {
                string doub = (predictedPrice[i]).ToString().Replace(",", ".");
                xlWorkSheet.Cells[i + Math.Abs(actualPrice.Length - predictedPrice.Length), 2] = "=" + doub;
            }

            Microsoft.Office.Interop.Excel.Range chartRange;

            ChartObjects xlCharts = (ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            ChartObject myChart = (ChartObject)xlCharts.Add(10, 80, 700, 350);
            Chart chartPage = myChart.Chart;

            chartRange = xlWorkSheet.get_Range("A1", "B" + actualPrice.Length);
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = XlChartType.xlLine;

            //export chart as picture file
            chartPage.Export(@path+"/prediction_graph.png", "PNG", misValue);

            xlWorkBook.SaveAs(@path+"/prediction.xls", XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        public static void CreateExcelFile(String path, float[] predictedPrice, int days)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                return;
            }


            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Actual Price";
            xlWorkSheet.Cells[1, 2] = "Predict";
            for (int i = 0; i < predictedPrice.Length; i++)
            {
                string doub = (predictedPrice[i]).ToString().Replace(",", ".");
                xlWorkSheet.Cells[i + 2, 2] = "=" + doub;
            }

            Microsoft.Office.Interop.Excel.Range chartRange;

            ChartObjects xlCharts = (ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            ChartObject myChart = (ChartObject)xlCharts.Add(10, 80, 700, 350);
            Chart chartPage = myChart.Chart;

            chartRange = xlWorkSheet.get_Range("A1", "B" + predictedPrice.Length);
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = XlChartType.xlLine;

            //export chart as picture file
            chartPage.Export(@path+"/ForecastGraph_"+days.ToString()+"Days.png", "PNG", misValue);

            xlWorkBook.SaveAs(@path+"/Forecast_"+days.ToString()+"Days.xls", XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        public static Sequential CreateModel()
        {
            ///Lstm Premier résultat ++
            /*var model = new Sequential();
            model.Add(new LSTM(50, return_sequences: true, input_shape: new Keras.Shape(100,1)));
            model.Add(new LSTM(50, return_sequences: true));
            model.Add(new LSTM(50));
            model.Add(new Dense(1));
            */
            ///Lstm Deuxième moins bon résultat
            Sequential model = new Sequential();
            model.Add(new LSTM(50, return_sequences: true, input_shape: new Keras.Shape(100, 1)));
            model.Add(new Dropout(0.2));
            model.Add(new LSTM(50, return_sequences: true));
            model.Add(new Dropout(0.2));
            model.Add(new LSTM(50, return_sequences: true));
            model.Add(new Dropout(0.2));
            model.Add(new LSTM(50));
            model.Add(new Dropout(0.2));
            model.Add(new Dense(1));

            model.Compile(optimizer: "adam", loss: "mean_squared_error");
            return model;
        }

        public static List<float> Forecast(Sequential model, double[] ScaledData, MinMaxScaler scaler, int days)
        {
            List<double> predictMonth = new List<double>(ScaledData.TakeLast(102));
            List<float> predicted = new List<float>();
            for (int i = 0; i < days; i++)
            {
                var (X_pred, y_pred) = np.array(predictMonth.ToArray()).reshape(-1, 1).Create_Dataset(100);
                X_pred = X_pred.reshape(X_pred.shape[0], X_pred.shape[1], 1);
                NDarray prediction = model.Predict(X_pred);
                float[] array = prediction.GetData<float>();
                float[] ScaledPrediction = scaler.Reverse(array);
                predictMonth.RemoveAt(0);
                predictMonth.Add((double)array[0]);
                predicted.Add(ScaledPrediction[0]);
            }
            return predicted;
        }
        static void Main(string[] args)
        {
            Console.WriteLine("What is the path of the csv file ?");
            String SourceData = Console.ReadLine();
            Console.WriteLine("Where do you want files to be created ?");
            String ExportData = Console.ReadLine();

            double[] high = getHigh(SourceData);
            MinMaxScaler scaler = new MinMaxScaler(0, 1);
            double[] scaled = scaler.FitTransform(high);
            double[] train_data = scaled.Take((int)(scaled.Length * 0.65)).ToArray();
            double[] test_data = scaled.Skip((int)(scaled.Length * 0.65)).ToArray();
            NDarray trainNumpy = np.array(train_data).reshape(-1, 1);
            NDarray testNumpy = np.array(test_data).reshape(-1, 1);

            var (X_train, y_train) = trainNumpy.Create_Dataset(100);

            var (X_test, y_test) = testNumpy.Create_Dataset(100);

            X_train = X_train.reshape(X_train.shape[0], X_train.shape[1], 1);
            X_test = X_test.reshape(X_test.shape[0], X_test.shape[1], 1);

            Sequential model = CreateModel();

            model.Fit(X_train, y_train, validation_data: new NDarray[] { X_test, y_test }, epochs: 100, batch_size: 64, verbose: 1);

            NDarray train_predict = model.Predict(X_train);
            NDarray test_predict = model.Predict(X_test);

            float[] train_predictScaled = scaler.Reverse(train_predict);
            float[] test_predictScaled = scaler.Reverse(test_predict);
            List<float> predict = new List<float>(train_predictScaled);
            predict.AddRange(test_predictScaled);

            CreateExcelFile(ExportData, high, predict.ToArray());
            Console.WriteLine("Graph and csv file created");

            Console.WriteLine("How many days do you want to forecast ?");
            int days = int.Parse(Console.ReadLine());
            
            List<float> forecast = Forecast(model, scaled, scaler, days);

            CreateExcelFile(ExportData, forecast.ToArray(), days);

            Console.WriteLine("Graph and csv file created for Forecast");
        }
    }
}
