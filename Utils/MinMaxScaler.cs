using Numpy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Keras_Stock_Prediction.Utils
{
    public class MinMaxScaler
    {
        private readonly double RangeMin;
        private readonly double RangeMax;
        private readonly double finalRange;

        private double dataMin;
        private double dataMax;

        public MinMaxScaler(double rangeMin, double rangeMax)
        {
            RangeMin = rangeMin;
            RangeMax = rangeMax;
            finalRange = RangeMax - RangeMin;
        }

        public void Fit(IEnumerable<double> data)
        {
            dataMin = data.Min();
            dataMax = data.Max();
        }

        public double Transform(double data)
        {
            var x = (data - dataMin) / (dataMax - dataMin);
            x = x * finalRange + RangeMin;

            return x;
        }

        public double[] FitTransform(double[] data)
        {
            Fit(data);
            double[] res = new double[data.Length];
            for (int i = 0; i < data.Length; i++)
            {
                res[i] = Transform(data[i]);
            }
            return res;
        }

        public double[] Reverse(double[] data)
        {
            double[] res = new double[data.Length];
            for (int i = 0; i < data.Length; i++)
            {
                res[i] = Reverse(data[i]);
            }
            return res;
        }

        public float[] Reverse(float[] data)
        {
            float[] res = new float[data.Length];
            for (int i = 0; i < data.Length; i++)
            {
                res[i] = Reverse(data[i]);
            }
            return res;
        }


        public float[] Reverse(NDarray data)
        {
            float[] array = data.GetData<float>();

            return Reverse(array);
        }

        public float Reverse(float data)
        {
            data = (float)((data - RangeMin) / finalRange);
            data = (float)(data * (dataMax - dataMin) + dataMin);


            return data;
        }

        public double Reverse(double data)
        {
            data = (data - RangeMin) / finalRange;
            data = data * (dataMax - dataMin) + dataMin;


            return data;
        }

        public override string ToString()
        {
            return $"Range: {RangeMax}-{RangeMin}, Data: {dataMax}-{dataMin}";
        }
    }
}
