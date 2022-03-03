using Numpy;
using Python.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Keras_Stock_Prediction.Utils
{
    public static class Extensions
    {
        public const string create_dataset =
@"
import numpy
# convert an array of values into a dataset matrix
def create_dataset(dataset, time_step=1):
	dataX, dataY = [], []
	for i in range(len(dataset)-time_step-1):
		a = dataset[i:(i+time_step), 0]   ###i=0, X=0,1,2,3-----99   Y=100 
		dataX.append(a)
		dataY.append(dataset[i + time_step, 0])
	return numpy.array(dataX), numpy.array(dataY)
X_data, Y_data = create_dataset(data, timestep)";

        public static (NDarray X_data, NDarray Y_data) Create_Dataset(this NDarray nDarray, int timestep)
        {
            NDarray X_data;
            NDarray Y_data;

            using (Py.GIL())
            {
                using (PyScope scope = Py.CreateScope())
                {
                    scope.Set("data", nDarray.self.ToPython());
                    scope.Set("timestep", timestep.ToPython());
                    scope.Exec(create_dataset);
                    X_data = new NDarray((PyObject)scope.Get<object>("X_data"));
                    Y_data = new NDarray((PyObject)scope.Get<object>("Y_data"));
                }
            }
            return (X_data, Y_data);
        }

    }
}
