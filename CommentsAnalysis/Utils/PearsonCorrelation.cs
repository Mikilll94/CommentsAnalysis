using System;
using System.Collections.Generic;
using System.Linq;

namespace CommentsAnalysis.Utils
{
    public static class PearsonCorrelation
    {
        public static double Compute(List<double> x, List<double> y)
        {
            if (x.Count != y.Count)
            {
                throw new Exception("Lists have different number of elements");
            }

            double n = x.Count;

            double sum_xy = 0;
            for (int i = 0; i < x.Count; i++)
            {
                sum_xy += (x[i] * y[i]);
            }

            double sum_x = x.Sum();
            double sum_y = y.Sum();

            double sum_x_2 = x.Select(el => el * el).Sum();
            double sum_y_2 = y.Select(el => el * el).Sum();

            return (n * sum_xy - sum_x * sum_y) / Math.Sqrt((n * sum_x_2 - sum_x * sum_x) * (n * sum_y_2 - sum_y * sum_y));
        }
    }
}
