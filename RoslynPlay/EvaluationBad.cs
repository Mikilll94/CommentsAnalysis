namespace RoslynPlay
{
    public class EvaluationBad
    {
        private Metrics _metrics;

        public EvaluationBad(Metrics metrics)
        {
            _metrics = metrics;
        }

        public bool IsBad()
        {
            return HasNothing() == true || HasQuestionMark() == true ||
                HasExclamationMark() == true || HasCode() == true ||
                CoherenceCoefficient() == true || WordsCount() == true;
        }

        public bool? HasNothing()
        {
            return _metrics.HasNothing;
        }

        public bool? HasQuestionMark()
        {
            return _metrics.HasQuestionMark;
        }

        public bool? HasExclamationMark()
        {
            return _metrics.HasExclamationMark;
        }

        public bool? HasCode()
        {
            return _metrics.HasCode;
        }

        public bool? CoherenceCoefficient()
        {
            if (_metrics.CoherenceCoefficient == null) return null;

            return _metrics.CoherenceCoefficient == 0 || _metrics.CoherenceCoefficient > 0.5;
        }

        public bool? WordsCount()
        {
            if (_metrics.WordsCount == null || _metrics.WordsCount == 0) return null;

            return _metrics.WordsCount <= 2 || _metrics.WordsCount > 30;
        }
    }
}
