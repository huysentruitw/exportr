namespace Exportr
{
    public static class Formula
    {
        public static Formula<T> Create<T>(string expression)
            => new Formula<T>(expression);
    }

    public sealed class Formula<TCell> : IFormula
    {
        internal Formula(string expression)
        {
            Expression = expression;
        }

        public string Expression { get; }

        public object DefaultCellValue => default(TCell);

        public override string ToString() => Expression;
    }

    public interface IFormula
    {
        string Expression { get; }

        object DefaultCellValue { get; }
    }
}
