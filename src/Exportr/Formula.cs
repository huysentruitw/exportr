namespace Exportr
{
    public sealed class Formula<TCell> : IFormula
    {
        private Formula(string expression)
        {
            Expression = expression;
        }

        public static Formula<T> Create<T>(string expression)
            => new Formula<T>(expression);

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
