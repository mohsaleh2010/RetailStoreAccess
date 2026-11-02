namespace RetailStore.Domain.Common;

public static class MoneyMath
{
    /// <summary>modCommon.RoundMoney: half away from zero (2.345 -> 2.35), not banker's rounding.</summary>
    public static decimal RoundMoney(decimal value, int digits = 2) => Math.Round(value, digits, MidpointRounding.AwayFromZero);

    /// <summary>CCur: Access keeps 4 decimals.</summary>
    public static decimal Currency(decimal value) => Math.Round(value, 4, MidpointRounding.ToEven);
}
