namespace RetailStore.Domain.Journal;

/// <summary>The place of an account in the tree: its level, the sort key (each level code on 10 digits) and the codes
/// of levels 1..5.</summary>
public sealed record TreePlace(int Level, string TreeKey, int?[] LevelCodes);

/// <summary>modAccounts.RebuildAccountTree (tools/sim.py rebuild_account_tree) without the database.</summary>
public static class AccountTree
{
    public const int MaxLevels = 5;

    /// <param name="parents">AccountCode -> ParentCode (0 = none).</param>
    /// <returns>The place of each account; the accounts in a loop or deeper than 5 levels are in Problems.</returns>
    public static (Dictionary<int, TreePlace> Places, List<int> Problems) Build(IReadOnlyDictionary<int, int> parents)
    {
        var places = new Dictionary<int, TreePlace>();
        var problems = new List<int>();
        foreach (int account in parents.Keys)
        {
            var chain = new List<int>();
            int code = account;
            while (code != 0 && chain.Count < MaxLevels + 1)
            {
                chain.Add(code);
                code = parents[code];
                if (code != 0 && !parents.ContainsKey(code))
                {
                    code = 0;                         // a parent that does not exist ends the chain
                }
            }
            if (chain.Count > MaxLevels)
            {
                problems.Add(account);
                continue;
            }
            chain.Reverse();                          // level 1 first
            var levels = new int?[MaxLevels];
            for (int k = 0; k < chain.Count; k++)
            {
                levels[k] = chain[k];
            }
            places[account] = new TreePlace(chain.Count, string.Concat(chain.Select(c => c.ToString("0000000000"))), levels);
        }
        return (places, problems);
    }
}
