using Microsoft.EntityFrameworkCore;

namespace RetailStore.Data.Services;

public static class NumberSequences
{
    /// <summary>modCommon.NextNumber: the next number of a sequence (prefix + padded value), saved by the caller.</summary>
    public static async Task<string> NextNumberAsync(RetailDbContext db, string sequenceName)
    {
        var s = await db.Sequences.FirstOrDefaultAsync(x => x.SequenceName == sequenceName)
                ?? throw new InvalidOperationException("Sequence not found: " + sequenceName);
        int n = s.NextValue;
        s.NextValue = n + 1;
        await db.SaveChangesAsync();
        int pad = s.PadLength;
        return (s.Prefix ?? "") + (pad > 0 ? n.ToString(new string('0', pad)) : n.ToString());
    }
}
