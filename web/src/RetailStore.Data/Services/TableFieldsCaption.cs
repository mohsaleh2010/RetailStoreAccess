namespace RetailStore.Data;

public static partial class TableFields
{
    /// <summary>modAudit.FieldCaption: the Arabic caption of a field, else its name.</summary>
    public static string Caption(string table, string field) =>
        All.TryGetValue(table, out var fields) && fields.FirstOrDefault(f => f.Field == field) is { Caption: { Length: > 0 } c } ? c : field;
}
