namespace ToggJirazer.Services;

/// <summary>Builds the Budget/Account field dropdowns of the setup wizard and maps a choice back to a setting.</summary>
public static class FieldChoices
{
    public const string None = "(none)";
    private const string CurrentSuffix = " (current)";

    /// <param name="currentName">
    /// The saved field name. Empty or whitespace means "(none)" was chosen and selects "(none)";
    /// null means there is no saved choice and selects <paramref name="defaultName"/> when listed.
    /// </param>
    public static (IReadOnlyList<string> Items, int SelectedIndex) Build(
        IReadOnlyList<string> siteFieldNames, string? currentName, string defaultName)
    {
        var items = new List<string> { None };
        items.AddRange(siteFieldNames.Distinct(StringComparer.OrdinalIgnoreCase));

        if (currentName != null)
        {
            if (string.IsNullOrWhiteSpace(currentName)) return (items, 0);

            var index = items.FindIndex(1, i => string.Equals(i, currentName, StringComparison.OrdinalIgnoreCase));
            if (index >= 0) return (items, index);

            items.Add(currentName + CurrentSuffix);
            return (items, items.Count - 1);
        }

        var defaultIndex = string.IsNullOrEmpty(defaultName)
            ? -1
            : items.FindIndex(1, i => string.Equals(i, defaultName, StringComparison.OrdinalIgnoreCase));
        return (items, Math.Max(defaultIndex, 0));
    }

    public static string ToSetting(string item)
    {
        if (item == None) return "";
        return item.EndsWith(CurrentSuffix, StringComparison.Ordinal) ? item[..^CurrentSuffix.Length] : item;
    }
}
