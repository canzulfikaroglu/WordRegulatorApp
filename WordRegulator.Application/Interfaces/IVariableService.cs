using WordRegulator.Domain.Entities;

namespace WordRegulator.Application.Interfaces;

public interface IVariableService
{
    IReadOnlyList<VariableDefinition> GetDefaultDefinitions();
    IDictionary<string, string> RecalculateDependentVariables(IDictionary<string, string> currentValues, HashSet<string> userOverriddenKeys);
    string CalculateVs30(IEnumerable<Vs30LayerItem> layers, double df);
    (string totalDrilling, string groundwaterLevel) SummarizeDrillingWells(IEnumerable<DrillingWellItem> wells);
    IDictionary<string, string> GenerateDrillingWellReplacements(IEnumerable<DrillingWellItem> wells);
    IDictionary<string, string> GenerateVs30Replacements(IEnumerable<Vs30LayerItem> layers, double df);
}
