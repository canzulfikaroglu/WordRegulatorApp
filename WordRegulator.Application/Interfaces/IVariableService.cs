using WordRegulator.Domain.Entities;

namespace WordRegulator.Application.Interfaces;

public interface IVariableService
{
    IReadOnlyList<VariableDefinition> GetDefaultDefinitions();
    IDictionary<string, string> RecalculateDependentVariables(IDictionary<string, string> currentValues, HashSet<string> userOverriddenKeys);
}
