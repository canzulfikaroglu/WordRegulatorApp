using WordRegulator.Domain.Entities;

namespace WordRegulator.Application.Services;

public static class ResidentialSuitabilityRepository
{
    public static IReadOnlyList<ResidentialSuitabilityDefinition> All { get; } = new List<ResidentialSuitabilityDefinition>
    {
        new(
            name: "Uygun Alan(UA)",
            description: """
            Uygun Alan(UA)
            """
        ),
          new(
            name: "Önlemli Alan-a(ÖA-a)",
            description: """
            Önlemli Alan-a(ÖA-a)
            """
        ),
           new(
            name: "Önlemli Alan-b(ÖA-b)",
            description: """
            Önlemli Alan-b(ÖA-b)
            """
        ),
           new(
            name: "Önlemli Alan-5b(ÖA-5b)",
            description: """
            Önlemli Alan-5b(ÖA-5b)
            """
        ),
           new(
            name: "Önlemli Alan-5a(ÖA-5a)",
            description: """
            Önlemli Alan-5a(ÖA-5a)
            """
        ),
           new(
            name: "Önlemli Alan-1a(ÖA-1a)",
            description: """
            Önlemli Alan-1a(ÖA-1a)
            """
        ),
           new(
            name: "Önlemli Alan-1b(ÖA-1b)",
            description: """
            Önlemli Alan-1b(ÖA-1b)
            """
        ),
    };
}
