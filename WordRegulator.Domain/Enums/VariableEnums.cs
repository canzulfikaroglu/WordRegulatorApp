namespace WordRegulator.Domain.Enums;

public enum VariableCategory
{
    OnKapak,
    VeriKapak,
    Rapor
}

public enum VariableType
{
    Independent, // Bağımsız (Kullanıcı manuel girer)
    Dependent    // Bağımlı (Otomatik hesaplanır/birleştirilir, istenirse düzenlenebilir)
}
