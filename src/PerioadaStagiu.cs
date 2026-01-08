public struct PerioadaStagiu
{
    private DateOnly dataStart;
    private DateOnly dataFinal;
    private int orePeZi;
    private TimeOnly oraStart;  
    private TimeOnly oraFinal;

    public PerioadaStagiu(DateOnly dataStart, DateOnly dataFinal, int orePeZi, TimeOnly oraStart, TimeOnly oraFinal)
    {
        this.dataStart = dataStart;
        this.dataFinal = dataFinal;
        this.orePeZi = orePeZi;
        this.oraStart = oraStart;
        this.oraFinal = oraFinal;
    }

    public override string ToString()
    {
        return $"{dataStart} - {dataFinal}; {orePeZi}h/zi; {oraStart} - {oraFinal}";
    }
}