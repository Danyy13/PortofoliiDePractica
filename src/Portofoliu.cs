public class Portofoliu
{
    private int id;
    public StagiuDePractica stagiuDePractica { get; }
    public StudentPracticant studentPracticant { get; }
    public Tutore tutore { get; }
    public ProfesorCoordonator profesorCoordonator { get; }
    public static int idCount = 0;

    public Portofoliu(StagiuDePractica stagiuDePractica, StudentPracticant studentPracticant, Tutore tutore, ProfesorCoordonator profesorCoordonator)
    {
        this.id = idCount++; // id-urile sunt asignate automat crescator incepand de la 0
        this.stagiuDePractica = stagiuDePractica;
        this.studentPracticant = studentPracticant;
        this.tutore = tutore;
        this.profesorCoordonator = profesorCoordonator;
    }
}