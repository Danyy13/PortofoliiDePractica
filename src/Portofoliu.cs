public class Portofoliu
{
    private int id;
    private StagiuDePractica stagiuDePractica;
    private StudentPracticant studentPracticant;
    private Tutore tutore;
    private ProfesorCoordonator profesorCoordonator;

    public Portofoliu(int id, StagiuDePractica stagiuDePractica, StudentPracticant studentPracticant, Tutore tutore, ProfesorCoordonator profesorCoordonator)
    {
        this.id = id;
        this.stagiuDePractica = stagiuDePractica;
        this.studentPracticant = studentPracticant;
        this.tutore = tutore;
        this.profesorCoordonator = profesorCoordonator;
    }
}