package app;

import service.ExcellService;

public class Actividad6 {

    private static final String FILE_PATH = "/home/angel/Documentos/superheroes.xlsx";

    public static void main(String[] args) {
        ExcellService excellService = new ExcellService();
        excellService.creaExcellSuperHeroes(FILE_PATH);
    }
}
