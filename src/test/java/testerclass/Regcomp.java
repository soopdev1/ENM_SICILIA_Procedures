package testerclass;

import java.util.ArrayList;
import static java.util.Collections.sort;
import java.util.List;
import java.util.logging.Level;
import static rc.so.exe.Constant.estraiEccezione;
import static rc.so.exe.Sicilia_gestione.log;
import rc.so.report.Complessivo;
import rc.so.report.FaseA;
import rc.so.report.FaseB;
import rc.so.report.Lezione;

/**
 *
 * @author Administrator
 */
public class Regcomp {

    public static void main(String[] args) {
        boolean testing = false;
        FaseA FA = new FaseA(testing);
        FaseB FB = new FaseB(testing);

        List<Integer> list_id_conclusi = new ArrayList<>();
        list_id_conclusi.add(482);

        Complessivo c1 = new Complessivo(FA.getHost());
        list_id_conclusi.forEach(idpr -> {
            try {
                log.log(Level.INFO, "REPORT COMPLESSIVO - IDPR {0}", idpr);

                List<Lezione> pr_a = FA.generaregistrofasea_PR(idpr, c1.getHost(), false, false, false);
                List<Lezione> pr_b = FB.generaregistrofasea_PR(idpr, c1.getHost(), false, false, false);

                List<Lezione> fad_a = FA.calcolaegeneraregistrofasea(idpr, c1.getHost(), false, true, false);
                List<Lezione> fad_b = FB.calcolaegeneraregistrofaseb(idpr, c1.getHost(), false, false, false);

                List<Lezione> ca = new ArrayList<>();
                if (pr_a != null && !pr_a.isEmpty()) {
                    ca.addAll(pr_a);
                }
                if (fad_a != null && !fad_a.isEmpty()) {
                    ca.addAll(fad_a);
                }

                List<Lezione> cb = new ArrayList<>();
                if (pr_b != null && !pr_b.isEmpty()) {
                    cb.addAll(pr_b);
                }
                if (fad_b != null && !fad_b.isEmpty()) {
                    cb.addAll(fad_b);
                }
                sort(ca, (emp1, emp2) -> emp1.getGiorno().compareTo(emp2.getGiorno()));
                sort(cb, (emp1, emp2) -> emp1.getGiorno().compareTo(emp2.getGiorno()));

                c1.registro_complessivo(idpr, c1.getHost(), ca, cb, false);

                log.log(Level.INFO, "COMPLETATO REPORT COMPLESSIVO - IDPR {0}", idpr);
            } catch (Exception e1) {
                log.severe(estraiEccezione(e1));
            }
        });
    }
}
