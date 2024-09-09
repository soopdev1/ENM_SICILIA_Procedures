package testerclass;

import java.io.FileWriter;
import java.io.PrintWriter;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.concurrent.atomic.AtomicInteger;
import org.joda.time.DateTime;
import org.joda.time.Years;
import static rc.so.exe.Constant.estraiEccezione;
import rc.so.exe.Db_Gest;
import rc.so.exe.Sicilia_gestione;
import static rc.so.exe.Sicilia_gestione.log;

/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
/**
 *
 * @author Administrator
 */
public class ExportClaudio {

    private static final String SEP = "|";
    private static final String DF = "yyyyMMddHHmmssSSS";

    public static void main(String[] args) {
        DateTime now = new DateTime();
        Sicilia_gestione tg = new Sicilia_gestione(false);

        Db_Gest db = new Db_Gest(tg.host);

        try {

            //AULA NORMALE
            String sql3 = "SELECT s.idsedi,s.indirizzo,c.nome,a.ragionesociale,c.cod_provincia FROM sedi_formazione s, progetti_formativi p, soggetti_attuatori a, comuni c "
                    + "WHERE p.sedefisica IS NOT NULL AND s.idsedi = p.sedefisica AND a.idsoggetti_attuatori=s.idsoggetti_attuatori AND c.idcomune=s.comune GROUP BY s.idsedi";
            try (PreparedStatement ps3 = db.getConnection().prepareStatement(sql3); ResultSet rs3 = ps3.executeQuery(); 
                    FileWriter fileWriter = new FileWriter("C:\\mnt\\mcn\\yisu_sicilia\\estrazioni\\" + now.toString(DF) + "_AULA.csv"); PrintWriter printWriter = new PrintWriter(fileWriter)) {
                printWriter.println("ID AULA|INDIRIZZO|COMUNE|PROVINCIA|RAGIONE SOCIALE SOGGETTO ATTUATORE|CORSI ATTIVI|CORSI CONCLUSI|NUM ALLIEVI");
                while (rs3.next()) {
                    String sede_id = rs3.getString(1);
                    String sede_addr = rs3.getString(2).toUpperCase().trim();
                    String sede_comune = rs3.getString(3).toUpperCase().trim();
                    String sede_rag_soc = rs3.getString(4).toUpperCase().trim();
                    String sede_provincia = rs3.getString(5).toUpperCase().trim();

                    String sql4 = "SELECT p.idprogetti_formativi,p.stato FROM progetti_formativi p WHERE p.sedefisica = ?";

                    AtomicInteger sede_pf_incorso = new AtomicInteger(0);
                    AtomicInteger sede_a = new AtomicInteger(0);
                    AtomicInteger sede_pf_conclusi = new AtomicInteger(0);
                    try (PreparedStatement ps4 = db.getConnection().prepareStatement(sql4)) {
                        ps4.setString(1, sede_id);
                        try (ResultSet rs4 = ps4.executeQuery()) {
                            while (rs4.next()) {
                                String pf_id = rs4.getString(1);
                                String pf_stato = rs4.getString(2);
                                String sql6 = "SELECT a.idprogetti_formativi FROM allievi a WHERE a.idprogetti_formativi = ?";
                                try (PreparedStatement ps6 = db.getConnection().prepareStatement(sql6)) {
                                    ps6.setString(1, pf_id);
                                    try (ResultSet rs6 = ps6.executeQuery()) {
                                        while (rs6.next()) {
                                            sede_a.addAndGet(1);
                                        }
                                    }
                                }
                                switch (pf_stato) {
                                    case "F", "DBV", "IV", "CK", "EVI", "CO" ->
                                        sede_pf_conclusi.addAndGet(1);
                                    default ->
                                        sede_pf_incorso.addAndGet(1);
                                }
                            }
                        }
                    }

                    printWriter.println(sede_id + SEP + sede_addr + SEP + sede_comune + SEP + sede_provincia + SEP + sede_rag_soc + SEP + sede_pf_incorso.get()
                            + SEP + sede_pf_conclusi.get() + SEP + sede_a.get());
                    System.out.println(sede_id + SEP + sede_addr + SEP + sede_comune + SEP + sede_provincia + SEP + sede_rag_soc + SEP + sede_pf_incorso.get()
                            + SEP + sede_pf_conclusi.get() + SEP + sede_a.get());
                }
            }

            //SOGGETTI ATTUATORI
            String sql7 = "SELECT s.idsoggetti_attuatori,s.indirizzo,c.nome,s.ragionesociale,c.cod_provincia FROM soggetti_attuatori s, comuni c WHERE c.idcomune=s.comune";

            try (PreparedStatement ps7 = db.getConnection().prepareStatement(sql7); ResultSet rs7 = ps7.executeQuery(); FileWriter fileWriter = new FileWriter("C:\\mnt\\mcn\\yisu_sicilia\\estrazioni\\" + now.toString(DF) + "_SOGGETTIATTUATORI.csv"); PrintWriter printWriter = new PrintWriter(fileWriter)) {
                printWriter.println("ID SOGGETTO ATTUATORE|INDIRIZZO|COMUNE|PROVINCIA|RAGIONE SOCIALE SOGGETTO ATTUATORE|CORSI ATTIVI|CORSI CONCLUSI|NUM ALLIEVI");
                while (rs7.next()) {
                    String sa_id = rs7.getString(1);
                    String fad_addr = rs7.getString(2).toUpperCase().trim();
                    String fad_comune = rs7.getString(3).toUpperCase().trim();
                    String fad_rag_soc = rs7.getString(4).toUpperCase().trim();
                    String fad_provincia = rs7.getString(5).toUpperCase().trim();

                    String sql8 = "SELECT p.idprogetti_formativi,p.stato FROM progetti_formativi p WHERE p.idsoggetti_attuatori = ?";

                    AtomicInteger fad_pf_incorso = new AtomicInteger(0);
                    AtomicInteger fad_a = new AtomicInteger(0);
                    AtomicInteger fad_pf_conclusi = new AtomicInteger(0);

                    try (PreparedStatement ps8 = db.getConnection().prepareStatement(sql8)) {
                        ps8.setString(1, sa_id);

                        try (ResultSet rs8 = ps8.executeQuery()) {
                            while (rs8.next()) {
                                String pf_id = rs8.getString(1);
                                String pf_stato = rs8.getString(2);

//                                String sql9 = "SELECT SUM(lc.ore) FROM lezioni_modelli lm, modelli_progetti m, lezione_calendario lc "
//                                        + "WHERE m.id_modello=lm.id_modelli_progetto AND lm.tipolez='F' AND lc.id_lezionecalendario=lm.id_lezionecalendario "
//                                        + "AND m.id_progettoformativo = ?";
//                                try (PreparedStatement ps9 = db.getConnection().prepareStatement(sql9)) {
//                                    ps9.setString(1, pf_id);
//
//                                    try (ResultSet rs9 = ps9.executeQuery()) {
//                                        if (rs9.next()) {
//                                            switch (pf_stato) {
//                                                case "F", "DBV", "IV", "CK", "EVI", "CO" ->
//                                                    fad_ore_conclusi.addAndGet(rs9.getDouble(1));
//                                                default ->
//                                                    fad_ore_incorso.addAndGet(rs9.getDouble(1));
//                                            }
//                                        }
//                                    }
//                                }
                                String sql10 = "SELECT a.idprogetti_formativi FROM allievi a WHERE a.idprogetti_formativi = ?";

                                try (PreparedStatement ps10 = db.getConnection().prepareStatement(sql10)) {
                                    ps10.setString(1, pf_id);
                                    try (ResultSet rs10 = ps10.executeQuery()) {
                                        while (rs10.next()) {
                                            fad_a.addAndGet(1);
                                        }
                                    }
                                }

                                switch (pf_stato) {
                                    case "F", "DBV", "IV", "CK", "EVI", "CO" ->
                                        fad_pf_conclusi.addAndGet(1);
                                    default ->
                                        fad_pf_incorso.addAndGet(1);
                                }

                            }
                        }
                    }

//                    if (fad_ag.get() > 0 || fad_ap.get() > 0) {
                    printWriter.println(sa_id + SEP + fad_addr + SEP + fad_comune + SEP + fad_provincia + SEP + fad_rag_soc + SEP + fad_pf_incorso.get()
                            + SEP + fad_pf_conclusi.get() + SEP + fad_a.get());
                    System.out.println(sa_id + SEP + fad_addr + SEP + fad_comune + SEP + fad_provincia + SEP + fad_rag_soc + SEP + fad_pf_incorso.get()
                            + SEP + fad_pf_conclusi.get() + SEP + fad_a.get());

                    //                  }
                }
            }

            String sql11 = "SELECT a.idallievi,a.sesso,a.comune_residenza,a.comune_domicilio,a.datanascita,s.descrizione,a.idsoggetto_attuatore FROM allievi a, stato_partecipazione s WHERE a.id_statopartecipazione=s.codice";

            try (PreparedStatement ps11 = db.getConnection().prepareStatement(sql11); ResultSet rs11 = ps11.executeQuery(); 
                    FileWriter fileWriter = new FileWriter("C:\\mnt\\mcn\\yisu_sicilia\\estrazioni\\" + now.toString(DF) + "_ALLIEVI.csv"); PrintWriter printWriter = new PrintWriter(fileWriter)) {
                
                printWriter.println("ID ALLIEVO|CITTA' DOMICILIO|PROVINCIA|ETA'|SESSO|STATO|ID SOGGETTO ATTUATORE|RAGIONE SOCIALE SOGGETTO ATTUATORE");

                while (rs11.next()) {
                    String allievo_id = rs11.getString(1);
                    String allievo_sesso = rs11.getString(2);
                    String allievo_idcomune = (rs11.getString(4) == null) ? rs11.getString(3) : rs11.getString(4);
                    String allievo_provincia = "";
                    String allievo_eta = String.valueOf(Years.yearsBetween(new DateTime(rs11.getDate(5).getTime()), new DateTime()).getYears());
                    String allievo_stato = rs11.getString(6).toUpperCase().trim();
                    String allievo_sa = (rs11.getString(7) == null) ? "NON ASSEGNATO" : rs11.getString(7);
                    String allievo_saragsoc = "NON ASSEGNATO";

                    String sql12 = "SELECT c.nome,c.cod_provincia FROM comuni c WHERE c.idcomune = ?";

                    try (PreparedStatement ps12 = db.getConnection().prepareStatement(sql12)) {
                        ps12.setString(1, allievo_idcomune);
                        try (ResultSet rs12 = ps12.executeQuery()) {
                            if (rs12.next()) {
                                allievo_idcomune = rs12.getString(1);
                                allievo_provincia = rs12.getString(2);
                            }
                        }
                    }
                    if (!allievo_sa.equals("NON ASSEGNATO")) {
                        String sql13 = "SELECT s.ragionesociale FROM soggetti_attuatori s WHERE s.idsoggetti_attuatori= ?";
                        try (PreparedStatement ps13 = db.getConnection().prepareStatement(sql13)) {
                            ps13.setString(1, allievo_sa);
                            try (ResultSet rs13 = ps13.executeQuery()) {
                                if (rs13.next()) {
                                    allievo_saragsoc = rs13.getString(1).toUpperCase().trim();
                                }
                            }
                        }
                    }

                    System.out.println(allievo_id + SEP + allievo_idcomune + SEP + allievo_provincia + SEP + allievo_eta + SEP + allievo_sesso
                            + SEP + allievo_stato + SEP + allievo_sa + SEP + allievo_saragsoc);
                    printWriter.println(allievo_id + SEP + allievo_idcomune + SEP + allievo_provincia + SEP + allievo_eta + SEP + allievo_sesso
                            + SEP + allievo_stato + SEP + allievo_sa + SEP + allievo_saragsoc);
                }

            }

        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }

        db.closeDB();

    }
}
