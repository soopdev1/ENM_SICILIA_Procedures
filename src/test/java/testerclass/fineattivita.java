package testerclass;

import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import rc.so.exe.Constant;
import static rc.so.exe.Constant.estraiEccezione;
import rc.so.exe.Db_Gest;
import rc.so.exe.Sicilia_gestione;
import static rc.so.exe.Sicilia_gestione.log;
import rc.so.exe.Utils;
import static rc.so.exe.Utils.parseIntR;

/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
/**
 *
 * @author Administrator
 */
public class fineattivita {

    public static void main(String[] args) {

        Sicilia_gestione tg = new Sicilia_gestione(false);
        Db_Gest db = new Db_Gest(tg.host);
        int soglia = parseIntR(db.getPath("fc.sogliaore"));
        String sql0 = "SELECT p.idprogetti_formativi FROM progetti_formativi p WHERE p.stato = 'ATB' AND p.`end` < CURDATE()";
        try (Statement st0 = db.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0)) {
            while (rs0.next()) {
                String idpr = rs0.getString(1);
                String sql1 = "SELECT a.idallievi,a.gruppo_faseB,a.idprogetti_formativi,a.orec_faseb FROM allievi a "
                        + "WHERE a.id_statopartecipazione = '15' "
                        + "AND a.idprogetti_formativi = " + idpr;
                boolean fineattivita = true;
                boolean okallievi = false;
                List<String> idlezioniconvalidate = new ArrayList<>();
                try (Statement st1 = db.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                    while (rs1.next()) {
                        String idallievi = rs1.getString(1);
                        String gruppo_fb = rs1.getString(2);

                        try {
                            if (Double.compare(rs1.getDouble(4), soglia) == 0) {
                                okallievi = true;
                            }
                        } catch (Exception ex1) {
                            log.severe(estraiEccezione(ex1));
                        }

                        String sql2 = "SELECT la.convalidata,p.datalezione,p.idlezioneriferimento FROM presenzelezioniallievi la, presenzelezioni p WHERE la.idallievi=" + idallievi + " AND p.idpresenzelezioni=la.idpresenzelezioni "
                                + "AND p.idlezioneriferimento IN (SELECT m.id_lezionimodelli FROM lezioni_modelli m WHERE m.id_modelli_progetto IN "
                                + "(SELECT mp.id_modello FROM modelli_progetti mp WHERE mp.id_progettoformativo = " + idpr + ") AND m.gruppo_faseB IN (0," + gruppo_fb + "))"
                                + " ORDER BY p.datalezione ASC, la.convalidata DESC";
                        try (Statement st2 = db.getConnection().createStatement(); ResultSet rs2 = st2.executeQuery(sql2)) {
                            while (rs2.next()) {
                                boolean conv = rs2.getBoolean(1);
                                if (conv) {
                                    idlezioniconvalidate.add(rs2.getString(3));
                                } else {
                                    if (!idlezioniconvalidate.contains(rs2.getString(3))) {
                                        System.out.println(idallievi + " PR " + idpr + " NON CONVALIDATA () " + rs2.getString(2));
                                        fineattivita = false;
                                        break;
                                    }
                                }
                            }
                        } catch (Exception ex2) {
                            log.severe(estraiEccezione(ex2));
                        }
                    }
                } catch (Exception ex2) {
                    log.severe(estraiEccezione(ex2));
                }
                System.out.println(idpr + " fineattivita ---> " + fineattivita);

                if (fineattivita) {
                    System.out.println(idpr + " okallievi ---> " + okallievi);
                    if (okallievi) {

                    } else {

                    }
                }

            }
        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }
        db.closeDB();
    }
}
