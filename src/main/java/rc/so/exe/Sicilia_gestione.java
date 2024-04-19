/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package rc.so.exe;

import com.google.common.base.Splitter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.codec.digest.DigestUtils;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.commons.lang3.StringUtils;
import static org.apache.commons.lang3.StringUtils.remove;
import static org.apache.commons.lang3.StringUtils.removeEnd;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.joda.time.Years;
import static rc.so.exe.SendMailJet.sendMail;
import static rc.so.exe.Utils.conf;
import static rc.so.exe.Utils.estraiEccezione;
import static rc.so.exe.Utils.getCell;
import static rc.so.exe.Utils.setCell;
import static rc.so.exe.Utils.getRow;
import static rc.so.exe.Utils.patternITA;
import static rc.so.exe.Utils.patternSql;
import static rc.so.exe.Utils.sdfITA;
import static rc.so.exe.Utils.sdfSQL;
import static rc.so.exe.Utils.timestamp;

public class Sicilia_gestione {

    ////////////////////////////////////////////////////////////////////////////
    private static final String startroom = "FADSICILIA_";
    public String host;
    boolean test;
    public static final Logger log = Utils.createLog("Sicilia_GEST_PR");

    ////////////////////////////////////////////////////////////////////////////
    public Sicilia_gestione(boolean test) {
        this.host = conf.getString("db.host") + ":3306/enm_gestione_sicilia_prod";
        this.test = test;
        if (test) {
            this.host = conf.getString("db.host") + ":3306/enm_gestione_sicilia";
        }
        log.log(Level.INFO, "HOST: {0}", this.host);
    }

    public List<String> ore_convalidateAllievi() {
        List<String> report = new ArrayList<>();
        Db_Gest db1 = new Db_Gest(this.host);
        String sql1 = "SELECT a.idallievi FROM allievi a WHERE a.id_statopartecipazione IN (10,12,13,14,15,18,19)";
        try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
            while (rs1.next()) {
                int idallievi = rs1.getInt(1);
                String sql2 = "SELECT r.totaleorerendicontabili,r.fase FROM registro_completo r WHERE r.idutente='"
                        + idallievi + "' AND r.ruolo='ALLIEVO'";
                String sql3 = "SELECT a.durataconvalidata FROM presenzelezioniallievi a WHERE a.idallievi = '"
                        + idallievi + "' AND a.convalidata=1 AND a.durataconvalidata > 0";
//                String sql3 = "SELECT p.durataconvalidata,z.codice_ud FROM presenzelezioniallievi p, presenzelezioni l, lezione_calendario z "
//                        + "WHERE p.idallievi = '" + idallievi + "' AND p.convalidata=1 AND l.idpresenzelezioni=p.idpresenzelezioni "
//                        + "AND l.idlezioneriferimento=z.id_lezionecalendario ";

                Long presenze = 0L;

                try (Statement st2 = db1.getConnection().createStatement(); ResultSet rs2 = st2.executeQuery(sql2)) {

                    while (rs2.next()) {
                        report.add(idallievi + ";" + rs2.getString(2) + ";" + rs2.getLong(1));
                        presenze += rs2.getLong(1);
                    }
                }

                try (Statement st3 = db1.getConnection().createStatement(); ResultSet rs3 = st3.executeQuery(sql3)) {
                    while (rs3.next()) {
                        Long conv = rs3.getString(1) == null ? 0L : rs3.getLong(1);
//                        report.add(idallievi + ";" + StringUtils.substring(rs3.getString(2), 0, 1) + ";" + conv);
                        presenze += conv;
                    }
                }

                double res = presenze / 3600000.00;

                String upd4 = "UPDATE allievi SET importo = '" + res + "', orec_totali = '" + res + "' WHERE idallievi='" + idallievi + "'";

                try (Statement st4 = db1.getConnection().createStatement()) {
                    st4.executeUpdate(upd4);
                }

            }

        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }
        db1.closeDB();

        return report;
    }

    ////////////////////////////////////////////////////////////////////////////
    public void verifica_stanze(int idprogetti_formativi) {
        Db_Gest db1 = new Db_Gest(this.host);
        try {
            String dataoggi = new DateTime().toString(patternSql);
            String sql1 = "SELECT ud.fase,lm.gruppo_faseB FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, progetti_formativi pf"
                    + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                    + " AND mp.id_progettoformativo=" + idprogetti_formativi
                    + " AND pf.idprogetti_formativi=mp.id_progettoformativo AND ((pf.stato='ATA' AND ud.fase='Fase A') OR (pf.stato='ATB' AND ud.fase='Fase B'))"
                    + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' ORDER BY lm.gruppo_faseB,lm.orario_start";
            try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                while (rs1.next()) {
                    String fase = rs1.getString("ud.fase");
                    if (fase.endsWith("A")) {
                        String sql2 = "SELECT nomestanza,stato FROM fad_multi WHERE idprogetti_formativi = " + idprogetti_formativi + " AND numerocorso = 1";
                        try (Statement st2 = db1.getConnection().createStatement(); ResultSet rs2 = st2.executeQuery(sql2)) {
                            String nomestanza = startroom + idprogetti_formativi + "_A1";
                            if (test) {
                                nomestanza = "TESTING_" + nomestanza;
                            }
                            if (rs2.next()) {//VERIFICO SE ATTIVA
                                nomestanza = rs2.getString("nomestanza");
                                String stato = rs2.getString("stato");
                                if (stato.equals("0")) {
                                    //UPDATE ATTIVA
                                    try (Statement st3 = db1.getConnection().createStatement()) {
                                        String upd = "UPDATE fad_multi SET stato = '1' WHERE nomestanza = '" + nomestanza + "'";
                                        st3.executeUpdate(upd);
                                    }
                                }
                            } else { //INSERISCO
                                try (Statement st3 = db1.getConnection().createStatement()) {
                                    String ins = "INSERT INTO fad_multi VALUES ('" + nomestanza + "'," + idprogetti_formativi + ",'1','" + new DateTime().toString("yyyy-MM-dd HH:mm:ss") + "','1')";
                                    st3.executeUpdate(ins);
                                }
                            }
                        }

                    } else if (fase.endsWith("B")) {
                        int gruppo_faseB = rs1.getInt("lm.gruppo_faseB");

                        String sql2 = "SELECT nomestanza,stato FROM fad_multi WHERE idprogetti_formativi = " + idprogetti_formativi + " AND numerocorso = " + gruppo_faseB;

                        try (Statement st2 = db1.getConnection().createStatement(); ResultSet rs2 = st2.executeQuery(sql2)) {

                            String nomestanza = startroom + idprogetti_formativi + "_B" + gruppo_faseB;
                            if (test) {
                                nomestanza = "TESTING_" + nomestanza;
                            }
                            if (rs2.next()) {//VERIFICO SE ATTIVA
                                nomestanza = rs2.getString("nomestanza");
                                String stato = rs2.getString("stato");
                                if (stato.equals("0")) {
                                    //UPDATE ATTIVA
                                    try (Statement st3 = db1.getConnection().createStatement()) {
                                        String upd = "UPDATE fad_multi SET stato = '1' WHERE nomestanza = '" + nomestanza + "'";
                                        st3.executeUpdate(upd);
                                    }
                                }
                            } else { //INSERISCO
                                try (Statement st3 = db1.getConnection().createStatement()) {
                                    String ins = "INSERT INTO fad_multi VALUES ('" + nomestanza + "'," + idprogetti_formativi + ",'" + gruppo_faseB + "','" + new DateTime().toString("yyyy-MM-dd HH:mm:ss") + "','1')";
                                    st3.executeUpdate(ins);
                                }
                            }
                        }
                    }

                }
            }

        } catch (Exception e) {
            log.severe(estraiEccezione(e));
        }
        db1.closeDB();
    }

    ////////////////////////////////////////////////////////////////////////////
    public void fad_allievi(int idprogetti_formativi, boolean manual) {
        DbSSO dbs = new DbSSO();
        Db_Gest db1 = new Db_Gest(this.host);
        try {
            String mailsender = db1.getPath("mailsender");
            String dataoggi = new DateTime().toString(patternSql);
            String datainvito = new DateTime().toString(patternITA);
            String sql1 = "SELECT ud.fase,lm.gruppo_faseB,f.nomestanza,ud.codice FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, fad_multi f, progetti_formativi pf"
                    + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                    + " AND mp.id_progettoformativo=" + idprogetti_formativi
                    + " AND f.idprogetti_formativi=mp.id_progettoformativo AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso)"
                    + " AND pf.idprogetti_formativi=f.idprogetti_formativi"
                    + " AND ((pf.stato='ATA' AND ud.fase='Fase A') OR (pf.stato='ATB' AND ud.fase='Fase B'))"
                    + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' GROUP BY f.nomestanza";

            if (manual) {
                sql1 = "SELECT ud.fase,lm.gruppo_faseB,f.nomestanza,ud.codice FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, fad_multi f"
                        + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                        + " AND mp.id_progettoformativo=" + idprogetti_formativi
                        + " AND f.idprogetti_formativi=mp.id_progettoformativo AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso)"
                        + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' GROUP BY f.nomestanza";
            }

            try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {

                while (rs1.next()) {

                    String fase = rs1.getString("ud.fase");
                    String nomestanza = rs1.getString("f.nomestanza");
                    String ud = rs1.getString("ud.codice");
                    String sql3;
                    if (fase.endsWith("A")) {
                        sql3 = "SELECT idallievi,email,nome,cognome,codicefiscale FROM allievi WHERE id_statopartecipazione IN (10,12,13,14,15,18,19) AND idprogetti_formativi = " + idprogetti_formativi;
                    } else if (fase.endsWith("B")) {
                        int gruppo_faseB = rs1.getInt("lm.gruppo_faseB");
                        sql3 = "SELECT idallievi,email,nome,cognome,codicefiscale FROM allievi WHERE id_statopartecipazione IN (10,12,13,14,15,18,19) AND idprogetti_formativi = " + idprogetti_formativi + " AND gruppo_faseB = " + gruppo_faseB;
                    } else {
                        continue;
                    }

                    String sql1A = "SELECT lm.orario_start,lm.orario_end FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, fad_multi f"
                            + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                            + " AND mp.id_progettoformativo=" + idprogetti_formativi
                            + " AND f.idprogetti_formativi=mp.id_progettoformativo AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso) "
                            + " AND f.nomestanza = '" + nomestanza + "'"
                            + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' ORDER BY lm.orario_start";
                    StringBuilder orainvitosb = new StringBuilder("");
                    try (Statement st1A = db1.getConnection().createStatement(); ResultSet rs1A = st1A.executeQuery(sql1A)) {
                        while (rs1A.next()) {
                            orainvitosb.append(StringUtils.substring(rs1A.getString(1), 0, 5)).append("-").append(StringUtils.substring(rs1A.getString(2), 0, 5)).append("<br>");
                        }
                    }
                    String orainvito = StringUtils.removeEnd(orainvitosb.toString(), "<br>");
                    try (Statement st3 = db1.getConnection().createStatement(); ResultSet rs3 = st3.executeQuery(sql3)) {
                        while (rs3.next()) {
                            String codicefiscale = rs3.getString("codicefiscale").toUpperCase();
                            String nomecognome = rs3.getString("nome").toUpperCase() + " " + rs3.getString("cognome").toUpperCase();
                            int idsoggetto = rs3.getInt("idallievi");
                            String email = rs3.getString("email").toLowerCase();
                            //VERIFICA
                            String sql4 = "SELECT user FROM fad_access WHERE type='S' "
                                    + "AND idprogetti_formativi = " + idprogetti_formativi + " "
                                    + "AND idsoggetto = " + idsoggetto + " "
                                    + "AND data ='" + dataoggi + "' "
                                    + "AND ud ='" + ud + "' "
                                    + "AND room = '" + nomestanza + "'";
                            try (Statement st4 = db1.getConnection().createStatement(); ResultSet rs4 = st4.executeQuery(sql4)) {
                                String user = RandomStringUtils.randomAlphabetic(8);
                                String psw = RandomStringUtils.randomAlphanumeric(6);
                                String md5psw = DigestUtils.md5Hex(psw);
                                if (!rs4.next()) {
                                    try (Statement st5 = db1.getConnection().createStatement()) {
                                        String ins = "INSERT INTO fad_access VALUES (" + idprogetti_formativi + "," + idsoggetto + ",'" + dataoggi
                                                + "','S','" + nomestanza + "','" + user + "','" + md5psw + "','" + ud + "')";
                                        st5.executeUpdate(ins);

                                        String ins_SSO = "INSERT INTO fad_access VALUES (" + idprogetti_formativi + "," + idsoggetto + ",'" + dataoggi
                                                + "','S','" + nomestanza + "','" + user + "','" + md5psw + "','" + codicefiscale + "')";
                                        log.log(Level.INFO, "SSO ALLIEVO ) {0} : {1}", new Object[]{nomecognome, dbs.executequery(ins_SSO)});
                                        log.log(Level.INFO, "NUOVE CREDENZIALI DISCENTE ) {0}", nomecognome);
                                    }
                                } else {
                                    user = rs4.getString(1);
                                    try (Statement st5 = db1.getConnection().createStatement()) {
                                        String upd = "UPDATE fad_access SET psw = '" + md5psw + "' WHERE idsoggetto = " + idsoggetto + " AND data = '" + dataoggi + "' AND ud='" + ud + "' AND type = 'S' ";
                                        st5.executeUpdate(upd);
                                        log.log(Level.INFO, "SSO ALLIEVO ) {0} : {1}", new Object[]{nomecognome, dbs.executequery(upd)});
                                        log.log(Level.INFO, "RECUPERO CREDENZIALI DISCENTE ) {0}", nomecognome);
                                    }
                                }
                                //INVIO MAIL
                                String sql5 = "SELECT oggetto,testo FROM email WHERE chiave ='fad3.0'";
                                try (Statement st5 = db1.getConnection().createStatement(); ResultSet rs5 = st5.executeQuery(sql5)) {
                                    if (rs5.next()) {
                                        String emailtesto = rs5.getString(2);
                                        String emailoggetto = rs5.getString(1);
                                        String linkweb = db1.getPath("linkfad");
                                        String linknohttpweb = remove(linkweb, "https://");
                                        linknohttpweb = remove(linknohttpweb, "http://");
                                        linknohttpweb = removeEnd(linknohttpweb, "/");

                                        emailtesto = StringUtils.replace(emailtesto, "@nomecognome", nomecognome);
                                        emailtesto = StringUtils.replace(emailtesto, "@username", user);
                                        emailtesto = StringUtils.replace(emailtesto, "@password", psw);
                                        emailtesto = StringUtils.replace(emailtesto, "@datainvito", datainvito);

                                        emailtesto = StringUtils.replace(emailtesto, "@orainvito", orainvito);
                                        emailtesto = StringUtils.replace(emailtesto, "@nomestanza", nomestanza);
                                        emailtesto = StringUtils.replace(emailtesto, "@linkweb", linkweb);
                                        emailtesto = StringUtils.replace(emailtesto, "@linknohttpweb", linknohttpweb);
                                        boolean es = sendMail(mailsender, new String[]{email}, new String[]{}, emailtesto, emailoggetto, db1, log);
                                        if (es) {
                                            log.log(Level.INFO, "MAIL DISCENTE INVIATA A : {0}", email);
                                        } else {
                                            log.log(Level.SEVERE, "MAIL DISCENTE ERROR {0}", email);
                                        }
                                    }
                                }

                            }

                        }
                    }

                }
            }
        } catch (Exception e) {
            log.severe(estraiEccezione(e));
        }
        db1.closeDB();
        dbs.closeDB();
    }

    ////////////////////////////////////////////////////////////////////////////
    public void fad_docenti(int idprogetti_formativi, boolean manual) {
        DbSSO dbs = new DbSSO();
        Db_Gest db1 = new Db_Gest(this.host);
        try {
            String mailsender = db1.getPath("mailsender");
            String dataoggi = new DateTime().toString(patternSql);
            String datainvito = new DateTime().toString(patternITA);

            String sql1 = "SELECT ud.fase,lm.gruppo_faseB,f.nomestanza,ud.codice,lm.id_docente FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, fad_multi f, progetti_formativi pf"
                    + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                    + " AND mp.id_progettoformativo=" + idprogetti_formativi
                    + " AND f.idprogetti_formativi=mp.id_progettoformativo AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso)"
                    + " AND pf.idprogetti_formativi=f.idprogetti_formativi AND ((pf.stato='ATA' AND ud.fase='Fase A') OR (pf.stato='ATB' AND ud.fase='Fase B'))"
                    + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' GROUP BY lm.id_docente,f.nomestanza";
            if (manual) {
                sql1 = "SELECT ud.fase,lm.gruppo_faseB,f.nomestanza,ud.codice,lm.id_docente FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, fad_multi f"
                        + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                        + " AND mp.id_progettoformativo=" + idprogetti_formativi
                        + " AND f.idprogetti_formativi=mp.id_progettoformativo AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso)"
                        + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' GROUP BY lm.id_docente,f.nomestanza";
            }
            try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                while (rs1.next()) {

                    int id_docente = rs1.getInt("lm.id_docente");
                    String nomestanza = rs1.getString("f.nomestanza");
                    String ud = rs1.getString("ud.codice");

                    String sql1A = "SELECT lm.orario_start,lm.orario_end FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, fad_multi f"
                            + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                            + " AND mp.id_progettoformativo=" + idprogetti_formativi
                            + " AND f.idprogetti_formativi=mp.id_progettoformativo AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso) "
                            + " AND f.nomestanza = '" + nomestanza + "'"
                            + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' AND lm.id_docente = " + id_docente + " ORDER BY lm.orario_start";
                    StringBuilder orainvitosb = new StringBuilder("");
                    try (Statement st1A = db1.getConnection().createStatement(); ResultSet rs1A = st1A.executeQuery(sql1A)) {
                        while (rs1A.next()) {
                            orainvitosb.append(StringUtils.substring(rs1A.getString(1), 0, 5)).append("-").append(StringUtils.substring(rs1A.getString(2), 0, 5)).append("<br>");
                        }
                    }
                    String orainvito = StringUtils.removeEnd(orainvitosb.toString(), "<br>");

                    String sql4 = "SELECT iddocenti,email,nome,cognome,codicefiscale FROM docenti WHERE iddocenti = " + id_docente;
                    try (Statement st4 = db1.getConnection().createStatement(); ResultSet rs4 = st4.executeQuery(sql4)) {
                        if (rs4.next()) {
                            String codicefiscale = rs4.getString("codicefiscale").toUpperCase();
                            String nomecognome = rs4.getString("nome").toUpperCase() + " " + rs4.getString("cognome").toUpperCase();
                            int idsoggetto = rs4.getInt("iddocenti");
                            String email = rs4.getString("email").toLowerCase();
                            String sql5 = "SELECT user FROM fad_access WHERE type='D' "
                                    + "AND idprogetti_formativi = " + idprogetti_formativi + " "
                                    + "AND idsoggetto = " + idsoggetto + " "
                                    + "AND data ='" + dataoggi + "' "
                                    + "AND room = '" + nomestanza + "'";
                            try (Statement st5 = db1.getConnection().createStatement(); ResultSet rs5 = st5.executeQuery(sql5)) {
                                String user = RandomStringUtils.randomAlphabetic(8);
                                String psw = RandomStringUtils.randomAlphanumeric(6);
                                String md5psw = DigestUtils.md5Hex(psw);

                                if (!rs5.next()) {
                                    //CREO CREDENZIALI
                                    try (Statement st6 = db1.getConnection().createStatement()) {
                                        String ins = "INSERT INTO fad_access VALUES (" + idprogetti_formativi + "," + idsoggetto + ",'" + dataoggi
                                                + "','D','" + nomestanza + "','" + user + "','" + md5psw + "','" + ud + "')";
                                        st6.executeUpdate(ins);
                                        String ins_SSO = "INSERT INTO fad_access VALUES (" + idprogetti_formativi + "," + idsoggetto + ",'" + dataoggi
                                                + "','D','" + nomestanza + "','" + user + "','" + md5psw + "','" + codicefiscale + "')";
                                        log.log(Level.INFO, "SSO DOCENTE ) {0} : {1}", new Object[]{nomecognome, dbs.executequery(ins_SSO)});
                                        log.log(Level.INFO, "NUOVE CREDENZIALI DOCENTE ) {0}", nomecognome);
                                    }
                                } else { //CREDENZIALI GIA presenti
                                    user = rs5.getString(1);
                                    try (Statement st6 = db1.getConnection().createStatement()) {
                                        String upd = "UPDATE fad_access SET psw = '" + md5psw + "' WHERE idsoggetto = " + idsoggetto + " AND data = '" + dataoggi + "' AND ud='" + ud
                                                + "' AND type = 'D' ";
                                        st6.executeUpdate(upd);
                                        log.log(Level.INFO, "SSO DOCENTE ) {0} : {1}", new Object[]{nomecognome, dbs.executequery(upd)});
                                        log.log(Level.INFO, "RECUPERO CREDENZIALI DOCENTE ) {0}", nomecognome);
                                    }
                                }

                                //INVIO MAIL
                                String sql6 = "SELECT oggetto,testo FROM email WHERE chiave ='fad3.0_DOCENTE'";
                                try (Statement st6 = db1.getConnection().createStatement(); ResultSet rs6 = st6.executeQuery(sql6)) {
                                    if (rs6.next()) {
                                        String emailtesto = rs6.getString(2);
                                        String emailoggetto = rs6.getString(1);

                                        String linkweb = db1.getPath("linkfad");
                                        String linknohttpweb = remove(linkweb, "https://");
                                        linknohttpweb = remove(linknohttpweb, "http://");
                                        linknohttpweb = removeEnd(linknohttpweb, "/");
//
                                        emailtesto = StringUtils.replace(emailtesto, "@nomecognome", nomecognome);
                                        emailtesto = StringUtils.replace(emailtesto, "@username", user);
                                        emailtesto = StringUtils.replace(emailtesto, "@password", psw);
                                        emailtesto = StringUtils.replace(emailtesto, "@datainvito", datainvito);
                                        emailtesto = StringUtils.replace(emailtesto, "@orainvito", orainvito);
                                        emailtesto = StringUtils.replace(emailtesto, "@nomestanza", nomestanza);
                                        emailtesto = StringUtils.replace(emailtesto, "@linkweb", linkweb);
                                        emailtesto = StringUtils.replace(emailtesto, "@linknohttpweb", linknohttpweb);
//
                                        boolean es = sendMail(mailsender, new String[]{email}, new String[]{}, emailtesto, emailoggetto, db1, log);
                                        if (es) {
                                            log.log(Level.INFO, "MAIL DOCENTE INVIATA A : {0}", email);
                                        } else {
                                            log.log(Level.SEVERE, "MAIL DOCENTE ERROR {0}", email);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.severe(estraiEccezione(e));
        }
        db1.closeDB();
        dbs.closeDB();
    }

    ////////////////////////////////////////////////////////////////////////////
    public void fad_ospiti(int idprogetti_formativi, boolean manual) {
        Db_Gest db1 = new Db_Gest(this.host);
        try {
            String mailsender = db1.getPath("mailsender");
            String dataoggi = new DateTime().toString(patternSql);
            String datainvito = new DateTime().toString(patternITA);

            String sql1 = "SELECT ud.fase,lm.gruppo_faseB,f.nomestanza,ud.codice,lm.id_docente FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, fad_multi f, progetti_formativi pf"
                    + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                    + " AND mp.id_progettoformativo=" + idprogetti_formativi
                    + " AND f.idprogetti_formativi=mp.id_progettoformativo AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso) "
                    + " AND pf.idprogetti_formativi=f.idprogetti_formativi AND ((pf.stato='ATA' AND ud.fase='Fase A') OR (pf.stato='ATB' AND ud.fase='Fase B'))"
                    + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' GROUP BY f.nomestanza";
            if (manual) {
                sql1 = "SELECT ud.fase,lm.gruppo_faseB,f.nomestanza,ud.codice,lm.id_docente FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, fad_multi f"
                        + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                        + " AND mp.id_progettoformativo=" + idprogetti_formativi
                        + " AND f.idprogetti_formativi=mp.id_progettoformativo AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso) "
                        + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' GROUP BY f.nomestanza";
            }
            try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                while (rs1.next()) {

                    String nomestanza = rs1.getString("f.nomestanza");
                    String ud = rs1.getString("ud.codice");

                    String sql1A = "SELECT lm.orario_start,lm.orario_end FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud, fad_multi f"
                            + " WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario AND ud.codice=lc.codice_ud"
                            + " AND mp.id_progettoformativo=" + idprogetti_formativi
                            + " AND f.idprogetti_formativi=mp.id_progettoformativo AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso)"
                            + " AND f.nomestanza = '" + nomestanza + "'"
                            + " AND lm.tipolez='F' AND lm.giorno = '" + dataoggi + "' ORDER BY lm.orario_start";
                    StringBuilder orainvitosb = new StringBuilder("");
                    try (Statement st1A = db1.getConnection().createStatement(); ResultSet rs1A = st1A.executeQuery(sql1A)) {
                        while (rs1A.next()) {
                            orainvitosb.append(StringUtils.substring(rs1A.getString(1), 0, 5)).append("-").append(StringUtils.substring(rs1A.getString(2), 0, 5)).append("<br>");
                        }
                    }
                    String orainvito = StringUtils.removeEnd(orainvitosb.toString(), "<br>");

                    String sql4O = "SELECT id_staff, nome, cognome, email FROM staff_modelli WHERE id_progettoformativo = " + idprogetti_formativi;
                    try (Statement st4 = db1.getConnection().createStatement(); ResultSet rs4 = st4.executeQuery(sql4O)) {
                        while (rs4.next()) {
                            String nomecognome = rs4.getString("nome").toUpperCase() + " " + rs4.getString("cognome").toUpperCase();
                            int idsoggetto = rs4.getInt("id_staff");
                            String email = rs4.getString("email").toLowerCase();
                            String sql5 = "SELECT user FROM fad_access WHERE type='O' "
                                    + "AND idprogetti_formativi = " + idprogetti_formativi + " "
                                    + "AND idsoggetto = " + idsoggetto + " "
                                    + "AND data ='" + dataoggi + "' "
                                    + "AND room = '" + nomestanza + "'";
                            try (Statement st5 = db1.getConnection().createStatement(); ResultSet rs5 = st5.executeQuery(sql5)) {
                                String user = RandomStringUtils.randomAlphabetic(8);
                                String psw = RandomStringUtils.randomAlphanumeric(6);
                                String md5psw = DigestUtils.md5Hex(psw);
                                if (!rs5.next()) {
                                    //CREO CREDENZIALI
                                    try (Statement st6 = db1.getConnection().createStatement()) {
                                        String ins = "INSERT INTO fad_access VALUES ("
                                                + idprogetti_formativi + "," + idsoggetto + ",'" + dataoggi
                                                + "','O','" + nomestanza + "','" + user + "','"
                                                + md5psw + "','" + ud + "')";

                                        st6.executeUpdate(ins);
                                    }
                                    log.log(Level.INFO, "NUOVE CREDENZIALI OSPITE ) {0}", nomecognome);
                                } else { //CREDENZIALI GIA presenti
                                    user = rs5.getString(1);
                                    try (Statement st6 = db1.getConnection().createStatement()) {
                                        String upd = "UPDATE fad_access SET psw = '"
                                                + md5psw + "' WHERE idsoggetto = "
                                                + idsoggetto + " AND data = '"
                                                + dataoggi + "' AND ud='" + ud + "' AND type = 'O'";
                                        st6.executeUpdate(upd);
                                    }
                                    log.log(Level.INFO, "RECUPERO CREDENZIALI OSPITE ) {0}", nomecognome);
                                }
                                //INVIO MAIL
                                String sql6 = "SELECT oggetto,testo FROM email WHERE chiave ='fad3.0'";
                                try (Statement st6 = db1.getConnection().createStatement(); ResultSet rs6 = st6.executeQuery(sql6)) {
                                    if (rs6.next()) {
                                        String emailtesto = rs6.getString(2);
                                        String emailoggetto = rs6.getString(1);

                                        String linkweb = db1.getPath("linkfad");
                                        String linknohttpweb = remove(linkweb, "https://");
                                        linknohttpweb = remove(linknohttpweb, "http://");
                                        linknohttpweb = removeEnd(linknohttpweb, "/");

                                        emailtesto = StringUtils.replace(emailtesto, "@nomecognome", "OSPITE " + nomecognome);
                                        emailtesto = StringUtils.replace(emailtesto, "@username", user);
                                        emailtesto = StringUtils.replace(emailtesto, "@password", psw);
                                        emailtesto = StringUtils.replace(emailtesto, "@datainvito", datainvito);
                                        emailtesto = StringUtils.replace(emailtesto, "@orainvito", orainvito);
                                        emailtesto = StringUtils.replace(emailtesto, "@nomestanza", nomestanza);
                                        emailtesto = StringUtils.replace(emailtesto, "@linkweb", linkweb);
                                        emailtesto = StringUtils.replace(emailtesto, "@linknohttpweb", linknohttpweb);

                                        boolean es = sendMail(mailsender, new String[]{email}, new String[]{}, emailtesto, emailoggetto, db1, log);
                                        if (es) {
                                            log.log(Level.INFO, "MAIL OSPITE INVIATA A : {0}", email);
                                        } else {
                                            log.log(Level.SEVERE, "MAIL OSPITE ERROR {0}", email);
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
            }
        } catch (Exception e) {
            log.severe(estraiEccezione(e));
        }
        db1.closeDB();
    }

    ////////////////////////////////////////////////////////////////////////////
    public void fad_gestione() {
        List<Integer> elenco = new ArrayList<>();
        Db_Gest db1 = new Db_Gest(this.host);
        try {
            String sql0 = "SELECT idprogetti_formativi FROM progetti_formativi "
                    + "WHERE CURDATE()>=start AND CURDATE()<=end "
                    + "AND (stato='ATA' OR stato = 'ATB')";
            try (Statement st0 = db1.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0)) {
                while (rs0.next()) {
                    elenco.add(rs0.getInt("idprogetti_formativi"));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        db1.closeDB();

        elenco.forEach(pf -> {
            try {
                log.log(Level.WARNING, "VERIFICA STANZE PF -> {0}", pf);
                this.verifica_stanze(pf);
            } catch (Exception e) {
                log.log(Level.SEVERE, "ERRORE VERIFICA STANZE PF ->{0}", pf);
                log.severe(estraiEccezione(e));
            }
            try {
                log.log(Level.WARNING, "FAD ALLIEVI -> {0}", pf);
                this.fad_allievi(pf, false);
            } catch (Exception e) {
                log.log(Level.SEVERE, "ERRORE FAD ALLIEVI PF ->{0}", pf);
                log.severe(estraiEccezione(e));
            }
            try {
                log.log(Level.WARNING, "FAD DOCENTI -> {0}", pf);
                this.fad_docenti(pf, false);
            } catch (Exception e) {
                log.log(Level.SEVERE, "ERRORE FAD DOCENTI PF ->{0}", pf);
                log.severe(estraiEccezione(e));
            }
            try {
                log.log(Level.WARNING, "FAD OSPITI -> {0}", pf);
                this.fad_ospiti(pf, false);
            } catch (Exception e) {
                log.log(Level.SEVERE, "ERRORE FAD OSPITI PF ->{0}", pf);
                log.severe(estraiEccezione(e));
            }
        });
    }

    //REPORT
    private static final String formatdataCell = "#.0";
    private static final String formatdataCellINT = "##";
    private static final String formatdataCellCUR = "€#,##0.00";

    // da completare
    public void report_allievi() {
        try {
//            List<String> presenzeconv = ore_convalidateAllievi();
            String fileing = "/mnt/mcn/yisu_sicilia/estrazioni/Report_Allievi_Sicilia.xlsx";
            String fileout = "/mnt/mcn/yisu_sicilia/estrazioni/Report_Allievi_" + new DateTime().toString(timestamp) + ".xlsx";

            String sql0 = "SELECT * FROM allievi a WHERE a.id_statopartecipazione<>'00' ORDER BY a.cognome";
            Db_Gest db1 = new Db_Gest(this.host);

            try (OutputStream outputStream = new FileOutputStream(new File(fileout)); XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(new File(fileing)))) {

                XSSFCellStyle cellStyle = wb.createCellStyle();
                cellStyle.setDataFormat(wb.createDataFormat().getFormat(formatdataCell));
                XSSFCellStyle cellStyleINT = wb.createCellStyle();
                cellStyleINT.setDataFormat(wb.createDataFormat().getFormat(formatdataCellINT));
                XSSFCellStyle cellStyleCUR = wb.createCellStyle();
                cellStyleCUR.setDataFormat(wb.createDataFormat().getFormat(formatdataCellCUR));

                HashMap<Integer, String> lista_istat = new HashMap<>();
                XSSFSheet sh1 = wb.getSheetAt(0);
                AtomicInteger maxrow = new AtomicInteger(1);
                AtomicInteger indiceriga = new AtomicInteger(1);
                try (Statement st0 = db1.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0)) {
                    while (rs0.next()) {
                        int idallievo = rs0.getInt("a.idallievi");
                        int idpr = rs0.getInt("a.idprogetti_formativi");
                        int idsa = rs0.getInt("a.idsoggetto_attuatore");
                        String sa = "";

                        String sql01 = "SELECT ragionesociale FROM soggetti_attuatori WHERE idsoggetti_attuatori = " + idsa;
                        try (ResultSet rs01 = db1.getConnection().createStatement().executeQuery(sql01)) {
                            if (rs01.next()) {
                                sa = rs01.getString(1).toUpperCase();
                            }
                        }

                        String comune_nascita = "";
                        String provincia_nascita = "";
                        String comune_residenza = "";
                        String comune_domicilio = "";
                        String provincia_residenza = "";
                        String provincia_domicilio = "";
                        String regione_residenza = "";
                        String regione_domicilio = "";

                        String codice_istat_residenza = "";
                        String codice_istat_domicilio = "";
                        String sql1 = "SELECT nome,idcomune,nome_provincia,regione FROM comuni WHERE idcomune IN ("
                                + rs0.getInt("a.comune_nascita") + ","
                                + rs0.getInt("a.comune_residenza") + ","
                                + rs0.getInt("a.comune_domicilio") + ")";

                        try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                            while (rs1.next()) {
                                if (rs1.getInt(2) == rs0.getInt("a.comune_nascita")) {
                                    comune_nascita = rs1.getString(1).toUpperCase();
                                    provincia_nascita = rs1.getString(3).toUpperCase();
                                }
                                if (rs1.getInt(2) == rs0.getInt("a.comune_residenza")) {
                                    comune_residenza = rs1.getString(1).toUpperCase();
                                    provincia_residenza = rs1.getString(3).toUpperCase();
                                    regione_residenza = rs1.getString(4).toUpperCase();

                                    if (lista_istat.get(rs1.getInt(2)) == null) {
                                        String sql2A = "SELECT t.COD_ISTAT FROM comuni c, TC16 t WHERE c.nome=t.DESCRIZIONE_COMUNE "
                                                + "AND c.regione=t.DESCRIZIONE_REGIONE AND c.cittadinanza=0 AND c.idcomune =" + rs1.getInt(2);
                                        try (ResultSet rs2A = db1.getConnection().createStatement().executeQuery(sql2A)) {
                                            if (rs2A.next()) {
                                                codice_istat_residenza = rs2A.getString(1);
                                                lista_istat.put(rs1.getInt(2), rs2A.getString(1));
                                            }
                                        }
                                    } else {
                                        codice_istat_residenza = lista_istat.get(rs1.getInt(2));
                                    }

                                }
                                if (rs1.getInt(2) == rs0.getInt("a.comune_domicilio")) {
                                    comune_domicilio = rs1.getString(1).toUpperCase();
                                    provincia_domicilio = rs1.getString(3).toUpperCase();
                                    regione_domicilio = rs1.getString(4).toUpperCase();
                                    if (lista_istat.get(rs1.getInt(2)) == null) {
                                        String sql2A = "SELECT t.COD_ISTAT FROM comuni c, TC16 t WHERE c.nome=t.DESCRIZIONE_COMUNE "
                                                + "AND c.regione=t.DESCRIZIONE_REGIONE AND c.cittadinanza=0 AND c.idcomune =" + rs1.getInt(2);
                                        try (ResultSet rs2A = db1.getConnection().createStatement().executeQuery(sql2A)) {
                                            if (rs2A.next()) {
                                                codice_istat_domicilio = rs2A.getString(1);
                                                lista_istat.put(rs1.getInt(2), rs2A.getString(1));
                                            }
                                        }
                                    } else {
                                        codice_istat_domicilio = lista_istat.get(rs1.getInt(2));
                                    }
                                }
                            }
                        }

                        String statonascita = rs0.getString("a.stato_nascita");
                        if (statonascita.equals("100")) {
                            statonascita = "ITALIA";
                        } else {
                            String sql2 = "SELECT nome FROM nazioni_rc WHERE codicefiscale = '" + statonascita + "'";
                            try (Statement st2 = db1.getConnection().createStatement(); ResultSet rs2 = st2.executeQuery(sql2)) {
                                if (rs2.next()) {
                                    statonascita = rs2.getString(1).toUpperCase();
                                }
                            }
                        }

                        String titolo_studio = rs0.getString("a.titolo_studio");
                        String sql4 = "SELECT descrizione FROM titoli_studio WHERE codice = '" + titolo_studio + "'";
                        try (Statement st4 = db1.getConnection().createStatement(); ResultSet rs4 = st4.executeQuery(sql4)) {
                            if (rs4.next()) {
                                titolo_studio = rs4.getString(1).toUpperCase();
                            }
                        }

                        String condizione_lavorativa = rs0.getString("a.idcondizione_lavorativa");
                        String sql6 = "SELECT descrizione FROM condizione_lavorativa WHERE idcondizione_lavorativa = '" + condizione_lavorativa + "'";
                        try (ResultSet rs6 = db1.getConnection().createStatement().executeQuery(sql6)) {
                            if (rs6.next()) {
                                condizione_lavorativa = rs6.getString(1).toUpperCase();
                            }
                        }

                        String canale_conoscenza = rs0.getString("a.idcanale");
                        String sql7 = "SELECT descrizione FROM canale WHERE idcanale = '" + canale_conoscenza + "'";
                        try (ResultSet rs7 = db1.getConnection().createStatement().executeQuery(sql7)) {
                            if (rs7.next()) {
                                canale_conoscenza = rs7.getString(1).toUpperCase();
                            }
                        }

                        String motivazione = rs0.getString("a.motivazione");
                        String sql8 = "SELECT descrizione FROM motivazione WHERE idmotivazione = '" + motivazione + "'";
                        try (ResultSet rs8 = db1.getConnection().createStatement().executeQuery(sql8)) {
                            if (rs8.next()) {
                                motivazione = rs8.getString(1).toUpperCase();
                            }
                        }

                        String statopartecipazione = rs0.getString("a.id_statopartecipazione");
                        String sql13 = "SELECT descrizione FROM stato_partecipazione WHERE codice = '" + statopartecipazione + "'";
                        try (Statement st13 = db1.getConnection().createStatement(); ResultSet rs13 = st13.executeQuery(sql13)) {
                            if (rs13.next()) {
                                statopartecipazione = rs13.getString(1).toUpperCase();
                            }
                        }

                        String cip = "";
                        String statopr = "";
                        String dataavvio = "";
                        String datachiusura = "";
//                        String assegnazione = "";
                        String sql10 = "SELECT p.cip,p.start,p.END,s.descrizione,p.assegnazione FROM progetti_formativi p, stati_progetto s WHERE p.stato=s.idstati_progetto AND idprogetti_formativi=" + idpr;

                        try (ResultSet rs10 = db1.getConnection().createStatement().executeQuery(sql10)) {
                            if (rs10.next()) {
                                if (rs10.getString(1) != null) {
                                    cip = rs10.getString(1).toUpperCase();
                                }
                                if (rs10.getDate(2) != null) {
                                    dataavvio = sdfITA.format(rs10.getDate(2));
                                }
                                if (rs10.getDate(3) != null) {
                                    if (new DateTime().withMillisOfDay(0).isAfter(new DateTime(rs10.getDate(3).getTime()))) {
                                        datachiusura = sdfITA.format(rs10.getDate(3));
                                    }
                                }
                                if (rs10.getString(4) != null) {
                                    statopr = rs10.getString(4).toUpperCase();
                                }
//                                if (rs10.getString(5) != null) {
//                                    assegnazione = rs10.getString(5).toUpperCase();
//                                }
                            }
                        }

                        String domiciliouguale = "NO";

                        try {
                            boolean checkdomiciliouguale = rs0.getString("a.indirizzodomicilio").equalsIgnoreCase(rs0.getString("a.indirizzoresidenza"))
                                    && rs0.getString("a.civicodomicilio").equalsIgnoreCase(rs0.getString("a.civicoresidenza"))
                                    && rs0.getInt("a.comune_domicilio") == rs0.getInt("a.comune_residenza");

                            if (checkdomiciliouguale) {
                                domiciliouguale = "SI";
                            }
                        } catch (Exception ex1) {
                            domiciliouguale = "SI";
                        }

                        String sesso = rs0.getString("a.sesso").toUpperCase();
                        if (sesso.equals("M")) {
                            sesso = "UOMO";
                        } else if (sesso.equals("F")) {
                            sesso = "DONNA";
                        }
                        String cittadinanza = rs0.getString("a.cittadinanza").toUpperCase();
                        String istat_cittadinanza = rs0.getString("a.cittadinanza").toUpperCase();

                        String sql3 = "SELECT nome,istat FROM nazioni_rc WHERE idnazione = " + cittadinanza;
                        try (ResultSet rs3 = db1.getConnection().createStatement().executeQuery(sql3)) {
                            if (rs3.next()) {
                                cittadinanza = rs3.getString(1).toUpperCase();
                                istat_cittadinanza = rs3.getString(2).toUpperCase();
                            }
                        }

                        String target = rs0.getString("a.partecipazione") == null ? "" : rs0.getString("a.partecipazione");

                        target = switch (target) {
                            case "01" ->
                                "TARGET 1 - Disoccupato/inoccupato";
                            case "02" ->
                                "TARGET 2 - Persona in condizione di disagio socio-economico";
                            case "03" ->
                                "TARGET 3 - Persona destinataria di politiche passive (CIG)";
                            default ->
                                "";
                        };

                        AtomicInteger indicecolonna = new AtomicInteger(0);
                        XSSFRow row = getRow(sh1, indiceriga.get());
                        indiceriga.addAndGet(1);
                        setCell(getCell(row, 0), String.valueOf(idallievo));
                        setCell(getCell(row, indicecolonna.addAndGet(1)), sa);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.nome").trim().toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.cognome").trim().toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getDate("a.datanascita") == null ? "" : sdfITA.format(rs0.getDate("a.datanascita")));
                        setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(Years.yearsBetween(new DateTime(rs0.getDate("a.datanascita").getTime()), new DateTime()).getYears()));
                        setCell(getCell(row, indicecolonna.addAndGet(1)), statonascita);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), comune_nascita);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), provincia_nascita);

                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.codicefiscale").trim().toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.telefono").trim().toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.email").trim().toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.indirizzoresidenza").trim().toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), comune_residenza);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), codice_istat_residenza);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.capresidenza") == null ? "" : rs0.getString("a.capresidenza").trim());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), provincia_residenza);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), regione_residenza);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), domiciliouguale);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.indirizzodomicilio") == null ? "" : rs0.getString("a.indirizzodomicilio").trim().toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), comune_domicilio);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), codice_istat_domicilio);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), provincia_domicilio);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), regione_domicilio);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), sesso);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), cittadinanza);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), istat_cittadinanza);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), target);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), titolo_studio);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), condizione_lavorativa);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), canale_conoscenza);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), motivazione);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), "SI");
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.privacy2"));
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("a.privacy3"));
                        setCell(getCell(row, indicecolonna.addAndGet(1)), statopartecipazione);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), cip);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), statopr);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), dataavvio);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), datachiusura);

                        setCell(getCell(row, indicecolonna.addAndGet(1)), cellStyle, rs0.getString("a.orec_totali") == null ? "0.0" : rs0.getString("a.orec_totali").trim(), false, true);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), cellStyle, rs0.getString("a.orec_fasea") == null ? "0.0" : rs0.getString("a.orec_fasea").trim(), false, true);

                        //MANCA MODELLO 5
                        maxrow.set(indicecolonna.get());
                    }
                }
                for (int i = 0; i < 56; i++) {
                    sh1.autoSizeColumn(i);
                }
                wb.write(outputStream);

            }
            log.log(Level.WARNING, "{0} RILASCIATO CORRETTAMENTE.", fileout);

            String upd = "UPDATE estrazioni SET path = '" + fileout + "' WHERE idestrazione=2";
            db1.getConnection().createStatement().executeUpdate(upd);
            db1.closeDB();
        } catch (Exception e) {
            log.severe(estraiEccezione(e));
        }
    }

    public void report_docenti() {
        try {

            String hostbando = conf.getString("db.host") + ":3306/enm_sicilia_prod";
            if (this.test) {
                hostbando = conf.getString("db.host") + ":3306/enm_sicilia";
            }

            String fileing = "/mnt/mcn/yisu_sicilia/estrazioni/Report_Docenti_Sicilia.xlsx";
            String fileout = "/mnt/mcn/yisu_sicilia/estrazioni/Report_Docenti_" + new DateTime().toString(timestamp) + ".xlsx";

            Db_Gest db1 = new Db_Gest(this.host);

            String sql0 = "SELECT s.ragionesociale,s.piva,s.protocollo,d.nome,d.cognome,d.codicefiscale,d.stato,d.tipo_inserimento,d.datawebinair,d.motivo,d.fascia,d.email FROM docenti d, soggetti_attuatori s WHERE d.idsoggetti_attuatori=s.idsoggetti_attuatori";

            try (OutputStream outputStream = new FileOutputStream(new File(fileout)); XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(new File(fileing)))) {

                XSSFSheet sh1 = wb.getSheetAt(0);
                AtomicInteger indiceriga = new AtomicInteger(1);
                try (Statement st0 = db1.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0)) {
                    while (rs0.next()) {

                        String stato = rs0.getString("d.stato") == null ? "" : rs0.getString("d.stato");
                        stato = switch (stato) {
                            case "A" ->
                                "ACCREDITATO";
                            case "DV" ->
                                "DA VALIDARE";
                            case "W" ->
                                "IN ATTESA WEBINAIR";
                            case "R" ->
                                "RIGETTATO";
                            default ->
                                "";
                        };

                        String tipo_inserimento;
                        String fasciaproposta = rs0.getString("d.fascia").toUpperCase();

                        if (rs0.getString("d.tipo_inserimento") == null || rs0.getString("d.tipo_inserimento").trim().equals("")
                                || rs0.getString("d.tipo_inserimento").trim().equals("-")) {
                            tipo_inserimento = "ACCREDITAMENTO";

                            Db_Accreditamento db2 = new Db_Accreditamento(hostbando);
                            String sql2 = "SELECT CONCAT('F',fascia) "
                                    + " FROM allegato_b a , bando_sicilia_mcn b WHERE a.username=b.username "
                                    + " AND b.protocollo = '" + rs0.getString("s.protocollo") + "' AND a.cf = '" + rs0.getString("d.codicefiscale") + "'";
                            try (ResultSet rs2 = db2.getConnection().createStatement().executeQuery(sql2)) {
                                if (rs2.next()) {
                                    fasciaproposta = rs2.getString(1);
                                }
                            }
                            db2.closeDB();

                        } else {
                            tipo_inserimento = "GESTIONALE";
                        }

                        AtomicInteger indicecolonna = new AtomicInteger(0);
                        XSSFRow row = getRow(sh1, indiceriga.get());
                        indiceriga.addAndGet(1);
                        setCell(getCell(row, indicecolonna.get()), rs0.getString("s.ragionesociale").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("d.nome").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("d.cognome").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("d.codicefiscale").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("d.email").toLowerCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), fasciaproposta);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("d.fascia").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getDate("d.datawebinair") == null ? "" : sdfITA.format(rs0.getDate("d.datawebinair")));
                        setCell(getCell(row, indicecolonna.addAndGet(1)), stato);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), tipo_inserimento);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("d.motivo") == null ? "" : rs0.getString("d.motivo").toUpperCase());

                    }
                }

                for (int i = 0; i < 20; i++) {
                    sh1.autoSizeColumn(i);
                }

                wb.write(outputStream);
            }
            log.log(Level.WARNING, "{0} RILASCIATO CORRETTAMENTE.", fileout);

            String upd = "UPDATE estrazioni SET path = '" + fileout + "' WHERE idestrazione=1";
            db1.getConnection().createStatement().executeUpdate(upd);
            db1.closeDB();

        } catch (Exception e) {
            log.severe(estraiEccezione(e));
        }
    }

    public void report_pf() {
        try {

            String fileing = "/mnt/mcn/yisu_sicilia/estrazioni/Report_ProgettiFormativi_Sicilia.xlsx";
            String fileout = "/mnt/mcn/yisu_sicilia/estrazioni/Report_ProgettiFormativi_Sicilia_" + new DateTime().toString(timestamp) + ".xlsx";

            Db_Gest db1 = new Db_Gest(this.host);
            try (OutputStream outputStream = new FileOutputStream(new File(fileout)); XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(new File(fileing)))) {
                //FOGLIO 1
                XSSFSheet sh1 = wb.getSheetAt(0);
                String sql0_foglio1 = "SELECT sa.ragionesociale,sa.piva,sa.idsoggetti_attuatori,sa.comune FROM soggetti_attuatori sa ORDER BY sa.ragionesociale";
                AtomicInteger indiceriga = new AtomicInteger(2);
                try (Statement st0 = db1.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0_foglio1)) {
                    while (rs0.next()) {
                        int idsa = rs0.getInt("sa.idsoggetti_attuatori");
                        Long idcomune = rs0.getLong("sa.comune");
                        String regione_sedelegale = "";
                        String sqla_foglio1 = "SELECT regione FROM comuni WHERE idcomune = " + idcomune;
                        try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sqla_foglio1)) {
                            if (rs1.next()) {
                                regione_sedelegale = rs1.getString(1).toUpperCase();
                            }
                        }

                        int docenti = 0;
                        String sql1_foglio1 = "SELECT COUNT(iddocenti) FROM docenti d WHERE stato='A' AND d.idsoggetti_attuatori=" + idsa;
                        try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1_foglio1)) {
                            if (rs1.next()) {
                                docenti = rs1.getInt(1);
                            }
                        }

                        String sql2_foglio1 = "SELECT p.idprogetti_formativi,p.stato FROM progetti_formativi p WHERE p.idsoggetti_attuatori=" + idsa;

                        int DAVALIDARE_p = 0, DAVALIDARE_a = 0;
                        int PROGRAMMATO_p = 0, PROGRAMMATO_a = 0;
                        int DACONFERMARE_p = 0, DACONFERMARE_a = 0;
                        int FASEA_p = 0, FASEA_a = 0;
                        int FASEB_p = 0, FASEB_a = 0;
                        int SOSPESO_p = 0, SOSPESO_a = 0;
                        int RIGETTATO_p = 0, RIGETTATO_a = 0;
                        int FINEATTIVITA_p = 0, FINEATTIVITA_a = 0;
                        int DAVALIDAREMODELLO6_p = 0, DAVALIDAREMODELLO6_a = 0;
                        int INATTESADIMAPPATURA_p = 0, INATTESADIMAPPATURA_a = 0;
                        int INVERIFICA_p = 0, INVERIFICA_a = 0;
                        int ESITOVERIFICACONCLUSO_p = 0, ESITOVERIFICACONCLUSO_a = 0;
                        int ESITOVERIFICAINVIATO_p = 0, ESITOVERIFICAINVIATO_a = 0;
                        int CONCLUSO_p = 0, CONCLUSO_a = 0;

                        try (ResultSet rs2 = db1.getConnection().createStatement().executeQuery(sql2_foglio1)) {

                            while (rs2.next()) {
                                String stato = rs2.getString("p.stato");
                                int idpr = rs2.getInt("p.idprogetti_formativi");
                                switch (stato) {
                                    case "DV" -> {
                                        DAVALIDARE_p++;
                                        DAVALIDARE_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "P" -> {
                                        PROGRAMMATO_p++;
                                        PROGRAMMATO_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "DC" -> {
                                        DACONFERMARE_p++;
                                        DACONFERMARE_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "ATA" -> {
                                        FASEA_p++;
                                        FASEA_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "ATB" -> {
                                        FASEB_p++;
                                        FASEB_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "SOA", "SOB" -> {
                                        SOSPESO_p++;
                                        SOSPESO_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "DCE", "DVE", "ATAE", "DVAE", "ATBE", "DVBE" -> {
                                        RIGETTATO_p++;
                                        RIGETTATO_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "F" -> {
                                        FINEATTIVITA_p++;
                                        FINEATTIVITA_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "DVB" -> {
                                        DAVALIDAREMODELLO6_p++;
                                        DAVALIDAREMODELLO6_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "MA" -> {
                                        INATTESADIMAPPATURA_p++;
                                        INATTESADIMAPPATURA_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "IV" -> {
                                        INVERIFICA_p++;
                                        INVERIFICA_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "CK" -> {
                                        ESITOVERIFICACONCLUSO_p++;
                                        ESITOVERIFICACONCLUSO_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "EVI" -> {
                                        ESITOVERIFICAINVIATO_p++;
                                        ESITOVERIFICAINVIATO_a += db1.get_allievi_accreditati(idpr);
                                    }
                                    case "CO" -> {
                                        CONCLUSO_p++;
                                        CONCLUSO_a += db1.get_allievi_accreditati(idpr);
                                    }
                                }
                            }

                            XSSFRow row = getRow(sh1, indiceriga.get());
                            indiceriga.addAndGet(1);

                            AtomicInteger indicecolonna = new AtomicInteger(0);
                            setCell(getCell(row, indicecolonna.get()), rs0.getString("sa.ragionesociale").toUpperCase());
                            setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("sa.piva").toUpperCase());
                            setCell(getCell(row, indicecolonna.addAndGet(1)), regione_sedelegale);

                            //NUOVI 3 campi
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FINEATTIVITA_p + DAVALIDAREMODELLO6_p + INATTESADIMAPPATURA_p + INVERIFICA_p + ESITOVERIFICACONCLUSO_p + ESITOVERIFICAINVIATO_p + CONCLUSO_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FINEATTIVITA_a + DAVALIDAREMODELLO6_a + INATTESADIMAPPATURA_a
                                    + INVERIFICA_a + ESITOVERIFICACONCLUSO_a + ESITOVERIFICAINVIATO_a + CONCLUSO_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FASEA_p + FASEB_p + SOSPESO_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FASEA_a + FASEB_a + SOSPESO_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(DAVALIDARE_p + PROGRAMMATO_p + DACONFERMARE_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(DAVALIDARE_a + PROGRAMMATO_a + DACONFERMARE_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(docenti));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(DAVALIDARE_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(DAVALIDARE_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(PROGRAMMATO_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(PROGRAMMATO_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(DACONFERMARE_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(DACONFERMARE_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FASEA_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FASEA_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FASEB_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FASEB_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(SOSPESO_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(SOSPESO_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(RIGETTATO_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(RIGETTATO_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FINEATTIVITA_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(FINEATTIVITA_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(DAVALIDAREMODELLO6_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(DAVALIDAREMODELLO6_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(ESITOVERIFICACONCLUSO_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(ESITOVERIFICACONCLUSO_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(ESITOVERIFICAINVIATO_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(ESITOVERIFICAINVIATO_a));

                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(CONCLUSO_p));
                            setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(CONCLUSO_a));

                        }

                    }
                }

                //FOGLIO 2
                XSSFSheet sh2 = wb.getSheetAt(1);

                String sql0_foglio2 = "SELECT sa.ragionesociale,sa.piva,co.regione,pf.cip,st.descrizione,pf.start,"
                        + "pf.end,pf.idprogetti_formativi,sa.idsoggetti_attuatori,pf.extract "
                        + "FROM progetti_formativi pf, soggetti_attuatori sa, comuni co, stati_progetto st "
                        + "WHERE sa.idsoggetti_attuatori=pf.idsoggetti_attuatori AND co.idcomune=sa.comune "
                        + "AND st.idstati_progetto=pf.stato";

                AtomicInteger indiceriga2 = new AtomicInteger(1);
                try (Statement st0 = db1.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0_foglio2)) {
                    while (rs0.next()) {
                        String rendicontato = "NO";
                        if (rs0.getInt("pf.extract") != 0) {
                            rendicontato = switch (rs0.getInt("pf.extract")) {
                                case 1 ->
                                    "SI";
                                case 2 ->
                                    "IN ATTESA";
                                default ->
                                    "NO";
                            };
                        }

                        XSSFRow row = getRow(sh2, indiceriga2.get());
                        indiceriga2.addAndGet(1);

                        AtomicInteger indicecolonna = new AtomicInteger(0);
                        setCell(getCell(row, indicecolonna.get()), rs0.getString("sa.ragionesociale").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("sa.piva").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("co.regione").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("pf.cip") == null ? "" : rs0.getString("pf.cip").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("st.descrizione").toUpperCase());
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rendicontato);
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("pf.start") == null ? "" : sdfITA.format(rs0.getDate("pf.start")));
                        setCell(getCell(row, indicecolonna.addAndGet(1)), rs0.getString("pf.end") == null ? "" : sdfITA.format(rs0.getDate("pf.end")));

                        AtomicInteger ALLIEVIISCRITTI = new AtomicInteger(0);
                        AtomicInteger ALLIEVIVALIDATI = new AtomicInteger(0);
                        int idpr = rs0.getInt("pf.idprogetti_formativi");
                        String sql1_foglio2 = "SELECT a.idallievi,a.id_statopartecipazione FROM allievi a WHERE a.idprogetti_formativi=" + idpr;
                        try (ResultSet rs1 = db1.getConnection().createStatement().executeQuery(sql1_foglio2)) {
                            while (rs1.next()) {
                                String statoallievo = rs1.getString("a.id_statopartecipazione");
                                if (statoallievo.equals("15")) {
                                    ALLIEVIVALIDATI.addAndGet(1);
                                }
                                ALLIEVIISCRITTI.addAndGet(1);
                            }
                        }

                        setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(ALLIEVIISCRITTI.get()));
                        setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(ALLIEVIVALIDATI.get()));

                        //  DA CALCOLARE IN BASE ALLE ORE FASE_A > 48
                        setCell(getCell(row, indicecolonna.addAndGet(1)), String.valueOf(0));

                    }
                }

                wb.write(outputStream);
            }

            log.log(Level.WARNING, "{0} RILASCIATO CORRETTAMENTE.", fileout);

            String upd = "UPDATE estrazioni SET path = '" + fileout + "' WHERE idestrazione=3";
            db1.getConnection().createStatement().executeUpdate(upd);
            db1.closeDB();
        } catch (Exception e) {
            log.severe(estraiEccezione(e));
        }

    }

    public void impostaritiratounder48oreA() {
        try {

            Db_Gest db1 = new Db_Gest(this.host);

            String sql1 = "SELECT a.idallievi FROM allievi a WHERE a.idprogetti_formativi IS NOT NULL AND a.id_statopartecipazione IN ('10','12','13','14','15','18','19') "
                    + "AND a.idprogetti_formativi IN(SELECT p.idprogetti_formativi FROM progetti_formativi p "
                    + "WHERE p.stato IN (SELECT s.idstati_progetto FROM stati_progetto s "
                    + "WHERE s.ordine_processo IS NOT NULL AND s.ordine_processo>4)) AND a.importo < 48";

            try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                while (rs1.next()) {
                    long idallievi = rs1.getLong(1);
                    String upd1 = "UPDATE allievi SET id_statopartecipazione='11' WHERE idallievi=" + idallievi;
                    try (Statement st3 = db1.getConnection().createStatement()) {
                        st3.executeUpdate(upd1);
                    }
                }
            }

            db1.closeDB();

        } catch (Exception e) {
            log.severe(estraiEccezione(e));
        }
    }

    public void ore_ud() {
        Db_Gest db0 = new Db_Gest(this.host);

        List<Long> idallievi = new ArrayList<>();
        String sql = "SELECT a.idallievi FROM allievi a WHERE a.id_statopartecipazione IN (10,12,13,14,15,18,19)";
        try (Statement st = db0.getConnection().createStatement(); ResultSet rs = st.executeQuery(sql)) {
            while (rs.next()) {
                idallievi.add(rs.getLong(1));
            }
        } catch (Exception ex) {
            log.severe(estraiEccezione(ex));
        }

        //SALVA UD
        String sql0 = "SELECT u.descrizione,u.fase FROM unita_didattiche u GROUP BY u.descrizione,u.fase ORDER BY ordine;";
        try (Statement st0 = db0.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0)) {
            while (rs0.next()) {
                String UD_desc = rs0.getString(1);
                String UD_fase = rs0.getString(2);
                String sql1 = "SELECT u.codice,u.ore FROM unita_didattiche u WHERE u.descrizione = '" + UD_desc + "' AND u.fase = '" + UD_fase + "'";
                Map<String, String> MOD_ore = new HashMap<>();
                List<String> MOD_codici = new ArrayList<>();
                int ore = 0;
                try (Statement st1 = db0.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                    while (rs1.next()) {
                        MOD_ore.put(rs1.getString(1), rs1.getString(2));
                        MOD_codici.add(rs1.getString(1));
                        ore += rs1.getInt(2);
                    }
                }
//                if (MOD_codici.size() == 4 && MOD_codici.get(0).startsWith("A")) {
//                    String ud11_1 = "A-11A_A-11B";
//                    String ud11_2 = "A-11B_A-11C";
//                    String ud11_3 = "A-11C_A-11D";
//                    MOD_codici.add(ud11_1);
//                    MOD_codici.add(ud11_2);
//                    MOD_codici.add(ud11_3);
//                    MOD_ore.put(ud11_1, "5");
//                    MOD_ore.put(ud11_2, "5");
//                    MOD_ore.put(ud11_3, "5");
//                } else if (MOD_codici.size() == 4 && MOD_codici.get(0).startsWith("B")) {
//                    String ud11_1 = "B-1A_B-1B";
//                    String ud11_2 = "B-1B_B-1C";
//                    String ud11_3 = "B-1C_B-1D";
//                    MOD_codici.add(ud11_1);
//                    MOD_codici.add(ud11_2);
//                    MOD_codici.add(ud11_3);
//                    MOD_ore.put(ud11_1, "5");
//                    MOD_ore.put(ud11_2, "5");
//                    MOD_ore.put(ud11_3, "5");
//                }
                String nomemodulo = "";
                for (String m1 : MOD_codici) {
                    nomemodulo += m1 + "_";
                }
                nomemodulo = StringUtils.chop(nomemodulo);

                for (Long idallievo : idallievi) {

                    String sql2_F0 = "SELECT r.totaleorerendicontabili,r.data FROM registro_completo r WHERE r.idutente=" + idallievo
                            + " AND r.ruolo='ALLIEVO' AND r.nud='" + nomemodulo + "'";

                    String datainizio;
                    String datafine;
                    try (Statement st2F0 = db0.getConnection().createStatement(); ResultSet rs2F0 = st2F0.executeQuery(sql2_F0)) {
                        if (rs2F0.next()) {
                            //FAD 2 MODULI IN UNO
                            double res = rs2F0.getLong(1) / 3600000.00;
                            datainizio = sdfSQL.format(rs2F0.getDate(2));
                            datafine = sdfSQL.format(rs2F0.getDate(2));

                            db0.insert_UD_presenza(String.valueOf(Double.compare(Double.parseDouble(String.valueOf(ore)), res)),
                                    datafine, datainizio, nomemodulo.split("-")[0], String.valueOf(res), String.valueOf(ore), nomemodulo, String.valueOf(idallievo));
                            //System.out.println(idallievo+" : "+nomemodulo + " - " + ore + " : " + res + " - " + datainizio + " - " + datafine + " -- " + nomemodulo.split("-")[0]);
                        } else {

                            for (String m1 : MOD_codici) {
                                String sql2_F1 = "SELECT r.totaleorerendicontabili,r.data FROM registro_completo r WHERE r.idutente=" + idallievo
                                        + " AND r.ruolo='ALLIEVO' AND r.nud='" + m1 + "'";
//                            System.out.println(sql2_F1);
                                try (Statement st2F1 = db0.getConnection().createStatement(); ResultSet rs2F1 = st2F1.executeQuery(sql2_F1)) {
                                    if (rs2F1.next()) {
                                        //FAD 2 MODULI SINGOLI
                                        double res = rs2F1.getLong(1) / 3600000.00;
                                        datainizio = sdfSQL.format(rs2F1.getDate(2));
                                        datafine = sdfSQL.format(rs2F1.getDate(2));
                                        db0.insert_UD_presenza(String.valueOf(Double.compare(Double.parseDouble(String.valueOf(MOD_ore.get(m1))), res)),
                                                datafine, datainizio, m1.split("-")[0], String.valueOf(res), String.valueOf(MOD_ore.get(m1)), m1, String.valueOf(idallievo));
//                                        System.out.println(idallievo+" : "+idallievo+" : "+m1 + " - " + MOD_ore.get(m1) + " : " + res + " - " + datainizio + " - " + datafine + " -- " + m1.split("-")[0]);
                                    } else {
                                        String sql2_P = "SELECT r.durataconvalidata,p.datalezione FROM presenzelezioni p, presenzelezioniallievi r, lezioni_modelli l , lezione_calendario lc "
                                                + "WHERE p.idpresenzelezioni=r.idpresenzelezioni AND l.id_lezionimodelli=p.idlezioneriferimento AND l.id_lezionecalendario=lc.id_lezionecalendario "
                                                + "AND r.idallievi=" + idallievo + " AND r.convalidata=1 AND lc.codice_ud='" + m1 + "' AND r.durataconvalidata > 0";
                                        try (Statement st2P = db0.getConnection().createStatement(); ResultSet rs2P = st2P.executeQuery(sql2_P)) {
                                            if (rs2P.next()) {
                                                //PRESENZA MODULI SINGOLI
                                                double res = rs2P.getLong(1) / 3600000.00;
                                                datainizio = sdfSQL.format(rs2P.getDate(2));
                                                datafine = sdfSQL.format(rs2P.getDate(2));
                                                db0.insert_UD_presenza((Double.compare(Double.parseDouble(String.valueOf(MOD_ore.get(m1))), res) == 0) ? "1" : "0",
                                                        datafine, datainizio, m1.split("-")[0], String.valueOf(res), String.valueOf(MOD_ore.get(m1)), m1, String.valueOf(idallievo));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception ex) {
            log.severe(estraiEccezione(ex));
        }
        //UD
        List<UD> elencoud = new ArrayList<>();
        String sql3A = "SELECT descrizione,sum(ore),fase,GROUP_CONCAT(codice SEPARATOR ';') AS moduli FROM unita_didattiche GROUP BY descrizione,fase ORDER BY fase,ordine";
        try (Statement st3A = db0.getConnection().createStatement(); ResultSet rs3A = st3A.executeQuery(sql3A)) {
            while (rs3A.next()) {
                elencoud.add(new UD(rs3A.getString(1), rs3A.getDouble(2), rs3A.getString(3), rs3A.getString(4)));
            }
        } catch (Exception ex) {
            log.severe(estraiEccezione(ex));
        }
        for (Long idallievo : idallievi) {
////            ORE
            String sql3 = "SELECT SUM(p.orepresenze),p.fase FROM presenzeudallievi p WHERE p.idallievi=" + idallievo + " GROUP BY p.fase";
            try (Statement st3 = db0.getConnection().createStatement(); ResultSet rs3 = st3.executeQuery(sql3)) {
                while (rs3.next()) {
                    String upd3 = "UPDATE allievi SET orec_fase" + rs3.getString(2).toLowerCase() + " = '" + rs3.getDouble(1) + "' WHERE idallievi = " + idallievo;
                    try (Statement st3A = db0.getConnection().createStatement()) {
                        boolean es = st3A.executeUpdate(upd3) > 0;
                        log.log(Level.WARNING, "{0} ) {1} -- {2} : {3}", new Object[]{idallievo, rs3.getString(2), rs3.getDouble(1), es});
                    }
                }
            } catch (Exception ex1) {
                log.severe(estraiEccezione(ex1));
            }
////            UNITA DIDATTICHE
            int ud_ok_A = 0;
            int ud_ok_B = 0;

            for (UD unit : elencoud) {

                double oreallievo = 0.0;

                String sql4 = "SELECT * FROM presenzeudallievi p WHERE p.idallievi=" + idallievo + " AND (";
                List<String> moduli = Splitter.on(";").splitToList(unit.getModuli());

                for (String m1 : moduli) {
                    sql4 += " ud LIKE '%" + m1 + "%' OR";
                }

                sql4 = StringUtils.substring(sql4, 0, sql4.length() - 2).trim() + ")";
//                    System.out.println(sql4);
                try (Statement st4 = db0.getConnection().createStatement(); ResultSet rs4 = st4.executeQuery(sql4)) {
                    while (rs4.next()) {
                        oreallievo += rs4.getDouble("orepresenze");
                    }
                } catch (Exception ex2) {
                    log.severe(estraiEccezione(ex2));
                }
                if (unit.getOre() == oreallievo) {
                    if (unit.getFase().endsWith("A")) {
                        ud_ok_A++;
                    } else if (unit.getFase().endsWith("B")) {
                        ud_ok_B++;
                    }
                }
            }

            String upd4 = "UPDATE allievi SET ud_ok_A = ?, ud_ok_B = ? WHERE idallievi = ?";

            try (PreparedStatement ps4 = db0.getConnection().prepareStatement(upd4)) {
                ps4.setInt(1, ud_ok_A);
                ps4.setInt(2, ud_ok_B);
                ps4.setLong(3, idallievo);
                boolean es = ps4.executeUpdate() > 0;
                log.log(Level.WARNING, "{0} - {1}", new Object[]{ps4.toString(), es});
            } catch (Exception ex3) {
                log.severe(estraiEccezione(ex3));
            }

            //ASSENZE
            int assenzeOK = 0;
            String sql5 = "SELECT p.datarealelezione,p.durataconvalidata FROM presenzelezioniallievi p "
                    + "WHERE p.convalidata=1 AND (p.durataconvalidata=0 OR p.durataconvalidata=-1) "
                    + "AND p.idallievi=" + idallievo + " GROUP BY LEFT(p.datarealelezione,10)";

            try (Statement st5 = db0.getConnection().createStatement(); ResultSet rs5 = st5.executeQuery(sql5)) {
                while (rs5.next()) {
                    assenzeOK++;
                }
            } catch (Exception ex4) {
                log.severe(estraiEccezione(ex4));
            }

            if (assenzeOK > 0) {
                String upd5 = "UPDATE allievi SET assenzeOK = ?, assenzeKO = ? WHERE idallievi = ?";
                try (PreparedStatement ps5 = db0.getConnection().prepareStatement(upd5)) {
                    ps5.setInt(1, assenzeOK);
                    ps5.setInt(2, 0);
                    ps5.setLong(3, idallievo);
                    boolean es = ps5.executeUpdate() > 0;
                    log.log(Level.WARNING, "{0} - {1}", new Object[]{ps5.toString(), es});
                } catch (Exception ex5) {
                    log.severe(estraiEccezione(ex5));
                }
            }

        }

        db0.closeDB();
    }

    public void imposta_fineattivita() {
        Db_Gest db = new Db_Gest(this.host);
        String sql0 = "SELECT p.idprogetti_formativi FROM progetti_formativi p WHERE p.stato = 'ATB' AND p.`end` < CURDATE()";
        try (Statement st0 = db.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0)) {
            while (rs0.next()) {
                String idpr = rs0.getString(1);
                String sql1 = "SELECT a.idallievi,a.gruppo_faseB,a.idprogetti_formativi FROM allievi a "
                        + "WHERE a.id_statopartecipazione = '15' "
                        + "AND a.idprogetti_formativi = " + idpr;
                boolean fineattivita = true;
                try (Statement st1 = db.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                    while (rs1.next()) {
                        String idallievi = rs1.getString(1);
                        String gruppo_fb = rs1.getString(2);
                        String sql2 = "SELECT la.convalidata,p.datalezione FROM presenzelezioniallievi la, presenzelezioni p WHERE la.idallievi=" + idallievi + " AND p.idpresenzelezioni=la.idpresenzelezioni "
                                + "AND p.idlezioneriferimento IN (SELECT m.id_lezionimodelli FROM lezioni_modelli m WHERE m.id_modelli_progetto IN "
                                + "(SELECT mp.id_modello FROM modelli_progetti mp WHERE mp.id_progettoformativo = " + idpr + ") AND m.gruppo_faseB IN (0," + gruppo_fb + "))";
                        try (Statement st2 = db.getConnection().createStatement(); ResultSet rs2 = st2.executeQuery(sql2)) {
                            while (rs2.next()) {
                                boolean conv = rs2.getBoolean(1);
                                if (!conv) {
                                    log.log(Level.WARNING, "PROGETTO {0} - ALLIEVO {1} : NON ANCORA CONVALIDATA LEZIONE {2}", new Object[]{idpr, idallievi, rs2.getString(2)});
                                    fineattivita = false;
                                    break;
                                }
                            }
                        } catch (Exception ex2) {
                            log.severe(estraiEccezione(ex2));
                        }
                    }
                } catch (Exception ex2) {
                    log.severe(estraiEccezione(ex2));
                }
                if (fineattivita) {
                    try (Statement st3 = db.getConnection().createStatement()) {
                        String upd = "UPDATE progetti_formativi SET stato = 'F' WHERE idprogetti_formativi=" + idpr;
                        boolean es3 = st3.executeUpdate(upd) > 0;
                        if (es3) {
                            log.log(Level.INFO, "PROGETTO {0} IMPOSTATO IN FINE ATTIVITA'.", idpr);
                        }else{
                            log.log(Level.SEVERE, "ERRORE NELL''UPDATE DEL PROGETTO {0}", idpr);
                        }
                    }
                } else {
                    log.log(Level.INFO, "PROGETTO {0} ANCORA DA VALIDARE.", idpr);
                }
            }
        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }
        db.closeDB();
    }

//    public static void main(String[] args) {
//        Sicilia_gestione sg = new Sicilia_gestione(false);
//        sg.imposta_fineattivita();
//    }
}
