package rc.so.exe;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.stream.Collectors;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import static rc.so.exe.Constant.estraiEccezione;
import static rc.so.exe.Constant.sdfITA;
import static rc.so.exe.Constant.zipListFiles;
import static rc.so.exe.Sicilia_gestione.log;
import static rc.so.exe.Utils.calcolaintervallomillis;
import static rc.so.exe.Utils.getCell;
import static rc.so.exe.Utils.getRow;
import static rc.so.exe.Utils.parseIntR;
import static rc.so.exe.Utils.setCell;
import static rc.so.exe.Utils.timestamp;
import rc.so.report.FaseA;
import rc.so.report.Utenti;

public class Rendicontazione {

    private static final String separator = "###";
    private static final String formatdataCellint = "#,#";
    //COSTANTI
    public static final String A_VALUE = "2021.IT.05.SFPR.014/1/4.1/9.1.2/YSUP/001";//CIP Operazione
    public static final String C_VALUE = "YES I START UP-Formarsi per diventare imprenditore/imprenditrice in Sicilia";//Titolo del corso
    public static final int E_VALUE = 22;//Tipo attività corsuale
    public static final int G_VALUE = 0;//Numero allievi disabili previsti
    public static final String H_VALUE = "n";//Prove di selezione (s/n)
    public static final String I_VALUE = "n";//Esami finali (s/n)
    public static final int J_VALUE = 5;//Esami finali (s/n)
    public static final String K_VALUE = "Attestato di frequenza";//Specifica attestazione finale
    public static final String L_VALUE = "ALTRO";//Tipo qualifica
    public static final String M_VALUE = "autoimprenditorialità";//Specifica qualifica
    public static final int N_VALUE = 32;//Tipo corso
    public static final String O_VALUE = "periodi di sospensione momentanea dell'attività lavorativa";//Tipo collocazione temporale
    public static final String P_VALUE = "Disoccupato/inoccupato";//Descrizione dei destinatari
    public static final String S_VALUE = "";//Codice ATECO

    public static final int T_VALUE_A = 0;//Ore
    public static final int T_VALUE_B = 20;//Ore
    public static final int U_VALUE_A = 80;//Ore
    public static final int U_VALUE_B = 0;//Ore
    public static final int V_VALUE = 0;//Ore
    public static final String W_VALUE = "Percorsi di accompagnamento all’autoimpiego ed auto imprenditorialità";//Descrizione tipologia del corso
    public static final String X_VALUE = "aula";//Tipo registro

    private static final int soglia = 48;

    public static final int ALL_H_VALUE = 4;
    public static final int ALL_N_VALUE = 10;
    public static final String ALL_BFN_VALUE = "NO";

    private static final String CAL_IDCORSO = "";
    private static final String CAL_STAGE = "N";

    public static void generaRendicontazione(boolean complete) {
        try {
            FaseA FA = new FaseA(false);
            Db_Gest db1 = new Db_Gest(FA.getHost());

            if (complete) {
                String sql1 = "SELECT p.idprogetti_formativi FROM progetti_formativi p WHERE p.stato='CO'";
                List<Integer> idpr = new ArrayList<>();
                try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                    while (rs1.next()) {
                        idpr.add(rs1.getInt(1));
                    }
                }

                File xlsx1 = prospetto_riepilogo(0, idpr);
                File xlsx2 = prospetto_riepilogo_allievi(0, idpr);
                List<File> calendar = list_calendar(idpr);
                File xlsx3 = prospetto_riepilogo_calendario_allievi(0, idpr);

                List<File> output = new ArrayList<>();
                output.add(xlsx1);
                output.add(xlsx2);
                output.addAll(calendar);
                output.add(xlsx3);

                String pathdest = db1.getPath("output_excel_archive");
                DateTime oggi = new DateTime();
                String nomerend_cod = "R0";
                String nomerend = nomerend_cod + "_" + oggi.toString("ddMMyyyy");
                String filezip = pathdest + "/" + nomerend + ".zip";
                File zip = new File(filezip);

                zipListFiles(output, zip);
                try {
                    for (File f1 : output) {
                        f1.delete();
                        f1.deleteOnExit();
                    }
                } catch (Exception e) {
                }
                System.out.println("rc.so.exe.Rendicontazione.generaRendicontazione() " + zip.getPath());
            } else {
                Gson gson = new Gson();
                String sql0 = "SELECT e.idestrazione,e.progetti FROM estrazioni e WHERE e.path IS NULL";

                try (Statement st0 = db1.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0)) {
                    while (rs0.next()) {
                        int idestrazione = rs0.getInt(1);
                        List<Integer> idpr = new ArrayList<>();
                        List<String> progetti = gson.fromJson(rs0.getString(2), new TypeToken<List<String>>() {
                        }.getType());
                        progetti.forEach(cip -> {
                            try {
                                String sql1 = "SELECT e.idprogetti_formativi FROM progetti_formativi e WHERE e.cip = ?";
                                try (PreparedStatement ps1 = db1.getConnection().prepareStatement(sql1)) {
                                    ps1.setString(1, cip);
                                    try (ResultSet rs1 = ps1.executeQuery()) {
                                        if (rs1.next()) {
                                            idpr.add(rs1.getInt(1));
                                        }
                                    }
                                }
                            } catch (Exception ex) {
                                log.severe(estraiEccezione(ex));
                            }
                        });

                        File xlsx1 = prospetto_riepilogo(idestrazione, idpr);
                        File xlsx2 = prospetto_riepilogo_allievi(idestrazione, idpr);

                        if (xlsx1 != null && xlsx2 != null) {

                            List<File> calendar = list_calendar(idpr);

                            File xlsx3 = prospetto_riepilogo_calendario_allievi(idestrazione, idpr);

                            List<File> output = new ArrayList<>();
                            output.add(xlsx1);
                            output.add(xlsx2);
                            output.add(xlsx3);
                            output.addAll(calendar);
                            String pathdest = db1.getPath("output_excel_archive");
                            DateTime oggi = new DateTime();
                            String nomerend_cod = "R" + idestrazione;
                            String nomerend = nomerend_cod + "_" + oggi.toString("ddMMyyyy");
                            String filezip = pathdest + "/" + nomerend + ".zip";

                            File zip = new File(filezip);

                            if (zipListFiles(output, zip)) {
                                String update1 = "UPDATE estrazioni SET path = ? WHERE idestrazione = ?";
                                try (PreparedStatement ps1 = db1.getConnection().prepareStatement(update1)) {
                                    ps1.setString(1, StringUtils.replace(zip.getPath(), "\\", "/"));
                                    ps1.setInt(2, idestrazione);
                                    ps1.executeUpdate();
                                }
                                for (int i = 0; i < idpr.size(); i++) {
                                    String update2 = "UPDATE progetti_formativi SET extract = 1 WHERE idprogetti_formativi = ?";
                                    try (PreparedStatement ps2 = db1.getConnection().prepareStatement(update2)) {
                                        ps2.setInt(1, idpr.get(i));
                                        ps2.executeUpdate();
                                    }
                                }
                                try {
                                    for (File f1 : output) {
                                        f1.delete();
                                        f1.deleteOnExit();
                                    }
                                } catch (Exception e) {
                                }
                            }

                        }

                    }
                }
            }

        } catch (Exception e) {
            log.severe(estraiEccezione(e));
        }
    }

    private static int mapping_i_cittadinanza_id(String ing) {
        return switch (ing) {
            case "99" ->
                90;
            default ->
                0;
        };
    }

    private static int mapping_i_tipo_titolo_studio_id(String ing) {
        return switch (ing) {
            case "00" ->
                21;
            case "01" ->
                22;
            case "02" ->
                23;
            case "03" ->
                24;
            case "04" ->
                25;
            case "05" ->
                26;
            case "06" ->
                27;
            case "07" ->
                28;
            case "08" ->
                29;
            case "09" ->
                30;
            default ->
                0;
        };
    }

    private static List<File> list_calendar(List<Integer> list_idpr) {
        List<File> output = new ArrayList<>();
        DateTime oggi = new DateTime();
        try {
            FaseA FA = new FaseA(false);
            Db_Gest db1 = new Db_Gest(FA.getHost());
            String pathdest = db1.getPath("output_excel_archive");
            String fileing = pathdest + "YISUS_Prospetto_Calendario_v1.xlsx";
            File ing = new File(fileing);

            for (int ss = 0; ss < list_idpr.size(); ss++) {

                int idpr = list_idpr.get(ss);

                String sql1 = "SELECT m.id_modello,m.modello,p.cip FROM modelli_progetti m, progetti_formativi p WHERE m.id_progettoformativo=p.idprogetti_formativi AND m.id_progettoformativo = ? AND m.modello IN (3,4) AND m.stato='OK' ORDER BY m.modello";
                try (PreparedStatement ps1 = db1.getConnection().prepareStatement(sql1)) {
                    ps1.setInt(1, idpr);
                    try (ResultSet rs1 = ps1.executeQuery()) {
                        while (rs1.next()) {
                            int modello = rs1.getInt(2);
                            int id_modello = rs1.getInt(1);
                            String CIP = rs1.getString(3);
                            if (modello == 3) { // FASE A

                                String sql2 = "SELECT l.giorno,l.orario_start,l.orario_end,d.codicefiscale,c.codice_ud "
                                        + "FROM lezioni_modelli l, docenti d, lezione_calendario c "
                                        + "WHERE l.id_docente=d.iddocenti AND l.id_lezionecalendario=c.id_lezionecalendario "
                                        + "AND l.id_modelli_progetto = ? ORDER BY l.id_lezionecalendario,l.giorno,l.orario_start";

                                AtomicInteger index_row = new AtomicInteger(1);
                                try (PreparedStatement ps2 = db1.getConnection().prepareStatement(sql2)) {
                                    ps2.setInt(1, id_modello);
                                    try (InputStream is = new FileInputStream(ing); XSSFWorkbook wb = new XSSFWorkbook(is, false)) {
                                        XSSFSheet sh_corso = wb.getSheet("calendario");
                                        XSSFFont font_string = wb.createFont();
                                        font_string.setFontHeightInPoints((short) 12);

                                        XSSFCellStyle style_normal = wb.createCellStyle();
                                        style_normal.setVerticalAlignment(VerticalAlignment.CENTER);
                                        style_normal.setAlignment(HorizontalAlignment.CENTER);
                                        style_normal.setBorderBottom(BorderStyle.THIN);
                                        style_normal.setBorderTop(BorderStyle.THIN);
                                        style_normal.setBorderRight(BorderStyle.THIN);
                                        style_normal.setBorderLeft(BorderStyle.THIN);
                                        style_normal.setFont(font_string);

                                        XSSFDataFormat xssfDataFormat = wb.createDataFormat();
                                        XSSFCellStyle cellStyle_int = wb.createCellStyle();
                                        cellStyle_int.setBorderBottom(BorderStyle.THIN);
                                        cellStyle_int.setBorderTop(BorderStyle.THIN);
                                        cellStyle_int.setBorderRight(BorderStyle.THIN);
                                        cellStyle_int.setBorderLeft(BorderStyle.THIN);
                                        cellStyle_int.setVerticalAlignment(VerticalAlignment.CENTER);
                                        cellStyle_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));

                                        XSSFCellStyle style_int = wb.createCellStyle();
                                        style_int.setVerticalAlignment(VerticalAlignment.CENTER);
                                        style_int.setAlignment(HorizontalAlignment.CENTER);
                                        style_int.setBorderBottom(BorderStyle.THIN);
                                        style_int.setBorderTop(BorderStyle.THIN);
                                        style_int.setBorderRight(BorderStyle.THIN);
                                        style_int.setBorderLeft(BorderStyle.THIN);
                                        style_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));
                                        style_int.setFont(font_string);

                                        try (ResultSet rs2 = ps2.executeQuery()) {
                                            while (rs2.next()) {

                                                String giorno = rs2.getString(1);
                                                String orario_start = rs2.getString(2);
                                                String orario_end = rs2.getString(3);
                                                String codicefiscale = rs2.getString(4);
                                                String codice_ud = rs2.getString(5);

                                                String ore = String.valueOf(calcolaintervallomillis(orario_start, orario_end) / 3600000.00);

//                                                System.out.println(giorno.split("-")[2] + separator + giorno.split("-")[1] + separator + giorno.split("-")[0] + separator
//                                                        + orario_start.split(":")[0] + separator + orario_start.split(":")[1] + separator + ore + separator + CAL_IDCORSO + separator
//                                                        + CAL_STAGE + separator + codicefiscale + separator + codice_ud);
                                                XSSFRow riga_A = getRow(sh_corso, index_row.get());
                                                index_row.addAndGet(1);

                                                AtomicInteger index_column = new AtomicInteger(0);

                                                setCell(getCell(riga_A, index_column.get()), style_int, giorno.split("-")[2], true, false);
                                                setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, giorno.split("-")[1], true, false);
                                                setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, giorno.split("-")[0], true, false);
                                                setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, orario_start.split(":")[0], true, false);
                                                setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, orario_start.split(":")[1], true, false);
                                                setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, ore, false, true);
                                                setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, CAL_IDCORSO, false, false);
                                                setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, CAL_STAGE, false, false);
                                                setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, codicefiscale, false, false);
                                                setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, codice_ud, false, false);

                                            }
                                        }

                                        for (int ix = 0; ix < 10; ix++) {
                                            sh_corso.autoSizeColumn(ix);
                                        }
                                        String CIP_VALUE = CIP + "_FASE A";
                                        File output_xlsx = new File(pathdest + "/" + CIP_VALUE + "_CALENDARIO_" + oggi.toString(timestamp) + ".xlsx");
                                        try (FileOutputStream outputStream = new FileOutputStream(output_xlsx)) {
                                            wb.write(outputStream);
                                            log.log(Level.INFO, "FILE RILASCIATO: {0}", output_xlsx.getPath());
                                            output.add(output_xlsx);
                                        }
                                    }
                                }

                            } else { //FASE B . GRUPPI

                                String sql2 = "SELECT DISTINCT(l.gruppo_faseB) FROM lezioni_modelli l WHERE l.id_modelli_progetto = ? ORDER BY l.gruppo_faseB";

                                try (PreparedStatement ps2 = db1.getConnection().prepareStatement(sql2)) {
                                    ps2.setInt(1, id_modello);
                                    try (ResultSet rs2 = ps2.executeQuery()) {
                                        while (rs2.next()) {
                                            int gruppo_faseb = rs2.getInt(1);

                                            String sql3 = "SELECT l.giorno,l.orario_start,l.orario_end,d.codicefiscale,c.codice_ud "
                                                    + "FROM lezioni_modelli l, docenti d, lezione_calendario c "
                                                    + "WHERE l.id_docente=d.iddocenti AND l.id_lezionecalendario=c.id_lezionecalendario "
                                                    + "AND l.id_modelli_progetto = ? AND l.gruppo_faseB = ? ORDER BY l.id_lezionecalendario,l.giorno,l.orario_start";

                                            try (InputStream is = new FileInputStream(ing); XSSFWorkbook wb = new XSSFWorkbook(is, false)) {

                                                XSSFSheet sh_corso = wb.getSheet("calendario");
                                                XSSFFont font_string = wb.createFont();
                                                font_string.setFontHeightInPoints((short) 12);

                                                XSSFCellStyle style_normal = wb.createCellStyle();
                                                style_normal.setVerticalAlignment(VerticalAlignment.CENTER);
                                                style_normal.setAlignment(HorizontalAlignment.CENTER);
                                                style_normal.setBorderBottom(BorderStyle.THIN);
                                                style_normal.setBorderTop(BorderStyle.THIN);
                                                style_normal.setBorderRight(BorderStyle.THIN);
                                                style_normal.setBorderLeft(BorderStyle.THIN);
                                                style_normal.setFont(font_string);

                                                XSSFDataFormat xssfDataFormat = wb.createDataFormat();
                                                XSSFCellStyle cellStyle_int = wb.createCellStyle();
                                                cellStyle_int.setBorderBottom(BorderStyle.THIN);
                                                cellStyle_int.setBorderTop(BorderStyle.THIN);
                                                cellStyle_int.setBorderRight(BorderStyle.THIN);
                                                cellStyle_int.setBorderLeft(BorderStyle.THIN);
                                                cellStyle_int.setVerticalAlignment(VerticalAlignment.CENTER);
                                                cellStyle_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));

                                                XSSFCellStyle style_int = wb.createCellStyle();
                                                style_int.setVerticalAlignment(VerticalAlignment.CENTER);
                                                style_int.setAlignment(HorizontalAlignment.CENTER);
                                                style_int.setBorderBottom(BorderStyle.THIN);
                                                style_int.setBorderTop(BorderStyle.THIN);
                                                style_int.setBorderRight(BorderStyle.THIN);
                                                style_int.setBorderLeft(BorderStyle.THIN);
                                                style_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));
                                                style_int.setFont(font_string);
                                                AtomicInteger index_row = new AtomicInteger(1);

                                                try (PreparedStatement ps3 = db1.getConnection().prepareStatement(sql3)) {
                                                    ps3.setInt(1, id_modello);
                                                    ps3.setInt(2, gruppo_faseb);
                                                    try (ResultSet rs3 = ps3.executeQuery()) {
                                                        while (rs3.next()) {
                                                            String giorno = rs3.getString(1);
                                                            String orario_start = rs3.getString(2);
                                                            String orario_end = rs3.getString(3);
                                                            String codicefiscale = rs3.getString(4);
                                                            String codice_ud = rs3.getString(5);
                                                            String ore = String.valueOf(calcolaintervallomillis(orario_start, orario_end) / 3600000.00);

                                                            XSSFRow riga_A = getRow(sh_corso, index_row.get());
                                                            index_row.addAndGet(1);

                                                            AtomicInteger index_column = new AtomicInteger(0);

                                                            setCell(getCell(riga_A, index_column.get()), style_int, giorno.split("-")[2], true, false);
                                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, giorno.split("-")[1], true, false);
                                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, giorno.split("-")[0], true, false);
                                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, orario_start.split(":")[0], true, false);
                                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, orario_start.split(":")[1], true, false);
                                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, ore, false, true);
                                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, CAL_IDCORSO, false, false);
                                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, CAL_STAGE, false, false);
                                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, codicefiscale, false, false);
                                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, codice_ud, false, false);
                                                        }
                                                    }
                                                }
                                                for (int ix = 0; ix < 10; ix++) {
                                                    sh_corso.autoSizeColumn(ix);
                                                }
                                                String CIP_VALUE = CIP + "_FASE B" + gruppo_faseb;
                                                File output_xlsx = new File(pathdest + "/" + CIP_VALUE + "_CALENDARIO_" + oggi.toString(timestamp) + ".xlsx");
                                                try (FileOutputStream outputStream = new FileOutputStream(output_xlsx)) {
                                                    wb.write(outputStream);
                                                    log.log(Level.INFO, "FILE RILASCIATO: {0}", output_xlsx.getPath());
                                                    output.add(output_xlsx);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }
        return output;
    }

    private static File prospetto_riepilogo_allievi(int idestrazione, List<Integer> list_idpr) {
        File output_xlsx = null;
        DateTime oggi = new DateTime();
        String nomerend_cod = "R" + idestrazione;
        String nomerend = nomerend_cod + "_" + oggi.toString("ddMMyyyy");
        try {
            FaseA FA = new FaseA(false);
            Db_Gest db1 = new Db_Gest(FA.getHost());
            String pathdest = db1.getPath("output_excel_archive");
            String fileing = pathdest + "YISUS_Prospetto_Allievi_v1.xlsx";
            File ing = new File(fileing);

            try (InputStream is = new FileInputStream(ing); XSSFWorkbook wb = new XSSFWorkbook(is, false)) {

                XSSFSheet sh_corso = wb.getSheet("Tracciato partecipante");
                XSSFFont font_string = wb.createFont();
                font_string.setFontHeightInPoints((short) 12);

                XSSFCellStyle style_normal = wb.createCellStyle();
                style_normal.setVerticalAlignment(VerticalAlignment.CENTER);
                style_normal.setAlignment(HorizontalAlignment.CENTER);
                style_normal.setBorderBottom(BorderStyle.THIN);
                style_normal.setBorderTop(BorderStyle.THIN);
                style_normal.setBorderRight(BorderStyle.THIN);
                style_normal.setBorderLeft(BorderStyle.THIN);
                style_normal.setFont(font_string);

                XSSFDataFormat xssfDataFormat = wb.createDataFormat();
                XSSFCellStyle cellStyle_int = wb.createCellStyle();
                cellStyle_int.setBorderBottom(BorderStyle.THIN);
                cellStyle_int.setBorderTop(BorderStyle.THIN);
                cellStyle_int.setBorderRight(BorderStyle.THIN);
                cellStyle_int.setBorderLeft(BorderStyle.THIN);
                cellStyle_int.setVerticalAlignment(VerticalAlignment.CENTER);
                cellStyle_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));

                XSSFCellStyle style_int = wb.createCellStyle();
                style_int.setVerticalAlignment(VerticalAlignment.CENTER);
                style_int.setAlignment(HorizontalAlignment.CENTER);
                style_int.setBorderBottom(BorderStyle.THIN);
                style_int.setBorderTop(BorderStyle.THIN);
                style_int.setBorderRight(BorderStyle.THIN);
                style_int.setBorderLeft(BorderStyle.THIN);
                style_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));
                style_int.setFont(font_string);

                AtomicInteger index_row = new AtomicInteger(2);

                for (int ss = 0; ss < list_idpr.size(); ss++) {

                    int idpr = list_idpr.get(ss);

                    String sql1 = "SELECT p.cip FROM progetti_formativi p WHERE p.idprogetti_formativi = " + idpr;

                    try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                        if (rs1.next()) {
                            String cip = rs1.getString(1).toUpperCase();

                            String sql2 = "SELECT * FROM allievi a WHERE a.idprogetti_formativi = " + idpr + " AND a.orec_fasea >= " + soglia + " ORDER BY codicefiscale";

                            try (Statement st2 = db1.getConnection().createStatement(); ResultSet rs2 = st2.executeQuery(sql2)) {
                                while (rs2.next()) {
                                    String CIP_VALUE = cip + "_FASE A";
                                    int gb = rs2.getInt("a.gruppo_faseB");
                                    String ALL_C_VALUE = rs2.getString("a.codicefiscale").toUpperCase().trim();
                                    String ALL_D_VALUE = rs2.getString("a.nome").toUpperCase().trim();
                                    String ALL_E_VALUE = rs2.getString("a.cognome").toUpperCase().trim();
                                    String ALL_F_VALUE = sdfITA.format(rs2.getDate("a.datanascita"));
                                    int ALL_G_VALUE = mapping_i_tipo_titolo_studio_id(rs2.getString("a.titolo_studio").trim());
                                    String ALL_AE_VALUE = rs2.getString("a.email").toLowerCase().trim();
                                    String ALL_AFG_VALUE = rs2.getString("a.telefono").toLowerCase().trim();
                                    String ALL_AH_VALUE = rs2.getString("a.sesso").toUpperCase().trim().equals("M") ? "UOMO" : "DONNA";
                                    int ALL_AI_VALUE = mapping_i_cittadinanza_id(rs2.getString("a.cittadinanza").trim());
                                    String ALL_AK_VALUE = rs2.getString("a.indirizzoresidenza").toUpperCase().trim();

                                    String sql3 = "SELECT t.CODICE_PROVINCIA,t.CODICE_COMUNE,c.provincia FROM comuni c, TC16 t WHERE c.nome=t.DESCRIZIONE_COMUNE "
                                            + "AND c.regione=t.DESCRIZIONE_REGIONE AND c.cittadinanza=0 AND c.idcomune = " + rs2.getInt("a.comune_residenza");

                                    int ALL_AL_VALUE = 0;
                                    String ALL_AM_VALUE = "";

                                    try (Statement st3 = db1.getConnection().createStatement(); ResultSet rs3 = st3.executeQuery(sql3)) {
                                        if (rs3.next()) {
                                            ALL_AL_VALUE = parseIntR(rs3.getString(1) + rs3.getString(2));
                                            ALL_AM_VALUE = rs3.getString(3).toUpperCase();
                                        }
                                    }
                                    String ALL_AN_VALUE = rs2.getString("a.capresidenza").toUpperCase().trim();

                                    int ALL_AU_VALUE = 0;
                                    String ALL_AV_VALUE = "";
                                    String ALL_AW_VALUE = "";

                                    if (rs2.getString("a.stato_nascita").trim().equals("99") || rs2.getString("a.stato_nascita").trim().equals("100")) { //ITALIA
                                        String sql4 = "SELECT t.CODICE_PROVINCIA,t.CODICE_COMUNE,c.provincia FROM comuni c, TC16 t WHERE c.nome=t.DESCRIZIONE_COMUNE "
                                                + "AND c.regione=t.DESCRIZIONE_REGIONE AND c.cittadinanza=0 AND c.idcomune = " + rs2.getInt("a.comune_nascita");
                                        try (Statement st4 = db1.getConnection().createStatement(); ResultSet rs4 = st4.executeQuery(sql4)) {
                                            if (rs4.next()) {
                                                ALL_AU_VALUE = parseIntR(rs4.getString(1) + rs4.getString(2));
                                                ALL_AW_VALUE = rs4.getString(3).toUpperCase();
                                            }
                                        }
                                    } else {
                                        String sql4 = "SELECT nome FROM nazioni_rc WHERE codicefiscale='" + rs2.getString("a.stato_nascita").trim() + "'";
                                        try (Statement st4 = db1.getConnection().createStatement(); ResultSet rs4 = st4.executeQuery(sql4)) {
                                            if (rs4.next()) {
                                                ALL_AV_VALUE = rs4.getString(1).toUpperCase();
                                            }
                                        }
                                    }

                                    XSSFRow riga_A = getRow(sh_corso, index_row.get());
                                    index_row.addAndGet(1);

                                    AtomicInteger index_column = new AtomicInteger(0);

                                    setCell(getCell(riga_A, index_column.get()), style_normal, A_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, CIP_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_C_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_D_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_E_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_F_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(ALL_G_VALUE), true, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(ALL_H_VALUE), true, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(6)), style_int, String.valueOf(ALL_N_VALUE), true, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(17)), style_normal, ALL_AE_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_AFG_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_AFG_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_AH_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(ALL_AI_VALUE), true, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(2)), style_normal, ALL_AK_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(ALL_AL_VALUE), true, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_AM_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_AN_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(7)), style_int, String.valueOf(ALL_AU_VALUE), true, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_AV_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_AW_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(9)), style_normal, ALL_BFN_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);

                                    if (gb > 0) {

                                        XSSFRow riga_B = getRow(sh_corso, index_row.get());
                                        index_row.addAndGet(1);

                                        AtomicInteger index_column_B = new AtomicInteger(0);
                                        CIP_VALUE = cip + "_FASE B" + gb;
                                        setCell(getCell(riga_B, index_column_B.get()), style_normal, A_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, CIP_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_C_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_D_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_E_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_F_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(ALL_G_VALUE), true, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(ALL_H_VALUE), true, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(6)), style_int, String.valueOf(ALL_N_VALUE), true, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(17)), style_normal, ALL_AE_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_AFG_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_AFG_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_AH_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(ALL_AI_VALUE), true, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(2)), style_normal, ALL_AK_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(ALL_AL_VALUE), true, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_AM_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_AN_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(7)), style_int, String.valueOf(ALL_AU_VALUE), true, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_AV_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_AW_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(9)), style_normal, ALL_BFN_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                        setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, ALL_BFN_VALUE, false, false);
                                    }
                                }
                            }

                        }
                    }
                }
                for (int ix = 0; ix < 50; ix++) {
                    sh_corso.autoSizeColumn(ix);
                }
                output_xlsx = new File(pathdest + "/" + nomerend + "_RiepilogoAllievi_" + new DateTime().toString(timestamp) + ".xlsx");
                try (FileOutputStream outputStream = new FileOutputStream(output_xlsx)) {
                    wb.write(outputStream);
                    log.log(Level.INFO, "FILE RILASCIATO: {0}", output_xlsx.getPath());
                }
            }

        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }
        return output_xlsx;
    }

    private static File prospetto_riepilogo(int idestrazione, List<Integer> list_idpr) {
        File output_xlsx = null;
        DateTime oggi = new DateTime();
        String nomerend_cod = "R" + idestrazione;
        String nomerend = nomerend_cod + "_" + oggi.toString("ddMMyyyy");
        try {
            FaseA FA = new FaseA(false);
            Db_Gest db1 = new Db_Gest(FA.getHost());
            String pathdest = db1.getPath("output_excel_archive");
            String fileing = pathdest + "YISUS_Prospetto_Riepilogo_v1.xlsx";
            File ing = new File(fileing);

            try (InputStream is = new FileInputStream(ing); XSSFWorkbook wb = new XSSFWorkbook(is, false)) {

                XSSFSheet sh_corso = wb.getSheet("CORSO");
                XSSFFont font_string = wb.createFont();
                font_string.setFontHeightInPoints((short) 12);

                XSSFCellStyle style_normal = wb.createCellStyle();
                style_normal.setVerticalAlignment(VerticalAlignment.CENTER);
                style_normal.setAlignment(HorizontalAlignment.CENTER);
                style_normal.setBorderBottom(BorderStyle.THIN);
                style_normal.setBorderTop(BorderStyle.THIN);
                style_normal.setBorderRight(BorderStyle.THIN);
                style_normal.setBorderLeft(BorderStyle.THIN);
                style_normal.setFont(font_string);

                XSSFDataFormat xssfDataFormat = wb.createDataFormat();
                XSSFCellStyle cellStyle_int = wb.createCellStyle();
                cellStyle_int.setBorderBottom(BorderStyle.THIN);
                cellStyle_int.setBorderTop(BorderStyle.THIN);
                cellStyle_int.setBorderRight(BorderStyle.THIN);
                cellStyle_int.setBorderLeft(BorderStyle.THIN);
                cellStyle_int.setVerticalAlignment(VerticalAlignment.CENTER);
                cellStyle_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));

                XSSFCellStyle style_int = wb.createCellStyle();
                style_int.setVerticalAlignment(VerticalAlignment.CENTER);
                style_int.setAlignment(HorizontalAlignment.CENTER);
                style_int.setBorderBottom(BorderStyle.THIN);
                style_int.setBorderTop(BorderStyle.THIN);
                style_int.setBorderRight(BorderStyle.THIN);
                style_int.setBorderLeft(BorderStyle.THIN);
                style_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));
                style_int.setFont(font_string);

                AtomicInteger index_row = new AtomicInteger(1);
//                AtomicInteger indice = new AtomicInteger(0);

                for (int ss = 0; ss < list_idpr.size(); ss++) {

                    int idpr = list_idpr.get(ss);

                    String sql1 = "SELECT p.idprogetti_formativi,p.cip,s.ragionesociale,c.cod_provincia,p.start,p.end "
                            + "FROM progetti_formativi p, soggetti_attuatori s,comuni c WHERE p.stato='CO' "
                            + "AND p.idsoggetti_attuatori=s.idsoggetti_attuatori AND c.idcomune=s.comune AND p.idprogetti_formativi = " + idpr;

                    try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                        if (rs1.next()) {
                            String cip = rs1.getString(2).toUpperCase();
                            String cod_provincia = rs1.getString(4).toUpperCase();

                            List<Utenti> allievi_OK = db1.list_Allievi_OK(idpr);

                            String sql2 = "SELECT DISTINCT(l.gruppo_faseB) FROM lezioni_modelli l WHERE l.id_modelli_progetto IN (SELECT m.id_modello FROM modelli_progetti m WHERE m.id_progettoformativo = "
                                    + idpr + " AND m.modello=4)";
                            List<Integer> gruppiB = new ArrayList<>();
                            try (Statement st2 = db1.getConnection().createStatement(); ResultSet rs2 = st2.executeQuery(sql2)) {
                                while (rs2.next()) {
                                    gruppiB.add(rs2.getInt(1));
                                }
                            }

                            String sql3 = "SELECT l.giorno FROM lezioni_modelli l WHERE l.id_lezionecalendario IN (1,16) AND l.id_modelli_progetto IN "
                                    + "(SELECT m.id_modello FROM modelli_progetti m WHERE m.id_progettoformativo = " + idpr + " AND m.modello=3) ORDER BY l.giorno";
                            String start = "";
                            String end = "";
                            try (Statement st3 = db1.getConnection().createStatement(); ResultSet rs3 = st3.executeQuery(sql3)) {
                                while (rs3.next()) {
                                    if (start.equals("")) {
                                        start = sdfITA.format(rs3.getDate(1));
                                    }
                                    end = sdfITA.format(rs3.getDate(1));
                                }
                            }
                            String CIP_VALUE = cip + "_FASE A";

                            XSSFRow riga_A = getRow(sh_corso, index_row.get());
                            index_row.addAndGet(1);

                            AtomicInteger index_column = new AtomicInteger(0);

                            setCell(getCell(riga_A, index_column.get()), style_normal, A_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, CIP_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, C_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, cod_provincia, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(E_VALUE), true, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, String.valueOf(allievi_OK.size()), true, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(G_VALUE), true, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, H_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, I_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(J_VALUE), true, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, K_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, L_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, M_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(N_VALUE), true, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, O_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, P_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, start, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, end, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, S_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(T_VALUE_A), true, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(U_VALUE_A), true, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, String.valueOf(V_VALUE), true, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, W_VALUE, false, false);
                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, X_VALUE, false, false);

                            for (Integer gb : gruppiB) {
                                CIP_VALUE = cip + "_FASE B" + gb;
                                int F_VALUE = allievi_OK.stream().filter(a -> a.getGruppofaseB().equals(String.valueOf(gb))).collect(Collectors.toList()).size();

                                start = "";
                                end = "";
                                String sql3A = "SELECT l.giorno FROM lezioni_modelli l WHERE l.gruppo_faseB = " + gb + " AND l.id_lezionecalendario IN (17,20) AND l.id_modelli_progetto IN "
                                        + "(SELECT m.id_modello FROM modelli_progetti m WHERE m.id_progettoformativo = " + idpr + " AND m.modello=4) ORDER BY l.giorno";
                                try (Statement st3A = db1.getConnection().createStatement(); ResultSet rs3A = st3A.executeQuery(sql3A)) {
                                    while (rs3A.next()) {
                                        if (start.equals("")) {
                                            start = sdfITA.format(rs3A.getDate(1));
                                        }
                                        end = sdfITA.format(rs3A.getDate(1));
                                    }
                                }

                                XSSFRow riga_B = getRow(sh_corso, index_row.get());
                                index_row.addAndGet(1);

                                AtomicInteger index_column_B = new AtomicInteger(0);

                                setCell(getCell(riga_B, index_column_B.get()), style_normal, A_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, CIP_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, C_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, cod_provincia, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(E_VALUE), true, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, String.valueOf(F_VALUE), true, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(G_VALUE), true, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, H_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, I_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(J_VALUE), true, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, K_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, L_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, M_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(N_VALUE), true, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, O_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, P_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, start, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, end, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, S_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(T_VALUE_B), true, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(U_VALUE_B), true, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_int, String.valueOf(V_VALUE), true, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, W_VALUE, false, false);
                                setCell(getCell(riga_B, index_column_B.addAndGet(1)), style_normal, X_VALUE, false, false);

                            }

                        }
                    }
                }
                for (int ix = 0; ix < 24; ix++) {
                    sh_corso.autoSizeColumn(ix);
                }
                output_xlsx = new File(pathdest + "/" + nomerend + "_Riepilogo_" + new DateTime().toString(timestamp) + ".xlsx");
                try (FileOutputStream outputStream = new FileOutputStream(output_xlsx)) {
                    wb.write(outputStream);
                    log.log(Level.INFO, "FILE RILASCIATO: {0}", output_xlsx.getPath());
                }
            }

        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }
        return output_xlsx;
    }

    private static File prospetto_riepilogo_calendario_allievi(int idestrazione, List<Integer> list_idpr) {
        DateTime oggi = new DateTime();
        String nomerend_cod = "R" + idestrazione;
        String nomerend = nomerend_cod + "_" + oggi.toString("ddMMyyyy");
        try {
            FaseA FA = new FaseA(false);
            Db_Gest db1 = new Db_Gest(FA.getHost());
            String pathdest = db1.getPath("output_excel_archive");
            String fileing = pathdest + "YISUS_Prospetto_Allievi_Calendario_v1.xlsx";
            File ing = new File(fileing);
            try (InputStream is = new FileInputStream(ing); XSSFWorkbook wb = new XSSFWorkbook(is, false)) {

                XSSFSheet sh_registro = wb.getSheet("REGISTRO NEW");
                XSSFFont font_string = wb.createFont();
                font_string.setFontHeightInPoints((short) 12);

                XSSFCellStyle style_normal = wb.createCellStyle();
                style_normal.setVerticalAlignment(VerticalAlignment.CENTER);
                style_normal.setAlignment(HorizontalAlignment.CENTER);
                style_normal.setBorderBottom(BorderStyle.THIN);
                style_normal.setBorderTop(BorderStyle.THIN);
                style_normal.setBorderRight(BorderStyle.THIN);
                style_normal.setBorderLeft(BorderStyle.THIN);
                style_normal.setFont(font_string);

                XSSFDataFormat xssfDataFormat = wb.createDataFormat();
                XSSFCellStyle cellStyle_int = wb.createCellStyle();
                cellStyle_int.setBorderBottom(BorderStyle.THIN);
                cellStyle_int.setBorderTop(BorderStyle.THIN);
                cellStyle_int.setBorderRight(BorderStyle.THIN);
                cellStyle_int.setBorderLeft(BorderStyle.THIN);
                cellStyle_int.setVerticalAlignment(VerticalAlignment.CENTER);
                cellStyle_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));

                XSSFCellStyle style_int = wb.createCellStyle();
                style_int.setVerticalAlignment(VerticalAlignment.CENTER);
                style_int.setAlignment(HorizontalAlignment.CENTER);
                style_int.setBorderBottom(BorderStyle.THIN);
                style_int.setBorderTop(BorderStyle.THIN);
                style_int.setBorderRight(BorderStyle.THIN);
                style_int.setBorderLeft(BorderStyle.THIN);
                style_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));
                style_int.setFont(font_string);

                AtomicInteger index_row = new AtomicInteger(1);

                for (int ss = 0; ss < list_idpr.size(); ss++) {

                    int idpr = list_idpr.get(ss);

                    String sql1 = "SELECT m.idallievi,m.codicefiscale,p.cip FROM allievi m, progetti_formativi p WHERE p.idprogetti_formativi=m.idprogetti_formativi AND m.id_statopartecipazione=15 AND m.idprogetti_formativi = ? ORDER BY m.cognome";
                    try (PreparedStatement ps1 = db1.getConnection().prepareStatement(sql1)) {
                        ps1.setInt(1, idpr);
                        try (ResultSet rs1 = ps1.executeQuery()) {
                            while (rs1.next()) {

                                Long idallievo = rs1.getLong(1);
                                String cf = rs1.getString(2).toUpperCase().trim();
                                String cip = rs1.getString(3).toUpperCase().trim();

                                String sql2 = "SELECT p.datalezione,c.codice_ud,p.orainizio,p.orafine,d.codicefiscale,l.durataconvalidata,m.giorno,m.orario_start,m.orario_end "
                                        + "FROM presenzelezioniallievi l, presenzelezioni p, docenti d, lezioni_modelli m, lezione_calendario c "
                                        + "WHERE c.id_lezionecalendario=m.id_lezionecalendario AND m.id_lezionimodelli=p.idlezioneriferimento "
                                        + "AND p.idpresenzelezioni=l.idpresenzelezioni AND p.iddocente=d.iddocenti "
                                        + "AND l.presente=1 AND p.idprogetto = ? AND l.idallievi = ?";

                                try (PreparedStatement ps2 = db1.getConnection().prepareStatement(sql2)) {
                                    ps2.setInt(1, idpr);
                                    ps2.setLong(2, idallievo);
                                    try (ResultSet rs2 = ps2.executeQuery()) {
                                        while (rs2.next()) {
                                            String ud = rs2.getString(2);
                                            DateTime d1 = new DateTime(rs2.getDate(1).getTime());
                                            DateTime d2 = new DateTime(rs2.getDate(7).getTime());
                                            String idr = cip + "_" + ud + "_" + d1.toString("yyyyMMdd");
                                            String datainizioPRES = d1.toString("dd/MM/yyyy") + " " + rs2.getString(3);
                                            String datafinePRES = d1.toString("dd/MM/yyyy") + " " + rs2.getString(4);

                                            String datainizioLEZ = d2.toString("dd/MM/yyyy") + " " + StringUtils.substring(rs2.getString(8), 0, 5);
                                            String datafineLEZ = d2.toString("dd/MM/yyyy") + " " + StringUtils.substring(rs2.getString(9), 0, 5);

                                            String cfdoc = rs2.getString(5);
                                            long presenza = rs2.getLong(6);
                                            String[] ore_minuti = Utils.calcoladurata(presenza);

                                            XSSFRow riga_A = getRow(sh_registro, index_row.get());
                                            index_row.addAndGet(1);

                                            AtomicInteger index_column = new AtomicInteger(0);

                                            setCell(getCell(riga_A, index_column.get()), style_normal, cf, false, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, cip, false, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, idr, false, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ud, false, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, datainizioLEZ, false, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, datafineLEZ, false, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, datainizioPRES, false, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, datafinePRES, false, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, ore_minuti[0], true, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, ore_minuti[1], true, false);
                                            setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, cfdoc, false, false);
                                        }
                                    }
                                }
                                String sql3 = "SELECT r.idriunione,r.nud,r.data,r.totaleorerendicontabili FROM registro_completo r WHERE r.ruolo='ALLIEVO' "
                                        + "AND r.idutente = ? AND r.idprogetti_formativi = ?";
                                try (PreparedStatement ps3 = db1.getConnection().prepareStatement(sql3)) {
                                    ps3.setLong(1, idallievo);
                                    ps3.setInt(2, idpr);
                                    try (ResultSet rs2 = ps3.executeQuery()) {
                                        while (rs2.next()) {
                                            String idr = rs2.getString(1);
                                            String ud = rs2.getString(2);
                                            DateTime d1 = new DateTime(rs2.getDate(3).getTime());
                                            long presenza = rs2.getLong(4);
                                            String[] ore_minuti = Utils.calcoladurata(presenza);
                                            String sql4 = "SELECT d.codicefiscale,c.orainizio,c.orafine,d.iddocenti "
                                                    + "FROM registro_completo c, docenti d WHERE c.idutente=d.iddocenti "
                                                    + "AND c.idriunione = ? AND c.ruolo<>'ALLIEVO' ";

                                            try (PreparedStatement ps4 = db1.getConnection().prepareStatement(sql4)) {
                                                ps4.setString(1, idr);
                                                try (ResultSet rs4 = ps4.executeQuery()) {
                                                    if (rs4.next()) {

                                                        String datainizioPRES = d1.toString("dd/MM/yyyy") + " " + rs4.getString(2);
                                                        String datafinePRES = d1.toString("dd/MM/yyyy") + " " + rs4.getString(3);
                                                        String cfdoc = rs4.getString(1);

                                                        String sql5 = "SELECT lm.orario_start,lm.orario_end "
                                                                + "FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, unita_didattiche ud "
                                                                + "WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario "
                                                                + "AND ud.codice=lc.codice_ud AND lm.tipolez='F' AND lm.id_docente = ? AND mp.id_progettoformativo = ? AND lm.giorno = ?";

                                                        try (PreparedStatement ps5 = db1.getConnection().prepareStatement(sql5)) {
                                                            ps5.setLong(1, rs4.getLong(4));
                                                            ps5.setInt(2, idpr);
                                                            ps5.setString(3, rs2.getString(3));
                                                            try (ResultSet rs5 = ps5.executeQuery()) {
                                                                if (rs5.next()) {
                                                                    String datainizioLEZ = d1.toString("dd/MM/yyyy") + " " + StringUtils.substring(rs5.getString(1), 0, 5);
                                                                    String datafineLEZ = d1.toString("dd/MM/yyyy") + " " + StringUtils.substring(rs5.getString(2), 0, 5);
                                                                    XSSFRow riga_A = getRow(sh_registro, index_row.get());
                                                                    index_row.addAndGet(1);

                                                                    AtomicInteger index_column = new AtomicInteger(0);

                                                                    setCell(getCell(riga_A, index_column.get()), style_normal, cf, false, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, cip, false, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, idr, false, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, ud, false, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, datainizioLEZ, false, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, datafineLEZ, false, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, datainizioPRES, false, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, datafinePRES, false, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, ore_minuti[0], true, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_int, ore_minuti[1], true, false);
                                                                    setCell(getCell(riga_A, index_column.addAndGet(1)), style_normal, cfdoc, false, false);
                                                                }
                                                            }
                                                        }

                                                    }
                                                }
                                            }

                                        }
                                    }
                                }
                            }
                        }
                    }

                }

//                for (int ix = 0; ix < 13; ix++) {
//                    sh_registro.autoSizeColumn(ix);
//                }
                File output_xlsx = new File(pathdest + "/" + nomerend + "_CalendarioAllievi_" + new DateTime().toString(timestamp) + ".xlsx");
                try (FileOutputStream outputStream = new FileOutputStream(output_xlsx)) {
                    wb.write(outputStream);
                    log.log(Level.INFO, "FILE RILASCIATO: {0}", output_xlsx.getPath());
                    return output_xlsx;
                }

            }
        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }
        return null;
    }

    public static void main(String[] args) {
        generaRendicontazione(true);

//        List<Integer> start = new ArrayList<>();
//        start.add(10);
////
//        prospetto_riepilogo_calendario_allievi(0, start);
    }
}
