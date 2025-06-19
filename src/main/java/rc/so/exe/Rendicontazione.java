package rc.so.exe;

import com.google.common.util.concurrent.AtomicDouble;
import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import static org.apache.commons.lang3.StringUtils.substring;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.stream.Collectors;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import static rc.so.exe.Constant.estraiEccezione;
import static rc.so.exe.Constant.patternITA;
import static rc.so.exe.Constant.sdfITA;
import static rc.so.exe.Constant.zipListFiles;
import static rc.so.exe.Sicilia_gestione.log;
import static rc.so.exe.Utils.calcolaintervallomillis;
import static rc.so.exe.Utils.createFile;
import static rc.so.exe.Utils.getCell;
import static rc.so.exe.Utils.getRow;
import static rc.so.exe.Utils.parseIntR;
import static rc.so.exe.Utils.setCell;
import static rc.so.exe.Utils.timestamp;
import rc.so.report.FaseA;
import rc.so.report.Registro_completo;
import rc.so.report.Utenti;

public class Rendicontazione {

    private static final String separator = "###";
    private static final String formatdataCellint = "#,#";
    private static final String formatdataCell = "#,#.00";

    private static final byte[] bianco = {(byte) 255, (byte) 255, (byte) 255};
    private static final byte[] color1 = {(byte) 49, (byte) 134, (byte) 155};
    private static final byte[] color2 = {(byte) 83, (byte) 141, (byte) 213};
    private static final byte[] color3 = {(byte) 197, (byte) 217, (byte) 241};
    private static final byte[] color4 = {(byte) 238, (byte) 30, (byte) 30};
    private static final byte[] color5 = {(byte) 0, (byte) 204, (byte) 0};
    private static final byte[] color6 = {(byte) 0, (byte) 128, (byte) 128};
    private static final byte[] color7 = {(byte) 192, (byte) 192, (byte) 192};
    private static final XSSFColor myColor1 = new XSSFColor(color1, new DefaultIndexedColorMap());
    private static final XSSFColor myColor2 = new XSSFColor(color2, new DefaultIndexedColorMap());
    private static final XSSFColor myColor3 = new XSSFColor(color3, new DefaultIndexedColorMap());
    private static final XSSFColor myColor4 = new XSSFColor(color4, new DefaultIndexedColorMap());
    private static final XSSFColor myColor5 = new XSSFColor(color5, new DefaultIndexedColorMap());
    private static final XSSFColor myColor6 = new XSSFColor(color6, new DefaultIndexedColorMap());
    private static final XSSFColor myColor7 = new XSSFColor(color7, new DefaultIndexedColorMap());
    private static final XSSFColor white = new XSSFColor(bianco, new DefaultIndexedColorMap());
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

    private static final double costo_ora_docenza_A = 164.53;
    private static final String S_costo_ora_docenza_A = "164,53€";
    private static final double costo_ora_docenza_B = 131.63;
    private static final String S_costo_ora_docenza_B = "131,63€";
    private static final double costo_ora_docenza_C = 82.27;
    private static final String S_costo_ora_docenza_C = "82,27€";
    private static final double costo_ora_fasea = 0.90;
    private static final String S_costo_ora_fasea = "0,90€";
    private static final double costo_ora_faseb = 45;
    private static final String S_costo_ora_faseb = "45€";

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

                File xlsx0 = prospetto_riepilogo_2025(0, idpr);
                File xlsx1 = prospetto_riepilogo(0, idpr);
                File xlsx2 = prospetto_riepilogo_allievi(0, idpr, true);
                List<File> calendar = list_calendar(idpr);
                File xlsx3 = prospetto_riepilogo_calendario_allievi(0, idpr);

                List<File> output = new ArrayList<>();
                output.add(xlsx0);
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
                        File xlsx0 = prospetto_riepilogo_2025(0, idpr);
                        File xlsx1 = prospetto_riepilogo(idestrazione, idpr);
                        File xlsx2 = prospetto_riepilogo_allievi(idestrazione, idpr, false);

                        if (xlsx0 != null && xlsx1 != null && xlsx2 != null) {

                            List<File> calendar = list_calendar(idpr);

                            File xlsx3 = prospetto_riepilogo_calendario_allievi(idestrazione, idpr);

                            List<File> output = new ArrayList<>();
                            output.add(xlsx0);
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

    private static File prospetto_riepilogo_allievi(int idestrazione, List<Integer> list_idpr, boolean includeall) {
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
                            if (includeall) {
                                sql2 = "SELECT * FROM allievi a WHERE a.idprogetti_formativi = " + idpr + " ORDER BY codicefiscale";
                            }

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

    private static void cleanBeforeMergeOnValidCells(XSSFSheet sheet, CellRangeAddress region, XSSFCellStyle cellStyle, boolean add) {
        try {
            for (int rowNum = region.getFirstRow(); rowNum <= region.getLastRow(); rowNum++) {
                XSSFRow row = getRow(sheet, rowNum);
                for (int colNum = region.getFirstColumn(); colNum <= region.getLastColumn(); colNum++) {
                    XSSFCell currentCell = getCell(row, colNum);
                    currentCell.setCellStyle(cellStyle);
                }
            }

        } catch (Exception ex) {
            log.severe(estraiEccezione(ex));
        }
        if (add) {
            try {
                sheet.addMergedRegion(region);
            } catch (Exception ex) {
                log.severe(estraiEccezione(ex));
            }
        }

    }

    private static double roundDouble(double f, boolean converttoHours) {
        try {
            if (converttoHours) {
                double hours = f / 1000.0 / 60.0 / 60.0;
                BigDecimal bigDecimal = new BigDecimal(hours);
                bigDecimal = bigDecimal.setScale(2, RoundingMode.HALF_EVEN);
                return bigDecimal.doubleValue();
            } else {
                BigDecimal bigDecimal = new BigDecimal(f);
                bigDecimal = bigDecimal.setScale(2, RoundingMode.HALF_EVEN);
                return bigDecimal.doubleValue();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return 0.0;
    }

    private static File prospetto_riepilogo_2025(int idestrazione, List<Integer> list_idpr) {
        File output_xlsx = null;
        DateTime oggi = new DateTime();
        String nomerend_cod = "R" + idestrazione;
        String nomerend = nomerend_cod + "_" + oggi.toString("ddMMyyyy");
        try {
            FaseA FA = new FaseA(false);
            Db_Gest db1 = new Db_Gest(FA.getHost());
            String pathdest = db1.getPath("output_excel_archive");
            String fileing = pathdest + "YISUS_Prospetto_Riepilogo_v2.xlsx";
            AtomicDouble total_ore = new AtomicDouble(0.0);
            AtomicDouble total_rend = new AtomicDouble(0.0);

            File ing = createFile(fileing);
            if (ing == null) {
                db1.closeDB();
                return null;
            }
            try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(ing), false)) {

                XSSFSheet sh1 = wb.getSheet("Prospetto di riepilogo DdR XX");
                wb.setSheetName(sh1.getWorkbook().getSheetIndex(sh1.getSheetName()), "Prospetto di riepilogo DdR " + nomerend_cod);

                XSSFFont font_total = wb.createFont();
                font_total.setFontHeightInPoints((short) 12);
                font_total.setBold(true);

                XSSFFont font_white = wb.createFont();
                font_white.setFontHeightInPoints((short) 14);
                font_white.setBold(true);
                font_white.setColor(white);
                XSSFFont font_int = wb.createFont();
                font_int.setFontHeightInPoints((short) 14);
                font_int.setBold(true);

                XSSFCellStyle intestazione_1 = wb.createCellStyle();
                intestazione_1.setVerticalAlignment(VerticalAlignment.CENTER);
                intestazione_1.setAlignment(HorizontalAlignment.CENTER);
                intestazione_1.setBorderBottom(BorderStyle.THIN);
                intestazione_1.setBorderTop(BorderStyle.THIN);
                intestazione_1.setBorderRight(BorderStyle.THIN);
                intestazione_1.setBorderLeft(BorderStyle.THIN);
                intestazione_1.setFillForegroundColor(myColor1);
                intestazione_1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                intestazione_1.setFont(font_white);

                XSSFCellStyle intestazione_2 = wb.createCellStyle();
                intestazione_2.setVerticalAlignment(VerticalAlignment.CENTER);
                intestazione_2.setAlignment(HorizontalAlignment.CENTER);
                intestazione_2.setBorderBottom(BorderStyle.THIN);
                intestazione_2.setBorderTop(BorderStyle.THIN);
                intestazione_2.setBorderRight(BorderStyle.THIN);
                intestazione_2.setBorderLeft(BorderStyle.THIN);
                intestazione_2.setFillForegroundColor(myColor2);
                intestazione_2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                intestazione_2.setFont(font_white);

                XSSFCellStyle intestazione_2A = wb.createCellStyle();
                intestazione_2A.setVerticalAlignment(VerticalAlignment.CENTER);
                intestazione_2A.setAlignment(HorizontalAlignment.CENTER);
                intestazione_2A.setBorderBottom(BorderStyle.THIN);
                intestazione_2A.setBorderTop(BorderStyle.THIN);
                intestazione_2A.setBorderRight(BorderStyle.THIN);
                intestazione_2A.setBorderLeft(BorderStyle.THIN);
                intestazione_2A.setFillForegroundColor(myColor6);
                intestazione_2A.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                intestazione_2A.setFont(font_white);

                XSSFCellStyle intestazione_2B = wb.createCellStyle();
                intestazione_2B.setVerticalAlignment(VerticalAlignment.CENTER);
                intestazione_2B.setAlignment(HorizontalAlignment.CENTER);
                intestazione_2B.setBorderBottom(BorderStyle.THIN);
                intestazione_2B.setBorderTop(BorderStyle.THIN);
                intestazione_2B.setBorderRight(BorderStyle.THIN);
                intestazione_2B.setBorderLeft(BorderStyle.THIN);
                intestazione_2B.setFillForegroundColor(myColor4);
                intestazione_2B.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                intestazione_2B.setFont(font_white);

                XSSFCellStyle intestazione_3 = wb.createCellStyle();
                intestazione_3.setVerticalAlignment(VerticalAlignment.CENTER);
                intestazione_3.setAlignment(HorizontalAlignment.CENTER);
                intestazione_3.setBorderBottom(BorderStyle.THIN);
                intestazione_3.setBorderTop(BorderStyle.THIN);
                intestazione_3.setBorderRight(BorderStyle.THIN);
                intestazione_3.setBorderLeft(BorderStyle.THIN);
                intestazione_3.setFillForegroundColor(myColor3);
                intestazione_3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                intestazione_3.setFont(font_int);

                XSSFCellStyle intestazione_4 = wb.createCellStyle();
                intestazione_4.setVerticalAlignment(VerticalAlignment.CENTER);
                intestazione_4.setAlignment(HorizontalAlignment.CENTER);
                intestazione_4.setBorderBottom(BorderStyle.THIN);
                intestazione_4.setBorderTop(BorderStyle.THIN);
                intestazione_4.setBorderRight(BorderStyle.THIN);
                intestazione_4.setBorderLeft(BorderStyle.THIN);
                intestazione_4.setFillForegroundColor(myColor4);
                intestazione_4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                intestazione_4.setFont(font_white);

                XSSFCellStyle intestazione_5 = wb.createCellStyle();
                intestazione_5.setVerticalAlignment(VerticalAlignment.CENTER);
                intestazione_5.setAlignment(HorizontalAlignment.CENTER);
                intestazione_5.setBorderBottom(BorderStyle.THIN);
                intestazione_5.setBorderTop(BorderStyle.THIN);
                intestazione_5.setBorderRight(BorderStyle.THIN);
                intestazione_5.setBorderLeft(BorderStyle.THIN);
                intestazione_5.setFillForegroundColor(myColor5);
                intestazione_5.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                intestazione_5.setFont(font_white);

                XSSFDataFormat xssfDataFormat = wb.createDataFormat();
                XSSFCellStyle cellStyle_int = wb.createCellStyle();
                cellStyle_int.setBorderBottom(BorderStyle.THIN);
                cellStyle_int.setBorderTop(BorderStyle.THIN);
                cellStyle_int.setBorderRight(BorderStyle.THIN);
                cellStyle_int.setBorderLeft(BorderStyle.THIN);
                cellStyle_int.setVerticalAlignment(VerticalAlignment.CENTER);
                cellStyle_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));

                XSSFCellStyle cellStyle_double = wb.createCellStyle();
                cellStyle_double.setBorderBottom(BorderStyle.THIN);
                cellStyle_double.setBorderTop(BorderStyle.THIN);
                cellStyle_double.setBorderRight(BorderStyle.THIN);
                cellStyle_double.setBorderLeft(BorderStyle.THIN);
                cellStyle_double.setVerticalAlignment(VerticalAlignment.CENTER);
                cellStyle_double.setDataFormat(xssfDataFormat.getFormat(formatdataCell));

                XSSFCellStyle cs = wb.createCellStyle();
                cs.setVerticalAlignment(VerticalAlignment.CENTER);
                cs.setAlignment(HorizontalAlignment.CENTER);
                cs.setBorderBottom(BorderStyle.THIN);
                cs.setBorderTop(BorderStyle.THIN);
                cs.setBorderRight(BorderStyle.THIN);
                cs.setBorderLeft(BorderStyle.THIN);

                XSSFCellStyle cstotal = wb.createCellStyle();
                cstotal.setVerticalAlignment(VerticalAlignment.CENTER);
                cstotal.setAlignment(HorizontalAlignment.CENTER);
                cstotal.setBorderBottom(BorderStyle.THIN);
                cstotal.setBorderTop(BorderStyle.THIN);
                cstotal.setBorderRight(BorderStyle.THIN);
                cstotal.setBorderLeft(BorderStyle.THIN);
                cstotal.setFillForegroundColor(myColor7);
                cstotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cstotal.setFont(font_total);

                XSSFCellStyle cstotal_double = wb.createCellStyle();
                cstotal_double.setVerticalAlignment(VerticalAlignment.CENTER);
                cstotal_double.setAlignment(HorizontalAlignment.CENTER);
                cstotal_double.setBorderBottom(BorderStyle.THIN);
                cstotal_double.setBorderTop(BorderStyle.THIN);
                cstotal_double.setBorderRight(BorderStyle.THIN);
                cstotal_double.setBorderLeft(BorderStyle.THIN);
                cstotal_double.setFillForegroundColor(myColor7);
                cstotal_double.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cstotal_double.setDataFormat(xssfDataFormat.getFormat(formatdataCell));
                cstotal_double.setFont(font_total);

                XSSFCellStyle cstotal_int = wb.createCellStyle();

                cstotal_int.setVerticalAlignment(VerticalAlignment.CENTER);
                cstotal_int.setAlignment(HorizontalAlignment.CENTER);
                cstotal_int.setBorderBottom(BorderStyle.THIN);
                cstotal_int.setBorderTop(BorderStyle.THIN);
                cstotal_int.setBorderRight(BorderStyle.THIN);
                cstotal_int.setBorderLeft(BorderStyle.THIN);
                cstotal_int.setFillForegroundColor(myColor7);
                cstotal_int.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cstotal_int.setDataFormat(xssfDataFormat.getFormat(formatdataCellint));
                cstotal_int.setFont(font_total);

                AtomicInteger index_row = new AtomicInteger(9);
                AtomicInteger indice = new AtomicInteger(1);
                for (int ss = 0; ss < list_idpr.size(); ss++) {

                    AtomicDouble total_pr = new AtomicDouble(0.0);

                    int idpr = list_idpr.get(ss);

                    String sql1 = "SELECT p.idprogetti_formativi,p.cip,s.ragionesociale,c.cod_provincia,p.start,p.end "
                            + "FROM progetti_formativi p, soggetti_attuatori s,comuni c WHERE p.stato='CO' "
                            + "AND p.idsoggetti_attuatori=s.idsoggetti_attuatori AND c.idcomune=s.comune AND p.idprogetti_formativi = " + idpr;

                    try (Statement st1 = db1.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {
                        if (rs1.next()) {

                            String cip = rs1.getString(2).toUpperCase();

                            String ragionesociale = rs1.getString(3).toUpperCase();
//                            String cod_provincia = rs1.getString(4).toUpperCase();
                            String start = sdfITA.format(rs1.getDate(5));
                            String end = sdfITA.format(rs1.getDate(6));

                            List<Utenti> allievi_OK = db1.list_Allievi_OK(idpr);
                            List<Utenti> docenti_tab = db1.list_Docenti(idpr);

                            List<Registro_completo> registrocompleto = db1.registro_modello6(idpr);
                            List<Items> calendario = db1.calendario(idpr);
                            LinkedList<Items> calendarioA = calendario.stream().filter(l1
                                    -> l1.getFase().endsWith("A")).collect(Collectors.toCollection(LinkedList::new));

                            LinkedList<Items> calendarioB = calendario.stream().filter(l1
                                    -> l1.getFase().endsWith("B")).collect(Collectors.toCollection(LinkedList::new));

                            HashMap<Integer, String> bus_pla = new HashMap<>();
                            String sql1A = "SELECT m.allievo,m.businessplan_presente FROM maschera_m5 m WHERE m.progetto_formativo = " + idpr;
                            try (Statement st2 = db1.getConnection().createStatement(); ResultSet rs1A = st2.executeQuery(sql1A)) {
                                while (rs1A.next()) {
                                    bus_pla.put(rs1A.getInt(1), rs1A.getString(2));
                                }
                            }

                            String DATAINIZIOFASEA = calendarioA.getFirst().getData();
                            String DATAFINEFASEA = calendarioA.getLast().getData();

                            List<Registro_completo> faseA = registrocompleto.stream().filter(r1
                                    -> r1.getRuolo().contains("ALLIEVO")
                                    && r1.getFase().equalsIgnoreCase("A"))
                                    .collect(Collectors.toList());

                            List<Registro_completo> faseB = registrocompleto.stream().filter(r1
                                    -> r1.getRuolo().contains("ALLIEVO")
                                    && r1.getFase().equalsIgnoreCase("B"))
                                    .collect(Collectors.toList());

                            int numpartecipanti = allievi_OK.size();

                            XSSFSheet sh_pr = wb.createSheet(cip);

                            AtomicInteger indici_docenti = new AtomicInteger(15 + numpartecipanti);

                            //docenti
                            XSSFRow row_docenti = getRow(sh_pr, indici_docenti.get());
                            CellRangeAddress region_20 = new CellRangeAddress(row_docenti.getRowNum(), row_docenti.getRowNum(), 5, (12 + calendarioA.size()));
                            cleanBeforeMergeOnValidCells(sh_pr, region_20, cs, true);

                            setCell(getCell(row_docenti, 5), intestazione_1, "DOCENTI FASE A", false, false);

                            indici_docenti.addAndGet(1);
                            XSSFRow row_docenti2 = getRow(sh_pr, indici_docenti.get());

                            setCell(getCell(row_docenti2, 5), cs, "N.", false, false);
                            CellRangeAddress region_21 = new CellRangeAddress(row_docenti2.getRowNum(), row_docenti2.getRowNum(), 6, 7);
                            cleanBeforeMergeOnValidCells(sh_pr, region_21, cs, true);
                            setCell(getCell(row_docenti2, 6), cs, "COGNOME", false, false);
                            setCell(getCell(row_docenti2, 8), cs, "NOME", false, false);
                            setCell(getCell(row_docenti2, 9), cs, "CODICE FISCALE", false, false);
                            setCell(getCell(row_docenti2, 10), cs, "COSTO ORA/DOCENZA", false, false);

                            for (int x = 0; x < calendarioA.size(); x++) {
                                String cal1 = calendarioA.get(x).getData();
                                setCell(getCell(row_docenti2, 11 + x), cs, cal1, false, false);
                            }

                            setCell(getCell(row_docenti2, 11 + calendarioA.size()), cs, "TOTALE ORE (MAX 80h)", false, false);
                            setCell(getCell(row_docenti2, 12 + calendarioA.size()), cs, "TOTALE IMPORTO DOCENZA FASE A", false, false);

                            indici_docenti.addAndGet(1);

                            AtomicInteger numdocenti = new AtomicInteger(0);
                            AtomicDouble tot_ore_docenti = new AtomicDouble(0.0);
                            AtomicDouble tot_docenti = new AtomicDouble(0.0);

                            List<Registro_completo> docentifaseA = registrocompleto.stream().filter(r1
                                    -> r1.getRuolo().equalsIgnoreCase("DOCENTE")
                                    && r1.getFase().equalsIgnoreCase("A")).collect(Collectors.toList());

                            List<Integer> docentiid = docentifaseA.stream().map(r1
                                    -> r1.getIdutente()).distinct().collect(Collectors.toList());

                            docentiid.forEach(r1 -> {
                                Utenti docente = docenti_tab.stream().filter(d1 -> d1.getId() == r1).findAny().orElse(null);
                                if (docente != null) {
                                    numdocenti.addAndGet(1);
                                    XSSFRow row_d = getRow(sh_pr, indici_docenti.get());
                                    setCell(getCell(row_d, 5), cs, String.valueOf(numdocenti.get()), true, false);

                                    CellRangeAddress region_22 = new CellRangeAddress(row_d.getRowNum(), row_d.getRowNum(), 6, 7);
                                    cleanBeforeMergeOnValidCells(sh_pr, region_22, cs, true);

                                    setCell(getCell(row_d, 6), cs, docente.getCognome(), false, false);

                                    setCell(getCell(row_d, 8), cs, docente.getNome(), false, false);
                                    setCell(getCell(row_d, 9), cs, docente.getCf(), false, false);

                                    if (docente.getFascia().endsWith("A")) {
                                        setCell(getCell(row_d, 10), cellStyle_double, String.valueOf(costo_ora_docenza_A), false, true);
                                    } else if (docente.getFascia().endsWith("B")) {
                                        setCell(getCell(row_d, 10), cellStyle_double, String.valueOf(costo_ora_docenza_B), false, true);
                                    } else {
                                        setCell(getCell(row_d, 10), cellStyle_double, String.valueOf(costo_ora_docenza_C), false, true);
                                    }

                                    AtomicDouble tot_ore_fase_A = new AtomicDouble(0.0);

                                    for (int x = 0; x < calendarioA.size(); x++) {
                                        String cal1 = calendarioA.get(x).getData();
                                        List<Registro_completo> docente_ore = docentifaseA.stream().filter(
                                                r3 -> r3.getIdutente() == r1 && r3.getData().toString(patternITA).equals(cal1))
                                                .collect(Collectors.toList());

                                        if (!docente_ore.isEmpty()) {
                                            long ore = 0L;

                                            for (Registro_completo rr : docente_ore) {
                                                ore += rr.getTotaleorerendicontabili();
                                            }

                                            setCell(getCell(row_d, 11 + x),
                                                    cellStyle_double,
                                                    String.valueOf(roundDouble(ore, true)),
                                                    false, true);

                                            tot_ore_fase_A.addAndGet(roundDouble(ore, true));

                                        } else {
                                            setCell(getCell(row_d, 11 + x),
                                                    cellStyle_double,
                                                    "0.00",
                                                    false, true);
                                        }

                                    }

                                    setCell(getCell(row_d, (11 + calendarioA.size())), cellStyle_double,
                                            String.valueOf(tot_ore_fase_A.get()),
                                            false, true);

                                    double tot_d1;

                                    if (docente.getFascia().endsWith("A")) {
                                        tot_d1 = tot_ore_fase_A.get() * costo_ora_docenza_A;
                                    } else if (docente.getFascia().endsWith("B")) {
                                        tot_d1 = tot_ore_fase_A.get() * costo_ora_docenza_B;
                                    } else {
                                        tot_d1 = tot_ore_fase_A.get() * costo_ora_docenza_C;
                                    }

                                    tot_docenti.addAndGet(tot_d1);
                                    tot_ore_docenti.addAndGet(tot_ore_fase_A.get());
                                    setCell(getCell(row_d, (12 + calendarioA.size())), cellStyle_double, String.valueOf(tot_d1), false, true);

                                    indici_docenti.addAndGet(1);

                                }
                            });

                            XSSFRow row_dt = getRow(sh_pr, indici_docenti.get());
                            setCell(getCell(row_dt, (11 + calendarioA.size())), cstotal_double,
                                    String.valueOf(tot_ore_docenti.get()), false, true);
                            setCell(getCell(row_dt, (12 + calendarioA.size())), cstotal_double,
                                    String.valueOf(tot_docenti.get()), false, true);

                            //ALLIEVI
                            XSSFRow row_intest = getRow(sh_pr, 2);
                            XSSFRow row_intest2 = getRow(sh_pr, 3);
                            XSSFRow row_intest3 = getRow(sh_pr, 4);
                            XSSFRow row_intest4 = getRow(sh_pr, 5);
                            XSSFRow row_intest5 = getRow(sh_pr, 6);

                            CellRangeAddress region_1 = new CellRangeAddress(2, 3, 1, 6);
                            CellRangeAddress region_2 = new CellRangeAddress(2, 3, 7, 10);
                            CellRangeAddress region_3 = new CellRangeAddress(2, 2, 11, 13 + calendarioA.size());
                            CellRangeAddress region_4 = new CellRangeAddress(4, 6, 1, 1);
                            CellRangeAddress region_5 = new CellRangeAddress(4, 6, 2, 2);
                            CellRangeAddress region_6 = new CellRangeAddress(4, 6, 3, 3);
                            CellRangeAddress region_7 = new CellRangeAddress(4, 6, 4, 4);
                            CellRangeAddress region_8 = new CellRangeAddress(4, 6, 5, 5);
                            CellRangeAddress region_9 = new CellRangeAddress(4, 6, 6, 6);
                            CellRangeAddress region_10 = new CellRangeAddress(4, 6, 7, 7);
                            CellRangeAddress region_11 = new CellRangeAddress(4, 6, 8, 8);
                            CellRangeAddress region_12 = new CellRangeAddress(4, 6, 9, 9);
                            CellRangeAddress region_13 = new CellRangeAddress(4, 6, 10, 10);

                            cleanBeforeMergeOnValidCells(sh_pr, region_1, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_2, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_3, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_4, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_5, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_6, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_7, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_8, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_9, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_10, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_11, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_12, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_13, cs, true);

                            setCell(getCell(row_intest, 1), intestazione_1, "ANAGRAFICA PERCORSO", false, false);
                            setCell(getCell(row_intest, 7), intestazione_2, "ANAGRAFICA PARTECIPANTI", false, false);
                            setCell(getCell(row_intest, 11), intestazione_3, "FASE A DATE", false, false);
                            setCell(getCell(row_intest3, 1), cs, "N.", false, false);
                            setCell(getCell(row_intest3, 2), cs, "ID", false, false);
                            setCell(getCell(row_intest3, 3), cs, "CIP", false, false);
                            setCell(getCell(row_intest3, 4), cs, "SA", false, false);
                            setCell(getCell(row_intest3, 5), cs, "DATA INIZIO CORSO", false, false);
                            setCell(getCell(row_intest3, 6), cs, "DATA FINE CORSO", false, false);
                            setCell(getCell(row_intest3, 7), cs, "COGNOME", false, false);
                            setCell(getCell(row_intest3, 8), cs, "NOME", false, false);
                            setCell(getCell(row_intest3, 9), cs, "CODICE FISCALE", false, false);
                            setCell(getCell(row_intest3, 10), cs, "BUSINESS PLAN (SI/NO)", false, false);

                            for (int x = 0; x < calendarioA.size(); x++) {
                                Items it1 = calendarioA.get(x);
                                String cal1 = it1.getData();
                                setCell(getCell(row_intest3, 11 + x), cs, cal1, false, false);
                                setCell(getCell(row_intest4, 11 + x), cs, substring(it1.getOrainizio(), 0, 5), false, false);
                                setCell(getCell(row_intest5, 11 + x), cs, substring(it1.getOrafine(), 0, 5), false, false);
                            }

                            CellRangeAddress region_14 = new CellRangeAddress(3, 3, 11, 13 + calendarioA.size());
                            CellRangeAddress region_16 = new CellRangeAddress(4, 6, 11 + calendarioA.size(), 11 + calendarioA.size());
                            CellRangeAddress region_17 = new CellRangeAddress(4, 6, 12 + calendarioA.size(), 12 + calendarioA.size());
                            CellRangeAddress region_18 = new CellRangeAddress(4, 6, 13 + calendarioA.size(), 13 + calendarioA.size());

                            cleanBeforeMergeOnValidCells(sh_pr, region_14, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_16, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_17, cs, true);
                            cleanBeforeMergeOnValidCells(sh_pr, region_18, cs, true);

                            setCell(getCell(row_intest2, 11), intestazione_3, "DAL " + DATAINIZIOFASEA + " AL " + DATAFINEFASEA, false, false);
                            setCell(getCell(row_intest3, 11 + calendarioA.size()), cs, "TOTALE ORE (MAX 80h)", false, false);
                            setCell(getCell(row_intest3, 12 + calendarioA.size()), cs, "TOTALE ORE * " + S_costo_ora_fasea, false, false);
                            setCell(getCell(row_intest3, 13 + calendarioA.size()), cs, "TOTALE FASE A DOCENZA", false, false);

                            List<String> elencogruppi_B = calendario.stream().filter(it1 -> it1.getFase().equals("B"))
                                    .map(Items::getGruppo).distinct().sorted().collect(Collectors.toList());

                            AtomicInteger iniziofaseb = new AtomicInteger(15 + calendarioA.size());

                            HashMap<String, Integer> indicifaseB = new HashMap<>();
                            HashMap<String, Double> totaliFaseB = new HashMap<>();

                            int iniziofaseb_i;
                            for (int ix = 0; ix < elencogruppi_B.size(); ix++) {
                                String numerogruppo = elencogruppi_B.get(ix);
                                iniziofaseb_i = iniziofaseb.get();
                                indicifaseB.put(numerogruppo, iniziofaseb_i);

                                LinkedList<Items> calendarioB2 = calendarioB.stream().filter(l1
                                        -> l1.getFase().endsWith("B") && l1.getGruppo().equals(numerogruppo))
                                        .collect(Collectors.toCollection(LinkedList::new));

                                String datainizioB = calendarioB2.getFirst().getData();
                                String datafineB = calendarioB2.getLast().getData();

                                for (int x = 0; x < calendarioB2.size(); x++) {
                                    Items it1 = calendarioB2.get(x);
                                    setCell(getCell(row_intest3, iniziofaseb.get()), cs, it1.getData(), false, false);
                                    setCell(getCell(row_intest4, iniziofaseb.get()), cs, substring(it1.getOrainizio(), 0, 5), false, false);
                                    setCell(getCell(row_intest5, iniziofaseb.get()), cs, substring(it1.getOrafine(), 0, 5), false, false);
                                    iniziofaseb.addAndGet(1);
                                }
                                iniziofaseb.addAndGet(2);

                                CellRangeAddress region_B1 = new CellRangeAddress(2, 2, iniziofaseb_i, iniziofaseb.get() - 1);
                                cleanBeforeMergeOnValidCells(sh_pr, region_B1, cs, true);
                                CellRangeAddress region_B2 = new CellRangeAddress(3, 3, iniziofaseb_i, iniziofaseb.get() - 1);
                                cleanBeforeMergeOnValidCells(sh_pr, region_B2, cs, true);
                                setCell(getCell(row_intest, iniziofaseb_i), intestazione_2, "FASE B GRUPPO " + numerogruppo + " DATE", false, false);
                                setCell(getCell(row_intest2, iniziofaseb_i), intestazione_2, "DAL " + datainizioB + " AL " + datafineB, false, false);
                                CellRangeAddress region_B3 = new CellRangeAddress(4, 6, iniziofaseb.get() - 2, iniziofaseb.get() - 2);
                                cleanBeforeMergeOnValidCells(sh_pr, region_B3, cs, true);
                                setCell(getCell(row_intest3, iniziofaseb.get() - 2), cs, "TOTALE ORE (MAX 40h)", false, false);
                                CellRangeAddress region_B4 = new CellRangeAddress(4, 6, iniziofaseb.get() - 1, iniziofaseb.get() - 1);
                                cleanBeforeMergeOnValidCells(sh_pr, region_B4, cs, true);
                                setCell(getCell(row_intest3, iniziofaseb.get() - 1), cs, "TOTALE ORE FASE B * " + S_costo_ora_faseb, false, false);
                            }

                            AtomicInteger index_allievo = new AtomicInteger(7);

                            AtomicDouble componente_allievi_A = new AtomicDouble(0.0);
                            AtomicDouble componente_docenti_A = new AtomicDouble(0.0);

                            for (int y = 0; y < allievi_OK.size(); y++) {
                                Utenti allievo = allievi_OK.get(y);
//                                        if (rc != null) {
                                int indicepartenzafaseBallievo;

                                XSSFRow row_allievo = getRow(sh_pr, index_allievo.get());
                                setCell(getCell(row_allievo, 1), cellStyle_int, String.valueOf(index_allievo.get() - 6), true, false);
                                setCell(getCell(row_allievo, 2), cellStyle_int, String.valueOf(idpr), true, false);
                                setCell(getCell(row_allievo, 3), cs, cip, false, false);
                                setCell(getCell(row_allievo, 4), cs, ragionesociale, false, false);
                                setCell(getCell(row_allievo, 5), cs, start, false, false);
                                setCell(getCell(row_allievo, 6), cs, end, false, false);
                                setCell(getCell(row_allievo, 7), cs, allievo.getCognome().toUpperCase(), false, false);
                                setCell(getCell(row_allievo, 8), cs, allievo.getNome().toUpperCase(), false, false);
                                setCell(getCell(row_allievo, 9), cs, allievo.getCf().toUpperCase(), false, false);

                                if (bus_pla.get(allievo.getId()) == null) {
                                    setCell(getCell(row_allievo, 10), cs, "NO", false, false);
                                } else {
                                    if (bus_pla.get(allievo.getId()).equals("1")) {
                                        setCell(getCell(row_allievo, 10), cs, "SI", false, false);
                                    } else {
                                        setCell(getCell(row_allievo, 10), cs, "NO", false, false);
                                    }
                                }

                                AtomicDouble a_tot_ore = new AtomicDouble(0.0);
                                AtomicDouble b_tot_ore = new AtomicDouble(0.0);

                                AtomicDouble a_tot_al = new AtomicDouble(0.0);
                                AtomicDouble b_tot_al = new AtomicDouble(0.0);

                                AtomicDouble d_tot = new AtomicDouble(0.0);

                                //FASE A
                                for (int x = 0; x < calendarioA.size(); x++) {
                                    Items it1 = calendarioA.get(x);
                                    String cal1 = it1.getData();

                                    List<Registro_completo> rc_A = faseA.stream().filter(
                                            a1 -> a1.getIdutente() == allievo.getId() && a1.getData().toString(patternITA).equals(cal1))
                                            .collect(Collectors.toList());

                                    if (!rc_A.isEmpty()) {
                                        long ore = 0L;

                                        for (Registro_completo rr : rc_A) {
                                            ore += rr.getTotaleorerendicontabili();
                                        }

                                        setCell(getCell(row_allievo, 11 + x),
                                                cellStyle_double,
                                                String.valueOf(roundDouble(ore, true)),
                                                false, true);
                                        a_tot_ore.addAndGet(roundDouble(ore, true));
                                    } else {
                                        setCell(getCell(row_allievo, 11 + x),
                                                cellStyle_double,
                                                "0.0",
                                                false, true);
                                    }

                                }

                                setCell(getCell(row_allievo, 11 + calendarioA.size()),
                                        cellStyle_double,
                                        String.valueOf(a_tot_ore.get()),
                                        false, true);

                                a_tot_al.addAndGet(a_tot_ore.get() * costo_ora_fasea);
                                componente_allievi_A.addAndGet(a_tot_ore.get() * costo_ora_fasea);
                                setCell(getCell(row_allievo, 12 + calendarioA.size()),
                                        cellStyle_double,
                                        String.valueOf(a_tot_ore.get() * costo_ora_fasea),
                                        false, true);

                                d_tot.addAndGet(tot_docenti.get() / numpartecipanti);
                                componente_docenti_A.addAndGet(tot_docenti.get() / numpartecipanti);
                                setCell(getCell(row_allievo, 13 + calendarioA.size()),
                                        cellStyle_double,
                                        String.valueOf(tot_docenti.get() / numpartecipanti),
                                        false, true);

//                                        System.out.println("Rendicontazione.prospetto_riepilogo() " + indicifaseB.toString());
                                //FASE B
                                for (int ix = 0; ix < elencogruppi_B.size(); ix++) {
                                    String numerogruppo = elencogruppi_B.get(ix);
                                    if (numerogruppo.equals(allievo.getGruppofaseB())) {

                                        indicepartenzafaseBallievo = indicifaseB.get(numerogruppo);

                                        LinkedList<Items> calendarioB2 = calendarioB.stream().filter(l1
                                                -> l1.getFase().endsWith("B") && l1.getGruppo().equals(numerogruppo))
                                                .collect(Collectors.toCollection(LinkedList::new));

                                        for (int x = 0; x < calendarioB2.size(); x++) {
                                            Items it1 = calendarioB2.get(x);

                                            List<Registro_completo> rc_B = faseB.stream().filter(
                                                    a1 -> a1.getIdutente() == allievo.getId() && a1.getData().toString(patternITA).equals(it1.getData()))
                                                    .collect(Collectors.toList());

                                            if (!rc_B.isEmpty()) {
                                                long ore = 0L;

                                                for (Registro_completo rr : rc_B) {
                                                    ore += rr.getTotaleorerendicontabili();
                                                }

                                                setCell(getCell(row_allievo, indicepartenzafaseBallievo + x),
                                                        cellStyle_double,
                                                        String.valueOf(roundDouble(ore, true)),
                                                        false, true);

                                                b_tot_ore.addAndGet(roundDouble(ore, true));
                                            } else {
                                                setCell(getCell(row_allievo, indicepartenzafaseBallievo + x),
                                                        cellStyle_double,
                                                        "0.0",
                                                        false, true);
                                            }

                                        }

                                        setCell(getCell(row_allievo, indicepartenzafaseBallievo + calendarioB2.size()),
                                                cellStyle_double,
                                                String.valueOf(b_tot_ore.get()),
                                                false, true);

                                        b_tot_al.addAndGet(b_tot_ore.get() * costo_ora_faseb);

                                        setCell(getCell(row_allievo, indicepartenzafaseBallievo + calendarioB2.size() + 1),
                                                cellStyle_double,
                                                String.valueOf(b_tot_ore.get() * costo_ora_faseb),
                                                false, true);

                                        if (totaliFaseB.get(numerogruppo) == null) {
                                            totaliFaseB.put(numerogruppo, b_tot_ore.get() * costo_ora_faseb);
                                        } else {
                                            totaliFaseB.put(numerogruppo, totaliFaseB.get(numerogruppo) + (b_tot_ore.get() * costo_ora_faseb));
                                        }

                                    }

                                }

                                //
                                setCell(getCell(row_allievo, iniziofaseb.get()),
                                        cellStyle_double,
                                        String.valueOf(a_tot_ore.get() + b_tot_ore.get()),
                                        false, true);

                                index_allievo.addAndGet(1);

                                //PRIMO FOGLIO
                                XSSFRow row = getRow(sh1, index_row.get());
                                setCell(getCell(row, 0), cellStyle_int, String.valueOf(indice.get()), true, false);
                                setCell(getCell(row, 1), cellStyle_int, String.valueOf(idpr), true, false);
                                setCell(getCell(row, 2), cs, cip, false, false);
                                setCell(getCell(row, 3), cs, ragionesociale, false, false);
                                setCell(getCell(row, 4), cs, allievo.getCognome(), false, false);
                                setCell(getCell(row, 5), cs, allievo.getNome(), false, false);
                                setCell(getCell(row, 6), cs, allievo.getCf(), false, false);
                                setCell(getCell(row, 7), cs, start, false, false);
                                setCell(getCell(row, 8), cs, end, false, false);
                                setCell(getCell(row, 9),
                                        cellStyle_double,
                                        String.valueOf(a_tot_ore.get() + b_tot_ore.get()),
                                        false, true);
                                total_ore.addAndGet(a_tot_ore.get() + b_tot_ore.get());
                                setCell(getCell(row, 10), cellStyle_double, String.valueOf(a_tot_al.get() + b_tot_al.get() + d_tot.get()), false, true);
                                total_rend.addAndGet(a_tot_al.get() + b_tot_al.get() + d_tot.get());
                                total_pr.addAndGet(a_tot_al.get() + b_tot_al.get() + d_tot.get());
                                index_row.addAndGet(1);
                                indice.addAndGet(1);

                            }

                            //RIGA TOTALI CIP
                            XSSFRow row_total = getRow(sh_pr, index_allievo.get());

                            for (int x = 1; x <= iniziofaseb.get(); x++) {
                                if (x == 12 + calendarioA.size()) {
                                    setCell(getCell(row_total, x), cstotal_double, String.valueOf(componente_allievi_A.get()), false, true);
                                } else if (x == 13 + calendarioA.size()) {
                                    setCell(getCell(row_total, x), cstotal_double, String.valueOf(componente_docenti_A.get()), false, true);
                                } else {
                                    setCell(getCell(row_total, x), cstotal, " ", false, false);
                                    for (int ix = 0; ix < elencogruppi_B.size(); ix++) {
                                        String numerogruppo = elencogruppi_B.get(ix);
                                        if (indicifaseB.get(numerogruppo) != null) {
                                            LinkedList<Items> calendarioB2 = calendarioB.stream().filter(l1
                                                    -> l1.getFase().endsWith("B") && l1.getGruppo().equals(numerogruppo))
                                                    .collect(Collectors.toCollection(LinkedList::new));
                                            if (x == indicifaseB.get(numerogruppo) + calendarioB2.size() + 1) {
                                                setCell(getCell(row_total, x), cstotal_double, String.valueOf(totaliFaseB.get(numerogruppo)), false, true);
                                            }
                                        }
                                    }
                                }
                            }

                            CellRangeAddress region_C1 = new CellRangeAddress(2, 3, iniziofaseb.get(), iniziofaseb.get());
                            cleanBeforeMergeOnValidCells(sh_pr, region_C1, cs, true);
                            CellRangeAddress region_C2 = new CellRangeAddress(4, 6, iniziofaseb.get(), iniziofaseb.get());
                            cleanBeforeMergeOnValidCells(sh_pr, region_C2, cs, true);
                            setCell(getCell(row_intest, iniziofaseb.get()), intestazione_2A, "TOTALE ORE", false, false);
                            setCell(getCell(row_intest3, iniziofaseb.get()), cs, "FASE A + FASE B", false, false);
                            XSSFRow row_recap = getRow(sh_pr, index_allievo.get() + 2);
                            setCell(getCell(row_recap, 7), intestazione_5, "TOTALE PARTECIPANTI", false, false);
                            setCell(getCell(row_recap, 8), cstotal_int, String.valueOf(numpartecipanti), true, false);
                            XSSFRow row_recap1 = getRow(sh_pr, index_allievo.get() + 3);
                            setCell(getCell(row_recap1, 7), intestazione_5, "TOTALE IMPORTO CORSO €", false, false);
                            setCell(getCell(row_recap1, 8), cstotal_double, String.valueOf((total_pr.get())), false, true);
                            for (int ix = 1; ix < iniziofaseb.get() + 10; ix++) {
                                sh_pr.autoSizeColumn(ix);
                            }
                        }
                    }
                }

                XSSFRow row_total = getRow(sh1, index_row.get());
                setCell(getCell(row_total, 0), cstotal, "", false, false);
                setCell(getCell(row_total, 1), cstotal, "", false, false);
                setCell(getCell(row_total, 2), cstotal, "", false, false);
                setCell(getCell(row_total, 3), cstotal, "", false, false);
                setCell(getCell(row_total, 4), cstotal, "", false, false);
                setCell(getCell(row_total, 5), cstotal, "", false, false);
                setCell(getCell(row_total, 6), cstotal, "", false, false);
                setCell(getCell(row_total, 7), cstotal, "", false, false);
                setCell(getCell(row_total, 8), cstotal, "", false, false);
                setCell(getCell(row_total, 9), cstotal, "", false, false);
                setCell(getCell(row_total, 10), cstotal_double, String.valueOf(total_rend.get()), false, true);
                for (int i = 0; i < 13; i++) {
                    sh1.autoSizeColumn(i);
                }
                output_xlsx = createFile(pathdest + "/" + nomerend + "_NEW_Riepilogo_" + new DateTime().toString(timestamp) + ".xlsx");
                if (output_xlsx == null) {
                    db1.closeDB();
                    return null;
                }
                try (FileOutputStream outputStream = new FileOutputStream(output_xlsx)) {
                    wb.write(outputStream);
                    log.log(Level.INFO, "FILE RILASCIATO: {0}", output_xlsx.getPath());
                }

            }
            db1.closeDB();
        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }

        return output_xlsx;
    }

    public static void main(String[] args) {
        generaRendicontazione(true);

//        List<Integer> start = new ArrayList<>();
//        start.add(10);
    }
}
