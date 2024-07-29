//package rc.so.report;
//
//import com.google.common.util.concurrent.AtomicDouble;
//import static rc.so.exe.Constant.cf_soggetto_DD;
//import static rc.so.exe.Constant.codice_bb;
//import static rc.so.exe.Constant.codice_yisu_neet;
//import static rc.so.exe.Constant.codice_yisu_ded;
//import static rc.so.exe.Constant.coeff_ddr_dd;
//import static rc.so.exe.Constant.coeff_docfasciaA;
//import static rc.so.exe.Constant.coeff_docfasciaB;
//import static rc.so.exe.Constant.coeff_faseA;
//import static rc.so.exe.Constant.coeff_faseB;
//import static rc.so.exe.Constant.contoallievifaseA;
//import static rc.so.exe.Constant.contoallievifaseA_DD;
//import static rc.so.exe.Constant.contoallievifaseB;
//import static rc.so.exe.Constant.contoallievifaseB_DD;
//import static rc.so.exe.Constant.contodocentiA;
//import static rc.so.exe.Constant.contodocentiA_DD;
//import static rc.so.exe.Constant.contodocentiB;
//import static rc.so.exe.Constant.contodocentiB_DD;
//import static rc.so.exe.Constant.getCell;
//import static rc.so.exe.Constant.getRow;
//import static rc.so.exe.Constant.percentuale_attribuzioneDD;
//import static rc.so.exe.Constant.setCell;
//import static rc.so.exe.Constant.tipologia_costo;
//import static rc.so.exe.Constant.tipologia_giustificativo;
//import static rc.so.exe.Constant.zipListFiles;
//import rc.so.exe.Db_Accreditamento;
//import rc.so.exe.Items;
//import java.io.BufferedWriter;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.FileWriter;
//import java.math.BigDecimal;
//import java.math.RoundingMode;
//import java.sql.Connection;
//import java.sql.PreparedStatement;
//import java.sql.ResultSet;
//import java.sql.Statement;
//import java.text.SimpleDateFormat;
//import java.util.ArrayList;
//import java.util.Arrays;
//import java.util.HashMap;
//import java.util.List;
//import java.util.Locale;
//import java.util.Map;
//import java.util.concurrent.atomic.AtomicInteger;
//import java.util.stream.Collectors;
//import org.apache.commons.lang3.StringUtils;
//import org.apache.poi.ss.usermodel.BorderStyle;
//import org.apache.poi.ss.usermodel.FillPatternType;
//import org.apache.poi.ss.usermodel.HorizontalAlignment;
//import org.apache.poi.ss.usermodel.IndexedColors;
//import org.apache.poi.ss.usermodel.VerticalAlignment;
//import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFCellStyle;
//import org.apache.poi.xssf.usermodel.XSSFColor;
//import org.apache.poi.xssf.usermodel.XSSFDataFormat;
//import org.apache.poi.xssf.usermodel.XSSFFont;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.joda.time.DateTime;
//
///*
// * To change this license header, choose License Headers in Project Properties.
// * To change this template file, choose Tools | Templates
// * and open the template in the editor.
// */
///**
// *
// * @author rcosco
// */
//public class Excel {
//
//    private static final String separator = "|";
//    private static final SimpleDateFormat sdfHHmm = new SimpleDateFormat("HH:mm");
//    private static final SimpleDateFormat sdfHHmmss = new SimpleDateFormat("HH:mm:ss");
//    private static final SimpleDateFormat sdfITA = new SimpleDateFormat("dd/MM/yyyy");
//    private static final SimpleDateFormat sdfSQL = new SimpleDateFormat("yyyy-MM-dd");
//    private static final String formatdataCell = "#,#.00";
//    private static final String formatdataCellint = "#,#";
//    private static final byte[] bianco = {(byte) 255, (byte) 255, (byte) 255};
//    private static final byte[] color1 = {(byte) 49, (byte) 134, (byte) 155};
//    private static final byte[] color2 = {(byte) 83, (byte) 141, (byte) 213};
//    private static final byte[] color3 = {(byte) 197, (byte) 217, (byte) 241};
//    private static final byte[] color4 = {(byte) 238, (byte) 30, (byte) 30};
//    private static final byte[] color5 = {(byte) 0, (byte) 204, (byte) 0};
//    private static final XSSFColor myColor1 = new XSSFColor(color1, new DefaultIndexedColorMap());
//    private static final XSSFColor myColor2 = new XSSFColor(color2, new DefaultIndexedColorMap());
//    private static final XSSFColor myColor3 = new XSSFColor(color3, new DefaultIndexedColorMap());
//    private static final XSSFColor myColor4 = new XSSFColor(color4, new DefaultIndexedColorMap());
//    private static final XSSFColor myColor5 = new XSSFColor(color5, new DefaultIndexedColorMap());
//    private static final XSSFColor white = new XSSFColor(bianco, new DefaultIndexedColorMap());
//    private static final Long hh36 = Long.valueOf(129600000);
//
//    private static String get_incremental(int index) {
//
//        try {
//
//            String end = String.valueOf(index);
//            String start = new DateTime().toString("YYMMdd");
//            int mancanti = 10 - end.length() - start.length();
//            String middle = StringUtils.leftPad("", mancanti, "0");
//            return start + middle + end;
//        } catch (Exception e) {
//        }
//        return "000000000";
//    }
//
//    private static String getDoubleforTXT(double ing) {
//        try {
//            BigDecimal bd = new BigDecimal(ing).setScale(2, RoundingMode.HALF_EVEN);
//
//            if (bd.doubleValue() % 1 == 0) {
//                return String.valueOf(bd.intValue());
//            } else {
//                return (String.format(Locale.ITALIAN, "%.2f", bd.doubleValue()));
//            }
//        } catch (Exception e) {
//            return (String.format(Locale.ITALIAN, "%.2f", ing));
//        }
//    }
//
//    private static void cleanBeforeMergeOnValidCells(XSSFSheet sheet, CellRangeAddress region, XSSFCellStyle cellStyle) {
//        try {
//            for (int rowNum = region.getFirstRow(); rowNum <= region.getLastRow(); rowNum++) {
//                XSSFRow row = getRow(sheet, rowNum);
//                for (int colNum = region.getFirstColumn(); colNum <= region.getLastColumn(); colNum++) {
//                    XSSFCell currentCell = getCell(row, colNum);
//                    currentCell.setCellStyle(cellStyle);
//                }
//            }
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//
//    }
//
//    private static List<Items> calendario(int idpr, Connection conn) {
//
//        List<Items> out = new ArrayList<>();
//        List<Items> temp = new ArrayList<>();
//        try {
//
//            String sql = "SELECT lc.lezione,lm.giorno,lm.orario_start,lm.orario_end,ud.fase,lm.gruppo_faseB "
//                    + "FROM lezioni_modelli lm, modelli_progetti mp, lezione_calendario lc, "
//                    + "unita_didattiche ud, fad_multi f, "
//                    + "progetti_formativi p  "
//                    + "WHERE mp.id_modello=lm.id_modelli_progetto AND lc.id_lezionecalendario=lm.id_lezionecalendario "
//                    + "AND ud.codice=lc.codice_ud "
//                    + "AND p.idprogetti_formativi=f.idprogetti_formativi AND f.idprogetti_formativi=mp.id_progettoformativo "
//                    + "AND (lm.gruppo_faseB = 0 OR lm.gruppo_faseB=f.numerocorso) "
//                    + "AND mp.id_progettoformativo=" + idpr
//                    + " GROUP BY lm.gruppo_faseB,lm.giorno,lm.id_lezionecalendario "
//                    + " ORDER BY lm.gruppo_faseB,lm.giorno";
//            try (Statement st = conn.createStatement(); ResultSet rs = st.executeQuery(sql)) {
//                while (rs.next()) {
//
//                    String fase = rs.getString("ud.fase").endsWith("A") ? "A" : "B";
//                    String gruppo = fase.equals("A") ? "1" : rs.getString("lm.gruppo_faseB");
//                    Items itm = new Items(fase, rs.getString("lm.giorno"), rs.getString("lm.orario_start"), rs.getString("lm.orario_end"), gruppo);
//                    temp.add(itm);
//                }
//            }
//        } catch (Exception ex1) {
//            ex1.printStackTrace();
//        }
//
//        for (int i = 0; i < temp.size(); i++) {
//            if (i == temp.size() - 1) {
//                out.add(temp.get(i));
//            } else {
//                if (temp.get(i).getData().equals(temp.get(i + 1).getData())) {
//                    out.add(new Items(temp.get(i).getFase(), temp.get(i).getData(), temp.get(i).getOrainizio(), temp.get(i + 1).getOrafine(), temp.get(i).getGruppo()));
//                    i++;
//                } else {
//                    out.add(temp.get(i));
//                }
//            }
//        }
//
//        return out;
//    }
//
//    private static Map<Long, Long> OreRendicontabiliAlunni_faseA(Connection conn, int pf) {
//        Map result = new HashMap();
//        try {
//
//            String sql = "SELECT sum(totaleorerendicontabili) as totOre,idutente FROM registro_completo WHERE fase = 'A' AND ruolo like 'ALLIEVO%' "
//                    + " AND idutente IN (SELECT a.idallievi FROM allievi a WHERE a.idprogetti_formativi = ? AND a.id_statopartecipazione='01')"
//                    + "AND idprogetti_formativi = ? GROUP BY idutente";
//            PreparedStatement ps = conn.prepareStatement(sql);
//            ps.setInt(1, pf);
//            ps.setInt(2, pf);
//            ResultSet rs = ps.executeQuery();
//            while (rs.next()) {
//                result.put(rs.getLong("idutente"), rs.getLong("totOre"));
//            }
//
//            return result;
//
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        return result;
//    }
//
//    private static Map<Long, Long> OreRendicontabiliDocentiFASEA(Connection conn, int pf) {
//        Map result = new HashMap();
//        try {
//
//            String sql = "SELECT sum(totaleorerendicontabili) as totOre,idutente FROM registro_completo WHERE ruolo = 'DOCENTE' "
//                    + "AND fase='A' AND idprogetti_formativi = ? GROUP BY idutente";
//            PreparedStatement ps = conn.prepareStatement(sql);
//            ps.setInt(1, pf);
//            ResultSet rs = ps.executeQuery();
//            while (rs.next()) {
//                result.put(rs.getLong("idutente"), rs.getLong("totOre"));
//            }
//
//        } catch (Exception e) {
//            e.printStackTrace();
//
//        }
//        return result;
//    }
//
//    private static double roundFloatAndFormat(float f) {
//        try {
//
//            double hours = f / 1000.0 / 60.0 / 60.0;
//            BigDecimal bigDecimal = new BigDecimal(hours);
//            bigDecimal = bigDecimal.setScale(2, RoundingMode.HALF_EVEN);
//            return bigDecimal.doubleValue();
//
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        return 0.0;
//
//    }
//
//}
