/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package testerclass;

import java.sql.ResultSet;
import java.sql.Statement;
import static rc.so.exe.Constant.estraiEccezione;
import rc.so.exe.Db_Gest;
import rc.so.exe.Sicilia_gestione;
import static rc.so.exe.Sicilia_gestione.log;
import rc.so.exe.Utils;

/**
 *
 * @author Administrator
 */
public class verificaconvalidate {

    public static void main(String[] args) {

        Sicilia_gestione tg = new Sicilia_gestione(false);
        Db_Gest db = new Db_Gest(tg.host);
        String sql0 = "SELECT l.idpresenzelezioniallievi,l.durata,l.orainizio,l.orafine,l.durataconvalidata FROM presenzelezioniallievi l WHERE l.convalidata=1 AND l.durata > 0";

        try (Statement st0 = db.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0)) {
            while (rs0.next()) {
                long idpresenzelezioniallievi = rs0.getLong("l.idpresenzelezioniallievi");
                String orainizio = rs0.getString("l.orainizio");
                String orafine = rs0.getString("l.orafine");
                long durata = rs0.getLong("l.durata");
                long durataconvalidata = rs0.getLong("l.durataconvalidata");

                long check = Utils.calcolaintervallomillis(orainizio, orafine);
                if (durata != check) {
                    String upd1 = "UPDATE presenzelezioniallievi SET durata='" + check + "' WHERE idpresenzelezioniallievi=" + idpresenzelezioniallievi;
                    try (Statement st1 = db.getConnection().createStatement()) {
                        st1.executeUpdate(upd1);
                    }
                    System.out.println("A>"+idpresenzelezioniallievi + " -- (" + orainizio + "->" + orafine + ") " + durata + " () " + check + " ** " + durataconvalidata);
                } else if (durataconvalidata > check) {
                    String upd2 = "UPDATE presenzelezioniallievi SET durataconvalidata='" + check + "' WHERE idpresenzelezioniallievi=" + idpresenzelezioniallievi;
                    try (Statement st1 = db.getConnection().createStatement()) {
                        st1.executeUpdate(upd2);
                    }
                    System.out.println("B>"+idpresenzelezioniallievi + " -- (" + orainizio + "->" + orafine + ") " + durata + " () " + check + " ** " + durataconvalidata);

                }
            }
        } catch (Exception ex1) {
            log.severe(estraiEccezione(ex1));
        }

        db.closeDB();

    }
}
