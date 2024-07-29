/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package testerclass;

import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import rc.so.exe.Db_Gest;
import rc.so.exe.Sicilia_gestione;

/**
 *
 * @author Administrator
 */
public class nuovoinvio {
    public static void main(String[] args) {
//        List<Integer> elenco = new ArrayList<>();
        Sicilia_gestione sg = new Sicilia_gestione(false);
//        Db_Gest db1 = new Db_Gest(sg.host);
//        try {
//            String sql0 = "SELECT idprogetti_formativi FROM progetti_formativi "
//                    + "WHERE CURDATE()>=start AND CURDATE()<=end "
//                    + "AND (stato='ATA' OR stato = 'ATB') ORDER BY idprogetti_formativi";
//            try (Statement st0 = db1.getConnection().createStatement(); ResultSet rs0 = st0.executeQuery(sql0)) {
//                while (rs0.next()) {
//                    elenco.add(rs0.getInt("idprogetti_formativi"));
//                }
//            }
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        db1.closeDB();

//        CREARE STANZA PER I PROGETTI CHE HANNO SOLO PRESENZA

//        elenco.forEach(pf -> {
            sg.fad_docenti_PRESENZA(264);
//        });
        
        
    }
}
