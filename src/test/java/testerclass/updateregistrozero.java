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
public class updateregistrozero {

    public static void main(String[] args) {
        Sicilia_gestione n = new Sicilia_gestione(false);

        Db_Gest db = new Db_Gest(n.host);

        String sql1 = "SELECT r.idallievi,r.cognome,r.nome,r.idprogetti_formativi FROM allievi r "
                + "WHERE r.idprogetti_formativi IS NOT NULL AND r.id_statopartecipazione ='15' ORDER BY r.cognome,r.nome";

        List<Datiallievo> la = new ArrayList<>();

        try (Statement st1 = db.getConnection().createStatement(); ResultSet rs1 = st1.executeQuery(sql1)) {

            while (rs1.next()) {
                String idallievi = rs1.getString(1);
                String cognome = rs1.getString(2);
                String nome = rs1.getString(3);
                String idprogetti_formativi = rs1.getString(4);

                la.add(new Datiallievo(idallievi, cognome, nome, idprogetti_formativi));

            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        String sql2 = "SELECT r.id,r.cognome,r.nome,r.idprogetti_formativi,r.idutente FROM registro_completo r WHERE r.ruolo='ALLIEVO' AND r.idutente=0 ORDER BY r.cognome,r.nome";

        try (Statement st2 = db.getConnection().createStatement(); ResultSet rs2 = st2.executeQuery(sql2)) {

            while (rs2.next()) {
                String idregistro = rs2.getString(1);
                String cognome = rs2.getString(2);
                String nome = rs2.getString(3);

                Datiallievo da = la.stream().filter(d1 -> d1.getCognome().equalsIgnoreCase(cognome) && d1.getNome().equals(nome)).findAny().orElse(null);
                if (da != null) {
                    String upd = "UPDATE registro_completo SET idutente = " + da.getIdallievi() + " WHERE id = " + idregistro;
                    try (Statement st3 = db.getConnection().createStatement()) {
                        boolean es1 = st3.executeUpdate(upd) > 0;
                        System.out.println(es1 + ") " + upd);
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        db.closeDB();

    }
}

class Datiallievo {

    String idallievi, cognome, nome, idprogetti_formativi;

    public Datiallievo(String idallievi, String cognome, String nome, String idprogetti_formativi) {
        this.idallievi = idallievi;
        this.cognome = cognome;
        this.nome = nome;
        this.idprogetti_formativi = idprogetti_formativi;
    }

    public String getIdallievi() {
        return idallievi;
    }

    public void setIdallievi(String idallievi) {
        this.idallievi = idallievi;
    }

    public String getCognome() {
        return cognome;
    }

    public void setCognome(String cognome) {
        this.cognome = cognome;
    }

    public String getNome() {
        return nome;
    }

    public void setNome(String nome) {
        this.nome = nome;
    }

    public String getIdprogetti_formativi() {
        return idprogetti_formativi;
    }

    public void setIdprogetti_formativi(String idprogetti_formativi) {
        this.idprogetti_formativi = idprogetti_formativi;
    }

}
