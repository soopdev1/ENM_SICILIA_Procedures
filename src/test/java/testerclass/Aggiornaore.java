/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package testerclass;

import rc.so.exe.Sicilia_gestione;

/**
 *
 * @author Administrator
 */
public class Aggiornaore {

    public static void main(String[] args) {
        String sql = "SELECT a.idallievi FROM allievi a WHERE a.idprogetti_formativi IN (434)";
        Sicilia_gestione sg = new Sicilia_gestione(false);
        sg.ore_convalidateAllievi(sql);
        sg.ore_ud(sql);
//        sg.ritira_allievi_zero();
    }
}
