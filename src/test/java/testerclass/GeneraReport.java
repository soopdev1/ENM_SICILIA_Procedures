package testerclass;

import rc.so.report.FaseA;
import rc.so.report.Lezione;
import java.util.List;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author rcosco
 */
public class GeneraReport {

    public static void main(String[] args) {
        try {

            boolean testing = false;
            int idpr = 73;

            FaseA FA = new FaseA(testing);
////            FaseB FB = new FaseB(testing, neet);

//        //  FASE A
            List<Lezione> ca = FA.calcolaegeneraregistrofasea(idpr, FA.getHost(), false, true, false);
            FA.registro_aula_FaseA(idpr, FA.getHost(), false, false);

//          //  FASE B
//            List<Lezione> cb = FB.calcolaegeneraregistrofaseb(idpr, FA.getHost(), false, true, false);
//            FB.registro_aula_FaseB(idpr, FA.getHost(), false, cb, neet);
//
////        //  COMPLESSIVO
//            new Complessivo(FA.getHost()).registro_complessivo(idpr, FA.getHost(), ca, cb, false, neet);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
