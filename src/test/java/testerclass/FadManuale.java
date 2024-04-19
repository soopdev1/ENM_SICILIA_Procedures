package testerclass;




import rc.so.exe.Sicilia_gestione;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author rcosco
 */
public class FadManuale {

    public static void main(String[] args) {

        int pf = 49;
        
//        DeD_gestione d = new DeD_gestione(false);
////        
////        d.verifica_stanze(pf);
//        d.fad_allievi(pf, true);
//        d.fad_docenti(pf, true);
//        d.fad_ospiti(pf, true);
        
        
        Sicilia_gestione n = new Sicilia_gestione(false);
        n.verifica_stanze(pf);
        n.fad_allievi(pf, true);
        n.fad_docenti(pf, true);
        n.fad_ospiti(pf, true);
        
    }
}
