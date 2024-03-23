/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package rc.so.exe;

import static rc.so.exe.Constant.estraiEccezione;
//import it.refill.otp.Sms;
//import rc.so.report.Create;
import java.util.logging.Logger;
import static rc.so.report.Create.crearegistri;

/**
 *
 * @author Administrator
 */
public class MainSelector {

    public static final Logger log = Constant.createLog("ENM_SICILIA_Procedura", "/mnt/mcn/test/log/");

    public static void main(String[] args) {

        try {
            System.setProperty("org.apache.commons.logging.Log", "org.apache.commons.logging.impl.NoOpLog");
            java.util.logging.Logger.getLogger("org.apache.pdfbox").setLevel(java.util.logging.Level.SEVERE);
        } catch (Exception e) {
        }

        //ARGS[0] TEST
        //////////////////////////////
        //ARGS[1] OPERAZIONE 1 - 3
        //1 - GESTIONE - COMUNICAZIONI
        //2 - GESTIONE - FAD
        //3 - GESTIONE - ESTRAZIONI
        //4 - GESTIONE - REPORT FAD
        //5 - ACCREDITAMENTO
        //6 - REPAIR
        //7 - RENDICONTAZIONE
        //////////////////////////////
        //ARGS[2] NEET - DD
        //////////////////////////////
        boolean testing;
        try {
            testing = args[0].trim().equals("test");
        } catch (Exception e) {
            testing = false;
        }

        int select_action;
        try {
            select_action = Integer.parseInt(args[1].trim());
        } catch (Exception e) {
            select_action = 4;
        }

        Sicilia_gestione sg = new Sicilia_gestione(testing);
        switch (select_action) {
            case 2 -> {
                log.warning("GESTIONE SICILIA - FAD");
                try {
                    sg.fad_gestione();
                } catch (Exception e) {
                }
                break;
            }
            case 3 -> {
                log.warning("GESTIONE SICILIA - ESTRAZIONI");
                log.warning("GENERAZIONE FILE REPORT... INIZIO");
                try {
                        sg.report_docenti();
                } catch (Exception e) {
                }
                try {
                    sg.report_allievi();
                } catch (Exception e) {

                }
                try {
                    sg.report_pf();
                } catch (Exception e) {

                }
                log.warning("GENERAZIONE FILE REPORT... FINE");
            }
            case 4 -> {
                log.warning("GESTIONE SICILIA - REPORT FAD");
                crearegistri(testing);
                log.warning("GESTIONE SICILIA - UPDATE ORE CONVALIDATE");
                sg.ore_convalidateAllievi();
                break;
            }
            case 5 -> {
                log.warning("ACCREDITAMENTO SICILIA");
                Engine accreditamento = new Engine(testing);

                try {
                    log.info("START ELENCO DOMANDE");
                    accreditamento.elenco_domande_fase1();
                    log.info("FINE ELENCO DOMANDE");
                } catch (Exception e) {
                    log.severe(estraiEccezione(e));
                }

                try {
                    log.info("START UPDATE DOMANDE");
                    accreditamento.update_domande_fase1();
                    log.info("FINE UPDATE DOMANDE");
                } catch (Exception e) {
                    log.severe(estraiEccezione(e));
                }

                try {
                    log.info("START AGGIORNA DATA CONVENZIONE");
                    accreditamento.aggiorna_dataconvenzione_fase1();
                    log.info("FINE AGGIORNA DATA CONVENZIONE");
                } catch (Exception e) {
                    log.severe(estraiEccezione(e));
                }

                try {
                    log.info("START AGGIORNA REPORTISTICA");
                    accreditamento.aggiorna_reportistica();
                    log.info("FINE AGGIORNA REPORTISTICA");
                } catch (Exception e) {
                    log.severe(estraiEccezione(e));
                }
                try {
                    log.info("START CREA REPORT");
                    accreditamento.crea_report();
                    log.info("FINE CREA REPORT");
                } catch (Exception e) {
                    log.severe(estraiEccezione(e));
                }
            }
            default ->
                log.severe("GESTIONE SICILIA - NESSUN METODO SELEZIONATO");

        }
//                case 0:
//                    log.severe("GESTIONE NEET - NESSUN METODO SELEZIONATO");
//                    break;
//                case 1:
//                    log.warning("GESTIONE NEET - COMUNICAZIONI");
//                    try {
//                        log.warning("MAIL REMIND 1 GIORNO... INIZIO");
//                        ne.mail_remind(1);
//                        log.warning("MAIL REMIND 1 GIORNO... FINE");
//                    } catch (Exception e) {
//                    }
//                    try {
//                        log.warning("MAIL REMIND QUESTIONARI INGRESSO... INIZIO");
//                        ne.mail_questionario_INGRESSO();
//                        log.warning("MAIL REMIND QUESTIONARI INGRESSO... FINE");
//                    } catch (Exception e) {
//                    }
//                    try {
//                        log.warning("MAIL REMIND QUESTIONARI USCITA... INIZIO");
//                        ne.mail_questionario_USCITA();
//                        log.warning("MAIL REMIND QUESTIONARI USCITA... FINE");
//                    } catch (Exception e) {
//                    }
//                    break;
//                case 2:
//                    log.warning("GESTIONE NEET - FAD");
//                    try {
//                        ne.fad_gestione();
//                    } catch (Exception e) {
//                    }
//                    break;
//                case 4:
//                    log.warning("GESTIONE NEET - REPORT FAD");
//                    Create.crea(true, testing);
//                    break;
//                case 6:
//                    log.warning("REPAIR NEET");
//                    Repair neetr = new Repair(testing, true);
//
//                    try {
//                        neetr.imposta_progetti_finettivita();
//                    } catch (Exception e) {
//                    }
//
//                    try {
//                        neetr.impostaritiratounder36oreA();
//                    } catch (Exception e) {
//
//                    }
//
//                    try {
//                        neetr.copiadocumentidocenti();
//                    } catch (Exception e) {
//
//                    }
//                    try {
//                        neetr.crea_pdf_unico_ANPAL(true);
//                    } catch (Exception e) {
//
//                    }
//                    break;
//                case 7: 
//                    try {
//                    log.info("START RENDICONTAZIONE NEET");
//                    new Rendicontazione(false, true).generaRendicontazione(true);
//                    log.info("END RENDICONTAZIONE NEET");
//                } catch (Exception e) {
//                }
//                break;
    }
}
