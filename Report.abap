***********************************************************************
*
*  ESA Zusatzentwicklung
*
*  Entwicklungs-ID    : D...
*  Beschreibung       : Dieses Report soll die Adressdaten der Kunden beschaffen und Sie sortiert ausgeben
*  Jira Ticket        : RICEFF-183
*  Funktion           : Beschreibung der Funktionen im Einzelnen
*                       inkl. der Integration in die SAP Logik
*  Autor              : Rippstein
*  Datum              : 03.06.2019
*  CR                 : Referenz auf den CR, falls relevant
*  Verantwortlicher   : Mihail, Thiel
*  NetWeaver Release  : <RELEASE>
*
***********************************************************************

REPORT z_abap_test.
*Desktop als Default Ordner setzen
"PERFORM set_default_directory.
"INCLUDE zmm_mig_material_con_top.
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
***********************************************************************
* Globale Daten Definierung
**********************************************************************
TYPES: BEGIN OF ty_mara,
         matnr TYPE matnr,
         meins TYPE meins,
         ean11 TYPE ean11,
         numtp TYPE numtp,
       END OF ty_mara.

TYPES: BEGIN OF ty_mvke,
         matnr TYPE matnr,
         vkorg TYPE vkorg,
         vtweg TYPE vtweg,
         vrkme TYPE vrkme,
         scmng TYPE scmng,
       END OF ty_mvke.

TYPES: BEGIN OF ty_file,
         matnr  TYPE matnr,
         kwmeng TYPE kwmeng,
       END OF ty_file.

TYPES: BEGIN OF ty_errors,
          status TYPE zsd_status,
          kunnr TYPE kunnr,
          zeile TYPE SYST_TABIX,
      END OF ty_errors.

* Globale interne Tabellen
*--------------------------------------------------------------------*
DATA: gt_mara TYPE TABLE OF ty_mara.
DATA: gt_makt TYPE TABLE OF makt.

DATA: gt_mvke TYPE TABLE OF ty_mvke.

DATA: gt_output TYPE TABLE OF zmm_mig_material_conv_output.

DATA: gt_message TYPE TABLE OF bapiret2.

DATA: gt_file TYPE TABLE OF ty_file.

DATA: gt_errors TYPE TABLE OF ty_errors.

DATA: lt_errors TYPE ty_errors.
DATA: ld_errors TYPE ty_errors.
* Globale Objekte
*--------------------------------------------------------------------*
DATA: go_excel  TYPE REF TO zcl_excel,
      go_reader TYPE REF TO zif_excel_reader.

DATA: go_worksheet TYPE REF TO zcl_excel_worksheet.
* Globale Strukturen und Daten Referenz
*--------------------------------------------------------------------*

* Globale Variablen
*--------------------------------------------------------------------*
DATA: gv_finished TYPE abap_bool.

DATA: gv_etext_update TYPE abap_bool.

DATA: gv_file TYPE string.
* Feldsymbolen
*--------------------------------------------------------------------*
*Subroutin Ablauf
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-



TABLES: kna1, knvv, zbwbpcust_01.

TYPES: BEGIN OF ty_file1,
         kunnr  TYPE kunnr,
         name1 TYPE name1_gp,
         pstlz TYPE pstlz,
         ort01 TYPE ort01_gp,
       END OF ty_file1.

TYPES: BEGIN OF ls_kunnr,
         kunnr  TYPE kunnr,
       END OF ls_kunnr.

DATA: gt_file1 TYPE TABLE OF ty_file1.

TYPES: BEGIN OF ty_output,
         kunnr  TYPE kunnr,
       END OF ty_output.

DATA: ls_output_data TYPE  ty_output,
      gt_output_data TYPE TABLE OF ty_output,
      ls_output_file TYPE  ty_output,
      gt_output_file TYPE TABLE OF ty_output,
      ls_output_carro TYPE  ty_output,
      gt_output_carro TYPE TABLE OF ty_output,
      ls_output_mitinh TYPE  ty_output,
      gt_output_mitinh TYPE TABLE OF ty_output.

DATA: fm_name           TYPE rs38l_fnam,      " CHAR 30 0 Name of Function Module
      lt_kunden_adressen TYPE zsd_kunde_adresse_tab,
      ld_kunden_adressen TYPE zsd_kunde_adresse,
      fp_docparams      TYPE sfpdocparams,    " Structure  SFPDOCPARAMS Short Description  Form Parameters for Form Processing
      fp_outputparams   TYPE sfpoutputparams. " Structure  SFPOUTPUTPARAMS Short Description  Form Processing Output Parameter

*Sortierung...
DATA order_by  TYPE string.

******************************************************************************************************************************
* Selektion Screen
******************************************************************************************************************************
SELECTION-SCREEN BEGIN OF BLOCK bg_005 WITH FRAME TITLE TEXT-500.
PARAMETERS: p_filx    AS CHECKBOX USER-COMMAND act_file DEFAULT ''. "FILE
"IF p_filx EQ 'X'.
  PARAMETERS: p_file1 RADIOBUTTON GROUP file .
  PARAMETERS: p_file2 RADIOBUTTON GROUP file .
  PARAMETERS: p_file TYPE string LOWER CASE.

"ELSE.
  "PARAMETERS: p_file TYPE string LOWER CASE NO-DISPLAY.
"ENDIF.
"PARAMETERS: p_file TYPE string LOWER CASE.
PARAMETERS: p_carro    AS CHECKBOX DEFAULT 'X'.  "Carro
PARAMETERS: p_mitinh    AS CHECKBOX DEFAULT 'X'.  "Mitinhaber
*PARAMETERS: p_file TYPE string OBLIGATORY LOWER CASE.
SELECTION-SCREEN END OF BLOCK bg_005.

SELECTION-SCREEN BEGIN OF BLOCK bg_001 WITH FRAME TITLE TEXT-100.

SELECT-OPTIONS: zkdgrp FOR knvv-kdgrp.   "Kunden Gruppe (alt ESA, Dista usw.)
SELECT-OPTIONS: zlifsd FOR kna1-lifsd.   "Aktiv/inaktiv
SELECT-OPTIONS: zkatr2 FOR kna1-katr2.
SELECT-OPTIONS: zspras FOR kna1-spras.
SELECT-OPTIONS: zvkbur FOR knvv-vkbur.   "SALESOFFICE
SELECT-OPTIONS: zkunnr FOR kna1-kunnr.
SELECT-OPTIONS: zpstlz FOR kna1-pstlz.
SELECT-OPTIONS: zvtweg FOR knvv-vtweg.  "Vertriebsweg
SELECT-OPTIONS: zkatr1 FOR kna1-katr1.
SELECT-OPTIONS: zkdkg1 FOR kna1-kdkg1.
SELECT-OPTIONS: zkvgr1 FOR knvv-kvgr1.
SELECTION-SCREEN END OF BLOCK bg_001.

SELECTION-SCREEN BEGIN OF BLOCK bg_002 WITH FRAME TITLE TEXT-200.
PARAMETERS : p_sort1 RADIOBUTTON GROUP rb1.
PARAMETERS : p_sort2 RADIOBUTTON GROUP rb1.
SELECTION-SCREEN END OF BLOCK bg_002.

SELECTION-SCREEN BEGIN OF BLOCK bg_006 WITH FRAME TITLE TEXT-600.
PARAMETERS: p_list AS CHECKBOX  DEFAULT 'X',
            p_stat AS CHECKBOX  DEFAULT 'X',
            p_form AS CHECKBOX  DEFAULT 'X'.
*            p_downl AS CHECKBOX,
*            p_excel AS CHECKBOX.

SELECTION-SCREEN END OF BLOCK bg_006.


******************************************************************************************************************************
* Initialisation der Screen-Variablen und allen anderen Variablen wenn nötig.
******************************************************************************************************************************
INITIALIZATION.
* Default = ESA und Grossisten.
zkdgrp-low = '10'.
zkdgrp-high = '20'.
zkdgrp-sign = 'I'.
zkdgrp-option =  'EQ'.
APPEND zkdgrp .

* Default damit keine Inaktiven verarbeitet werden.
zlifsd-low = '01'.
zlifsd-sign ='I'.
zlifsd-option = 'NE'.
APPEND zlifsd.

* Default will keine Werbung
zkatr2-low = '3'.
zkatr2-sign ='I'.
zkatr2-option = 'NE'.
APPEND zkatr2.

******************************************************************************************************************************
* Loop über Feld, um Sichbar oder nicht Sichtbar.
******************************************************************************************************************************
AT SELECTION-SCREEN OUTPUT.

  LOOP AT SCREEN.
    IF p_filx EQ abap_true.
      IF screen-name CS 'P_FILE'.
         screen-active = 1.
      ENDIF.
    ELSE.
      IF screen-name CS 'P_FILE'.
         screen-active = 0.
      ENDIF.
    ENDIF.
    MODIFY SCREEN.
  ENDLOOP.

*LOAD-OF-PROGRAM.


*p_file = 'C:\Users\philip.rippstein\Desktop\Expresskunden01.xlsx'.

******************************************************************************************************************************
* Alles um das File einzulesen.
******************************************************************************************************************************
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f4_file.

AT SELECTION-SCREEN ON p_file.
*  IF p_filx = 'X' AND sy-ucomm = 'ACT_FILE' AND p_file NE space.
  IF p_filx = 'X' AND p_file NE space.
    PERFORM check_file.
  ENDIF.


"Hier abfragen IF Kunde aktiv usw... Dann abfüllen


******************************************************************************************************************************
* Main Programm
******************************************************************************************************************************
START-OF-SELECTION.
************************************************************************
* Alle Tabellen Initialisieeren
************************************************************************
  CLEAR: gt_output_file, gt_output_carro,gt_output_mitinh.

************************************************************************
* Einlesen des Files in die Temp Tabelle gt_output_filed
************************************************************************
  IF p_filx = 'X' AND p_file NE space.
    PERFORM upload_excel_data.
  ENDIF.
************************************************************************
* Ermitteln der Carro (Klassifikation) gt_output_carro
************************************************************************
  IF p_carro = 'X'.
    PERFORM get_carro.
  ENDIF.
************************************************************************
* Ermitteln der Mitinhaber gt_output_mitinh
************************************************************************
  IF p_mitinh = 'X'.
    PERFORM get_mitinhaber.
  ENDIF.
************************************************************************
* Verdichten der Kundennumern ( Jeder Kunde nur einmal )
* dann sind alle Nummer in der Tabelle gt_output_data
************************************************************************
IF p_file1 ='X' AND p_file NE space.
LOOP AT gt_output_file INTO ls_output_file.
  COLLECT ls_output_file INTO gt_output_data.
ENDLOOP.

ELSEIF p_file2 ='X' AND p_file NE space.

LOOP AT gt_output_file INTO ls_output_file.
  COLLECT ls_output_file INTO gt_output_data.
ENDLOOP.

LOOP AT gt_output_carro INTO ls_output_carro.
  COLLECT ls_output_carro INTO gt_output_data.
ENDLOOP.

LOOP AT gt_output_mitinh INTO ls_output_mitinh.
  COLLECT ls_output_mitinh INTO gt_output_data.
ENDLOOP.
ENDIF.


    "FEHLER Excel nicht lesbar oder keine Daten.
***    MESSAGE e022(zsd_ig).

************************************************************************
*  Sortieren in order_by abgelegt (kann wieder gemacht werden !?)
************************************************************************
*  PERFORM daten_sortieren.
************************************************************************
* Die eigentlichen Adressen beschaffen inkl. Sortierung
************************************************************************
  PERFORM read_data.
************************************************************************
* Formular mit C4 ausgeben
************************************************************************
  IF p_form NE space.
    PERFORM show_formular.
  ENDIF.
************************************************************************
  IF p_stat NE space.
    PERFORM get_statistik.
  ENDIF.
************************************************************************
END-OF-SELECTION.
************************************************************************
* Excel Output erstellen.
************************************************************************
*  IF p_excel NE space.
*    PERFORM excel_datei_erstellen.
*  ENDIF.
************************************************************************
* Zum Schluss evt. Kontroll-Liste erstellen.
************************************************************************
  IF p_list NE space.
    PERFORM liste_erstellen.
  ENDIF.
************************************************************************
* Schluss vom Programm
************************************************************************


*&---------------------------------------------------------------------*
*& Form F4_FILE
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM f4_file .

DATA: lt_filetable TYPE filetable.
DATA: lv_rc TYPE i.

CALL METHOD cl_gui_frontend_services=>file_open_dialog
*  EXPORTING
*    WINDOW_TITLE            =
*    DEFAULT_EXTENSION       =
*    DEFAULT_FILENAME        =
*    FILE_FILTER             =
*    WITH_ENCODING           =
*    INITIAL_DIRECTORY       =
*    MULTISELECTION          =
  CHANGING
    file_table              = lt_filetable
    rc                      = lv_rc
*    USER_ACTION             =
*    FILE_ENCODING           =
  EXCEPTIONS
    file_open_dialog_failed = 1
    cntl_error              = 2
    error_no_gui            = 3
    not_supported_by_gui    = 4
    OTHERS                  = 5
        .
IF sy-subrc <> 0.
* Implement suitable error handling here
ENDIF.

IF lv_rc EQ 1.
  READ TABLE lt_filetable INTO p_file INDEX 1.
ELSE.
  MESSAGE e052(zsd_ig).
ENDIF.

ENDFORM.
*********************************************************************************************************************************
* File name muss geprüft werden ob es überhaupt exisitert
*********************************************************************************************************************************
FORM check_file .
*Lokale Variablen
"  CHECK p_filx EQ abap_true AND sy-ucomm NE 'ACT_FILE' AND p_file IS INITIAL.

  DATA: lv_result TYPE abap_bool.
*
  DATA: lv_file(1024) TYPE c.
  DATA: lv_extension(10) TYPE c.
  DATA: lv_file_temp TYPE string.

  lv_file_temp = p_file.

  CALL METHOD cl_gui_frontend_services=>file_exist
    EXPORTING
      file                 = lv_file_temp
    RECEIVING
      result               = lv_result
    EXCEPTIONS
      cntl_error           = 1
      error_no_gui         = 2
      wrong_parameter      = 3
      not_supported_by_gui = 4
      OTHERS               = 5.
  IF sy-subrc <> 0.
* Implement suitable error handling here
   MESSAGE e040(zsd_ig).
    EXIT.
  ENDIF.

  IF lv_result NE abap_true.
    MESSAGE e021(zsd_ig).
*   Datei ist nicht vorhanden.
    RETURN.
  ENDIF.

  lv_file = p_file.

  CALL FUNCTION 'TRINT_FILE_GET_EXTENSION'
    EXPORTING
      filename  = lv_file
*     UPPERCASE = 'X'
    IMPORTING
      extension = lv_extension.
  IF p_filx = 'X'.
  CHECK lv_extension NE 'XLSX'.

  MESSAGE e021(zsd_ig).

  ENDIF.
* Datei ist kein XLSX.Bitte prüfen Sie die ausgewählte Datei.
ENDFORM.



FORM upload_excel_data .
*Lokale Objekte
  DATA: lx_error TYPE REF TO cx_root,
        lx_excel TYPE REF TO zcx_excel.

  DATA: lv_extension TYPE string.
  DATA: ls_output TYPE zmm_mig_material_conv_output.
  DATA: ls_message TYPE bapiret2.


*gt_errors

  TRY.

      FIND REGEX '(\.xlsx|\.xlsm\.XLSX)\s*$' IN p_file SUBMATCHES lv_extension.
      TRANSLATE lv_extension TO UPPER CASE.

      CASE lv_extension.

        WHEN zsdig_extension_xlsx_2.
          CREATE OBJECT go_reader TYPE zcl_excel_reader_2007.
          go_excel = go_reader->load_file(  p_file ).
          "Use template for charts
          go_excel->use_template = abap_true.

        WHEN zsdig_extension_xlsm_2.
          CREATE OBJECT go_reader TYPE zcl_excel_reader_xlsm.
          go_excel = go_reader->load_file(  p_file ).
          "Use template for charts
          go_excel->use_template = abap_true.

        WHEN OTHERS.
          MESSAGE i021(zsd_ig).
           "Datei ist kein XLSX.Bitte prüfen Sie die ausgewählte Datei.
          RETURN.

      ENDCASE.

      go_worksheet = go_excel->get_active_worksheet( ).

      TRY.
          go_worksheet->get_table( EXPORTING iv_skipped_rows = 1
                                   IMPORTING et_table        = gt_file1 ).

*         IF gt_file1-kunnr is INITIAL.
*         IF gt_file1->kunnr IS INITIAL.


*          READ TABLE gt_file1 INTO gw_file1 INDEX 1.

*                       MESSAGE i021(zsd_ig).
*                       " Datei ist kein XLSX.Bitte prüfen Sie die ausgewählte Datei.
*                       RETURN.


          LOOP AT gt_file1 REFERENCE INTO DATA(lds_file1).

            DATA(lv_tabix) = sy-tabix + 1.

            DATA: lv_string TYPE string.

*            ls_output-row = lv_tabix.

            CLEAR: ls_output.

            ls_output-status = icon_led_inactive.
*           Wieso ?
*            IF lds_file1->kunnr IS INITIAL OR
            "IF lds_file1->kunnr CA '[A-Z]' OR lds_file1->kunnr EQ ''.
            IF lds_file1->kunnr CA SY-ABCDE OR lds_file1->kunnr EQ ''OR lds_file1->kunnr CA '[abcdefghijklmnopqrstuvwxyzäöü]'.
              "Dann was ???
              "BEI FEHLER IN ZEILE IGNORIEREN UND WEITERMACHEN
              "ls_output-status = icon_led_red.
              MESSAGE s040(zsd_ig).

*              set SCREEN '0'.
*              LEAVE SCREEN.

*              gt_errors-status = icon_led_red.
*              gt_errors-kunnr = lds_file1->kunnr.


               lt_errors-status = icon_led_red.
               lt_errors-kunnr = lds_file1->kunnr.
               lt_errors-zeile = sy-tabix + 1.

               "Folgende konnten nicht gelesen werden ->gt_errors
               APPEND lt_errors TO gt_errors.
            ELSE.
              "Error Catch cx_root
              CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
               EXPORTING
                   input         = lds_file1->kunnr
                   IMPORTING
                   output        = lds_file1->kunnr.

                "WERTE IN NEUE ZEILE SCHREIBEN
                ls_output_file-kunnr = lds_file1->kunnr.
*                 APPEND
*                 NEW-LINE
                "APPEND ls_output_file TO gt_output_file.
            ENDIF.
            "Only append in ELSE FALL???
            APPEND ls_output_file TO gt_output_file.


          ENDLOOP.
        CATCH zcx_excel INTO lx_excel.
          MESSAGE s022(zsd_ig).
*         Datei konnte nicht geöffnet werden.
          RETURN.
      ENDTRY.

    CATCH cx_root INTO lx_error.
      MESSAGE s022(zsd_ig).
*         Datei konnte nicht geöffnet werden.
      RETURN.
  ENDTRY.

ENDFORM.

FORM read_data.
  CLEAR lt_kunden_adressen.

 SELECT language AS spras, country AS land1, postalcode AS pstlz, businesspartner AS kunnr, addressnumber AS adrnr
   FROM zbw_bp_cds_b_bpcust_01
   FOR ALL ENTRIES IN @gt_output_data
   WHERE businesspartner = @gt_output_data-kunnr       "Geschäftspartner
    AND businesspartner IN @zkunnr          "Geschäftspartner
    AND distributionchannel IN @zvtweg      "Vertriebsweg (Alt ESA Dista usw.)
    AND customergroup   IN @zkdgrp          "Kundengruppe
    AND lifsd           IN @zlifsd          "Aktiv/inaktiv
    AND katr2           IN @zkatr2          "Werbung oder Mailing Code
    AND katr1           IN @zkatr1          "Vignetten
    AND language        IN @zspras          "Sprache
    AND salesoffice     IN @zvkbur          "Verkaufsbüro
    AND postalcode      IN @zpstlz          "PLZ
    AND language        IN @zspras          "Sprache
    AND kdkg1           IN @zkdkg1          "Garantie Kulanz
    AND kvgr1           IN @zkvgr1          "Kundenart ... Garage, Carroserie usw.
    AND sperr EQ ' '             " (Ohne gelöschte)
   INTO TABLE @lt_kunden_adressen.

   SORT lt_kunden_adressen BY kunnr .

IF p_sort1 = 'X'.
  SORT lt_kunden_adressen BY kunnr.
ELSEIF p_sort2 = 'X'.
  SORT lt_kunden_adressen BY pstlz.
ELSE.
  SORT lt_kunden_adressen BY kunnr.
ENDIF.


ENDFORM.

FORM show_formular.
*********************************************************************************************************************************
* Sets the output parameters and opens the spool job
*********************************************************************************************************************************
CALL FUNCTION 'FP_JOB_OPEN'                   "& Form Processing: Call Form
  CHANGING
    ie_outputparams = fp_outputparams

  EXCEPTIONS
    cancel          = 1
    usage_error     = 2
    system_error    = 3
    internal_error  = 4
    OTHERS          = 5.
IF sy-subrc <> 0.
*            <error handling>
ENDIF.
*&---- Get the name of the generated function module
CALL FUNCTION 'FP_FUNCTION_MODULE_NAME'           "& Form Processing Generation
  EXPORTING
    i_name     = 'ZSD_FORMS_ETTIKETEN'
  IMPORTING
    e_funcname = fm_name.
IF sy-subrc <> 0.
*  <error handling>
ENDIF.
*-- Fetch the Data and store it in the Internal Table
*SELECT * FROM mari INTO TABLE it_mari UP TO 15 ROWS.
** Language and country setting (here US as an example)
*fp_docparams-langu   = 'E'.
*fp_docparams-country = 'US'.
*&--- Call the generated function module
CALL FUNCTION fm_name
  EXPORTING
    /1bcdwb/docparams = fp_docparams
    lt_kunde_adresse  = lt_kunden_adressen

*    IMPORTING
*     /1BCDWB/FORMOUTPUT       =
  EXCEPTIONS
    usage_error           = 1
    system_error          = 2
    internal_error           = 3.
IF sy-subrc <> 0.
*  <error handling>
ENDIF.
*&---- Close the spool job
CALL FUNCTION 'FP_JOB_CLOSE'
*    IMPORTING
*     E_RESULT             =
  EXCEPTIONS
    usage_error           = 1
    system_error          = 2
    internal_error        = 3
    OTHERS               = 4.
IF sy-subrc <> 0.
*            <error handling>
ENDIF.

ENDFORM.

FORM set_default_directory .

  CHECK sy-batch EQ abap_false.

  CALL METHOD cl_gui_frontend_services=>get_desktop_directory
    CHANGING
      desktop_directory    = p_file
    EXCEPTIONS
      cntl_error           = 1
      error_no_gui         = 2
      not_supported_by_gui = 3
      OTHERS               = 4.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid
    TYPE 'E'
    NUMBER sy-msgno
    WITH sy-msgv1
    sy-msgv2
    sy-msgv3
    sy-msgv4.
  ENDIF.

  cl_gui_cfw=>update_view( ).

  gv_file = p_file.

ENDFORM.

*&---------------------------------------------------------------------*
*&      Form  DATEN_SORTIEREN
*&---------------------------------------------------------------------*
FORM daten_sortieren.
IF p_sort1 = 'X'.
  order_by = 'kunnr'.
ELSEIF p_sort2 = 'X'.
  order_by = 'pstlz'.
ELSE.
  order_by = 'kunnr'.
ENDIF.
ENDFORM.                    " DATEN_SORTIEREN

FORM liste_erstellen.
  LOOP AT lt_kunden_adressen INTO ld_kunden_adressen.
    WRITE: /(8) ld_kunden_adressen-kunnr,
                ld_kunden_adressen-adrnr,
                ld_kunden_adressen-spras,
                ld_kunden_adressen-land1,
                ld_kunden_adressen-pstlz.
 ENDLOOP.
  ULINE.

  if p_filx EQ 'X'.
WRITE: / 'Fehler bei Excel File "Kundennummer Konvention" Fehlgeschlagen bei:'.
  ULINE.
WRITE: / 'Status   Kundenummer        Excel Zeile'.
LOOP AT gt_errors INTO ld_errors.
WRITE: /(8) ld_errors-status,
            ld_errors-kunnr,
            ld_errors-zeile.
ENDLOOP.
ULINE.
ENDIF.
**********************  WRITE: / 'Anzahl Adressen auf Liste:      ', zclist.
**********************  ULINE.
ENDFORM.

FORM get_mitinhaber .
  CLEAR gt_output_mitinh.

  SELECT businesspartner FROM zbw_bp_cds_b_bpcust_01 INTO TABLE @gt_output_mitinh
   WHERE lifsd              IN @zlifsd          "Aktiv/inaktiv
     AND deletionindicator  EQ ''               "Löschvermerk
     AND zzgenossenschafter EQ 'X'.             "Genossenschafter

ENDFORM.

FORM get_carro .
  CLEAR gt_output_carro.

  SELECT businesspartner FROM zbw_bp_cds_b_bpcust_01 INTO TABLE @gt_output_carro
   WHERE lifsd          IN @zlifsd          "Aktiv/inaktiv
     AND deletionindicator EQ ''            "Löschvermerk
     AND zzcarrosserie_konzept EQ 'X'.

ENDFORM.
FORM get_statistik.
  DATA: itab TYPE i.
*    NEW-PAGE.
DESCRIBE TABLE lt_kunden_adressen LINES itab.
WRITE: / 'Anzal Kunden: ', itab.
*    NEW-PAGE.
ENDFORM.

FORM excel_datei_erstellen.

ENDFORM.
*INCLUDE z_abap_test_get_mitinhaberf01.
