*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL51
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel51.

TYPES: ty_sflight_lines TYPE TABLE OF sflight.

DATA: lo_excel         TYPE REF TO zcl_excel,
      lv_excel_xstring TYPE xstring,
      lo_writer        TYPE REF TO zif_excel_writer,
      lo_reader        TYPE REF TO zif_excel_reader,
      lo_worksheet     TYPE REF TO zcl_excel_worksheet.

DATA: lt_field_catalog  TYPE zexcel_t_fieldcatalog,
      ls_table_settings TYPE zexcel_s_table_settings.



START-OF-SELECTION.

  FIELD-SYMBOLS: <fs_field_catalog> TYPE zexcel_s_fieldcatalog.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Internal table' ).


  DATA lt_test TYPE ty_sflight_lines.
  PERFORM load_fixed_data CHANGING lt_test.

  lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = lt_test ).

  ls_table_settings-table_style  = zcl_excel_table=>builtinstyle_medium5.

  lo_worksheet->bind_table( ip_table          = lt_test
                            is_table_settings = ls_table_settings
                            it_field_catalog  = lt_field_catalog ).

  lo_worksheet->freeze_panes( ip_num_rows = 1 ).

  " Create output
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  lv_excel_xstring = lo_writer->write_file( lo_excel ).


**********************************************************************
  CLEAR: lt_test, lo_excel, lo_worksheet.
**********************************************************************
  " Import
  CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
  lo_excel = lo_reader->load( lv_excel_xstring ).
  lo_worksheet = lo_excel->get_worksheet_by_index( 1 ).
  lo_worksheet->convert_to_table(
    IMPORTING
      et_data = lt_test
  ).

  DATA: lo_salv TYPE REF TO cl_salv_table.
  cl_salv_table=>factory( IMPORTING r_salv_table = lo_salv
                          CHANGING t_table = lt_test ).
  lo_salv->display( ).


FORM load_fixed_data CHANGING ct_test TYPE ty_sflight_lines.
  DATA: lt_lines  TYPE TABLE OF string,
        lv_line   TYPE string,
        lt_fields TYPE TABLE OF string,
        lv_comp   TYPE i,
        lv_field  TYPE string,
        ls_test   TYPE sflight.
  FIELD-SYMBOLS: <lv_field> TYPE simple.

  APPEND 'AA 0017 20171219  422 USD 747-400  385 371 191334 31  28  21  21' TO lt_lines.
  APPEND 'AA 0017 20180309  422 USD 747-400  385 365 189984 31  29  21  20' TO lt_lines.
  APPEND 'AA 0017 20180528  422 USD 747-400  385 374 193482 31  30  21  20' TO lt_lines.
  APPEND 'AA 0017 20180816  422 USD 747-400  385 372 193127 31  30  21  20' TO lt_lines.
  APPEND 'AA 0017 20181104  422 USD 747-400  385  44  23908 31   4  21   3' TO lt_lines.
  APPEND 'AA 0017 20190123  422 USD 747-400  385  40  20347 31   3  21   2' TO lt_lines.
  APPEND 'AZ 0555 20171219  185 EUR 737-800  140 133  32143 12  12  10  10' TO lt_lines.
  APPEND 'AZ 0555 20180309  185 EUR 737-800  140 137  32595 12  12  10  10' TO lt_lines.
  APPEND 'AZ 0555 20180528  185 EUR 737-800  140 134  31899 12  11  10  10' TO lt_lines.
  APPEND 'AZ 0555 20180816  185 EUR 737-800  140 128  29775 12  10  10   9' TO lt_lines.
  APPEND 'AZ 0555 20181104  185 EUR 737-800  140   0      0 12   0  10   0' TO lt_lines.
  APPEND 'AZ 0555 20190123  185 EUR 737-800  140  23   5392 12   1  10   2' TO lt_lines.
  APPEND 'AZ 0789 20171219 1030 EUR 767-200  260 250 307176 21  20  11  11' TO lt_lines.
  APPEND 'AZ 0789 20180309 1030 EUR 767-200  260 252 306054 21  20  11  10' TO lt_lines.
  APPEND 'AZ 0789 20180528 1030 EUR 767-200  260 252 307063 21  20  11  10' TO lt_lines.
  APPEND 'AZ 0789 20180816 1030 EUR 767-200  260 249 300739 21  19  11  10' TO lt_lines.
  APPEND 'AZ 0789 20181104 1030 EUR 767-200  260 104 127647 21   8  11   5' TO lt_lines.
  APPEND 'AZ 0789 20190123 1030 EUR 767-200  260  18  22268 21   1  11   1' TO lt_lines.
  APPEND 'DL 0106 20171217  611 USD A380-800 475 458 324379 30  29  20  20' TO lt_lines.
  APPEND 'DL 0106 20180307  611 USD A380-800 475 458 324330 30  30  20  20' TO lt_lines.
  APPEND 'DL 0106 20180526  611 USD A380-800 475 459 328149 30  29  20  20' TO lt_lines.
  APPEND 'DL 0106 20180814  611 USD A380-800 475 462 326805 30  30  20  18' TO lt_lines.
  APPEND 'DL 0106 20181102  611 USD A380-800 475 167 115554 30  10  20   6' TO lt_lines.
  APPEND 'DL 0106 20190121  611 USD A380-800 475  11   9073 30   1  20   1' TO lt_lines.
  LOOP AT lt_lines INTO lv_line.
    CONDENSE lv_line.
    SPLIT lv_line AT space INTO TABLE lt_fields.
    lv_comp = 2.
    LOOP AT lt_fields INTO lv_field.
      ASSIGN COMPONENT lv_comp OF STRUCTURE ls_test TO <lv_field>.
      <lv_field> = lv_field.
      lv_comp = lv_comp + 1.
    ENDLOOP.
    APPEND ls_test TO ct_test.
  ENDLOOP.
ENDFORM.
