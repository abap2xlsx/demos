CLASS zcl_excel_demo30 DEFINITION PUBLIC.
  PUBLIC SECTION.
    INTERFACES zif_excel_demo_output.
ENDCLASS.

CLASS zcl_excel_demo30 IMPLEMENTATION.

  METHOD zif_excel_demo_output~run.

    DATA: lo_worksheet TYPE REF TO zcl_excel_worksheet,
          lo_column    TYPE REF TO zcl_excel_column.

    DATA: lv_value  TYPE string,
          lv_count  TYPE i VALUE 10,
          lv_packed TYPE p LENGTH 16 DECIMALS 1 VALUE '1234567890.5'.

    CONSTANTS: lc_typekind_string TYPE abap_typekind VALUE cl_abap_typedescr=>typekind_string,
               lc_typekind_packed TYPE abap_typekind VALUE cl_abap_typedescr=>typekind_packed,
               lc_typekind_num    TYPE abap_typekind VALUE cl_abap_typedescr=>typekind_num,
               lc_typekind_date   TYPE abap_typekind VALUE cl_abap_typedescr=>typekind_date.

  " Creates active sheet
    CREATE OBJECT ro_excel.

  " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).
    lo_worksheet->set_title( ip_title = 'Cell data types' ).

    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Number as String'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = '11'
                          ip_abap_type = lc_typekind_string ).

    lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'String'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = ' String with leading spaces'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'String without leading spaces'
                          ip_abap_type = lc_typekind_string ).

    lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Packed'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = '50000.01-'
                          ip_abap_type = lc_typekind_packed ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = '5000.02'
                          ip_abap_type = lc_typekind_packed ).

    lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Number with Percentage'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = 2 ip_value = '0 %'
                          ip_abap_type = lc_typekind_num ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = 3 ip_value = '50 %'
                          ip_abap_type = lc_typekind_num ).

    lo_worksheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = 'Date'
                          ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'E' ip_row = 2 ip_value = '20110831'
                          ip_abap_type = lc_typekind_date ).

    WHILE lv_count <= 15.
      lv_value = lv_count.
      CONCATENATE 'Positive Value with' lv_value 'Digits' INTO lv_value SEPARATED BY space.
      lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_count ip_value = lv_value
                            ip_abap_type = lc_typekind_string ).
      lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_count ip_value = lv_packed
                            ip_abap_type = lc_typekind_packed ).
      CONCATENATE 'Positive Value with' lv_value 'Digits formated as string' INTO lv_value SEPARATED BY space.
      lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_count ip_value = lv_value
                            ip_abap_type = lc_typekind_string ).
      lo_worksheet->set_cell( ip_column = 'E' ip_row = lv_count ip_value = lv_packed
                            ip_abap_type = lc_typekind_string ).
      lv_packed = lv_packed * 10.
      lv_count  = lv_count + 1.
    ENDWHILE.

    lo_column = lo_worksheet->get_column( ip_column = 'A' ).
    lo_column->set_auto_size( abap_true ).
    lo_column = lo_worksheet->get_column( ip_column = 'B' ).
    lo_column->set_auto_size( abap_true ).
    lo_column = lo_worksheet->get_column( ip_column = 'C' ).
    lo_column->set_auto_size( abap_true ).
    lo_column = lo_worksheet->get_column( ip_column = 'D' ).
    lo_column->set_auto_size( abap_true ).
    lo_column = lo_worksheet->get_column( ip_column = 'E' ).
    lo_column->set_auto_size( abap_true ).


  ENDMETHOD.

ENDCLASS.
