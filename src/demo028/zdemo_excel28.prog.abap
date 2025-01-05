*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL28
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel28.

DATA: lo_excel        TYPE REF TO zcl_excel,
      lo_excel_writer TYPE REF TO zif_excel_writer,
      lo_worksheet    TYPE REF TO zcl_excel_worksheet,
      lo_column       TYPE REF TO zcl_excel_column.

DATA: lv_file      TYPE xstring,
      lv_bytecount TYPE i,
      lt_file_tab  TYPE solix_tab.

DATA: lv_full_path      TYPE string,
      lv_workdir        TYPE string,
      lv_file_separator TYPE c.

DATA lo_autofilter TYPE REF TO zcl_excel_autofilter.
DATA ls_area       TYPE zexcel_s_autofilter_area.
DATA lo_row        TYPE REF TO zcl_excel_row.

CONSTANTS c_initial_date TYPE d VALUE IS INITIAL.

CONSTANTS: lv_default_file_name TYPE string VALUE '28_HelloWorld.csv'.

PARAMETERS: p_path TYPE string LOWER CASE.

PARAMETERS p_skip_c AS CHECKBOX DEFAULT abap_true.
PARAMETERS p_skip_r AS CHECKBOX DEFAULT abap_true.
PARAMETERS p_xlsx   AS CHECKBOX DEFAULT abap_false.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.

  cl_gui_frontend_services=>directory_browse( EXPORTING initial_folder  = p_path
                                              CHANGING  selected_folder = p_path ).

INITIALIZATION.
  cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

START-OF-SELECTION.

  IF p_path IS INITIAL.
    p_path = lv_workdir.
  ENDIF.
  cl_gui_frontend_services=>get_file_separator( CHANGING file_separator = lv_file_separator ).
  CONCATENATE p_path lv_file_separator lv_default_file_name INTO lv_full_path.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet1' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Name' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Value 1' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Value 2' ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Column hidden' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'Text' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'Date and time' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = sy-datum ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = sy-uzeit ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = 'Initial Date' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = c_initial_date ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 5 ip_value = 'Row filtered out' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 6 ip_value = 'Row hidden' ).

  lo_autofilter = lo_excel->add_new_autofilter( io_sheet = lo_worksheet ) .
  ls_area-row_start = 1.
  ls_area-col_start = 1.
  ls_area-row_end = lo_worksheet->get_highest_row( ).
  ls_area-col_end = lo_worksheet->get_highest_column( ).
  lo_autofilter->set_filter_area( is_area = ls_area ).
  lo_autofilter->set_value( i_column = 1 " column A
                            i_value  = 'Text' ).
  lo_autofilter->set_value( i_column = 1 " column A
                            i_value  = 'Date and time' ).
  lo_autofilter->set_value( i_column = 1 " column A
                            i_value  = 'Initial Date' ).
  lo_autofilter->get_filter_area( ).

  lo_row = lo_worksheet->get_row( 6 ).
  lo_row->set_visible( abap_false ).

  lo_column = lo_worksheet->get_column( 'D' ).
  lo_column->set_visible( abap_false ).

  lo_column = lo_worksheet->get_column( 'B' ).
  lo_column->set_width( 11 ).

  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the second sheet' ).

  CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_csv.
  zcl_excel_writer_csv=>set_delimiter( ip_value = cl_abap_char_utilities=>horizontal_tab ).
  zcl_excel_writer_csv=>set_enclosure( ip_value = '''' ).
  zcl_excel_writer_csv=>set_endofline( ip_value = cl_abap_char_utilities=>cr_lf ).
  zcl_excel_writer_csv=>set_initial_ext_date( ip_value = '' ).
  zcl_excel_writer_csv=>set_skip_hidden_columns( p_skip_c ).
  zcl_excel_writer_csv=>set_skip_hidden_rows( p_skip_r ).

  zcl_excel_writer_csv=>set_active_sheet_index( i_active_worksheet = 2 ).

  lv_file = lo_excel_writer->write_file( lo_excel ).

  " Convert to binary
  CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
    EXPORTING
      buffer        = lv_file
    IMPORTING
      output_length = lv_bytecount
    TABLES
      binary_tab    = lt_file_tab.

  " Save the file
  REPLACE FIRST OCCURRENCE OF '.csv'  IN lv_full_path WITH '_Sheet2.csv'.
  cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                          CHANGING  data_tab     = lt_file_tab ).

  zcl_excel_writer_csv=>set_active_sheet_index_by_name( i_worksheet_name = 'Sheet1' ).
  lv_file = lo_excel_writer->write_file( lo_excel ).
  REPLACE FIRST OCCURRENCE OF '_Sheet2.csv'  IN lv_full_path WITH '_Sheet1.csv'.

  " Convert to binary
  CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
    EXPORTING
      buffer        = lv_file
    IMPORTING
      output_length = lv_bytecount
    TABLES
      binary_tab    = lt_file_tab.

  " Save the file
  cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                          CHANGING  data_tab     = lt_file_tab ).

  IF p_xlsx = abap_true.
    REPLACE FIRST OCCURRENCE OF '_Sheet1.csv' IN lv_full_path WITH '.xlsx'.
    CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
    lv_file = lo_excel_writer->write_file( lo_excel ).
    lt_file_tab = cl_bcs_convert=>xstring_to_solix( lv_file ).
    cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = xstrlen( lv_file )
                                                      filename     = lv_full_path
                                                      filetype     = 'BIN'
                                            CHANGING  data_tab     = lt_file_tab ).
  ENDIF.
