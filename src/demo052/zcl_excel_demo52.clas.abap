CLASS zcl_excel_demo52 DEFINITION PUBLIC.
  PUBLIC SECTION.
    INTERFACES zif_excel_demo_output.


  PROTECTED SECTION.
    CLASS-METHODS read_xlsx_from_web_repository
      IMPORTING
        objid         TYPE wwwdata-objid
      RETURNING
        VALUE(result) TYPE xstring
      RAISING
        zcx_excel.

ENDCLASS.

CLASS zcl_excel_demo52 IMPLEMENTATION.

  METHOD zif_excel_demo_output~run.

    DATA lo_original_worksheet TYPE REF TO zcl_excel_worksheet.
    DATA lo_worksheet TYPE REF TO zcl_excel_worksheet.
    DATA lv_source TYPE xstring.
    DATA lo_reader TYPE REF TO zcl_excel_reader_2007.

    "Read template file from SMW0 repository
    "This file is 'hardcoded' file containing various types of content to test the worksheet copy process.
    "You can download/upload the file with SMW0 for package $ABAP2XLSX_DEMOS_DEMO052, object name 'ZDEMO_EXCEL52_TEMPLATE'.
    "You might need to maintain MIME type application/vnd.openxmlformats-officedocument.spreadsheetml.sheet = *.xlsx in the settings
    "of SMW0 before uploading files.
    lv_source = read_xlsx_from_web_repository( 'ZDEMO_EXCEL52_TEMPLATE' ).

    "Open template file
    CREATE OBJECT lo_reader.
    ro_excel = lo_reader->zif_excel_reader~load( lv_source ).
    lo_original_worksheet = ro_excel->get_worksheet_by_name( 'original' ).

    "Clone 1
    lo_worksheet = ro_excel->clone_worksheet(
                     io_source  = lo_original_worksheet
                     iv_title   = 'copy1'
                   ).

    "Clone 2
    DATA ls_clone_options TYPE zcl_excel_worksheet=>ts_clone_options.
    ls_clone_options-skip_cells = abap_true.
    lo_worksheet = ro_excel->clone_worksheet(
                     io_source  = lo_original_worksheet
                     iv_title   = 'copy2'
                     is_options = ls_clone_options
                   ).

  ENDMETHOD.



  METHOD read_xlsx_from_web_repository.

    DATA: query_string   TYPE w3query,
          query_table    TYPE TABLE OF w3query,
          html_table     TYPE TABLE OF w3html,
          return_code    TYPE w3param-ret_code,
          content_type   TYPE w3param-cont_type,
          content_length TYPE w3param-cont_len,
          mime_table     TYPE TABLE OF w3mime.

    CLEAR: query_table, query_string.
    query_string-name = '_OBJECT_ID'.
    query_string-value = objid.
    APPEND query_string TO query_table.

    CALL FUNCTION 'WWW_GET_MIME_OBJECT'
      TABLES
        query_string        = query_table
        html                = html_table
        mime                = mime_table
      CHANGING
        return_code         = return_code
        content_type        = content_type
        content_length      = content_length
      EXCEPTIONS
        object_not_found    = 1
        parameter_not_found = 2
        OTHERS              = 3.
    IF sy-subrc = 1.
      RETURN.
    ELSEIF sy-subrc >= 2.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'WWW_GET_MIME_OBJECT'.
    ENDIF.

    result = cl_bcs_convert=>solix_to_xstring( it_solix = mime_table iv_size = content_length ).

  ENDMETHOD.

ENDCLASS.
