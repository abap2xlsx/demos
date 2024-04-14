*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL16
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel16.

DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_drawing   TYPE REF TO zcl_excel_drawing.


DATA: ls_io TYPE skwf_io.

CONSTANTS: gc_save_file_name TYPE string VALUE '16_Drawings.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

SELECTION-SCREEN BEGIN OF BLOCK b01 WITH FRAME TITLE TEXT-b01.
  PARAMETERS p_wobjid TYPE w3objid DEFAULT 'SAPLOGO.GIF'.
SELECTION-SCREEN END OF BLOCK b01.

SELECTION-SCREEN BEGIN OF BLOCK b02 WITH FRAME TITLE TEXT-b02.
  PARAMETERS p_file TYPE string LOWER CASE DEFAULT '<SAPGUI-directory>\wwi\graphics\W_bio.bmp'.
SELECTION-SCREEN END OF BLOCK b02.

SELECTION-SCREEN BEGIN OF BLOCK b03 WITH FRAME TITLE TEXT-b03.
  PARAMETERS p_mobjid TYPE sdok_docid DEFAULT '456694429165174BE10000000A1550C0'. " Question mark in standard Web Dynpro WDT_QUIZ
  PARAMETERS p_class TYPE sdok_class DEFAULT 'M_IMAGE_P'.
  PARAMETERS pobjtype TYPE skwf_ioty DEFAULT 'P'.
SELECTION-SCREEN END OF BLOCK b03.

START-OF-SELECTION.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  "Load samle image
  DATA: lt_bin     TYPE solix_tab,
        lv_len     TYPE i,
        lv_content TYPE xstring,
        ls_key     TYPE wwwdatatab.
  DATA gv_file TYPE string.
  DATA gv_sapgui_directory TYPE string.
  DATA gv_message TYPE string.

  cl_gui_frontend_services=>get_sapgui_directory( CHANGING   sapgui_directory = gv_sapgui_directory
                                                  EXCEPTIONS cntl_error       = 1
                                                             error_no_gui     = 2 ).
  IF sy-subrc <> 0.
    gv_message = |Exception { sy-subrc } during CL_GUI_FRONTEND_SERVICES=>GET_SAPGUI_DIRECTORY|.
    MESSAGE gv_message TYPE 'I' DISPLAY LIKE 'E'.
    STOP.
  ENDIF.
  " flush to send previous call to frontend
  cl_gui_cfw=>flush( EXCEPTIONS cntl_system_error = 1
                                cntl_error        = 2
                                OTHERS            = 3 ).
  IF sy-subrc <> 0.
    gv_message = |Exception { sy-subrc } during CL_GUI_CFW=>FLUSH|.
    MESSAGE gv_message TYPE 'I' DISPLAY LIKE 'E'.
    STOP.
  ENDIF.

  gv_file = replace( val  = p_file
                     sub  = '<SAPGUI-directory>'
                     with = gv_sapgui_directory ).
  CALL METHOD cl_gui_frontend_services=>gui_upload
    EXPORTING
      filename                = gv_file
      filetype                = 'BIN'
    IMPORTING
      filelength              = lv_len
    CHANGING
      data_tab                = lt_bin
    EXCEPTIONS
      file_open_error         = 1
      file_read_error         = 2
      no_batch                = 3
      gui_refuse_filetransfer = 4
      invalid_type            = 5
      no_authority            = 6
      unknown_error           = 7
      bad_data_format         = 8
      header_not_allowed      = 9
      separator_not_allowed   = 10
      header_too_long         = 11
      unknown_dp_error        = 12
      access_denied           = 13
      dp_out_of_memory        = 14
      disk_full               = 15
      dp_timeout              = 16
      not_supported_by_gui    = 17
      error_no_gui            = 18
      OTHERS                  = 19.
  IF sy-subrc <> 0.
    gv_message = |Exception { sy-subrc } while uploading the file at "{ p_file }"|.
    MESSAGE gv_message TYPE 'I' DISPLAY LIKE 'E'.
    STOP.
  ENDIF.

  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = lv_len
    IMPORTING
      buffer       = lv_content
    TABLES
      binary_tab   = lt_bin
    EXCEPTIONS
      failed       = 1
      OTHERS       = 2.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4
               DISPLAY LIKE 'E'.
    STOP.
  ENDIF.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Sheet1' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Image from web repository (SMW0)' ).

  " create global drawing, set position and media from web repository
  lo_drawing = lo_excel->add_new_drawing( ).
  lo_drawing->set_position( ip_from_row = 3
                            ip_from_col = 'B' ).

  ls_key-relid = 'MI'.
  ls_key-objid = p_wobjid.
  lo_drawing->set_media_www( ip_key = ls_key
                             ip_width = 166
                             ip_height = 75 ).

  " assign drawing to the worksheet
  lo_worksheet->add_drawing( lo_drawing ).

  " another drawing from a XSTRING read from a file
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 8 ip_value = |Image from a file ({ p_file })| ).
  lo_drawing = lo_excel->add_new_drawing( ).
  lo_drawing->set_position( ip_from_row = 9
                            ip_from_col = 'B' ).
  lo_drawing->set_media( ip_media       = lv_content
                         ip_media_type = zcl_excel_drawing=>c_media_type_bmp
                         ip_width = 83
                         ip_height = 160 ).

  lo_worksheet->add_drawing( lo_drawing ).

  ls_io-objid   = p_mobjid.
  ls_io-class   = p_class.
  ls_io-objtype = pobjtype.
  IF ls_io IS NOT INITIAL.
    " another drawing from a XSTRING read from a file
    lo_worksheet->set_cell( ip_column = 'B' ip_row = 18 ip_value = 'Mime repository (by default Question mark in standard Web Dynpro WDT_QUIZ)' ).
    lo_drawing = lo_excel->add_new_drawing( ).
    lo_drawing->set_position( ip_from_row = 19
                              ip_from_col = 'B' ).
    lo_drawing->set_media_mime( ip_io       = ls_io
                                ip_width = 126
                                ip_height = 145 ).

    lo_worksheet->add_drawing( lo_drawing ).
  ENDIF.

*** Create output
  lcl_output=>output( lo_excel ).
