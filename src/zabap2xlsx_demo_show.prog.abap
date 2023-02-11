*&---------------------------------------------------------------------*
*& Report  ZABAP2XLSX_DEMO_SHOW
*&---------------------------------------------------------------------*
REPORT zabap2xlsx_demo_show.


*----------------------------------------------------------------------*
*       CLASS lcl_perform DEFINITION
*----------------------------------------------------------------------*
CLASS lcl_perform DEFINITION CREATE PRIVATE.
  PUBLIC SECTION.
    CLASS-METHODS: setup_objects,
      collect_reports,

      handle_nav FOR EVENT double_click OF cl_gui_alv_grid
        IMPORTING e_row.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_reports,
             progname    TYPE reposrc-progname,
             sort        TYPE reposrc-progname,
             description TYPE repti,
             filename    TYPE string,
           END OF ty_reports.

    CLASS-DATA:
      mo_grid     TYPE REF TO cl_gui_alv_grid,
      mo_editor   TYPE REF TO cl_gui_abapedit,
      mo_document TYPE REF TO i_oi_document_proxy,

      t_reports   TYPE STANDARD TABLE OF ty_reports WITH NON-UNIQUE DEFAULT KEY,
      error       TYPE REF TO i_oi_error,
      t_errors    TYPE STANDARD TABLE OF REF TO i_oi_error WITH NON-UNIQUE DEFAULT KEY,
      mo_control  TYPE REF TO i_oi_container_control.   "Office Dokument

    CLASS-METHODS add_error
      IMPORTING
        error TYPE REF TO i_oi_error.

    CLASS-METHODS show_errors.
    CLASS-METHODS set_grid
      IMPORTING
        container TYPE REF TO cl_gui_container.
    CLASS-METHODS set_editor
      IMPORTING
        container TYPE REF TO cl_gui_container.
    CLASS-METHODS set_preview
      IMPORTING
        container TYPE REF TO cl_gui_container.

ENDCLASS.                    "lcl_perform DEFINITION


START-OF-SELECTION.
  lcl_perform=>collect_reports( ).
  lcl_perform=>setup_objects( ).

END-OF-SELECTION.

  WRITE '.'.  " Force output


*----------------------------------------------------------------------*
*       CLASS lcl_perform IMPLEMENTATION
*----------------------------------------------------------------------*
CLASS lcl_perform IMPLEMENTATION.
  METHOD setup_objects.

    DATA: lo_split     TYPE REF TO cl_gui_splitter_container.

    CREATE OBJECT lo_split
      EXPORTING
        parent                  = cl_gui_container=>screen0
        rows                    = 1
        columns                 = 3
        no_autodef_progid_dynnr = 'X'.
    lo_split->set_column_width( id = 1 width = 20 ).
    lo_split->set_column_width( id = 2 width = 40 ).

    " Left:   List of reports
    set_grid( lo_split->get_container( row = 1 column = 1 ) ).

    " Middle: ABAP Editor with coding
    set_editor( lo_split->get_container( row = 1 column = 2 ) ).

    " right:  DemoOutput
    set_preview( lo_split->get_container( row = 1 column = 3 ) ).


  ENDMETHOD.                    "setup_objects

  "collect_reports
  METHOD collect_reports.
    FIELD-SYMBOLS <report> LIKE LINE OF t_reports.
    DATA t_source TYPE STANDARD TABLE OF text255 WITH NON-UNIQUE DEFAULT KEY.
    DATA texts TYPE STANDARD TABLE OF textpool.
    DATA description TYPE textpool.

* Get all demoreports
    SELECT progname
      INTO CORRESPONDING FIELDS OF TABLE t_reports
      FROM reposrc
      WHERE progname LIKE 'ZDEMO_EXCEL%'
        AND progname <> sy-repid
        AND subc     = '1'.

    LOOP AT t_reports ASSIGNING <report>.

* Check if already switched to new outputoptions
      READ REPORT <report>-progname INTO t_source.
      IF sy-subrc = 0.
        FIND 'INCLUDE zdemo_excel_outputopt_incl.' IN TABLE t_source IGNORING CASE.
      ENDIF.
      IF sy-subrc <> 0.
        DELETE t_reports.
        CONTINUE.
      ENDIF.


* Build half-numeric sort
      <report>-sort = <report>-progname.
      REPLACE REGEX '(ZDEMO_EXCEL)(\d\d)\s*$' IN <report>-sort WITH '$1\0$2'. "      REPLACE REGEX '(ZDEMO_EXCEL)([^][^])*$' IN <report>-sort WITH '$1$2'.REPLACE REGEX '(ZDEMO_EXCEL)([^][^])*$' IN <report>-sort WITH '$1$2'.REPLACE

      REPLACE REGEX '(ZDEMO_EXCEL)(\d)\s*$'      IN <report>-sort WITH '$1\0\0$2'.

* get report text
      READ TEXTPOOL <report>-progname INTO texts LANGUAGE sy-langu.
      READ TABLE texts INTO description WITH KEY id = 'R'.
      IF sy-subrc > 0.
        "If not available in logon language, use english
        READ TEXTPOOL <report>-progname INTO texts LANGUAGE 'E'.
        READ TABLE texts INTO description WITH KEY id = 'R'.
      ENDIF.
      "set report title
      <report>-description = description-entry.

    ENDLOOP.

    SORT t_reports BY sort progname.

  ENDMETHOD.  "collect_reports

  METHOD handle_nav.

    CONSTANTS: filename TYPE text80 VALUE 'ZABAP2XLSX_DEMO_SHOW.xlsx'.
    DATA: wa_report  LIKE LINE OF t_reports,
          t_source   TYPE STANDARD TABLE OF text255,
          t_rawdata  TYPE solix_tab,
          wa_rawdata LIKE LINE OF t_rawdata,
          bytecount  TYPE i,
          length     TYPE i,
          add_selopt TYPE flag.

    READ TABLE t_reports INTO wa_report INDEX e_row-index.
    CHECK sy-subrc = 0.

* Set new text into middle frame
    READ REPORT wa_report-progname INTO t_source.
    mo_editor->set_text( t_source ).


* Unload old xls-file
    mo_document->close_document( ).

* Get the demo
* If additional parameters found on selection screen, start via selection screen , otherwise start w/o
    CLEAR add_selopt.
    FIND 'PARAMETERS' IN TABLE t_source.
    IF sy-subrc = 0.
      add_selopt = 'X'.
    ELSE.
      FIND 'SELECT-OPTIONS' IN TABLE t_source.
      IF sy-subrc = 0.
        add_selopt = 'X'.
      ENDIF.
    ENDIF.
    IF add_selopt IS INITIAL.
      SUBMIT (wa_report-progname) AND RETURN             "#EC CI_SUBMIT
              WITH p_backfn = filename
              WITH rb_back  = 'X'
              WITH rb_down  = ' '
              WITH rb_send  = ' '
              WITH rb_show  = ' '.
    ELSE.
      SUBMIT (wa_report-progname) VIA SELECTION-SCREEN AND RETURN "#EC CI_SUBMIT
              WITH p_backfn = filename
              WITH rb_back  = 'X'
              WITH rb_down  = ' '
              WITH rb_send  = ' '
              WITH rb_show  = ' '.
    ENDIF.

    OPEN DATASET filename FOR INPUT IN BINARY MODE.
    IF sy-subrc = 0.
      DO.
        CLEAR wa_rawdata.
        READ DATASET filename INTO wa_rawdata LENGTH length.
        IF sy-subrc <> 0.
          APPEND wa_rawdata TO t_rawdata.
          ADD length TO bytecount.
          EXIT.
        ENDIF.
        APPEND wa_rawdata TO t_rawdata.
        ADD length TO bytecount.
      ENDDO.
      CLOSE DATASET filename.
    ENDIF.

    mo_control->get_document_proxy(
      EXPORTING
        document_type  = 'Excel.Sheet'                " EXCEL
        no_flush       = ' '
        register_container = abap_true
      IMPORTING
        document_proxy = mo_document
        error          = error ).

    mo_document->open_document_from_table(
        document_size    = bytecount
        document_table   = t_rawdata
        open_inplace     = 'X' ).

  ENDMETHOD.                    "handle_nav

  METHOD add_error.
    IF error->has_failed = abap_true.
      APPEND error TO t_errors.
    ENDIF.
  ENDMETHOD.

  METHOD show_errors.
    IF t_errors IS NOT INITIAL.
      MESSAGE 'There were errors. See internal T_ERRORS table.'(001) TYPE 'I'.
    ENDIF.
  ENDMETHOD.

  METHOD set_grid.

    DATA: lt_fieldcat           TYPE lvc_t_fcat,
          ls_layout             TYPE lvc_s_layo,
          ls_variant            TYPE disvariant,
          lt_excluded_functions TYPE ui_functions.
    FIELD-SYMBOLS: <fc> LIKE LINE OF lt_fieldcat.

    CREATE OBJECT mo_grid
      EXPORTING
        i_parent = container.
    SET HANDLER lcl_perform=>handle_nav FOR mo_grid.

    ls_variant-report = sy-repid.
    ls_variant-handle = '0001'.

    ls_layout-cwidth_opt = 'X'.

    APPEND INITIAL LINE TO lt_fieldcat ASSIGNING <fc>.
    <fc>-fieldname   = 'PROGNAME'.
    <fc>-tabname     = 'REPOSRC'.

    APPEND INITIAL LINE TO lt_fieldcat ASSIGNING <fc>.
    <fc>-fieldname   = 'SORT'.
    <fc>-ref_field   = 'PROGNAME'.
    <fc>-ref_table   = 'REPOSRC'.
    <fc>-tech        = abap_true. "No need to display this help field

    APPEND INITIAL LINE TO lt_fieldcat ASSIGNING <fc>.
    <fc>-fieldname   = 'DESCRIPTION'.
    <fc>-ref_field   = 'REPTI'.
    <fc>-ref_table   = 'RS38M'.

    APPEND cl_gui_alv_grid=>mc_mb_subtot  TO lt_excluded_functions.
    APPEND cl_gui_alv_grid=>mc_mb_sum     TO lt_excluded_functions.
    APPEND cl_gui_alv_grid=>mc_mb_variant TO lt_excluded_functions.
    APPEND cl_gui_alv_grid=>mc_mb_view    TO lt_excluded_functions.
    APPEND cl_gui_alv_grid=>mc_fc_detail  TO lt_excluded_functions.
    APPEND cl_gui_alv_grid=>mc_fc_graph   TO lt_excluded_functions.
    APPEND cl_gui_alv_grid=>mc_fc_info    TO lt_excluded_functions.

    mo_grid->set_table_for_first_display(
      EXPORTING
        is_variant                    = ls_variant
        i_save                        = 'A'
        is_layout                     = ls_layout
        it_toolbar_excluding          = lt_excluded_functions
      CHANGING
        it_outtab                     = t_reports
        it_fieldcatalog               = lt_fieldcat
      EXCEPTIONS
        invalid_parameter_combination = 1
        program_error                 = 2
        too_many_lines                = 3
        OTHERS                        = 4 ).

  ENDMETHOD.

  METHOD set_editor.

    CREATE OBJECT mo_editor
      EXPORTING
        parent = container.
    mo_editor->set_readonly_mode( 1 ).

  ENDMETHOD.

  METHOD set_preview.
    c_oi_container_control_creator=>get_container_control(
       IMPORTING
         control = mo_control
         error   = error ).
    add_error( error ).

    mo_control->init_control(
      EXPORTING
        inplace_enabled     = 'X'
        no_flush            = 'X'
        r3_application_name = 'Demo Document Container'
        parent              = container
      IMPORTING
        error               = error
      EXCEPTIONS
        OTHERS              = 2 ).
    add_error( error ).

    mo_control->get_document_proxy(
      EXPORTING
        document_type  = 'Excel.Sheet'                " EXCEL
        no_flush       = ' '
      IMPORTING
        document_proxy = mo_document
        error          = error ).
    add_error( error ).

    show_errors( ).

  ENDMETHOD.

ENDCLASS.                    "lcl_perform IMPLEMENTATION
