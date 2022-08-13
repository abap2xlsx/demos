*&---------------------------------------------------------------------*
*& Report zdemo_excel_checker
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel_checker.


CLASS lcl_xlsx_diff_file_ext DEFINITION DEFERRED.


CLASS lcx_zip_diff DEFINITION
  INHERITING FROM cx_static_check
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    METHODS constructor
      IMPORTING
        text     TYPE clike
        previous TYPE REF TO cx_root OPTIONAL.
    METHODS get_text REDEFINITION.
    METHODS get_longtext REDEFINITION.
  PROTECTED SECTION.
  PRIVATE SECTION.
    DATA text TYPE string.
ENDCLASS.


CLASS lcl_zip_cleanup_for_diff DEFINITION
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES : BEGIN OF ty_zip_structure,
              ref_to_structure TYPE REF TO data,
              ref_to_x         TYPE REF TO data,
              length           TYPE i,
              view             TYPE REF TO cl_abap_view_offlen,
              charset_bit      TYPE i,
              conv_in_utf8     TYPE REF TO cl_abap_conv_in_ce,
              conv_in_ibm437   TYPE REF TO cl_abap_conv_in_ce,
              conv_out_utf8    TYPE REF TO cl_abap_conv_out_ce,
              conv_out_ibm437  TYPE REF TO cl_abap_conv_out_ce,
            END OF ty_zip_structure.

    METHODS run
      IMPORTING
        zip_xstring   TYPE xstring
      RETURNING
        VALUE(result) TYPE xstring
      RAISING
        zcx_excel.

  PRIVATE SECTION.

    METHODS init_structure
      IMPORTING
        length        TYPE i
        charset_bit   TYPE i
        structure     TYPE any
      RETURNING
        VALUE(result) TYPE ty_zip_structure.

    METHODS write_zip
      IMPORTING
        offset        TYPE i
      CHANGING
        zip_structure TYPE ty_zip_structure
        zip_xstring   TYPE xstring.

    METHODS read_zip
      IMPORTING
        zip_xstring   TYPE xstring
        offset        TYPE i
      CHANGING
        zip_structure TYPE ty_zip_structure.

ENDCLASS.


CLASS lcl_xlsx_cleanup_for_diff DEFINITION
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS run
      IMPORTING
        xstring       TYPE xstring

      RETURNING
        VALUE(result) TYPE xstring
      RAISING
        zcx_excel.

ENDCLASS.


CLASS lcl_xlsx_diff_item DEFINITION
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES : ty_diff_state TYPE i,
            BEGIN OF ty_diff_item_attr,
              date TYPE cl_abap_zip=>t_file-date,
              time TYPE cl_abap_zip=>t_file-time,
              size TYPE cl_abap_zip=>t_file-size,
            END OF ty_diff_item_attr,
            BEGIN OF ty_diff_item,
              local_name  TYPE string,
              full_path   TYPE cl_abap_zip=>t_file-name,
              attr_1      TYPE ty_diff_item_attr,
              attr_2      TYPE ty_diff_item_attr,
              diff_state  TYPE ty_diff_state,
              folder_diff TYPE REF TO lcl_xlsx_diff_item,
            END OF ty_diff_item,
            ty_diff_items TYPE STANDARD TABLE OF ty_diff_item WITH DEFAULT KEY,
            BEGIN OF ty_item,
              local_name TYPE string,
              full_path  TYPE cl_abap_zip=>t_file-name,
              date       TYPE cl_abap_zip=>t_file-date,
              time       TYPE cl_abap_zip=>t_file-time,
              size       TYPE cl_abap_zip=>t_file-size,
            END OF ty_item,
            ty_items TYPE SORTED TABLE OF ty_item WITH UNIQUE KEY local_name.
    CONSTANTS: BEGIN OF state,
                 same                   TYPE ty_diff_state VALUE 1,
                 only_content_changed   TYPE ty_diff_state VALUE 2,
                 only_attribute_changed TYPE ty_diff_state VALUE 3,
                 changed                TYPE ty_diff_state VALUE 4,
                 added                  TYPE ty_diff_state VALUE 5,
                 deleted                TYPE ty_diff_state VALUE 6,
                 undefined              TYPE ty_diff_state VALUE 7,
                 folder_undefined       TYPE ty_diff_state VALUE 8,
               END OF state.

    CLASS-METHODS get_diff
      IMPORTING
        zip_1         TYPE REF TO cl_abap_zip
        zip_2         TYPE REF TO cl_abap_zip
      RETURNING
        VALUE(result) TYPE REF TO lcl_xlsx_diff_item.

    DATA: path         TYPE string READ-ONLY,
          items        TYPE ty_diff_items READ-ONLY,
          is_different TYPE abap_bool READ-ONLY.

  PROTECTED SECTION.
  PRIVATE SECTION.

    METHODS compare_items_at_path
      IMPORTING
        path          TYPE string OPTIONAL
      RETURNING
        VALUE(result) TYPE REF TO lcl_xlsx_diff_item.

    METHODS get_items_at_path
      IMPORTING
        items             TYPE cl_abap_zip=>t_files
        path              TYPE string OPTIONAL
      RETURNING
        VALUE(path_items) TYPE ty_items.

    METHODS compare_two_items
      IMPORTING
        zip_1         TYPE REF TO cl_abap_zip
        zip_2         TYPE REF TO cl_abap_zip
        name_1        TYPE string
        name_2        TYPE string
      RETURNING
        VALUE(result) TYPE ty_diff_item-diff_state.

    DATA: zip_1 TYPE REF TO cl_abap_zip,
          zip_2 TYPE REF TO cl_abap_zip.

ENDCLASS.


CLASS lcl_xlsx_diff_viewer DEFINITION
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS constructor
      IMPORTING
        io_container TYPE REF TO cl_gui_container.

    METHODS diff_and_view
      IMPORTING
        title_old TYPE csequence OPTIONAL
        title_new TYPE csequence OPTIONAL
        zip_old   TYPE REF TO cl_abap_zip
        zip_new   TYPE REF TO cl_abap_zip.

    EVENTS selection_changed
      EXPORTING
        VALUE(node) TYPE lcl_xlsx_diff_item=>ty_diff_item
        VALUE(zip_old) TYPE REF TO cl_abap_zip
        VALUE(zip_new) TYPE REF TO cl_abap_zip.

  PROTECTED SECTION.
  PRIVATE SECTION.

    CONSTANTS: state    LIKE lcl_xlsx_diff_item=>state VALUE lcl_xlsx_diff_item=>state,
               icon_old TYPE tv_image VALUE icon_create,
               icon_new TYPE tv_image VALUE icon_msg.

    METHODS add_tree_nodes
      IMPORTING
        parent_node_key TYPE tm_nodekey
        path            TYPE string
        diff_object     TYPE REF TO lcl_xlsx_diff_item.

    METHODS view
      IMPORTING
        i_zip_diff TYPE REF TO lcl_xlsx_diff_item.

    METHODS on_selection_changed
                FOR EVENT selection_changed OF cl_tree_model
      IMPORTING node_key.

    DATA: container   TYPE REF TO cl_gui_container,
          go_splitter TYPE REF TO cl_gui_splitter_container,
          go_toolbar  TYPE REF TO cl_gui_toolbar,
          go_tree     TYPE REF TO cl_column_tree_model,
          node_key    TYPE i,
          title_old   TYPE text40,
          title_new   TYPE text40,
          zip_old     TYPE REF TO cl_abap_zip,
          zip_new     TYPE REF TO cl_abap_zip,
          diff_file   TYPE REF TO lcl_xlsx_diff_file_ext.

ENDCLASS.


CLASS lcl_xlsx_diff_file_ext DEFINITION
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    CLASS-METHODS class_constructor.

    METHODS gui_download
      IMPORTING
        i_content   TYPE xstring
        i_file_path TYPE string.

    METHODS xml_pretty_print
      CHANGING
        c_content TYPE xstring.

    METHODS on_selection_changed
                FOR EVENT selection_changed OF lcl_xlsx_diff_viewer
      IMPORTING node zip_old zip_new.

  PROTECTED SECTION.

  PRIVATE SECTION.

    CLASS-DATA: temp_dir TYPE string.

    METHODS zip_get_and_gui_download
      IMPORTING
        zip           TYPE REF TO cl_abap_zip
        name          TYPE string
        prefix        TYPE string
      RETURNING
        VALUE(result) TYPE string
      RAISING
        lcx_zip_diff.

    METHODS gui_execute
      IMPORTING
        i_parameter TYPE string
      RAISING
        lcx_zip_diff.

    CONSTANTS: state LIKE lcl_xlsx_diff_item=>state VALUE lcl_xlsx_diff_item=>state.

ENDCLASS.


CLASS lcl_app DEFINITION.

  PUBLIC SECTION.

    METHODS at_selection_screen.

    METHODS at_selection_screen_on_exit.

    METHODS at_selection_screen_output.

    METHODS set_sscrfields
      CHANGING
        sscrfields TYPE sscrfields.

  PRIVATE SECTION.

    TYPES : BEGIN OF ty_demo,
              program  TYPE trdir-name,
              objid    TYPE wwwdata-objid,
              text     TYPE wwwdata-text,
              filename TYPE string,
            END OF ty_demo,
            ty_demos TYPE STANDARD TABLE OF ty_demo WITH DEFAULT KEY,
            BEGIN OF ty_check_result,
              diff                   TYPE abap_bool,
              xlsx_just_now          TYPE xstring,
              xlsx_reference         TYPE xstring,
              compare_xlsx_just_now  TYPE xstring,
              compare_xlsx_reference TYPE xstring,
            END OF ty_check_result,
            BEGIN OF ty_alv_line,
              status_icon            TYPE string,
              xlsx_diff              TYPE string,
              write_smw0             TYPE string,
              objid                  TYPE wwwdata-objid,
              obj_text               TYPE wwwdata-text,
              program                TYPE trdir-name,
              prog_text              TYPE trdirt-text,
              filename               TYPE string,
              xlsx_just_now          TYPE xstring,
              xlsx_reference         TYPE xstring,
              compare_xlsx_just_now  TYPE xstring,
              compare_xlsx_reference TYPE xstring,
              cell_types             TYPE salv_t_int4_column,
            END OF ty_alv_line,
            ty_alv_table              TYPE STANDARD TABLE OF ty_alv_line WITH DEFAULT KEY,
            ty_popup_confirm_question TYPE c LENGTH 400.

    METHODS at_selection_screen_output1000.

    METHODS at_selection_screen_output1001
      RAISING
        zcx_excel
        cx_salv_data_error
        cx_salv_not_found
        cx_salv_msg.

    METHODS check_regression
      IMPORTING
        demo          TYPE ty_demo
      RETURNING
        VALUE(result) TYPE ty_check_result
      RAISING
        zcx_excel.

    METHODS get_list_of_demo_files
      RETURNING
        VALUE(result) TYPE ty_demos.

    METHODS gui_upload
      IMPORTING
        file_name     TYPE string
      RETURNING
        VALUE(result) TYPE xstring
      RAISING
        zcx_excel.

    METHODS load_alv_table
      RAISING
        zcx_excel.

    METHODS on_link_clicked FOR EVENT link_click OF cl_salv_events_table IMPORTING column row.

    METHODS popup_confirm
      IMPORTING
        question TYPE ty_popup_confirm_question
      RAISING
        zcx_excel.

    METHODS read_screen_fields.

    METHODS read_xlsx_from_web_repository
      IMPORTING
        objid         TYPE wwwdata-objid
      RETURNING
        VALUE(result) TYPE xstring
      RAISING
        zcx_excel.

    METHODS screen_1001_pbo_first_time
      RAISING
        cx_salv_data_error
        cx_salv_msg
        cx_salv_not_found
        zcx_excel .

    METHODS write_screen_fields.

    METHODS write_xlsx_to_web_repository
      IMPORTING
        objid    TYPE wwwdata-objid
        text     TYPE wwwdata-text
        xstring  TYPE xstring
        filename TYPE clike
      RAISING
        zcx_excel.

    DATA: ref_sscrfields     TYPE REF TO sscrfields,
          p_path             TYPE zexcel_export_dir,
          splitter           TYPE REF TO cl_gui_splitter_container,
          alv_container      TYPE REF TO cl_gui_container,
          zip_diff_container TYPE REF TO cl_gui_container,
          viewer             TYPE REF TO lcl_xlsx_diff_viewer,
          salv               TYPE REF TO cl_salv_table,
          alv_table          TYPE ty_alv_table,
          lv_filesep         TYPE c LENGTH 1.

ENDCLASS.


CLASS lcx_zip_diff IMPLEMENTATION.
  METHOD constructor.
    super->constructor( textid = textid previous = previous ).
    me->text = text.
  ENDMETHOD.
  METHOD get_text.
    result = text.
  ENDMETHOD.
  METHOD get_longtext.
    result = get_text( ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_zip_cleanup_for_diff IMPLEMENTATION.


  METHOD run.

    TYPES : BEGIN OF ty_local_file_header,
              local_file_header_signature TYPE x LENGTH 4,  " 04034b50
              version_needed_to_extract   TYPE x LENGTH 2,
              general_purpose_bit_flag    TYPE x LENGTH 2,
              compression_method          TYPE x LENGTH 2,
              last_mod_file_time          TYPE int2,
              last_mod_file_date          TYPE int2,
              crc_32                      TYPE x LENGTH 4,
              compressed_size             TYPE i,
              uncompressed_size           TYPE i,
              file_name_length            TYPE int2,
              extra_field_length          TYPE int2,
              " file name (variable size)
              " extra field (variable size)
            END OF ty_local_file_header,
            BEGIN OF ty_central_file_header,
              central_file_header_signature TYPE x LENGTH 4, " 02014b50
              version_made_by               TYPE x LENGTH 2,
              version_needed_to_extract     TYPE x LENGTH 2,
              general_purpose_bit_flag      TYPE x LENGTH 2,
              compression_method            TYPE x LENGTH 2,
              last_mod_file_time            TYPE int2,
              last_mod_file_date            TYPE int2,
              crc_32                        TYPE x LENGTH 4,
              compressed_size               TYPE i,
              uncompressed_size             TYPE i,
              file_name_length              TYPE int2, " field 12
              extra_field_length            TYPE int2, " field 13
              file_comment_length           TYPE int2, " field 14
              disk_number_start             TYPE int2,
              internal_file_attributes      TYPE x LENGTH 2,
              external_file_attributes      TYPE x LENGTH 4,
              rel_offset_of_local_header    TYPE x LENGTH 4,
              " file name                       (variable size defined in 12)
              " extra field                     (variable size defined in 13)
              " file comment                    (variable size defined in 14)
            END OF ty_central_file_header,
            BEGIN OF ty_end_of_central_dir,
              signature                      TYPE x LENGTH 4, " 0x06054b50
              number_of_this_disk            TYPE int2,
              disk_num_start_of_central_dir  TYPE int2,
              n_of_entries_in_central_dir_dk TYPE int2,
              n_of_entries_in_central_dir    TYPE int2,
              size_of_central_dir            TYPE i,
              offset_start_of_central_dir    TYPE i,
              file_comment_length            TYPE int2,
            END OF ty_end_of_central_dir.

    FIELD-SYMBOLS:
      <local_file_header_x>   TYPE x,
      <central_file_header_x> TYPE x,
      <end_of_central_dir_x>  TYPE x,
      <local_file_header>     TYPE ty_local_file_header,
      <central_file_header>   TYPE ty_central_file_header,
      <end_of_central_dir>    TYPE ty_end_of_central_dir.
    CONSTANTS:
      local_file_header_signature   TYPE x LENGTH 4 VALUE '504B0304',
      central_file_header_signature TYPE x LENGTH 4 VALUE '504B0102',
      end_of_central_dir_signature  TYPE x LENGTH 4 VALUE '504B0506'.
    DATA:
      dummy_local_file_header   TYPE ty_local_file_header,
      dummy_central_file_header TYPE ty_central_file_header,
      dummy_end_of_central_dir  TYPE ty_end_of_central_dir,
      local_file_header         TYPE ty_zip_structure,
      central_file_header       TYPE ty_zip_structure,
      end_of_central_dir        TYPE ty_zip_structure,
      offset                    TYPE i,
      max_offset                TYPE i.



    local_file_header = init_structure( length = 30 charset_bit = 60 structure = dummy_local_file_header ).
    ASSIGN local_file_header-ref_to_structure->* TO <local_file_header>.
    ASSIGN local_file_header-ref_to_x->* TO <local_file_header_x>.

    central_file_header = init_structure( length = 46 charset_bit = 76 structure = dummy_central_file_header ).
    ASSIGN central_file_header-ref_to_structure->* TO <central_file_header>.
    ASSIGN central_file_header-ref_to_x->* TO <central_file_header_x>.

    end_of_central_dir = init_structure( length = 22 charset_bit = 0 structure = dummy_end_of_central_dir ).
    ASSIGN end_of_central_dir-ref_to_structure->* TO <end_of_central_dir>.
    ASSIGN end_of_central_dir-ref_to_x->* TO <end_of_central_dir_x>.

    result = zip_xstring.

    offset = 0.
    max_offset = xstrlen( result ) - 4.
    WHILE offset <= max_offset.

      CASE result+offset(4).

        WHEN local_file_header_signature.

          read_zip( EXPORTING zip_xstring = result offset = offset CHANGING zip_structure = local_file_header ).

          CLEAR <local_file_header>-last_mod_file_date.
          CLEAR <local_file_header>-last_mod_file_time.

          write_zip( EXPORTING offset = offset CHANGING zip_structure = local_file_header zip_xstring = result ).

          offset = offset + local_file_header-length + <local_file_header>-file_name_length + <local_file_header>-extra_field_length + <local_file_header>-compressed_size.

        WHEN central_file_header_signature.

          read_zip( EXPORTING zip_xstring = result offset = offset CHANGING zip_structure = central_file_header ).

          CLEAR <central_file_header>-last_mod_file_date.
          CLEAR <central_file_header>-last_mod_file_time.

          write_zip( EXPORTING offset = offset CHANGING zip_structure = central_file_header zip_xstring = result ).

          offset = offset + central_file_header-length + <central_file_header>-file_name_length + <central_file_header>-extra_field_length + <central_file_header>-file_comment_length.

        WHEN end_of_central_dir_signature.

          read_zip( EXPORTING zip_xstring = result offset = offset CHANGING zip_structure = end_of_central_dir ).

          offset = offset + end_of_central_dir-length + <end_of_central_dir>-file_comment_length.

        WHEN OTHERS.
          RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'Invalid ZIP file'.

      ENDCASE.

    ENDWHILE.

  ENDMETHOD.


  METHOD init_structure.

    DATA:
      offset      TYPE i,
      rtts_struct TYPE REF TO cl_abap_structdescr.
    FIELD-SYMBOLS:
      <component> TYPE abap_compdescr.

    CREATE DATA result-ref_to_structure LIKE structure.
    result-length = length.
    result-charset_bit = charset_bit.
    CREATE DATA result-ref_to_x TYPE x LENGTH length.

    result-view = cl_abap_view_offlen=>create( ).
    offset = 0.
    rtts_struct ?= cl_abap_typedescr=>describe_by_data( structure ).
    LOOP AT rtts_struct->components ASSIGNING <component>.
      result-view->append( off = offset len = <component>-length ).
      offset = offset + <component>-length.
    ENDLOOP.

  ENDMETHOD.


  METHOD read_zip.

    DATA:
      charset TYPE i.
    FIELD-SYMBOLS:
      <zip_structure_x> TYPE x,
      <zip_structure>   TYPE any.

    ASSIGN zip_structure-ref_to_x->* TO <zip_structure_x>.
    ASSIGN zip_structure-ref_to_structure->* TO <zip_structure>.

    <zip_structure_x> = zip_xstring+offset.

    IF zip_structure-charset_bit >= 1.
      GET BIT zip_structure-charset_bit OF <zip_structure_x> INTO charset.
    ENDIF.

    IF charset = 0.
      IF zip_structure-conv_in_ibm437 IS NOT BOUND.
        zip_structure-conv_in_ibm437 = cl_abap_conv_in_ce=>create(
                  encoding = '1107'
                  endian = 'L' ).
      ENDIF.
      zip_structure-conv_in_ibm437->convert_struc(
            EXPORTING input = <zip_structure_x>
                      view = zip_structure-view
            IMPORTING data = <zip_structure> ).
    ELSE.
      IF zip_structure-conv_in_utf8 IS NOT BOUND.
        zip_structure-conv_in_utf8 = cl_abap_conv_in_ce=>create(
                  encoding = '4110'
                  endian = 'L' ).
      ENDIF.
      zip_structure-conv_in_utf8->convert_struc(
            EXPORTING input = <zip_structure_x>
                      view = zip_structure-view
            IMPORTING data = <zip_structure> ).
    ENDIF.

  ENDMETHOD.


  METHOD write_zip.

    DATA:
      charset TYPE i.
    FIELD-SYMBOLS:
      <zip_structure_x> TYPE x,
      <zip_structure>   TYPE any.

    ASSIGN zip_structure-ref_to_x->* TO <zip_structure_x>.
    ASSIGN zip_structure-ref_to_structure->* TO <zip_structure>.

    IF zip_structure-charset_bit >= 1.
      GET BIT zip_structure-charset_bit OF <zip_structure_x> INTO charset.
    ENDIF.

    IF charset = 0.
      IF zip_structure-conv_out_ibm437 IS NOT BOUND.
        zip_structure-conv_out_ibm437 = cl_abap_conv_out_ce=>create(
                  encoding = '1107'
                  endian = 'L' ).
      ENDIF.
      zip_structure-conv_out_ibm437->convert_struc(
            EXPORTING data = <zip_structure>
                      view = zip_structure-view
            IMPORTING buffer = <zip_structure_x> ).
    ELSE.
      IF zip_structure-conv_out_utf8 IS NOT BOUND.
        zip_structure-conv_out_utf8 = cl_abap_conv_out_ce=>create(
                  encoding = '4110'
                  endian = 'L' ).
      ENDIF.
      zip_structure-conv_out_utf8->convert_struc(
            EXPORTING data = <zip_structure>
                      view = zip_structure-view
            IMPORTING buffer = <zip_structure_x> ).
    ENDIF.

    REPLACE SECTION OFFSET offset LENGTH zip_structure-length OF zip_xstring WITH <zip_structure_x> IN BYTE MODE.

  ENDMETHOD.


ENDCLASS.


CLASS lcl_xlsx_cleanup_for_diff IMPLEMENTATION.

  METHOD run.

    TYPES: BEGIN OF ty_docprops_core,
             creator          TYPE string,
             description      TYPE string,
             last_modified_by TYPE string,
             created          TYPE string,
             modified         TYPE string,
           END OF ty_docprops_core.
    TYPES: BEGIN OF ty_file,
             name    TYPE string,
             content TYPE xstring,
           END OF ty_file.
    DATA: zip                  TYPE REF TO cl_abap_zip,
          content              TYPE xstring,
          docprops_core        TYPE ty_docprops_core,
          ls_file              TYPE ty_file,
          lt_file              TYPE TABLE OF ty_file,
          lo_ixml              TYPE REF TO if_ixml,
          lo_streamfactory     TYPE REF TO if_ixml_stream_factory,
          lo_istream           TYPE REF TO if_ixml_istream,
          lo_parser            TYPE REF TO if_ixml_parser,
          lo_renderer          TYPE REF TO if_ixml_renderer,
          lo_ostream           TYPE REF TO if_ixml_ostream,
          lo_document          TYPE REF TO if_ixml_document,
          lo_element           TYPE REF TO if_ixml_element,
          lo_filter            TYPE REF TO if_ixml_node_filter,
          lo_iterator          TYPE REF TO if_ixml_node_iterator,
          zip_cleanup_for_diff TYPE REF TO lcl_zip_cleanup_for_diff.
    FIELD-SYMBOLS:
      <file>     TYPE cl_abap_zip=>t_file,
      <ls_file2> TYPE ty_file.

    CREATE OBJECT zip.
    zip->load(
      EXPORTING
        zip             = xstring
      EXCEPTIONS
        zip_parse_error = 1
        OTHERS          = 2 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'zip load'.
    ENDIF.

    zip->get(
      EXPORTING
        name                    = 'docProps/core.xml'
      IMPORTING
        content                 = content
      EXCEPTIONS
        zip_index_error         = 1
        zip_decompression_error = 2
        OTHERS                  = 3 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'docProps/core.xml not found'.
    ENDIF.

    CALL TRANSFORMATION zexcel_tr_docprops_core SOURCE XML content RESULT root = docprops_core.

    CLEAR: docprops_core-creator,
           docprops_core-description,
           docprops_core-created,
           docprops_core-modified.

    CALL TRANSFORMATION zexcel_tr_docprops_core SOURCE root = docprops_core RESULT XML content.

    zip->delete(
      EXPORTING
        name            = 'docProps/core.xml'
      EXCEPTIONS
        zip_index_error = 1
        OTHERS          = 2 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |delete before add of docProps/core.xml|.
    ENDIF.

    zip->add(
        name    = 'docProps/core.xml'
        content = content ).

    LOOP AT zip->files ASSIGNING <file>
        WHERE name CP 'xl/drawings/drawing*.xml'.

      zip->get(
        EXPORTING
          name                    = <file>-name
        IMPORTING
          content                 = content
        EXCEPTIONS
          zip_index_error         = 1
          zip_decompression_error = 2
          OTHERS                  = 3 ).
      IF sy-subrc <> 0.
        RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |{ <file>-name } not found|.
      ENDIF.

      lo_ixml = cl_ixml=>create( ).
      lo_streamfactory = lo_ixml->create_stream_factory( ).
      lo_istream = lo_streamfactory->create_istream_xstring( content ).
      lo_document = lo_ixml->create_document( ).
      lo_parser = lo_ixml->create_parser(
                  document       = lo_document
                  istream        = lo_istream
                  stream_factory = lo_streamfactory ).
      lo_parser->parse( ).

      lo_filter = lo_document->create_filter_name_ns(
                  name      = 'cNvPr'
                  namespace = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' ).
      lo_iterator = lo_document->create_iterator_filtered( lo_filter ).
      DO.
        lo_element ?= lo_iterator->get_next( ).
        IF lo_element IS NOT BOUND.
          EXIT.
        ENDIF.
        lo_element->set_attribute_ns( name = 'name' value = '' ).
      ENDDO.

      CLEAR content.
      lo_ostream = lo_streamfactory->create_ostream_xstring( content ).
      lo_renderer = lo_ixml->create_renderer(
                  document = lo_document
                  ostream  = lo_ostream ).
      lo_renderer->render( ).

      ls_file-name = <file>-name.
      ls_file-content = content.
      APPEND ls_file TO lt_file.

    ENDLOOP.

    LOOP AT lt_file ASSIGNING <ls_file2>.
      zip->delete( name = <ls_file2>-name ).
      zip->add( name = <ls_file2>-name content = <ls_file2>-content ).
    ENDLOOP.

    result = zip->save( ).

    CREATE OBJECT zip_cleanup_for_diff.
    result = zip_cleanup_for_diff->run( result ).

  ENDMETHOD.


ENDCLASS.


CLASS lcl_xlsx_diff_item IMPLEMENTATION.


  METHOD compare_items_at_path.

    DATA: diff_item     TYPE ty_diff_item,
          folder_diff   TYPE REF TO lcl_xlsx_diff_item,
          item_1        TYPE REF TO lcl_xlsx_diff_item=>ty_item,
          item_2        TYPE REF TO lcl_xlsx_diff_item=>ty_item,
          zip_1_items   TYPE lcl_xlsx_diff_item=>ty_items,
          zip_2_items   TYPE lcl_xlsx_diff_item=>ty_items,
          zip_1_index   TYPE i,
          zip_2_index   TYPE i,
          zip_1_read    TYPE abap_bool,
          zip_2_read    TYPE abap_bool,
          diff_state    TYPE lcl_xlsx_diff_item=>ty_diff_state,
          temp_string_1 TYPE string,
          temp_string_2 TYPE string,
          item          TYPE REF TO lcl_xlsx_diff_item=>ty_item.

    CREATE OBJECT result TYPE lcl_xlsx_diff_item.
    result->path = path.

    zip_1_items = get_items_at_path( items = zip_1->files path = path ).
    zip_2_items = get_items_at_path( items = zip_2->files path = path ).

    zip_1_index = 0.
    zip_2_index = 0.
    zip_1_read = abap_true.
    zip_2_read = abap_true.

    WHILE zip_1_read = abap_true OR zip_2_read = abap_true.

      IF zip_1_read = abap_true.
        zip_1_index = zip_1_index + 1.
        CLEAR item_1.
        IF zip_1_index > lines( zip_1_items ).
          zip_1_read = abap_false.
        ELSE.
*          item_1 = REF #( zip_1_items[ zip_1_index ] OPTIONAL ).
          READ TABLE zip_1_items INDEX zip_1_index REFERENCE INTO item_1.
        ENDIF.
      ENDIF.
      IF zip_2_read = abap_true.
        zip_2_index = zip_2_index + 1.
        CLEAR item_2.
        IF zip_2_index > lines( zip_2_items ).
          zip_2_read = abap_false.
        ELSE.
*          item_2 = REF #( zip_2_items[ zip_2_index ] OPTIONAL ).
          READ TABLE zip_2_items INDEX zip_2_index REFERENCE INTO item_2.
        ENDIF.
      ENDIF.

      CLEAR folder_diff.
      diff_state = state-undefined.
      zip_1_read = abap_false.
      zip_2_read = abap_false.
      IF zip_1_index <= lines( zip_1_items ) AND zip_2_index <= lines( zip_2_items ).
        IF item_1->local_name < item_2->local_name.
          diff_state = state-deleted.
          zip_1_read = abap_true.
        ELSEIF item_1->local_name > item_2->local_name.
          diff_state = state-added.
          zip_2_read = abap_true.
        ELSE.
          " same item -> compare the contents of the 2 items
          zip_1_read = abap_true.
          zip_2_read = abap_true.
          IF item_1->local_name CP '*/'.
*          IF contains( val = item_1->local_name end = '/' ).
            " folder
            temp_string_1 = |{ path }{ item_1->local_name }|.
            folder_diff = compare_items_at_path( path = temp_string_1 ).
            LOOP AT folder_diff->items TRANSPORTING NO FIELDS WHERE diff_state <> state-same.
              EXIT.
            ENDLOOP.
            IF sy-subrc = 0.
              diff_state = state-changed.
            ELSE.
              diff_state = state-same.
            ENDIF.
          ELSE.
            temp_string_1 = |{ path }{ item_1->local_name }|.
            temp_string_2 = |{ path }{ item_2->local_name }|.
            diff_state = compare_two_items( zip_1 = zip_1 zip_2 = zip_2 name_1 = temp_string_1 name_2 = temp_string_2 ).
          ENDIF.
        ENDIF.
      ELSEIF zip_1_index <= lines( zip_1_items ).
        diff_state = state-deleted.
        zip_1_read = abap_true.
        IF item_1->local_name CP '*/'.
*        IF contains( val = item_1->local_name end = '/' ).
          " folder
          temp_string_1 = |{ path }{ item_1->local_name }|.
          folder_diff = compare_items_at_path( path = temp_string_1 ).
        ENDIF.
      ELSEIF zip_2_index <= lines( zip_2_items ).
        diff_state = state-added.
        zip_2_read = abap_true.
        IF item_2->local_name CP '*/'.
*        IF contains( val = item_2->local_name end = '/' ).
          " folder
          temp_string_2 = |{ path }{ item_2->local_name }|.
          folder_diff = compare_items_at_path( path = temp_string_2 ).
        ENDIF.
      ELSE.
        " no more file to compare
        EXIT.
      ENDIF.

      CLEAR diff_item.
      IF diff_state = state-added.
        item = item_2.
      ELSE.
        item = item_1.
        diff_item-attr_1-date = item_1->date.
        diff_item-attr_1-time = item_1->time.
        diff_item-attr_1-size = item_1->size.
      ENDIF.
      IF diff_state <> state-deleted.
        diff_item-attr_2-date = item_2->date.
        diff_item-attr_2-time = item_2->time.
        diff_item-attr_2-size = item_2->size.
      ENDIF.
      diff_item-local_name  = item->local_name.
      diff_item-full_path   = path && item->local_name.
      diff_item-diff_state  = diff_state.
      diff_item-folder_diff = folder_diff.
*      diff_item = VALUE #(
*          LET item = SWITCH #( diff_state WHEN state-added THEN item_2 ELSE item_1 ) IN
*          local_name  = item->local_name
*          full_path   = path && item->local_name
*          attr_1      = COND #( WHEN diff_state <> state-added THEN VALUE #(
*                          date = item_1->date
*                          time = item_1->time
*                          size = item_1->size ) )
*          attr_2      = COND #( WHEN diff_state <> state-deleted THEN VALUE #(
*                          date = item_2->date
*                          time = item_2->time
*                          size = item_2->size ) )
*          diff_state  = diff_state
*          folder_diff = folder_diff ).
      INSERT diff_item INTO TABLE result->items.

    ENDWHILE.

    LOOP AT result->items TRANSPORTING NO FIELDS WHERE diff_state <> state-same.
      EXIT.
    ENDLOOP.
    result->is_different = boolc( sy-subrc = 0 ).

  ENDMETHOD.


  METHOD compare_two_items.
    DATA: content_1 TYPE xstring,
          content_2 TYPE xstring.

    zip_1->get(
      EXPORTING
        name                    = name_1
      IMPORTING
        content                 = content_1
      EXCEPTIONS
        zip_index_error         = 1
        zip_decompression_error = 2
        OTHERS                  = 3 ).
    IF sy-subrc <> 0.
      " todo
    ENDIF.
    zip_2->get(
      EXPORTING
        name                    = name_2
      IMPORTING
        content                 = content_2
      EXCEPTIONS
        zip_index_error         = 1
        zip_decompression_error = 2
        OTHERS                  = 3 ).
    IF sy-subrc <> 0.
      " todo
    ENDIF.

    IF content_1 <> content_2.
      result = state-only_content_changed.
    ELSE.
      result = state-same.
    ENDIF.
*    result = COND #( WHEN content_1 <> content_2 THEN state-only_content_changed ELSE state-same ).

  ENDMETHOD.


  METHOD get_diff.
    DATA: engine TYPE REF TO lcl_xlsx_diff_item.

    CREATE OBJECT engine TYPE lcl_xlsx_diff_item.
    engine->zip_1 = zip_1.
    engine->zip_2 = zip_2.
    result = engine->compare_items_at_path( ).

  ENDMETHOD.


  METHOD get_items_at_path.

    DATA: starter        TYPE string,
          item_or_folder TYPE string VALUE '^[^/]*(?:/|$)',
          local_name     TYPE string,
          path_item      TYPE lcl_xlsx_diff_item=>ty_item.
    FIELD-SYMBOLS: <item> TYPE cl_abap_zip=>t_file.

    CLEAR path_items.

    LOOP AT items ASSIGNING <item>.
      IF path = ``.
        local_name = substring_to( val = <item>-name regex = item_or_folder ).
      ELSEIF strlen( <item>-name ) >= strlen( path ) AND substring( val = <item>-name len = strlen( path ) ) = path.
*      ELSEIF contains( val = <item>-name start = path ).
        local_name = substring_to( val = substring( val = <item>-name off = strlen( path ) ) regex = item_or_folder ).
      ENDIF.
*      local_name = COND #( WHEN path = `` THEN substring_to( val = <item>-name regex = item_or_folder )
*                           WHEN contains( val = <item>-name start = path ) THEN substring_to( val = substring( val = <item>-name off = strlen( path ) ) regex = item_or_folder ) ).
      IF local_name IS NOT INITIAL.
        READ TABLE path_items WITH TABLE KEY local_name = local_name TRANSPORTING NO FIELDS.
        IF sy-subrc <> 0.
*        IF NOT line_exists( path_items[ local_name = local_name ] ).
          CLEAR path_item.
          IF local_name NP '*/'.
*          IF NOT contains( val = local_name end = '/' ).
            path_item-local_name = local_name.
            path_item-full_path  = <item>-name.
            path_item-date       = <item>-date.
            path_item-time       = <item>-time.
            path_item-size       = <item>-size.
          ELSE.
            path_item-local_name = local_name.
            path_item-full_path  = path && local_name.
          ENDIF.
          INSERT path_item INTO TABLE path_items.
*          INSERT COND ty_item(
*                WHEN NOT contains( val = local_name end = '/' ) THEN VALUE #(
*                  local_name = local_name
*                  full_path  = <item>-name
*                  date       = <item>-date
*                  time       = <item>-time
*                  size       = <item>-size )
*                ELSE VALUE #(
*                  local_name = local_name
*                  full_path  = path && local_name ) )
*              INTO TABLE path_items.
        ENDIF.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


ENDCLASS.


CLASS lcl_xlsx_diff_viewer IMPLEMENTATION.


  METHOD constructor.

    me->container = io_container.
    CREATE OBJECT diff_file TYPE lcl_xlsx_diff_file_ext.
    SET HANDLER diff_file->on_selection_changed FOR me.

  ENDMETHOD.


  METHOD diff_and_view.
    DATA: zip_diff TYPE REF TO lcl_xlsx_diff_item.

    me->title_old = title_old.
    me->title_new = title_new.
    me->zip_old = zip_old.
    me->zip_new = zip_new.

    zip_diff = lcl_xlsx_diff_item=>get_diff( zip_1 = zip_old zip_2 = zip_new ).

    view( zip_diff ).

  ENDMETHOD.


  METHOD view.

    DATA: ls_hierarchy_header TYPE treemhhdr,
          event               TYPE cntl_simple_event,
          events              TYPE cntl_simple_events,
          isfolder            TYPE as4flag,
          item_table          TYPE treemcitab,
          item_line           TYPE treemcitem,
          temp_container      TYPE REF TO cl_gui_container.
    CONSTANTS: cntb_btype_button TYPE i VALUE 0.

    IF go_tree IS NOT BOUND.
      ls_hierarchy_header-heading = 'ZIP Hierarchy'(001).

      CREATE OBJECT go_splitter
        EXPORTING
          parent  = container
          rows    = 2
          columns = 1.
      go_splitter->set_row_height( id = 1 height = 5 ).
      go_splitter->set_row_sash( id = 1 type = go_splitter->type_sashvisible value = go_splitter->false ).

      temp_container = go_splitter->get_container( row = 1 column = 1 ).
      CREATE OBJECT go_toolbar
        EXPORTING
          parent       = temp_container
          display_mode = cl_gui_toolbar=>m_mode_horizontal.
      go_toolbar->add_button(
                fcode       = 'TITLE_OLD'
                icon        = icon_old
                butn_type   = cntb_btype_button
                text        = title_old ).
      go_toolbar->add_button(
                fcode       = 'TITLE_NEW'
                icon        = icon_new
                butn_type   = cntb_btype_button
                text        = title_new ).

      CREATE OBJECT go_tree
        EXPORTING
          node_selection_mode   = cl_gui_simple_tree=>node_sel_mode_single
          hierarchy_column_name = 'C1'
          hierarchy_header      = ls_hierarchy_header.

      go_tree->add_column(
            name         = 'DATE_1'
            width        = 0
            header_image = icon_old
            header_text  = 'Date' ).

      go_tree->add_column(
            name         = 'DATE_2'
            width        = 0
            header_image = icon_new
            header_text  = 'Date' ).

      go_tree->add_column(
            name         = 'TIME_1'
            width        = 0
            header_image = icon_old
            header_text  = 'Time' ).

      go_tree->add_column(
            name         = 'TIME_2'
            width        = 0
            header_image = icon_new
            header_text  = 'Time' ).

      go_tree->add_column(
            name         = 'SIZE_1'
            width        = 0
            header_image = icon_old
            header_text  = 'Size' ).

      go_tree->add_column(
            name         = 'SIZE_2'
            width        = 0
            header_image = icon_new
            header_text  = 'Size' ).

      temp_container = go_splitter->get_container( row = 2 column = 1 ).
      go_tree->create_tree_control( parent = temp_container ).

      SET HANDLER on_selection_changed FOR go_tree.
      CLEAR event.
      event-eventid = cl_tree_model=>eventid_selection_changed.
      INSERT event INTO TABLE events.
      go_tree->set_registered_events(
        EXPORTING
          events                    = events
        EXCEPTIONS
          illegal_event_combination = 1
          unknown_event             = 2
          OTHERS                    = 3 ).
      IF sy-subrc <> 0.
*       MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                  WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.

    ELSE.

      go_tree->delete_all_nodes( ).

    ENDIF.

    IF lines( i_zip_diff->items ) >= 1.
      isfolder = abap_true.
    ELSE.
      isfolder = abap_false.
    ENDIF.
    CLEAR item_table.
    CLEAR item_line.
    item_line-item_name = 'C1'.
    item_line-class     = cl_column_tree_model=>item_class_text.
    item_line-text      = 'Root'.
    INSERT item_line INTO TABLE item_table.
    go_tree->add_node(
          node_key   = '1'
          isfolder   = isfolder
          item_table = item_table ).

    node_key = 1.
    add_tree_nodes(
          diff_object     = i_zip_diff
          path            = ''
          parent_node_key = '1' ).

    go_tree->expand_root_nodes(
      EXCEPTIONS
        illegal_level_count = 1
        OTHERS              = 2 ).
    IF sy-subrc <> 0.
*     MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    go_tree->adjust_column_width(
      EXPORTING
        all_columns               = abap_true
        include_heading           = abap_true
      EXCEPTIONS
        control_not_existing      = 1
        control_dead              = 2
        cntl_system_error         = 3
        failed                    = 4
        start_column_not_found    = 5
        end_column_not_found      = 6
        start_column_in_hierarchy = 7
        end_column_in_hierarchy   = 8
        start_column_empty        = 9
        OTHERS                    = 10 ).
    IF sy-subrc <> 0.
*       MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                  WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

  ENDMETHOD.


  METHOD add_tree_nodes.

    FIELD-SYMBOLS:
          <sub_files> TYPE lcl_xlsx_diff_item=>ty_diff_items.
    DATA: lt_column          TYPE treemcitab,
          diff_item          TYPE REF TO lcl_xlsx_diff_item=>ty_diff_item,
          save_node_key      TYPE i,
          item_line          TYPE treemcitem,
          temp_node_key      TYPE string,
          temp_string        TYPE string,
          temp_save_node_key TYPE string,
          isfolder           TYPE as4flag.

    LOOP AT diff_object->items REFERENCE INTO diff_item.

      CLEAR lt_column.

      CLEAR item_line.
      item_line-item_name = 'C1'.
      item_line-class     = cl_column_tree_model=>item_class_text.
      item_line-text      = diff_item->local_name.
      CASE diff_item->diff_state.
        WHEN state-deleted.
          item_line-style = cl_tree_control_base=>style_emphasized_negative.
        WHEN state-added.
          item_line-style = cl_tree_control_base=>style_emphasized_positive.
        WHEN state-changed
            OR state-only_attribute_changed
            OR state-only_content_changed.
          IF diff_item->folder_diff IS NOT BOUND.
            item_line-style = cl_tree_control_base=>style_emphasized.
          ELSE.
            item_line-style = cl_tree_control_base=>style_intensified.
          ENDIF.
        WHEN OTHERS.
          item_line-style = cl_tree_control_base=>style_default.
      ENDCASE.
      INSERT item_line INTO TABLE lt_column.

      CLEAR item_line.
      item_line-item_name = 'DATE_1'.
      item_line-class = cl_column_tree_model=>item_class_text.
      IF diff_item->attr_1-date IS NOT INITIAL.
        item_line-text = |{ diff_item->attr_1-date DATE = USER }|.
      ENDIF.
      INSERT item_line INTO TABLE lt_column.

      CLEAR item_line.
      item_line-item_name = 'DATE_2'.
      item_line-class     = cl_column_tree_model=>item_class_text.
      IF diff_item->attr_2-date IS NOT INITIAL.
        item_line-text = |{ diff_item->attr_2-date DATE = USER }|.
      ENDIF.
      INSERT item_line INTO TABLE lt_column.

      CLEAR item_line.
      item_line-item_name = 'TIME_1'.
      item_line-class     = cl_column_tree_model=>item_class_text.
      IF diff_item->attr_1-time IS NOT INITIAL.
        item_line-text = |{ diff_item->attr_1-time TIME = USER }|.
      ENDIF.
      INSERT item_line INTO TABLE lt_column.

      CLEAR item_line.
      item_line-item_name = 'TIME_2'.
      item_line-class     = cl_column_tree_model=>item_class_text.
      IF diff_item->attr_2-time IS NOT INITIAL.
        item_line-text = |{ diff_item->attr_2-time TIME = USER }|.
      ENDIF.
      INSERT item_line INTO TABLE lt_column.

      CLEAR item_line.
      item_line-item_name = 'SIZE_1'.
      item_line-class     = cl_column_tree_model=>item_class_text.
      item_line-text      = |{ diff_item->attr_1-size }|.
      INSERT item_line INTO TABLE lt_column.

      CLEAR item_line.
      item_line-item_name = 'SIZE_2'.
      item_line-class     = cl_column_tree_model=>item_class_text.
      item_line-text      = |{ diff_item->attr_2-size }|.
      INSERT item_line INTO TABLE lt_column.

*      lt_column = VALUE #(
*      ( item_name = 'C1'
*        class     = cl_column_tree_model=>item_class_text
*        text      = diff_item->local_name
*        style     = SWITCH #( diff_item->diff_state
*                      WHEN state-deleted THEN cl_tree_control_base=>style_emphasized_negative
*                      WHEN state-added THEN cl_tree_control_base=>style_emphasized_positive
*                      WHEN state-changed
*                        OR state-only_attribute_changed
*                        OR state-only_content_changed
*                        THEN COND #( WHEN diff_item->folder_diff IS NOT BOUND
*                                               THEN cl_tree_control_base=>style_emphasized
*                                               ELSE cl_tree_control_base=>style_intensified )
*                      ELSE cl_tree_control_base=>style_default ) )
*      ( item_name = 'DATE_1'
*        class     = cl_column_tree_model=>item_class_text
*        text      = COND #( WHEN diff_item->attr_1-date IS NOT INITIAL THEN |{ diff_item->attr_1-date DATE = USER }| ) )
*      ( item_name = 'DATE_2'
*        class     = cl_column_tree_model=>item_class_text
*        text      = COND #( WHEN diff_item->attr_2-date IS NOT INITIAL THEN |{ diff_item->attr_2-date DATE = USER }| ) )
*      ( item_name = 'TIME_1'
*        class     = cl_column_tree_model=>item_class_text
*        text      = COND #( WHEN diff_item->attr_2-time IS NOT INITIAL THEN |{ diff_item->attr_1-time TIME = USER }| ) )
*      ( item_name = 'TIME_2'
*        class     = cl_column_tree_model=>item_class_text
*        text      = COND #( WHEN diff_item->attr_2-time IS NOT INITIAL THEN |{ diff_item->attr_2-time TIME = USER }| ) )
*      ( item_name = 'SIZE_1'
*        class     = cl_column_tree_model=>item_class_text
*        text      = |{ diff_item->attr_1-size }| )
*      ( item_name = 'SIZE_2'
*        class     = cl_column_tree_model=>item_class_text
*        text      = |{ diff_item->attr_2-size }| ) ).

      ADD 1 TO node_key.
      save_node_key = node_key.
      temp_node_key = |{ node_key }|.

      IF diff_item->folder_diff IS BOUND.
        isfolder = abap_true.
      ELSE.
        isfolder = abap_false.
      ENDIF.
      go_tree->add_node(
            node_key          = temp_node_key
            relationship      = cl_column_tree_model=>relat_first_child
            relative_node_key = parent_node_key
            isfolder          = isfolder
            item_table        = lt_column
            user_object       = diff_object ).

      IF diff_item->folder_diff IS BOUND
            AND lines( diff_item->folder_diff->items ) > 0.
        temp_string = path && diff_item->local_name.
        add_tree_nodes(
            diff_object     = diff_item->folder_diff
            path            = temp_string
            parent_node_key = temp_node_key ).
        IF diff_item->diff_state <> state-same.
          temp_save_node_key = |{ save_node_key }|.
          go_tree->expand_node(
            EXPORTING
              node_key       = temp_save_node_key
            EXCEPTIONS
              node_not_found = 1
              OTHERS         = 2 ).
          IF sy-subrc <> 0.
*     MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
          ENDIF.
        ENDIF.
      ENDIF.

    ENDLOOP.

  ENDMETHOD.


  METHOD on_selection_changed.
    DATA: item        TYPE treemcitem,
          user_object TYPE REF TO object,
          diff_item   TYPE REF TO lcl_xlsx_diff_item,
          item_2      TYPE lcl_xlsx_diff_item=>ty_diff_item.

    go_tree->node_get_item(
      EXPORTING
        node_key       = node_key
        item_name      = 'C1'
      IMPORTING
        item           = item
      EXCEPTIONS
        node_not_found = 1
        item_not_found = 2
        OTHERS         = 3 ).
    IF sy-subrc <> 0.
*     MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
    go_tree->node_get_user_object(
      EXPORTING
        node_key       = node_key
      IMPORTING
        user_object    = user_object
      EXCEPTIONS
        node_not_found = 1
        OTHERS         = 2
    ).
    IF sy-subrc <> 0.
*     MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    IF user_object IS BOUND.

      diff_item ?= user_object.

      READ TABLE diff_item->items WITH KEY local_name = item-text INTO item_2.
      RAISE EVENT selection_changed
          EXPORTING
              node    = item_2
*              node    = diff_item->items[ local_name = item-text ]
              zip_old = zip_old
              zip_new = zip_new.

    ENDIF.

  ENDMETHOD.


ENDCLASS.


CLASS lcl_xlsx_diff_file_ext IMPLEMENTATION.


  METHOD class_constructor.

    CALL METHOD cl_gui_frontend_services=>get_temp_directory
      CHANGING
        temp_dir     = temp_dir
      EXCEPTIONS
        cntl_error   = 1
        error_no_gui = 2.
    IF sy-subrc <> 0.
      " Error handling
    ENDIF.

  ENDMETHOD.


  METHOD on_selection_changed.
    DATA: downloaded_file_new TYPE string,
          downloaded_file_old TYPE string,
          error               TYPE REF TO cx_root.

    TRY.

        CASE node-diff_state.
          WHEN state-added.

            downloaded_file_new = zip_get_and_gui_download( zip = zip_new name = node-full_path prefix = 'added' ).

            gui_execute( |"{ downloaded_file_new }"| ).

          WHEN state-changed
            OR state-only_attribute_changed
            OR state-only_content_changed.

            downloaded_file_old = zip_get_and_gui_download( zip = zip_old name = node-full_path prefix = 'old' ).
            downloaded_file_new = zip_get_and_gui_download( zip = zip_new name = node-full_path prefix = 'new' ).

            gui_execute( |-d "{ downloaded_file_old }" "{ downloaded_file_new }"| ).

          WHEN state-deleted.

            downloaded_file_old = zip_get_and_gui_download( zip = zip_old name = node-full_path prefix = 'deleted' ).

            gui_execute( |"{ downloaded_file_old }"| ).

          WHEN state-same.

            downloaded_file_old = zip_get_and_gui_download( zip = zip_old name = node-full_path prefix = 'same' ).

            gui_execute( |"{ downloaded_file_old }"| ).

        ENDCASE.

      CATCH cx_root INTO error.
        MESSAGE error TYPE 'I' DISPLAY LIKE 'E'.
    ENDTRY.

  ENDMETHOD.


  METHOD gui_download.
    DATA: solix_tab TYPE solix_tab.

    solix_tab = cl_bcs_convert=>xstring_to_solix( i_content ).

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        bin_filesize              = xstrlen( i_content )
        filename                  = i_file_path
        filetype                  = 'BIN'
      CHANGING
        data_tab                  = solix_tab
      EXCEPTIONS
        file_write_error          = 1
        no_batch                  = 2
        gui_refuse_filetransfer   = 3
        invalid_type              = 4
        no_authority              = 5
        unknown_error             = 6
        header_not_allowed        = 7
        separator_not_allowed     = 8
        filesize_not_allowed      = 9
        header_too_long           = 10
        dp_error_create           = 11
        dp_error_send             = 12
        dp_error_write            = 13
        unknown_dp_error          = 14
        access_denied             = 15
        dp_out_of_memory          = 16
        disk_full                 = 17
        dp_timeout                = 18
        file_not_found            = 19
        dataprovider_exception    = 20
        control_flush_error       = 21
        not_supported_by_gui      = 22
        error_no_gui              = 23
        OTHERS                    = 24
    ).
    IF sy-subrc <> 0.
*        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                   WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

  ENDMETHOD.


  METHOD xml_pretty_print.

    DATA xml_document TYPE REF TO if_ixml_document.

    CALL FUNCTION 'SDIXML_XML_TO_DOM'
      EXPORTING
        xml      = c_content
      IMPORTING
        document = xml_document
      EXCEPTIONS
        OTHERS   = 1.
    CALL FUNCTION 'SDIXML_DOM_TO_XML'
      EXPORTING
        document      = xml_document
        pretty_print  = abap_true
      IMPORTING
        xml_as_string = c_content
      EXCEPTIONS
        OTHERS        = 2.

  ENDMETHOD.



  METHOD gui_execute.
    DATA: temp_string TYPE string.

    cl_gui_frontend_services=>execute(
        EXPORTING
          application = 'code'
          parameter   = i_parameter
          minimized   = 'X'
          synchronous = ''
        EXCEPTIONS
          cntl_error             = 1
          error_no_gui           = 2
          bad_parameter          = 3
          file_not_found         = 4
          path_not_found         = 5
          file_extension_unknown = 6
          error_execute_failed   = 7
          synchronous_failed     = 8
          not_supported_by_gui   = 9
          OTHERS                 = 10 ).
    IF sy-subrc <> 0.
      temp_string = |cl_gui_frontend_services=>execute { sy-subrc } { i_parameter }|.
      RAISE EXCEPTION TYPE lcx_zip_diff EXPORTING text = temp_string.
    ENDIF.

  ENDMETHOD.


  METHOD zip_get_and_gui_download.
    DATA: content     TYPE xstring,
          parts       TYPE string_table,
          new_parts   TYPE string_table,
          part        TYPE REF TO string,
          temp_string TYPE string.

    zip->get(
        EXPORTING
          name                    = name
        IMPORTING
          content                 = content
        EXCEPTIONS
          zip_index_error         = 1
          zip_decompression_error = 2
          OTHERS                  = 3 ).
    IF sy-subrc <> 0.
      temp_string = |zip->get { sy-subrc } { name }|.
      RAISE EXCEPTION TYPE lcx_zip_diff EXPORTING text = temp_string.
    ENDIF.

    SPLIT name AT '/' INTO TABLE parts.
    CLEAR new_parts.
    LOOP AT parts REFERENCE INTO part.
      INSERT part->* INTO new_parts INDEX 1.
    ENDLOOP.

    result = |{ temp_dir }\\{ prefix }_{ concat_lines_of( table = new_parts sep = '_' ) }.xml|.
    IF name CS '.xml'.
      xml_pretty_print( CHANGING c_content = content ).
    ENDIF.

    gui_download( i_content   = content
                  i_file_path = result ).

  ENDMETHOD.

ENDCLASS.


CLASS lcl_app IMPLEMENTATION.


  METHOD at_selection_screen.

    DATA: error TYPE REF TO cx_root.

    TRY.

        CASE sy-dynnr.

          WHEN 1000.

            CASE ref_sscrfields->ucomm.

              WHEN 'ONLI'.

                read_screen_fields( ).
                SUBMIT zdemo_excel WITH p_path = p_path WITH p_checkr = abap_true AND RETURN.
                CALL SELECTION-SCREEN 1001.

            ENDCASE.

          WHEN 1001.

            CASE ref_sscrfields->ucomm.

              WHEN 'FC01'. " REFRESH

                SUBMIT (sy-repid) WITH p_path = p_path.

            ENDCASE.
        ENDCASE.

      CATCH cx_root INTO error.
        MESSAGE error TYPE 'E'.
    ENDTRY.

  ENDMETHOD.


  METHOD at_selection_screen_on_exit.

    CASE sy-dynnr.

      WHEN 1001.

        CALL SELECTION-SCREEN 1000.

    ENDCASE.

  ENDMETHOD.


  METHOD at_selection_screen_output.

    DATA: error TYPE REF TO cx_root.

    TRY.

        CASE sy-dynnr.

          WHEN 1000.

            at_selection_screen_output1000( ).

          WHEN 1001.

            at_selection_screen_output1001( ).

        ENDCASE.

      CATCH cx_root INTO error.
        MESSAGE error TYPE 'I' DISPLAY LIKE 'E'.
    ENDTRY.

  ENDMETHOD.


  METHOD at_selection_screen_output1000.

    DATA: lv_workdir TYPE string.

    cl_gui_frontend_services=>get_file_separator( CHANGING file_separator = lv_filesep ).

    IF p_path IS INITIAL.

      cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = lv_workdir ).
      cl_gui_cfw=>flush( ).
      p_path = lv_workdir.

    ENDIF.

    write_screen_fields( ).

  ENDMETHOD.


  METHOD at_selection_screen_output1001.

    DATA: excluded_functions TYPE ui_functions.

    LOOP AT SCREEN.
      screen-active = '0'.
      MODIFY SCREEN.
    ENDLOOP.

    ref_sscrfields->functxt_01 = icon_refresh.

    APPEND 'ONLI' TO excluded_functions.
    APPEND 'PRIN' TO excluded_functions.
    APPEND 'SPOS' TO excluded_functions.
    CALL FUNCTION 'RS_SET_SELSCREEN_STATUS'
      EXPORTING
        p_status  = sy-pfkey
      TABLES
        p_exclude = excluded_functions.

    load_alv_table( ).

    IF alv_container IS NOT BOUND.

      screen_1001_pbo_first_time( ).

    ENDIF.

  ENDMETHOD.


  METHOD check_regression.

    DATA: xlsx_cleanup_for_diff TYPE REF TO lcl_xlsx_cleanup_for_diff.


    result-xlsx_just_now = gui_upload( file_name = p_path && lv_filesep && demo-filename ).

    result-xlsx_reference = read_xlsx_from_web_repository( objid = demo-objid ).

    IF result-xlsx_reference IS INITIAL.

      result-diff = abap_true.

    ELSE.

      CREATE OBJECT xlsx_cleanup_for_diff.
      result-compare_xlsx_just_now = xlsx_cleanup_for_diff->run( result-xlsx_just_now ).
      result-compare_xlsx_reference = xlsx_cleanup_for_diff->run( result-xlsx_reference ).

      result-diff = boolc( result-compare_xlsx_just_now <> result-compare_xlsx_reference ).

    ENDIF.

  ENDMETHOD.


  METHOD get_list_of_demo_files.

    DATA: line TYPE ty_demo.

    line-program = 'ZDEMO_EXCEL1'.
    line-objid = 'ZDEMO_EXCEL1'.
    line-filename = '01_HelloWorld.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL2'.
    line-objid = 'ZDEMO_EXCEL2'.
    line-filename = '02_Styles.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL3'.
    line-objid = 'ZDEMO_EXCEL3'.
    line-filename = '03_iTab.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL4'.
    line-objid = 'ZDEMO_EXCEL4'.
    line-filename = '04_Sheets.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL5'.
    line-objid = 'ZDEMO_EXCEL5'.
    line-filename = '05_Conditional.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL6'.
    line-objid = 'ZDEMO_EXCEL6'.
    line-filename = '06_Formulas.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL7'.
    line-objid = 'ZDEMO_EXCEL7'.
    line-filename = '07_ConditionalAll.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL8'.
    line-objid = 'ZDEMO_EXCEL8'.
    line-filename = '08_Range.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL9'.
    line-objid = 'ZDEMO_EXCEL9'.
    line-filename = '09_DataValidation.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL10'.
    line-objid = 'ZDEMO_EXCEL10'.
    line-filename = '10_iTabFieldCatalog.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL12'.
    line-objid = 'ZDEMO_EXCEL12'.
    line-filename = '12_HideSizeOutlineRowsAndColumns.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL13'.
    line-objid = 'ZDEMO_EXCEL13'.
    line-filename = '13_MergedCells.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL14'.
    line-objid = 'ZDEMO_EXCEL14'.
    line-filename = '14_Alignment.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL16'.
    line-objid = 'ZDEMO_EXCEL16'.
    line-filename = '16_Drawings.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL17'.
    line-objid = 'ZDEMO_EXCEL17'.
    line-filename = '17_SheetProtection.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL18'.
    line-objid = 'ZDEMO_EXCEL18'.
    line-filename = '18_BookProtection.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL19'.
    line-objid = 'ZDEMO_EXCEL19'.
    line-filename = '19_SetActiveSheet.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL21'.
    line-objid = 'ZDEMO_EXCEL21'.
    line-filename = '21_BackgroundColorPicker.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL22'.
    line-objid = 'ZDEMO_EXCEL22'.
    line-filename = '22_itab_fieldcatalog.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL23'.
    line-objid = 'ZDEMO_EXCEL23'.
    line-filename = '23_Sheets_with_and_without_grid_lines.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL24'.
    line-objid = 'ZDEMO_EXCEL24'.
    line-filename = '24_Sheets_with_different_default_date_formats.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL27'.
    line-objid = 'ZDEMO_EXCEL27'.
    line-filename = '27_ConditionalFormatting.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL30'.
    line-objid = 'ZDEMO_EXCEL30'.
    line-filename = '30_CellDataTypes.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL31'.
    line-objid = 'ZDEMO_EXCEL31'.
    line-filename = '31_AutosizeWithDifferentFontSizes.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL33'.
    line-objid = 'ZDEMO_EXCEL33'.
    line-filename = '33_autofilter.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL34'.
    line-objid = 'ZDEMO_EXCEL34'.
    line-filename = '34_Static Styles_Chess.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL35'.
    line-objid = 'ZDEMO_EXCEL35'.
    line-filename = '35_Static_Styles.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL36'.
    line-objid = 'ZDEMO_EXCEL36'.
    line-filename = '36_DefaultStyles.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL37'.
    line-objid = 'ZDEMO_EXCEL37'.
    line-filename = '37- Read template and output.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL38'.
    line-objid = 'ZDEMO_EXCEL38'.
    line-filename = '38_SAP-Icons.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL39'.
    line-objid = 'ZDEMO_EXCEL39'.
    line-filename = '39_Charts.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL40'.
    line-objid = 'ZDEMO_EXCEL40'.
    line-filename = '40_Printsettings.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL41'.
    line-objid = 'ZDEMO_EXCEL41'.
    line-filename = 'ABAP2XLSX Inheritance.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL49'.
    line-objid = 'ZDEMO_EXCEL49'.
    line-filename = '49_Bind_Table_Conversion_Exit.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL_COMMENTS'.
    line-objid = 'ZDEMO_EXCEL_COMMENTS'.
    line-filename = 'Comments.xlsx'.
    APPEND line TO result.
    line-program = 'ZTEST_EXCEL_IMAGE_HEADER'.
    line-objid = 'ZTEST_EXCEL_IMAGE_HEADER'.
    line-filename = 'Image_Header_Footer.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_01'.
    line-filename = '15_01_HelloWorldFromReader.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_02'.
    line-filename = '15_02_StylesFromReader.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_03'.
    line-filename = '15_03_iTabFromReader.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_04'.
    line-filename = '15_04_SheetsFromReader.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_05'.
    line-filename = '15_05_ConditionalFromReader.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_07'.
    line-filename = '15_07_ConditionalAllFromReader.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_08'.
    line-filename = '15_08_RangeFromReader.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_13'.
    line-filename = '15_13_MergedCellsFromReader.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_24'.
    line-filename = '15_24_Sheets_with_different_default_date_formatsFromReader.xlsx'.
    APPEND line TO result.
    line-program = 'ZDEMO_EXCEL15'.
    line-objid = 'ZDEMO_EXCEL15_31'.
    line-filename = '15_31_AutosizeWithDifferentFontSizesFromReader.xlsx'.
    APPEND line TO result.

  ENDMETHOD.


  METHOD gui_upload.

    DATA: solix_tab   TYPE solix_tab,
          file_length TYPE i.

    cl_gui_frontend_services=>gui_upload(
      EXPORTING
        filename                = file_name
        filetype                = 'BIN'
      IMPORTING
        filelength              = file_length
      CHANGING
        data_tab                = solix_tab
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
        OTHERS                  = 19 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |gui_upload error { file_name }|.
    ENDIF.

    result = cl_bcs_convert=>solix_to_xstring( it_solix = solix_tab iv_size = file_length ).

  ENDMETHOD.


  METHOD load_alv_table.

    DATA: demos        TYPE ty_demos,
          alv_line     TYPE ty_alv_line,
          cell_type    TYPE salv_s_int4_column,
          error        TYPE REF TO cx_root,
          check_result TYPE ty_check_result.
    FIELD-SYMBOLS:
      <demo> TYPE ty_demo.

    CLEAR alv_table.

    demos = get_list_of_demo_files( ).
    LOOP AT demos ASSIGNING <demo>.

      TRY.

          CLEAR alv_line.
          alv_line-objid = <demo>-objid.
          alv_line-obj_text = |{ <demo>-filename } ({ <demo>-program })|.
          alv_line-filename = <demo>-filename.
          alv_line-program = <demo>-program.
          SELECT SINGLE text FROM trdirt INTO alv_line-prog_text
              WHERE sprsl = sy-langu
                AND name  = alv_line-program.

          check_result = check_regression( <demo> ).

          CASE check_result-diff.
            WHEN abap_true.
              alv_line-status_icon = '@0W\QFiles are different@'.
            WHEN abap_false.
              alv_line-status_icon = '@0V\QFiles are identical@'.
          ENDCASE.
          IF check_result-diff = abap_true.
            alv_line-xlsx_diff = '@46\QShow differences@'.
            alv_line-write_smw0 = '@2L\QSave XLSX to Web Repository@'.
            cell_type-columnname = 'XLSX_DIFF'.
            cell_type-value      = if_salv_c_cell_type=>hotspot.
            APPEND cell_type TO alv_line-cell_types.
            cell_type-columnname = 'WRITE_SMW0'.
            cell_type-value      = if_salv_c_cell_type=>hotspot.
            APPEND cell_type TO alv_line-cell_types.
          ENDIF.
          alv_line-xlsx_just_now = check_result-xlsx_just_now.
          alv_line-xlsx_reference = check_result-xlsx_reference.
          alv_line-compare_xlsx_just_now = check_result-compare_xlsx_just_now.
          alv_line-compare_xlsx_reference = check_result-compare_xlsx_reference.
          APPEND alv_line TO alv_table.

        CATCH cx_root INTO error.
          alv_line-status_icon = |{ icon_cancel }{ error->get_text( ) }|.
      ENDTRY.

    ENDLOOP.

  ENDMETHOD.


  METHOD on_link_clicked.

    DATA: alv_line       TYPE ty_alv_line,
          error          TYPE REF TO cx_root,
          zip_old        TYPE REF TO cl_abap_zip,
          zip_new        TYPE REF TO cl_abap_zip,
          refresh_stable TYPE lvc_s_stbl,
          question       TYPE ty_popup_confirm_question.

    TRY.

        READ TABLE alv_table INDEX row INTO alv_line.
        ASSERT sy-subrc = 0.

        CASE column.

          WHEN 'XLSX_DIFF'.

            TRY.
                IF viewer IS NOT BOUND.
                  CREATE OBJECT viewer TYPE lcl_xlsx_diff_viewer
                    EXPORTING
                      io_container = zip_diff_container.
                ENDIF.

                CREATE OBJECT zip_old.
                zip_old->load(
                  EXPORTING
                    zip             = alv_line-compare_xlsx_reference
                  EXCEPTIONS
                    zip_parse_error = 1
                    OTHERS          = 2 ).
                IF sy-subrc <> 0.
                  RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |zip load old { alv_line-obj_text }|.
                ENDIF.

                CREATE OBJECT zip_new.
                zip_new->load(
                  EXPORTING
                    zip             = alv_line-compare_xlsx_just_now
                  EXCEPTIONS
                    zip_parse_error = 1
                    OTHERS          = 2 ).
                IF sy-subrc <> 0.
                  RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |zip load new { alv_line-filename }|.
                ENDIF.

                viewer->diff_and_view(
                    title_old = '"old" (Web Repository/expected)'
                    zip_old   = zip_old
                    title_new = '"new" (current execution)'
                    zip_new   = zip_new ).

              CATCH cx_root INTO error.
                RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |Viewer error: { error->get_text( ) }|.
            ENDTRY.

          WHEN 'WRITE_SMW0'.

            question = |Are you sure you want to overwrite { alv_line-objid } in Web repository?|.
            popup_confirm( question ).

            write_xlsx_to_web_repository(
                objid    = alv_line-objid
                text     = alv_line-obj_text
                xstring  = alv_line-xlsx_just_now
                filename = alv_line-filename ).
            COMMIT WORK.

            alv_line-status_icon = icon_okay.
            CLEAR alv_line-xlsx_diff.
            CLEAR alv_line-write_smw0.
            CLEAR alv_line-cell_types.
            MODIFY alv_table INDEX row FROM alv_line.

            refresh_stable-row = abap_true.
            refresh_stable-col = abap_true.
            salv->refresh( s_stable = refresh_stable ).

        ENDCASE.

      CATCH cx_root INTO error.
        MESSAGE error TYPE 'I' DISPLAY LIKE 'E'.
    ENDTRY.

  ENDMETHOD.


  METHOD popup_confirm.

    DATA: l_answer TYPE c LENGTH 1.

    CALL FUNCTION 'POPUP_TO_CONFIRM'
      EXPORTING
        text_question = question
      IMPORTING
        answer        = l_answer. "1 = button 1, 2 = button 2, A = cancel
    CASE l_answer.
      WHEN '2' OR 'A'.
        RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'Action cancelled by user'.
    ENDCASE.

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
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'WWW_GET_MIME_OBJECT'.
    ENDIF.

    result = cl_bcs_convert=>solix_to_xstring( it_solix = mime_table iv_size = content_length ).

  ENDMETHOD.


  METHOD read_screen_fields.

    DATA: lt_dummy   TYPE TABLE OF rsparams,
          lt_sel_255 TYPE TABLE OF rsparamsl_255.
    FIELD-SYMBOLS:
      <ls_sel_255> TYPE rsparamsl_255.

    CALL FUNCTION 'RS_REFRESH_FROM_SELECTOPTIONS'
      EXPORTING
        curr_report         = sy-repid
      TABLES
        selection_table     = lt_dummy
        selection_table_255 = lt_sel_255
      EXCEPTIONS
        not_found           = 1
        no_report           = 2
        OTHERS              = 3.

    READ TABLE lt_sel_255 WITH KEY selname = 'P_PATH' ASSIGNING <ls_sel_255>.
    ASSERT sy-subrc = 0.
    p_path = <ls_sel_255>-low.

  ENDMETHOD.


  METHOD screen_1001_pbo_first_time.

    DATA: columns TYPE REF TO cl_salv_columns_table,
          events  TYPE REF TO cl_salv_events_table.

    CREATE OBJECT splitter
      EXPORTING
        parent            = cl_gui_container=>screen0
        rows              = 1
        columns           = 2
      EXCEPTIONS
        cntl_error        = 1
        cntl_system_error = 2
        OTHERS            = 3.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'create splitter'.
    ENDIF.

    alv_container = splitter->get_container( row = 1 column = 1 ).
    zip_diff_container = splitter->get_container( row = 1 column = 2 ).

    cl_salv_table=>factory(
      EXPORTING
        r_container    = alv_container
      IMPORTING
        r_salv_table   = salv
      CHANGING
        t_table        = alv_table ).

    columns = salv->get_columns( ).
    columns->set_cell_type_column( value = 'CELL_TYPES' ).
    columns->get_column( 'STATUS_ICON' )->set_medium_text( 'Diff status' ).
    columns->get_column( 'STATUS_ICON' )->set_output_length( 2 ).
    columns->get_column( 'XLSX_DIFF' )->set_medium_text( 'View diff' ).
    columns->get_column( 'XLSX_DIFF' )->set_output_length( 5 ).
    columns->get_column( 'XLSX_DIFF' )->set_alignment( if_salv_c_alignment=>centered ).
    columns->get_column( 'WRITE_SMW0' )->set_medium_text( 'Web repository' ).
    columns->get_column( 'WRITE_SMW0' )->set_output_length( 5 ).
    columns->get_column( 'WRITE_SMW0' )->set_alignment( if_salv_c_alignment=>centered ).
    columns->get_column( 'PROGRAM' )->set_output_length( 15 ).
    columns->get_column( 'PROG_TEXT' )->set_output_length( 30 ).
    columns->get_column( 'OBJID' )->set_output_length( 20 ).
    columns->get_column( 'OBJ_TEXT' )->set_output_length( 50 ).
    columns->get_column( 'FILENAME' )->set_output_length( 50 ).
    columns->get_column( 'XLSX_JUST_NOW' )->set_technical( ).
    columns->get_column( 'XLSX_REFERENCE' )->set_technical( ).
    columns->get_column( 'COMPARE_XLSX_JUST_NOW' )->set_technical( ).
    columns->get_column( 'COMPARE_XLSX_REFERENCE' )->set_technical( ).

    events = salv->get_event( ).
    SET HANDLER on_link_clicked FOR events.

    salv->display( ).

  ENDMETHOD.


  METHOD set_sscrfields.

    GET REFERENCE OF sscrfields INTO ref_sscrfields.

  ENDMETHOD.


  METHOD write_screen_fields.

    DATA: fieldname TYPE string.
    FIELD-SYMBOLS:
      <field> TYPE simple.

    fieldname = |({ sy-repid })P_PATH|.
    ASSIGN (fieldname) TO <field>.
    <field> = p_path.

  ENDMETHOD.


  METHOD write_xlsx_to_web_repository.

    DATA: key                TYPE wwwdatatab,
          mime               TYPE TABLE OF w3mime,
          devclass           TYPE devclass,
          wwwparam           TYPE wwwparams,
          table_of_wwwparams TYPE TABLE OF wwwparams.

    SELECT SINGLE devclass FROM tadir
        INTO devclass
        WHERE pgmid    = 'R3TR'
          AND object   = 'W3MI'
          AND obj_name = objid.

    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |Object must be first created manually|.
    ENDIF.

    mime = cl_bcs_convert=>xstring_to_solix( xstring ).
    key-relid = 'MI'.
    key-objid = objid.
    key-chname = sy-uname.
    key-text = text.
    key-tdate = sy-datum.
    key-ttime = sy-uzeit.

    CALL FUNCTION 'WWWDATA_EXPORT'
      EXPORTING
        key               = key
      TABLES
        mime              = mime
      EXCEPTIONS
        wrong_object_type = 1
        export_error      = 2
        OTHERS            = 3.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'WWWDATA_EXPORT'.
    ENDIF.

    wwwparam-relid = 'MI'.
    wwwparam-objid = objid.
    wwwparam-name = 'mimetype'.
    wwwparam-value = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'.
    APPEND wwwparam TO table_of_wwwparams.
    wwwparam-name = 'filename'.
    wwwparam-value = filename.
    APPEND wwwparam TO table_of_wwwparams.
    wwwparam-name = 'filesize'.
    wwwparam-value = |{ xstrlen( xstring ) }|.
    APPEND wwwparam TO table_of_wwwparams.
    wwwparam-name = 'version'.
    wwwparam-value = |00001|.
    APPEND wwwparam TO table_of_wwwparams.
    wwwparam-name = 'fileextension'.
    wwwparam-value = |.xlsx|.
    APPEND wwwparam TO table_of_wwwparams.

    CALL FUNCTION 'WWWPARAMS_UPDATE'
      TABLES
        params       = table_of_wwwparams
      EXCEPTIONS
        update_error = 1
        OTHERS       = 2.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |WWWPARAMS_UPDATE { objid } { filename }|.
    ENDIF.

  ENDMETHOD.


ENDCLASS.



TABLES sscrfields.
DATA: app TYPE REF TO lcl_app.

PARAMETERS p_path TYPE zexcel_export_dir.

SELECTION-SCREEN BEGIN OF SCREEN 1001.
SELECTION-SCREEN FUNCTION KEY 1.
PARAMETERS dummy.
SELECTION-SCREEN END OF SCREEN 1001.

INITIALIZATION.
  CREATE OBJECT app.
  app->set_sscrfields( CHANGING sscrfields = sscrfields ).

AT SELECTION-SCREEN OUTPUT.
  app->at_selection_screen_output( ).

AT SELECTION-SCREEN.
  app->at_selection_screen( ).

AT SELECTION-SCREEN ON EXIT-COMMAND.
  app->at_selection_screen_on_exit( ).
