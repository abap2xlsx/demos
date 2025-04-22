METHOD onactionbtn_download .

  DATA: lo_excel        TYPE REF TO zcl_excel,
        lo_excel_writer TYPE REF TO zcl_excel_writer_2007,
        lo_worksheet    TYPE REF TO zcl_excel_worksheet.
  DATA: lv_content                    TYPE xstring.

  TRY.
      CREATE OBJECT lo_excel.
      lo_worksheet = lo_excel->get_active_worksheet( ).

      lo_worksheet->set_cell( ip_column = 'B'
      ip_row    = '2'
      ip_value  = 'Welcome to Web Dynpro and abap2xlsx.' ).

      CREATE OBJECT lo_excel_writer.
      lv_content = lo_excel_writer->zif_excel_writer~write_file( lo_excel ).
    CATCH zcx_excel.
      "Unlikely, ignore to keep demo simple.
  ENDTRY.

  DATA: lv_filename                   TYPE string.
  lv_filename = 'wda01.xlsx'.

  CALL METHOD cl_wd_runtime_services=>attach_file_to_response
    EXPORTING
      i_filename      = lv_filename
      i_content       = lv_content
      i_mime_type     = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      i_in_new_window = abap_false
      i_inplace       = abap_false.

ENDMETHOD.

method WDDOAFTERACTION .
endmethod.

method WDDOBEFOREACTION .
*  data lo_api_controller type ref to if_wd_view_controller.
*  data lo_action         type ref to if_wd_action.

*  lo_api_controller = wd_this->wd_get_api( ).
*  lo_action = lo_api_controller->get_current_action( ).

*  if lo_action is bound.
*    case lo_action->name.
*      when '...'.

*    endcase.
*  endif.
endmethod.

method WDDOEXIT .
endmethod.

method WDDOINIT .
endmethod.

method WDDOMODIFYVIEW .
  DATA lo_button TYPE REF TO cl_wd_button.
  lo_button ?= view->get_element( id = 'BTN_DOWNLOAD' ).
  lo_button->set_text( 'Download' ).
endmethod.

method WDDOONCONTEXTMENU .
endmethod.

