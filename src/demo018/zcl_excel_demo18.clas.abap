CLASS zcl_excel_demo18 DEFINITION PUBLIC.
  PUBLIC SECTION.
    INTERFACES zif_excel_demo_output.
ENDCLASS.

CLASS zcl_excel_demo18 IMPLEMENTATION.

  METHOD zif_excel_demo_output~run.

    DATA: lo_worksheet             TYPE REF TO zcl_excel_worksheet,
          lv_style_protection_guid TYPE zexcel_cell_style.

    CREATE OBJECT ro_excel.

    " Get active sheet
    ro_excel->zif_excel_book_protection~protected     = zif_excel_book_protection=>c_protected.
    ro_excel->zif_excel_book_protection~lockrevision  = zif_excel_book_protection=>c_locked.
    ro_excel->zif_excel_book_protection~lockstructure = zif_excel_book_protection=>c_locked.
    ro_excel->zif_excel_book_protection~lockwindows   = zif_excel_book_protection=>c_locked.

    lo_worksheet = ro_excel->get_active_worksheet( ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 'This cell is unlocked' ip_style = lv_style_protection_guid ).

  ENDMETHOD.

ENDCLASS.
