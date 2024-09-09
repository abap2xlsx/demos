CLASS zcl_excel_demo6 DEFINITION PUBLIC.
  PUBLIC SECTION.
    INTERFACES zif_excel_demo_output.
ENDCLASS.

CLASS zcl_excel_demo6 IMPLEMENTATION.

  METHOD zif_excel_demo_output~run.

    DATA: lo_worksheet TYPE REF TO zcl_excel_worksheet,
          lv_row       TYPE i,
          lv_formula   TYPE string.


    CREATE OBJECT ro_excel.

  " Get active sheet
    lo_worksheet = ro_excel->get_active_worksheet( ).

*--------------------------------------------------------------------*
*  Get some testdata
*--------------------------------------------------------------------*
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 100  ).
    lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 1000  ).
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 150 ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = -10  ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 500  ).


*--------------------------------------------------------------------*
*  Demonstrate using formulas
*--------------------------------------------------------------------*
    lo_worksheet->set_cell( ip_row = 9 ip_column = 'C' ip_formula = 'SUM(C4:C8)' ).


*--------------------------------------------------------------------*
* Demonstrate standard EXCEL-behaviour when copying a formula to another cell
* by calculating the resulting formula to put into another cell
*--------------------------------------------------------------------*
    DO 10 TIMES.

      lv_formula = zcl_excel_common=>shift_formula( iv_reference_formula = 'SUM(C4:C8)'
                                                  iv_shift_cols        = 0                " Offset in Columns - here we copy in same column --> 0
                                                  iv_shift_rows        = sy-index ).      " Offset in Row     - here we copy downward --> sy-index
      lv_row = 9 + sy-index.                                                                " Absolute row = sy-index rows below reference cell
      lo_worksheet->set_cell( ip_row = lv_row ip_column = 'C' ip_formula = lv_formula ).

    ENDDO.
  ENDMETHOD.

ENDCLASS.