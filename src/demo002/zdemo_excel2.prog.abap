*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL2
*& Test Styles for ABAP2XLSX
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel2.

DATA lo_excel TYPE REF TO zcl_excel.

CONSTANTS: gc_save_file_name TYPE string VALUE '02_Styles.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.
  lo_excel = zcl_excel_demo2=>zif_excel_demo_output~run( ).

  lcl_output=>output( lo_excel ).
