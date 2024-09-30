*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL14
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel14.

DATA lo_excel TYPE REF TO zcl_excel.

CONSTANTS: gc_save_file_name TYPE string VALUE '14_Alignment.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.
  lo_excel = zcl_excel_demo14=>zif_excel_demo_output~run( ).

*** Create output
  lcl_output=>output( lo_excel ).
