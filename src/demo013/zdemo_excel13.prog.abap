*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL13
*&
*&---------------------------------------------------------------------*
*& Example by: Alvaro "Blag" Tejada Galindo.
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel13.

DATA lo_excel TYPE REF TO zcl_excel.

CONSTANTS: gc_save_file_name TYPE string VALUE '13_MergedCells.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

START-OF-SELECTION.
  lo_excel = zcl_excel_demo13=>zif_excel_demo_output~run( ).

*** Create output
  lcl_output=>output( lo_excel ).
