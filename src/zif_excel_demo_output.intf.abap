INTERFACE zif_excel_demo_output PUBLIC.
  CLASS-METHODS run
    RETURNING
      VALUE(ro_excel) TYPE REF TO zcl_excel
    RAISING
      zcx_excel.
ENDINTERFACE.
