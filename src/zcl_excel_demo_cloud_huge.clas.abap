"! <p class="shorttext synchronized">abap2xlsx: compare ZCL_EXCEL_WRITER_2007 / _HUGE_FILE</p>
CLASS zcl_excel_demo_cloud_huge DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    INTERFACES if_oo_adt_classrun.
  PROTECTED SECTION.
  PRIVATE SECTION.
    TYPES: BEGIN OF ty_s_case,
             tag  TYPE string,
             file TYPE xstring,
           END OF ty_s_case,
           ty_t_cases TYPE STANDARD TABLE OF ty_s_case WITH EMPTY KEY.
    CLASS-METHODS out_base64
      IMPORTING
        !iv_tag  TYPE string
        !iv_file TYPE xstring
        !io_out  TYPE REF TO if_oo_adt_classrun_out.
ENDCLASS.



CLASS zcl_excel_demo_cloud_huge IMPLEMENTATION.


  METHOD out_base64.
    DATA lv_off TYPE i.
    DATA(lv_b64) = cl_web_http_utility=>encode_x_base64( iv_file ).
    io_out->write( |<<{ iv_tag } bytes={ xstrlen( iv_file ) }>>| ).
    WHILE lv_off < strlen( lv_b64 ).
      io_out->write( substring( val = lv_b64 off = lv_off len = nmin( val1 = 1000 val2 = strlen( lv_b64 ) - lv_off ) ) ).
      lv_off = lv_off + 1000.
    ENDWHILE.
    io_out->write( |<<END { iv_tag }>>| ).
  ENDMETHOD.


  METHOD if_oo_adt_classrun~main.
    DATA lt_cases TYPE ty_t_cases.

    TRY.
        " passwords for the protection test are generated at runtime (no hard coded passwords)
        DATA(lo_random) = cl_abap_random_int=>create( seed = cl_abap_random=>seed( ) min = 100000 max = 999999 ).
        DATA(lv_sheet_pwd) = CONV zexcel_aes_password( |S{ lo_random->get_next( ) }| ).
        DATA(lv_book_pwd)  = CONV zexcel_aes_password( |B{ lo_random->get_next( ) }| ).

        lt_cases = VALUE #(
          ( tag = `JOURNAL`   file = zcl_excel_demo_cloud_test=>build_journal_xlsx( zcl_excel_demo_cloud_test=>get_demo_data( ) ) )
          ( tag = `CHART`     file = zcl_excel_demo_cloud_test=>build_chart_validation_xlsx( ) )
          ( tag = `CONDFMT`   file = zcl_excel_demo_cloud_test=>build_condformat_link_xlsx( ) )
          ( tag = `PROTECT`   file = zcl_excel_demo_cloud_test=>build_protect_pagesetup_xlsx( iv_sheet_password = lv_sheet_pwd
                                                                                             iv_book_password  = lv_book_pwd ) )
          ( tag = `OUTLINE`   file = zcl_excel_demo_cloud_test=>build_comment_outline_xlsx( ) )
          ( tag = `AFTAB`     file = zcl_excel_demo_cloud_test=>build_autofilter_table_xlsx( ) )
          ( tag = `RTMERGE`   file = zcl_excel_demo_cloud_test=>build_richtext_merge_xlsx( ) )
          ( tag = `IMAGES`    file = zcl_excel_demo_cloud_test=>build_image_drawing_xlsx( ) )
          ( tag = `FREEZE`    file = zcl_excel_demo_cloud_test=>build_freeze_width_xlsx( ) )
          ( tag = `NAMES`     file = zcl_excel_demo_cloud_test=>build_named_range_xlsx( ) )
          ( tag = `ADVSTYLE`  file = zcl_excel_demo_cloud_test=>build_advanced_style_xlsx( ) )
          ( tag = `XLFILE`    file = zcl_excel_demo_cloud_test=>get_excel_created_file( ) ) ) ##NO_TEXT.

        LOOP AT lt_cases INTO DATA(ls_case).
          TRY.
              DATA(lo_excel) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( ls_case-file ).
              DATA(lv_std) = CAST zif_excel_writer( NEW zcl_excel_writer_2007( ) )->write_file( lo_excel ).
              lo_excel = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( ls_case-file ).
              DATA(lv_huge) = CAST zif_excel_writer( NEW zcl_excel_writer_huge_file( ) )->write_file( lo_excel ).
              out->write( |HUGE { ls_case-tag }: 2007 { xstrlen( lv_std ) } bytes, huge { xstrlen( lv_huge ) } bytes| ) ##NO_TEXT.
              out_base64( iv_tag = |{ ls_case-tag }_STD|  iv_file = lv_std  io_out = out ).
              out_base64( iv_tag = |{ ls_case-tag }_HUGE| iv_file = lv_huge io_out = out ).
            CATCH cx_root INTO DATA(lx_case).
              out->write( |HUGE { ls_case-tag }: EXCEPTION { cl_abap_classdescr=>get_class_name( lx_case ) }: { lx_case->get_text( ) }| ) ##NO_TEXT.
              DATA(lx_prev) = lx_case->previous.
              WHILE lx_prev IS BOUND.
                out->write( |   previous { cl_abap_classdescr=>get_class_name( lx_prev ) }: { lx_prev->get_text( ) }| ) ##NO_TEXT.
                lx_prev = lx_prev->previous.
              ENDWHILE.
          ENDTRY.
        ENDLOOP.
      CATCH cx_root INTO DATA(lx_root).
        out->write( |HUGE EXCEPTION { cl_abap_classdescr=>get_class_name( lx_root ) }: { lx_root->get_text( ) }| ) ##NO_TEXT.
    ENDTRY.
  ENDMETHOD.
ENDCLASS.

