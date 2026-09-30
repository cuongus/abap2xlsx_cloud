"! <p class="shorttext synchronized">abap2xlsx performance test (large workbook)</p>
CLASS zcl_excel_demo_cloud_perf DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    INTERFACES if_oo_adt_classrun.

    "! Build a workbook with iv_rows data rows x 10 columns (text, numbers, date, amount, formula)
    CLASS-METHODS build_large_xlsx
      IMPORTING
        !iv_rows       TYPE i
        !io_out        TYPE REF TO if_oo_adt_classrun_out OPTIONAL
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Time per SET_CELL call for different value types
    CLASS-METHODS benchmark_set_cell
      IMPORTING
        !iv_count TYPE i
        !io_out   TYPE REF TO if_oo_adt_classrun_out
      RAISING
        zcx_excel .
  PROTECTED SECTION.
  PRIVATE SECTION.
    CLASS-METHODS seconds_since
      IMPORTING
        !iv_start         TYPE utclong
      RETURNING
        VALUE(rv_seconds) TYPE decfloat34 .
ENDCLASS.



CLASS zcl_excel_demo_cloud_perf IMPLEMENTATION.


  METHOD benchmark_set_cell.
    " time per SET_CELL call for the different value types (iv_count cells each, column A of a new sheet)
    DATA: lo_excel TYPE REF TO zcl_excel,
          lo_sheet TYPE REF TO zcl_excel_worksheet.
    CREATE OBJECT lo_excel.
    DATA(lo_st_date) = lo_excel->add_new_style( ).
    lo_st_date->number_format->format_code = 'dd/mm/yyyy' ##NO_TEXT.
    DATA(lv_st_date) = lo_st_date->get_guid( ).
    DATA(lv_text) = `Khách hàng số 123` ##NO_TEXT.
    DATA(lv_int) = 12345.
    DATA(lv_dec) = CONV decfloat34( '1234.25' ).
    DATA(lv_pack) = VALUE zexcel_s_fieldcatalog-width( ).
    DATA(lv_date) = CONV d( '20260929' ).
    DO 8 TIMES.
      DATA(lv_case) = sy-index.
      lo_sheet = lo_excel->add_new_worksheet( ).
      DATA(lv_start) = utclong_current( ).
      DO iv_count TIMES.
        CASE lv_case.
          WHEN 1. lo_sheet->set_cell( ip_column = 1 ip_row = sy-index ip_value = lv_text ).
          WHEN 2. lo_sheet->set_cell( ip_column = 1 ip_row = sy-index ip_value = lv_int ).
          WHEN 3. lo_sheet->set_cell( ip_column = 1 ip_row = sy-index ip_value = lv_dec ).
          WHEN 4. lo_sheet->set_cell( ip_column = 1 ip_row = sy-index ip_value = lv_date ).
          WHEN 5. lo_sheet->set_cell( ip_column = 1 ip_row = sy-index ip_value = lv_date ip_style = lv_st_date ).
          WHEN 6. lo_sheet->set_cell( ip_column = 1 ip_row = sy-index ip_formula = 'B1*C1' ) ##NO_TEXT.
          WHEN 7. lo_sheet->set_cell( ip_column = 'A' ip_row = sy-index ip_value = lv_text ) ##NO_TEXT.
          WHEN 8. lo_sheet->set_cell( ip_column = 1 ip_row = sy-index ip_value = lv_pack ).
        ENDCASE.
      ENDDO.
      io_out->write( |BENCH set_cell { SWITCH string( lv_case WHEN 1 THEN `string` WHEN 2 THEN `int` WHEN 3 THEN `decfloat34`
                                                             WHEN 4 THEN `date` WHEN 5 THEN `date+style` WHEN 6 THEN `formula`
                                                             WHEN 7 THEN `string col 'A'` WHEN 8 THEN `packed` ) }:| &&
                     | { round( val = utclong_diff( high = utclong_current( ) low = lv_start ) * 1000000 / iv_count dec = 1 ) } µs/cell| ) ##NO_TEXT.
    ENDDO.
  ENDMETHOD.


  METHOD seconds_since.
    rv_seconds = round( val = utclong_diff( high = utclong_current( ) low = iv_start ) dec = 2 ).
  ENDMETHOD.


  METHOD build_large_xlsx.
    DATA: lo_excel  TYPE REF TO zcl_excel,
          lo_sheet  TYPE REF TO zcl_excel_worksheet,
          lo_writer TYPE REF TO zif_excel_writer,
          lv_row    TYPE i.

    DATA(lv_start) = utclong_current( ).
    CREATE OBJECT lo_excel.
    DATA(lo_st_head) = lo_excel->add_new_style( ).
    lo_st_head->font->bold = abap_true.
    lo_st_head->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_st_head->fill->fgcolor-rgb = 'FFD9E1F2' ##NO_TEXT.
    DATA(lo_st_date) = lo_excel->add_new_style( ).
    lo_st_date->number_format->format_code = 'dd/mm/yyyy' ##NO_TEXT.
    DATA(lo_st_amt) = lo_excel->add_new_style( ).
    lo_st_amt->number_format->format_code = '#,##0.00' ##NO_TEXT.
    DATA(lv_st_head) = lo_st_head->get_guid( ).
    DATA(lv_st_date) = lo_st_date->get_guid( ).
    DATA(lv_st_amt) = lo_st_amt->get_guid( ).

    lo_sheet = lo_excel->get_active_worksheet( ).
    lo_sheet->set_title( 'Data' ) ##NO_TEXT.
    DATA(lt_head) = VALUE string_table( ( `STT` ) ( `Số chứng từ` ) ( `Ngày` ) ( `Mã KH` ) ( `Tên khách hàng` )
                                        ( `Số lượng` ) ( `Đơn giá` ) ( `Thành tiền` ) ( `Ghi chú` ) ( `Kiểm tra` ) ) ##NO_TEXT.
    LOOP AT lt_head INTO DATA(lv_head).
      lo_sheet->set_cell( ip_column = sy-tabix ip_row = 1 ip_value = lv_head ip_style = lv_st_head ).
    ENDLOOP.

    DATA(lv_date) = CONV d( '20260101' ).
    DO iv_rows TIMES.
      lv_row = sy-index + 1.
      lo_sheet->set_cell( ip_column = 1  ip_row = lv_row ip_value = sy-index ).
      lo_sheet->set_cell( ip_column = 2  ip_row = lv_row ip_value = |CT{ sy-index WIDTH = 8 ALIGN = RIGHT PAD = '0' }| ) ##NO_TEXT.
      lo_sheet->set_cell( ip_column = 3  ip_row = lv_row ip_value = CONV d( lv_date + sy-index MOD 365 ) ip_style = lv_st_date ).
      lo_sheet->set_cell( ip_column = 4  ip_row = lv_row ip_value = |KH{ sy-index MOD 500 WIDTH = 4 ALIGN = RIGHT PAD = '0' }| ) ##NO_TEXT.
      lo_sheet->set_cell( ip_column = 5  ip_row = lv_row ip_value = |Khách hàng số { sy-index MOD 500 }| ) ##NO_TEXT.
      lo_sheet->set_cell( ip_column = 6  ip_row = lv_row ip_value = sy-index MOD 97 + 1 ).
      lo_sheet->set_cell( ip_column = 7  ip_row = lv_row ip_value = CONV decfloat34( 1000 + sy-index MOD 1000 / 4 ) ip_style = lv_st_amt ).
      lo_sheet->set_cell( ip_column = 8  ip_row = lv_row ip_formula = |F{ lv_row }*G{ lv_row }| ip_style = lv_st_amt ) ##NO_TEXT.
      lo_sheet->set_cell( ip_column = 9  ip_row = lv_row ip_value = COND string( WHEN sy-index MOD 10 = 0 THEN `Đơn hàng ưu tiên` ) ) ##NO_TEXT.
      lo_sheet->set_cell( ip_column = 10 ip_row = lv_row ip_value = xsdbool( sy-index MOD 3 = 0 ) ).
    ENDDO.
    lo_sheet->freeze_panes( ip_num_rows = 1 ).
    lo_excel->add_new_autofilter( lo_sheet )->set_filter_area( VALUE #( row_start = 1 col_start = 1 row_end = lv_row col_end = 10 ) ).
    lo_sheet->get_column( 'B' )->set_width( 14 ).
    lo_sheet->get_column( 'C' )->set_width( 12 ).
    lo_sheet->get_column( 'E' )->set_width( 24 ).
    lo_sheet->get_column( 'H' )->set_width( 16 ).
    IF io_out IS BOUND.
      io_out->write( |PERF { iv_rows } rows: build { seconds_since( lv_start ) } s| ) ##NO_TEXT.
    ENDIF.

    lv_start = utclong_current( ).
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
    IF io_out IS BOUND.
      io_out->write( |PERF { iv_rows } rows: write { seconds_since( lv_start ) } s, { xstrlen( rv_file ) } bytes| ) ##NO_TEXT.
    ENDIF.
  ENDMETHOD.


  METHOD if_oo_adt_classrun~main.
    DATA lv_b64 TYPE string.
    DATA lv_off TYPE i.

    TRY.
        benchmark_set_cell( iv_count = 20000 io_out = out ).
        DATA lt_sizes TYPE STANDARD TABLE OF i WITH EMPTY KEY.
        lt_sizes = VALUE #( ( 10000 ) ( 50000 ) ).
        LOOP AT lt_sizes INTO DATA(lv_rows).
          DATA(lv_total) = utclong_current( ).
          DATA(lv_file) = build_large_xlsx( iv_rows = lv_rows io_out = out ).
          " same workbook with the streaming writer / reader for large files
          DATA(lo_excel_again) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_file ).
          DATA(lv_start) = utclong_current( ).
          DATA(lv_file_huge) = CAST zif_excel_writer( NEW zcl_excel_writer_huge_file( ) )->write_file( lo_excel_again ).
          out->write( |PERF { lv_rows } rows: write ZCL_EXCEL_WRITER_HUGE_FILE { seconds_since( lv_start ) } s, { xstrlen( lv_file_huge ) } bytes| ) ##NO_TEXT.
          FREE lo_excel_again.
          lv_start = utclong_current( ).
          DATA(lo_read) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_file ).
          out->write( |PERF { lv_rows } rows: read ZCL_EXCEL_READER_2007 { seconds_since( lv_start ) } s, highest row { lo_read->get_active_worksheet( )->get_highest_row( ) }| ) ##NO_TEXT.
          FREE lo_read.
          lv_start = utclong_current( ).
          lo_read = NEW zcl_excel_reader_huge_file( )->zif_excel_reader~load( lv_file ).
          out->write( |PERF { lv_rows } rows: read ZCL_EXCEL_READER_HUGE_FILE { seconds_since( lv_start ) } s, highest row { lo_read->get_active_worksheet( )->get_highest_row( ) }| ) ##NO_TEXT.
          FREE lo_read.
          out->write( |PERF { lv_rows } rows: total { seconds_since( lv_total ) } s| ) ##NO_TEXT.
        ENDLOOP.
        " the 50.000 rows file for Excel
        lv_b64 = cl_web_http_utility=>encode_x_base64( lv_file ).
        out->write( |<<PERF50K bytes={ xstrlen( lv_file ) }>>| ).
        WHILE lv_off < strlen( lv_b64 ).
          out->write( substring( val = lv_b64 off = lv_off len = nmin( val1 = 1000 val2 = strlen( lv_b64 ) - lv_off ) ) ).
          lv_off = lv_off + 1000.
        ENDWHILE.
        out->write( |<<END PERF50K>>| ).
      CATCH cx_root INTO DATA(lx_root).
        out->write( |PERF EXCEPTION { cl_abap_classdescr=>get_class_name( lx_root ) }: { lx_root->get_text( ) }| ) ##NO_TEXT.
    ENDTRY.
  ENDMETHOD.
ENDCLASS.

