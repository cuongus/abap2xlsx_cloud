CLASS zcl_excel_demo_cloud_test DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    INTERFACES:
      if_oo_adt_classrun ,
      if_rap_query_provider .

    " rows of custom entity ZC_DEMO_ABAP2XLSX (query implemented by this class)
    TYPES ty_t_entity TYPE STANDARD TABLE OF zc_demo_abap2xlsx WITH EMPTY KEY.
    " demo row of the journal (layout Mau1.xlsx)
    TYPES: ty_s_journal TYPE zc_demo_abap2xlsx,
           ty_t_journal TYPE STANDARD TABLE OF ty_s_journal WITH EMPTY KEY.

    "! Demo data for ZC_DEMO_ABAP2XLSX
    CLASS-METHODS get_demo_data
      RETURNING
        VALUE(rt_data) TYPE ty_t_journal .
    "! Macro enabled test workbook (.xlsm) with the embedded VBA project
    CLASS-METHODS build_xlsm
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Embedded vbaProject.bin (XlsxWriter example, macro say_hello)
    CLASS-METHODS get_test_vba_project
      RETURNING
        VALUE(rv_file) TYPE xstring .
    "! Workbook created by Microsoft Excel (embedded) to test the reader with a file not written by abap2xlsx
    CLASS-METHODS get_excel_created_file
      RETURNING
        VALUE(rv_file) TYPE xstring .
    "! Test workbook for global / sheet-local named ranges and cross-sheet formulas
    CLASS-METHODS build_named_range_xlsx
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Test workbook for gradient fills, text rotation, indent and other alignments
    CLASS-METHODS build_advanced_style_xlsx
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Test workbook for freeze panes, column widths and row heights
    CLASS-METHODS build_freeze_width_xlsx
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Test workbook for images (anchors, reuse, header / footer images)
    CLASS-METHODS build_image_drawing_xlsx
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Embedded test images: PNG (240x80) or JPG (320x200)
    CLASS-METHODS get_test_image
      IMPORTING
        !iv_kind        TYPE string
      RETURNING
        VALUE(rv_image) TYPE xstring .
    "! Test workbook for rich text cells and merged cells
    CLASS-METHODS build_richtext_merge_xlsx
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Test workbook for autofilter (value + text filters) and tables (totals row, column formula)
    CLASS-METHODS build_autofilter_table_xlsx
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Test workbook for comments and row / column outlines
    CLASS-METHODS build_comment_outline_xlsx
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Test workbook for sheet / workbook protection and page setup (passwords are supplied by the caller)
    CLASS-METHODS build_protect_pagesetup_xlsx
      IMPORTING
        !iv_sheet_password TYPE zexcel_aes_password
        !iv_book_password  TYPE zexcel_aes_password
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Test workbook for conditional formatting and hyperlinks
    CLASS-METHODS build_condformat_link_xlsx
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Test workbook for charts (bar / line / pie) and data validations
    CLASS-METHODS build_chart_validation_xlsx
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .
    "! Transaction journal (layout Mau1.xlsx), one worksheet per company code + customer
    CLASS-METHODS build_journal_xlsx
      IMPORTING
        !it_data       TYPE ty_t_journal
      RETURNING
        VALUE(rv_file) TYPE xstring
      RAISING
        zcx_excel .

  PROTECTED SECTION.
  PRIVATE SECTION.
    TYPES:
      BEGIN OF ty_s_row,
        title    TYPE zexcel_sheet_title,
        amount   TYPE p LENGTH 13 DECIMALS 2,
        qty      TYPE i,
        doc_date TYPE d,
        note     TYPE string,
      END OF ty_s_row,
      ty_t_row TYPE STANDARD TABLE OF ty_s_row WITH EMPTY KEY.

    METHODS build_workbook
      RETURNING
        VALUE(ro_excel) TYPE REF TO zcl_excel
      RAISING
        zcx_excel .
    CLASS-METHODS new_style
      IMPORTING
        !io_excel      TYPE REF TO zcl_excel
        !iv_font       TYPE zexcel_style_font_name
        !iv_size       TYPE zexcel_style_font_size DEFAULT 12
        !iv_bold       TYPE abap_bool DEFAULT abap_false
        !iv_halign     TYPE zexcel_alignment OPTIONAL
        !iv_valign     TYPE zexcel_alignment DEFAULT zcl_excel_style_alignment=>c_vertical_center
        !io_border     TYPE REF TO zcl_excel_style_border OPTIONAL
        !iv_format     TYPE zexcel_number_format OPTIONAL
      RETURNING
        VALUE(rv_guid) TYPE zexcel_cell_style .
    CLASS-METHODS set_line
      IMPORTING
        !io_sheet TYPE REF TO zcl_excel_worksheet
        !iv_row   TYPE zexcel_cell_row
        !iv_style TYPE zexcel_cell_style
        !iv_text  TYPE string
      RAISING
        zcx_excel .
    CLASS-METHODS set_amount
      IMPORTING
        !io_sheet  TYPE REF TO zcl_excel_worksheet
        !iv_column TYPE zexcel_cell_column_alpha
        !iv_row    TYPE zexcel_cell_row
        !iv_amount TYPE zc_demo_abap2xlsx-amount_dr
        !iv_style  TYPE zexcel_cell_style
      RAISING
        zcx_excel .
    CLASS-METHODS matches_ranges
      IMPORTING
        !iv_value       TYPE any
        !it_range       TYPE if_rap_query_filter=>tt_range_option
      RETURNING
        VALUE(rv_match) TYPE abap_bool .
    CLASS-METHODS random_password
      IMPORTING
        !io_random         TYPE REF TO cl_abap_random_int
      RETURNING
        VALUE(rv_password) TYPE zexcel_aes_password .
    METHODS out_base64
      IMPORTING
        !iv_tag  TYPE string
        !iv_file TYPE xstring
        !io_out  TYPE REF TO if_oo_adt_classrun_out .
ENDCLASS.



CLASS zcl_excel_demo_cloud_test IMPLEMENTATION.


  METHOD build_workbook.
    DATA: lo_sheet   TYPE REF TO zcl_excel_worksheet,
          lo_sheet2  TYPE REF TO zcl_excel_worksheet,
          lo_style   TYPE REF TO zcl_excel_style,
          lo_date    TYPE REF TO zcl_excel_style,
          lo_drawing TYPE REF TO zcl_excel_drawing,
          lt_rows    TYPE ty_t_row,
          lv_png     TYPE xstring.

    CREATE OBJECT ro_excel.

    " Sheet 1: plain cells, styles, formula, merge, unicode text, image
    lo_sheet = ro_excel->get_active_worksheet( ).
    lo_sheet->set_title( 'Report' ) ##NO_TEXT.

    lo_style = ro_excel->add_new_style( ).
    lo_style->font->bold = abap_true.
    lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb = zcl_excel_style_color=>c_yellow.

    lo_date = ro_excel->add_new_style( ).
    lo_date->number_format->format_code = zcl_excel_style_number_format=>c_format_date_std.

    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'abap2xlsx ABAP Cloud test' ip_style = lo_style->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_merge( ip_range = 'A1:D1' ).
    lo_sheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'Tiếng Việt có dấu: Hóa đơn điện tử' ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'Ký tự đặc biệt: <tag> & "quote" €' ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = 1234 ).
    lo_sheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = '56.78' ).
    lo_sheet->set_cell( ip_column = 'C' ip_row = 4 ip_formula = 'A4+B4' ).
    lo_sheet->set_cell( ip_column = 'D' ip_row = 4 ip_formula = 'A6&" - "&A2' ).
    lo_sheet->set_cell( ip_column = 'A' ip_row = 5 ip_value = CONV d( '20260927' ) ip_style = lo_date->get_guid( ) ).
    lo_sheet->set_cell( ip_column = 'A' ip_row = 6 ip_value = 'Mã: 000123' ) ##NO_TEXT.

    " 1x1 PNG
    lv_png = '89504E470D0A1A0A0000000D4948445200000001000000010806000000'
          && '1F15C4890000000D4944415478DA63F8FFFF3F0005FE02FEA7D2A0D800'
          && '00000049454E44AE426082'.
    lo_drawing = ro_excel->add_new_drawing( ).
    lo_drawing->set_position( ip_from_row = 8 ip_from_col = 'B' ).
    lo_drawing->set_media( ip_media      = lv_png
                           ip_media_type = 'png'
                           ip_width      = 32
                           ip_height     = 32 ).
    lo_sheet->add_drawing( lo_drawing ).

    " Sheet 2: table binding (field catalog via RTTI + XCO labels)
    lt_rows = VALUE #(
      ( title = 'Dòng 1' amount = '1000.50' qty = 3  doc_date = '20260101' note = 'Alpha' )
      ( title = 'Dòng 2' amount = '-20.25'  qty = 10 doc_date = '20261231' note = 'Bravo' ) ) ##NO_TEXT.
    lo_sheet2 = ro_excel->add_new_worksheet( ).
    lo_sheet2->set_title( 'Table' ) ##NO_TEXT.
    lo_sheet2->bind_table( ip_table = lt_rows ).
  ENDMETHOD.


  METHOD random_password.
    CONSTANTS lc_chars TYPE string VALUE `ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz0123456789ABCDE` ##NO_TEXT.
    DATA lv_password TYPE string.

    DO 8 TIMES.
      lv_password = lv_password && substring( val = lc_chars off = io_random->get_next( ) len = 1 ).
    ENDDO.
    rv_password = lv_password.
  ENDMETHOD.


  METHOD out_base64.
    DATA lv_b64 TYPE string.
    DATA lv_off TYPE i.

    lv_b64 = cl_web_http_utility=>encode_x_base64( iv_file ).
    io_out->write( |<<{ iv_tag } bytes={ xstrlen( iv_file ) }>>| ).
    WHILE lv_off < strlen( lv_b64 ).
      io_out->write( substring( val = lv_b64 off = lv_off len = nmin( val1 = 1000 val2 = strlen( lv_b64 ) - lv_off ) ) ).
      lv_off = lv_off + 1000.
    ENDWHILE.
    io_out->write( |<<END { iv_tag }>>| ).
  ENDMETHOD.


  METHOD if_oo_adt_classrun~main.
    DATA: lo_excel   TYPE REF TO zcl_excel,
          lo_read    TYPE REF TO zcl_excel,
          lo_writer  TYPE REF TO zif_excel_writer,
          lo_reader  TYPE REF TO zif_excel_reader,
          lo_sheet   TYPE REF TO zcl_excel_worksheet,
          lv_xlsx    TYPE xstring,
          lv_csv     TYPE xstring,
          lv_value   TYPE zexcel_cell_value,
          lv_formula TYPE zexcel_cell_formula,
          lv_row     TYPE i,
          lx_excel   TYPE REF TO zcx_excel,
          lx_root    TYPE REF TO cx_root.

    TRY.
        " 1) write XLSX
        lo_excel = build_workbook( ).
        CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
        lv_xlsx = lo_writer->write_file( lo_excel ).
        out->write( |WRITE XLSX OK, { xstrlen( lv_xlsx ) } bytes| ).
        out_base64( iv_tag = `XLSX` iv_file = lv_xlsx io_out = out ).

        " 2) write CSV (first sheet)
        CREATE OBJECT lo_writer TYPE zcl_excel_writer_csv.
        lv_csv = lo_writer->write_file( lo_excel ).
        out->write( |WRITE CSV OK, { xstrlen( lv_csv ) } bytes| ).
        out_base64( iv_tag = `CSV` iv_file = lv_csv io_out = out ).

        " 3) read the XLSX back
        CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
        lo_read = lo_reader->load( lv_xlsx ).
        out->write( |READ XLSX OK, worksheets: { lo_read->get_worksheets_size( ) }| ) ##NO_TEXT.
        lo_sheet = lo_read->get_worksheet_by_name( 'Report' ) ##NO_TEXT.
        DO 6 TIMES.
          lv_row = sy-index.
          lo_sheet->get_cell( EXPORTING ip_column = 'A' ip_row = lv_row IMPORTING ep_value = lv_value ).
          out->write( |READ Report!A{ lv_row } = { lv_value }| ) ##NO_TEXT.
        ENDDO.
        lo_sheet->get_cell( EXPORTING ip_column = 'C' ip_row = 4 IMPORTING ep_formula = lv_formula ).
        out->write( |READ Report!C4 formula = { lv_formula }| ) ##NO_TEXT.
        lo_sheet->get_cell( EXPORTING ip_column = 'D' ip_row = 4 IMPORTING ep_formula = lv_formula ).
        out->write( |READ Report!D4 formula = { lv_formula }| ) ##NO_TEXT.
        lo_sheet = lo_read->get_worksheet_by_name( 'Table' ) ##NO_TEXT.
        DO 3 TIMES.
          lv_row = sy-index.
          lo_sheet->get_cell( EXPORTING ip_column = 'A' ip_row = lv_row IMPORTING ep_value = lv_value ).
          out->write( |READ Table!A{ lv_row } = { lv_value }| ) ##NO_TEXT.
          lo_sheet->get_cell( EXPORTING ip_column = 'B' ip_row = lv_row IMPORTING ep_value = lv_value ).
          out->write( |READ Table!B{ lv_row } = { lv_value }| ) ##NO_TEXT.
        ENDDO.

        " console check: literal ampersand (built from code point 38), and its length
        DATA(lv_amp) = zcl_excel_common=>unicode_char_from_code_point( 38 ).
        out->write( |LITERAL = [a { lv_amp } b] len={ strlen( |a { lv_amp } b| ) }| ) ##NO_TEXT.
        lo_sheet = lo_read->get_worksheet_by_name( 'Report' ) ##NO_TEXT.
        lo_sheet->get_cell( EXPORTING ip_column = 'A' ip_row = 3 IMPORTING ep_value = lv_value ).
        out->write( |A3 amp-entity-text={ xsdbool( lv_value CS |{ lv_amp }amp;| ) } count-amp={ count( val = lv_value sub = lv_amp ) } len={ strlen( lv_value ) }| ) ##NO_TEXT.

        " 5) journal workbook of custom entity ZC_DEMO_ABAP2XLSX (layout Mau1.xlsx)
        DATA(lv_journal) = build_journal_xlsx( get_demo_data( ) ).
        out->write( |JOURNAL XLSX OK, { xstrlen( lv_journal ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `JOURNAL` iv_file = lv_journal io_out = out ).

        " 6) charts + data validations: write, then read back with the reader
        DATA(lv_charts) = build_chart_validation_xlsx( ).
        out->write( |CHARTS XLSX OK, { xstrlen( lv_charts ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `CHARTS` iv_file = lv_charts io_out = out ).
        lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_charts ).
        lo_sheet = lo_read->get_worksheet_by_name( 'Data' ) ##NO_TEXT.
        DATA(lo_chart_iter) = lo_sheet->get_drawings_iterator( zcl_excel_drawing=>type_chart ).
        WHILE lo_chart_iter->has_next( ) = abap_true.
          DATA(lo_chart) = CAST zcl_excel_drawing( lo_chart_iter->get_next( ) ).
          IF lo_chart->graph IS BOUND.
            out->write( |READ chart type={ lo_chart->graph_type } title={ lo_chart->graph->title } series={ lines( lo_chart->graph->series ) }| ) ##NO_TEXT.
          ELSE.
            out->write( |READ chart type={ lo_chart->graph_type } (no graph object)| ) ##NO_TEXT.
          ENDIF.
        ENDWHILE.
        lo_sheet = lo_read->get_worksheet_by_name( 'Validation' ) ##NO_TEXT.
        out->write( |READ data validations = { lo_sheet->get_data_validations_size( ) }| ) ##NO_TEXT.
        DATA(lo_dv_iter) = lo_sheet->get_data_validations_iterator( ).
        WHILE lo_dv_iter->has_next( ) = abap_true.
          DATA(lo_dv) = CAST zcl_excel_data_validation( lo_dv_iter->get_next( ) ).
          out->write( |READ dv { lo_dv->cell_column }{ lo_dv->cell_row }:{ lo_dv->cell_column_to }{ lo_dv->cell_row_to }| &&
                      | type={ lo_dv->type } op={ lo_dv->operator } f1={ lo_dv->formula1 } f2={ lo_dv->formula2 } style={ lo_dv->errorstyle }| ) ##NO_TEXT.
        ENDWHILE.

        " 7) conditional formatting + hyperlinks: write, then read back with the reader
        DATA(lv_cond) = build_condformat_link_xlsx( ).
        out->write( |CONDFMT XLSX OK, { xstrlen( lv_cond ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `CONDFMT` iv_file = lv_cond io_out = out ).
        lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_cond ).
        lo_sheet = lo_read->get_worksheet_by_name( 'CondFormat' ) ##NO_TEXT.
        DATA(lo_cond_iter) = lo_sheet->get_style_cond_iterator( ).
        WHILE lo_cond_iter->has_next( ) = abap_true.
          DATA(lo_cond) = CAST zcl_excel_style_cond( lo_cond_iter->get_next( ) ).
          out->write( |READ cond { lo_cond->get_dimension_range( ) } rule={ lo_cond->rule } prio={ lo_cond->priority }| ) ##NO_TEXT.
        ENDWHILE.
        " round trip: the workbook read back is written again (conditional formats must survive)
        CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
        DATA(lv_cond_rt) = lo_writer->write_file( lo_read ).
        out->write( |CONDFMT ROUNDTRIP OK, { xstrlen( lv_cond_rt ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `CONDRT` iv_file = lv_cond_rt io_out = out ).
        lo_sheet = lo_read->get_worksheet_by_name( 'Links' ) ##NO_TEXT.
        out->write( |READ hyperlinks = { lo_sheet->get_hyperlinks_size( ) }| ) ##NO_TEXT.
        DATA(lo_link_iter) = lo_sheet->get_hyperlinks_iterator( ).
        WHILE lo_link_iter->has_next( ) = abap_true.
          DATA(lo_link) = CAST zcl_excel_hyperlink( lo_link_iter->get_next( ) ).
          out->write( |READ link ref={ lo_link->get_ref( ) } internal={ lo_link->is_internal( ) } url={ lo_link->get_url( ) }| ) ##NO_TEXT.
        ENDWHILE.

        " 8) protection + page setup: write, then read back with the reader
        " no hard-coded passwords: random test passwords, printed for the tester
        DATA(lo_random) = cl_abap_random_int=>create( seed = cl_abap_random=>seed( ) min = 0 max = 61 ).
        DATA(lv_sheet_pwd) = random_password( lo_random ).
        DATA(lv_book_pwd)  = random_password( lo_random ).
        out->write( |TEST PASSWORDS sheet={ lv_sheet_pwd } book={ lv_book_pwd }| ) ##NO_TEXT.
        DATA(lv_prot) = build_protect_pagesetup_xlsx( iv_sheet_password = lv_sheet_pwd iv_book_password = lv_book_pwd ).
        out->write( |PROTECT XLSX OK, { xstrlen( lv_prot ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `PROTECT` iv_file = lv_prot io_out = out ).
        lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_prot ).
        out->write( |READ book protected={ lo_read->zif_excel_book_protection~protected } lockstructure={ lo_read->zif_excel_book_protection~lockstructure }| &&
                    | pwd={ lo_read->zif_excel_book_protection~workbookpassword }| ) ##NO_TEXT.
        lo_sheet = lo_read->get_worksheet_by_name( 'Protected' ) ##NO_TEXT.
        out->write( |READ sheet protected={ lo_sheet->zif_excel_sheet_protection~protected } pwd={ lo_sheet->zif_excel_sheet_protection~password }| &&
                    | sheet={ lo_sheet->zif_excel_sheet_protection~sheet } objects={ lo_sheet->zif_excel_sheet_protection~objects }| ) ##NO_TEXT.
        lo_sheet = lo_read->get_worksheet_by_name( 'PageSetup' ) ##NO_TEXT.
        out->write( |READ setup orientation={ lo_sheet->sheet_setup->orientation } paper={ lo_sheet->sheet_setup->paper_size }| &&
                    | fit={ lo_sheet->sheet_setup->fit_to_page } { lo_sheet->sheet_setup->fit_to_width }x{ lo_sheet->sheet_setup->fit_to_height }| &&
                    | margin_top={ lo_sheet->sheet_setup->margin_top } gridlines={ lo_sheet->print_gridlines }| ) ##NO_TEXT.
        lo_sheet->sheet_setup->get_header_footer_string( IMPORTING ep_odd_header = DATA(lv_odd_header) ep_odd_footer = DATA(lv_odd_footer) ).
        out->write( |READ header={ lv_odd_header } footer={ lv_odd_footer }| ) ##NO_TEXT.
        lo_sheet->zif_excel_sheet_printsettings~get_print_repeat_rows( IMPORTING ev_rows_from = DATA(lv_rep_from) ev_rows_to = DATA(lv_rep_to) ).
        out->write( |READ repeat rows={ lv_rep_from }-{ lv_rep_to } pagebreaks={ lines( lo_sheet->get_pagebreaks( )->get_all_pagebreaks( ) ) }| ) ##NO_TEXT.
        " round trip: protection and header / footer must survive read + write
        CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
        DATA(lv_prot_rt) = lo_writer->write_file( lo_read ).
        out->write( |PROTECT ROUNDTRIP OK, { xstrlen( lv_prot_rt ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `PROTRT` iv_file = lv_prot_rt io_out = out ).

        " 9) comments + outlines: write, read back, write again (round trip)
        DATA(lv_outline) = build_comment_outline_xlsx( ).
        out->write( |OUTLINE XLSX OK, { xstrlen( lv_outline ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `OUTLINE` iv_file = lv_outline io_out = out ).
        lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_outline ).
        lo_sheet = lo_read->get_worksheet_by_name( 'Outline' ) ##NO_TEXT.
        DATA(lo_comment_iter) = lo_sheet->get_comments( )->get_iterator( ).
        WHILE lo_comment_iter->has_next( ) = abap_true.
          DATA(lo_comment) = CAST zcl_excel_comment( lo_comment_iter->get_next( ) ).
          out->write( |READ comment { lo_comment->get_ref( ) } len={ strlen( lo_comment->get_text( ) ) }| ) ##NO_TEXT.
        ENDWHILE.
        LOOP AT lo_sheet->get_row_outlines( ) INTO DATA(ls_outline).
          out->write( |READ row outline { ls_outline-row_from }-{ ls_outline-row_to } collapsed={ ls_outline-collapsed }| ) ##NO_TEXT.
        ENDLOOP.
        out->write( |READ column B outline level={ lo_sheet->get_column( 'B' )->get_outline_level( ) }| ) ##NO_TEXT.
        CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
        DATA(lv_outline_rt) = lo_writer->write_file( lo_read ).
        out->write( |OUTLINE ROUNDTRIP OK, { xstrlen( lv_outline_rt ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `OUTLRT` iv_file = lv_outline_rt io_out = out ).

        " 10) autofilter + tables: write, read back, write again (round trip)
        DATA(lv_aftab) = build_autofilter_table_xlsx( ).
        out->write( |AFTAB XLSX OK, { xstrlen( lv_aftab ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `AFTAB` iv_file = lv_aftab io_out = out ).
        lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_aftab ).
        lo_sheet = lo_read->get_worksheet_by_name( 'AutoFilter' ) ##NO_TEXT.
        DATA(lo_af_read) = lo_read->get_autofilters_reference( )->get( lo_sheet ).
        IF lo_af_read IS BOUND.
          out->write( |READ autofilter ref={ lo_af_read->get_filter_reference( ) } values={ lines( lo_af_read->get_values( ) ) }| ) ##NO_TEXT.
          LOOP AT lo_af_read->get_filters( ) INTO DATA(ls_af_filter) USING KEY rule WHERE rule = zcl_excel_autofilter=>mc_filter_rule_text_pattern.
            out->write( |READ text filter column { ls_af_filter-column }: { VALUE #( ls_af_filter-tr_textfilter1[ 1 ]-low OPTIONAL ) }| &&
                        | { ls_af_filter-logical_operator } { VALUE #( ls_af_filter-tr_textfilter2[ 1 ]-low OPTIONAL ) }| ) ##NO_TEXT.
          ENDLOOP.
          DATA(lv_hidden) = ``.
          DO 13 TIMES.
            IF lo_af_read->is_row_hidden( sy-index ) = abap_true.
              lv_hidden = |{ lv_hidden } { sy-index }|.
            ENDIF.
          ENDDO.
          out->write( |READ autofilter hidden rows:{ lv_hidden }| ) ##NO_TEXT.
        ELSE.
          out->write( |READ autofilter: none| ) ##NO_TEXT.
        ENDIF.
        lo_sheet = lo_read->get_worksheet_by_name( 'Table' ) ##NO_TEXT.
        DATA(lo_tab_iter) = lo_sheet->get_tables_iterator( ).
        WHILE lo_tab_iter->has_next( ) = abap_true.
          DATA(lo_table) = CAST zcl_excel_table( lo_tab_iter->get_next( ) ).
          out->write( |READ table { lo_table->settings-table_name } style={ lo_table->settings-table_style }| &&
                      | range={ lo_table->settings-top_left_column }{ lo_table->settings-top_left_row }:| &&
                      |{ lo_table->settings-bottom_right_column }{ lo_table->settings-bottom_right_row }| &&
                      | rows={ lo_table->settings-show_row_stripes } cols={ lo_table->settings-show_column_stripes }| &&
                      | nofilters={ lo_table->settings-nofilters }| ) ##NO_TEXT.
          LOOP AT lo_table->fieldcat INTO DATA(ls_tab_fc).
            out->write( |READ   column { ls_tab_fc-column_name } totals={ ls_tab_fc-totals_function }| ) ##NO_TEXT.
          ENDLOOP.
        ENDWHILE.
        CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
        DATA(lv_aftab_rt) = lo_writer->write_file( lo_read ).
        out->write( |AFTAB ROUNDTRIP OK, { xstrlen( lv_aftab_rt ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `AFTABRT` iv_file = lv_aftab_rt io_out = out ).

        " 11) rich text + merged cells: write, read back, write again (round trip)
        DATA(lv_rtmerge) = build_richtext_merge_xlsx( ).
        out->write( |RTMERGE XLSX OK, { xstrlen( lv_rtmerge ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `RTMERGE` iv_file = lv_rtmerge io_out = out ).
        lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_rtmerge ).
        lo_sheet = lo_read->get_worksheet_by_name( 'RichText' ) ##NO_TEXT.
        DO 7 TIMES.
          lo_sheet->get_cell( EXPORTING ip_column = 'A' ip_row = sy-index
                              IMPORTING ep_value = DATA(lv_rt_value) et_rtf = DATA(lt_rt_runs) ).
          DATA(lv_rt_desc) = ``.
          LOOP AT lt_rt_runs INTO DATA(ls_rt_run).
            lv_rt_desc = |{ lv_rt_desc } [{ ls_rt_run-offset },{ ls_rt_run-length }| &&
                         |{ COND #( WHEN ls_rt_run-font-bold = abap_true THEN ` b` ) }| &&
                         |{ COND #( WHEN ls_rt_run-font-italic = abap_true THEN ` i` ) }| &&
                         |{ COND #( WHEN ls_rt_run-font-underline = abap_true THEN ` u` ) }| &&
                         |{ COND #( WHEN ls_rt_run-font-strikethrough = abap_true THEN ` s` ) }| &&
                         |{ COND #( WHEN ls_rt_run-font-color-rgb IS NOT INITIAL THEN ` ` && ls_rt_run-font-color-rgb ) }| &&
                         | { ls_rt_run-font-name } { ls_rt_run-font-size }]| ##NO_TEXT.
          ENDLOOP.
          out->write( |READ A{ sy-index } len={ strlen( lv_rt_value ) } runs={ lines( lt_rt_runs ) }:{ lv_rt_desc }| ) ##NO_TEXT.
        ENDDO.
        lo_sheet = lo_read->get_worksheet_by_name( 'Merge' ) ##NO_TEXT.
        out->write( |READ merges: { concat_lines_of( table = lo_sheet->get_merge( ) sep = ` ` ) }| ) ##NO_TEXT.
        lo_sheet->get_cell( EXPORTING ip_columnrow = 'B7' IMPORTING ep_value = DATA(lv_m_value) ep_formula = DATA(lv_m_formula) ) ##NO_TEXT.
        out->write( |READ B7 value={ lv_m_value } formula={ lv_m_formula }| ) ##NO_TEXT.
        lo_sheet->get_cell( EXPORTING ip_columnrow = 'E7' IMPORTING ep_value = lv_m_value ep_formula = lv_m_formula ) ##NO_TEXT.
        out->write( |READ E7 value={ lv_m_value } formula={ lv_m_formula }| ) ##NO_TEXT.
        lo_sheet->get_cell( EXPORTING ip_columnrow = 'H1' IMPORTING ep_value = lv_m_value ) ##NO_TEXT.
        out->write( |READ H1 { lv_m_value }| ) ##NO_TEXT.
        CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
        DATA(lv_rtmerge_rt) = lo_writer->write_file( lo_read ).
        out->write( |RTMERGE ROUNDTRIP OK, { xstrlen( lv_rtmerge_rt ) } bytes| ) ##NO_TEXT.
        out_base64( iv_tag = `RTMERGERT` iv_file = lv_rtmerge_rt io_out = out ).

        " 12) images + drawings: write, read back, write again (round trip)
        TRY.
            DATA(lv_images) = build_image_drawing_xlsx( ).
            out->write( |IMAGES XLSX OK, { xstrlen( lv_images ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `IMAGES` iv_file = lv_images io_out = out ).
            lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_images ).
            DATA(lo_ws_iter) = lo_read->get_worksheets_iterator( ).
            WHILE lo_ws_iter->has_next( ) = abap_true.
              lo_sheet = CAST zcl_excel_worksheet( lo_ws_iter->get_next( ) ).
              DATA(lo_dr_iter) = lo_sheet->get_drawings( )->get_iterator( ).
              WHILE lo_dr_iter->has_next( ) = abap_true.
                DATA(lo_dr) = CAST zcl_excel_drawing( lo_dr_iter->get_next( ) ).
                DATA(ls_dr_pos) = lo_dr->get_position( ).
                out->write( |READ { lo_sheet->get_title( ) } drawing type={ lo_dr->get_type( ) } anchor={ ls_dr_pos-anchor }| &&
                            | from={ ls_dr_pos-from-col }/{ ls_dr_pos-from-row }+{ ls_dr_pos-from-col_offset }/{ ls_dr_pos-from-row_offset }| &&
                            | to={ ls_dr_pos-to-col }/{ ls_dr_pos-to-row } size={ ls_dr_pos-size-width }x{ ls_dr_pos-size-height }| &&
                            | media={ lo_dr->get_media_type( ) } { xstrlen( lo_dr->get_media( ) ) } bytes| &&
                            | same_png={ xsdbool( lo_dr->get_media( ) = get_test_image( 'PNG' ) ) }| &&
                            | same_jpg={ xsdbool( lo_dr->get_media( ) = get_test_image( 'JPG' ) ) }| ) ##NO_TEXT.
              ENDWHILE.
            ENDWHILE.
            lo_sheet = lo_read->get_worksheet_by_name( 'HeaderFooter' ) ##NO_TEXT.
            out->write( |READ HeaderFooter header/footer images={ lines( lo_sheet->get_header_footer_drawings( ) ) }| ) ##NO_TEXT.
            lo_sheet->sheet_setup->get_header_footer_string( IMPORTING ep_odd_header = DATA(lv_img_header) ep_odd_footer = DATA(lv_img_footer) ).
            out->write( |READ HeaderFooter header={ lv_img_header } footer={ lv_img_footer }| ) ##NO_TEXT.
            LOOP AT lo_sheet->get_header_footer_drawings( ) INTO DATA(ls_hf_drawing).
              DATA(lo_hf_drawing) = CAST zcl_excel_drawing( ls_hf_drawing-drawing ).
              out->write( |READ HeaderFooter image { lo_hf_drawing->get_media_type( ) } { xstrlen( lo_hf_drawing->get_media( ) ) } bytes| &&
                          | size={ lo_hf_drawing->get_position( )-size-width }x{ lo_hf_drawing->get_position( )-size-height }| ) ##NO_TEXT.
            ENDLOOP.
            CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
            DATA(lv_images_rt) = lo_writer->write_file( lo_read ).
            out->write( |IMAGES ROUNDTRIP OK, { xstrlen( lv_images_rt ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `IMAGESRT` iv_file = lv_images_rt io_out = out ).
          CATCH cx_root INTO DATA(lx_images).
            out->write( |IMAGES EXCEPTION { cl_abap_classdescr=>get_class_name( lx_images ) }: { lx_images->get_text( ) }| ) ##NO_TEXT.
        ENDTRY.

        " 13) freeze panes + column widths: write, read back, write again (round trip)
        TRY.
            DATA(lv_freeze) = build_freeze_width_xlsx( ).
            out->write( |FREEZE XLSX OK, { xstrlen( lv_freeze ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `FREEZE` iv_file = lv_freeze io_out = out ).
            lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_freeze ).
            DATA(lo_fw_iter) = lo_read->get_worksheets_iterator( ).
            WHILE lo_fw_iter->has_next( ) = abap_true.
              lo_sheet = CAST zcl_excel_worksheet( lo_fw_iter->get_next( ) ).
              lo_sheet->get_freeze_cell( IMPORTING ep_row = DATA(lv_fr_row) ep_column = DATA(lv_fr_col) ).
              out->write( |READ { lo_sheet->get_title( ) } freeze cell row={ lv_fr_row } col={ lv_fr_col }| &&
                          | pane_top_left={ lo_sheet->pane_top_left_cell }| ) ##NO_TEXT.
            ENDWHILE.
            lo_sheet = lo_read->get_worksheet_by_name( 'Widths' ) ##NO_TEXT.
            out->write( |READ Widths default width={ lo_sheet->get_default_column( )->get_width( ) DECIMALS = 2 }| ) ##NO_TEXT.
            DATA(lo_col_iter) = lo_sheet->get_columns_iterator( ).
            WHILE lo_col_iter->has_next( ) = abap_true.
              DATA(lo_fw_col) = CAST zcl_excel_column( lo_col_iter->get_next( ) ).
              out->write( |READ Widths column { zcl_excel_common=>convert_column2alpha( lo_fw_col->get_column_index( ) ) }| &&
                          | width={ lo_fw_col->get_width( ) DECIMALS = 2 } auto={ lo_fw_col->get_auto_size( ) }| &&
                          | visible={ lo_fw_col->get_visible( ) }| ) ##NO_TEXT.
            ENDWHILE.
            DO 5 TIMES.
              DATA(lo_fw_row) = lo_sheet->get_row( sy-index ).
              out->write( |READ Widths row { sy-index } height={ lo_fw_row->get_row_height( ) } visible={ lo_fw_row->get_visible( ) }| ) ##NO_TEXT.
            ENDDO.
            CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
            DATA(lv_freeze_rt) = lo_writer->write_file( lo_read ).
            out->write( |FREEZE ROUNDTRIP OK, { xstrlen( lv_freeze_rt ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `FREEZERT` iv_file = lv_freeze_rt io_out = out ).
          CATCH cx_root INTO DATA(lx_freeze).
            out->write( |FREEZE EXCEPTION { cl_abap_classdescr=>get_class_name( lx_freeze ) }: { lx_freeze->get_text( ) }| ) ##NO_TEXT.
        ENDTRY.

        " 14) named ranges + cross-sheet formulas / 15) advanced styles: write, read back, write again
        TRY.
            DATA(lv_names) = build_named_range_xlsx( ).
            out->write( |NAMES XLSX OK, { xstrlen( lv_names ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `NAMES` iv_file = lv_names io_out = out ).
            lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_names ).
            DATA(lo_rng_iter) = lo_read->get_ranges_iterator( ).
            WHILE lo_rng_iter->has_next( ) = abap_true.
              DATA(lo_rng) = CAST zcl_excel_range( lo_rng_iter->get_next( ) ).
              out->write( |READ global range { lo_rng->name } = { lo_rng->get_value( ) }| ) ##NO_TEXT.
            ENDWHILE.
            DATA(lo_nm_ws_iter) = lo_read->get_worksheets_iterator( ).
            WHILE lo_nm_ws_iter->has_next( ) = abap_true.
              lo_sheet = CAST zcl_excel_worksheet( lo_nm_ws_iter->get_next( ) ).
              lo_rng_iter = lo_sheet->get_ranges_iterator( ).
              WHILE lo_rng_iter->has_next( ) = abap_true.
                lo_rng = CAST zcl_excel_range( lo_rng_iter->get_next( ) ).
                out->write( |READ local range { lo_sheet->get_title( ) }!{ lo_rng->name } = { lo_rng->get_value( ) }| ) ##NO_TEXT.
              ENDWHILE.
            ENDWHILE.
            lo_sheet = lo_read->get_worksheet_by_name( 'Report' ) ##NO_TEXT.
            DO 11 TIMES.
              lo_sheet->get_cell( EXPORTING ip_column = 'B' ip_row = sy-index IMPORTING ep_formula = DATA(lv_nm_formula) ).
              IF lv_nm_formula IS NOT INITIAL.
                out->write( |READ Report!B{ sy-index } formula={ lv_nm_formula }| ) ##NO_TEXT.
              ENDIF.
            ENDDO.
            CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
            DATA(lv_names_rt) = lo_writer->write_file( lo_read ).
            out->write( |NAMES ROUNDTRIP OK, { xstrlen( lv_names_rt ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `NAMESRT` iv_file = lv_names_rt io_out = out ).
          CATCH cx_root INTO DATA(lx_names).
            out->write( |NAMES EXCEPTION { cl_abap_classdescr=>get_class_name( lx_names ) }: { lx_names->get_text( ) }| ) ##NO_TEXT.
        ENDTRY.

        TRY.
            DATA(lv_adv) = build_advanced_style_xlsx( ).
            out->write( |ADVSTYLE XLSX OK, { xstrlen( lv_adv ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `ADVSTYLE` iv_file = lv_adv io_out = out ).
            lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_adv ).
            lo_sheet = lo_read->get_worksheet_by_name( 'Styles' ) ##NO_TEXT.
            DO 9 TIMES.
              lo_sheet->get_cell( EXPORTING ip_column = 'A' ip_row = sy-index + 1 IMPORTING ep_value = DATA(lv_adv_value) ep_style = DATA(lo_adv_style) ).
              IF lo_adv_style IS BOUND AND lv_adv_value IS NOT INITIAL.
                out->write( |READ A{ sy-index + 1 } { lv_adv_value } filltype={ lo_adv_style->fill->filltype }| &&
                            | gradtype={ lo_adv_style->fill->gradtype-type } degree={ lo_adv_style->fill->gradtype-degree }| &&
                            | fg={ lo_adv_style->fill->fgcolor-rgb } bg={ lo_adv_style->fill->bgcolor-rgb }| ) ##NO_TEXT.
              ENDIF.
            ENDDO.
            DO 6 TIMES.
              lo_sheet->get_cell( EXPORTING ip_column = sy-index ip_row = 12 IMPORTING ep_style = lo_adv_style ).
              IF lo_adv_style IS BOUND.
                out->write( |READ row12 col{ sy-index } textRotation={ lo_adv_style->alignment->textrotation }| ) ##NO_TEXT.
              ENDIF.
            ENDDO.
            DO 4 TIMES.
              lo_sheet->get_cell( EXPORTING ip_column = 'A' ip_row = 14 + sy-index IMPORTING ep_style = lo_adv_style ).
              IF lo_adv_style IS BOUND.
                out->write( |READ A{ 14 + sy-index } indent={ lo_adv_style->alignment->indent } h={ lo_adv_style->alignment->horizontal }| ) ##NO_TEXT.
              ENDIF.
            ENDDO.
            lo_sheet->get_cell( EXPORTING ip_columnrow = 'B20' IMPORTING ep_style = lo_adv_style ) ##NO_TEXT.
            out->write( |READ B20 shrink={ lo_adv_style->alignment->shrinktofit }| ) ##NO_TEXT.
            lo_sheet->get_cell( EXPORTING ip_columnrow = 'C20' IMPORTING ep_style = lo_adv_style ) ##NO_TEXT.
            out->write( |READ C20 h={ lo_adv_style->alignment->horizontal } v={ lo_adv_style->alignment->vertical } wrap={ lo_adv_style->alignment->wraptext }| ) ##NO_TEXT.
            lo_sheet->get_cell( EXPORTING ip_columnrow = 'D20' IMPORTING ep_style = lo_adv_style ) ##NO_TEXT.
            out->write( |READ D20 h={ lo_adv_style->alignment->horizontal }| ) ##NO_TEXT.
            CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
            DATA(lv_adv_rt) = lo_writer->write_file( lo_read ).
            out->write( |ADVSTYLE ROUNDTRIP OK, { xstrlen( lv_adv_rt ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `ADVSTYLERT` iv_file = lv_adv_rt io_out = out ).
          CATCH cx_root INTO DATA(lx_adv).
            out->write( |ADVSTYLE EXCEPTION { cl_abap_classdescr=>get_class_name( lx_adv ) }: { lx_adv->get_text( ) }| ) ##NO_TEXT.
        ENDTRY.

        " 16) read a workbook created by Microsoft Excel, then write it again
        TRY.
            DATA(lv_xl_file) = get_excel_created_file( ).
            lo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( lv_xl_file ).
            out->write( |XLREAD OK, { xstrlen( lv_xl_file ) } bytes, sheets={ lo_read->get_worksheets_size( ) }| ) ##NO_TEXT.
            DATA(lo_xl_iter) = lo_read->get_worksheets_iterator( ).
            WHILE lo_xl_iter->has_next( ) = abap_true.
              lo_sheet = CAST zcl_excel_worksheet( lo_xl_iter->get_next( ) ).
              lo_sheet->get_freeze_cell( IMPORTING ep_row = DATA(lv_xl_frow) ep_column = DATA(lv_xl_fcol) ).
              DATA(lo_xl_cf_iter) = lo_sheet->get_style_cond_iterator( ).
              DATA(lv_xl_cf) = 0.
              WHILE lo_xl_cf_iter->has_next( ) = abap_true.
                lo_xl_cf_iter->get_next( ).
                lv_xl_cf = lv_xl_cf + 1.
              ENDWHILE.
              out->write( |XLREAD sheet { lo_sheet->get_title( ) } dim={ lo_sheet->get_dimension_range( ) }| &&
                          | merges={ concat_lines_of( table = lo_sheet->get_merge( ) sep = ` ` ) }| &&
                          | freeze={ lv_xl_frow }/{ lv_xl_fcol } tables={ lo_sheet->get_tables_size( ) }| &&
                          | drawings={ lo_sheet->get_drawings( )->size( ) } comments={ lo_sheet->get_comments( )->size( ) }| &&
                          | hyperlinks={ lo_sheet->get_hyperlinks_size( ) } validations={ lo_sheet->get_data_validations_size( ) }| &&
                          | condformats={ lv_xl_cf }| ) ##NO_TEXT.
            ENDWHILE.
            lo_sheet = lo_read->get_worksheet_by_name( 'DuLieu' ) ##NO_TEXT.
            DO 13 TIMES.
              DATA(lv_xl_row) = sy-index.
              DATA(lv_xl_line) = ``.
              DO 8 TIMES.
                lo_sheet->get_cell( EXPORTING ip_column = sy-index ip_row = lv_xl_row
                                    IMPORTING ep_value = DATA(lv_xl_value) ep_formula = DATA(lv_xl_formula) ep_style = DATA(lo_xl_style) ).
                IF lv_xl_formula IS NOT INITIAL.
                  lv_xl_line = |{ lv_xl_line } \| ={ lv_xl_formula }|.
                ELSEIF lv_xl_value IS NOT INITIAL.
                  lv_xl_line = |{ lv_xl_line } \| { lv_xl_value }|.
                  IF lo_xl_style IS BOUND AND lo_xl_style->number_format->format_code <> 'General'.
                    lv_xl_line = |{ lv_xl_line } [{ lo_xl_style->number_format->format_code }]|.
                  ENDIF.
                ENDIF.
              ENDDO.
              IF lv_xl_line IS NOT INITIAL.
                out->write( |XLREAD row { lv_xl_row }:{ lv_xl_line }| ) ##NO_TEXT.
              ENDIF.
            ENDDO.
            lo_sheet->get_cell( EXPORTING ip_columnrow = 'B6' IMPORTING et_rtf = DATA(lt_xl_rtf) ) ##NO_TEXT.
            out->write( |XLREAD B6 rich text runs={ lines( lt_xl_rtf ) }| ) ##NO_TEXT.
            DATA(lo_xl_rng_iter) = lo_read->get_ranges_iterator( ).
            WHILE lo_xl_rng_iter->has_next( ) = abap_true.
              DATA(lo_xl_rng) = CAST zcl_excel_range( lo_xl_rng_iter->get_next( ) ).
              out->write( |XLREAD range { lo_xl_rng->name } = { lo_xl_rng->get_value( ) }| ) ##NO_TEXT.
            ENDWHILE.
            CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
            DATA(lv_xl_rt) = lo_writer->write_file( lo_read ).
            out->write( |XLREAD ROUNDTRIP OK, { xstrlen( lv_xl_rt ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `XLREADRT` iv_file = lv_xl_rt io_out = out ).
          CATCH cx_root INTO DATA(lx_xl).
            out->write( |XLREAD EXCEPTION { cl_abap_classdescr=>get_class_name( lx_xl ) }: { lx_xl->get_text( ) }| ) ##NO_TEXT.
        ENDTRY.

        " 17) macro enabled workbook (.xlsm): write with ZCL_EXCEL_WRITER_XLSM, read with ZCL_EXCEL_READER_XLSM, write again
        TRY.
            DATA(lv_xlsm) = build_xlsm( ).
            out->write( |XLSM OK, { xstrlen( lv_xlsm ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `XLSM` iv_file = lv_xlsm io_out = out ).
            lo_read = NEW zcl_excel_reader_xlsm( )->zif_excel_reader~load( lv_xlsm ).
            lo_sheet = lo_read->get_active_worksheet( ).
            out->write( |READ XLSM vba={ xstrlen( lo_read->zif_excel_book_vba_project~vbaproject ) } bytes| &&
                        | same={ xsdbool( lo_read->zif_excel_book_vba_project~vbaproject = get_test_vba_project( ) ) }| &&
                        | book codename={ lo_read->zif_excel_book_vba_project~codename } / { lo_read->zif_excel_book_vba_project~codename_pr }| &&
                        | sheet { lo_sheet->get_title( ) } codename={ lo_sheet->zif_excel_sheet_vba_project~codename_pr }| ) ##NO_TEXT.
            CREATE OBJECT lo_writer TYPE zcl_excel_writer_xlsm.
            DATA(lv_xlsm_rt) = lo_writer->write_file( lo_read ).
            out->write( |XLSM ROUNDTRIP OK, { xstrlen( lv_xlsm_rt ) } bytes| ) ##NO_TEXT.
            out_base64( iv_tag = `XLSMRT` iv_file = lv_xlsm_rt io_out = out ).
          CATCH cx_root INTO DATA(lx_xlsm).
            out->write( |XLSM EXCEPTION { cl_abap_classdescr=>get_class_name( lx_xlsm ) }: { lx_xlsm->get_text( ) }| ) ##NO_TEXT.
        ENDTRY.

        " 4) reader: LOAD_FILE must raise in ABAP Cloud
        TRY.
            lo_reader->load_file( 'dummy.xlsx' ).
            out->write( |LOAD_FILE: no exception (unexpected)| ) ##NO_TEXT.
          CATCH zcx_excel INTO lx_excel.
            out->write( |LOAD_FILE raised as expected: { lx_excel->get_text( ) }| ) ##NO_TEXT.
        ENDTRY.

      CATCH zcx_excel INTO lx_excel.
        out->write( |ZCX_EXCEL: { lx_excel->get_text( ) }| ).
      CATCH cx_root INTO lx_root.
        out->write( |EXCEPTION { cl_abap_classdescr=>get_class_name( lx_root ) }: { lx_root->get_text( ) }| ).
    ENDTRY.
  ENDMETHOD.

  METHOD get_demo_data.
    " Demo data for custom entity ZC_DEMO_ABAP2XLSX (group 1 = figures of template Mau1.xlsx)
    CONSTANTS: lc_comp_1000 TYPE c LENGTH 255 VALUE 'Công ty TNHH Thương mại & Dịch vụ AB' ##NO_TEXT,
               lc_addr_1000 TYPE c LENGTH 255 VALUE 'Đường số 7, KCN Tiên Sơn, Phường Từ Sơn, Thành phố Bắc Ninh' ##NO_TEXT,
               lc_comp_2000 TYPE c LENGTH 255 VALUE 'Công ty Cổ phần Sản xuất AB Miền Nam' ##NO_TEXT,
               lc_addr_2000 TYPE c LENGTH 255 VALUE 'Lô B2, KCN Tân Tạo, Quận Bình Tân, Thành phố Hồ Chí Minh' ##NO_TEXT.
    DATA ls_row TYPE ty_s_journal.

    " customer 0100098012 in company code 1000 (Mau1.xlsx)
    ls_row = VALUE #( companycode    = '1000'
                      fiscalyear     = '2026'
                      accounting_doc = '1000900001'
                      customer       = '0100098012'
                      customer_name  = 'Công ty Dịch vụ H&M'
                      customer_addr  = 'số 1A, đường Hoàng Quốc Việt, phường Cầu Giấy, Thành phố Hà Nội'
                      identification = '00102234556'
                      phone_number   = '(+84) 230405679'
                      email_addr     = 'coqa.hm@hm.com'
                      dvt            = 'KG'
                      currency       = 'VND'
                      company_name   = lc_comp_1000
                      company_addr   = lc_addr_1000
                      postingdate    = '20260928' ) ##NO_TEXT.
    rt_data = VALUE #( BASE rt_data
      ( VALUE #( BASE ls_row document_itm = '001' quantity = 1 amount_dr = 100000 ) )
      ( VALUE #( BASE ls_row document_itm = '002' quantity = 2 amount_cr = 12500 ) )
      ( VALUE #( BASE ls_row document_itm = '003' quantity = 3 amount_dr = 10000 ) )
      ( VALUE #( BASE ls_row document_itm = '004' quantity = 1 amount_cr = 23456 ) )
      ( VALUE #( BASE ls_row document_itm = '005' quantity = 4 amount_dr = 111 ) )
      ( VALUE #( BASE ls_row document_itm = '006' quantity = 3 amount_dr = 111 ) )
      ( VALUE #( BASE ls_row document_itm = '007' quantity = 4 amount_cr = 77599 ) )
      ( VALUE #( BASE ls_row document_itm = '008' quantity = 5 amount_dr = 3333 ) ) ).

    " customer 0100098020 in company code 1000
    ls_row = VALUE #( companycode    = '1000'
                      fiscalyear     = '2026'
                      accounting_doc = '1000900015'
                      customer       = '0100098020'
                      customer_name  = 'Công ty TNHH Bao bì Minh Phát'
                      customer_addr  = 'số 25, đường Nguyễn Văn Linh, phường Long Biên, Thành phố Hà Nội'
                      identification = '0109876543'
                      phone_number   = '(+84) 243556677'
                      email_addr     = 'kt@minhphat.vn'
                      dvt            = 'CAI'
                      currency       = 'VND'
                      company_name   = lc_comp_1000
                      company_addr   = lc_addr_1000
                      postingdate    = '20260925' ) ##NO_TEXT.
    rt_data = VALUE #( BASE rt_data
      ( VALUE #( BASE ls_row document_itm = '001' quantity = 10 amount_dr = 2500000 ) )
      ( VALUE #( BASE ls_row document_itm = '002' quantity = 5  amount_dr = 1250000 ) )
      ( VALUE #( BASE ls_row document_itm = '003' quantity = 2  amount_cr = 500000 ) ) ).

    " customer 0100098012 again, but in company code 2000 -> own sheet
    ls_row = VALUE #( companycode    = '2000'
                      fiscalyear     = '2026'
                      accounting_doc = '2000100007'
                      customer       = '0100098012'
                      customer_name  = 'Công ty Dịch vụ H&M'
                      customer_addr  = 'số 1A, đường Hoàng Quốc Việt, phường Cầu Giấy, Thành phố Hà Nội'
                      identification = '00102234556'
                      phone_number   = '(+84) 230405679'
                      email_addr     = 'coqa.hm@hm.com'
                      dvt            = 'KG'
                      currency       = 'VND'
                      company_name   = lc_comp_2000
                      company_addr   = lc_addr_2000
                      postingdate    = '20260926' ) ##NO_TEXT.
    rt_data = VALUE #( BASE rt_data
      ( VALUE #( BASE ls_row document_itm = '001' quantity = 7 amount_dr = 700000 ) )
      ( VALUE #( BASE ls_row document_itm = '002' quantity = 1 amount_cr = 150000 ) ) ).
  ENDMETHOD.


  METHOD build_journal_xlsx.
    " One worksheet per company code + customer, layout of template Mau1.xlsx
    CONSTANTS: lc_font_text   TYPE zexcel_style_font_name VALUE 'Times New Roman' ##NO_TEXT,
               lc_font_head   TYPE zexcel_style_font_name VALUE 'Consolas' ##NO_TEXT,
               lc_fmt_amount  TYPE zexcel_number_format VALUE '_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)' ##NO_TEXT,
               lc_fmt_date    TYPE zexcel_number_format VALUE 'dd/mm/yyyy' ##NO_TEXT,
               lc_fmt_text    TYPE zexcel_number_format VALUE '@' ##NO_TEXT,
               lc_first_row   TYPE i VALUE 14,
               lc_header_row  TYPE i VALUE 13.
    TYPES: BEGIN OF ty_s_column,
             column TYPE c LENGTH 1,
             title  TYPE string,
             width  TYPE string,
           END OF ty_s_column,
           ty_t_column TYPE STANDARD TABLE OF ty_s_column WITH EMPTY KEY.
    TYPES: BEGIN OF ty_s_customer_cocode,
             customer TYPE kunnr,
             count    TYPE i,
           END OF ty_s_customer_cocode.
    DATA: lo_excel       TYPE REF TO zcl_excel,
          lo_sheet       TYPE REF TO zcl_excel_worksheet,
          lo_writer      TYPE REF TO zif_excel_writer,
          lt_data        TYPE ty_t_journal,
          lt_columns     TYPE ty_t_column,
          lt_cust_count  TYPE HASHED TABLE OF ty_s_customer_cocode WITH UNIQUE KEY customer,
          lv_row         TYPE i,
          lv_last_row    TYPE i,
          lv_title       TYPE zexcel_sheet_title,
          lv_first_sheet TYPE abap_bool VALUE abap_true.

    lt_data = it_data.
    SORT lt_data BY companycode customer fiscalyear accounting_doc document_itm.

    " a customer that exists in more than one company code gets the company code in the sheet name
    LOOP AT lt_data INTO DATA(ls_cnt) GROUP BY ( companycode = ls_cnt-companycode customer = ls_cnt-customer ) INTO DATA(lg_cnt).
      READ TABLE lt_cust_count ASSIGNING FIELD-SYMBOL(<ls_cust_count>) WITH TABLE KEY customer = lg_cnt-customer.
      IF sy-subrc <> 0.
        INSERT VALUE #( customer = lg_cnt-customer ) INTO TABLE lt_cust_count ASSIGNING <ls_cust_count>.
      ENDIF.
      <ls_cust_count>-count = <ls_cust_count>-count + 1.
    ENDLOOP.

    CREATE OBJECT lo_excel.

    " ---- styles (shared by all worksheets)
    DATA(lo_border) = NEW zcl_excel_style_border( ).
    lo_border->border_style = zcl_excel_style_border=>c_border_thin.
    lo_border->border_color-rgb = zcl_excel_style_color=>c_black.

    DATA(lv_st_company) = new_style( io_excel = lo_excel iv_font = lc_font_text iv_bold = abap_true
                                     iv_halign = zcl_excel_style_alignment=>c_horizontal_left ).
    DATA(lv_st_line)    = new_style( io_excel = lo_excel iv_font = lc_font_text
                                     iv_halign = zcl_excel_style_alignment=>c_horizontal_left ).
    DATA(lv_st_title)   = new_style( io_excel = lo_excel iv_font = lc_font_text iv_bold = abap_true
                                     iv_halign = zcl_excel_style_alignment=>c_horizontal_center ).
    DATA(lv_st_head)    = new_style( io_excel = lo_excel iv_font = lc_font_head iv_size = 10 iv_bold = abap_true
                                     iv_halign = zcl_excel_style_alignment=>c_horizontal_center io_border = lo_border ).
    DATA(lv_st_center)  = new_style( io_excel = lo_excel iv_font = lc_font_text io_border = lo_border
                                     iv_halign = zcl_excel_style_alignment=>c_horizontal_center iv_valign = space ).
    DATA(lv_st_docno)   = new_style( io_excel = lo_excel iv_font = lc_font_text io_border = lo_border iv_format = lc_fmt_text
                                     iv_halign = zcl_excel_style_alignment=>c_horizontal_left iv_valign = space ).
    DATA(lv_st_date)    = new_style( io_excel = lo_excel iv_font = lc_font_text io_border = lo_border iv_format = lc_fmt_date
                                     iv_halign = zcl_excel_style_alignment=>c_horizontal_center iv_valign = space ).
    DATA(lv_st_number)  = new_style( io_excel = lo_excel iv_font = lc_font_text io_border = lo_border iv_valign = space ).
    DATA(lv_st_amount)  = new_style( io_excel = lo_excel iv_font = lc_font_text io_border = lo_border iv_format = lc_fmt_amount
                                     iv_valign = space ).
    DATA(lv_st_tot_lbl) = new_style( io_excel = lo_excel iv_font = lc_font_text iv_bold = abap_true io_border = lo_border
                                     iv_halign = zcl_excel_style_alignment=>c_horizontal_center iv_valign = space ).
    DATA(lv_st_tot_amt) = new_style( io_excel = lo_excel iv_font = lc_font_text iv_bold = abap_true io_border = lo_border
                                     iv_format = lc_fmt_amount iv_valign = space ).

    lt_columns = VALUE #( ( column = 'A' title = 'Document_Itm'   width = '13.89' )
                          ( column = 'B' title = 'Accounting_doc' width = '17.11' )
                          ( column = 'C' title = 'PostingDate'    width = '14.89' )
                          ( column = 'D' title = 'Quantity'       width = '11.11' )
                          ( column = 'E' title = 'DVT'            width = '9.78' )
                          ( column = 'F' title = 'Amount_Dr'      width = '16.78' )
                          ( column = 'G' title = 'Amount_Cr'      width = '16.11' )
                          ( column = 'H' title = 'Currency'       width = '11.44' ) ) ##NO_TEXT.

    " ---- one worksheet per company code + customer
    LOOP AT lt_data INTO DATA(ls_first) GROUP BY ( companycode = ls_first-companycode customer = ls_first-customer )
         INTO DATA(lg_sheet).

      IF lv_first_sheet = abap_true.
        lo_sheet = lo_excel->get_active_worksheet( ).
        lv_first_sheet = abap_false.
      ELSE.
        lo_sheet = lo_excel->add_new_worksheet( ).
      ENDIF.
      IF lt_cust_count[ customer = lg_sheet-customer ]-count > 1.
        lv_title = |{ lg_sheet-customer }_{ lg_sheet-companycode }|.
      ELSE.
        lv_title = lg_sheet-customer.
      ENDIF.
      lo_sheet->set_title( lv_title ).

      LOOP AT lt_columns INTO DATA(ls_column).
        lo_sheet->get_column( ls_column-column )->set_width( ls_column-width ).
      ENDLOOP.

      " header block (company, title, customer)
      LOOP AT GROUP lg_sheet INTO DATA(ls_head).
        EXIT.
      ENDLOOP.
      set_line( io_sheet = lo_sheet iv_row = 1  iv_style = lv_st_company iv_text = CONV #( ls_head-company_name ) ).
      set_line( io_sheet = lo_sheet iv_row = 2  iv_style = lv_st_line    iv_text = CONV #( ls_head-company_addr ) ).
      set_line( io_sheet = lo_sheet iv_row = 5  iv_style = lv_st_title   iv_text = `NHẬT KÝ GIAO DỊCH` ) ##NO_TEXT.
      set_line( io_sheet = lo_sheet iv_row = 6  iv_style = lv_st_line
                iv_text = |Nhà cung cấp: { ls_head-customer_name }| ) ##NO_TEXT.
      set_line( io_sheet = lo_sheet iv_row = 7  iv_style = lv_st_line
                iv_text = |Địa chỉ nhà cung cấp: { ls_head-customer_addr }| ) ##NO_TEXT.
      set_line( io_sheet = lo_sheet iv_row = 8  iv_style = lv_st_line
                iv_text = |Mã số thuế: { ls_head-identification }| ) ##NO_TEXT.
      set_line( io_sheet = lo_sheet iv_row = 9  iv_style = lv_st_line
                iv_text = |Số điện thoại: { ls_head-phone_number }| ) ##NO_TEXT.
      set_line( io_sheet = lo_sheet iv_row = 10 iv_style = lv_st_line
                iv_text = |Email: { ls_head-email_addr }| ) ##NO_TEXT.

      " table header
      lo_sheet->get_row( lc_header_row )->set_row_height( '19.2' ).
      LOOP AT lt_columns INTO ls_column.
        lo_sheet->set_cell( ip_column = ls_column-column ip_row = lc_header_row
                            ip_value  = ls_column-title  ip_style = lv_st_head ).
      ENDLOOP.

      " items
      lv_row = lc_first_row.
      LOOP AT GROUP lg_sheet INTO DATA(ls_item).
        lo_sheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = CONV i( ls_item-document_itm ) ip_style = lv_st_center ).
        lo_sheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_item-accounting_doc ip_style = lv_st_docno ).
        lo_sheet->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_item-postingdate ip_style = lv_st_date ).
        lo_sheet->set_cell( ip_column = 'D' ip_row = lv_row ip_value = ls_item-quantity ip_style = lv_st_number ).
        lo_sheet->set_cell( ip_column = 'E' ip_row = lv_row ip_value = ls_item-dvt ip_style = lv_st_center ).
        set_amount( io_sheet = lo_sheet iv_column = 'F' iv_row = lv_row iv_amount = ls_item-amount_dr iv_style = lv_st_amount ).
        set_amount( io_sheet = lo_sheet iv_column = 'G' iv_row = lv_row iv_amount = ls_item-amount_cr iv_style = lv_st_amount ).
        lo_sheet->set_cell( ip_column = 'H' ip_row = lv_row ip_value = ls_item-currency ip_style = lv_st_center ).
        lv_row = lv_row + 1.
      ENDLOOP.
      lv_last_row = lv_row - 1.

      " total line
      lo_sheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = `Tổng cộng` ip_style = lv_st_tot_lbl ) ##NO_TEXT.
      lo_sheet->set_cell_style( ip_column = 'B' ip_row = lv_row ip_style = lv_st_tot_lbl ).
      lo_sheet->set_cell_style( ip_column = 'C' ip_row = lv_row ip_style = lv_st_tot_lbl ).
      lo_sheet->set_cell_style( ip_column = 'D' ip_row = lv_row ip_style = lv_st_tot_lbl ).
      lo_sheet->set_cell_style( ip_column = 'E' ip_row = lv_row ip_style = lv_st_tot_lbl ).
      lo_sheet->set_merge( ip_range = |A{ lv_row }:E{ lv_row }| ).
      lo_sheet->set_cell( ip_column = 'F' ip_row = lv_row ip_style = lv_st_tot_amt
                          ip_formula = |SUM(F{ lc_first_row }:F{ lv_last_row })| ) ##NO_TEXT.
      lo_sheet->set_cell( ip_column = 'G' ip_row = lv_row ip_style = lv_st_tot_amt
                          ip_formula = |SUM(G{ lc_first_row }:G{ lv_last_row })| ) ##NO_TEXT.
      lo_sheet->set_cell_style( ip_column = 'H' ip_row = lv_row ip_style = lv_st_tot_lbl ).
    ENDLOOP.

    IF lv_first_sheet = abap_true.
      " no data: keep an empty sheet with the title only
      set_line( io_sheet = lo_excel->get_active_worksheet( ) iv_row = 5 iv_style = lv_st_title
                iv_text = `NHẬT KÝ GIAO DỊCH` ) ##NO_TEXT.
    ENDIF.

    lo_excel->set_active_sheet_index( 1 ).
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD build_chart_validation_xlsx.
    " Test workbook: bar / line / pie chart on sheet "Data", data validations on sheet "Validation"
    TYPES: BEGIN OF ty_s_month,
             month   TYPE string,
             revenue TYPE i,
             cost    TYPE i,
           END OF ty_s_month.
    DATA: lo_excel    TYPE REF TO zcl_excel,
          lo_data     TYPE REF TO zcl_excel_worksheet,
          lo_valid    TYPE REF TO zcl_excel_worksheet,
          lo_lists    TYPE REF TO zcl_excel_worksheet,
          lo_drawing  TYPE REF TO zcl_excel_drawing,
          lo_bars     TYPE REF TO zcl_excel_graph_bars,
          lo_line     TYPE REF TO zcl_excel_graph_line,
          lo_pie      TYPE REF TO zcl_excel_graph_pie,
          lo_dv       TYPE REF TO zcl_excel_data_validation,
          lo_writer   TYPE REF TO zif_excel_writer,
          lt_months   TYPE STANDARD TABLE OF ty_s_month WITH EMPTY KEY,
          lv_row      TYPE i,
          lv_sheet    TYPE zexcel_sheet_title.

    CREATE OBJECT lo_excel.

    " ---- sheet "Data": source values of the charts
    lo_data = lo_excel->get_active_worksheet( ).
    lv_sheet = 'Data' ##NO_TEXT.
    lo_data->set_title( lv_sheet ).
    lt_months = VALUE #( ( month = 'T1' revenue = 120 cost = 80 )
                         ( month = 'T2' revenue = 150 cost = 95 )
                         ( month = 'T3' revenue = 90  cost = 70 )
                         ( month = 'T4' revenue = 180 cost = 110 )
                         ( month = 'T5' revenue = 210 cost = 130 )
                         ( month = 'T6' revenue = 160 cost = 100 ) ) ##NO_TEXT.
    lo_data->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Tháng` ) ##NO_TEXT.
    lo_data->set_cell( ip_column = 'B' ip_row = 1 ip_value = `Doanh thu` ) ##NO_TEXT.
    lo_data->set_cell( ip_column = 'C' ip_row = 1 ip_value = `Chi phí` ) ##NO_TEXT.
    lv_row = 2.
    LOOP AT lt_months INTO DATA(ls_month).
      lo_data->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_month-month ).
      lo_data->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_month-revenue ).
      lo_data->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_month-cost ).
      lv_row = lv_row + 1.
    ENDLOOP.

    " bar chart (clustered): 2 series, categories A2:A7
    lo_drawing = lo_excel->add_new_drawing( ip_type = zcl_excel_drawing=>type_chart ip_title = 'BarChart' ) ##NO_TEXT.
    CREATE OBJECT lo_bars.
    lo_bars->create_serie( ip_order = 0 ip_sheet = lv_sheet ip_sername = `Doanh thu`
                           ip_lbl_from_col = 'A' ip_lbl_from_row = 2 ip_lbl_to_col = 'A' ip_lbl_to_row = 7
                           ip_ref_from_col = 'B' ip_ref_from_row = 2 ip_ref_to_col = 'B' ip_ref_to_row = 7 ) ##NO_TEXT.
    lo_bars->create_serie( ip_order = 1 ip_sheet = lv_sheet ip_sername = `Chi phí`
                           ip_lbl_from_col = 'A' ip_lbl_from_row = 2 ip_lbl_to_col = 'A' ip_lbl_to_row = 7
                           ip_ref_from_col = 'C' ip_ref_from_row = 2 ip_ref_to_col = 'C' ip_ref_to_row = 7 ) ##NO_TEXT.
    lo_bars->create_ax( ip_type = zcl_excel_graph_bars=>c_catax ).
    lo_bars->create_ax( ip_type = zcl_excel_graph_bars=>c_valax ).
    lo_bars->set_title( `Doanh thu và chi phí` ) ##NO_TEXT.
    lo_drawing->graph = lo_bars.
    lo_drawing->graph_type = zcl_excel_drawing=>c_graph_bars.
    lo_drawing->set_position2( ip_from = VALUE #( col = 4 row = 0 ) ip_to = VALUE #( col = 11 row = 15 ) ).
    lo_drawing->set_media( ip_media_type = zcl_excel_drawing=>c_media_type_xml ).
    lo_data->add_drawing( lo_drawing ).

    " line chart: revenue
    lo_drawing = lo_excel->add_new_drawing( ip_type = zcl_excel_drawing=>type_chart ip_title = 'LineChart' ) ##NO_TEXT.
    CREATE OBJECT lo_line.
    lo_line->create_serie( ip_order = 0 ip_sheet = lv_sheet ip_sername = `Doanh thu` ip_symbol = zcl_excel_graph_line=>c_symbol_auto
                           ip_lbl_from_col = 'A' ip_lbl_from_row = 2 ip_lbl_to_col = 'A' ip_lbl_to_row = 7
                           ip_ref_from_col = 'B' ip_ref_from_row = 2 ip_ref_to_col = 'B' ip_ref_to_row = 7 ) ##NO_TEXT.
    lo_line->create_ax( ip_type = zcl_excel_graph_line=>c_catax ).
    lo_line->create_ax( ip_type = zcl_excel_graph_line=>c_valax ).
    lo_line->set_title( `Xu hướng doanh thu` ) ##NO_TEXT.
    lo_drawing->graph = lo_line.
    lo_drawing->graph_type = zcl_excel_drawing=>c_graph_line.
    lo_drawing->set_position2( ip_from = VALUE #( col = 4 row = 16 ) ip_to = VALUE #( col = 11 row = 31 ) ).
    lo_drawing->set_media( ip_media_type = zcl_excel_drawing=>c_media_type_xml ).
    lo_data->add_drawing( lo_drawing ).

    " pie chart: cost share with percentages
    lo_drawing = lo_excel->add_new_drawing( ip_type = zcl_excel_drawing=>type_chart ip_title = 'PieChart' ) ##NO_TEXT.
    CREATE OBJECT lo_pie.
    lo_pie->create_serie( ip_order = 0 ip_sheet = lv_sheet ip_sername = `Chi phí`
                          ip_lbl_from_col = 'A' ip_lbl_from_row = 2 ip_lbl_to_col = 'A' ip_lbl_to_row = 7
                          ip_ref_from_col = 'C' ip_ref_from_row = 2 ip_ref_to_col = 'C' ip_ref_to_row = 7 ) ##NO_TEXT.
    lo_pie->set_show_percent( zcl_excel_graph_pie=>c_show_true ).
    lo_pie->set_title( `Cơ cấu chi phí` ) ##NO_TEXT.
    lo_drawing->graph = lo_pie.
    lo_drawing->graph_type = zcl_excel_drawing=>c_graph_pie.
    lo_drawing->set_position2( ip_from = VALUE #( col = 12 row = 0 ) ip_to = VALUE #( col = 19 row = 15 ) ).
    lo_drawing->set_media( ip_media_type = zcl_excel_drawing=>c_media_type_xml ).
    lo_data->add_drawing( lo_drawing ).

    " ---- sheet "Lists": source of a cross-sheet list validation
    lo_lists = lo_excel->add_new_worksheet( ).
    lo_lists->set_title( 'Lists' ) ##NO_TEXT.
    lo_lists->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Hà Nội` ) ##NO_TEXT.
    lo_lists->set_cell( ip_column = 'A' ip_row = 2 ip_value = `Bắc Ninh` ) ##NO_TEXT.
    lo_lists->set_cell( ip_column = 'A' ip_row = 3 ip_value = `TP. Hồ Chí Minh` ) ##NO_TEXT.

    " ---- sheet "Validation"
    lo_valid = lo_excel->add_new_worksheet( ).
    lo_valid->set_title( 'Validation' ) ##NO_TEXT.
    lo_valid->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Trạng thái (list)` ) ##NO_TEXT.
    lo_valid->set_cell( ip_column = 'B' ip_row = 1 ip_value = `Số lượng 1-100` ) ##NO_TEXT.
    lo_valid->set_cell( ip_column = 'C' ip_row = 1 ip_value = `Ngày >= 01/01/2026` ) ##NO_TEXT.
    lo_valid->set_cell( ip_column = 'D' ip_row = 1 ip_value = `Khu vực (sheet Lists)` ) ##NO_TEXT.
    lo_valid->set_cell( ip_column = 'E' ip_row = 1 ip_value = `Mã <= 10 ký tự` ) ##NO_TEXT.
    lo_valid->set_cell( ip_column = 'F' ip_row = 1 ip_value = `Đơn giá > 0 (custom)` ) ##NO_TEXT.

    " A: fixed list with dropdown, stop on error
    lo_dv = lo_valid->add_new_data_validation( ).
    lo_dv->type = zcl_excel_data_validation=>c_type_list.
    lo_dv->formula1 = '"Mới,Đang xử lý,Hoàn tất"' ##NO_TEXT.
    lo_dv->cell_column = 'A'.
    lo_dv->cell_row = 2.
    lo_dv->cell_column_to = 'A'.
    lo_dv->cell_row_to = 20.
    lo_dv->errorstyle = zcl_excel_data_validation=>c_style_stop.
    lo_dv->errortitle = `Giá trị không hợp lệ` ##NO_TEXT.
    lo_dv->error = `Chọn một trạng thái trong danh sách` ##NO_TEXT.
    lo_dv->prompttitle = `Trạng thái` ##NO_TEXT.
    lo_dv->prompt = `Chọn Mới / Đang xử lý / Hoàn tất` ##NO_TEXT.

    " B: whole number between 1 and 100
    lo_dv = lo_valid->add_new_data_validation( ).
    lo_dv->type = zcl_excel_data_validation=>c_type_whole.
    lo_dv->operator = zcl_excel_data_validation=>c_operator_between.
    lo_dv->formula1 = '1'.
    lo_dv->formula2 = '100'.
    lo_dv->cell_column = 'B'.
    lo_dv->cell_row = 2.
    lo_dv->cell_column_to = 'B'.
    lo_dv->cell_row_to = 20.
    lo_dv->errorstyle = zcl_excel_data_validation=>c_style_stop.
    lo_dv->error = `Nhập số nguyên từ 1 đến 100` ##NO_TEXT.

    " C: date on or after 01.01.2026 (warning only)
    lo_dv = lo_valid->add_new_data_validation( ).
    lo_dv->type = zcl_excel_data_validation=>c_type_date.
    lo_dv->operator = zcl_excel_data_validation=>c_operator_greaterthanorequal.
    lo_dv->formula1 = zcl_excel_common=>date_to_excel_string( CONV d( '20260101' ) ).
    lo_dv->cell_column = 'C'.
    lo_dv->cell_row = 2.
    lo_dv->cell_column_to = 'C'.
    lo_dv->cell_row_to = 20.
    lo_dv->errorstyle = zcl_excel_data_validation=>c_style_warning.
    lo_dv->error = `Ngày nên từ 01/01/2026` ##NO_TEXT.

    " D: list from another sheet
    lo_dv = lo_valid->add_new_data_validation( ).
    lo_dv->type = zcl_excel_data_validation=>c_type_list.
    lo_dv->formula1 = 'Lists!$A$1:$A$3' ##NO_TEXT.
    lo_dv->cell_column = 'D'.
    lo_dv->cell_row = 2.
    lo_dv->cell_column_to = 'D'.
    lo_dv->cell_row_to = 20.

    " E: text length <= 10
    lo_dv = lo_valid->add_new_data_validation( ).
    lo_dv->type = zcl_excel_data_validation=>c_type_textlength.
    lo_dv->operator = zcl_excel_data_validation=>c_operator_lessthanorequal.
    lo_dv->formula1 = '10'.
    lo_dv->cell_column = 'E'.
    lo_dv->cell_row = 2.
    lo_dv->cell_column_to = 'E'.
    lo_dv->cell_row_to = 20.
    lo_dv->errorstyle = zcl_excel_data_validation=>c_style_information.
    lo_dv->error = `Mã tối đa 10 ký tự` ##NO_TEXT.

    " F: custom formula
    lo_dv = lo_valid->add_new_data_validation( ).
    lo_dv->type = zcl_excel_data_validation=>c_type_custom.
    lo_dv->formula1 = 'AND(ISNUMBER(F2),F2>0)' ##NO_TEXT.
    lo_dv->cell_column = 'F'.
    lo_dv->cell_row = 2.
    lo_dv->cell_column_to = 'F'.
    lo_dv->cell_row_to = 20.
    lo_dv->error = `Đơn giá phải là số dương` ##NO_TEXT.

    lo_excel->set_active_sheet_index( 1 ).
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD build_condformat_link_xlsx.
    " Test workbook: conditional formatting (sheet "CondFormat") and hyperlinks (sheet "Links")
    TYPES: BEGIN OF ty_s_line,
             item   TYPE string,
             score  TYPE i,
             amount TYPE i,
             temp   TYPE i,
             growth TYPE i,
             status TYPE string,
           END OF ty_s_line.
    DATA: lo_excel  TYPE REF TO zcl_excel,
          lo_cf     TYPE REF TO zcl_excel_worksheet,
          lo_links  TYPE REF TO zcl_excel_worksheet,
          lo_cond   TYPE REF TO zcl_excel_style_cond,
          lo_writer TYPE REF TO zif_excel_writer,
          lt_lines  TYPE STANDARD TABLE OF ty_s_line WITH EMPTY KEY,
          lv_row    TYPE i.

    CREATE OBJECT lo_excel.

    " ---- styles used by the rules (written as differential formats)
    DATA(lo_st_red) = lo_excel->add_new_style( ).
    lo_st_red->font->bold = abap_true.
    lo_st_red->font->color-rgb = 'FF9C0006' ##NO_TEXT.
    lo_st_red->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_st_red->fill->bgcolor-rgb = 'FFFFC7CE' ##NO_TEXT.
    DATA(lo_st_yellow) = lo_excel->add_new_style( ).
    lo_st_yellow->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_st_yellow->fill->bgcolor-rgb = 'FFFFEB9C' ##NO_TEXT.
    DATA(lo_st_gray) = lo_excel->add_new_style( ).
    lo_st_gray->font->italic = abap_true.
    lo_st_gray->font->color-rgb = 'FF808080' ##NO_TEXT.
    DATA(lo_st_green) = lo_excel->add_new_style( ).
    lo_st_green->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_st_green->fill->bgcolor-rgb = 'FFC6EFCE' ##NO_TEXT.
    DATA(lo_st_link) = lo_excel->add_new_style( ).
    lo_st_link->font->color-rgb = 'FF0563C1' ##NO_TEXT.
    lo_st_link->font->underline = abap_true.
    lo_st_link->font->underline_mode = zcl_excel_style_font=>c_underline_single.

    " ---- sheet "CondFormat": data
    lo_cf = lo_excel->get_active_worksheet( ).
    lo_cf->set_title( 'CondFormat' ) ##NO_TEXT.
    lt_lines = VALUE #(
      ( item = 'SP01' score = 95 amount = 1200 temp = 12 growth = -15 status = 'OK' )
      ( item = 'SP02' score = 72 amount = 800  temp = 25 growth = 5   status = 'Lỗi giao hàng' )
      ( item = 'SP03' score = 25 amount = 300  temp = 31 growth = 40  status = 'OK' )
      ( item = 'SP04' score = 55 amount = 1500 temp = 18 growth = 12  status = 'Chờ duyệt' )
      ( item = 'SP05' score = 88 amount = 950  temp = 35 growth = -5  status = 'Lỗi hóa đơn' )
      ( item = 'SP06' score = 10 amount = 150  temp = 22 growth = 60  status = 'OK' )
      ( item = 'SP07' score = 64 amount = 700  temp = 28 growth = 25  status = 'OK' )
      ( item = 'SP08' score = 81 amount = 2000 temp = 15 growth = 0   status = 'Chờ duyệt' ) ) ##NO_TEXT.
    lo_cf->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Mã hàng` ) ##NO_TEXT.
    lo_cf->set_cell( ip_column = 'B' ip_row = 1 ip_value = `Điểm` ) ##NO_TEXT.
    lo_cf->set_cell( ip_column = 'C' ip_row = 1 ip_value = `Doanh số` ) ##NO_TEXT.
    lo_cf->set_cell( ip_column = 'D' ip_row = 1 ip_value = `Nhiệt độ` ) ##NO_TEXT.
    lo_cf->set_cell( ip_column = 'E' ip_row = 1 ip_value = `Tăng trưởng %` ) ##NO_TEXT.
    lo_cf->set_cell( ip_column = 'F' ip_row = 1 ip_value = `Trạng thái` ) ##NO_TEXT.
    lv_row = 2.
    LOOP AT lt_lines INTO DATA(ls_line).
      lo_cf->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_line-item ).
      lo_cf->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_line-score ).
      lo_cf->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_line-amount ).
      lo_cf->set_cell( ip_column = 'D' ip_row = lv_row ip_value = ls_line-temp ).
      lo_cf->set_cell( ip_column = 'E' ip_row = lv_row ip_value = ls_line-growth ).
      lo_cf->set_cell( ip_column = 'F' ip_row = lv_row ip_value = ls_line-status ).
      lv_row = lv_row + 1.
    ENDLOOP.

    " 1) cellIs > 80 -> red
    lo_cond = lo_cf->add_new_style_cond( 'B2:B9' ) ##NO_TEXT.
    lo_cond->rule = zcl_excel_style_cond=>c_rule_cellis.
    lo_cond->priority = 1.
    lo_cond->mode_cellis = VALUE #( operator = zcl_excel_style_cond=>c_operator_greaterthan
                                    formula  = '80'
                                    cell_style = lo_st_red->get_guid( ) ).
    " 2) cellIs between 50 and 80 -> yellow
    lo_cond = lo_cf->add_new_style_cond( 'B2:B9' ) ##NO_TEXT.
    lo_cond->rule = zcl_excel_style_cond=>c_rule_cellis.
    lo_cond->priority = 2.
    lo_cond->mode_cellis = VALUE #( operator = zcl_excel_style_cond=>c_operator_between
                                    formula  = '50'
                                    formula2 = '80'
                                    cell_style = lo_st_yellow->get_guid( ) ).
    " 3) expression: item gray when score < 30
    lo_cond = lo_cf->add_new_style_cond( 'A2:A9' ) ##NO_TEXT.
    lo_cond->rule = zcl_excel_style_cond=>c_rule_expression.
    lo_cond->priority = 3.
    lo_cond->mode_expression = VALUE #( formula = '$B2<30' cell_style = lo_st_gray->get_guid( ) ) ##NO_TEXT.
    " 4) data bar on amount
    lo_cond = lo_cf->add_new_style_cond( 'C2:C9' ) ##NO_TEXT.
    lo_cond->rule = zcl_excel_style_cond=>c_rule_databar.
    lo_cond->priority = 4.
    lo_cond->mode_databar = VALUE #( cfvo1_type = zcl_excel_style_cond=>c_cfvo_type_min
                                     cfvo2_type = zcl_excel_style_cond=>c_cfvo_type_max
                                     colorrgb   = 'FF638EC6' ) ##NO_TEXT.
    " 5) 3 color scale on temperature
    lo_cond = lo_cf->add_new_style_cond( 'D2:D9' ) ##NO_TEXT.
    lo_cond->rule = zcl_excel_style_cond=>c_rule_colorscale.
    lo_cond->priority = 5.
    lo_cond->mode_colorscale = VALUE #( cfvo1_type = zcl_excel_style_cond=>c_cfvo_type_min
                                        cfvo2_type = zcl_excel_style_cond=>c_cfvo_type_percentile cfvo2_value = '50'
                                        cfvo3_type = zcl_excel_style_cond=>c_cfvo_type_max
                                        colorrgb1  = 'FF63BE7B' colorrgb2 = 'FFFFEB84' colorrgb3 = 'FFF8696B' ) ##NO_TEXT.
    " 6) icon set (3 arrows) on growth
    lo_cond = lo_cf->add_new_style_cond( 'E2:E9' ) ##NO_TEXT.
    lo_cond->rule = zcl_excel_style_cond=>c_rule_iconset.
    lo_cond->priority = 6.
    lo_cond->mode_iconset = VALUE #( iconset    = zcl_excel_style_cond=>c_iconset_3arrows
                                     cfvo1_type = zcl_excel_style_cond=>c_cfvo_type_percent cfvo1_value = '0'
                                     cfvo2_type = zcl_excel_style_cond=>c_cfvo_type_percent cfvo2_value = '33'
                                     cfvo3_type = zcl_excel_style_cond=>c_cfvo_type_percent cfvo3_value = '67'
                                     showvalue  = zcl_excel_style_cond=>c_showvalue_true ).
    " 7) top 3 amounts -> green
    lo_cond = lo_cf->add_new_style_cond( 'C2:C9' ) ##NO_TEXT.
    lo_cond->rule = zcl_excel_style_cond=>c_rule_top10.
    lo_cond->priority = 7.
    lo_cond->mode_top10 = VALUE #( topxx_count = 3 cell_style = lo_st_green->get_guid( ) ).
    " 8) above average temperature -> red
    lo_cond = lo_cf->add_new_style_cond( 'D2:D9' ) ##NO_TEXT.
    lo_cond->rule = zcl_excel_style_cond=>c_rule_above_average.
    lo_cond->priority = 8.
    lo_cond->mode_above_average = VALUE #( above_average = abap_true cell_style = lo_st_red->get_guid( ) ).
    " 9) text contains "Lỗi" -> red
    lo_cond = lo_cf->add_new_style_cond( 'F2:F9' ) ##NO_TEXT.
    lo_cond->rule = zcl_excel_style_cond=>c_rule_textfunction.
    lo_cond->priority = 9.
    lo_cond->mode_textfunction = VALUE #( textfunction = zcl_excel_style_cond=>c_textfunction_containstext
                                          text         = `Lỗi`
                                          cell_style   = lo_st_red->get_guid( ) ) ##NO_TEXT.

    " ---- sheet "Links": hyperlinks
    lo_links = lo_excel->add_new_worksheet( ).
    lo_links->set_title( 'Links' ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Loại liên kết` ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'B' ip_row = 1 ip_value = `Liên kết` ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'A' ip_row = 2 ip_value = `Web` ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'B' ip_row = 2 ip_value = `Trang SAP` ip_style = lo_st_link->get_guid( )
                        ip_hyperlink = zcl_excel_hyperlink=>create_external_link( `https://www.sap.com` ) ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'A' ip_row = 3 ip_value = `URL có tham số` ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'B' ip_row = 3 ip_value = `Tìm kiếm abap2xlsx` ip_style = lo_st_link->get_guid( )
                        ip_hyperlink = zcl_excel_hyperlink=>create_external_link(
                          |https://github.com/search?q=abap2xlsx{ zcl_excel_common=>unicode_char_from_code_point( 38 ) }type=repositories| ) ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'A' ip_row = 4 ip_value = `Email` ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'B' ip_row = 4 ip_value = `Gửi email` ip_style = lo_st_link->get_guid( )
                        ip_hyperlink = zcl_excel_hyperlink=>create_external_link( `mailto:coqa.hm@hm.com?subject=abap2xlsx` ) ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'A' ip_row = 5 ip_value = `Trong file (sheet khác)` ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'B' ip_row = 5 ip_value = `Tới CondFormat!F2` ip_style = lo_st_link->get_guid( )
                        ip_hyperlink = zcl_excel_hyperlink=>create_internal_link( `CondFormat!F2` ) ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'A' ip_row = 6 ip_value = `Trong file (cùng sheet)` ) ##NO_TEXT.
    lo_links->set_cell( ip_column = 'B' ip_row = 6 ip_value = `Tới A1` ip_style = lo_st_link->get_guid( )
                        ip_hyperlink = zcl_excel_hyperlink=>create_internal_link( `Links!A1` ) ) ##NO_TEXT.

    lo_excel->set_active_sheet_index( 1 ).
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD build_protect_pagesetup_xlsx.
    " Test workbook: sheet protection + workbook protection (sheet "Protected"), page setup (sheet "PageSetup")
    DATA: lo_excel    TYPE REF TO zcl_excel,
          lo_prot     TYPE REF TO zcl_excel_worksheet,
          lo_page     TYPE REF TO zcl_excel_worksheet,
          lo_writer   TYPE REF TO zif_excel_writer,
          ls_header   TYPE zexcel_s_worksheet_head_foot,
          ls_footer   TYPE zexcel_s_worksheet_head_foot,
          lv_row      TYPE i.

    CREATE OBJECT lo_excel.

    " ---- styles: unlocked input cells, hidden formula cells, header
    DATA(lo_st_input) = lo_excel->add_new_style( ).
    lo_st_input->protection->locked = zcl_excel_style_protection=>c_protection_unlocked.
    lo_st_input->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_st_input->fill->fgcolor-rgb = 'FFFFF2CC' ##NO_TEXT.
    DATA(lo_st_hidden) = lo_excel->add_new_style( ).
    lo_st_hidden->protection->locked = zcl_excel_style_protection=>c_protection_locked.
    lo_st_hidden->protection->hidden = zcl_excel_style_protection=>c_protection_hidden.
    DATA(lo_st_head) = lo_excel->add_new_style( ).
    lo_st_head->font->bold = abap_true.
    lo_st_head->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_st_head->fill->fgcolor-rgb = 'FFD9E1F2' ##NO_TEXT.

    " ---- sheet "Protected"
    lo_prot = lo_excel->get_active_worksheet( ).
    lo_prot->set_title( 'Protected' ) ##NO_TEXT.
    lo_prot->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Mã (khóa)` ip_style = lo_st_head->get_guid( ) ) ##NO_TEXT.
    lo_prot->set_cell( ip_column = 'B' ip_row = 1 ip_value = `Số lượng (được nhập)` ip_style = lo_st_head->get_guid( ) ) ##NO_TEXT.
    lo_prot->set_cell( ip_column = 'C' ip_row = 1 ip_value = `Thành tiền (công thức ẩn)` ip_style = lo_st_head->get_guid( ) ) ##NO_TEXT.
    DO 4 TIMES.
      lv_row = sy-index + 1.
      lo_prot->set_cell( ip_column = 'A' ip_row = lv_row ip_value = |SP{ sy-index WIDTH = 2 ALIGN = RIGHT PAD = '0' }| ) ##NO_TEXT.
      lo_prot->set_cell( ip_column = 'B' ip_row = lv_row ip_value = sy-index * 10 ip_style = lo_st_input->get_guid( ) ).
      lo_prot->set_cell( ip_column = 'C' ip_row = lv_row ip_formula = |B{ lv_row }*1500| ip_style = lo_st_hidden->get_guid( ) ) ##NO_TEXT.
    ENDDO.
    lo_prot->get_column( 'A' )->set_width( 14 ).
    lo_prot->get_column( 'B' )->set_width( 22 ).
    lo_prot->get_column( 'C' )->set_width( 26 ).

    lo_prot->zif_excel_sheet_protection~protected = zif_excel_sheet_protection=>c_protected.
    lo_prot->zif_excel_sheet_protection~password  = zcl_excel_common=>encrypt_password( iv_sheet_password ).
    lo_prot->zif_excel_sheet_protection~sheet     = zif_excel_sheet_protection=>c_active.
    lo_prot->zif_excel_sheet_protection~objects   = zif_excel_sheet_protection=>c_active.
    lo_prot->zif_excel_sheet_protection~scenarios = zif_excel_sheet_protection=>c_active.
    " users may still format columns (0 = allowed in OOXML)
    lo_prot->zif_excel_sheet_protection~format_columns = zif_excel_sheet_protection=>c_noactive.

    " ---- workbook protection (structure: no add / delete / rename of sheets)
    lo_excel->zif_excel_book_protection~protected        = zif_excel_book_protection=>c_protected.
    lo_excel->zif_excel_book_protection~lockstructure    = zif_excel_book_protection=>c_locked.
    lo_excel->zif_excel_book_protection~workbookpassword = zcl_excel_common=>encrypt_password( iv_book_password ).

    " ---- sheet "PageSetup": 60 lines -> 2 printed pages
    lo_page = lo_excel->add_new_worksheet( ).
    lo_page->set_title( 'PageSetup' ) ##NO_TEXT.
    lo_page->set_cell( ip_column = 'A' ip_row = 1 ip_value = `STT` ip_style = lo_st_head->get_guid( ) ) ##NO_TEXT.
    lo_page->set_cell( ip_column = 'B' ip_row = 1 ip_value = `Diễn giải` ip_style = lo_st_head->get_guid( ) ) ##NO_TEXT.
    lo_page->set_cell( ip_column = 'C' ip_row = 1 ip_value = `Số tiền` ip_style = lo_st_head->get_guid( ) ) ##NO_TEXT.
    DO 60 TIMES.
      lv_row = sy-index + 1.
      lo_page->set_cell( ip_column = 'A' ip_row = lv_row ip_value = sy-index ).
      lo_page->set_cell( ip_column = 'B' ip_row = lv_row ip_value = |Dòng chi tiết số { sy-index }| ) ##NO_TEXT.
      lo_page->set_cell( ip_column = 'C' ip_row = lv_row ip_value = sy-index * 125000 ).
    ENDDO.
    lo_page->get_column( 'B' )->set_width( 40 ).
    lo_page->get_column( 'C' )->set_width( 18 ).

    lo_page->sheet_setup->orientation = zcl_excel_sheet_setup=>c_orientation_landscape.
    lo_page->sheet_setup->paper_size  = zcl_excel_sheet_setup=>c_papersize_a4.
    lo_page->sheet_setup->fit_to_page = abap_true.
    lo_page->sheet_setup->fit_to_width  = 1.
    lo_page->sheet_setup->fit_to_height = 0.
    lo_page->sheet_setup->horizontal_centered = abap_true.
    lo_page->sheet_setup->set_page_margins( ip_top = CONV f( '2.0' ) ip_bottom = CONV f( '2.0' ) ip_left = CONV f( '1.5' ) ip_right = CONV f( '1.5' )
                                            ip_header = CONV f( '0.8' ) ip_footer = CONV f( '0.8' ) ip_unit = 'cm' ) ##NO_TEXT.
    ls_header-left_value   = `&A` ##NO_TEXT.                           " sheet name
    ls_header-center_value = `BÁO CÁO CHI TIẾT` ##NO_TEXT.
    ls_header-center_font-bold = abap_true.
    ls_header-center_font-size = 14.
    ls_header-right_value  = `In ngày &D` ##NO_TEXT.                   " print date
    ls_footer-center_value = `Trang &P / &N` ##NO_TEXT.                " page x of y
    lo_page->sheet_setup->set_header_footer( ip_odd_header = ls_header ip_odd_footer = ls_footer ).
    lo_page->zif_excel_sheet_printsettings~set_print_repeat_rows( iv_rows_from = 1 iv_rows_to = 1 ).
    lo_page->set_print_gridlines( abap_true ).
    lo_page->get_pagebreaks( )->add_pagebreak( ip_column = 'A' ip_row = 31 ).

    lo_excel->set_active_sheet_index( 1 ).
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD get_test_vba_project.
    " vbaProject.bin of the XlsxWriter examples (github.com/jmcnamara/XlsxWriter, BSD license):
    " module Module1 with Sub say_hello (MsgBox), code names ThisWorkbook / Sheet1
    DATA lv_b64 TYPE string.
    lv_b64 = `0M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAPgADAP7/CQAGAAAAAAAAAAAAAAABAAAAAQAAAAAAAAAAEAAAAgAAAAIAAAD+////AAAAAAAAAAD/////////////////////////////////////////////////////////////////////////////////////////////`
          && `////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `///////////////////////////////////////////////////////////////////////////////////9////BwAAABQAAAAEAAAABQAAAAYAAAAIAAAADQAAAAkAAAAKAAAACwAAAAwAAAAOAAAAGwAAAA8AAAAQAAAAEQAAABIAAAATAAAAFQAAAP7///8WAAAA`
          && `FwAAABgAAAAZAAAAGgAAABwAAAD+////HQAAAP7/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////1IAbwBvAHQAIABFAG4AdAByAHkAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWAAUA//////////8OAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALCnnEHE988BAwAAAMAsAAAAAAAAVgBCAEEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAgAAQD//////////wQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALCnnEHE988BsKecQcT3zwEAAAAAAAAAAAAAAABUAGgAaQBzAFcAbwByAGsAYgBvAG8AawAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGgACAP//`
          && `/////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADhAwAAAAAAAFMAaABlAGUAdAAxAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOAAIBCgAAAAgAAAD/////AAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAANoDAAAAAAAAAQAAAAIAAAADAAAABAAAAAUAAAAGAAAABwAAAAgAAAAJAAAACgAAAAsAAAAMAAAADQAAAA4AAAAPAAAA/v///xEAAAASAAAAEwAAABQAAAAVAAAAFgAAABcAAAAYAAAAGQAAABoAAAAbAAAAHAAAAB0A`
          && `AAAeAAAAHwAAAP7///8hAAAAIgAAACMAAAAkAAAAJQAAACYAAAAnAAAAKAAAACkAAAAqAAAAKwAAACwAAAAtAAAALgAAAC8AAAAwAAAA/v///zIAAAAzAAAANAAAAP7///82AAAA/v///zgAAAA5AAAAOgAAADsAAAA8AAAAPQAAAD4AAAA/AAAAQAAAAEEAAABCAAAA`
          && `QwAAAEQAAABFAAAARgAAAP7///9IAAAASQAAAEoAAABLAAAATAAAAE0AAABOAAAATwAAAFAAAABRAAAAUgAAAFMAAABUAAAAVQAAAFYAAAD+////WAAAAFkAAABaAAAAWwAAAFwAAABdAAAAXgAAAF8AAABgAAAAYQAAAGIAAABjAAAAZAAAAGUAAABmAAAAZwAAAGgA`
          && `AABpAAAAagAAAGsAAABsAAAAbQAAAG4AAABvAAAAcAAAAHEAAAByAAAAcwAAAHQAAAB1AAAAdgAAAHcAAAB4AAAAeQAAAHoAAAB7AAAAfAAAAH0AAAB+AAAAfwAAAIAAAAABFgEAAfAAAADMAgAA1AAAAAACAAD/////0wIAACcDAAAAAAAAAQAAANqdJiAAAP//IwEA`
          && `AIgAAAC2AP//AQEAAAAA/////wAAAAD///////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAwAAAAUAAAAHAAAA//////////8BAQgAAAD/////eAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAP//AAAAAE1FAAD///////8AAAAA//8AAAAA//8BAQAAAADfAP//AAAAAAwA////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `//////////////////////////////////8oAAAAAgBTTP////8AAAEAUxD/////AAABAFOU/////wAAAAACPP////8AAP//AQEAAAAAAQBOADAAewAwADAAMAAyADAAOAAxADkALQAwADAAMAAwAC0AMAAwADAAMAAtAEMAMAAwADAALQAwADAAMAAwADAAMAAwADAA`
          && `MAAwADQANgB9AAYAAAD/////AQFAAAAAAoD+//////8QAP//KAAAAAIB//8AAAAAAAAAAP//////////AAAAAB0AAAAlAAAA/////0gAAAD/////QAAAAAAAAAAAAAEAAAAAAAAAAAD///////////////8AAAAA//////////////////////////8AAAAA////////`
          && `//////////////////8AAAAAAQAAAP//////////AAAAAP///////////////////////////////wEAMAAAAFEfZ1IgAN8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+ygEAAAD/////AQEIAAAA////`
          && `/3gAAAD/////AAABsLAAQXR0cmlidXQAZSBWQl9OYW0AZSA9ICJUaGkAc1dvcmtib28QayINCgqMQmFzAQKMMHswMDAyMFA4MTktABAwAwhDBwAUAhIBJDAwNDZ9gQ18R2xvYmFsAdAQU3BhYwGSRmFsBHNlDGRDcmVhdAhhYmwVH1ByZWSQZWNsYQAGSWQAsQhUcnUN`
          && `QkV4cG8Ec2UUHFRlbXBsAGF0ZURlcml2AwISkkJ1c3RvbWkGegREAzIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARYBAAHwAAAAzAIAANQAAAAAAgAA/////9MCAAAnAwAAAAAAAAEAAADanYU6AAD//yMBAACIAAAAtgD//wEBAAAAAP////8AAAAA////`
          && `////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAMAAAAFAAAABwAAAP//////////AQEIAAAA/////3gAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//wAAAABNRQAA////////AAAAAP//`
          && `AAAAAP//AQEAAAAA3wD//wAAAAAMAP//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `KAAAAAIAU0z/////AAABAFMQ/////wAAAQBTlP////8AAAAAAjz/////AAD//wEBAAAAAAEATgAwAHsAMAAwADAAMgAwADgAMQA5AC0AMAAwADAAMAAtADAAMAAwADAALQBDADAAMAAwAC0AMAAwADAAMAAwADAAMAAwADAAMAA0ADYAfQAGAAAA/////wEBQAAAAAKA`
          && `/v//////EAD//ygAAAACAf//AAAAAAAAAAD//////////wAAAAAdAAAAJQAAAP////9IAAAA/////0AAAAAAAAAAAAABAAAAAAAAAAAA////////////////AAAAAP//////////////////////////AAAAAP//////////////////////////AAAAAAEAAAD/////`
          && `/////wAAAAD///////////////////////////////8BADAAAABRH2dSIADfAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/soBAAAA/////wEBCAAAAP////94AAAA/////wAAAamwAEF0dHJpYnV0AGUg`
          && `VkJfTmFtAGUgPSAiU2hlQGV0MSINCgroQgRhcwJ0MHswMDBAMjA4MTktACAwHQMIQwAUAhIBJDAwNAQ2fQ18R2xvYmFCbAHEU3BhYwGSRhBhbHNlDGRDcmUgYXRhYmwVH1ByQGVkZWNsYQAGSSJkAKtUcnUNQkV4EHBvc2UUHFRlbQBwbGF0ZURlcgxpdgIkkkJ1c3Rv`
          && `GG1pegREAzIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAE0AbwBkAHUAbABlADEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAIBAwAAAAYAAAD/////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAIAAAACAEAAAAAAAAXwBfAFMAUgBQAF8AMgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAgD///////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAxAAAA4AAAAAAA`
          && `AABfAF8AUwBSAFAAXwAzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAACAAwAAAAJAAAA/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADUAAABnAAAAAAAAAFQAaABpAHMAVwBvAHIAawBiAG8A`
          && `bwBrADEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAcAAIA////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANwAAANoDAAAAAAAAARYBAATwAAAAxAIAANQAAACwAQAA/////wUDAAC1AwAAAAAAAAEAAADa`
          && `ndDbAAD//wMAAAAAAAAAtgD//wEBAAAAAP////8AAAAA//8EAP//AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAMAAAAFAAAABwAAAP//////////AQEIAAAA/////3gAAAACAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAD//wAAAABNRQAA////////AAAAAP//AAAAAP//AQEAAAAA3wD//wAAAAAIAP//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `////////////////////////////////////////////////KAAAAAAANgr/////AAAAABoI/////wAAAAACPAwA//8AAAAAAjz/////AAD//wEBAAAAAAAAAQAAAP////8BAYgAAAA4AAAA/////wKDHgL/////CAD//zAAAAAAAP///////wAAAAD//////////wAA`
          && `AAAdAAAAJQAAAAsSIAL/////AAAAYAAAAAD//////////wAAAAAAAAAA//////z+SAD/////////////AAADAAMAAACEAAADAAD/////wAEAAP////+4AQAAAQABAAAAAAAAAAAAAAAAADgAAAD//////////////////////////zgAAAD/////////////////////`
          && `////////////////CAAAAAAAAAAAAAAACAAEAP////8AAAAA////////////////////////////////AgAAAAAAUR9nUiAAASQAKgBcAFIAZgBmAGYAZgAqADAAYgA1ADIANgA3ADEAZgA1ADEAARAAKgBcAFIAMAAqACMAMQA0AN8AAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+ygEAAwAigQgABgACAAAAAAAAgQgEHgAAABgAAAAEgQgAAgADAAgAAAD/////AQFAAAAAlgQ4AAAAAABvAP//cAAAAP////9IAAAAtgASAEhlbGxvIGZyb20gUHl0aG9uIR0AQUAiAgEAYwD/`
          && `////MAAAAP////8AAAFhsABBdHRyaWJ1dABlIFZCX05hbQBlID0gIk1vZAB1bGUxIg0KUwB1YiBzYXlfaABlbGxvKCkNCgIgAABNc2dCb3gQICgiSAFUIGZyAG9tIFB5dGhvCG4hIgBERW5kIAEAbg0KAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAByVYAA`
          && `AACAAAAAgAAAAIAAAAACAAB+fAAAfwAAAAAOAAAACQAAAAAAAAAJAAAAAAADAAgAAAAAAAIAAQABAAMAAADBBQAAAAAAADEGAAAAAAAAoQYAAAAAAABwAAAAMAAAAAACACsn/P4nHP8nPP/1AAAAADps/wIATlz/BFz/CgEAFAA2CABc/zz/HP/8/gAAFAAAAAAEAIAA`
          && `MAAkAAgAAAAAABMAAAAAAAwAAAAAAAAAAAAAABwAAAAAAAQAAAAAAFz/AgA8/wIAHP8CAPz+AgAEAAASAABrAAB/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAclWAAAAAAAAAAIAAAACAAAAAAAAAABAAAAAJAAAAAAACAP//////////AAAAAAgA`
          && `AAAEACQAgQAAAAAAAgAAAABgAAD9//////////////8AAAAAAAAAAAAAAAAAAABuAAB/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABFgEAAPAAAADEAgAA1AAAAAACAAD/////ywIAAB8DAAAAAAAAAQAAANqdvlAAAP//IwEAAIgAAAC2AP//AQEAAAAA////`
          && `/wAAAAD///////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAwAAAAUAAAAHAAAA//////////8BAQgAAAD/////eAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//AAAAAE1FAAD/////`
          && `//8AAAAA//8AAAAA//8BAQAAAADfAP//AAAAAAwA////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `//////////8oAAAAAgBTTP////8AAAEAUxD/////AAABAFOU/////wAAAAACPP////8AAP//AQEAAAAAAQBOADAAewAwADAAMAAyADAAOAAxADkALQAwADAAMAAwAC0AMAAwADAAMAAtAEMAMAAwADAALQAwADAAMAAwADAAMAAwADAAMAAwADQANgB9AAAAAAD/////`
          && `AQE4AAAAAoD+//////8QAP//KAAAAAIB//8AAAAAAAAAAP//////////AAAAAB0AAAAlAAAA/////0gAAAAAAP//AAABAAAAAAAAAAAA////////////////AAAAAP//////////////////////////AAAAAP//////////////////////////AAAAAAAAAAD/////`
          && `/////wAAAAD///////////////////////////////////////8AAAAAAADfAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/soBAAAA/////wEBCAAAAP////94AAAA/////wAAAbGwAEF0dHJpYnV0AGUg`
          && `VkJfTmFtAGUgPSAiVGhpAHNXb3JrYm9vIGsxIg0KCpBCYQJzApAwezAwMDKgMDgxOS0AEDADCA5DABQCEgEkMDA0NgJ9DXxHbG9iYWwhAdJTcGFjAZJGYQhsc2UMZENyZWEQdGFibBUfUHJlIGRlY2xhAAZJZBEAslRydQ1CRXhwCG9zZRQcVGVtcABsYXRlRGVyaQZ2`
          && `AhKSQnVzdG9tDGl6BEQDMgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARYBAADwAAAAxAIAANQAAAAAAgAA/////8sCAAAfAwAAAAAAAAEAAADanTJ/AAD//yMBAACIAAAAtgD//wEBAFMAaABlAGUAdAAyAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOAAIA////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARwAAANEDAAAAAAAAXwBWAEIAQQBfAFAAUgBPAEoARQBDAFQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAABoAAgECAAAABwAAAP////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABXAAAASAsAAAAAAABkAGkAcgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAACAP//////////`
          && `/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIUAAABqAgAAAAAAAF8AXwBTAFIAUABfADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAIA////////////////AAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAjwAAACMFAAAAAAAAAAAA/////wAAAAD///////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAwAAAAUAAAAHAAAA//////////8BAQgAAAD/////eAAAAAgAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//AAAAAE1FAAD///////8AAAAA//8AAAAA//8BAQAAAADfAP//AAAAAAwA////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `//////////////////////////////////////////////////////////8oAAAAAgBTTP////8AAAEAUxD/////AAABAFOU/////wAAAAACPP////8AAP//AQEAAAAAAQBOADAAewAwADAAMAAyADAAOAAyADAALQAwADAAMAAwAC0AMAAwADAAMAAtAEMAMAAwADAA`
          && `LQAwADAAMAAwADAAMAAwADAAMAAwADQANgB9AAAAAAD/////AQE4AAAAAoD+//////8QAP//KAAAAAIB//8AAAAAAAAAAP//////////AAAAAB0AAAAlAAAA/////0gAAAAAAP//AAABAAAAAAAAAAAA////////////////AAAAAP//////////////////////////`
          && `AAAAAP//////////////////////////AAAAAAEAAAD//////////wAAAAD///////////////////////////////////////8AAAAAAADfAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/soBAAAA////`
          && `/wEBCAAAAP////94AAAA/////wAAAaiwAEF0dHJpYnV0AGUgVkJfTmFtAGUgPSAiU2hlQGV0MiINCgroQgRhcwJ0MHswMDDAMjA4MjAtACAECA5DABQCHAEkMDA0NgJ9DXxHbG9iYWwhAcRTcGFjAZJGYQhsc2UMZENyZWEQdGFibBUfUHJlIGRlY2xhAAZJZBEAq1Ry`
          && `dQ1CRXhwCG9zZRQcVGVtcABsYXRlRGVyaQZ2AiSSQnVzdG9tDGl6BEQDMgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAzGGFAAABAP8JCAAACQQAAOQEAQAAAAAAAAAAAAEABAACABwBKgBcAEcAewAwADAAMAAyADAANABFAEYA`
          && `LQAwADAAMAAwAC0AMAAwADAAMAAtAEMAMAAwADAALQAwADAAMAAwADAAMAAwADAAMAAwADQANgB9ACMANAAuADAAIwA5ACMAQwA6AFwAUAByAG8AZwByAGEAbQAgAEYAaQBsAGUAcwBcAEMAbwBtAG0AbwBuACAARgBpAGwAZQBzAFwATQBpAGMAcgBvAHMAbwBmAHQA`
          && `IABTAGgAYQByAGUAZABcAFYAQgBBAFwAVgBCAEEANgBcAFYAQgBFADYALgBEAEwATAAjAFYAaQBzAHUAYQBsACAAQgBhAHMAaQBjACAARgBvAHIAIABBAHAAcABsAGkAYwBhAHQAaQBvAG4AcwAAAAAAAAAAAAAAAAAaASoAXABHAHsAMAAwADAAMgAwADgAMQAzAC0A`
          && `MAAwADAAMAAtADAAMAAwADAALQBDADAAMAAwAC0AMAAwADAAMAAwADAAMAAwADAAMAA0ADYAfQAjADEALgA2ACMAMAAjAEMAOgBcAHAAcgBvAGcAcgBhAG0AIABmAGkAbABlAHMAXABtAGkAYwByAG8AcwBvAGYAdAAgAG8AZgBmAGkAYwBlACAAMgAwADAANwBcAE8A`
          && `ZgBmAGkAYwBlADEAMgBcAEUAWABDAEUATAAuAEUAWABFACMATQBpAGMAcgBvAHMAbwBmAHQAIABFAHgAYwBlAGwAIAAxADIALgAwACAATwBiAGoAZQBjAHQAIABMAGkAYgByAGEAcgB5AAAAAAAAAAAAAAAAALwAKgBcAEcAewAwADAAMAAyADAANAAzADAALQAwADAA`
          && `MAAwAC0AMAAwADAAMAAtAEMAMAAwADAALQAwADAAMAAwADAAMAAwADAAMAAwADQANgB9ACMAMgAuADAAIwAwACMAQwA6AFwAVwBJAE4ARABPAFcAUwBcAHMAeQBzAHQAZQBtADMAMgBcAHMAdABkAG8AbABlADIALgB0AGwAYgAjAE8ATABFACAAQQB1AHQAbwBtAGEA`
          && `dABpAG8AbgAAAAAAAAAAAAAAAAAoASoAXABHAHsAMgBEAEYAOABEADAANABDAC0ANQBCAEYAQQAtADEAMAAxAEIALQBCAEQARQA1AC0AMAAwAEEAQQAwADAANAA0AEQARQA1ADIAfQAjADIALgA0ACMAMAAjAEMAOgBcAFAAcgBvAGcAcgBhAG0AIABGAGkAbABlAHMA`
          && `XABDAG8AbQBtAG8AbgAgAEYAaQBsAGUAcwBcAE0AaQBjAHIAbwBzAG8AZgB0ACAAUwBoAGEAcgBlAGQAXABPAEYARgBJAEMARQAxADIAXABNAFMATwAuAEQATABMACMATQBpAGMAcgBvAHMAbwBmAHQAIABPAGYAZgBpAGMAZQAgADEAMgAuADAAIABPAGIAagBlAGMA`
          && `dAAgAEwAaQBiAHIAYQByAHkAAAAAAAAAAAAAAAAABQACAAIAAQACAAIABAAEAgAABgIBAAgCAAAKAgEAEAL///////8AAAAA//8AAFEfZ1IgAAEA//8CAAAA//////////8EAAMA//////////////////////////////////////////////////////////8BAAAA`
          && `AAAAAAAAAAAAAAAAAAAAANqdBQAYAFQAaABpAHMAVwBvAHIAawBiAG8AbwBrABQAMABdADUAMgA2ADcAMQBmADIAMQD//xUCGABUAGgAaQBzAFcAbwByAGsAYgBvAG8AawD//yYgAAAAAAAAAAIAAAAtAwAA//8MAFMAaABlAGUAdAAxABQAMABeADUAMgA2ADcAMQBm`
          && `ADIAMQD//xkCDABTAGgAZQBlAHQAMQD//4U6AAAAAAAAGAIAAAAtAwAA//8OAE0AbwBkAHUAbABlADEAFAAwAGIANQAyADYANwAxAGYANQAxAP//HgIOAE0AbwBkAHUAbABlADEA///Q2wAAAAAAADACAAAAuwMAAP//GgBUAGgAaQBzAFcAbwByAGsAYgBvAG8AawAx`
          && `ABQAMAA1ADUANgAxADEAYwAwAGIAZQD//yUCGgBUAGgAaQBzAFcAbwByAGsAYgBvAG8AawAxAP//vlAAAAAAAABQAgAAACUDAAD//wwAUwBoAGUAZQB0ADIAFAAwADYANQA2ADEAMQBjADAAYgBlAP//JwIMAFMAaABlAGUAdAAyAP//Mn8AAAAAAABoAgAAACUDAAD/`
          && `//////8BAYgCAAAYAgAA/////////////////////2gCAAD/////////////////////////////////////////////////////////////////////////////////////////////////////MAIAAP///////////////////////////////wACAAD/////////`
          && `//////////////////////////////////////////////////////////////////////////////////////9QAgAA////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////gQAAAIIAAACDAAAAhAAAAP7///+GAAAAhwAAAIgAAACJAAAAigAAAIsAAACMAAAA`
          && `jQAAAI4AAAD+////kAAAAJEAAACSAAAAkwAAAJQAAACVAAAAlgAAAJcAAACYAAAAmQAAAJoAAACbAAAAnAAAAJ0AAACeAAAAnwAAAKAAAAChAAAAogAAAKMAAAD+////pQAAAP7///+nAAAAqAAAAP7///+qAAAAqwAAAKwAAACtAAAArgAAAK8AAACwAAAAsQAAALIA`
          && `AAD+////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////`
          && `////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////rNztGtDOSEyMTMXxwDvw`
          && `FP////8BAAAAGt2hIqKi/km0kB0c45MtUv////8BAAAAU4biDKx9xkiedsKr70dxO/////8BAAAA/////zAAAAD/A1s5f1x+RIZ2wPAlw/Xp/////wEAAAB/NCo06yLUQaTZXrMNyxhP/////wEAAAD/////GAIAAIAAAAAAABMBFAD/ALgmAAAFBEV4Y2VsgCsQAAME`
          && `VkJB9+IQAAUEV2luMTbBfhAABQRXaW4zMgd/EAADBE1hY7OyEAAEBFZCQTatIxAACARQcm9qZWN0MQoXEAAGBHN0ZG9sZZNgEAAKBFZCQVByb2plY3S+vxAABgRPZmZpY2UVdRAADARUaGlzV29ya2Jvb2t84xAACYAAAP8DAQBfRXZhbHVhdGUY2RAABgRTaGVldDHo`
          && `GhAACARXb3JrYm9va2sYEAAJBFdvcmtzaGVldMH+EAAHBE1vZHVsZTFiERAACYQIAP8D//9zYXlfaGVsbG8wdBAABgBNc2dCb3iXUhAADQRUaGlzV29ya2Jvb2sxNV0QAAYEU2hlZXQy6RoQAAL//wEBeAAAAP///////////////////////w4CAgD//xAC/////xIC`
          && `AwD//xUCAAAOAP///////xkCAQAGAP///////////////x8CAgALAP///////////////yUCAwALACcCBAAIAAACAQD//wICAAD//////////////////w8AFAAAAAEAWgAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAEAAAAAAAAA`
          && `AAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWaygAEABAAAAAEAMCoCApAJAHAUBkgDAIICAGTkBAQACgAcAFZCQVByb2pliGN0BQA0AABA`
          && `AhRqBgIKPQIKBwJyARQIBQYSCQISUR9nUiCUAAwCSjwCChYAAXKAc3Rkb2xlPgIZAHMAdABkAG8AgGwAZQANAGgAJQJeAAMqXEd7MDCAMDIwNDMwLQAIHQQEQwAKAg4BEjAwNAA2fSMyLjAjMAAjQzpcV0lORABPV1Ncc3lzdCBlbTMyXANlMi4AdGxiI09MRSAAQXV0`
          && `b21hdGkcb24AYAABg0VPZmZEaWOERU8AZoAAadQAY4JFnoARlIABgUUAMkRGOEQwNEMALTVCRkEtMTCAMUItQkRFNYBF1EFBgEM0gAUyiEWAmABncmFtIEZpbABlc1xDb21tbwJuBAZNaWNyb3MAb2Z0IFNoYXIAZWRcT0ZGSUMARTEyXE1TTy4wRExMI4cQg00gMRGA`
          && `cSBPYoHBIExpYGJyYXJ5AEsAAQ8RgtQFABOCA9qdGQGCqFRoaXNXb3IAa2Jvb2tHABgFgBNUgKtpAHMAV0GAs3IAawBiwAFvVcABGs4LMtoLHMASABQASEIBMUJ4LQMALAAeQgIBBSzCISYgKiJCCCtCARlCfFNooGVldDFHwhtTQCOCZUBYdAAxABpIB6YyTgfjG4U6`
          && `yxsHwB1ATW9kdWxlABwO4wEDgDtkAHWCmIEcCAgsMgAPCE8du1A50NvSIYAWAABFOQ1AA0levcAeGoEElF7BIY4MMsBDpAAAVWsxAE8mJVAm9L5Q1F8yzV/hDeFxojOaMu0vMv8N4j0yf+kNAhCAHgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJNLKoUBABAAAAD//wAA`
          && `AAABAAIA//8AAAAAAQAAAAIAAAAAAAEAAgACAAAAAAABAAUABQAFAAUABQAFAAUABQAFAAUABQAFAAAAclUAAQAAgAAAAIAAAACAAAAABAAAfgEAAH4BAAB+AQAAfgEAAH4CAAB+AwAAfmsAAH8AAAAAFQAAAAkAAAAAAAEACAAAAAAAAAARAQAAAAAAACcCvjBhkNFL`
          && `ukWsHOgxMYkBAAkEAAAJCAAA5AQAAAAAAAABAP////8FAAIKAAD///////////////8AAAAAuQAAAAAAAAACCgAA////////////////AAAAAOEAAAAAAAAAAQhBAAkAAAAAAAIAoQUAAAAAAAD/////cAAAAAAAAAD/////AQCBBQAAAAAAAP//AAD5AAAAAAAAAAMK`
          && `BAD///////////////8AAAAA+QYAAAAAAAADCgQA////////////////AAAAACEHAAAAAAAABADRAQAAAAAAAAAAAAAAAAAAAAAAAAAAAADpAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAApBQAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAEAAACDAAAAFRoaXNXb3JrYm9vawIAAAIGAAAAU2hlZXQxAgAAAgcAAABNb2R1bGUxAwAAAgoAAABWQkFQcm9qZWN0BAAAA+8EAgAAAAAAwAAAAAAAAEYMAAACLwAAAEM6XFBST0dSQX4xXENPTU1PTn4xXE1JQ1JPU34xXFZCQVxWQkE2XFZCRTYuRExMAQAAAgMA`
          && `AABWQkEKAAAKMQEAAAAAAAD/////BAAAAAkAAABZAQAAAAAAAMEBAAAAAAAAMAAAAAAAAAAAAAAAAAAEAAADEwgCAAAAAADAAAAAAAAARg8AAAI5AAAAQzpccHJvZ3JhbSBmaWxlc1xtaWNyb3NvZnQgb2ZmaWNlIDIwMDdcT2ZmaWNlMTJcRVhDRUwuRVhFAgAAAgUA`
          && `AABFeGNlbAoAAAopAgAAAAAAAP////8BAAYAAAAAAFECAAAAAAAA0QIAAAAAAABAAAAAAAAAAAAAAAAAAAQAAAMwBAIAAAAAAMAAAAAAAABGCAAAAh8AAABDOlxXSU5ET1dTXHN5c3RlbTMyXHN0ZG9sZTIudGxiAgAAAgYAAABzdGRvbGUHAAB/CgAACkEDAAAAAAAA`
          && `/////wIAAAAAAAAAaQMAAAAAAACxAwAAAAAAAFAAAAAAAAAAAAAAAAAABAAAA0zQ+C36WxsQveUAqgBE3lIQAAACPwAAAEM6XFByb2dyYW0gRmlsZXNcQ29tbW9uIEZpbGVzXE1pY3Jvc29mdCBTaGFyZWRcT0ZGSUNFMTJcTVNPLkRMTAIAAAIGAAAAT2ZmaWNlCgAA`
          && `CmEEAAAAAAAA/////wIABAAAAAAAiQQAAAAAAAARBQAAAAAAAGAAAAAAAAAAAAAAAAAAAwAAAgkAAABzYXlfaGVsbG8DAAANDAAMAAgAAAAAAAAAAAAJAAALHgAAAEgAZQBsAGwAbwAgAGYAcgBvAG0AIABQAGUAcgBsAAMAAAIIAAAAVkJFNi5ETEwNAAAHEQYAAAAA`
          && `AAD/////UwILAIgAAAAAAAAACgAACyAAAABIAGUAbABsAG8AIABmAHIAbwBtACAAUABlAHIAbAAhAAQAAAINAAAAVGhpc1dvcmtib29rMQIAAAIGAAAAU2hlZXQyGQAAfwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHJVgAAAAAAAAACAAAAAgAAAAAAA`
          && `AAAKAAAACQAAAAAAAAD///////////////8AAAAA//////////8JAAAAAAADAP//////////dAAAfwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVGhpc1dvcmtib29rAFQAaABpAHMAVwBvAHIAawBiAG8AbwBrAAAAU2hl`
          && `ZXQxAFMAaABlAGUAdAAxAAAATW9kdWxlMQBNAG8AZAB1AGwAZQAxAAAAVGhpc1dvcmtib29rMQBUAGgAaQBzAFcAbwByAGsAYgBvAG8AawAxAAAAU2hfAF8AUwBSAFAAXwAxAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `EAACAQsAAAAFAAAA/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKQAAABSAAAAAAAAAFAAUgBPAEoARQBDAFQAdwBtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAIA////////////////AAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAApgAAAJUAAAAAAAAAUABSAE8ASgBFAEMAVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAgEBAAAADQAAAP////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAACpAAAARgIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGVl`
          && `dDIAUwBoAGUAZQB0ADIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABJRD0iezhEODA3MTIyLTA2NTctNDJDOC1CQzZGLTRCNUZEMDgwMzFDOX0iDQpEb2N1bWVudD1UaGlzV29ya2Jvb2svJkgwMDAwMDAwMA0KRG9jdW1lbnQ9`
          && `U2hlZXQxLyZIMDAwMDAwMDANCk1vZHVsZT1Nb2R1bGUxDQpEb2N1bWVudD1UaGlzV29ya2Jvb2sxLyZIMDAwMDAwMDANCkRvY3VtZW50PVNoZWV0Mi8mSDAwMDAwMDAwDQpOYW1lPSJWQkFQcm9qZWN0Ig0KSGVscENvbnRleHRJRD0iMCINClZlcnNpb25Db21wYXRp`
          && `YmxlMzI9IjM5MzIyMjAwMCINCkNNRz0iREJEOTY2RUNFNEYwRTRGMEU0RjBFNEYwIg0KRFBCPSI1NTU3RTg2RTE4RTkxOUU5MTlFOSINCkdDPSJDRkNENzJGMDk2MTAxMTExMTExMUVFIg0KDQpbSG9zdCBFeHRlbmRlciBJbmZvXQ0KJkgwMDAwMDAwMT17MzgzMkQ2`
          && `NDAtQ0Y5MC0xMUNGLThFNDMtMDBBMEM5MTEwMDVBfTtWQkU7JkgwMDAwMDAwMA0KDQpbV29ya3NwYWNlXQ0KVGhpc1dvcmtib29rPTAsIDAsIDAsIDAsIEMNClNoZWV0MT0wLCAwLCAwLCAwLCBDDQpNb2R1bGUxPTExMCwgMTQ1LCA5OTQsIDcyMSwgDQpUaGlzV29y`
          && `a2Jvb2sxPTAsIDAsIDAsIDAsIEMNClNoZWV0Mj0wLCAwLCAwLCAwLCBDDQoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=` ##NO_TEXT.
    rv_file = cl_web_http_utility=>decode_x_base64( lv_b64 ).
  ENDMETHOD.

  METHOD build_xlsm.
    " Test workbook with a VBA project (macro enabled workbook, .xlsm)
    DATA: lo_excel  TYPE REF TO zcl_excel,
          lo_sheet  TYPE REF TO zcl_excel_worksheet,
          lo_writer TYPE REF TO zif_excel_writer.

    CREATE OBJECT lo_excel.
    lo_sheet = lo_excel->get_active_worksheet( ).
    lo_sheet->set_title( 'Macro' ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Workbook có macro (vbaProject.bin)` ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = `Bật macro rồi chạy: View > Macros (Alt+F8) > say_hello` ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = 123 ).
    lo_sheet->set_cell( ip_column = 'B' ip_row = 4 ip_formula = 'A4*2' ) ##NO_TEXT.
    lo_sheet->get_column( 'A' )->set_width( 55 ).

    " VBA project and the code names used inside it (ThisWorkbook / Sheet1)
    lo_excel->zif_excel_book_vba_project~set_vbaproject( get_test_vba_project( ) ).
    lo_excel->zif_excel_book_vba_project~set_codename( '{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}' ) ##NO_TEXT.
    lo_excel->zif_excel_book_vba_project~set_codename_pr( 'ThisWorkbook' ) ##NO_TEXT.
    lo_sheet->zif_excel_sheet_vba_project~set_codename( 'Sheet1' ) ##NO_TEXT.
    lo_sheet->zif_excel_sheet_vba_project~set_codename_pr( 'Sheet1' ) ##NO_TEXT.

    CREATE OBJECT lo_writer TYPE zcl_excel_writer_xlsm.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD get_excel_created_file.
    " workbook created and saved by Microsoft Excel 16 (not by abap2xlsx): values, formulas, merge, styles,
    " comment, hyperlink, validation, conditional formats, autofilter, freeze, table, chart, defined name
    DATA lv_b64 TYPE string.
    lv_b64 = `UEsDBBQABgAIAAAAIQA3FDkb/QEAAEkKAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADMlttuGjEQhu8r9R1Wvq1Yk1SqqoolF0l7VbWRkj6A`
          && `sQfWwid5HAJv37FZIhQBC2Kl9gazh/m/fzzrsSd3a2uqFUTU3jXsph6zCpz0SrtFw/48/xh9ZRUm4ZQw3kHDNoDsbvrxw+R5EwArinbYsDal8I1zlC1YgbUP4OjJ3EcrEl3GBQ9CLsUC+O14/IVL7xK4NEpZg00nDzAXLyZV39d0e+tkph2r7rfvZVTDRAhGS5HIKF85`
          && `9Q4y8vO5lqC8fLEkXWOIIBS2AMmaOkRNxPgEKVFiyPhBZgSDl0G7rGqKLMaw1QE/UepHCCt6ck1WFP8QxSvlcASQ0ccBnbHfVO+oFVSPIqZfwtLk8rXhrz4uZ94v69Mil859qUFthXa7iTnBLy8jL8PNwEZyfkX4Qh+3/4mPz//IB9CqjE6Yn9otke9fDV2hfe2eIiXq`
          && `NMDL7/UuikwPENPGAA69NopoH7kVEdRToh62GNzAvnaPD+ltbq04dNV3un0VFzMqAE95GNpCEe3hq23jRd79ud5DJ9Q37VSihFzmYTAm7YlF8Dx2+fbPYVvs9uGteok7Mz1vfDzr03rHkCUwr6M+kDDyvqVtaOA1LHe6p/h0LnmMPlAZfYTLDezOGTl6FEgIYtLwdtI4`
          && `tKG+EencdHXGuTE7BeoAm5eD4PQvAAAA//8DAFBLAwQUAAYACAAAACEAtVUwI/QAAABMAgAACwAIAl9yZWxzLy5yZWxzIKIEAiigAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKySTU/DMAyG70j8`
          && `h8j31d2QEEJLd0FIuyFUfoBJ3A+1jaMkG92/JxwQVBqDA0d/vX78ytvdPI3qyCH24jSsixIUOyO2d62Gl/pxdQcqJnKWRnGs4cQRdtX11faZR0p5KHa9jyqruKihS8nfI0bT8USxEM8uVxoJE6UchhY9mYFaxk1Z3mL4rgHVQlPtrYawtzeg6pPPm3/XlqbpDT+IOUzs`
          && `0pkVyHNiZ9mufMhsIfX5GlVTaDlpsGKecjoieV9kbMDzRJu/E/18LU6cyFIiNBL4Ms9HxyWg9X9atDTxy515xDcJw6vI8MmCix+o3gEAAP//AwBQSwMEFAAGAAgAAAAhAO6eZ1GDBAAAOwsAAA8AAAB4bC93b3JrYm9vay54bWysVmtP6zgQ/b7S/odsxFcTJ3HSJKJc`
          && `5amLBAhBgV2pEnITt4nIax0HitD97ztOn9CrVZe7qORhj4/PzJwZ5+zbsiqVF8a7oqnHqn6KVYXVaZMV9WKs3k8S5KhKJ2id0bKp2Vh9Y5367fz3385eG/48a5pnBQDqbqzmQrSepnVpziranTYtq2Fm3vCKCnjlC61rOaNZlzMmqlIzMLa1iha1ukLw+DEYzXxepCxq`
          && `0r5itViBcFZSAfS7vGi7DVqVHgNXUf7ctyhtqhYgZkVZiLcBVFWq1LtY1A2nsxLcXuqWsuTws+Ffx3AxNjvB1MFWVZHypmvm4hSgtRXpA/91rOn6hxAsD2NwHBLROHspZA63rLj9RVb2Fsvegen4l9F0kNagFQ+C90U0a8vNUM/P5kXJHlbSVWjbXtNKZqpUlZJ2Is4K`
          && `wbKxOoLX5pV9GOB9G/RFCbMmNg1b1c63cr7hSsbmtC/FBIS8gR+rBjZMjKUlCMMvBeM1FSxsagE6XPv1q5obsMO8AYUrt+zvvuAMCgv0Bb7ClaYenXU3VORKz8uxGnrT+w7cn4Zvi7rvnq7ZrC/p1G/biAo6vWxSWk4nrGqnaUn7jE0jhMLhCYV9Uy/uOxT6d5f+1NRd`
          && `yhzLRq6dYUQsd45m+txGFkup7TrYdRiZdimnIs1bmk3nBbCZ7umfHhbbf6gAmsqwahDXle+r588xhhBwb6PyG8EVeL6ILiHTd/QF8g7qytZt4QIS6zy9j7A7wokfoxgTHRGb+CiwHRfZoR4kekgcy7d/gBfc9tKG9iJfa0lijlUCwjmYuqLLzYyOvb7Idvu/4/UfkvdP`
          && `l83cD+mp7JoPBXvtdqqTr8rysaiz5nWsIh1D1337+Po6TD4WmchBjaZhQXWuxr6zYpEDY910XFmy3JDMxuo70bGBbewga+S7iEQRZHiEDRTHjm+4gWH6STQw0vYoDf0ZqA13pR5qKuovC9bDQSB7t4yurirck3vwi0yXPu1bB7Re7NlCo9zaGge2gBs1e9bmnrU56GJD`
          && `iC2Hoitv2ZxxOKMYkDwY221FhsU/WwT1XdQsk+0CIPbe1t4+Lcu6On1KClnlspRmtGOyi0A93W0iAHLLiyxj8tBUz1cB+uPEPzG9k+8nzpm2B/uzPaKG1vkk73dLkxPinSQHS/eBgC1QSKFFyduQBxdS7EpPwdHLTgx36A4FZF8n2IcSIAjHpoWI4xrIIaaBQhIZsTWK`
          && `oziwpPzl8e39H4fY0KS8zXeBZJlTLiacps/wNQFpCyCOMlyy2oHvPtnAcgJsAkWS6AkiuotRENgEWVFiWiM9CmMr2ZGV7s+/eIQ42rCaUdFDe5WddXj35DVZj24H56uBtTA+9DTvNpKOrFf/m+EdeF+yI42ThyMNw+urydWRtpfx5OkxOdbYvwoi/3h7//bW/2sS/7nZ`
          && `QvtpQD8nPNKJi83YR6YZEkRGyQg5CbaQSUYktEgQ63i0S3j5mr58Ld8G0TaKDPe/8ta9XuZfgnvrT2ClY2I9BV12p1RJf6ivLdr5PwAAAP//AwBQSwMEFAAGAAgAAAAhAGqqelQnAQAA8wQAABoACAF4bC9fcmVscy93b3JrYm9vay54bWwucmVscyCiBAEooAABAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALyUwUrEMBCG74LvUHK3abu6imy6B0VY8KTrA4R02oSmSclE3b69oYd2C0u8lL0E`
          && `Zob8/8dMJrv9qdPJDzhU1jCSpxlJwAhbKdMw8nV8u3siCXpuKq6tAUYGQLIvb292H6C5D5dQqh6ToGKQEel9/0wpCgkdx9T2YEKltq7jPoSuoT0XLW+AFlm2pe5cg5QLzeRQMeIOVfA/Dn1w/l/b1rUS8GrFdwfGX7CggmvxIrkyQZS7BjwjUyoNpIRehtisCfFrXYsS`
          && `wM8QUwrpWNnEYB7XhEHJHVSf3oWB4wy0SMdgiit3pojB5FeGyWMw21XH5AcdNm96tTjGMfuHNe192GeY3ceQjme0BfdrMsDJgzNcvyvTzijnWaTn0URGF19V+QcAAP//AwBQSwMEFAAGAAgAAAAhAHJi6Xi4BwAAnxYAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54`
          && `bWycWF2O47gRfg+QO2iVxaLTaFum/uW1vbBladzYmUxjunvyEOSBLdO20LKkpeT+SbB3SI6QG+Q5z4vcIzdJkdQv7en17GDapllflao+FilWTX542SfKE6FFnKVTFQ1HqkLSKFvH6Xaq3t+FA1dVihKna5xkKZmqr6RQf5j9/neT54w+FjtCSgUspMVU3ZVlPta0ItqR`
          && `PS6GWU5SkGwyuscl/KRbrcgpwWuutE80fTSytT2OU1VYGNNzbGSbTRyRZRYd9iQthRFKElyC/8Uuzova2sv6LHtrip8h1tqfjotLIWnsIfPIv30c0azINuUwyvaacO04Sk/zenHuo3MC3WP6eMgHYDiH4B7iJC5febiqso/G19s0o/ghgRV5QSaOlBcK/3X4MzoO4+Mn`
          && `ne8yjhpLx0yeZQaZGiVPMUut1pT+21hEVmNLb40Zv9GY3RhjdNHxIV5P1b/7xsKeI98cBIFnDcyF5Q3c0dwfzBfIco2FMQ/Q6Gd1NlnHkHssKoWSzVSdo/EKGao2m/DU/hyT56IzVkr8cEsSEpUEHoJUhe2chyx7ZMBrmBqByRynRHm9zWGVpyr4VGb5e7IpfZIk8ABT`
          && `VXBUxk/kBmBT9SEry2zP5HxvljC1odnfSMp94I9izjGbfbAw0hotfhIBmExRa7zvjutIQr6Pb6iyJht8SMpP2fOKxNsduIvMITjI03C8fl2SIoKdCWENOSdRloAN+FT2MTthIH3xi+AhXpc7GI2GjuMYpuFYqhIdCgjtz5WEudVowrJzTfh+FnLdPU8T+OSa8F1pIv08`
          && `TfCIa8J3rWmep2lXmvBda9pvamqCKL4IS1zi2YRmzwrsPGCsgKWEExaN4TDmjMP3Kb4tyKSI6cyZEleFhShg9mk2mmhPsLBRhVgIhM4pZiq+PLGUJwJ5IuxMaOBt4zIQLbtsGkO2b9lZJSdJ7TPTEslf+4wkn48Reh/hHyOMPmJ5jDD7iOAYYfUR4THC7iPeHSOcPmJ1`
          && `jHAbRI9M2FtdMt+kkGGnKnw2y+5JFB4jkJQZvoCwXHqambZuS7Ethdzmclk5EEJHCHVrNBoNZfZayGa2NC8D4H/DnoU4XEK/E2gXDG5moand3n+4+Db81hzDh/vHStMeWrpjIst2Xc/RTc8xg4GUGatu4A/CvdN8w14/m2+G7fON5Jw9AZGTVkAawh3JxFLIBeG6tFqB`
          && `ENaE26cIbyFAuHUZAMWccN2CO5i8gu8EuiLcOk34aIgMZHm67dkI6UC8Lvm86oYt6G4d76U3HJBn082wEt3S9l4IiNfZAboE8QWkpVvObyEXdBsy3UJY0+2coruFAN32ZQD2Od2Gi4BuOb8FuqLb/iLdnutZlo0c23Z0z3Dk7O4y82Z2O19BN8NKdEtn5eIERArQF5CG`
          && `blfObiEXdJsy3UJY0+2eoruFAN3OZQAHLafbQpDd8nZ5J9AV3c6X6NZtx7B1DzLbshBcTGS6u2G/md2seuq8vN88vBlWoltKzcUJiPRe8QWkpVvObiEXdFsy3UJY0+2doruFAN3uZQBvLU63bcJhr8vZLdAV3e6X6DYM20WmPYJ7IJzh8Ce9K7thv5ndsO/Pppthp2r3`
          && `pEDtG5jfohYCwm/HfMKXJ5byRCBPhGICsUv+ZnZ7v7j7eDd/f+FdheY4bF5gyEPAniOzJxtbdSZ6hyiYPzvwBQfD5RBub801QW8zvG9Yun2+lcDLnkUk3TwCkDIKPr//+PHH+5sL9ccVHIeGejVnRFxZV+H8/W1Qv9GRLh2tfa/gKnnuOi8RgNs45YyvpSXl3l2HF6H3`
          && `3bb8Hnxj/67UX/7xv//8q1SvVH/3339j5Zd/8p+1n0J6+iIBNeFXeMmKvvreK79NA2YKpKRKog8Xf0F/XeB0+42vj32zduYPn4LwG8kXUdWJggIfyiyMk5LQqnA1xitWSXxd/Qs12Z7QLWFlZKFE2YEVe2xtm9mmLIYCgRWWLZyVc+k6Zs0anIiSsoTui1IVonxPsBpm`
          && `8+mQEKV8zaGAhbooo7cRToiq5DTOKDRCxANbCVN5yioFqDN5VdPO5ISymjRmJp4wVNTWSEJAlVcVmxmQs32YqmEYunC9WRxPh8HCFWcC86xG28YicDiaVXKVx6zg4rGwwanIf4WQpTlessOzTwgQf12oyvplI9oHHVrYxoDewfXmjh6AO9i90IejuMygwttC9w3W/m6H`
          && `U5ZHUNAfEjxjN5x6/OverqEy/YyTGL5Zw627/H1RtRZJXECTAidJ9rxIcPooytld9nyd5ofyAykKvBV+FjAZUJrR3qRoUKzMfqLa5nzpBUtjgAKkD8wgmA9cPTAHLoJhGNiW79isUVPFhWbq3af7QJwwahMugkXpO300AZ2LHSQhTeL0sTsWGc5PUjpmvSN6vWbHaZYl`
          && `ZZxP1TsKm1O5nd90dhdaQFkCr9SBMff9ganr0FiybGtgueHIDEaWPdedn/l+6T4yB34+YLqNgewEOj6srwK3DypaL3wM681n4cotOkP1rx10WwksPOvEKJssg8Vv2jLM7i0pD7kCGwr2Bl/PqQprtC4inMNGaQLjPQJhK+RGZpNsva6G3+F9/r0v4uXjG0VT+OBPE62F`
          && `TbS+gYRscfRatVjbR/GOkda0lmf/BwAA//8DAFBLAwQUAAYACAAAACEAtn42WacCAACLBgAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQyLnhtbJyV327TMBTG75F4BytXINE6f9smSjN1zTomgZig2w3iwk2c1loSR7bbbkI8AU/Bc3DJk/AmHCc0WdshlUmtFOecfOfn`
          && `4+8k4dl9kaMNFZLxcmxYfdNAtEx4ysrl2LiZz3ojA0lFypTkvKRj44FK4yx6+SLccnEnV5QqBAqlHBsrpaoAY5msaEFkn1e0hEjGRUEULMUSy0pQktYPFTm2TXOAC8JKo1EIxCkaPMtYQmOerAtaqkZE0Jwo4JcrVsmdWpGcIlcQcbeuegkvKpBYsJyph1rUQEUSXC1L`
          && `Lsgih33fWy5J0L2Anw1/Z1emvn9UqWCJ4JJnqg/KuGE+3r6PfUySVul4/yfJWC4WdMP0AXZS9vOQLK/Vsjsx55lig1ZMt0sEa5aOja/+yLy4mMSjnjO1/J47nc1657570RvFZhxb5sCferNvRhSmDE5Y7woJmo2NiRVMPQNHYe2fW0a38tE10nZccH6nA1dQxtSp+Ch3`
          && `VtvxWqCUZmSdq498+5ay5UqB992+C5j6nIP0IaYyAYOBUN9pq8ZEkSgUfIvgrCwYjIpo51uB3t9TD0ZholMnkAtSEna1iWw3xBtgS/4Gz/eC3n5wuhcctEEMEC0JHNXJJJDbkQwPSPaCowMSu6a3PNN8mgJ6cDIF5HYU/gHF46DTFas7OXWaHg69f1DAEZ5MAbkthWMd`
          && `UOwF7YNeuDUFOPnpVnj/AQG5HYRzUMeDOln06eZ8/mE+effKMv03apHfEjVff75kv34gJX7//P7ldYgzbS3PftSvxiLNCDTGrciSvidiyUqJcprV5h4aSDT+N/twrXilLT8EqgVXihe71Qre3BSsrMcBZZyr3QLGTOlX5DURSqKEr/XQWMDd3kUi0IMvrlKrHsouHaag`
          && `/YxEfwAAAP//AwBQSwMEFAAGAAgAAAAhACtlAMzEAQAAqwMAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0My54bWyck8GK2zAQhu+FvoPQ3ZbtOMnGxFm2a4fmUCil7V2Rx7aIJRlJ2SSUvnvHDvEWcgkLEmik0Tczml/r57PqyBtYJ43OaRxGlIAWppK6yemvn9vgiRLn`
          && `ua54ZzTk9AKOPm8+f1qfjD24FsATJGiX09b7PmPMiRYUd6HpQeNJbaziHk3bMNdb4NV4SXUsiaIFU1xqeiVk9hGGqWspoDDiqED7K8RCxz3m71rZuxtNiUdwitvDsQ+EUT0i9rKT/jJCKVEi2zXaWL7vsO5znHJBzhZHgnN2CzPu30VSUljjTO1DJLNrzvflr9iKcTGR`
          && `7ut/CBOnzMKbHBr4jko+llI8n1jJO2z2Qdhigg3PZbOjrHL6p0yL1/lTvAwWi6QM0uILauxltQ228SIq03JZJq+zv3SzriR2eKiKWKhz+hJTtlmP4vkt4eT+W5NBi3tjDsPBDmNEgyu7892OWvxuSQU1P3b+hzl9Bdm0HoWfhinmODQ5qy4FOIHqQlA4m6IW3HPE9ryB`
          && `b9w2UjvSQT06LSmxV04U4tqbfri6nFOyN94bdbNalD9gmwcsqY3xNwO5leUn/HTEZsMr2V011sumb7b5BwAA//8DAFBLAwQUAAYACAAAACEA9mC0QbgHAAARIgAAEwAAAHhsL3RoZW1lL3RoZW1lMS54bWzsWs2PG7cVvwfI/0DMXdbM6HthOdCnN/bueuGVXeRISZSG`
          && `Xs5wQFK7KxQBCufUS4ECadFLgd56KIoGaIAGueSPMWAjTf+IPHJGmuGKir3+QJJidy8z1O89/ua9x8c3j3P3k6uYoQsiJOVJ1wvu+B4iyYzPabLsek8m40rbQ1LhZI4ZT0jXWxPpfXLv44/u4gMVkZggkE/kAe56kVLpQbUqZzCM5R2ekgR+W3ARYwW3YlmdC3wJemNW`
          && `DX2/WY0xTTyU4BjUPlos6IygiVbp3dsoHzG4TZTUAzMmzrRqYkkY7Pw80Ai5lgMm0AVmXQ/mmfPLCblSHmJYKvih6/nmz6veu1vFB7kQU3tkS3Jj85fL5QLz89DMKZbT7aT+KGzXg61+A2BqFzdq6/+tPgPAsxk8acalrDNoNP12mGNLoOzSobvTCmo2vqS/tsM56DT7`
          && `Yd3Sb0CZ/vruM447o2HDwhtQhm/s4Ht+2O/ULLwBZfjmDr4+6rXCkYU3oIjR5HwX3Wy1280cvYUsODt0wjvNpt8a5vACBdGwjS49xYInal+sxfgZF2MAaCDDiiZIrVOywDOI4l6quERDKlOG1x5KccIlDPthEEDo1f1w+28sjg8ILklrXsBE7gxpPkjOBE1V13sAWr0S`
          && `5OU337x4/vWL5/958cUXL57/Cx3RZaQyVZbcIU6WZbkf/v7H//31d+i///7bD1/+yY2XZfyrf/7+1bff/ZR6WGqFKV7++atXX3/18i9/+P4fXzq09wSeluETGhOJTsglesxjeEBjCps/mYqbSUwiTC0JHIFuh+qRiizgyRozF65PbBM+FZBlXMD7q2cW17NIrBR1zPww`
          && `ii3gMeesz4XTAA/1XCULT1bJ0j25WJVxjzG+cM09wInl4NEqhfRKXSoHEbFonjKcKLwkCVFI/8bPCXE83WeUWnY9pjPBJV8o9BlFfUydJpnQqRVIhdAhjcEvaxdBcLVlm+OnqM+Z66mH5MJGwrLAzEF+Qphlxvt4pXDsUjnBMSsb/AiryEXybC1mZdxIKvD0kjCORnMi`
          && `pUvmkYDnLTn9IYbE5nT7MVvHNlIoeu7SeYQ5LyOH/HwQ4Th1cqZJVMZ+Ks8hRDE65coFP+b2CtH34Aec7HX3U0osd78+ETyBBFemVASI/mUlHL68T7i9HtdsgYkry/REbGXXnqDO6OivllZoHxHC8CWeE4KefOpg0OepZfOC9IMIssohcQXWA2zHqr5PiIQySdc1uyny`
          && `iEorZM/Iku/hc7y+lnjWOImx2Kf5BLxuhe5UwGJ0UHjEZudl4AmF8g/ixWmURxJ0lIJ7tE/raYStvUvfS3e8roXlvzdZY7Aun910XYIMubEMJPY3ts0EM2uCImAmmKIjV7oFEcv9hYjeV43Yyim3sBdt4QYojKx6J6bJ64qfEywEv/x5ap8PVvW4Fb9LvbMvrxxeq3L2`
          && `4X6Ftc0Qr5JTAtvJbuK6LW1uSxvv/7602beWbwua24LmtqBxvYJ9kIKmqGGgvClaPabxE+/t+ywoY2dqzciRNK0fCa818zEMmp6UaUxu+4BpBJf6eWACC7cU2MggwdVvqIrOIpxCfygwXcylzFUvJUq5hLaRGTb9VHJNt2k+reJjPs/anaa/5GcmlFgV434DGk/ZOLSq`
          && `VIZutvJBzW9D3bBdmlbrhoCWvQmJ0mQ2iZqDRGsz+BoSunP2flh0HCzaWv3GVTumAGpbr8B7N4K39a7XqGeMoCMHNfpc+ylz9ca72jnv1dP7jMnKEQCtxV1PdzTXvY+nny4LtTfwtEXCOCULK5uE8ZUp8GQEb8N5dJb77j8VcDf1dadwqUVPm2KzGgoarfaH8LVOItdy`
          && `A0vKmYIl6BLWeAiLzkMznHa9BfSN4TJOIXikfvfCbAmHLzMlshX/NqklFVINsYwyi5usk/knpooIxGjc9fTzb8OBJSaJZOQ6sHR/qeRCveB+aeTA67aXyWJBZqrs99KItnR2Cyk+SxbOX43424O1JF+Bu8+i+SWaspV4jCHEGq1Ae3dOJRwfBJmr5xTOw7aZrIi/aztT`
          && `nv2tQ64iH2OWRjjfUsrZPIObDWVLx9xtbVC6y58ZDLprwulS77DvvO2+fq/Wliv2x06xaVppRW+b7mz64Xb5EqtiF7VYZbn7es7tbJIdBKpzm3j3vb9ErZjMoqYZ7+ZhnbTzUZvae6wISrtPc4/dtpuE0xJvu/WD3PWo1TvEprA0gW8Ozstn23z6DJLHEE4RVyw77WYJ`
          && `3JnSMj0VxrdTPl/nl0xmiSbzuS5Ks1T+mCwQnV91vdBVOeaHx3k1wBJAm5oXVthW0Fnt2YJ6s8tFswW7Fc7K2Gv1qi28ldgcs26FTWvRRVtdbU7Uda1uZtYOy57apGFjKbjatSK0yQWG0jk7zM1yL+SZK5VX2nCFVoJ2vd/6jV59EDYGFb/dGFXqtbpfaTd6tUqv0agF`
          && `o0bgD/vh50BPRXHQyL58GMNpEFvn3z+Y8Z1vIOLNgdedGY+r3HzjUDXeN99ABOH+byDAkUArHAX1sBcOKoNh0KzUw2Gz0m7VepVB2ByGPdi0m+Pe5x66MOCgPxyOx42w0hwAru73GpVevzaoNNujfjgORvWhD+B8+7mCtxidc3NbwKXhde9HAAAA//8DAFBLAwQUAAYA`
          && `CAAAACEADB1NRGoEAAC+DwAADQAAAHhsL3N0eWxlcy54bWzEV1lv4zYQfi/Q/yAI6KOiw5ZrG5YXsR2hC2yDBeIC+0pLlE2Eh0DRibxF/3uHpGTJOb1JtvWDLFLDmfnm5uxTzahzh2VFBE/c8CJwHcwzkRO+Tdy/1qk3dp1KIZ4jKjhO3AOu3E/zX3+ZVepA8c0OY+UA`
          && `C14l7k6pcur7VbbDDFUXosQcvhRCMqRgKbd+VUqM8kofYtSPgmDkM0S4azlMWXYOE4bk7b70MsFKpMiGUKIOhpfrsGz6ecuFRBsKqtbhEGVOHY5k5NSyFWJ2H8lhJJOiEoW6AL6+KAqS4cfqTvyJj7KOE3B+G6cw9oPoBHst38hp6Et8R7T73PmM71nKVOVkYs9V4kbH`
          && `Lcd++ZyDj0dD17FeWYoc7JTnPmP+AX6u37I4oY9P6YOL4DdN6TfS5rNC8E7oCOyjLT+95eKep/oTCAVNNNV8Vn137hCFnVDzyAQV0lEQMaCI2eGIYUtxWSpROddISnGvaQvECD3Yb5HeMJHWEDMCfjdaWTH2udFU/5vA4X+McNLJIzzHNQZvjx8YdY12gqEnzen3Lbd/`
          && `xXKBiZWP9NVTfpLbTeKm6TLQvw+UaKBWEFKE0mO2DHSMwsZ8BoVFYclTWDjN+/pQQoRyqIE2yAzdK9RbiQ5hFJ9/oBKU5FqL7dLkRQN/NbkKUxPxm+bD0b2QyzoTewrDyqpl/gDjRsgc6nu/Jtit+YziQsFxSbY7/a9ECc+NUApq4HyWE7QVHFEtoD3RPwl9AVpA4qod`
          && `lPA2lR9qpkU0Es6iN7oYVc4iB5Vbjc+it+CextaABJNlmNIbDe5bcWK3uujVReiUOmF0SdWvEBDNq7WRXYDtnjukq/DThxxUlvRwvWcbLFPTPo0Is6tjslstjG+79SUlW86wrv2gkznwVQqFM2Xau0khv4/OYu3BDHXP+HGcTl28Cjh6HnB72qqse4buBnbVw6S7CGoh`
          && `OjshyXcg1d0kA8zY1v+6eBeAjxcKQFpHgwm66ID9l3BbT7dWaD39qlXuJSrXuDYG1Jn7FnM80qyT/jCgm2ni5Tx4xO80th9jfChl8HquvVuGmYp+Mo5wZGapnyylV5l6sQe7Z8fe8/7uWels3k97+zTLn5fYQ/NCybSZZIYdU+CgpPXq+EkVP9Y/R8+aifsHtHZJCb9t`
          && `8xGAbfaEKsJ1TRubqartCM2Za12WaS+BewceVFnQI6+7PqJrGKzbcVhPqMTOtOacHYibxp+m7dzTzGa+OamfwFXpa45pU0d4YKscF2hP1fr4MXG79z9xTvYMqlBD9ZXcCWVYJG73/kXPA+FI44ZC8qWCBg7/zl6SxP37avH7ZHWVRt44WIy94QDH3iRerLx4uFysVukk`
          && `iILlP73L1juuWuZuCMU8HE4rChcy2YBtlL/p9hK3t7Dqm7kI1O7rPolGwWUcBl46CEJvOEJjbzwaxF4ah9FqNFxcxWnc0z1+45Us8MPQXu608vFUEYYhvlpftR7q74KTYPkCCL/1hN9dvOf/AgAA//8DAFBLAwQUAAYACAAAACEAbmbZc2YCAADcBQAAFAAAAHhsL3No`
          && `YXJlZFN0cmluZ3MueG1sxFRPa9RAFL8LfofH3G12t7VoyaZsV7aB6rqw0fuYTpNgMrPOzFbXU4uIBy/twXuXnqwWighCc/CQZb/HfBNfEmhpZg8eBEMIk9/7vf/vjbv9LkvhkEmVCN4l7bUWAcZDsZ/wqEteBIMHjwgoTfk+TQVnXTJjimx79++5SmlAXa66JNZ6suU4`
          && `KoxZRtWamDCOkgMhM6rxV0aOmkhG91XMmM5Sp9NqbToZTTiBUEy57pL1DQJTnryZsv4N4Lkq8Vzt7RTHz6FffvA0BL84Gu5C4ON5Fx6jrc6m62jPdUp2rfGsOIc9v4kGxQWH13ExD2OIizMeNQnDqDibNcGxyU8hXV6Z/NzWWJws5xyipJhbzkoPMejE5MfcEpr8F6Qm`
          && `/9QULE4w8iBownt+u9VqN9F+8ZNHoGcQDH0fimN092R5hQFZiVX6HUvf5Je0rgQsTkz+NYQRVkev9L6+Et2wqpUUF1PQsck/wzCazkz+kcNekq3Ufmg1AIsGqmrQKK5TgSG1lAOTf7GT3MUuwHiEjcdi2eFiIz5koCVd0Qy0Btrk36AaUNihtvWnmBhOj7n+rWHcG901`
          && `IssxDcz15QQWpwIbX4sRLt8RftR7OKQp7lebOJ4bilRIrBLLcKEqRA4E1zWlN9FCYdpSircl+YBmSTqrZZ0SqNaM1UCWcCFL0KncVBu5pSY0RMO4corJQ0Y8+Kt4ZPSqSwaDfqt8Spv/KijPL35QXN7v/68uXnxTgrv3hDUN2OYmVk2WljjTTUnP3sqXOAYa9PIKejbb`
          && `2sFb9o7Ntob4lt234haapregg7ez9wcAAP//AwBQSwMEFAAGAAgAAAAhAGE8THJpAgAAmgQAABsAAAB4bC9kcmF3aW5ncy92bWxEcmF3aW5nMS52bWyMVMFu2zAMvQ/YPwjaIZeksd01aFXbQNeht23ANmCHoSgUS4nVyqJhMa7Trx9lOWk7FMOCOJZJ6onv8Tn50FhG`
          && `l/OiL/iuc8JXtW6kXzSm6sDDBhcVNKJvLH//bqqEf1XCZmMqLeLtec/wH3v0UGnLSzonB+Fr2Wor97BD1gs9YMG1MjimQ96oRravMkxJlAVP+XKEWL7CKPM+QuK+1cyogt8NCX3uMEsyziqATnnzpAuepaskmY+/nBFGSyeHGmqLtRLrgjdzG/NdrLXxNuipOToJO3jQ`
          && `7B6M87i3hNoY1F3sjFErAYhtO6mMdjhShYeCYziwAud0haHPgne0OvB5QeDI5iUTnybZGWdx44dX9GITsxa8QQNOyLUHu0N9GUg1stsat7B6gyLLspNVi5dTDKEVq4sx8mgU1iJNkpNzytfabGsUq2TMPS2MU3oQ6YjXG2/Wxhrci9oopd2MbYy1FVjoCm7cBj7J6mHb`
          && `wc4p9vs8ueUs6jVVOHCa4untKDlN2nmNDSiSQ+4QjiIHUBocoWZvw44+CGqTwAoeY23Bn/EZrH216zS5YVL5OJu/5hD2PE8PyY5rGKjtMNtZ42GhTBgVabuQFkXocxZcyFiuTH8oDPsob7ZOBLVnZb6k7FiXL3sxwcbnQVzbYI7PZGr2bX1P6D9HS3wFnIzG8kF8gV7/`
          && `Mlhfa2t9JBzCP8jKb4SvXFVDFxtj2ZyldNH348Wcnc7Z2cc5W1HwLF8O4mUpAV4RoxsSvLyR1uux4BCJNAfxHR7L05AJi0PwmnzWuDIN8WkdUuPjkV98Xyd/kyj0h1T+AQAA//8DAFBLAwQUAAYACAAAACEAolONViQCAADOBAAAGAAAAHhsL2RyYXdpbmdzL2RyYXdp`
          && `bmcxLnhtbJxUUW+bMBB+n7T/YPmdgAklCQpUJIGp0tT1YfsBljEBCduR7Sapqv732ca0UZZJW1/gfMd3d/6+O9b3ZzaAI5WqFzyHaBZBQDkRTc/3Ofz1sw6WECiNeYMHwWkOX6iC98XXL+tzI7OT2klgEnCVmWMOO60PWRgq0lGG1UwcKDfRVkiGtTnKfdhIfDKp2RDG`
          && `UZSG6iApblRHqd6NEejz4U9kY7jnsHCd6ZPY0mEoOemEHF2tFGy0iBiKaB3aG1jTAYzxo22L+C6Joo+YdbmwFKcCjRBrTj4bXyAUe4QJOYRL/VFPi4+6q9t1UYTSv9b1mOvCSxQvbxWeyu0lPnQ9qSVmFDBMpMihp4cfv10EnzxD5PH4JEHf5DCGgBtUDrcdlhogA8MZ`
          && `PevvSnsLPMs+h691HW/uqjoJamMFSbRJgk2VrII6ni+reFFv43n6ZtEozYjRWpsxe2gmjVH6h8qsN30q0eoZESwUbdsTOk2NmRmUhE5l1+VrvEs3ySItg7RalMEWVZugikwz81WJ5lVUpuVq+wbDYh267qe3u8Uov72xF+fxihMLsyNym6tzK8004cx0CM45NGvzYp8G`
          && `5EoBMjrJ5HWpJpBXxn7rzR3WeKT0v3eIWIkMxSRzlueWfGJ/xkw+wT+t86jPTpBnRrked1rSwcmsuv6gIJCZHSj50CAvxMWNnSDvZDiKLofWCzP0JrclaFLkarkdzv6Kit8AAAD//wMAUEsDBBQABgAIAAAAIQBCQEuVnQcAAD4aAAAUAAAAeGwvY2hhcnRzL2NoYXJ0`
          && `MS54bWzsWc1uI8UW3iPxDn17suM67va/rXFQ3E4YRCYTjcMgsStXl+0i1VU9VdVJPIgFC1Z3c3kDECsQSAixIsugeY+8Cad+2nEcG5IZJBhEFknX36lzTp3z1XcqD989z1hwSqSigvfDeDsKA8KxSCmf9sMPj/crnTBQGvEUMcFJP5wTFb678/ZbD3EPz5DUoxxhEoAQ`
          && `rnq4H860znvVqsIzkiG1LXLCYWwiZIY0NOW0mkp0BsIzVq1FUatqhYReAHoFARmivFwv77JeTCYUk6HARUa4dlpIwpAGD6gZzVUpDcctWbslMaNYCiUmehuLrOqElUaBsLhZXVi1A05KkSZxN2oEp4j1wyismk6G+NR1EF75cOQ6pSh4StJESA7HsTQ/w71dponkICoR`
          && `XIPW3l/ZnTyeIXlS5BVQNwcjx5RRPbdmhzsPQXYyE+CP4Cl5XlBJVD/EcePaBY37OiBqVzvVmj9XMDZu9JSeM+IMiqOasba62NeqsI8YGyN8YnyzNHkx9XrcLFx1hlllw8h8aKoZsR/n5rekeLbzEPXGIp0fyUAKbQ4hUDnep1LpA6T0EZIQd3FoskA/gV8TJs76IWEM`
          && `YoFCNJh+8IKQL8LgTKK8H6rnBZIkDBDH0A0e07JsJBrasbER9ZjSI2O6beSmJz+S5k9KJk9BG/UCpjYi0GdstaL2d9EPOaSaSTtJTyDluBjZrzA4gSCAJZA41gQ7fYwUYdSkZgTniXpKMJruU8Zsw+QhSZh03tfnsZ3DiuyxSF1fqxmBOKdvkT2ZTFx3veyugshSChzc`
          && `ygYmbXig5zmZAA70w3cyXmHaiSNoZYAgN4DVygBWNihKx9htvKusvyQ4y+RMPzyllWeHToreGV1dfBmwlz9eXXwD+aRnRAQns8uv8SyYXX7Fp0ZZbWWBEPiGE6iWEQEf2gaIgNNlaH4jPZXfmwvjR+cabty56ABpzPaQyYRgfaC0CWoIXrvSyHYH/W/YGdR7A8OO8NTg`
          && `gslSF3gLoF4EkjtjE0gl5KBCi2PTGBJGNPEJ5jE/Z0LvSoIc/s9FYUIG98ZIJuYe9d9D6lMVC+auhSncCzlcmC5CMSsUXAUkdYOnSM4TwcSN+wLCkEDE4x5Nz2/EtZAp8eK9Ui4HAGiekolZMdkZFgeUFP/ZGm7VTUDbXhhPEGCAmZHrBO4p7dHcaZHrALZyAIR7pzfz`
          && `0kg5tcmRm2Q0AF9Kcw279SIfy9z7PRhDGMMdaEH2Fh7dL08pN+D+/uSQTAHITv0l5Z2DkT2Vdd4ZbDV6W4Otzh+7qLneRcnlzway5sHx4aNHweXnfBYMX/748muDWsv+Mg53zjXIbUaSq4sfkAO44Nf/X118i4MjAD29aV3Nrzu++uWHPPj1SwHQGDy6/AkFg8vv+WzT`
          && `srpfNqKX3xWArVcX/wsOp8X86uILHnxAs03rGn7dIUBwoCwYH82cZcEhWln2O9HgnQ+syZjNi2xNhMIZDJfOACYtwtQxzkSkZOfBfx88iOxBWRZq+1YDecMpxXbddfTePo3axhml3+sbZ5QubmycUTqzuTID/HZtrGuUWeQdRs7NpWRcB19BIYFafJrU650oSRqV5rC1`
          && `V2lE3aQy2KvXKnvt2rDWbdeb9ST5bIn83pv6NpZYX6tXcPq8IO97FPwUWIb9qbS7u7B70ogqnQh2b3b2h3vdWjIYDIefWRpgdbZQUVph4sRBWnowZsrSxJk4OyBTwOkPyMrtDSPPEBQ1S4RbQV+CNATgzRQ3/SMi1/YfEWlg5pacQTEeMzKiL5ZFgYYL1aYo/4imeuZW`
          && `1uKugwDHNHLXW6m1XS86Lz3UaHQ7cdxottYO1OK2HYCNlq8MSJNdy2M2ylEYAT2cGp8JScEgW+g4LTLKH6Nz7/Sliam9v25Yjs6PhC9Kxk5BiMD9TAfXmdYP3yNQuiAG5FQU4LwDyk9I6ukw7mXoEyGPKT55DAWJE26Jrr0HQZXNgxoWwcEvNOAQGMfCqbGWrAVA4LvN`
          && `WjMMsOHrE6Cp8JnloIziU+DqbApE2tD2V6HM8XrK3LkbZbZ1nqGLr8AhKy2fRX+nEqb7ihXMP7xyuTOFhOoVXhTUridrFgfKdPdjxKeeYZgfE+lD37RuEDE2ZrsQ2Z4qQnS714Yxg7JOkZK0uXIPrhDxuGCaHpwyyK2lZAeIWcAKgOhafLnW78/HF098LVy8J2lqqluH`
          && `+b4uY/zNz29w8m0D10CqZS9/PaC+dvX7L3K9OW8ur4Vc1wzGotp65LJDA6LPCPFoNXYNz0UWsPO6TzGQZsuFt/l+RtUTzjxntEUkcDeq8gEU+idq1yMtkDjPL4AiDg3rNE+DwJZusKJ1NLvZGnT3knhYqXeaQHTjRqfSjeL9Srsz7ERRrd1uJ9EyzZb1exPtdjWqL3Nt`
          && `WTfvzGgIVsBb1pPcPmXHLeB88IQNY9B/iHaVtXDpwqj64bVLwXNg3SoNX7yz3qVOH0831Oj/PHp2l/dex/n8HXaPHMsl5XpENLy0Tu09OCMIHnL2hYCXIHvD52hKgFRPKVf2KXm7DczX/Lthux0G8B5t/5qXbzvglptWPQT6bqS4xkIWbFbk7kVzZXPDDRb/9tn5DQAA`
          && `//8DAFBLAwQUAAYACAAAACEA1S0YMZwEAADCJQAAFAAAAHhsL2NoYXJ0cy9zdHlsZTEueG1s7Fptb+I4EP4rkX9AA/RoKWoqdVutdBK9rfZWus8mccC7jp2zzVL662/svBCThJct0MLtNzwJtueZ8TPjmdyGahhOsdR/6wUj3kvCOAhUgKZap0PfV+GUJFhdJDSUQolY`
          && `X4Qi8UUc05D4kcRzyid+r9Pt+ctZUD4Nrs0iUsJhiVjIBGt1IeSkmCNhMEvnyk8w5cijUYBgUnR3C9vDL1R9o5oRO2L8K4nhhZcAdZBvRTFlrCYkcUxCXX9X8KUwoVxIWAQPrZrkgUnvJ2YB0i9mbTxks+RJRJnsqt/p2BWt+EscZ+LLQuxXZrm79WHjcbaW3WNE4q/P`
          && `0lOvAeqaebwfRHL4DUobLczrrp4h1mQi5OIetD9lxVX6LC2U3JsH6Kbf6yMvxGmAYoY1/ExSsLXiE+RhNgFEQp1bRDAafQbLbmmebrN5Bi3mAWNVF5BixiNjBzAuz8yRbRzsUjHdTbPlarayZ+FeEuwlIoLDhBkT87+EUefLTyIljQioa2Ujykkhy/z9aB6e+13VTQtr`
          && `OeCsHI7xpGtddgVCdg4GbjubpUGzw4w1HuExAd8E/3g3RrpudvnezozU4taRo2Y5ejCeO9PH0z36sW823sbNmW5xc0tIa9ipabulTdzgUVpwJXg4R2t3TgKnHItoAfFGCm3ipKfS8DOVSo+w0s9YQmTuIg+YSBvqiYGIgHgZTZE3FfJ1VWbeg1AOT5A3l4a71b8zLAny`
          && `2J8c2O3yqn99hTxtB91BbzBAnqw+GVefYB7CVBnRe9ngQcM4s7FK72caiFLn5JTpkTFyowca4bOgfDt/zFMKZZKdMtpjWLEIwnkuYZ37zSnEKsGaKO/utxxdPm51nj7Q/vMdlxqYQFbXwUra4bbcWU/i7MkDsPZuAZjTTUh6g/51kZFIHmXJ4JrDnU7BbZrCX2sG4Ri9`
          && `BKmE7QlLSAX3B9wGwA/q3zV0bbq3iS9bIK3RngNkBbYVKEd4AdHJU4tkLCCVD6kMGVCVoq8kQP38ylD+5x8qSSxxcoKuW8mlj+C5LlAGv294fNp3soIJuE3MgaxN8n++txTXZpGY80+4gXgOdac2adEv3jj2kAC2X1Y20dMeagHr0znLaxVrRFKkzcH0kOWONaY59KEo`
          && `SyhuclwWXNYlx3i4PvBWsYSLt5BH9XnjO+8I7I4lq21rIsZhHSzhAiGOyCQbUHXo3JzuUtCYUiw3P4G6DIMs9gl//0DqnGhIMk5SB7SUmMrrcasn73gOm+s0N28sTToIl3hO6Uj8r6LHjlWwXUjOxZLgiMjfgTlA+wjMbAVNMiE8OpMGT0s5lS11TJnQJ9ceMHxTbNxa`
          && `qhhcPp5Um6OqR17EUkRSok691ba5UbWiZzb8TWn7obQVNPU5taz/MI3PsW0lVBrXpq1gZWOsiMmc8568OWFL9bUEajdP91XkLPr+h68Od286fVDc9qvfVGNLoffyiNU0+2ZALdSj0HkV2e03W+gcwEr4PkDHcW/fQLSEyAZdZ+mpVKlaG4ZlSfFXSk1v+a6h8b67BBQ+`
          && `dZmR8w96rppz6F5/mIvnjqWKbO+GIpYfW939BwAA//8DAFBLAwQUAAYACAAAACEAHBSnqP0AAABuAwAAFQAAAHhsL2NoYXJ0cy9jb2xvcnMxLnhtbKSTQW7CMBBFrxL5AHESIK0iwqbrigUnGE1sYsn2INulcHucUGhDIRLBO/vrv/l/JC/RV0ia3CYctUgORtv44GvW`
          && `hrCrOPfYCgM+NQodeZIhRTKcpFQoeOPgW9ktL7K84NiCCz2F/WDgH4V2wsYRkpyB4FNy2wvD6EjJSm5AWZYYEdqa4RG1YIlqapZnbLWEqk8jPrRL9qBrBojChpzxh1oxos1GtPmIthjRyk6LK92DUxAU2ZtrF1R/mU9qzg3KLJ7Ow/+ahohbz/vF06PWUp5RcX3TUU+M`
          && `v0Yejp9PGL+Y4Hm73372CuqJ9tfIw/aPmnTk3++1OgEAAP//AwBQSwMEFAAGAAgAAAAhAKV9JRAaAQAA/wIAACMAAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0MS54bWwucmVsc6ySwU7DMAyG70i8Q5Q7STsQQmjpLgNpBy4wHiCkbhutcaIkrNvb4wlNrNMmLj3a`
          && `Tn5/9u/5Yud6toWYrEfFS1FwBmh8bbFV/HP9evfEWcoaa917BMX3kPiiur2Zv0OvM31KnQ2JkQomxbucw7OUyXTgdBI+AFKl8dHpTGFsZdBmo1uQs6J4lPFUg1cjTbaqFY+r+p6z9T5Q5/+1fdNYA0tvvh1gvtBCbl2/jHqg4UhVxxay4kLI+jeXTuqloLdcXmaaTckU`
          && `osUM8QNyJq40AjurybO4FF8Wr0GWU0J2ZEHsLW7+8A5+JDJ7GAaRdBDGO3msvvmaHHvZ0Vyor67xYUpCan9wfby/Y7IUdIaHRcnR2VY/AAAA//8DAFBLAwQUAAYACAAAACEAqJz1ALwAAAAlAQAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQyLnhtbC5yZWxz`
          && `hI/BCsIwEETvgv8Q9m7SehCRpr2I0KvoB6zptg22SchG0b834EVB8DTsDvtmp2oe8yTuFNl6p6GUBQhyxnfWDRrOp8NqC4ITug4n70jDkxiaermojjRhykc82sAiUxxrGFMKO6XYjDQjSx/IZaf3ccaUxziogOaKA6l1UWxU/GRA/cUUbachtl0J4vQMOfk/2/e9NbT3`
          && `5jaTSz8iVMLLRBmIcaCkQcr3ht9SyvwsqLpSX+XqFwAAAP//AwBQSwMEFAAGAAgAAAAhABWvqjy9AAAAKwEAACMAAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0My54bWwucmVsc4SPzQrCMBCE74LvEPZu0noQkaZeRPAq9QGWZPuDbRKy8advby6CguBtZ5f9Zqba`
          && `P6dR3Cny4J2GUhYgyBlvB9dpuDTH1RYEJ3QWR+9Iw0wM+3q5qM40YspP3A+BRaY41tCnFHZKselpQpY+kMuX1scJU5axUwHNFTtS66LYqPjJgPqLKU5WQzzZEkQzh+z8n+3bdjB08OY2kUs/LJSN+MjNMhJjR0mDlO8dv4dS5sig6kp9VaxfAAAA//8DAFBLAwQUAAYA`
          && `CAAAACEADkT037wAAAAlAQAAIwAAAHhsL2RyYXdpbmdzL19yZWxzL2RyYXdpbmcxLnhtbC5yZWxzhI/NCsIwEITvgu8Q9m7SehCRpr2I0KvUB1jS7Q+2SchGsW9voBcFwdOwO+w3O0X1mifxpMCjsxpymYEga1w72l7DrbnsjiA4om1xcpY0LMRQldtNcaUJYzriYfQs`
          && `EsWyhiFGf1KKzUAzsnSebHI6F2aMaQy98mju2JPaZ9lBhU8GlF9MUbcaQt3mIJrFp+T/bNd1o6GzM4+ZbPwRocyAISYghp6iBinXDa+Sy/QsqLJQX+XKNwAAAP//AwBQSwMEFAAGAAgAAAAhAIAb3IHSAAAAjwEAAB8AAAB4bC9jaGFydHMvX3JlbHMvY2hhcnQxLnht`
          && `bC5yZWxzrJDBSsQwEIbvgu8Q5m7T7EFENt2DIOxV1wcI6bQNm2TCTBD79qYXsYtHjzPD/30/czx9pag+kSVQtmC6HhRmT2PIs4WPy+vDEyipLo8uUkYLKwqchvu74xtGV1tIllBENUoWC0ut5Vlr8QsmJx0VzO0yESdX28izLs5f3Yz60PePmn8zYNgx1Xm0wOfxAOqy`
          && `lma+YafgmYSm2nlKmqYp+I1qzJ6q/eK4vlAkfq9rxEZzPGO14LedmK4VBP232/yH+0YrW4sfq969cfgGAAD//wMAUEsDBBQABgAIAAAAIQDHjixBQwEAAIECAAAiAAAAeGwvZXh0ZXJuYWxMaW5rcy9leHRlcm5hbExpbmsxLnhtbJyS32rCMBSH7wd7h5B7TVvGmMVW`
          && `EDcQxm62PUBMT9tgklNOotS3X1qdc3gj3uXkz3d++TjzRW8N2wN5ja7g6TThDJzCSrum4N9fb5MXznyQrpIGHRT8AJ4vyseHOfQByEnzrt2WRYjzBW9D6HIhvGrBSj/FDlw8qZGsDLGkRviOQFa+BQjWiCxJnoWV2vEjIbfqFoiVtN11E4W2k0FvtNHhMLI4sypfNw5J`
          && `bkwM26dPrO9Nlv7y48ZVA6sVocc6TCNQYF1rBdc5Z2L2L+mIvY+VpSLKM9Gbz5KYrTy7XCKeXOZ0i4hj2BWqnQUXjjoJTJSCzre685xRrquC07oa+ozeP6QFf7Fme2kKvpSu4aKci6s7KxnkJ4TTi6Fi45115MZpIagJfPtKhDFz+sc4vxt+O47K8L3Y4XJyyh8AAAD/`
          && `/wMAUEsDBBQABgAIAAAAIQCy3uc6zwIAALgNAAAnAAAAeGwvcHJpbnRlclNldHRpbmdzL3ByaW50ZXJTZXR0aW5nczEuYmlu7Fa9bhNBEP5sBxQjS1BQUEaIwhT+C07AEU2wDTKy45PtRC6x4ktysrk73Tk2P+INeABqGkooKPwIeQAKilSUiHcI3+zdEf+JJMIRImRP`
          && `uzs7Nzs78+3M3FVhQscGLPQ4LyGOAmcXHa4t2LhLXoJdg+PLbCvJNrwWWohcPcRhLJIAwgjh+zVrsc35OpphWTfDEY5l6uypfY6/70+mUHA25zC7rI/YnpTq2VG9hdLG5m0c4E5kGF37+O3D7868ol5Gla45mHip4h9EIIir05h+QKF6pfFUZG9giNfII4PHWMEqMyiH`
          && `LLNmneMqeQnOOSyjqHgZUjlS95V0ltIZvimSe0/tfkONtWK9UC5j0zQc3RVKa9m6Uzde6SgXG41iDVXH0M1eq2dYJrRqrVFbLzVQ012ru694JKu2UBnkra7lVKy27lHj3g2jwFa2UAl839K00tH72Jc4xW4yKZa8xMCPBWCRvBj7V3aTNSLtP6Ma91Se21hDio+Lbewx`
          && `95+jRTrJ2SBHKonLvkPZJNcW+SkM+M5Em6sB36aIR5rIrJBKk06xFjlKoqfG3THO6DkdnveSOizKt5Wm6VOPdXlVrk1bgM/sNncIPdnm61eady1+PTiDXzv0p0WkdN+3jjLx2Oblc7Y5o2yW6D39XUzbLFZ+8nH2PBg3++/jPBk/YHZ7cdFhPE6iPO94l7g4G8aT9vZ/`
          && `WblHu18wYronxvOAMTVgLg54x0mVOZJfTVT45RZ7cnwkVk0/Al3mSYuZqyvNz9Q5LrPtpLyZfY5kuuR6cGJ9pGqIRtvXb5yD/oSqJi4rUYveeR5JjZMm8zvWPamNQr8lLf8bUiM6rHAa9+wSgyr2ud9WY57odRWCM4y94KwAl/8Zg1lX/FCChk3iKM4hwEn+tPvqi6jj`
          && `Eece40ojz+FoMHP1mZl7cYPolu/aNE7y5e8rZC5xAgSnnwAAAP//AwBQSwMEFAAGAAgAAAAhAHI+G1KeAQAAgQIAABAAAAB4bC9jb21tZW50czEueG1sjJBNbtswEIX3BXIHgnuZkmIniiEpaH4MGC2CLNJuC5oaiUT5I5CUIafoXZprZOuT9CYdR3aBIpsCBDHzZvDN`
          && `wyuvR6PJFnxQzlY0m6WUgBWuUbar6JenVVJQEiK3DdfOQkV3EOh1ffahFM4YsDEQBNhQURljv2QsCAmGh5nrweKkdd7wiK3vWOg98CZIgGg0y9P0ghmuLJ0ISyP+B2K4/z70CV7veVQbpVXcvbEoMWK57qzzfKPR6OhP4NG/AxslvAuujTMEMde2SsA7f9mcediqQzK0`
          && `LvkQpfPhVNS3u84O4dsDbAbNSzaN61OBe8eAPqsQ/zbEQ1vRmzkl0/66qSgmHiTv4ViPfjkolH8sVlc3aZpmyWWxOE/mxeoiKbIsT/JFXszvby8Xd/nHn2gswogHPL5H/DasLsMz2XJd0SuKjXDaeaJsAyMgtsgOol85G6elJy6d4Qex5Ubp3aTmB4G9EWP9Se5fhCRy`
          && `/8t25Ov6kSSkk4oIuX8ljSP3owBN4u/XF1cytMLQBptcsWMIKPwTx6kL9R8AAAD//wMAUEsDBBQABgAIAAAAIQDURqooLgIAABAEAAAUAAAAeGwvdGFibGVzL3RhYmxlMS54bWyck01u2zAQhfcFegeBe0Y/lGTJiBxQPywCJF00bve0TMdERVIgqcRG0RPkFEV3vUKX`
          && `OUlvUspOlLTpou1OeqK+N284c3q2E513w7ThShYgPAmAx2Sr1lxeF+D9ksAMeMZSuaadkqwAe2bA2eL1q1NLVx3z3N/SFGBrbT/3fdNumaDmRPVMui8bpQW17lVf+6bXjK7NljErOj8KgtQXlEtwJMxF+zcQQfXHoYetEj21fMU7bvcHFvBEOz+/lkqPVRVgp72dRo/w`
          && `nX4BF7zVyqiNPXEwX202vGUvagxjX7MbPrbmCYX+k5VOLFcXX7teO6aeD+Pjp6rBJG+aCIYoD2BMqhBmJU5gPEsrFDd1gIP6M/AkFS6cXXUfqF0OwFtz03d0//Y3WbNNAXA4rxLgWWVpZ96p20oN0o6ui1M6WEV4Z5n2pqPxP1fjL44zUKluENJ47dEAOYPDbBz1p6jo`
          && `IWuYIjJLcAgrlBAYZ2kAS4IiiPNohlFaYkzKKevl/ddnGS7oinUFWI6ZwK/+B59oDPHok2KESDYrIQrrEsZ1GMEywQjWeU2iJMryoCGTz/L+m/wTcZyhiThDddXEQQWzCEcwzrMZzOsGwzRKCUZhE1dJMBHf8PsvntU/vt89C0AG2drDpplBjH7+s1aZh8Zd2X3HzuVG`
          && `Pdz3cpzpg3jJ1nwQudvIrbolXBt7bHIB3NqO2gV9Ibmrv7Ka98ytqRu58dTxp0kNngpZ/AQAAP//AwBQSwMEFAAGAAgAAAAhAAVyt4XSAAAAoAEAABAAAAB4bC9jYWxjQ2hhaW4ueG1sXNBLbsMgEAbgfaXeAc0+xk6aNKmMs4hqLpAeAOFJsMTDAlS1t++ojROXJR8/`
          && `86A9fjnLPjGmMXgBTVUDQ6/DMPqrgI9zv9oDS1n5QdngUcA3Jjh2z0+tVlafjBo9owo+CTA5T2+cJ23QqVSFCT3dXEJ0KtMxXnmaIqohGcTsLF/X9Y47KgBdq1kUcNoCGwWsgVmaBPiN35vNry+FMpRcSlNIfyhA0iL/3vQlyNcyUYLc3RKM9n1072eeR5Z/mywSJciX`
          && `stcC+P1vux8AAAD//wMAUEsDBBQABgAIAAAAIQBogQelRAEAAGsCAAARAAgBZG9jUHJvcHMvY29yZS54bWwgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUklFPgzAUhd9N/A+k71BgcToCLHFmT86YuEXji6ntHWukpWk7Gf/eAhui88XH9pz79ZybpvODKL1P0IZXMkNRECIPJK0Yl0WGNuulf4M8Y4lkpKwkZKgBg+b55UVKVUIrDY+6UqAtB+M5kjQJVRnaWasSjA3dgSAmcA7p`
          && `xG2lBbHuqAusCP0gBeA4DKdYgCWMWIJboK8GIjoiGR2Qaq/LDsAohhIESGtwFET422tBC/PnQKeMnILbRrlOx7hjNqO9OLgPhg/Guq6DetLFcPkj/LK6f+qq+ly2u6KA8pTRhGogttL5oink3rw9wPu+JCkeKe0WS2Lsyi18y4HdNr/N5wZH7or0eGCei5b0RU7K82Rx`
          && `t16iPA7jqR/O/Hi2DuMkmiTx1Wv7/o/5Nmp/IY4p/kW8HhFPgDzFZ98j/wIAAP//AwBQSwMEFAAGAAgAAAAhADtkFJqtAQAAxQMAABAACAFkb2NQcm9wcy9hcHAueG1sIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`
          && `AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAnFPBbtswDL0P2D8Yujdy1qIYAlnFFm/ooduCxW3PnEzHQm3JkBgj2dePjtHU2YIN3Y3kox8fnyl1s2ubpMcQrXeZmM9SkaAzvrRuk4n74vPFe5FEAldC4x1m`
          && `Yo9R3Oi3b9Qq+A4DWYwJU7iYiZqoW0gZTY0txBnDjpHKhxaI07CRvqqswdybbYuO5Ls0vZa4I3QllhfdkVCMjIue/pe09GbQFx+KfceCtfrQdY01QLyl/mJN8NFXlHzaGWyUnIKK1a3RbIOlvU6VnKZqbaDBJRPrCpqISr4U1C3CYNoKbIha9bTo0ZAPSbQ/2bYrkfyA`
          && `iIOcTPQQLDhiWUPbmBziposU9KMPT7FGpKgkN4zFQzjtncb2Sl8eGjj4a+PI9RVaLJPv4Db4mhHz8yMGjeOuPPvUhcJSg/FbtYJA/zLlIG20ZFSZb+8sbqcWHM34yNLPA/xF7s9CuQdXF/UJ4Yn038QufduB2/NOx+jOuqd43xU+B8LnEzgtqnUNAUu+muOJHAvqlv9+`
          && `aAaSZT2YXz73/AkMB/swvko9v56llynf4qSm5Mv7078AAAD//wMAUEsDBBQABgAIAAAAIQDgX3c80wAAADkBAAAtAAAAeGwvZXh0ZXJuYWxMaW5rcy9fcmVscy9leHRlcm5hbExpbmsxLnhtbC5yZWxzbI9BSwMxEIXvgv8hzN3NtgcR2WxBsFCwIFJ/wJDM7oYmmZAJ`
          && `sv33xoNoxdPw5vG+xxt2awzqg4p4TgY2XQ+KkmXn02zg/bS/ewAlFZPDwIkMXEhgN97eDG8UsLaQLD6LapQkBpZa86PWYheKKB1nSs2ZuESsTZZZZ7RnnElv+/5el98MGK+Y6uAMlIPbgDpdcmv+w47eFhaeamc5ap4mb/+j6jU8r5VKwvDi0/kV69JeX+foRdrKhscy`
          && `UzXwhD/qyK5VfidBj4O+Gjx+AgAA//8DAFBLAQItABQABgAIAAAAIQA3FDkb/QEAAEkKAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhALVVMCP0AAAATAIAAAsAAAAAAAAAAAAAAAAANgQAAF9yZWxzLy5yZWxzUEsB`
          && `Ai0AFAAGAAgAAAAhAO6eZ1GDBAAAOwsAAA8AAAAAAAAAAAAAAAAAWwcAAHhsL3dvcmtib29rLnhtbFBLAQItABQABgAIAAAAIQBqqnpUJwEAAPMEAAAaAAAAAAAAAAAAAAAAAAsMAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc1BLAQItABQABgAIAAAAIQByYul4`
          && `uAcAAJ8WAAAYAAAAAAAAAAAAAAAAAHIOAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwECLQAUAAYACAAAACEAtn42WacCAACLBgAAGAAAAAAAAAAAAAAAAABgFgAAeGwvd29ya3NoZWV0cy9zaGVldDIueG1sUEsBAi0AFAAGAAgAAAAhACtlAMzEAQAAqwMAABgA`
          && `AAAAAAAAAAAAAAAAPRkAAHhsL3dvcmtzaGVldHMvc2hlZXQzLnhtbFBLAQItABQABgAIAAAAIQD2YLRBuAcAABEiAAATAAAAAAAAAAAAAAAAADcbAAB4bC90aGVtZS90aGVtZTEueG1sUEsBAi0AFAAGAAgAAAAhAAwdTURqBAAAvg8AAA0AAAAAAAAAAAAAAAAAICMA`
          && `AHhsL3N0eWxlcy54bWxQSwECLQAUAAYACAAAACEAbmbZc2YCAADcBQAAFAAAAAAAAAAAAAAAAAC1JwAAeGwvc2hhcmVkU3RyaW5ncy54bWxQSwECLQAUAAYACAAAACEAYTxMcmkCAACaBAAAGwAAAAAAAAAAAAAAAABNKgAAeGwvZHJhd2luZ3Mvdm1sRHJhd2luZzEu`
          && `dm1sUEsBAi0AFAAGAAgAAAAhAKJTjVYkAgAAzgQAABgAAAAAAAAAAAAAAAAA7ywAAHhsL2RyYXdpbmdzL2RyYXdpbmcxLnhtbFBLAQItABQABgAIAAAAIQBCQEuVnQcAAD4aAAAUAAAAAAAAAAAAAAAAAEkvAAB4bC9jaGFydHMvY2hhcnQxLnhtbFBLAQItABQABgAI`
          && `AAAAIQDVLRgxnAQAAMIlAAAUAAAAAAAAAAAAAAAAABg3AAB4bC9jaGFydHMvc3R5bGUxLnhtbFBLAQItABQABgAIAAAAIQAcFKeo/QAAAG4DAAAVAAAAAAAAAAAAAAAAAOY7AAB4bC9jaGFydHMvY29sb3JzMS54bWxQSwECLQAUAAYACAAAACEApX0lEBoBAAD/AgAA`
          && `IwAAAAAAAAAAAAAAAAAWPQAAeGwvd29ya3NoZWV0cy9fcmVscy9zaGVldDEueG1sLnJlbHNQSwECLQAUAAYACAAAACEAqJz1ALwAAAAlAQAAIwAAAAAAAAAAAAAAAABxPgAAeGwvd29ya3NoZWV0cy9fcmVscy9zaGVldDIueG1sLnJlbHNQSwECLQAUAAYACAAAACEA`
          && `Fa+qPL0AAAArAQAAIwAAAAAAAAAAAAAAAABuPwAAeGwvd29ya3NoZWV0cy9fcmVscy9zaGVldDMueG1sLnJlbHNQSwECLQAUAAYACAAAACEADkT037wAAAAlAQAAIwAAAAAAAAAAAAAAAABsQAAAeGwvZHJhd2luZ3MvX3JlbHMvZHJhd2luZzEueG1sLnJlbHNQSwEC`
          && `LQAUAAYACAAAACEAgBvcgdIAAACPAQAAHwAAAAAAAAAAAAAAAABpQQAAeGwvY2hhcnRzL19yZWxzL2NoYXJ0MS54bWwucmVsc1BLAQItABQABgAIAAAAIQDHjixBQwEAAIECAAAiAAAAAAAAAAAAAAAAAHhCAAB4bC9leHRlcm5hbExpbmtzL2V4dGVybmFsTGluazEu`
          && `eG1sUEsBAi0AFAAGAAgAAAAhALLe5zrPAgAAuA0AACcAAAAAAAAAAAAAAAAA+0MAAHhsL3ByaW50ZXJTZXR0aW5ncy9wcmludGVyU2V0dGluZ3MxLmJpblBLAQItABQABgAIAAAAIQByPhtSngEAAIECAAAQAAAAAAAAAAAAAAAAAA9HAAB4bC9jb21tZW50czEueG1s`
          && `UEsBAi0AFAAGAAgAAAAhANRGqiguAgAAEAQAABQAAAAAAAAAAAAAAAAA20gAAHhsL3RhYmxlcy90YWJsZTEueG1sUEsBAi0AFAAGAAgAAAAhAAVyt4XSAAAAoAEAABAAAAAAAAAAAAAAAAAAO0sAAHhsL2NhbGNDaGFpbi54bWxQSwECLQAUAAYACAAAACEAaIEHpUQB`
          && `AABrAgAAEQAAAAAAAAAAAAAAAAA7TAAAZG9jUHJvcHMvY29yZS54bWxQSwECLQAUAAYACAAAACEAO2QUmq0BAADFAwAAEAAAAAAAAAAAAAAAAAC2TgAAZG9jUHJvcHMvYXBwLnhtbFBLAQItABQABgAIAAAAIQDgX3c80wAAADkBAAAtAAAAAAAAAAAAAAAAAJlRAAB4`
          && `bC9leHRlcm5hbExpbmtzL19yZWxzL2V4dGVybmFsTGluazEueG1sLnJlbHNQSwUGAAAAABwAHACxBwAAt1IAAAAA` ##NO_TEXT.
    rv_file = cl_web_http_utility=>decode_x_base64( lv_b64 ).
  ENDMETHOD.


  METHOD build_named_range_xlsx.
    " Test workbook: global / sheet-local named ranges and formulas referencing other sheets
    DATA: lo_excel  TYPE REF TO zcl_excel,
          lo_data   TYPE REF TO zcl_excel_worksheet,
          lo_price  TYPE REF TO zcl_excel_worksheet,
          lo_rep    TYPE REF TO zcl_excel_worksheet,
          lo_range  TYPE REF TO zcl_excel_range,
          lo_writer TYPE REF TO zif_excel_writer,
          lv_row    TYPE i.

    CREATE OBJECT lo_excel.
    DATA(lo_st_bold) = lo_excel->add_new_style( ).
    lo_st_bold->font->bold = abap_true.

    " ---- sheet "Data": products B = quantity, C = price
    lo_data = lo_excel->get_active_worksheet( ).
    lo_data->set_title( 'Data' ) ##NO_TEXT.
    lo_data->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Mã SP` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_data->set_cell( ip_column = 'B' ip_row = 1 ip_value = `Số lượng` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_data->set_cell( ip_column = 'C' ip_row = 1 ip_value = `Đơn giá` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    DO 5 TIMES.
      lv_row = sy-index + 1.
      lo_data->set_cell( ip_column = 'A' ip_row = lv_row ip_value = |SP{ sy-index WIDTH = 2 ALIGN = RIGHT PAD = '0' }| ) ##NO_TEXT.
      lo_data->set_cell( ip_column = 'B' ip_row = lv_row ip_value = sy-index * 10 ).
      lo_data->set_cell( ip_column = 'C' ip_row = lv_row ip_value = sy-index * 1500 ).
    ENDDO.

    " ---- sheet with blank and Vietnamese characters in the name (needs quotes in references)
    lo_price = lo_excel->add_new_worksheet( ).
    lo_price->set_title( 'Bảng giá 2026' ) ##NO_TEXT.
    lo_price->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Thuế suất` ) ##NO_TEXT.
    lo_price->set_cell( ip_column = 'B' ip_row = 1 ip_value = CONV decfloat34( '0.08' ) ).

    " ---- global named ranges
    lo_range = lo_excel->add_new_range( ).
    lo_range->name = 'SoLuong' ##NO_TEXT.
    lo_range->set_value( ip_sheet_name = 'Data' ip_start_column = 'B' ip_start_row = 2 ip_stop_column = 'B' ip_stop_row = 6 ) ##NO_TEXT.
    lo_range = lo_excel->add_new_range( ).
    lo_range->name = 'DonGia' ##NO_TEXT.
    lo_range->set_value( ip_sheet_name = 'Data' ip_start_column = 'C' ip_start_row = 2 ip_stop_column = 'C' ip_stop_row = 6 ) ##NO_TEXT.
    lo_range = lo_excel->add_new_range( ).
    lo_range->name = 'ThueSuat' ##NO_TEXT.
    lo_range->set_value( ip_sheet_name = 'Bảng giá 2026' ip_start_column = 'B' ip_start_row = 1 ip_stop_column = 'B' ip_stop_row = 1 ) ##NO_TEXT.
    " ---- sheet-local named range (only visible in sheet "Data")
    lo_range = lo_data->add_new_range( ).
    lo_range->name = 'MaSP' ##NO_TEXT.
    lo_range->set_value( ip_sheet_name = 'Data' ip_start_column = 'A' ip_start_row = 2 ip_stop_column = 'A' ip_stop_row = 6 ) ##NO_TEXT.
    lo_data->set_cell( ip_column = 'E' ip_row = 1 ip_value = `Số mã (tên cục bộ MaSP)` ) ##NO_TEXT.
    lo_data->set_cell( ip_column = 'E' ip_row = 2 ip_formula = 'COUNTA(MaSP)' ) ##NO_TEXT.

    " ---- sheet "Report": formulas referencing the other sheets and the names
    lo_rep = lo_excel->add_new_worksheet( ).
    lo_rep->set_title( 'Report' ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Tổng số lượng =SUM(SoLuong)` ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'B' ip_row = 1 ip_formula = 'SUM(SoLuong)' ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'A' ip_row = 2 ip_value = `Doanh thu =SUMPRODUCT(SoLuong,DonGia)` ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'B' ip_row = 2 ip_formula = 'SUMPRODUCT(SoLuong,DonGia)' ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'A' ip_row = 3 ip_value = `Thuế =B2*ThueSuat` ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'B' ip_row = 3 ip_formula = 'B2*ThueSuat' ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'A' ip_row = 4 ip_value = `SP02 =Data!B3*Data!C3` ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'B' ip_row = 4 ip_formula = 'Data!B3*Data!C3' ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'A' ip_row = 5 ip_value = `Giá SP04 =VLOOKUP("SP04",Data!$A$2:$C$6,3,FALSE)` ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'B' ip_row = 5 ip_formula = 'VLOOKUP("SP04",Data!$A$2:$C$6,3,FALSE)' ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'A' ip_row = 6 ip_value = `Thuế suất ='Bảng giá 2026'!B1` ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'B' ip_row = 6 ip_formula = |'Bảng giá 2026'!B1| ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'A' ip_row = 7 ip_value = `Có dữ liệu? =IF(COUNT(Data!B:B)>0,"Có","Không")` ) ##NO_TEXT.
    lo_rep->set_cell( ip_column = 'B' ip_row = 7 ip_formula = 'IF(COUNT(Data!B:B)>0,"Có","Không")' ) ##NO_TEXT.
    " formulas copied down with relative references (shared formula handling of the writer is not involved)
    DO 3 TIMES.
      lv_row = 8 + sy-index.
      lo_rep->set_cell( ip_column = 'A' ip_row = lv_row ip_formula = |Data!A{ sy-index + 1 }| ) ##NO_TEXT.
      lo_rep->set_cell( ip_column = 'B' ip_row = lv_row ip_formula = |Data!B{ sy-index + 1 }*Data!C{ sy-index + 1 }| ) ##NO_TEXT.
    ENDDO.
    lo_rep->get_column( 'A' )->set_width( 48 ).
    lo_rep->get_column( 'B' )->set_width( 14 ).

    lo_excel->set_active_sheet_index( 3 ).
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD build_advanced_style_xlsx.
    " Test workbook: gradient fills, text rotation, indent, shrink to fit, justify / vertical alignment
    TYPES: BEGIN OF ty_s_grad,
             type TYPE zexcel_fill_type,
             text TYPE string,
           END OF ty_s_grad.
    DATA: lo_excel  TYPE REF TO zcl_excel,
          lo_sheet  TYPE REF TO zcl_excel_worksheet,
          lo_writer TYPE REF TO zif_excel_writer,
          lt_grad   TYPE STANDARD TABLE OF ty_s_grad WITH EMPTY KEY,
          lv_row    TYPE i.

    CREATE OBJECT lo_excel.
    lo_sheet = lo_excel->get_active_worksheet( ).
    lo_sheet->set_title( 'Styles' ) ##NO_TEXT.

    " ---- column A: gradient fills (rows 2 - 9)
    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Gradient fill` ) ##NO_TEXT.
    lt_grad = VALUE #( ( type = zcl_excel_style_fill=>c_fill_gradient_horizontal90  text = `horizontal90` )
                       ( type = zcl_excel_style_fill=>c_fill_gradient_horizontal270 text = `horizontal270` )
                       ( type = zcl_excel_style_fill=>c_fill_gradient_horizontalb   text = `horizontalb` )
                       ( type = zcl_excel_style_fill=>c_fill_gradient_vertical      text = `vertical` )
                       ( type = zcl_excel_style_fill=>c_fill_gradient_diagonal45    text = `diagonal45` )
                       ( type = zcl_excel_style_fill=>c_fill_gradient_diagonal135b  text = `diagonal135b` )
                       ( type = zcl_excel_style_fill=>c_fill_gradient_fromcenter    text = `fromCenter` )
                       ( type = zcl_excel_style_fill=>c_fill_gradient_cornerrb      text = `cornerRB` ) ) ##NO_TEXT.
    lv_row = 1.
    LOOP AT lt_grad INTO DATA(ls_grad).
      lv_row = lv_row + 1.
      DATA(lo_st_grad) = lo_excel->add_new_style( ).
      lo_st_grad->fill->filltype = ls_grad-type.
      lo_st_grad->fill->fgcolor-rgb = 'FF1F4E78' ##NO_TEXT.
      lo_st_grad->fill->bgcolor-rgb = 'FFDDEBF7' ##NO_TEXT.
      lo_st_grad->font->bold = abap_true.
      lo_st_grad->font->color-rgb = 'FFFFFFFF' ##NO_TEXT.
      lo_st_grad->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
      lo_st_grad->alignment->vertical   = zcl_excel_style_alignment=>c_vertical_center.
      lo_sheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_grad-text ip_style = lo_st_grad->get_guid( ) ).
      lo_sheet->set_row_height( ip_row = lv_row ip_height_fix = 30 ).
    ENDLOOP.
    lo_sheet->get_column( 'A' )->set_width( 22 ).

    " ---- row 12: text rotation 0 / 45 / 90 / 135 (= -45) / 180 (= -90) / 255 (vertical stacked)
    lo_sheet->set_cell( ip_column = 'A' ip_row = 11 ip_value = `Xoay chữ (textRotation)` ) ##NO_TEXT.
    DATA lt_rot TYPE STANDARD TABLE OF i WITH EMPTY KEY.
    lt_rot = VALUE #( ( 0 ) ( 45 ) ( 90 ) ( 135 ) ( 180 ) ( 255 ) ).
    DATA(lv_col) = 1.
    LOOP AT lt_rot INTO DATA(lv_rot).
      DATA(lo_st_rot) = lo_excel->add_new_style( ).
      lo_st_rot->alignment->textrotation = lv_rot.
      lo_st_rot->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
      lo_st_rot->borders->allborders = NEW zcl_excel_style_border( ).
      lo_st_rot->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
      lo_sheet->set_cell( ip_column = lv_col ip_row = 12 ip_value = |Xoay { lv_rot }| ip_style = lo_st_rot->get_guid( ) ) ##NO_TEXT.
      lv_col = lv_col + 1.
    ENDLOOP.
    lo_sheet->set_row_height( ip_row = 12 ip_height_fix = 70 ).

    " ---- rows 14 - 17: indent 0..3 (left) and 2 (right)
    lo_sheet->set_cell( ip_column = 'A' ip_row = 14 ip_value = `Thụt lề (indent)` ) ##NO_TEXT.
    DO 4 TIMES.
      DATA(lo_st_ind) = lo_excel->add_new_style( ).
      lo_st_ind->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_left.
      lo_st_ind->alignment->indent = sy-index - 1.
      lo_sheet->set_cell( ip_column = 'A' ip_row = 14 + sy-index ip_value = |Cấp { sy-index - 1 }| ip_style = lo_st_ind->get_guid( ) ) ##NO_TEXT.
    ENDDO.
    DATA(lo_st_rind) = lo_excel->add_new_style( ).
    lo_st_rind->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_right.
    lo_st_rind->alignment->indent = 2.
    lo_sheet->set_cell( ip_column = 'B' ip_row = 15 ip_value = CONV decfloat34( '12345.5' ) ip_style = lo_st_rind->get_guid( ) ).

    " ---- row 20: shrink to fit, justify, vertical top, center across selection
    DATA(lo_st_shrink) = lo_excel->add_new_style( ).
    lo_st_shrink->alignment->shrinktofit = abap_true.
    lo_sheet->set_cell( ip_column = 'B' ip_row = 20 ip_value = `Chữ rất dài được thu nhỏ vừa ô (shrinkToFit)` ip_style = lo_st_shrink->get_guid( ) ) ##NO_TEXT.
    DATA(lo_st_just) = lo_excel->add_new_style( ).
    lo_st_just->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_justify.
    lo_st_just->alignment->vertical   = zcl_excel_style_alignment=>c_vertical_top.
    lo_st_just->alignment->wraptext   = abap_true.
    lo_sheet->set_cell( ip_column = 'C' ip_row = 20 ip_value = `Đoạn văn căn đều hai bên, canh trên, tự xuống dòng trong ô.` ip_style = lo_st_just->get_guid( ) ) ##NO_TEXT.
    DATA(lo_st_cc) = lo_excel->add_new_style( ).
    lo_st_cc->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center_continuous.
    lo_sheet->set_cell( ip_column = 'D' ip_row = 20 ip_value = `Center across D20:F20` ip_style = lo_st_cc->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'E' ip_row = 20 ip_value = `` ip_style = lo_st_cc->get_guid( ) ).
    lo_sheet->set_cell( ip_column = 'F' ip_row = 20 ip_value = `` ip_style = lo_st_cc->get_guid( ) ).
    lo_sheet->set_row_height( ip_row = 20 ip_height_fix = 60 ).
    lo_sheet->get_column( 'B' )->set_width( 16 ).
    lo_sheet->get_column( 'C' )->set_width( 18 ).

    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD build_freeze_width_xlsx.
    " Test workbook: freeze panes (rows + columns / rows only / columns only), column widths
    " (fixed, auto size, hidden, default width) and row heights
    DATA: lo_excel  TYPE REF TO zcl_excel,
          lo_sheet  TYPE REF TO zcl_excel_worksheet,
          lo_writer TYPE REF TO zif_excel_writer,
          lv_row    TYPE i,
          lv_col    TYPE i.

    CREATE OBJECT lo_excel.
    DATA(lo_st_head) = lo_excel->add_new_style( ).
    lo_st_head->font->bold = abap_true.
    lo_st_head->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_st_head->fill->fgcolor-rgb = 'FFD9E1F2' ##NO_TEXT.

    " ---- 3 sheets with a 40 x 12 grid and different freeze settings
    DO 3 TIMES.
      DATA(lv_sheet_no) = sy-index.
      IF lv_sheet_no = 1.
        lo_sheet = lo_excel->get_active_worksheet( ).
      ELSE.
        lo_sheet = lo_excel->add_new_worksheet( ).
      ENDIF.
      CASE lv_sheet_no.
        WHEN 1.
          lo_sheet->set_title( 'Freeze_B2' ) ##NO_TEXT.
          " 1 row + 1 column frozen, scrolled pane starts at D10
          lo_sheet->freeze_panes( ip_num_columns = 1 ip_num_rows = 1 ).
          lo_sheet->set_pane_top_left_cell( 'D10' ) ##NO_TEXT.
        WHEN 2.
          lo_sheet->set_title( 'Freeze_Rows3' ) ##NO_TEXT.
          lo_sheet->freeze_panes( ip_num_rows = 3 ).
        WHEN 3.
          lo_sheet->set_title( 'Freeze_ColsB' ) ##NO_TEXT.
          lo_sheet->freeze_panes( ip_num_columns = 2 ).
      ENDCASE.
      lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Mã \\ Tháng` ip_style = lo_st_head->get_guid( ) ) ##NO_TEXT.
      DO 12 TIMES.
        lv_col = sy-index + 1.
        lo_sheet->set_cell( ip_column = lv_col ip_row = 1 ip_value = |T{ sy-index WIDTH = 2 ALIGN = RIGHT PAD = '0' }| ip_style = lo_st_head->get_guid( ) ) ##NO_TEXT.
      ENDDO.
      DO 40 TIMES.
        lv_row = sy-index + 1.
        lo_sheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = |KH{ sy-index WIDTH = 4 ALIGN = RIGHT PAD = '0' }| ip_style = lo_st_head->get_guid( ) ) ##NO_TEXT.
        DO 12 TIMES.
          lv_col = sy-index + 1.
          lo_sheet->set_cell( ip_column = lv_col ip_row = lv_row ip_value = lv_row * 100 + sy-index ).
        ENDDO.
      ENDDO.
    ENDDO.

    " ---- sheet "Widths"
    lo_sheet = lo_excel->add_new_worksheet( ).
    lo_sheet->set_title( 'Widths' ) ##NO_TEXT.
    DATA(lo_st_bold) = lo_excel->add_new_style( ).
    lo_st_bold->font->bold = abap_true.
    DATA(lo_st_big) = lo_excel->add_new_style( ).
    lo_st_big->font->size = 16.
    DATA(lo_st_date) = lo_excel->add_new_style( ).
    lo_st_date->number_format->format_code = zcl_excel_style_number_format=>c_format_date_std.
    DATA(lo_st_wrap) = lo_excel->add_new_style( ).
    lo_st_wrap->alignment->wraptext = abap_true.

    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `A=8` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = `B=40 (set_column_width)` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = `C auto` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = `D auto (đậm)` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = `E auto số/ngày` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'F' ip_row = 1 ip_value = `F ẩn` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'G' ip_row = 1 ip_value = `G auto 16pt` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'H' ip_row = 1 ip_value = `H auto wrap` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'I' ip_row = 1 ip_value = `I mặc định 15` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.

    lo_sheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = `Chữ dài bị cắt trong cột hẹp` ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = `Cột cố định 40` ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = `ngắn` ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = `Dòng dài nhất của cột C quyết định độ rộng` ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'D' ip_row = 2 ip_value = `CÔNG TY TNHH ABAP CLOUD` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'E' ip_row = 2 ip_value = CONV decfloat34( '1234567890.55' ) ).
    lo_sheet->set_cell( ip_column = 'E' ip_row = 3 ip_value = CONV d( '20260928' ) ip_style = lo_st_date->get_guid( ) ).
    lo_sheet->set_cell( ip_column = 'F' ip_row = 2 ip_value = `Cột này bị ẩn` ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'G' ip_row = 2 ip_value = `Chữ cỡ 16` ip_style = lo_st_big->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'H' ip_row = 2 ip_value = |Dòng 1{ cl_abap_char_utilities=>newline }Dòng thứ hai dài hơn| ip_style = lo_st_wrap->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'I' ip_row = 2 ip_value = `Rộng mặc định` ) ##NO_TEXT.

    lo_sheet->get_column( 'A' )->set_width( 8 ).
    lo_sheet->set_column_width( ip_column = 'B' ip_width_fix = 40 ).
    lo_sheet->get_column( 'C' )->set_auto_size( abap_true ).
    lo_sheet->get_column( 'D' )->set_auto_size( abap_true ).
    lo_sheet->get_column( 'E' )->set_auto_size( abap_true ).
    lo_sheet->get_column( 'F' )->set_visible( abap_false ).
    lo_sheet->get_column( 'G' )->set_auto_size( abap_true ).
    lo_sheet->get_column( 'H' )->set_auto_size( abap_true ).
    lo_sheet->get_default_column( )->set_width( 15 ).

    " row heights: row 1 = 30 pt, row 2 = 45 pt (wrap), row 4 hidden
    lo_sheet->set_row_height( ip_row = 1 ip_height_fix = 30 ).
    lo_sheet->get_row( 2 )->set_row_height( 45 ).
    lo_sheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = `Dòng 4 bị ẩn` ) ##NO_TEXT.
    lo_sheet->get_row( 4 )->set_visible( abap_false ).
    lo_sheet->set_cell( ip_column = 'A' ip_row = 5 ip_value = `Dòng 5 (dòng 4 ẩn phía trên)` ) ##NO_TEXT.

    lo_excel->set_active_sheet_index( 1 ).
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD get_test_image.
    " test images generated with System.Drawing: logo.png 240x80 (transparent), photo.jpg 320x200
    DATA lv_b64 TYPE string.
    CASE iv_kind.
      WHEN 'PNG'.
        lv_b64 = `iVBORw0KGgoAAAANSUhEUgAAAPAAAABQCAYAAAAnSfh8AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAx/SURBVHhe7Z09bNtIFoBTbnHABtjIkizRSrUIDshecNsscPauq3OuWUu29ePYe3L+HNkSLdn6W6S5dAEkIWkW`
              && `SGMgZbq4TJnChUuXBtK4dOnSpQ5vyJFGj4+URP1Z1PuAByQiORyO5+MbDknpzh2GYRiGYRiGYRiGYSaCP52/H9oqLotIt8OfzosIpo8e4G0YhpkAkXgpMP/kaCe0VTwJbh+dzW8fNUX8YcZ/2xFMHxqxY0Rg5/A88PTwJJAuZEB6XDbDMCMguHn0IJQ6LMxvFs9CT4rN`
              && `0JYRLXl7FFjEU4hCM/BMxHngxeHb+Rf5R3ifDMMMQCSdvxtKFt+EUqWL0GaxKcSVMTyBm4HnZrzIX/mf5z/40vsBXBeGYXokkv7fd6FUqRBKla7DqVIztAkxFoFF+F/mbwIv82+/T+fv4roxDOOAlizthBOlSxBXxrgFDrw0wr97cD23qxci6fR3uJ4MwyiEN0qPw4ny`
              && `eThZboaTbXknK7CQuDm3q1/6Xx6kcJ0Zhrlz504oUSqEE+WmiFsosP+VDP0DrjvDzCxwrRveKH9syXvrBT5ozmUOvn6f52tjZsaJrObvahvlMy2uyDsFAvszILF++cNujh8MYWaT+XjpUXijfKnFK81pFNiQOHft29cf42NjGE9jyFu51jZA3ikWeE83IxfFx8gwngQe`
              && `gRSZF+T1iMC+Pf3Gt7/PT3Ex3gYmrLT18llLXo8ILCTez1369vnpLcbDaGvlT9q6Iq+HBJ7bF/GVH/hgPIkWrf4p5PW2wE1f7nbcJ368VG+q8e/Fxs94HYYG2gq3H15npgjHSo8X1qqGvB4XeC4L/89mcBuMG9wBp1Xg/yw2flxZbLxeWap/Vo9nZal2urJUy/72219/`
              && `w9sMCguM0Naq5zMlcFa/mvRQGnfAaRQYS2sXIDnedhBYYAUtVimAvDMmMExqvcdtMU5wB5xGgfExOMUwMzELbCJmnWPVq5kUOCtuLU1sVhp3wKkVeLH+7fFibXtl6V0QPgNR4f/4+OAzvL1bWGATbbWdfWdNYCFxbnJZGHfAaRR4ZbFet8usK4u1Y/X4YF28jltYYPOB`
              && `DS1auZlxgZuDPuCx8q/6su0kzmLjtV0Hxx0QOiVkMejonR2/dmwnN5S98mvjd9gG9te5//pnu6xnJ4CogyreYv2bk6ROwAQWro/dMgh8jHDdjNeBtoZldvVXEW2zVMvSfxdoL2PEMGhdJsbCWvX9Qqwt76wKfC+bPcFt0yvUH58KahIHr0MOO5UAUXEZuHNSAR0WC0gJ`
              && `AOXjz9SgjsEJ64mo8VouM4bZ9W8dyxXBxfZYvMXasVxG1V/dlhIOh5R00LpMDC1WuWSBc817udyN2xnpXgUGifC2eJ1eQmYNCe5YdqHKA1ACdAvqGOyAelq2Rycgqg5SKmqZehKilstlQkii/jjULEuV12tdJkIkVvlFyMsCi5jLuXvZQQzRYDgGw2hFLvFHR2d1nMFw`
              && `p4BQ75tSGRmLCJkAPoP9qZ2KyqbqdlSnFOWbkonhJ8qgEHhoSWEOXTuG89AWeD3Acp1sZj58YsLDVar+TsuwrHBs+Fjc1mUiaKvVNyxwW+B7uv4Rt9GgWLIzuh7FnYya5MES4aGdE7jjqScQqpPjDAlgEeGY8DoqYuiKTlx43ypkpsbXzsRwlaq/0zI8cqFwW5eJoEWr`
              && `5yxwW2BfLnuN26gfxFldZmObYS3u/Hg5zggA1RnxOiLjiREA7L92bLd/nIXw8s5SDSwd2OEEQmV9CDt5JXgfOKjhqlP97a5/5Uips6RO3NRl7ERWq/cXoqa8LLAReq75Qzbr+MelsMs4VIxCYDtpqBilwHikYKxbO+0mr8SuDalRAdCt/ng43BEws+4gcr91GTtatJxh`
              && `ga0C39Ozb3FbOUENuWCYbGTjd0Fr5x+uwNAJUfmn0MlgG+M6tDMTj0pgUl7icsAJPFRvl9N5zS/ppf647jjshOy3LmNnIVr9wgJTAufOcVs5AX/Qjj8wOqvjDuRGYJxh1Zlg3NHw8G4oAuNrcNSJyYk2GzHswO2Eg8rivdZfnMjEfXIiIxMTa27qMna0aOWCBSYFvsFt`
              && `5YSTINTybgLj5YClDCWz4e3V7ahbKd0ExvWnylDl7La8F6jrVSwbdfuKqj9eB0Ptq9vyXuoyVuDZZyEvC2wRWFwH53r/FktKLujUxi2Yzuxs/vEdBRZhzlT3chsHL5PbGsN360RWN4HVdcxLAMtQUs3yOPu66dx4H8YTbfSlibodVX+lzNatvX5urbmty1iJrJYescD2`
              && `AvdzPxh34G7Rk8AOga8/cXboFr0I7Bio41InCaew1J8YrkrhqJOXKiNVf6dyqVD/HtQ2vdZlrERWq8sssL3AvlxuB7eZHWIIaTdjaT4HbddhAGob/Jm6DHcaasjXWp+4ndVNYMf9E5NSuPxuoQpM1V1tH3J4rtSBqr9cRsmIQ72fO2hdxko4VkmxwI4C/4nbzAlqqCuH`
              && `0rgjdRMYOqWYcFFEEicC4tpYAp3P8vKBuT4WrJvArTorJyVxX9nmlgvevluoAlMnC3yCwu0HIY+Bqr/czmgT6wnMeDHD+mLIoHUZK9pqdYcFdhBYz77BbeZFnARgbjGte8AsMAvMAk8fnIFZYIAFnlJYYBYYYIGnFJ6F7iJwH7PQ0wwLPKWI94BZYFuB72WzKdxmXoQF`
              && `nlLg935ZYHuBB/1+LIYZOVq0esUC0wK7/WodhhkbC9HKVxaYEjh7gduKYW4d4ut0WGCLwKP4Wh2GGTqtFxpY4A6B+3mRgWEmihY1v1KWBRYB7wLz9S8zNSzEqh9YYEXgXO4LbiM3/KMW3/6pnmz+892W5VsQH9ZSWVimBnyG1wN+aiR/FMsbKcsL8g8b8Z+t5SSOqX1i`
              && `YFtYt7VdPfEN6iyXwWd4m2Ezrv14GvF7wCxwS+C57HB+LxiEeFhPfH7YSFq+PwlkhWWt/5sdmZS0nqiLcmpJy0vyWAAQ11zf8hUxKrAfub+//xUXb9vAiUKeRHC5o2Jc+/E04lcJ4XeRWGAR/nz+Pm6jfhEyQEarJ5cpmbDA4jMhqTULiwxpZmGcWSkBQEhDzjj5ilsr`
              && `azeStt/pRJU7Csa1H8+zEK1+YoHF01dnuG3cAFlXZl5KJlLgWvJUDmEl4gRgZl4Y7mLB7QSg9ikR5dQTji+hU+XKugj5jSF3XWZvgNon/gz+LcsQIxTzUkLdhnGBmI1mgYc2+6xmOHNI2yEMFhjEhQ6tCmGslzhuX5emfsfZnBINsrRThoVlICP+XAWXC+urMoqhei15`
              && `qh4DlhV/Juslj0eWgevPuERk4RkWeFjZFw+bZedX1+llEksOheWwWf5fFROL1hKrlrD92Q9KNAwuV2ZLdR18oqDKVT8ToxJUL6ptGJeI3wiOtX8jeNYEHtazzyLjookroyO3J6hwBqbEExkXTVzh4a8UDZ8IcCZXoUTDYIHttlE/p9bpEJi4xsf7YQZE/E7wDArsy+qu`
              && `fxdYpZUlieiQk7oGxllPud7EYbdNL4gTBTEzroLLFfVngW8/RhauXs2SwL6sftPPd0A7QV2nGp8bHVUOh7sJLIeneNa5PcNsZHM3Ajjdn5bgctV7xBJcRznrLpfj2XA8egB4CD0CtFilMFMC7+fe4zZwi1N263hQossQ2rhetN73BcQQvbVe/wIDpkwdwkFZMkPictv3`
              && `jS2TWO3hvKhX8hTEhYD/dwhslinbQN5qc1N/pgsLserJLAjsy+bOhvXYJJ7UwahSkpNYjeRree1KZTyJzFpCEpcCA+Ikot4WgttYptBUuXJ00V7fOukGJ6X28XRKL8toLa8nPsvRgFoOMwTEwx1rlXNPC7ynX/lfDf7QBsPcSuC3g8PrlWsvCuzb12/mMvov+JgZxlOE`
              && `1qvL4fXyjdcEntsbzvPODHPr0dbLGS8J7Nsb3qQVw0wF2kal4AWBfRn9Az42hpkJxHA6XrmeRoHnMgc3PGxmZp7I2tGDcLx8MVUCvzq44gkrhjGB75MOJcpfpkNg/ZxvFTEMQThRfhuKl25ur8D6h2E9pMEwngSenQ4nyh9vlcAvD06Cu8N5tplhZoLg5tGDUKp0MmGB`
              && `z/yv8o4vqjMM40AoWVwOp0pfxymw/3n+IribH8o3aTAMYw6t51PFTOhJ8WQkAj8rnAWf5gs8VGaYEQMvRsxvFlOhreLH+SfFazcCB54eXgeeFj75n+V3Iun8XbwPhmHGBAgd2iouz28f7QT/OHxjibQRgXQh40/nl1lYhmEYhmEYhmEYpk/+D+9aCyOavaoAAAAAAElF`
              && `TkSuQmCC` ##NO_TEXT.
      WHEN 'JPG'.
        lv_b64 = `/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAoHBwkHBgoJCAkLCwoMDxkQDw4ODx4WFxIZJCAmJSMgIyIoLTkwKCo2KyIjMkQyNjs9QEBAJjBGS0U+Sjk/QD3/2wBDAQsLCw8NDx0QEB09KSMpPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09`
              && `PT09PT09PT3/wAARCADIAUADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVW`
              && `V1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAEC`
              && `AxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq`
              && `8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDssUYp2Kqarff2Zpst35fmeXj5d2M5IHX8a9aUlFNvoeNCnKpJQju9CzijFZt9qtzBqYsrSw+1SGHzj++CYGSO4+n51DJ4hYWSyxWTtOboWjQtIFxJjnDcgjPGeKh14JtN7eTOiOCrSSaW/mv89PmbGKMVRg1CdLeefVLQWMUQ`
              && `BDGYPu9en4fXNV7bWL66MUqaRILWVwFlMo3bScbtuM+/9cc0e2jp/kyVhKrva1l1ureid7N+SNbFGKr6hcXFtEhtLQ3Ujvt2BwmBgnJJ47frUNhqhvLG4ne2kSW3d45IVIc7l5wuOv8AjVOpFS5SFh5uHtFt6r8ty9ijFZltrUn2qaDUbM2bxQG4z5gkGwHBPFMsvEK3`
              && `FwqXNrJaRzI0kEkh4kUZJJ/u8DP0/DM+3hpqavBVkm+Xbs0/us9fka2KMVE19aLbrO1zAIWOFkMg2k+gPTsfyqjqesz6c8jnT5HtIiokn8wL1x91erdR+OR2qpVYxV2zOnhqlSXLFa/d+fXyNPFGKytQ177HdyQw2ctyLdQ9y6ceUpxj6nHP+Th17rghMKWNrJeyyxee`
              && `ETK4jx97OO/p/wDWzLrwV9di1gq75Wo7+nrr2072NPFGKisbyLULOO5gJMcgyMjBHYj86sYrRSTV0c8oOLcZKzQzFGKfijFO4rDMUYp+KMUXCwzFGKfijFFwsMxRin4oxRcLDMUYp+KMUXCwzFGKfijFFwsMxRin4oxRcLDMUYp+KMUXCwzFGKfiozLEpIaRAR2LCi4W`
              && `FxRik8+H/nqn/fQo8+H/AJ6p/wB9Ci4WFxRik8+H/nqn/fQp4IYAqQQe4ouFhuKMU/FGKLhYXFY3iES3tu2lwW07SThSJtv7pMNk7m7cD+VbeKXFZTjzxce5vQqeymqlrtar16HP32mS33icN5l3BELPHnQMVy28/Lux75xUWuaVHBpVpDDbT3Ea3ayThAXkkGDuYkck`
              && `n1/lXS4oxWboRafmdEMdUi4do9DmVsYb7QL6x0+wu7LOHC3CEb24OBk/7IHtmm3E97farZGG21S2kSVFnQnEG0EluRwfr3FdRijFJ0PMpY5ptuN97XberST9dkZmtmMWIjnsp7yGVwrxwglgOSDx7gdx1qhYm90zw9dPFZSErKzWtuwywjYjGQMnjJJB54rosUYq5U7y`
              && `5r62MoYjlpqm43V7vV/0u1+xzNjD9shvYjBqH2ye3ZTc3sW1eeNq4ztGTnA/oKZp6Xk15pxnsZrePS4XWVnBO87dvyAD5umePX6Z6nFGKhUNtdv+HNXjm+b3d/wurO3y+44izsLm31G2vpbKcWX2uV1gEGWiDAbWKjnsO3G3jqKvXd1Lc6lKL+w1OS2hkKxwwQkxSAH7`
              && `znjdkgEDoMd8mupxRipWH5VZM0nmLnJTlDVKy121v9/mcvqcV7a3+prFZSXKalEqRvGeEIXad3p1Jz7fXDpYLzRbuzultHu1SxFqywZJDjnOMdDjGf8A6wPTYoxVOgrtp/8AA1uZrHOyi4pq1nvrpZemi6GZoNjLp+i29vcACVASwBzjLE4/WtHFOxRitopRiorocdWb`
              && `qzc5bt3+8bijFOxRiquRYbijFOxRii4WG4oxTsUYouFhuKMU7FGKLhYbijFOxRii4WG4oxTsUYouFhuKMU7FGKLhYbijFOxRii4WG4rjNW/5Ctx/v12uK4rV/wDkK3H+/TQFOiiiqAK7PSR/xKrf/crjK7TSB/xKrf8A3KUgLeKMU7FGKm4WHYoxTsUYqblWG4oxTsUY`
              && `ouFhuKMU7FGKLhYbijFOxRii4WG4oxTsUYouFhuKMU7FGKLhYbijFOxRii4WG4oxTsUYouFhuKMU7FGKLhYbijFOxRii4WG4rAvfEMtreSwrAjBGxkk810OK4fWP+Qvc/wC/VR1Bmh/wlM3/AD7x/maP+Epm/wCfeP8AM1hUVVkSbv8AwlM3/PvH+ZrfspzdWcUzKFLr`
              && `nA7Vwddxo4/4lFt/uUpKw0W8UYp2KMVFx2G4oxTsUYouFhuKMU7FGKLhYbiuI1f/AJC1z/v13OK4fWP+Qvc/79VHcGUqKKKskK7bRx/xKbb/AHK4mu40cf8AEotv9yplsNFvFGKdijFRcdh2KKdijFSVYbRTsUYoCw2inYoxQFhtFOxRigLDaKdijFAWG0U7FGKAsNop`
              && `2KMUBYbRTsUYoCw2inYoxQFhtFOxRigLDa4XWf8AkL3P+/XeYrhNa/5DFz/v1cNxSKNFFFaEBXc6N/yCLb/crhq7vRR/xJ7b/cqJ7FRLlRtPEpIaVAR1BYVNiuE1r/kMXP8Av1MVcb0O0+0Q/wDPaP8A76FH2iH/AJ7R/wDfQrz2iq5Bcx6F9oh/57R/99CpFIYAqQQe`
              && `hFec13eij/iT23+5UyjYady5XC6z/wAhe5/367zFcJrX/IYuf9+nDcJFGiiitCArudG/5BFt/uVw1d3oo/4k9t/uVE9iolyinYoxWZVh2KMU7FGKm5dhuKMU7FGKLhYbijFOxRii4WG4oxTsUYouFhuKMU7Fc5f+J5rO+mgW3jYRtgEk800m9hN2OhxRiuW/4TCf/n1j`
              && `/wC+jR/wmE//AD6x/wDfRquSQuZHU4oxXLf8JhP/AM+sf/fRro7C4N5YwzsoUyLkgdqTi1uNNMmxRinYoxU3HYbijFOxRii4WG4oxTsUYouFhuK4LWv+Qzdf79d/iuA1v/kM3X+/WlPciexRooorUzCu90Qf8Sa1/wByuCrv9EH/ABJrX/cqKmxcNy5iuC1r/kM3X+/X`
              && `f4rgNb/5DN1/v1NPcc9ijRRRWpmFd7og/wCJNa/7lcFXf6IP+JNa/wC5UVNi4blzFcFrX/IZuv8Afrv8VwGt/wDIZuv9+pp7jnsUaKKK1MwrvdEH/Emtf9yuCrv9EH/Emtf9yoqbFw3LmKMU7FGKxuaWHYoxTsUYqblWG4oxTsUYouFhuKMU7FGKLhYbijFOxRii4WG4`
              && `rz/XP+Q1df79ehYrz3Xf+Q3d/wC/WtLcipsUKKKK2MQr0DQx/wASW1/3K8/r0LQh/wASS0/3Kzq7GlPcu4oxTsUYrC5rYbijFOxRii4WG4oxTsUYouFhuK8/1z/kNXX+/XoWK8913/kN3f8Av1rS3IqbFCiiitjEK9A0Mf8AEltf9yvP69C0If8AEktP9ys6uxpT3LuK`
              && `8/1z/kNXX+/XoWK8913/AJDd3/v1NLcqpsUKKKK2MQr0DQx/xJbX/crz+vQtCH/EktP9ys6uxpT3LuK8/wBc/wCQ1df79ehYrz3Xf+Q3d/79TS3KqbFCiiitjEK9A0Mf8SW1/wByvP69C0If8SS0/wBys6uxpT3LuKMU7FGKwua2H4oxTsUYqblWG4oxTsUYouFhuKMU`
              && `7FGKLhYbijFOxRii4WG4rzvXf+Q3d/8AXSvRsV5zr3/Icu/+ula0dzOrsZ9FFFdBgFeiaCP+JHaf7led16NoI/4kdp/1zrKtsa0ty9isG98VQ2V5LbtbyMY2wSCOa6DFec69/wAhy7/66VnTSk9S5tpaG/8A8Jpb/wDPrL/30KP+E0t/+fWX/voVyFFbeziZe0kdf/wm`
              && `lv8A8+sv/fQresrkXtnFcKpUSLkA9q8yr0bQR/xI7T/rnWdSKitC4SbepexXneu/8hu7/wCulejYrznXv+Q5d/8AXSlR3HV2M+iiiugwCvRNBH/EjtP9yvO69G0Ef8SO0/651lW2NaW5exXneu/8hu7/AOulejYrznXv+Q5d/wDXSoo7lVdjPoooroMAr0TQR/xI7T/c`
              && `rzuvRtBH/EjtP+udZVtjWluXsV53rv8AyG7v/rpXo2K8517/AJDl3/10qKO5VXYz6KKK6DAK9E0Ef8SO0/3K87r0bQR/xI7T/rnWVbY1pbl7FGKdijFc9zaw7FGKdijFTcuw3FGKdisi68TafZ3MkErSeZGcNhM00m9hNpbmrijFYv8Awl+l/wB6X/v3R/wl+l/3pf8A`
              && `v3T5JdieaPc2sUYrF/4S/S/70v8A37rWtbiO8to54cmOQZXIxSaa3GmnsSYrzfX/APkO3f8A10r0rFebeIP+Q9ef9dK1obszrbGdRRRXSc4V6PoA/wCJFaf9c684r0nw+P8AiQ2f/XOsa+yNqO5fxXm+v/8AIdu/+ulelYrzbxB/yHrz/rpUUN2VW2M6iiiuk5wr0fQB`
              && `/wASK0/6515xXpPh8f8AEhs/+udY19kbUdy/ivN9f/5Dt3/10r0rFebeIP8AkPXn/XSoobsqtsZ1FFFdJzhXo+gD/iRWn/XOvOK9J8Pj/iQ2f/XOsa+yNqO5fxXm+v8A/Idu/wDrpXpWK828Qf8AIevP+ulRQ3ZVbYzqKKK6TnCvR9AH/EitP+udecV6T4fH/Ehs/wDr`
              && `nWNfZG1Hcv4rzfX/APkO3f8A10r0rFebeIP+Q9ef9dKihuyq2xnUUUV0nOFej6AP+JFaf9c684r0nw+P+JDZ/wDXOsa+yNqO5fxRinYoxXNc6LDsUYp2KMVNyrDcV5p4g/5D15/10r03FeZ+If8AkP3n/XSt8P8AEzCvsjNooorqOYK9L8Pj/iQ2f/XOvNK9M8Pf8gCz`
              && `/wCudYYj4Ub0NzQxXmniD/kPXn/XSvTcV5n4h/5D95/10qMP8TKr7IzaKKK6jmCvS/D4/wCJDZ/9c680r0zw9/yALP8A651hiPhRvQ3NDFeaeIP+Q9ef9dK9NxXmfiH/AJD95/10qMP8TKr7IzaKKK6jmCvS/D4/4kNn/wBc680r0zw9/wAgCz/651hiPhRvQ3NDFeae`
              && `IP8AkPXn/XSvTcV5n4h/5D95/wBdKjD/ABMqvsjNooorqOYK9L8Pj/iQ2f8A1zrzSvTPD3/IAs/+udYYj4Ub0NzQxXmniD/kPXn/AF0r03FeZ+If+Q/ef9dKjD/Eyq+yM2iiiuo5gr0vw+P+JDZ/9c680r0zw9/yALP/AK51hiPhRvQ3NDFeaeIP+Q9ef9dK9NxXmfiH`
              && `/kP3n/XSow/xMqvsjNooorqOYK9L8Pj/AIkNn/1zrzSvTPD3/IAs/wDrnWGI+FG9Dc0MUYp2KMVy3Omw7FGKdijFSVYbivMfEX/IwXn/AF0r1DFeX+I/+Rgvf+uldGH+JmGI+FGbRRRXWcgV6d4dH/FP2f8A1zrzGvUPDg/4p+y/651hiPhR0Yf4maGK8x8Rf8jBef8A`
              && `XSvUMV5f4j/5GC9/66Vnh/iZeI+FGbRRRXWcgV6d4dH/ABT9n/1zrzGvUPDg/wCKfsv+udYYj4UdGH+JmhivMfEX/IwXn/XSvUMV5f4j/wCRgvf+ulZ4f4mXiPhRm0UUV1nIFeneHR/xT9n/ANc68xr1Dw4P+Kfsv+udYYj4UdGH+JmhivMfEX/IwXn/AF0r1DFeX+I/`
              && `+Rgvf+ulZ4f4mXiPhRm0UUV1nIFeneHR/wAU/Z/9c68xr1Dw4P8Ain7L/rnWGI+FHRh/iZoYrzHxF/yMF5/10r1DFeX+I/8AkYL3/rpWeH+Jl4j4UZtFFFdZyBXp3h0f8U/Z/wDXOvMa9Q8OD/in7L/rnWGI+FHRh/iZoYrzHxF/yMF5/wBdK9QxXl/iP/kYL3/rpWeH`
              && `+Jl4j4UZtFFFdZyBXp3h0f8AFP2f/XOvMa9Q8OD/AIp+y/651hiPhR0Yf4maGKMU7FGK47nVYdijFOxRipuVYbivLvEf/Iw3v/XSvU8V5Z4k/wCRivf+uldOG+JnPifhRmUUUV2HGFeo+HB/xT1l/wBc68ur1Pw2P+Kdsv8ArnXPifhR0Yb4maOK8u8R/wDIw3v/AF0r`
              && `1PFeWeJP+Rivf+ulZ4b4maYn4UZlFFFdhxhXqPhwf8U9Zf8AXOvLq9T8Nj/inbL/AK51z4n4UdGG+JmjivLvEf8AyMN7/wBdK9TxXlniT/kYr3/rpWeG+JmmJ+FGZRRRXYcYV6j4cH/FPWX/AFzry6vU/DY/4p2y/wCudc+J+FHRhviZo4ry7xH/AMjDe/8AXSvU8V5Z`
              && `4k/5GK9/66VnhviZpifhRmUUUV2HGFeo+HB/xT1l/wBc68ur1Pw2P+Kdsv8ArnXPifhR0Yb4maOK8u8R/wDIw3v/AF0r1PFeWeJP+Rivf+ulZ4b4maYn4UZlFFFdhxhXqPhwf8U9Zf8AXOvLq9T8Nj/inbL/AK51z4n4UdGG+JmjivLvEf8AyMN7/wBdK9TxXlniT/kY`
              && `r3/rpWeG+JmmJ+FGZRRRXYcYV6j4cH/FPWX/AFzry6vU/DY/4p2y/wCudc+J+FHRhviZo4oxTsUYriudlh2KMU7FGKkuw3FeV+Jf+Rivv+uhr1bFZ8+gaZczvNNZRPI5yzEck1tRqKm7sxrU3NWR5LRXq/8AwjOkf8+EP5Uf8IzpH/PhD+VdH1qPYw+qy7nlFeq+Gh/x`
              && `Ttj/ANc6d/wjOkf8+EP5VoQW8dtAkMKBI0GFUdAKxrVlUVka0aLg7sXFeV+Jf+Rivv8Aroa9WxWfPoGmXM7zTWUTyOcsxHJNTRqKm7sqtTc1ZHktFer/APCM6R/z4Q/lR/wjOkf8+EP5V0fWo9jD6rLueUV6r4aH/FO2P/XOnf8ACM6R/wA+EP5VoQW8dtAkMKBI0GFU`
              && `dAKxrVlUVka0aLg7sXFeV+Jf+Rivv+uhr1bFZ8+gaZczvNNZRPI5yzEck1NGoqbuyq1NzVkeS0V6v/wjOkf8+EP5Uf8ACM6R/wA+EP5V0fWo9jD6rLueUV6r4aH/ABTtj/1zp3/CM6R/z4Q/lWhBbx20CQwoEjQYVR0ArGtWVRWRrRouDuxcV5X4l/5GK+/66GvVsVnz`
              && `6BplzO801lE8jnLMRyTU0aipu7KrU3NWR5LRXq//AAjOkf8APhD+VH/CM6R/z4Q/lXR9aj2MPqsu55RXqvhof8U7Y/8AXOnf8IzpH/PhD+VaEFvHbQJDCgSNBhVHQCsa1ZVFZGtGi4O7FxXlfiX/AJGK+/66GvVsVnz6BplzO801lE8jnLMRyTU0aipu7KrU3NWR5LRX`
              && `q/8AwjOkf8+EP5Uf8IzpH/PhD+VdH1qPYw+qy7nlFeq+Gh/xTtj/ANc6d/wjOkf8+EP5VoQW8dtAkMKBI0GFUdAKxrVlUVka0aLg7sXFeV+Jf+Rivv8Aroa9WxWfPoGmXM7zTWUTyOcsxHJNTRqKm7sqtTc1ZHktFer/APCM6R/z4Q/lR/wjOkf8+EP5V0fWo9jD6rLu`
              && `eUV6r4aH/FO2P/XOnf8ACM6R/wA+EP5VoQW8dtAkMKBI0GFUdAKxrVlUVka0aLg7sXFGKdijFc50WH4oxRRUjDFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFGKKKADFG`
              && `KKKADFGKKKADFGKKKAP/2Q==` ##NO_TEXT.
    ENDCASE.
    rv_image = cl_web_http_utility=>decode_x_base64( lv_b64 ).
  ENDMETHOD.


  METHOD build_image_drawing_xlsx.
    " Test workbook: images with one-cell / two-cell / absolute anchor (sheet "Images"),
    " same image on a 2nd sheet (sheet "Reuse"), header / footer images (sheet "HeaderFooter")
    DATA: lo_excel   TYPE REF TO zcl_excel,
          lo_sheet   TYPE REF TO zcl_excel_worksheet,
          lo_drawing TYPE REF TO zcl_excel_drawing,
          lo_writer  TYPE REF TO zif_excel_writer,
          ls_header  TYPE zexcel_s_worksheet_head_foot,
          ls_footer  TYPE zexcel_s_worksheet_head_foot.

    DATA(lv_png) = get_test_image( 'PNG' ).
    DATA(lv_jpg) = get_test_image( 'JPG' ).

    CREATE OBJECT lo_excel.

    " ---- sheet "Images"
    lo_sheet = lo_excel->get_active_worksheet( ).
    lo_sheet->set_title( 'Images' ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `1) PNG one-cell anchor B2, 240x80 px` ) ##NO_TEXT.
    lo_drawing = lo_excel->add_new_drawing( ip_title = 'Logo PNG' ) ##NO_TEXT.
    lo_drawing->set_media( ip_media = lv_png ip_media_type = 'png' ip_width = 240 ip_height = 80 ) ##NO_TEXT.
    lo_drawing->set_position( ip_from_row = 2 ip_from_col = 'B' ).
    lo_sheet->add_drawing( lo_drawing ).

    lo_sheet->set_cell( ip_column = 'A' ip_row = 7 ip_value = `2) JPEG two-cell anchor B8:G20 (co giãn theo ô)` ) ##NO_TEXT.
    lo_drawing = lo_excel->add_new_drawing( ip_title = 'Photo JPEG' ) ##NO_TEXT.
    lo_drawing->set_media( ip_media = lv_jpg ip_media_type = zcl_excel_drawing=>c_media_type_jpg ip_width = 320 ip_height = 200 ).
    lo_drawing->set_position2( ip_from = VALUE #( col = 1 row = 7 )
                               ip_to   = VALUE #( col = 7 row = 20 ) ).
    lo_sheet->add_drawing( lo_drawing ).

    lo_sheet->set_cell( ip_column = 'I' ip_row = 1 ip_value = `3) PNG one-cell anchor I2 + offset 20/10 px, 120x40 px` ) ##NO_TEXT.
    lo_drawing = lo_excel->add_new_drawing( ip_title = 'Logo small' ) ##NO_TEXT.
    lo_drawing->set_media( ip_media = lv_png ip_media_type = 'png' ip_width = 120 ip_height = 40 ) ##NO_TEXT.
    lo_drawing->set_position( ip_from_row = 2 ip_from_col = 'I'
                              ip_coloff   = CONV #( zcl_excel_drawing=>pixel2emu( 20 ) )
                              ip_rowoff   = CONV #( zcl_excel_drawing=>pixel2emu( 10 ) ) ).
    lo_sheet->add_drawing( lo_drawing ).

    lo_sheet->set_cell( ip_column = 'I' ip_row = 7 ip_value = `4) PNG absolute anchor (set_position2 ABS)` ) ##NO_TEXT.
    lo_drawing = lo_excel->add_new_drawing( ip_title = 'Logo absolute' ) ##NO_TEXT.
    lo_drawing->set_media( ip_media = lv_png ip_media_type = 'png' ip_width = 240 ip_height = 80 ) ##NO_TEXT.
    lo_drawing->set_position2( ip_from   = VALUE #( col = 8 row = 7 )
                               ip_to     = VALUE #( )
                               ip_anchor = zcl_excel_drawing=>anchor_absolute ).
    lo_sheet->add_drawing( lo_drawing ).

    " ---- sheet "Reuse": the same PNG content again
    lo_sheet = lo_excel->add_new_worksheet( ).
    lo_sheet->set_title( 'Reuse' ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Cùng nội dung PNG trên sheet thứ 2` ) ##NO_TEXT.
    lo_drawing = lo_excel->add_new_drawing( ip_title = 'Logo again' ) ##NO_TEXT.
    lo_drawing->set_media( ip_media = lv_png ip_media_type = 'png' ip_width = 240 ip_height = 80 ) ##NO_TEXT.
    lo_drawing->set_position( ip_from_row = 3 ip_from_col = 'A' ).
    lo_sheet->add_drawing( lo_drawing ).

    " ---- sheet "HeaderFooter": image in the header (left) and in the footer (center) -> Page Layout view
    lo_sheet = lo_excel->add_new_worksheet( ).
    lo_sheet->set_title( 'HeaderFooter' ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Xem bằng View > Page Layout: logo ở header trái, ảnh ở footer giữa` ) ##NO_TEXT.
    lo_drawing = lo_excel->add_new_drawing( ip_type = zcl_excel_drawing=>type_image_header_footer ip_title = 'Header logo' ) ##NO_TEXT.
    lo_drawing->set_media( ip_media = lv_png ip_media_type = 'png' ip_width = 120 ip_height = 40 ) ##NO_TEXT.
    ls_header-left_image   = lo_drawing.
    ls_header-right_value  = `abap2xlsx - header image` ##NO_TEXT.
    lo_drawing = lo_excel->add_new_drawing( ip_type = zcl_excel_drawing=>type_image_header_footer ip_title = 'Footer photo' ) ##NO_TEXT.
    lo_drawing->set_media( ip_media = lv_jpg ip_media_type = zcl_excel_drawing=>c_media_type_jpg ip_width = 96 ip_height = 60 ).
    ls_footer-center_image = lo_drawing.
    ls_footer-center_value = `Ảnh:` ##NO_TEXT.   " text and image in the same section
    ls_footer-right_value  = `Trang &P` ##NO_TEXT.
    lo_sheet->sheet_setup->set_header_footer( ip_odd_header = ls_header ip_odd_footer = ls_footer ).

    lo_excel->set_active_sheet_index( 1 ).
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD build_richtext_merge_xlsx.
    " Test workbook: rich text runs (sheet "RichText"), merged cells (sheet "Merge")
    DATA: lo_excel  TYPE REF TO zcl_excel,
          lo_sheet  TYPE REF TO zcl_excel_worksheet,
          lo_writer TYPE REF TO zif_excel_writer,
          lt_rtf    TYPE zexcel_t_rtf,
          ls_rtf    TYPE zexcel_s_rtf,
          lv_text   TYPE string.

    CREATE OBJECT lo_excel.

    " ---- sheet "RichText"
    lo_sheet = lo_excel->get_active_worksheet( ).
    lo_sheet->set_title( 'RichText' ) ##NO_TEXT.
    lo_sheet->get_column( 'A' )->set_width( 60 ).

    " A1: bold / italic red / underline blue 14pt / strikethrough / Times New Roman
    lv_text = `Đậm, nghiêng đỏ, gạch chân xanh 14, gạch ngang, Times` ##NO_TEXT.
    CLEAR lt_rtf.
    ls_rtf = VALUE #( offset = 0  length = 3 font = VALUE #( bold = abap_true ) ).
    APPEND ls_rtf TO lt_rtf.
    ls_rtf = VALUE #( offset = 5  length = 10 font = VALUE #( italic = abap_true color-rgb = 'FFFF0000' ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    ls_rtf = VALUE #( offset = 17 length = 17 font = VALUE #( underline = abap_true underline_mode = zcl_excel_style_font=>c_underline_single
                                                             size = 14 color-rgb = 'FF0070C0' ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    ls_rtf = VALUE #( offset = 36 length = 10 font = VALUE #( strikethrough = abap_true ) ).
    APPEND ls_rtf TO lt_rtf.
    ls_rtf = VALUE #( offset = 48 length = 5 font = VALUE #( name = 'Times New Roman' size = 12 bold = abap_true italic = abap_true ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = lv_text it_rtf = lt_rtf ).

    " A2: runs with leading / trailing blanks (xml:space="preserve")
    lv_text = `Mã:  VT0001  - còn hàng ` ##NO_TEXT.
    CLEAR lt_rtf.
    ls_rtf = VALUE #( offset = 0 length = 3 font = VALUE #( bold = abap_true ) ).
    APPEND ls_rtf TO lt_rtf.
    ls_rtf = VALUE #( offset = 3 length = 10 font = VALUE #( color-rgb = 'FF00B050' bold = abap_true ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = lv_text it_rtf = lt_rtf ).

    " A3: XML special characters inside runs
    lv_text = |Điều kiện: a < b { zcl_excel_common=>unicode_char_from_code_point( 38 ) } c > "d"| ##NO_TEXT.
    CLEAR lt_rtf.
    ls_rtf = VALUE #( offset = 11 length = 5 font = VALUE #( bold = abap_true color-rgb = 'FFC00000' ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    ls_rtf = VALUE #( offset = 19 length = 7 font = VALUE #( italic = abap_true ) ).
    APPEND ls_rtf TO lt_rtf.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = lv_text it_rtf = lt_rtf ).

    " A4 / A5: same text, different runs -> two shared strings; A6: same text without runs -> third one
    lv_text = `Cùng nội dung khác định dạng` ##NO_TEXT.
    CLEAR lt_rtf.
    ls_rtf = VALUE #( offset = 0 length = 4 font = VALUE #( bold = abap_true ) ).
    APPEND ls_rtf TO lt_rtf.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = lv_text it_rtf = lt_rtf ).
    CLEAR lt_rtf.
    ls_rtf = VALUE #( offset = 5 length = 8 font = VALUE #( italic = abap_true color-rgb = 'FF7030A0' ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 5 ip_value = lv_text it_rtf = lt_rtf ).
    lo_sheet->set_cell( ip_column = 'A' ip_row = 6 ip_value = lv_text ).
    " A7: rich text in a cell with style font Consolas (gaps take the cell font)
    DATA(lo_st_mono) = lo_excel->add_new_style( ).
    lo_st_mono->font->name = 'Consolas' ##NO_TEXT.
    lo_st_mono->font->size = 10.
    lv_text = `SELECT * FROM zc_demo WHERE bukrs = '1000'` ##NO_TEXT.
    CLEAR lt_rtf.
    ls_rtf = VALUE #( offset = 0  length = 6 font = VALUE #( name = 'Consolas' size = 10 bold = abap_true color-rgb = 'FF0000FF' ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    ls_rtf = VALUE #( offset = 9  length = 4 font = VALUE #( name = 'Consolas' size = 10 bold = abap_true color-rgb = 'FF0000FF' ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    ls_rtf = VALUE #( offset = 22 length = 5 font = VALUE #( name = 'Consolas' size = 10 bold = abap_true color-rgb = 'FF0000FF' ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 7 ip_value = lv_text it_rtf = lt_rtf ip_style = lo_st_mono->get_guid( ) ).

    " ---- sheet "Merge"
    lo_sheet = lo_excel->add_new_worksheet( ).
    lo_sheet->set_title( 'Merge' ) ##NO_TEXT.
    DATA(lo_st_title) = lo_excel->add_new_style( ).
    lo_st_title->font->bold = abap_true.
    lo_st_title->font->size = 14.
    lo_st_title->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_st_title->alignment->vertical   = zcl_excel_style_alignment=>c_vertical_center.
    lo_st_title->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    lo_st_title->fill->fgcolor-rgb = 'FFD9E1F2' ##NO_TEXT.
    DATA(lo_st_box) = lo_excel->add_new_style( ).
    lo_st_box->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_st_box->alignment->vertical   = zcl_excel_style_alignment=>c_vertical_center.
    lo_st_box->alignment->wraptext   = abap_true.
    lo_st_box->borders->allborders = NEW zcl_excel_style_border( ).
    lo_st_box->borders->allborders->border_style = zcl_excel_style_border=>c_border_medium.
    lo_st_box->borders->allborders->border_color-rgb = 'FF1F4E78' ##NO_TEXT.

    " 1) horizontal title A1:F1 with value + style (set_merge ip_range)
    lo_sheet->set_merge( ip_range = 'A1:F1' ip_value = `BÁO CÁO TỔNG HỢP (A1:F1)` ip_style = lo_st_title->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_row_height( ip_row = 1 ip_height_fix = 28 ).
    " 2) vertical A3:A8 (column / row parameters)
    lo_sheet->set_merge( ip_column_start = 'A' ip_column_end = 'A' ip_row = 3 ip_row_to = 8
                         ip_value = `Dọc A3:A8` ip_style = lo_st_box->get_guid( ) ) ##NO_TEXT.
    " 3) block B3:C5 with border on all cells of the area
    lo_sheet->set_merge( ip_column_start = 'B' ip_column_end = 'C' ip_row = 3 ip_row_to = 5
                         ip_value = `Khối B3:C5 viền đậm` ip_style = lo_st_box->get_guid( ) ) ##NO_TEXT.
    " 4) set_area with ip_merge: E3:F5
    lo_sheet->set_area( ip_range = 'E3:F5' ip_value = `set_area merge E3:F5` ip_style = lo_st_box->get_guid( ) ip_merge = abap_true ) ##NO_TEXT.
    " 5) merge with a formula in the top-left cell: B7:C7 = SUM(D7:D8)
    lo_sheet->set_cell( ip_column = 'D' ip_row = 7 ip_value = 100 ).
    lo_sheet->set_cell( ip_column = 'D' ip_row = 8 ip_value = 250 ).
    lo_sheet->set_merge( ip_range = 'B7:C7' ip_formula = 'SUM(D7:D8)' ip_style = lo_st_box->get_guid( ) ) ##NO_TEXT.
    " 6) set_area_formula with ip_merge: E7:F8
    lo_sheet->set_area_formula( ip_range = 'E7:F8' ip_formula = 'D7*2' ip_merge = abap_true ) ##NO_TEXT.
    " 7) merged cell with rich text in the top-left cell: B10:F10
    lv_text = `Ghi chú: ô gộp chứa rich text` ##NO_TEXT.
    CLEAR lt_rtf.
    ls_rtf = VALUE #( offset = 0 length = 8 font = VALUE #( bold = abap_true color-rgb = 'FFFF0000' ) ) ##NO_TEXT.
    APPEND ls_rtf TO lt_rtf.
    lo_sheet->set_cell( ip_column = 'B' ip_row = 10 ip_value = lv_text it_rtf = lt_rtf ).
    lo_sheet->set_merge( ip_range = 'B10:F10' ).
    " 8) merge that is removed again: B12:C12 must not appear in the file
    lo_sheet->set_merge( ip_range = 'B12:C12' ip_value = `Không gộp (đã delete_merge)` ) ##NO_TEXT.
    lo_sheet->delete_merge( ip_cell_column = 'B' ip_cell_row = 12 ).
    " 9) overlapping merge must raise ZCX_EXCEL
    TRY.
        lo_sheet->set_merge( ip_range = 'C4:D4' ).
        lo_sheet->set_cell( ip_column = 'H' ip_row = 1 ip_value = `Overlap: no exception` ) ##NO_TEXT.
      CATCH zcx_excel.
        lo_sheet->set_cell( ip_column = 'H' ip_row = 1 ip_value = `Overlap C4:D4 raised ZCX_EXCEL` ) ##NO_TEXT.
    ENDTRY.
    lo_sheet->get_column( 'A' )->set_width( 12 ).
    lo_sheet->get_column( 'B' )->set_width( 14 ).
    lo_sheet->get_column( 'C' )->set_width( 14 ).
    lo_sheet->get_column( 'H' )->set_width( 32 ).

    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD build_autofilter_table_xlsx.
    " Test workbook: autofilter with value + text filters (sheet "AutoFilter"), two tables (sheet "Table")
    TYPES: BEGIN OF ty_s_sale,
             region  TYPE string,
             product TYPE string,
             qty     TYPE i,
             amount  TYPE p LENGTH 15 DECIMALS 2,
           END OF ty_s_sale,
           BEGIN OF ty_s_item,
             matnr TYPE c LENGTH 10,
             descr TYPE string,
             qty   TYPE i,
             price TYPE p LENGTH 15 DECIMALS 2,
             total TYPE p LENGTH 15 DECIMALS 2,
           END OF ty_s_item,
           BEGIN OF ty_s_unit,
             code TYPE c LENGTH 3,
             text TYPE string,
           END OF ty_s_unit.
    DATA: lo_excel      TYPE REF TO zcl_excel,
          lo_sheet      TYPE REF TO zcl_excel_worksheet,
          lo_autofilter TYPE REF TO zcl_excel_autofilter,
          lo_writer     TYPE REF TO zif_excel_writer,
          lt_sales      TYPE STANDARD TABLE OF ty_s_sale WITH EMPTY KEY,
          lt_items      TYPE STANDARD TABLE OF ty_s_item WITH EMPTY KEY,
          lt_units      TYPE STANDARD TABLE OF ty_s_unit WITH EMPTY KEY,
          lt_fieldcat   TYPE zexcel_t_fieldcatalog,
          ls_settings   TYPE zexcel_s_table_settings,
          lv_row        TYPE i.

    CREATE OBJECT lo_excel.
    DATA(lo_st_bold) = lo_excel->add_new_style( ).
    lo_st_bold->font->bold = abap_true.
    DATA(lo_st_amount) = lo_excel->add_new_style( ).
    lo_st_amount->number_format->format_code = '#,##0' ##NO_TEXT.

    " ---- sheet "AutoFilter": A1:D13, column A = Hà Nội / Đà Nẵng, column B = Laptop* or *Mouse
    lo_sheet = lo_excel->get_active_worksheet( ).
    lo_sheet->set_title( 'AutoFilter' ) ##NO_TEXT.
    lt_sales = VALUE #( ( region = `Hà Nội`    product = `Laptop Dell`    qty = 5  amount = 125000000 )
                        ( region = `Hà Nội`    product = `Wireless Mouse` qty = 40 amount = 12000000 )
                        ( region = `Hà Nội`    product = `Monitor LG`     qty = 8  amount = 32000000 )
                        ( region = `Đà Nẵng`   product = `Laptop HP`      qty = 3  amount = 69000000 )
                        ( region = `Đà Nẵng`   product = `Keyboard`       qty = 25 amount = 7500000 )
                        ( region = `Đà Nẵng`   product = `Gaming Mouse`   qty = 12 amount = 9600000 )
                        ( region = `TP.HCM`    product = `Laptop Asus`    qty = 9  amount = 198000000 )
                        ( region = `TP.HCM`    product = `Office Mouse`   qty = 60 amount = 9000000 )
                        ( region = `TP.HCM`    product = `Printer Canon`  qty = 4  amount = 22000000 )
                        ( region = `Cần Thơ`   product = `Laptop Lenovo`  qty = 2  amount = 41000000 )
                        ( region = `Cần Thơ`   product = `Webcam`         qty = 15 amount = 10500000 )
                        ( region = `Hải Phòng` product = `Laptop Acer`    qty = 6  amount = 90000000 ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Khu vực` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = `Sản phẩm` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = `Số lượng` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = `Doanh thu` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lv_row = 2.
    LOOP AT lt_sales INTO DATA(ls_sale).
      lo_sheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_sale-region ).
      lo_sheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_sale-product ).
      lo_sheet->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_sale-qty ).
      lo_sheet->set_cell( ip_column = 'D' ip_row = lv_row ip_value = ls_sale-amount ip_style = lo_st_amount->get_guid( ) ).
      lv_row = lv_row + 1.
    ENDLOOP.
    lo_sheet->get_column( 'A' )->set_width( 14 ).
    lo_sheet->get_column( 'B' )->set_width( 20 ).
    lo_sheet->get_column( 'D' )->set_width( 16 ).

    lo_autofilter = lo_excel->add_new_autofilter( lo_sheet ).
    lo_autofilter->set_filter_area( VALUE #( row_start = 1 col_start = 1 row_end = lv_row - 1 col_end = 4 ) ).
    lo_autofilter->set_value( i_column = 1 i_value = `Hà Nội` ) ##NO_TEXT.
    lo_autofilter->set_value( i_column = 1 i_value = `Đà Nẵng` ) ##NO_TEXT.
    lo_autofilter->set_text_filter( i_column            = 2
                                    iv_textfilter1      = 'Laptop*'
                                    iv_logical_operator = zcl_excel_autofilter=>mc_logical_operator_or
                                    iv_textfilter2      = '*Mouse' ) ##NO_TEXT.

    " ---- sheet "Table": table 1 with totals row + column formula, table 2 without filter buttons
    lo_sheet = lo_excel->add_new_worksheet( ).
    lo_sheet->set_title( 'Table' ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = `DANH SÁCH VẬT TƯ` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lt_items = VALUE #( ( matnr = 'VT0001' descr = `Xi măng PCB40`       qty = 120  price = 85000 )
                        ( matnr = 'VT0002' descr = `Thép D10`            qty = 350  price = 17500 )
                        ( matnr = 'VT0003' descr = `Cát vàng (m3)`       qty = 40   price = 320000 )
                        ( matnr = 'VT0004' descr = `Gạch ống 8x8x18`     qty = 5000 price = 1250 )
                        ( matnr = 'VT0005' descr = `Sơn nước ngoại thất` qty = 18   price = 1450000 ) ) ##NO_TEXT.
    lt_fieldcat = zcl_excel_common=>get_fieldcatalog( ip_table = lt_items ).
    LOOP AT lt_fieldcat ASSIGNING FIELD-SYMBOL(<ls_fc>).
      CASE <ls_fc>-fieldname.
        WHEN 'MATNR'.
          <ls_fc>-scrtext_l = 'Mã VT' ##NO_TEXT.
          <ls_fc>-totals_function = zcl_excel_table=>totals_function_count.
        WHEN 'DESCR'.
          <ls_fc>-scrtext_l = 'Tên vật tư' ##NO_TEXT.
        WHEN 'QTY'.
          <ls_fc>-scrtext_l = 'Số lượng' ##NO_TEXT.
          <ls_fc>-totals_function = zcl_excel_table=>totals_function_sum.
        WHEN 'PRICE'.
          <ls_fc>-scrtext_l = 'Đơn giá' ##NO_TEXT.
          <ls_fc>-style = lo_st_amount->get_guid( ).
          <ls_fc>-totals_function = zcl_excel_table=>totals_function_average.
        WHEN 'TOTAL'.
          <ls_fc>-scrtext_l = 'Thành tiền' ##NO_TEXT.
          <ls_fc>-style = lo_st_amount->get_guid( ).
          <ls_fc>-column_formula = 'D4*E4' ##NO_TEXT.   " relative to the first data row (row 4)
          <ls_fc>-totals_function = zcl_excel_table=>totals_function_sum.
      ENDCASE.
      <ls_fc>-scrtext_m = <ls_fc>-scrtext_l.
      <ls_fc>-scrtext_s = <ls_fc>-scrtext_l.
    ENDLOOP.
    ls_settings = VALUE #( table_style      = zcl_excel_table=>builtinstyle_medium2
                           table_name       = 'tblVatTu'
                           top_left_column  = 'B'
                           top_left_row     = 3
                           show_row_stripes = abap_true ) ##NO_TEXT.
    lo_sheet->bind_table( EXPORTING ip_table          = lt_items
                                    it_field_catalog  = lt_fieldcat
                                    is_table_settings = ls_settings
                                    iv_default_descr  = 'L'
                          IMPORTING es_table_settings = ls_settings ).

    lt_units = VALUE #( ( code = 'KG'  text = `Kilogram` )
                        ( code = 'M3'  text = `Mét khối` )
                        ( code = 'CAI' text = `Cái` ) ) ##NO_TEXT.
    lt_fieldcat = zcl_excel_common=>get_fieldcatalog( ip_table = lt_units ).
    LOOP AT lt_fieldcat ASSIGNING <ls_fc>.
      <ls_fc>-scrtext_l = SWITCH #( <ls_fc>-fieldname WHEN 'CODE' THEN 'Mã ĐVT' ELSE 'Diễn giải' ) ##NO_TEXT.
      <ls_fc>-scrtext_m = <ls_fc>-scrtext_l.
      <ls_fc>-scrtext_s = <ls_fc>-scrtext_l.
    ENDLOOP.
    lo_sheet->bind_table( ip_table          = lt_units
                          it_field_catalog  = lt_fieldcat
                          is_table_settings = VALUE #( table_style         = zcl_excel_table=>builtinstyle_light9
                                                       table_name          = 'tblDonVi'
                                                       top_left_column     = 'I'
                                                       top_left_row        = 3
                                                       show_column_stripes = abap_true
                                                       nofilters           = abap_true )
                          iv_default_descr  = 'L' ) ##NO_TEXT.
    lo_sheet->get_column( 'C' )->set_width( 24 ).
    lo_sheet->get_column( 'E' )->set_width( 14 ).
    lo_sheet->get_column( 'F' )->set_width( 16 ).
    lo_sheet->get_column( 'J' )->set_width( 14 ).

    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD build_comment_outline_xlsx.
    " Test workbook: row outlines (2 levels, one group collapsed), column outline, comments
    TYPES: BEGIN OF ty_s_line,
             text  TYPE string,
             m1    TYPE i,
             m2    TYPE i,
             m3    TYPE i,
             total TYPE abap_bool,
           END OF ty_s_line.
    DATA: lo_excel   TYPE REF TO zcl_excel,
          lo_sheet   TYPE REF TO zcl_excel_worksheet,
          lo_comment TYPE REF TO zcl_excel_comment,
          lo_writer  TYPE REF TO zif_excel_writer,
          lt_lines   TYPE STANDARD TABLE OF ty_s_line WITH EMPTY KEY,
          lv_row     TYPE i.

    CREATE OBJECT lo_excel.
    DATA(lo_st_bold) = lo_excel->add_new_style( ).
    lo_st_bold->font->bold = abap_true.

    lo_sheet = lo_excel->get_active_worksheet( ).
    lo_sheet->set_title( 'Outline' ) ##NO_TEXT.

    " rows 2-4 region North (details), 5 subtotal; 6-8 region South, 9 subtotal; 10 grand total
    lt_lines = VALUE #( ( text = 'Hà Nội'         m1 = 120 m2 = 130 m3 = 125 )
                        ( text = 'Hải Phòng'      m1 = 80  m2 = 85  m3 = 90 )
                        ( text = 'Bắc Ninh'       m1 = 60  m2 = 70  m3 = 65 )
                        ( text = 'Tổng Miền Bắc'  total = abap_true )
                        ( text = 'TP. Hồ Chí Minh' m1 = 200 m2 = 210 m3 = 220 )
                        ( text = 'Bình Dương'     m1 = 90  m2 = 95  m3 = 100 )
                        ( text = 'Đồng Nai'       m1 = 70  m2 = 75  m3 = 80 )
                        ( text = 'Tổng Miền Nam'  total = abap_true )
                        ( text = 'TỔNG CỘNG'      total = abap_true ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Khu vực` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = `Tháng 1` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = `Tháng 2` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = `Tháng 3` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lo_sheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = `Quý 1` ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
    lv_row = 2.
    LOOP AT lt_lines INTO DATA(ls_line).
      IF ls_line-total = abap_true.
        lo_sheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_line-text ip_style = lo_st_bold->get_guid( ) ).
        CASE lv_row.
          WHEN 5.
            lo_sheet->set_cell( ip_column = 'E' ip_row = lv_row ip_formula = 'SUM(E2:E4)' ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
          WHEN 9.
            lo_sheet->set_cell( ip_column = 'E' ip_row = lv_row ip_formula = 'SUM(E6:E8)' ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
          WHEN OTHERS.
            lo_sheet->set_cell( ip_column = 'E' ip_row = lv_row ip_formula = 'E5+E9' ip_style = lo_st_bold->get_guid( ) ) ##NO_TEXT.
        ENDCASE.
      ELSE.
        lo_sheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_line-text ).
        lo_sheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_line-m1 ).
        lo_sheet->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_line-m2 ).
        lo_sheet->set_cell( ip_column = 'D' ip_row = lv_row ip_value = ls_line-m3 ).
        lo_sheet->set_cell( ip_column = 'E' ip_row = lv_row ip_formula = |SUM(B{ lv_row }:D{ lv_row })| ) ##NO_TEXT.
      ENDIF.
      lv_row = lv_row + 1.
    ENDLOOP.
    lo_sheet->get_column( 'A' )->set_width( 20 ).

    " row outlines: level 1 = rows 2-9 (both regions), level 2 = details; South details collapsed
    lo_sheet->set_row_outline( iv_row_from = 2 iv_row_to = 9 iv_collapsed = abap_false ).
    lo_sheet->set_row_outline( iv_row_from = 2 iv_row_to = 4 iv_collapsed = abap_false ).
    lo_sheet->set_row_outline( iv_row_from = 6 iv_row_to = 8 iv_collapsed = abap_true ).

    " column outline: months B:D grouped (level 1), visible
    lo_sheet->get_column( 'B' )->set_outline_level( 1 ).
    lo_sheet->get_column( 'C' )->set_outline_level( 1 ).
    lo_sheet->get_column( 'D' )->set_outline_level( 1 ).

    " comments
    CREATE OBJECT lo_comment.
    lo_comment->set_text( ip_ref = 'A1' ip_text = `Khu vực bán hàng (nhấn +/- bên trái để mở / thu gọn)` ) ##NO_TEXT.
    lo_sheet->add_comment( lo_comment ).
    CREATE OBJECT lo_comment.
    lo_comment->set_text( ip_ref = 'E5' ip_text = |Tổng Miền Bắc{ cl_abap_char_utilities=>newline }= SUM(E2:E4)| ) ##NO_TEXT.
    lo_sheet->add_comment( lo_comment ).
    CREATE OBJECT lo_comment.
    lo_comment->set_text( ip_ref = 'A9' ip_text = `Các tỉnh Miền Nam đang được thu gọn` ) ##NO_TEXT.
    lo_sheet->add_comment( lo_comment ).
    CREATE OBJECT lo_comment.
    lo_comment->set_text( ip_ref = 'E10' ip_text = |Ký tự đặc biệt: <b> "x" { zcl_excel_common=>unicode_char_from_code_point( 38 ) } y €| ) ##NO_TEXT.
    lo_sheet->add_comment( lo_comment ).

    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    rv_file = lo_writer->write_file( lo_excel ).
  ENDMETHOD.


  METHOD new_style.
    DATA(lo_style) = io_excel->add_new_style( ).
    lo_style->font->name = iv_font.
    lo_style->font->size = iv_size.
    lo_style->font->bold = iv_bold.
    lo_style->font->color-rgb = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal = iv_halign.
    lo_style->alignment->vertical = iv_valign.
    IF io_border IS BOUND.
      lo_style->borders->allborders = io_border.
    ENDIF.
    IF iv_format IS NOT INITIAL.
      lo_style->number_format->format_code = iv_format.
    ENDIF.
    rv_guid = lo_style->get_guid( ).
  ENDMETHOD.


  METHOD set_line.
    " text line merged over the columns A:H of the table
    io_sheet->set_cell( ip_column = 'A' ip_row = iv_row ip_value = iv_text ip_style = iv_style ).
    io_sheet->set_merge( ip_range = |A{ iv_row }:H{ iv_row }| ).
  ENDMETHOD.


  METHOD set_amount.
    " zero amounts stay empty (bordered) like in the template
    IF iv_amount IS INITIAL.
      io_sheet->set_cell_style( ip_column = iv_column ip_row = iv_row ip_style = iv_style ).
    ELSE.
      io_sheet->set_cell( ip_column = iv_column ip_row = iv_row ip_value = iv_amount ip_style = iv_style ).
    ENDIF.
  ENDMETHOD.


  METHOD matches_ranges.
    " value against RAP filter ranges (I-entries: any must match / E-entries: none may match)
    DATA: lr_low     TYPE REF TO data,
          lr_high    TYPE REF TO data,
          lv_has_inc TYPE abap_bool,
          lv_inc_hit TYPE abap_bool,
          lv_hit     TYPE abap_bool.
    FIELD-SYMBOLS: <lv_low>  TYPE any,
                   <lv_high> TYPE any.

    rv_match = abap_true.
    CREATE DATA lr_low LIKE iv_value.
    CREATE DATA lr_high LIKE iv_value.
    ASSIGN lr_low->* TO <lv_low>.
    ASSIGN lr_high->* TO <lv_high>.

    LOOP AT it_range INTO DATA(ls_range).
      TRY.
          <lv_low>  = ls_range-low.
          <lv_high> = ls_range-high.
        CATCH cx_sy_conversion_error.
          CLEAR: <lv_low>, <lv_high>.
      ENDTRY.
      CASE ls_range-option.
        WHEN 'EQ'.
          lv_hit = xsdbool( iv_value = <lv_low> ).
        WHEN 'NE'.
          lv_hit = xsdbool( iv_value <> <lv_low> ).
        WHEN 'GT'.
          lv_hit = xsdbool( iv_value > <lv_low> ).
        WHEN 'GE'.
          lv_hit = xsdbool( iv_value >= <lv_low> ).
        WHEN 'LT'.
          lv_hit = xsdbool( iv_value < <lv_low> ).
        WHEN 'LE'.
          lv_hit = xsdbool( iv_value <= <lv_low> ).
        WHEN 'BT'.
          lv_hit = xsdbool( iv_value BETWEEN <lv_low> AND <lv_high> ).
        WHEN 'NB'.
          lv_hit = xsdbool( iv_value NOT BETWEEN <lv_low> AND <lv_high> ).
        WHEN 'CP'.
          lv_hit = xsdbool( CONV string( iv_value ) CP ls_range-low ).
        WHEN 'NP'.
          lv_hit = xsdbool( CONV string( iv_value ) NP ls_range-low ).
        WHEN OTHERS.
          lv_hit = abap_false.
      ENDCASE.
      IF ls_range-sign = 'E'.
        IF lv_hit = abap_true.
          rv_match = abap_false.
          RETURN.
        ENDIF.
      ELSE.
        lv_has_inc = abap_true.
        IF lv_hit = abap_true.
          lv_inc_hit = abap_true.
        ENDIF.
      ENDIF.
    ENDLOOP.
    IF lv_has_inc = abap_true AND lv_inc_hit = abap_false.
      rv_match = abap_false.
    ENDIF.
  ENDMETHOD.


  METHOD if_rap_query_provider~select.
    TYPES: BEGIN OF ty_s_file,
             companycode TYPE bukrs,
             file        TYPE xstring,
           END OF ty_s_file.
    DATA: lt_all_data TYPE ty_t_journal,
          lt_journal  TYPE ty_t_journal,
          lt_result   TYPE ty_t_entity,
          lt_sort     TYPE abap_sortorder_tab,
          lt_files    TYPE HASHED TABLE OF ty_s_file WITH UNIQUE KEY companycode,
          lv_offset   TYPE int8,
          lv_size     TYPE int8.
    FIELD-SYMBOLS <lv_value> TYPE any.

    TRY.
        " 1) data + filter
        lt_all_data = get_demo_data( ).
        lt_journal = lt_all_data.
        DATA(lt_filter) = io_request->get_filter( )->get_as_ranges( ).
        LOOP AT lt_journal INTO DATA(ls_journal).
          DATA(lv_tabix) = sy-tabix.
          LOOP AT lt_filter INTO DATA(ls_filter).
            ASSIGN COMPONENT ls_filter-name OF STRUCTURE ls_journal TO <lv_value>.
            IF sy-subrc = 0 AND matches_ranges( iv_value = <lv_value> it_range = ls_filter-range ) = abap_false.
              DELETE lt_journal INDEX lv_tabix.
              EXIT.
            ENDIF.
          ENDLOOP.
        ENDLOOP.

        " 2) total count
        IF io_request->is_total_numb_of_rec_requested( ).
          io_response->set_total_number_of_records( lines( lt_journal ) ).
        ENDIF.

        IF io_request->is_data_requested( ).
          " 3) sort
          LOOP AT io_request->get_sort_elements( ) INTO DATA(ls_sort).
            APPEND VALUE #( name = ls_sort-element_name descending = ls_sort-descending ) TO lt_sort.
          ENDLOOP.
          IF lt_sort IS INITIAL.
            lt_sort = VALUE #( ( name = 'COMPANYCODE' ) ( name = 'FISCALYEAR' )
                               ( name = 'ACCOUNTING_DOC' ) ( name = 'DOCUMENT_ITM' ) ) ##NO_TEXT.
          ENDIF.
          SORT lt_journal BY (lt_sort).

          " 4) the Excel file content is only built if the stream is requested (download link / media read)
          DATA(lt_requested) = io_request->get_requested_elements( ).
          DATA(lv_attachment_requested) = xsdbool( line_exists( lt_requested[ table_line = 'ATTACHMENT' ] ) ).

          " 5) paging
          lv_offset = io_request->get_paging( )->get_offset( ).
          lv_size   = io_request->get_paging( )->get_page_size( ).
          LOOP AT lt_journal INTO ls_journal FROM lv_offset + 1.
            IF lv_size <> if_rap_query_paging=>page_size_unlimited AND lines( lt_result ) >= lv_size.
              EXIT.
            ENDIF.
            APPEND CORRESPONDING #( ls_journal ) TO lt_result ASSIGNING FIELD-SYMBOL(<ls_result>).
            " media type + file name are always sent: Fiori shows the download link only with a media type
            <ls_result>-mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ##NO_TEXT.
            <ls_result>-filename = |NhatKyGiaoDich_{ ls_journal-companycode }.xlsx| ##NO_TEXT.
            IF lv_attachment_requested = abap_true.
              " journal of the whole company code (one sheet per customer), independent of the key filter
              " of a single media read -> built once per company code and request
              READ TABLE lt_files INTO DATA(ls_file) WITH TABLE KEY companycode = ls_journal-companycode.
              IF sy-subrc <> 0.
                ls_file-companycode = ls_journal-companycode.
                ls_file-file = build_journal_xlsx( VALUE #( FOR ls_all IN lt_all_data
                                                            WHERE ( companycode = ls_journal-companycode ) ( ls_all ) ) ).
                INSERT ls_file INTO TABLE lt_files.
              ENDIF.
              <ls_result>-attachment = ls_file-file.
            ENDIF.
          ENDLOOP.
          io_response->set_data( lt_result ).
        ENDIF.

      CATCH cx_rap_query_filter_no_range INTO DATA(lx_filter).
        RAISE EXCEPTION NEW lcx_query( previous = lx_filter ).
      CATCH zcx_excel INTO DATA(lx_excel).
        RAISE EXCEPTION NEW lcx_query( previous = lx_excel ).
    ENDTRY.
  ENDMETHOD.

ENDCLASS.

