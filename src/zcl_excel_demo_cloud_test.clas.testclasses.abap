*"* use this source file for your ABAP unit test classes

"! Test double for the RAP query request / filter / paging / response
CLASS ltd_query DEFINITION FOR TESTING.
  PUBLIC SECTION.
    INTERFACES: if_rap_query_request  PARTIALLY IMPLEMENTED,
                if_rap_query_filter   PARTIALLY IMPLEMENTED,
                if_rap_query_paging   PARTIALLY IMPLEMENTED,
                if_rap_query_response PARTIALLY IMPLEMENTED.
    DATA: mt_ranges    TYPE if_rap_query_filter=>tt_name_range_pairs,
          mt_sort      TYPE if_rap_query_request=>tt_sort_elements,
          mt_requested TYPE if_rap_query_request=>tt_requested_elements,
          mv_offset    TYPE int8,
          mv_page_size TYPE int8 VALUE if_rap_query_paging=>page_size_unlimited,
          mt_data      TYPE zcl_excel_demo_cloud_test=>ty_t_entity,
          mv_count     TYPE int8 VALUE -1.
ENDCLASS.

CLASS ltd_query IMPLEMENTATION.
  METHOD if_rap_query_request~is_data_requested.
    rv_is_requested = abap_true.
  ENDMETHOD.
  METHOD if_rap_query_request~is_total_numb_of_rec_requested.
    rv_is_requested = abap_true.
  ENDMETHOD.
  METHOD if_rap_query_request~get_filter.
    ro_filter = me.
  ENDMETHOD.
  METHOD if_rap_query_request~get_paging.
    ro_paging = me.
  ENDMETHOD.
  METHOD if_rap_query_request~get_sort_elements.
    rt_sort_elements = mt_sort.
  ENDMETHOD.
  METHOD if_rap_query_request~get_requested_elements.
    rt_requested_elements = mt_requested.
  ENDMETHOD.
  METHOD if_rap_query_filter~get_as_ranges.
    rt_ranges = mt_ranges.
  ENDMETHOD.
  METHOD if_rap_query_paging~get_offset.
    rv_offset = mv_offset.
  ENDMETHOD.
  METHOD if_rap_query_paging~get_page_size.
    rv_page_size = mv_page_size.
  ENDMETHOD.
  METHOD if_rap_query_response~set_data.
    mt_data = it_data.
  ENDMETHOD.
  METHOD if_rap_query_response~set_total_number_of_records.
    mv_count = iv_total_number_of_records.
  ENDMETHOD.
ENDCLASS.


CLASS ltc_query DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.
  PRIVATE SECTION.
    DATA mo_cut TYPE REF TO zcl_excel_demo_cloud_test.
    DATA mo_query TYPE REF TO ltd_query.
    METHODS setup.
    METHODS filter_with_attachment FOR TESTING RAISING cx_static_check.
    METHODS sort_and_paging FOR TESTING RAISING cx_static_check.
    METHODS exclude_customer FOR TESTING RAISING cx_static_check.
    METHODS media_read_single_key FOR TESTING RAISING cx_static_check.
ENDCLASS.

CLASS ltc_query IMPLEMENTATION.

  METHOD setup.
    mo_cut = NEW #( ).
    mo_query = NEW #( ).
  ENDMETHOD.

  METHOD filter_with_attachment.
    DATA lt_titles TYPE string_table.

    mo_query->mt_ranges = VALUE #( ( name = 'COMPANYCODE' range = VALUE #( ( sign = 'I' option = 'EQ' low = '1000' ) ) ) ).
    mo_query->mt_requested = VALUE #( ( `COMPANYCODE` ) ( `ATTACHMENT` ) ( `MIMETYPE` ) ( `FILENAME` ) ).

    mo_cut->if_rap_query_provider~select( io_request = mo_query io_response = mo_query ).

    cl_abap_unit_assert=>assert_equals( act = mo_query->mv_count exp = 11 ).
    cl_abap_unit_assert=>assert_equals( act = lines( mo_query->mt_data ) exp = 11 ).
    DATA(ls_row) = mo_query->mt_data[ 1 ].
    cl_abap_unit_assert=>assert_not_initial( ls_row-attachment ).
    cl_abap_unit_assert=>assert_equals( act = ls_row-filename exp = 'NhatKyGiaoDich_1000.xlsx' ).
    cl_abap_unit_assert=>assert_equals( act = ls_row-mimetype
                                        exp = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ).

    " attachment is a valid workbook with one sheet per company code + customer
    " (filtered to company code 1000, customer 0100098012 occurs once -> sheet name without company code)
    DATA(lo_excel) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( ls_row-attachment ).
    DATA(lo_iterator) = lo_excel->get_worksheets_iterator( ).
    WHILE lo_iterator->has_next( ) = abap_true.
      APPEND CAST zcl_excel_worksheet( lo_iterator->get_next( ) )->get_title( ) TO lt_titles.
    ENDWHILE.
    cl_abap_unit_assert=>assert_equals( act = lt_titles exp = VALUE string_table( ( `0100098012` ) ( `0100098020` ) ) ).

    DATA(lo_sheet) = lo_excel->get_worksheet_by_name( '0100098020' ).
    lo_sheet->get_cell( EXPORTING ip_column = 'A' ip_row = 6 IMPORTING ep_value = DATA(lv_value) ).
    cl_abap_unit_assert=>assert_equals( act = lv_value exp = 'Nhà cung cấp: Công ty TNHH Bao bì Minh Phát' ).
    lo_sheet->get_cell( EXPORTING ip_column = 'F' ip_row = 17 IMPORTING ep_formula = DATA(lv_formula) ).
    cl_abap_unit_assert=>assert_equals( act = lv_formula exp = 'SUM(F14:F16)' ).
  ENDMETHOD.

  METHOD sort_and_paging.
    mo_query->mt_sort = VALUE #( ( element_name = 'AMOUNT_DR' descending = abap_true ) ).
    mo_query->mv_offset = 1.
    mo_query->mv_page_size = 2.

    mo_cut->if_rap_query_provider~select( io_request = mo_query io_response = mo_query ).

    cl_abap_unit_assert=>assert_equals( act = mo_query->mv_count exp = 13 ).
    cl_abap_unit_assert=>assert_equals( act = lines( mo_query->mt_data ) exp = 2 ).
    " highest amount 2.500.000 is skipped by the offset, then 1.250.000 and 700.000
    cl_abap_unit_assert=>assert_equals( act = mo_query->mt_data[ 1 ]-amount_dr exp = 1250000 ).
    cl_abap_unit_assert=>assert_equals( act = mo_query->mt_data[ 2 ]-amount_dr exp = 700000 ).
    " attachment not requested -> not built, but media type / file name are always set
    cl_abap_unit_assert=>assert_initial( mo_query->mt_data[ 1 ]-attachment ).
    cl_abap_unit_assert=>assert_not_initial( mo_query->mt_data[ 1 ]-mimetype ).
    cl_abap_unit_assert=>assert_not_initial( mo_query->mt_data[ 1 ]-filename ).
  ENDMETHOD.

  METHOD exclude_customer.
    mo_query->mt_ranges = VALUE #( ( name = 'CUSTOMER' range = VALUE #( ( sign = 'E' option = 'EQ' low = '0100098012' ) ) ) ).

    mo_cut->if_rap_query_provider~select( io_request = mo_query io_response = mo_query ).

    cl_abap_unit_assert=>assert_equals( act = lines( mo_query->mt_data ) exp = 3 ).
    LOOP AT mo_query->mt_data INTO DATA(ls_row).
      cl_abap_unit_assert=>assert_equals( act = ls_row-customer exp = '0100098020' ).
    ENDLOOP.
  ENDMETHOD.

  METHOD media_read_single_key.
    " media read of one row: filter on all keys, only the stream is requested
    mo_query->mt_ranges = VALUE #(
      ( name = 'COMPANYCODE'    range = VALUE #( ( sign = 'I' option = 'EQ' low = '1000' ) ) )
      ( name = 'FISCALYEAR'     range = VALUE #( ( sign = 'I' option = 'EQ' low = '2026' ) ) )
      ( name = 'ACCOUNTING_DOC' range = VALUE #( ( sign = 'I' option = 'EQ' low = '1000900001' ) ) )
      ( name = 'DOCUMENT_ITM'   range = VALUE #( ( sign = 'I' option = 'EQ' low = '001' ) ) ) ).
    mo_query->mt_requested = VALUE #( ( `ATTACHMENT` ) ).

    mo_cut->if_rap_query_provider~select( io_request = mo_query io_response = mo_query ).

    cl_abap_unit_assert=>assert_equals( act = lines( mo_query->mt_data ) exp = 1 ).
    " the file still contains the whole company code 1000: 2 customers = 2 sheets, 8 lines on the first one
    DATA(lo_excel) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( mo_query->mt_data[ 1 ]-attachment ).
    cl_abap_unit_assert=>assert_equals( act = lo_excel->get_worksheets_size( ) exp = 2 ).
    lo_excel->get_worksheet_by_name( '0100098012' )->get_cell( EXPORTING ip_column = 'A' ip_row = 22
                                                               IMPORTING ep_value  = DATA(lv_total) ).
    cl_abap_unit_assert=>assert_equals( act = lv_total exp = 'Tổng cộng' ).
  ENDMETHOD.

ENDCLASS.

CLASS ltc_autofilter DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.
  PRIVATE SECTION.
    DATA mo_sheet TYPE REF TO zcl_excel_worksheet.
    DATA mo_filter TYPE REF TO zcl_excel_autofilter.
    METHODS setup RAISING cx_static_check.
    METHODS pattern_to_excel FOR TESTING RAISING cx_static_check.
    METHODS pattern_from_excel FOR TESTING RAISING cx_static_check.
    METHODS text_filter_or FOR TESTING RAISING cx_static_check.
    METHODS text_filter_and FOR TESTING RAISING cx_static_check.
    METHODS hidden_rows RETURNING VALUE(rv_rows) TYPE string.
ENDCLASS.


CLASS ltc_autofilter IMPLEMENTATION.

  METHOD setup.
    DATA(lo_excel) = NEW zcl_excel( ).
    mo_sheet = lo_excel->get_active_worksheet( ).
    mo_sheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = `Product` ) ##NO_TEXT.
    mo_sheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = `Laptop Dell` ) ##NO_TEXT.
    mo_sheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = `Wireless Mouse` ) ##NO_TEXT.
    mo_sheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = `Keyboard` ) ##NO_TEXT.
    mo_sheet->set_cell( ip_column = 'A' ip_row = 5 ip_value = `Laptop Mouse` ) ##NO_TEXT.
    mo_filter = lo_excel->add_new_autofilter( mo_sheet ).
    mo_filter->set_filter_area( VALUE #( row_start = 1 col_start = 1 row_end = 5 col_end = 1 ) ).
  ENDMETHOD.

  METHOD hidden_rows.
    DO 5 TIMES.
      IF mo_filter->is_row_hidden( sy-index ) = abap_true.
        rv_rows = |{ rv_rows }{ sy-index }|.
      ENDIF.
    ENDDO.
  ENDMETHOD.

  METHOD pattern_to_excel.
    cl_abap_unit_assert=>assert_equals( act = zcl_excel_autofilter=>excel_pattern_from_filter( iv_option = 'CP' iv_low = 'Lap+op*' )
                                        exp = `Lap?op*` ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = zcl_excel_autofilter=>excel_pattern_from_filter( iv_option = 'CP' iv_low = 'C#+#+*' )
                                        exp = `C++*` ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = zcl_excel_autofilter=>excel_pattern_from_filter( iv_option = 'EQ' iv_low = 'Why?*~' )
                                        exp = `Why~?~*~~` ) ##NO_TEXT.
  ENDMETHOD.

  METHOD pattern_from_excel.
    cl_abap_unit_assert=>assert_equals( act = zcl_excel_autofilter=>filter_from_excel_pattern( `Lap?op*` )
                                        exp = `Lap+op*` ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = zcl_excel_autofilter=>filter_from_excel_pattern( `C++~*` )
                                        exp = `C#+#+#*` ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = zcl_excel_autofilter=>filter_from_excel_pattern( `Keyboard` )
                                        exp = `Keyboard` ) ##NO_TEXT.
  ENDMETHOD.

  METHOD text_filter_or.
    mo_filter->set_text_filter( i_column            = 1
                                iv_textfilter1      = 'Laptop*'
                                iv_logical_operator = zcl_excel_autofilter=>mc_logical_operator_or
                                iv_textfilter2      = '*Mouse' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = hidden_rows( ) exp = `4` ).
  ENDMETHOD.

  METHOD text_filter_and.
    mo_filter->set_text_filter( i_column            = 1
                                iv_textfilter1      = 'Laptop*'
                                iv_logical_operator = zcl_excel_autofilter=>mc_logical_operator_and
                                iv_textfilter2      = '*Mouse' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = hidden_rows( ) exp = `234` ).
    " single criterion (operator none): second filter is ignored
    mo_filter->set_text_filter( i_column = 1 iv_textfilter1 = 'Keyboard' iv_textfilter2 = 'x' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = hidden_rows( ) exp = `235` ).
  ENDMETHOD.

ENDCLASS.

CLASS ltc_richtext_merge DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.
  PRIVATE SECTION.
    METHODS same_text_different_runs FOR TESTING RAISING cx_static_check.
    METHODS merge_with_formula FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltc_richtext_merge IMPLEMENTATION.

  METHOD same_text_different_runs.
    " A4 / A5 / A6 of sheet RichText: same text, bold run / italic run / no runs
    DATA(lo_read) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( zcl_excel_demo_cloud_test=>build_richtext_merge_xlsx( ) ).
    DATA(lo_sheet) = lo_read->get_worksheet_by_name( 'RichText' ) ##NO_TEXT.
    lo_sheet->get_cell( EXPORTING ip_columnrow = 'A4' IMPORTING et_rtf = DATA(lt_rtf) ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = lines( lt_rtf ) exp = 2 ).
    cl_abap_unit_assert=>assert_true( lt_rtf[ 1 ]-font-bold ).
    lo_sheet->get_cell( EXPORTING ip_columnrow = 'A5' IMPORTING et_rtf = lt_rtf ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = lines( lt_rtf ) exp = 3 ).
    cl_abap_unit_assert=>assert_true( lt_rtf[ 2 ]-font-italic ).
    lo_sheet->get_cell( EXPORTING ip_columnrow = 'A6' IMPORTING et_rtf = lt_rtf ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_initial( lt_rtf ).
  ENDMETHOD.

  METHOD merge_with_formula.
    DATA(lo_read) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( zcl_excel_demo_cloud_test=>build_richtext_merge_xlsx( ) ).
    DATA(lo_sheet) = lo_read->get_worksheet_by_name( 'Merge' ) ##NO_TEXT.
    lo_sheet->get_cell( EXPORTING ip_columnrow = 'B7' IMPORTING ep_value = DATA(lv_value) ep_formula = DATA(lv_formula) ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = lv_formula exp = 'SUM(D7:D8)' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_initial( lv_value ).
    cl_abap_unit_assert=>assert_equals( act = lines( lo_sheet->get_merge( ) ) exp = 7 ).
  ENDMETHOD.

ENDCLASS.

CLASS ltc_images DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.
  PRIVATE SECTION.
    DATA mo_read TYPE REF TO zcl_excel.
    METHODS setup RAISING cx_static_check.
    METHODS absolute_anchor FOR TESTING RAISING cx_static_check.
    METHODS header_footer_images FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltc_images IMPLEMENTATION.

  METHOD setup.
    mo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( zcl_excel_demo_cloud_test=>build_image_drawing_xlsx( ) ).
  ENDMETHOD.

  METHOD absolute_anchor.
    DATA lv_found TYPE abap_bool.
    DATA(lo_iter) = mo_read->get_worksheet_by_name( 'Images' )->get_drawings( )->get_iterator( ) ##NO_TEXT.
    WHILE lo_iter->has_next( ) = abap_true.
      DATA(ls_pos) = CAST zcl_excel_drawing( lo_iter->get_next( ) )->get_position( ).
      IF ls_pos-anchor = zcl_excel_drawing=>anchor_absolute.
        lv_found = abap_true.
        " col 8 / row 7 of the demo -> 8 * 64 px and 7 * 20 px in EMU
        cl_abap_unit_assert=>assert_equals( act = ls_pos-from-col_offset exp = 4876800 ).
        cl_abap_unit_assert=>assert_equals( act = ls_pos-from-row_offset exp = 1333500 ).
        cl_abap_unit_assert=>assert_equals( act = ls_pos-size-width exp = 240 ).
      ENDIF.
    ENDWHILE.
    cl_abap_unit_assert=>assert_true( lv_found ).
  ENDMETHOD.

  METHOD header_footer_images.
    DATA(lo_sheet) = mo_read->get_worksheet_by_name( 'HeaderFooter' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = lines( lo_sheet->get_header_footer_drawings( ) ) exp = 2 ).
    lo_sheet->sheet_setup->get_header_footer_string( IMPORTING ep_odd_header = DATA(lv_header) ep_odd_footer = DATA(lv_footer) ).
    cl_abap_unit_assert=>assert_char_cp( act = lv_header exp = '&L&G&R*abap2xlsx - header image' ) ##NO_TEXT.
    " footer: text and image in the center section, right text kept
    cl_abap_unit_assert=>assert_char_cp( act = lv_footer exp = '&C*Ảnh:&G&R*Trang &P' ) ##NO_TEXT.
  ENDMETHOD.

ENDCLASS.

CLASS ltc_freeze_width DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.
  PRIVATE SECTION.
    DATA mo_read TYPE REF TO zcl_excel.
    METHODS setup RAISING cx_static_check.
    METHODS auto_size_and_default FOR TESTING RAISING cx_static_check.
    METHODS freeze_rows_only FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltc_freeze_width IMPLEMENTATION.

  METHOD setup.
    mo_read = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( zcl_excel_demo_cloud_test=>build_freeze_width_xlsx( ) ).
  ENDMETHOD.

  METHOD auto_size_and_default.
    DATA(lo_sheet) = mo_read->get_worksheet_by_name( 'Widths' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = lo_sheet->get_default_column( )->get_width( ) exp = 15 ).
    " column C: longest text (42 characters) -> 42 + 0.75 padding
    cl_abap_unit_assert=>assert_equals( act = round( val = CONV decfloat34( lo_sheet->get_column( 'C' )->get_width( ) ) dec = 2 )
                                        exp = CONV decfloat34( '42.75' ) ).
    " fixed width stays
    cl_abap_unit_assert=>assert_equals( act = lo_sheet->get_column( 'B' )->get_width( ) exp = 40 ).
  ENDMETHOD.

  METHOD freeze_rows_only.
    DATA(lo_sheet) = mo_read->get_worksheet_by_name( 'Freeze_Rows3' ) ##NO_TEXT.
    lo_sheet->get_freeze_cell( IMPORTING ep_row = DATA(lv_row) ep_column = DATA(lv_column) ).
    cl_abap_unit_assert=>assert_equals( act = lv_row exp = 4 ).
    cl_abap_unit_assert=>assert_equals( act = lv_column exp = 1 ).
  ENDMETHOD.

ENDCLASS.

CLASS ltc_xlsm DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.
  PRIVATE SECTION.
    METHODS package_parts FOR TESTING RAISING cx_static_check.
    METHODS read_vba_project FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltc_xlsm IMPLEMENTATION.

  METHOD package_parts.
    DATA lv_rels TYPE xstring.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( zcl_excel_demo_cloud_test=>build_xlsm( ) ).
    lo_zip->get( EXPORTING name = 'xl/vbaProject.bin' IMPORTING content = DATA(lv_vba) ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = lv_vba exp = zcl_excel_demo_cloud_test=>get_test_vba_project( ) ).
    " a sheet without drawings must not reference xl/drawings/drawing1.xml
    lo_zip->get( EXPORTING name = 'xl/worksheets/_rels/sheet1.xml.rels' IMPORTING content = lv_rels ) ##NO_TEXT.
    DATA(lv_rels_text) = cl_abap_conv_codepage=>create_in( )->convert( lv_rels ).
    cl_abap_unit_assert=>assert_char_np( act = lv_rels_text exp = '*drawing*' ) ##NO_TEXT.
  ENDMETHOD.

  METHOD read_vba_project.
    DATA(lo_read) = NEW zcl_excel_reader_xlsm( )->zif_excel_reader~load( zcl_excel_demo_cloud_test=>build_xlsm( ) ).
    cl_abap_unit_assert=>assert_equals( act = lo_read->zif_excel_book_vba_project~codename_pr exp = 'ThisWorkbook' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = lo_read->get_active_worksheet( )->zif_excel_sheet_vba_project~codename_pr exp = 'Sheet1' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_equals( act = xstrlen( lo_read->zif_excel_book_vba_project~vbaproject ) exp = 15872 ).
  ENDMETHOD.

ENDCLASS.

CLASS ltc_excel_file_roundtrip DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.
  PRIVATE SECTION.
    DATA mo_zip TYPE REF TO cl_abap_zip.
    METHODS setup RAISING cx_static_check.
    METHODS part_text IMPORTING iv_name TYPE string RETURNING VALUE(rv_text) TYPE string RAISING cx_static_check.
    METHODS external_link_kept FOR TESTING RAISING cx_static_check.
    METHODS chart_parts_kept FOR TESTING RAISING cx_static_check.
    METHODS no_empty_cols FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltc_excel_file_roundtrip IMPLEMENTATION.

  METHOD setup.
    " workbook created by Microsoft Excel: read and write again
    DATA(lo_excel) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( zcl_excel_demo_cloud_test=>get_excel_created_file( ) ).
    cl_abap_unit_assert=>assert_equals( act = lines( lo_excel->external_links ) exp = 1 ).
    mo_zip = NEW cl_abap_zip( ).
    mo_zip->load( CAST zif_excel_writer( NEW zcl_excel_writer_2007( ) )->write_file( lo_excel ) ).
  ENDMETHOD.

  METHOD part_text.
    mo_zip->get( EXPORTING name = iv_name IMPORTING content = DATA(lv_content) ).
    rv_text = cl_abap_conv_codepage=>create_in( )->convert( lv_content ).
  ENDMETHOD.

  METHOD external_link_kept.
    cl_abap_unit_assert=>assert_char_cp( act = part_text( `xl/externalLinks/externalLink1.xml` ) exp = '*<sheetName val="Bang"/>*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_cp( act = part_text( `xl/workbook.xml` ) exp = '*<externalReferences><externalReference r:id="rIdExt1"/></externalReferences>*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_cp( act = part_text( `xl/_rels/workbook.xml.rels` ) exp = '*Id="rIdExt1"*Target="externalLinks/externalLink1.xml"*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_cp( act = part_text( `[Content_Types].xml` ) exp = '*/xl/externalLinks/externalLink1.xml*' ) ##NO_TEXT.
  ENDMETHOD.

  METHOD chart_parts_kept.
    cl_abap_unit_assert=>assert_char_cp( act = part_text( `xl/charts/_rels/chart1.xml.rels` ) exp = '*chartStyle*style1.xml*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_cp( act = part_text( `xl/charts/style1.xml` ) exp = '*chartStyle*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_cp( act = part_text( `[Content_Types].xml` ) exp = '*/xl/charts/colors1.xml*chartcolorstyle*' ) ##NO_TEXT.
  ENDMETHOD.

  METHOD no_empty_cols.
    " sheet "Bang" has no column settings in the Excel file -> no <cols> (previously width 9.10 for A:C)
    cl_abap_unit_assert=>assert_char_np( act = part_text( `xl/worksheets/sheet2.xml` ) exp = '*<cols>*' ) ##NO_TEXT.
  ENDMETHOD.

ENDCLASS.

CLASS ltc_writer_huge_file DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.
  PRIVATE SECTION.
    METHODS sheet_xml
      IMPORTING io_excel       TYPE REF TO zcl_excel
                iv_sheet       TYPE i DEFAULT 1
      RETURNING VALUE(rv_xml)  TYPE string
      RAISING   cx_static_check.
    METHODS formulas_and_types FOR TESTING RAISING cx_static_check.
    METHODS header_footer_no_dump FOR TESTING RAISING cx_static_check.
    METHODS column_formula_of_table FOR TESTING RAISING cx_static_check.
    METHODS sheet_parts FOR TESTING RAISING cx_static_check.
    METHODS rows_without_cells FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltc_writer_huge_file IMPLEMENTATION.

  METHOD sheet_xml.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( CAST zif_excel_writer( NEW zcl_excel_writer_huge_file( ) )->write_file( io_excel ) ).
    lo_zip->get( EXPORTING name = |xl/worksheets/sheet{ iv_sheet }.xml| IMPORTING content = DATA(lv_content) ) ##NO_TEXT.
    rv_xml = cl_abap_conv_codepage=>create_in( )->convert( lv_content ).
  ENDMETHOD.

  METHOD formulas_and_types.
    DATA(lo_excel) = NEW zcl_excel( ).
    DATA(lo_sheet) = lo_excel->get_active_worksheet( ).
    lo_sheet->set_cell( ip_columnrow = 'A1' ip_value = 10 ) ##NO_TEXT.
    lo_sheet->set_cell( ip_columnrow = 'B1' ip_formula = 'A1*2' ) ##NO_TEXT.
    DATA(lv_xml) = sheet_xml( lo_excel ).
    cl_abap_unit_assert=>assert_char_cp( act = lv_xml exp = '*<c r="B1"*><f>A1*2</f></c>*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_np( act = lv_xml exp = '*t=""*' ) ##NO_TEXT.
  ENDMETHOD.

  METHOD header_footer_no_dump.
    " the transformation raised CX_ST_REF_ACCESS for every header / footer
    DATA(lo_excel) = NEW zcl_excel( ).
    DATA(lo_sheet) = lo_excel->get_active_worksheet( ).
    lo_sheet->set_cell( ip_columnrow = 'A1' ip_value = `x` ) ##NO_TEXT.
    lo_sheet->sheet_setup->set_header_footer( ip_odd_header = VALUE #( center_value = `Header` ) ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_cp( act = sheet_xml( lo_excel ) exp = '*<oddHeader>&amp;C*Header</oddHeader>*' ) ##NO_TEXT.
  ENDMETHOD.

  METHOD column_formula_of_table.
    TYPES: BEGIN OF ty_s_line,
             qty   TYPE i,
             total TYPE i,
           END OF ty_s_line.
    DATA lt_lines TYPE STANDARD TABLE OF ty_s_line WITH EMPTY KEY.
    lt_lines = VALUE #( ( qty = 1 ) ( qty = 2 ) ).
    DATA(lo_excel) = NEW zcl_excel( ).
    DATA(lo_sheet) = lo_excel->get_active_worksheet( ).
    DATA(lt_fieldcat) = zcl_excel_common=>get_fieldcatalog( ip_table = lt_lines ).
    LOOP AT lt_fieldcat ASSIGNING FIELD-SYMBOL(<ls_fc>) WHERE fieldname = 'TOTAL'.
      <ls_fc>-column_formula = 'A2*10' ##NO_TEXT.
    ENDLOOP.
    lo_sheet->bind_table( ip_table = lt_lines it_field_catalog = lt_fieldcat
                          is_table_settings = VALUE #( top_left_column = 'A' top_left_row = 1 ) ) ##NO_TEXT.
    DATA(lv_xml) = sheet_xml( lo_excel ).
    cl_abap_unit_assert=>assert_char_cp( act = lv_xml exp = '*<c r="B3"*><f>A3*10</f>*' ) ##NO_TEXT.
  ENDMETHOD.

  METHOD sheet_parts.
    " autofilter, conditional formatting, data validation and comments come from ZCL_EXCEL_WRITER_2007,
    " inserted at the position of the OOXML schema
    DATA(lo_excel) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( zcl_excel_demo_cloud_test=>build_autofilter_table_xlsx( ) ).
    cl_abap_unit_assert=>assert_char_cp( act = sheet_xml( lo_excel ) exp = '*<sheetPr filterMode="1">*</sheetData><autoFilter ref="A1:D13">*' ) ##NO_TEXT.
    lo_excel = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( zcl_excel_demo_cloud_test=>build_condformat_link_xlsx( ) ).
    cl_abap_unit_assert=>assert_char_cp( act = sheet_xml( lo_excel ) exp = '*</sheetData>*<conditionalFormatting*<pageMargins*' ) ##NO_TEXT.
    lo_excel = NEW zcl_excel_reader_2007( )->zif_excel_reader~load( zcl_excel_demo_cloud_test=>build_comment_outline_xlsx( ) ).
    cl_abap_unit_assert=>assert_char_cp( act = sheet_xml( lo_excel ) exp = '*outlineLevel="2"*<legacyDrawing r:id=*' ) ##NO_TEXT.
  ENDMETHOD.

  METHOD rows_without_cells.
    DATA(lo_excel) = NEW zcl_excel( ).
    DATA(lo_sheet) = lo_excel->get_active_worksheet( ).
    lo_sheet->set_cell( ip_columnrow = 'A1' ip_value = `x` ) ##NO_TEXT.
    lo_sheet->set_cell( ip_columnrow = 'A10' ip_value = `y` ) ##NO_TEXT.
    lo_sheet->get_row( 3 )->set_row_height( '25.5' ).                     " empty row with height
    lo_sheet->get_row( 4 )->set_visible( abap_false ).                   " empty hidden row
    lo_sheet->set_row_outline( iv_row_from = 6 iv_row_to = 7 iv_collapsed = abap_true ).
    lo_sheet->get_column( 'A' )->set_width( '12.5' ).
    DATA(lv_xml) = sheet_xml( lo_excel ).
    cl_abap_unit_assert=>assert_char_cp( act = lv_xml exp = '*<row r="3"*customHeight="1" ht="25.5"*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_cp( act = lv_xml exp = '*<row r="4"*hidden="true"*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_cp( act = lv_xml exp = '*<row r="6"*outlineLevel="1"*<row r="8"*collapsed="true"*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_cp( act = lv_xml exp = '*<col min="1" max="1" width="12.5"*' ) ##NO_TEXT.
    cl_abap_unit_assert=>assert_char_np( act = lv_xml exp = '*E+0*' ) ##NO_TEXT.
    " rows are in ascending order
    cl_abap_unit_assert=>assert_char_cp( act = lv_xml exp = '*<row r="1"*<row r="3"*<row r="4"*<row r="10"*' ) ##NO_TEXT.
  ENDMETHOD.

ENDCLASS.

