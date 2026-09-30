*"* use this source file for your ABAP unit test classes

CLASS ltc_unescape_string_value DEFINITION DEFERRED.
CLASS ltc_header_footer DEFINITION DEFERRED.
CLASS zcl_excel_reader_2007 DEFINITION LOCAL FRIENDS
    ltc_unescape_string_value ltc_header_footer.


CLASS ltc_unescape_string_value DEFINITION
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.
  PRIVATE SECTION.

    METHODS escaped_character_inside_text FOR TESTING.
    METHODS no_escaping FOR TESTING.
    METHODS one_escaped_character FOR TESTING.
    METHODS two_escaped_characters FOR TESTING.
    METHODS skip_when_not_escaped FOR TESTING RAISING cx_static_check.

    METHODS run_cut
      IMPORTING
        input TYPE string
        exp   TYPE string.

ENDCLASS.


CLASS ltc_unescape_string_value IMPLEMENTATION.

  METHOD escaped_character_inside_text.
    run_cut( input = 'start _x0000_ end' exp = |start { zcl_excel_common=>unicode_char_from_hex( '0000' ) } end| ).
  ENDMETHOD.

  METHOD no_escaping.
    run_cut( input = 'no escaping' exp = 'no escaping' ).
  ENDMETHOD.

  METHOD one_escaped_character.
    run_cut( input = '_x0000_' exp = zcl_excel_common=>unicode_char_from_hex( '0000' ) ).
  ENDMETHOD.

  METHOD run_cut.

    DATA: lo_excel TYPE REF TO zcl_excel_reader_2007.

    CREATE OBJECT lo_excel.

    cl_abap_unit_assert=>assert_equals( act = lo_excel->unescape_string_value( input ) exp = exp msg = |input: { input }| ).

  ENDMETHOD.

  METHOD two_escaped_characters.
    run_cut( input = '_x0000_ and _xFFFF_' exp = |{ zcl_excel_common=>unicode_char_from_hex( '0000' ) } and { zcl_excel_common=>unicode_char_from_hex( 'FFFF' ) }| ).
  ENDMETHOD.

  METHOD skip_when_not_escaped.
    DATA: lo_excel TYPE REF TO zcl_excel_reader_2007,
          value TYPE string VALUE 'TEST_X'.

    CREATE OBJECT lo_excel.

    "Method is used to check for "_x", but its not an escaped charcater, output should input.
    lo_excel->unescape_string_value( i_value = value ).

    cl_abap_unit_assert=>assert_equals(
    exp = value
    act = value ).
  ENDMETHOD.

ENDCLASS.


CLASS ltc_header_footer DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.
  PRIVATE SECTION.
    METHODS three_sections_with_fonts FOR TESTING.
    METHODS center_only FOR TESTING.
    METHODS plain_text_is_center FOR TESTING.
    METHODS font_name_and_ampersand FOR TESTING.
ENDCLASS.


CLASS ltc_header_footer IMPLEMENTATION.

  METHOD three_sections_with_fonts.
    DATA(ls_hf) = zcl_excel_reader_2007=>parse_header_footer(
      `&L&"-,Standard"&A&C&"-,Bold"&14BÁO CÁO&R&"-,Standard"In ngày &D` ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-left_value   exp = `&A` ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-center_value exp = `BÁO CÁO` ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-center_font-bold exp = abap_true ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-center_font-size exp = 14 ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-right_value  exp = `In ngày &D` ).
    cl_abap_unit_assert=>assert_initial( ls_hf-left_font-bold ).
  ENDMETHOD.

  METHOD center_only.
    DATA(ls_hf) = zcl_excel_reader_2007=>parse_header_footer( `&C&"-,Standard"Trang &P / &N` ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-center_value exp = `Trang &P / &N` ).
    cl_abap_unit_assert=>assert_initial( ls_hf-left_value ).
    cl_abap_unit_assert=>assert_initial( ls_hf-right_value ).
  ENDMETHOD.

  METHOD plain_text_is_center.
    DATA(ls_hf) = zcl_excel_reader_2007=>parse_header_footer( `Confidential` ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-center_value exp = `Confidential` ).
  ENDMETHOD.

  METHOD font_name_and_ampersand.
    DATA(ls_hf) = zcl_excel_reader_2007=>parse_header_footer( `&L&"Arial,Bold Italic"&10A&&B` ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-left_value exp = `A&&B` ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-left_font-name exp = `Arial` ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-left_font-bold exp = abap_true ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-left_font-italic exp = abap_true ).
    cl_abap_unit_assert=>assert_equals( act = ls_hf-left_font-size exp = 10 ).
  ENDMETHOD.

ENDCLASS.

