CLASS zcl_excel_sheet_setup DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_SHEET_SETUP
*"* do not include other source files here!!!

    DATA black_and_white TYPE abap_bool .
    DATA cell_comments TYPE string .
    DATA copies TYPE int2 .
    CONSTANTS c_break_column TYPE zexcel_break VALUE 2.     "#EC NOTEXT
    CONSTANTS c_break_none TYPE zexcel_break VALUE 0.       "#EC NOTEXT
    CONSTANTS c_break_row TYPE zexcel_break VALUE 1.        "#EC NOTEXT
    CONSTANTS c_cc_as_displayed TYPE string VALUE 'asDisplayed'. "#EC NOTEXT
    CONSTANTS c_cc_at_end TYPE string VALUE 'atEnd'.        "#EC NOTEXT
    CONSTANTS c_cc_none TYPE string VALUE 'none'.           "#EC NOTEXT
    CONSTANTS c_ord_downthenover TYPE string VALUE 'downThenOver'. "#EC NOTEXT
    CONSTANTS c_ord_overthendown TYPE string VALUE 'overThenDown'. "#EC NOTEXT
    CONSTANTS c_orientation_default TYPE zexcel_sheet_orienatation VALUE 'default'. "#EC NOTEXT
    CONSTANTS c_orientation_landscape TYPE zexcel_sheet_orienatation VALUE 'landscape'. "#EC NOTEXT
    CONSTANTS c_orientation_portrait TYPE zexcel_sheet_orienatation VALUE 'portrait'. "#EC NOTEXT
    CONSTANTS c_papersize_6_3_4_envelope TYPE zexcel_sheet_paper_size VALUE 38. "#EC NOTEXT
    CONSTANTS c_papersize_a2_paper TYPE zexcel_sheet_paper_size VALUE 64. "#EC NOTEXT
    CONSTANTS c_papersize_a3 TYPE zexcel_sheet_paper_size VALUE 8. "#EC NOTEXT
    CONSTANTS c_papersize_a3_extra_paper TYPE zexcel_sheet_paper_size VALUE 61. "#EC NOTEXT
    CONSTANTS c_papersize_a3_extra_tv_paper TYPE zexcel_sheet_paper_size VALUE 66. "#EC NOTEXT
    CONSTANTS c_papersize_a3_tv_paper TYPE zexcel_sheet_paper_size VALUE 65. "#EC NOTEXT
    CONSTANTS c_papersize_a4 TYPE zexcel_sheet_paper_size VALUE 9. "#EC NOTEXT
    CONSTANTS c_papersize_a4_extra_paper TYPE zexcel_sheet_paper_size VALUE 51. "#EC NOTEXT
    CONSTANTS c_papersize_a4_plus_paper TYPE zexcel_sheet_paper_size VALUE 58. "#EC NOTEXT
    CONSTANTS c_papersize_a4_small TYPE zexcel_sheet_paper_size VALUE 10. "#EC NOTEXT
    CONSTANTS c_papersize_a4_tv_paper TYPE zexcel_sheet_paper_size VALUE 53. "#EC NOTEXT
    CONSTANTS c_papersize_a5 TYPE zexcel_sheet_paper_size VALUE 11. "#EC NOTEXT
    CONSTANTS c_papersize_a5_extra_paper TYPE zexcel_sheet_paper_size VALUE 62. "#EC NOTEXT
    CONSTANTS c_papersize_a5_tv_paper TYPE zexcel_sheet_paper_size VALUE 59. "#EC NOTEXT
    CONSTANTS c_papersize_b4 TYPE zexcel_sheet_paper_size VALUE 12. "#EC NOTEXT
    CONSTANTS c_papersize_b4_envelope TYPE zexcel_sheet_paper_size VALUE 33. "#EC NOTEXT
    CONSTANTS c_papersize_b5 TYPE zexcel_sheet_paper_size VALUE 13. "#EC NOTEXT
    CONSTANTS c_papersize_b5_envelope TYPE zexcel_sheet_paper_size VALUE 34. "#EC NOTEXT
    CONSTANTS c_papersize_b6_envelope TYPE zexcel_sheet_paper_size VALUE 35. "#EC NOTEXT
    CONSTANTS c_papersize_c TYPE zexcel_sheet_paper_size VALUE 24. "#EC NOTEXT
    CONSTANTS c_papersize_c3_envelope TYPE zexcel_sheet_paper_size VALUE 29. "#EC NOTEXT
    CONSTANTS c_papersize_c4_envelope TYPE zexcel_sheet_paper_size VALUE 30. "#EC NOTEXT
    CONSTANTS c_papersize_c5_envelope TYPE zexcel_sheet_paper_size VALUE 28. "#EC NOTEXT
    CONSTANTS c_papersize_c65_envelope TYPE zexcel_sheet_paper_size VALUE 32. "#EC NOTEXT
    CONSTANTS c_papersize_c6_envelope TYPE zexcel_sheet_paper_size VALUE 31. "#EC NOTEXT
    CONSTANTS c_papersize_d TYPE zexcel_sheet_paper_size VALUE 25. "#EC NOTEXT
    CONSTANTS c_papersize_de_leg_fanfold TYPE zexcel_sheet_paper_size VALUE 41. "#EC NOTEXT
    CONSTANTS c_papersize_de_std_fanfold TYPE zexcel_sheet_paper_size VALUE 40. "#EC NOTEXT
    CONSTANTS c_papersize_dl_envelope TYPE zexcel_sheet_paper_size VALUE 27. "#EC NOTEXT
    CONSTANTS c_papersize_e TYPE zexcel_sheet_paper_size VALUE 26. "#EC NOTEXT
    CONSTANTS c_papersize_executive TYPE zexcel_sheet_paper_size VALUE 7. "#EC NOTEXT
    CONSTANTS c_papersize_folio TYPE zexcel_sheet_paper_size VALUE 14. "#EC NOTEXT
    CONSTANTS c_papersize_invite_envelope TYPE zexcel_sheet_paper_size VALUE 47. "#EC NOTEXT
    CONSTANTS c_papersize_iso_b4 TYPE zexcel_sheet_paper_size VALUE 42. "#EC NOTEXT
    CONSTANTS c_papersize_iso_b5_extra_paper TYPE zexcel_sheet_paper_size VALUE 63. "#EC NOTEXT
    CONSTANTS c_papersize_italy_envelope TYPE zexcel_sheet_paper_size VALUE 36. "#EC NOTEXT
    CONSTANTS c_papersize_jis_b5_tv_paper TYPE zexcel_sheet_paper_size VALUE 60. "#EC NOTEXT
    CONSTANTS c_papersize_jpn_dbl_postcard TYPE zexcel_sheet_paper_size VALUE 43. "#EC NOTEXT
    CONSTANTS c_papersize_ledger TYPE zexcel_sheet_paper_size VALUE 4. "#EC NOTEXT
    CONSTANTS c_papersize_legal TYPE zexcel_sheet_paper_size VALUE 5. "#EC NOTEXT
    CONSTANTS c_papersize_legal_extra_paper TYPE zexcel_sheet_paper_size VALUE 49. "#EC NOTEXT
    CONSTANTS c_papersize_letter TYPE zexcel_sheet_paper_size VALUE 1. "#EC NOTEXT
    CONSTANTS c_papersize_letter_extra_paper TYPE zexcel_sheet_paper_size VALUE 48. "#EC NOTEXT
    CONSTANTS c_papersize_letter_extv_paper TYPE zexcel_sheet_paper_size VALUE 54. "#EC NOTEXT
    CONSTANTS c_papersize_letter_plus_paper TYPE zexcel_sheet_paper_size VALUE 57. "#EC NOTEXT
    CONSTANTS c_papersize_letter_small TYPE zexcel_sheet_paper_size VALUE 2. "#EC NOTEXT
    CONSTANTS c_papersize_letter_tv_paper TYPE zexcel_sheet_paper_size VALUE 52. "#EC NOTEXT
    CONSTANTS c_papersize_monarch_envelope TYPE zexcel_sheet_paper_size VALUE 37. "#EC NOTEXT
    CONSTANTS c_papersize_no10_envelope TYPE zexcel_sheet_paper_size VALUE 20. "#EC NOTEXT
    CONSTANTS c_papersize_no11_envelope TYPE zexcel_sheet_paper_size VALUE 21. "#EC NOTEXT
    CONSTANTS c_papersize_no12_envelope TYPE zexcel_sheet_paper_size VALUE 22. "#EC NOTEXT
    CONSTANTS c_papersize_no14_envelope TYPE zexcel_sheet_paper_size VALUE 23. "#EC NOTEXT
    CONSTANTS c_papersize_no9_envelope TYPE zexcel_sheet_paper_size VALUE 19. "#EC NOTEXT
    CONSTANTS c_papersize_note TYPE zexcel_sheet_paper_size VALUE 18. "#EC NOTEXT
    CONSTANTS c_papersize_quarto TYPE zexcel_sheet_paper_size VALUE 15. "#EC NOTEXT
    CONSTANTS c_papersize_standard_1 TYPE zexcel_sheet_paper_size VALUE 16. "#EC NOTEXT
    CONSTANTS c_papersize_standard_2 TYPE zexcel_sheet_paper_size VALUE 17. "#EC NOTEXT
    CONSTANTS c_papersize_standard_paper_1 TYPE zexcel_sheet_paper_size VALUE 44. "#EC NOTEXT
    CONSTANTS c_papersize_standard_paper_2 TYPE zexcel_sheet_paper_size VALUE 45. "#EC NOTEXT
    CONSTANTS c_papersize_standard_paper_3 TYPE zexcel_sheet_paper_size VALUE 46. "#EC NOTEXT
    CONSTANTS c_papersize_statement TYPE zexcel_sheet_paper_size VALUE 6. "#EC NOTEXT
    CONSTANTS c_papersize_supera_a4_paper TYPE zexcel_sheet_paper_size VALUE 55. "#EC NOTEXT
    CONSTANTS c_papersize_superb_a3_paper TYPE zexcel_sheet_paper_size VALUE 56. "#EC NOTEXT
    CONSTANTS c_papersize_tabloid TYPE zexcel_sheet_paper_size VALUE 3. "#EC NOTEXT
    CONSTANTS c_papersize_tabl_extra_paper TYPE zexcel_sheet_paper_size VALUE 50. "#EC NOTEXT
    CONSTANTS c_papersize_us_std_fanfold TYPE zexcel_sheet_paper_size VALUE 39. "#EC NOTEXT
    CONSTANTS c_pe_blank TYPE string VALUE 'blank'.         "#EC NOTEXT
    CONSTANTS c_pe_dash TYPE string VALUE 'dash'.           "#EC NOTEXT
    CONSTANTS c_pe_displayed TYPE string VALUE 'displayed'. "#EC NOTEXT
    CONSTANTS c_pe_na TYPE string VALUE 'NA'.               "#EC NOTEXT
    DATA diff_oddeven_headerfooter TYPE abap_bool .
    DATA draft TYPE abap_bool .
    DATA errors TYPE string .
    DATA even_footer TYPE zexcel_s_worksheet_head_foot .
    DATA even_header TYPE zexcel_s_worksheet_head_foot .
    DATA first_page_number TYPE int2 .
    DATA fit_to_height TYPE int2 .
    DATA fit_to_page TYPE abap_bool .
    DATA fit_to_width TYPE int2 .
    DATA horizontal_centered TYPE abap_bool .
    DATA horizontal_dpi TYPE int2 .
    DATA margin_bottom TYPE zexcel_dec_8_2 .
    DATA margin_footer TYPE zexcel_dec_8_2 .
    DATA margin_header TYPE zexcel_dec_8_2 .
    DATA margin_left TYPE zexcel_dec_8_2 .
    DATA margin_right TYPE zexcel_dec_8_2 .
    DATA margin_top TYPE zexcel_dec_8_2 .
    DATA odd_footer TYPE zexcel_s_worksheet_head_foot .
    DATA odd_header TYPE zexcel_s_worksheet_head_foot .
    DATA orientation TYPE zexcel_sheet_orienatation .
    DATA page_order TYPE string .
    DATA paper_height TYPE string .
    DATA paper_size TYPE int2 .
    DATA paper_width TYPE string .
    DATA scale TYPE int2 .
    DATA use_first_page_num TYPE abap_bool .
    DATA use_printer_defaults TYPE abap_bool .
    DATA vertical_centered TYPE abap_bool .
    DATA vertical_dpi TYPE int2 .

    METHODS constructor .
    METHODS set_page_margins
      IMPORTING
        !ip_bottom TYPE f OPTIONAL
        !ip_footer TYPE f OPTIONAL
        !ip_header TYPE f OPTIONAL
        !ip_left   TYPE f OPTIONAL
        !ip_right  TYPE f OPTIONAL
        !ip_top    TYPE f OPTIONAL
        !ip_unit   TYPE csequence DEFAULT 'in' .
    METHODS set_header_footer
      IMPORTING
        !ip_odd_header  TYPE zexcel_s_worksheet_head_foot OPTIONAL
        !ip_odd_footer  TYPE zexcel_s_worksheet_head_foot OPTIONAL
        !ip_even_header TYPE zexcel_s_worksheet_head_foot OPTIONAL
        !ip_even_footer TYPE zexcel_s_worksheet_head_foot OPTIONAL .
    METHODS get_header_footer_string
      EXPORTING
        !ep_odd_header  TYPE string
        !ep_odd_footer  TYPE string
        !ep_even_header TYPE string
        !ep_even_footer TYPE string .
    METHODS get_header_footer
      EXPORTING
        !ep_odd_header  TYPE zexcel_s_worksheet_head_foot
        !ep_odd_footer  TYPE zexcel_s_worksheet_head_foot
        !ep_even_header TYPE zexcel_s_worksheet_head_foot
        !ep_even_footer TYPE zexcel_s_worksheet_head_foot .
  PROTECTED SECTION.

*"* protected components of class ZCL_EXCEL_SHEET_SETUP
*"* do not include other source files here!!!
    METHODS process_header_footer
      IMPORTING
        !ip_header                 TYPE zexcel_s_worksheet_head_foot
        !ip_side                   TYPE string
      RETURNING
        VALUE(rv_processed_string) TYPE string .
    "! Build the OOXML string of one header / footer: &L, &C, &R sections with text and image (&G)
    METHODS build_header_footer_string
      IMPORTING
        !is_header_footer TYPE zexcel_s_worksheet_head_foot
      RETURNING
        VALUE(rv_string)  TYPE string .
  PRIVATE SECTION.
*"* private components of class ZCL_EXCEL_SHEET_SETUP
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_SHEET_SETUP IMPLEMENTATION.


  METHOD constructor.
    orientation = me->c_orientation_default.

* default margins
    margin_bottom = '0.75'.
    margin_footer = '0.3'.
    margin_header = '0.3'.
    margin_left   = '0.7'.
    margin_right  = '0.7'.
    margin_top    = '0.75'.

* clear page settings
    CLEAR: black_and_white,
           cell_comments,
           copies,
           draft,
           errors,
           first_page_number,
           fit_to_page,
           fit_to_height,
           fit_to_width,
           horizontal_dpi,
           orientation,
           page_order,
           paper_height,
           paper_size,
           paper_width,
           scale,
           use_first_page_num,
           use_printer_defaults,
           vertical_dpi.
  ENDMETHOD.


  METHOD get_header_footer.

* Only Basic font/text formatting possible:
* Bold (yes / no), Font Type, Font Size
*
* usefull placeholders, which can be used in header/footer value strings
* '&P' - page number
* '&N' - total number of pages
* '&D' - Date
* '&T' - Time
* '&F' - File Name
* '&Z' - Path
* '&A' - Sheet name
* new line via class constant CL_ABAP_CHAR_UTILITIES=>newline
*
* Example Value String 'page &P of &N'
*
* DO NOT USE &L , &C or &R which automatically created as position markers

    ep_odd_header = me->odd_header.
    ep_odd_footer = me->odd_footer.
    ep_even_header = me->even_header.
    ep_even_footer = me->even_footer.

  ENDMETHOD.


  METHOD get_header_footer_string.
* ----------------------------------------------------------------------
* each section is written once: &L<text>&G &C<text>&G &R<text>&G
* (previously the image markers were appended to EP_ODD_HEADER, so the footer and the even
*  header / footer lost their text and the section marker appeared twice)
* ----------------------------------------------------------------------
    IF ep_odd_header IS SUPPLIED.
      ep_odd_header = build_header_footer_string( me->odd_header ).
    ENDIF.
    IF ep_odd_footer IS SUPPLIED.
      ep_odd_footer = build_header_footer_string( me->odd_footer ).
    ENDIF.
    IF ep_even_header IS SUPPLIED.
      ep_even_header = build_header_footer_string( me->even_header ).
    ENDIF.
    IF ep_even_footer IS SUPPLIED.
      ep_even_footer = build_header_footer_string( me->even_footer ).
    ENDIF.
  ENDMETHOD.


  METHOD build_header_footer_string.
    TYPES: BEGIN OF ty_s_section,
             side   TYPE string,
             marker TYPE string,
           END OF ty_s_section.
    DATA: lt_sections TYPE STANDARD TABLE OF ty_s_section WITH EMPTY KEY,
          lv_section  TYPE string,
          lv_value    TYPE string,
          lv_fname    TYPE string.
    FIELD-SYMBOLS: <lv_text>  TYPE string,
                   <lo_image> TYPE any.

    lt_sections = VALUE #( ( side = `LEFT`   marker = `&L` )
                           ( side = `CENTER` marker = `&C` )
                           ( side = `RIGHT`  marker = `&R` ) ) ##NO_TEXT.
    LOOP AT lt_sections INTO DATA(ls_section).
      CLEAR lv_section.
      lv_fname = |{ ls_section-side }_VALUE|.
      ASSIGN COMPONENT lv_fname OF STRUCTURE is_header_footer TO <lv_text>.
      IF <lv_text> IS ASSIGNED AND <lv_text> IS NOT INITIAL.
        lv_section = me->process_header_footer( ip_header = is_header_footer ip_side = ls_section-side ).
      ENDIF.
      lv_fname = |{ ls_section-side }_IMAGE|.
      ASSIGN COMPONENT lv_fname OF STRUCTURE is_header_footer TO <lo_image>.
      IF <lo_image> IS ASSIGNED AND <lo_image> IS NOT INITIAL AND lv_section NS '&G'.
        lv_section = lv_section && `&G`.
      ENDIF.
      IF lv_section IS NOT INITIAL.
        rv_string = rv_string && ls_section-marker && lv_section.
      ENDIF.
      UNASSIGN: <lv_text>, <lo_image>.
    ENDLOOP.
  ENDMETHOD.


  METHOD process_header_footer.

* ----------------------------------------------------------------------
* Only Basic font/text formatting possible:
* Bold (yes / no), Font Type, Font Size

    DATA:   lv_fname(12) TYPE c
          , lv_string    TYPE string
          .

    FIELD-SYMBOLS:   <lv_value> TYPE string
                   , <ls_font>  TYPE zexcel_s_style_font
                   .

* ----------------------------------------------------------------------
    CONCATENATE ip_side '_VALUE' INTO lv_fname.
    ASSIGN COMPONENT lv_fname OF STRUCTURE ip_header TO <lv_value>.

    CONCATENATE ip_side '_FONT' INTO lv_fname.
    ASSIGN COMPONENT lv_fname OF STRUCTURE ip_header TO <ls_font>.

    IF <ls_font> IS ASSIGNED AND <lv_value> IS ASSIGNED.

      IF <lv_value> = '&G'. "image header
        rv_processed_string = <lv_value>.
      ELSE.

        IF <ls_font>-name IS NOT INITIAL.
          CONCATENATE '&"' <ls_font>-name ',' INTO rv_processed_string.
        ELSE.
          rv_processed_string = '&"-,'.
        ENDIF.

        IF <ls_font>-bold = abap_true.
          CONCATENATE rv_processed_string 'Bold"' INTO rv_processed_string ##NO_TEXT.
        ELSE.
          CONCATENATE rv_processed_string 'Standard"' INTO rv_processed_string ##NO_TEXT.
        ENDIF.

        IF <ls_font>-size IS NOT INITIAL.
          lv_string = <ls_font>-size.
          CONDENSE lv_string NO-GAPS.
          CONCATENATE rv_processed_string '&' lv_string INTO rv_processed_string.
        ENDIF.

        CONCATENATE rv_processed_string <lv_value> INTO rv_processed_string.
      ENDIF.
    ENDIF.
* ----------------------------------------------------------------------

  ENDMETHOD.


  METHOD set_header_footer.

* Only Basic font/text formatting possible:
* Bold (yes / no), Font Type, Font Size
*
* usefull placeholders, which can be used in header/footer value strings
* '&P' - page number
* '&N' - total number of pages
* '&D' - Date
* '&T' - Time
* '&F' - File Name
* '&Z' - Path
* '&A' - Sheet name
* new line via class constant CL_ABAP_CHAR_UTILITIES=>newline
*
* Example Value String 'page &P of &N'
*
* DO NOT USE &L , &C or &R which automatically created as position markers

    me->odd_header = ip_odd_header.
    me->odd_footer = ip_odd_footer.
    me->even_header = ip_even_header.
    me->even_footer = ip_even_footer.

    IF me->even_header IS NOT INITIAL OR me->even_footer IS NOT INITIAL.
      me->diff_oddeven_headerfooter = abap_true.
    ENDIF.


  ENDMETHOD.


  METHOD set_page_margins.
    DATA: lv_coef TYPE f,
          lv_unit TYPE string.

    lv_unit = ip_unit.
    TRANSLATE lv_unit TO UPPER CASE.

    CASE lv_unit.
      WHEN 'IN'. lv_coef = 1.
      WHEN 'CM'. lv_coef = '0.393700787'.
      WHEN 'MM'. lv_coef = '0.0393700787'.
    ENDCASE.

    IF ip_bottom IS SUPPLIED. margin_bottom = lv_coef * ip_bottom. ENDIF.
    IF ip_footer IS SUPPLIED. margin_footer = lv_coef * ip_footer. ENDIF.
    IF ip_header IS SUPPLIED. margin_header = lv_coef * ip_header. ENDIF.
    IF ip_left IS SUPPLIED.   margin_left   = lv_coef * ip_left. ENDIF.
    IF ip_right IS SUPPLIED.  margin_right  = lv_coef * ip_right. ENDIF.
    IF ip_top IS SUPPLIED.    margin_top    = lv_coef * ip_top. ENDIF.

  ENDMETHOD.
ENDCLASS.
