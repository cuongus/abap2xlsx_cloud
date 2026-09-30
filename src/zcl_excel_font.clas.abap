CLASS zcl_excel_font DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES ty_font_height TYPE n LENGTH 3.
    CONSTANTS lc_default_font_height TYPE ty_font_height VALUE '110' ##NO_TEXT.
    CONSTANTS lc_default_font_name TYPE zexcel_style_font_name VALUE 'Calibri' ##NO_TEXT.

    CLASS-METHODS calculate_text_width
      IMPORTING
        !iv_font_name   TYPE zexcel_style_font_name
        !iv_font_height TYPE ty_font_height
        !iv_flag_bold   TYPE abap_bool
        !iv_flag_italic TYPE abap_bool
        !iv_cell_value  TYPE zexcel_cell_value
      RETURNING
        VALUE(rv_width) TYPE f .
  PROTECTED SECTION.
  PRIVATE SECTION.
    TYPES:
      BEGIN OF mty_s_font_metric,
        char       TYPE c LENGTH 1,
        char_width TYPE i,
      END OF mty_s_font_metric .
    TYPES:
      mty_th_font_metrics
               TYPE HASHED TABLE OF mty_s_font_metric
               WITH UNIQUE KEY char .
    TYPES:
      BEGIN OF mty_s_font_cache,
        font_name       TYPE zexcel_style_font_name,
        font_height     TYPE ty_font_height,
        flag_bold       TYPE abap_bool,
        flag_italic     TYPE abap_bool,
        th_font_metrics TYPE mty_th_font_metrics,
      END OF mty_s_font_cache .
    TYPES:
      mty_th_font_cache
               TYPE HASHED TABLE OF mty_s_font_cache
               WITH UNIQUE KEY font_name font_height flag_bold flag_italic .

    CLASS-DATA mth_font_cache TYPE mty_th_font_cache .

ENDCLASS.



CLASS ZCL_EXCEL_FONT IMPLEMENTATION.


  METHOD calculate_text_width.

    CONSTANTS lc_excel_cell_padding TYPE f VALUE '0.75'.

    DATA: ld_current_character       TYPE c LENGTH 1,
          ld_offset                  TYPE i,
          ld_length                  TYPE i,
          ld_width_from_font_metrics TYPE i,
          ls_font_cache              TYPE mty_s_font_cache.

    FIELD-SYMBOLS: <ls_font_cache>  TYPE mty_s_font_cache,
                   <ls_font_metric> TYPE mty_s_font_metric.

    " Check if the same font (font name and font attributes) was already
    " used before
    READ TABLE mth_font_cache
      WITH TABLE KEY
        font_name   = iv_font_name
        font_height = iv_font_height
        flag_bold   = iv_flag_bold
        flag_italic = iv_flag_italic
      ASSIGNING <ls_font_cache>.

    IF sy-subrc <> 0.
      " Font is used for the first time
      " Add the font to our local font cache
      " ABAP Cloud: SAPscript font metrics (TFO01 / LOAD_FONT) are not available,
      " the width is calculated from the font size only
      ls_font_cache-font_name   = iv_font_name.
      ls_font_cache-font_height = iv_font_height.
      ls_font_cache-flag_bold   = iv_flag_bold.
      ls_font_cache-flag_italic = iv_flag_italic.
      INSERT ls_font_cache INTO TABLE mth_font_cache
        ASSIGNING <ls_font_cache>.
    ENDIF.

    " Calculate the cell width
    " If available, use font metrics
    IF lines( <ls_font_cache>-th_font_metrics ) = 0.
      " Font metrics are not available
      " -> Calculate the cell width using only the font size
      " a cell with line breaks is as wide as its longest line; bold text is about 10 % wider
      ld_length = 0.
      SPLIT iv_cell_value AT cl_abap_char_utilities=>newline INTO TABLE DATA(lt_lines).
      LOOP AT lt_lines INTO DATA(ld_line).
        ld_length = nmax( val1 = ld_length val2 = strlen( ld_line ) ).
      ENDLOOP.
      rv_width = ld_length * iv_font_height / lc_default_font_height.
      IF iv_flag_bold = abap_true.
        rv_width = rv_width * '1.1'.
      ENDIF.
      rv_width = rv_width + lc_excel_cell_padding.

    ELSE.
      " Font metrics are available

      " Calculate the size of the text by adding the sizes of each
      " letter
      ld_length = strlen( iv_cell_value ).
      DO ld_length TIMES.
        " Subtract 1, because the first character is at offset 0
        ld_offset = sy-index - 1.

        " Read the current character from the cell value
        ld_current_character = iv_cell_value+ld_offset(1).

        " Look up the size of the current letter
        READ TABLE <ls_font_cache>-th_font_metrics
          WITH TABLE KEY char = ld_current_character
          ASSIGNING <ls_font_metric>.
        IF sy-subrc = 0.
          " The size of the letter is known
          " -> Add the actual size of the letter
          ld_width_from_font_metrics = ld_width_from_font_metrics + <ls_font_metric>-char_width.
        ELSE.
          " The size of the letter is unknown
          " -> Add the font height as the default letter size
          ld_width_from_font_metrics = ld_width_from_font_metrics + iv_font_height.
        ENDIF.
      ENDDO.

      " Add cell padding (Excel makes columns a bit wider than the space
      " that is needed for the text itself) and convert unit
      " (division by 100)
      rv_width = ld_width_from_font_metrics / 100 + lc_excel_cell_padding.
    ENDIF.

  ENDMETHOD.
ENDCLASS.
