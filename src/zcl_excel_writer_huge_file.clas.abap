CLASS zcl_excel_writer_huge_file DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_writer_2007
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF ty_cell,
        name    TYPE c LENGTH 10, "AAA1234567"
        style   TYPE i,
        type    TYPE c LENGTH 9,
        formula TYPE string,
        value   TYPE string,
      END OF ty_cell .

    DATA:
      cells TYPE STANDARD TABLE OF ty_cell .

    METHODS get_cells
      IMPORTING
        !i_row   TYPE i
        !i_index TYPE i .
  PROTECTED SECTION.

    METHODS create_xl_sharedstrings
        REDEFINITION .
    METHODS create_xl_sheet
        REDEFINITION .
  PRIVATE SECTION.

    DATA worksheet TYPE REF TO zcl_excel_worksheet .

    "! Insert XML parts into the sheet XML: after the end tag IV_AFTER, or before the first found of IT_BEFORE
    METHODS insert_sheet_parts
      IMPORTING
        !it_parts  TYPE ty_t_sheet_parts
        !it_names  TYPE string_table
        !iv_after  TYPE string OPTIONAL
        !it_before TYPE string_table OPTIONAL
      CHANGING
        !cv_xml    TYPE xstring .
ENDCLASS.



CLASS ZCL_EXCEL_WRITER_HUGE_FILE IMPLEMENTATION.


  METHOD create_xl_sharedstrings.
*
* Redefinition using simple transformation instead of CL_IXML
*
** Constant node name

    TYPES:
      BEGIN OF ts_root,
        count        TYPE string,
        unique_count TYPE string,
      END OF ts_root.

    DATA:
      lo_iterator  TYPE REF TO zcl_excel_collection_iterator,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

    DATA:
      ls_root           TYPE ts_root,
      lt_cell_data      TYPE zexcel_t_cell_data_unsorted,
      ls_shared_string  TYPE zexcel_s_shared_string,
      lv_sytabix        TYPE i,
      lt_shared_strings TYPE TABLE OF zexcel_s_shared_string.

    FIELD-SYMBOLS:
      <sheet_content>     TYPE zexcel_s_cell_data.

**********************************************************************
* STEP 1: Collect strings from each worksheet

    lo_iterator = excel->get_worksheets_iterator( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      " rich text (runs) is only supported by the shared strings of ZCL_EXCEL_WRITER_2007
      " (unique strings only - much smaller than the sheet data)
      LOOP AT lo_worksheet->sheet_content TRANSPORTING NO FIELDS WHERE rtf_tab IS NOT INITIAL. "#EC CI_SORTSEQ
        ep_content = super->create_xl_sharedstrings( ).
        RETURN.
      ENDLOOP.
      APPEND LINES OF lo_worksheet->sheet_content TO lt_cell_data.
    ENDWHILE.

    DELETE lt_cell_data WHERE cell_formula IS NOT INITIAL " delete formula content
                           OR data_type    NE 's'.        " MvC: Only shared strings

    ls_root-count = lines( lt_cell_data ).
    CONDENSE ls_root-count.

    SORT lt_cell_data BY cell_value data_type.
    DELETE ADJACENT DUPLICATES FROM lt_cell_data COMPARING cell_value data_type.

    ls_root-unique_count = lines( lt_cell_data ).
    CONDENSE ls_root-unique_count.

    LOOP AT lt_cell_data ASSIGNING <sheet_content>.

      lv_sytabix = sy-tabix - 1.
      ls_shared_string-string_no = lv_sytabix.
      ls_shared_string-string_value = <sheet_content>-cell_value.
      INSERT ls_shared_string INTO TABLE shared_strings.

    ENDLOOP.

**********************************************************************
* STEP 2: Create XML

    " Escape the string values, use of standard table in order to keep same sort order as in sorted table SHARED_STRINGS.
    CLEAR lt_shared_strings.
    LOOP AT shared_strings INTO ls_shared_string.
      ls_shared_string-string_value = escape_string_value( ls_shared_string-string_value ).
      APPEND ls_shared_string TO lt_shared_strings.
    ENDLOOP.

    CALL TRANSFORMATION zexcel_tr_shared_strings
      SOURCE root = ls_root
             shared_strings = lt_shared_strings
      OPTIONS xml_header = 'full'
      RESULT XML ep_content.


  ENDMETHOD.


  METHOD create_xl_sheet.
*
* Build Sheet#.xml with Simple Transformation ZEXCEL_TR_SHEET
*
* This is an adaption of ZCL_EXCEL_WRITER_2007.
* Not all features are supported, notably the autofilter settings,
* conditional formatting and sheet protection.
*
* Bug reports to marcus.voncube AT deutschebahn.com
*
    TYPES:
      lty_bool TYPE c LENGTH 5.

    CONSTANTS:
      lc_true              TYPE lty_bool   VALUE 'true',
      lc_zero              TYPE c LENGTH 1 VALUE '0',
      lc_one               TYPE c LENGTH 1 VALUE '1',
      lc_default_col_width TYPE string VALUE `9.1`.

    TYPES:
      BEGIN OF lty_column,
        min          TYPE i,
        max          TYPE i,
        width        TYPE string,
        hidden       TYPE lty_bool,
        customwidth  TYPE lty_bool,
        bestfit      TYPE lty_bool,
        collapsed    TYPE lty_bool,
        outlinelevel TYPE i,
        style        TYPE i,
      END OF lty_column,

      BEGIN OF lty_row,
        row          TYPE i,
        index        TYPE i,
        spans        TYPE c LENGTH 11,  "12345:12345"
        hidden       TYPE lty_bool,
        customheight TYPE lty_bool,
        height       TYPE string,
        collapsed    TYPE lty_bool,
        outlinelevel TYPE i,
        customformat TYPE lty_bool,
        style        TYPE i,
      END OF lty_row,

      BEGIN OF lty_mergecell,
        ref TYPE c LENGTH 21, "AAA1234567:BBB1234567"
      END OF lty_mergecell,

      BEGIN OF lty_hyperlink,
        ref      TYPE string,
        location TYPE string,
        r_id     TYPE string,
      END OF lty_hyperlink,

      BEGIN OF lty_table,
        r_id TYPE string,
      END OF lty_table,

      BEGIN OF lty_table_area,
        left   TYPE i,
        right  TYPE i,
        top    TYPE i,
        bottom TYPE i,
      END OF lty_table_area,

      BEGIN OF ty_missing_columns,
        first_column TYPE zexcel_cell_column,
        last_column  TYPE zexcel_cell_column,
      END OF ty_missing_columns.

*
* Root node for transformation
*
    DATA:
      BEGIN OF l_worksheet,
        dimension          TYPE string,
        tabcolor           TYPE string,
        summarybelow       TYPE c,
        summaryright       TYPE c,
        fittopage          TYPE c,
        filtermode         TYPE c,
        showzeros          TYPE c,
        tabselected        TYPE c,
        zoomscale          TYPE i,
        zoomscalenormal    TYPE i,
        zoomscalepageview  TYPE i,
        zoomscalesheetview TYPE i,
        righttoleft        TYPE i,
        workbookviewid     TYPE c,
        showgridlines      TYPE c,
        showrowcolheaders  TYPE c,
        activepane         TYPE string,
        state              TYPE string,
        ysplit             TYPE i,
        xsplit             TYPE i,
        topleftcell        TYPE c LENGTH 10,
        activecell         TYPE c LENGTH 10,
        customheight       TYPE lty_bool,
        defaultrowheight   TYPE string,
        defaultcolwidth    TYPE string,
        outlinelevelrow    TYPE i,
        outlinelevelcol    TYPE i,
        cols               TYPE STANDARD TABLE OF lty_column,
        rows               TYPE STANDARD TABLE OF lty_row,
        mergecells_count   TYPE i,
        mergecells         TYPE STANDARD TABLE OF lty_mergecell,
        hyperlinks_count   TYPE i,
        hyperlinks         TYPE STANDARD TABLE OF lty_hyperlink,
        BEGIN OF printoptions,
          gridlines          TYPE lty_bool,
          horizontalcentered TYPE lty_bool,
          verticalcentered   TYPE lty_bool,
        END OF printoptions,
        BEGIN OF pagemargins,
          left   TYPE zexcel_dec_8_2,
          right  TYPE zexcel_dec_8_2,
          top    TYPE zexcel_dec_8_2,
          bottom TYPE zexcel_dec_8_2,
          header TYPE zexcel_dec_8_2,
          footer TYPE zexcel_dec_8_2,
        END OF pagemargins,
        BEGIN OF pagesetup,
          blackandwhite      TYPE c,
          cellcomments       TYPE string,
          copies             TYPE i,
          draft              TYPE c,
          errors             TYPE string,
          firstpagenumber    TYPE i,
          fittoheight        TYPE i,
          fittowidth         TYPE i,
          horizontaldpi      TYPE i,
          orientation        TYPE string,
          pageorder          TYPE string,
          paperheight        TYPE string,
          papersize          TYPE i,
          paperwidth         TYPE string,
          scale              TYPE i,
          usefirstpagenumber TYPE c,
          useprinterdefaults TYPE c,
          verticaldpi        TYPE i,
        END OF pagesetup,
        BEGIN OF headerfooter,
          differentoddeven TYPE c,
          oddheader        TYPE string,
          oddfooter        TYPE string,
          evenheader       TYPE string,
          evenfooter       TYPE string,
        END OF headerfooter,
        drawings           TYPE string,
        tables_count       TYPE i,
        tables             TYPE STANDARD TABLE OF lty_table,
      END OF l_worksheet.

*
* Local data
*
    DATA:
      lo_iterator                 TYPE REF TO zcl_excel_collection_iterator,
      lo_table                    TYPE REF TO zcl_excel_table,
      lo_column_default           TYPE REF TO zcl_excel_column,
      lo_row_default              TYPE REF TO zcl_excel_row,
      lv_value                    TYPE string,
      lv_index                    TYPE i,
      lv_spans                    TYPE string,
      lt_range_merge              TYPE string_table,
      lv_column                   TYPE zexcel_cell_column,
      lv_style_guid               TYPE zexcel_cell_style,
      ls_last_row                 TYPE zexcel_s_cell_data,
      lv_freeze_cell_row          TYPE zexcel_cell_row,
      lv_freeze_cell_column       TYPE zexcel_cell_column,
      lv_freeze_cell_column_alpha TYPE zexcel_cell_column_alpha,
      lo_column_iterator          TYPE REF TO zcl_excel_collection_iterator,
      lo_column                   TYPE REF TO zcl_excel_column,
      lo_row_iterator             TYPE REF TO zcl_excel_collection_iterator,
      lo_row                      TYPE REF TO zcl_excel_row,
      lv_relation_id              TYPE i VALUE 0,
      outline_level_row           TYPE i VALUE 0,
      outline_level_col           TYPE i VALUE 0,
      col_count                   TYPE int4,
      lt_table_areas              TYPE SORTED TABLE OF lty_table_area
                                       WITH NON-UNIQUE KEY left right top bottom,
      ls_table_area               LIKE LINE OF lt_table_areas,
      lts_sorted_columns          TYPE SORTED TABLE OF zexcel_cell_column
                                       WITH UNIQUE KEY table_line,
      t_missing_columns           TYPE STANDARD TABLE OF ty_missing_columns
                                       WITH NON-UNIQUE DEFAULT KEY,
      missing_column              LIKE LINE OF t_missing_columns,
      lo_link                     TYPE REF TO zcl_excel_hyperlink,
      lo_drawings                 TYPE REF TO zcl_excel_drawings,
      lo_autofilter               TYPE REF TO zcl_excel_autofilter,
      lt_parts                    TYPE ty_t_sheet_parts,
      lth_rows                    TYPE HASHED TABLE OF i WITH UNIQUE KEY table_line,
      lth_candidates              TYPE SORTED TABLE OF i WITH UNIQUE KEY table_line,
      lv_no_cells_index           TYPE i,
      lt_part_names               TYPE string_table.

    FIELD-SYMBOLS:
      <sheet_content> TYPE zexcel_s_cell_data,
      <range_merge>   LIKE LINE OF lt_range_merge,
      <col>           TYPE lty_column,
      <row>           TYPE lty_row,
      <hyperlink>     TYPE lty_hyperlink,
      <mergecell>     TYPE lty_mergecell,
      <table>         TYPE lty_table.

**********************************************************************
* STEP 0: Sheet parts not supported by the transformation are created like in ZCL_EXCEL_WRITER_2007
*         (without cell data, i.e. cheap) and inserted into the transformation result in STEP 3
*
    lt_part_names = VALUE #( ( `sheetPr` ) ( `sheetProtection` ) ( `autoFilter` ) ( `conditionalFormatting` )
                             ( `dataValidations` ) ( `rowBreaks` ) ( `colBreaks` ) ( `ignoredErrors` )
                             ( `legacyDrawing` ) ( `legacyDrawingHF` ) ) ##NO_TEXT.
    lt_parts = create_xl_sheet_parts( io_worksheet = io_worksheet
                                      iv_active    = iv_active
                                      it_names     = lt_part_names ).
    READ TABLE lt_parts INTO DATA(ls_sheetpr) WITH KEY name = `sheetPr` ##NO_TEXT.
    IF sy-subrc = 0 AND ls_sheetpr-xml CS 'filterMode="1"'.
      l_worksheet-filtermode = lc_one.
    ENDIF.
    lo_autofilter = excel->get_autofilters_reference( )->get( io_worksheet ).
    IF lo_autofilter IS BOUND.
      lo_autofilter->get_filter_area( ).   " validates the filter area used by IS_ROW_HIDDEN
    ENDIF.

**********************************************************************
* STEP 1: Fill root node
*
    l_worksheet-tabcolor      = io_worksheet->tabcolor-rgb.
    l_worksheet-summarybelow  = io_worksheet->zif_excel_sheet_properties~summarybelow.
    l_worksheet-summaryright  = io_worksheet->zif_excel_sheet_properties~summaryright.

    IF io_worksheet->sheet_setup->fit_to_page IS NOT INITIAL.
      l_worksheet-fittopage   = lc_one.
    ENDIF.

    l_worksheet-dimension     = io_worksheet->get_dimension_range( ).

    IF io_worksheet->zif_excel_sheet_properties~show_zeros EQ abap_true.
      l_worksheet-showzeros   = lc_one.
    ELSE.
      l_worksheet-showzeros   = lc_zero.
    ENDIF.

    IF iv_active = abap_true
      OR io_worksheet->zif_excel_sheet_properties~selected EQ abap_true.
      l_worksheet-tabselected = lc_one.
    ELSE.
      l_worksheet-tabselected = lc_zero.
    ENDIF.

    IF io_worksheet->zif_excel_sheet_properties~zoomscale GT 400.
      io_worksheet->zif_excel_sheet_properties~zoomscale = 400.
    ELSEIF io_worksheet->zif_excel_sheet_properties~zoomscale LT 10.
      io_worksheet->zif_excel_sheet_properties~zoomscale = 10.
    ENDIF.
    l_worksheet-zoomscale = io_worksheet->zif_excel_sheet_properties~zoomscale.

    IF io_worksheet->zif_excel_sheet_properties~zoomscale_normal NE 0.
      IF io_worksheet->zif_excel_sheet_properties~zoomscale_normal GT 400.
        io_worksheet->zif_excel_sheet_properties~zoomscale_normal = 400.
      ELSEIF io_worksheet->zif_excel_sheet_properties~zoomscale_normal LT 10.
        io_worksheet->zif_excel_sheet_properties~zoomscale_normal = 10.
      ENDIF.
      l_worksheet-zoomscalenormal = io_worksheet->zif_excel_sheet_properties~zoomscale_normal.
    ENDIF.

    IF io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview NE 0.
      IF io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview GT 400.
        io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview = 400.
      ELSEIF io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview LT 10.
        io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview = 10.
      ENDIF.
      l_worksheet-zoomscalepageview = io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview.
    ENDIF.

    IF io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview NE 0.
      IF io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview GT 400.
        io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview = 400.
      ELSEIF io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview LT 10.
        io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview = 10.
      ENDIF.
      l_worksheet-zoomscalesheetview = io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview.
    ENDIF.

    IF io_worksheet->zif_excel_sheet_properties~get_right_to_left( ) EQ abap_true.
      l_worksheet-righttoleft = lc_one.
    ELSE.
      l_worksheet-righttoleft = lc_zero.
    ENDIF.

    l_worksheet-workbookviewid = lc_zero.

    IF io_worksheet->show_gridlines = abap_true.
      l_worksheet-showgridlines = lc_one.
    ELSE.
      l_worksheet-showgridlines = lc_zero.
    ENDIF.

    IF io_worksheet->show_rowcolheaders = abap_true.
      l_worksheet-showrowcolheaders = lc_one.
    ELSE.
      l_worksheet-showrowcolheaders = lc_zero.
    ENDIF.

*
* Freeze
*
    io_worksheet->get_freeze_cell(
      IMPORTING ep_row    = lv_freeze_cell_row
                ep_column = lv_freeze_cell_column ).

    IF lv_freeze_cell_row IS NOT INITIAL AND lv_freeze_cell_column IS NOT INITIAL.
      IF lv_freeze_cell_row > 1.
        l_worksheet-ysplit = lv_freeze_cell_row - 1.
      ENDIF.

      IF lv_freeze_cell_column > 1.
        lv_value = lv_freeze_cell_column - 1.
        l_worksheet-xsplit = lv_freeze_cell_column - 1.
      ENDIF.

      IF io_worksheet->pane_top_left_cell IS NOT INITIAL.
        l_worksheet-topleftcell = io_worksheet->pane_top_left_cell.
      ELSE.
        lv_freeze_cell_column_alpha = zcl_excel_common=>convert_column2alpha( ip_column = lv_freeze_cell_column ).
        lv_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_freeze_cell_row  ).
        CONCATENATE lv_freeze_cell_column_alpha lv_value INTO lv_value.
        l_worksheet-topleftcell = lv_value.
      ENDIF.

      " active pane depends on the split (like ZCL_EXCEL_WRITER_2007)
      IF lv_freeze_cell_row > 1 AND lv_freeze_cell_column > 1.
        l_worksheet-activepane = 'bottomRight'.
      ELSEIF lv_freeze_cell_row > 1.
        l_worksheet-activepane = 'bottomLeft'.
      ELSE.
        l_worksheet-activepane = 'topRight'.
      ENDIF.
      l_worksheet-state      = 'frozen'.
    ENDIF.

    l_worksheet-activecell = io_worksheet->get_active_cell( ).

*
* Row and column info
*
    io_worksheet->calculate_column_widths( ).
    lo_column_iterator = io_worksheet->get_columns_iterator( ).

    lo_column_default = io_worksheet->get_default_column( ).
    IF lo_column_default IS BOUND AND lo_column_default->get_width( ) >= 0.
      l_worksheet-defaultcolwidth = |{ round( val = CONV decfloat34( lo_column_default->get_width( ) ) dec = 8 ) STYLE = SIMPLE }|.
    ENDIF.

    lo_row_default = io_worksheet->get_default_row( ).
    IF lo_row_default IS BOUND.
      IF lo_row_default->get_row_height( ) >= 0.
        l_worksheet-customheight = lc_true.
        lv_value = |{ round( val = CONV decfloat34( lo_row_default->get_row_height( ) ) dec = 8 ) STYLE = SIMPLE }|.
      ELSE.
        lv_value = '12.75'.
      ENDIF.
    ELSE.
      lv_value = '12.75'.
    ENDIF.
    CONDENSE lv_value.
    l_worksheet-defaultrowheight = lv_value.

    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    WHILE lo_row_iterator->has_next( ) = abap_true.
      lo_row ?= lo_row_iterator->get_next( ).
      IF lo_row->get_outline_level( ) > outline_level_row.
        l_worksheet-outlinelevelrow = lo_row->get_outline_level( ).
      ENDIF.
    ENDWHILE.

* Set column information (width, style, ...)
    WHILE lo_column_iterator->has_next( ) = abap_true.
      lo_column ?= lo_column_iterator->get_next( ).
      IF lo_column->get_outline_level( ) > outline_level_col.
        outline_level_col = lo_column->get_outline_level( ).
        l_worksheet-outlinelevelcol = outline_level_col.
      ENDIF.
      " a column without own properties is not written (like ZCL_EXCEL_WRITER_2007)
      IF lo_column->get_width( ) < 0
          AND lo_column->get_visible( ) = abap_true
          AND lo_column->get_auto_size( ) = abap_false
          AND lo_column->get_collapsed( ) = abap_false
          AND lo_column->get_outline_level( ) = 0
          AND io_worksheet->zif_excel_sheet_properties~style IS INITIAL
          AND lo_column->get_column_style_guid( ) = io_worksheet->zif_excel_sheet_properties~get_style( ).
        CONTINUE.
      ENDIF.
      APPEND INITIAL LINE TO l_worksheet-cols ASSIGNING <col>.
      <col>-min = <col>-max = lo_column->get_column_index( ).
      IF lo_column->get_width( ) >= 0.
        <col>-width = |{ round( val = CONV decfloat34( lo_column->get_width( ) ) dec = 8 ) STYLE = SIMPLE }|.
      ELSEIF lo_column_default IS BOUND AND lo_column_default->get_width( ) >= 0.
        <col>-width = |{ round( val = CONV decfloat34( lo_column_default->get_width( ) ) dec = 8 ) STYLE = SIMPLE }|.
      ELSE.
        <col>-width = lc_default_col_width.
      ENDIF.
      IF lo_column->get_visible( ) = abap_false.
        <col>-hidden = lc_true.
      ENDIF.
      IF lo_column->get_auto_size( ) = abap_true.
        <col>-bestfit = lc_true.
      ENDIF.
      IF lo_column->get_width( ) >= 0.
        IF lo_column_default IS BOUND.
          IF lo_column->get_width( ) <> lo_column_default->get_width( ).
            <col>-customwidth = lc_true.
          ENDIF.
        ELSE.
          <col>-customwidth = lc_true.
        ENDIF.
      ENDIF.
      IF lo_column->get_collapsed( ) = abap_true.
        <col>-collapsed = lc_true.
      ENDIF.
      <col>-outlinelevel = lo_column->get_outline_level( ).
      lv_style_guid = lo_column->get_column_style_guid( ).
      <col>-style = me->excel->get_style_index_in_styles( lv_style_guid ) - 1.
*
* Missing columns
*
* First collect columns that were already handled before.
* The rest has to be inserted now.
*

      lv_column = zcl_excel_common=>convert_column2int( lo_column->get_column_index( ) ).
      INSERT lv_column INTO TABLE lts_sorted_columns.
    ENDWHILE.

*
* Now find all columns that were missing so far
*
    missing_column-first_column = 1.

    LOOP AT lts_sorted_columns INTO lv_column.
      IF lv_column > missing_column-first_column.
        missing_column-last_column = lv_column - 1.
        APPEND missing_column TO t_missing_columns.
      ENDIF.
      missing_column-first_column = lv_column + 1.
    ENDLOOP.

    missing_column-last_column = zcl_excel_common=>c_excel_sheet_max_col.
    APPEND missing_column TO t_missing_columns.

*
* Now apply stylesetting and other defaults (only needed for a sheet style, like ZCL_EXCEL_WRITER_2007)
*
    IF io_worksheet->zif_excel_sheet_properties~style IS INITIAL.
      CLEAR t_missing_columns.
    ENDIF.
    LOOP AT t_missing_columns INTO missing_column.
      APPEND INITIAL LINE TO l_worksheet-cols ASSIGNING <col>.
      <col>-min = missing_column-first_column.
      <col>-max = missing_column-last_column.
      IF lo_column_default IS BOUND AND lo_column_default->get_width( ) >= 0.
        <col>-width = |{ round( val = CONV decfloat34( lo_column_default->get_width( ) ) dec = 8 ) STYLE = SIMPLE }|.
      ELSE.
        <col>-width = lc_default_col_width.
      ENDIF.
      lv_style_guid = io_worksheet->zif_excel_sheet_properties~get_style( ).
      <col>-style = me->excel->get_style_index_in_styles( lv_style_guid ) - 1.
    ENDLOOP.

*
* Build table to hold all table-areas attached to this sheet
*
    lo_iterator = io_worksheet->get_tables_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_table ?= lo_iterator->get_next( ).
      ls_table_area-left   = zcl_excel_common=>convert_column2int( lo_table->settings-top_left_column ).
      ls_table_area-right  = lo_table->get_right_column_integer( ).
      ls_table_area-top    = lo_table->settings-top_left_row.
      ls_table_area-bottom = lo_table->get_bottom_row_integer( ).
      INSERT ls_table_area INTO TABLE lt_table_areas.
    ENDWHILE.

*
* Build sheet data node
*
* Spans is constant amongst all rows
*
    col_count = io_worksheet->get_highest_column( ).
    lv_spans = col_count.
    CONCATENATE '1:' lv_spans INTO lv_spans.
    CONDENSE lv_spans.

    LOOP AT io_worksheet->sheet_content ASSIGNING <sheet_content>.
      IF ls_last_row-cell_row NE <sheet_content>-cell_row.
*
*     Row with cells - the cells are filled in by callback GET_CELLS called from the transformation
*
        APPEND VALUE #( row = <sheet_content>-cell_row index = sy-tabix spans = lv_spans ) TO l_worksheet-rows.
        INSERT <sheet_content>-cell_row INTO TABLE lth_rows.
      ENDIF.
      ls_last_row = <sheet_content>.
    ENDLOOP.

*
*   Rows without cells but with row properties (height, hidden, style) or in a row outline
*   (like ZCL_EXCEL_WRITER_2007); index behind the last cell -> GET_CELLS returns no cells
*
    lv_no_cells_index = lines( io_worksheet->sheet_content ) + 1.
    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    WHILE lo_row_iterator->has_next( ) = abap_true.
      lo_row ?= lo_row_iterator->get_next( ).
      INSERT lo_row->get_row_index( ) INTO TABLE lth_candidates.
    ENDWHILE.
    LOOP AT io_worksheet->get_row_outlines( ) ASSIGNING FIELD-SYMBOL(<ls_row_outline>).
      DATA(lv_outline_row) = <ls_row_outline>-row_from.
      WHILE lv_outline_row <= <ls_row_outline>-row_to + 1.   " + 1: collapsed flag of the summary row
        INSERT lv_outline_row INTO TABLE lth_candidates.
        lv_outline_row = lv_outline_row + 1.
      ENDWHILE.
    ENDLOOP.
    LOOP AT lth_candidates INTO DATA(lv_candidate).
      IF NOT line_exists( lth_rows[ table_line = lv_candidate ] ).
        APPEND VALUE #( row = lv_candidate index = lv_no_cells_index spans = lv_spans ) TO l_worksheet-rows.
      ENDIF.
    ENDLOOP.
    SORT l_worksheet-rows BY row.

*
*   Row dimension attributes
*
    LOOP AT l_worksheet-rows ASSIGNING <row>.
      lo_row = io_worksheet->get_row( <row>-row ).
      IF lo_row->get_visible( io_worksheet ) = abap_false
          OR ( lo_autofilter IS BOUND AND lo_autofilter->is_row_hidden( <row>-row ) = abap_true ).
        <row>-hidden = lc_true.
      ENDIF.
      IF lo_row->get_row_height( ) >= 0.
        <row>-customheight = lc_one.
        <row>-height       = |{ round( val = CONV decfloat34( lo_row->get_row_height( ) ) dec = 8 ) STYLE = SIMPLE }|.
      ENDIF.
      IF lo_row->get_collapsed( io_worksheet ) = abap_true.
        <row>-collapsed = lc_true.
      ENDIF.
      " outline level from the row outlines of the worksheet (SET_ROW_OUTLINE)
      <row>-outlinelevel = lo_row->get_outline_level( io_worksheet ).
      IF <row>-outlinelevel > l_worksheet-outlinelevelrow.
        l_worksheet-outlinelevelrow = <row>-outlinelevel.
      ENDIF.
      <row>-style = lo_row->get_xf_index( ).
      IF <row>-style <> 0.
        <row>-customformat = lc_one.
      ENDIF.
      " a row without cells is only written if it has any row property
      IF <row>-index = lv_no_cells_index
          AND <row>-hidden IS INITIAL AND <row>-customheight IS INITIAL AND <row>-collapsed IS INITIAL
          AND <row>-outlinelevel = 0 AND <row>-style = 0.
        DELETE l_worksheet-rows.
      ENDIF.
    ENDLOOP.

*
* Merged cells
*
    lt_range_merge = io_worksheet->get_merge( ).
    IF lt_range_merge IS NOT INITIAL.
      l_worksheet-mergecells_count = lines( lt_range_merge ).

      LOOP AT lt_range_merge ASSIGNING <range_merge>.
        APPEND INITIAL LINE TO l_worksheet-mergecells ASSIGNING <mergecell>.
        <mergecell>-ref = <range_merge>.
      ENDLOOP.
    ENDIF.

*
* Hyperlinks
*
    l_worksheet-hyperlinks_count = io_worksheet->get_hyperlinks_size( ).
    IF l_worksheet-hyperlinks_count > 0.
      lo_iterator = io_worksheet->get_hyperlinks_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.
        lo_link ?= lo_iterator->get_next( ).

        APPEND INITIAL LINE TO l_worksheet-hyperlinks ASSIGNING <hyperlink>.
        <hyperlink>-ref = lo_link->get_ref( ).
        IF lo_link->is_internal( ) = abap_true.
          <hyperlink>-location = lo_link->get_url( ).
        ELSE.
          lv_relation_id = lv_relation_id + 1.
          lv_value = lv_relation_id.
          CONDENSE lv_value.
          CONCATENATE 'rId' lv_value INTO lv_value.
          <hyperlink>-r_id = lv_value.
        ENDIF.
      ENDWHILE.
    ENDIF.

*
* Print options
*
    IF io_worksheet->print_gridlines = abap_true.
      l_worksheet-printoptions-gridlines = lc_true.
    ENDIF.

    IF io_worksheet->sheet_setup->horizontal_centered = abap_true.
      l_worksheet-printoptions-horizontalcentered = lc_true.
    ENDIF.

    IF io_worksheet->sheet_setup->vertical_centered = abap_true.
      l_worksheet-printoptions-verticalcentered = lc_true.
    ENDIF.

*
* Page margins
*
    l_worksheet-pagemargins-left   = io_worksheet->sheet_setup->margin_left.
    l_worksheet-pagemargins-right  = io_worksheet->sheet_setup->margin_right.
    l_worksheet-pagemargins-top    = io_worksheet->sheet_setup->margin_top.
    l_worksheet-pagemargins-bottom = io_worksheet->sheet_setup->margin_bottom.
    l_worksheet-pagemargins-header = io_worksheet->sheet_setup->margin_header.
    l_worksheet-pagemargins-footer = io_worksheet->sheet_setup->margin_footer.

*
* Page setup
*
    l_worksheet-pagesetup-cellcomments       = io_worksheet->sheet_setup->cell_comments.
    l_worksheet-pagesetup-copies             = io_worksheet->sheet_setup->copies.
    l_worksheet-pagesetup-firstpagenumber    = io_worksheet->sheet_setup->first_page_number.
    l_worksheet-pagesetup-fittoheight        = io_worksheet->sheet_setup->fit_to_height.
    l_worksheet-pagesetup-fittowidth         = io_worksheet->sheet_setup->fit_to_width.
    l_worksheet-pagesetup-horizontaldpi      = io_worksheet->sheet_setup->horizontal_dpi.
    l_worksheet-pagesetup-orientation        = io_worksheet->sheet_setup->orientation.
    l_worksheet-pagesetup-pageorder          = io_worksheet->sheet_setup->page_order.
    l_worksheet-pagesetup-paperheight        = io_worksheet->sheet_setup->paper_height.
    l_worksheet-pagesetup-papersize          = io_worksheet->sheet_setup->paper_size.
    l_worksheet-pagesetup-paperwidth         = io_worksheet->sheet_setup->paper_width.
    l_worksheet-pagesetup-scale              = io_worksheet->sheet_setup->scale.
    l_worksheet-pagesetup-usefirstpagenumber = io_worksheet->sheet_setup->use_first_page_num.
    l_worksheet-pagesetup-verticaldpi        = io_worksheet->sheet_setup->vertical_dpi.

    IF io_worksheet->sheet_setup->black_and_white IS NOT INITIAL.
      l_worksheet-pagesetup-blackandwhite = lc_one.
    ENDIF.

    IF io_worksheet->sheet_setup->draft IS NOT INITIAL.
      l_worksheet-pagesetup-draft = lc_one.
    ENDIF.

    IF io_worksheet->sheet_setup->errors IS NOT INITIAL.
      l_worksheet-pagesetup-errors = io_worksheet->sheet_setup->errors.
    ENDIF.

    IF io_worksheet->sheet_setup->use_printer_defaults IS NOT INITIAL.
      l_worksheet-pagesetup-useprinterdefaults = lc_one.
    ENDIF.

*
* Header and footer
*
    IF io_worksheet->sheet_setup->diff_oddeven_headerfooter = abap_true.
      l_worksheet-headerfooter-differentoddeven = lc_one.
    ENDIF.

    io_worksheet->sheet_setup->get_header_footer_string(
      IMPORTING
        ep_odd_header  = l_worksheet-headerfooter-oddheader
        ep_odd_footer  = l_worksheet-headerfooter-oddfooter
        ep_even_header = l_worksheet-headerfooter-evenheader
        ep_even_footer = l_worksheet-headerfooter-evenfooter ).

*
* Drawings
*
    lo_drawings = io_worksheet->get_drawings( ).
    IF lo_drawings->is_empty( ) = abap_false.
      lv_relation_id = lv_relation_id + 1.
      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO l_worksheet-drawings.
    ENDIF.

*
* Comments (legacyDrawing + comments part) and header / footer images (legacyDrawingHF) - the elements
* are inserted in STEP 3, their relationships come before the tables (see CREATE_XL_SHEET_RELS)
*
    IF io_worksheet->get_comments( )->is_empty( ) = abap_false.
      lv_relation_id = lv_relation_id + 2.
    ENDIF.
    IF lines( io_worksheet->get_header_footer_drawings( ) ) > 0.
      lv_relation_id = lv_relation_id + 1.
    ENDIF.

*
* Tables
*
    l_worksheet-tables_count = io_worksheet->get_tables_size( ).
    IF l_worksheet-tables_count > 0.
      lo_iterator = io_worksheet->get_tables_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.
        lo_table ?= lo_iterator->get_next( ).
        APPEND INITIAL LINE TO l_worksheet-tables ASSIGNING <table>.
        lv_relation_id = lv_relation_id + 1.
        lv_value = lv_relation_id.
        CONDENSE lv_value.
        CONCATENATE 'rId' lv_value INTO <table>-r_id.
      ENDWHILE.
    ENDIF.

**********************************************************************
* STEP 2: Create XML

    me->worksheet = io_worksheet. "Neccessary for callback GET_CELL

    CALL TRANSFORMATION zexcel_tr_sheet
      SOURCE worksheet = l_worksheet
             cells     = me->cells
             writer    = me
      OPTIONS xml_header = 'full'
      RESULT XML ep_content.

**********************************************************************
* STEP 3: Insert the parts of STEP 0 in the order of the OOXML schema
*         (sheetData, sheetProtection, autoFilter, mergeCells, conditionalFormatting, dataValidations,
*          hyperlinks, printOptions, pageMargins, pageSetup, headerFooter, rowBreaks, colBreaks,
*          ignoredErrors, drawing, legacyDrawing, legacyDrawingHF, tableParts)
    insert_sheet_parts( EXPORTING it_parts  = lt_parts
                                  it_names  = VALUE #( ( `sheetProtection` ) ( `autoFilter` ) )
                                  iv_after  = `</sheetData>`
                                  it_before = VALUE #( ( `<mergeCells` ) ( `<hyperlinks` ) ( `<printOptions` ) ( `<pageMargins` ) )
                        CHANGING  cv_xml    = ep_content ) ##NO_TEXT.
    insert_sheet_parts( EXPORTING it_parts  = lt_parts
                                  it_names  = VALUE #( ( `conditionalFormatting` ) ( `dataValidations` ) )
                                  it_before = VALUE #( ( `<hyperlinks` ) ( `<printOptions` ) ( `<pageMargins` ) )
                        CHANGING  cv_xml    = ep_content ) ##NO_TEXT.
    insert_sheet_parts( EXPORTING it_parts  = lt_parts
                                  it_names  = VALUE #( ( `rowBreaks` ) ( `colBreaks` ) ( `ignoredErrors` ) )
                                  it_before = VALUE #( ( `<drawing ` ) ( `<tableParts` ) ( `</worksheet>` ) )
                        CHANGING  cv_xml    = ep_content ) ##NO_TEXT.
    insert_sheet_parts( EXPORTING it_parts  = lt_parts
                                  it_names  = VALUE #( ( `legacyDrawing` ) ( `legacyDrawingHF` ) )
                                  it_before = VALUE #( ( `<tableParts` ) ( `</worksheet>` ) )
                        CHANGING  cv_xml    = ep_content ) ##NO_TEXT.

  ENDMETHOD.                    "CREATE_XL_SHEET


  METHOD insert_sheet_parts.
    DATA: lv_parts  TYPE string,
          lv_offset TYPE i VALUE -1,
          lv_marker TYPE xstring.

    LOOP AT it_names INTO DATA(lv_name).
      LOOP AT it_parts INTO DATA(ls_part) WHERE name = lv_name.
        lv_parts = lv_parts && ls_part-xml.
      ENDLOOP.
    ENDLOOP.
    CHECK lv_parts IS NOT INITIAL.

    DATA(lo_conv) = cl_abap_conv_codepage=>create_out( ).
    IF iv_after IS NOT INITIAL.
      lv_marker = lo_conv->convert( iv_after ).
      FIND FIRST OCCURRENCE OF lv_marker IN cv_xml IN BYTE MODE MATCH OFFSET lv_offset.
      IF sy-subrc = 0.
        lv_offset = lv_offset + xstrlen( lv_marker ).
      ELSE.
        lv_offset = -1.
      ENDIF.
    ENDIF.
    IF lv_offset < 0.
      LOOP AT it_before INTO DATA(lv_before).
        lv_marker = lo_conv->convert( lv_before ).
        FIND FIRST OCCURRENCE OF lv_marker IN cv_xml IN BYTE MODE MATCH OFFSET lv_offset.
        IF sy-subrc = 0.
          EXIT.
        ENDIF.
        lv_offset = -1.
      ENDLOOP.
    ENDIF.
    CHECK lv_offset >= 0.

    DATA(lv_parts_x) = lo_conv->convert( lv_parts ).
    cv_xml = cv_xml(lv_offset) && lv_parts_x && cv_xml+lv_offset.
  ENDMETHOD.


  METHOD get_cells.
*
* Callback method from transformation ZEXCEL_TR_SHEET
*
* The method fills the data cells for each row.
* This saves memory if there are many rows.
*
    DATA:
      lv_cell_style TYPE zexcel_cell_style.

    FIELD-SYMBOLS:
      <cell>    TYPE ty_cell,
      <content> TYPE zexcel_s_cell_data,
      <style>   TYPE zexcel_s_styles_mapping.

    CLEAR cells.

    LOOP AT worksheet->sheet_content FROM i_index ASSIGNING <content>.
      IF <content>-cell_row <> i_row.
*
*     End of row
*
        EXIT.
      ENDIF.

*
*   Determine style index
*
      IF lv_cell_style <> <content>-cell_style.
        lv_cell_style = <content>-cell_style.
        UNASSIGN <style>.
        IF lv_cell_style IS NOT INITIAL.
          READ TABLE styles_mapping ASSIGNING <style> WITH KEY guid = lv_cell_style.
        ENDIF.
      ENDIF.
*
*   Add a new cell
*
      APPEND INITIAL LINE TO cells ASSIGNING <cell>.
      <cell>-name    = <content>-cell_coords.
      <cell>-formula = <content>-cell_formula.
      " column formula of a table (BIND_TABLE with COLUMN_FORMULA): formula of the first data row, shifted
      IF <cell>-formula IS INITIAL AND <content>-column_formula_id <> 0.
        READ TABLE worksheet->column_formulas ASSIGNING FIELD-SYMBOL(<ls_column_formula>)
             WITH TABLE KEY id = <content>-column_formula_id.
        IF sy-subrc = 0.
          TRY.
              <cell>-formula = zcl_excel_common=>shift_formula(
                                   iv_reference_formula = <ls_column_formula>-formula
                                   iv_shift_cols        = 0
                                   iv_shift_rows        = <content>-cell_row - <ls_column_formula>-table_top_left_row - 1 ).
            CATCH zcx_excel.
              <cell>-formula = <ls_column_formula>-formula.
          ENDTRY.
        ENDIF.
      ENDIF.
      <cell>-type    = <content>-data_type.   " initial for numbers -> no attribute t (see ZEXCEL_TR_SHEET)
      IF <cell>-type = 's'.
        <cell>-value = me->get_shared_string_index( ip_cell_value = <content>-cell_value
                                                    it_rtf        = <content>-rtf_tab ).
      ELSE.
        <cell>-value = <content>-cell_value.
      ENDIF.
      IF <style> IS ASSIGNED.
        <cell>-style = <style>-style.
      ELSE.
        <cell>-style = -1.
      ENDIF.

    ENDLOOP.

  ENDMETHOD.
ENDCLASS.
