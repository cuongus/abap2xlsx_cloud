CLASS zcl_excel_writer_csv DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
  PUBLIC SECTION.

    TYPES ty_ext_date TYPE c LENGTH 10.

    INTERFACES zif_excel_writer .

    "! Default value for initial dates e.g. user's format (DD.MM.YYYY, MM.DD.YYYY, etc.)
    CONSTANTS c_default TYPE c LENGTH 10 VALUE 'DEFAULT' ##NO_TEXT.

    CLASS-METHODS set_delimiter
      IMPORTING
        VALUE(ip_value) TYPE c DEFAULT ';' .
    CLASS-METHODS set_enclosure
      IMPORTING
        VALUE(ip_value) TYPE c DEFAULT '"' .
    CLASS-METHODS set_endofline
      IMPORTING
        VALUE(ip_value) TYPE any DEFAULT cl_abap_char_utilities=>cr_lf .
    CLASS-METHODS set_active_sheet_index
      IMPORTING
        !i_active_worksheet TYPE zexcel_active_worksheet .
    CLASS-METHODS set_active_sheet_index_by_name
      IMPORTING
        !i_worksheet_name TYPE zexcel_worksheets_name .
    CLASS-METHODS set_initial_ext_date
      IMPORTING
        !ip_value TYPE ty_ext_date DEFAULT c_default .
    CLASS-METHODS set_skip_hidden_rows
      IMPORTING
        !ip_value TYPE abap_bool.
    CLASS-METHODS set_skip_hidden_columns
      IMPORTING
        !ip_value TYPE abap_bool.
*"* protected components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_WRITER_CSV
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA excel TYPE REF TO zcl_excel .
    CLASS-DATA delimiter TYPE c VALUE ';' ##NO_TEXT.
    CLASS-DATA enclosure TYPE c VALUE '"' ##NO_TEXT.
    CLASS-DATA:
      eol TYPE c LENGTH 2 VALUE cl_abap_char_utilities=>cr_lf ##NO_TEXT.
    CLASS-DATA worksheet_name TYPE zexcel_worksheets_name .
    CLASS-DATA worksheet_index TYPE zexcel_active_worksheet .
    CLASS-DATA initial_ext_date TYPE ty_ext_date VALUE c_default.
    CLASS-DATA skip_hidden_rows TYPE abap_bool.
    CLASS-DATA skip_hidden_columns TYPE abap_bool.

    METHODS create
      RETURNING
        VALUE(ep_excel) TYPE xstring
      RAISING
        zcx_excel.
    METHODS create_csv
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
ENDCLASS.



CLASS ZCL_EXCEL_WRITER_CSV IMPLEMENTATION.


  METHOD create.

* .csv format with ; delimiter

* Start of insertion # issue 1134 - Dateretention of cellstyles(issue #139)
    me->excel->add_static_styles( ).
* End of insertion # issue 1134 - Dateretention of cellstyles(issue #139)

    ep_excel = me->create_csv( ).

  ENDMETHOD.


  METHOD create_csv.

    TYPES: BEGIN OF lty_format,
             cmpname  TYPE abap_attrname,
             attvalue TYPE string,
           END OF lty_format.
    DATA: lt_format TYPE STANDARD TABLE OF lty_format,
          ls_format LIKE LINE OF lt_format,
          lv_date   TYPE d,
          lv_tmp    TYPE string,
          lv_time   TYPE c LENGTH 8.

    DATA: lo_iterator  TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet TYPE REF TO zcl_excel_worksheet.

    DATA: lo_autofilter TYPE REF TO zcl_excel_autofilter.
    DATA: lv_row_hidden TYPE abap_bool.

    DATA: lt_cell_data TYPE zexcel_t_cell_data_unsorted,
          lv_row       TYPE i,
          lv_col       TYPE i,
          lv_string    TYPE string,
          lc_value     TYPE string,
          lv_attrname  TYPE abap_attrname,
          lv_previous_row TYPE zexcel_cell_row.

    DATA: lo_classdescr TYPE REF TO cl_abap_classdescr,
          lo_attr_type  TYPE REF TO cl_abap_datadescr,
          ls_attribute  TYPE abap_attrdescr,
          lv_date_fmt   TYPE string.

    FIELD-SYMBOLS <lv_constant> TYPE any.

    DATA: ls_numfmt TYPE zexcel_s_style_numfmt,
          lo_style  TYPE REF TO zcl_excel_style.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data.

* --- Retrieve supported cell format
*     ABAP Cloud: class repository table SEOCOMPODF is not released -> read the constants by RTTI
    lo_classdescr ?= cl_abap_typedescr=>describe_by_name( 'ZCL_EXCEL_STYLE_NUMBER_FORMAT' ).
    LOOP AT lo_classdescr->attributes INTO ls_attribute
         WHERE is_constant = abap_true AND visibility = cl_abap_objectdescr=>public.
      lo_attr_type = lo_classdescr->get_attribute_type( ls_attribute-name ).
      CHECK lo_attr_type->get_relative_name( ) = 'ZEXCEL_NUMBER_FORMAT'.
      ASSIGN zcl_excel_style_number_format=>(ls_attribute-name) TO <lv_constant>.
      CHECK sy-subrc = 0.
      CLEAR ls_format.
      ls_format-cmpname  = ls_attribute-name.
      ls_format-attvalue = |'{ <lv_constant> }'|.
      APPEND ls_format TO lt_format.
    ENDLOOP.

* --- Retrieve user date format
*     ABAP Cloud: domain texts of XUDATFM (DD07T) are not released -> derive the pattern
*     from the user specific formatting of a known date
    CLEAR ls_format.
    lv_date = '20231231'.
    lv_date_fmt = |{ lv_date DATE = USER }|.
    REPLACE '2023' IN lv_date_fmt WITH 'YYYY'.
    REPLACE '12' IN lv_date_fmt WITH 'MM'.
    REPLACE '31' IN lv_date_fmt WITH 'DD'.
    ls_format-cmpname = 'DATE'.
    ls_format-attvalue = |'{ lv_date_fmt }'|.
    APPEND ls_format TO lt_format.
    CLEAR lv_date.


    LOOP AT lt_format INTO ls_format.
      TRANSLATE ls_format-attvalue TO UPPER CASE.
      MODIFY lt_format FROM ls_format.
    ENDLOOP.


* STEP 1: Collect strings from the first worksheet
    lo_iterator = excel->get_worksheets_iterator( ).
    DATA: current_worksheet_title TYPE zexcel_sheet_title.

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      IF worksheet_name IS NOT INITIAL.
        current_worksheet_title = lo_worksheet->get_title( ).
        CHECK current_worksheet_title = worksheet_name.
      ELSE.
        IF worksheet_index IS INITIAL.
          worksheet_index = 1.
        ENDIF.
        CHECK worksheet_index = sy-index.
      ENDIF.
      APPEND LINES OF lo_worksheet->sheet_content TO lt_cell_data.
      EXIT. " Take first worksheet only
    ENDWHILE.

    DELETE lt_cell_data WHERE cell_formula IS NOT INITIAL. " delete formula content

    SORT lt_cell_data BY cell_row
                         cell_column.

    IF skip_hidden_rows = abap_true.
* --- Retrieve autofilters (to identify hidden rows)
      lo_autofilter = excel->get_autofilters_reference( )->get( io_worksheet = lo_worksheet ).
      IF lo_autofilter IS NOT INITIAL.
        lo_autofilter->get_filter_area( ). " trigger filter area validation
      ENDIF.
    ENDIF.

    lv_row = 1.
    lv_col = 1.
    CLEAR lv_string.
    LOOP AT lt_cell_data ASSIGNING <fs_sheet_content>.

* --- Check, if row is hidden (group change on CELL_ROW, replaces the obsolete AT NEW)
      IF sy-tabix = 1 OR <fs_sheet_content>-cell_row <> lv_previous_row.
        IF lo_autofilter IS NOT INITIAL.
          lv_row_hidden = lo_autofilter->is_row_hidden( iv_row = <fs_sheet_content>-cell_row ).
        ENDIF.
      ENDIF.
      lv_previous_row = <fs_sheet_content>-cell_row.

* --- Add empty rows
      WHILE lv_row < <fs_sheet_content>-cell_row.
        CONCATENATE lv_string zcl_excel_writer_csv=>eol INTO lv_string.
        lv_row = lv_row + 1.
        lv_col = 1.
      ENDWHILE.

* --- Skip hidden rows
      IF lv_row_hidden = abap_true.
        lv_row = <fs_sheet_content>-cell_row + 1.
        lv_col = 1.
        CONTINUE.
      ENDIF.

* --- Add empty columns
      WHILE lv_col < <fs_sheet_content>-cell_column.
        CONCATENATE lv_string zcl_excel_writer_csv=>delimiter INTO lv_string.
        lv_col = lv_col + 1.
      ENDWHILE.

* --- Skip hidden columns
      IF skip_hidden_columns = abap_true AND
          lo_worksheet->get_column( ip_column = <fs_sheet_content>-cell_column )->get_visible( ) = abap_false.
        lv_col = <fs_sheet_content>-cell_column + 1.
        CONTINUE.
      ENDIF.

*   --- Retrieve Cell Style format and data type
      CLEAR ls_numfmt.
      IF <fs_sheet_content>-data_type IS INITIAL AND <fs_sheet_content>-cell_style IS NOT INITIAL.
        lo_iterator = excel->get_styles_iterator( ).
        WHILE lo_iterator->has_next( ) EQ abap_true.
          lo_style ?= lo_iterator->get_next( ).
          CHECK lo_style->get_guid( ) = <fs_sheet_content>-cell_style.
          ls_numfmt     = lo_style->number_format->get_structure( ).
          EXIT.
        ENDWHILE.
      ENDIF.
      IF <fs_sheet_content>-data_type IS INITIAL AND ls_numfmt IS NOT INITIAL.
        " determine data-type
        CLEAR lv_attrname.
        CONCATENATE '''' ls_numfmt-numfmt '''' INTO ls_numfmt-numfmt.
        TRANSLATE ls_numfmt-numfmt TO UPPER CASE.
        READ TABLE lt_format INTO ls_format WITH KEY attvalue = ls_numfmt-numfmt.
        IF sy-subrc = 0.
          lv_attrname = ls_format-cmpname.
        ENDIF.

        IF lv_attrname IS NOT INITIAL.
          FIND FIRST OCCURRENCE OF 'DATETIME' IN lv_attrname.
          IF sy-subrc = 0.
            <fs_sheet_content>-data_type = 'd'.
          ELSE.
            FIND FIRST OCCURRENCE OF 'TIME' IN lv_attrname.
            IF sy-subrc = 0.
              <fs_sheet_content>-data_type = 't'.
            ELSE.
              FIND FIRST OCCURRENCE OF 'DATE' IN lv_attrname.
              IF sy-subrc = 0.
                <fs_sheet_content>-data_type = 'd'.
              ELSE.
                FIND FIRST OCCURRENCE OF 'CURRENCY' IN lv_attrname.
                IF sy-subrc = 0.
                  <fs_sheet_content>-data_type = 'n'.
                ELSE.
                  FIND FIRST OCCURRENCE OF 'NUMBER' IN lv_attrname.
                  IF sy-subrc = 0.
                    <fs_sheet_content>-data_type = 'n'.
                  ELSE.
                    FIND FIRST OCCURRENCE OF 'PERCENTAGE' IN lv_attrname.
                    IF sy-subrc = 0.
                      <fs_sheet_content>-data_type = 'n'.
                    ENDIF. " Purcentage
                  ENDIF. " Number
                ENDIF. " Currency
              ENDIF. " Date
            ENDIF. " TIME
          ENDIF. " DATETIME
        ENDIF. " lv_attrname IS NOT INITIAL.
      ENDIF. " <fs_sheet_content>-data_type IS INITIAL AND ls_numfmt IS NOT INITIAL.

* ----- Use format to determine the data type and display format.
      CASE <fs_sheet_content>-data_type.

        WHEN 'd' OR 'D'.
          IF <fs_sheet_content>-cell_value IS INITIAL AND initial_ext_date <> c_default.
            lc_value = initial_ext_date.
          ELSE.
            lc_value = zcl_excel_common=>excel_string_to_date( ip_value = <fs_sheet_content>-cell_value ).
            TRY.
                lv_date = lc_value.
                " ABAP Cloud: CONVERT_DATE_TO_EXTERNAL is not released -> user specific date format
                lv_tmp = |{ lv_date DATE = USER }|.
                lc_value = lv_tmp.

              CATCH cx_sy_conversion_no_number ##NO_HANDLER.

            ENDTRY.
          ENDIF.

        WHEN 't' OR 'T'.
          lc_value = zcl_excel_common=>excel_string_to_time( ip_value = <fs_sheet_content>-cell_value ).
          IF strlen( lc_value ) >= 6.
            lv_time = |{ lc_value(2) }:{ lc_value+2(2) }:{ lc_value+4(2) }|.
            lc_value = lv_time.
          ENDIF.
        WHEN OTHERS.
          lc_value = <fs_sheet_content>-cell_value.

      ENDCASE.

      CONCATENATE zcl_excel_writer_csv=>enclosure zcl_excel_writer_csv=>enclosure INTO lv_tmp.
      CONDENSE lv_tmp.
      REPLACE ALL OCCURRENCES OF zcl_excel_writer_csv=>enclosure IN lc_value WITH lv_tmp.

      FIND FIRST OCCURRENCE OF zcl_excel_writer_csv=>delimiter IN lc_value.
      IF sy-subrc = 0.
        CONCATENATE lv_string zcl_excel_writer_csv=>enclosure lc_value zcl_excel_writer_csv=>enclosure INTO lv_string.
      ELSE.
        CONCATENATE lv_string lc_value INTO lv_string.
      ENDIF.

    ENDLOOP.

    CLEAR ep_content.

    ep_content = cl_abap_conv_codepage=>create_out( )->convert( lv_string ).

  ENDMETHOD.


  METHOD set_active_sheet_index.
    CLEAR worksheet_name.
    worksheet_index = i_active_worksheet.
  ENDMETHOD.


  METHOD set_active_sheet_index_by_name.
    CLEAR worksheet_index.
    worksheet_name = i_worksheet_name.
  ENDMETHOD.


  METHOD set_delimiter.
    delimiter = ip_value.
  ENDMETHOD.


  METHOD set_enclosure.
    zcl_excel_writer_csv=>enclosure = ip_value.
  ENDMETHOD.


  METHOD set_endofline.
    zcl_excel_writer_csv=>eol = ip_value.
  ENDMETHOD.


  METHOD set_initial_ext_date.
    initial_ext_date = ip_value.
  ENDMETHOD.


  METHOD set_skip_hidden_rows.
    skip_hidden_rows = ip_value.
  ENDMETHOD.


  METHOD set_skip_hidden_columns.
    skip_hidden_columns = ip_value.
  ENDMETHOD.


  METHOD zif_excel_writer~write_file.
    me->excel = io_excel.
    ep_file = me->create( ).
  ENDMETHOD.
ENDCLASS.
