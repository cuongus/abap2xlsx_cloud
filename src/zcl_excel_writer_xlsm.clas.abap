CLASS zcl_excel_writer_xlsm DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_writer_2007
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_WRITER_XLSM
*"* do not include other source files here!!!
  PROTECTED SECTION.

*"* protected components of class ZCL_EXCEL_WRITER_XLSM
*"* do not include other source files here!!!
    CONSTANTS c_xl_vbaproject TYPE string VALUE 'xl/vbaProject.bin'. "#EC NOTEXT

    " CREATE is not redefined: ZCL_EXCEL_WRITER_2007->CREATE writes all parts (drawings, comments,
    " header / footer images, tables ...) and calls ADD_FURTHER_DATA_TO_ZIP, which adds xl/vbaProject.bin
    " (the former copy of CREATE wrote a drawing relation also for sheets without drawings)
    METHODS add_further_data_to_zip
        REDEFINITION .
    METHODS create_content_types
        REDEFINITION .
    METHODS create_xl_relationships
        REDEFINITION .
    METHODS create_xl_sheet
        REDEFINITION .
    METHODS create_xl_workbook
        REDEFINITION .
  PRIVATE SECTION.
*"* private components of class ZCL_EXCEL_WRITER_XLSM
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_WRITER_XLSM IMPLEMENTATION.


  METHOD add_further_data_to_zip.

    super->add_further_data_to_zip( io_zip = io_zip ).

* Add vbaProject.bin to zip
    io_zip->add( name    = me->c_xl_vbaproject
                 content = me->excel->zif_excel_book_vba_project~vbaproject ).

  ENDMETHOD.


  METHOD create_content_types.
** Constant node name
    DATA: lc_xml_node_workb_ct    TYPE string VALUE 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' ##NO_TEXT,
          lc_xml_node_default     TYPE string VALUE 'Default' ##NO_TEXT,
          " Node attributes
          lc_xml_attr_partname    TYPE string VALUE 'PartName',
          lc_xml_attr_extension   TYPE string VALUE 'Extension' ##NO_TEXT,
          lc_xml_attr_contenttype TYPE string VALUE 'ContentType',
          lc_xml_node_workb_pn    TYPE string VALUE '/xl/workbook.xml',
          lc_xml_node_bin_ext     TYPE string VALUE 'bin',
          lc_xml_node_bin_ct      TYPE string VALUE 'application/vnd.ms-office.vbaProject' ##NO_TEXT.


    DATA: lo_ixml          TYPE REF TO if_ixml_core,
          lo_document      TYPE REF TO if_ixml_document,
          lo_parser        TYPE REF TO if_ixml_parser_core,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_collection    TYPE REF TO if_ixml_node_collection,
          lo_iterator      TYPE REF TO if_ixml_node_iterator,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory_core,
          lo_ostream       TYPE REF TO if_ixml_ostream_core,
          lo_renderer      TYPE REF TO if_ixml_renderer_core.

    DATA: lv_contenttype TYPE string.

**********************************************************************
* STEP 3: Create standard contentType
    ep_content = super->create_content_types( ).

**********************************************************************
* STEP 2: modify XML adding the extension bin definition

    " ABAP Cloud: CL_XML_DOCUMENT is not released -> parse with the iXML core API
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_document = lo_ixml->create_document( ).
    lo_parser = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                        istream        = lo_streamfactory->create_istream_xstring( ep_content )
                                        document       = lo_document ).
    lo_parser->parse( ).
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    " extension node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_bin_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_bin_ct ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 3: modify XML changing the contentType of node Override /xl/workbook.xml

    lo_collection = lo_document->get_elements_by_tag_name( 'Override' ) ##NO_TEXT.
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lv_contenttype = lo_element->get_attribute_ns( lc_xml_attr_partname ).
      IF lv_contenttype EQ lc_xml_node_workb_pn.
        lo_element->remove_attribute_ns( lc_xml_attr_contenttype ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                      value = lc_xml_node_workb_ct ).
        EXIT.
      ENDIF.
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).

  ENDMETHOD.


  METHOD create_xl_relationships.

** Constant node name
    DATA: lc_xml_node_relationship TYPE string VALUE 'Relationship' ##NO_TEXT,
          " Node attributes
          lc_xml_attr_id           TYPE string VALUE 'Id',
          lc_xml_attr_type         TYPE string VALUE 'Type' ##NO_TEXT,
          lc_xml_attr_target       TYPE string VALUE 'Target' ##NO_TEXT,
          " Node id
          lc_xml_node_ridx_id      TYPE string VALUE 'rId#',
          " Node type
          lc_xml_node_rid_vba_tp   TYPE string VALUE 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
          " Node target
          lc_xml_node_rid_vba_tg   TYPE string VALUE 'vbaProject.bin'.

    DATA: lo_ixml          TYPE REF TO if_ixml_core,
          lo_document      TYPE REF TO if_ixml_document,
          lo_parser        TYPE REF TO if_ixml_parser_core,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory_core,
          lo_ostream       TYPE REF TO if_ixml_ostream_core,
          lo_renderer      TYPE REF TO if_ixml_renderer_core.

    DATA: lv_xml_node_ridx_id TYPE string,
          lv_size             TYPE i,
          lv_syindex(2)       TYPE c.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_relationships( ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    " ABAP Cloud: CL_XML_DOCUMENT is not released -> parse with the iXML core API
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_document = lo_ixml->create_document( ).
    lo_parser = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                        istream        = lo_streamfactory->create_istream_xstring( ep_content )
                                        document       = lo_document ).
    lo_parser->parse( ).
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).


    lv_size = excel->get_worksheets_size( ).

    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lv_size = lv_size + 4.
    lv_syindex = lv_size.
    SHIFT lv_syindex RIGHT DELETING TRAILING space.
    SHIFT lv_syindex LEFT DELETING LEADING space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_vba_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_vba_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).

  ENDMETHOD.


  METHOD create_xl_sheet.

** Constant node name
    DATA: lc_xml_attr_codename      TYPE string VALUE 'codeName'.

    DATA: lo_ixml          TYPE REF TO if_ixml_core,
          lo_document      TYPE REF TO if_ixml_document,
          lo_parser        TYPE REF TO if_ixml_parser_core,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_collection    TYPE REF TO if_ixml_node_collection,
          lo_iterator      TYPE REF TO if_ixml_node_iterator,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory_core,
          lo_ostream       TYPE REF TO if_ixml_ostream_core,
          lo_renderer      TYPE REF TO if_ixml_renderer_core.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_sheet( io_worksheet = io_worksheet
                                         iv_active    = iv_active ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    " ABAP Cloud: CL_XML_DOCUMENT is not released -> parse with the iXML core API
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_document = lo_ixml->create_document( ).
    lo_parser = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                        istream        = lo_streamfactory->create_istream_xstring( ep_content )
                                        document       = lo_document ).
    lo_parser->parse( ).
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    lo_collection = lo_document->get_elements_by_tag_name( 'sheetPr' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = io_worksheet->zif_excel_sheet_vba_project~codename_pr ).
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).
  ENDMETHOD.


  METHOD create_xl_workbook.

** Constant node name
    DATA: lc_xml_attr_codename      TYPE string VALUE 'codeName'.

    DATA: lo_ixml          TYPE REF TO if_ixml_core,
          lo_document      TYPE REF TO if_ixml_document,
          lo_parser        TYPE REF TO if_ixml_parser_core,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_collection    TYPE REF TO if_ixml_node_collection,
          lo_iterator      TYPE REF TO if_ixml_node_iterator,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory_core,
          lo_ostream       TYPE REF TO if_ixml_ostream_core,
          lo_renderer      TYPE REF TO if_ixml_renderer_core.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_workbook( ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    " ABAP Cloud: CL_XML_DOCUMENT is not released -> parse with the iXML core API
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_document = lo_ixml->create_document( ).
    lo_parser = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                        istream        = lo_streamfactory->create_istream_xstring( ep_content )
                                        document       = lo_document ).
    lo_parser->parse( ).
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    lo_collection = lo_document->get_elements_by_tag_name( 'fileVersion' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = me->excel->zif_excel_book_vba_project~codename ).
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

    lo_collection = lo_document->get_elements_by_tag_name( 'workbookPr' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = me->excel->zif_excel_book_vba_project~codename_pr ).
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml_core=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).
  ENDMETHOD.
ENDCLASS.
