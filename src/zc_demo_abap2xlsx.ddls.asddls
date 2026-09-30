@EndUserText.label: 'Demo ABAP2XLSX'
@ObjectModel.query.implementedBy: 'ABAP:ZCL_EXCEL_DEMO_CLOUD_TEST'
@Metadata.allowExtensions: true
define custom entity ZC_DEMO_ABAP2XLSX
  // with parameters parameter_name : parameter_type
{
      @UI.lineItem   : [{ position: 10 }]
      @UI.identification: [{ position: 10 }]
      @EndUserText.label: 'Company Code'
  key CompanyCode    : bukrs;

      @UI.lineItem   : [{ position: 20 }]
      @UI.identification: [{ position: 20 }]
      @EndUserText.label: 'Fiscal Year'
  key FiscalYear     : gjahr;

      @UI.lineItem   : [{ position: 30 }]
      @UI.identification: [{ position: 30 }]
      @EndUserText.label: 'Accounting Document'
  key Accounting_doc : belnr_d;

      @UI.lineItem   : [{ position: 40 }]
      @UI.identification: [{ position: 40 }]
      @EndUserText.label: 'Item Line'
  key Document_Itm   : abap.char(3);

      @UI.lineItem   : [{ position: 50 }]
      @UI.identification: [{ position: 50 }]
      @EndUserText.label: 'Mã Khách'
      Customer       : kunnr;

      @UI.lineItem   : [{ position: 60 }]
      @UI.identification: [{ position: 60 }]
      @EndUserText.label: 'Tên Khách hàng'
      Customer_Name  : abap.char( 255 );

      @UI.lineItem   : [{ position: 70 }]
      @UI.identification: [{ position: 70 }]
      @EndUserText.label: 'Địa chỉ khách hàng'
      Customer_Addr  : abap.char( 255 );

      @UI.lineItem   : [{ position: 80 }]
      @UI.identification: [{ position: 80 }]
      @EndUserText.label: 'Mã số thuế'
      Identification : abap.char( 25 );

      @UI.lineItem   : [{ position: 90 }]
      @UI.identification: [{ position: 90 }]
      @EndUserText.label: 'Số điện thoại'
      Phone_Number   : abap.char( 15 );

      @UI.lineItem   : [{ position: 100 }]
      @UI.identification: [{ position: 100 }]
      @EndUserText.label: 'Email'
      Email_Addr     : abap.char( 15 );

      @UI.lineItem   : [{ position: 110 }]
      @UI.identification: [{ position: 110 }]
      @EndUserText.label: 'Số lượng'
      Quantity       : abap.dec( 15, 3 );

      @UI.lineItem   : [{ position: 120 }]
      @UI.identification: [{ position: 120 }]
      @EndUserText.label: 'Đơn vị tính'
      DVT            : abap.char(3);

      @UI.lineItem   : [{ position: 130 }]
      @UI.identification: [{ position: 130 }]
      @EndUserText.label: 'Phát sinh có'
      Amount_Dr      : abap.dec( 15, 3 );

      @UI.lineItem   : [{ position: 140 }]
      @UI.identification: [{ position: 140 }]
      @EndUserText.label: 'Phát sinh nợ'
      Amount_Cr      : abap.dec( 15, 3 );

      @UI.lineItem   : [{ position: 150 }]
      @UI.identification: [{ position: 150 }]
      @EndUserText.label: 'Loại tiền'
      Currency       : abap.char( 3 );

      @UI.lineItem   : [{ position: 160 }]
      @UI.identification: [{ position: 160 }]
      @EndUserText.label: 'Company Name'
      Company_Name   : abap.char( 255 );

      @UI.lineItem   : [{ position: 170 }]
      @UI.identification: [{ position: 170 }]
      @Semantics.largeObject:
                       { mimeType     : 'MimeType',
      fileName       : 'Filename',
      contentDispositionPreference: #INLINE }
      Attachment     : abap.rawstring( 0 );
      @Semantics.mimeType    : true
      Mimetype       : abap.char(128);
      Filename       : abap.char(128);

      @UI.lineItem   : [{ position: 45 }]
      @UI.identification: [{ position: 45 }]
      @EndUserText.label: 'Ngày chứng từ'
      PostingDate    : abap.dats;
      @UI.lineItem   : [{ position: 165 }]
      @UI.identification: [{ position: 165 }]
      @EndUserText.label: 'Địa chỉ công ty'
      Company_Addr   : abap.char( 255 );

}
