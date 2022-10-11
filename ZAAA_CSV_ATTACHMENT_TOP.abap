*&---------------------------------------------------------------------*
*&  Include           ZAAA_CSV_ATTACHMENT_TOP
*&---------------------------------------------------------------------*



TABLES : adr6.
TYPE-POOLS: truxs.

DATA: gr_table   TYPE REF TO cl_salv_table,
      gt_sflight TYPE TABLE OF sflight.

DATA: gv_xstring      TYPE xstring,  " cl_salv_table wich will be converted to xstring
      email           TYPE adr6-smtp_addr,
      gv_xlen         TYPE int4,
      gr_request      TYPE REF TO cl_bcs, "to create the send request
      gv_body_text    TYPE bcsy_text,
      gv_subject      TYPE so_obj_des,
      gv_att          TYPE so_obj_des,
      gr_recipient    TYPE REF TO if_recipient_bcs,  "create list of recipient to distribute emails
      gr_document     TYPE REF TO cl_document_bcs,   "message
      gv_size         TYPE so_obj_len,
      it_receivers    TYPE STANDARD TABLE OF  somlreci1,
      wa_it_receivers LIKE LINE OF it_receivers,
      it_message      TYPE STANDARD TABLE OF solisti1,
      wa_it_message   LIKE LINE OF it_message,
      c1(99)          TYPE c,
      c2(15)          TYPE c,
      ls_sflight      TYPE sflight,
      gt_data_csv     TYPE TABLE OF string,
      ls_data_csv     TYPE string,
      lv_price        TYPE string,
      lv_seatsmax     TYPE string,
      lv_seatsocc     TYPE string,
      main_text       TYPE bcsy_text.
