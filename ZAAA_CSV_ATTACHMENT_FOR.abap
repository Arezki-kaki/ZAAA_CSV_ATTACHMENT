*&---------------------------------------------------------------------*
*&  Include           ZAAA_CSV_ATTACHMENT_FOR
*&---------------------------------------------------------------------*


DATA lv_string      TYPE string.
DATA binary_content TYPE solix_tab.

CONSTANTS:
  gc_tab  TYPE c VALUE cl_bcs_convert=>gc_tab,
  gc_crlf TYPE c VALUE cl_bcs_convert=>gc_crlf.



LOOP AT so_mail INTO DATA(t_so_mail).
  CLEAR wa_it_receivers.
  wa_it_receivers-receiver   = t_so_mail-low.
  wa_it_receivers-rec_type   = 'U'.
  wa_it_receivers-com_type   = 'INT'.
  wa_it_receivers-notif_del  = 'X'.
  wa_it_receivers-notif_ndel = 'X'.
  APPEND wa_it_receivers TO it_receivers.
ENDLOOP.

**********************************************************************
" Get data
SELECT *
  FROM sflight
  INTO CORRESPONDING FIELDS OF TABLE gt_sflight.
IF sy-subrc EQ 0. ENDIF.


TRY.
    "Send request
    cl_salv_table=>factory(
    IMPORTING
      r_salv_table    = gr_table
    CHANGING
      t_table         = gt_sflight ).
  CATCH cx_salv_msg ##NO_HANDLER.
ENDTRY.


"Testing display
*gr_table->display( ).

TRY.

    gv_xstring = gr_table->to_xml( if_salv_bs_xml=>c_type_xlsx ).


    LOOP AT gt_sflight INTO ls_sflight.
      CLEAR : ls_data_csv,
              lv_price,
              lv_seatsmax,
              lv_seatsocc.

      lv_price    = ls_sflight-price.
      lv_seatsmax = ls_sflight-seatsmax.
      lv_seatsocc = ls_sflight-seatsocc.

      CONCATENATE ls_sflight-carrid
                  ls_sflight-connid
                  ls_sflight-fldate
                  lv_price
                  ls_sflight-currency
                  ls_sflight-planetype
                  lv_seatsmax
                  lv_seatsocc
             INTO ls_data_csv
     SEPARATED BY ';'.

      APPEND ls_data_csv TO gt_data_csv.
    ENDLOOP.

    LOOP AT gt_data_csv INTO ls_data_csv .
      CASE sy-tabix.
        WHEN 1.
          CONCATENATE ls_data_csv gc_crlf INTO lv_string .
        WHEN OTHERS.
          CONCATENATE lv_string ls_data_csv gc_crlf INTO lv_string .
      ENDCASE.
    ENDLOOP.


    TRY.
        cl_bcs_convert=>string_to_solix(
          EXPORTING
            iv_string   = lv_string
            iv_codepage = '4103'  "suitable for MS Excel, leave empty
            iv_add_bom  = 'X'     "for other doc types
          IMPORTING
            et_solix  = binary_content
            ev_size   = gv_size ).
      CATCH cx_bcs.
        MESSAGE e445(so).
    ENDTRY.

    gr_request = cl_bcs=>create_persistent( ).

    APPEND 'Sample body text' TO gv_body_text.
    gv_subject = 'MY ATTACHMENT'.


    gv_subject = 'Send Mail from ABAP Program.'.
    gv_att = 'ATTACHMENT'.

    CLEAR wa_it_message.
    c1 = 'Bonjour'(005).
    c2 = 'Nom 1'.
    CONCATENATE c1 c2 ',' INTO
    wa_it_message-line SEPARATED BY space.
    APPEND wa_it_message TO it_message.

    CLEAR wa_it_message.
    wa_it_message-line = '                               '.
    APPEND wa_it_message TO it_message.

    CLEAR wa_it_message.
    wa_it_message-line = 'Veuillez trouver ci-joint le detail de la facture !'(002).
    APPEND wa_it_message TO it_message.

    "Create the email document
    gr_document = cl_document_bcs=>create_document(
      i_type        = 'HTM'
      i_text        = it_message
      i_subject     = gv_subject ).

    "Create the Attachment
    gv_size = gv_xlen.
    gr_document->add_attachment(
      i_attachment_type = 'csv'
      i_attachment_subject = gv_att
      i_attachment_size = gv_size
      i_att_content_hex = binary_content ).

    gr_request->set_document( gr_document ).

    LOOP AT so_mail .
      "Create recipient object.
      email = so_mail-low .
      gr_recipient = cl_cam_address_bcs=>create_internet_address( email ).
      "add recipient object to send request
      gr_request->add_recipient( gr_recipient ).
    ENDLOOP.

    gr_request->send( ).

    MESSAGE 'Email sent.' TYPE 'S'.

    COMMIT WORK.


  CATCH cx_bcs.
    MESSAGE 'Email not sent' TYPE 'A'.
ENDTRY.
