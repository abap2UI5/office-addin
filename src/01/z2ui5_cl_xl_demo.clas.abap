CLASS z2ui5_cl_xl_demo DEFINITION PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES z2ui5_if_app.

    TYPES:
      BEGIN OF ty_s_material,
        matnr    TYPE string,
        maktx    TYPE string,
        mtart    TYPE string,
        quantity TYPE p LENGTH 13 DECIMALS 3,
        unit     TYPE string,
        price    TYPE p LENGTH 13 DECIMALS 2,
        created  TYPE d,
      END OF ty_s_material.
    TYPES ty_t_material TYPE STANDARD TABLE OF ty_s_material WITH EMPTY KEY.
    TYPES:
      BEGIN OF ty_s_column,
        key           TYPE string,
        header        TYPE string,
        number_format TYPE string,
      END OF ty_s_column.
    TYPES ty_t_column TYPE STANDARD TABLE OF ty_s_column WITH EMPTY KEY.
    TYPES:
      BEGIN OF ty_s_cell,
        col1 TYPE string,
        col2 TYPE string,
        col3 TYPE string,
        col4 TYPE string,
        col5 TYPE string,
        col6 TYPE string,
      END OF ty_s_cell.
    TYPES ty_t_cell TYPE STANDARD TABLE OF ty_s_cell WITH EMPTY KEY.

    DATA materials         TYPE ty_t_material.
    DATA columns           TYPE ty_t_column.
    DATA sheet_name        TYPE string.
    DATA selection         TYPE ty_t_cell.
    DATA selection_address TYPE string.
    DATA follow_selection  TYPE abap_bool.
    DATA excel_available   TYPE abap_bool.
    DATA status            TYPE string.
    DATA status_type       TYPE string.
    DATA status_visible    TYPE abap_bool.

  PROTECTED SECTION.
    DATA client  TYPE REF TO z2ui5_if_client.
    DATA exports TYPE i.

    METHODS view_display.
    METHODS on_event.
    METHODS show_status
      IMPORTING
        text TYPE string
        type TYPE string.
    METHODS model_init.

  PRIVATE SECTION.
ENDCLASS.


CLASS z2ui5_cl_xl_demo IMPLEMENTATION.

  METHOD z2ui5_if_app~main.

    " The demo of abap2UI5/office-addin: an abap2UI5 app in the task pane of
    " an Excel add-in, talking to the workbook through the frontend's custom
    " control z2ui5.cc.ExcelBridge (abap2UI5 #2847). The control reads its
    " bindings and does what the app asks for as a frontend action on its id:
    " write - the materials into a new sheet, as an Excel table - and read -
    " the user's selection into SELECTION. Both run without a roundtrip; the
    " control's events report back. Outside Excel the same view runs, and
    " every call ends in OnError.

    me->client = client.
    IF client->check_on_init( ).
      model_init( ).
      view_display( ).
    ELSEIF client->check_on_navigated( ).
      view_display( ).
    ELSEIF client->check_on_event( ).
      on_event( ).
    ENDIF.

  ENDMETHOD.

  METHOD view_display.

    DATA(view) = z2ui5_cl_ui5_view_builder=>factory(
        )->ele( n = `View` ns = `mvc`
            )->a( n = `xmlns`        v = `sap.m`
            )->a( n = `xmlns:mvc`    v = `sap.ui.core.mvc`
            )->a( n = `xmlns:z2ui5`  v = `z2ui5.cc`
            )->a( n = `displayBlock` v = `true`
            )->a( n = `height`       v = `100%` ).

    DATA(page) = view->ele( `Page`
        )->a( n = `title` v = `Materials` ).

    page->tag( n = `ExcelBridge` ns = `z2ui5`
        )->a( n = `id`                v = `excel`
        )->a( n = `rows`              v = client->_bind( materials )
        )->a( n = `columns`           v = client->_bind( columns )
        )->a( n = `target`            v = `newSheet`
        )->a( n = `sheetName`         v = client->_bind( sheet_name )
        )->a( n = `asTable`           v = `true`
        )->a( n = `selection`         v = client->_bind( selection )
        )->a( n = `selectionAddress`  v = client->_bind( selection_address )
        )->a( n = `selectionChange`   v = client->_bind( follow_selection )
        )->a( n = `available`         v = client->_bind( excel_available )
        )->a( n = `OnWritten`         v = client->_event( val = `XL_WRITTEN` arg = `${$parameters>/address}` )
        )->a( n = `OnRead`            v = client->_event( val = `XL_READ` arg = `${$parameters>/address}` )
        )->a( n = `OnSelectionChange` v = client->_event( val = `XL_SELECTION` arg = `${$parameters>/address}` )
        )->a( n = `OnError`           v = client->_event( val   = `XL_ERROR`
                                                          t_arg = VALUE #( ( `${$parameters>/message}` )
                                                                           ( `${$parameters>/code}` ) ) ) ).

    page->tag( `MessageStrip`
        )->a( n = `text`     v = client->_bind( status )
        )->a( n = `type`     v = client->_bind( status_type )
        )->a( n = `visible`  v = client->_bind( status_visible )
        )->a( n = `showIcon` v = `true`
        )->a( n = `class`    v = `sapUiTinyMargin` ).

    " a task pane is about 350 px wide: the actions wrap instead of
    " disappearing into an overflow menu
    page->ele( `FlexBox`
        )->a( n = `wrap`       v = `Wrap`
        )->a( n = `alignItems` v = `Center`
        )->a( n = `class`      v = `sapUiTinyMarginBeginEnd`
        )->ele( `items`
            )->tag( `Button`
                )->a( n = `id`    v = `exportButton`
                )->a( n = `text`  v = `Export to Excel`
                )->a( n = `icon`  v = `sap-icon://excel-attachment`
                )->a( n = `type`  v = `Emphasized`
                )->a( n = `class` v = `sapUiTinyMarginEnd`
                )->a( n = `press` v = client->follow_up_action( val   = client->cs_event-control_by_id
                                                                t_arg = VALUE #( ( `excel` )
                                                                                 ( `write` ) ) )
            )->tag( `Button`
                )->a( n = `id`    v = `readButton`
                )->a( n = `text`  v = `Take selection`
                )->a( n = `icon`  v = `sap-icon://download`
                )->a( n = `class` v = `sapUiTinyMarginEnd`
                )->a( n = `press` v = client->follow_up_action( val   = client->cs_event-control_by_id
                                                                t_arg = VALUE #( ( `excel` )
                                                                                 ( `read` ) ) )
            )->tag( `Switch`
                )->a( n = `id`      v = `followSwitch`
                )->a( n = `state`   v = client->_bind( follow_selection )
                )->a( n = `enabled` v = client->_bind( excel_available )
            )->tag( `Label`
                )->a( n = `text`     v = `Follow selection`
                )->a( n = `labelFor` v = `followSwitch` ).

    page->ele( `Table`
        )->a( n = `id`    v = `materialTable`
        )->a( n = `items` v = client->_bind( materials )
        )->ele( `columns`
            )->ele( `Column`
                )->tag( `Text`
                    )->a( n = `text` v = `Material`
            )->end(
            )->ele( `Column`
                )->tag( `Text`
                    )->a( n = `text` v = `Description`
            )->end(
            )->ele( `Column`
                )->a( n = `minScreenWidth` v = `Tablet`
                )->a( n = `demandPopin`    v = `true`
                )->a( n = `hAlign`         v = `End`
                )->tag( `Text`
                    )->a( n = `text` v = `Quantity`
            )->end(
            )->ele( `Column`
                )->a( n = `minScreenWidth` v = `Tablet`
                )->a( n = `demandPopin`    v = `true`
                )->a( n = `hAlign`         v = `End`
                )->tag( `Text`
                    )->a( n = `text` v = `Price`
            )->end(
            )->ele( `Column`
                )->a( n = `minScreenWidth` v = `Tablet`
                )->a( n = `demandPopin`    v = `true`
                )->tag( `Text`
                    )->a( n = `text` v = `Created`
            )->end(
        )->end(
        )->ele( `items`
            )->ele( `ColumnListItem`
                )->ele( `cells`
                    )->tag( `ObjectIdentifier`
                        )->a( n = `title` v = `{MATNR}`
                        )->a( n = `text`  v = `{MTART}`
                    )->tag( `Text`
                        )->a( n = `text` v = `{MAKTX}`
                    )->tag( `Text`
                        )->a( n = `text` v = `{QUANTITY} {UNIT}`
                    )->tag( `Text`
                        )->a( n = `text` v = `{PRICE} EUR`
                    )->tag( `Text`
                        )->a( n = `text` v = `{CREATED}` ).

    page->tag( `Title`
        )->a( n = `text`  v = client->_bind( selection_address )
        )->a( n = `class` v = `sapUiSmallMarginBegin sapUiSmallMarginTop` ).

    page->ele( `Table`
        )->a( n = `id`              v = `selectionTable`
        )->a( n = `items`           v = client->_bind( selection )
        )->a( n = `noDataText`      v = `Select cells in Excel and press Take selection`
        )->ele( `columns`
            )->ele( `Column`
                )->tag( `Text`
                    )->a( n = `text` v = `A`
            )->end(
            )->ele( `Column`
                )->tag( `Text`
                    )->a( n = `text` v = `B`
            )->end(
            )->ele( `Column`
                )->tag( `Text`
                    )->a( n = `text` v = `C`
            )->end(
            )->ele( `Column`
                )->a( n = `minScreenWidth` v = `Tablet`
                )->a( n = `demandPopin`    v = `true`
                )->tag( `Text`
                    )->a( n = `text` v = `D`
            )->end(
            )->ele( `Column`
                )->a( n = `minScreenWidth` v = `Tablet`
                )->a( n = `demandPopin`    v = `true`
                )->tag( `Text`
                    )->a( n = `text` v = `E`
            )->end(
            )->ele( `Column`
                )->a( n = `minScreenWidth` v = `Tablet`
                )->a( n = `demandPopin`    v = `true`
                )->tag( `Text`
                    )->a( n = `text` v = `F`
            )->end(
        )->end(
        )->ele( `items`
            )->ele( `ColumnListItem`
                )->ele( `cells`
                    )->tag( `Text`
                        )->a( n = `text` v = `{COL1}`
                    )->tag( `Text`
                        )->a( n = `text` v = `{COL2}`
                    )->tag( `Text`
                        )->a( n = `text` v = `{COL3}`
                    )->tag( `Text`
                        )->a( n = `text` v = `{COL4}`
                    )->tag( `Text`
                        )->a( n = `text` v = `{COL5}`
                    )->tag( `Text`
                        )->a( n = `text` v = `{COL6}` ).

    client->view_display( view->stringify( ) ).

  ENDMETHOD.

  METHOD on_event.

    CASE client->get_event( ).
      WHEN `XL_WRITTEN`.
        exports = exports + 1.
        sheet_name = |Materials { exports + 1 }|.
        show_status( text = |Written to { client->get_event_arg( ) }| type = `Success` ).
      WHEN `XL_READ`.
        show_status( text = |{ lines( selection ) } rows taken from { client->get_event_arg( ) }|
                     type = `Information` ).
      WHEN `XL_SELECTION`.
        show_status( text = |Selected in Excel: { client->get_event_arg( ) }|
                     type = `Information` ).
      WHEN `XL_ERROR`.
        show_status( text = |Excel: { client->get_event_arg( ) } ({ client->get_event_arg( 2 ) })|
                     type = `Error` ).
    ENDCASE.

  ENDMETHOD.

  METHOD show_status.

    status = text.
    status_type = type.
    status_visible = abap_true.

  ENDMETHOD.

  METHOD model_init.

    sheet_name = `Materials 1`.
    status_type = `Information`.
    columns = VALUE #(
        ( key = `MATNR`    header = `Material`    number_format = `` )
        ( key = `MAKTX`    header = `Description` number_format = `` )
        ( key = `MTART`    header = `Type`        number_format = `` )
        ( key = `QUANTITY` header = `Quantity`    number_format = `#,##0.000` )
        ( key = `UNIT`     header = `Unit`        number_format = `` )
        ( key = `PRICE`    header = `Price (EUR)` number_format = `#,##0.00` )
        ( key = `CREATED`  header = `Created`     number_format = `yyyy-mm-dd` ) ).
    materials = VALUE #(
        ( matnr = `000000000000001001` maktx = `Hex bolt M8 x 40`        mtart = `ROH`  quantity = '1250'   unit = `PC` price = '0.12'    created = '20260105' )
        ( matnr = `000000000000001002` maktx = `Hex nut M8`              mtart = `ROH`  quantity = '4800'   unit = `PC` price = '0.04'    created = '20260105' )
        ( matnr = `000000000000001003` maktx = `Washer 8.4 mm`           mtart = `ROH`  quantity = '5200'   unit = `PC` price = '0.02'    created = '20260112' )
        ( matnr = `000000000000002001` maktx = `Steel sheet 2 mm`        mtart = `ROH`  quantity = '86.5'   unit = `M2` price = '24.90'   created = '20260203' )
        ( matnr = `000000000000002002` maktx = `Aluminium profile 40x40` mtart = `ROH`  quantity = '312.25' unit = `M`  price = '8.75'    created = '20260217' )
        ( matnr = `000000000000003001` maktx = `Gear housing`            mtart = `HALB` quantity = '40'     unit = `PC` price = '118.00'  created = '20260302' )
        ( matnr = `000000000000003002` maktx = `Drive shaft`             mtart = `HALB` quantity = '65'     unit = `PC` price = '74.30'   created = '20260310' )
        ( matnr = `000000000000004001` maktx = `Gearbox GX-200`          mtart = `FERT` quantity = '12'     unit = `PC` price = '1490.00' created = '20260401' )
        ( matnr = `000000000000004002` maktx = `Gearbox GX-400`          mtart = `FERT` quantity = '7'      unit = `PC` price = '2380.00' created = '20260415' )
        ( matnr = `000000000000005001` maktx = `Lubricant ISO VG 220`    mtart = `HIBE` quantity = '18.4'   unit = `L`  price = '6.95'    created = '20260502' ) ).

  ENDMETHOD.

ENDCLASS.
