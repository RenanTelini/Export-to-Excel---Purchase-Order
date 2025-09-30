report 50105 "Export to Excel - Purchase"
{
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = All;
    CaptionML = ENU = 'Export to Excel - Purchase',
    PTB = 'Exportar para Excel - Compras',
    ESS = 'Exportar a Excel - Compras';
    ProcessingOnly = True;

    dataset
    {
        dataitem(PurchaseHeader; "Purchase Header")
        {
            DataItemTableView = SORTING("Document Type", "No.");
            dataitem(PurchaseLine; "Purchase Line")
            {
                DataItemTableView = SORTING("Document Type", "Document No.", "Line No.");
                DataItemLinkReference = PurchaseHeader;
                DataItemLink = "Document No." = FIELD("No.");

                trigger OnPreDataItem()
                begin

                    gvRow := gvRow + 1;
                    EnterCell(gvRow, 2, FieldCaption("Document No."), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 3, FieldCaption(Type), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 4, FieldCaption("No."), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 5, FieldCaption("Line No."), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 6, FieldCaption(Description), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 7, FieldCaption("Location Code"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 8, FieldCaption(Quantity), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 9, FieldCaption("Reserved Quantity"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 10, FieldCaption("Unit of Measure Code"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 11, FieldCaption("Direct Unit Cost"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 12, FieldCaption("Area"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 13, FieldCaption("Line Amount"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 14, FieldCaption("Line Discount %"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 15, FieldCaption("Qty. to Receive"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 16, FieldCaption("Quantity Received"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 17, FieldCaption("Qty. to Invoice"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 18, FieldCaption("Quantity Invoiced"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 19, FieldCaption("Qty. to Assign"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 20, FieldCaption("Qty. Assigned"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 21, FieldCaption("Order Date"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 22, FieldCaption("Planned Receipt Date"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 23, FieldCaption("Expected Receipt Date"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 24, FieldCaption("Promised Receipt Date"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 25, FieldCaption("Requested Receipt Date"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 26, FieldCaption("Drop Shipment"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 27, FieldCaption("Variant Code"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 28, FieldCaption("Shortcut Dimension 1 Code"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 29, FieldCaption("Shortcut Dimension 2 Code"), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 30, FieldCaption("Blanket Order No."), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 31, FieldCaption("Blanket Order Line No."), gvLineBold, gvLineItalic, gvLineUnderline);
                    EnterCell(gvRow, 32, FieldCaption("Bin Code"), gvLineBold, gvLineItalic, gvLineUnderline);

                end;

                trigger OnAfterGetRecord()
                begin

                    gvRow := gvRow + 1;
                    EnterCell(gvRow, 2, "Document No.", False, False, False);
                    EnterCell(gvRow, 3, Format(Type), False, False, False);
                    EnterCell(gvRow, 4, "No.", False, False, False);
                    EnterCell(gvRow, 5, Format("Line No."), False, False, False);
                    EnterCell(gvRow, 6, Description, False, False, False);
                    EnterCell(gvRow, 7, "Location Code", False, False, False);
                    EnterCell(gvRow, 8, Format(Quantity), False, False, False);
                    EnterCell(gvRow, 9, Format("Reserved Quantity"), False, False, False);
                    EnterCell(gvRow, 10, "Unit of Measure Code", False, False, False);
                    EnterCell(gvRow, 11, Format("Direct Unit Cost"), False, False, False);
                    EnterCell(gvRow, 12, Area, False, False, False);
                    EnterCell(gvRow, 13, Format("Line Amount"), False, False, False);
                    EnterCell(gvRow, 14, Format("Line Discount %"), False, False, False);
                    EnterCell(gvRow, 15, Format("Qty. to Receive"), False, False, False);
                    EnterCell(gvRow, 16, Format("Quantity Received"), False, False, False);
                    EnterCell(gvRow, 17, Format("Qty. to Invoice"), False, False, False);
                    EnterCell(gvRow, 18, Format("Quantity Invoiced"), False, False, False);
                    EnterCell(gvRow, 19, Format("Qty. to Assign"), False, False, False);
                    EnterCell(gvRow, 20, Format("Qty. Assigned"), False, False, False);
                    EnterCell(gvRow, 21, Format("Order Date"), False, False, False);
                    EnterCell(gvRow, 22, Format("Planned Receipt Date"), False, False, False);
                    EnterCell(gvRow, 23, Format("Expected Receipt Date"), False, False, False);
                    EnterCell(gvRow, 24, Format("Promised Receipt Date"), False, False, False);
                    EnterCell(gvRow, 25, Format("Requested Receipt Date"), False, False, False);
                    EnterCell(gvRow, 26, Format("Drop Shipment"), False, False, False);
                    EnterCell(gvRow, 27, Format("Variant Code"), False, False, False);
                    EnterCell(gvRow, 28, "Shortcut Dimension 1 Code", False, False, False);
                    EnterCell(gvRow, 29, "Shortcut Dimension 2 Code", False, False, False);
                    EnterCell(gvRow, 30, "Blanket Order No.", False, False, False);
                    EnterCell(gvRow, 31, Format("Blanket Order Line No."), False, False, False);
                    EnterCell(gvRow, 32, "Bin Code", False, False, False);

                end;

                trigger OnPostDataItem()
                begin

                end;

            }

            trigger OnPreDataItem()
            begin

                grTempExcelBuffer.DELETEALL();
                CLEAR(grTempExcelBuffer);

                EnterCell(1, 1, gcText000, True, True, False);
                EnterCell(3, 1, FieldCaption("No."), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline); //header part of worksheet
                EnterCell(3, 2, FieldCaption("Buy-from Vendor No."), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 3, FieldCaption("Buy-from Contact No."), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 4, FieldCaption("Buy-from Vendor Name"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 5, FieldCaption("Buy-from Address"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 6, FieldCaption("Buy-from Address 2"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 7, FieldCaption("Buy-from Post Code"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 8, FieldCaption("Buy-from City"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 9, FieldCaption("Buy-from Contact No."), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 10, FieldCaption("Buy-from Contact"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 11, FieldCaption("No. of Archived Versions"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 12, FieldCaption("Posting Date"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 13, FieldCaption("Order Date"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 14, FieldCaption("Document Date"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 15, FieldCaption("Quote No."), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 16, FieldCaption("Vendor Order No."), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 17, FieldCaption("Vendor Shipment No."), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 18, FieldCaption("Vendor Invoice No."), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 19, FieldCaption("Order Address Code"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 20, FieldCaption("Purchaser Code"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 21, FieldCaption("Responsibility Center"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 22, FieldCaption("Assigned User ID"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 23, FieldCaption("Job Queue Status"), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);
                EnterCell(3, 24, FieldCaption(Status), gvHeaderBold, gvHeaderItalic, gvHeaderUnderline);

                gvRow := 3;

            end;

            trigger OnAfterGetRecord()
            begin

                gvRow := gvRow + 1;

                EnterCell(gvRow, 1, "No.", False, False, False); //rows part of worksheet
                EnterCell(gvRow, 2, "Buy-from Vendor No.", False, False, False);
                EnterCell(gvRow, 3, "Buy-from Contact No.", False, False, False);
                EnterCell(gvRow, 4, "Buy-from Vendor Name", False, False, False);
                EnterCell(gvRow, 5, "Buy-from Address", False, False, False);
                EnterCell(gvRow, 6, "Buy-from Address 2", False, False, False);
                EnterCell(gvRow, 7, "Buy-from Post Code", False, False, False);
                EnterCell(gvRow, 8, "Buy-from City", False, False, False);
                EnterCell(gvRow, 9, "Buy-from Contact No.", False, False, False);
                EnterCell(gvRow, 10, "Buy-from Contact", False, False, False);
                EnterCell(gvRow, 11, Format("No. of Archived Versions"), False, False, False);
                EnterCell(gvRow, 12, Format("Posting Date"), False, False, False);
                EnterCell(gvRow, 13, Format("Order Date"), False, False, False);
                EnterCell(gvRow, 14, Format("Document Date"), False, False, False);
                EnterCell(gvRow, 15, "Quote No.", False, False, False);
                EnterCell(gvRow, 16, "Vendor Order No.", False, False, False);
                EnterCell(gvRow, 17, "Vendor Shipment No.", False, False, False);
                EnterCell(gvRow, 18, "Vendor Invoice No.", False, False, False);
                EnterCell(gvRow, 19, "Order Address Code", False, False, False);
                EnterCell(gvRow, 20, "Purchaser Code", False, False, False);
                EnterCell(gvRow, 21, "Responsibility Center", False, False, False);
                EnterCell(gvRow, 22, "Assigned User ID", False, False, False);
                EnterCell(gvRow, 23, Format("Job Queue Status"), False, False, False);
                EnterCell(gvRow, 24, Format(Status), False, False, False);

            end;

            trigger OnPostDataItem()
            begin

                grTempExcelBuffer.CreateNewBook(gcText000);
                grTempExcelBuffer.WriteSheet(gcText000, COMPANYNAME, USERID);
                grTempExcelBuffer.CloseBook();
                grTempExcelBuffer.OpenExcel();

            end;
        }
    }

    requestpage
    {
        layout
        {
            area(Content)
            {
                group(Header)
                {
                    CaptionML = ENU = 'Header',
                    PTB = 'Cabeçalho',
                    ESS = 'Encabezado';
                    field(gvHeaderBold; gvHeaderBold)
                    {
                        ApplicationArea = All;
                        CaptionML = ENU = 'Bold header?',
                        PTB = 'Cabeçalho em negrito?',
                        ESS = '¿Encabezado en negrita?';

                    }
                    field(gvHeaderItalic; gvHeaderItalic)
                    {
                        ApplicationArea = All;
                        CaptionML = ENU = 'Header in italics?',
                        PTB = 'Cabeçalho em itálico?',
                        ESS = '¿Encabezado en cursiva?';

                    }
                    field(gvHeaderUnderline; gvHeaderUnderline)
                    {
                        ApplicationArea = All;
                        CaptionML = ENU = 'Underline header?',
                        PTB = 'Cabeçalho sublinhado?',
                        ESS = '¿Subrayar encabezado?';

                    }
                }
                group(Lines)
                {
                    CaptionML = ENU = 'Lines',
                    PTB = 'Linhas',
                    ESS = 'Líneas';
                    field(gvLineBold; gvLineBold)
                    {
                        ApplicationArea = All;
                        CaptionML = ENU = 'Bold lines?',
                        PTB = 'Linhas em negrito?',
                        ESS = '¿Líneas en negrita?';

                    }
                    field(gvLineItalic; gvLineItalic)
                    {
                        ApplicationArea = All;
                        CaptionML = ENU = 'Lines in italics?',
                        PTB = 'Linhas em itálico?',
                        ESS = '¿Líneas en cursiva?';

                    }
                    field(gvLineUnderline; gvLineUnderline)
                    {
                        ApplicationArea = All;
                        CaptionML = ENU = 'Lines with underline?',
                        PTB = 'Linhas com sublinhado?',
                        ESS = 'Líneas con subrayado?';

                    }

                }
            }
        }

    }


    var
        gvRow: Integer;
        grTempExcelBuffer: Record "Excel Buffer" Temporary;
        gvHeaderBold, gvHeaderItalic, gvHeaderUnderline : Boolean;
        gvLineBold, gvLineItalic, gvLineUnderline : Boolean;
        gcText000: TextConst ENU = 'Purchase Order',
        PTB = 'Pedido de Compra',
        ESS = 'Pedido de Compra';

    trigger OnInitReport()
    begin

    end;

    trigger OnPreReport()
    begin

    end;

    trigger OnPostReport()
    begin

    end;

    local procedure EnterCell(piRowNo: Integer; piColumnNo: Integer; piCellValue: Text[250];
        piBold: Boolean; piItalic: Boolean; piUnderLine: Boolean)
    begin
        //Insert on Excel Buffer
        grTempExcelBuffer.INIT;
        grTempExcelBuffer.VALIDATE("Row No.", piRowNo);
        grTempExcelBuffer.VALIDATE("Column No.", piColumnNo);
        grTempExcelBuffer."Cell Value as Text" := piCellValue;
        grTempExcelBuffer.Bold := piBold;
        grTempExcelBuffer.Italic := piItalic;
        grTempExcelBuffer.Underline := piUnderLine;
        grTempExcelBuffer.INSERT;
    end;
}