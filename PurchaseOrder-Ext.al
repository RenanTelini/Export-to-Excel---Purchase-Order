pageextension 50101 "Export Purchase Order to Excel" extends "Purchase Order"
{
    actions
    {
        addafter(AttachAsPDF)
        {
            action(ExportExcel)
            {
                Caption = 'Export to Excel';
                Image = Export;
                ApplicationArea = All;
                ToolTip = 'Export the purchase order to an Excel file.';
                RunObject = Report "Export to Excel - Purchase";
            }
        }
    }
}