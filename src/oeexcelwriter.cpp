#include "oeexcelwriter.h"

#ifdef QT_IDE
#include <QString>
#endif
#include <sstream>
#include <string>
#include <xlslib.h>
using namespace std;
using namespace xlslib_core;

class OEExcelWriterPrivate {
    OEExcelWriterPrivate(OEExcelWriter* parent)
        :q_ptr(parent){

    }

public:

    cell_t * cell;
    font_t * _font;
    xf_t * xf;
private:

    workbook pWB;

    worksheet *pWS;

    OEExcelWriter * const q_ptr;
    OE_DECLARE_PUBLIC(OEExcelWriter)
};


OEExcelWriter::OEExcelWriter()
    :d_ptr(new OEExcelWriterPrivate(this))
{

}

void OEExcelWriter::CreateWorkSheet(const string &SheetName)
{
    OE_D(OEExcelWriter);
	//OEExcelWriterPrivate* d = d_func();

    std::string sheetname = SheetName;  
    d->pWS = d->pWB.sheet(sheetname);
    d->pWS->defaultColwidth(25); 
    d->pWS->defaultRowHeight(15); 
    d->_font = d->pWB.font("Arial");
    d->_font->SetBoldStyle(BOLDNESS_BOLD); 
    d->_font->SetHeight(220);             

    d->xf = d->pWB.xformat();
    d->xf->SetFont(d->_font);
    d->xf->SetFillBGColor(CLR_WHITE);
    d->xf->SetFillFGColor(CLR_RED);
    d->pWS->MakeActive();
}

void OEExcelWriter::SaveExcelFile(const string &FileName)
{
    OE_D(OEExcelWriter);

    d->pWB.Dump(FileName);
}

void OEExcelWriter::SetCellNumber(int row, int col, int number)
{
    OE_D(OEExcelWriter);
    std::stringstream oss;
    std::string Num;
    oss << number;
    oss >> Num;
    d->cell = d->pWS->label(row, col, Num, OE_NULLPTR);
}

void OEExcelWriter::SetCellText(int row, int col, const string &text, bool setFont)
{
    OE_D(OEExcelWriter);
    std::string value = text;


    if (setFont == true) 
    {
        d->cell = d->pWS->label(row, col,value,d->xf);
    }
    else
    {
        d->cell = d->pWS->label(row, col, value,OE_NULLPTR);
    }
}

void OEExcelWriter::SetCellText(int row, int col, const wstring &text, bool setFont)
{
    OE_D(OEExcelWriter);
    std::wstring value = text;


    if (setFont == true)  
    {
        d->cell = d->pWS->label(row, col,value,d->xf);
    }
    else
    {
        d->cell = d->pWS->label(row, col, value,OE_NULLPTR);
    }
}

void OEExcelWriter::SetCellSize(int row, int col, int row_width, int col_hight)
{
    OE_D(OEExcelWriter);
    if(row_width!=0)
    {
        d->pWS->rowheight(row,row_width);
    }

    if (col_hight!=0)
    {
        d->pWS->colwidth(col,col_hight);
    }
}

void OEExcelWriter::SetCellCenter()
{
    OE_D(OEExcelWriter);
    d->cell->halign(HALIGN_CENTER); 
    d->cell->valign(VALIGN_CENTER); 
}

void OEExcelWriter::SetCellLeft()
{
    OE_D(OEExcelWriter);
    d->cell->halign(HALIGN_LEFT); 
    d->cell->valign(VALIGN_CENTER);
}

void OEExcelWriter::MergeCells(int begin_row, int begin_col, int end_row, int end_col)
{
    OE_D(OEExcelWriter);
    d->pWS->merge(begin_row,begin_col,end_row,end_col); 
}






