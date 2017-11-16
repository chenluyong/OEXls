#include "oeexcelreader.h"

#include <libxl.h>


using namespace libxl;

class OEExcelReaderPrivate {
    OEExcelReaderPrivate(OEExcelReader* parent = OE_NULLPTR)
        :q_ptr(parent),lastSheet_(0), readNum_(0){
    }
private:


    Book *pWB;

    Sheet *pWS;

	std::string lastFileName_;

    int lastSheet_;
    int readNum_;
private:

    OEExcelReader * const q_ptr;
    OE_DECLARE_PUBLIC(OEExcelReader)
};

OEExcelReader::OEExcelReader(std::string _fileName)
    :d_ptr(new OEExcelReaderPrivate(this))
{
    OE_D(OEExcelReader);

    d->pWB = xlCreateBook();

    if (!d->pWB->load(_fileName.data()))
        return;


    d->lastFileName_ = _fileName;

    d->pWS = d->pWB->getSheet(0);
    if (!d->pWS)
        return;
}

int OEExcelReader::createWorkbooks()
{
    OE_D(OEExcelReader);


    if (d->pWB != OE_NULLPTR)
        return 1247;

    d->pWB = xlCreateBook();

    if (d->pWB == OE_NULLPTR)
        return -1;
    return 0;
}

int OEExcelReader::load(const std::string& _fileName)
{
    OE_D(OEExcelReader);

    if (!d->pWB->load(_fileName.data()))
        return -1;

    d->lastFileName_ = _fileName;

    return 0;
}

int OEExcelReader::selectSheet(int sheetIndex)
{
    OE_D(OEExcelReader);

    d->pWS = d->pWB->getSheet(sheetIndex);
    if (!d->pWS)
        return -1;

    d->lastSheet_ = sheetIndex;

    return 0;
}

int OEExcelReader::getSheetsCount()
{
    OE_D(OEExcelReader);

    if (d->pWB == OE_NULLPTR)
        return -1;

    return d->pWB->sheetCount();
}

std::string OEExcelReader::getSheetName(int sheetIndex)
{
    OE_D(OEExcelReader);

    if (d->pWB != OE_NULLPTR)
        return d->pWB->getSheet(sheetIndex)->name();

    return "";
}

double OEExcelReader::getCellDoubleValue(int row, int column)
{

    OE_D(OEExcelReader);


    if (d->readNum_ > 200) {

        closeWorkBook();
        if (0 != createWorkbooks()) {
            return -1;
        }

        if (0 != load(d->lastFileName_)) {
            return -1;
        }

        // ??????
        if (0 != selectSheet(d->lastSheet_)) {
            return -1;
        }
    }


    double ret_data;

    CellType data_type = d->pWS->cellType(row,column);

    switch(data_type) {

    case CellType::CELLTYPE_NUMBER: {
        ret_data = d->pWS->readNum(row,column);
        break;
    }

    default:
        ret_data = -1;
    }

    ++(d->readNum_);

    return ret_data;
}

std::string OEExcelReader::getCellValue(int row, int column)
{
    OE_D(OEExcelReader);

    // ?????300?��?????
    if (d->readNum_ > 200) {

        closeWorkBook();
        if (0 != createWorkbooks()) {
            return "error";
        }

        if (0 != load(d->lastFileName_)) {
            return "error";
        }

        // ??????
        if (0 != selectSheet(d->lastSheet_)) {
            return "error";
        }
    }


	std::string ret_data;

    CellType data_type = d->pWS->cellType(row,column);

    switch(data_type) {

    case CellType::CELLTYPE_NUMBER: {
        double num = d->pWS->readNum(row,column);
#ifdef QT_IDE
        ret_data = QString::number(num,10,4).toStdString();
#else
        char temp_str[256] = {};
        sprintf(temp_str, "%lf", num);
        ret_data = temp_str;
#endif
        break;
    }
    case CellType::CELLTYPE_STRING:
        ret_data = d->pWS->readStr(row,column);
        break;

    default:
        ret_data = "not read";
    }

    ++(d->readNum_);

    return ret_data;
}

int OEExcelReader::getUsedRowsCount()
{
    OE_D(OEExcelReader);
    if (d->pWS == OE_NULLPTR)
        return -1;

    return d->pWS->lastRow();
}

int OEExcelReader::getUsedColsCount()
{
    OE_D(OEExcelReader);
    if (d->pWS == OE_NULLPTR)
        return -1;

    return d->pWS->lastCol();
}

void OEExcelReader::closeWorkBook()
{
    OE_D(OEExcelReader);

    d->pWB->release();
    d->pWB = OE_NULLPTR;
    d->pWS = OE_NULLPTR;
}
