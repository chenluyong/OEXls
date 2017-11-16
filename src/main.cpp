
#include <QDebug>
#include "oeexcelwriter.h"
#include "oeexcelreader.h"

int main(void) {
    OEExcelWriter* writer = new OEExcelWriter();

    writer->CreateWorkSheet("asd");
    for (int i = 0; i < 1000; ++i)
        writer->SetCellText(i,0,QString::number(i).toStdString(),false);

    writer->SaveExcelFile("F:/xlstest.xls");

    qDebug() << "write done";

    OEExcelReader* reader = new OEExcelReader("F:/xlstest.xls");

    for (int i = 0; i < 1000; ++i)
        qDebug() << i << reader->getCellValue(i,0).c_str();



    reader->closeWorkBook();


}
