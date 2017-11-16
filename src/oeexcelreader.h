#ifndef OEEXCELREADER_H
#define OEEXCELREADER_H

#include "auditxls_global.h"

#ifdef QT_IDE
#include <QString>
#else
#include <iostream>
#endif


class OEExcelReaderPrivate;

class AUDITXLSSHARED_EXPORT OEExcelReader
{
public:
    OEExcelReader(std::string _fileName);

    
    int createWorkbooks(void);

    int load(const std::string& _fileName);

    
    int selectSheet(int sheetIndex);

    
    int getSheetsCount();

    
    std::string getSheetName(int sheetIndex);

    
    double getCellDoubleValue(int row, int column);
    std::string getCellValue(int row, int column);

    
    int getUsedRowsCount();

    
    int getUsedColsCount();

    
    void closeWorkBook();

private:

    OEExcelReaderPrivate * const d_ptr;

	OE_DECLARE_PRIVATE(OEExcelReader)
	OE_DISABLE_COPY(OEExcelReader)
};

#endif // OEEXCELREADER_H
