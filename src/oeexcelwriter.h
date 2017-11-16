#ifndef OEEXCELWRITER_H
#define OEEXCELWRITER_H

#include "auditxls_global.h"

#ifdef QT_IDE
#include <QString>
#endif
#include <string>

class OEExcelWriterPrivate;

class AUDITXLSSHARED_EXPORT OEExcelWriter
{
public:
    OEExcelWriter();

    void CreateWorkSheet(const std::string& SheetName); 
    void SaveExcelFile(const std::string& FileName); 
    void SetCellNumber(int row,int col,int number);
    void SetCellText(int row,int col,const std::string & text,bool setFont);
    void SetCellText(int row,int col,const std::wstring & text,bool setFont);

    void SetCellSize(int row,int col,int row_width ,int col_hight); 
    void SetCellCenter();
    void SetCellLeft(); 
    void MergeCells(int begin_row,int begin_col,int end_row,int end_col);

private:
    OEExcelWriterPrivate * const d_ptr;

	OE_DECLARE_PRIVATE(OEExcelWriter)
	OE_DISABLE_COPY(OEExcelWriter)

};

#endif // OEEXCELWRITER_H
