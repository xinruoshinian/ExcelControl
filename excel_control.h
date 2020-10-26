#ifndef EXCEL_CONTROL_H
#define EXCEL_CONTROL_H

/*
操作流程：
1、打开文件（OpenFile）   ->  读写     ->  保存（Save）      ->  关闭文件（CloseFile）
2、创建文件（CreateFile） ->  读写     ->  另存为(SaveAt)    ->  关闭文件（CloseFile）

注意事项：
1、更改数据以后一定要保存（Save）,不然文件数据不会发生改变
2、坐标起始位置是X：1 ，Y：1；
*/

#include <QObject>
#include <QAxObject>

class ExcelControl : public QObject
{
    Q_OBJECT
public:
    explicit ExcelControl(QObject *parent = nullptr);
    ~ExcelControl();

public:
    void NewFile();
    bool OpenFile(QString filePath);

    //读写单个单元格
    bool ReadCell(int row, int column, QString& data);
    bool WriteCell(int row, int column, QString& data);
    bool WriteCell(int row, int column, char* data);

    void SwitchWorksheet(int index);    //切换工作表

    //数据量大时采用这种方式
    QVariant ReadCells();                                           //读取全部数据
    QString  FindDataFromAll(int row, int column, QVariant allData);//查找数据，和ReadCells配套使用
    bool     WriteCells(const QList<QList<QVariant>> &cells);       //写入全部数据
    void     ConvertToColName(int data, QString &res);              //把列数转换为excel的字母列号
    QString  To26AlphabetString(int data);                          //数字转换为26字母

    void Save();
    void SaveAt(QString savePath);
    void CloseFile();

private:
    QAxObject*       m_excel;       //excel操作对象
    QAxObject*       m_workbooks;   //工作簿集合
    QAxObject*       m_workbook;    //工作簿（excel文件）
    QAxObject*       m_worksheets;  //工作表集合
    QAxObject*       m_worksheet;   //工作表
};

#endif // EXCEL_CONTROL_H
