#include "excel_control.h"
#include <qdir.h>
#include <qdebug.h>

ExcelControl::ExcelControl(QObject *parent) : QObject(parent)
{
    m_excel = nullptr;
    m_workbooks = nullptr;
    m_workbook = nullptr;

    m_worksheets = nullptr;
    m_worksheet = nullptr;

    m_excel = new QAxObject(this);                              //建立excel操作对象
    m_excel->setControl("Excel.Application");                   //连接Excel控件
    m_excel->dynamicCall("SetVisible (bool Visible)","false");  //设置为不显示窗体
    m_excel->setProperty("DisplayAlerts", false);               //设置为不显示任何警告信息，如关闭时的是否保存提示
}

ExcelControl::~ExcelControl()
{

}

void ExcelControl::NewFile()
{
    m_workbooks = m_excel->querySubObject("WorkBooks");     //获取工作簿(excel文件)集合
    m_workbooks->dynamicCall("Add");                        //新建一个工作簿
    m_workbook = m_excel->querySubObject("ActiveWorkBook");//获取当前工作簿

    m_worksheets = m_workbook->querySubObject("Sheets");        //获取excel文件里面所有工作表
    m_worksheet = m_worksheets->querySubObject("Item(int)",1);  //默认获取第一个工作表
}

//功能：打开一个excel文件
//参数filePath：文件路径      如：C:\\test.xlsx
bool ExcelControl::OpenFile(QString filePath)
{
    m_workbooks = m_excel->querySubObject("WorkBooks");                         //获取工作簿(excel文件)集合
    m_workbook = m_workbooks->querySubObject("Open(const QString&)", filePath); //获取excel文件
    if(nullptr == m_workbook) return false;

    m_worksheets = m_workbook->querySubObject("Sheets");        //获取excel文件里面所有工作表
    m_worksheet = m_worksheets->querySubObject("Item(int)",1);  //默认获取第一个工作表
    if(nullptr == m_worksheet) return false;

    return true;
}

//功能： 根据坐标读入单个单元格
//参数row：   行坐标
//参数column：列坐标
//参数data：  保存读取单元格的数据
bool ExcelControl::ReadCell(int row, int column, QString& data)
{
    if(row == 0 || column == 0) return false;

    QAxObject* cell = m_worksheet->querySubObject("Cells(int,int)", row, column);//基于行列坐标
    data = cell->property("Value").toString();
    delete cell;

    if(data.isEmpty())
    {
        return false;
    }
    else
    {
        return true;
    }
}

//功能：根据坐标写入单个单元格   坐标位置从x：1， y：1开始
//参数row：   行坐标
//参数column：列坐标
//参数data：  要写入的数据
bool ExcelControl::WriteCell(int row, int column, QString& data)
{
    if(row == 0 || column == 0) return false;

    QAxObject* cell = m_worksheet->querySubObject("Cells(int,int)", row, column);//基于行列坐标
    int ret = cell->setProperty("Value", data);
    delete cell;

    return ret;
}

//功能：根据坐标写入单个单元格    重载函数
bool ExcelControl::WriteCell(int row, int column, char* data)
{
    if(row == 0 || column == 0) return false;

    QString str = data;
    QAxObject* cell = m_worksheet->querySubObject("Cells(int,int)", row, column);//基于行列坐标，x和y坐标从0开始
    int ret = cell->setProperty("Value", str);
    delete cell;

    return ret;
}

//功能：切换工作表      注意事项：切换索引必须时有效的，当切换的不存在的工作表时会出问题！
void ExcelControl::SwitchWorksheet(int index)
{
    m_worksheet = m_worksheets->querySubObject("Item(int)",index);
}


//功能：一次性读取所有单元格数据
//返回数据解析方式：通过FindDataFromAll()进行查找相应坐标的数据
QVariant ExcelControl::ReadCells()
{
    QAxObject* usedRange = m_worksheet->querySubObject("UsedRange");//获取用户区域范围
    QVariant allData = usedRange->dynamicCall("Value");//读取区域内所有值
    delete usedRange;

    return allData;
}


//功能：返回所有单元格数据中对应的坐标    注意：起始坐标x：1，y：1
//参数row：    行坐标
//参数column： 列坐标
//参数allData：通过ReasCells()得到的QVariant
QString ExcelControl::FindDataFromAll(int row, int column, QVariant allData)
{
    if(row == 0 || column == 0) return "Coordinate error";

    //将数据放入容器
    QList<QList<QVariant>> excelRows;
    auto rows = allData.toList();
    for(auto row:rows)
    {
        excelRows.append(row.toList());
    }

    //遍历查找数据数据
    for(int rowIndex=0; rowIndex<excelRows.count(); rowIndex++)//行
    {
        for(int columnIndex=0; columnIndex<excelRows.at(rowIndex).count(); columnIndex++)//列
        {
            if(rowIndex == (row - 1) && columnIndex == (column - 1))
            {
                return excelRows.at(rowIndex).at(columnIndex).toString();
            }
        }
    }
}


/*超大数据快速写入示例：
    ExcelControl*   excel = new ExcelControl;
    excel->CreateFile();

    QVariant var;
    QList<QVariant> rowList;
    QList<QList<QVariant>> rowsList;


    for(int i=0; i<5; i++)
    {
        for(int j=0; j<5; j++)
        {
            var = QString::number(j + i * 5);
            rowList.append(var);
        }
        rowsList.append(rowList);
    }

    excel->WriteCells(rowsList);

    excel->SaveAt("C:\\Users\\huang\\Desktop\\test3.xlsx");
    excel->CloseFile();

    delete excel;
*/
bool ExcelControl::WriteCells(const QList<QList<QVariant> > &cells)
{
    //将QList<QList<QVariant>转换成QList<QVariant>
    QVariant rowList;
    QList<QVariant> rowsList;
    for(int i=0; i<cells.count(); i++)
    {
        rowList = QVariant::fromValue(cells.at(i));
        rowsList.append(rowList);
    }

    //将QList<QVariant>转换为QVariant
    QVariant var = QVariant::fromValue(rowsList);
    int row = cells.size();
    int col = cells.at(0).size();

    //按照区域写入一大块数据
    QString rangStr;
    ConvertToColName(col,rangStr);
    rangStr += QString::number(row);
    rangStr = "A1:" + rangStr;      //从A1开始写起
    QAxObject *range = m_worksheet->querySubObject("Range(const QString&)",rangStr);
    bool ret;
    ret = range->setProperty("Value", var);

    delete range;
    return ret;
}

//功能：把列数转换为excel的字母列号
//参数data： 大于0的数
//返回参数res:  字母列号，如1->A 26->Z 27 AA
void ExcelControl::ConvertToColName(int data, QString &res)
{
    Q_ASSERT(data>0 && data<65535);
    int tempData = data / 26;
    if(tempData > 0)
    {
        int mode = data % 26;
        ConvertToColName(mode,res);
        ConvertToColName(tempData,res);
    }
    else
    {
        res=(To26AlphabetString(data)+res);
    }
}

//功能：数字转换为26字母
//参数data：  需要转换的数字
//返回值：    转换后的字符串   1->A  26->Z  27->AA
QString ExcelControl::To26AlphabetString(int data)
{
    QChar ch = data + 0x40;//A对应0x41
    return QString(ch);
}

//功能：保存文件
void ExcelControl::Save()
{
    m_workbook->dynamicCall("Save()");
}

//功能：另存为文件
//参数savePath：   保存路径        如： C:\\test.xlsx
void ExcelControl::SaveAt(QString savePath)
{
    qDebug()<<"savePath: "<<QDir::toNativeSeparators(savePath);
    m_workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(savePath));
}

void ExcelControl::CloseFile()
{
    m_workbook->dynamicCall("Close(Boolean)", false);//关闭excel文件
    m_excel->dynamicCall("Quit(void)");              //关闭excel操作对象
}
