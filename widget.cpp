#include "widget.h"
#include "ui_widget.h"
#include "excel_control.h"
#include <QDebug>
#include <QTime>
#include <QFileDialog>
#include <QtConcurrent>
#include "windows.h"


Widget::Widget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Widget)
{
    ui->setupUi(this);



}

Widget::~Widget()
{
    delete ui;
}

void Widget::on_pushButton_clicked()
{
    QtConcurrent::run([&]()
    {
        HRESULT r = OleInitialize(0);
        if(r != S_OK && r != S_FALSE)
        {
            qDebug()<<"Qt: Could not initialize OLE(error %x)"<<(unsigned int)r;
        }

        ExcelControl*   excel = new ExcelControl;
        excel->NewFile();

        QVariant var;
        QList<QVariant> rowList;
        QList<QList<QVariant>> rowsList;


        for(int i=1; i<11; i++)
        {
            for(int j=1; j<11; j++)
            {
                var = QString("测试数据");
                rowList.append(var);
            }
            rowsList.append(rowList);
        }
        QTime time;
        time.start();

        excel->WriteCells(rowsList);
        excel->SaveAt("C:/Users/huangxi/Desktop/test1.xlsx");
        excel->CloseFile();

        qDebug()<<"time: "<<time.elapsed();

        delete excel;

        OleUninitialize();
    });


}

void Widget::on_pushButton_2_clicked()
{
    QtConcurrent::run([&]()
    {
        HRESULT r = OleInitialize(0);
        if(r != S_OK && r != S_FALSE)
        {
            qDebug()<<"Qt: Could not initialize OLE(error %x)"<<(unsigned int)r;
        }

        ExcelControl*   excel = new ExcelControl;
        excel->NewFile();


        QTime time;
        time.start();

        for(int i=1; i<11; i++)
        {
            for(int j=1; j<11; j++)
            {
                excel->WriteCell(i, j, "测试数据");
            }
        }

        excel->SaveAt("C:/Users/huangxi/Desktop/test2.xlsx");
        excel->CloseFile();

        qDebug()<<"time: "<<time.elapsed();

        delete excel;

        OleUninitialize();
    });
}

void Widget::on_pushButton_3_clicked()
{
    ExcelControl*   excel = new ExcelControl;
    QString path = QFileDialog::getOpenFileName(this, tr("选择excel文件"),".\\","excel(*.xlsx)");
    if(!excel->OpenFile(QDir::toNativeSeparators(path)))
    {
        qDebug()<<"open excel file error!";
        return;
    }

    QVariant allData = excel->ReadCells();

    for(int i=1; i<6; i++)
    {
        qDebug()<<excel->FindDataFromAll(1, i, allData);
    }
    excel->CloseFile();
    delete excel;
}

void Widget::on_pushButton_4_clicked()
{
    ExcelControl*   excel = new ExcelControl;
    QString path = QFileDialog::getOpenFileName(this, tr("选择excel文件"),".\\","excel(*.xlsx)");
    if(!excel->OpenFile(QDir::toNativeSeparators(path)))
    {
        qDebug()<<"open excel file error!";
        return;
    }

    QString data;
    for(int i=1; i<6; i++)
    {
        bool ret = excel->ReadCell(1, i, data);
        qDebug()<<"ret: "<<ret <<"data: "<<data;
    }

    excel->CloseFile();
    delete excel;
}
