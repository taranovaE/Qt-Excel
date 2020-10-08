/**
  * @file mainwindow.cpp
  * @brief window with table
  * @author Agapova Ekaterina
  */
#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QFileDialog>
#include <QUrl>
#include <QMessageBox>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    connect(ui->create, &QPushButton::clicked,
            this, &MainWindow::createTable);
    connect(ui->save, &QPushButton::clicked,
            this, &MainWindow::saveTable);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::createTable()
{
    ui->tableWidget->setColumnCount(ui->col_edit->displayText().toInt());
    ui->tableWidget->setRowCount(ui->row_edit->displayText().toInt());
    ui->tableWidget->resizeRowsToContents();
    ui->tableWidget->resizeColumnsToContents();
}

void MainWindow::openExcel(QString file)
{
    mExcel = new QAxObject("Excel.Application", this);
    QAxObject *workbooks = mExcel->querySubObject( "Workbooks" );
    workbook = workbooks->querySubObject( "Open(const QString&)", QUrl::fromLocalFile(file) );
    mSheets = workbook->querySubObject( "Sheets" );
    int count = mSheets->dynamicCall("Count()").toInt();
    QString name;
    for (int i=1; i<=count; i++){
        QAxObject* sheet = mSheets->querySubObject( "Item( int )", i );
        name = sheet->dynamicCall("Name()").toString();
    }
    StatSheet = mSheets->querySubObject( "Item(const QVariant&)", QVariant(name) );
    StatSheet->setProperty("Name", "My table");
}

void MainWindow::saveTable()
{

    QString file = QFileDialog::getOpenFileName(this, tr("Open files"),
                                                QString(),
                                                tr("Excel Files (*.xlsx *.xls)"));
    this->openExcel(file);
    this->cellsDef();
    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");
    mExcel->dynamicCall("Quit()");
    QMessageBox::information(NULL,QObject::tr("Information"), tr("Таблица сохранена успешно"));
}

void MainWindow::cellsDef()
{
    QAxObject* Cell1 = StatSheet->querySubObject("Cells(QVariant&,QVariant&)", 1, 1);
    QAxObject* Cell2 = StatSheet->querySubObject("Cells(QVariant&,QVariant&)", ui->row_edit->text().toInt(), ui->col_edit->text().toInt());
    QAxObject* range = StatSheet->querySubObject("Range(const QVariant&,const QVariant&)", Cell1->asVariant(), Cell2->asVariant() );
    QList<QVariant> cellsList;
    QList<QVariant> rowsList;
    for (int i = 0; i < ui->row_edit->text().toInt(); i++)
    {
        cellsList.clear();
        for (int j = 0; j < ui->col_edit->text().toInt(); j++){
            QVariant myData;
            QModelIndex myIndex;
            myIndex = ui->tableWidget->model()->index(i, j, QModelIndex());
            myData = ui->tableWidget->model()->data(myIndex, Qt::DisplayRole);
            cellsList << myData;
        }
        rowsList << QVariant(cellsList);
    }
    range->setProperty("Value", QVariant(rowsList) );
}

