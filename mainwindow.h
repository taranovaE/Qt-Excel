/**
  * @file mainwindow.h
  * @brief window with table
  * @author Agapova Ekaterina
  */
#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <ActiveQt/QAxObject>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private:
    Ui::MainWindow *ui;
    QAxObject *mExcel;
    QAxObject *workbook;
    QAxObject *mSheets;
    QAxObject *StatSheet;

private slots:
    void createTable();
    void openExcel(QString file);
    void saveTable();
    void cellsDef();
};
#endif // MAINWINDOW_H
