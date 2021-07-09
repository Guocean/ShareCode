#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "exchange.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}


MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_btnCreateExcel_clicked()
{
    Exchange change;
    change.CreateExcel();
}

void MainWindow::on_btnAddData_clicked()
{
    Exchange change;
    change.AddData();
}
