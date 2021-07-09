#ifndef EXCHANGE_H
#define EXCHANGE_H

//添加头文件
#include <QVariant>
#include <ActiveQt/QAxObject>//Excel
#include <QDebug>//debug输出
#include <QDir>//保存路径

class Exchange
{
public:
    Exchange();
    void SetTitle1(const QVariant &var,const QVariant &value);
    void SetTitle2(const QVariant &var,const QVariant &value);
    void SetInteriorColor(const QVariant &var,const QVariant &value);
    void SetFontProperty(const QVariant &var,const QVariant &value1,const QVariant &value2,const QVariant &value3,const QVariant &value4);
    void WriteData(const QVariant &var,const QVariant &value);
    void CreateExcel();
    void AddData();
};

#endif // EXCHANGE_H
