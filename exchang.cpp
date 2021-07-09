#include "exchange.h"

QAxObject *excel;
QAxObject *workbooks;
QAxObject *workbook;
QAxObject *worksheets;
QAxObject *worksheet;
QAxObject *range;
QAxObject *interior;
QAxObject *cell;
QAxObject *font;
//Excel保存路径

QString path="C:/Users/dell/Desktop/Excel.xlsx";
Exchange::Exchange()
{

}
/***********************************
 * 函数功能：设置一级标题
 * var：单元格范围
 * value：一级标题名称
 * 创建时间：2018/10/17
 * 创建者：OYXL
************************************/
void Exchange::SetTitle1(const QVariant &var,const QVariant &value)
{
    range=worksheet->querySubObject("Range(const QString&)",var);
    range->setProperty("MergeCells",true);
    range->setProperty("Value",value);
}
/***********************************
 * 函数功能：设置二级标题
 * var：单元格范围
 * value：二级标题名称
 * 创建时间：2018/10/17
 * 创建者：OYXL
************************************/
void Exchange::SetTitle2(const QVariant &var,const QVariant &value)
{
    range=worksheet->querySubObject("Range(const QString&)",var);
    range->setProperty("Value",value);
}
/***********************************
 * 函数功能：按颜色序号设置背景色
 * var：单元格范围
 * value：颜色序号
 * 创建时间：2018/10/17
 * 创建者：OYXL
************************************/
void Exchange::SetInteriorColor(const QVariant &var,const QVariant &value)
{
    range=worksheet->querySubObject("Range(const QString&)",var);
    interior=range->querySubObject("Interior");
    interior->setProperty("ColorIndex",value);                              //按颜色序号进行颜色填充
}
/***********************************
 * 函数功能：设置字体属性
 * var：单元格范围
 * value1：列宽
 * value2：自动换行true或者false
 * value3：加粗true或者false
 * value4：颜色序号
 * 创建时间：2018/10/17
 * 创建者：OYXL
************************************/
void Exchange::SetFontProperty(const QVariant &var,const QVariant &value1,const QVariant &value2,const QVariant &value3,const QVariant &value4)
{
    range=worksheet->querySubObject("Range(const QString&)",var);
    range->setProperty("ColumnWidth",value1);
    range->setProperty("WrapText", value2);
    range->setProperty("HorizontalAlignment", -4108);//水平对齐：默认＝1,居中＝-4108,左＝-4131,右＝-4152
    range->setProperty("VerticalAlignment", -4108);//垂直对齐：默认＝2,居中＝-4108,左＝-4160,右＝-4107
    font = range->querySubObject("Font");//获取单元格字体
    font->setProperty("Name", QStringLiteral("微软雅黑"));//设置单元格字体
    font->setProperty("Bold", value3);//设置单元格字体加粗
    font->setProperty("Size", 12);//设置单元格字体大小
    font->setProperty("ColorIndex",value4);//按颜色序号进行颜色填充
}
/***********************************
 * 函数功能：将数据写入EXCEL
 * var：单元格范围
 * value：数据值
 * 创建时间：2018/10/17
 * 创建者：OYXL
************************************/
void Exchange::WriteData(const QVariant &var,const QVariant &value)
{
    range=worksheet->querySubObject("Range(const QString&)",var);
    range->setProperty("Value",value);
}
/***********************************
 * 函数功能：创建EXCEL表格
 * 创建时间：2018/7/5
 * 创建者：OYXL
************************************/
void Exchange::CreateExcel()
{
    Exchange change;

    excel = new QAxObject("Excel.Application");
    if (!excel)
    {
       qDebug()<<"创建Excel失败！";
    }

    excel->dynamicCall("SetVisible(bool Visible)", true);       //是否可视化excel
    excel->dynamicCall("SetUserControl(bool UserControl)", false);             //是否用户可操作
    //excel->setProperty("DisplayAlerts", true);                //是否弹出警告窗口
    workbooks = excel->querySubObject("WorkBooks");             //获取工作簿集合
    workbooks->dynamicCall("Add");                              //新建一个工作簿
    workbook = excel->querySubObject("ActiveWorkBook");         //获取当前工作簿
    worksheets = workbook->querySubObject("Sheets");            //获取工作表格集合
    worksheet  = worksheets->querySubObject("Item(int)", 1);    //获取当前工作表格1，即sheet1
    worksheet->setProperty("Name","恋爱数据");                  //修改sheet名称

    //<添加数据一级标题
    change.SetTitle1("A1:A2","序号");//<序号
    change.SetTitle1("B1:B2","时间和日期");//<时间和日期
    change.SetTitle1("C1:C2","恋爱模式");//<恋爱模式
    change.SetTitle1("D1:D2","姓名");//<姓名
    change.SetTitle1("E1:E2","性别");//<性别
    change.SetTitle1("F1:F2","年龄");//<年龄
    change.SetTitle1("G1:G2","签名");//<签名
    change.SetTitle1("H1:K1","爱好");//爱好

    //<添加数据二级标题
    change.SetTitle2("H2:H2","运动");//运动
    change.SetTitle2("I2:I2","音乐");//音乐
    change.SetTitle2("J2:J2","舞蹈");//舞蹈
    change.SetTitle2("K2:K2","游戏");//游戏

    //<颜色填充
    change.SetInteriorColor("A1:G1",3);
    change.SetInteriorColor("A2:G2",3);
    change.SetInteriorColor("H1:K1",4);
    change.SetInteriorColor("H2:K2",4);

    change.SetFontProperty("A1:A2",5,true,true,2);
    change.SetFontProperty("B1:B2",20,true,true,2);
    change.SetFontProperty("C1:F2",8,true,true,2);
    change.SetFontProperty("G1:G2",20,true,true,2);
    change.SetFontProperty("H1:K1",8,true,true,2);
    change.SetFontProperty("H2:K2",8,true,true,2);

    workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(path));
}
/*
 * 函数功能：添加实验数据至EXCEL
 * 创建时间：2018/7/5
 * 创建者：OYXL
*/
void Exchange::AddData()
{
    Exchange change;
    QString rowsNum="3";

    workbooks->dynamicCall("Open(const QString&)", QDir::toNativeSeparators(path));//打开工作簿
    workbook = excel->querySubObject("ActiveWorkBook");         //获取当前工作簿
    worksheets = workbook->querySubObject("Sheets");            //获取工作表格集合
    worksheet  = worksheets->querySubObject("Item(int)", 1);    //获取当前工作表格1，即sheet1

    change.WriteData("A"+rowsNum+":"+"A"+rowsNum,"1");//序号
    change.WriteData("B"+rowsNum+":"+"B"+rowsNum,"2018/11/6");//时间和日期
    change.WriteData("C"+rowsNum+":"+"C"+rowsNum,"日久生情");//恋爱模式
    change.WriteData("D"+rowsNum+":"+"D"+rowsNum,"可乐");//姓名
    change.WriteData("E"+rowsNum+":"+"E"+rowsNum,"女");//性别
    change.WriteData("F"+rowsNum+":"+"F"+rowsNum,"18");//年龄
    change.WriteData("G"+rowsNum+":"+"G"+rowsNum,"一只会拆家的二哈");//签名

    //<CH1实验数据
    change.WriteData("H"+rowsNum+":"+"H"+rowsNum,"狗刨");
    change.WriteData("I"+rowsNum+":"+"I"+rowsNum,"God is gril");
    change.WriteData("J"+rowsNum+":"+"J"+rowsNum,"转圈");
    change.WriteData("K"+rowsNum+":"+"K"+rowsNum,"LOL");

    //<整行处理
    //COLORREF ColorFont1=RGB(255,255,255);
    change.SetFontProperty("A"+rowsNum+":"+"A"+rowsNum,4,true,false,1);
    change.SetFontProperty("B"+rowsNum+":"+"B"+rowsNum,20,true,false,1);
    change.SetFontProperty("C"+rowsNum+":"+"F"+rowsNum,8,true,false,1);
    change.SetFontProperty("G"+rowsNum+":"+"G"+rowsNum,20,true,false,1);
    change.SetFontProperty("H"+rowsNum+":"+"K"+rowsNum,8,true,false,1);

    workbook->dynamicCall("Save()");//保存EXCEL
    //workbook->dynamicCall("Close()");//关闭工作簿
    //excel->dynamicCall("Quit()");//退出
}
