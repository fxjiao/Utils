package com.example.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 *
 * 文件需要是.xls文件,不然jxl解析不了
 * Created by fxjiao on 16/10/28.
 */
public class HDBuildPoint {

    /** 埋点文件名*/
    private static final String BUILD_POINT_EXCEL_PATH = "/Users/fxjiao/Downloads/1111_owner_20161021.xls";
    /** 输出结果样式:模块 */
    private static final String RESULT_STYLE_MODE = "// ------------------------%s模块 --------------------";
    /** 输出结果样式:版本 */
    private static final String RESULT_STYLE_VERSION = "//迭代1111";
    /** 输出结果样式:注释 */
    private static final String RESULT_STYLE_NOTE = "/**%s%s*/";
    /** 输出结果样式:变量定义 */
    private static final String RESULT_STYLE_DEFINE = "public static final String  %s = " + "\"" + "%s" + "\";";

    public static void main(String args[]) {
        readExcel();
    }

    public static String readExcel() {
        StringBuffer sb = new StringBuffer();
        Workbook wb = ExcelUtil.getReadableWorkBook(BUILD_POINT_EXCEL_PATH);
        if (wb != null) {
            Sheet[] sheets = wb.getSheets();
            if (sheets != null && sheets.length > 0) {
                for (int i = 0; i < sheets.length; i++) {
                    Sheet sheet = wb.getSheet(i);
                    int rows_len = sheet.getRows();

                    //每个sheet 首行不做处理:首行一般包括:0编号,1版本,2模块,3中文名,4英文名,5属性,6类型,7产品经理,8对应开发,9备注
                    // 其中需要用到的是:2模块,3中文名,4英文名,5属性
                    //首行处理:找出所用到属性所在的列
                    int posMode = 0;
                    int posChinease = 0;
                    int posEng = 0;
                    int posAttr = 0;
                    Cell[] cells0 = sheet.getRow(0);
                    int pos = 0;
                    for (Cell cell : cells0) {
                        String content = cell.getContents();
                        if (content != null) {
                            content = content.trim();
                            if (content.equals("模块")) {
                                posMode = pos;
                            } else if (content.equals("中文名")) {
                                posChinease = pos;
                            } else if (content.equals("英文名")) {
                                posEng = pos;
                            } else if (content.equals("属性")) {
                                posAttr = pos;
                            }
                        }
                        pos++;
                    }

                    //记录上一个模块的名称
                    String lastMode = "";
                    for (int j = 1; j < rows_len; j++) {
                        //获取所有的列
                        Cell[] cells = sheet.getRow(j);
                        if (cells != null && cells.length != 0) {
                            //需要将不同的2模块分区归类,归类为 //迭代xxxx
                            String mode = cells[posMode].getContents().trim();
                            String chs = cells[posChinease].getContents().trim();
                            String eng = cells[posEng].getContents().trim();
                            String attr = cells[posAttr].getContents().trim();
                            attr = attr.equals("操作类") ? "" : "(" + attr + ")";

                            if(chs.equals("")||eng.equals("")){
                                continue;
                            }
                            if (!mode.equals(lastMode)) {
                                lastMode = mode;
                                sb.append("\n");
                                sb.append(String.format(RESULT_STYLE_MODE,mode));
                                sb.append("\n");
                                sb.append(RESULT_STYLE_VERSION);
                            }

                            // 注释样式: /** 中文名(5属性)*/
                            sb.append("\n");
                            sb.append(String.format(RESULT_STYLE_NOTE,chs,attr));
                            String bianliangName = "";
                            String bianliangTag = "";
                            //存在中文的情况时需要做特殊处理
                            if (eng.length() < eng.getBytes().length) {
                                bianliangTag = eng.substring(eng.lastIndexOf("-"), eng.length()).replace("-", "+");
                                eng = eng.substring(0, eng.lastIndexOf("-") + 1);//通常情况带参数都是最好一节数据做说明
                            }
                            bianliangName = eng.toUpperCase().replace("-", "_") + bianliangTag;
                            sb.append("\n");
                            sb.append(String.format(RESULT_STYLE_DEFINE,bianliangName,eng));
                        }
                    }
                    //最后一列换行
                    sb.append("\t\n");
                }
                String result   =  sb.toString();
                System.out.print("--- result ---"+result.length()+"  " + sb.toString().substring(0,result.length()));
            }
        }
        return sb.toString();
    }


    //写Excel表格
    public static void writeExcel(String fileName) {
        String[] title = {"姓名", "英文姓名", "职位", "手机", "年龄", "工资"};
        try {
            OutputStream os = new FileOutputStream(fileName);
            //创建工作薄,一个参数
            WritableWorkbook wb = Workbook.createWorkbook(os);
            //创建工作表
            WritableSheet sheet = wb.createSheet("员工基本信息", 0);
            Label label;
            //填充表头
            for (int i = 0; i < title.length; i++) {
                // Label(x,y,z)其中x代表单元格的第x+1列，y代表单元格的第y+1行, 单元格的内容是z
                label = new Label(i, 0, title[i]);
                sheet.addCell(label);
            }

            //填充数据
            sheet.addCell(new Label(0, 1, "离歌"));
            sheet.addCell(new Label(1, 1, "bruce"));
            sheet.addCell(new Label(2, 1, "web programmer"));
            sheet.addCell(new Label(3, 1, "123456789"));

            jxl.write.Number number = new jxl.write.Number(4, 1, 25);
            sheet.addCell(number);

            //jxl会自动实现四舍五入
            jxl.write.NumberFormat format = new jxl.write.NumberFormat("#.##");
            jxl.write.WritableCellFormat wcf = new jxl.write.WritableCellFormat(format);
            jxl.write.Number nb = new jxl.write.Number(5, 1, 3200.000, wcf);
            sheet.addCell(nb);

            //定义字体等样式
            CellFormat cf = wb.getSheet(0).getCell(1, 0).getCellFormat();
            WritableCellFormat wcf2 = new WritableCellFormat(cf);
            wcf2.setAlignment(Alignment.CENTRE);
            wcf2.setBorder(Border.TOP, BorderLineStyle.THIN);
            wcf2.setBackground(Colour.RED);

            WritableFont font = new WritableFont(WritableFont.createFont("隶书"), 15);
            WritableCellFormat wfont = new WritableCellFormat(font);
            //这种方式是填充数据后设置的样式
            for (int j = 0; j < title.length; j++) {
                String content = sheet.getCell(j, 1).getContents();
                sheet.addCell(new Label(j, 1, content, wfont));
                //sheet.addCell(new Label(j,1,content,wcf2));
            }

            //进行写操作
            wb.write();
            wb.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

}
