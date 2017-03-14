package com.chinadaas.riskbell;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Formatter;

/**
 * Created by pc on 2017/3/13.
 * @para 输入路径（文件名），输出路径（文件夹名）
 * 完成文件的对应输出
 *
 */
public class Els_x_toHtml {

    private final Workbook wb;
    private HSSFWorkbook hwb;
    private XSSFWorkbook xwb;
    private Appendable output;
    private Formatter out;
    private int sheetNo;
    private String sheetNames[];

    private String outputPath;
    private String inputPath;

    public Els_x_toHtml(String inputPath,String outputPath) throws IOException{
       this.outputPath=outputPath;
       this.inputPath=inputPath;
            try {
                wb = WorkbookFactory.create(new FileInputStream(inputPath));
                output= new PrintWriter(new FileWriter(outputPath));
                if (wb == null)
                    throw new NullPointerException("wb");
                if (output == null)
                    throw new NullPointerException("output");
                sheetNo = wb.getNumberOfSheets();
                getNames(sheetNo);
                printPage();
            } catch (Exception e){
                throw new IllegalArgumentException("Cannot create workbook from stream", e);
            }
        }
    private void constructTabHeader()
    {
        out.format("<div id=\"main\">");
        out.format("<div class=\"ui-widget-header ui-corner-top\" >%n");
        out.format("<ul>%n");
        for(int i=0;i<sheetNo;i++)
        {
            out.format("<li>%n");
            out.format("<a href =\"#tabs-"+(i+1)+"\">"+sheetNames[i]+"</a>%n");
            out.format("</li>%n");
        }
        out.format("</div>");
    }

    private void getNames(int num)
    {
        sheetNames = new String[num];
        for(int i=0;i<num;i++)
        {
            String name = wb.getSheetName(i);
            sheetNames[i]=name;
        }
    }

    //判断Excel版本
    private void printsheets() throws IOException{
        if (wb instanceof HSSFWorkbook)
        {
            hwb=new HSSFWorkbook(new FileInputStream(inputPath));
            int sheetnum=hwb.getNumberOfSheets();
            for(int nums=0;nums<sheetnum;nums++){
                HSSFSheet sheet = hwb.getSheetAt(nums);
                String name = sheet.getSheetName();
                int numr =sheet.getPhysicalNumberOfRows();
                int numc =sheet.getLeftCol();
                out.format("<p>工作页码"+nums+"</p>\n" +
                        "<p>.工作空间名称:"+name+"</p>\n"+
                        "<p>行列数："+ numr+" ,"+numc+"<p>");
                for(int i=0;i<numr;i++){
                    //System.out.print("i:"+i);
                    out.format("<table ><tr>\n");
                    HSSFRow hssfRow = sheet.getRow(i);
                    numc=hssfRow.getPhysicalNumberOfCells();
                    for(int j=0;j<numc;j++){
                        //System.out.print("j:"+j);
                        HSSFCell cell =hssfRow.getCell(j);
                        String str =getCellValue(cell);
                        out.format("<td>"+str+"</td>\n" );
                    }
                    out.format("</tr>\n</table>");
                }
            }
        }
        else if (wb instanceof XSSFWorkbook)
        {
            xwb=new XSSFWorkbook(new FileInputStream(inputPath));
            int sheetnum=xwb.getNumberOfSheets();
            for(int nums=0;nums<sheetnum;nums++){
                XSSFSheet sheet = xwb.getSheetAt(nums);
                String name = sheet.getSheetName();
                int numr =sheet.getPhysicalNumberOfRows();
                int numc =sheet.getLeftCol();
                out.format("<p>工作页码"+nums+"</p>\n" +
                        "<p>.工作空间名称:"+name+"</p>\n"+
                        "<p>行列数："+ numr+" ,"+numc+"<p>");
                for(int i=0;i<numr;i++){
                    XSSFRow xssfRow = sheet.getRow(i);
                    numc=xssfRow.getPhysicalNumberOfCells();
                    out.format("<table ><tr>\n");
                    for(int j=0;j<numc;j++){
                        XSSFCell cell =xssfRow.getCell(j);
                        String str =getCellValue(cell);
                        out.format("<td>"+str+"</td>\n" );
                    }
                    out.format("</tr>\n</table>");
                }
            }
        }
        else
            throw new IllegalArgumentException(
                    "unknown workbook type: " + wb.getClass().getSimpleName());
    }

    public void printPage() throws IOException {
        try {
            ensureOut();
            int sheetnum = wb.getNumberOfSheets();
            out.format("<!DOCTYPE html>%n");
            out.format("<html>%n");
            out.format("<head>%n");
            out.format("<meta charset=\"utf-8\">%n");
            out.format("<title>Excel转换HTML测试version1</title>%n");
            out.format("</head>%n");
            out.format("<body>%n");
            out.format("<style type=\"text/css\">\n" +
                    "td\n" +
                    "  {\n" +
                    "  height:18px;\n" +
                    "  width:130px;\n" +
                    "  vertical-align:middle;\n" +
                    "  }table,th\n" +
                    "  {\n" +
                    "  border: 1px solid black;\n" +
                    "  font-size:10px;\n" +
                    "  }table\n" +
                    "  {\n" +
                    "  border-collapse:collapse;\n" +
                    "  width:100%%;\n" +
                    "  }\n" +
                    "  </style>\n" +
                    "  </body>");

            printsheets();
            if (true) {
                out.format("</body>%n");
                out.format("</html>%n");
            }
        } finally {
            if (out != null)
                out.close();
            if (output instanceof Closeable) {
                Closeable closeable = (Closeable) output;
                closeable.close();
                wb.close();
            }
        }
    }

    private void ensureOut() {
        if (out == null)
            out = new Formatter(output);
    }

    private String getCellValue(Cell cell) {
        String cellvalue = "";
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellType()) {
                // 如果当前Cell的Type为NUMERIC
                case HSSFCell.CELL_TYPE_NUMERIC: {
                    short format = cell.getCellStyle().getDataFormat();
                    if(format == 14 || format == 31 || format == 57 || format == 58){   //excel中的时间格式
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        double value = cell.getNumericCellValue();
                        Date date = DateUtil.getJavaDate(value);
                        cellvalue = sdf.format(date);
                    }
                    // 判断当前的cell是否为Date
                    else if (HSSFDateUtil.isCellDateFormatted(cell)) {  //先注释日期类型的转换，在实际测试中发现HSSFDateUtil.isCellDateFormatted(cell)只识别2014/02/02这种格式。
                        // 如果是Date类型则，取得该Cell的Date值           // 对2014-02-02格式识别不出是日期格式
                        Date date = cell.getDateCellValue();
                        DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue= formater.format(date);
                    } else { // 如果是纯数字
                        // 取得当前Cell的数值
                        cellvalue = NumberToTextConverter.toText(cell.getNumericCellValue());

                    }
                    break;
                }
                // 如果当前Cell的Type为STRIN
                case HSSFCell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getStringCellValue().replaceAll("'", "''");
                    break;
                case  HSSFCell.CELL_TYPE_BLANK:
                    cellvalue = "";
                    break;
                // 默认的Cell值
                default:{
                    cellvalue = "";
                }
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;
    }
}

