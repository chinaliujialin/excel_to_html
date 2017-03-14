package com.chinadaas.riskbell;

import java.io.File;
import java.util.List;

/**
 * Created by pc on 2017/3/13.
 */
public class Test {
    private static List<File> filelist;
    public static void main(String[] args){
        File filedir;
        getFileList("F:\\EXCEL表格文件夹","F:\\EXCEL表格文件夹","F:\\testhahaha");
        System.out.print("done");
    }
    /**
     * 为文件生成多级目录
     *@param strPath,源目录文件根地址
     *@param ipath 源文件地址（用于剪切字符）
     *@param opath 目标文件根目录
     */
    public static void getFileList(String strPath,String ipath,String opath) {
        File dir = new File(strPath);
        File dirin = new File(ipath);
        File[] files = dir.listFiles(); // 该文件目录下文件全部放入数组
        if (files != null) {
            for (int i = 0; i < files.length; i++) {
                String fileName = files[i].getName();
                if (files[i].isDirectory()) { // 判断是文件还是文件夹,如果是文件夹获取在子目录中的文件夹名称
                    String filename = files[i].getAbsolutePath().replace(dirin.getAbsolutePath(),"") ;
                    System.out.println("---" + filename);
                    File outp = new File(opath+filename);
                    outp.mkdirs();
                    getFileList(files[i].getAbsolutePath(),ipath,opath); // 迭代
                } else if (fileName.endsWith("xlsx")) { // 判断文件名是否以.xlsx结尾
                    String strFileName = files[i].getName();
                    String strFiledir = files[i].getAbsolutePath();
                    //加工文件名称  1、去除输入路径的文件名：F:\AAAA\AAASS\A.XLDX----->F:\AAAA\AAASS\
                    String str1=strFiledir.replace(strFileName,"");
                    //加工文件名称  2、去除输入路径的根目录：F:\AAAA\AAASS\------>\AAASS\
                    String str2=str1.replace(ipath,"");
                    //加工文件名称  3、增加输出目录的根目录：\AAASS\------>F:\asd\AAASS\
                    String str3=opath+str2+strFileName.substring(0,strFileName.indexOf(".xlsx")) + ".html";
                    //输出文件  Els_x_toHtml（输入路径，输出路径）
                    File filetest = new File(str3);
                    filetest.mkdirs();
                    //System.out.println("---" + strFileName);
                    //filelist.add(files[i]);
                }else if (fileName.endsWith("xls")) { // 判断文件名是否以.xls结尾
                    String strFileName = files[i].getName();
                    String strFiledir = files[i].getAbsolutePath();
                    //加工文件名称  1、去除输入路径的文件名：F:\AAAA\AAASS\A.XLDX----->F:\AAAA\AAASS\
                    String str1=strFiledir.replace(strFileName,"");
                    //加工文件名称  2、去除输入路径的根目录：F:\AAAA\AAASS\------>\AAASS\
                    String str2=str1.replace(ipath,"");
                    //加工文件名称  3、增加输出目录的根目录：\AAASS\------>F:\asd\AAASS\
                    String str3=opath+str2+strFileName.substring(0,strFileName.indexOf(".xls")) + ".html";
                    //输出文件  Els_x_toHtml（输入路径，输出路径）
                    File filetest = new File(str3);
                    filetest.mkdirs();
                    //System.out.println("---" + strFileName);
                    //filelist.add(files[i]);
                } else {
                    continue;
                }
            }
        }
    }
}
